using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Décrit et exécute les outils accessibles au modèle pour inspecter et modifier le projet VBE.</summary>
    internal sealed partial class LlmVbeTools
    {
        /// <summary>Session VBE utilisée pour exécuter les commandes sur le thread approprié.</summary>
        private readonly VbeSession session;
        /// <summary>Fenêtre propriétaire des demandes de confirmation et dialogues.</summary>
        private readonly IWin32Window owner;
        /// <summary>Paramètres utilisés par les outils et règles de modification.</summary>
        private readonly LlmSettings settings;
        /// <summary>Sérialiseur des arguments des outils et des réponses du pont.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        /// <summary>Requêtes explicites de l’utilisateur, utilisées pour limiter la lecture de fichiers locaux.</summary>
        private readonly List<string> userRequests = new List<string>();
        /// <summary>Se produit lorsqu’un outil a appliqué une modification de code.</summary>
        public event Action<CodeChange> CodeEdited;
        /// <summary>Signale une coupe récupérable effectuée dans le concepteur.</summary>
        public event Action<FormCutChange> FormCut;
        /// <summary>Noms des outils dont les opérations sont en lecture seule.</summary>
        private static readonly HashSet<string> ReadOnlyTools = new HashSet<string>(StringComparer.Ordinal) {
            "procedure_values_status", "preview_procedure_rename", "preview_class_member_rename", "read_project_protection", "open_native_ide_dialog", "preview_fit_form_content", "read_navigation_surface", "change_navigation_surface", "list_macros", "project_collection_state", "open_project_help", "certificate_trust", "verify_vba_signature_file", "preview_parameter_rename", "preview_local_rename", "toolbar_controls", "procedure_run_status", "form_clipboard_state", "list_toolbars", "read_runtime_forms", "read_code_clipboard", "native_code_history_state", "form_run_status", "list_object_browser", "select_object_browser", "read_object_browser", "code_pane_layout", "editor_layout", "window_layout", "project_symbols", "navigate_code", "code_bookmark", "preview_form_layout", "preview_code_edit", "status", "read_user_file", "list_projects", "list_modules", "list_references", "list_reference_types", "list_type_members", "read_module", "debug_state", "debug_windows", "debug_dialog", "debug_item", "read_debug_options", "read_vbe_options", "compile_project", "open_debug_pane", "list_commands", "select_code", "select_code_range",
            "project_properties", "project_persistence_status", "project_signature_status", "read_project_signature_dialog", "list_signing_certificates", "component_properties", "component_property_value", "vbe_windows", "vbe_environment", "list_addins", "focus_vbe_window", "window_linkage", "code_panes", "open_object_browser", "list_procedures", "find_code", "inspect_code_file", "select_procedure", "list_forms",
            "git_status", "git_history", "git_branches", "git_checkpoints", "git_conflicts", "git_conflict_read", "git_commit_read", "git_pull_requests",
            "form_state", "form_tree", "form_list_items", "form_properties", "form_control_properties", "form_event_catalog",
            "list_form_control_types", "open_form"
        };
        /// <summary>Obtient ou définit le nom du fournisseur courant utilisé pour les décisions d’accès.</summary>
        /// <value>Nom du fournisseur actif.</value>
        public string CurrentProviderName { get; set; }
        /// <summary>Obtient ou définit le mode de conversation qui autorise ou bloque les opérations d’écriture.</summary>
        /// <value>Mode courant, Agent par défaut.</value>
        public ChatMode Mode { get; set; } = ChatMode.Agent;
        /// <summary>Obtient ou définit le contrôle de portée appelé avant un outil.</summary>
        /// <value>Action de validation facultative.</value>
        public Action ValidateScope { get; set; }
        /// <summary>Obtient ou définit l’identifiant du projet auquel les outils sont limités.</summary>
        /// <value>Nom ou chemin du projet lié à la conversation.</value>
        public string BoundProject { get; set; }
        /// <summary>Indique qu’une restauration interne est en cours et peut contourner certaines gardes d’édition.</summary>
        private bool restoring;
        /// <summary>Adaptateurs natifs du débogueur, remplaçables par instance à la frontière UI.</summary>
        internal readonly VbeToolNativeBoundary Native = new VbeToolNativeBoundary();
        /// <summary>Exécute une commande sur la session hôte, sans remplacer l’orchestration de l’outil.</summary>
        internal Func<Request, Response> Execute;
        /// <summary>Interroge la disponibilité native d’une récupération dans le concepteur VBE.</summary>
        internal Func<Request, bool> CanRecoverDesignerCut;
        /// <summary>Demande la sauvegarde de signature au document hôte.</summary>
        internal Func<string, object> PersistSignature;
        /// <summary>Écrit les erreurs de lecture de diff dans le journal de chargement.</summary>
        internal Action<string> WriteLog = LoadLog.Write;
        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<VbeApprovalDialog, IWin32Window, DialogResult> ShowApproval = (dialog, window) => dialog.ShowDialog(window);
        /// <summary>Frontière native injectable, initialisée avec le comportement de production.</summary>
        internal Func<IWin32Window, string, string, DialogResult> ConfirmFile = (window, text, title) =>
            MessageBox.Show(window, text, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

        /// <summary>Valide la portée et refuse les outils d’écriture hors du mode Agent.</summary>
        /// <param name="name">Nom de l’outil demandé.</param>
        /// <exception cref="InvalidOperationException">Le mode courant interdit l’outil d’écriture.</exception>
        private void GuardMode(string name)
        {
            ValidateScope?.Invoke();
            if (!restoring && Mode != ChatMode.Agent && !ReadOnlyTools.Contains(name))
                throw new InvalidOperationException(UiText.Get("Mode ") + Mode + UiText.Get(" does not allow this editing or execution tool: ") + name);
        }

        /// <summary>Refuse qu’un outil cible un projet différent de celui lié à la conversation.</summary>
        /// <param name="name">Nom de l’outil demandé.</param>
        /// <param name="arguments">Arguments JSON contenant éventuellement le champ Project.</param>
        /// <exception cref="InvalidOperationException">Le projet demandé diffère de la portée courante.</exception>
        private void GuardProject(string name, string arguments)
        {
            if (string.IsNullOrEmpty(BoundProject) || (ReadOnlyTools.Contains(name) && name != "compile_project")) return;
            var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
            object project;
            if (values != null && values.TryGetValue("Project", out project) && !string.Equals(Convert.ToString(project), BoundProject, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Cette action vise un autre projet que celui de la conversation.");
        }

        /// <summary>Crée le catalogue d’outils lié à une session VBE et à sa fenêtre propriétaire.</summary>
        /// <param name="session">Session de commandes VBE.</param>
        /// <param name="owner">Fenêtre propriétaire des dialogues.</param>
        /// <param name="settings">Configuration des règles et fournisseurs.</param>
        /// <exception cref="ArgumentNullException">La configuration est nulle.</exception>
        public LlmVbeTools(VbeSession session, IWin32Window owner, LlmSettings settings)
        {
            this.session = session;
            Execute = request => session.Execute(request);
            CanRecoverDesignerCut = request => session.CanRecoverFormCut(request);
            PersistSignature = project => session.PersistProjectSignature(project);
            this.owner = owner;
            this.settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>Construit une définition de fonction JSON avec les champs et propriétés connus.</summary>
        /// <param name="name">Nom de l’outil.</param>
        /// <param name="description">Description affichée au fournisseur.</param>
        /// <param name="required">Noms des champs obligatoires.</param>
        /// <param name="fields">Noms des champs autorisés.</param>
        /// <returns>Objet conforme à la définition d’outil du fournisseur.</returns>
        private static object Definition(string name, string description, string[] required, params string[] fields)
        {
            var properties = new Dictionary<string, object>();
            foreach (string field in fields)
                properties[field] = field == "Temporary" ? (object)new { type = "boolean" } : field == "Value" ? (object)new { anyOf = new object[] {
                    new { type = "string" }, new { type = "number" }, new { type = "boolean" } } } :
                    field == "Arguments" && name == "run_procedure_values" ? ProcedureValuesArgumentSchema() :
                    field == "Arguments" ? (object)new { type = "array", maxItems = 30, items = new { anyOf = new object[] { new { type = "string" }, new { type = "number" }, new { type = "boolean" }, new { type = "null" } } } } :
                    field == "PathSegments" ? (object)new { type = "array", items = new { type = "string" }, minItems = 1, maxItems = 16 } :
                    field == "Rows" ? (object)new { type = "array", items = new { type = "array", items = new { type = "string", maxLength = 256 }, minItems = 1, maxItems = 10 }, minItems = 0, maxItems = 64 } :
                    field == "ArgumentNames" ? (object)new { type = "array", items = new { type = "string", maxLength = 255 }, minItems = 0, maxItems = 30 } :
                    field == "Items" ? (object)new { type = "array", items = new { type = "string", maxLength = 256 }, minItems = 0, maxItems = 64 } :
                    new { type = field == "StartLine" || field == "StartColumn" || field == "EndLine" || field == "EndColumn" || field == "Count" || field == "ExpectedMode" || field == "ControlId" || field == "WindowType" || field == "TargetWindowType" || field == "ProcKind" || field == "InsertIndex" ||
                        field == "ToolbarLeft" || field == "ToolbarTop" || field == "Offset" || field == "Limit" || field == "RowIndex" || field == "TypeIndex" || field == "ZPosition" ||
                        field == "Major" || field == "Minor" ? "integer" :
                    field == "Left" || field == "Top" || field == "Width" || field == "Height" || field == "FontSize" ? "number" :
                    field == "FontBold" || field == "WholeWord" || field == "MatchCase" || field == "PatternSearch" || field == "IncludeCallStack" ? "boolean" : "string" };
            return new { type = "function", function = new {
                name, description,
                parameters = new { type = "object", properties, required, additionalProperties = false }
            } };
        }

        /// <summary>Obtient l’ensemble des définitions d’outils proposées aux fournisseurs compatibles.</summary>
        /// <value>Définitions JSON des opérations autorisées.</value>
        public static object[] Definitions { get { return new object[] {
            Definition("status", "Read the live host process and currently open VBA projects; call before acting on VBE.", new string[0]),
            Definition("read_user_file", "Request separate user approval before reading and transmitting up to 64 KiB of a text file at a path explicitly supplied by the user.",
                new[] { "Path" }, "Path"),
            Definition("list_projects", "List open VBA projects, FileName and modes. Use the absolute FileName as Project selector when several projects have the same Name.", new string[0]),
            Definition("list_modules", "List modules in one VBA project.", new[] { "Project" }, "Project"),
            Definition("vbe_windows", "Read the native VBIDE Windows collection and the active window, including window type, visibility, state and position. Collection indexes are transient; no window is activated.", new string[0]),
            Definition("code_panes", "Read the already open VBIDE CodePanes collection and active code pane, with project/module, view, visible range and selection. Does not create or activate a pane.", new string[0]),
            Definition("debug_windows", "Read visible native VBE Locals, Watches and Immediate windows via accessibility. Optional IncludeCallStack opens the native Call Stack dialog through the Locals button, reads its frames, then closes it. Missing windows are reported as unavailable, not empty. No shortcuts or coordinate clicks are used.",
                new string[0], "IncludeCallStack"),
            Definition("debug_dialog", "Read a visible native VBA diagnostic dialog, including its exact message and button labels. Works while the VBE UI thread is modal; does not dismiss the dialog.", new string[0]),
            Definition("debug_item", "Expand or collapse exactly one row in a visible Locals or Watches pane by its PathSegments from debug_windows. Pane is locals or watches; Action is expand or collapse. Optional Context disambiguates watches with the same expression. Returns observed direct child count, then re-read debug_windows. Uses UI Automation, no shortcuts or coordinates.",
                new[] { "Pane", "Action", "PathSegments" }, "Pane", "Action", "PathSegments", "Context"),
            Definition("immediate_execute", "Execute one line in the visible VBE Immediate window through native character and Enter messages, without shortcuts or coordinates. Requires Project and current ExpectedMode (1 break or 2 design). Returns exact text before/after; arbitrary side effects require separate verification. Automatic VBE edit policy is required.",
                new[] { "Project", "ExpectedMode", "Text" }, "Project", "ExpectedMode", "Text"),
            Definition("respond_debug_dialog", "Activate one button on a visible native VBA run-time or compile diagnostic. Supply the exact Diagnostic and Button strings returned by debug_dialog, then read debug_state separately. Requires automatic VBE edit policy; no shortcut or coordinate click is used.",
                new[] { "Diagnostic", "Button" }, "Diagnostic", "Button"),
            Definition("open_debug_pane", "Open the native Locals, Watches or Immediate pane. Action is locals, watches or immediate. The effect may be asynchronous; verify with debug_windows in a separate request.",
                new[] { "Action" }, "Action"),
            Definition("debug_state", "Read design/run/break mode and the active code location for one project. Mode 1 is break; mode 2 is design. A modeless UserForm can remain visible while mode is 2; this state does not prove that every runtime form is closed.",
                new[] { "Project" }, "Project"),
            Definition("read_debug_options", "Read the VBE-wide error trapping setting from Tools > Options > General through the native dialog, then close with Cancel. No preference is changed. Returns the exact selected radio label and available choices; no shortcuts or coordinates.",
                new string[0]),
            Definition("certificate_trust", "Inspect the exact CertificateThumbprint in CurrentUser/My using Windows code-signing chain and Authenticode policy, cached URLs and cached revocation only. No downloads, private-key access or trust-store changes. Missing cached revocation produces an indeterminate result, not trust. This does not verify any VBA signature, digest, signer binding or timestamp.", new[] { "CertificateThumbprint" }, "CertificateThumbprint"),
            Definition("set_vbe_option", "Set one recognized native Editor or General preference from read_vbe_options using its exact Pane tab, Property control label and ExpectedOptionsVersion. Boolean Value for supported checkboxes, true for error-trapping radios, integer 1-32 for tab width. Native French/English labels are currently supported. Recognized Editor/General/Docking checkboxes and Editor Format controls are supported. Font/size/code colors require one exact observed Choices entry. NativeIndex:n identifies an unlabeled native palette choice, never an inferred RGB value. Query optionally selects one exact Code Colors category for the same palette mutation; read_vbe_options versions all categories together. An empty size catalogue refuses writes. Grid width/height are 2-60. Theme, security and arbitrary controls are excluded. InvokeAsync is required; normal VBE edit approval applies. Reopen read_vbe_options to verify persistence.", new[] { "Pane", "Property", "Value", "ExpectedOptionsVersion" }, "Pane", "Property", "Value", "ExpectedOptionsVersion", "Query"),
            Definition("read_vbe_options", "Read visible controls and values on every tab of the native VBE Tools > Options dialog, then close with Cancel. Returns native labels and read errors; no preference is changed, no shortcut or coordinates are used. This is a UI observation, not proof of persistence or of unavailable controls.",
                new string[0]),
            Definition("compile_project", "Compile the named VBA project using the native VBE command in design mode. Captures and dismisses a native compile error dialog; on failure read debug_state to locate the selected token. A successful response means no native diagnostic was observed. ExpectedMode must be 2.",
                new[] { "Project", "ExpectedMode" }, "Project", "ExpectedMode"),
            Definition("run_procedure", "Schedule one public Sub or Function in a standard module with up to 30 scalar JSON Arguments (string, finite number, boolean or null for VBA Null). Optional ArgumentNames pairs every value with a distinct parameter in the inspected signature, permitting omission of optional parameters and reordered named arguments; conditional/ParamArray named calls are refused. Requires current SHA and ExpectedMode=2. Uses a fully qualified native Immediate call, without generating source code. Functions print their result to Immediate. Poll procedure_run_status using Project and returned Query; delivery is not proof of runtime success. Do not retry automatically.",
                new[] { "Project", "Module", "Procedure", "ExpectedSha256", "ExpectedMode", "Arguments" },
                "Project", "Module", "Procedure", "ExpectedSha256", "ExpectedMode", "Arguments", "ArgumentNames"),
            Definition("procedure_run_status", "Read a previously scheduled procedure call by exact Project and Query operation identifier. Includes native Immediate output or delivery error; inspect debug_dialog and debug_state separately for runtime diagnostics.",
                new[] { "Project", "Query" }, "Project", "Query"),
            Definition("run_sub", "Run one parameterless Sub in a standard module by exact project, module and procedure name through the native VBE Run command. Requires the current module SHA-256 and ExpectedMode=2. Selects its declaration in the code pane; read debug_state separately for asynchronous effects. VBE edit policy applies.",
                new[] { "Project", "Module", "Procedure", "ExpectedSha256", "ExpectedMode" },
                "Project", "Module", "Procedure", "ExpectedSha256", "ExpectedMode"),
            Definition("list_commands", "List one page of VBE CommandBars controls matching an optional caption/path Query. Offset is zero-based and Limit defaults to 200 (maximum 200); request following pages until a page is short or empty. Returns transient Id, caption and enabled state; use these exact values for invoke_debug. The menu may change between requests.",
                new string[0], "Query", "Offset", "Limit"),
            Definition("select_code", "Activate a code pane and select an exact line or single-line text range after checking the current module SHA-256. Optional StartColumn and EndColumn are one-based, with an exclusive end; Expression can assert the selected source text. Does not edit source code.",
                new[] { "Project", "Module", "ExpectedSha256", "StartLine" },
                "Project", "Module", "ExpectedSha256", "StartLine", "StartColumn", "EndColumn", "Expression"),
            Definition("select_code_range", "Select an exact multi-line code range in a native CodePane after checking the current module SHA-256. StartLine/EndLine and StartColumn/EndColumn are one-based; the end position is exclusive. Re-reads the native selection. Does not modify code.",
                new[] { "Project", "Module", "ExpectedSha256", "StartLine", "StartColumn", "EndLine", "EndColumn" },
                "Project", "Module", "ExpectedSha256", "StartLine", "StartColumn", "EndLine", "EndColumn"),
            Definition("quick_watch", "Evaluate exactly the selected single-line VBA expression in break mode through the native Quick Watch dialog. Requires current module SHA-256, one-based selection columns and exact Expression; optional Procedure asserts context. The expression can call VBA code and have side effects. The dialog value is read and closed without shortcuts or coordinates; Automatic VBE edit policy is required.",
                new[] { "Project", "Module", "ExpectedSha256", "ExpectedMode", "StartLine", "StartColumn", "EndColumn", "Expression" },
                "Project", "Module", "ExpectedSha256", "ExpectedMode", "StartLine", "StartColumn", "EndColumn", "Expression", "Procedure"),
            Definition("invoke_debug", "Execute a native VBE debugger command on an exact project/module/line after checking SHA-256, expected mode, command Id and caption. Action is toggle_breakpoint, run, continue, step_into, step_over, step_out, run_to_cursor or set_next_statement. The last four require break mode. Before set_next_statement, use debug_global show_next_statement and read debug_state in a separate request; the target must be in that selected procedure, while the VBE still enforces the actual execution procedure. Setting the next statement changes control flow; verify by subsequent execution. VBE may apply the effect after return: read debug_state and debug_windows separately. Breakpoint toggle is not yet independently verifiable.",
                new[] { "Project", "Module", "ExpectedSha256", "StartLine", "ExpectedMode", "Action", "ControlId", "ControlCaption" },
                "Project", "Module", "ExpectedSha256", "StartLine", "ExpectedMode", "Action", "ControlId", "ControlCaption"),
            Definition("add_watch", "Add a native VBE watch in the active break-mode project and module, using the current procedure context. ExpectedMode must be 1. Optional Procedure must match the native dialog context. WatchType is expression (default), break_when_true or break_when_changed. The expression is evaluated by VBE and can call VBA functions. The dialog is completed through native controls without shortcuts or coordinates; then call debug_windows to re-read its value.",
                new[] { "Project", "Module", "ExpectedMode", "Expression" },
                "Project", "Module", "ExpectedMode", "Expression", "Procedure", "WatchType"),
            Definition("edit_watch", "Edit one native VBE watch selected by exact Expression and Context from debug_windows. NewExpression replaces its expression in the same context; optional WatchType changes its break condition. The new expression is read back but a break condition needs runtime verification. Requires automatic VBE edit policy, no shortcuts or coordinates.",
                new[] { "Project", "ExpectedMode", "Expression", "Context", "NewExpression" },
                "Project", "ExpectedMode", "Expression", "Context", "NewExpression", "WatchType"),
            Definition("remove_watch", "Remove one native VBE watch selected by exact Expression and Context from debug_windows. Requires the Watches pane visible and ExpectedMode from debug_state. Confirms disappearance separately through UI accessibility.",
                new[] { "Project", "ExpectedMode", "Expression", "Context" },
                "Project", "ExpectedMode", "Expression", "Context"),
            Definition("debug_global", "Execute break in run mode, reset or show_next_statement in break mode, or clear_all_breakpoints in break/design mode, through native VBE commands. Break and Clear All Breakpoints affect the entire VBE. Clear All Breakpoints cannot be verified from a VBIDE inventory; check subsequent execution on a disposable procedure. Requires ExpectedMode from debug_state and VBE edit policy.",
                new[] { "Project", "ExpectedMode", "Action" },
                "Project", "ExpectedMode", "Action"),
            Definition("open_object_browser", "Open the native VBE Object Browser through CommandBars Id 473 and read vbe_windows immediately. Opening may be asynchronous: if VerificationPending is true, call vbe_windows again in a separate request and confirm a visible Type 2 window. This command does not read libraries, classes or members.", new string[0]),
            Definition("vbe_environment", "Read the VBE version, active project and counts of projects, windows, code panes and VBE add-ins. Per-field COM failures are reported in Errors.", new string[0]),
            Definition("list_addins", "List VBE-registered add-ins with ProgId, Guid, Description and current Connect state. This is the VBE Add-In Manager collection, not the host application's COM add-ins. Per-field COM failures are reported; no add-in is loaded or unloaded.", new string[0]),
            Definition("focus_vbe_window", "Focus exactly one already-visible VBE window by the exact WindowCaption and WindowType returned by vbe_windows. Refuses absent, hidden or ambiguous windows and reads ActiveWindow after SetFocus. No shortcut or coordinate is used.",
                new[] { "WindowCaption", "WindowType" }, "WindowCaption", "WindowType"),
            Definition("show_vbe_window", "Show and focus exactly one VBE window already present in vbe_windows, including a hidden permanent View pane. Requires its exact WindowCaption and WindowType, then verifies Visible and ActiveWindow. Does not create a code pane or designer; no shortcut or coordinate is used.",
                new[] { "WindowCaption", "WindowType" }, "WindowCaption", "WindowType"),
            Definition("window_linkage", "Inspect whether one exact native VBE window from vbe_windows has a LinkedWindowFrame, and list that frame's LinkedWindows. Unsupported COM getters are reported in Errors; no window is moved or focused.",
                new[] { "WindowCaption", "WindowType" }, "WindowCaption", "WindowType"),
            Definition("close_vbe_window", "Close exactly one visible native VBE window by WindowCaption and WindowType from vbe_windows. A code pane or designer is destroyed as a window, while permanent View windows are hidden; VBA code and components are not deleted. Refuses the CodexVBE tool window and ambiguous targets, then checks vbe_windows. Subject to VBE edit policy; no shortcut or coordinate is used.",
                new[] { "WindowCaption", "WindowType" }, "WindowCaption", "WindowType"),
            Definition("list_procedures", "List Sub, Function and Property Get/Let/Set procedures from CodeModule without opening a code pane; returns exact VBIDE line ranges and module SHA-256.",
                new[] { "Project", "Module" }, "Project", "Module"),
            Definition("find_code", "Search one module or all modules of a project from CodeModule.Lines without opening a pane. Returns up to 200 locations and source SHA-256 values. Optional PatternSearch treats * as any number of characters and ? as one character within each line; MatchCase and WholeWord also apply. A pattern of only asterisks is refused.",
                new[] { "Project", "Query" }, "Project", "Module", "Query", "WholeWord", "MatchCase", "PatternSearch"),
            Definition("select_procedure", "Navigate the VBE to a procedure declaration using an exact project/module/name/ProcKind and a current ExpectedSha256. ProcKind: 0 Sub or Function, 1 Property Let, 2 Property Set, 3 Property Get. Changes only UI selection, not code.",
                new[] { "Project", "Module", "Procedure", "ProcKind", "ExpectedSha256" },
                "Project", "Module", "Procedure", "ProcKind", "ExpectedSha256"),
            Definition("create_event_procedure", "Create a UserForm or control event stub through CodeModule.CreateEventProc. ObjectName is UserForm or one control name from form_tree. Requires current code ExpectedSha256 and ExpectedTreeVersion, design mode and VBE edit policy; duplicates and invalid events are refused, then code is re-read.",
                new[] { "Project", "Form", "ObjectName", "EventName", "ExpectedSha256", "ExpectedTreeVersion" },
                "Project", "Form", "ObjectName", "EventName", "ExpectedSha256", "ExpectedTreeVersion"),
            Definition("create_procedure", "Append one complete Sub, Function or Property Let/Set/Get to a standard or class module in design mode. Text contains the full declaration, body and matching End statement. Procedure and ProcKind (0=Sub/Function, 1=Let, 2=Set, 3=Get) must match. Requires current module ExpectedSha256; VBIDE recognition and resulting SHA are read back. Compilation is not implied.",
                new[] { "Project", "Module", "Procedure", "ProcKind", "Text", "ExpectedSha256" },
                "Project", "Module", "Procedure", "ProcKind", "Text", "ExpectedSha256"),
            Definition("replace_procedure", "Replace only an existing Sub, Function or Property Get/Let/Set declaration through its End statement in a standard or class module. Comments and blank lines before or after remain in place. Text is the full replacement procedure with the same name and ProcKind. Requires the current module ExpectedSha256; VBIDE identity and new SHA are read back. Compilation is separate.",
                new[] { "Project", "Module", "Procedure", "ProcKind", "Text", "ExpectedSha256" },
                "Project", "Module", "Procedure", "ProcKind", "Text", "ExpectedSha256"),
            Definition("remove_procedure", "Remove only one existing Sub, Function or Property Get/Let/Set declaration through its End statement in a standard or class module. Preceding comments and other procedures remain. Requires the current module ExpectedSha256, design mode and VBE edit policy; VBIDE absence and the resulting SHA are read back. Compilation is separate.",
                new[] { "Project", "Module", "Procedure", "ProcKind", "ExpectedSha256" },
                "Project", "Module", "Procedure", "ProcKind", "ExpectedSha256"),
            Definition("inspect_code_file", "Inspect bytes of an absolute local source file explicitly named by the user without transmitting its content. Returns BOM, strict UTF-8 validity, non-ASCII/NUL flags, SHA-256, host ANSI code page and whether SourceEncoding must be explicit. Use before insert_code_file when encoding is uncertain.",
                new[] { "Path" }, "Path"),
            Definition("insert_code_file", "Insert up to 256 KiB of local VBA source text at StartLine in a design-mode module. Path must be an absolute path explicitly supplied by the user; the file is read locally and its bytes are not included in tool arguments. BOM or ASCII is auto-detected; a non-ASCII file without BOM requires explicit SourceEncoding from inspection or known provenance. Supported values: utf-8, utf-16le, utf-16be, windows-1252, system-ansi. Non-ASCII characters are checked in VBE after insertion, with rollback on mismatch. Requires current ExpectedSha256 and VBE edit policy. Returns source file SHA-256, encoding, inserted line count and new module SHA-256; compilation is separate.",
                new[] { "Project", "Module", "Path", "StartLine", "ExpectedSha256" },
                "Project", "Module", "Path", "StartLine", "ExpectedSha256", "SourceEncoding"),
            Definition("project_properties", "Read all exposed VBProject properties, component identities and a project revision.",
                new[] { "Project" }, "Project"),
            Definition("project_persistence_status", "Read VBProject.Saved and the exact Excel workbook, Word document, PowerPoint presentation or native standalone SWP project (Type=101) state. Documents are matched by native project identity and host PID. Unsupported hosts remain unavailable. Word/PowerPoint and SWP runtime qualification is pending. This does not write to disk.",
                new[] { "Project" }, "Project"),
            Definition("save_host_document", "Save the already-named writable Excel workbook, Word macro document, PowerPoint macro presentation or native standalone SWP project (Type=101) owning the exact design-mode VBE project. Requires ExpectedProjectVersion and ExpectedHostPath from project_persistence_status. Word/PowerPoint guard native identity, host PID and unchanged VBA after saving. Unsupported host projects and unsaved paths are refused. Word/PowerPoint and SWP runtime/reload qualification is pending; never treat saved flags as reload proof.",
                new[] { "Project", "ExpectedProjectVersion", "ExpectedHostPath" },
                "Project", "ExpectedProjectVersion", "ExpectedHostPath"),
            Definition("save_host_document_as", "First-save an unsaved Excel VBA project as .xlsm, Word as .docm/.dotm, PowerPoint as .pptm/.potm/.ppsm, or a native standalone project (Type=101) as .swp to a new Path explicitly supplied by the user. Refuses overwrite and checks ExpectedProjectVersion, design mode, native identity and saved paths. Unsupported host projects are refused. Word/PowerPoint and SWP runtime/reload remains unqualified; reopen the file to prove persistence. VBE edit policy applies.",
                new[] { "Project", "ExpectedProjectVersion", "Path" },
                "Project", "ExpectedProjectVersion", "Path"),
            Definition("project_signature_status", "Read whether the exact Excel workbook owning this VBE project has a signed VBA project. Returns Available=false when the host is not Excel, the registered Excel instance differs from this VBE, or its project cannot be matched. This does not sign, validate the certificate, or inspect pending edits.",
                new[] { "Project" }, "Project"),
            Definition("list_signing_certificates", "List public metadata and thumbprints of CurrentUser/My certificates with a private key and Code Signing EKU. EligibleNow indicates the validity dates only; it does not establish certificate trust. Use an exact thumbprint with sign_project.",
                new string[0]),
            Definition("read_project_signature_dialog", "Read the labels and buttons of the native VBE Digital Signature dialog for the currently active project, then close it through its accessible Cancel action. Project must exactly match ActiveVBProject and ExpectedMode must be 2. This is a read-only observation of the displayed certificate names; it does not assign, remove or validate a certificate. No shortcuts or coordinate clicks.",
                new[] { "Project", "ExpectedMode" }, "Project", "ExpectedMode"),
            Definition("sign_project", "Sign an unsigned, saved VBE project with an explicit code-signing certificate thumbprint from CurrentUser/My. Requires current project version, design mode and VBE edit policy. If the certificate is already associated with the project, the native VBE dialog can reuse it directly. Otherwise Windows Security asks the user to confirm it; the add-in waits up to two minutes, verifies the selected name and confirms the VBE dialog. For Excel, the add-in saves the exact workbook and reports SaveRequired=false on success; other hosts need their native save. Host signature status and a dialog readback are not cryptographic trust validation.",
                new[] { "Project", "ExpectedProjectVersion", "ExpectedMode", "CertificateThumbprint" },
                "Project", "ExpectedProjectVersion", "ExpectedMode", "CertificateThumbprint"),
            Definition("component_properties", "Read all exposed VBComponent and designer properties, code SHA-256, and a component revision. Works for document, standard, class and form components when VBIDE allows access.",
                new[] { "Project", "Module" }, "Project", "Module"),
            Definition("component_property_value", "Read one named VBComponent host property on demand. Type 100 document properties belong to the host object, not the common VBE editor; Excel MailEnvelope returns an explicit error because its getter blocks COM inspection.",
                new[] { "Project", "Module", "Property" }, "Project", "Module", "Property"),
            Definition("set_project_property", "Set a writable scalar VBProject property using ExpectedProjectVersion in design mode. Name uses a verified typed setter only for an unprotected, saved, writable Excel workbook targeted by its exact rooted path; source text is preserved, qualified references are not renamed. Other hosts and protected projects refuse Name. VBE edit policy applies.",
                new[] { "Project", "ExpectedProjectVersion", "Property", "Value" },
                "Project", "ExpectedProjectVersion", "Property", "Value"),
            Definition("set_component_property", "Set a writable scalar VBComponent property by its live descriptor and ExpectedComponentVersion in design mode; VBE edit policy applies.",
                new[] { "Project", "Module", "ExpectedComponentVersion", "Property", "Value" },
                "Project", "Module", "ExpectedComponentVersion", "Property", "Value"),
            Definition("set_class_instancing", "Set a VBA class module's Instancing property to 1 (Private) or 2 (PublicNotCreatable). Requires the current component revision, design mode and VBE edit policy; read back the resulting class properties. Excel and other hosts may restrict this setting.",
                new[] { "Project", "Module", "ExpectedComponentVersion", "Value" },
                "Project", "Module", "ExpectedComponentVersion", "Value"),
            Definition("rename_component", "Rename a VBComponent after checking ExpectedComponentVersion and VBE edit policy.",
                new[] { "Project", "Module", "ExpectedComponentVersion", "NewName" },
                "Project", "Module", "ExpectedComponentVersion", "NewName"),
            Definition("remove_component", "Remove an editable non-document component after checking project and component revisions; VBE edit policy applies. Removal cannot be undone in VBE.",
                new[] { "Project", "Module", "ExpectedProjectVersion", "ExpectedComponentVersion" },
                "Project", "Module", "ExpectedProjectVersion", "ExpectedComponentVersion"),
            Definition("import_component", "Import a VBA component from an absolute file path explicitly supplied by the user, with ExpectedProjectVersion and VBE edit policy. The file stays local. Applied=true with VerificationPending=true means the import took effect but needs a separate component_properties read; never retry Import automatically.",
                new[] { "Project", "ExpectedProjectVersion", "Path" },
                "Project", "ExpectedProjectVersion", "Path"),
            Definition("export_component", "Export a VBA component to a new absolute path explicitly supplied by the user, with ExpectedComponentVersion and VBE edit policy. Existing files are not overwritten; a UserForm may also create an FRX companion.",
                new[] { "Project", "Module", "ExpectedComponentVersion", "Path" },
                "Project", "Module", "ExpectedComponentVersion", "Path"),
            Definition("list_references", "List the type-library references actually selected by one VBA project, including identity, version, path and broken status.",
                new[] { "Project" }, "Project"),
            Definition("list_reference_types", "Read one page of up to 50 COM types from a reference selected by the VBA project. Use the exact Guid, Major and Minor returned by list_references. Offset is zero-based; Limit defaults to 50. Returns TypeIndex and TypeIdentity for list_type_members, and records whether the type library came from the reference file or registry. VBAProject itself is outside this reference catalog.",
                new[] { "Project", "Guid", "Major", "Minor" }, "Project", "Guid", "Major", "Minor", "Offset", "Limit"),
            Definition("list_type_members", "Read one page of up to 50 raw COM functions and variables for a type from list_reference_types. Supply that type's TypeIndex and TypeIdentity with the exact selected reference. A coclass resolves its default non-source interface. COM accessors and hidden members are not filtered like the VBE Object Browser; this is read-only metadata, not a live object property read.",
                new[] { "Project", "Guid", "Major", "Minor", "TypeIndex", "TypeIdentity" },
                "Project", "Guid", "Major", "Minor", "TypeIndex", "TypeIdentity", "Offset", "Limit"),
            Definition("add_reference_guid", "Add a project type-library reference by exact GUID and requested major/minor version in design mode; requires the current references revision and VBE edit policy. Major=Minor=0 requests the latest installed version.",
                new[] { "Project", "ExpectedReferencesVersion", "Guid", "Major", "Minor" },
                "Project", "ExpectedReferencesVersion", "Guid", "Major", "Minor"),
            Definition("add_reference_file", "Add a project reference from a fully qualified local type-library path explicitly supplied by the user; requires the current references revision and VBE edit policy. The file remains local.",
                new[] { "Project", "ExpectedReferencesVersion", "Path" },
                "Project", "ExpectedReferencesVersion", "Path"),
            Definition("remove_reference", "Remove one exactly identified project reference by GUID and major/minor after checking the current references revision; subject to VBE edit policy.",
                new[] { "Project", "ExpectedReferencesVersion", "Guid", "Major", "Minor" },
                "Project", "ExpectedReferencesVersion", "Guid", "Major", "Minor"),
            Definition("read_module", "Read complete VBA code and its SHA-256 revision.", new[] { "Project", "Module" }, "Project", "Module"),
            Definition("create_module", "Create a named standard VBA module in the selected design-mode project. ExpectedMode must be 2 from list_projects.",
                new[] { "Project", "Module", "ExpectedMode" }, "Project", "Module", "ExpectedMode"),
            Definition("create_class", "Create a named VBA class module in the selected design-mode project. ExpectedMode must be 2 from list_projects.",
                new[] { "Project", "Module", "ExpectedMode" }, "Project", "Module", "ExpectedMode"),
            Definition("list_forms", "List UserForms in a project.", new[] { "Project" }, "Project"),
            Definition("list_form_control_types", "List native MSForms controls and installed x64 CATID_Control candidates. Extra ActiveX hosting is unverified until Controls.Add succeeds.", new string[0]),
            Definition("form_state", "Read a UserForm and all its controls with geometry, caption and font.", new[] { "Project", "Form" }, "Project", "Form"),
            Definition("form_tree", "Read the recursive UserForm hierarchy, including Frame controls, MultiPage pages and TabStrip tabs. FormVersion and TreeVersion are the same recursive revision; use it for ExpectedFormVersion or ExpectedTreeVersion.",
                new[] { "Project", "Form" }, "Project", "Form"),
            Definition("form_list_items", "Read a bounded page of live indexed items from a design-time MSForms ComboBox or ListBox selected by canonical form_tree ControlPath. Offset is a zero-based row index; Limit defaults to 20 and is capped so at most 128 cells are read. Returns item values and per-cell errors. The list contents are not saved with the UserForm in the tested Excel VBE; this is a live-state read, not a persistence guarantee.",
                new[] { "Project", "Form", "ControlPath" }, "Project", "Form", "ControlPath", "Offset", "Limit"),
            Definition("set_form_list_initializer", "Generate or replace only the marked list block in UserForm_Initialize for an unbound native ComboBox or ListBox, including canonical paths nested in Frames and MultiPage Pages. Supply exactly one of Items (one column) or Rows (rectangular string matrix matching the existing ColumnCount, 1-10 columns), at most 64 rows and 256 characters per single-line cell. An empty array clears the list at runtime. Designer properties are not changed. Requires current form_tree TreeVersion and read_module SHA-256, design mode and VBE edit policy. Existing user code is preserved; edited managed blocks are refused. VBA code persists with the workbook, but this command does not populate the designer's live List. Verify the returned code with read_module; runtime execution remains a separate check.",
                new[] { "Project", "Form", "ControlPath", "ExpectedTreeVersion", "ExpectedSha256" },
                "Project", "Form", "ControlPath", "Items", "Rows", "ExpectedTreeVersion", "ExpectedSha256"),
            Definition("form_event_catalog", "Read the COM source-interface event names for a UserForm or a control selected by canonical form_tree ControlPath. This is only the COM source catalog: VBA/VBE events such as UserForm.Initialize can be absent. SourceInterfacesComplete does not mean all usable VBE events are listed; CreateEventProc validates a requested event name.",
                new[] { "Project", "Form" }, "Project", "Form", "ControlPath"),
            Definition("form_properties", "Read the designer properties of a UserForm.", new[] { "Project", "Form" }, "Project", "Form"),
            Definition("set_form_property", "Set a typed UserForm property or one object member path (for example Font.Name) after reading form_properties and form_state; uses ExpectedFormVersion and VBE edit policy. Read-only or unsupported objects return explicit errors.",
                new[] { "Project", "Form", "ExpectedFormVersion", "Property", "Value" },
                "Project", "Form", "ExpectedFormVersion", "Property", "Value"),
            Definition("set_form_picture", "Set a UserForm Picture from an absolute local image path explicitly supplied by the user. The image remains local and is not sent to the LLM. Requires form revision and VBE edit policy.",
                new[] { "Project", "Form", "ExpectedFormVersion", "Path" },
                "Project", "Form", "ExpectedFormVersion", "Path"),
            Definition("form_control_properties", "Read all exposed design properties, types and read-only flags of one UserForm control.",
                new[] { "Project", "Form", "Control" }, "Project", "Form", "Control"),
            Definition("set_form_node_property", "Set any writable scalar designer property, or a writable COM object member path, on a control, Frame, MultiPage Page or TabStrip Tab selected by its canonical form_tree path. Type is taken from the live property descriptor; requires ExpectedTreeVersion and VBE edit policy.",
                new[] { "Project", "Form", "ControlPath", "ExpectedTreeVersion", "Property", "Value" },
                "Project", "Form", "ControlPath", "ExpectedTreeVersion", "Property", "Value"),
            Definition("set_form_node_picture", "Set a writable OLE Picture or MouseIcon property on a form_tree node from a local image path explicitly supplied by the user. Image bytes remain local; requires ExpectedTreeVersion and VBE edit policy.",
                new[] { "Project", "Form", "ControlPath", "ExpectedTreeVersion", "Property", "Path" },
                "Project", "Form", "ControlPath", "ExpectedTreeVersion", "Property", "Path"),
            Definition("z_order_form_control", "Move a UserForm control to front (ZPosition=0) or back (ZPosition=1) in the native MSForms designer. Requires a canonical form_tree ControlPath, current ExpectedTreeVersion, design mode and VBE edit policy. Tested visually on two overlapping Labels. VBIDE does not expose a z-order readback; the tool reports Executed and Unverified, and TreeVersion may stay unchanged.",
                new[] { "Project", "Form", "ControlPath", "ExpectedTreeVersion", "ZPosition" },
                "Project", "Form", "ControlPath", "ExpectedTreeVersion", "ZPosition"),
            Definition("remove_form_control", "Remove a UserForm control at its canonical form_tree ControlPath, including one inside a Frame or MultiPage Page. Requires ExpectedTreeVersion and VBE edit policy; a Page or Tab itself is not accepted. The control and its descendants are deleted from the designer.",
                new[] { "Project", "Form", "ControlPath", "ExpectedTreeVersion" },
                "Project", "Form", "ControlPath", "ExpectedTreeVersion"),
            Definition("add_form_page", "Add a named Page to a MultiPage at canonical form_tree ParentPath. Optional InsertIndex is zero-based; without it the Page is appended. Requires ExpectedTreeVersion and VBE edit policy; re-reads the tree.",
                new[] { "Project", "Form", "ParentPath", "NewName", "ExpectedTreeVersion" },
                "Project", "Form", "ParentPath", "NewName", "Caption", "InsertIndex", "ExpectedTreeVersion"),
            Definition("add_form_tab", "Add a named Tab to a TabStrip at canonical form_tree ParentPath. Optional InsertIndex is zero-based; without it the Tab is appended. Requires ExpectedTreeVersion and VBE edit policy; re-reads the tree.",
                new[] { "Project", "Form", "ParentPath", "NewName", "ExpectedTreeVersion" },
                "Project", "Form", "ParentPath", "NewName", "Caption", "InsertIndex", "ExpectedTreeVersion"),
            Definition("remove_form_page_tab", "Remove a MultiPage Page or TabStrip Tab by its canonical form_tree ControlPath. A Page's child controls are deleted with it. Requires ExpectedTreeVersion and VBE edit policy; resolves the item's numeric index and re-reads the tree.",
                new[] { "Project", "Form", "ControlPath", "ExpectedTreeVersion" },
                "Project", "Form", "ControlPath", "ExpectedTreeVersion"),
            Definition("open_form", "Open a UserForm designer window in VBE.", new[] { "Project", "Form" }, "Project", "Form"),
            Definition("create_form", "Create a named UserForm in the selected design-mode project, subject to VBE edit policy.",
                new[] { "Project", "Form" }, "Project", "Form"),
            Definition("replace_lines", "Replace VBA lines only when ExpectedSha256 matches the current module; subject to VBE edit policy.",
                new[] { "Project", "Module", "ExpectedSha256", "StartLine", "Count", "Text" },
                "Project", "Module", "ExpectedSha256", "StartLine", "Count", "Text"),
            Definition("add_form_control", "Add a built-in MSForms control in design mode; requires form revision and VBE edit policy.",
                new[] { "Project", "Form", "ExpectedFormVersion", "ControlType", "Control", "Left", "Top", "Width", "Height" },
                "Project", "Form", "ExpectedFormVersion", "ControlType", "Control", "Left", "Top", "Width", "Height", "Caption"),
            Definition("add_nested_form_control", "Add a built-in MSForms control inside a Frame or MultiPage Page named by a canonical form_tree ParentPath. Requires ExpectedTreeVersion, design mode and VBE edit policy.",
                new[] { "Project", "Form", "ParentPath", "ExpectedTreeVersion", "ControlType", "Control", "Left", "Top", "Width", "Height" },
                "Project", "Form", "ParentPath", "ExpectedTreeVersion", "ControlType", "Control", "Left", "Top", "Width", "Height", "Caption"),
            Definition("rename_form_control", "Rename a UserForm control; requires form revision and VBE edit policy.",
                new[] { "Project", "Form", "ExpectedFormVersion", "Control", "NewName" },
                "Project", "Form", "ExpectedFormVersion", "Control", "NewName"),
            Definition("set_form_control_caption", "Set a UserForm control caption; requires form revision and VBE edit policy.",
                new[] { "Project", "Form", "ExpectedFormVersion", "Control", "Caption" },
                "Project", "Form", "ExpectedFormVersion", "Control", "Caption"),
            Definition("set_form_control_font", "Set a UserForm control font; requires form revision and VBE edit policy.",
                new[] { "Project", "Form", "ExpectedFormVersion", "Control", "FontName", "FontSize", "FontBold" },
                "Project", "Form", "ExpectedFormVersion", "Control", "FontName", "FontSize", "FontBold"),
            Definition("set_form_control_geometry", "Place and size a UserForm control; requires form revision and VBE edit policy.",
                new[] { "Project", "Form", "ExpectedFormVersion", "Control", "Left", "Top", "Width", "Height" },
                "Project", "Form", "ExpectedFormVersion", "Control", "Left", "Top", "Width", "Height")
        }.Concat(GitDefinitions).Concat(EditorDefinitions).ToArray(); } }

        /// <summary>Valide les gardes puis exécute synchroniquement un outil et sérialise sa réponse.</summary>
        /// <param name="name">Nom de l’outil demandé.</param>
        /// <param name="arguments">Arguments JSON de l’appel.</param>
        /// <returns>JSON d’une réponse réussie ou d’erreur.</returns>
        public string Invoke(string name, string arguments)
        {
            if (name.StartsWith("git_", StringComparison.Ordinal)) return json.Serialize(Response.Failure("Git tools require InvokeAsync."));
            try { GuardMode(name); GuardProject(name, arguments); }
            catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            var definition = Definitions.Cast<dynamic>().FirstOrDefault(item => (string)item.function.name == name);
            if (definition == null) return json.Serialize(Response.Failure("Unknown tool: " + name));
            try
            {
                var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                if (values == null) throw new ArgumentException("Tool arguments must be an object.");
                var required = (string[])definition.function.parameters.required;
                var fields = (Dictionary<string, object>)definition.function.parameters.properties;
                foreach (string field in required)
                    if (!values.ContainsKey(field) || values[field] == null ||
                        (field != "Text" && field != "Caption" && field != "Value" && values[field] is string &&
                            string.IsNullOrWhiteSpace((string)values[field])))
                        throw new ArgumentException(field + " is required.");
                foreach (string field in values.Keys)
                {
                    if (!fields.ContainsKey(field)) throw new ArgumentException("Unexpected argument: " + field);
                    object value = values[field];
                    if (field == "Value")
                    {
                        if (!(value is string) && !(value is bool) && !(value is int) &&
                            !(value is long) && !(value is double) && !(value is decimal))
                            throw new ArgumentException("Value must be a string, number or boolean.");
                        continue;
                    }
                    if (field == "Arguments")
                    {
                        var scalars = value as object[];
                        if (scalars == null || scalars.Length > 30) throw new ArgumentException("Arguments must contain at most 30 scalar values.");
                        if (name == "run_procedure_values") VbaProcedureValues.Capture(scalars);
                        else foreach (object scalar in scalars) VbeDebug.ProcedureLiteral(scalar);
                        continue;
                    }
                    if (field == "ArgumentNames")
                    {
                        var names = value as object[];
                        if (names == null || names.Length > 30 || names.Any(item => !(item is string) ||
                            ((string)item).Length > 255 || ((string)item).Any(char.IsControl)))
                            throw new ArgumentException("ArgumentNames must contain at most 30 single-line parameter names of at most 255 characters.");
                        continue;
                    }
                    if (field == "Items")
                    {
                        var items = value as object[];
                        if (items == null || items.Length > 64 || items.Any(item => !(item is string) ||
                            ((string)item).Length > 256 || ((string)item).Any(char.IsControl)))
                            throw new ArgumentException("Items must contain at most 64 single-line strings of at most 256 characters.");
                        continue;
                    }
                    if (field == "Rows")
                    {
                        var rows = value as object[];
                        if (rows == null || rows.Length > 64 || rows.Any(row =>
                        {
                            var cells = row as object[];
                            return cells == null || cells.Length < 1 || cells.Length > 10 ||
                                cells.Any(cell => !(cell is string) || ((string)cell).Length > 256 || ((string)cell).Any(char.IsControl));
                        }) || (rows.Length > 0 && rows.Any(row => ((object[])row).Length != ((object[])rows[0]).Length)))
                            throw new ArgumentException("Rows must be a rectangular matrix of at most 64 rows, 1-10 columns and single-line strings of at most 256 characters.");
                        continue;
                    }
                    string type = (string)((dynamic)fields[field]).type;
                    if (value == null ||
                        (type == "string" && !(value is string)) ||
                        (type == "boolean" && !(value is bool)) ||
                        (type == "integer" && !(value is int) && !(value is long)) ||
                        (type == "number" && !(value is int) && !(value is long) && !(value is double) && !(value is decimal)))
                        throw new ArgumentException(field + " must be a " + type + ".");
                }
                if (name == "read_user_file")
                    return json.Serialize(ReadUserFile((string)values["Path"]));
                if ((name == "set_form_picture" || name == "set_form_node_picture" ||
                    name == "add_reference_file" || name == "insert_code_file" || name == "import_component" || name == "open_standalone_project" || (name == "set_project_protection" && values.ContainsKey("Path")) ||
                    name == "save_host_document_as" ||
                    name == "inspect_code_file" || name == "export_component") &&
                    !IsExplicitUserPath((string)values["Path"]))
                    return json.Serialize(Response.Failure("L'utilisateur doit fournir explicitement le chemin absolu du fichier."));
                var normalized = new Dictionary<string, object>(values) { ["Command"] = name };
                var request = json.Deserialize<Request>(json.Serialize(normalized));
                // All newly registered tools are treated as edits unless explicitly classified as read-only.
                bool edit = !ReadOnlyTools.Contains(name);
                if (edit && settings.VbeEditApproval == "ReadOnly")
                    return json.Serialize(Response.Failure(UiText.Get("VBE edits are disabled (Read-only mode).")));
                if (edit && settings.VbeEditApproval != "Automatic" && settings.VbeEditApproval != "AskEachTime")
                    return json.Serialize(Response.Failure(UiText.Get("Unknown VBE edit policy; action refused.")));
                Dictionary<string, CodeSnapshot> procedureRenameBefore = name == "apply_procedure_rename" || name == "apply_class_member_rename" ? ReadProcedureRenameBefore(request) : null;
                CodeSnapshot beforeCode = null;
                bool formCodeEdit = name == "set_form_list_initializer" || name == "set_form_list_binding";
                string editedModule = formCodeEdit ? request.Form : request.Module;
                if (formCodeEdit || name == "cut_code" || name == "paste_code" || name == "replace_lines" || name == "apply_code_edit" || name == "apply_local_rename" || name == "apply_parameter_rename" || name == "undo_code_edit" || name == "redo_code_edit")
                {
                    beforeCode = ReadCode(request.Project, editedModule);
                    if (!string.Equals(beforeCode.Sha256, request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                        return json.Serialize(Response.Failure(UiText.Get("The module changed since the model read it.")));
                    if (name == "replace_lines") CodeChange.PreviewRows(beforeCode.Code, request);
                }
                if (edit && settings.VbeEditApproval == "AskEachTime" && name != "replace_lines")
                {
                    string summary = name + "\r\n\r\n" + json.Serialize(values);
                    using (var approval = new VbeApprovalDialog(summary))
                    {
                        if (ShowApproval(approval, owner) != DialogResult.Yes)
                            return json.Serialize(Response.Failure("User rejected the edit."));
                    }
                }
                Response result;
                try
                {
                    result = name == "status"
                        ? Response.Success(LlmVbeContext.LiveSnapshot(session))
                        : Execute(request);
                }
                catch (Exception error) when (name == "cut_code" || name == "paste_code" || name == "apply_procedure_rename" || name == "apply_class_member_rename")
                {
                    // An edit can fail after a partial native mutation. Still read back
                    // below so the user receives a diff and a recovery opportunity.
                    result = Response.Failure(error.Message);
                }
                if (procedureRenameBefore != null && !restoring) PublishProcedureRenameChanges(request.Project, procedureRenameBefore);
                if (result.Ok && name == "native_form_clipboard" && request.Action == "cut")
                {
                    dynamic cut = result.Data;
                    if (!string.IsNullOrEmpty((string)cut.DesignerClipboardRecoveryId) && (bool)cut.DesignerChangeObserved)
                        FormCut?.Invoke(new FormCutChange { Project = request.Project, Form = request.Form, ParentPath = request.ParentPath,
                            RecoveryId = (string)cut.DesignerClipboardRecoveryId, ControlCount = ((System.Collections.ICollection)cut.Before.Selected).Count, Owner = this });
                }
                if (result.Ok && name == "native_code_history" && !restoring)
                {
                    dynamic history = result.Data;
                    foreach (CodeChange change in history.Changes) CodeEdited?.Invoke(change);
                }
                if ((result.Ok || name == "cut_code" || name == "paste_code") && beforeCode != null && !restoring)
                {
                    try
                    {
                        CodeSnapshot afterCode = ReadCode(request.Project, editedModule);
                        if (!string.Equals(beforeCode.Sha256, afterCode.Sha256, StringComparison.OrdinalIgnoreCase))
                        {
                            int lineCount = CodeRollback.Lines(afterCode.Code).Length;
                            CodeEdited?.Invoke(new CodeChange(request.Project, editedModule,
                                beforeCode.Code, beforeCode.Sha256, afterCode.Code, afterCode.Sha256, lineCount));
                        }
                    }
                    catch (Exception ex) { WriteLog("Code diff readback failed: " + ex.Message); }
                }
                return json.Serialize(result);
            }
            catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
        }

        /// <summary>Couple le contenu actuel d’un module à son empreinte de version.</summary>
        private sealed class CodeSnapshot
        {
            /// <summary>Code source courant.</summary>
            public string Code;
            /// <summary>Empreinte SHA-256 du code courant.</summary>
            public string Sha256;
        }

        /// <summary>Lit le code d’un module et son empreinte auprès de la session VBE.</summary>
        /// <param name="project">Projet ciblé.</param>
        /// <param name="module">Module ciblé.</param>
        /// <returns>Code et empreinte lus.</returns>
        private CodeSnapshot ReadCode(string project, string module)
        {
            Response response = Execute(new Request { Command = "read_module",
                Project = project, Module = module });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            dynamic data = response.Data;
            return new CodeSnapshot { Code = (string)data.Code, Sha256 = (string)data.Sha256 };
        }

        /// <summary>Restaure toutes les zones restantes d’un changement après vérification et lecture du code courant.</summary>
        /// <param name="change">Changement à restaurer.</param>
        /// <returns>Réponse indiquant le résultat de la restauration.</returns>
        public Response RestoreCodeChange(CodeChange change)
        {
            return RestoreChanges(new[] { change }, null);
        }

        /// <summary>Restaure un bloc ou les changements d’un tour après vérification des conflits pour chaque module.</summary>
        /// <param name="changes">Modifications ordonnées à traiter.</param>
        /// <param name="hunk">Index de bloc facultatif à restaurer pour une modification unique.</param>
        /// <returns>Résultat des opérations ou réponse d’échec lorsque la prévalidation détecte un conflit.</returns>
        public Response RestoreChanges(CodeChange[] changes, int? hunk)
        {
            try
            {
                ValidateScope?.Invoke();
                var pending = changes.Where(x => x != null && !x.Restored).Reverse().ToArray();
                if (pending.Length == 0) return Response.Failure(UiText.Get("These changes have already been undone."));
                var snapshots = new Dictionary<string, CodeSnapshot>();
                var planned = new Dictionary<string, string>();
                foreach (var change in pending)
                {
                    string key = change.Project + "\0" + change.Module;
                    if (!snapshots.ContainsKey(key)) { snapshots[key] = ReadCode(change.Project, change.Module); planned[key] = snapshots[key].Code; }
                    planned[key] = CodeRollback.Apply(change, planned[key], hunk);
                }
                // Preflight every module before applying anything. Each write still checks its live SHA.
                restoring = true;
                int applied = 0;
                foreach (var group in pending.GroupBy(x => x.Project + "\0" + x.Module))
                {
                    var change = group.First(); var snapshot = snapshots[group.Key];
                    var arguments = new { Project = change.Project, Module = change.Module, ExpectedSha256 = snapshot.Sha256,
                        StartLine = 1, Count = CodeRollback.Lines(snapshot.Code).Length, Text = planned[group.Key] };
                    var result = ReadToolResponse(Invoke("replace_lines", json.Serialize(arguments)));
                    if (!result.Ok) return Response.Failure(UiText.Get("Undo stopped after ") + applied + " module(s). " + result.Error);
                    foreach (var item in group)
                    {
                        foreach (var block in CodeRollback.Hunks(item.Before, item.After).Where(x => !hunk.HasValue || x.Index == hunk.Value))
                            if (!item.RestoredHunks.Contains(block.Index)) item.RestoredHunks.Add(block.Index);
                        item.Restored = item.RestoredHunks.Count == CodeRollback.Hunks(item.Before, item.After).Length;
                    }
                    applied++;
                }
                return Response.Success(new { RestoredModules = applied });
            }
            catch (Exception ex) { return Response.Failure(ex.Message); }
            finally { restoring = false; }
        }

        /// <summary>Exécute un outil en tenant compte des opérations asynchrones et des confirmations d’édition.</summary>
        /// <param name="name">Nom de l’outil demandé.</param>
        /// <param name="arguments">Arguments JSON.</param>
        /// <returns>JSON de la réponse d’outil.</returns>
        public async Task<string> InvokeAsync(string name, string arguments)
        {
            try { GuardMode(name); GuardProject(name, arguments); }
            catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            if (name.StartsWith("git_", StringComparison.Ordinal)) return await InvokeGitAsync(name, arguments);
            if (name == "sign_project")
            {
                try
                {
                    await Task.Run(() => Native.EnsureNoSignatureDialog());
                    string scheduled = Invoke(name, arguments);
                    Response initial = ReadToolResponse(scheduled);
                    if (!initial.Ok) return scheduled;
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    var scheduledData = initial.Data as IDictionary<string, object>;
                    if (scheduledData == null || !scheduledData.ContainsKey("CertificateName"))
                        throw new InvalidOperationException("The certificate name was not returned by the VBE.");
                    string certificateName = (string)scheduledData["CertificateName"];
                    bool unsignedVerified = (bool)scheduledData["UnsignedVerified"];
                    object signed = await Task.Run(() => Native.CompleteProjectSignature(
                        (string)values["Project"], (string)values["CertificateThumbprint"], certificateName,
                        unsignedVerified));
                    object persistence = null;
                    string persistenceError = null;
                    for (int attempt = 0; attempt < 12; attempt++)
                    {
                        try
                        {
                            persistence = PersistSignature((string)values["Project"]);
                            persistenceError = null;
                            break;
                        }
                        catch (Exception ex)
                        {
                            persistenceError = ex.Message;
                            if (attempt == 11 || ex.ToString().IndexOf("0x800AC472",
                                StringComparison.OrdinalIgnoreCase) < 0) break;
                            await Task.Delay(250);
                        }
                    }
                    Response status = Execute(new Request { Command = "project_signature_status",
                        Project = (string)values["Project"] });
                    return json.Serialize(Response.Success(new { Signature = signed,
                        Persistence = persistence, PersistenceError = persistenceError,
                        SaveRequired = persistence == null || !((bool)((dynamic)persistence).Saved),
                        HostStatus = status.Ok ? status.Data : null,
                        HostStatusError = status.Ok ? null : status.Error }));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "read_project_protection" || name == "set_project_protection")
            {
                try
                {
                    await Task.Run(() => Native.EnsureNoProjectPropertiesDialog());
                    string scheduled = Invoke(name, arguments);
                    var initial = ReadToolResponse(scheduled);
                    if (!initial.Ok) return scheduled;
                    var data = json.DeserializeObject(json.Serialize(initial.Data)) as IDictionary<string, object>;
                    var request = json.Deserialize<Request>(arguments);
                    request.Caption = (string)data["ProjectName"];
                    return json.Serialize(Response.Success(await Task.Run(() => name == "set_project_protection" ?
                        Native.SetProjectProtection(request) : Native.ReadProjectProtection(request))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "read_project_signature_dialog")
            {
                try
                {
                    await Task.Run(() => Native.EnsureNoSignatureDialog());
                    string scheduled = Invoke(name, arguments);
                    Response initial = ReadToolResponse(scheduled);
                    if (!initial.Ok) return scheduled;
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    return json.Serialize(Response.Success(await Task.Run(() =>
                        Native.ReadSignatureDialog((string)values["Project"]))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "read_debug_options" || name == "read_vbe_options" || name == "set_vbe_option")
            {
                try
                {
                    await Task.Run(() => Native.EnsureNoDebugOptionsDialog());
                    string scheduled = Invoke(name, arguments);
                    Response initial = ReadToolResponse(scheduled);
                    if (!initial.Ok) return scheduled;
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    Request optionRequest = name == "set_vbe_option" ? json.Deserialize<Request>(json.Serialize(values)) : null;
                    return json.Serialize(Response.Success(await Task.Run(() => name == "set_vbe_option" ? Native.SetVbeOption(optionRequest) : name == "read_vbe_options"
                        ? Native.ReadVbeOptions() : Native.ReadDebugOptions())));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "quick_watch")
            {
                try
                {
                    if (settings.VbeEditApproval != "Automatic")
                        return json.Serialize(Response.Failure("Automatic VBE edit policy is required for Quick Watch evaluation."));
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null) throw new ArgumentException("Tool arguments must be an object.");
                    var requestValues = new Dictionary<string, object>(values) { ["Command"] = name };
                    Request request = json.Deserialize<Request>(json.Serialize(requestValues));
                    string scheduled = Invoke(name, arguments);
                    Response initial = ReadToolResponse(scheduled);
                    if (!initial.Ok) return scheduled;
                    return json.Serialize(Response.Success(await Task.Run(() => Native.CompleteQuickWatch(request))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "edit_watch")
            {
                try
                {
                    if (settings.VbeEditApproval != "Automatic")
                        return json.Serialize(Response.Failure("Automatic VBE edit policy is required for native watch editing."));
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null) throw new ArgumentException("Tool arguments must be an object.");
                    var requestValues = new Dictionary<string, object>(values) { ["Command"] = name };
                    Request request = json.Deserialize<Request>(json.Serialize(requestValues));
                    await Task.Run(() => Native.SelectWatch(request));
                    string scheduled = Invoke(name, arguments);
                    Response initial = ReadToolResponse(scheduled);
                    if (!initial.Ok) return scheduled;
                    return json.Serialize(Response.Success(await Task.Run(() => Native.CompleteEditWatch(request))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "read_navigation_surface" || name == "change_navigation_surface")
            {
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    string[] allowed = name == "read_navigation_surface" ? new[] { "Pane", "Query", "Offset", "Limit" } : new[] { "Pane", "Control", "Action", "ExpectedWindowVersion" };
                    if (values == null || !values.ContainsKey("Pane") || values.Keys.Any(k => !allowed.Contains(k)) ||
                        values.Any(pair => pair.Key == "Offset" || pair.Key == "Limit" ? !(pair.Value is int) : !(pair.Value is string)))
                        throw new ArgumentException("Invalid native navigation arguments.");
                    var request = json.Deserialize<Request>(arguments);
                    return json.Serialize(Response.Success(await Task.Run(() => name == "read_navigation_surface" ?
                        Native.ReadNavigationSurface(request) : Native.ChangeNavigationSurface(request))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "debug_item")
            {
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null || (values.Count != 3 && values.Count != 4) || !values.ContainsKey("Pane") ||
                        !values.ContainsKey("Action") || !values.ContainsKey("PathSegments"))
                        throw new ArgumentException("Pane, Action and PathSegments are required.");
                    if (values.Keys.Any(key => key != "Pane" && key != "Action" &&
                        key != "PathSegments" && key != "Context"))
                        throw new ArgumentException("Unexpected debug_item argument.");
                    var requestValues = new Dictionary<string, object>(values) { ["Command"] = name };
                    Request request = json.Deserialize<Request>(json.Serialize(requestValues));
                    return json.Serialize(Response.Success(await Task.Run(() => Native.ChangeDebugItem(request))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "immediate_execute")
            {
                try
                {
                    if (settings.VbeEditApproval != "Automatic")
                        return json.Serialize(Response.Failure("Automatic VBE edit policy is required for Immediate execution."));
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null || values.Count != 3 || !values.ContainsKey("Project") ||
                        !values.ContainsKey("ExpectedMode") || !values.ContainsKey("Text") ||
                        !(values["Project"] is string) || !(values["ExpectedMode"] is int) ||
                        !(values["Text"] is string) ||
                        ((int)values["ExpectedMode"] != 1 && (int)values["ExpectedMode"] != 2))
                        throw new ArgumentException("Project, ExpectedMode and Text are required.");
                    var state = Execute(new Request { Command = "debug_state", Project = (string)values["Project"] });
                    if (!state.Ok) return json.Serialize(state);
                    if ((int)((dynamic)state.Data).Mode != (int)values["ExpectedMode"])
                        return json.Serialize(Response.Failure("Project mode changed before Immediate execution."));
                    return json.Serialize(Response.Success(await Task.Run(() =>
                        Native.ExecuteImmediate((string)values["Text"]))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "debug_dialog" || name == "respond_debug_dialog")
            {
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null) throw new ArgumentException("Tool arguments must be an object.");
                    if (name == "debug_dialog")
                    {
                        if (values.Count != 0) throw new ArgumentException("debug_dialog has no arguments.");
                        return json.Serialize(Response.Success(await Task.Run(() => Native.ReadDebugDialog())));
                    }
                    if (settings.VbeEditApproval != "Automatic")
                        return json.Serialize(Response.Failure("Automatic VBE edit policy is required to respond to a diagnostic dialog."));
                    if (values.Count != 2 || !values.ContainsKey("Diagnostic") || !values.ContainsKey("Button") ||
                        !(values["Diagnostic"] is string) || !(values["Button"] is string))
                        throw new ArgumentException("Exact Diagnostic and Button strings are required.");
                    var requestValues = new Dictionary<string, object>(values) { ["Command"] = name };
                    Request request = json.Deserialize<Request>(json.Serialize(requestValues));
                    return json.Serialize(Response.Success(await Task.Run(() => Native.RespondDebugDialog(request))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "compile_project")
            {
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null || values.Count != 2 || !values.ContainsKey("Project") ||
                        !values.ContainsKey("ExpectedMode") || !(values["Project"] is string) ||
                        string.IsNullOrWhiteSpace((string)values["Project"]) ||
                        !(values["ExpectedMode"] is int) || (int)values["ExpectedMode"] != 2)
                        throw new ArgumentException("Project and ExpectedMode=2 are required.");
                    var requestValues = new Dictionary<string, object>(values) { ["Command"] = name };
                    Request request = json.Deserialize<Request>(json.Serialize(requestValues));
                    SynchronizationContext context = SynchronizationContext.Current;
                    if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
                    Native.EnsureNoCompileDialog();
                    Response compileResponse = null;
                    var completed = new System.Threading.ManualResetEventSlim(false);
                    context.Post(_ => {
                        try { compileResponse = Execute(request); }
                        catch (Exception ex) { compileResponse = Response.Failure(ex.Message); }
                        finally { completed.Set(); }
                    }, null);
                    string diagnostic = await Task.Run(() => Native.AwaitCompileDialog(completed));
                    if (compileResponse == null) return json.Serialize(Response.Failure("The native Compile command did not return a result."));
                    if (!compileResponse.Ok) return json.Serialize(compileResponse);
                    return json.Serialize(Response.Success(new {
                        Project = request.Project, Compiled = diagnostic == null, Diagnostic = diagnostic,
                        Verification = diagnostic == null ? "NoNativeDiagnosticObserved" : "NativeDiagnosticCaptured",
                        Command = compileResponse.Data,
                        NextRead = diagnostic == null ? null : "Read debug_state to locate the selected token."
                    }));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "remove_watch")
            {
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null) throw new ArgumentException("Tool arguments must be an object.");
                    if (settings.VbeEditApproval != "Automatic")
                        return json.Serialize(Response.Failure("Automatic VBE edit policy is required for native watch removal."));
                    var requestValues = new Dictionary<string, object>(values) { ["Command"] = name };
                    Request request = json.Deserialize<Request>(json.Serialize(requestValues));
                    await Task.Run(() => Native.SelectWatch(request));
                    string executed = Invoke(name, arguments);
                    Response response = ReadToolResponse(executed);
                    if (!response.Ok) return executed;
                    object result = await Task.Run(() => Native.VerifyWatchRemoved(request));
                    return json.Serialize(Response.Success(result));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "add_watch")
            {
                string scheduled = Invoke(name, arguments);
                Response initial = ReadToolResponse(scheduled);
                if (!initial.Ok) return scheduled;
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    var requestValues = new Dictionary<string, object>(values) { ["Command"] = name };
                    Request request = json.Deserialize<Request>(json.Serialize(requestValues));
                    object result = await Task.Run(() => Native.CompleteAddWatch(request));
                    return json.Serialize(Response.Success(result));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "list_object_browser")
            {
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null || !values.ContainsKey("Pane") || values.Keys.Any(key => key != "Pane" && key != "Query" && key != "Offset" && key != "Limit") ||
                        values.Any(pair => (pair.Key == "Pane" || pair.Key == "Query") ? !(pair.Value is string) : !(pair.Value is int)))
                        throw new ArgumentException("Pane is required; Query is an optional string; Offset and Limit must be integers.");
                    return json.Serialize(Response.Success(await Task.Run(() => Native.ListObjectBrowser(json.Deserialize<Request>(arguments)))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "select_object_browser")
            {
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null || (!values.ContainsKey("ObjectName") && !values.ContainsKey("Context")) || values.Keys.Any(key => key != "ObjectName" && key != "Procedure" && key != "Context") || values.Values.Any(value => !(value is string)))
                        throw new ArgumentException("ObjectName or library Context is required; Procedure requires ObjectName. All values must be strings.");
                    var request = json.Deserialize<Request>(arguments);
                    return json.Serialize(Response.Success(await Task.Run(() => Native.SelectObjectBrowser(request))));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name == "read_object_browser" || name == "read_runtime_forms")
            {
                try
                {
                    var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                    if (values == null || values.Count != 0) throw new ArgumentException("This tool takes an empty argument object.");
                    return json.Serialize(Response.Success(await Task.Run(() => name == "read_runtime_forms" ? Native.ReadRuntimeForms() : Native.ReadObjectBrowser())));
                }
                catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
            }
            if (name != "debug_windows") return Invoke(name, arguments);
            try
            {
                var values = json.DeserializeObject(arguments) as IDictionary<string, object>;
                if (values == null) throw new ArgumentException("Tool arguments must be an object.");
                foreach (string field in values.Keys)
                    if (field != "IncludeCallStack") throw new ArgumentException("Unexpected argument: " + field);
                object raw;
                bool stack = values.TryGetValue("IncludeCallStack", out raw) && raw is bool && (bool)raw;
                if (values.ContainsKey("IncludeCallStack") && !(raw is bool))
                    throw new ArgumentException("IncludeCallStack must be a boolean.");
                object result = await Task.Run(() => Native.Capture(stack));
                return json.Serialize(Response.Success(result));
            }
            catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
        }

        /// <summary>Désérialise une réponse d’outil et conserve la défense contre un JSON null.</summary>
        /// <param name="serialized">Réponse JSON à lire.</param>
        /// <returns>Réponse toujours non nulle ; un JSON null devient un échec sans message.</returns>
        internal Response ReadToolResponse(string serialized)
        {
            return json.Deserialize<Response>(serialized) ?? Response.Failure(null);
        }

        /// <summary>Retourne l’instantané JSON du contexte VBE actuellement ouvert.</summary>
        /// <returns>Instantané sérialisé du contexte.</returns>
        public string LiveContextJson()
        {
            return json.Serialize(LlmVbeContext.LiveSnapshot(session));
        }

        /// <summary>Mémorise une demande utilisateur non vide pour autoriser la lecture de son chemin explicite.</summary>
        /// <param name="request">Texte de la demande.</param>
        public void NoteUserRequest(string request)
        {
            if (!string.IsNullOrWhiteSpace(request)) userRequests.Add(request);
        }

        /// <summary>Demande confirmation, puis lit au plus 64 Kio du fichier texte explicitement demandé.</summary>
        /// <param name="requestedPath">Chemin absolu fourni explicitement dans une requête utilisateur.</param>
        /// <returns>Réponse contenant le texte, sa longueur et l’indication de troncature, ou une erreur de refus.</returns>
        private Response ReadUserFile(string requestedPath)
        {
            if (!IsExplicitUserPath(requestedPath))
                return Response.Failure(UiText.Get("The user did not provide this exact path."));
            string fullPath = Path.GetFullPath(requestedPath);
            if (!File.Exists(fullPath)) return Response.Failure(UiText.Get("The provided file could not be found."));
            const int limit = 65536;
            string provider = string.IsNullOrWhiteSpace(CurrentProviderName) ? UiText.Get("the active LLM provider") : CurrentProviderName;
            var choice = ConfirmFile(owner,
                UiText.Get("Allow reading and sending to ") + provider + UiText.Get(" up to the first 64 KiB of this file?\r\n\r\n") + fullPath,
                UiText.Get("VBAi — allow sending a file"));
            if (choice != DialogResult.Yes) return Response.Failure(UiText.Get("The user declined sending the file."));
            using (var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var bytes = new byte[limit + 1];
                int count = 0;
                while (count < bytes.Length)
                {
                    int read = stream.Read(bytes, count, bytes.Length - count);
                    if (read == 0) break;
                    count += read;
                }
                int included = Math.Min(count, limit);
                using (var memory = new MemoryStream(bytes, 0, included))
                using (var reader = new StreamReader(memory, Encoding.UTF8, true))
                {
                    string content = reader.ReadToEnd();
                    if (content.IndexOf('\0') >= 0)
                        return Response.Failure(UiText.Get("The file appears to be binary; text reading refused."));
                    return Response.Success(new { Path = fullPath, Text = content,
                        Truncated = count > limit, ByteLength = stream.Length });
                }
            }
        }

        /// <summary>Vérifie que le chemin absolu demandé figure littéralement dans une requête utilisateur mémorisée.</summary>
        /// <param name="requestedPath">Chemin à vérifier.</param>
        /// <returns><see langword="true"/> si une requête antérieure mentionne ce chemin.</returns>
        private bool IsExplicitUserPath(string requestedPath)
        {
            return !string.IsNullOrWhiteSpace(requestedPath) &&
                Regex.IsMatch(requestedPath, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])") &&
                userRequests.Any(request => request.IndexOf(requestedPath, StringComparison.OrdinalIgnoreCase) >= 0);
        }
    }
}
