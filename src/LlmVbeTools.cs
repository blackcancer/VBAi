using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed class LlmVbeTools
    {
        private readonly VbeSession session;
        private readonly IWin32Window owner;
        private readonly LlmSettings settings;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        private readonly List<string> userRequests = new List<string>();
        private static readonly HashSet<string> ReadOnlyTools = new HashSet<string>(StringComparer.Ordinal) {
            "status", "read_user_file", "list_projects", "list_modules", "list_references", "read_module", "list_forms",
            "form_state", "form_tree", "form_properties", "form_control_properties",
            "list_form_control_types", "open_form"
        };
        public string CurrentProviderName { get; set; }

        public LlmVbeTools(VbeSession session, IWin32Window owner, LlmSettings settings)
        {
            this.session = session;
            this.owner = owner;
            this.settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        private static object Definition(string name, string description, string[] required, params string[] fields)
        {
            var properties = new Dictionary<string, object>();
            foreach (string field in fields)
                properties[field] = field == "Value" ? (object)new { anyOf = new object[] {
                    new { type = "string" }, new { type = "number" }, new { type = "boolean" } } } :
                    new { type = field == "StartLine" || field == "Count" || field == "ExpectedMode" ||
                        field == "Major" || field == "Minor" ? "integer" :
                    field == "Left" || field == "Top" || field == "Width" || field == "Height" || field == "FontSize" ? "number" :
                    field == "FontBold" ? "boolean" : "string" };
            return new { type = "function", function = new {
                name, description,
                parameters = new { type = "object", properties, required, additionalProperties = false }
            } };
        }

        public static object[] Definitions { get { return new object[] {
            Definition("status", "Read the live host process and currently open VBA projects; call before acting on VBE.", new string[0]),
            Definition("read_user_file", "Request separate user approval before reading and transmitting up to 64 KiB of a text file at a path explicitly supplied by the user.",
                new[] { "Path" }, "Path"),
            Definition("list_projects", "List open VBA projects and their modes.", new string[0]),
            Definition("list_modules", "List modules in one VBA project.", new[] { "Project" }, "Project"),
            Definition("list_references", "List the type-library references actually selected by one VBA project, including identity, version, path and broken status.",
                new[] { "Project" }, "Project"),
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
        }; } }

        public string Invoke(string name, string arguments)
        {
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
                    name == "add_reference_file") &&
                    !IsExplicitUserPath((string)values["Path"]))
                    return json.Serialize(Response.Failure("L'utilisateur doit fournir explicitement le chemin absolu du fichier."));
                var normalized = new Dictionary<string, object>(values) { ["Command"] = name };
                var request = json.Deserialize<Request>(json.Serialize(normalized));
                // All newly registered tools are treated as edits unless explicitly classified as read-only.
                bool edit = !ReadOnlyTools.Contains(name);
                if (edit && settings.VbeEditApproval == "ReadOnly")
                    return json.Serialize(Response.Failure("Les modifications VBE sont désactivées (mode Lecture seule)."));
                if (edit && settings.VbeEditApproval != "Automatic" && settings.VbeEditApproval != "AskEachTime")
                    return json.Serialize(Response.Failure("Politique de modification VBE inconnue ; action refusée."));
                if (edit && settings.VbeEditApproval == "AskEachTime")
                {
                    string summary = name + "\r\n\r\n" + json.Serialize(values);
                    using (var approval = new Form { Text = "CodexVBE — valider la modification", Width = 740,
                        Height = 530, StartPosition = FormStartPosition.CenterParent, MinimizeBox = false,
                        MaximizeBox = true, ShowInTaskbar = false })
                    {
                        var details = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both,
                            WordWrap = false, Dock = DockStyle.Fill, Text = summary };
                        var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44,
                            FlowDirection = FlowDirection.RightToLeft };
                        var approve = new Button { Text = "Autoriser", DialogResult = DialogResult.Yes, Width = 100 };
                        var reject = new Button { Text = "Refuser", DialogResult = DialogResult.No, Width = 100 };
                        actions.Controls.Add(approve);
                        actions.Controls.Add(reject);
                        approval.Controls.Add(details);
                        approval.Controls.Add(actions);
                        approval.CancelButton = reject;
                        if (approval.ShowDialog(owner) != DialogResult.Yes)
                            return json.Serialize(Response.Failure("User rejected the edit."));
                    }
                }
                return json.Serialize(name == "status"
                    ? Response.Success(LlmVbeContext.LiveSnapshot(session))
                    : session.Execute(request));
            }
            catch (Exception ex) { return json.Serialize(Response.Failure(ex.Message)); }
        }

        public string LiveContextJson()
        {
            return json.Serialize(LlmVbeContext.LiveSnapshot(session));
        }

        public void NoteUserRequest(string request)
        {
            if (!string.IsNullOrWhiteSpace(request)) userRequests.Add(request);
        }

        private Response ReadUserFile(string requestedPath)
        {
            if (!IsExplicitUserPath(requestedPath))
                return Response.Failure("L'utilisateur n'a pas fourni ce chemin exact.");
            string fullPath = Path.GetFullPath(requestedPath);
            if (!File.Exists(fullPath)) return Response.Failure("Le fichier fourni est introuvable.");
            const int limit = 65536;
            string provider = string.IsNullOrWhiteSpace(CurrentProviderName) ? "le fournisseur LLM actif" : CurrentProviderName;
            var choice = MessageBox.Show(owner,
                "Autoriser la lecture et la transmission à " + provider + " des 64 premiers Kio au maximum de ce fichier ?\r\n\r\n" + fullPath,
                "CodexVBE — autoriser la transmission du fichier", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            if (choice != DialogResult.Yes) return Response.Failure("L'utilisateur a refusé la transmission du fichier.");
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
                        return Response.Failure("Le fichier semble binaire ; lecture textuelle refusée.");
                    return Response.Success(new { Path = fullPath, Text = content,
                        Truncated = count > limit, ByteLength = stream.Length });
                }
            }
        }

        private bool IsExplicitUserPath(string requestedPath)
        {
            return !string.IsNullOrWhiteSpace(requestedPath) &&
                Regex.IsMatch(requestedPath, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])") &&
                userRequests.Any(request => request.IndexOf(requestedPath, StringComparison.OrdinalIgnoreCase) >= 0);
        }
    }
}
