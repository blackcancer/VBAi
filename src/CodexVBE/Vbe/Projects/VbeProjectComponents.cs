using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    /// <summary>Inspecte les projets VBIDE et applique les opérations de gestion de composants explicitement validées.</summary>
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>Obtient le PID du processus propriétaire d’une fenêtre Win32.</summary>
        /// <param name="window">Handle de la fenêtre.</param>
        /// <param name="processId">Reçoit l’identifiant du processus propriétaire.</param>
        /// <returns>Identifiant du thread créateur de la fenêtre.</returns>
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        /// <summary>Fournit les informations Excel nécessaires pour associer sûrement un projet VBE à son classeur.</summary>
        internal interface IExcelHostProbe
        {
            /// <summary>Indique si l’hôte courant est Excel.</summary>
            /// <value><see langword="true"/> si le processus hôte est Excel.</value>
            bool IsExcel { get; }
            /// <summary>Obtient l’identifiant du processus courant.</summary>
            /// <value>PID du processus hôte.</value>
            int CurrentProcessId { get; }
            /// <summary>Obtient l’objet d’application Excel actif.</summary>
            /// <returns>Application Excel exposant la collection de classeurs.</returns>
            object ExcelApplication();
            /// <summary>Obtient le PID propriétaire d’un handle de fenêtre.</summary>
            /// <param name="window">Handle de fenêtre à vérifier.</param>
            /// <returns>Identifiant du processus propriétaire.</returns>
            uint WindowProcessId(IntPtr window);
        }

        /// <summary>Implémente le sondage Excel avec les API Windows et l’objet d’application actif.</summary>
        private sealed class NativeExcelHostProbe : IExcelHostProbe
        {
            /// <summary>Indique si le processus courant est EXCEL.EXE.</summary>
            /// <value><see langword="true"/> lorsque le nom du processus est EXCEL.</value>
            public bool IsExcel => string.Equals(Process.GetCurrentProcess().ProcessName,
                "EXCEL", StringComparison.OrdinalIgnoreCase);
            /// <summary>Obtient le PID du processus courant.</summary>
            /// <value>PID du processus courant.</value>
            public int CurrentProcessId => Process.GetCurrentProcess().Id;
            /// <summary>Obtient l’application Excel du processus via NativeOM, avec repli ROT protégé par PID.</summary>
            /// <returns>Objet COM Excel.Application.</returns>
            public object ExcelApplication() { return ExcelOwnedApplication.Resolve(CurrentProcessId, RegisteredExcel); }
            /// <summary>Obtient l’entrée ROT historique si aucune fenêtre de document du processus ne fournit NativeOM.</summary>
            /// <returns>Objet COM enregistré, dont le PID doit encore être vérifié par l’appelant.</returns>
            private static object RegisteredExcel() { return Marshal.GetActiveObject("Excel.Application"); }
            /// <summary>Retourne le PID propriétaire de la fenêtre native.</summary>
            /// <param name="window">Handle de la fenêtre Excel.</param>
            /// <returns>Identifiant du processus associé à la fenêtre.</returns>
            public uint WindowProcessId(IntPtr window)
            {
                uint processId;
                GetWindowThreadProcessId(window, out processId);
                return processId;
            }
        }

        /// <summary>Instance VBE interrogée pour résoudre les projets.</summary>
        private readonly dynamic vbe;
        /// <summary>Gestionnaire de contrôles MSForms, utilisé pour lire l’état des formulaires.</summary>
        private readonly VbeForms forms;
        /// <summary>Accès injectable aux informations du processus hôte Excel.</summary>
        private readonly IExcelHostProbe host;
        /// <summary>Sérialiseur des états utilisés pour calculer les versions de projet et composant.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };

        /// <summary>Crée le service avec le sondage Excel natif.</summary>
        /// <param name="vbe">Instance VBIDE.</param>
        /// <param name="forms">Service de lecture des formulaires MSForms.</param>
        public VbeProjectComponents(object vbe, VbeForms forms)
            : this(vbe, forms, new NativeExcelHostProbe()) { }

        /// <summary>Crée le service avec un sondage d’hôte injectable.</summary>
        /// <param name="vbe">Instance VBIDE.</param>
        /// <param name="forms">Service de lecture des formulaires MSForms.</param>
        /// <param name="host">Sonde Excel, obligatoire.</param>
        /// <exception cref="ArgumentNullException">Le sondage d’hôte est nul.</exception>
        internal VbeProjectComponents(object vbe, VbeForms forms, IExcelHostProbe host)
        {
            this.vbe = vbe;
            this.forms = forms;
            this.host = host ?? throw new ArgumentNullException(nameof(host));
        }

        /// <summary>Retourne les propriétés, composants et références d’un projet avec une empreinte de version.</summary>
        /// <param name="projectName">Nom du projet à inspecter.</param>
        /// <returns>Instantané contenant les propriétés, composants, références et version calculée.</returns>
        public object ProjectProperties(string projectName)
        {
            dynamic project = GetProject(projectName);
            var properties = ReadProperties((object)project);
            var components = new List<object>();
            foreach (dynamic component in project.VBComponents)
                components.Add(new { Name = (string)component.Name, Type = (int)component.Type });
            var references = new List<object>();
            foreach (dynamic reference in project.References)
                references.Add(new { Guid = (string)reference.GUID,
                    Major = (int)reference.Major, Minor = (int)reference.Minor,
                    IsBroken = (bool)reference.IsBroken, BuiltIn = (bool)reference.BuiltIn });
            string version = Hash(json.Serialize(new { Mode = (int)project.Mode,
                Properties = properties, Components = components, References = references }));
            return new { Project = (string)project.Name, Mode = (int)project.Mode,
                Version = version, Properties = properties, Components = components,
                References = references };
        }

        /// <summary>Retourne l’instantané détaillé d’un composant de projet.</summary>
        /// <param name="projectName">Nom du projet contenant le composant.</param>
        /// <param name="componentName">Nom du composant à inspecter.</param>
        /// <returns>Propriétés, état du code et version du composant.</returns>
        public object ComponentProperties(string projectName, string componentName)
        {
            dynamic component = GetComponent(GetProject(projectName), componentName);
            return ComponentSnapshot(projectName, component);
        }

        /// <summary>Détermine si le projet correspond à un classeur Excel et si celui-ci signale un projet VBA signé.</summary>
        /// <param name="projectName">Nom du projet VBIDE sélectionné.</param>
        /// <returns>État de disponibilité, signature, source et raison en cas d’indisponibilité.</returns>
        public object SignatureStatus(string projectName)
        {
            dynamic project = GetProject(projectName);
            if (!host.IsExcel)
                return new { Project = projectName, Available = false, Signed = (bool?)null,
                    Source = "Host", Reason = "This host does not expose Excel.Workbook.VBASigned." };
            try
            {
                dynamic excel = host.ExcelApplication();
                uint excelProcessId = host.WindowProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)));
                if (excelProcessId != (uint)host.CurrentProcessId)
                    return new { Project = projectName, Available = false, Signed = (bool?)null,
                        Source = "Excel.Workbook.VBASigned", Reason = "The registered Excel instance is not this VBE host. Registered PID=" + excelProcessId + ", VBE PID=" + host.CurrentProcessId };
                string projectPath = null;
                try { projectPath = (string)project.FileName; }
                catch { /* An unsaved workbook may have no project path. */ }
                bool singleUnsavedProject = string.IsNullOrWhiteSpace(projectPath) &&
                    (int)excel.Workbooks.Count == 1 && (int)vbe.VBProjects.Count == 1;
                foreach (dynamic workbook in excel.Workbooks)
                {
                    bool matches = singleUnsavedProject ||
                        (!string.IsNullOrWhiteSpace(projectPath) &&
                         string.Equals(Path.GetFullPath((string)workbook.FullName),
                             Path.GetFullPath(projectPath), StringComparison.OrdinalIgnoreCase));
                    if (!matches) continue;
                    return new { Project = projectName, Available = true, Signed = (bool?)((bool)workbook.VBASigned),
                        Source = "Excel.Workbook.VBASigned", Reason = (string)null };
                }
                return new { Project = projectName, Available = false, Signed = (bool?)null,
                    Source = "Excel.Workbook.VBASigned", Reason = "No workbook matches the selected VBE project." };
            }
            catch (Exception ex)
            {
                return new { Project = projectName, Available = false, Signed = (bool?)null,
                    Source = "Excel.Workbook.VBASigned", Reason = ex.Message };
            }
        }

        /// <summary>Retourne les états de sauvegarde du projet VBIDE et du classeur Excel correspondant.</summary>
        /// <param name="projectName">Nom du projet à examiner.</param>
        /// <returns>État du projet et état du document hôte, ou raison de l’indisponibilité hôte.</returns>
        public object PersistenceStatus(string projectName)
        {
            dynamic project = GetProject(projectName);
            bool projectSaved = (bool)project.Saved;
            if (!host.IsExcel && SupportsStandaloneMacro((object)project)) return StandalonePersistence(projectName, (object)project);
            if (!host.IsExcel)
                return new { Project = projectName, ProjectSaved = projectSaved,
                    HostAvailable = false, HostPath = (string)null, HostSaved = (bool?)null,
                    HostReadOnly = (bool?)null, HostHasPath = (bool?)null,
                    Reason = "The host is not Excel; its document save state is unavailable." };
            try
            {
                dynamic workbook = MatchExcelWorkbook(project, true);
                string path = (string)workbook.Path;
                bool hasPath = !string.IsNullOrWhiteSpace(path);
                return new { Project = projectName, ProjectSaved = projectSaved,
                    HostAvailable = true, HostPath = hasPath ? (string)workbook.FullName : null,
                    HostSaved = (bool?)((bool)workbook.Saved),
                    HostReadOnly = (bool?)((bool)workbook.ReadOnly),
                    HostHasPath = (bool?)hasPath, Reason = (string)null };
            }
            catch (Exception ex)
            {
                return new { Project = projectName, ProjectSaved = projectSaved,
                    HostAvailable = false, HostPath = (string)null, HostSaved = (bool?)null,
                    HostReadOnly = (bool?)null, HostHasPath = (bool?)null, Reason = ex.Message };
            }
        }

        /// <summary>Enregistre le classeur Excel associé après vérification du chemin et de la version du projet.</summary>
        /// <param name="request">Requête portant le projet, sa version attendue et le chemin hôte attendu.</param>
        /// <returns>Résultat de sauvegarde avec états lus après l’appel.</returns>
        /// <exception cref="ArgumentException">Le chemin attendu est manquant ou non absolu.</exception>
        /// <exception cref="InvalidOperationException">L’hôte, le chemin, la version ou l’état inscriptible ne permet pas l’enregistrement.</exception>
        public object SaveHostDocument(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedHostPath) ||
                !Path.IsPathRooted(request.ExpectedHostPath))
                throw new ArgumentException("ExpectedHostPath must be the absolute path read from project_persistence_status.");
            if (!host.IsExcel) return SaveStandaloneMacro(request, false);
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            string projectPath = (string)project.FileName;
            if (string.IsNullOrWhiteSpace(projectPath) || !Path.IsPathRooted(projectPath) ||
                !string.Equals(Path.GetFullPath(projectPath), Path.GetFullPath(request.ExpectedHostPath),
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected VBA project's workbook path changed since it was read.");
            dynamic workbook = MatchExcelWorkbook(project, false);
            if ((bool)workbook.ReadOnly) throw new InvalidOperationException("The workbook is read-only and cannot be saved.");
            bool beforeHostSaved = (bool)workbook.Saved;
            bool beforeProjectSaved = (bool)project.Saved;
            workbook.Save();
            bool hostSaved = (bool)workbook.Saved;
            bool projectSaved = (bool)project.Saved;
            if (!hostSaved || !projectSaved)
                throw new InvalidOperationException("Excel did not mark the workbook and VBA project as saved; a BeforeSave handler may have cancelled the save.");
            return new { Project = request.Project, HostPath = projectPath,
                SaveInvoked = true, HostSavedBefore = beforeHostSaved,
                ProjectSavedBefore = beforeProjectSaved, HostSaved = hostSaved,
                ProjectSaved = projectSaved,
                Verification = "ExcelWorkbookSaveAndSavedReadback",
                Limit = "The host reported Saved=true. Reopen the file to verify that a specific code edit persisted on disk." };
        }

        /// <summary>Enregistre pour la première fois un projet Excel non enregistré au chemin .xlsm demandé.</summary>
        /// <param name="request">Requête comprenant le chemin de destination et la version attendue du projet.</param>
        /// <returns>Chemins lus après SaveAs, taille du fichier et état de sauvegarde.</returns>
        /// <exception cref="ArgumentException">Le chemin ou la version attendue est invalide, ou l’extension n’est pas .xlsm.</exception>
        /// <exception cref="IOException">La destination existe déjà.</exception>
        /// <exception cref="DirectoryNotFoundException">Le dossier de destination n’existe pas.</exception>
        /// <exception cref="InvalidOperationException">L’hôte, le classeur ou la vérification après sauvegarde échoue.</exception>
        public object SaveHostDocumentAs(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Path) ||
                string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("Path and ExpectedProjectVersion are required.");
            if (!host.IsExcel) return SaveStandaloneMacro(request, true);
            string path = RequireAbsolutePath(request.Path);
            if (!string.Equals(Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("The first Excel SaveAs supports only a macro-enabled .xlsm workbook.");
            if (File.Exists(path)) throw new IOException("SaveAs destination already exists: " + path);
            if (!Directory.Exists(Path.GetDirectoryName(path)))
                throw new DirectoryNotFoundException("The SaveAs destination directory does not exist.");
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            dynamic workbook = MatchExcelWorkbook(project, true);
            if ((bool)workbook.ReadOnly)
                throw new InvalidOperationException("The workbook is read-only and cannot be saved.");
            string oldProjectPath = null;
            try { oldProjectPath = (string)project.FileName; }
            catch { } // VBProject.FileName can fail before the first save.
            if (!string.IsNullOrWhiteSpace((string)workbook.Path) ||
                !string.IsNullOrWhiteSpace(oldProjectPath))
                throw new InvalidOperationException("This command is only for the first save of an unsaved Excel VBA project.");
            workbook.SaveAs(path, 52); // xlOpenXMLWorkbookMacroEnabled
            string actual = Path.GetFullPath((string)workbook.FullName);
            string projectPath = Path.GetFullPath((string)project.FileName);
            if (!string.Equals(actual, path, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(projectPath, path, StringComparison.OrdinalIgnoreCase) ||
                !File.Exists(path) || !(bool)workbook.Saved || !(bool)project.Saved)
                throw new InvalidOperationException("Excel SaveAs returned without matching saved workbook and project paths.");
            return new { Project = request.Project, HostPath = actual, ProjectPath = projectPath,
                SaveAsInvoked = true, Bytes = new FileInfo(path).Length,
                HostSaved = true, ProjectSaved = true,
                Verification = "ExcelWorkbookAndProjectPathReadback",
                Limit = "Reopen the .xlsm to verify persistence of a specific VBA edit." };
        }

        /// <summary>Associe le projet au classeur Excel du même hôte en comparant leurs chemins complets.</summary>
        /// <param name="project">Projet VBIDE à associer.</param>
        /// <param name="allowUnsaved">Autorise l’unique classeur et l’unique projet VBE si le projet n’a pas encore de chemin.</param>
        /// <returns>Classeur Excel correspondant sans ambiguïté.</returns>
        /// <exception cref="InvalidOperationException">Le processus Excel est différent, le chemin manque ou aucun classeur unique ne correspond.</exception>
        private dynamic MatchExcelWorkbook(dynamic project, bool allowUnsaved)
        {
            dynamic excel = host.ExcelApplication();
            uint excelProcessId = host.WindowProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)));
            if (excelProcessId != (uint)host.CurrentProcessId)
                throw new InvalidOperationException("The registered Excel instance is not this VBE host. Registered PID=" + excelProcessId + ", VBE PID=" + host.CurrentProcessId);
            string projectPath = null;
            try { projectPath = (string)project.FileName; }
            catch { }
            if (string.IsNullOrWhiteSpace(projectPath) && allowUnsaved &&
                (int)excel.Workbooks.Count == 1 && (int)vbe.VBProjects.Count == 1)
                return excel.Workbooks.Item(1);
            if (string.IsNullOrWhiteSpace(projectPath) || !Path.IsPathRooted(projectPath))
                throw new InvalidOperationException("The project has no saved workbook path.");
            dynamic match = null;
            foreach (dynamic workbook in excel.Workbooks)
            {
                if (!string.Equals(Path.GetFullPath((string)workbook.FullName),
                    Path.GetFullPath(projectPath), StringComparison.OrdinalIgnoreCase)) continue;
                if (match != null) throw new InvalidOperationException("Multiple workbooks match the selected VBA project.");
                match = workbook;
            }
            if (match == null) throw new InvalidOperationException("No workbook matches the selected VBA project.");
            return match;
        }

        /// <summary>Enregistre le classeur associé à un projet signé et vérifie que la signature reste signalée par Excel.</summary>
        /// <param name="projectName">Nom du projet VBE.</param>
        /// <returns>Disponibilité, état de signature après sauvegarde et limites de vérification.</returns>
        /// <exception cref="InvalidOperationException">Le projet n’a pas de classeur enregistré correspondant, est en lecture seule ou perd son état signé.</exception>
        public object PersistExcelSignature(string projectName)
        {
            if (!host.IsExcel)
                return new { Available = false, Saved = false,
                    Reason = "The host is not Excel; save the host document with its native command." };
            dynamic project = GetProject(projectName);
            string projectPath = (string)project.FileName;
            if (string.IsNullOrWhiteSpace(projectPath) || !Path.IsPathRooted(projectPath))
                throw new InvalidOperationException("The Excel VBA project has no saved workbook path.");
            dynamic excel = host.ExcelApplication();
            uint excelProcessId = host.WindowProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)));
            if (excelProcessId != (uint)host.CurrentProcessId)
                throw new InvalidOperationException("The registered Excel instance is not this VBE host. Registered PID=" + excelProcessId + ", VBE PID=" + host.CurrentProcessId);
            dynamic match = null;
            foreach (dynamic workbook in excel.Workbooks)
            {
                if (!string.Equals(Path.GetFullPath((string)workbook.FullName),
                    Path.GetFullPath(projectPath), StringComparison.OrdinalIgnoreCase)) continue;
                if (match != null) throw new InvalidOperationException("Multiple workbooks match the selected VBA project.");
                match = workbook;
            }
            if (match == null) throw new InvalidOperationException("No workbook matches the selected VBA project.");
            if ((bool)match.ReadOnly) throw new InvalidOperationException("The signed workbook is read-only and cannot be saved.");
            if (!(bool)match.VBASigned) throw new InvalidOperationException("Excel does not report a signed VBA project before saving.");
            match.Save();
            if (!(bool)match.VBASigned) throw new InvalidOperationException("Excel no longer reports the VBA project as signed after saving.");
            return new { Available = true, Saved = true, Path = projectPath,
                Signed = true, Verification = "ExcelWorkbookSaveAndVBASignedReadback",
                Limit = "A reopening check is needed to prove the signature persisted on disk." };
        }

        /// <summary>Lit individuellement une propriété de concepteur du composant nommé.</summary>
        /// <param name="projectName">Nom du projet.</param>
        /// <param name="componentName">Nom du composant.</param>
        /// <param name="propertyName">Nom de la propriété à lire.</param>
        /// <returns>Résultat du sondage ciblé de la propriété.</returns>
        /// <exception cref="ArgumentException">Le nom de propriété est vide.</exception>
        public object ComponentPropertyValue(string projectName, string componentName, string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property is required.");
            return ComponentProbe(projectName, componentName, "designer_property_value", propertyName);
        }

        /// <summary>Exécute une étape de diagnostic COM ciblée sur un composant afin d’isoler les lectures bloquantes.</summary>
        /// <param name="projectName">Nom du projet.</param>
        /// <param name="componentName">Nom du composant.</param>
        /// <param name="stage">Étape prise en charge : identité, noms de propriétés, compte de code, empreinte ou valeur ciblée.</param>
        /// <param name="propertyName">Nom de propriété requis par les étapes de lecture d’une valeur.</param>
        /// <returns>Résultat propre à l’étape demandée.</returns>
        /// <exception cref="ArgumentException">Une propriété demandée est absente ou l’étape n’est pas prise en charge.</exception>
        /// <exception cref="InvalidOperationException">La lecture de propriété est bloquée pour le MailEnvelope d’un module de document Excel.</exception>
        public object ComponentProbe(string projectName, string componentName, string stage, string propertyName)
        {
            dynamic component = GetComponent(GetProject(projectName), componentName);
            switch (stage)
            {
                case "identity":
                    return new { Name = (string)component.Name, Type = (int)component.Type };
                case "descriptor_names":
                    return TypeDescriptor.GetProperties((object)component)
                        .Cast<PropertyDescriptor>().Select(item => new { item.Name,
                            Type = item.PropertyType?.FullName, item.IsReadOnly }).ToArray();
                case "designer_property_names":
                    var names = new List<string>();
                    foreach (dynamic property in component.Properties) names.Add((string)property.Name);
                    return names;
                case "code_count":
                    return new { Lines = (int)component.CodeModule.CountOfLines };
                case "code_sha":
                    dynamic module = component.CodeModule;
                    int lines = (int)module.CountOfLines;
                    return new { Lines = lines,
                        Sha256 = Hash(lines == 0 ? "" : (string)module.Lines(1, lines)) };
                case "descriptor_value":
                    PropertyDescriptor descriptor = TypeDescriptor.GetProperties((object)component).Find(propertyName, true);
                    if (descriptor == null) throw new ArgumentException("Descriptor not found: " + propertyName);
                    object value = descriptor.GetValue((object)component);
                    return new { Name = descriptor.Name, Type = descriptor.PropertyType?.FullName,
                        Kind = value != null && Marshal.IsComObject(value) ? "object" : "scalar",
                        Value = value != null && Marshal.IsComObject(value) ? null : Scalar(value) };
                case "designer_property_value":
                    if ((int)component.Type == 100 &&
                        string.Equals(propertyName, "MailEnvelope", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("MailEnvelope blocks the Excel document-component COM inspector and is not read automatically.");
                    dynamic selectedProperty = component.Properties.Item(propertyName);
                    object designerValue = selectedProperty.Value;
                    return new { Name = (string)selectedProperty.Name,
                        Kind = designerValue != null && Marshal.IsComObject(designerValue) ? "object" : "scalar",
                        Value = designerValue != null && Marshal.IsComObject(designerValue) ? null : Scalar(designerValue) };
                default:
                    throw new ArgumentException("Unsupported component probe Action.");
            }
        }

        /// <summary>Construit un instantané de propriétés, code, formulaire et empreinte pour un composant.</summary>
        /// <param name="projectName">Nom du projet contenant le composant.</param>
        /// <param name="component">Composant VBIDE déjà résolu.</param>
        /// <returns>État sérialisable et version calculée du composant.</returns>
        private object ComponentSnapshot(string projectName, dynamic component)
        {
            string name = (string)component.Name;
            int type = (int)component.Type;
            var properties = ReadProperties((object)component);
            var designerProperties = new List<VbePropertyInfo>();
            var hostProperties = new List<VbePropertyInfo>();
            foreach (dynamic property in component.Properties)
            {
                var info = new VbePropertyInfo { Name = (string)property.Name };
                if (type == 100)
                {
                    info.Kind = "host property";
                    info.Display = "Value not read in bulk; use component_property_value for a named property.";
                    if (string.Equals(info.Name, "MailEnvelope", StringComparison.OrdinalIgnoreCase))
                        info.Error = "Getter blocks Excel document-component inspection.";
                    hostProperties.Add(info);
                    continue;
                }
                try
                {
                    object value = property.Value;
                    info.Type = value?.GetType().FullName;
                    info.Kind = value != null && Marshal.IsComObject(value) ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = Scalar(value);
                }
                catch (Exception ex) { info.Error = ex.Message; }
                designerProperties.Add(info);
            }
            dynamic codeModule = component.CodeModule;
            int lineCount = (int)codeModule.CountOfLines;
            string code = lineCount == 0 ? "" : (string)codeModule.Lines(1, lineCount);
            string codeHash = Hash(code);
            string formVersion = null;
            if (type == 3)
            {
                dynamic state = forms.State(projectName, name);
                formVersion = (string)state.Version;
            }
            string version = Hash(json.Serialize(new { Name = name, Type = type,
                Properties = properties, DesignerProperties = designerProperties,
                HostProperties = hostProperties,
                CodeSha256 = codeHash, FormVersion = formVersion }));
            return new { Project = projectName, Component = name, Type = type,
                Version = version, CodeSha256 = codeHash, CodeLines = lineCount,
                FormVersion = formVersion, Properties = properties,
                DesignerProperties = designerProperties, HostProperties = hostProperties };
        }

        /// <summary>Modifie une propriété scalaire exposée du projet en mode conception et retourne son nouvel instantané.</summary>
        /// <param name="request">Requête avec propriété, valeur et version de projet attendue.</param>
        /// <returns>Instantané actualisé du projet.</returns>
        /// <exception cref="InvalidOperationException">Le renommage du projet est désactivé ou le projet a changé depuis sa lecture.</exception>
        /// <exception cref="ArgumentException">La valeur ne peut pas être convertie dans le type scalaire attendu.</exception>
        public object SetProjectProperty(Request request)
        {
            if (string.Equals(request.Property, "Name", StringComparison.OrdinalIgnoreCase)) return RenameSavedExcelProject(request);
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            SetScalar((object)project, request.Property, request.Value);
            return ProjectProperties(request.Project);
        }

        /// <summary>Modifie une propriété scalaire du composant après vérification de sa version.</summary>
        /// <param name="request">Requête avec projet, composant, propriété, valeur et version attendue.</param>
        /// <returns>Instantané actualisé du composant.</returns>
        /// <exception cref="ArgumentException">La valeur ou l’identifiant de composant est invalide.</exception>
        /// <exception cref="InvalidOperationException">Le projet n’est pas en mode conception ou le composant a changé.</exception>
        public object SetComponentProperty(Request request)
        {
            dynamic project = GetDesignProject(request.Project);
            dynamic component = GetComponent(project, request.Module);
            AssertComponentVersion(request, component);
            if (string.Equals(request.Property, "Name", StringComparison.OrdinalIgnoreCase))
                ValidateIdentifier(request.Value as string);
            SetScalar((object)component, request.Property, request.Value);
            return ComponentSnapshot(request.Project, component);
        }

        /// <summary>Définit Instancing d’un module de classe à Private (1) ou PublicNotCreatable (2).</summary>
        /// <param name="request">Requête avec la classe, sa version attendue et la valeur numérique 1 ou 2.</param>
        /// <returns>Valeur précédente, valeur retenue et instantané actualisé du composant.</returns>
        /// <exception cref="ArgumentException">La valeur n’est pas exactement 1 ou 2.</exception>
        /// <exception cref="InvalidOperationException">Le composant n’est pas une classe, n’est pas en mode conception ou n’a pas conservé la valeur.</exception>
        public object SetClassInstancing(Request request)
        {
            if (!(request.Value is int) && !(request.Value is long) && !(request.Value is double) &&
                !(request.Value is decimal))
                throw new ArgumentException("Value must be the number 1 (Private) or 2 (PublicNotCreatable).");
            int requested;
            try { requested = Convert.ToInt32(request.Value, CultureInfo.InvariantCulture); }
            catch (Exception ex) { throw new ArgumentException("Value must be 1 or 2.", ex); }
            if ((requested != 1 && requested != 2) ||
                Convert.ToDecimal(request.Value, CultureInfo.InvariantCulture) != requested)
                throw new ArgumentException("Value must be 1 (Private) or 2 (PublicNotCreatable).");
            dynamic project = GetDesignProject(request.Project);
            dynamic component = GetComponent(project, request.Module);
            if ((int)component.Type != 2) throw new InvalidOperationException("The component must be a class module.");
            AssertComponentVersion(request, component);
            dynamic property = component.Properties.Item("Instancing");
            int before = Convert.ToInt32(property.Value, CultureInfo.InvariantCulture);
            if (before != requested) property.Value = requested;
            int actual = Convert.ToInt32(property.Value, CultureInfo.InvariantCulture);
            if (actual != requested) throw new InvalidOperationException("The VBE did not retain class Instancing.");
            return new { Project = request.Project, Class = request.Module, Before = before,
                Instancing = actual, Meaning = actual == 1 ? "Private" : "PublicNotCreatable",
                Component = ComponentSnapshot(request.Project, component) };
        }

        /// <summary>Renomme un composant de projet après validation de l’identifiant et de la version attendue.</summary>
        /// <param name="request">Requête avec le nouveau nom et la version du composant.</param>
        /// <returns>Instantané du composant renommé.</returns>
        /// <exception cref="ArgumentException">Le nom demandé ne respecte pas le format d’identifiant VBA.</exception>
        /// <exception cref="InvalidOperationException">Le nom existe déjà, le projet a changé ou le VBE ne retient pas le nouveau nom.</exception>
        public object RenameComponent(Request request)
        {
            ValidateIdentifier(request.NewName);
            dynamic project = GetDesignProject(request.Project);
            foreach (dynamic existing in project.VBComponents)
                if (!string.Equals((string)existing.Name, request.Module, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals((string)existing.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A component with this name already exists.");
            dynamic component = GetComponent(project, request.Module);
            AssertComponentVersion(request, component);
            component.Name = request.NewName;
            if (!string.Equals((string)component.Name, request.NewName, StringComparison.Ordinal))
                throw new InvalidOperationException("The VBE did not retain the requested component name.");
            return ComponentSnapshot(request.Project, component);
        }

        /// <summary>Supprime un composant non protégé après validation des versions du projet et du composant.</summary>
        /// <param name="request">Requête contenant le nom du module et les deux versions attendues.</param>
        /// <returns>Instantané du projet après suppression.</returns>
        /// <exception cref="InvalidOperationException">Le projet ou composant a changé, ou le composant est un module de document hôte.</exception>
        public object RemoveComponent(Request request)
        {
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            dynamic component = GetComponent(project, request.Module);
            AssertComponentVersion(request, component);
            if ((int)component.Type == 100)
                throw new InvalidOperationException("Host document modules cannot be removed from their project.");
            project.VBComponents.Remove(component);
            return ProjectProperties(request.Project);
        }

        /// <summary>Importe un fichier de composant et vérifie l’ajout avant de retourner son état.</summary>
        /// <param name="request">Requête avec le chemin absolu et la version attendue du projet.</param>
        /// <returns>État du composant et du projet, avec indication si les lectures de confirmation restent en attente.</returns>
        /// <exception cref="InvalidOperationException">Le projet a changé ou le résultat de l’import ne peut être confirmé sans ambiguïté.</exception>
        public object ImportComponent(Request request)
        {
            string path = RequireExistingPath(request.Path);
            dynamic project = GetDesignProject(request.Project);
            AssertProjectVersion(request, project);
            var before = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic existing in project.VBComponents) before.Add((string)existing.Name);

            Exception importError = null;
            try { project.VBComponents.Import(path); }
            catch (Exception ex) { importError = ex; }

            // A COM error can occur after the component was added. Never call Import
            // again merely because its result or the immediate readback failed.
            var added = new List<string>();
            try
            {
                foreach (dynamic existing in project.VBComponents)
                {
                    string name = (string)existing.Name;
                    if (!before.Contains(name)) added.Add(name);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Import outcome is uncertain after COM error. " +
                    "Read list_modules before retrying; automatic re-import is disabled. " + ex.Message, ex);
            }
            if (added.Count != 1)
            {
                string cause = importError == null ? "" : " COM error: " + importError.Message;
                throw new InvalidOperationException("Import outcome is uncertain (new components: " +
                    added.Count + "). Read list_modules before retrying; automatic re-import is disabled." + cause);
            }

            string importedName = added[0];
            object importedState = TryImmediateRead(() => ComponentSnapshot(request.Project,
                GetComponent(project, importedName)), out string componentError);
            object projectState = TryImmediateRead(() => ProjectProperties(request.Project),
                out string projectError);
            bool verified = importedState != null && projectState != null;
            return new { Applied = true, Verified = verified, VerificationPending = !verified,
                ImportedName = importedName,
                Imported = importedState, Project = projectState,
                ImportError = importError?.Message, ComponentReadbackError = componentError,
                ProjectReadbackError = projectError,
                NextRead = verified ? null : "Call component_properties and project_properties in a separate request before another mutation." };
        }

        /// <summary>Tente une lecture immédiate et retourne son erreur sous forme de texte sans propager l’exception.</summary>
        /// <param name="read">Lecture à exécuter.</param>
        /// <param name="error">Reçoit le message de l’exception, ou nul en cas de succès.</param>
        /// <returns>Valeur lue, ou nul si la lecture échoue.</returns>
        private static object TryImmediateRead(Func<object> read, out string error)
        {
            error = null;
            try { return read(); }
            catch (Exception ex) { error = ex.Message; return null; }
        }

        /// <summary>Exporte un composant vers un nouveau chemin absolu après vérification de sa version.</summary>
        /// <param name="request">Requête contenant le projet, module, chemin et version attendue.</param>
        /// <returns>Chemin, taille du fichier créé et état du composant.</returns>
        /// <exception cref="IOException">La destination ou le fichier compagnon FRX existe déjà, ou le VBE ne crée pas l’export.</exception>
        /// <exception cref="InvalidOperationException">La version du composant a changé.</exception>
        public object ExportComponent(Request request)
        {
            string path = RequireAbsolutePath(request.Path);
            if (File.Exists(path)) throw new IOException("Export destination already exists: " + path);
            dynamic project = GetDesignProject(request.Project);
            dynamic component = GetComponent(project, request.Module);
            AssertComponentVersion(request, component);
            if ((int)component.Type == 3 &&
                string.Equals(Path.GetExtension(path), ".frm", StringComparison.OrdinalIgnoreCase) &&
                File.Exists(Path.ChangeExtension(path, ".frx")))
                throw new IOException("The UserForm FRX companion destination already exists.");
            component.Export(path);
            if (!File.Exists(path)) throw new IOException("The VBE did not create the export file.");
            return new { Project = request.Project, Component = request.Module,
                Path = path, Bytes = new FileInfo(path).Length,
                ComponentState = ComponentSnapshot(request.Project, component) };
        }

        /// <summary>Résout un projet dans l’instance VBE configurée.</summary>
        /// <param name="name">Nom du projet recherché.</param>
        /// <returns>Projet VBIDE correspondant.</returns>
        private dynamic GetProject(string name)
        {
            return VbeProjectResolver.Resolve(vbe, name);
        }

        /// <summary>Résout un projet et exige qu’il soit en mode conception.</summary>
        /// <param name="name">Nom du projet recherché.</param>
        /// <returns>Projet VBIDE en mode conception.</returns>
        /// <exception cref="InvalidOperationException">Le projet n’est pas en mode conception.</exception>
        private dynamic GetDesignProject(string name)
        {
            dynamic project = GetProject(name);
            if ((int)project.Mode != 2) throw new InvalidOperationException("The project must be in design mode.");
            return project;
        }

        /// <summary>Recherche un composant par nom sans tenir compte de la casse et refuse les noms ambigus.</summary>
        /// <param name="project">Projet dans lequel effectuer la recherche.</param>
        /// <param name="name">Nom du composant demandé.</param>
        /// <returns>Composant correspondant.</returns>
        /// <exception cref="ArgumentException">Le nom de composant est vide.</exception>
        /// <exception cref="InvalidOperationException">Le nom est ambigu ou aucun composant ne correspond.</exception>
        private static dynamic GetComponent(dynamic project, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Component is required.");
            dynamic match = null;
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if (match != null) throw new InvalidOperationException("Component name is ambiguous.");
                    match = component;
                }
            if (match == null) throw new InvalidOperationException("Component not found: " + name);
            return match;
        }

        /// <summary>Vérifie que l’empreinte actuelle du projet correspond à celle fournie par la requête.</summary>
        /// <param name="request">Requête portant l’empreinte attendue.</param>
        /// <param name="project">Projet déjà résolu.</param>
        /// <exception cref="ArgumentException">L’empreinte attendue est absente.</exception>
        /// <exception cref="InvalidOperationException">Le projet a changé depuis sa lecture.</exception>
        private void AssertProjectVersion(Request request, dynamic project)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("ExpectedProjectVersion is required.");
            dynamic state = ProjectProperties(request.Project);
            if (!string.Equals((string)state.Version, request.ExpectedProjectVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project changed since it was read.");
        }

        /// <summary>Vérifie que l’empreinte actuelle du composant correspond à celle fournie par la requête.</summary>
        /// <param name="request">Requête portant l’empreinte attendue et le projet source.</param>
        /// <param name="component">Composant déjà résolu.</param>
        /// <exception cref="ArgumentException">L’empreinte attendue est absente.</exception>
        /// <exception cref="InvalidOperationException">Le composant a changé depuis sa lecture.</exception>
        private void AssertComponentVersion(Request request, dynamic component)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedComponentVersion))
                throw new ArgumentException("ExpectedComponentVersion is required.");
            dynamic state = ComponentSnapshot(request.Project, component);
            if (!string.Equals((string)state.Version, request.ExpectedComponentVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The component changed since it was read.");
        }

        /// <summary>Convertit et définit une propriété scalaire modifiable puis vérifie la valeur relue.</summary>
        /// <param name="target">Objet dont la propriété sera modifiée.</param>
        /// <param name="name">Nom de la propriété exposée.</param>
        /// <param name="value">Valeur à convertir dans le type de propriété.</param>
        /// <exception cref="ArgumentException">Le nom ou la valeur est absent, ou une chaîne est exigée.</exception>
        /// <exception cref="InvalidOperationException">La propriété manque, est en lecture seule, n’est pas scalaire ou n’a pas retenu la valeur.</exception>
        private static void SetScalar(object target, string name, object value)
        {
            if (string.IsNullOrWhiteSpace(name) || value == null)
                throw new ArgumentException("Property and non-null Value are required.");
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(target).Find(name, true);
            if (descriptor == null) throw new InvalidOperationException("Property is not exposed: " + name);
            if (descriptor.IsReadOnly) throw new InvalidOperationException("Property is read-only: " + name);
            Type type = descriptor.PropertyType;
            if (type == typeof(object)) type = descriptor.GetValue(target)?.GetType() ?? value.GetType();
            object converted;
            if (type == typeof(string))
            {
                if (!(value is string)) throw new ArgumentException("A string is required.");
                converted = value;
            }
            else if (type.IsEnum)
                converted = value is string ? Enum.Parse(type, (string)value, true) :
                    Enum.ToObject(type, Convert.ToInt32(value, CultureInfo.InvariantCulture));
            else if (type.IsPrimitive || type == typeof(decimal))
                converted = Convert.ChangeType(value, type, CultureInfo.InvariantCulture);
            else throw new InvalidOperationException("Property is not an editable scalar: " + name);
            descriptor.SetValue(target, converted);
            object actual = descriptor.GetValue(target);
            if (!Equals(actual, converted) &&
                !string.Equals(Convert.ToString(actual, CultureInfo.InvariantCulture),
                    Convert.ToString(converted, CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The VBE did not retain property " + name + ".");
        }

        /// <summary>Énumère les propriétés exposées sans invoquer les accesseurs des types non scalaires.</summary>
        /// <param name="target">Objet VBIDE à inspecter.</param>
        /// <returns>Métadonnées et valeurs sécurisées des propriétés lisibles.</returns>
        private static List<VbePropertyInfo> ReadProperties(object target)
        {
            var result = new List<VbePropertyInfo>();
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(target))
            {
                var info = new VbePropertyInfo { Name = descriptor.Name,
                    Type = descriptor.PropertyType?.FullName, ReadOnly = descriptor.IsReadOnly };
                if (!IsSafeScalarType(descriptor.PropertyType))
                {
                    info.Kind = "object";
                    info.Display = "Object getter not invoked; use a dedicated VBIDE inspection command.";
                    result.Add(info);
                    continue;
                }
                try
                {
                    object value = descriptor.GetValue(target);
                    info.Kind = value != null && Marshal.IsComObject(value) ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = Scalar(value);
                }
                catch (Exception ex) { info.Error = ex.Message; }
                result.Add(info);
            }
            return result;
        }

        /// <summary>Indique si un type peut être lu comme valeur scalaire sans appeler un accesseur d’objet COM.</summary>
        /// <param name="type">Type de propriété examiné.</param>
        /// <returns><see langword="true"/> pour les primitifs, énumérations, chaînes, décimaux et dates.</returns>
        private static bool IsSafeScalarType(Type type)
        {
            return type != null && (type.IsPrimitive || type.IsEnum ||
                type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime));
        }

        /// <summary>Normalise une valeur connue en scalaire directement sérialisable.</summary>
        /// <param name="value">Valeur à convertir.</param>
        /// <returns>Valeur scalaire d’origine, ou représentation invariant-culture pour les autres types.</returns>
        private static object Scalar(object value)
        {
            if (value == null || value is string || value is bool || value is byte ||
                value is short || value is int || value is long || value is float ||
                value is double || value is decimal) return value;
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        /// <summary>Valide un chemin Windows absolu et retourne sa forme complète normalisée.</summary>
        /// <param name="path">Chemin à valider.</param>
        /// <returns>Chemin complet.</returns>
        /// <exception cref="ArgumentException">Le chemin n’est pas un chemin Windows absolu.</exception>
        private static string RequireAbsolutePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) ||
                !Regex.IsMatch(path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("A fully qualified file path is required.");
            return Path.GetFullPath(path);
        }

        /// <summary>Valide un chemin absolu et exige que le fichier existe.</summary>
        /// <param name="path">Chemin de fichier à valider.</param>
        /// <returns>Chemin complet existant.</returns>
        /// <exception cref="ArgumentException">Le chemin n’est pas absolu.</exception>
        /// <exception cref="FileNotFoundException">Le fichier n’existe pas.</exception>
        private static string RequireExistingPath(string path)
        {
            string fullPath = RequireAbsolutePath(path);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("Import file not found.", fullPath);
            return fullPath;
        }

        /// <summary>Valide un identifiant VBA commençant par une lettre et limité à quarante caractères.</summary>
        /// <param name="name">Identifiant à vérifier.</param>
        /// <exception cref="ArgumentException">Le nom ne respecte pas le format autorisé.</exception>
        private static void ValidateIdentifier(string name)
        {
            if (string.IsNullOrWhiteSpace(name) ||
                !Regex.IsMatch(name, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                throw new ArgumentException("Name must start with a letter and contain at most 40 letters, digits or underscores.");
        }

        /// <summary>Calcule le SHA-256 hexadécimal minuscule d’une chaîne UTF-8.</summary>
        /// <param name="value">Texte à hacher.</param>
        /// <returns>Empreinte hexadécimale.</returns>
        private static string Hash(string value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value)))
                    .Replace("-", "").ToLowerInvariant();
        }
    }
}
