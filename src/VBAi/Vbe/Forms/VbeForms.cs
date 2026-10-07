using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Expose les opérations d’inspection et de modification des UserForms du VBE.</summary>
    internal sealed partial class VbeForms
    {

        /// <summary>Instance VBE utilisée pour résoudre les projets à inspecter.</summary>
        private readonly dynamic vbe;

        /// <summary>Fonction injectable qui calcule l’empreinte de la sérialisation de l’arbre.</summary>
        internal static Func<byte[], byte[]> HashTree = bytes =>
        {
            using (var sha = SHA256.Create()) return sha.ComputeHash(bytes);
        };

        /// <summary>ProgIDs des contrôles MSForms pris en charge par les opérations de création.</summary>
        private static readonly HashSet<string> BuiltInControls = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Forms.CheckBox.1", "Forms.ComboBox.1", "Forms.CommandButton.1", "Forms.Frame.1",
            "Forms.Image.1", "Forms.Label.1", "Forms.ListBox.1", "Forms.MultiPage.1",
            "Forms.OptionButton.1", "Forms.ScrollBar.1", "Forms.SpinButton.1",
            "Forms.TabStrip.1", "Forms.TextBox.1", "Forms.ToggleButton.1"
        };

        /// <summary>Crée le service pour l’instance VBE fournie.</summary>
        /// <param name="vbe">Instance VBE à utiliser pour résoudre les projets.</param>
        public VbeForms(object vbe) { this.vbe = vbe; }

        /// <summary>Retourne le catalogue des types de contrôles MSForms disponibles.</summary>
        /// <returns>Catalogue des ProgIDs MSForms pris en charge.</returns>
        public object ControlTypes() { return VbeControlCatalog.List(BuiltInControls); }

        /// <summary>Énumère les UserForms du projet et indique si leur concepteur est ouvert.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <returns>Formulaires du projet avec leur état de concepteur.</returns>
        public object List(string projectName)
        {
            dynamic project = GetProject(projectName);
            var result = new List<object>();
            foreach (dynamic component in project.VBComponents)
                if ((int)component.Type == 3)
                    result.Add(new { Name = (string)component.Name, DesignerOpen = (bool)component.HasOpenDesigner });
            return result;
        }

        /// <summary>Retourne l’état synthétique du UserForm et son empreinte de version.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <returns>Instantané de l’état du formulaire.</returns>
        public object State(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            return Snapshot(projectName, form);
        }

        /// <summary>Décrit les propriétés du UserForm sélectionné.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <returns>Propriétés du formulaire décrites par le service.</returns>
        public object Properties(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            return DescribeProperties(form);
        }

        /// <summary>Retourne l’arbre hiérarchique des contrôles et sa version.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <returns>Arbre du formulaire, propriétés, chemins canoniques et empreinte.</returns>
        public object Tree(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            string version = TreeVersion(form, out List<object> nodes, out List<VbePropertyInfo> properties, out int nodeCount);
            return new
            {
                Project = projectName,
                Form = formName,
                FormVersion = version,
                TreeVersion = version,
                NodeCount = nodeCount,
                Properties = properties,
                Controls = nodes
            };
        }

        /// <summary>Retourne les événements COM disponibles pour le UserForm ou le contrôle canonique indiqué.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <param name="controlPath">Chemin hiérarchique canonique du contrôle dans form_tree.</param>
        /// <returns>Événements COM disponibles pour la cible choisie.</returns>
        public object EventCatalog(string projectName, string formName, string controlPath)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            dynamic tree = Tree(projectName, formName);
            bool userForm = string.IsNullOrWhiteSpace(controlPath) ||
                string.Equals(controlPath, "UserForm", StringComparison.OrdinalIgnoreCase);
            if (!userForm && !TreeContainsPath((IEnumerable)tree.Controls, controlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            object target = userForm ? (object)form.Designer : ResolveTreeItem(form.Designer, controlPath);
            object catalog = VbeComEvents.Read(target);
            return new
            {
                Project = projectName,
                Form = formName,
                ControlPath = userForm ? "UserForm" : controlPath,
                ObjectName = userForm ? "UserForm" : (string)((dynamic)target).Name,
                Type = TypeDescriptor.GetClassName(target),
                TreeVersion = (string)tree.TreeVersion,
                Catalog = catalog
            };
        }

        /// <summary>Construit l’arbre du concepteur et calcule sa version à partir de son contenu.</summary>
        /// <param name="form">Formulaire VBE inspecté.</param>
        /// <param name="nodes">Nœuds parcourus dans l’arbre.</param>
        /// <param name="properties">Liste des propriétés décrites pour le formulaire.</param>
        /// <param name="nodeCount">Compteur alimenté avec le nombre de nœuds parcourus.</param>
        /// <returns>Empreinte de l’arbre et de ses propriétés.</returns>
        private static string TreeVersion(dynamic form, out List<object> nodes,
            out List<VbePropertyInfo> properties, out int nodeCount)
        {
            nodeCount = 0;
            object designer = form.Designer;
            nodes = ReadChildControls(((dynamic)designer).Controls, designer, (string)form.Name,
                "Controls", 0, ref nodeCount);
            properties = DescribeProperties(form);
            string json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }
                .Serialize(new { Form = (string)form.Name, Properties = properties, Controls = nodes });
            return BitConverter.ToString(HashTree(Encoding.UTF8.GetBytes(json)))
                .Replace("-", "").ToLowerInvariant();
        }

        /// <summary>Retourne la sonde de parent de contrôle prise en charge par le conteneur.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <returns>Sonde du parent du formulaire ou du contrôle, si l’interface COM l’expose.</returns>
        public object ParentProbe(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            object designer = form.Designer;
            var rows = new List<object>();
            foreach (dynamic control in form.Designer.Controls)
            {
                object parent = control.Parent;
                rows.Add(new
                {
                    Control = (string)control.Name,
                    ParentName = SafeComName(parent),
                    ParentType = TypeDescriptor.GetClassName(parent),
                    DesignerName = SafeComName(designer),
                    DesignerType = TypeDescriptor.GetClassName(designer),
                    SameDesigner = SameComIdentity(parent, designer),
                    SameComponent = SameComIdentity(parent, (object)form)
                });
            }
            return new { Project = projectName, Form = formName, Rows = rows };
        }

        /// <summary>Lit le nom COM sans laisser une erreur d’accès empêcher l’inspection.</summary>
        /// <param name="item">Objet COM ou descripteur à inspecter.</param>
        /// <returns>Nom COM accessible, ou valeur de repli en cas d’erreur.</returns>
        private static string SafeComName(object item)
        {
            try { return (string)((dynamic)item).Name; }
            catch { return null; }
        }

        /// <summary>Parcourt les contrôles enfants et construit leurs descriptions hiérarchiques.</summary>
        /// <param name="collection">Collection COM des contrôles à parcourir.</param>
        /// <param name="owner">Objet parent attendu des contrôles de la collection.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <param name="path">Chemin à rechercher dans les nœuds.</param>
        /// <param name="depth">Profondeur courante de l’arbre.</param>
        /// <param name="nodeCount">Compteur alimenté avec le nombre de nœuds parcourus.</param>
        /// <returns>Nœuds des contrôles enfants et de leurs descendants.</returns>
        private static List<object> ReadChildControls(dynamic collection, object owner, string formName,
            string path, int depth, ref int nodeCount)
        {
            var result = new List<object>();
            foreach (dynamic control in collection)
                if (SameContainer((object)control.Parent, owner, formName))
                    result.Add(ReadTreeNode(control, "Control", formName, path, depth, ref nodeCount));
            return result;
        }

        /// <summary>Compare deux conteneurs de contrôle à partir de leur identité COM.</summary>
        /// <param name="actualParent">Parent observé du contrôle.</param>
        /// <param name="expectedOwner">Conteneur attendu selon l’arbre du formulaire.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <returns>true si les objets appartiennent au même conteneur.</returns>
        private static bool SameContainer(object actualParent, object expectedOwner, string formName)
        {
            if (SameComIdentity(actualParent, expectedOwner)) return true;
            string actualPath = ContainerIdentity(actualParent, formName);
            string expectedPath = ContainerIdentity(expectedOwner, formName);
            return actualPath != null && string.Equals(actualPath, expectedPath, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Construit une identité stable du conteneur à partir de ses éléments accessibles.</summary>
        /// <param name="item">Objet COM ou descripteur à inspecter.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <returns>Identité calculée du conteneur.</returns>
        private static string ContainerIdentity(object item, string formName)
        {
            if (item == null) return null;
            var path = new List<string>();
            for (int depth = 0; depth < 16 && item != null; depth++)
            {
                string type = TypeDescriptor.GetClassName(item);
                string name = SafeComName(item);
                if (string.Equals(type, "UserForm", StringComparison.OrdinalIgnoreCase) &&
                    string.IsNullOrEmpty(name)) name = formName;
                if (string.IsNullOrEmpty(type) || string.IsNullOrEmpty(name)) return null;
                path.Add(type + ":" + name);
                if (string.Equals(type, "UserForm", StringComparison.OrdinalIgnoreCase))
                    return string.Join("/", path);
                try { item = ((dynamic)item).Parent; }
                catch { return null; }
            }
            return null;
        }

        /// <summary>Compare deux objets COM en utilisant leur identité IUnknown.</summary>
        /// <param name="left">Premier objet COM à comparer.</param>
        /// <param name="right">Second objet COM à comparer.</param>
        /// <returns>true si les objets COM désignent la même identité.</returns>
        private static bool SameComIdentity(object left, object right)
        {
            if (left == null || right == null) return false;
            if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right))
                return ReferenceEquals(left, right);
            IntPtr leftIdentity = IntPtr.Zero;
            IntPtr rightIdentity = IntPtr.Zero;
            try
            {
                leftIdentity = Marshal.GetIUnknownForObject(left);
                rightIdentity = Marshal.GetIUnknownForObject(right);
                return leftIdentity == rightIdentity;
            }
            finally
            {
                if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
                if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
            }
        }

        /// <summary>Lit un nœud de contrôle, ses propriétés et ses enfants.</summary>
        /// <param name="item">Objet COM ou descripteur à inspecter.</param>
        /// <param name="kind">Type de nœud dans l’arbre.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <param name="parentPath">Chemin du parent dans l’arbre.</param>
        /// <param name="depth">Profondeur courante de l’arbre.</param>
        /// <param name="nodeCount">Compteur alimenté avec le nombre de nœuds parcourus.</param>
        /// <returns>Description du nœud et de ses propriétés enfant.</returns>
        private static object ReadTreeNode(object item, string kind, string formName,
            string parentPath, int depth, ref int nodeCount)
        {
            if (++nodeCount > 512 || depth > 16)
                throw new InvalidOperationException("The UserForm control hierarchy exceeds the inspection limit.");
            dynamic current = item;
            string name = (string)current.Name;
            string path = parentPath + "/" + name;
            var children = new List<object>();
            PropertyDescriptorCollection descriptors = TypeDescriptor.GetProperties(item);
            if (kind == "Control")
            {
                PropertyDescriptor controls = descriptors.Find("Controls", true);
                if (controls != null)
                {
                    object nested = controls.GetValue(item);
                    if (nested != null)
                        children.AddRange(ReadChildControls(nested, item, formName,
                            path + "/Controls", depth + 1, ref nodeCount));
                }
                PropertyDescriptor pages = descriptors.Find("Pages", true);
                if (pages != null)
                    foreach (dynamic page in (dynamic)pages.GetValue(item))
                        children.Add(ReadTreeNode(page, "Page", formName, path + "/Pages", depth + 1, ref nodeCount));
                PropertyDescriptor tabs = descriptors.Find("Tabs", true);
                if (tabs != null)
                    foreach (dynamic tab in (dynamic)tabs.GetValue(item))
                        children.Add(ReadTreeNode(tab, "Tab", formName, path + "/Tabs", depth + 1, ref nodeCount));
            }
            else if (kind == "Page")
            {
                PropertyDescriptor controls = descriptors.Find("Controls", true);
                if (controls != null)
                    children.AddRange(ReadChildControls(controls.GetValue(item), item, formName,
                        path + "/Controls", depth + 1, ref nodeCount));
            }
            return new
            {
                Path = path,
                Name = name,
                Kind = kind,
                Type = TypeDescriptor.GetClassName(item),
                Properties = ReadObjectProperties(item),
                Children = children
            };
        }

        /// <summary>Lit les propriétés descriptibles d’un objet du concepteur.</summary>
        /// <param name="item">Objet COM ou descripteur à inspecter.</param>
        /// <returns>Propriétés lisibles de l’objet.</returns>
        private static List<VbePropertyInfo> ReadObjectProperties(object item)
        {
            var result = new List<VbePropertyInfo>();
            string targetType = TypeDescriptor.GetClassName(item);
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(item))
            {
                var info = new VbePropertyInfo
                {
                    Name = descriptor.Name,
                    Type = descriptor.PropertyType?.FullName,
                    ReadOnly = descriptor.IsReadOnly,
                    AllowedValues = EnumChoices(descriptor.PropertyType),
                    SetterStatus = DesignerSetterStatus(targetType, descriptor)
                };
                if (TryDescribeReservedFontWriteOnly(item, descriptor, info))
                { result.Add(info); continue; }
                try
                {
                    object value = descriptor.GetValue(item);
                    info.Kind = value != null && (Marshal.IsComObject(value) || value is Font || value is Image)
                        ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = NormalizeScalar(value);
                    else if (value is Image img)
                    {
                        info.Digest = ImageDigest(img);
                    }
                    else if (value is Font)
                    {
                        try { info.Members = DescribeObjectMembers((object)((dynamic)item).Font); }
                        catch (Exception ex) { info.Error = ex.Message; }
                    }
                }
                catch (Exception ex) { info.Error = ex.Message; }
                result.Add(info);
            }
            return result;
        }

        /// <summary>Valide et affecte une propriété de formulaire puis retourne son état actualisé.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>État mis à jour après la modification de la propriété.</returns>
        public object SetProperty(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Property) || request.Value == null)
                throw new ArgumentException("Property and a non-null Value are required.");
            dynamic project = GetDesignProject(request.Project);
            dynamic form = GetForm(project, request.Form);
            AssertVersion(request, form);
            string[] path = request.Property.Split('.');
            if (path.Length < 1 || path.Length > 2 || path.Any(part => string.IsNullOrWhiteSpace(part)))
                throw new ArgumentException("Use a form property name or one member path such as Font.Name.");
            string propertyName = path[0];
            if (path.Length == 1 && string.Equals(propertyName, "Name", StringComparison.OrdinalIgnoreCase))
            {
                string newName = request.Value as string;
                if (string.IsNullOrWhiteSpace(newName) || !Regex.IsMatch(newName, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                    throw new ArgumentException("Name must start with a letter and contain at most 40 letters, digits or underscores.");
                foreach (dynamic component in project.VBComponents)
                    if (!string.Equals((string)component.Name, (string)form.Name, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals((string)component.Name, newName, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("A component with this name already exists.");
                form.Name = newName;
            }
            else
            {
                dynamic property = null;
                foreach (dynamic candidate in form.Properties)
                    if (string.Equals((string)candidate.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                    { property = candidate; break; }
                if (property == null) throw new InvalidOperationException("Unknown UserForm property: " + propertyName);
                if ((int)property.NumIndices != 0)
                    throw new InvalidOperationException("Indexed properties need an indexed editor; this path cannot be written as a scalar.");
                if (path.Length == 1)
                {
                    PropertyDescriptor descriptor = TypeDescriptor.GetProperties((object)form.Designer).Find(propertyName, true);
                    if (descriptor != null && descriptor.IsReadOnly)
                        throw new InvalidOperationException("This UserForm property is read-only: " + propertyName);
                    object previous = null;
                    try { previous = property.Value; }
                    catch (Exception ex) { throw new InvalidOperationException("Cannot read the current value of " + propertyName + ": " + ex.Message); }
                    if (previous != null && Marshal.IsComObject(previous))
                        throw new InvalidOperationException("Object property: use a member path such as Font.Name or a dedicated object command.");
                    // VBIDE exposes Tag as an untyped null Variant until it is assigned.
                    Type targetType = previous?.GetType() ?? descriptor?.PropertyType;
                    if (targetType == null && string.Equals(propertyName, "Tag", StringComparison.OrdinalIgnoreCase))
                        targetType = typeof(string);
                    if (targetType == null || targetType == typeof(object))
                        throw new InvalidOperationException("Cannot determine a safe scalar type for " + propertyName);
                    property.Value = ConvertScalar(request.Value, targetType);
                }
                else
                {
                    if (!string.Equals(propertyName, "Font", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("This object property needs a dedicated editor; scalar member writes are unavailable.");
                    SetFormFontMember(form.Designer, path[1], request.Value);
                }
            }
            string currentName = (string)form.Name;
            var properties = DescribeProperties(form);
            return new
            {
                request.Project,
                Form = currentName,
                request.Property,
                Properties = properties,
                State = Snapshot(request.Project, form)
            };
        }

        /// <summary>Affecte et vérifie une img OLE sur le contrôle ciblé.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>État mis à jour après le remplacement de l’img.</returns>
        public object SetPicture(Request request)
        {
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            object picture = OlePictureLoader.Load(request.Path);
            dynamic designer = form.Designer;
            designer.Picture = picture;
            object installed = designer.Picture;
            string expected = OlePictureLoader.Fingerprint(picture);
            string actual = OlePictureLoader.Fingerprint(installed);
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
                throw new InvalidOperationException("The UserForm did not retain the requested OLE picture.");
            return new
            {
                request.Project,
                request.Form,
                Picture = actual,
                Properties = DescribeProperties(form),
                State = Snapshot(request.Project, form)
            };
        }

        /// <summary>Convertit une valeur JSON vers le type scalaire de la propriété COM.</summary>
        /// <param name="value">Valeur demandée ou lue pour la propriété.</param>
        /// <param name="targetType">Type .NET attendu pour la valeur.</param>
        /// <returns>Valeur convertie vers le type cible.</returns>
        private static object ConvertScalar(object value, Type targetType)
        {
            if (targetType == null || targetType == typeof(object) || targetType.IsArray ||
                (!targetType.IsPrimitive && targetType != typeof(string) && targetType != typeof(decimal) && !targetType.IsEnum))
                throw new InvalidOperationException("This property is not a supported scalar type.");
            if (targetType.IsEnum)
            {
                if (value is string v) return Enum.Parse(targetType, v, true);
                return Enum.ToObject(targetType, Convert.ToInt32(value, CultureInfo.InvariantCulture));
            }
            if (targetType == typeof(bool) && value is string v1)
            {
                return !bool.TryParse(v1, out bool parsed) ? throw new ArgumentException("Boolean value must be true or false.") : (object)parsed;
            }
            if (targetType == typeof(string))
            {
                if (!(value is string)) throw new ArgumentException("This property requires a string.");
                return value;
            }
            if (targetType == typeof(float) || targetType == typeof(double) || targetType == typeof(decimal))
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (double.IsNaN(number) || double.IsInfinity(number))
                    throw new ArgumentException("Numeric value must be finite.");
            }
            return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
        }

        /// <summary>Modifie un membre pris en charge de la police du UserForm.</summary>
        /// <param name="designer">Objet Designer qui contient le contrôle.</param>
        /// <param name="member">Membre de police à modifier.</param>
        /// <param name="value">Valeur demandée ou lue pour la propriété.</param>
        private static void SetFormFontMember(dynamic designer, string member, object value)
        {
            dynamic font = designer.Font;
            if (string.Equals(member, "Name", StringComparison.OrdinalIgnoreCase))
            {
                string name = value as string;
                if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Font.Name requires a nonempty string.");
                font.Name = name;
                if (!string.Equals((string)font.Name, name, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The VBE did not retain Font.Name.");
            }
            else if (string.Equals(member, "Size", StringComparison.OrdinalIgnoreCase))
            {
                double size = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (double.IsNaN(size) || double.IsInfinity(size) || size <= 0 || size > 200)
                    throw new ArgumentException("Font.Size must be greater than 0 and at most 200.");
                font.Size = size;
                if (Math.Abs(Convert.ToDouble(font.Size, CultureInfo.InvariantCulture) - size) > 0.01)
                    throw new InvalidOperationException("The VBE did not retain Font.Size.");
            }
            else
            {
                bool enabled = (bool)ConvertScalar(value, typeof(bool));
                if (string.Equals(member, "Bold", StringComparison.OrdinalIgnoreCase))
                { font.Bold = enabled; if ((bool)font.Bold != enabled) throw new InvalidOperationException("The VBE did not retain Font.Bold."); }
                else if (string.Equals(member, "Italic", StringComparison.OrdinalIgnoreCase))
                { font.Italic = enabled; if ((bool)font.Italic != enabled) throw new InvalidOperationException("The VBE did not retain Font.Italic."); }
                else if (string.Equals(member, "Underline", StringComparison.OrdinalIgnoreCase))
                { font.Underline = enabled; if ((bool)font.Underline != enabled) throw new InvalidOperationException("The VBE did not retain Font.Underline."); }
                else if (string.Equals(member, "Strikethrough", StringComparison.OrdinalIgnoreCase))
                { font.Strikethrough = enabled; if ((bool)font.Strikethrough != enabled) throw new InvalidOperationException("The VBE did not retain Font.Strikethrough."); }
                else throw new InvalidOperationException("Unsupported Font member: " + member);
            }
        }

        /// <summary>Produit les descriptions de propriétés lisibles du UserForm.</summary>
        /// <param name="form">Formulaire VBE inspecté.</param>
        /// <returns>Descriptions des propriétés du formulaire.</returns>
        private static List<VbePropertyInfo> DescribeProperties(dynamic form)
        {
            var result = new List<VbePropertyInfo>();
            object designer = form.Designer;
            PropertyDescriptorCollection descriptors = TypeDescriptor.GetProperties(designer);
            foreach (dynamic property in form.Properties)
            {
                string name = (string)property.Name;
                PropertyDescriptor descriptor = descriptors.Find(name, true);
                var info = new VbePropertyInfo
                {
                    Name = name,
                    Type = descriptor?.PropertyType?.FullName,
                    ReadOnly = descriptor == null ? (bool?)null : descriptor.IsReadOnly,
                    AllowedValues = EnumChoices(descriptor?.PropertyType),
                    SetterStatus = descriptor == null ? "Unknown" : DesignerSetterStatus("UserForm", descriptor)
                };
                if (string.Equals(name, "Picture", StringComparison.OrdinalIgnoreCase))
                {
                    info.Kind = "object";
                    info.Type = "stdole.IPictureDisp";
                    try
                    {
                        object installed = ((dynamic)designer).Picture;
                        info.Digest = OlePictureLoader.Fingerprint(installed);
                        info.Display = installed == null ? "(empty)" : info.Digest;
                    }
                    catch (Exception ex) { info.Error = ex.Message; }
                    result.Add(info);
                    continue;
                }
                if (TryDescribeReservedFontWriteOnly(designer, descriptor, info))
                { result.Add(info); continue; }
                try { info.NumIndices = (int)property.NumIndices; }
                catch (Exception ex) { info.Error = ex.Message; }
                object raw = null;
                try { raw = property.Value; }
                catch (Exception ex) { info.Error = ex.Message; }
                object managed = null;
                if (descriptor != null)
                {
                    try { managed = descriptor.GetValue(designer); }
                    catch (Exception ex) { if (info.Error == null) info.Error = ex.Message; }
                }
                object inspected = managed ?? raw;
                if (inspected != null && info.Type == null) info.Type = inspected.GetType().FullName;
                if (info.NumIndices > 0) info.Kind = "indexed";
                else if (inspected is Image img)
                {
                    info.Kind = "object";
                    info.Display = inspected.GetType().Name;
                    info.Digest = ImageDigest(img);
                    info.Members = DescribeObjectMembers(inspected);
                }
                else if (inspected != null &&
                    (Marshal.IsComObject(inspected) || inspected is Font || inspected is System.Collections.IEnumerable && !(inspected is string)))
                {
                    info.Kind = "object";
                    info.Display = inspected.GetType().Name;
                    info.Members = DescribeObjectMembers(inspected);
                }
                else
                {
                    info.Kind = "scalar";
                    info.Value = NormalizeScalar(raw ?? managed);
                    if (managed is Color clr) info.Display = clr.Name;
                }
                result.Add(info);
            }
            return result;
        }

        /// <summary>Décrit les membres accessibles d’un objet imbriqué du concepteur.</summary>
        /// <param name="source">Objet dont les membres sont décrits.</param>
        /// <returns>Descriptions des membres publics accessibles.</returns>
        private static List<VbePropertyInfo> DescribeObjectMembers(object source)
        {
            var members = new List<VbePropertyInfo>();
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(source))
            {
                if (members.Count >= 64) break;
                var info = new VbePropertyInfo
                {
                    Name = descriptor.Name,
                    Type = descriptor.PropertyType?.FullName,
                    ReadOnly = descriptor.IsReadOnly,
                    AllowedValues = EnumChoices(descriptor.PropertyType)
                };
                try
                {
                    object value = descriptor.GetValue(source);
                    info.Kind = value != null && Marshal.IsComObject(value) ? "object" : "scalar";
                    if (info.Kind == "scalar") info.Value = NormalizeScalar(value);
                }
                catch (Exception ex) { info.Error = ex.Message; }
                members.Add(info);
            }
            return members;
        }

        /// <summary>Retourne les noms symboliques possibles pour une propriété enum.</summary>
        /// <param name="type">Type de propriété à énumérer.</param>
        /// <returns>Noms symboliques des valeurs enum.</returns>
        private static string[] EnumChoices(Type type)
        {
            if (type == null || !type.IsEnum) return null;
            try { return Enum.GetNames(type); }
            catch { return null; }
        }

        /// <summary>Classe l’écriture d’une propriété selon son support observé par le concepteur.</summary>
        /// <param name="targetType">Type .NET attendu pour la valeur.</param>
        /// <param name="descriptor">Descripteur de propriété retourné par le concepteur.</param>
        /// <returns>Statut descriptif de prise en charge de l’écriture.</returns>
        private static string DesignerSetterStatus(string targetType, PropertyDescriptor descriptor)
        {
            if (descriptor.IsReadOnly) return "DescriptorReadOnly";
            if (string.Equals(descriptor.Name, "_Font_Reserved", StringComparison.OrdinalIgnoreCase))
                return "GetterUnavailable";
            if (string.Equals(targetType, "Label", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(descriptor.Name, "Cancel", StringComparison.OrdinalIgnoreCase))
                return "BlockedNativeSetterFailure";
            if (string.Equals(targetType, "ToggleButton", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(descriptor.Name, "Value", StringComparison.OrdinalIgnoreCase))
                return "BlockedAfterHostCrash";
            if (string.Equals(targetType, "TextBox", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(descriptor.Name, "ScrollBars", StringComparison.OrdinalIgnoreCase))
                return "BlockedAfterHostCrash";
            if (string.Equals(targetType, "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(descriptor.Name, "ColumnCount", StringComparison.OrdinalIgnoreCase))
                return "BlockedAfterHostCrash";
            if (string.Equals(targetType, "SpinButton", StringComparison.OrdinalIgnoreCase) &&
                new[] { "Min", "Max", "Value", "Delay", "SmallChange" }
                    .Any(name => string.Equals(descriptor.Name, name, StringComparison.OrdinalIgnoreCase)))
                return "BlockedAfterHostCrash";
            return "DescriptorCandidateUnverified";
        }

        /// <summary>Normalise une valeur retournée par COM en valeur sérialisable.</summary>
        /// <param name="value">Valeur demandée ou lue pour la propriété.</param>
        /// <returns>Valeur normalisée et sérialisable.</returns>
        private static object NormalizeScalar(object value)
        {
            if (value == null) return null;
            if (value is string || value is bool || value is byte || value is sbyte ||
                value is short || value is ushort || value is int || value is uint ||
                value is long || value is ulong || value is float || value is double || value is decimal)
                return value;
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        /// <summary>Calcule l’empreinte de l’img actuellement exposée par le contrôle.</summary>
        /// <param name="image">Valeur img à empreinter.</param>
        /// <returns>Empreinte de l’img lue.</returns>
        private static string ImageDigest(Image image)
        {
            try
            {
                using (var memory = new MemoryStream())
                using (var sha = SHA256.Create())
                {
                    image.Save(memory, ImageFormat.Png);
                    return BitConverter.ToString(sha.ComputeHash(memory.ToArray())).Replace("-", "").ToLowerInvariant();
                }
            }
            catch { return null; }
        }

        /// <summary>Retourne les propriétés d’un contrôle sélectionné par son chemin canonique.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <param name="controlName">Nom du contrôle à résoudre.</param>
        /// <returns>Données des propriétés du contrôle.</returns>
        public object ControlProperties(string projectName, string formName, string controlName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            object control = GetControl(form.Designer, controlName);
            string targetType = TypeDescriptor.GetClassName(control);
            var result = new List<object>();
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(control))
            {
                object value = null;
                string error = null;
                try
                {
                    object raw = descriptor.GetValue(control);
                    if (raw != null && !Marshal.IsComObject(raw))
                        value = Convert.ToString(raw, CultureInfo.InvariantCulture);
                }
                catch (Exception ex) { error = ex.Message; }
                result.Add(new
                {
                    descriptor.Name,
                    Type = descriptor.PropertyType?.FullName,
                    ReadOnly = descriptor.IsReadOnly,
                    SetterStatus = DesignerSetterStatus(targetType, descriptor),
                    AllowedValues = EnumChoices(descriptor.PropertyType),
                    Value = value,
                    Error = error
                });
            }
            return result;
        }

        /// <summary>Crée un UserForm ou une ressource de formulaire selon la commande.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Résultat de la création du formulaire.</returns>
        public object Create(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Form) ||
                !Regex.IsMatch(request.Form, @"^[A-Za-z][A-Za-z0-9_]{0,39}$"))
                throw new ArgumentException("Form must start with a letter and contain at most 40 letters, digits or underscores.");
            dynamic project = GetDesignProject(request.Project);
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, request.Form, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A component with this name already exists.");
            int countBefore = (int)project.VBComponents.Count;
            dynamic form = null;
            try
            {
                form = project.VBComponents.Add(3); // vbext_ct_MSForm
                form.Name = request.Form;
                form.DesignerWindow().Visible = true;
                return Snapshot(request.Project, form);
            }
            catch (Exception error)
            {
                string rollbackError = null;
                try { if (form != null) project.VBComponents.Remove(form); }
                catch (Exception rollback) { rollbackError = rollback.Message; }
                bool remains;
                try
                {
                    remains = (int)project.VBComponents.Count != countBefore;
                    foreach (dynamic component in project.VBComponents)
                        if (string.Equals((string)component.Name, request.Form, StringComparison.OrdinalIgnoreCase))
                            remains = true;
                }
                catch (Exception inspection)
                {
                    remains = true;
                    rollbackError = rollbackError == null ? inspection.Message : rollbackError + "; " + inspection.Message;
                }
                if (remains)
                    throw new InvalidOperationException("UserForm creation failed and rollback could not be verified. " +
                        "Inspect list_modules and list_forms before retrying. Cause: " + error.Message +
                        (rollbackError == null ? "" : " Rollback: " + rollbackError), error);
                throw new InvalidOperationException("UserForm creation failed; the component is absent after rollback. Cause: " +
                    error.Message, error);
            }
        }

        /// <summary>Ouvre le concepteur du UserForm ciblé.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <param name="formName">Nom du UserForm ciblé.</param>
        /// <returns>État de l’ouverture du concepteur.</returns>
        public object Open(string projectName, string formName)
        {
            dynamic form = GetForm(GetProject(projectName), formName);
            form.DesignerWindow().Visible = true;
            return Snapshot(projectName, form);
        }

        /// <summary>Ajoute un contrôle MSForms après validation du nom, type et version de l’arbre.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Contrôle créé et arbre actualisé.</returns>
        public object AddControl(Request request)
        {
            ValidateName(request.Control, "Control");
            ValidateGeometry(request);
            if (!VbeControlCatalog.IsCandidate(request.ControlType ?? string.Empty, BuiltInControls))
                throw new ArgumentException("ControlType must be a native MSForms or installed x64 CATID_Control ProgID.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic designer = form.Designer;
            foreach (dynamic existing in designer.Controls)
                if (string.Equals((string)existing.Name, request.Control, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A control with this name already exists.");
            try
            {
                dynamic control = designer.Controls.Add(request.ControlType, request.Control, true);
                // The control already exists at this point. Reject unsupported
                // properties before the first setter, then roll the control
                // back if any setter or final readback still fails.
                RequireWritableControlProperty((object)control, "Left");
                RequireWritableControlProperty((object)control, "Top");
                RequireWritableControlProperty((object)control, "Width");
                RequireWritableControlProperty((object)control, "Height");
                if (request.Caption != null)
                    RequireWritableControlProperty((object)control, "Caption");
                control.Left = request.Left;
                control.Top = request.Top;
                control.Width = request.Width;
                control.Height = request.Height;
                if (request.Caption != null) control.Caption = request.Caption;
                form.DesignerWindow().Visible = true;
                return Snapshot(request.Project, form);
            }
            catch (Exception error)
            {
                string rollbackError = null;
                try
                {
                    if (ControlNameExists(designer, request.Control))
                        designer.Controls.Remove(request.Control);
                }
                catch (Exception rollback) { rollbackError = rollback.Message; }
                bool remains;
                try { remains = ControlNameExists(designer, request.Control); }
                catch (Exception inspection)
                {
                    remains = true;
                    rollbackError = rollbackError == null ? inspection.Message : rollbackError + "; " + inspection.Message;
                }
                if (remains)
                    throw new InvalidOperationException("Control creation failed and rollback could not be verified. " +
                        "Inspect form_tree before retrying. Cause: " + error.Message +
                        (rollbackError == null ? "" : " Rollback: " + rollbackError), error);
                throw new InvalidOperationException("Control creation failed; no control with the requested name remains. Cause: " +
                    error.Message, error);
            }
        }

        /// <summary>Refuse une propriété de contrôle qui n’est pas modifiable par le concepteur.</summary>
        /// <param name="control">Contrôle dont l’écriture de propriété doit être vérifiée.</param>
        /// <param name="name">Nom de la propriété, du composant ou du contrôle.</param>
        private static void RequireWritableControlProperty(object control, string name)
        {
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(control).Find(name, true) ?? throw new InvalidOperationException("The selected control type does not expose " + name + ".");
            if (descriptor.IsReadOnly)
                throw new InvalidOperationException("The selected control type exposes " + name + " as read-only.");
        }

        /// <summary>Indique si un contrôle portant ce nom existe dans le formulaire.</summary>
        /// <param name="designer">Objet Designer qui contient le contrôle.</param>
        /// <param name="name">Nom de la propriété, du composant ou du contrôle.</param>
        /// <returns>true si le nom existe dans le formulaire.</returns>
        private static bool ControlNameExists(dynamic designer, string name)
        {
            return ControlNameExistsInCollection(designer.Controls, name);
        }

        /// <summary>Recherche le nom dans une collection de contrôles du concepteur.</summary>
        /// <param name="controls">Collection COM dans laquelle effectuer la recherche.</param>
        /// <param name="name">Nom du contrôle à rechercher.</param>
        /// <returns>true si un élément de la collection porte ce nom.</returns>
        private static bool ControlNameExistsInCollection(dynamic controls, string name)
        {
            foreach (dynamic item in controls)
                if (string.Equals((string)item.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        /// <summary>Valide et applique la géométrie au contrôle demandé.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>État du contrôle après application des coordonnées et dimensions.</returns>
        public object SetControlGeometry(Request request)
        {
            ValidateName(request.Control, "Control");
            ValidateGeometry(request);
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic control = GetControl(form.Designer, request.Control);
            control.Left = request.Left;
            control.Top = request.Top;
            control.Width = request.Width;
            control.Height = request.Height;
            form.DesignerWindow().Visible = true;
            return Snapshot(request.Project, form);
        }

        /// <summary>Renomme un contrôle après vérification de la version et de l’unicité du nom.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Nouvel état du contrôle renommé.</returns>
        public object RenameControl(Request request)
        {
            ValidateName(request.Control, "Control");
            ValidateName(request.NewName, "NewName");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic designer = form.Designer;
            foreach (dynamic existing in designer.Controls)
                if (string.Equals((string)existing.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A control with the new name already exists.");
            dynamic control = GetControl(designer, request.Control);
            control.Name = request.NewName;
            return Snapshot(request.Project, form);
        }

        /// <summary>Modifie la légende du contrôle ciblé.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Nouvel état après affectation de la légende.</returns>
        public object SetControlCaption(Request request)
        {
            if (request.Caption == null) throw new ArgumentException("Caption is required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic control = GetControl(form.Designer, request.Control);
            control.Caption = request.Caption;
            return Snapshot(request.Project, form);
        }

        /// <summary>Modifie les propriétés de police prises en charge du contrôle.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Nouvel état après mise à jour de la police.</returns>
        public object SetControlFont(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.FontName) || !IsFinite(request.FontSize) ||
                request.FontSize <= 0 || request.FontSize > 200)
                throw new ArgumentException("FontName and a FontSize between 0 and 200 are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            AssertVersion(request, form);
            dynamic control = GetControl(form.Designer, request.Control);
            dynamic font = control.Font;
            font.Name = request.FontName;
            font.Size = request.FontSize;
            font.Bold = request.FontBold;
            return Snapshot(request.Project, form);
        }

        /// <summary>Résout un projet VBE par son nom.</summary>
        /// <param name="name">Nom de la propriété, du composant ou du contrôle.</param>
        /// <returns>Projet VBE résolu.</returns>
        private dynamic GetProject(string name)
        {
            return VbeProjectResolver.Resolve(vbe, name);
        }

        /// <summary>Résout le projet et exige le mode conception avant modification.</summary>
        /// <returns>Projet correspondant, validé en mode conception.</returns>
        /// <param name="name">Nom du projet VBE à valider.</param>
        private dynamic GetDesignProject(string name)
        {
            dynamic project = GetProject(name);
            if ((int)project.Mode != 2) throw new InvalidOperationException("The project must be in design mode.");
            return project;
        }

        /// <summary>Recherche le UserForm nommé dans les composants du projet.</summary>
        /// <param name="project">Projet VBE résolu.</param>
        /// <param name="name">Nom de la propriété, du composant ou du contrôle.</param>
        /// <returns>UserForm correspondant au nom demandé.</returns>
        private static dynamic GetForm(dynamic project, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Form is required.");
            foreach (dynamic component in project.VBComponents)
                if (string.Equals((string)component.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if ((int)component.Type != 3) throw new InvalidOperationException("The component is not a UserForm.");
                    return component;
                }
            throw new InvalidOperationException("UserForm not found: " + name);
        }

        /// <summary>Résout le contrôle désigné par son chemin hiérarchique canonique.</summary>
        /// <param name="designer">Objet Designer qui contient le contrôle.</param>
        /// <param name="name">Nom de la propriété, du composant ou du contrôle.</param>
        /// <returns>Contrôle résolu dans le concepteur.</returns>
        private static dynamic GetControl(dynamic designer, string name)
        {
            foreach (dynamic control in designer.Controls)
                if (string.Equals((string)control.Name, name, StringComparison.OrdinalIgnoreCase)) return control;
            throw new InvalidOperationException("Control not found: " + name);
        }

        /// <summary>Valide le nom VBA d’un formulaire ou d’un contrôle.</summary>
        /// <param name="name">Nom de la propriété, du composant ou du contrôle.</param>
        /// <param name="label">Nom logique utilisé dans le message de validation.</param>
        private static void ValidateName(string name, string label)
        {
            if (string.IsNullOrWhiteSpace(name) || !Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                throw new ArgumentException(label + " must be a VBA identifier.");
        }

        /// <summary>Vérifie que la géométrie demandée est finie et dans les bornes admises.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        private static void ValidateGeometry(Request request)
        {
            if (!IsFinite(request.Left) || !IsFinite(request.Top) || !IsFinite(request.Width) || !IsFinite(request.Height) ||
                request.Left < 0 || request.Top < 0 || request.Width <= 0 || request.Height <= 0 ||
                request.Left > 32767 || request.Top > 32767 || request.Width > 32767 || request.Height > 32767)
                throw new ArgumentException("Control geometry must use finite nonnegative point coordinates and positive sizes.");
        }

        /// <summary>Indique si la valeur flottante est finie.</summary>
        /// <param name="value">Valeur demandée ou lue pour la propriété.</param>
        /// <returns>true si la valeur est finie.</returns>
        private static bool IsFinite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }

        /// <summary>Vérifie que l’arbre n’a pas changé depuis la lecture fournie par le client.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <param name="form">Formulaire VBE inspecté.</param>
        private static void AssertVersion(Request request, dynamic form)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedFormVersion))
                throw new ArgumentException("ExpectedFormVersion is required for form edits.");
            string current = Version(form);
            if (!string.Equals(current, request.ExpectedFormVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form changed since it was read.");
        }

        /// <summary>Crée un instantané de l’état et de la version du formulaire.</summary>
        /// <param name="projectName">Nom du projet contenant le formulaire.</param>
        /// <param name="form">Formulaire VBE inspecté.</param>
        /// <returns>Instantané sérialisable du formulaire.</returns>
        private static object Snapshot(string projectName, dynamic form)
        {
            dynamic designer = form.Designer;
            var controls = new List<object>();
            foreach (dynamic control in designer.Controls)
            {
                if (!SameContainer((object)control.Parent, (object)designer, (string)form.Name)) continue;
                string caption = null;
                try { caption = (string)control.Caption; } catch { }
                string fontName = null;
                double? fontSize = null;
                bool? fontBold = null;
                try
                {
                    dynamic font = control.Font;
                    fontName = (string)font.Name;
                    fontSize = Convert.ToDouble(font.Size, CultureInfo.InvariantCulture);
                    fontBold = (bool)font.Bold;
                }
                catch { }
                controls.Add(new
                {
                    Name = (string)control.Name,
                    Caption = caption,
                    Left = (double)control.Left,
                    Top = (double)control.Top,
                    Width = (double)control.Width,
                    Height = (double)control.Height,
                    FontName = fontName,
                    FontSize = fontSize,
                    FontBold = fontBold
                });
            }
            return new
            {
                Project = projectName,
                Form = (string)form.Name,
                Caption = (string)form.Properties.Item("Caption").Value,
                Width = Convert.ToDouble(form.Properties.Item("Width").Value, CultureInfo.InvariantCulture),
                Height = Convert.ToDouble(form.Properties.Item("Height").Value, CultureInfo.InvariantCulture),
                Version = Version(form),
                Controls = controls
            };
        }

        /// <summary>Calcule l’empreinte de version du formulaire et de ses contrôles.</summary>
        /// <param name="form">Formulaire VBE inspecté.</param>
        /// <returns>Empreinte de version du formulaire.</returns>
        private static string Version(dynamic form)
        {
            return TreeVersion(form, out List<object> _, out List<VbePropertyInfo> _, out int _);
        }

        // Create inside a verified canonical container and confirm rollback if any post-add step fails.
        /// <summary>Ajoute un contrôle dans le conteneur sélectionné et vérifie le résultat dans le nouvel arbre.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Résultat de l’ajout et arbre actualisé.</returns>
        public object AddNestedControl(Request request)
        {
            ValidateName(request.Control, "Control");
            ValidateGeometry(request);
            if (!VbeControlCatalog.IsCandidate(request.ControlType ?? string.Empty, BuiltInControls))
                throw new ArgumentException("ControlType must be a native MSForms or installed x64 CATID_Control ProgID.");
            if (string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ExpectedTreeVersion is required from form_tree.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ParentPath))
                throw new InvalidOperationException("ParentPath is not a canonical path in form_tree.");
            dynamic controls = ResolveNestedControls(form.Designer, request.ParentPath);
            foreach (dynamic existing in controls)
                if (string.Equals((string)existing.Name, request.Control, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("A control with this name already exists in the container.");
            try
            {
                dynamic control = controls.Add(request.ControlType, request.Control, true);
                RequireWritableControlProperty((object)control, "Left");
                RequireWritableControlProperty((object)control, "Top");
                RequireWritableControlProperty((object)control, "Width");
                RequireWritableControlProperty((object)control, "Height");
                if (request.Caption != null)
                    RequireWritableControlProperty((object)control, "Caption");
                control.Left = request.Left;
                control.Top = request.Top;
                control.Width = request.Width;
                control.Height = request.Height;
                if (request.Caption != null) control.Caption = request.Caption;
                dynamic after = Tree(request.Project, request.Form);
                if (string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                    StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The nested control was not reflected in the UserForm tree.");
                return after;
            }
            catch (Exception error)
            {
                string rollbackError = null;
                try
                {
                    if (ControlNameExistsInCollection(controls, request.Control))
                        controls.Remove(request.Control);
                }
                catch (Exception rollback) { rollbackError = rollback.Message; }
                bool remains;
                try { remains = ControlNameExistsInCollection(controls, request.Control); }
                catch (Exception inspection)
                {
                    remains = true;
                    rollbackError = rollbackError == null ? inspection.Message : rollbackError + "; " + inspection.Message;
                }
                if (remains)
                    throw new InvalidOperationException("Nested control creation failed and rollback could not be verified. " +
                        "Inspect form_tree before retrying. Cause: " + error.Message +
                        (rollbackError == null ? "" : " Rollback: " + rollbackError), error);
                throw new InvalidOperationException("Nested control creation failed; no control with the requested name remains. Cause: " +
                    error.Message, error);
            }
        }

        /// <summary>Modifie une propriété autorisée du nœud sélectionné puis retourne l’arbre actualisé.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Résultat de l’affectation et arbre actualisé.</returns>
        public object SetNodeProperty(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.Property) || request.Value == null ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath, Property, Value and ExpectedTreeVersion are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            string[] propertyPath = request.Property.Split('.');
            if (propertyPath.Length < 1 || propertyPath.Length > 2 ||
                propertyPath.Any(part => string.IsNullOrWhiteSpace(part)))
                throw new ArgumentException("Property must be a property name or one object member path.");
            PropertyDescriptor root = TypeDescriptor.GetProperties(target).Find(propertyPath[0], true) ?? throw new InvalidOperationException("Property is not exposed: " + propertyPath[0]);
            if (string.Equals(root.Name, "_Font_Reserved", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("_Font_Reserved is a COM reserved member whose getter is unavailable in the tested VBE.");
            if (propertyPath.Length == 1)
            {
                // A ComboBox.ColumnCount=2 write followed by three AddItem
                // calls was followed by Excel heap corruption. The exact
                // trigger in that sequence has not been isolated.
                if (string.Equals(TypeDescriptor.GetClassName(target), "ComboBox", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(root.Name, "ColumnCount", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "ComboBox.ColumnCount editing is temporarily disabled: Excel crashed after a multicolumn design-time write sequence.");
                // A single TextBox.ScrollBars=Vertical write was followed by
                // Excel heap corruption after a successful readback.
                if (string.Equals(TypeDescriptor.GetClassName(target), "TextBox", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(root.Name, "ScrollBars", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "TextBox.ScrollBars editing is temporarily disabled: Excel crashed after a design-time write.");
                // A disposable Excel session crashed with heap corruption
                // shortly after a sequence of these SpinButton setters. The
                // responsible member has not yet been isolated.
                if (string.Equals(TypeDescriptor.GetClassName(target), "SpinButton", StringComparison.OrdinalIgnoreCase) &&
                    new[] { "Min", "Max", "Value", "Delay", "SmallChange" }
                        .Any(name => string.Equals(root.Name, name, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException(
                        "SpinButton." + root.Name + " editing is temporarily disabled: Excel crashed after a design-time property write sequence containing this member.");
                // Two disposable Excel sessions crashed during teardown after
                // ToggleButton.Value=true was set in the designer. Until the
                // host interaction is isolated, refuse every Value write for
                // this type before invoking its COM setter.
                if (string.Equals(root.Name, "Value", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(TypeDescriptor.GetClassName(target), "ToggleButton", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "ToggleButton.Value editing is temporarily disabled: Excel crashed during teardown after a design-time Value=true write.");
                // A real Excel Label exposed Cancel as writable through
                // PropertyDescriptor, but the COM setter returned member-not-found.
                // Reject the known bad path before invoking native COM.
                if (string.Equals(root.Name, "Cancel", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(TypeDescriptor.GetClassName(target), "Label", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Cancel has no usable designer setter in the tested Excel VBE.");
                if (root.IsReadOnly) throw new InvalidOperationException("Property is read-only: " + root.Name);
                object oldValue = root.GetValue(target);
                object converted = ConvertDescriptorValue(request.Value, root.PropertyType, oldValue);
                SetDesignerScalar(target, root, converted);
                object actual = root.GetValue(target);
                if (!SameDescriptorValue(actual, converted))
                    throw new InvalidOperationException("The VBE did not retain property " + root.Name + ".");
            }
            else
            {
                object owner = string.Equals(root.Name, "Font", StringComparison.OrdinalIgnoreCase)
                    ? (object)((dynamic)target).Font : root.GetValue(target);
                if (owner == null || !Marshal.IsComObject(owner))
                    throw new InvalidOperationException("The object property is not exposed as an editable COM object.");
                PropertyDescriptor member = TypeDescriptor.GetProperties(owner).Find(propertyPath[1], true) ?? throw new InvalidOperationException("Object member is not exposed: " + request.Property);
                if (member.IsReadOnly) throw new InvalidOperationException("Object member is read-only: " + request.Property);
                object oldValue = member.GetValue(owner);
                object converted = ConvertDescriptorValue(request.Value, member.PropertyType, oldValue);
                SetDesignerScalar(owner, member, converted);
                if (!SameDescriptorValue(member.GetValue(owner), converted))
                    throw new InvalidOperationException("The VBE did not retain object member " + request.Property + ".");
            }
            dynamic after = Tree(request.Project, request.Form);
            return new { request.ControlPath, request.Property, Tree = after };
        }

        /// <summary>Affecte une img OLE au nœud sélectionné après vérification de son chemin et de la version.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Résultat de l’affectation de l’img et arbre actualisé.</returns>
        public object SetNodePicture(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.Property) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath, Property and ExpectedTreeVersion are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(target).Find(request.Property, true);
            if (descriptor == null || descriptor.IsReadOnly ||
                (descriptor.PropertyType != typeof(Bitmap) && descriptor.PropertyType != typeof(Icon)))
                throw new InvalidOperationException("The selected node has no writable OLE img property by that name.");
            object picture = OlePictureLoader.Load(request.Path);
            target.GetType().InvokeMember(descriptor.Name, BindingFlags.SetProperty,
                null, target, new[] { picture });
            object installed = target.GetType().InvokeMember(descriptor.Name, BindingFlags.GetProperty,
                null, target, null);
            if (!string.Equals(OlePictureLoader.Fingerprint(installed), OlePictureLoader.Fingerprint(picture),
                StringComparison.Ordinal))
                throw new InvalidOperationException("The VBE did not retain the requested OLE img.");
            return new
            {
                request.ControlPath,
                Property = descriptor.Name,
                Tree = Tree(request.Project, request.Form)
            };
        }

        /// <summary>Retire un contrôle après validation du chemin canonique et de la version de l’arbre.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Confirmation du retrait et arbre actualisé.</returns>
        public object RemoveControl(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 ||
                !string.Equals(parts[parts.Length - 2], "Controls", StringComparison.Ordinal))
                throw new ArgumentException("ControlPath must identify a control, not a Page or Tab.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            string name = parts[parts.Length - 1];
            object owner = parts.Length == 2 ? (object)form.Designer :
                ResolveTreeItem(form.Designer, string.Join("/", parts.Take(parts.Length - 2)));
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(owner).Find("Controls", true) ?? throw new InvalidOperationException("The selected parent has no Controls collection.");
            dynamic controls = descriptor.GetValue(owner);
            controls.Remove(name);
            dynamic after = Tree(request.Project, request.Form);
            if (TreeContainsPath((IEnumerable)after.Controls, request.ControlPath) ||
                string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The control removal was not reflected in the UserForm tree.");
            return new { RemovedPath = request.ControlPath, Applied = true, Tree = after };
        }

        /// <summary>Place le contrôle sélectionné au premier plan ou à l’arrière-plan de son conteneur.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Résultat du changement d’ordre visuel et arbre actualisé.</returns>
        public object ZOrderControl(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            if (request.ZPosition != 0 && request.ZPosition != 1)
                throw new ArgumentOutOfRangeException("ZPosition", "Use 0 for front or 1 for back.");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 2 || parts[parts.Length - 2] != "Controls")
                throw new ArgumentException("ControlPath must identify a control, not a Page or Tab.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            object control = ResolveTreeItem(form.Designer, request.ControlPath);
            ((dynamic)control).ZOrder(request.ZPosition);
            dynamic after = Tree(request.Project, request.Form);
            if (!TreeContainsPath((IEnumerable)after.Controls, request.ControlPath))
                throw new InvalidOperationException("The control is no longer present after ZOrder.");
            return new
            {
                request.ControlPath,
                request.ZPosition,
                Executed = true,
                Verification = "Unverified",
                VerificationPending = true,
                VerificationLimit = "MSForms does not expose z-order through Controls or form_tree; compare the visible overlap in the designer.",
                TreeBefore = before,
                TreeAfter = after
            };
        }

        /// <summary>Ajoute une page de MultiPage ou un onglet de TabStrip après validation de l’arbre.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <param name="collectionName">Nom de la collection contrôlée.</param>
        /// <returns>Résultat de l’ajout et arbre actualisé.</returns>
        public object AddPageOrTab(Request request, string collectionName)
        {
            ValidateName(request.NewName, "NewName");
            if (string.IsNullOrWhiteSpace(request.ParentPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ParentPath and ExpectedTreeVersion are required.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ParentPath))
                throw new InvalidOperationException("ParentPath is not a canonical path in form_tree.");
            object parent = ResolveTreeItem(form.Designer, request.ParentPath);
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(parent).Find(collectionName, true);
            if (descriptor == null ||
                !string.Equals(TypeDescriptor.GetClassName(parent),
                    collectionName == "Pages" ? "MultiPage" : "TabStrip",
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected parent is not a compatible MultiPage or TabStrip.");
            dynamic collection = descriptor.GetValue(parent);
            foreach (dynamic item in collection)
                if (string.Equals((string)item.Name, request.NewName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("An item with this name already exists in the collection.");
            int count = (int)collection.Count;
            if (request.InsertIndex.HasValue &&
                (request.InsertIndex.Value < 0 || request.InsertIndex.Value > count))
                throw new ArgumentOutOfRangeException("InsertIndex", "InsertIndex must be between zero and Count.");
            string caption = request.Caption ?? request.NewName;
            string newPath = request.ParentPath + "/" + collectionName + "/" + request.NewName;
            try
            {
                if (request.InsertIndex.HasValue)
                    collection.Add(request.NewName, caption, request.InsertIndex.Value);
                else collection.Add(request.NewName, caption);
                dynamic after = Tree(request.Project, request.Form);
                if (!TreeContainsPath((IEnumerable)after.Controls, newPath) ||
                    string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The added Page or Tab was not reflected in the UserForm tree.");
                return new { AddedPath = newPath, Applied = true, Tree = after };
            }
            catch (Exception error)
            {
                string rollbackError = null;
                try
                {
                    int index = FindCollectionIndexByName(collection, request.NewName);
                    if (index >= 0) collection.Remove(index);
                }
                catch (Exception rollback) { rollbackError = rollback.Message; }
                bool remains;
                try
                {
                    remains = FindCollectionIndexByName(collection, request.NewName) >= 0 ||
                        (int)collection.Count != count;
                }
                catch (Exception inspection)
                {
                    remains = true;
                    rollbackError = rollbackError == null ? inspection.Message : rollbackError + "; " + inspection.Message;
                }
                if (remains)
                    throw new InvalidOperationException("Page or Tab creation failed and rollback could not be verified. " +
                        "Inspect form_tree before retrying. Cause: " + error.Message +
                        (rollbackError == null ? "" : " Rollback: " + rollbackError), error);
                throw new InvalidOperationException("Page or Tab creation failed; the item is absent after rollback. Cause: " +
                    error.Message, error);
            }
        }

        /// <summary>Retourne l’index de l’élément correspondant au nom dans la collection.</summary>
        /// <param name="collection">Collection COM des contrôles à parcourir.</param>
        /// <param name="name">Nom de la propriété, du composant ou du contrôle.</param>
        /// <returns>Index de l’élément correspondant, ou -1.</returns>
        private static int FindCollectionIndexByName(dynamic collection, string name)
        {
            int index = 0;
            foreach (dynamic item in collection)
            {
                if (string.Equals((string)item.Name, name, StringComparison.Ordinal)) return index;
                index++;
            }
            return -1;
        }

        /// <summary>Retire la page ou l’onglet exact après vérification de la structure et de la version.</summary>
        /// <param name="request">Paramètres de la commande et version attendue par le client.</param>
        /// <returns>Confirmation du retrait et arbre actualisé.</returns>
        public object RemovePageOrTab(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 4 || parts.Length % 2 != 0 ||
                (parts[parts.Length - 2] != "Pages" && parts[parts.Length - 2] != "Tabs"))
                throw new ArgumentException("ControlPath must identify a Page or Tab.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)before.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not a canonical path in form_tree.");
            string collectionName = parts[parts.Length - 2];
            string name = parts[parts.Length - 1];
            string parentPath = string.Join("/", parts.Take(parts.Length - 2));
            object parent = ResolveTreeItem(form.Designer, parentPath);
            PropertyDescriptor descriptor = TypeDescriptor.GetProperties(parent).Find(collectionName, true) ?? throw new InvalidOperationException("The selected parent has no " + collectionName + " collection.");
            dynamic collection = descriptor.GetValue(parent);
            int index = -1;
            int current = 0;
            foreach (dynamic item in collection)
            {
                if (string.Equals((string)item.Name, name, StringComparison.Ordinal))
                {
                    if (index >= 0) throw new InvalidOperationException("Page or Tab name is ambiguous.");
                    index = current;
                }
                current++;
            }
            if (index < 0) throw new InvalidOperationException("Page or Tab no longer exists in its collection.");
            collection.Remove(index);
            dynamic after = Tree(request.Project, request.Form);
            if (TreeContainsPath((IEnumerable)after.Controls, request.ControlPath) ||
                string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The Page or Tab removal was not reflected in the UserForm tree.");
            return new { RemovedPath = request.ControlPath, Applied = true, Tree = after };
        }

        /// <summary>Convertit la valeur JSON selon le type de la propriété décrite.</summary>
        /// <param name="value">Valeur demandée ou lue pour la propriété.</param>
        /// <param name="declaredType">Type déclaré de la propriété COM.</param>
        /// <param name="previous">Valeur existante utilisée pour préserver les membres non modifiés.</param>
        /// <returns>Valeur convertie vers le type de propriété.</returns>
        private static object ConvertDescriptorValue(object value, Type declaredType, object previous)
        {
            bool variant = declaredType != null &&
                declaredType.FullName == "System.Windows.Forms.ComponentModel.Com2Interop.Com2Variant";
            Type type = declaredType == typeof(object) || declaredType == null || variant
                ? previous?.GetType() ?? value.GetType() : declaredType;
            if (type == typeof(Color))
            {
                if (value is string v && v.StartsWith("#", StringComparison.Ordinal))
                    return ColorTranslator.FromHtml(v);
                return ColorTranslator.FromOle(Convert.ToInt32(value, CultureInfo.InvariantCulture));
            }
            return ConvertScalar(value, type);
        }

        /// <summary>Compare deux valeurs selon les règles de la propriété décrite.</summary>
        /// <param name="actual">Valeur effectivement lue après modification.</param>
        /// <param name="expected">Valeur attendue après modification.</param>
        /// <returns>true si les valeurs sont identiques selon le descripteur.</returns>
        private static bool SameDescriptorValue(object actual, object expected)
        {
            if (Equals(actual, expected)) return true;
            if (actual == null || expected == null) return false;
            if (actual is Color clr && expected is Color clr2)
                return clr.ToArgb() == clr2.ToArgb();
            if (actual is IConvertible && expected is IConvertible)
                return string.Equals(Convert.ToString(actual, CultureInfo.InvariantCulture),
                    Convert.ToString(expected, CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase);
            return false;
        }

        /// <summary>Indique si l’arbre contient exactement le chemin canonique demandé.</summary>
        /// <param name="nodes">Nœuds parcourus dans l’arbre.</param>
        /// <param name="path">Chemin à rechercher dans les nœuds.</param>
        /// <returns>true si le chemin canonique existe dans l’arbre.</returns>
        private static bool TreeContainsPath(IEnumerable nodes, string path)
        {
            foreach (dynamic node in nodes)
            {
                if (string.Equals((string)node.Path, path, StringComparison.Ordinal)) return true;
                if (TreeContainsPath((IEnumerable)node.Children, path)) return true;
            }
            return false;
        }

        /// <summary>Résout un contrôle ou une page en suivant chaque segment du chemin canonique.</summary>
        /// <param name="designer">Objet Designer qui contient le contrôle.</param>
        /// <param name="path">Chemin à rechercher dans les nœuds.</param>
        /// <returns>Objet résolu au chemin canonique.</returns>
        private static object ResolveTreeItem(object designer, string path)
        {
            string[] parts = path.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 || parts.Length > 16)
                throw new ArgumentException("Invalid control hierarchy path.");
            object current = designer;
            for (int i = 0; i < parts.Length; i += 2)
            {
                PropertyDescriptor collection = TypeDescriptor.GetProperties(current).Find(parts[i], true);
                if (collection == null || (parts[i] != "Controls" && parts[i] != "Pages" && parts[i] != "Tabs"))
                    throw new InvalidOperationException("Path collection is unavailable: " + parts[i]);
                object found = null;
                foreach (dynamic candidate in (dynamic)collection.GetValue(current))
                    if (string.Equals((string)candidate.Name, parts[i + 1], StringComparison.Ordinal))
                    { found = candidate; break; }

                current = found ?? throw new InvalidOperationException("Path item is unavailable: " + parts[i + 1]);
            }
            return current;
        }

        /// <summary>Résout les segments de collections imbriquées du chemin de contrôle.</summary>
        /// <param name="designer">Objet Designer qui contient le contrôle.</param>
        /// <param name="path">Chemin à rechercher dans les nœuds.</param>
        /// <returns>Contrôle ou page obtenu au terme du chemin.</returns>
        private static dynamic ResolveNestedControls(dynamic designer, string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("ParentPath is required.");
            string[] parts = path.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 || parts[0] != "Controls" || parts.Length > 12)
                throw new ArgumentException("ParentPath must start with Controls/<name> and use Controls or Pages segments.");
            object current = designer;
            for (int i = 0; i < parts.Length; i += 2)
            {
                string collectionName = parts[i];
                string itemName = parts[i + 1];
                if ((collectionName != "Controls" && collectionName != "Pages") ||
                    !Regex.IsMatch(itemName, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                    throw new ArgumentException("ParentPath contains an invalid collection or name.");
                PropertyDescriptor descriptor = TypeDescriptor.GetProperties(current).Find(collectionName, true) ?? throw new InvalidOperationException("The path element has no " + collectionName + " collection.");
                object collection = descriptor.GetValue(current);
                object match = null;
                foreach (dynamic candidate in (dynamic)collection)
                    if (string.Equals((string)candidate.Name, itemName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (match != null) throw new InvalidOperationException("Ambiguous parent path element: " + itemName);
                        match = candidate;
                    }

                current = match ?? throw new InvalidOperationException("Parent path element not found: " + itemName);
            }
            PropertyDescriptor controls = TypeDescriptor.GetProperties(current).Find("Controls", true) ?? throw new InvalidOperationException("The selected parent has no Controls collection.");
            return controls.GetValue(current);
        }

    }
}
