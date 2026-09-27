using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        private sealed class VerifiedChildProfile
        {
            public string Type;
            public string ProgId;
            public string[] Fields;
            public string Limitation;
        }

        private sealed class ProfiledChildSnapshot
        {
            public VerifiedChildProfile Profile;
            public string SourcePath;
            public string ProposedName;
            public string ProposedPath;
            public string Caption;
            public string TextValue;
            public bool BooleanValue;
            public string ListWidth;
            public double Left;
            public double Top;
            public double Width;
            public double Height;
            public int BackColor;
            public string FontName;
            public double FontSize;
            public bool FontBold;
        }

        private sealed class ProfiledFrameSnapshot
        {
            public string SourcePath;
            public string NewName;
            public string Caption;
            public double Left;
            public double Top;
            public double Width;
            public double Height;
            public string TreeVersion;
            public int NodeCount;
            public readonly List<ProfiledChildSnapshot> Children = new List<ProfiledChildSnapshot>();
            public readonly List<string> Issues = new List<string>();
        }

        // A positive registry. No profile is inferred from COM IsReadOnly or
        // PROPERTYPUT metadata; every field below came from a prior Excel probe.
        private static readonly Dictionary<string, VerifiedChildProfile> FrameChildProfiles =
            new Dictionary<string, VerifiedChildProfile>(StringComparer.OrdinalIgnoreCase)
            {
                ["Label"] = new VerifiedChildProfile { Type = "Label", ProgId = "Forms.Label.1",
                    Fields = new[] { "Name", "Caption", "Left", "Top", "Width", "Height",
                        "BackColor", "Font.Name", "Font.Size", "Font.Bold" },
                    Limitation = "Other Label properties, image, and z-order are not copied." },
                ["TextBox"] = new VerifiedChildProfile { Type = "TextBox", ProgId = "Forms.TextBox.1",
                    Fields = new[] { "Name", "Left", "Top", "Width", "Height", "Value (text or empty)" },
                    Limitation = "Bindings, validation, and other TextBox properties are not copied." },
                ["CheckBox"] = new VerifiedChildProfile { Type = "CheckBox", ProgId = "Forms.CheckBox.1",
                    Fields = new[] { "Name", "Caption", "Left", "Top", "Width", "Height",
                        "Value (Boolean only)" },
                    Limitation = "TriState/null and other CheckBox properties are not copied." },
                ["CommandButton"] = new VerifiedChildProfile { Type = "CommandButton", ProgId = "Forms.CommandButton.1",
                    Fields = new[] { "Name", "Caption", "Left", "Top", "Width", "Height" },
                    Limitation = "Event procedures and other CommandButton properties are not copied." },
                ["ComboBox"] = new VerifiedChildProfile { Type = "ComboBox", ProgId = "Forms.ComboBox.1",
                    Fields = new[] { "Name", "Left", "Top", "Width", "Height", "ListWidth (text)" },
                    Limitation = "Items, bindings, selection, and other ComboBox properties are not copied." },
                ["OptionButton"] = new VerifiedChildProfile { Type = "OptionButton", ProgId = "Forms.OptionButton.1",
                    Fields = new[] { "Name", "Caption", "Left", "Top", "Width", "Height" },
                    Limitation = "Value and GroupName are not copied; copying selection can affect siblings." }
            };

        public object FrameProfileCopyPlan(Request request)
        {
            ProfiledFrameSnapshot plan = ReadFrameProfilePlan(request);
            return new { Project = request.Project, Form = request.Form,
                SourcePath = plan.SourcePath, ProposedFrameName = plan.NewName,
                ExpectedTreeVersion = plan.TreeVersion, DirectChildCount = plan.Children.Count,
                Children = plan.Children.Select(child => new { SourcePath = child.SourcePath,
                    ProposedPath = child.ProposedPath, Type = child.Profile.Type,
                    CopiedFields = child.Profile.Fields, Limitation = child.Profile.Limitation }).ToArray(),
                Issues = plan.Issues, EligibleForLimitedProbe = plan.Issues.Count == 0,
                Completeness = "Partial", MutationVerified = false, ReadOnly = true,
                Scope = "Root Frame and direct leaf children with explicit positive profiles only." };
        }

        private ProfiledFrameSnapshot ReadFrameProfilePlan(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                request.ControlPath.Split('/').Length != 2 ||
                !request.ControlPath.StartsWith("Controls/", StringComparison.Ordinal))
                throw new ArgumentException("Only a root Controls/<Frame> path is supported.");
            if (string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ExpectedTreeVersion is required from form_tree.");
            ValidateName(request.NewName, "NewName");
            dynamic form = GetForm(GetProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)tree.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not canonical in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            if (!string.Equals(TypeDescriptor.GetClassName(target), "Frame", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected control is not a native MSForms Frame.");

            dynamic frame = target;
            var plan = new ProfiledFrameSnapshot {
                SourcePath = request.ControlPath, NewName = request.NewName,
                TreeVersion = (string)tree.TreeVersion, NodeCount = (int)tree.NodeCount,
                Caption = (string)frame.Caption,
                Left = Convert.ToDouble(frame.Left, CultureInfo.InvariantCulture),
                Top = Convert.ToDouble(frame.Top, CultureInfo.InvariantCulture),
                Width = Convert.ToDouble(frame.Width, CultureInfo.InvariantCulture),
                Height = Convert.ToDouble(frame.Height, CultureInfo.InvariantCulture)
            };
            try { ValidateCopyBox(plan.Left, plan.Top, plan.Width, plan.Height, "Frame"); }
            catch (Exception ex) { plan.Issues.Add(ex.Message); }
            if (request.NewName.Length > 40)
                plan.Issues.Add("New Frame name exceeds 40 characters.");

            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic existing in form.Designer.Controls)
                names.Add((string)existing.Name);
            if (names.Contains(request.NewName))
                plan.Issues.Add("New Frame name already exists: " + request.NewName);
            int directCount = 0;
            foreach (dynamic child in frame.Controls)
            {
                if (!SameContainer((object)child.Parent, target, (string)form.Name)) continue;
                directCount++;
                if (directCount > 128)
                    throw new InvalidOperationException("Frame exceeds the 128 direct-child copy limit.");
                string sourceName = (string)child.Name;
                string type = TypeDescriptor.GetClassName((object)child);
                VerifiedChildProfile profile;
                if (!FrameChildProfiles.TryGetValue(type, out profile))
                {
                    plan.Issues.Add(sourceName + ": no positive copy profile for " + type + ".");
                    continue;
                }
                string newName = request.NewName + "_" + sourceName;
                if (newName.Length > 40)
                    plan.Issues.Add(sourceName + ": proposed child name exceeds 40 characters.");
                if (!names.Add(newName))
                    plan.Issues.Add(sourceName + ": proposed child name already exists: " + newName);
                try
                {
                    plan.Children.Add(ReadProfiledChild(child, profile,
                        request.ControlPath + "/Controls/" + sourceName,
                        "Controls/" + request.NewName + "/Controls/" + newName, newName));
                }
                catch (Exception ex)
                {
                    plan.Issues.Add(sourceName + ": " + ex.Message);
                }
            }
            if (directCount == 0)
                plan.Issues.Add("Use duplicate_empty_form_frame for a Frame without children.");
            if (plan.Children.Count != directCount)
                plan.Issues.Add("Not all direct children have readable positive profiles.");
            return plan;
        }

        private static ProfiledChildSnapshot ReadProfiledChild(dynamic child,
            VerifiedChildProfile profile, string sourcePath, string proposedPath, string newName)
        {
            var item = new ProfiledChildSnapshot { Profile = profile, SourcePath = sourcePath,
                ProposedName = newName, ProposedPath = proposedPath,
                Left = Convert.ToDouble(child.Left, CultureInfo.InvariantCulture),
                Top = Convert.ToDouble(child.Top, CultureInfo.InvariantCulture),
                Width = Convert.ToDouble(child.Width, CultureInfo.InvariantCulture),
                Height = Convert.ToDouble(child.Height, CultureInfo.InvariantCulture) };
            ValidateCopyBox(item.Left, item.Top, item.Width, item.Height, profile.Type);
            switch (profile.Type)
            {
                case "Label":
                    item.Caption = (string)child.Caption;
                    item.BackColor = CopyOleColor((object)child.BackColor);
                    item.FontName = (string)child.Font.Name;
                    item.FontSize = Convert.ToDouble(child.Font.Size, CultureInfo.InvariantCulture);
                    item.FontBold = (bool)child.Font.Bold;
                    if (string.IsNullOrWhiteSpace(item.FontName) || !IsFinite(item.FontSize) ||
                        item.FontSize <= 0 || item.FontSize > 200)
                        throw new InvalidOperationException("Label font is outside the tested range.");
                    break;
                case "TextBox":
                    object text = child.Value;
                    if (text != null && !(text is string))
                        throw new InvalidOperationException("TextBox.Value must be text or empty.");
                    item.TextValue = (string)text;
                    break;
                case "CheckBox":
                    item.Caption = (string)child.Caption;
                    object check = child.Value;
                    if (!(check is bool))
                        throw new InvalidOperationException("CheckBox.Value must be Boolean.");
                    item.BooleanValue = (bool)check;
                    break;
                case "CommandButton":
                case "OptionButton":
                    item.Caption = (string)child.Caption;
                    break;
                case "ComboBox":
                    object listWidth = child.ListWidth;
                    if (!(listWidth is string) || string.IsNullOrWhiteSpace((string)listWidth) ||
                        ((string)listWidth).Length > 64)
                        throw new InvalidOperationException("ComboBox.ListWidth must be nonempty text.");
                    item.ListWidth = (string)listWidth;
                    break;
                default: throw new InvalidOperationException("Profile has no copy implementation.");
            }
            return item;
        }

        public object DuplicateFrameProfiled(Request request)
        {
            ProfiledFrameSnapshot plan = ReadFrameProfilePlan(request);
            if (plan.Issues.Count != 0)
                throw new InvalidOperationException("Frame profile copy plan is ineligible: " +
                    string.Join("; ", plan.Issues));
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic fresh = Tree(request.Project, request.Form);
            if (!string.Equals((string)fresh.TreeVersion, plan.TreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed during copy preflight.");
            dynamic rootControls = form.Designer.Controls;
            dynamic copiedFrame = null;
            bool frameCreated = false;
            var createdChildren = new List<string>();
            try
            {
                copiedFrame = rootControls.Add("Forms.Frame.1", plan.NewName, true);
                frameCreated = true;
                copiedFrame.Left = plan.Left;
                copiedFrame.Top = plan.Top;
                copiedFrame.Width = plan.Width;
                copiedFrame.Height = plan.Height;
                copiedFrame.Caption = plan.Caption;
                foreach (ProfiledChildSnapshot item in plan.Children)
                {
                    dynamic copy = copiedFrame.Controls.Add(item.Profile.ProgId, item.ProposedName, true);
                    createdChildren.Add(item.ProposedName);
                    ApplyProfiledChild(copy, item);
                }
                dynamic after = Tree(request.Project, request.Form);
                string framePath = "Controls/" + plan.NewName;
                if (!TreeContainsPath((IEnumerable)after.Controls, framePath) ||
                    (int)after.NodeCount != plan.NodeCount + 1 + plan.Children.Count ||
                    string.Equals((string)after.TreeVersion, plan.TreeVersion, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The new Frame hierarchy was not reflected in form_tree.");
                dynamic installedFrame = ResolveTreeItem(form.Designer, framePath);
                if ((string)installedFrame.Caption != plan.Caption ||
                    Convert.ToInt32(installedFrame.Controls.Count, CultureInfo.InvariantCulture) != plan.Children.Count ||
                    !SameCopyBox(installedFrame, plan.Left, plan.Top, plan.Width, plan.Height))
                    throw new InvalidOperationException("Copied Frame properties differ from the preflight snapshot.");
                foreach (ProfiledChildSnapshot item in plan.Children)
                {
                    if (!TreeContainsPath((IEnumerable)after.Controls, item.ProposedPath))
                        throw new InvalidOperationException("Copied child missing from form_tree: " + item.ProposedPath);
                    VerifyProfiledChild(ResolveTreeItem(form.Designer, item.ProposedPath), item);
                }
                return new { SourcePath = plan.SourcePath, NewPath = framePath,
                    DirectChildrenCopied = plan.Children.Count,
                    CopiedChildren = plan.Children.Select(item => new { Path = item.ProposedPath,
                        Type = item.Profile.Type, CopiedFields = item.Profile.Fields,
                        Limitation = item.Profile.Limitation }).ToArray(),
                    Completeness = "Partial", Tree = after };
            }
            catch (Exception originalError)
            {
                var rollbackErrors = new List<string>();
                if (frameCreated)
                {
                    for (int index = createdChildren.Count - 1; index >= 0; index--)
                        try { copiedFrame.Controls.Remove(createdChildren[index]); }
                        catch (Exception ex) { rollbackErrors.Add(createdChildren[index] + ": " + ex.Message); }
                    try { rootControls.Remove(plan.NewName); }
                    catch (Exception ex) { rollbackErrors.Add(plan.NewName + ": " + ex.Message); }
                }
                if (rollbackErrors.Count > 0)
                    throw new InvalidOperationException("Profiled Frame copy failed and rollback was incomplete. " +
                        string.Join("; ", rollbackErrors) + " Inspect form_tree before retrying.", originalError);
                throw;
            }
        }

        private static void ApplyProfiledChild(dynamic copy, ProfiledChildSnapshot item)
        {
            copy.Left = item.Left;
            copy.Top = item.Top;
            copy.Width = item.Width;
            copy.Height = item.Height;
            switch (item.Profile.Type)
            {
                case "Label":
                    copy.Caption = item.Caption;
                    copy.BackColor = item.BackColor;
                    dynamic font = copy.Font;
                    font.Name = item.FontName;
                    font.Size = item.FontSize;
                    font.Bold = item.FontBold;
                    break;
                case "TextBox": if (item.TextValue != null) copy.Value = item.TextValue; break;
                case "CheckBox":
                    copy.Caption = item.Caption;
                    if (item.BooleanValue) copy.Value = true;
                    break;
                case "CommandButton":
                case "OptionButton": copy.Caption = item.Caption; break;
                case "ComboBox": copy.ListWidth = item.ListWidth; break;
                default: throw new InvalidOperationException("Profile has no setter implementation.");
            }
        }

        private static void VerifyProfiledChild(dynamic actual, ProfiledChildSnapshot item)
        {
            if (!SameCopyBox(actual, item.Left, item.Top, item.Width, item.Height))
                throw new InvalidOperationException("Copied child geometry differs: " + item.ProposedPath);
            switch (item.Profile.Type)
            {
                case "Label":
                    if ((string)actual.Caption != item.Caption ||
                        CopyOleColor((object)actual.BackColor) != item.BackColor ||
                        (string)actual.Font.Name != item.FontName ||
                        Math.Abs(Convert.ToDouble(actual.Font.Size, CultureInfo.InvariantCulture) - item.FontSize) > 0.01 ||
                        (bool)actual.Font.Bold != item.FontBold)
                        throw new InvalidOperationException("Copied Label differs: " + item.ProposedPath);
                    break;
                case "TextBox":
                    if (!string.Equals(item.TextValue, actual.Value as string, StringComparison.Ordinal))
                        throw new InvalidOperationException("Copied TextBox value differs: " + item.ProposedPath);
                    break;
                case "CheckBox":
                    if ((string)actual.Caption != item.Caption ||
                        !(actual.Value is bool) || (bool)actual.Value != item.BooleanValue)
                        throw new InvalidOperationException("Copied CheckBox differs: " + item.ProposedPath);
                    break;
                case "CommandButton":
                case "OptionButton":
                    if ((string)actual.Caption != item.Caption)
                        throw new InvalidOperationException("Copied button caption differs: " + item.ProposedPath);
                    break;
                case "ComboBox":
                    if (!string.Equals(item.ListWidth, actual.ListWidth as string, StringComparison.Ordinal))
                        throw new InvalidOperationException("Copied ComboBox ListWidth differs: " + item.ProposedPath);
                    break;
                default: throw new InvalidOperationException("Profile has no readback implementation.");
            }
        }
    }
}
