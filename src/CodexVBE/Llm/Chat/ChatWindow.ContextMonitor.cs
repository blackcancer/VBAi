using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private Timer contextMonitorTimer;
        private VbeContextMonitor contextMonitor;
        private bool contextDirty;
        private VbeReferenceEvents referenceEvents;
        private VbeCollectionEvents projectEvents, componentEvents;

        private void InitializeContextMonitor(VbeSession session)
        {
            projectEvents = new VbeCollectionEvents(() => contextDirty = true, false);
            componentEvents = new VbeCollectionEvents(() => contextDirty = true, true);
            referenceEvents = new VbeReferenceEvents(() => contextDirty = true);
            contextMonitor = new VbeContextMonitor(session.Execute);
            contextMonitor.Changed += () => contextDirty = true;
            contextMonitorTimer = new Timer(components) { Interval = 1500 };
            contextMonitorTimer.Tick += (sender, args) => {
                if (busy || loadingSession || !Visible || IsDisposed) return;
                try
                {
                    string project = (scopePicker.SelectedItem as MacroScope)?.Project;
                    ObserveCollectionEvents(projectEvents, session.ProjectsEventSource);
                    ObserveCollectionEvents(componentEvents, () => session.ComponentsEventSource(project));
                    try { referenceEvents.Observe(session.ReferenceEventSource(project)); }
                    catch (Exception error)
                    {
                        referenceEvents.Observe(null);
                        LoadLog.Write("Reference events unavailable; polling continues: " + error.Message);
                    }
                    // A project-removed event must refresh the scope before polling
                    // its old selector; otherwise that failed read starves refresh forever.
                    RefreshDirtyContext(session);
                    contextMonitor.Step((scopePicker.SelectedItem as MacroScope)?.Project);
                    RefreshDirtyContext(session);
                }
                catch (Exception error) { LoadLog.Write("Context refresh deferred: " + error.Message); }
            };
            contextMonitorTimer.Start();
        }

        private static void ObserveCollectionEvents(VbeCollectionEvents observer, Func<object> source)
        {
            try { observer.Observe(source()); }
            catch (Exception error)
            {
                observer.Observe(null);
                LoadLog.Write("Collection events unavailable; polling continues: " + error.Message);
            }
        }

        private void RefreshDirtyContext(VbeSession session)
        {
            if (!contextDirty) return;
            contextDirty = false;
            try
            {
                if (!RefreshAvailableScopes(session)) { contextDirty = true; return; }
                referenceIndexReady = false;
                if (referencePopup != null && referencePopup.IsOpen) UpdateReferences();
            }
            catch { contextDirty = true; throw; }
        }

        private bool RefreshAvailableScopes(VbeSession session)
        {
            Response response = session.Execute(new Request { Command = "list_projects" });
            if (!response.Ok) return false;
            var entries = json.DeserializeObject(json.Serialize(response.Data)) as object[];
            if (entries == null) return false;
            var selected = scopePicker.SelectedItem as MacroScope;
            var old = scopePicker.Items.Cast<MacroScope>().ToArray();
            var updated = new List<MacroScope>();
            foreach (var raw in entries)
            {
                var fields = raw as IDictionary<string, object>;
                if (fields == null) continue;
                string name = Convert.ToString(fields["Name"]), path = Convert.ToString(fields["FileName"]);
                bool saved = !string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path);
                string project = saved ? Path.GetFullPath(path) : name;
                var previous = old.FirstOrDefault(x => string.Equals(x.Project, project, StringComparison.OrdinalIgnoreCase) && !updated.Contains(x));
                if (previous != null)
                {
                    previous.Name = name;
                    previous.Label = name + " · " + (saved ? Path.GetFileName(path) : UiText.Get("unsaved document"));
                }
                updated.Add(previous ?? new MacroScope { Project = project, Name = name,
                    Key = saved ? project.ToUpperInvariant() : "temporary:" + Guid.NewGuid().ToString("N"),
                    Label = name + " · " + (saved ? Path.GetFileName(path) : UiText.Get("unsaved document")) });
            }
            loadingSession = true;
            try
            {
                scopePicker.Items.Clear(); scopePicker.Items.AddRange(updated.ToArray());
                if (selected != null && updated.Contains(selected)) scopePicker.SelectedItem = selected;
                else scopePicker.SelectedIndex = -1;
                send.Enabled = scopePicker.SelectedItem != null;
            }
            finally { loadingSession = false; }
            // Never switch an existing conversation to a different macro implicitly.
            return true;
        }
    }
}
