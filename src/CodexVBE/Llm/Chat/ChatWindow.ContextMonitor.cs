using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Fenêtre du chat et coordination des mises à jour de contexte VBE.</summary>
internal sealed partial class ChatWindow
    {
        /// <summary>Cadence les lectures incrémentales du contexte de projet.</summary>
private Timer contextMonitorTimer;
        /// <summary>Surveille les empreintes des projets, modules et références.</summary>
private VbeContextMonitor contextMonitor;
        /// <summary>Indique qu’un événement ou un sondage impose de rafraîchir le contexte affiché.</summary>
private bool contextDirty;
        /// <summary>Abonnement aux événements d’ajout et retrait de références.</summary>
private VbeReferenceEvents referenceEvents;
        /// <summary>Abonnements aux modifications de projets et de composants.</summary>
private VbeCollectionEvents projectEvents, componentEvents;

        /// <summary>Installe les observateurs COM et le sondage périodique sur la session VBE.</summary>
        /// <param name="session">Session hôte dont les événements et requêtes sont surveillés.</param>
private void InitializeContextMonitor(VbeSession session)
        {
            projectEvents = new VbeCollectionEvents(() => contextDirty = true, false);
            componentEvents = new VbeCollectionEvents(() => contextDirty = true, true);
            referenceEvents = new VbeReferenceEvents(() => contextDirty = true);
            contextMonitor = new VbeContextMonitor(request => ReadHost(session, request));
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

        /// <summary>Actualise la collection COM observée et maintient le sondage si les événements natifs sont indisponibles.</summary>
        /// <param name="observer">Observateur à reconfigurer.</param><param name="source">Accès différé à la collection source.</param>
private static void ObserveCollectionEvents(VbeCollectionEvents observer, Func<object> source)
        {
            try { observer.Observe(source()); }
            catch (Exception error)
            {
                observer.Observe(null);
                LoadLog.Write("Collection events unavailable; polling continues: " + error.Message);
            }
        }

        /// <summary>Réconcilie le sélecteur de projet et invalide l’index des références lorsqu’un changement est en attente.</summary>
        /// <param name="session">Session VBE utilisée pour relire les projets.</param>
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

        /// <summary>Reconstruit les choix de projet à partir de l’hôte sans basculer implicitement la conversation.</summary>
        /// <param name="session">Session VBE interrogée.</param>
        /// <returns><see langword="true"/> si la liste a pu être relue et actualisée.</returns>
private bool RefreshAvailableScopes(VbeSession session)
        {
            Response response = ReadHost(session, new Request { Command = "list_projects" });
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
