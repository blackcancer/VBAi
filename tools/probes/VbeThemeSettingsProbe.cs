using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;

public static class VbeThemeSettingsProbe
{
    private static Thread worker;
    public static string Error;
    public static bool Saved;

    public static void Begin(int processId, bool enabled)
    {
        if (worker != null && worker.IsAlive) throw new InvalidOperationException("The previous settings worker is still running.");
        Error = null;
        Saved = false;
        worker = new Thread(() => Run(processId, enabled)) { IsBackground = true };
        worker.SetApartmentState(ApartmentState.MTA);
        worker.Start();
    }

    public static bool Wait() { return worker.Join(15000); }

    private static AutomationElement ById(AutomationElement root, string id)
    {
        return root.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
    }

    private static void Run(int processId, bool enabled)
    {
        AutomationElement settings = null;
        try
        {
            var timeout = Stopwatch.StartNew();
            while (settings == null && timeout.ElapsedMilliseconds < 10000)
            {
                foreach (AutomationElement window in AutomationElement.RootElement.FindAll(TreeScope.Children,
                    new PropertyCondition(AutomationElement.ProcessIdProperty, processId)))
                    if (ById(window, "saveButton") != null && ById(window, "cancelButton") != null)
                    { settings = window; break; }
                if (settings == null) Thread.Sleep(100);
            }
            if (settings == null) throw new InvalidOperationException("The add-in settings window was not found.");
            AutomationElement check = null;
            foreach (AutomationElement tab in settings.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem)))
            {
                ((SelectionItemPattern)tab.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
                check = ById(settings, "nativeVbeDark");
                if (check != null && !check.Current.IsOffscreen) break;
                check = null;
            }
            if (check == null) throw new InvalidOperationException("The native theme checkbox is not visible.");
            var toggle = (TogglePattern)check.GetCurrentPattern(TogglePattern.Pattern);
            ToggleState expected = enabled ? ToggleState.On : ToggleState.Off;
            if (toggle.Current.ToggleState != expected) toggle.Toggle();
            if (toggle.Current.ToggleState != expected) throw new InvalidOperationException("The native theme checkbox did not change.");
            ((InvokePattern)ById(settings, "saveButton").GetCurrentPattern(InvokePattern.Pattern)).Invoke();
            Saved = true;
        }
        catch (Exception error)
        {
            Error = error.ToString();
            if (settings != null)
                try { ((InvokePattern)ById(settings, "cancelButton").GetCurrentPattern(InvokePattern.Pattern)).Invoke(); } catch { }
        }
    }
}
