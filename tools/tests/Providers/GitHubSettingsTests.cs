using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CodexVBE;

internal static partial class ProviderTests
{
    private static void GitHubSettingsUi()
    {
        Assert(GitHubAccountService.ParseAccounts("alice\r\nbob\r\nalice\r\n").Length == 2, "Account listing deduplication");
        bool rejected = false;
        try { GitHubAccountService.ParseAccounts("https://secret@github.com"); } catch (InvalidOperationException) { rejected = true; }
        Assert(rejected && !GitHubAccountService.ValidAccount("user\nhelper=other"), "Unexpected account output rejected");
        var settings = new LlmSettings { ProviderName = "Claude", GitHubAccount = "alice" };
        var flags = BindingFlags.Instance | BindingFlags.NonPublic;
        int logins = 0; string accounts = "alice\nbob\n";
        using (var form = new LlmSettingsWindow(settings))
        {
            typeof(LlmSettingsWindow).GetField("githubService", flags).SetValue(form, new GitHubAccountService((command, token) => {
                token.ThrowIfCancellationRequested();
                if (command.Contains(" login ")) { Assert(command.Contains("--browser") && !command.Contains("--pat"), "Browser authentication without token argument"); logins++; return Task.FromResult(""); }
                Assert(command.Contains(" list ") && command.Contains("--no-ui"), "Account refresh must be noninteractive");
                return Task.FromResult(accounts);
            }));
            Func<string, object> field = name => typeof(LlmSettingsWindow).GetField(name, flags).GetValue(form);
            Action<bool> refresh = login => ((Task)typeof(LlmSettingsWindow).GetMethod("RefreshGitHubAsync", flags).Invoke(form, new object[] { login })).GetAwaiter().GetResult();
            refresh(false);
            var picker = (ComboBox)field("githubAccount");
            Assert((string)picker.SelectedItem == "alice" && logins == 0, "Refresh preserves preferred account without login");
            picker.SelectedItem = "bob";
            ((ComboBox)field("provider")).SelectedItem = Provider("LM Studio");
            Assert((string)picker.SelectedItem == "bob", "GitHub account independent of AI provider");
            refresh(true); Assert(logins == 1 && (string)picker.SelectedItem == "bob", "Explicit browser login and refresh");
            Assert(((Label)field("githubStatus")).Text.StartsWith(UiText.Get("GitHub sign-in completed. "), StringComparison.Ordinal),
                "A completed browser flow must show an explicit confirmation in Settings");
            accounts = "alice\n"; refresh(false);
            Assert(((Label)field("githubStatus")).Text == UiText.Get("The selected account is no longer saved. Sign in again or choose another account."), "Missing preferred account not silently switched");
            accounts = "alice\nbob\n";
            form.Show(); Application.DoEvents(); form.Refresh();
            Assert(picker.Visible && ((Button)field("githubLogin")).Visible, "GitHub section visible for local AI provider");
            var save = (Button)field("saveButton");
            Assert(save.Visible && form.RectangleToScreen(form.ClientRectangle).Contains(save.PointToScreen(new System.Drawing.Point(0, save.Height))), "Save button accessible after GitHub layout");
            using (var bitmap = new System.Drawing.Bitmap(form.Width, form.Height))
            {
                form.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0, 0, form.Width, form.Height));
                bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "github-settings.png"));
            }
            form.Close(); form.Dispose();
        }
        Assert(settings.GitHubAccount == "alice", "Cancel must not persist the selected account");
        var roundTrip = Json.Deserialize<LlmSettings>(Json.Serialize(settings));
        Assert(roundTrip.GitHubAccount == "alice", "Account selection survives settings serialization");
        SynchronizationContext.SetSynchronizationContext(null);
        Console.WriteLine("PASS GitHub settings: provider independence, browser login command, noninteractive listing, stale account, cancellation and persistence");
    }
}
