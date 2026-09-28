param([string]$AssemblyPath = 'artifacts/chat-build/CodexVBE/Debug/net48/CodexVBE.dll')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Design, WindowsFormsIntegration
Add-Type -ReferencedAssemblies System.dll, System.Design.dll -TypeDefinition @'
using System.ComponentModel.Design;
public static class ChatDesignerSelection {
    public static void Select(object service, object component) { ((ISelectionService)service).SetSelectedComponents(new object[] { component }, SelectionTypes.Replace); }
    public static object Primary(object service) { return ((ISelectionService)service).PrimarySelection; }
}
'@
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$type = $assembly.GetType('CodexVBE.ChatWindow', $true)
$flags = [Reflection.BindingFlags]'Instance,NonPublic,Public'
function Field($object, [string]$name) { ,$type.GetField($name,$flags).GetValue($object) }
function Assert($condition,[string]$message) { if (-not $condition) { throw $message } }
$form = [Activator]::CreateInstance($type, $true)
try {
    foreach ($name in @('send','newChat','history','selection','compile','modules','methods','contextToggle')) {
        Assert ((Field $form $name) -is [Windows.Forms.Button]) "$name is not a designer-editable WinForms button."
    }
    foreach ($name in @('scopePicker','providerPicker','modelPicker','effortPicker','modePicker')) {
        Assert ((Field $form $name) -is [Windows.Forms.ComboBox]) "$name is not a designer-editable WinForms picker."
    }
    Assert ((Field $form transcriptHost).Child -eq $null) 'Parameterless constructor started the transcript engine.'
    $input = Field $form promptHost
    Assert ($input.GetType().Name -eq 'ChatInputView' -and $input.Controls.Find('previewEditor', $true).Count -eq 1) 'Designer input representation is missing.'
    $editorField = $input.GetType().GetField('editor', $flags)
    Assert ($editorField.GetValue($input) -eq $null) 'Designer constructor started the input engine.'
    Assert ((Field $form settings) -eq $null -and (Field $form tools) -eq $null -and (Field $form sessionStore) -eq $null) 'Designer constructor started runtime services.'
    $button = Field $form send
    Assert ($button.Parent -ne $null -and (Field $form historyPanel).Parent -ne $null) 'Designer control hierarchy is incomplete.'
    $text = [ComponentModel.TypeDescriptor]::GetProperties($button)['Text']
    Assert (-not $text.IsReadOnly) 'Button caption is read-only in the property grid.'
    $before = $text.GetValue($button)
    $text.SetValue($button, 'Designer edit probe')
    Assert ($button.Text -eq 'Designer edit probe') 'Designer property edit was not applied.'
    $text.SetValue($button, $before)
    $source = [IO.File]::ReadAllText((Join-Path (Get-Location) 'src/CodexVBE/Llm/Chat/ChatWindow.Designer.cs'))
    Assert (-not $source.Contains('InitializeShell') -and -not $source.Contains('BuildWorkflowControls')) 'InitializeComponent calls a runtime UI factory.'
    Assert (-not $source.Contains('Maj+Entrée') -and -not $source.Contains('Entrée : envoyer')) 'Unwanted keyboard annotation remains.'
    Write-Output 'PASS complete WinForms hierarchy, inert designer constructor and removed keyboard annotation'
} finally { $form.Dispose() }
$surface = [ComponentModel.Design.DesignSurface]::new()
try {
    $surface.BeginLoad($type)
    if (-not $surface.IsLoaded -or $surface.LoadErrors.Count -gt 0) { throw ($surface.LoadErrors -join "`n") }
    $designerHost = $surface.GetService([ComponentModel.Design.IDesignerHost])
    Assert ($designerHost.RootComponent -ne $null) 'Designer has no root component.'
    Assert ($designerHost.GetDesigner($designerHost.RootComponent) -ne $null) 'WinForms root designer was not loaded.'
    $view = $surface.View
    Assert ($view -is [Windows.Forms.Control]) 'WinForms designer did not expose its editable surface.'
    $button = Field $designerHost.RootComponent send
    $selection = $surface.GetService([ComponentModel.Design.ISelectionService])
    [ChatDesignerSelection]::Select($selection, $button)
    Assert ([ChatDesignerSelection]::Primary($selection) -eq $button) 'Designer cannot select the send button.'
    Write-Output 'PASS WinForms DesignSurface load and control selection (compiled inherited view)'
} finally { $surface.Dispose() }
$settingsSurface = [ComponentModel.Design.DesignSurface]::new()
try {
    $settingsType = $assembly.GetType('CodexVBE.LlmSettingsWindow', $true)
    $settingsSurface.BeginLoad($settingsType)
    Assert ($settingsSurface.IsLoaded -and $settingsSurface.LoadErrors.Count -eq 0) 'Settings designer failed to load.'
    $settingsHost = $settingsSurface.GetService([ComponentModel.Design.IDesignerHost])
    foreach ($name in @('manualModels','customName','azureEntra','githubAccount','githubLogin','githubRefresh','githubStatus')) {
        $control = $settingsType.GetField($name, $flags).GetValue($settingsHost.RootComponent)
        Assert ($control -is [Windows.Forms.Control] -and $control.Parent -ne $null) "Missing provider designer control: $name"
    }
    Write-Output 'PASS provider settings DesignSurface and fixed provider controls'
} finally { $settingsSurface.Dispose() }
$approvalSurface = [ComponentModel.Design.DesignSurface]::new()
try {
    $approvalType = $assembly.GetType('CodexVBE.VbeApprovalDialog', $true)
    $approvalSurface.BeginLoad($approvalType)
    Assert ($approvalSurface.IsLoaded -and $approvalSurface.LoadErrors.Count -eq 0) 'Approval dialog designer failed to load.'
    $approvalHost = $approvalSurface.GetService([ComponentModel.Design.IDesignerHost])
    foreach ($name in @('details','actions','approve','reject')) {
        $control = $approvalType.GetField($name, $flags).GetValue($approvalHost.RootComponent)
        Assert ($control -is [Windows.Forms.Control] -and $control.Parent -ne $null) "Missing approval designer control: $name"
    }
    Write-Output 'PASS approval dialog DesignSurface and fixed controls'
} finally { $approvalSurface.Dispose() }
$gitSurface = [ComponentModel.Design.DesignSurface]::new()
try {
    $gitType = $assembly.GetType('CodexVBE.GitWindow', $true)
    $gitSurface.BeginLoad($gitType)
    Assert ($gitSurface.IsLoaded -and $gitSurface.LoadErrors.Count -eq 0) 'Git designer failed to load.'
    $gitHost = $gitSurface.GetService([ComponentModel.Design.IDesignerHost])
    Assert ($gitHost.RootComponent -ne $null -and $gitSurface.View -is [Windows.Forms.Control]) 'Git designer has no editable form.'
    Write-Output 'PASS Git DesignSurface load'
} finally { $gitSurface.Dispose() }
