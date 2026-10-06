# One ordered, synthetic help capture batch. Import does not load an assembly,
# show a window, connect an account or launch an Office host.
function Invoke-HelpCapture {
    param([Parameter(Mandatory=$true)][string]$AssemblyPath,
          [Parameter(Mandatory=$true)][string]$OutputDirectory,
          [string]$Culture = 'fr-FR')
    if ($Culture -ne 'fr-FR') { throw 'Validate the French manual before adding help languages.' }
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing,PresentationFramework,WindowsFormsIntegration
    [Windows.Forms.Application]::SetUnhandledExceptionMode([Windows.Forms.UnhandledExceptionMode]::ThrowException)
    if (-not ('HelpCaptureFrame' -as [type])) {
        Add-Type -ReferencedAssemblies System.Windows.Forms,System.Drawing,System.Core -TypeDefinition @'
public sealed class HelpCaptureFrame : System.Windows.Forms.Form {
    protected override bool ShowWithoutActivation { get { return true; } }
    protected override System.Windows.Forms.CreateParams CreateParams {
        get { var value = base.CreateParams; value.ExStyle |= 0x08000000; return value; }
    }
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern System.IntPtr GetForegroundWindow();
    public static System.Delegate CompletedTask(System.Type delegateType, object result) {
        var resultType=delegateType.GetMethod("Invoke").ReturnType.GetGenericArguments()[0];
        var method=typeof(HelpCaptureFrame).GetMethod("Completed").MakeGenericMethod(resultType);
        return System.Linq.Expressions.Expression.Lambda(delegateType,
            System.Linq.Expressions.Expression.Call(method,System.Linq.Expressions.Expression.Constant(result,resultType))).Compile();
    }
    public static System.Threading.Tasks.Task<T> Completed<T>(T result) { return System.Threading.Tasks.Task.FromResult(result); }
}
'@
    }
    $foregroundBefore=[HelpCaptureFrame]::GetForegroundWindow()
    $assembly = Open-UiAssembly $AssemblyPath
    $flags = [Reflection.BindingFlags]'Instance,Public,NonPublic'
    $static = [Reflection.BindingFlags]'Static,Public,NonPublic'
    $originalContext = [ComponentModel.LicenseManager]::CurrentContext
    $designContext = [ComponentModel.Design.DesigntimeLicenseContext]::new()
    $uiText = $assembly.GetType('VBAi.UiText')
    $uiText.GetProperty('Culture',$static).SetValue($null,[Globalization.CultureInfo]::GetCultureInfo($Culture))
    $theme = $assembly.GetType('VBAi.UiTheme')
    $theme.GetField('<Choice>k__BackingField',$static).SetValue($null,[Enum]::Parse($assembly.GetType('VBAi.ThemeChoice'),'Light'))
    $settingsType=$assembly.GetType('VBAi.LlmSettingsWindow')
    $statusObject=[Activator]::CreateInstance($assembly.GetType('VBAi.CodexAccountStatus'),$flags,$null,@($true,'Compte de démonstration · connecté'),$null)
    $statusField=$settingsType.GetField('ReadCodexStatus',$static)
    $originalReadStatus=$statusField.GetValue($null)
    $statusField.SetValue($null,[HelpCaptureFrame]::CompletedTask($statusField.FieldType,$statusObject))
    $directory = [IO.Path]::GetFullPath($OutputDirectory)
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $before = "Option Explicit`r`n`r`nPublic Function CalculerTotal(ByVal derniereLigne As Long) As Double`r`n    Dim i As Long`r`n    For i = 1 To derniereLigne - 1`r`n        CalculerTotal = CalculerTotal + Cells(i, 2).Value`r`n    Next i`r`nEnd Function"
    $after = $before.Replace('derniereLigne - 1','derniereLigne')
    $catalog=New-UiObject $assembly VbaTestCatalog
    $support=$assembly.GetType('VBAi.VbaTestRuntimeSource').GetMethod('Generate',$static).Invoke($null,@($catalog))
    $values = @{
        'remote'='https://github.com/exemple/SuiviBudget'; 'branch'='main'; 'branchName'='corriger-total';
        'commitMessage'='Inclure la dernière ligne dans le total'; 'checkpointName'='Avant correction du total';
        'resolutionText'=$after; 'baseContent'=$before; 'importSummary'="ModuleCalcul.bas : 1 procédure modifiée`r`nClassBudget.cls : inchangée`r`nFrmBudget.frm : inchangé`r`nSauvegarde requise avant application";
        'repositoryName'='SuiviBudget'; 'repositorySearch'='SuiviBudget'; 'pullTitle'='Inclure la dernière ligne dans CalculerTotal';
        'pullBody'="## Modification`r`nLa borne supérieure inclut la dernière ligne.`r`n`r`n## Validation`r`nCompilation et jeu de données de test. Exemple pédagogique.";
        'commentBody'='La signature publique est conservée. Merci de vérifier le cas sans données.';
        'pullDetails'="PR #12 · Inclure la dernière ligne dans CalculerTotal`r`nBranche source : corriger-total → branche cible : main`r`nÉtat : brouillon · auteur fictif`r`n`r`nModification proposée`r`nLa boucle parcourt désormais toutes les lignes. La signature publique est conservée.`r`n`r`nVérification proposée`r`nCompiler le projet, vérifier le total sur trois lignes puis tester une liste vide.`r`n`r`nExemple pédagogique : aucune pull request réelle n'a été créée.";
        'checks'="Exemple pédagogique`r`nCompilation VBA : réussie`r`nTests du projet : 2 réussis";
        'historyDetails'="Exemple pédagogique`r`nInclure la dernière ligne dans le total`r`nModuleCalcul.bas : une borne de boucle modifiée";
        'codexStatus'='Exemple : compte connecté'; 'githubStatus'='Exemple : compte GitHub connecté';
        'sourceLabel'='Branche source : corriger-total'; 'openAiEndpoint'='https://api.openai.com/v1'; 'ollamaEndpoint'='http://localhost:11434';
        'ollamaTemperature'='0'; 'ollamaTopP'='0.8'; 'customName'='Mon provider'; 'manualModels'='modèle-exemple';
        'titleInput'='Le catalogue des modèles reste vide'; 'descriptionInput'="Après sélection du provider, la liste des modèles reste vide.`r`nAttendu : voir les modèles disponibles.`r`nReproduction sur un projet jetable.";
        'preview'="VBAi — Rapport pédagogique`r`nHôte : Microsoft Excel / VBE`r`nAucun code, chemin personnel ni identifiant joint automatiquement.`r`n`r`nLa liste des modèles reste vide après sélection.";
        'reportIdentity'='Rapport pédagogique'; 'sessionTitle'='Fiabiliser le calcul du total'; 'chatTitleEditor'='Fiabiliser le calcul du total';
        'memoryEditor'='Le module ModuleCalcul contient les totaux. Conserver la signature publique.';
        'message'='Inclure la dernière ligne et conserver la signature publique.'; 'detail'='La borne exclut la dernière ligne. Je compare les deux versions avant correction.';
        'text'='ModuleCalcul.CalculerTotal — procédure ajoutée au contexte'; 'content'=$after;
        'module'='ModuleCalcul · CalculerTotal'; 'count'='1 bloc modifié'; 'projectLabel'='SuiviBudget — projet fictif';
        'details'="Projet : SuiviBudget`r`nAction : corriger la borne de boucle de CalculerTotal.`r`nModule : ModuleCalcul`r`nExemple pédagogique : aucune mutation exécutée.";
        'state'='Terminé'; 'status'='Exemple pédagogique · aucune connexion externe';
        'version'='Version installée : exemple de prévisualisation'; 'notes'='Exemple de notes de version : correction de synchronisation et amélioration des messages.';
        'humanReport'="Tests de ModuleCalcul`r`nPASS — TestTotalInclutDerniereLigne`r`nPASS — TestTotalSansDonnees`r`nExemple pédagogique, aucun test VBA exécuté.";
        'compactReport'='{"example":true,"passed":2,"failed":0}';
        'summary'='Exemple · 2 tests · 2 réussites'; 'coverage'='Exemple : couverture des procédures VBA, distincte de la couverture C#.';
        'hostValue'='Microsoft Excel · Visual Basic Editor (exemple)'; 'languageValue'='français (France)'
    }
    $lists = @{
        'provider'=@('Codex'); 'providerPicker'=@('Codex'); 'modelPicker'=@('Modèle de démonstration'); 'effortPicker'=@('Moyen');
        'themePicker'=@('Clair','Système','Sombre');
        'scopePicker'=@('SuiviBudget · Budget.xlsm'); 'githubAccount'=@('Compte de démonstration');
        'repositoryList'=@('exemple/SuiviBudget — privé'); 'repositoryBranch'=@('main','corriger-total'); 'organization'=@('exemple');
        'targetBranch'=@('main'); 'branchList'=@('main','corriger-total'); 'checkpointList'=@('Avant correction du total');
        'changes'=@('ModuleCalcul.bas — modifié');
        'files'=@('ModuleCalcul.bas — une borne de boucle modifiée'); 'checks'=@('Exemple : compilation VBA réussie');
        'comments'=@('Exemple : conserver la signature publique');
        'conflictList'=@('ModuleCalcul.bas'); 'projectList'=@('SuiviBudget (projet lié)','BibliothequeUtilitaire');
        'sessionList'=@('Fiabiliser le calcul du total','Créer le formulaire de saisie'); 'targets'=@('ModuleCalcul.CalculerTotal')
        'history'=@('abc1234 · Inclure la dernière ligne dans le total','def5678 · Première version des sources');
        'pulls'=@('#12 · Inclure la dernière ligne · corriger-total → main')
    }
    function Fill-HelpControl($root) {
        if ($root.GetType().Name -eq 'ChatTextContentView') { Invoke-UiMethod $root ShowPlain @($after,$true) }
        if ($root.GetType().Name -eq 'ChatDisclosureView' -and $root.ContentPanel.Controls.Count -eq 0) {
            $root.Expanded=$true
            $body=New-UiObject $assembly ChatTextContentView
            Invoke-UiMethod $body ShowPlain @('Exemple : lire ModuleCalcul, comparer la borne de boucle, vérifier le résultat.',$false)
            $root.ContentPanel.Controls.Add($body)
        }
        if ($root.GetType().Name -eq 'CodeDiffView') { $root.ShowDiff($before,$after) }
        if ($root.GetType().Name -eq 'ChatActivityGroupView') {
            $section=Get-UiField $root section
            if (-not @($section.ContentPanel.Controls | Where-Object { $_.GetType().Name -eq 'ChatActivityStepView' }).Count) {
                foreach ($old in @($section.ContentPanel.Controls)) { $section.ContentPanel.Controls.Remove($old); $old.Dispose() }
                foreach ($title in @('Lire ModuleCalcul.CalculerTotal','Comparer les bornes de la boucle')) {
                    $step=New-UiObject $assembly ChatActivityStepView
                    (Get-UiField $step section).Title=$title
                    $step.Dock=[Windows.Forms.DockStyle]::Top
                    $section.ContentPanel.Controls.Add($step)
                    Fill-HelpControl $step
                }
            }
            $section.Expanded=$true
        }
        foreach ($field in $root.GetType().GetFields($flags)) {
            if (-not [Windows.Forms.Control].IsAssignableFrom($field.FieldType)) { continue }
            $control = $field.GetValue($root)
            if ($null -eq $control) { continue }
            if ($control.GetType().Name -eq 'CodeDiffView') { $control.ShowDiff($before,$after); continue }
            if ($values.ContainsKey($field.Name)) {
                if ($control -is [Windows.Forms.TextBoxBase] -or $control -is [Windows.Forms.Label]) { $control.Text=$values[$field.Name] }
            }
            if ($lists.ContainsKey($field.Name) -and ($control -is [Windows.Forms.ComboBox] -or $control -is [Windows.Forms.ListBox])) {
                if ($null -eq $control.DataSource -and $control.Items.Count -eq 0) {
                    foreach ($item in $lists[$field.Name]) { $control.Items.Add($item) | Out-Null }
                    if ($control -isnot [Windows.Forms.CheckedListBox]) { $control.SelectedIndex=0 }
                }
                if ($control -is [Windows.Forms.ComboBox] -and $control.Items.Count -gt 0 -and $control.SelectedIndex -lt 0) { $control.SelectedIndex=0 }
            }
            if ($control -is [Windows.Forms.DataGridView] -and -not $control.VirtualMode -and $null -eq $control.DataSource -and $control.Columns.Count -gt 0) {
                $row = [object[]]::new($control.Columns.Count)
                for ($i=0; $i -lt $row.Length; $i++) {
                    $column = $control.Columns[$i]
                    $row[$i] = if ($column -is [Windows.Forms.DataGridViewCheckBoxColumn]) { $true } else { '' }
                    $caption = $column.HeaderText.ToLowerInvariant()
                    if ($caption -match 'file|fichier|module') { $row[$i]='ModuleCalcul.bas' }
                    elseif ($caption -match 'message|title|titre') { $row[$i]='Inclure la dernière ligne' }
                    elseif ($caption -match 'status|state|état|result|résultat') { $row[$i]='Exemple : réussi' }
                    elseif ($caption -match 'author|auteur') { $row[$i]='Auteur fictif' }
                    elseif ($caption -match 'sha|commit|revision|révision') { $row[$i]='abc1234' }
                }
                if ($control.Name -eq 'conflictDiff') {
                    $control.Rows.Clear()
                    $control.Rows.Add([object[]]@('For i = 1 To derniereLigne - 1','For i = 1 To derniereLigne')) | Out-Null
                } elseif ($control.Rows.Count -eq 0) { $control.Rows.Add($row) | Out-Null }
            }
            if ($control -is [Windows.Forms.TreeView] -and $control.Nodes.Count -eq 0) {
                $node=$control.Nodes.Add('ModuleCalcul'); $node.Nodes.Add('TestTotalInclutDerniereLigne — exemple réussi') | Out-Null
                $node.Nodes.Add('TestTotalSansDonnees — exemple réussi') | Out-Null; $control.ExpandAll()
            }
            if ($control.GetType().Name -eq 'ChatTextContentView') { Invoke-UiMethod $control ShowPlain @($after,$true) }
            if ($control.GetType().Name -eq 'ChatDisclosureView') {
                $control.Expanded=$true
                if ($control.ContentPanel.Controls.Count -eq 0) {
                    $body=[Windows.Forms.Label]::new(); $body.AutoSize=$true; $body.MaximumSize=[Drawing.Size]::new(600,0)
                    $body.Text='Exemple : lire le module, comparer la borne de boucle, vérifier le résultat.'
                    $control.ContentPanel.Controls.Add($body)
                }
            }
        }
        # Nested views also sit inside panels and tab pages, not just UserControls.
        foreach ($child in $root.Controls) { Fill-HelpControl $child }
        if ($root.GetType().Name -eq 'ChatActivityStepView') {
            Invoke-UiMethod (Get-UiField $root detail) ShowPlain @('ModuleCalcul.CalculerTotal : lecture de la procédure. La borne actuelle exclut la dernière ligne. Aucune modification dans cette étape fictive.',$false)
        }
        if ($root.GetType().Name -eq 'ChatContextChipView') { (Get-UiField $root open).Text='#ModuleCalcul.CalculerTotal' }
        if ($root.GetType().Name -eq 'ChatLinkView') { (Get-UiField $root link).Text='Ouvrir ModuleCalcul.CalculerTotal' }
    }
    function Wait-HelpTask($task) {
        $watch=[Diagnostics.Stopwatch]::StartNew()
        while (-not $task.IsCompleted -and $watch.Elapsed.TotalSeconds -lt 30) { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 15 }
        if (-not $task.IsCompleted) { throw 'Embedded renderer did not complete within the help capture budget.' }
        $task.GetAwaiter().GetResult()
    }
    $types=@($assembly.GetTypes() | Where-Object { -not $_.IsAbstract -and $_.Namespace -eq 'VBAi' -and
        ([Windows.Forms.Form].IsAssignableFrom($_) -or [Windows.Forms.UserControl].IsAssignableFrom($_)) -and
        $null -ne $_.GetConstructor($flags,$null,[Type[]]@(),$null) } | Sort-Object Name)
    $captures=[Collections.Generic.List[object]]::new(); $failures=[Collections.Generic.List[object]]::new()
    try {
        foreach ($type in $types) {
            # Abstract visual shells are covered through their actual children;
            # embedded WebView2 and WPF surfaces need the renderer-specific bank.
            if ($type.Name -in @('ChatDesignerView','ChatToolWindow')) { continue }
            $view=$null; $frame=$null; $settingsOwner=$null; $chatOwner=$null
            try {
                [ComponentModel.LicenseManager]::CurrentContext=$designContext
                if ($type.Name -in @('LlmSettingsWindow','ProviderSettingsView')) {
                    $settings=New-UiObject $assembly LlmSettings
                    $settings.ProviderName='Codex'
                    $settingsOwner=[Activator]::CreateInstance($settingsType,$flags,$null,@($settings),$null)
                    Set-UiField $settingsOwner githubLoaded $true
                    if ($type.Name -eq 'LlmSettingsWindow') { $view=$settingsOwner; $settingsOwner=$null }
                    else { $view=Get-UiField $settingsOwner providerSettingsView; $view.Parent.Controls.Remove($view) }
                } elseif ($type.Name -eq 'ChatSuggestionsView') {
                    $chatOwner=New-UiObject $assembly ChatWindow
                    Invoke-UiMethod $chatOwner InitializeShell @()
                    Invoke-UiMethod $chatOwner InitializeComposer @($null)
                    $view=Get-UiField $chatOwner referenceView
                    $reference=New-UiObject $assembly VbeChatReference
                    $reference.Project='SuiviBudget'; $reference.Module='ModuleCalcul'; $reference.Name='CalculerTotal'; $reference.Kind='Function'
                    (Get-UiField $view targets).Items.Add($reference) | Out-Null
                } else { $view=[Activator]::CreateInstance($type,$true) }
                if ($type.Name -eq 'ChatWindow') {
                    Invoke-UiMethod $view InitializeShell @()
                    Invoke-UiMethod $view InitializeComposer @($null)
                    Invoke-UiMethod $view InitializeTranscript @()
                    Invoke-UiMethod $view AddTranscriptMessage @('Vous','Explique la borne de boucle de CalculerTotal puis propose la correction.')
                    Invoke-UiMethod $view AddTranscriptMessage @('Assistant','La boucle exclut la dernière ligne car sa borne est derniereLigne - 1. Utilisez derniereLigne pour inclure cette ligne. Compilez puis vérifiez un jeu de données jetable.')
                    $view.Size=[Drawing.Size]::new(780,900)
                }
                if ($type.Name -eq 'ChatInputView') {
                    $editor=$type.GetProperty('Editor',$flags).GetValue($view)
                    $editor.Text='Explique le calcul du total et conserve la signature publique.'
                }
                Fill-HelpControl $view
                if ($type.Name -eq 'ChatWindow') { Invoke-UiMethod $view RefreshModelSummary @() }
                [ComponentModel.LicenseManager]::CurrentContext=$originalContext
                $componentsField=$type.GetField('components',$flags)
                $components=if ($null -ne $componentsField) { $componentsField.GetValue($view) } else { $null }
                $uiText.GetMethod('Apply',$static).Invoke($null,@($view,$components,[Windows.Forms.ToolTip[]]@())) | Out-Null
                $theme.GetMethod('Apply',$static).Invoke($null,@($view)) | Out-Null
                $frame=[HelpCaptureFrame]::new()
                $initialHeight=$view.Height
                $frame.Text=if ($view -is [Windows.Forms.Form]) { $view.Text } else { 'VBAi · exemple pédagogique' }
                if ($view -is [Windows.Forms.Form]) {
                    $view.TopLevel=$false; $view.FormBorderStyle=[Windows.Forms.FormBorderStyle]::None
                    $frame.ClientSize=$view.Size
                } else {
                    $captureWidth=if ($type.Name -in @('CodeDiffView','GitChangesView')) { 1100 } else { [Math]::Max(760,$view.Width) }
                    $frame.ClientSize=[Drawing.Size]::new($captureWidth,[Math]::Max(280,$view.Height))
                }
                $view.Dock=[Windows.Forms.DockStyle]::Fill; $frame.Controls.Add($view)
                $frame.StartPosition=[Windows.Forms.FormStartPosition]::Manual
                $frame.Location=[Drawing.Point]::new(-30000,-30000)
                # Keep runtime OnShown services disabled. The actual child controls
                # still receive native handles and visibility for PrintWindow.
                [ComponentModel.LicenseManager]::CurrentContext=$designContext
                $frame.Show(); $view.Show(); [Windows.Forms.Application]::DoEvents()
                if ($view -is [Windows.Forms.Form]) { $frame.ClientSize=$view.Size }
                Fill-HelpControl $view
                if ($type.Name -eq 'TestSupportReviewDialog') { (Get-UiField $view diff).ShowDiff('',$support) }
                if ($type.Name -eq 'CodeDiffView') { $view.UnifiedDiff=$true }
                if ($type.Name -eq 'GitChangesView') { (Get-UiField $view diff).UnifiedDiff=$true }
                if ($type.Name -eq 'ChatActivityGroupView') {
                    foreach ($step in (Get-UiField $view section).ContentPanel.Controls) { (Get-UiField $step section).Expanded=$false }
                }
                if ($type.Name -eq 'TestExplorerWindow') {
                    (Get-UiField $view details).Text="TestTotalInclutDerniereLigne`r`nAttendu : somme de toutes les lignes.`r`nExemple : résultat conforme.`r`nAucun test VBA exécuté pour cette capture."
                }
                if ($view -isnot [Windows.Forms.Form]) {
                    $view.PerformLayout()
                    $preferred=$view.GetPreferredSize([Drawing.Size]::new($frame.ClientSize.Width,0)).Height
                    $height=if ($type.Name -eq 'ChatChangeCardView') { [Math]::Max(480,$preferred) }
                            elseif ($type.Name -eq 'ChatActivityGroupView') { 220 }
                            elseif ($type.Name -eq 'ChatDisclosureView') { 220 }
                            elseif ($type.Name -eq 'CodeDiffView') { 350 }
                            elseif ($type.Name -eq 'ProviderSettingsView') { [Math]::Max(220,$preferred) }
                            else { [Math]::Max($initialHeight,[Math]::Max(180,[Math]::Max($preferred,$view.MinimumSize.Height))) }
                    $frame.ClientSize=[Drawing.Size]::new($frame.ClientSize.Width,[Math]::Min(1300,$height))
                }
                if ($type.Name -eq 'ModernEditorWindow') {
                    [ComponentModel.LicenseManager]::CurrentContext=$originalContext
                    Wait-HelpTask (Invoke-UiMethod $view InitializeBrowser @())
                    $watch=[Diagnostics.Stopwatch]::StartNew()
                    $ready=$type.GetProperty('Ready',$flags)
                    while (-not $ready.GetValue($view) -and $watch.Elapsed.TotalSeconds -lt 30) { [Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 15 }
                    if (-not $ready.GetValue($view)) { throw ('The actual Monaco renderer is not ready: ' + (Get-UiField $view status).Text) }
                    $browser=$type.GetProperty('Browser',$flags).GetValue($view)
                    foreach ($field in $type.GetFields($flags)) { if ($field.FieldType -eq [Windows.Forms.Timer]) { $field.GetValue($view).Stop() } }
                    $serializer=[Web.Script.Serialization.JavaScriptSerializer]::new()
                    $script='window.vbai.open("help-module",'+$serializer.Serialize($after)+');'
                    Wait-HelpTask ($browser.CoreWebView2.ExecuteScriptAsync($script)) | Out-Null
                    (Get-UiField $view status).Text='Exemple pédagogique · module fictif, aucune synchronisation VBE'
                    [ComponentModel.LicenseManager]::CurrentContext=$designContext
                }
                $frame.PerformLayout(); $view.PerformLayout(); $frame.Refresh()
                [Windows.Forms.Application]::DoEvents()
                $path=Join-Path $directory ($type.Name + '.png')
                Save-UiControlBitmap $frame $path -RenderRichText -RenderHostedSurfaces
                if ($type.Name -eq 'ModernEditorWindow') {
                    $stream=[IO.MemoryStream]::new()
                    try {
                        Wait-HelpTask ($browser.CoreWebView2.CapturePreviewAsync([Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat]::Png,$stream)) | Out-Null
                        $stream.Position=0; $browserImage=[Drawing.Image]::FromStream($stream)
                        $full=[Drawing.Bitmap]::new($path)
                        try {
                            $graphics=[Drawing.Graphics]::FromImage($full)
                            try { $point=$browser.PointToScreen([Drawing.Point]::Empty); $graphics.DrawImageUnscaled($browserImage,$point.X-$frame.Left,$point.Y-$frame.Top) }
                            finally { $graphics.Dispose() }
                            $full.Save((Join-Path $directory 'monaco-complete.png'),[Drawing.Imaging.ImageFormat]::Png)
                        } finally { $full.Dispose(); $browserImage.Dispose() }
                        [IO.File]::Copy((Join-Path $directory 'monaco-complete.png'),$path,$true)
                        [IO.File]::Delete((Join-Path $directory 'monaco-complete.png'))
                    } finally { $stream.Dispose() }
                }
                $probe=[Drawing.Bitmap]::new($path)
                try {
                    $colors=[Collections.Generic.HashSet[int]]::new()
                    for ($x=20; $x -lt $probe.Width-20 -and $colors.Count -lt 24; $x+=2) {
                        for ($y=35; $y -lt $probe.Height-20 -and $colors.Count -lt 24; $y+=2) { $colors.Add($probe.GetPixel($x,$y).ToArgb()) | Out-Null }
                    }
                    if ($colors.Count -lt 4) { throw ('Capture is blank or lacks rendered child controls: ' + $type.Name) }
                } finally { $probe.Dispose() }
                $captures.Add([pscustomobject]@{ id=$type.Name; file=$type.Name+'.png'; type=$type.FullName;
                    caption=$frame.Text; synthetic=$true; culture=$Culture; theme='Light';
                    width=$frame.Width; height=$frame.Height; sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() })
                if ($type.Name -eq 'ChatWindow') {
                    $history=Get-UiField $view historyPanel
                    $history.Bounds=[Drawing.Rectangle]::new($view.Width-345,48,335,720)
                    $history.Visible=$true; $history.BringToFront()
                    (Get-UiField $view memoryPanel).Visible=$true
                    (Get-UiField $view historyLayout).RowStyles[7].Height=226
                    $frame.PerformLayout(); $history.PerformLayout(); $frame.Refresh()
                    [Windows.Forms.Application]::DoEvents()
                    $historyPath=Join-Path $directory 'ChatWindowHistory.png'
                    # Render the actual native history panel separately: a WPF
                    # transcript overlay must not cover its sibling controls.
                    Save-UiControlBitmap $history $historyPath -RenderRichText
                    $captures.Add([pscustomobject]@{ id='ChatWindowHistory'; file='ChatWindowHistory.png'; type=$type.FullName;
                        caption='Historique de conversation et notes locales du document'; synthetic=$true; culture=$Culture; theme='Light';
                        width=$history.Width; height=$history.Height; sha256=(Get-FileHash -LiteralPath $historyPath -Algorithm SHA256).Hash.ToLowerInvariant() })
                }
            } catch { $failures.Add([pscustomobject]@{ type=$type.Name; error=$_.Exception.ToString() }) }
            finally { if ($null -ne $frame) { $frame.Dispose() } elseif ($null -ne $view) { $view.Dispose() }; if ($null -ne $settingsOwner) { $settingsOwner.Dispose() }; if ($null -ne $chatOwner) { $chatOwner.Dispose() } }
        }
    } finally {
        [ComponentModel.LicenseManager]::CurrentContext=$originalContext
        $statusField.SetValue($null,$originalReadStatus)
    }
    [pscustomobject]@{ assembly=[IO.Path]::GetFullPath($AssemblyPath); assemblySha256=(Get-FileHash $AssemblyPath).Hash.ToLowerInvariant();
        capturedUtc=[DateTime]::UtcNow.ToString('o'); renderer='Actual visible WinForms; DrawToBitmap on off-screen non-activating frame';
        foregroundBefore=$foregroundBefore.ToInt64(); foregroundAfter=[HelpCaptureFrame]::GetForegroundWindow().ToInt64();
        captures=@($captures); failures=@($failures) } | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath (Join-Path $directory 'manifest.json') -Encoding UTF8
    Write-Output ("Captured {0} interfaces; {1} failures" -f $captures.Count,$failures.Count)
    if ([HelpCaptureFrame]::GetForegroundWindow() -ne $foregroundBefore) { throw 'Foreground changed during capture; this bank is not accepted as non-activating.' }
    if ($failures.Count) { throw 'Help capture bank failed; inspect manifest.json before rerunning.' }
}
Export-ModuleMember -Function Invoke-HelpCapture
