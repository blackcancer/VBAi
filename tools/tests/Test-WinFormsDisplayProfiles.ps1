param([Parameter(Mandatory=$true)][string]$AssemblyPath,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$assembly=[Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$output=[IO.Path]::GetFullPath($OutputDirectory);[IO.Directory]::CreateDirectory($output)|Out-Null
$rows=@()
foreach($screen in [Windows.Forms.Screen]::AllScreens){
 foreach($name in @('ChatWindow','LlmSettingsWindow','GitWindow','VbeApprovalDialog')){
  $form=[Activator]::CreateInstance($assembly.GetType('CodexVBE.'+$name),$true)
  try{
   $form.StartPosition=[Windows.Forms.FormStartPosition]::Manual
   $form.Location=[Drawing.Point]::new($screen.WorkingArea.Left+20,$screen.WorkingArea.Top+20)
   $form.Show();[Windows.Forms.Application]::DoEvents()
   $actual=[Windows.Forms.Screen]::FromHandle($form.Handle)
   if($actual.DeviceName -ne $screen.DeviceName){throw "$name did not reach $($screen.DeviceName)."}
   $graphics=[Drawing.Graphics]::FromHwnd($form.Handle)
   try{$dpiX=$graphics.DpiX;$dpiY=$graphics.DpiY}finally{$graphics.Dispose()}
   $bitmap=[Drawing.Bitmap]::new($form.Width,$form.Height)
   $file=($screen.DeviceName -replace '[^a-zA-Z0-9]','')+'-'+$name+'.png'
   try{$form.DrawToBitmap($bitmap,[Drawing.Rectangle]::new(0,0,$form.Width,$form.Height));$bitmap.Save((Join-Path $output $file))}finally{$bitmap.Dispose()}
   $rows+=[pscustomobject]@{Display=$screen.DeviceName;Primary=$screen.Primary;WorkingArea=$screen.WorkingArea.ToString();Window=$name;Bounds=$form.Bounds.ToString();DeviceDpi=$form.DeviceDpi;GraphicsDpiX=$dpiX;GraphicsDpiY=$dpiY;CompletelyInsideWorkingArea=$screen.WorkingArea.Contains($form.Bounds);Screenshot=$file;Qualification='Actual detached WinForms display; not native VBE docking or a simulated DPI profile'}
  }finally{$form.Dispose()}
 }
}
$rows|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $output 'display-profiles.json') -Encoding UTF8
$rows|Format-Table Display,Window,DeviceDpi,CompletelyInsideWorkingArea
