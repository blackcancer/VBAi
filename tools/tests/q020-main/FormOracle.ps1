# Qualification-only paired projection; raw trees and revisions remain unchanged.
function Get-Q020CrossGenerationFormPair($leftTree,$rightTree) {
 $allow=ConvertFrom-Json -InputObject @'
{"Root/BorderStyle": {"Family": "fmBorderStyle", "Choices": ["None", "Single"]}, "Root/Cycle": {"Family": "fmCycle", "Choices": ["AllForms", "CurrentForm"]}, "Root/KeepScrollBarsVisible": {"Family": "fmScrollBars", "Choices": ["None", "Horizontal", "Vertical", "Both"]}, "Root/MousePointer": {"Family": "fmMousePointer", "Choices": ["Default", "Arrow", "Cross", "IBeam", "SizeNESW", "SizeNS", "SizeNWSE", "SizeWE", "UpArrow", "HourGlass", "NoDrop", "AppStarting", "Help", "SizeAll", "Custom"]}, "Root/PictureAlignment": {"Family": "fmPictureAlignment", "Choices": ["TopLeft", "TopRight", "Center", "BottomLeft", "BottomRight"]}, "Root/PictureSizeMode": {"Family": "fmPictureSizeMode", "Choices": ["Clip", "Stretch", "Zoom"]}, "Root/ScrollBars": {"Family": "fmScrollBars", "Choices": ["None", "Horizontal", "Vertical", "Both"]}, "Root/SpecialEffect": {"Family": "fmSpecialEffect", "Choices": ["Flat", "Raised", "Sunken", "Etched", "Bump"]}, "Root/VerticalScrollBarSide": {"Family": "fmVerticalScrollBarSide", "Choices": ["Right", "Left"]}, "Root/DesignMode": {"Family": "fmMode", "Choices": ["Off", "Inherit", "On"]}, "Root/ShowToolbox": {"Family": "fmMode", "Choices": ["Off", "Inherit", "On"]}, "Root/ShowGridDots": {"Family": "fmMode", "Choices": ["Off", "Inherit", "On"]}, "Root/SnapToGrid": {"Family": "fmMode", "Choices": ["Off", "Inherit", "On"]}, "Control/LayoutEffect": {"Family": "fmLayoutEffect", "Choices": ["None", "Initiate", "Respond"]}, "Control/BackStyle": {"Family": "fmBackStyle", "Choices": ["Transparent", "Opaque"]}, "Control/BorderStyle": {"Family": "fmBorderStyle", "Choices": ["None", "Single"]}, "Control/MousePointer": {"Family": "fmMousePointer", "Choices": ["Default", "Arrow", "Cross", "IBeam", "SizeNESW", "SizeNS", "SizeNWSE", "SizeWE", "UpArrow", "HourGlass", "NoDrop", "AppStarting", "Help", "SizeAll", "Custom"]}, "Control/PicturePosition": {"Family": "fmPicturePosition", "Choices": ["LeftTop", "LeftCenter", "LeftBottom", "RightTop", "RightCenter", "RightBottom", "AboveLeft", "AboveCenter", "AboveRight", "BelowLeft", "BelowCenter", "BelowRight", "Center"]}, "Control/SpecialEffect": {"Family": "fmSpecialEffect", "Choices": ["Flat", "Raised", "Sunken", "Etched", "Bump"]}, "Control/TextAlign": {"Family": "fmTextAlign", "Choices": ["Left", "Center", "Right"]}}
'@
 $left=ConvertFrom-Json -InputObject (ConvertTo-Json -InputObject $leftTree -Depth 80 -Compress)
 $right=ConvertFrom-Json -InputObject (ConvertTo-Json -InputObject $rightTree -Depth 80 -Compress)
 # Existing ROOT5 persisted-default proof remains exclusively in the original oracle.
 $root5=@('HelpContextID','ShowModal','WhatsThisButton','WhatsThisHelp','Visible')
 $groups=@(@{Scope='Root';Left=@($left.Properties|Where-Object Name -cnotin $root5|Sort-Object Name -CaseSensitive);Right=@($right.Properties|Where-Object Name -cnotin $root5|Sort-Object Name -CaseSensitive)})
 if(@($left.Controls).Count -ne @($right.Controls).Count){throw 'Ordered control count differs.'}
 for($i=0;$i -lt @($left.Controls).Count;$i++){
  if($left.Controls[$i].Name -cne $right.Controls[$i].Name){throw 'Ordered control identity differs.'}
  $groups+=@{Scope='Control';Left=@($left.Controls[$i].Properties);Right=@($right.Controls[$i].Properties)}
 }
 $changes=0
 foreach($g in $groups){
  if($g.Left.Count -ne $g.Right.Count){throw 'Ordered property count differs.'}
  for($i=0;$i -lt $g.Left.Count;$i++){
   $x=$g.Left[$i];$y=$g.Right[$i]
   if($x.Name -cne $y.Name){throw 'Ordered property identity differs.'}
   if($x.Type -ceq $y.Type){continue}
   $entry=$allow.PSObject.Properties[$g.Scope+'/'+$x.Name]
   if($null -eq $entry){throw 'Unknown differing COM enum property.'}
   $spec=$entry.Value;$pattern='\A-?[0-9]+_'+[regex]::Escape($spec.Family)+'\z'
   if($x.Type -cnotmatch $pattern -or $y.Type -cnotmatch $pattern){throw 'Unknown COM enum type family.'}
   foreach($r in @($x,$y)){
    if($r.Kind -cne 'scalar' -or $null -ne $r.Error){throw 'Unreadable/non-scalar enum descriptor.'}
    if((ConvertTo-Json -InputObject @($r.AllowedValues) -Compress) -cne (ConvertTo-Json -InputObject @($spec.Choices) -Compress)){throw 'Documented ordered enum choices differ.'}
    $v=$r.Value
    $integer=($v -is [int] -or $v -is [long]) -and $v -ge [int]::MinValue -and $v -le [int]::MaxValue
    $observedNull=$g.Scope -ceq 'Control' -and $r.Name -ceq 'LayoutEffect' -and $spec.Family -ceq 'fmLayoutEffect' -and $null -eq $v -and $r.ReadOnly -eq $true -and $r.SetterStatus -ceq 'DescriptorReadOnly' -and $r.NumIndices -eq 0 -and $null -eq $r.Display -and $null -eq $r.Digest -and $null -eq $r.Members
    if(-not ($integer -or ($v -is [string] -and $v -cin @($spec.Choices)) -or $observedNull)){throw 'Unsupported enum scalar representation.'}
   }
   $x.Type=$spec.Family;$y.Type=$spec.Family
   if((ConvertTo-Json -InputObject $x -Depth 80 -Compress) -cne (ConvertTo-Json -InputObject $y -Depth 80 -Compress)){throw 'Enum descriptor content differs beyond generated type identity.'}
   $changes++
  }
 }
 return @{Left=$left;Right=$right;ProjectedDescriptorCount=$changes}
}

# Independent qualification oracle. No product assembly/helper is loaded or consulted.
function Require-FormOracle($ok,[string]$message) { if(-not $ok){throw $message} }
function Assert-FormReadable($value,[int]$depth=0) {
    Require-FormOracle ($depth -le 80) 'Designer inventory depth bound exceeded.'
    if($null -eq $value -or $value -is [string] -or $value -is [ValueType]){return}
    if($value -is [Collections.IDictionary]){$members=@($value.Keys|ForEach-Object {@{Name=$_;Value=$value[$_]}})}
    elseif($value -is [Collections.IEnumerable]){foreach($item in $value){Assert-FormReadable $item ($depth+1)};return}
    else{$members=@($value.PSObject.Properties)}
    foreach($member in $members){
        if($member.Name -cin @('Error','ReadError')){Require-FormOracle ($null -eq $member.Value) 'Recursive designer getter error cannot be projected away.'}
        Assert-FormReadable $member.Value ($depth+1)
    }
}
function Get-CanonicalExportedForm([string]$path,[string]$name) {
    $bytes=[IO.File]::ReadAllBytes($path)
    Require-FormOracle ($bytes.Length -gt 0 -and $bytes.Length -le 16777216) 'Bounded complete FRM export required.'
    $encoding=[Text.Encoding]::GetEncoding(28591) # one byte per character, never a lossy decoded header
    $text=$encoding.GetString($bytes)
    $marker=[regex]::Match($text,'(?m)^Attribute VB_Name\s*=')
    Require-FormOracle $marker.Success 'Exact exported FRM attribute boundary required.'
    $header=$text.Substring(0,$marker.Index)
    Require-FormOracle ([regex]::IsMatch($header,'\AVERSION 5\.00\r\nBegin \{C62A69F0-16DC-11CE-9E98-00AA00574A4F\} '+[regex]::Escape($name)+' *\r\n')) 'Exact typed owned FRM header required.'
    $blobs=[regex]::Matches($header,'(?m)^\s*OleObjectBlob\s*=\s*"([^"\r\n]+)":([0-9A-Fa-f]+)[^\r\n]*\r?$')
    Require-FormOracle ($blobs.Count -eq 1) 'Unique FRX binding required.'
    $frx=[IO.Path]::ChangeExtension($path,'.frx');$filename=$blobs[0].Groups[1].Value
    Require-FormOracle ($filename -ceq [IO.Path]::GetFileName($frx) -and [IO.Path]::GetFileName($filename) -ceq $filename -and
        [IO.File]::Exists($frx) -and (Get-Item -LiteralPath $frx).Length -gt 0) 'Exact owned nonempty FRX companion required.'
    $match=$blobs[0].Groups[1]
    $canonical=$header.Substring(0,$match.Index)+'$owned-frx'+$header.Substring($match.Index+$match.Length)
    $values=[ordered]@{};$defaults=[ordered]@{HelpContextID=0;ShowModal=$true;WhatsThisButton=$false;WhatsThisHelp=$false;Visible=$true}
    foreach($key in $defaults.Keys){
        $hits=[regex]::Matches($header,'(?m)^\s*'+[regex]::Escape($key)+'\s*=\s*([^\r\n]*)\r?$')
        Require-FormOracle ($hits.Count -le 1) ('Duplicate root setting: '+$key)
        $value=$defaults[$key]
        if($hits.Count){
            $token=($hits[0].Groups[1].Value -split "'",2)[0].Trim()
            if($key -ceq 'HelpContextID'){$number=0;Require-FormOracle ([int]::TryParse($token,[Globalization.NumberStyles]::None,[Globalization.CultureInfo]::InvariantCulture,[ref]$number) -and $number -ge 0) 'Invalid persisted HelpContextID.';$value=$number}
            else{Require-FormOracle ($token -cin @('-1','0','True','False')) ('Invalid persisted Boolean: '+$key);$value=$token -cin @('-1','True')}
        }
        $values[$key]=$value
    }
    return @{ExportFrmPath=$path;FormHeaderCanonical=$canonical;MetadataDefaults5Values=$values;DocumentedDefaults=$defaults;
        HeaderByteEncoding='ISO-8859-1 bijective mapping; only proved owned OleObjectBlob filename replaced';ResourcePath=$frx}
}
function Get-IndependentFormRoot($tree,$export) {
    Require-FormOracle ($null -ne $tree -and $null -ne $export) 'Tree plus independently parsed FRM required.'
    $seen=@{};$stable=[ordered]@{};$rows=@($tree.Properties)
    foreach($row in $rows){
        $name=[string]$row.Name;Require-FormOracle ($name -and -not $seen.ContainsKey($name)) 'Missing/duplicate root property.';$seen[$name]=$true
        Require-FormOracle ($null -eq $row.Error) ('Unreadable root property: '+$name)
    }
    foreach($key in $export.DocumentedDefaults.Keys){
        $found=@($rows|Where-Object Name -ceq $key);$stored=$export.MetadataDefaults5Values[$key];$default=$export.DocumentedDefaults[$key]
        if($found.Count){
            $r=$found[0];$typed=$(if($key -ceq 'HelpContextID'){$r.Value -is [int]}else{$r.Value -is [bool]})
            Require-FormOracle ($r.Kind -ceq 'scalar' -and $typed -and $r.Value -ceq $stored) ('Getter and persisted root setting differ: '+$key)
        }else{Require-FormOracle ($stored -ceq $default) ('Missing nondefault persisted root getter: '+$key)}
        $stable[$key]=($stored|ConvertTo-Json -Compress)
    }
    # The three live clipboard/undo capabilities are not persisted designer content. Their raw values remain in receipts.
    foreach($row in @($rows|Sort-Object Name -CaseSensitive)){
        if($row.Name -cin @('HelpContextID','ShowModal','WhatsThisButton','WhatsThisHelp','Visible','CanPaste','CanUndo','CanRedo')){continue}
        $stable[$row.Name]=(ConvertTo-Json -InputObject $row -Depth 80 -Compress)
    }
    return $stable
}
function Assert-IndependentFormContent($left,$right) {
    Require-FormOracle ($null -ne $left.Form -and $null -ne $right.Form -and $null -ne $left.FormTree -and $null -ne $right.FormTree) 'Complete form_state and form_tree snapshots required.'
    $pair=Get-Q020CrossGenerationFormPair $left.FormTree $right.FormTree
    $leftTree=$pair.Left;$rightTree=$pair.Right
    Assert-FormReadable $leftTree;Assert-FormReadable $rightTree
    Require-FormOracle ($left.FormExport.FormHeaderCanonical -ceq $right.FormExport.FormHeaderCanonical) 'Persisted FRM header byte content differs.'
    $a=Get-IndependentFormRoot $leftTree $left.FormExport;$b=Get-IndependentFormRoot $rightTree $right.FormExport
    Require-FormOracle ((ConvertTo-Json -InputObject $a -Depth 80 -Compress) -ceq (ConvertTo-Json -InputObject $b -Depth 80 -Compress)) 'Complete stable root property content differs.'
    Require-FormOracle ((ConvertTo-Json -InputObject $leftTree.Controls -Depth 80 -Compress) -ceq (ConvertTo-Json -InputObject $rightTree.Controls -Depth 80 -Compress)) 'Exact ordered control/property/font/picture tree differs.'
    Require-FormOracle ($left.PictureDigest -ceq $right.PictureDigest) 'Real OLE picture content digest differs.'
    foreach($key in @('Form','Caption','Width','Height','Controls')){Require-FormOracle ($null -ne $left.Form.PSObject.Properties[$key] -and $null -ne $right.Form.PSObject.Properties[$key]) ('Required form_state field missing: '+$key)}
    foreach($key in @('Form','Caption','Width','Height')){Require-FormOracle ($left.Form.PSObject.Properties[$key].Value -ceq $right.Form.PSObject.Properties[$key].Value) ('Form state differs: '+$key)}
    Require-FormOracle ((ConvertTo-Json -InputObject $left.Form.Controls -Depth 80 -Compress) -ceq (ConvertTo-Json -InputObject $right.Form.Controls -Depth 80 -Compress)) 'Exact form-state control summary differs.'
    # Raw source-local FormVersion/TreeVersion remain stored and are still used for same-project concurrency guards.
}
