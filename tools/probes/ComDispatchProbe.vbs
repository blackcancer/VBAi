On Error Resume Next
Set addin = CreateObject("CodexVBE.AddIn")
If Err.Number <> 0 Then
    WScript.Echo "CreateObject failed: " & Hex(Err.Number) & " " & Err.Description
    WScript.Quit 1
End If
WScript.Echo "CreateObject OK; TypeName=" & TypeName(addin)
If Err.Number <> 0 Then
    WScript.Echo "TypeName failed: " & Hex(Err.Number) & " " & Err.Description
    WScript.Quit 1
End If
Dim custom(0)
addin.OnAddInsUpdate custom
If Err.Number <> 0 Then
    WScript.Echo "OnAddInsUpdate failed: " & Hex(Err.Number) & " " & Err.Description
    WScript.Quit 1
End If
WScript.Echo "OnAddInsUpdate OK"
