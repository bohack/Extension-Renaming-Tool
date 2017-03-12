' Jon Buhagiar
' 03/12/17
' Simple script to rename a group of files matched by extension

Option Explicit
Dim objFSO, objFile

If WScript.Arguments.Count < 3 Then
   Wscript.Echo "Usage:" & Wscript.ScriptName & " Path Old-Extension New-Extension"
   WScript.Quit
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each objFile In objFSO.GetFolder(Wscript.Arguments(0)).Files
  If LCase(objFSO.GetExtensionName(objFile.Name)) = Wscript.Arguments(1) Then
    Wscript.Echo objFile.Name & " -> " & Mid(objFile.Name, 1, Len(objFile.Name) - Len(Wscript.Arguments(1))) & Wscript.Arguments(2)
    objFSO.MoveFile objFile.Path, Mid(objFile.Path, 1, Len(objFile.Path) - Len(Wscript.Arguments(1))) & Wscript.Arguments(2)
  End If
Next

Set objFSO = Nothing