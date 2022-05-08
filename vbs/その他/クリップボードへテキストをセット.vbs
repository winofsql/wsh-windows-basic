Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")

TempPath = WshShell.ExpandEnvironmentStrings("%TEMP%") 
TempFile = TempPath & "\___.txt"

Wscript.Echo TempFile

Call WshShell.Run( "cmd /c ipconfig /all > """ & TempFile & """", 0, True )
Call WshShell.Run("cmd /c clip.exe < """ & TempFile & """", 0, True )

on error resume next
Set objHandle = fso.CreateTextFile( TempFile, True )
if Err.Number <> 0 then
    Wscript.Echo Err.Description
else
    objHandle.Write( TempFile )
    objHandle.Close
end if
on error goto 0

Call WshShell.Popup("‚²‹@Œ™‚¢‚©‚ª‚Å‚·‚©?", 2)

Call WshShell.Run("cmd /c clip.exe < """ & TempFile & """", 0, True )
