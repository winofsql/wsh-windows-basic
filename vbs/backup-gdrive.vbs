strName = "myfolder"
strTarget = "C:\app\web22"
strBackupFolder = "G:\�}�C�h���C�u\backup"
if MsgBox( strTarget & vbCrLf & "�̃o�b�N�A�b�v���J�n���܂��B��낵���ł���? (�ۑ���:" & strBackupFolder & "\" & strName & ")", 1 ) = 2 then
	Wscript.Quit
end if
Set WshShell = Wscript.CreateObject( "WScript.Shell" )
ExecCommand = "cmd.exe /C ""robocopy """ & strTarget & """ """ & strBackupFolder & "\" & strName & """ /E /XO /NP /R:3 /W:3 /XD dir C:\app\system\log /LOG:%TEMP%\robo.log & PAUSE"""

Call WshShell.Run( ExecCommand )

