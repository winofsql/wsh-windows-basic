' ***********************************************************
' �����J�n
' ***********************************************************
Set Fso = Wscript.CreateObject( "Scripting.FileSystemObject" )
Set Shell = Wscript.CreateObject( "Shell.Application" )

' ***********************************************************
' ���s���f�B���N�g���̎擾
' ***********************************************************
strPath = Wscript.ScriptFullName 
Set objFile = Fso.GetFile( strPath )
strBackupFolder = Fso.GetParentFolderName( objFile )

' ***********************************************************
' �o�b�N�A�b�v�Ώۃf�B���N�g���̎擾
' ***********************************************************
' �@ �ȗ�����ƁA���[�g�̓f�X�N�g�b�v
Set objFolder = Shell.BrowseForFolder( 0, "�o�b�N�A�b�v����t�H���_��I�����Ă�������", &H4B )

' �A ������ɂ�钼�ڎw��
' strRoot = "c:\"
' Set objFolder = Shell.BrowseForFolder( 0, "�o�b�N�A�b�v����t�H���_��I�����Ă�������", &H4B, strRoot )

' �B ���[�g��ԍ��Ŏw��( ���̏ꍇ�� C:\Users\username\AppData\Local )
' �� ���܂茻���I�ł͂Ȃ�����f�B���N�g���̑I��
' nRoot = &h1c
' Set objFolder = Shell.BrowseForFolder( 0, "�o�b�N�A�b�v����t�H���_��I�����Ă�������", &H4B, nRoot )

if objFolder is nothing then
	WScript.Quit
end if
if not objFolder.Self.IsFileSystem then
	WScript.Echo "�t�@�C���V�X�e���ł͂���܂���"
	WScript.Quit
end if

strTargetFolder = objFolder.Self.Path
strName = Replace( strTargetFolder, ":", "" )
strName = Replace( strName, "\", "_" )
strName = Replace( strName, " ", "" )
strName = "BK_" & strName

' ***********************************************************
' �X�N���v�g�쐬
' ***********************************************************
Set OutFile = Fso.OpenTextFile( strBackupFolder & "\" & strName & ".vbs", 2, True )

OutFile.WriteLine "strName = """ & strName & """"
OutFile.WriteLine "strTarget = """ & strTargetFolder & """"
OutFile.WriteLine "strBackupFolder = """ & strBackupFolder & """"
OutFile.Write "if MsgBox( strTarget & vbCrLf & ""�̃o�b�N�A�b�v���J�n���܂��B��낵���ł���? (�ۑ���:"" & strBackupFolder & ""\"" & strName & "")"""
OutFile.WriteLine ", 1 ) = 2 then"
OutFile.WriteLine "	Wscript.Quit"
OutFile.WriteLine "end if"

OutFile.WriteLine "Set WshShell = Wscript.CreateObject( ""WScript.Shell"" )"
OutFile.Write "ExecCommand = ""cmd.exe /C """"xcopy.exe """""" & strTarget & """""" """""" & strBackupFolder & ""\"" & strName & ""\"""""
OutFile.WriteLine " /D /E /C /S /Y & PAUSE"""""""
OutFile.WriteLine "Call WshShell.Run( ExecCommand )"

OutFile.Close

WScript.Echo "�o�b�N�A�b�v�X�N���v�g : " &  strName & ".vbs" & " ���쐬���܂���"


' ***********************************************************
' �f�B���N�g���w��p�ԍ�
' https://docs.microsoft.com/ja-jp/windows/desktop/api/shldisp/ne-shldisp-shellspecialfolderconstants
' typedef enum {
' 	ssfALTSTARTUP = 0x1d,
' 	ssfAPPDATA = 0x1a,
' 	ssfBITBUCKET = 0x0a,
' 	ssfCOMMONALTSTARTUP = 0x1e,
' 	ssfCOMMONAPPDATA = 0x23,
' 	ssfCOMMONDESKTOPDIR = 0x19,
' 	ssfCOMMONFAVORITES = 0x1f,
' 	ssfCOMMONPROGRAMS = 0x17,
' 	ssfCOMMONSTARTMENU = 0x16,
' 	ssfCOMMONSTARTUP = 0x18,
' 	ssfCONTROLS = 0x03,
' 	ssfCOOKIES = 0x21,
' 	ssfDESKTOP = 0x00,
' 	ssfDESKTOPDIRECTORY = 0x10,
' 	ssfDRIVES = 0x11,
' 	ssfFAVORITES = 0x06,
' 	ssfFONTS = 0x14,
' 	ssfHISTORY = 0x22,
' 	ssfINTERNETCACHE = 0x20,
' 	ssfLOCALAPPDATA = 0x1c,
' 	ssfMYPICTURES = 0x27,
' 	ssfNETHOOD = 0x13,
' 	ssfNETWORK = 0x12,
' 	ssfPERSONAL = 0x05,
' 	ssfPRINTERS = 0x04,
' 	ssfPRINTHOOD = 0x1b,
' 	ssfPROFILE = 0x28,
' 	ssfPROGRAMFILES = 0x26,
' 	ssfPROGRAMFILESx86 = 0x30,
' 	ssfPROGRAMS = 0x02,
' 	ssfRECENT = 0x08,
' 	ssfSENDTO = 0x09,
' 	ssfSTARTMENU = 0x0b,
' 	ssfSTARTUP = 0x07,
' 	ssfSYSTEM = 0x25,
' 	ssfSYSTEMx86 = 0x29,
' 	ssfTEMPLATES = 0x15,
' 	ssfWINDOWS = 0x24
' } ShellSpecialFolderConstants;
' ***********************************************************



