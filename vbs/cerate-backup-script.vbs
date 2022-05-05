' ***********************************************************
' 処理開始
' ***********************************************************
Set Fso = Wscript.CreateObject( "Scripting.FileSystemObject" )
Set Shell = Wscript.CreateObject( "Shell.Application" )

' ***********************************************************
' 実行中ディレクトリの取得
' ***********************************************************
strPath = Wscript.ScriptFullName 
Set objFile = Fso.GetFile( strPath )
strBackupFolder = Fso.GetParentFolderName( objFile )

' ***********************************************************
' バックアップ対象ディレクトリの取得
' ***********************************************************
' ① 省略すると、ルートはデスクトップ
Set objFolder = Shell.BrowseForFolder( 0, "バックアップするフォルダを選択してください", &H4B )

' ② 文字列による直接指定
' strRoot = "c:\"
' Set objFolder = Shell.BrowseForFolder( 0, "バックアップするフォルダを選択してください", &H4B, strRoot )

' ③ ルートを番号で指定( この場合は C:\Users\username\AppData\Local )
' ※ あまり現実的ではない特殊ディレクトリの選択
' nRoot = &h1c
' Set objFolder = Shell.BrowseForFolder( 0, "バックアップするフォルダを選択してください", &H4B, nRoot )

if objFolder is nothing then
	WScript.Quit
end if
if not objFolder.Self.IsFileSystem then
	WScript.Echo "ファイルシステムではありません"
	WScript.Quit
end if

strTargetFolder = objFolder.Self.Path
strName = Replace( strTargetFolder, ":", "" )
strName = Replace( strName, "\", "_" )
strName = Replace( strName, " ", "" )
strName = "BK_" & strName

' ***********************************************************
' スクリプト作成
' ***********************************************************
Set OutFile = Fso.OpenTextFile( strBackupFolder & "\" & strName & ".vbs", 2, True )

OutFile.WriteLine "strName = """ & strName & """"
OutFile.WriteLine "strTarget = """ & strTargetFolder & """"
OutFile.WriteLine "strBackupFolder = """ & strBackupFolder & """"
OutFile.Write "if MsgBox( strTarget & vbCrLf & ""のバックアップを開始します。よろしいですか? (保存先:"" & strBackupFolder & ""\"" & strName & "")"""
OutFile.WriteLine ", 1 ) = 2 then"
OutFile.WriteLine "	Wscript.Quit"
OutFile.WriteLine "end if"

OutFile.WriteLine "Set WshShell = Wscript.CreateObject( ""WScript.Shell"" )"
OutFile.Write "ExecCommand = ""cmd.exe /C """"xcopy.exe """""" & strTarget & """""" """""" & strBackupFolder & ""\"" & strName & ""\"""""
OutFile.WriteLine " /D /E /C /S /Y & PAUSE"""""""
OutFile.WriteLine "Call WshShell.Run( ExecCommand )"

OutFile.Close

WScript.Echo "バックアップスクリプト : " &  strName & ".vbs" & " を作成しました"


' ***********************************************************
' ディレクトリ指定用番号
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



