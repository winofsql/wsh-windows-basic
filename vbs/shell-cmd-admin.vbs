' �R�}���h�v�����v�g���Ǘ��Ҍ����ŋN��
Set WshShell = WScript.CreateObject("WScript.Shell")
Set obj = Wscript.CreateObject("Shell.Application")
obj.ShellExecute _
    "cmd.exe", _
    "/k cd " & WshShell.CurrentDirectory & " & prompt $G& title " & WshShell.CurrentDirectory, _
    , _
    "runas", _
    1
