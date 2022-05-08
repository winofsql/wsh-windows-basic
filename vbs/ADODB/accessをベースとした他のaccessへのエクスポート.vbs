' 販売管理.accdb:社員マスタ => sqlite3.accdb:社員マスタ001

' Access の接続文字列
strConnectBaseAccess = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\販売管理.accdb" & ";"

' Access の接続文字列
strConnectAccess = _
    "[;DATABASE=C:\app\workspace\sqlite3.accdb].社員マスタ" & CLng(Timer)

' 処理
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectBaseAccess
Query = "select * into " & strConnectAccess & " from 社員マスタ"
Cn.Execute Query
Cn.Close()
