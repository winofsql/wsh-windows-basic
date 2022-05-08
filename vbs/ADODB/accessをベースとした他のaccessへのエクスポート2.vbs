' 販売管理.accdb:社員マスタ => Database.accdb:社員マスタ

' Access の接続文字列(接続用)
strConnectBaseAccess = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\sqlite3.accdb" & ";"

' 読込みAccess の接続文字列
strConnectAccessIn = _
    "[;DATABASE=C:\app\workspace\販売管理.accdb].社員マスタ"

' 出力Access の接続文字列
strConnectAccessOut = _
    "[;DATABASE=C:\app\workspace\Database.accdb].社員マスタ" & CLng(Timer)

' 処理
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectBaseAccess
Query = "select * into " & strConnectAccessOut & " from " & strConnectAccessIn
Cn.Execute Query
Cn.Close()
