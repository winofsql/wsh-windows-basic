' Access の接続文字列
strConnectBaseAccess = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\販売管理.accdb" & ";"

' MySQL ODBC 8.0 Unicode Driver
strConnectSqlite3 = _
    "[ODBC;Provider=MSDASQL" & _
    ";Driver={SQLite3 ODBC Driver}" & _
    ";DATABASE=C:\app\workspace\lightbox.sqlite3" & _
    ";].社員マスタ" & CLng(Timer)

' 処理
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectBaseAccess
Query = "select * into " & strConnectSqlite3 & " from 社員マスタ"
Cn.Execute Query
Cn.Close()
