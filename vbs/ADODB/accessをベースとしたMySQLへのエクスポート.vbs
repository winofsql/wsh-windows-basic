' MDB の接続文字列
strConnectMdb = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\販売管理.accdb" & ";"

' MySQL ODBC 8.0 Unicode Driver
strConnectMySQL = _
    "[ODBC;Provider=MSDASQL" & _
    ";Driver={MySQL ODBC 8.0 Unicode Driver}" & _
    ";SERVER=localhost" & _
    ";DATABASE=lightbox" & _
    ";UID=root" & _
    ";PWD=" & _
    ";Charset=cp932" & _
    ";].社員マスタ001"

' 処理
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectMdb
Query = "select * into " & strConnectMySQL & " from 社員マスタ"
Cn.Execute Query
Cn.Close()
