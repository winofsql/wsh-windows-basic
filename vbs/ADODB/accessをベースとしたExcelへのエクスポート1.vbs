' Access の接続文字列
strConnectMdb = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\販売管理.accdb" & ";"

' Excel
strConnectExcel = _
    "[Excel 12.0 xml;DATABASE=C:\app\workspace\販売管理.xlsx].社員マスタ" & CLng(Timer)

' 処理
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectMdb
Query = "select * into " & strConnectExcel & " from 社員マスタ"
Cn.Execute Query
Cn.Close()
