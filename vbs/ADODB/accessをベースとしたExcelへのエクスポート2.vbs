' Access の接続文字列
strConnectMdb = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\販売管理.accdb" & ";"

' 読込みExcel
strConnectExcelIn = _
    "[Excel 12.0 xml;DATABASE=C:\app\workspace\販売管理.xlsx].社員マスタ"

' 出力Excel
strConnectExcelOut = _
    "[Excel 12.0 xml;DATABASE=C:\app\workspace\社員マスタ.xlsx].社員マスタ" & CLng(Timer)

' 処理
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectMdb
Query = "select * into " & strConnectExcelOut & " from " & strConnectExcelIn
Cn.Execute Query
Cn.Close()
