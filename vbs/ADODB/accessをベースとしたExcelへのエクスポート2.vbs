' Access ÌÚ±¶ñ
strConnectMdb = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\ÌÇ.accdb" & ";"

' ÇÝExcel
strConnectExcelIn = _
    "[Excel 12.0 xml;DATABASE=C:\app\workspace\ÌÇ.xlsx].Ðõ}X^"

' oÍExcel
strConnectExcelOut = _
    "[Excel 12.0 xml;DATABASE=C:\app\workspace\Ðõ}X^.xlsx].Ðõ}X^" & CLng(Timer)

' 
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectMdb
Query = "select * into " & strConnectExcelOut & " from " & strConnectExcelIn
Cn.Execute Query
Cn.Close()
