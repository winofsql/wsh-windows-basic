' Access ?̐ڑ???????
strConnectBaseAccess = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\?̔??Ǘ?.accdb" & ";"

' MySQL ODBC 8.0 Unicode Driver
strConnectMySQL = _
    "[ODBC;Provider=MSDASQL" & _
    ";Driver={MySQL ODBC 8.0 Unicode Driver}" & _
    ";SERVER=localhost" & _
    ";DATABASE=lightbox" & _
    ";UID=root" & _
    ";PWD=" & _
    ";Charset=cp932" & _
    ";].?Ј??}?X?^" & CLng(Timer)

' ????
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectBaseAccess
Query = "select * into " & strConnectMySQL & " from ?Ј??}?X?^"
Cn.Execute Query
Cn.Close()
