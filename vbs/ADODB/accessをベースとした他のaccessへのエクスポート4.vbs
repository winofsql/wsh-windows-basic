' sqlite3.accdb( ODBC:Sqlte3:Πυ}X^ ) => ΜΗ.mdb:Πυ}X^

' Access ΜΪ±Άρ(Ϊ±p)
strConnectBaseAccess = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\ΜΗ.accdb" & ";"

' ΗέAccess ΜΪ±Άρ
strConnectAccessIn = _
    "[;DATABASE=C:\app\workspace\sqlite3.accdb].Πυ}X^"

' oΝAccess ΜΪ±Άρ
strConnectAccessOut = _
    "[;DATABASE=C:\app\workspace\ΜΗ.mdb].Πυ}X^" & CLng(Timer)

' 
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectBaseAccess
Query = "select * into " & strConnectAccessOut & " from " & strConnectAccessIn
Cn.Execute Query
Cn.Close()
