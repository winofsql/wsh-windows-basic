' Access �̐ڑ�������
strConnectBaseAccess = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\�̔��Ǘ�.accdb" & ";"

' MySQL ODBC 8.0 Unicode Driver
strConnectSqlite3 = _
    "[ODBC;Provider=MSDASQL" & _
    ";Driver={SQLite3 ODBC Driver}" & _
    ";DATABASE=C:\app\workspace\lightbox.sqlite3" & _
    ";].�Ј��}�X�^" & CLng(Timer)

' ����
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectBaseAccess
Query = "select * into " & strConnectSqlite3 & " from �Ј��}�X�^"
Cn.Execute Query
Cn.Close()
