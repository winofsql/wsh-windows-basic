' �̔��Ǘ�.accdb:�Ј��}�X�^ => sqlite3.accdb:�Ј��}�X�^001

' Access �̐ڑ�������
strConnectBaseAccess = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\�̔��Ǘ�.accdb" & ";"

' Access �̐ڑ�������
strConnectAccess = _
    "[;DATABASE=C:\app\workspace\sqlite3.accdb].�Ј��}�X�^" & CLng(Timer)

' ����
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectBaseAccess
Query = "select * into " & strConnectAccess & " from �Ј��}�X�^"
Cn.Execute Query
Cn.Close()
