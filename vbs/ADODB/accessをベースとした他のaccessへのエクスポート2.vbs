' �̔��Ǘ�.accdb:�Ј��}�X�^ => Database.accdb:�Ј��}�X�^

' Access �̐ڑ�������(�ڑ��p)
strConnectBaseAccess = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\sqlite3.accdb" & ";"

' �Ǎ���Access �̐ڑ�������
strConnectAccessIn = _
    "[;DATABASE=C:\app\workspace\�̔��Ǘ�.accdb].�Ј��}�X�^"

' �o��Access �̐ڑ�������
strConnectAccessOut = _
    "[;DATABASE=C:\app\workspace\Database.accdb].�Ј��}�X�^" & CLng(Timer)

' ����
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectBaseAccess
Query = "select * into " & strConnectAccessOut & " from " & strConnectAccessIn
Cn.Execute Query
Cn.Close()
