' Access �̐ڑ�������
strConnectMdb = _
    "Provider=Microsoft.ACE.OLEDB.16.0;" & _
    "Data Source=C:\app\workspace\�̔��Ǘ�.accdb" & ";"

' Excel
strConnectExcel = _
    "[Excel 12.0 xml;DATABASE=C:\app\workspace\�̔��Ǘ�.xlsx].�Ј��}�X�^" & CLng(Timer)

' ����
Set Cn = Wscript.CreateObject("ADODB.Connection")
Cn.Open strConnectMdb
Query = "select * into " & strConnectExcel & " from �Ј��}�X�^"
Cn.Execute Query
Cn.Close()
