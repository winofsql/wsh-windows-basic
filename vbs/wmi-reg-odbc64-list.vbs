Const HKEY_LOCAL_MACHINE = &H80000002
Const adVarChar = 200

Dim ErrorMessage

Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")

strPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
bRet = WMIRegEnumValues( HKEY_LOCAL_MACHINE, strPath, aNames, aTypes )

if not bRet then
    Wscript.Echo ErrorMessage
    Wscript.Quit
end if

Set Rs = CreateObject("ADODB.Recordset")
Rs.Fields.Append "ソートキー", adVarChar,255
Rs.Open

For Each data In aNames

    Rs.AddNew
    Rs.Fields("ソートキー").value = data

Next

Rs.Sort = "ソートキー"
Rs.MoveFirst

Do while not Rs.EOF

    Wscript.Echo Rs.Fields("ソートキー").value & ""

    Rs.MoveNext

Loop

Rs.Close


' **********************************************************
' 列挙
' **********************************************************
Function WMIRegEnumValues ( nType, strPath, aNames, aTypes )
    WMIRegEnumValues = False

    on error resume next
    WMIRet = objRegistry.EnumValues( nType, strPath, aNames, aTypes )
    if Err.Number <> 0 then
        ErrorMessage = Err.Description
        Exit Function
    end if
    if WMIRet <> 0 then
        ErrorMessage = Hex( WMIRet )
        Exit Function
    end if
    on error goto 0

    WMIRegEnumValues = True
End Function

