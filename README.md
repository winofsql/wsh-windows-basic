# wsh-windows-basic

### [OLEDBプロバイダの一覧取得](https://qiita.com/yaju/items/7b0aa9e9f30005f60388)
```
(New-Object data.oledb.oledbenumerator).getElements() | select SOURCES_NAME, SOURCES_DESCRIPTION
```

## Shell.Application

<a href="https://docs.microsoft.com/en-us/windows/desktop/shell/shell-reference-bumper" target="_blank">Shell Reference</a>\
<a href="https://docs.microsoft.com/en-us/windows/desktop/shell/objects" target="_blank">Shell Objects for Scripting and Microsoft Visual Basic</a>\
<a href="https://docs.microsoft.com/en-us/windows/desktop/shell/shell" target="_blank">Shell object</a>\
<a href="https://docs.microsoft.com/en-us/windows/desktop/shell/shell-browseforfolder" target="_blank">Shell.BrowseForFolder method</a>

## WSH

[WScript オブジェクト](https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364475(v=msdn.10))\
[WshShell オブジェクト](https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364436(v=msdn.10))
