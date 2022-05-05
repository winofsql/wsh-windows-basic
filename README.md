# wsh-windows-basic

### [OLEDBプロバイダの一覧取得](https://qiita.com/yaju/items/7b0aa9e9f30005f60388)
```
(New-Object data.oledb.oledbenumerator).getElements() | select SOURCES_NAME, SOURCES_DESCRIPTION
```
