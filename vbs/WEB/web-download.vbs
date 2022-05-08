Set objSrvHTTP = Wscript.CreateObject("Msxml2.ServerXMLHTTP")
Set Stream = Wscript.CreateObject("ADODB.Stream")


webUrlFolder = "https://winofsql.jp/image/daz/"
webFileName = "2021-08-03-1.jpg"
savePath = ".\" & "2021-08-03-1.jpg"

Call objSrvHTTP.Open("GET", webUrlFolder & webFileName, False )
objSrvHTTP.Send
Stream.Open
Stream.Type = 1
Stream.Write objSrvHTTP.responseBody
Stream.SaveToFile savePath, 2
Stream.Close
