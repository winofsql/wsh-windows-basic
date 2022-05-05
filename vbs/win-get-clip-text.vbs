Set obj = Wscript.CreateObject("htmlfile")
text = obj.ParentWindow.ClipboardData.GetData("text")

Wscript.Echo text