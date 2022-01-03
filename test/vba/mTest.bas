Attribute VB_Name = "mTest"
'Requires JsBridge from here:             https://github.com/sancarn/VbaJsBridge
'ScriptLab test can be downloaded here:   https://gist.github.com/sancarn/b974b650f4b451ff2de51861af1671b1
Sub test()
  Dim js As JsBridge: Set js = JsBridge.Create("test")
  Call js.SendMessage("hello world")
End Sub
