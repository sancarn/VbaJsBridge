VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVbaJsBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents JSBridge As CustomXMLPart
Attribute JSBridge.VB_VarHelpID = -1
Public Initialised As Boolean
Public callbacks As Collection
Public Terminators As Collection
Private Handles As Object
Private PolEventLoop_Id As String
Private DefaultNamespace As String

Public Function Create(wb As Workbook, Optional ByVal Namespace As String = "VBA-JS-Bridge") As cVbaJsBridge
  'Register bridge:
  Call registerBridge(wb, Namespace)
  
  'Create new bridge
  Dim bridge As New cVbaJsBridge
  
  'Set Namespace
  Call bridge.SetNamespace(Namespace)
  
  'Set bridge of our new object (hooking into it's events)
  Set bridge.JSBridge = wb.CustomXMLParts.SelectByNamespace(Namespace)(1)
  
  'Set as initialised
  bridge.Initialised = True
  
  'return bridge
  Set Create = bridge
End Function

'Evaluate javascript asynchronously
Public Function evalJS(ByVal js As String, Optional id As String = "") As String
  If Initialised Then
    ' Create an ID
    If id = "" Then id = getGUID()
    
    'Use ID to be able to get result later.
    Dim sXML As String
    sXML = "<js-exec id=""" & id & """>" & _
               js & _
           "</js-exec>"
    
    ' Evaluate javascript via bridge
    Me.JSBridge.DocumentElement.AppendChildSubtree sXML
    
    'Return ID
    evalJS = id
  Else
    Err.Raise -1, "cVbaJsBridge.evalJS(...)", "VBA-JS bridge is not initialised. Use Create(...) to create a bridge on a document"
  End If
End Function

'Evaluate javascript synchronously
Public Function evalJSSync(ByVal js As String, Optional id As String) As String
  Err.Raise 0, "cVbaJsBridge.evalJsSync", "Currently not implemented"
  If Initialised Then
    ' Create an ID
    Dim id As String
    id = getGUID()
    
    'Add no callback to callbacks
    callbacks.add Nothing, id
    
    
    'Use ID to be able to get result later.
    Dim sXML As String
    sXML = "<js-exec id=""" & id & """>" & _
               js & _
           "</js-exec>"
    
    ' Evaluate javascript via bridge
    Me.JSBridge.DocumentElement.AppendChildSubtree sXML
    
    
    
  Else
    Err.Raise -1, "cVbaJsBridge.evalJS(...)", "VBA-JS bridge is not initialised. Use Create(...) to create a bridge on a document"
  End If
End Function

Public Sub addCallback(callback As Object)
  'CallByName(object as Object,[ProcName as String], [CallType as vbCallBack], [args() as variant])
  'Application.Run(ModuleName.SubName, arg1, arg2, arg3, ..., arg30)
  'Callback Options:
  '{type:"Object","object":Object,"ProcName":"DoSomething","CallType":..., args:[...]}
  '{type:"Module","module":"Module1","ProcName":"DoSomething", args:[...]}
  'Can use a dictionary or a collection (or any object who's "Attribute Value.VB_UserMemId = 0" relates to a collection)
  callbacks.add callback
End Sub


Sub test()
  'evalJS "console.log(""hello world"")"
  'evalJS "console.log(""hello mark"")"
  'evalJS "console.log(""hello johnny"")"
  evalJS "console.log(document)"
End Sub

Private Sub callCallback(callback As Object, Optional xmlNode As Object) '2nd param commonly: OFFICE.CustomXMLNode
  'If no callback supplied do nothing
  If callback Is Nothing Then Exit Sub
  
  'Prepare args array from callback("args")
  Dim args() As Variant
  args = callback("args")
  
  'Increase array length by 1
  ReDim Preserve args(UBound(args) + 1)
  
  'Loop through array. Set item i+1 to i
  Dim i As Integer
  For i = UBound(args) To LBound(args) Step -1
    If i <> LBound(args) Then
      If VarType(args(i - 1)) = vbObject Then
        Set args(i) = args(i - 1)
      Else
        args(i) = args(i - 1)
      End If
    End If
  Next
  
  'First argument should always be xmlNode
  Set args(LBound(args)) = xmlNode
  
  'Call callback based on type
  If callback("type") = "Module" Then
    'Arguments supplied to callback object
    Dim origArgs As Variant
    origArgs = callback("args")
    
    'Get missing for missing params
    Dim missing As Variant
    missing = getMissing()
    
    'Create argument array for Run command
    Dim args(0 To 29) As Variant
    Dim origLength, origLB As Long
    origLength = UBound(origArgs) - LBound(origArgs) + 1
    origLB = LBound(origArgs)
    
    'Set arguments as required.
    For i = 0 To 29
      'Set fixed arguments, else set dynamic args
      If i <= 0 Then
        Set args(i) = xmlNode
      Else
        'If i < original length then these are missing arguments. Assign these to missing. Else set object in params accordingly.
        If i > origLength Then
          args(i) = missing
        Else
          'If object set as object, else set as value.
          If VarType(origArgs(origLB + i - 1)) = vbObject Then
            Set args(i) = origArgs(origLB + i - 1)
          Else
            args(i) = origArgs(origLB + i - 1)
          End If
        End If
      End If
    Next
    
    'funcName - Module based or global
    Dim funcName As String
    If callback("module") <> "" Then
      funcName = callback("module") & "." & callback("ProcName")
    Else
      Debug.Print "WARN: cVbaJsBridge.callCallback() - No module name supplied, thus using a global function search."
      funcName = callback("ProcName")
    End If
    
    'If original length > 30 then warn dev that too many arguments submitted
    If origLength > (30 - 1) Then Debug.Print "WARN: cVbaJsBridge.callCallback() - Cannot pass more than 30 arguments to Application.Run. Assuming only 30 arguments required. Note: 1 argument is taken up by default by xmlNode. You are advised to pass a data collection reather than 29 extra arguments."
    
    'Run macro
    Application.Run funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28), args(29)
    
'    'Up to 30 arguments...
'    Select Case ub - lb + 1
'      Case 0: Application.Run funcName
'      Case 1: Application.Run funcName, args(lb + 0)
'      Case 2: Application.Run funcName, args(lb + 0), args(lb + 1)
'      Case 3: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2)
'      Case 4: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3)
'      Case 5: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4)
'      Case 6: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5)
'      Case 7: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6)
'      Case 8: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7)
'      Case 9: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8)
'      Case 10: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9)
'      Case 11: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10)
'      Case 12: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11)
'      Case 13: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12)
'      Case 14: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13)
'      Case 15: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14)
'      Case 16: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15)
'      Case 17: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16)
'      Case 18: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17)
'      Case 19: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18)
'      Case 20: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19)
'      Case 21: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20)
'      Case 22: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21)
'      Case 23: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22)
'      Case 24: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22), args(lb + 23)
'      Case 25: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22), args(lb + 23), args(lb + 24)
'      Case 26: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22), args(lb + 23), args(lb + 24), args(lb + 25)
'      Case 27: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22), args(lb + 23), args(lb + 24), args(lb + 25), args(lb + 26)
'      Case 28: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22), args(lb + 23), args(lb + 24), args(lb + 25), args(lb + 26), args(lb + 27)
'      Case 29: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22), args(lb + 23), args(lb + 24), args(lb + 25), args(lb + 26), args(lb + 27), args(lb + 28)
'      Case 30: Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22), args(lb + 23), args(lb + 24), args(lb + 25), args(lb + 26), args(lb + 27), args(lb + 28), args(lb + 29)
'      Case Else
'        Application.Run funcName, args(lb + 0), args(lb + 1), args(lb + 2), args(lb + 3), args(lb + 4), args(lb + 5), args(lb + 6), args(lb + 7), args(lb + 8), args(lb + 9), args(lb + 10), args(lb + 11), args(lb + 12), args(lb + 13), args(lb + 14), args(lb + 15), args(lb + 16), args(lb + 17), args(lb + 18), args(lb + 19), args(lb + 20), args(lb + 21), args(lb + 22), args(lb + 23), args(lb + 24), args(lb + 25), args(lb + 26), args(lb + 27), args(lb + 28), args(lb + 29)
'        Debug.Print "WARN: cVbaJsBridge.callCallback() - Cannot pass more than 30 arguments to Application.Run. Assuming only 30 arguments required."
'    End Select

  Else
    CallByName callback("object"), callback("ProcName"), callback("CallType"), callback("args")
  End If
End Sub



'Register bridge with selected workbook
Private Sub registerBridge(wb As Workbook, Optional Namespace As String = "VBA-JS-Bridge")
  'when JSBridge CustomXMLPart doesn't exist, add it.
  If wb.CustomXMLParts.SelectByNamespace(Namespace).count = 0 Then
    Dim xmlPart As CustomXMLPart
    Set xmlPart = wb.CustomXMLParts.add("<data xmlns=""" & Namespace & """></data>")
  End If
End Sub


'********************
'* HELPER FUNCTIONS *
'********************
'Create a unique ID
Public Function getGUID() As String
    Do While Len(getGUID) < 32
        If Len(getGUID) = 16 Then
            '17th character holds version information
            getGUID = getGUID & Hex$(8 + CInt(Rnd * 3))
        End If
        getGUID = getGUID & Hex$(CInt(Rnd * 15))
    Loop
    getGUID = Mid(getGUID, 1, 8) & "-" & Mid(getGUID, 9, 4) & "-" & Mid(getGUID, 13, 4) & "-" & Mid(getGUID, 17, 4) & "-" & Mid(getGUID, 21, 12)
End Function

'Gets the missing object. Not sure what the official method of getting this is... But this works if passed no params!
Private Function getMissing(Optional a As Variant) As Variant
  getMissing = a
End Function

Private Sub handleVBACommand(data As Object) ' xml As OFFICE.CustomXMLNode
  Dim objID As String
  objID = ""
  Select Case data.Command
    Case "accepts"
    
    Case "CreateObject"
    
    Case "CallByName"
      
    Case "Run"
    
    Case "Application"
      
  End Select
End Sub


Private Function colHasKey(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    HasKey = (Err.Number = 0)
    Err.Clear
End Function


'****************
'* CLASS EVENTS *
'****************
Private Sub Class_Initialize()
  Set Me.callbacks = New Collection
  Set Me.Terminators = New Collection
  Set Handles = CreateObject("Scripting.Dictionary")
  cVbaJsBridgeHelper.Timer_PollObject Me
End Sub

Private Sub Class_Terminate()
  Debug.Assert False
  'Execute all terminator callbacks
  Dim Terminator As Object
  For Each Terminator In Me.Terminators
    Call callCallback(Terminator, Me)
  Next
  
  'Stop timers
  cVbaJsBridgeHelper.Timer_StopObject Me
End Sub

Public Sub SetNamespace(ns As String)
  If Not initialized Then
    DefaultNamespace = ns
  Else
    Debug.Print "Namespace not allowed to change while in process. Change at creation instead!"
  End If
End Sub

Private Function getBridge() As CustomXMLPart
  Set getBridge = wb.CustomXMLParts.SelectByNamespace(Namespace)(1)
End Function

Public Sub PolUpdate()
  'Seek for callbacks
  Dim xmlNode As CustomXMLNode
  Dim callbacks As Collection
  On Error GoTo NoJSBridge
  For Each xmlNode In Me.JSBridge.DocumentElement.ChildNodes
    Select Case xmlNode.BaseName
      'If is a javascript result then call callback
      Case "js-result"
        'Call callback if exists
        If colHasKey(Me.callbacks, xmlNode.Attributes("id")) Then
          Call callCallback(Me.callbacks(xmlNode.Attributes("id")), xmlNode)
        End If
        
        'Delete node
        xmlNode.Delete
      'If "vba-exec" node exists, handle the command
      Case "vba-exec"
        'Handle VBA command
        Call handleVBACommand(xmlNode)
        
        'Delete XML Node
        xmlNode.Delete
      
      Case Else
        'Currently do nothing, perhaps there will be more handlers later?
    End Select
  Next
  Exit Sub
NoJSBridge:
  If Me.JSBridge Is Nothing Then
    Set Me.JSBridge = getBridge()
  Else
    Debug.Assert False
  End If
End Sub






''**************
''* XML EVENTS *
''**************
''Event to latch onto
'Private Sub JSBridge_NodeAfterInsert(ByVal NewNode As Office.CustomXMLNode, ByVal InUndoRedo As Boolean)
'  Debug.Assert False
'  If NewNode.BaseName = "js-result" Then
'    Call callCallback(Me.Callbacks(NewNode.Attributes("id")), NewNode)
'  End If
'End Sub
'Private Sub JSBridge_NodeAfterReplace(ByVal OldNode As Office.CustomXMLNode, ByVal NewNode As Office.CustomXMLNode, ByVal InUndoRedo As Boolean)
'  Debug.Assert False
'End Sub
'Private Sub JSBridge_NodeAfterDelete(ByVal OldNode As Office.CustomXMLNode, ByVal OldParentNode As Office.CustomXMLNode, ByVal OldNextSibling As Office.CustomXMLNode, ByVal InUndoRedo As Boolean)
'  Debug.Assert False
'End Sub
