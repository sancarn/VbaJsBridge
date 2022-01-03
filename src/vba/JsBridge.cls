VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec:
'A class used to automate  objects which implement IAccessible
'
'
'CONSTRUCTORS
'    [X] Create
'    [X] CreateClone
'PUBLIC INSTANCE METHODS:
'    [X] SendMessage
'EXAMPLE
'  Dim js as JsBridge: set js = JsBridge.Create("b19159c0-dab1-4803-b495-f2b8cd149406")
'  Call js.SendMessage("hello world")

Private pUID As String

'CONSTRUCTORS

'Create a bridge
'@param {string} UID. Suggest this is a GUID created at https://www.guidgenerator.com. This is the JsBridge identifier for your application and prevents conflicts with other applications.
'@returns {JsBridge} bridge to use to communicate with js runtime
Public Function Create(ByVal sUID As String) As JsBridge
    Set Create = New JsBridge
    Call Create.protInit(sUID)
End Function

'Create a clone of this bridge
'@returns {JsBridge} The clone returned
Public Function CreateClone() As JsBridge
    Set Create = New JsBridge
    Call Create.protInit(pUID)
End Function

'INSTANCE METHODS

'Syncronously send a message to JS environment
'@param {string} Message to send
'@remark The function will wait until `finish()` is called before returning.
Public Sub SendMessageSync(ByVal sMessage As String)
    Dim sNodeID As String: sNodeID = SendMessage(sMessage)
    Do
      Dim node As CustomXMLPart
      Set node = ThisWorkbook.CustomXMLParts.SelectByID(sNodeID)
      If node.XML Like "*finished=""true""*" Then
        'Ensure that attribute is true
        Dim attr As CustomXMLNode
        For Each attr In node.DocumentElement.Attributes
          If attr.XML = "finished=""true""" Then
            Call node.Delete
            Exit Sub
          End If
        Next
      End If
      
      DoEvents
    Loop
End Sub

'Asyncronously send a message to JS environment
'@param {string} Message to send
'@returns {string} The ID of the CustomXMLPart created. See source code of `SendMessageSync` to see how this can be used.
'@remark The function will wait until `finish()` is called before returning.
Public Function SendMessage(ByVal sMessage As String) As String
    SendMessage = AddAction(pUID, "js", sMessage).ID
End Function

'Delete all messages under the namespace ID
Public Sub DeleteAll()
  Dim col As CustomXMLParts
  Set col = ThisWorkbook.CustomXMLParts.SelectByNamespace(pUID)
  While col.Count > 0
    col.Item(1).Delete
  Wend
End Sub

'PROTECTED
Public Sub protInit(sUID As String)
    pUID = sUID
End Sub

'Create an XML action to send down the bridge
'@param {string} The applications UID
'@param {string} "js" to send a message to JS environment. Suggest "vba" to send a message to VBA environment.
'@param {strubg} The action / data to send
'@returns {string} The Action XML to send to the bridge
Private Function CreateAction(ByVal sAppUID As String, ByVal actionFor As String, ByVal action As String) As String
    CreateAction = "<action for=""" & actionFor & """ handled=""false"" finished=""false"" xmlns=""" & sAppUID & """>" & action & "</action>"
End Function

'Sends an XML action down the bridge
'@param {string} The applications UID
'@param {string} "js" to send a message to JS environment. Suggest "vba" to send a message to VBA environment.
'@param {strubg} The action / data to send
Private Function AddAction(ByVal sAppUID As String, ByVal actionFor As String, ByVal action As String) As CustomXMLPart
    Set AddAction = ThisWorkbook.CustomXMLParts.Add(CreateAction(sAppUID, actionFor, action))
End Function
