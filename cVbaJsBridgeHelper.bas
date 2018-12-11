Attribute VB_Name = "cVbaJsBridgeHelper"

' New event loop idea:
' -----------------------------------------------------------
' Public Bridges as collection  - List of bridges to update
' Public Initialised as Boolean - Does this require initialising?
' Public NextTime as DateTime   - Time of next scheduled update.
' Const Interval = ...          - Interval in seconds.
'
' Timer_PollObject()     --> Adds object to Collection and initialises timer if not initialised
' Timer_StopObject()     --> Stops polling an object.
' Timer_Start()          --> Initialises timer.
' Timer_Stop()           --> Stops timer.
' Timer_Handler()        --> Loops through Bridges collection, calling `PolUpdate()` on each object.


Public bridge As cVbaJsBridge
Public Bridges As Collection
Public Initialised As Boolean
Public NextTime As Date
Const Interval = 1

'Start polling an object
Public Sub Timer_PollObject(bridge As cVbaJsBridge)
  If Not cVbaJsBridgeHelper.Initialised Then
    Set cVbaJsBridgeHelper.Bridges = New Collection
    cVbaJsBridgeHelper.Timer_Start
    Initialised = True
  End If
  
  'Add object to Bridges
  cVbaJsBridgeHelper.Bridges.add bridge
End Sub

'Stop polling an object
Public Sub Timer_StopObject(bridge As cVbaJsBridge)
  Dim i As Long
  For i = 1 To cVbaJsBridgeHelper.Bridges.count
    If cVbaJsBridgeHelper.Bridges(i) Is bridge Then
      cVbaJsBridgeHelper.Bridges.remove (i)
      Exit Sub
    End If
  Next
End Sub

'Start timer / this is also the timer itself
Public Sub Timer_Start()
  'Schedules self
  NextTime = Now + TimeSerial(0, 0, Interval)
  Application.OnTime NextTime, "cVbaJsBridgeHelper.Timer_Start", Schedule:=True
  
  'Handle bridges...
  Call cVbaJsBridgeHelper.Timer_Handler
End Sub

'Stop the timer
Public Sub Timer_Stop()
  Application.OnTime NextTime, "cVbaJsBridgeHelper.Timer_Start", Schedule:=False
End Sub

'Handle each bridge in collection
Public Sub Timer_Handler()
  Dim bridge As cVbaJsBridge
  For Each bridge In cVbaJsBridgeHelper.Bridges
    If Not bridge Is Nothing Then
      Call bridge.PolUpdate
    End If
  Next
End Sub





'**************************************
' TESTS
'**************************************

Sub testCVbaJsBridge()
  Set bridge = cVbaJsBridge.Create(ActiveWorkbook)
  
  Debug.Print bridge.evalJS("""hello world""")
End Sub





'**************************************
' OLD EVENT LOOP
'**************************************

'Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long
'Public Handles As Collection
'Public bridge As cVbaJsBridge
''Public Function TimerStart(ByRef caller As cVbaJsBridge, Optional ByVal milliseconds As Long = 300) As LongPtr
'  'Initialise Handles if uninitialised
'  If Handles Is Nothing Then Set Handles = New Collection
'
'  'Original
'  TimerStart = SetTimer(0&, 0&, milliseconds, AddressOf cVbaJsBridgeHelper.TimerHandle)
'
'  'Preserve Timer ID
'  Handles.Add caller, CStr(TimerStart)
'End Function
'Public Sub TimerStop(TimerID As LongPtr)
'  On Error Resume Next
'  KillTimer 0&, TimerID
'  Handles.Remove CStr(TimerID)
'End Sub
'Public Sub TimerHandle(ByVal HWnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTimer As Long)
'  Dim obj As cVbaJsBridge
'  Set obj = Handles(CStr(nIDEvent))
'  Call obj.PolEventLoop_Handle
'End Sub



