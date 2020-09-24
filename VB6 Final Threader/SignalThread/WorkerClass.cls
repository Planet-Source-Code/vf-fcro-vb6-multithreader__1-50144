VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements InThreadCall

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long
End Function
Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long
Dim Message As Long
Dim Reason As Long
Dim Args As Variant
Dim IsValidCall As Boolean

Dim SignalRet As Long
'Translate call !!!
MULTITHREADER.TranslateArguments CallArgs, Reason, Message, Args

SignalRet = MULTITHREADER.SignalLocalObject(App.ThreadId) 'Signal object which owner is current thread! [Another object wait for this signal!]

If SignalRet = 1 Then
    MsgBox "Event signaled right now!" & vbCrLf & "I'm done,BYE!", vbExclamation, "Signal Thread!"
ElseIf SignalRet = 0 Or SignalRet = -1 Then
    MsgBox "Object not valid!", vbCritical, "Signal Thread!"
End If

MULTITHREADER.ExitMessagePump 'Exit itself
End Function


