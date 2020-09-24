VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit App"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Attach Information and Signal Another Thread"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall





Private Sub Command1_Click()
Dim PrevCallArg As Long
MULTITHREADER.AttachThreadCallArguments 0, 0, 0, 0, Array("My First Argument!", "My Last Argument!!!!"), UsingMove, PrevCallArg, False
MULTITHREADER.SignalLocalObject 0
End Sub

Private Sub Command2_Click()
MULTITHREADER.RemoveThread 0
End Sub

Private Sub Form_Load()
Caption = "Thread Id:" & App.ThreadId
End Sub

Private Sub Form_Unload(Cancel As Integer)
MULTITHREADER.RemoveThread App.ThreadId 'Remove Thread & Exit App
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long
Dim IsValidCall As Boolean
Dim THandle As Long

If ThreadNotify < 0 Then
    
    MsgBox "Thread terminated!" & vbCrLf & "I'll create another one!", vbCritical, "Internal Notify!"
    THandle = MULTITHREADER.CreateNewThread(Thread2, &HC000&, THREAD_PRIORITY_NORMAL, 0, 0, Empty, ObjectEnabled)
    CloseHandle THandle
    
    
ElseIf ThreadNotify > 0 And ThreadNotify <> App.ThreadId Then
    MULTITHREADER.CallThread ThreadNotify, 0, 1, Empty, UsingCopy, UsingAPC, IsValidCall

End If

End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long
Dim Reason As Long
Dim Message As Long
Dim Arguments As Variant
Dim IsValidCall As Boolean


MULTITHREADER.TranslateArguments CallArgs, Reason, Message, Arguments

If Message = 1 Then

    Caption = "Waiting For the Signal!! Thread Id:" & App.ThreadId
    DoEvents
    MULTITHREADER.WaitForLocalObject Thread2, -1
    Caption = "Signaled! Thread Id:" & App.ThreadId
    MULTITHREADER.DetachThreadCallArguments 0, 0, Reason, Message, Arguments, True
    MsgBox Arguments(0) & vbCrLf & Arguments(1), vbExclamation, "Detached information by ThreadId:" & App.ThreadId
    MULTITHREADER.CallThread Thread2, 0, 1, Empty, UsingCopy, UsingAPC, IsValidCall
End If
End Function
