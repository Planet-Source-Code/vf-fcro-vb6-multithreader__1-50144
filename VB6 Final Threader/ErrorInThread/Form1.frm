VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Serious Error Example!"
   ClientHeight    =   525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create Thread with Serious Error!"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall
Dim NewThreadId As Long

Private Sub Command1_Click()
Dim ThreadHandle As Long
ThreadHandle = MULTITHREADER.CreateNewThread(NewThreadId, &HC000&, THREAD_PRIORITY_NORMAL, 0, 0, Empty, ObjectEnabled)
CloseHandle ThreadHandle
End Sub





Private Sub Form_Load()
Caption = Caption & " :Thread Id:" & App.ThreadId
End Sub

Private Sub Form_Unload(Cancel As Integer)
MULTITHREADER.RemoveThread 0
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long
If ThreadNotify < 0 Then
NewThreadId = 0
MsgBox "Thread Id:" & CStr(ThreadNotify And &H7FFFFFFF) & " destroyed!", vbExclamation, "Information"
Else
If ThreadNotify = App.ThreadId Then Exit Function
MsgBox "Thread Id:" & ThreadNotify & " created!", vbExclamation, "Information"
End If

End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long

End Function
