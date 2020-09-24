VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Using Global Sync Object"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Try to get object while new thread ticking!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Thread,new thread take the Sync Object!"
      Height          =   375
      Left            =   120
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
If NewThreadId <> 0 Then MsgBox "Thread Allready exist!", vbExclamation, "Information": Exit Sub
Dim ThreadHandle As Long
ThreadHandle = MULTITHREADER.CreateNewThread(NewThreadId, &HC000&, THREAD_PRIORITY_NORMAL, 0, 0, Empty, ObjectEnabled)
CloseHandle ThreadHandle
End Sub



Private Sub Command2_Click()
Dim ret As Boolean
ret = MULTITHREADER.EnterSynchronization(True)
MsgBox ret, , "Did i get the object?"
If ret Then MULTITHREADER.LeaveSynchronization

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
