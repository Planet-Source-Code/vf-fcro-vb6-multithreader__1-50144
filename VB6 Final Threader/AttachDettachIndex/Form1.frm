VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simple Signalization"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Attach If Index Exist and Return Previous!"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Attach Information"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Signal Thread"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Thread"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall

Private Sub Command1_Click()
Dim ThreadId As Long
Dim ThreadHandle As Long
ThreadHandle = MULTITHREADER.CreateNewThread(ThreadId, &HC000&, THREAD_PRIORITY_NORMAL, 0, 0, Empty, ObjectEnabled)
CloseHandle ThreadHandle
End Sub

Private Sub Command2_Click()
MULTITHREADER.SignalLocalObject 0
End Sub

Private Sub Command3_Click()
Dim PreviousCallArgs As Long
MULTITHREADER.AttachThreadCallArguments 0, 12, 0, 0, CStr(Text1.Text), UsingCopy, PreviousCallArgs, Check1

    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long

If PreviousCallArgs = -1 Then
 'canceled attach because of flag:AttachIfExist=FALSE
 'just peek on Index!
    MULTITHREADER.DetachThreadCallArguments 0, 12, Reason, Message, Args, False
    MsgBox Args, , "Attach canceled,index Allready Exist! peek on previous Argument:"

ElseIf PreviousCallArgs <> 0 Then
    MULTITHREADER.TranslateArguments PreviousCallArgs, Reason, Message, Args
    MsgBox Args, , "Attach succes,previous Arguments returned:"

End If

End Sub

Private Sub Form_Load()
Caption = "Main Thread Id:" & App.ThreadId
End Sub

Private Sub Form_Unload(Cancel As Integer)
MULTITHREADER.RemoveThread 0
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long
Dim TID As Long
If ThreadNotify < 0 Then
TID = ThreadNotify And &H7FFFFFFF
MsgBox "Destroyed Thread:" & TID, vbExclamation, "Notify"
Else
MsgBox "Created Thread:" & ThreadNotify, vbExclamation, "Notify"
End If
End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long

End Function
