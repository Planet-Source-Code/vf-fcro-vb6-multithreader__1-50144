VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Thread Start Parameters"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create Thread with that Start Parameters"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Argument"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Message"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reason"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall

Private Sub Command1_Click()
On Error GoTo Dalje
Dim Reason As Long
Dim Message As Long
Dim Argument As Variant
Dim ThreadH As Long
Dim TID As Long
Reason = CLng(Text1(0))
Message = CLng(Text1(1))
Argument = CVar(Text1(2))

ThreadH = MULTITHREADER.CreateNewThread(TID, &HC000&, THREAD_PRIORITY_NORMAL, Reason, Message, Argument, ObjectEnabled)
CloseHandle ThreadH
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Reason or Message isn't 32bit value!", , "Error"
End Sub

Private Sub Form_Load()
Caption = Caption & "(TID:" & App.ThreadId & ")"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MULTITHREADER.RemoveThread 0
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long
End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long
End Function
