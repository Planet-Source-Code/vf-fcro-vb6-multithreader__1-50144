VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simple Signalization"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Action"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4695
      Begin VB.CommandButton Command4 
         Caption         =   "Send order to Signal Thread that fire Event"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Destroy Signal Thread which cause invalid Object"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Waitable Thread"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Signal Thread"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Notify CREATE / DESTROY Thread"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall
Private Thread1 As Long
Private Thread2 As Long

Private Sub Command1_Click()
Dim ThreadHandle As Long
If Thread1 <> 0 Then MsgBox "Signal Thread allready Exist!", vbExclamation, "Information!": Exit Sub

ThreadHandle = MULTITHREADER.CreateNewThread(Thread1, &HC000&, THREAD_PRIORITY_NORMAL, 0, 1, Empty, ObjectEnabled)
CloseHandle ThreadHandle

End Sub





Private Sub Command2_Click()
If Thread1 = 0 Then MsgBox "Signal Thread doesnt exist!", vbCritical, "Main Thread!": Exit Sub
MULTITHREADER.RemoveThread Thread1

End Sub

Private Sub Command3_Click()
Dim ThreadHandle As Long
If Thread1 = 0 Then
MsgBox "Create Signal Thread First!", vbExclamation, "Information!": Exit Sub
ElseIf Thread2 <> 0 Then
MsgBox "Waitable Thread allready Exist!", vbExclamation, "Information!": Exit Sub
End If

ThreadHandle = MULTITHREADER.CreateNewThread(Thread2, &HC000&, THREAD_PRIORITY_NORMAL, Thread1, 2, Empty, ObjectEnabled)
CloseHandle ThreadHandle
End Sub

Private Sub Command4_Click()
Dim IsValidCall As Boolean

If Thread1 = 0 Then MsgBox "Signal Thread doesnt exist!", vbCritical, "Main Thread!": Exit Sub
MULTITHREADER.CallThread Thread1, 1, 0, Empty, [UsingMove], UsingAPC, IsValidCall


End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close Application!

MULTITHREADER.CloseCaller App.HInstance
MULTITHREADER.RemoveThread 0 'Exit Application!
End Sub



Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long
Dim ThreadId As Long

If ThreadNotify < 0 Then
    'worker thread termination notify
    ThreadId = ThreadNotify And &H7FFFFFFF
    AddLine Text1, "Thread Id:" & ThreadId & " destroyed!"
    If ThreadId = Thread1 Then
        Thread1 = 0
    ElseIf ThreadId = Thread2 Then
        Thread2 = 0
    End If
    
Else
    'worker thread creation notify
    ThreadId = ThreadNotify
    If App.ThreadId = ThreadId Then
    AddLine Text1, "Main Thread Id:" & ThreadId & " multithread initialized!"
    Else
    AddLine Text1, "Thread Id:" & ThreadId & " created!"
    End If
End If


End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long
Dim Message As Long
Dim Reason As Long
Dim Args As Variant
'Translate call !!!
MULTITHREADER.TranslateArguments CallArgs, Reason, Message, Args
MsgBox Args(0) & vbCrLf & Args(1) & vbCrLf & Args(2), vbInformation, "Main Thread incomming Arguments!"
End Function


Private Sub AddLine(TXT As TextBox, ByRef TextX As String)
TXT.SelLength = 0
TXT.SelStart = 0
TXT.SelText = TextX & vbCrLf
End Sub
