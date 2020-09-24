VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   90
      Left            =   360
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create New Thread"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall
Public IsMainThread As Boolean
Public XVAL As Long

Private Sub Command1_Click()
Dim TID As Long
Dim THREADHANDLE As Long
THREADHANDLE = MULTITHREADER.CreateNewThread(TID, &HC000&, THREAD_PRIORITY_NORMAL, 0, 0, Empty, ObjectEnabled)
CloseHandle THREADHANDLE
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Caption = "Thread Id:" & App.ThreadId
If IsMainThread Then Caption = Caption & "(MAIN)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If IsMainThread Then
MULTITHREADER.CloseCaller App.HInstance
MULTITHREADER.RemoveThread 0
Else
MULTITHREADER.RemoveThread App.ThreadId
End If
'Each thread must clean itself from multithreader structure!
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long

End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long

End Function

Public Sub SetValues(ByVal Interval_ As Integer, ByVal Color_ As Long)
Timer1.Interval = Interval_
XVAL = Color_
End Sub


Private Sub Timer1_Timer()
Static XP As Boolean

If XP Then
    XP = False
    Picture1.BackColor = 0
Else
    XP = True
    Picture1.BackColor = XVAL
End If



End Sub
