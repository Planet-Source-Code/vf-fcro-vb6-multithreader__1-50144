VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Thread Synchronization!"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0AAA9&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CA9273&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0036D877&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall




Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim TID As Long
Dim THREADHANDLE As Long
Dim u As Long

For u = 0 To 4
THREADHANDLE = MULTITHREADER.CreateNewThread(TID, &HC000&, THREAD_PRIORITY_NORMAL, 0, 0, CVar(Label1(u)), ObjectEnabled)
CloseHandle THREADHANDLE
Next u
End Sub

Private Sub Form_Unload(Cancel As Integer)

MULTITHREADER.CloseCaller App.HInstance
MULTITHREADER.RemoveThread 0

'Each thread must clean itself from multithreader structure!
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long

End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long

End Function





