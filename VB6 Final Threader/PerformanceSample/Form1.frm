VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "InterThread Communication - Test Performace"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "InterThread Communcation direct"
      Height          =   5295
      Left            =   4080
      TabIndex        =   11
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox Check1 
         Caption         =   "Update at the End"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Non Synchronized"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Synchronized"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Create 5 Threads"
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"Form1.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "InterThread Communication through Message Loop"
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox Check3 
         Caption         =   "Synchronized"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Direct Call Object From Different Thread"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         Caption         =   "Arguments Transfer"
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   2415
         Begin VB.OptionButton Option3 
            Caption         =   "Move Data"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Copy Data"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Update at the End"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create 5 Threads"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Using PostMessage"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Using CallBacks"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Using APC"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Threading Concept:Each created thread send information to main thread [or any thread]  [THREAD SAFE COMMUNICATION]"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   4200
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements InThreadCall 'multithread communication interface







Private Sub Check1_Click()
Update1 = Check1.Value
End Sub

Private Sub Check2_Click()
Update2 = Check2.Value
End Sub

Private Sub Check3_Click()
SyncIt2 = CLng(Check3.Value)
End Sub

Private Sub Command1_Click()
Dim ThreadH As Long
Dim ThreadId As Long

Dim u As Long
For u = 0 To 4
Text1(u) = ""
Next u
DoEvents


For u = 0 To 4
ThreadH = MULTITHREADER.CreateNewThread(ThreadId, &HC000&, THREAD_PRIORITY_NORMAL, u, 1, Array(Text1(u), SyncIt2), ObjectEnabled)
CloseHandle ThreadH
Next u
End Sub

Private Sub Command2_Click()
Dim ThreadH As Long
Dim ThreadId As Long
Dim u As Long
For u = 0 To 4
Text2(u) = ""
Next u
DoEvents


For u = 0 To 4
ThreadH = MULTITHREADER.CreateNewThread(ThreadId, &HC000&, THREAD_PRIORITY_NORMAL, 1, SyncIt, Text2(u), ObjectEnabled)
CloseHandle ThreadH
Next u

End Sub

Private Sub Form_Load()
Option1(0).Value = True
Option2(0).Value = True
Option3(0).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MULTITHREADER.CloseCaller App.HInstance 'Remove itself and clean internals
MULTITHREADER.RemoveThread 0 'Exit Application
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long
End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long
Dim Reason As Long
Dim Message As Long
Dim Arguments As Variant
Dim TC As Long
MULTITHREADER.TranslateArguments CallArgs, Reason, Message, Arguments

If Arguments(1) = 1 Then


    If CLng(Arguments(3)) = 8000& Then
    TC = GetTickCount
    Arguments(0).Text = "Thread Id:" & App.ThreadId & " ,Execute Time:" & TC - CLng(Arguments(2)) & " msec"
    Else
    
    If Update2 Then Exit Function
    
    Arguments(0).Text = "Thread Id:" & Reason & Space(10) & ",Count:" & Message
    End If


    If SendType = 2 Then
    'DoEvents with Posting not allowed!!!!!
    Arguments(0).Refresh
    ElseIf SendType < 2 Then
    MULTITHREADER.FastDoEvents Arguments(0).hWnd

    End If

End If

InThreadCall_ThreadCall = &H9009&
End Function




Private Sub Option1_Click(Index As Integer)
SendType = Index
End Sub

Private Sub Option2_Click(Index As Integer)
SyncIt = Index + 2
End Sub

Private Sub Option3_Click(Index As Integer)
TransferType = Index
End Sub
