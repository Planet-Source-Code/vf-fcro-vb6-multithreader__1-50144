VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form2"
   ScaleHeight     =   570
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reply something"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall


Private Sub Command1_Click()
Dim PrevCallArg As Long
MULTITHREADER.AttachThreadCallArguments 0, 0, 0, 0, Array("Argument Back [FIRST]", "Argument Back [LAST]!"), UsingMove, PrevCallArg, False
MULTITHREADER.SignalLocalObject Thread2
End Sub

Private Sub Form_Load()
Caption = "Thread Id:" & App.ThreadId


End Sub

Private Sub Form_Unload(Cancel As Integer)
MULTITHREADER.RemoveThread App.ThreadId 'Remove Thread
End
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long
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
    MULTITHREADER.WaitForLocalObject 0, -1
    MULTITHREADER.DetachThreadCallArguments 0, 0, Reason, Message, Arguments, True
    Caption = "Signaled!  Thread Id:" & App.ThreadId
    MsgBox Arguments(0) & vbCrLf & Arguments(1), vbExclamation, "Detached information by ThreadId:" & App.ThreadId
    MULTITHREADER.CallThread 0, 0, 1, Empty, UsingCopy, UsingAPC, IsValidCall
End If

End Function

