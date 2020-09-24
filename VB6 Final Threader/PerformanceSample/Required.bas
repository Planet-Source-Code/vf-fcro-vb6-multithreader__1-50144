Attribute VB_Name = "Test"
'Required!!!
Option Explicit
Declare Sub GatherObject Lib "vb6multithread.dll" (Obj As Object)

Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public MULTITHREADER As New Threader
Public F1 As Form1

Public SyncIt As Long
Public SyncIt2 As Long

Public SendType As Long
Public Update1 As Boolean
Public Update2 As Boolean
Public TransferType As Long

'PROJECT MUST START FROM SUB-MAIN!!!
'KEEP MAIN THREAD ALIVE!!!


Sub Main()
Static IsInit As Boolean


If Not IsInit Then
Dim IsUpdated As Boolean
IsInit = True

    'Main Thread
    GatherObject MULTITHREADER 'Execute this always as first step with MAIN THREAD
    IsUpdated = MULTITHREADER.UpdateVirtualMachine
    
    If IsUpdated Then
    MsgBox "Virtual Machine updated!", vbExclamation, "Information"
    Else
    MsgBox "This version of Virtual Machine not supported!", vbCritical, "Info!"
    End If

    
    
    MULTITHREADER.InitCaller App.HInstance '2nd required step....
  
    Set F1 = New Form1

    MULTITHREADER.AddThread F1 'Set object which implements thread communicator

    F1.Show
    MULTITHREADER.AboutBox
Else

Dim u As Long
Dim IsValidCall As Boolean
    'New Threads!
Dim Message As Long
Dim Reason As Long
Dim Args As Variant
Dim TC As Long
Dim Ret As Long

MULTITHREADER.GetThreadParams Reason, Message, Args 'Get Thread Parameters!


MULTITHREADER.AddThread F1

If Message = 1 Then
TC = GetTickCount


For u = 0 To 8000
    
 If Args(1) = 1 Then MULTITHREADER.EnterSynchronization False
   Ret = MULTITHREADER.CallThread(0, App.ThreadId, u, Array(Args(0), 1, TC, u), TransferType, SendType, IsValidCall)
 If Args(1) = 1 Then MULTITHREADER.LeaveSynchronization
    Next u
    MULTITHREADER.RemoveThread App.ThreadId
    End 'Thread Must Exit with END!!!!
    
    
Else

TC = GetTickCount
    
    For u = 0 To 8000
   
If Message = 3 Then MULTITHREADER.EnterSynchronization False


If Not Update1 Then
'Update each one!
    Args.Text = "Thread Id:" & App.ThreadId & Space(10) & ",Count:" & u
    Args.Refresh
End If


If Message = 3 Then MULTITHREADER.LeaveSynchronization

   
    Next u
    
TC = GetTickCount - TC
    Args.Text = "Thread Id:" & App.ThreadId & " ,Execute Time:" & TC & " msec"
    Args.Refresh
    
    MULTITHREADER.RemoveThread App.ThreadId
    End 'Thread Must Exit with END!!!!
    
End If



End If



End Sub
