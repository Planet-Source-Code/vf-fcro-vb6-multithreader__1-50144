Attribute VB_Name = "MainModule"
Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public MULTITHREADER As New Threader
Public FrmMain As Form1
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Sub Main()
Static IsInit As Boolean
Dim IsUpdated As Boolean
If Not IsInit Then
    
    'Main Thread
    IsInit = True
    Set FrmMain = New Form1
    GatherObject MULTITHREADER 'First Step required!
    IsUpdated = MULTITHREADER.UpdateVirtualMachine
    
    If IsUpdated Then
    MsgBox "Virtual Machine updated!", vbExclamation, "Information"
    Else
    MsgBox "This version of Virtual Machine not supported!", vbCritical, "Info!"
    End If

    MULTITHREADER.InitCaller App.HInstance 'Second Step required!
    MULTITHREADER.AddThread FrmMain
    FrmMain.Show
    
Else
    
    'New Threads
    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long
    Dim WorkerC As New InThreadCall

    
    MULTITHREADER.GetThreadParams Reason, Message, Args 'Get Thread Parameters! {U can get it only once!}
    MULTITHREADER.AddThread WorkerC
    MULTITHREADER.EnterSynchronization False
    
    Dim u As Long
    For u = 0 To 9
    Beep 1222, 12
    Sleep 600
    Next u
    
    MULTITHREADER.LeaveSynchronization
    MULTITHREADER.RemoveThread App.ThreadId
    End

End If


End Sub
