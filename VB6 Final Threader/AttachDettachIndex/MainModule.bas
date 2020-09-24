Attribute VB_Name = "MainModule"
Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public MULTITHREADER As New Threader
Public FrmMain As Form1


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
    
    MULTITHREADER.WaitForLocalObject 0, -1
 
    MULTITHREADER.DetachThreadCallArguments 0, 12, Reason, Message, Args, True
    MsgBox Args, , "Arguments Detached and Removed by Thread:" & App.ThreadId
    MULTITHREADER.RemoveThread App.ThreadId
    End
End If


End Sub
