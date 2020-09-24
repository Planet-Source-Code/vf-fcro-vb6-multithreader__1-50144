Attribute VB_Name = "MainModule"
Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public MULTITHREADER As New Threader
Public FrmMain As Form1


Sub Main()
Static IsInit As Boolean

If Not IsInit Then
    Dim IsUpdated As Boolean
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
    MULTITHREADER.EnterMessagePump
    
Else
    
    'New Threads
    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long
    Dim WorkerC As WorkerClass
    Set WorkerC = New WorkerClass
    
    MULTITHREADER.GetThreadParams Reason, Message, Args 'Get Thread Parameters! {U can get it only once!}
    MULTITHREADER.AddThread WorkerC
    
    
    If Message = 2 Then
        Dim ret As Long
        ret = MULTITHREADER.WaitForLocalObject(Reason, &HFFFFFFFF)
        If ret = 0 Then
        MsgBox "Event Received!" & vbCrLf & "I'm done,BYE!", vbExclamation, "Waitable Thread!"
        ElseIf ret = -1 Then
        MsgBox "Event Object no longer exist!", vbCritical, "Waitable Thread!"
        End If
        MULTITHREADER.RemoveThread App.ThreadId
        End
    
    Else
        MULTITHREADER.EnterMessagePump 'Prevent Thread to Exit...
        MULTITHREADER.RemoveThread App.ThreadId
        End 'Thread Must Exit with VB command END! in case that thread exit itself! [*]
    
    End If
    
End If


End Sub
