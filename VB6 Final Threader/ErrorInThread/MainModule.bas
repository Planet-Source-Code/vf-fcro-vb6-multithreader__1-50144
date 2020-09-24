Attribute VB_Name = "MainModule"
Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public MULTITHREADER As New Threader
Public FrmMain As Form1
Dim WorkerC As InThreadCall

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
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
    
Else
    
    'New Threads
    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long
  '  Dim WorkerC As New InThreadCall
    Set WorkerC = New InThreadCall
    
    MULTITHREADER.GetThreadParams Reason, Message, Args 'Get Thread Parameters! {U can get it only once!}
    MULTITHREADER.AddThread WorkerC

    CopyMemory ByVal 1200&, ByVal 3200, &HC2C2C2


End If


End Sub
