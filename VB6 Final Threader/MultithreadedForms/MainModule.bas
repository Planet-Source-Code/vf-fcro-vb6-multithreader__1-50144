Attribute VB_Name = "MainModule"
Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public MULTITHREADER As New Threader
Dim MultithreadedF As Form1

Sub Main()
Static IsInit As Boolean

'************************
'REQUIRED INITIALIZATION!
'************************

If Not IsInit Then
    Dim IsUpdated As Boolean
    'Main Thread
    IsInit = True
    GatherObject MULTITHREADER 'First Step required!
    IsUpdated = MULTITHREADER.UpdateVirtualMachine 'Required!
    MULTITHREADER.InitCaller App.HInstance 'Second Step required!
    
    Set MultithreadedF = New Form1
    MULTITHREADER.AddThread MultithreadedF
    MultithreadedF.IsMainThread = True
    
    Randomize
    MultithreadedF.SetValues Int(Rnd * 300) + 20, Int(Rnd * &HFFFFFF)
    MultithreadedF.Show
    Randomize

Else
    
    'New Threads
    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long
    
    MULTITHREADER.GetThreadParams Reason, Message, Args 'Get Thread Parameters! {U can get it only once!}
    
    Set MultithreadedF = New Form1 'Create New
    MULTITHREADER.AddThread MultithreadedF
    
    Randomize
    MultithreadedF.SetValues Int(Rnd * 400) + 20, Int(Rnd * &HFFFFFF)
    
    MultithreadedF.Show
    Randomize

End If


End Sub
