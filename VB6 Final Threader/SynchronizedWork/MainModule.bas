Attribute VB_Name = "MainModule"
Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public MULTITHREADER As New Threader
Dim MultithreadedF As Form1

Public GLOBALCOUNTER As Long



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
    MultithreadedF.Show
    


Else
    
    'New Threads
    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long
    Dim DUMMY As New InThreadCall
    MULTITHREADER.GetThreadParams Reason, Message, Args 'Get Thread Parameters! {U can get it only once!}
    
    
    MULTITHREADER.AddThread DUMMY
    
    Do
    MULTITHREADER.EnterSynchronization False 'Enter SYNC work!
    
    If GLOBALCOUNTER = 50 Then
    GLOBALCOUNTER = 0
    Args.Caption = "Thread Id:" & App.ThreadId & ",reset counter to " & GLOBALCOUNTER
    Else
    GLOBALCOUNTER = GLOBALCOUNTER + 1
    Args.Caption = "Thread Id:" & App.ThreadId & ",increase value to " & GLOBALCOUNTER
    End If
    
    Sleep 125
    MULTITHREADER.LeaveSynchronization 'Exit SYNC work!
    
    Loop
    
    


End If


End Sub
