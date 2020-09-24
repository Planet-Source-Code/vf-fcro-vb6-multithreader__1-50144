Attribute VB_Name = "MainModule"
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Public MULTITHREADER As New Threader
Dim FMAIN As Form1
Sub Main()
Static IsInit As Boolean

If Not IsInit Then
    Dim IsUpdated As Boolean
    'Main Thread
    IsInit = True
    GatherObject MULTITHREADER 'First Step required!
    MULTITHREADER.UpdateVirtualMachine
    MULTITHREADER.InitCaller App.HInstance 'Second Step required!
    
    Set FMAIN = New Form1
    MULTITHREADER.AddThread FMAIN
    FMAIN.Show
    
Else
    
    'New Threads with Start Params
    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long

    MULTITHREADER.GetThreadParams Reason, Message, Args 'Get Thread Parameters! {U can get it only once!}
    
    MsgBox "Thread Start Parameters:" & vbCrLf & "Reason:" & _
    Reason & vbCrLf & "Message:" & Message & vbCrLf & "Argument:" & Args, vbExclamation, "ThreadId:" & App.ThreadId
    MULTITHREADER.RemoveThread App.ThreadId
    End
    
End If


End Sub
