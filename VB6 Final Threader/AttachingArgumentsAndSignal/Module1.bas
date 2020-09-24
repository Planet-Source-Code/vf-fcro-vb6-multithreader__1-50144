Attribute VB_Name = "Module1"
Option Explicit
Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public MULTITHREADER As New Threader

Public F1 As Form1
Public F2 As Form2
Public Thread1 As Long
Public Thread2 As Long

Sub Main()
Static IsInit As Boolean

If Not IsInit Then
    Dim IsUpdated As Boolean
    IsInit = True
    Set F1 = New Form1
    Thread1 = App.ThreadId
    Dim THandle As Long
    GatherObject MULTITHREADER
    
    IsUpdated = MULTITHREADER.UpdateVirtualMachine
    
    If IsUpdated Then
    MsgBox "Virtual Machine updated!", vbExclamation, "Information"
    Else
    MsgBox "This version of Virtual Machine not supported!", vbCritical, "Info!"
    End If
    
    MULTITHREADER.InitCaller App.HInstance
    MULTITHREADER.AddThread F1
    F1.Show
    THandle = MULTITHREADER.CreateNewThread(Thread2, &HC000&, THREAD_PRIORITY_NORMAL, 0, 0, Empty, ObjectEnabled)
    CloseHandle THandle
Else
    Dim Reason As Long
    Dim Message As Long
    Dim Arguments As Variant
    Set F2 = New Form2
    MULTITHREADER.GetThreadParams Reason, Message, Arguments
    F2.Show
    MULTITHREADER.AddThread F2
    
End If


End Sub



