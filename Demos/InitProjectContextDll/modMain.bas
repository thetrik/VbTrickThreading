Attribute VB_Name = "modMain"
Option Explicit

Public Sub CallbackProc()
    MsgBox "CallbackProc in EXE is being called from DLL ThreadID:0x" & Hex$(App.ThreadID)
End Sub

' // This function is called from initialized thread
Public Sub CallbackProcInit( _
           ByVal lUnused As Long)
    CallbackProc
End Sub

' // This function is called from DLL in compiled EXE
Public Sub CallbackProcEXE()
    ' // Initialize project context and call function CallbackProcInit
    InitCurrentThreadAndCallFunction AddressOf CallbackProcInit, 0, 0
End Sub
