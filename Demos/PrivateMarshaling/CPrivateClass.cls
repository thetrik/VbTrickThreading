VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrivateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CPrivateClass for marshaling demonstration
' //

Option Explicit

Private Declare Function WinSleep Lib "kernel32" Alias "Sleep" ( _
                         ByVal dwMilliseconds As Long) As Long

Private mcLogForm  As Object    ' // Marshaled instance (you can see it in Log method)

' // Set logging form
Public Property Set LogForm( _
                    ByVal cValue As Object)
    Log "CPrivateClass::LogForm_set"
    Set mcLogForm = cValue
End Property

Public Property Get ThreadID() As Long
    Log "CPrivateClass::ThreadID_get"
    ThreadID = App.ThreadID
End Property

Public Sub Sleep()
    Log "CPrivateClass::Sleep"
    WinSleep 3000
End Sub

Public Function Process( _
                ByRef lArray() As Long, _
                ByRef sInOutString As String) As Long
    Dim lRet    As Long
    Dim lIndex  As Long
    
    Log "CPrivateClass::Process InString: '" & sInOutString & "'"
    
    ' // Add all items
    For lIndex = 0 To UBound(lArray)
        lRet = lRet + lArray(lIndex)
    Next
    
    sInOutString = "Returned from 0x" & Hex$(App.ThreadID)
    
    Process = lRet
    
End Function

Private Sub Class_Terminate()
    Log "CPrivateClass::~CPrivateClass"
End Sub

Private Sub Log( _
            ByRef sText As String)
    
    If Not mcLogForm Is Nothing Then
        mcLogForm.Log sText & " {ThreadID: 0x" & Hex$(App.ThreadID) & "}"
    End If
    
End Sub

