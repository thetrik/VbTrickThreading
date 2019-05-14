VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CThreadHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Form As Object

Private m_cContext  As CContextHolder

Public Sub CreateFormFromExportedFunction()
    
    Set m_cContext = New CContextHolder
    
    ' // Init DLL project context
    If Not m_cContext.InitVbProjContext(App.Path & "\dll\TestDll.dll") Then
        MsgBox "Unable to init context"
        Exit Sub
    End If
    
    ' // Now we can call any exported functions without restrictions
    Set Form = m_cContext.CallFunc("CreateForm", vbObject)
    
End Sub

Private Sub Class_Terminate()
    
    ' // Unload form
    If Not Form Is Nothing Then
        Unload Form
    End If
    
    ' // Break the reference
    Set Form = Nothing
    
End Sub
