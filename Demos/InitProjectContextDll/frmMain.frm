VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DLL callback threading example by The trick"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.Label ldlDesc 
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   6075
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cThreads  As Collection       ' // Threads list which holders the forms in own threads
Private m_cIds      As Collection       ' // IDs for async call or waiting for thread completion

Private Sub Form_Initialize()
    modMultiThreading.Initialize
    modMultiThreading.EnablePrivateMarshaling True
End Sub

Private Sub Form_Load()
    Dim cThread     As Object
    Dim lIndex      As Long
    Dim lId         As Long
    Dim bIsInIDE    As Boolean
    
    Set m_cThreads = New Collection
    Set m_cIds = New Collection
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    ' // Create 5 threads with the CThreadHolder instance per thread.
    ' // Each instance lives in its own STA. We get the marshalled reference to the objects
    
    For lIndex = 0 To 4
    
        Set cThread = CreatePrivateObjectByNameInNewThread("CThreadHolder", , lId)
        
        ' // Create form from DLL in the object's STA
        cThread.CreateFormFromExportedFunction
        
        ' // Setup callback
        If bIsInIDE Then
            cThread.Form.SetCallback InitCurrentThreadAndCallFunctionIDEProc(AddressOf CallbackProcInit, 0)
        Else
            cThread.Form.SetCallback AddressOf CallbackProcEXE
        End If
        
        m_cThreads.Add cThread
        m_cIds.Add lId
        
    Next

End Sub

Private Sub Form_Terminate()
    modMultiThreading.Uninitialize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim vId As Variant
    
    ' // Breaks all the references
    Set m_cThreads = Nothing
    
    ' // Wait for all the threads completion
    For Each vId In m_cIds
        WaitForObjectThreadCompletion vId
    Next
    
End Sub

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True: bValue = True
End Function
