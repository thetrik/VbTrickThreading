VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Callback test by The Trick"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFree 
      Caption         =   "Free"
      Height          =   495
      Left            =   4860
      TabIndex        =   4
      Top             =   1200
      Width           =   1515
   End
   Begin VB.ListBox lstCallbackLog 
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      Top             =   2580
      Width           =   6195
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove callback"
      Height          =   495
      Left            =   4860
      TabIndex        =   2
      Top             =   660
      Width           =   1515
   End
   Begin VB.CommandButton cmdAddCallback 
      Caption         =   "Add callback..."
      Height          =   495
      Left            =   4860
      TabIndex        =   1
      Top             =   120
      Width           =   1515
   End
   Begin VB.ListBox lstCallback 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4635
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Callback example
' // By The Trick 2018
' // This code demonstrates the user callback function that is called in differents threads
' // In IDE it transmits all calls to main thread
' //

Option Explicit

' // Log data to listbox
Public Sub Log( _
           ByRef sText As String)
    lstCallbackLog.AddItem time & ": [0x" & Hex$(App.ThreadID) & "] " & sText
    lstCallbackLog.ListIndex = lstCallbackLog.ListCount - 1
End Sub

' // Add a callback
Private Sub cmdAddCallback_Click()
    Dim sInterval   As String
    Dim sKey        As String
    Dim lInterval   As Long
    Dim lId         As Long
    Dim bIsInIDE    As Boolean
    
    On Error GoTo err_handler
    
    Debug.Assert MakeTrue(bIsInIDE)

    sInterval = InputBox("Enter interval")
    If StrPtr(sInterval) = 0 Then Exit Sub
    
    sKey = InputBox("Enter key")
    lInterval = Val(sInterval)
    
    ' // In IDE we should use asm-thunks which transmit calls to main thread
    If bIsInIDE Then
        lId = SetCallback(InitCurrentThreadAndCallFunctionIDEProc(AddressOf CallBackProc_user, 12), lInterval, StrPtr(sKey))
    Else
        lId = SetCallback(AddressOf CallbackProc, lInterval, StrPtr(sKey))
    End If
    
    If lId < 0 Then
        MsgBox "Unable to add callback"
        Exit Sub
    End If

    lstCallback.AddItem "'" & sKey & ";' ID: 0x" & Hex$(lId) & ": {" & time & "}; Interval: " & lInterval
    lstCallback.ItemData(lstCallback.NewIndex) = lId
    
err_handler:
    
End Sub

' // Free callback
Private Sub cmdFree_Click()

    FreeAll
    lstCallback.Clear
    lstCallbackLog.Clear
    
End Sub

' // Remove seelected callback
Private Sub cmdRemove_Click()
    Dim lIndex  As Long
    Dim lId     As Long
    
    lIndex = lstCallback.ListIndex
    If lIndex = -1 Then Exit Sub
    
    lId = lstCallback.ItemData(lIndex)
    lstCallback.RemoveItem lIndex
    StopCallback lId
    
End Sub

Private Sub Form_Load()
    Dim cObj    As Object
    
    Me.Caption = Me.Caption & " [ThreadID: 0x" & Hex$(App.ThreadID) & "]"
    
    modMultiThreading.Initialize
    
    Set cObj = Me
    
    ' // Make multiple marshaling stream to call Log method
    ' // We can unmarshal it multiple times in the different threads
    gpMarshalData = Marshal2(cObj)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' // Clean up
    FreeAll
    FreeMarshalData gpMarshalData
    modMultiThreading.Uninitialize
    
End Sub

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True: bValue = True
End Function


