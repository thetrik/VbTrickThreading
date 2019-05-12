VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private object marshaling by The trick"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFreeList 
      Caption         =   "Free"
      Height          =   435
      Left            =   4980
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   435
      Left            =   4980
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CheckBox chkSynch 
      Caption         =   "Asynch"
      Height          =   315
      Left            =   4980
      TabIndex        =   6
      Top             =   3060
      Width           =   1695
   End
   Begin VB.CommandButton cmdSleep 
      Caption         =   "Sleep"
      Height          =   435
      Left            =   4980
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdThreadID 
      Caption         =   "Thread ID"
      Height          =   435
      Left            =   4980
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtLog 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4920
      Width           =   6495
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   435
      Left            =   4980
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddObject 
      Caption         =   "Add"
      Height          =   435
      Left            =   4980
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox lstObjects 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Private marshaling example
' // By The trick 2018
' // User can create a private object in other thread and call its method
' // (synchronously/asynchronously). In IDE all objects live ein the main thread
' //

Option Explicit

Private mcObjList   As Collection   ' // List of objects
Private mcIDs       As Collection   ' // List of object identifiers (used for asynch call)

' // Log message to textbox
Public Sub Log( _
           ByVal sText As String)
    Dim lPrev   As Long
    
    lPrev = Len(txtLog.Text)
    
    txtLog.Text = txtLog.Text & time & ": [0x" & Hex$(App.ThreadID) & "]: " & sText & vbNewLine
    txtLog.SelStart = lPrev + 2
    
End Sub

' // This is callback function of ThreadID proprety
' // When asynch method has been called this function is called
Public Sub ThreadID_CallBack( _
           ByVal vRet As Variant)
    Log "(CallBack ThreadID_CallBack) Returned value: 0x" & Hex$(vRet)
End Sub

' // This is callback function of Sleep method
Public Sub Sleep_CallBack( _
           ByVal vRet As Variant)
    Log "(CallBack Sleep_CallBack) Returned value: 0x" & Hex$(vRet)
End Sub

' // This is callback function of Process method
' // Notice, here we lost [out] string output parameter
Public Sub Process_CallBack( _
           ByVal vRet As Variant)
    Log "(CallBack Process_CallBack) Returned value: 0x" & Hex$(vRet)
End Sub

' // Create CPrivateClass instance and add it to list
Private Sub cmdAddObject_Click()
    Dim cNewObj     As Object
    Dim lAsynchId   As Long
    
    Set cNewObj = CreatePrivateObjectByNameInNewThread("CPrivateClass", , lAsynchId)
    
    ' // Set log form
    Set cNewObj.LogForm = Me

    mcObjList.Add cNewObj
    mcIDs.Add lAsynchId
    
    lstObjects.AddItem Hex$(lAsynchId) & " {ThreadID 0x" & Hex$(cNewObj.ThreadID) & "}"
    
End Sub

' // Free all object
Private Sub cmdFreeList_Click()

    Set mcObjList = New Collection
    Set mcIDs = New Collection
    lstObjects.Clear
    
End Sub

' // Remove object from list
Private Sub cmdRemove_Click()
    Dim lIndex  As Long
    
    lIndex = lstObjects.ListIndex + 1
    
    If lIndex <= 0 Then Exit Sub
    
    lstObjects.RemoveItem lIndex - 1
    mcIDs.Remove lIndex
    mcObjList.Remove lIndex
    
End Sub

' // Call CPrivateClass:Process
Private Sub cmdProcess_Click()
    Dim cObj        As Object
    Dim lId         As Long
    Dim lRet        As Long
    Dim sInput      As String
    Dim lArray()    As Long
    Dim lIndex      As Long
    
    Set cObj = GetSelectedObject(lId)
    
    If cObj Is Nothing Then Exit Sub
    
    ReDim lArray(10)
    
    For lIndex = 0 To UBound(lArray)
        lArray(lIndex) = lIndex
    Next
    
    sInput = "Hello from 0x" & Hex(App.ThreadID) & " thread!"
    
    If chkSynch.Value = vbChecked Then
            
        AsynchDispMethodCall lId, "Process", VbMethod, Me, "Process_CallBack", lArray(), sInput
            
    Else
    
        lRet = cObj.Process(lArray, sInput)
        Log "Returned value: 0x" & lRet & "; InOutString: '" & sInput & "'"
        
    End If
    
End Sub

' // Call Slep method
Private Sub cmdSleep_Click()
    Dim cObj    As Object
    Dim lId     As Long
    Dim lRet    As Long
    
    Set cObj = GetSelectedObject(lId)
    
    If cObj Is Nothing Then Exit Sub
    
    If chkSynch.Value = vbChecked Then
            
        AsynchDispMethodCall lId, "Sleep", VbMethod, Me, "Sleep_CallBack"
            
    Else
    
        lRet = cObj.Sleep
        Log "Returned value: 0x" & Hex$(lRet)
        
    End If
    
End Sub

' // Call ThreadID property
Private Sub cmdThreadID_Click()
    Dim cObj    As Object
    Dim lId     As Long
    Dim lRet    As Long
    
    Set cObj = GetSelectedObject(lId)
    
    If cObj Is Nothing Then Exit Sub
    
    If chkSynch.Value = vbChecked Then
            
        AsynchDispMethodCall lId, "ThreadID", VbGet, Me, "ThreadID_CallBack"
            
    Else
    
        lRet = cObj.ThreadID
        Log "Returned value: 0x" & Hex$(lRet)
        
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Caption = Me.Caption & " [ThreadID: 0x" & Hex$(App.ThreadID) & "]"
    
    modMultiThreading.Initialize
    
    ' // Enable private marshaling of VB6 objects
    modMultiThreading.EnablePrivateMarshaling True
    
    Set mcObjList = New Collection
    Set mcIDs = New Collection
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim vId As Variant
    
    Set mcObjList = Nothing
    
    For Each vId In mcIDs
        WaitForObjectThreadCompletion vId
    Next
    
    Set mcIDs = Nothing
    
    modMultiThreading.Uninitialize
    
End Sub

' // Get current selected object
Private Function GetSelectedObject( _
                 Optional ByRef lAsynchId As Long) As Object
    Dim lIndex  As Long
    
    lIndex = lstObjects.ListIndex + 1
    lAsynchId = 0
    
    If lIndex <= 0 Then Exit Function
    
    Set GetSelectedObject = mcObjList(lIndex)
    lAsynchId = mcIDs(lIndex)
    
End Function

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    bValue = True
    MakeTrue = True
End Function

