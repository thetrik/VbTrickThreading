VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Marshal user interface by The trick"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowMsg 
      Caption         =   "Show message..."
      Height          =   435
      Left            =   4980
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.ListBox lstObjects 
      Height          =   4740
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdAddObject 
      Caption         =   "Add"
      Height          =   435
      Left            =   4980
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   435
      Left            =   4980
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtLog 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   4920
      Width           =   6495
   End
   Begin VB.CommandButton cmdThreadID 
      Caption         =   "Thread ID"
      Height          =   435
      Left            =   4980
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdSetValue 
      Caption         =   "Set value..."
      Height          =   435
      Left            =   4980
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdGetValue 
      Caption         =   "Get value"
      Height          =   435
      Left            =   4980
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdFreeList 
      Caption         =   "Free"
      Height          =   435
      Left            =   4980
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // User interface marshaling example
' // By The trick 2018
' // User can create a private object in other thread and call its method by user interface
' // In IDE all objects live in the main thread
' // This example uses type library (user_typelib.tlb) with interfaces definitions (IUserInterface, ILogObject)
' // It uses Reg-Free manifest (in resources) in compiled form (manifest.xml)
' //

Option Explicit

Private mcObjList   As Collection   ' // List of objects
Private mcIDs       As Collection   ' // List of object identifiers (used for asynch call)

Implements ILogObject

' // Create CPrivateClass instance and add it to list
Private Sub cmdAddObject_Click()
    Dim cNewObj     As IUserInterface
    Dim lAsynchId   As Long
    
    ' // Create object in new thread and marshal interface pointer (IUserInterface)
    Set cNewObj = CreatePrivateObjectByNameInNewThread("CUserClass", VarPtr(IID_IUserInterface), lAsynchId)
    
    ' // Set log form
    cNewObj.SetLogObject Me

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

Private Sub cmdGetValue_Click()
    Dim cObj    As IUserInterface
    Dim vRet    As Variant
    
    Set cObj = GetSelectedObject()
    If cObj Is Nothing Then Exit Sub
    
    vRet = cObj.Value
    ILogObject_Log "Returned value: " & vRet

End Sub

Private Sub cmdSetValue_Click()
    Dim cObj    As IUserInterface
    Dim lRet    As Long
    Dim sValue  As String
    
    Set cObj = GetSelectedObject()
    If cObj Is Nothing Then Exit Sub
    
    sValue = InputBox("Enter value")
    If StrPtr(sValue) = 0 Then Exit Sub
    
    cObj.Value = sValue

End Sub

Private Sub cmdShowMsg_Click()
    Dim cObj    As IUserInterface
    Dim lRet    As Long
    Dim sMsg    As String
    
    Set cObj = GetSelectedObject()
    If cObj Is Nothing Then Exit Sub
    
    sMsg = InputBox("Enter message")
    If StrPtr(sMsg) = 0 Then Exit Sub
    
    cObj.ShowMessage sMsg
    
End Sub

Private Sub cmdThreadID_Click()
    Dim cObj    As IUserInterface
    Dim lRet    As Long
    
    Set cObj = GetSelectedObject()
    If cObj Is Nothing Then Exit Sub
    
    lRet = cObj.ThreadID
    ILogObject_Log "Returned value: 0x" & Hex$(lRet)

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
Private Function GetSelectedObject() As IUserInterface
    Dim lIndex  As Long
    
    lIndex = lstObjects.ListIndex + 1

    If lIndex <= 0 Then Exit Function
    
    Set GetSelectedObject = mcObjList(lIndex)
    
End Function

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    bValue = True
    MakeTrue = True
End Function

Private Sub ILogObject_Log( _
            ByVal sMsg As String)
    Dim lPrev   As Long
    
    lPrev = Len(txtLog.Text)
    
    txtLog.Text = txtLog.Text & time & ": [0x" & Hex$(App.ThreadID) & "]: " & sMsg & vbNewLine
    txtLog.SelStart = lPrev + 2
    
End Sub
