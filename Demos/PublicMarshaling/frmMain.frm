VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Public object marshaling by The trick"
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
   Begin VB.CommandButton cmdDicItems 
      Caption         =   "Dictionary::Items"
      Height          =   435
      Left            =   4980
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdDicKeys 
      Caption         =   "Dictionary::Keys"
      Height          =   435
      Left            =   4980
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdFreeList 
      Caption         =   "Free"
      Height          =   435
      Left            =   4980
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdDicRemove 
      Caption         =   "Dictionary::Remove"
      Height          =   435
      Left            =   4980
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox chkSynch 
      Caption         =   "Asynch"
      Height          =   315
      Left            =   4980
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdDicAdd 
      Caption         =   "Dictionary::Add"
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
' // Public marshaling example
' // By The trick 2018
' // User can create a public ActiveX object in other thread and call its method
' // (synchronously/asynchronously). In IDE all objects live in the main thread
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

Public Sub Add_Callback( _
           ByVal vRet As Variant)
    Log "Add_Callback; Return = " & vRet
End Sub

Public Sub Remove_Callback( _
           ByVal vRet As Variant)
    Log "Remove_Callback; Return = " & vRet
End Sub

Public Sub Items_Callback( _
           ByVal vRet As Variant)
    Log "Items_Callback; Return = " & DumpItems(vRet)
End Sub

Public Sub Keys_Callback( _
           ByVal vRet As Variant)
    Log "Keys_Callback; Return = " & DumpItems(vRet)
End Sub

' // Create CPrivateClass instance and add it to list
Private Sub cmdAddObject_Click()
    Dim cNewObj     As Object
    Dim lAsynchId   As Long
    
    Set cNewObj = CreateActiveXObjectInNewThread2("Scripting.Dictionary", , lAsynchId)

    mcObjList.Add cNewObj
    mcIDs.Add lAsynchId
    
    lstObjects.AddItem "0x" & Hex$(lAsynchId)
    
End Sub

Private Sub cmdDicAdd_Click()
    Dim sKey        As String
    Dim sItem       As String
    Dim cObj        As Object
    Dim lAsynchId   As Long
    
    On Error GoTo error_handler
    
    Set cObj = GetSelectedObject(lAsynchId)
    If cObj Is Nothing Then Exit Sub
    
    sKey = InputBox("Key")
    If StrPtr(sKey) = 0 Then Exit Sub
    
    sItem = InputBox("Item")
    If StrPtr(sItem) = 0 Then Exit Sub
    
    If chkSynch.Value = vbChecked Then
    
        AsynchDispMethodCall lAsynchId, "Add", VbMethod, Me, "Add_Callback", sKey, sItem
        
    Else
    
        cObj.Add sKey, sItem
        Log "Item added to 0x" & Hex(lAsynchId) & "; Key = '" & sKey & "'; Item = '" & sItem & "'"
        
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmdDicItems_Click()
    Dim sKey        As String
    Dim cObj        As Object
    Dim lAsynchId   As Long
    Dim vRet        As Variant
    
    On Error GoTo error_handler
    
    Set cObj = GetSelectedObject(lAsynchId)
    If cObj Is Nothing Then Exit Sub

    If chkSynch.Value = vbChecked Then
    
        AsynchDispMethodCall lAsynchId, "Items", VbMethod, Me, "Items_Callback"
        
    Else
    
        vRet = cObj.Items
        Log "Items method has been called " & Hex(lAsynchId) & "; Items = " & DumpItems(vRet)
        
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description, vbCritical
    
End Sub

Private Function DumpItems( _
                 ByRef vItems As Variant) As String
    Dim vItem   As Variant
    
    DumpItems = "{"
    
    For Each vItem In vItems
        DumpItems = DumpItems & vItem & ", "
    Next
    
    If Len(DumpItems) > 2 Then
        DumpItems = Left$(DumpItems, Len(DumpItems) - 2)
    End If
    
    DumpItems = DumpItems & "}"
    
End Function

Private Sub cmdDicKeys_Click()
    Dim sKey        As String
    Dim cObj        As Object
    Dim lAsynchId   As Long
    Dim vRet        As Variant
    
    On Error GoTo error_handler
    
    Set cObj = GetSelectedObject(lAsynchId)
    If cObj Is Nothing Then Exit Sub

    If chkSynch.Value = vbChecked Then
    
        AsynchDispMethodCall lAsynchId, "Keys", VbMethod, Me, "Keys_Callback"
        
    Else
    
        vRet = cObj.Keys
        Log "Keys method has been called " & Hex(lAsynchId) & "; Items = " & DumpItems(vRet)
        
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmdDicRemove_Click()
    Dim sKey        As String
    Dim cObj        As Object
    Dim lAsynchId   As Long
    
    On Error GoTo error_handler
    
    Set cObj = GetSelectedObject(lAsynchId)
    If cObj Is Nothing Then Exit Sub
    
    sKey = InputBox("Key")
    If StrPtr(sKey) = 0 Then Exit Sub

    If chkSynch.Value = vbChecked Then
    
        AsynchDispMethodCall lAsynchId, "Remove", VbMethod, Me, "Remove_Callback", sKey
        
    Else
    
        cObj.Remove sKey
        Log "Item removed from 0x" & Hex(lAsynchId) & "; Key = '" & sKey & "'"
        
    End If
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description, vbCritical
    
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

Private Sub Form_Load()
    
    Me.Caption = Me.Caption & " [ThreadID: 0x" & Hex$(App.ThreadID) & "]"
    
    modMultiThreading.Initialize

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

