VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InternetStatusCallback example by The trick"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   4380
      TabIndex        =   6
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   500
      Left            =   4980
      Top             =   2940
   End
   Begin ComctlLib.ListView lvwList 
      Height          =   1875
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3307
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "iglIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Progress"
         Object.Width           =   1428
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Percents"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Bytes"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "URL"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H8000000F&
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3180
      Width           =   5655
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4380
      TabIndex        =   2
      Top             =   360
      Width           =   1395
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin ComctlLib.ImageList iglIcons 
      Left            =   3720
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   327682
   End
   Begin VB.Label lblLog 
      Caption         =   "Callback log:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblURL 
      Caption         =   "Enter URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5595
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Async downloader using InternetStatusCallback
' // by The trick 2019
' //

Option Explicit

Implements ICallbackEvents

Private m_pfnPrevCallback       As Long             ' // Previous callback function
Private m_cRequestsQueue        As CAsynchQueue     ' // Queue

Private Sub cmdAbort_Click()
    Dim cObj    As CAsynchDownloader
    
    If lvwList.SelectedItem Is Nothing Then Exit Sub
        
    ' // Get downloader from tag
    vbaObjSetAddref cObj, ByVal CLng(lvwList.SelectedItem.Tag)

    cObj.Abort
    
End Sub

Private Sub cmdDownload_Click()
    Dim cObj    As CAsynchDownloader
    
    Set cObj = m_cRequestsQueue.Add(txtURL.Text, Me, txtLog)
    
    With lvwList.ListItems.Add(, "#" & ObjPtr(cObj), , , 1)
        
        .SubItems(3) = txtURL.Text
        .Tag = ObjPtr(cObj)
        
    End With
    
End Sub

Private Sub Form_Initialize()
    modMultiThreading.Initialize
    RegisterAsyncWindowClass
End Sub

Private Sub Form_Load()
    Dim bIsInIDE    As Boolean
    
    ' // Enable grid and full row select
    SendMessage lvwList.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES, _
                                ByVal LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES
    
    Set m_cRequestsQueue = New CAsynchQueue

    Debug.Assert MakeTrue(bIsInIDE)
    
    Me.Caption = Me.Caption & " TID: 0x" & Hex$(App.ThreadID)
    
    ' // Create icons
    CreateImageListProgressIconsCollection
    
    ' // Open session
    g_hSession = InternetOpen(StrPtr(App.ProductName), INTERNET_OPEN_TYPE_PRECONFIG, 0, 0, INTERNET_FLAG_ASYNC)
    
    If g_hSession = 0 Then
        MsgBox "InternetOpen failed " & CStr(Err.LastDllError)
    Else
        cmdDownload.Enabled = True
    End If
    
    ' // Setup callback
    If bIsInIDE Then
        m_pfnPrevCallback = InternetSetStatusCallback(g_hSession, InitCurrentThreadAndCallFunctionIDEProc( _
                                                                    AddressOf InternetCallback, 20))
    Else
        m_pfnPrevCallback = InternetSetStatusCallback(g_hSession, AddressOf InternetCallbackEXE)
    End If
    
End Sub

Private Sub Form_Terminate()
    modMultiThreading.Uninitialize
    UnregisterAsyncWindowClass
End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)
    Dim cObj    As CAsynchDownloader
    
    ' // Abort all
    For Each cObj In m_cRequestsQueue
        cObj.Abort
    Next
    
    Set m_cRequestsQueue = Nothing
    
    ' // Clean up
    If g_hSession Then
        InternetSetStatusCallback g_hSession, m_pfnPrevCallback
        InternetCloseHandle g_hSession
    End If
    
    g_hSession = 0
    
End Sub


' // Create icons
Private Sub CreateImageListProgressIconsCollection()
    Dim cPicBox As PictureBox
    Dim lState  As Long
    Dim lX      As Long
    
    Set cPicBox = Me.Controls.Add("VB.PictureBox", "picTemp")
    
    cPicBox.BackColor = vbMagenta
    cPicBox.AutoRedraw = True
    cPicBox.ScaleMode = vbPixels
    cPicBox.BorderStyle = 0
    
    cPicBox.Move 0, 0, Me.ScaleX(64, vbPixels, Me.ScaleMode), Me.ScaleY(16, vbPixels, Me.ScaleMode)
    
    ' // Draw progress bar (21 state)
    cPicBox.Line (0, 0)-(63, 15), &H209F20, B
    
    For lState = 0 To 100 Step 5
        
        cPicBox.Line (1, 1)-Step(lState / 100 * 61, 6), &H80FF80, BF
        cPicBox.Line (1, 7)-Step(lState / 100 * 61, 7), vbGreen, BF
        
        iglIcons.ListImages.Add , , cPicBox.Image

    Next
    
    ' // Draw progress bar for unknown progress
    For lState = 0 To 1
    
        cPicBox.Cls
        
        For lX = -16 To 64 Step 3
            cPicBox.Line (lX + lState * 2, 0)-Step(16, 16), vbGreen
        Next
        
        cPicBox.Line (0, 0)-(63, 15), &H209F20, B
        
        iglIcons.ListImages.Add , , cPicBox.Image
        
    Next
    
    ' // Draw check mark
    cPicBox.Cls
    
    cPicBox.DrawWidth = 3
    cPicBox.Line (20, 10)-(25, 13), vbGreen
    cPicBox.Line -(32, 3), vbGreen
    
    iglIcons.ListImages.Add , , cPicBox.Image
    
    ' // Draw cross (X)
    cPicBox.Cls
    cPicBox.Line (20, 3)-(30, 13), vbRed
    cPicBox.Line (30, 3)-(20, 13), vbRed
    
    iglIcons.ListImages.Add , , cPicBox.Image
    
    Me.Controls.Remove "picTemp"
    
End Sub

Private Sub ICallbackEvents_Complete( _
            ByVal cObj As CAsynchDownloader)
    Dim cItem   As ListItem
    
    Set cItem = lvwList.ListItems("#" & ObjPtr(cObj))
    
    ' // Update icon (check mark)
    cItem.SmallIcon = 24
    cItem.SubItems(1) = "100%"
    
End Sub

Private Sub ICallbackEvents_Error( _
            ByVal cObj As CAsynchDownloader, _
            ByVal lError As Long)
    Dim cItem   As ListItem
    
    Set cItem = lvwList.ListItems("#" & ObjPtr(cObj))
    
    ' // Update icon (X)
    cItem.SmallIcon = 25

End Sub

Private Sub tmrUpdate_Timer()
    Dim cObj    As CAsynchDownloader
    Dim cItem   As ListItem
    Dim sSize   As String

    If lvwList.ListItems.Count = 0 Then Exit Sub

    ' // Update all the items
    For Each cObj In m_cRequestsQueue
        
        If cObj.IsDownloading Then
            
            Set cItem = lvwList.ListItems("#" & ObjPtr(cObj))
            
            If cObj.Progress >= 0 Then
                cItem.SmallIcon = Int(cObj.Progress * 20) + 1
                cItem.SubItems(1) = Format$(cObj.Progress, "0.00%")
            Else
            
                If cItem.SmallIcon = 0 Then
                    cItem.SmallIcon = 22
                ElseIf cItem.SmallIcon = 22 Then
                    cItem.SmallIcon = 23
                Else
                    cItem.SmallIcon = 22
                End If
                
                cItem.SubItems(1) = "?"
                
            End If
            
            sSize = Space$(10)
            
            StrFormatByteSize cObj.BytesCount, 0, StrPtr(sSize), Len(sSize)
            
            cItem.SubItems(2) = sSize
            
        End If
        
    Next

End Sub
