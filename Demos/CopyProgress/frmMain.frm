VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy folder threading example by The trick"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSource 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   6555
   End
   Begin VB.CommandButton cmdSourceBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   360
      Width           =   1275
   End
   Begin VB.TextBox txtDestination 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   6555
   End
   Begin VB.CommandButton cmdDestinationBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   1020
      Width           =   1275
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   435
      Left            =   2700
      TabIndex        =   5
      Top             =   3900
      Width           =   1395
   End
   Begin VB.Frame fraProgress 
      Caption         =   "Progress"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   7875
      Begin VB.PictureBox picProgress 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   7635
         TabIndex        =   2
         Top             =   300
         Width           =   7635
      End
      Begin VB.Label lblInfo 
         Caption         =   "None"
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   180
         TabIndex        =   4
         Top             =   1140
         Width           =   7575
      End
      Begin VB.Label lblCurrentFile 
         ForeColor       =   &H0000C000&
         Height          =   435
         Left            =   180
         TabIndex        =   3
         Top             =   660
         Width           =   7455
      End
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   435
      Left            =   4140
      TabIndex        =   0
      Top             =   3900
      Width           =   1395
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3540
      Top             =   2640
   End
   Begin VB.Label Label1 
      Caption         =   "Source folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label2 
      Caption         =   "Destination folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   780
      Width           =   7935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // CopyFolder example
' // By The trick 2018
' // This example copies folder in separate thread using class instance
' //

Option Explicit

Private Declare Function StrFormatByteSizeW Lib "Shlwapi" ( _
                         ByVal qdw As Currency, _
                         ByVal pszBuf As Long, _
                         ByVal cchBuf As Long) As Long
    
Private Const BIF_RETURNONLYFSDIRS    As Long = &H1&
Private Const BIF_NEWDIALOGSTYLE      As Long = &H40&
Private Const ERROR_FILE_EXISTS       As Long = 80

Private mcCopyClass As Object
Private mlID        As Long
Private mbCancel    As Boolean
Private mbIsRunning As Boolean
Private mbIsPaused  As Boolean

' // This method is called by thread when Copy method is finished
Public Sub CopyComplete( _
           ByRef vRet As Variant)
    
    mbIsRunning = False
    cmdPause.Enabled = False
    cmdCopy.Caption = "Copy"
    picProgress.Cls
    
End Sub

' // This method is called when copying process begins
Public Sub Start()
    Dim bIsinIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    mbIsRunning = True
    If Not bIsinIDE Then cmdPause.Enabled = True
    cmdCopy.Caption = "Cancel"
    mbCancel = False
    
End Sub

' // This method is called by class when copying process is finished
Public Sub Complete( _
           ByVal lError As Long)

    lblInfo.Caption = "Complete [0x" & Hex$(lError) & "] " & Error(lError)
    lblCurrentFile.Caption = vbNullString
    Beep
    
End Sub

' // This method is called periodically by class to update information about files
Public Sub EnumrateProgress( _
           ByVal cNumberOfFiles As Currency, _
           ByVal cNumberOfFolders As Currency, _
           ByVal cDataCountInBytes As Currency, _
           ByRef sCurrentPath As String, _
           ByRef bCancelFlag As Boolean)
    
    lblCurrentFile.Caption = sCurrentPath
    lblInfo.Caption = "Total files: " & cNumberOfFiles * 10000 & vbNewLine & _
                      "Total folders: " & cNumberOfFolders * 10000 & vbNewLine & _
                      "Total size: " & GetSize(cDataCountInBytes)
    
    bCancelFlag = mbCancel
    
End Sub

' // This method is called when all files have been enumerated
Public Sub EnumrateComplete( _
           ByVal cNumberOfFiles As Currency, _
           ByVal cNumberOfFolders As Currency, _
           ByVal cDataCountInBytes As Currency, _
           ByRef bCancelFlag As Boolean)

    lblInfo.Caption = "Total files: " & cNumberOfFiles * 10000 & vbNewLine & _
                      "Total folders: " & cNumberOfFolders * 10000 & vbNewLine & _
                      "Total size: " & GetSize(cDataCountInBytes)
        
End Sub

' // This function is called periodically to update copying progress
Public Sub CopyProgress( _
           ByRef sCurrentFile As String, _
           ByVal cTotalFileSize As Currency, _
           ByVal cTotalBytesTransferred As Currency, _
           ByVal cDataCountInBytes As Currency, _
           ByVal cTransferedDataInBytes As Currency, _
           ByVal cNumberOfFiles As Currency, _
           ByVal cNumberOfFolders As Currency, _
           ByRef bCancelFlag As Boolean)
    Dim dAllProgress    As Double
    Dim dCurProgress    As Double
    
    If cDataCountInBytes > 0 Then
        dAllProgress = cTransferedDataInBytes / cDataCountInBytes
    Else
        dAllProgress = 1
    End If
    
    If cTotalFileSize > 0 Then
        dCurProgress = cTotalBytesTransferred / cTotalFileSize
    Else
        dCurProgress = 1
    End If
    
    picProgress.Cls
    
    If dAllProgress > dCurProgress Then
        picProgress.Line (0, 0)-(dCurProgress * picProgress.ScaleWidth, picProgress.ScaleHeight), RGB(0, 90, 0), BF
        picProgress.Line (dCurProgress * picProgress.ScaleWidth, 0)- _
                         (dAllProgress * picProgress.ScaleWidth, picProgress.ScaleHeight), RGB(0, 180, 0), BF
    Else
        picProgress.Line (0, 0)-(dAllProgress * picProgress.ScaleWidth, picProgress.ScaleHeight), RGB(0, 90, 0), BF
        picProgress.Line (dAllProgress * picProgress.ScaleWidth, 0)- _
                         (dCurProgress * picProgress.ScaleWidth, picProgress.ScaleHeight), RGB(0, 180, 0), BF
    End If
    
    lblInfo.Caption = "Remaining files: " & cNumberOfFiles * 10000 & vbNewLine & _
                      "Remaining folders: " & cNumberOfFolders * 10000 & vbNewLine & _
                      "Remaining size: " & GetSize(cDataCountInBytes - cTransferedDataInBytes)
    lblCurrentFile.Caption = sCurrentFile
    
    bCancelFlag = mbCancel
     
End Sub

' // This method is called when an error occured
Public Function FileCopyError( _
                ByRef sFileName As String, _
                ByVal lError As Long, _
                ByRef eFlags As eCopyFlags) As VbMsgBoxResult
    
    lblCurrentFile.Caption = sFileName
    
    If lError = ERROR_FILE_EXISTS Then
        
        FileCopyError = MsgBox("File already exists '" & sFileName & "'" & vbNewLine & "Overwrite?", vbQuestion Or vbYesNo)
        
        If FileCopyError = vbYes Then

            If MsgBox("Overwrite always?", vbQuestion Or vbYesNo) = vbYes Then
                eFlags = CF_OVERWRITEALWAYS
            Else
                eFlags = CF_OVERWRITE
            End If
            
            FileCopyError = vbRetry
            
        Else
            FileCopyError = vbIgnore
        End If

    Else
        FileCopyError = MsgBox("Unable to copy file" & vbNewLine & sFileName & vbNewLine & "Error: 0x" & _
                                Hex$(lError), vbExclamation Or vbAbortRetryIgnore)
    End If
    
End Function

' // Get string representation of file size (MB, GB, etc.)
Private Function GetSize( _
                 ByVal cValue As Currency) As String
    
    GetSize = Space$(32)
    
    If StrFormatByteSizeW(cValue, StrPtr(GetSize), Len(GetSize)) Then
        GetSize = Left$(GetSize, InStr(1, GetSize, vbNullChar) - 1)
    Else
        GetSize = "UNKNOWN"
    End If
    
End Function

Private Sub cmdCopy_Click()
        
    If mbIsRunning Then
        mbCancel = True
    Else
        AsynchDispMethodCall mlID, "Copy", VbMethod, Me, "CopyComplete", txtSource.Text, txtDestination.Text
    End If
    
End Sub

Private Sub cmdPause_Click()
    
    If mbIsPaused Then
    
        cmdPause.Caption = "Pause"
        SuspendResume mlID, False
        
    Else
    
        cmdPause.Caption = "Resume"
        SuspendResume mlID, True
        
    End If
    
    mbIsPaused = Not mbIsPaused
    
End Sub

Private Sub cmdSourceBrowse_Click()
    Dim sPath   As String
    
    sPath = BrowseForFolder()
    
    If Len(sPath) Then
        txtSource.Text = sPath
    End If
    
End Sub

Private Sub cmdDestinationBrowse_Click()
    Dim sPath   As String
    
    sPath = BrowseForFolder()
    
    If Len(sPath) Then
        txtDestination.Text = sPath
    End If
    
End Sub

Private Function BrowseForFolder() As String
    Dim Folder As Object
    
    With CreateObject("Shell.Application")
        Set Folder = .BrowseForFolder(Me.hwnd, "Pick a folder", BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE)
    End With
    
    If Not Folder Is Nothing Then
        BrowseForFolder = Folder.Self.Path
    End If
    
End Function

Private Sub Form_Load()
    
    modMultiThreading.Initialize
    modMultiThreading.EnablePrivateMarshaling True
    
    Set mcCopyClass = CreatePrivateObjectByNameInNewThread("CCopyProgress", , mlID)
    Set mcCopyClass.NotifyObject = Me
    
    ' // Update 10 times per second
    mcCopyClass.UpdateTime = 0.1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    mbCancel = True
    
    If mbIsPaused Then
        SuspendResume mlID, False
    End If
    
    Set mcCopyClass = Nothing
    
    WaitForObjectThreadCompletion mlID
    
    modMultiThreading.Uninitialize
    
End Sub

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True
    bValue = True
End Function


