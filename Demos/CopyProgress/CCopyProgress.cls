VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCopyProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CCopyProgress - class for copying with notification
' //

Option Explicit

Public Enum eCopyFlags
    CF_OVERWRITE = 1
    CF_OVERWRITEALWAYS
End Enum

Private Type tEnumerateData
    cNumberOfFiles          As Currency     ' // Number of files
    cNumberOfFolders        As Currency     ' // Number of folders
    cDataCountInBytes       As Currency     ' // Total bytes of all files
    cTransferedDataInBytes  As Currency     ' // Total transfered bytes
End Type

Private Const COPY_FILE_FAIL_IF_EXISTS  As Long = 1

Private Declare Function CopyFileEx Lib "kernel32" _
                         Alias "CopyFileExW" ( _
                         ByVal lpExistingFileName As Long, _
                         ByVal lpNewFileName As Long, _
                         ByVal lpProgressRoutine As Long, _
                         ByRef lpData As Any, _
                         ByRef pbCancel As Any, _
                         ByVal dwCopyFlags As Long) As Long
                    
Private mcNotifyObject          As Object               ' // Callback object (marshaled)
Private mfLastNotificationTime  As Single               ' // Last notification time
Private mtEnumerateData         As tEnumerateData       ' // Info about files
Private msCurrentFile           As String               ' // Current file
Private mbCancelFlag            As Boolean              ' // If true - abort copying
Private mcFso                   As FileSystemObject
Private meCopyFlags             As eCopyFlags           ' // Overwrite/yes/no
Private mfUpdateTime            As Single               ' // Updation interval

Public Property Set NotifyObject( _
                    ByVal cValue As Object)
    Set mcNotifyObject = cValue
End Property

Public Property Get UpdateTime() As Single
    UpdateTime = mfUpdateTime
End Property

Public Property Let UpdateTime( _
                    ByVal fValue As Single)
    mfUpdateTime = fValue
End Property

Public Sub Copy( _
           ByRef sSrcFolder As String, _
           ByRef sDstFolder As String)
    Dim cSrcFolder  As Folder
    Dim cDstFolder  As Folder
    
    meCopyFlags = 0
    
    On Error GoTo exit_proc
    
    mbCancelFlag = False
    mtEnumerateData.cDataCountInBytes = 0
    mtEnumerateData.cNumberOfFiles = 0
    mtEnumerateData.cNumberOfFolders = 0
    mtEnumerateData.cTransferedDataInBytes = 0
    
    Err.Clear
    
    mfLastNotificationTime = Timer
    
    Set mcFso = New FileSystemObject
    
    Set cSrcFolder = mcFso.GetFolder(sSrcFolder)
    Set cDstFolder = mcFso.GetFolder(sDstFolder)
    
    If Not mcNotifyObject Is Nothing Then
        mcNotifyObject.Start
    End If
    
    EnumerateFiles cSrcFolder   ' // Get info about all files
    
    If mbCancelFlag Then GoTo exit_proc
    
    If Not mcNotifyObject Is Nothing Then
        ' // Notify form
        mcNotifyObject.EnumrateComplete mtEnumerateData.cNumberOfFiles, mtEnumerateData.cNumberOfFolders, _
                                mtEnumerateData.cDataCountInBytes, mbCancelFlag
                                
    End If
    
    If mbCancelFlag Then GoTo exit_proc
    
    ' // Start copying
    CopyFolder cDstFolder, cSrcFolder
    
exit_proc:

    If Not mcNotifyObject Is Nothing Then
        mcNotifyObject.Complete Err.Number
    End If
    
End Sub

Public Sub ProgressRoutine( _
           ByVal cTotalFileSize As Currency, _
           ByVal cTotalBytesTransferred As Currency)
    Dim bIsinIDE    As Boolean

    Debug.Assert MakeTrue(bIsinIDE)
    
    If bIsinIDE Then DoEvents
    
    If Timer - mfLastNotificationTime > mfUpdateTime Then
    
        If Not mcNotifyObject Is Nothing Then
            mcNotifyObject.CopyProgress msCurrentFile, cTotalFileSize, cTotalBytesTransferred, _
                                        mtEnumerateData.cDataCountInBytes, mtEnumerateData.cTransferedDataInBytes + _
                                        cTotalBytesTransferred, mtEnumerateData.cNumberOfFiles, mtEnumerateData.cNumberOfFolders, _
                                        mbCancelFlag
        End If
        
        mfLastNotificationTime = Timer
    
    End If
    
End Sub

Private Sub CopyFolder( _
            ByVal cDst As Folder, _
            ByVal cSrc As Folder)
    Dim cFile       As File
    Dim cFolder     As Folder
    Dim cSubFolder  As Folder
    Dim lFlags      As Long
    
    For Each cFile In cSrc.Files
        
        msCurrentFile = cDst.Path & "\" & cFile.Name

        Do
            
            If meCopyFlags = CF_OVERWRITEALWAYS Then
                lFlags = 0
            Else
                lFlags = COPY_FILE_FAIL_IF_EXISTS
            End If
            
            If CopyFileEx(StrPtr(cFile.Path), StrPtr(msCurrentFile), AddressOf CopyProgressRoutine, _
                            Me, mbCancelFlag, lFlags) = 0 Then
                
                If mbCancelFlag Then Exit Sub
                
                If Not mcNotifyObject Is Nothing Then
                    
                    Select Case mcNotifyObject.FileCopyError(cFile.Path, Err.LastDllError, meCopyFlags)
                    Case VbMsgBoxResult.vbIgnore
                        Exit Do
                    Case VbMsgBoxResult.vbRetry
                    Case Else
                    
                        mbCancelFlag = True
                        Exit Sub
                        
                    End Select
                    
                Else
                    Exit Do
                End If
            
            Else
                Exit Do
            End If
            
            If meCopyFlags = CF_OVERWRITE Then
                lFlags = 0
            End If
            
        Loop

        mtEnumerateData.cTransferedDataInBytes = mtEnumerateData.cTransferedDataInBytes + CCur(cFile.Size / 10000)
        mtEnumerateData.cNumberOfFiles = mtEnumerateData.cNumberOfFiles - 0.0001@
            
    Next
    
    msCurrentFile = vbNullString
    
    For Each cFolder In cSrc.SubFolders
        
        If mbCancelFlag Then Exit Sub
        
        If Not mcFso.FolderExists(cDst.Path & "\" & cFolder.Name) Then
            Set cSubFolder = cDst.SubFolders.Add(cFolder.Name)
        Else
            Set cSubFolder = cDst.SubFolders(cFolder.Name)
        End If
        
        CopyFolder cSubFolder, cFolder

        mtEnumerateData.cNumberOfFolders = mtEnumerateData.cNumberOfFolders - 0.0001@
        
    Next

End Sub

Private Sub EnumerateFiles( _
            ByRef cFolder As Folder)
    Dim cFile       As File
    Dim cSubFolder  As Folder
    Dim bIsinIDE    As Boolean

    Debug.Assert MakeTrue(bIsinIDE)
    
    For Each cFile In cFolder.Files
    
        mtEnumerateData.cNumberOfFiles = mtEnumerateData.cNumberOfFiles + 0.0001@   ' // 1 - 64 bit
        mtEnumerateData.cDataCountInBytes = mtEnumerateData.cDataCountInBytes + CCur(cFile.Size / 10000)
        
        If Timer - mfLastNotificationTime > mfUpdateTime Then
            
            If Not mcNotifyObject Is Nothing Then
            
                mcNotifyObject.EnumrateProgress mtEnumerateData.cNumberOfFiles, mtEnumerateData.cNumberOfFolders, _
                                mtEnumerateData.cDataCountInBytes, cFolder.Path, mbCancelFlag
                                
                If mbCancelFlag Then Exit Sub
                                
            End If
            
            mfLastNotificationTime = Timer
            
        End If
        
        If bIsinIDE Then DoEvents
        
    Next
    
    For Each cSubFolder In cFolder.SubFolders
        
        mtEnumerateData.cNumberOfFolders = mtEnumerateData.cNumberOfFolders + 0.0001@
        EnumerateFiles cSubFolder
        
        If mbCancelFlag Then Exit Sub
        
    Next
    
End Sub

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True
    bValue = True
End Function

Private Sub Class_Initialize()
    mfUpdateTime = 0.25
End Sub
