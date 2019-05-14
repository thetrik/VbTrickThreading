VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAsynchDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // Downloader class
' //

Option Explicit

Private m_nIsDownloading    As Boolean           ' // Determine if downloading process is active
Private m_hUrl              As Long              ' // Request handle
Private m_bData()           As Byte              ' // Buffer of data
Private m_lSize             As Long              ' // The size of resource
Private m_lDownloaded       As Long              ' // The number of downloaded bytes
Private m_cLogBox           As TextBox           ' // Textbox for log
Private m_sFileName         As String            ' // File name of resource
Private m_tBuffer           As INTERNET_BUFFERS  ' // The async buffer descriptor
Private m_cCallback         As ICallbackEvents   ' // Callback event object
Private m_hWndAsync         As Long              ' // Handle of async window
Private m_hFile             As Long              ' // File handle

' // Set the request handle
' // This property is setted asynchronously from InternetStatusCallback
Public Property Let RequestHandle( _
                    ByVal hValue As Long)
    m_hUrl = hValue
End Property

' //
' // Download the file
' //
Public Sub StartDownloading( _
           ByRef sUrl As String)
    Dim lRet    As Long
    
    If m_hWndAsync = 0 Then
        Err.Raise 5
    End If
    
    ' // Abort all the previous downloading
    Abort
    
    ' // Start downloading
    ' // It'll call InternetStatusCallback function
    lRet = InternetOpenUrl(g_hSession, StrPtr(sUrl), 0, 0, INTERNET_FLAG_RELOAD, ByVal m_hWndAsync)
    
    If lRet = 0 Then
    
        If Err.LastDllError <> ERROR_IO_PENDING Then
            Err.Raise 5
        End If
        
    End If
    
    ' // Process is active
    m_nIsDownloading = True
    
End Sub

' //
' // Abort
' //
Public Sub Abort()
    
    ' // Close handles
    If m_hUrl Then
    
        InternetCloseHandle m_hUrl
        
        m_hUrl = 0
        
        WaitForHandleTermination
        
    End If
    
    If m_hFile Then
        CloseHandle m_hFile
        m_hFile = 0
    End If
        
    ' // Process has stopped
    m_nIsDownloading = False
    m_sFileName = vbNullString
    m_lDownloaded = 0
    m_lSize = 0
    
End Sub

' //
' // Get progress
' //
Public Property Get Progress() As Single
    
    If m_lSize = 0 Then
        ' // Use -1 to resources with unknown size
        Progress = -1
    Else
        Progress = m_lDownloaded / m_lSize
    End If
    
End Property

' //
' // Is downloading
' //
Public Property Get IsDownloading() As Boolean
    IsDownloading = m_nIsDownloading
End Property

' //
' // Get bytes count
' //
Public Property Get BytesCount() As Long
    BytesCount = m_lDownloaded
End Property

Public Property Set Callback( _
                    ByVal cObj As ICallbackEvents)
    Set m_cCallback = cObj
End Property

Public Property Set LogWindow( _
                    ByVal cTextBox As TextBox)
    Set m_cLogBox = cTextBox
End Property

Public Sub PutLog( _
           ByVal sText As String)
    Dim tTime   As SYSTEMTIME
    
    If Not m_cLogBox Is Nothing Then
            
        GetLocalTime tTime
        
        m_cLogBox.SelStart = Len(m_cLogBox.Text)
        m_cLogBox.SelLength = 0
        m_cLogBox.SelText = tTime.wHour & ":" & tTime.wMinute & ":" & tTime.wSecond & "." & tTime.wMilliseconds & ": " & sText & vbNewLine

    End If
    
End Sub

' // This function is called when INTERNET_STATUS_REQUEST_COMPLETE status is received
Public Sub OnStatusComplete( _
           ByVal lStatus As Long, _
           ByVal lError As Long)
    Dim lRet    As Long
    
    ' // Check if error
    If lStatus Then
        
        ' // The first
        If Len(m_sFileName) = 0 Then
            
            ' // Get original file name
            m_sFileName = GetRemoteFileName(m_hUrl)

            ' // Create file
            m_hFile = CreateFile(StrPtr(App.Path & "\" & m_sFileName), GENERIC_WRITE, FILE_SHARE_READ, _
                                 ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
            
            If m_hFile = INVALID_HANDLE_VALUE Then
                
                m_hFile = 0
                m_cCallback.Error Me, Err.LastDllError
                Abort
                Exit Sub
                
            End If
            
            ' // Get size of resource
            m_lSize = GetResourceSize(m_hUrl)
            
            ' // Setup buffer
            m_tBuffer.dwStructSize = Len(m_tBuffer)
            
            ' // 640KB buffer
            ReDim m_bData(655360 - 1)
    
            m_tBuffer.lpvBuffer = VarPtr(m_bData(0))
            m_tBuffer.dwBufferLength = UBound(m_bData) + 1
            
        Else
        
            m_lDownloaded = m_lDownloaded + m_tBuffer.dwBufferLength
            
            ' // Write buffer to file
            If WriteFile(m_hFile, m_bData(0), m_tBuffer.dwBufferLength, 0, ByVal 0&) = 0 Then

                m_cCallback.Error Me, Err.LastDllError
                Abort
                Exit Sub

            End If
            
        End If
        
        Do
            
            ' // Read the chunk
            lRet = InternetReadFileEx(m_hUrl, m_tBuffer, IRF_ASYNC, ByVal m_hWndAsync)
            
            If lRet = 0 Then
                If Err.LastDllError = ERROR_IO_PENDING Then
                    ' // InternetStatusCallback will called when data is available
                    Exit Do
                Else
                    If Not m_cCallback Is Nothing Then
                    
                        m_cCallback.Error Me, Err.LastDllError
                        Abort
                        Exit Do
                        
                    End If
                End If
            End If
            
            m_lDownloaded = m_lDownloaded + m_tBuffer.dwBufferLength
            
            ' // Write to file
            If WriteFile(m_hFile, m_bData(0), m_tBuffer.dwBufferLength, 0, ByVal 0&) = 0 Then

                m_cCallback.Error Me, Err.LastDllError
                Abort
                Exit Do

            End If
            
            ' // Process is complete
            If m_tBuffer.dwBufferLength = 0 Then
            
                ' // Complete
                If Not m_cCallback Is Nothing Then
                    m_cCallback.Complete Me
                End If
                
                Abort
                
                Exit Do
                
            End If
    
        Loop While True
        
    Else
    
        If Not m_cCallback Is Nothing Then

            m_cCallback.Error Me, lError
            
        End If
        
        Abort
        
    End If
    
End Sub

' // Wait until handle is being closed by tracking INTERNET_STATUS_HANDLE_CLOSING callback
Private Sub WaitForHandleTermination()
    Dim lRet        As Long
    Dim hEvent      As Long
    Dim bIsInIDE    As Boolean
    
    If m_hWndAsync = 0 Then Exit Sub
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    hEvent = GetWindowLong(m_hWndAsync, 4)
    
    Do
        
        ' // Wait until thread set event when it accepts INTERNET_STATUS_HANDLE_CLOSING callback
        ' // There is small probability to crash app in IDE because the callback function is called
        ' // from the window procedure allocated dynamicly. if the thread doesn't have time to return
        ' // from INTERNET_STATUS_HANDLE_CLOSING and we uninitialize MT module the thread'll return
        ' // to  the freed memory
        ' // In compiled form we don't have such behavior
        
        lRet = WaitForSingleObject(hEvent, 200)
        
        If bIsInIDE Then
            DoEvents
        End If
        
    Loop While lRet = WAIT_TIMEOUT
    
End Sub

Private Sub Class_Initialize()
    Dim hEvent  As Long
    
    ' // Initialize the message window
    ' // This window is used to manage async call from threads
    ' // Create window
    m_hWndAsync = CreateWindowEx(0, StrPtr("AsyncCaller"), 0, 0, 0, 0, 0, 0, _
                                 HWND_MESSAGE, 0, App.hInstance, ByVal 0&)
    
    If m_hWndAsync Then
        
        ' // Save reference to current instance
        SetWindowLong m_hWndAsync, 0, ObjPtr(Me)
        
        ' // Save event for waiting
        hEvent = CreateEvent(ByVal 0&, 0, 0, 0)
        
        SetWindowLong m_hWndAsync, 4, hEvent
        
    Else
        MsgBox "Unable to create async window", vbCritical
    End If
    
End Sub

Private Sub Class_Terminate()
    
    Abort
    
    If m_hWndAsync Then
    
        ' // Close event
        CloseHandle GetWindowLong(m_hWndAsync, 4)
        DestroyWindow m_hWndAsync
        
    End If
    
End Sub

' // Get resource size
Private Function GetResourceSize( _
                 ByVal hUrl As Long) As Long
    Dim lRet    As Long
    Dim cSize   As Currency
    
    lRet = HttpQueryInfo(hUrl, HTTP_QUERY_CONTENT_LENGTH Or HTTP_QUERY_FLAG_NUMBER, cSize, Len(cSize), 0)
    
    If lRet = 0 Or cSize > 1000000@ Then Exit Function
    
    GetResourceSize = cSize * 10000

End Function

' // Get remote file name from URL
Private Function GetRemoteFileName( _
                 ByVal hUrl As Long) As String
    Dim lSize   As Long
    Dim lRet    As Long
    Dim sHeader As String
    
    ' // Try to get name from Content-Disposition
    lRet = HttpQueryInfo(hUrl, HTTP_QUERY_CONTENT_DISPOSITION, ByVal 0&, lSize, 0)
    
    If lRet = 0 Then
        If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
            
            sHeader = Space$(lSize \ 2 - 1)
            
            lRet = HttpQueryInfo(hUrl, HTTP_QUERY_CONTENT_DISPOSITION, ByVal StrPtr(sHeader), lSize, 0)
            
            If lRet Then
                
                GetRemoteFileName = GetRemoteFileNameFromContentDisposition(sHeader)
                
                If Len(GetRemoteFileName) Then Exit Function
                
            End If
            
        End If
    End If
    
    ' // Try to get from URL
    lRet = InternetQueryOption(hUrl, INTERNET_OPTION_URL, ByVal 0&, lSize)
    
    If lRet = 0 Then
        If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
            
            sHeader = Space$(lSize \ 2 - 1)
            
            lRet = InternetQueryOption(hUrl, INTERNET_OPTION_URL, ByVal StrPtr(sHeader), lSize)
            
            If lRet Then
                
                lRet = PathFindFileName(StrPtr(sHeader))
                
                GetRemoteFileName = Mid$(sHeader, (lRet - StrPtr(sHeader)) \ 2 + 1)
                Exit Function
                
            End If
            
        End If
    End If
    
End Function

' // Extract file name
' // This is the simplified procedure and doesn't accept all the results
Private Function GetRemoteFileNameFromContentDisposition( _
                 ByRef sHeader As String) As String
    Dim lIndex  As Long
    Dim sName   As String
    Dim hr      As Long
    Dim lSize   As Long
    
    lIndex = InStr(1, sHeader, "attachment; filename=")
    
    If lIndex = 0 Then Exit Function
    
    If Mid$(sHeader, 22, 1) = """" Then
        GetRemoteFileNameFromContentDisposition = Mid$(sHeader, 23, Len(sHeader) - 23)
    Else
        GetRemoteFileNameFromContentDisposition = Mid$(sHeader, 22)
    End If
    
    lSize = Len(GetRemoteFileNameFromContentDisposition)
    
    If UrlUnescape(StrPtr(GetRemoteFileNameFromContentDisposition), 0, lSize, URL_UNESCAPE_INPLACE) < 0 Then
        GetRemoteFileNameFromContentDisposition = vbNullString
        Exit Function
    End If
    
    lIndex = InStr(1, GetRemoteFileNameFromContentDisposition, vbNullChar)
    
    If lIndex > 0 Then
        GetRemoteFileNameFromContentDisposition = Left$(GetRemoteFileNameFromContentDisposition, lIndex - 1)
    End If

End Function


