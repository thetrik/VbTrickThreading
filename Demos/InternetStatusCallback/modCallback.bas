Attribute VB_Name = "modCallback"
' //
' // Callback module
' //

Option Explicit

Public Const INTERNET_OPEN_TYPE_PRECONFIG            As Long = 0
Public Const INTERNET_FLAG_RELOAD                    As Long = &H80000000
Public Const CREATE_ALWAYS                           As Long = 2
Public Const FILE_ATTRIBUTE_NORMAL                   As Long = &H80
Public Const INVALID_HANDLE_VALUE                    As Long = -1
Public Const GENERIC_WRITE                           As Long = &H40000000
Public Const INTERNET_FLAG_ASYNC                     As Long = &H10000000
Public Const HTTP_QUERY_CONTENT_DISPOSITION          As Long = 47
Public Const HWND_MESSAGE                            As Long = -3
Public Const WM_ONCOMPLETE                           As Long = &H400
Public Const WM_PUTLOG                               As Long = &H401
Public Const WM_PUTHANDLE                            As Long = &H402
Public Const LVM_FIRST                               As Long = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE            As Long = (LVM_FIRST + 54)
Public Const LVS_EX_FULLROWSELECT                    As Long = &H20
Public Const LVS_EX_GRIDLINES                        As Long = &H1
Public Const FILE_SHARE_READ                         As Long = &H1
Public Const HTTP_QUERY_CONTENT_LENGTH               As Long = 5
Public Const HTTP_QUERY_FLAG_NUMBER                  As Long = &H20000000
Public Const INTERNET_STATUS_RESOLVING_NAME          As Long = 10
Public Const INTERNET_STATUS_NAME_RESOLVED           As Long = 11
Public Const INTERNET_STATUS_CONNECTING_TO_SERVER    As Long = 20
Public Const INTERNET_STATUS_CONNECTED_TO_SERVER     As Long = 21
Public Const INTERNET_STATUS_SENDING_REQUEST         As Long = 30
Public Const INTERNET_STATUS_REQUEST_SENT            As Long = 31
Public Const INTERNET_STATUS_RECEIVING_RESPONSE      As Long = 40
Public Const INTERNET_STATUS_RESPONSE_RECEIVED       As Long = 41
Public Const INTERNET_STATUS_CTL_RESPONSE_RECEIVED   As Long = 42
Public Const INTERNET_STATUS_PREFETCH                As Long = 43
Public Const INTERNET_STATUS_CLOSING_CONNECTION      As Long = 50
Public Const INTERNET_STATUS_CONNECTION_CLOSED       As Long = 51
Public Const INTERNET_STATUS_HANDLE_CREATED          As Long = 60
Public Const INTERNET_STATUS_HANDLE_CLOSING          As Long = 70
Public Const INTERNET_STATUS_REQUEST_COMPLETE        As Long = 100
Public Const INTERNET_STATUS_REDIRECT                As Long = 110
Public Const INTERNET_STATUS_STATE_CHANGE            As Long = 200
Public Const INTERNET_OPTION_URL                     As Long = 34
Public Const ERROR_INSUFFICIENT_BUFFER               As Long = 122
Public Const URL_UNESCAPE_INPLACE                    As Long = &H100000
Public Const IRF_ASYNC                               As Long = &H1
Public Const ERROR_IO_PENDING                        As Long = 997
Public Const WAIT_TIMEOUT                            As Long = &H102&
Public Const SB_VERT                                 As Long = 1

Public Type INTERNET_ASYNC_RESULT
    dwResult                    As Long
    dwError                     As Long
End Type

Public Type WNDCLASSEX
    cbSize                      As Long
    style                       As Long
    lpfnwndproc                 As Long
    cbClsextra                  As Long
    cbWndExtra2                 As Long
    hInstance                   As Long
    hIcon                       As Long
    hCursor                     As Long
    hbrBackground               As Long
    lpszMenuName                As Long
    lpszClassName               As Long
    hIconSm                     As Long
End Type

Public Type INTERNET_BUFFERS
    dwStructSize                As Long
    pNext                       As Long
    lpcszHeader                 As Long
    dwHeadersLength             As Long
    dwHeadersTotal              As Long
    lpvBuffer                   As Long
    dwBufferLength              As Long
    dwBufferTotal               As Long
    dwOffsetLow                 As Long
    dwOffsetHigh                As Long
End Type

Public Type SYSTEMTIME
    wYear                       As Integer
    wMonth                      As Integer
    wDayOfWeek                  As Integer
    wDay                        As Integer
    wHour                       As Integer
    wMinute                     As Integer
    wSecond                     As Integer
    wMilliseconds               As Integer
End Type

Public Type tInternetCallbackParams
    hInternet                   As Long
    dwContext                   As Long    ' // hWnd of async window
    dwInternetStatus            As Long
    lpvStatusInformation        As Long
    dwStatusInformationLength   As Long
End Type

Public Declare Function GetScrollPos Lib "user32" ( _
                        ByVal hwnd As Long, _
                        ByVal nBar As Long) As Long
Public Declare Function SetScrollPos Lib "user32" ( _
                        ByVal hwnd As Long, _
                        ByVal nBar As Long, _
                        ByVal nPos As Long, _
                        ByVal bRedraw As Long) As Long
Public Declare Function CreateEvent Lib "kernel32" _
                        Alias "CreateEventW" ( _
                        ByRef lpEventAttributes As Any, _
                        ByVal bManualReset As Long, _
                        ByVal bInitialState As Long, _
                        ByVal lpName As Long) As Long
Public Declare Function SetEvent Lib "kernel32" ( _
                        ByVal hEvent As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" ( _
                        ByVal hHandle As Long, _
                        ByVal dwMilliseconds As Long) As Long
Public Declare Function UnregisterClass Lib "user32" _
                        Alias "UnregisterClassW" ( _
                        ByVal lpClassName As Long, _
                        ByVal hInstance As Long) As Long
Public Declare Sub GetLocalTime Lib "kernel32" ( _
                   ByRef lpSystemTime As SYSTEMTIME)
Public Declare Function SysAllocString Lib "oleaut32" ( _
                        ByRef pOlechar As Any) As Long
Public Declare Function GetMem4 Lib "msvbvm60" ( _
                        ByRef src As Any, _
                        ByRef Dst As Any) As Long
Public Declare Function lstrlenA Lib "kernel32" ( _
                        ByRef lpString As Any) As Long
Public Declare Function lstrcpynA Lib "kernel32" ( _
                        ByRef lpString1 As Any, _
                        ByRef lpString2 As Any, _
                        ByVal iMaxLength As Long) As Long
Public Declare Function HttpQueryInfo Lib "wininet" _
                        Alias "HttpQueryInfoW" ( _
                        ByVal hRequest As Long, _
                        ByVal dwInfoLevel As Long, _
                        ByRef lpBuffer As Any, _
                        ByRef lpdwBufferLength As Long, _
                        ByRef lpdwIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" _
                        Alias "SendMessageW" ( _
                        ByVal hwnd As Long, _
                        ByVal wMsg As Long, _
                        ByVal wParam As Long, _
                        ByRef lParam As Any) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60" _
                        Alias "__vbaObjSetAddref" ( _
                        ByRef dstObject As Any, _
                        ByRef srcObjPtr As Any) As Long
Public Declare Function InternetCloseHandle Lib "wininet" ( _
                        ByVal hInternet As Long) As Boolean
Public Declare Function InternetOpen Lib "wininet" _
                        Alias "InternetOpenW" ( _
                        ByVal lpszAgent As Long, _
                        ByVal dwAccessType As Long, _
                        ByVal lpszProxy As Long, _
                        ByVal lpszProxyBypass As Long, _
                        ByVal dwFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet" _
                        Alias "InternetOpenUrlW" ( _
                        ByVal hInternet As Long, _
                        ByVal lpszUrl As Long, _
                        ByVal lpszHeaders As Long, _
                        ByVal dwHeadersLength As Long, _
                        ByVal dwFlags As Long, _
                        ByRef dwContext As Any) As Long
Public Declare Function InternetReadFileEx Lib "wininet" _
                        Alias "InternetReadFileExW" ( _
                        ByVal hFile As Long, _
                        ByRef lpBuffersOut As Any, _
                        ByVal dwFlags As Long, _
                        ByRef dwContext As Any) As Long
Public Declare Function CreateFile Lib "kernel32" _
                        Alias "CreateFileW" ( _
                        ByVal lpFileName As Long, _
                        ByVal dwDesiredAccess As Long, _
                        ByVal dwShareMode As Long, _
                        ByRef lpSecurityAttributes As Any, _
                        ByVal dwCreationDisposition As Long, _
                        ByVal dwFlagsAndAttributes As Long, _
                        ByVal hTemplateFile As Long) As Long
Public Declare Function WriteFile Lib "kernel32" ( _
                        ByVal hFile As Long, _
                        ByRef lpBuffer As Any, _
                        ByVal nNumberOfBytesToWrite As Long, _
                        ByRef lpNumberOfBytesWritten As Long, _
                        ByRef lpOverlapped As Any) As Long
Public Declare Function InternetSetStatusCallback Lib "wininet.dll" ( _
                        ByVal hInternetSession As Long, _
                        ByVal lpfnInternetCallback As Long) As Long
Public Declare Function ArrPtr Lib "msvbvm60" _
                        Alias "VarPtr" ( _
                        ByRef arr() As Any) As Long
Public Declare Function InternetQueryOption Lib "wininet.dll" _
                        Alias "InternetQueryOptionW" ( _
                        ByVal hInternet As Long, _
                        ByVal dwOption As Long, _
                        ByRef lpBuffer As Any, _
                        ByRef lpdwBufferLength As Long) As Long
Public Declare Function PathFindFileName Lib "Shlwapi.dll" _
                        Alias "PathFindFileNameW" ( _
                        ByVal pszPath As Long) As Long
Public Declare Function UrlUnescape Lib "Shlwapi.dll" _
                        Alias "UrlUnescapeW" ( _
                        ByVal pszUrl As Long, _
                        ByVal pszUnescaped As Long, _
                        ByRef pcchUnescaped As Long, _
                        ByVal dwFlags As Long) As Long
Public Declare Function vbaObjSet Lib "msvbvm60" _
                        Alias "__vbaObjSet" ( _
                        ByRef dstObject As Any, _
                        ByRef srcObjPtr As Any) As Long
Public Declare Function GetClassInfoEx Lib "user32" _
                        Alias "GetClassInfoExW" ( _
                        ByVal hInstance As Long, _
                        ByVal lpClassName As Long, _
                        ByRef lpWndClassEx As WNDCLASSEX) As Long
Public Declare Function RegisterClassEx Lib "user32" _
                        Alias "RegisterClassExW" ( _
                        ByRef pcWndClassEx As WNDCLASSEX) As Integer
Public Declare Function CreateWindowEx Lib "user32" _
                        Alias "CreateWindowExW" ( _
                        ByVal dwExStyle As Long, _
                        ByVal lpClassName As Long, _
                        ByVal lpWindowName As Long, _
                        ByVal dwStyle As Long, _
                        ByVal x As Long, _
                        ByVal y As Long, _
                        ByVal nWidth As Long, _
                        ByVal nHeight As Long, _
                        ByVal hWndParent As Long, _
                        ByVal hMenu As Long, _
                        ByVal hInstance As Long, _
                        ByRef lpParam As Any) As Long
Public Declare Function GetModuleHandle Lib "kernel32" _
                        Alias "GetModuleHandleW" ( _
                        ByVal lpModuleName As Long) As Long
Public Declare Function DestroyWindow Lib "user32" ( _
                        ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" _
                        Alias "SetWindowLongW" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long
Public Declare Function DefWindowProc Lib "user32" _
                        Alias "DefWindowProcW" ( _
                        ByVal hwnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" _
                        Alias "GetWindowLongW" ( _
                        ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
Public Declare Function PostMessage Lib "user32" _
                        Alias "PostMessageW" ( _
                        ByVal hwnd As Long, _
                        ByVal wMsg As Long, _
                        ByVal wParam As Long, _
                        ByRef lParam As Any) As Long
Public Declare Function StrFormatByteSize Lib "Shlwapi" _
                        Alias "StrFormatByteSizeW" ( _
                        ByVal qdwl As Long, _
                        ByVal qdwh As Long, _
                        ByVal pszBuf As Long, _
                        ByVal cchBuf As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" _
                   Alias "RtlMoveMemory" ( _
                   ByRef Destination As Any, _
                   ByRef Source As Any, _
                   ByVal Length As Long)

Public g_hSession   As Long     ' // Session handle

' //
' // Register async window class
' //
Public Function RegisterAsyncWindowClass() As Boolean
    Dim tClass      As WNDCLASSEX
    

    tClass.cbSize = Len(tClass)
    
    ' // Check if class already registered
    If GetClassInfoEx(App.hInstance, StrPtr("AsyncCaller"), tClass) = 0 Then

        tClass.hInstance = App.hInstance
        tClass.lpfnwndproc = FAR_PROC(AddressOf AsyncWndProc)
        tClass.lpszClassName = StrPtr("AsyncCaller")
        tClass.cbWndExtra2 = 8  ' // 4 bytes for store reference to object, 4 bytes for event handle
        
        If RegisterClassEx(tClass) = 0 Then Exit Function

    End If
    
    RegisterAsyncWindowClass = True
    
End Function

' //
' // Unregister
' //
Public Sub UnregisterAsyncWindowClass()
    UnregisterClass StrPtr("AsyncCaller"), App.hInstance
End Sub

' //
' // The callback function
' // Is called from arbitrary threads
' //
Public Function InternetCallback( _
                ByRef tParams As tInternetCallbackParams) As Long
    Dim tAsyncRes   As INTERNET_ASYNC_RESULT
    Dim lSize       As Long
    Dim lFlags      As Long
    Dim hwnd        As Long
    
    hwnd = tParams.dwContext    ' // Extract async hWnd
    
    ' // Check status
    Select Case tParams.dwInternetStatus

    Case INTERNET_STATUS_REQUEST_COMPLETE
        
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_REQUEST_COMPLETE"
        
        ' // Extract result
        CopyMemory tAsyncRes, ByVal tParams.lpvStatusInformation, Len(tAsyncRes)

        PutLogAsync hwnd, "Status: " & CBool(tAsyncRes.dwResult)
        PutLogAsync hwnd, "Error code: " & CStr(tAsyncRes.dwError)
        
        ' // Make async call to main thread
        OnStatusCompleteAsynch hwnd, tAsyncRes.dwResult, tAsyncRes.dwError

    Case INTERNET_STATUS_CLOSING_CONNECTION
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_CLOSING_CONNECTION"
        
    Case INTERNET_STATUS_CONNECTED_TO_SERVER
        
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_CONNECTED_TO_SERVER"
        PutLogAsync hwnd, "IP: " & StringFromPtr(tParams.lpvStatusInformation)
        
    Case INTERNET_STATUS_CONNECTING_TO_SERVER
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_CONNECTING_TO_SERVER"
        PutLogAsync hwnd, "IP: " & StringFromPtr(tParams.lpvStatusInformation)

    Case INTERNET_STATUS_CONNECTION_CLOSED
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_CONNECTION_CLOSED"

    Case INTERNET_STATUS_HANDLE_CLOSING
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_HANDLE_CLOSING"
        
        ' // The handle is closed. Set event to release main thread
        SetEvent GetWindowLong(hwnd, 4)
        
    Case INTERNET_STATUS_HANDLE_CREATED
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_HANDLE_CREATED"
        
        ' // Extract result
        CopyMemory tAsyncRes, ByVal tParams.lpvStatusInformation, Len(tAsyncRes)
        
        ' // Set handle to main thread
        PutHandleAsync hwnd, tAsyncRes.dwResult

        PutLogAsync hwnd, "Handle: 0x" & Hex$(tAsyncRes.dwResult)
        PutLogAsync hwnd, "Error code: " & CStr(tAsyncRes.dwError)
        
    Case INTERNET_STATUS_NAME_RESOLVED
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_NAME_RESOLVED"
        PutLogAsync hwnd, "Host: " & StringFromPtr(tParams.lpvStatusInformation)

    Case INTERNET_STATUS_RECEIVING_RESPONSE
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_RECEIVING_RESPONSE"
    
    Case INTERNET_STATUS_REDIRECT
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_REDIRECT"
        PutLogAsync hwnd, "New URL: " & StringFromPtr(tParams.lpvStatusInformation)

    Case INTERNET_STATUS_REQUEST_SENT
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_REQUEST_SENT"
        
        ' // Extract size
        GetMem4 ByVal tParams.lpvStatusInformation, lSize
        
        PutLogAsync hwnd, "Size: " & CStr(lSize) & "bytes"
        
    Case INTERNET_STATUS_RESOLVING_NAME
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_RESOLVING_NAME"
        PutLogAsync hwnd, "Host: " & StringFromPtr(tParams.lpvStatusInformation)
        
    Case INTERNET_STATUS_RESPONSE_RECEIVED
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_RESPONSE_RECEIVED"
    
    Case INTERNET_STATUS_SENDING_REQUEST
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_SENDING_REQUEST"
    
    Case INTERNET_STATUS_STATE_CHANGE
    
        PutLogAsync hwnd, "TID: 0x" & Hex$(App.ThreadID) & ": " & "INTERNET_STATUS_STATE_CHANGE"
        
        ' // Extract flags
        GetMem4 ByVal tParams.lpvStatusInformation, lFlags
        
        PutLogAsync hwnd, "Flags: 0x" & Hex$(lFlags)
        
    End Select

End Function

' // This function is called in EXE
Public Function InternetCallbackEXE( _
                ByVal hInternet As Long, _
                ByVal dwContext As Long, _
                ByVal dwInternetStatus As Long, _
                ByVal lpvStatusInformation As Long, _
                ByVal dwStatusInformationLength As Long) As Long
    ' // Initialize project context and call function InternetCallback
    InitCurrentThreadAndCallFunction AddressOf InternetCallback, VarPtr(hInternet), InternetCallbackEXE
End Function

Public Function FAR_PROC( _
                ByVal pfn As Long) As Long
    FAR_PROC = pfn
End Function

' // Window proc of async window in main thread
Public Function AsyncWndProc( _
                ByVal hwnd As Long, _
                ByVal uMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
    Dim cObj    As CAsynchDownloader
    
    ' // Get object from window bytes
    vbaObjSetAddref cObj, ByVal GetWindowLong(hwnd, 0)
    
    Select Case uMsg
    Case WM_PUTLOG  ' // Show log entry
        Dim bstrText    As String
        
        ' // Extract string
        GetMem4 wParam, ByVal VarPtr(bstrText)
        
        cObj.PutLog bstrText
        
    Case WM_PUTHANDLE   ' // Put request handle
    
        cObj.RequestHandle = wParam
        
    Case WM_ONCOMPLETE  ' // INTERNET_STATUS_REQUEST_COMPLETE event

        cObj.OnStatusComplete wParam, lParam
        
    Case Else
        AsyncWndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    End Select
    
End Function

Public Function MakeTrue( _
                ByRef bValue As Boolean) As Boolean
    bValue = True
    MakeTrue = True
End Function

Private Sub PutHandleAsync( _
            ByVal hwnd As Long, _
            ByVal hUrl As Long)
    PostMessage hwnd, WM_PUTHANDLE, hUrl, ByVal 0&
End Sub

Private Sub PutLogAsync( _
            ByVal hwnd As Long, _
            ByRef sText As String)
    PostMessage hwnd, WM_PUTLOG, SysAllocString(ByVal StrPtr(sText)), ByVal 0&
End Sub

Private Sub OnStatusCompleteAsynch( _
            ByVal hwnd As Long, _
            ByVal lStatus As Long, _
            ByVal lError As Long)
    PostMessage hwnd, WM_ONCOMPLETE, lStatus, ByVal lError
End Sub

' // Get ANSI string from ptr
Private Function StringFromPtr( _
                 ByVal ptr As Long) As String
    Dim lSize   As Long
    
    lSize = lstrlenA(ByVal ptr)
    
    If lSize > 0 Then
        
        StringFromPtr = Space$(lSize)
        lstrcpynA ByVal StringFromPtr, ByVal ptr, lSize + 1
        
    End If

End Function

