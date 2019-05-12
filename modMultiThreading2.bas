Attribute VB_Name = "modMultiThreading"
' //
' // modMultiThreading2.bas - The module provides support for multi-threading.
' // Version 2
' // © Krivous Anatoly Anatolevich (The trick), 2015-2018
' // No TLB, No additional DLLs, No Asm-thunks (in compiled form)
' // Private object creation based on NameBasedObjectFactory by firehacker
' //

Option Explicit

Private Const HEAP_NO_SERIALIZE           As Long = &H1
Private Const HEAP_CREATE_ENABLE_EXECUTE  As Long = &H40000
Private Const PAGE_READWRITE              As Long = 4&
Private Const CC_STDCALL                  As Long = 4
Private Const HEAP_ZERO_MEMORY            As Long = &H8
Private Const HKEY_CURRENT_USER           As Long = &H80000001
Private Const REG_OPTION_NON_VOLATILE     As Long = 0
Private Const KEY_WRITE                   As Long = &H20006
Private Const KEY_QUERY_VALUE             As Long = &H1
Private Const REG_SZ                      As Long = 1
Private Const CLSCTX_INPROC_SERVER        As Long = 1
Private Const CLSCTX_LOCAL_SERVER         As Long = 4
Private Const MSHLFLAGS_TABLESTRONG       As Long = 1
Private Const MSHCTX_INPROC               As Long = 3
Private Const WM_ASYNCH_CALL              As Long = &H8001&
Private Const TLS_OUT_OF_INDEXES          As Long = &HFFFF&
Private Const ERROR_NO_MORE_ITEMS         As Long = 259&
Private Const PROCESS_HEAP_ENTRY_BUSY     As Long = &H4
Private Const TH32CS_SNAPTHREAD           As Long = 4
Private Const MESSAGE_WINDOW_CLASS        As String = "MT2_VB6"
Private Const WM_ONCALLBACK               As Long = &H400
Private Const HWND_MESSAGE                As Long = -3
Private Const THREAD_SUSPEND_RESUME       As Long = 2
Private Const SYNCHRONIZE                 As Long = &H100000
Private Const PM_REMOVE                   As Long = &H1

' // Lazy GUID structure
Private Type tCurGUID
    c1          As Currency
    c2          As Currency
End Type

' // That structure is used for asynch call. Caller pass that structure to callee
' // Callee uses that structure to make asynch call and callback to caller
Private Type tAsynchCallData
    pStream         As Long         ' // Stream for marshaling of callback object
    sCallBackName   As String       ' // Callback method name
    eCallType       As VbCallType   ' // Call type (Method, Property, etc.)
    sMethodName     As String       ' // Method name to call
    vArgs()         As Variant      ' // Arguments
End Type

' // TLS data for new thread
Private Type tThreadData
    lpParameter As Long
    lpAddress   As Long
End Type

' // Object types flags
Private Enum eObjThreadDataFlags
    OTDF_ACTIVEX    ' // Create new ActiveX object (Public)
    OTDF_PRIVATE    ' // Create local object (Private)
End Enum

' // New object info. That structure uses to create new object in new thread
Private Type tNewObjectThreadData
    eFlags  As eObjThreadDataFlags  ' // Object type
    pStream As Long                 ' // Marshaling stream
    hEvent  As Long                 ' // Synchronization event
    pClsid  As Long                 ' // Class identifier (for ActiveX calsses - CLSID, for private calsses - name)
    pIID    As Long                 ' // Interface identifier
    hr      As Long                 ' // Result
End Type

Private Type CRITICAL_SECTION
    pDebugInfo      As Long
    LockCount       As Long
    RecursionCount  As Long
    OwningThread    As Long
    LockSemaphore   As Long
    SpinCount       As Long
End Type

Private Type tCriticalSection
    tWinApiSection  As CRITICAL_SECTION
    bIsInitialized  As Boolean
End Type

Private Type PROCESS_HEAP_ENTRY
    lpData              As Long
    cbData              As Long
    cbOverhead          As Byte
    iRegionIndex        As Byte
    wFlags              As Integer
    dwCommittedSize     As Long
    dwUnCommittedSize   As Long
    lpFirstBlock        As Long
    lpLastBlock         As Long
End Type
Private Type THREADENTRY32
    dwSize              As Long
    cntUsage            As Long
    th32ThreadID        As Long
    th32OwnerProcessID  As Long
    tpBasePri           As Long
    tpDeltaPri          As Long
    dwFlags             As Long
End Type

Private Type POINTAPI
    x                   As Long
    y                   As Long
End Type

Private Type msg
    hwnd                As Long
    message             As Long
    wParam              As Long
    lParam              As Long
    time                As Long
    pt                  As POINTAPI
End Type

Private Type WNDCLASSEX
    cbSize          As Long
    style           As Long
    lpfnwndproc     As Long
    cbClsextra      As Long
    cbWndExtra2     As Long
    hInstance       As Long
    hIcon           As Long
    hCursor         As Long
    hbrBackground   As Long
    lpszMenuName    As Long
    lpszClassName   As Long
    hIconSm         As Long
End Type

Private Declare Function MsgWaitForMultipleObjects Lib "user32" ( _
                         ByVal nCount As Long, _
                         ByRef pHandles As Long, _
                         ByVal fWaitAll As Long, _
                         ByVal dwMilliseconds As Long, _
                         ByVal dwWakeMask As Long) As Long
Private Declare Function Sleep Lib "kernel32" ( _
                         ByVal dwMilliseconds As Long) As Long
Private Declare Function vbaNew Lib "msvbvm60" _
                         Alias "__vbaNew" ( _
                         ByRef lpObjectInformation As Any) As IUnknown
Private Declare Function lstrcmp Lib "kernel32" _
                         Alias "lstrcmpA" ( _
                         ByVal lpString1 As Long, _
                         ByVal lpString2 As Long) As Long
Private Declare Function EbExecuteLine Lib "vba6.dll" ( _
                         ByVal pStringToExec As Long, _
                         ByVal Foo1 As Long, _
                         ByVal Foo2 As Long, _
                         ByVal fCheckOnly As Long) As Long
Private Declare Function InitializeCriticalSection Lib "kernel32" ( _
                         ByRef lpCriticalSection As CRITICAL_SECTION) As Long
Private Declare Sub EnterCriticalSection Lib "kernel32" ( _
                    ByRef lpCriticalSection As CRITICAL_SECTION)
Private Declare Sub LeaveCriticalSection Lib "kernel32" ( _
                    ByRef lpCriticalSection As CRITICAL_SECTION)
Private Declare Sub DeleteCriticalSection Lib "kernel32" ( _
                    ByRef lpCriticalSection As CRITICAL_SECTION)
Private Declare Function TryEnterCriticalSection Lib "kernel32" ( _
                         ByRef lpCriticalSection As CRITICAL_SECTION) As Long
Private Declare Function IStream_Reset Lib "Shlwapi" ( _
                         ByVal pstm As Any) As Long
Private Declare Function rtcCallByName Lib "msvbvm60" ( _
                         ByRef vRet As Variant, _
                         ByVal cObj As Object, _
                         ByVal sMethod As Long, _
                         ByVal eCallType As VbCallType, _
                         ByRef pArgs() As Variant, _
                         ByVal lcid As Long) As Long
Private Declare Function vbaObjSet Lib "msvbvm60" _
                         Alias "__vbaObjSet" ( _
                         ByRef dstObject As Any, _
                         ByRef srcObjPtr As Any) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" _
                         Alias "__vbaObjSetAddref" ( _
                         ByRef dstObject As Any, _
                         ByRef srcObjPtr As Any) As Long
Private Declare Sub CoFreeUnusedLibraries Lib "ole32" ()
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
                         Alias "RegCreateKeyExW" ( _
                         ByVal hKey As Long, _
                         ByVal lpSubKey As Long, _
                         ByVal Reserved As Long, _
                         ByVal lpClass As Long, _
                         ByVal dwOptions As Long, _
                         ByVal samDesired As Long, _
                         ByRef lpSecurityAttributes As Any, _
                         ByRef phkResult As Long, _
                         ByRef lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
                         ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
                         Alias "RegSetValueExW" ( _
                         ByVal hKey As Long, _
                         ByVal lpValueName As Long, _
                         ByVal Reserved As Long, _
                         ByVal dwType As Long, _
                         ByRef lpData As Any, _
                         ByVal cbData As Long) As Long
Private Declare Function CreateIExprSrvObj Lib "MSVBVM60.DLL" ( _
                         ByVal pUnk1 As Long, _
                         ByVal lUnk2 As Long, _
                         ByVal pUnk3 As Long) As IUnknown
Private Declare Function CreateIExprSrvObj2 Lib "vba6" _
                         Alias "CreateIExprSrvObj" ( _
                         ByVal pUnk1 As Long, _
                         ByVal lUnk2 As Long, _
                         ByVal pUnk3 As Long) As IUnknown
Private Declare Function VBDllGetClassObject Lib "MSVBVM60.DLL" ( _
                         ByRef phModule As Long, _
                         ByVal lReserved As Long, _
                         ByVal pVBHeader As Long, _
                         ByRef pClsid As Any, _
                         ByRef pIID As Any, _
                         ByRef pObject As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
                         ByRef pvInstance As Any, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long
Private Declare Function VirtualProtect Lib "kernel32" ( _
                         ByVal lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal flNewProtect As Long, _
                         ByRef lpflOldProtect As Long) As Long
Private Declare Function TlsAlloc Lib "kernel32" () As Long
Private Declare Function TlsFree Lib "kernel32" ( _
                         ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsSetValue Lib "kernel32" ( _
                         ByVal dwTlsIndex As Long, _
                         ByRef lpTlsValue As Any) As Long
Private Declare Function TlsGetValue Lib "kernel32" ( _
                         ByVal dwTlsIndex As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal dwBytes As Long) As Long
Private Declare Function HeapCreate Lib "kernel32" ( _
                         ByVal flOptions As Long, _
                         ByVal dwInitialSize As Long, _
                         ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function HeapLock Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function HeapUnlock Lib "kernel32" ( _
                         ByVal hHeap As Long) As Long
Private Declare Function HeapWalk Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByRef lpEntry As PROCESS_HEAP_ENTRY) As Long
Private Declare Function HeapFree Lib "kernel32" ( _
                         ByVal hHeap As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpMem As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function CreateThread Lib "kernel32" ( _
                         ByRef lpThreadAttributes As Any, _
                         ByVal dwStackSize As Long, _
                         ByVal lpStartAddress As Long, _
                         ByRef lpParameter As Any, _
                         ByVal dwCreationFlags As Long, _
                         ByRef lpThreadId As Long) As Long
Private Declare Function CoInitialize Lib "ole32" ( _
                         ByRef pvReserved As Any) As Long
Private Declare Function CopyMemory Lib "kernel32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal Length As Long) As Long
Private Declare Function CoMarshalInterThreadInterfaceInStream Lib "ole32.dll" ( _
                         ByRef riid As Any, _
                         ByVal pUnk As Any, _
                         ByRef ppstm As Any) As Long
Private Declare Function CoMarshalInterface Lib "ole32.dll" ( _
                         ByVal pstm As Long, _
                         ByRef riid As Any, _
                         ByVal pUnk As Any, _
                         ByVal dwDestContext As Long, _
                         ByRef pvDestContext As Any, _
                         ByVal mshlflags As Long) As Long
Private Declare Function CoGetInterfaceAndReleaseStream Lib "ole32.dll" ( _
                         ByVal pstm As Long, _
                         ByRef riid As Any, _
                         ByRef pUnk As Any) As Long
Private Declare Function CoUnmarshalInterface Lib "ole32.dll" ( _
                         ByVal pstm As Long, _
                         ByRef riid As Any, _
                         ByRef pUnk As Any) As Long
Private Declare Function CoReleaseMarshalData Lib "ole32" ( _
                         ByVal pstm As Long) As Long
Private Declare Function GetMessage Lib "user32" _
                         Alias "GetMessageW" ( _
                         ByRef lpMsg As Any, _
                         ByVal hwnd As Long, _
                         ByVal wMsgFilterMin As Long, _
                         ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" ( _
                         ByRef lpMsg As Any) As Long
Private Declare Function DispatchMessage Lib "user32" _
                         Alias "DispatchMessageW" ( _
                         ByRef lpMsg As Any) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" ( _
                         ByVal hHandle As Long, _
                         ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" _
                         Alias "CreateEventW" ( _
                         ByRef lpEventAttributes As Any, _
                         ByVal bManualReset As Long, _
                         ByVal bInitialState As Long, _
                         ByVal lpName As Long) As Long
Private Declare Function SetEvent Lib "kernel32" ( _
                         ByVal hEvent As Long) As Long
Private Declare Function CLSIDFromProgID Lib "ole32" ( _
                         ByVal TSzProgID As Long, _
                         ByRef pGuid As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" ( _
                         ByVal rclsid As Long, _
                         ByVal pUnkOuter As Long, _
                         ByVal dwClsContext As Long, _
                         ByVal riid As Long, _
                         ByRef ppv As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" ( _
                         ByVal hGlobal As Long, _
                         ByVal fDeleteOnRelease As Long, _
                         ByRef ppstm As Any) As Long
Private Declare Function VariantCopy Lib "oleaut32" ( _
                         ByRef pvargDest As Any, _
                         ByRef pvargSrc As Any) As Long
Private Declare Function VariantClear Lib "oleaut32" ( _
                         ByRef pvarg As Any) As Long
Private Declare Function VariantCopyInd Lib "oleaut32" ( _
                         ByRef pvarDest As Any, _
                         ByRef pvargSrc As Any) As Long
Private Declare Function PostThreadMessage Lib "user32" _
                         Alias "PostThreadMessageW" ( _
                         ByVal idThread As Long, _
                         ByVal msg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
Private Declare Function lstrlen Lib "kernel32" _
                         Alias "lstrlenA" ( _
                         ByVal lpString As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" ( _
                         ByRef m_pBase As Any, _
                         ByVal sz As Long) As String
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
                         ByVal dwFlags As Long, _
                         ByVal th32ProcessID As Long) As Long
Private Declare Function Thread32First Lib "kernel32" ( _
                         ByVal hSnapshot As Long, _
                         ByRef lpte As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" ( _
                         ByVal hSnapshot As Long, _
                         ByRef lpte As THREADENTRY32) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function GetClassLong Lib "user32" _
                         Alias "GetClassLongW" ( _
                         ByVal hwnd As Long, _
                         ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" _
                         Alias "SetClassLongW" ( _
                         ByVal hwnd As Long, _
                         ByVal nIndex As Long, _
                         ByVal dwNewLong As Long) As Long
Private Declare Function GetClassInfoEx Lib "user32" _
                         Alias "GetClassInfoExW" ( _
                         ByVal hInstance As Long, _
                         ByVal lpClassName As Long, _
                         ByRef lpWndClassEx As WNDCLASSEX) As Long
Private Declare Function UnregisterClass Lib "user32" _
                         Alias "UnregisterClassW" ( _
                         ByVal lpClassName As Long, _
                         ByVal hInstance As Long) As Long
Private Declare Function RegisterClassEx Lib "user32" _
                         Alias "RegisterClassExW" ( _
                         ByRef pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function CreateWindowEx Lib "user32" _
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
Private Declare Function GetModuleHandle Lib "kernel32" _
                         Alias "GetModuleHandleW" ( _
                         ByVal lpModuleName As Long) As Long
Private Declare Function DestroyWindow Lib "user32" ( _
                         ByVal hwnd As Long) As Long
Private Declare Function SuspendThread Lib "kernel32" ( _
                         ByVal hThread As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" ( _
                         ByVal hThread As Long) As Long
Private Declare Function OpenThread Lib "kernel32" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwThreadId As Long) As Long
Private Declare Function PeekMessage Lib "user32" _
                         Alias "PeekMessageW" ( _
                         ByRef lpMsg As msg, _
                         ByVal hwnd As Long, _
                         ByVal wMsgFilterMin As Long, _
                         ByVal wMsgFilterMax As Long, _
                         ByVal wRemoveMsg As Long) As Long


Private Declare Sub ZeroMemory Lib "kernel32" _
                    Alias "RtlZeroMemory" ( _
                    ByRef Destination As Any, _
                    ByVal Length As Long)
Private Declare Sub CoUninitialize Lib "ole32" ()
                         
' // Export to close handles from clients
Public Declare Function CloseHandle Lib "kernel32" ( _
                        ByVal hObject As Long) As Long

Private lTlsSlot        As Long             ' // Index of the item in the TLS. There will be data specific to the thread.
Private pVBHeader       As Long             ' // Pointer to VBHeader structure.
Private hModule         As Long             ' // Base address.
Private tLockMarshal    As tCriticalSection ' // Critical section for multiple-time marshaling
Private tLockHeap       As tCriticalSection ' // Critical section for cleaning heap
Private hHeadersHeap    As Long

' // IN IDE ONLY
Private hCodeHeap       As Long     ' // Heap for dynamic code
Private hMsgWindow      As Long     ' // Handle of message window

' // Initialize
Public Function Initialize() As Boolean
    Dim bIsinIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    If Not bIsinIDE Then
        
        InitializeCriticalSection tLockMarshal.tWinApiSection
        tLockMarshal.bIsInitialized = True
        
        InitializeCriticalSection tLockHeap.tWinApiSection
        tLockHeap.bIsInitialized = True
        
        hModule = App.hInstance
        
        pVBHeader = GetVBHeader()
        If pVBHeader = 0 Then GoTo CleanUp
        
        ModifyVBHeader AddressOf FakeMain
        
        hHeadersHeap = HeapCreate(0, 0, 65536)
        If hHeadersHeap = 0 Then GoTo CleanUp

    End If
    
    lTlsSlot = TlsAlloc()
    If lTlsSlot = TLS_OUT_OF_INDEXES Then GoTo CleanUp
    
    Initialize = True
    
CleanUp:
    
    If Not Initialize Then
        Uninitialize
    End If
    
End Function

' // Uninitialize resources
' // WARNING! Don't call it if a thread uses resources
Public Sub Uninitialize()
    Dim bIsinIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    If bIsinIDE Then
    
        If hCodeHeap Then HeapDestroy hCodeHeap
        If hMsgWindow Then
        
            DestroyWindow hMsgWindow
            UnregisterClass StrPtr(MESSAGE_WINDOW_CLASS), App.hInstance
            
        End If
        
        hMsgWindow = 0
        hCodeHeap = 0
        
    End If
    
    If tLockMarshal.bIsInitialized Then DeleteCriticalSection tLockMarshal.tWinApiSection
    If tLockHeap.bIsInitialized Then DeleteCriticalSection tLockHeap.tWinApiSection
    If hHeadersHeap Then HeapDestroy hHeadersHeap
    If lTlsSlot Then TlsFree lTlsSlot
    
    FreeUnusedHeaders
    
    tLockMarshal.bIsInitialized = False
    tLockHeap.bIsInitialized = False
    hHeadersHeap = 0
    lTlsSlot = 0
    
End Sub

' // Create a new thread
Public Function vbCreateThread(ByVal lpThreadAttributes As Long, _
                               ByVal dwStackSize As Long, _
                               ByVal lpStartAddress As Long, _
                               ByVal lpParameter As Long, _
                               ByVal dwCreationFlags As Long, _
                               ByRef lpThreadId As Long, _
                               Optional ByVal bIDEInSameThread As Boolean = True) As Long
    Dim bIsinIDE    As Boolean
    Dim hr          As Long
    Dim pThreadData As Long
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    If bIsinIDE Then

        If bIDEInSameThread Then
        
            ' // Run function in main thread
            hr = DispCallFunc(ByVal 0&, lpStartAddress, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(CVar(lpParameter)), CVar(0))
            
            If hr Then
                Err.Raise hr
            End If
            
            Exit Function

        End If
        
        Err.Raise 5
        
    End If
    
    pThreadData = PrepareData(lpStartAddress, lpParameter)
    If pThreadData = 0 Then Exit Function

    ' // Create thread
    vbCreateThread = CreateThread(ByVal lpThreadAttributes, _
                                  dwStackSize, _
                                  AddressOf ThreadProc, _
                                  ByVal pThreadData, _
                                  dwCreationFlags, _
                                  lpThreadId)
    
End Function

' // Allow marshaling of private classes
Public Function EnablePrivateMarshaling( _
                ByVal bEnable As Boolean) As Boolean
    Dim hKey        As Long
    Dim lType       As Long
    Dim bData(255)  As Byte
    Dim lSize       As Long
    Dim lRet        As Long
    Dim sValue      As String
    
    If bEnable Then
        sValue = "1"
    Else
        sValue = "0"
    End If
    
    If RegCreateKeyEx(HKEY_CURRENT_USER, StrPtr("Software\Microsoft\Visual Basic\6.0"), 0, 0, _
                      REG_OPTION_NON_VOLATILE, KEY_WRITE Or KEY_QUERY_VALUE, ByVal 0&, hKey, 0) Then
        Exit Function
    End If

    If RegSetValueEx(hKey, StrPtr("AllowUnsafeObjectPassing"), 0, REG_SZ, ByVal StrPtr(sValue), LenB(sValue) + 2) Then
        GoTo CleanUp
    End If
    
    EnablePrivateMarshaling = True
    
CleanUp:
        
    RegCloseKey hKey
    
End Function

' // FOR IDE ONLY
' // Get address of callback procedure that calls user function in main thread
' // It uses a window to transmit call from callback thread to main thread
' // When callback procedure is being called asm-thunk call SendMessage to
' // message-only window that is in main thread. Window proc of that window
' // calls pfnCallback function and passes pointer to stack of caller thread
Public Function InitCurrentThreadAndCallFunctionIDEProc( _
                ByVal pfnCallback As Long, _
                ByVal lParametersSize As Long) As Long
    Dim ptr             As Long
    Dim hUser32         As Long
    Dim pfnSendMessage  As Long
    
    ' // Initialize message window and heap
    hMsgWindow = InitializeMessageWindow()
    If hMsgWindow = 0 Then Exit Function
    
    ' // Check if thunk already exists
    ptr = FindThunk(pfnCallback)
    
    If ptr Then
        InitCurrentThreadAndCallFunctionIDEProc = ptr
        Exit Function
    End If
    
    ' // Make asm-thunk
    
    ' // LEA EAX, [ESP+4]
    ' // PUSH EAX
    ' // PUSH pfnCallback
    ' // PUSH WM_ONCALLBACK
    ' // PUSH hMsgWindow
    ' // Call SendMessageW
    ' // RETN lParametersSize

    hUser32 = GetModuleHandle(StrPtr("user32"))
    
    ptr = HeapAlloc(hCodeHeap, 0, &H1C)
    If ptr = 0 Then Exit Function
    
    pfnSendMessage = GetProcAddress(hUser32, "SendMessageW") - ptr - &H19
    
    GetMem8 11469288874.1005@, ByVal ptr
    GetMem8 749398979713119.0272@, ByVal ptr + &H8
    GetMem8 99643241.2672@, ByVal ptr + &H10
    GetMem4 &HC200&, ByVal ptr + &H18
    
    GetMem4 pfnCallback, ByVal ptr + &H6
    GetMem4 hMsgWindow, ByVal ptr + &H10
    GetMem2 lParametersSize, ByVal ptr + &H1A
    GetMem4 pfnSendMessage, ByVal ptr + &H15
    
    InitCurrentThreadAndCallFunctionIDEProc = ptr
    
End Function

' // Initialize message window
Private Function InitializeMessageWindow() As Long
    Dim tClass      As WNDCLASSEX
    Dim pWndProc    As Long
    Dim bRegistered As Boolean
    
    If hMsgWindow Then
    
        InitializeMessageWindow = hMsgWindow
        Exit Function
        
    End If
    
    tClass.cbSize = Len(tClass)
    
    ' // Check if class already registered
    If GetClassInfoEx(App.hInstance, StrPtr(MESSAGE_WINDOW_CLASS), tClass) = 0 Then
        
        ' // Create heap for asm-thunks
        If hCodeHeap = 0 Then
            
            hCodeHeap = HeapCreate(HEAP_CREATE_ENABLE_EXECUTE, 0, 0)
            If hCodeHeap = 0 Then GoTo CleanUp
            
        End If
        
        ' // Create window proc
        pWndProc = CreateWndProcCode()
        If pWndProc = 0 Then GoTo CleanUp
        
        tClass.hInstance = App.hInstance
        tClass.lpfnwndproc = pWndProc
        tClass.lpszClassName = StrPtr(MESSAGE_WINDOW_CLASS)
        tClass.cbClsextra = 8   ' // 0x00 - hCodeHeap
                                ' // 0x04 - hMsgWindow
                                ' // 0x08 - size
                                
        If RegisterClassEx(tClass) = 0 Then GoTo CleanUp
        
        bRegistered = True

    End If
    
    ' // Create window
    hMsgWindow = CreateWindowEx(0, StrPtr(MESSAGE_WINDOW_CLASS), 0, 0, 0, 0, 0, 0, _
                                HWND_MESSAGE, 0, App.hInstance, ByVal 0&)
    If hMsgWindow = 0 Then Exit Function
    
    If Not bRegistered Then
        
        ' // Destroy previous session window and get heap
        DestroyWindow GetClassLong(hMsgWindow, 4)
        hCodeHeap = GetClassLong(hMsgWindow, 0)
        
        ' // Free unused thunks
        FreeThunks tClass.lpfnwndproc
        
    End If
    
    ' // Save global data
    SetClassLong hMsgWindow, 0, hCodeHeap
    SetClassLong hMsgWindow, 4, hMsgWindow
    
    InitializeMessageWindow = hMsgWindow
    
    Exit Function
    
CleanUp:
    
    If pWndProc Then HeapFree hCodeHeap, 0, pWndProc
    If bRegistered Then UnregisterClass StrPtr(MESSAGE_WINDOW_CLASS), App.hInstance
    If hCodeHeap Then HeapDestroy hCodeHeap
    
    hCodeHeap = 0
    hMsgWindow = 0
    
End Function

' // Free thunks
Private Sub FreeThunks( _
            ByVal pfnWndProc As Long)
    Dim tEntry      As PROCESS_HEAP_ENTRY
    Dim pPointers() As Long
    Dim lIndex      As Long
    Dim lCount      As Long

    ReDim pPointers(255)
    
    If HeapLock(hCodeHeap) = 0 Then Exit Sub
    
    Do While HeapWalk(hCodeHeap, tEntry)

        If tEntry.wFlags And PROCESS_HEAP_ENTRY_BUSY Then

            If tEntry.lpData <> pfnWndProc Then
                
                If lIndex > UBound(pPointers) Then
                    ReDim Preserve pPointers(lIndex + 256)
                End If
                
                pPointers(lIndex) = tEntry.lpData
                
                lIndex = lIndex + 1
                
            End If
            
        End If

    Loop

    HeapUnlock hCodeHeap
    
    lCount = lIndex
    
    For lIndex = 0 To lCount - 1
        HeapFree hCodeHeap, 0, pPointers(lIndex)
    Next

End Sub

Private Function FindThunk( _
                 ByVal pfn As Long) As Long
    Dim tEntry  As PROCESS_HEAP_ENTRY
    Dim lIndex  As Long
    Dim pfnCur  As Long
    
    If HeapLock(hCodeHeap) = 0 Then Exit Function
    
    Do While HeapWalk(hCodeHeap, tEntry)

        If tEntry.wFlags And PROCESS_HEAP_ENTRY_BUSY Then

            GetMem4 ByVal tEntry.lpData + &H6, pfnCur
            
            If pfnCur = pfn Then
            
                FindThunk = tEntry.lpData
                Exit Do
                
            End If
            
        End If

    Loop

    HeapUnlock hCodeHeap
           
End Function

' // Create window proc
Private Function CreateWndProcCode() As Long
    Dim ptr                 As Long
    Dim hUser32             As Long
    Dim pfnDefWindowProc    As Long
    
    ptr = HeapAlloc(hCodeHeap, 0, &H1A)
    If ptr = 0 Then Exit Function
    
    hUser32 = GetModuleHandle(StrPtr("user32"))
    pfnDefWindowProc = GetProcAddress(hUser32, "DefWindowProcW") - ptr - &HF
    
    ' //    CMP DWORD [ESP+8], WM_ONCALLBACK
    ' //    JE SHORT L
    ' //    JMP DefWindowProcW
    ' // L: PUSH DWORD PTR SS:[ESP+10]
    ' //    CALL DWORD PTR SS:[ESP+10]
    ' //    RETN 10

    GetMem8 439819570.2913@, ByVal ptr
    GetMem8 -7205759402265.6652@, ByVal ptr + 8
    GetMem8 -446281617701647.8604@, ByVal ptr + &H10
    GetMem2 16&, ByVal ptr + &H18
    GetMem4 pfnDefWindowProc, ByVal ptr + &HB
    
    CreateWndProcCode = ptr
    
End Function

' // Init thread and call function
' // This function is useful for callback of WINAPI functions which can call function in arbitrary thread
Public Function InitCurrentThreadAndCallFunction( _
                ByVal pfnCallback As Long, _
                ByVal pParam As Long, _
                ByRef lReturnValue As Long) As Boolean
    Dim cExpSrv     As IUnknown
    Dim pThreadData As Long
    Dim hr          As Long
    Dim vRet        As Variant
    Dim bIsinIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    If bIsinIDE Then
        ' // Error
        Exit Function
    End If
    
    Set cExpSrv = CreateIExprSrvObj(0, 4, 0)
    
    If lTlsSlot Then
    
        ' // Check if thread already initialized
        pThreadData = TlsGetValue(lTlsSlot)
        
        If pThreadData <> 0 Then
            
            ' // Just call by pointer
            hr = DispCallFunc(ByVal 0&, pfnCallback, CC_STDCALL, vbLong, 1, vbLong, VarPtr(CVar(pParam)), vRet)
            
            If hr Then
                Err.Raise hr
            End If
            
            lReturnValue = vRet
            
            InitCurrentThreadAndCallFunction = True
            
            Exit Function
            
        End If
    
    End If
    
    pThreadData = PrepareData(pfnCallback, pParam)
    If pThreadData = 0 Then Exit Function
    
    Set cExpSrv = Nothing
    
    lReturnValue = ThreadProc(pThreadData)
    
End Function

' // Create new private object in the new thread by name
Public Function CreatePrivateObjectByNameInNewThread( _
                ByRef sClassName As String, _
                Optional ByVal pIID As Long, _
                Optional ByRef lAsynchId As Long) As IUnknown
    Dim tThreadData     As tNewObjectThreadData
    Dim hThread         As Long
    Dim tIID_IDispatch  As tCurGUID
    Dim sAnsiName       As String
    Dim bIsinIDE        As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    sAnsiName = StrConv(sClassName, vbFromUnicode)
    
    tThreadData.eFlags = OTDF_PRIVATE
    
    tThreadData.pClsid = StrPtr(sAnsiName)
    
    If pIID = 0 Then
    
        tIID_IDispatch.c1 = 13.2096@
        tIID_IDispatch.c2 = 504403158265495.5712@
        
        tThreadData.pIID = VarPtr(tIID_IDispatch)
        
    Else:   tThreadData.pIID = pIID
    End If
        
    If bIsinIDE Then
    
        ActiveXThreadProc tThreadData
        vbaObjSet CreatePrivateObjectByNameInNewThread, ByVal tThreadData.pStream
        lAsynchId = tThreadData.pStream
        
    Else
        
        tThreadData.hEvent = CreateEvent(ByVal 0&, 1, 0, 0)
        If tThreadData.hEvent = 0 Then Err.Raise 7
        
        hThread = vbCreateThread(0, 0, AddressOf ActiveXThreadProc, VarPtr(tThreadData), 0, lAsynchId)
        
        ' // Wait for object creation
        WaitForSingleObject tThreadData.hEvent, -1
        
        CloseHandle tThreadData.hEvent
        
        If tThreadData.hr Then
            Err.Raise tThreadData.hr
        End If
        
        If tThreadData.pStream = 0 Then Exit Function
        
        Set CreatePrivateObjectByNameInNewThread = UnMarshal(tThreadData.pStream, pIID)
        
    End If
    
End Function

' // Create new ActiveX object in the new thread by ProgID
Public Function CreateActiveXObjectInNewThread2( _
                ByRef sProgID As String, _
                Optional ByVal pIID As Long, _
                Optional ByRef lAsynchId As Long) As IUnknown
    Dim tClsId  As tCurGUID
    
    If CLSIDFromProgID(StrPtr(sProgID), tClsId) < 0 Then
        Err.Raise 5
    End If
    
    Set CreateActiveXObjectInNewThread2 = CreateActiveXObjectInNewThread(VarPtr(tClsId), pIID, lAsynchId)
    
End Function

' // Create new ActiveX object in the new thread
Public Function CreateActiveXObjectInNewThread( _
                ByVal pClsid As Long, _
                Optional ByVal pIID As Long, _
                Optional ByRef lAsynchId As Long) As IUnknown
    Dim tThreadData     As tNewObjectThreadData
    Dim hThread         As Long
    Dim tIID_IDispatch  As tCurGUID
    Dim bIsinIDE        As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    tThreadData.eFlags = OTDF_ACTIVEX
    tThreadData.pClsid = pClsid
    
    If pIID = 0 Then
    
        tIID_IDispatch.c1 = 13.2096@
        tIID_IDispatch.c2 = 504403158265495.5712@
        
        tThreadData.pIID = VarPtr(tIID_IDispatch)
        
    Else:   tThreadData.pIID = pIID
    End If
        
    If bIsinIDE Then
        
        ActiveXThreadProc tThreadData
        vbaObjSet CreateActiveXObjectInNewThread, ByVal tThreadData.pStream
        lAsynchId = tThreadData.pStream
        
    Else
    
        tThreadData.hEvent = CreateEvent(ByVal 0&, 1, 0, 0)
        If tThreadData.hEvent = 0 Then Err.Raise 7
        
        hThread = vbCreateThread(0, 0, AddressOf ActiveXThreadProc, VarPtr(tThreadData), 0, lAsynchId)
    
        WaitForSingleObject tThreadData.hEvent, -1
        
        CloseHandle tThreadData.hEvent
        
        If tThreadData.hr Then
            Err.Raise tThreadData.hr
        End If
        
        If tThreadData.pStream = 0 Then Exit Function
        
        Set CreateActiveXObjectInNewThread = UnMarshal(tThreadData.pStream, pIID)
    
    End If

End Function

' // Wait for object thread completion
Public Sub WaitForObjectThreadCompletion( _
           ByVal lAsynchId As Long)
    Dim hThread     As Long
    Dim bIsinIDE    As Boolean
    Dim tMSG        As msg
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    ' // Unsupported in IDE
    If bIsinIDE Then Exit Sub
    
    hThread = OpenThread(SYNCHRONIZE, 0, lAsynchId)
    If hThread = 0 Then Exit Sub
    
    Do
    
        Select Case MsgWaitForMultipleObjects(1, hThread, 0, -1, &H5FF)
        Case 0: Exit Do
        Case 1
            
            PeekMessage tMSG, 0, 0, 0, PM_REMOVE
            TranslateMessage tMSG
            DispatchMessage tMSG
            
        Case Else
            ' // Error
            Exit Do
        End Select
    
    Loop
    
    CloseHandle hThread
      
End Sub

' // Suspend/Resume object thread
Public Sub SuspendResume( _
           ByVal lAsynchId As Long, _
           ByVal bSuspend As Boolean)
    Dim hThread     As Long
    Dim bIsinIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    ' // Unsupported in IDE
    If bIsinIDE Then Exit Sub
    
    hThread = OpenThread(THREAD_SUSPEND_RESUME, 0, lAsynchId)
    If hThread = 0 Then Exit Sub
    
    If bSuspend Then
        SuspendThread hThread
    Else
        ResumeThread hThread
    End If
    
    CloseHandle hThread
    
End Sub

' // Asynch call of ActiveX object in thread
' // lAsynchId - thread identifier of object which object was created in (using current module)
' // sMethodName - method name to call
' // eCallType - call type (property, method, etc.)
' // cCallBackObject (optional) - object that receives callback
' // sCallBackMethod (optional) - callback method name
' // vArgs - arguments
' // WARNING! Marshaling of parameters is not supported yet, be careful passing object variables
' // If you need to pass an object variable make it through synch method and save a marhsaled
' // instance inside object.
Public Sub AsynchDispMethodCall( _
           ByVal lAsynchId As Long, _
           ByRef sMethodName As String, _
           ByVal eCallType As VbCallType, _
           ByVal cCallbackObject As Object, _
           ByRef sCallBackMethod As String, _
           ParamArray vArgs() As Variant)
    Dim tAsynchData As tAsynchCallData
    Dim pData       As Long
    Dim pStruct     As Long
    Dim bIsinIDE    As Boolean
    Dim cObj        As IUnknown
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    tAsynchData.sMethodName = sMethodName
    tAsynchData.sCallBackName = sCallBackMethod
    tAsynchData.eCallType = eCallType
    tAsynchData.vArgs = vArgs
        
    If bIsinIDE Then
        
        tAsynchData.pStream = ObjPtr(cCallbackObject)
        vbaObjSetAddref cObj, ByVal lAsynchId
        MakeAsynchCall cObj, VarPtr(tAsynchData)
        
    Else
    
        ' // Marshal callback object
        If Not cCallbackObject Is Nothing Then
            tAsynchData.pStream = Marshal(cCallbackObject)
        End If
        
        ' // Allocate memory to hold tAsynchCallData structure
        pData = HeapAlloc(GetProcessHeap(), 0, Len(tAsynchData))
        
        If pData = 0 Then
            
            FreeMarshalData tAsynchData.pStream
            Err.Raise 7
    
        End If
        
        ' // Avoid unicode/ansi conversion
        pStruct = VarPtr(tAsynchData)
        
        CopyMemory ByVal pData, ByVal pStruct, Len(tAsynchData)
        
        ' // Post message to thread
        If PostThreadMessage(lAsynchId, WM_ASYNCH_CALL, 0, pData) = 0 Then
        
            HeapFree GetProcessHeap(), 0, pData
            FreeMarshalData tAsynchData.pStream
            
            Err.Raise 5
            
        End If
        
        ' // Avoid memory cleaning (it will freed in new thread)
        ZeroMemory ByVal pStruct, Len(tAsynchData)
        
    End If
    
End Sub

' // Marshal an interface
Public Function Marshal( _
                ByVal cObject As IUnknown, _
                Optional ByVal pInterface As Long) As Long
    Dim tIID_IDispatch  As tCurGUID
    Dim hr              As Long
    
    If pInterface = 0 Then
    
        tIID_IDispatch.c1 = 13.2096@
        tIID_IDispatch.c2 = 504403158265495.5712@
        
        pInterface = VarPtr(tIID_IDispatch)
        
    End If

    hr = CoMarshalInterThreadInterfaceInStream(ByVal pInterface, cObject, Marshal)

    If hr Then
        Err.Raise hr
    End If
    
End Function

' // Marshal an interface for many times
' // That methond intends for multiple-times marshaling.
Public Function Marshal2( _
                ByVal cObject As IUnknown, _
                Optional ByVal pInterface As Long) As Long
    Dim tIID_IDispatch  As tCurGUID
    Dim hr              As Long
    Dim pstm            As Long
    
    ' // If interface is not specified use IDispatch one
    If pInterface = 0 Then
    
        tIID_IDispatch.c1 = 13.2096@
        tIID_IDispatch.c2 = 504403158265495.5712@
        
        pInterface = VarPtr(tIID_IDispatch)
        
    End If
    
    hr = CreateStreamOnHGlobal(0, 1, pstm)
    
    If hr Then
        Err.Raise hr
    End If
    
    hr = CoMarshalInterface(pstm, ByVal pInterface, cObject, MSHCTX_INPROC, ByVal 0&, MSHLFLAGS_TABLESTRONG)
     
    If hr Then
    
        vbaObjSet pstm, ByVal 0&
        Err.Raise hr
        
    End If
    
    Marshal2 = pstm
    
End Function
       
' // Free marshal data
Public Sub FreeMarshalData( _
           ByRef pStream As Long)
    Dim hr  As Long

    EnterCriticalSection tLockMarshal.tWinApiSection

    hr = IStream_Reset(pStream)
    
    If hr >= 0 Then
        hr = CoReleaseMarshalData(pStream)
    End If

    LeaveCriticalSection tLockMarshal.tWinApiSection

    vbaObjSet pStream, ByVal 0&
    
    If hr Then
        Err.Raise hr
    End If
    
End Sub
       
' // Unmarshal a interface
Public Function UnMarshal( _
                ByVal pStream As Long, _
                Optional ByVal pInterface As Long, _
                Optional ByVal bReleaseStream As Boolean = True) As IUnknown
    Dim tIID_IDispatch  As tCurGUID
    Dim hr              As Long
    Dim lStmSize        As Long
    Dim cSize           As Currency
    
    If pInterface = 0 Then
    
        tIID_IDispatch.c1 = 13.2096@
        tIID_IDispatch.c2 = 504403158265495.5712@
        
        pInterface = VarPtr(tIID_IDispatch)
        
    End If
    
    If bReleaseStream Then
        hr = CoGetInterfaceAndReleaseStream(pStream, ByVal pInterface, UnMarshal)
    Else
        
        If Not tLockMarshal.bIsInitialized Then
            
            InitializeCriticalSection tLockMarshal.tWinApiSection
            tLockMarshal.bIsInitialized = True
            
        End If
        
        ' // To ensure zero offset in stream lock access
        EnterCriticalSection tLockMarshal.tWinApiSection
        
        hr = IStream_Reset(pStream)
        
        If hr >= 0 Then
            hr = CoUnmarshalInterface(pStream, ByVal pInterface, UnMarshal)
        End If
        
        LeaveCriticalSection tLockMarshal.tWinApiSection
        
    End If
    
    If hr Then
        Err.Raise hr
    End If
    
End Function

' // ActiveX object thread
Private Function ActiveXThreadProc( _
                 ByRef tThreadData As tNewObjectThreadData) As Long
    Dim cObj        As IUnknown
    Dim cObjDead    As IUnknown
    Dim pObjDeadPtr As Long
    Dim tMSG        As msg
    Dim vRet        As Variant
    Dim lRet        As Long
    Dim hr          As Long
    Dim bIsinIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    If tThreadData.eFlags = OTDF_ACTIVEX Then
        tThreadData.hr = CoCreateInstance(tThreadData.pClsid, 0, CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER, _
                                        tThreadData.pIID, cObj)
    Else
        tThreadData.hr = CreatePrivateClass(tThreadData.pClsid, tThreadData.pIID, cObj)
    End If
    
    If bIsinIDE Then
        vbaObjSetAddref tThreadData.pStream, ByVal ObjPtr(cObj)
    Else
    
        If tThreadData.hr < 0 Then
            
            SetEvent tThreadData.hEvent
            Exit Function
            
        End If
        
        On Error Resume Next
        
        tThreadData.pStream = Marshal(cObj, tThreadData.pIID)
        
        If Err.Number Then
            tThreadData.hr = Err.Number
        End If
        
        On Error GoTo -1
        
        SetEvent tThreadData.hEvent
        
        Set cObjDead = cObj ' // Add the reference = 2
        
        pObjDeadPtr = ObjPtr(cObjDead)
        
        Do
        
            lRet = GetMessage(tMSG, 0, 0, 0)
            If lRet = -1 Or lRet = 0 Then Exit Do
            
            If tMSG.message = WM_ASYNCH_CALL And tMSG.hwnd = 0 Then
                MakeAsynchCall cObj, tMSG.lParam
            Else
    
                TranslateMessage tMSG
                DispatchMessage tMSG
            
            End If
            
            ' // Check if object if freed by calling IUnknown::Release
            hr = DispCallFunc(ByVal pObjDeadPtr, 8, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, vRet)
            
            GetMem4 0&, cObjDead
            
            ' // If only cObj references to object
            If CLng(vRet) = 1 Then
    
                Set cObj = Nothing      ' // Release object
                Exit Do
            
            Else: Set cObjDead = cObj   ' // Add reference
            End If
            
        Loop
    
        CoFreeUnusedLibraries
    
    End If
    
End Function

' // Create private class
Private Function CreatePrivateClass( _
                 ByVal pClassName As Long, _
                 ByVal pIID As Long, _
                 ByRef cObj As IUnknown) As Long
    Dim bIsinIDE        As Boolean
    Dim pProjInfo       As Long
    Dim pObjTable       As Long
    Dim pObjDesc        As Long
    Dim lTotalObjects   As Long
    Dim lIndex          As Long
    Dim lModuleType     As Long
    Dim sClassName      As String
    Dim pModname        As Long
    Dim pObjInfo        As Long
    Dim cTempObj        As IUnknown
    Dim iTypes(1)       As Integer
    Dim pArgs(1)        As Long
    Dim vArgs(1)        As Variant
    Dim vResult         As Variant
    
    Debug.Assert MakeTrue(bIsinIDE)

    If bIsinIDE Then
        
        sClassName = SysAllocStringByteLen(ByVal pClassName, lstrlen(pClassName))
        EbExecuteLine StrPtr("modMultithreading.QueueObject new " & sClassName), 0, 0, 0
        Set cTempObj = QueueObject(Nothing)
        
    Else
        
        ' // Go thru modules
        GetMem4 ByVal pVBHeader + &H30, pProjInfo
        GetMem4 ByVal pProjInfo + &H4, pObjTable
        GetMem2 ByVal pObjTable + &H2A, lTotalObjects
        GetMem4 ByVal pObjTable + &H30, pObjDesc
        
        For lIndex = 0 To lTotalObjects - 1
            
            GetMem4 ByVal pObjDesc + &H28, lModuleType
            GetMem4 ByVal pObjDesc + &H18, pModname
            
            ' // Only object modules
            If (lModuleType And 3) = 3 Then
                If lstrcmp(pClassName, pModname) = 0 Then
                    
                    GetMem4 ByVal pObjDesc, pObjInfo
                    Set cTempObj = vbaNew(ByVal pObjInfo)
                    
                    Exit For
                    
                End If
            End If
            
            pObjDesc = pObjDesc + &H30
            
        Next

    End If
    
    If cTempObj Is Nothing Then
    
        CreatePrivateClass = &H80040154
        Exit Function
        
    End If
    
    ' // Query interface
    iTypes(0) = vbLong:             iTypes(1) = vbLong
    vArgs(0) = pIID:                vArgs(1) = VarPtr(cObj)
    pArgs(0) = VarPtr(vArgs(0)):    pArgs(1) = VarPtr(vArgs(1))
    
    CreatePrivateClass = DispCallFunc(ByVal ObjPtr(cTempObj), 0, CC_STDCALL, vbLong, 2, iTypes(0), pArgs(0), vResult)
    
    If CreatePrivateClass >= 0 Then
        CreatePrivateClass = vResult
    End If
    
End Function

Private Function QueueObject( _
                 ByVal cObj As IUnknown) As IUnknown
    Static cTmpObj  As IUnknown
    Set QueueObject = cTmpObj
    Set cTmpObj = cObj
End Function

' // Make Asynch call
Private Function MakeAsynchCall( _
                 ByVal cObject As IUnknown, _
                 ByVal pData As Long) As Long
    Dim tData       As tAsynchCallData
    Dim cCallBack   As Object
    Dim vRet        As Variant
    Dim pStruct     As Long
    Dim pvArgs      As Long
    Dim bIsinIDE    As Boolean
    
    On Error GoTo error_handler
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    ' // Copy Structure
    pStruct = VarPtr(tData)
    
    CopyMemory ByVal pStruct, ByVal pData, Len(tData)
    
    If bIsinIDE Then
         
        vbaObjSetAddref cCallBack, ByVal tData.pStream
        
        rtcCallByName vRet, cObject, StrPtr(tData.sMethodName), tData.eCallType, tData.vArgs, &H409
        
        If Not cCallBack Is Nothing Then
            CallByName cCallBack, tData.sCallBackName, VbMethod, vRet
        End If
        
        ZeroMemory ByVal pStruct, Len(tData)
        
    Else
        
        ' // Unmarshal callback object pointer
        If tData.pStream Then
            Set cCallBack = UnMarshal(tData.pStream)
        End If
        
        ' // Call mathod
        rtcCallByName vRet, cObject, StrPtr(tData.sMethodName), tData.eCallType, tData.vArgs, &H409
        
        ' // Callback return value
        If Not cCallBack Is Nothing Then
            CallByName cCallBack, tData.sCallBackName, VbMethod, vRet
        End If
        
    End If
    
    Exit Function
    
error_handler:
    
    MakeAsynchCall = Err.Number
    
End Function

' // Prepare data
Private Function PrepareData( _
                 ByVal lpStartAddress As Long, _
                 ByVal lpParameter As Long) As Long
    Dim bIsinIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)

    Dim pThreadData As Long
    Dim tThreadData As tThreadData
    
    ' // Allocate thread-specific memory for tThreadData structure
    pThreadData = HeapAlloc(GetProcessHeap(), 0, Len(tThreadData))
    
    If pThreadData = 0 Then Exit Function
    
    ' // Set parameters
    tThreadData.lpAddress = lpStartAddress
    tThreadData.lpParameter = lpParameter

    ' // Copy parameters to thread-specific memory
    CopyMemory ByVal pThreadData, tThreadData, Len(tThreadData)
    
    PrepareData = pThreadData
    
End Function

' // Initialize runtime for new thread and run procedure
Private Function ThreadProc( _
                 ByVal pParameter As Long) As Long
    Dim cExpSrv     As IUnknown
    Dim hr          As Long
    Dim bIsinIDE    As Boolean
    Dim tClsId      As tCurGUID
    Dim tIID        As tCurGUID
    Dim tThreadData As tThreadData
    Dim hHeap       As Long
    Dim pContext    As Long
    Dim pProjInfo   As Long
    Dim lIsNative   As Long
    Dim pNewHeader  As Long
    
    Debug.Assert MakeTrue(bIsinIDE)
    
    If Not bIsinIDE Then
        Set cExpSrv = CreateIExprSrvObj(0, 4, 0)
    End If
    
    CoInitialize ByVal 0&
    
    hHeap = GetProcessHeap()
    
    TlsSetValue lTlsSlot, ByVal pParameter
    
    GetMem8 ByVal pParameter, tThreadData

    If bIsinIDE Then
        FakeMain
    Else
        
        pNewHeader = CreateVBHeaderCopy()
        
        If pNewHeader Then
        
            tIID.c2 = 504403158265495.5712@

            VBDllGetClassObject hModule, 0, pNewHeader, tClsId, tIID, 0

            ' // Becasue of a header will be used by MSVBVM60 in DllMain (with DLL_THREAD_DETACH)
            ' // we can't free it now. To avoid memory leak we will free it later
            
        End If
        
    End If
    
    GetMem8 ByVal pParameter, tThreadData
    
    ThreadProc = tThreadData.lpParameter
    
    HeapFree hHeap, 0, pParameter
    
    ' // Set state
    TlsSetValue lTlsSlot, ByVal 0&
    
    CoUninitialize
    
End Function

' // Free copy of header
Private Sub FreeHeaderCopy( _
            ByVal pHeader As Long)
    HeapFree GetProcessHeap(), 0, pHeader - 4
End Sub

' // Free unused headers
' // If other thread already is being cleaning return true
Private Function FreeUnusedHeaders() As Boolean
    Dim tEntry              As PROCESS_HEAP_ENTRY
    Dim lTID                As Long
    Dim lCurrentThreads()   As Long
    Dim pPointers()         As Long
    Dim lIndex              As Long
    Dim lCount              As Long

    ' // Try to get exclusive access
    If TryEnterCriticalSection(tLockHeap.tWinApiSection) = 0 Then
    
        FreeUnusedHeaders = True
        Exit Function
        
    End If
    
    If GetThreadsList(lCurrentThreads()) = 0 Then GoTo unlock_access
    
    ReDim pPointers(255)
    
    If HeapLock(hHeadersHeap) = 0 Then GoTo unlock_access
    
    Do While HeapWalk(hHeadersHeap, tEntry)

        If tEntry.wFlags And PROCESS_HEAP_ENTRY_BUSY Then
            
            ' // Check if header is unused
            GetMem4 ByVal tEntry.lpData, lTID

            If Not TIDIsInList(lTID, lCurrentThreads()) Then
                
                If lIndex > UBound(pPointers) Then
                    ReDim Preserve pPointers(lIndex + 256)
                End If
                
                pPointers(lIndex) = tEntry.lpData
                
                lIndex = lIndex + 1
                
            End If
            
        End If

    Loop

    HeapUnlock hHeadersHeap
    
    lCount = lIndex
    
    For lIndex = 0 To lCount - 1
        HeapFree hHeadersHeap, 0, pPointers(lIndex)
    Next
    
unlock_access:
    
    LeaveCriticalSection tLockHeap.tWinApiSection

End Function

' // Check if TID in the list
Private Function TIDIsInList( _
                 ByVal lTID As Long, _
                 ByRef lList() As Long) As Boolean
    Dim lIndex  As Long
    
    For lIndex = 0 To UBound(lList)
        
        If lList(lIndex) = lTID Then
            TIDIsInList = True
            Exit Function
        End If
        
    Next
    
End Function

' // Get threads list of current process
Private Function GetThreadsList( _
                 ByRef lTIDs() As Long) As Long
    Dim hSnap   As Long
    Dim tEntry  As THREADENTRY32
    Dim lIndex  As Long
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0)
    If hSnap = -1 Then Exit Function
    
    ReDim lTIDs(255)
    
    tEntry.dwSize = Len(tEntry)
    
    If Thread32First(hSnap, tEntry) Then
        
        Do
            
            If lIndex > UBound(lTIDs) Then
                ReDim Preserve lTIDs(lIndex + 256)
            End If
            
            If tEntry.th32OwnerProcessID = GetCurrentProcessId() Then
            
                lTIDs(lIndex) = tEntry.th32ThreadID
                lIndex = lIndex + 1
                
            End If
            
        Loop While Thread32Next(hSnap, tEntry)
    
    End If
    
    CloseHandle hSnap
    
    If lIndex Then
        ReDim Preserve lTIDs(lIndex - 1)
    End If
    
    GetThreadsList = lIndex
    
End Function

' // Create copy of VBHeader and other structures
' // The first four bytes contain thread ID. We use that ID to clean unused headers
Private Function CreateVBHeaderCopy() As Long
    Dim pHeader         As Long
    Dim pOldProjInfo    As Long
    Dim pProjInfo       As Long
    Dim pObjTable       As Long
    Dim pOldObjTable    As Long
    Dim lDifference     As Long
    Dim lIndex          As Long
    Dim pNames(3)       As Long
    Dim lModulesCount   As Long
    Dim pDescriptors    As Long
    Dim pOldDesc        As Long
    Dim pVarBlock       As Long
    Dim lSizeOfHeaders  As Long
    
    ' // Get size of all headers
    lSizeOfHeaders = &H6A + &H23C + &H54 + &HC + 4
    
    GetMem4 ByVal pVBHeader + &H30, pOldProjInfo
    GetMem4 ByVal pOldProjInfo + &H4, pOldObjTable
    GetMem4 ByVal pOldObjTable + &H30, pOldDesc
    GetMem2 ByVal pOldObjTable + &H2A, lModulesCount
    
    lSizeOfHeaders = lSizeOfHeaders + &H30 * lModulesCount
    
    ' // Allocate memory for header
    Do
        
        pHeader = HeapAlloc(hHeadersHeap, HEAP_ZERO_MEMORY, lSizeOfHeaders)
        
        ' // If there is not enough memory - free unused headers
        If pHeader = 0 Then
            
            If FreeUnusedHeaders() Then
                ' // Other threead frees memory
                Sleep 100
            Else
                
                pHeader = HeapAlloc(hHeadersHeap, HEAP_ZERO_MEMORY, lSizeOfHeaders)
                If pHeader = 0 Then GoTo CleanUp
                
                Exit Do
                
            End If
        Else
            Exit Do
        End If
        
    Loop
    
    pHeader = pHeader + 4
    
    lDifference = pHeader - pVBHeader
    
    CopyMemory ByVal pHeader, ByVal pVBHeader, &H6A
    
    ' // Update strings offsets
    CopyMemory pNames(0), ByVal pVBHeader + &H58, &H10
    
    For lIndex = 0 To 3
        pNames(lIndex) = pNames(lIndex) - lDifference
    Next
        
    CopyMemory ByVal pHeader + &H58, pNames(0), &H10

    ' // In order to keep global variables
    ' // Change the VbPublicObjectDescriptor.lpPublicBytes, VbPublicObjectDescriptor.lpStaticBytes
    pProjInfo = pHeader + &H6A

    CopyMemory ByVal pProjInfo, ByVal pOldProjInfo, &H23C

    ' // Update on VBHeader
    GetMem4 pProjInfo, ByVal pHeader + &H30

    ' // Create copy of Object table
    pObjTable = pProjInfo + &H23C

    CopyMemory ByVal pObjTable, ByVal pOldObjTable, &H54

    ' // Update
    GetMem4 pObjTable, ByVal pProjInfo + &H4

    ' // Allocate descriptors
    pDescriptors = pObjTable + &H54

    CopyMemory ByVal pDescriptors, ByVal pOldDesc, lModulesCount * &H30

    ' // Update
    GetMem4 pDescriptors, ByVal pObjTable + &H30

    ' // Allocate variables block
    pVarBlock = pDescriptors + lModulesCount * &H30

    For lIndex = 0 To lModulesCount - 1

        ' // Zero number of public and local variables
        GetMem4 pVarBlock, ByVal pDescriptors + lIndex * &H30 + &H8
        GetMem4 0&, ByVal pDescriptors + lIndex * &H30 + &HC

    Next

    CreateVBHeaderCopy = pHeader
    
CleanUp:

End Function

' // Callback
Private Sub FakeMain()
    Dim hr          As Long
    Dim pData       As Long
    Dim tThreadData As tThreadData
    Dim vRet        As Variant
    Dim bIsinIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsinIDE)

    pData = TlsGetValue(lTlsSlot)
    
    GetMem8 ByVal pData, tThreadData
    
    hr = DispCallFunc(ByVal 0&, tThreadData.lpAddress, CC_STDCALL, vbLong, 1, vbLong, _
                      VarPtr(CVar(tThreadData.lpParameter)), vRet)
    
    ' // Thread returned value
    tThreadData.lpParameter = vRet
    
    GetMem8 tThreadData, ByVal pData

End Sub

' // Get VBHeader structure
Private Function GetVBHeader() As Long
    Dim ptr     As Long
   
    ' // Get e_lfanew
    GetMem4 ByVal hModule + &H3C, ptr
    ' // Get AddressOfEntryPoint
    GetMem4 ByVal ptr + &H28 + hModule, ptr
    ' // Get VBHeader
    GetMem4 ByVal ptr + hModule + 1, GetVBHeader
    
End Function

' // Modify VBHeader to replace Sub Main
Private Sub ModifyVBHeader( _
            ByVal pNewAddress As Long)
    Dim ptr             As Long
    Dim lOldProtect     As Long
    Dim lFlags          As Long
    Dim lFormsCount     As Long
    Dim lModulesCount   As Long
    Dim lStructSize     As Long
    
    ptr = pVBHeader + &H2C
    ' // Allow to write to that page
    VirtualProtect ByVal ptr, 4, PAGE_READWRITE, lOldProtect
    
    ' // Set new Sub Main
    GetMem4 pNewAddress, ByVal ptr
    VirtualProtect ByVal ptr, 4, lOldProtect, 0
    
    ' // Remove startup form
    GetMem4 ByVal pVBHeader + &H4C, ptr
    ' // Get number of forms
    GetMem2 ByVal pVBHeader + &H44, lFormsCount
    
    Do While lFormsCount > 0
    
        ' // Get structure size
        GetMem4 ByVal ptr, lStructSize
        
        ' // Get flag (unknown5) from current form
        GetMem4 ByVal ptr + &H28, lFlags
        
        ' // When set, bit 5,
        If lFlags And &H10 Then
        
            ' // Unset bit 5
            lFlags = lFlags And &HFFFFFFEF
            ' // Are allowed to write in the page
            VirtualProtect ByVal ptr, 4, PAGE_READWRITE, lOldProtect
            ' // Write changet lFlags
            GetMem4 lFlags, ByVal ptr + &H28
            ' // Restoring the memory attributes
            VirtualProtect ByVal ptr, 4, lOldProtect, 0
            
        End If
        
        lFormsCount = lFormsCount - 1
        ptr = ptr + lStructSize
        
    Loop

End Sub

Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True: bValue = True
End Function
