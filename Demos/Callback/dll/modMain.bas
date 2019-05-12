Attribute VB_Name = "modMain"
' //
' // CallbackDll - the dll that periodically calls a callback function in different thread
' //

Option Explicit

' // Item that represents a callback
Private Type tCallbackItem
    pfn         As Long     ' // Pointer to user callback function
    lInterval   As Long     ' // Interval in MS
    hThread     As Long     ' // Thread handle of periodic proc
    bStopFlag   As Boolean  ' // If true - end thread
    sKey        As String   ' // sKey - just data
    fTimer      As Single   ' // Last timer value
End Type

' // Parameters for DispCallFunc
Private Type tCallbackParams
    vArgs(2)    As Variant
    iTypes(2)   As Integer
    pArgs(2)    As Long
End Type

' // Callbacks list
Private Type tList
    tCallSlots(255) As tCallbackItem
    bLocked         As Boolean  ' // If true - forbid change list
End Type

Private tList       As tList
Private tLockList   As CRITICAL_SECTION

Sub Main()

End Sub

Private Function DllMain( _
                 ByVal hInstDll As Long, _
                 ByVal fdwReason As Long, _
                 ByVal lpvReserved As Long) As Long
    
    If fdwReason = DLL_PROCESS_ATTACH Then
        
        InitializeCriticalSection tLockList
        
    ElseIf fdwReason = DLL_PROCESS_DETACH Then
        
        FreeAll
        DeleteCriticalSection tLockList
        
    End If
    
    DllMain = 1
    
End Function

' // Set user callback procedure
Private Function SetCallback( _
                 ByVal pfn As Long, _
                 ByVal lInterval As Long, _
                 ByRef sKey As String) As Long
    Dim lIndex      As Long
    Dim lFreeItem   As Long
    
    lFreeItem = -1
    
    ' // Exclusive access
    LockList
    
    If tList.bLocked Then
        UnlockList
        Exit Function
    End If
    
    ' // Search for free slot
    For lIndex = 0 To 255
        
        If tList.tCallSlots(lIndex).hThread = 0 Then
            lFreeItem = lIndex
            Exit For
        End If
        
    Next
                     
    If lFreeItem >= 0 Then
        
        With tList.tCallSlots(lFreeItem)
        
        ' // Create thread
        .bStopFlag = False
        .lInterval = lInterval
        .sKey = sKey
        .pfn = pfn
        .hThread = CreateThread(ByVal 0&, 0, AddressOf PriodicProc, tList.tCallSlots(lFreeItem), 0, 0)
        
        End With
        
    End If
        
    SetCallback = lFreeItem
    
    UnlockList
    
End Function

Private Sub StopCallback( _
            ByVal bId As Byte)
    Dim bValid As Boolean
    
    LockList
    
    If tList.bLocked Then
        UnlockList
        Exit Sub
    End If

    If tList.tCallSlots(bId).hThread Then
        
        ' // Stop periodic proc
        tList.tCallSlots(bId).bStopFlag = True
        bValid = True
        
    End If
    
    UnlockList
    
    If bValid Then
        
        ' // Wait for periodic proc completion
        WaitCallbackCompletion tList.tCallSlots(bId)
        ClearSlot bId
        
    End If
    
End Sub

' // Priodic proc. This function periodically calls CallProc procedure in different threads
Private Function PriodicProc( _
                 ByRef tItem As tCallbackItem) As Long
    Dim hThread As Long
    
    Do
        tItem.fTimer = Timer
        
        Sleep tItem.lInterval

        hThread = CreateThread(ByVal 0&, 0, AddressOf CallProc, tItem, 0, 0)
        WaitForSingleObject hThread, -1
        CloseHandle hThread
    
    Loop Until tItem.bStopFlag
    
End Function

' // This procedure calls user callback function
Private Function CallProc( _
                 ByRef tItem As tCallbackItem) As Long
    Dim tArgs   As tCallbackParams
    Dim vRet    As Variant
    
    tArgs.iTypes(0) = vbLong
    tArgs.iTypes(1) = vbString
    tArgs.iTypes(2) = vbSingle
    
    tArgs.vArgs(0) = GetCurrentThreadId()
    tArgs.vArgs(1) = tItem.sKey
    tArgs.vArgs(2) = Timer - tItem.fTimer
    
    tArgs.pArgs(0) = VarPtr(tArgs.vArgs(0))
    tArgs.pArgs(1) = VarPtr(tArgs.vArgs(1))
    tArgs.pArgs(2) = VarPtr(tArgs.vArgs(2))
    
    DispCallFunc ByVal 0&, tItem.pfn, 4, vbEmpty, 3, tArgs.iTypes(0), tArgs.pArgs(0), vRet
        
End Function

' // End up all callback threads
Private Sub FreeAll()
    Dim lIndex  As Long
    
    LockList
    tList.bLocked = True
    UnlockList
    
    For lIndex = 0 To 255
        
        If tList.tCallSlots(lIndex).hThread Then
        
            WaitCallbackCompletion tList.tCallSlots(lIndex)
            ClearSlot lIndex
            
        End If
        
    Next
    
    LockList
    tList.bLocked = False
    UnlockList
    
End Sub

' // Wait
Private Sub WaitCallbackCompletion( _
            ByRef tItem As tCallbackItem)
    Dim lWaitState  As Long
    Dim tMsg        As MSG
    
    tItem.bStopFlag = True
    
    Do
        
        ' // Because of user function makes call to main thread (because marshaling)
        ' // We need to process windows messages
        
        lWaitState = MsgWaitForMultipleObjects(1, tItem.hThread, 0, -1, &H5FF)
    
        If lWaitState = 1 Then
        
            PeekMessage tMsg, 0, 0, 0, 1
            TranslateMessage tMsg
            DispatchMessage tMsg
            
        Else
            Exit Do
        End If
    
    Loop
            
End Sub

Private Sub ClearSlot( _
            ByVal lIndex As Long)
            
    LockList
    
    CloseHandle tList.tCallSlots(lIndex).hThread
    
    tList.tCallSlots(lIndex).hThread = 0
    tList.tCallSlots(lIndex).pfn = 0
    tList.tCallSlots(lIndex).sKey = vbNullString
    
    UnlockList
    
End Sub

Private Sub LockList()
    EnterCriticalSection tLockList
End Sub

Private Sub UnlockList()
    LeaveCriticalSection tLockList
End Sub

