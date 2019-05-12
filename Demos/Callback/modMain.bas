Attribute VB_Name = "modMain"
' //
' // Callback procedures
' //

Option Explicit

Private Type tCallbackParams
    lThreadId           As Long
    sKey                As String
    fTimeFromLastTick   As Single
End Type

Public Declare Function SetCallback Lib "dll\CallbackDll" ( _
                        ByVal pfn As Long, _
                        ByVal lInterval As Long, _
                        ByRef pszKey As Long) As Long
Public Declare Sub StopCallback Lib "dll\CallbackDll" ( _
                   ByVal bId As Byte)
Public Declare Sub FreeAll Lib "dll\CallbackDll" ()

' // Marshal stream. A threads uses that data to unmarshal frmMain object in different threads
Public gpMarshalData    As Long

' // This function is used in compiled form
Public Function CallbackProc( _
                ByVal lThreadId As Long, _
                ByVal sKey As String, _
                ByVal fTimeFromLastTick As Single) As Long
    ' // Init runtime and call CallBackProc_user with VarPtr(lThreadId) parameter
    InitCurrentThreadAndCallFunction AddressOf CallBackProc_user, VarPtr(lThreadId), CallbackProc
End Function

' // Callback function is called by runtime/window proc (in IDE)
Public Function CallBackProc_user( _
                ByRef tParam As tCallbackParams) As Long
    Dim cObj    As Object
    
    ' // Get unmarshaled pointer of frmMain
    Set cObj = UnMarshal(gpMarshalData, , False)
    
    ' // Log parameters
    cObj.Log "Callback TID [0x" & Hex$(App.ThreadID) & "] Params: 0x" & Hex$(tParam.lThreadId) & "; '" & _
                                        tParam.sKey & "'; " & Format$(tParam.fTimeFromLastTick, "0.0000")
    
End Function

