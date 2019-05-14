VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContextHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CC_STDCALL As Long = 4


Private Declare Function LoadLibrary Lib "kernel32" _
                         Alias "LoadLibraryW" ( _
                         ByVal lpLibFileName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
                         ByVal hModule As Long, _
                         ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
                         ByVal hLibModule As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function lstrcmpA Lib "kernel32" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any) As Long
Private Declare Function memcpy Lib "kernel32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal Length As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
                         ByVal pvInstance As IUnknown, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long
                         
Private m_cFactory  As IUnknown
Private m_cContext  As Object
Private m_hLibrary  As Long

' // Call exported function from DLL
Public Function CallFunc( _
                ByRef sFuncName As String, _
                ByVal eRetType As VbVarType, _
                ParamArray vParams() As Variant) As Variant
    Dim iTypes()    As Integer
    Dim lList()     As Long
    Dim vParam()    As Variant
    Dim lIndex      As Long
    Dim hr          As Long
    Dim hLib        As Long
    Dim pfn         As Long
    Dim pList       As Long
    Dim pTypes      As Long
    
    If m_hLibrary = 0 Then Err.Raise 5
    
    pfn = GetProcAddress(m_hLibrary, sFuncName)
    If pfn = 0 Then Err.Raise 5
 
    If LBound(vParams) <= UBound(vParams) Then
        
        ReDim lList(UBound(vParams))
        ReDim iTypes(UBound(vParams))
        ReDim vParam(UBound(vParams))
        
        For lIndex = 0 To UBound(vParams)
        
            vParam(lIndex) = vParams(lIndex)
            lList(lIndex) = VarPtr(vParams(lIndex))
            iTypes(lIndex) = VarType(vParams(lIndex))
            
        Next
        
        pList = VarPtr(lList(0))
        pTypes = VarPtr(iTypes(0))
        
    End If
 
    hr = DispCallFunc(Nothing, pfn, CC_STDCALL, eRetType, UBound(vParams) - LBound(vParams) + 1, ByVal pTypes, ByVal pList, CallFunc)

    If hr Then Err.Raise 5: Exit Function
    
End Function

' // Init project context for dll
' // DLL should has the dummy class called "CInitContext"
Public Function InitVbProjContext( _
                ByRef sLibName As String) As Boolean
    Dim pVBHeader       As Long
    Dim pCOMData        As Long
    Dim lOfstRegInfo    As Long
    Dim pRegInfo        As Long
    Dim pfn             As Long
    Dim hLib            As Long
    Dim pClassName      As Long
    Dim iTypes(2)       As Integer
    Dim vParams(2)      As Variant
    Dim lList(2)        As Long
    Dim bIID(1)         As Currency
    Dim cFactory        As IUnknown
    Dim cObject         As IUnknown
    Dim hr              As Long
    Dim vRet            As Variant
    
    Free
    
    hLib = LoadLibrary(StrPtr(sLibName))
    If hLib = 0 Then Exit Function
    
    pfn = GetProcAddress(hLib, "DllGetClassObject")
    If pfn = 0 Then GoTo CleanUp
    
    ' // Get VBHeader
    GetMem4 ByVal pfn + 2, pVBHeader
    ' // Get COM data
    GetMem4 ByVal pVBHeader + &H54, pCOMData
    ' // Get Reg info
    GetMem4 ByVal pCOMData, lOfstRegInfo
    
    pRegInfo = pCOMData + lOfstRegInfo
    
    ' // Search for CInitContext CLSID
    
    Do
        
        GetMem4 ByVal pRegInfo + 4, pClassName
        
        pClassName = pCOMData + pClassName
        
        ' // Check the class name
        If lstrcmpA(ByVal pClassName, ByVal "CInitContext") = 0 Then
            
            ' // Setup DllGetClassObject
            iTypes(0) = vbLong
            iTypes(1) = vbLong
            iTypes(2) = vbLong
            
            ' // IClassFactory
            bIID(0) = 0.0001@
            bIID(1) = 504403158265495.5712@
            
            vParams(0) = pRegInfo + &H14
            vParams(1) = VarPtr(bIID(0))
            vParams(2) = VarPtr(cFactory)
            
            lList(0) = VarPtr(vParams(0))
            lList(1) = VarPtr(vParams(1))
            lList(2) = VarPtr(vParams(2))
            
            hr = DispCallFunc(Nothing, pfn, CC_STDCALL, vbLong, 3, iTypes(0), lList(0), vRet)
            
            If hr >= 0 And vRet >= 0 Then
                
                ' // Call IClassFactory::CreateInstance
                
                ' // IUnknown
                bIID(0) = 0
                
                vParams(0) = 0&
                vParams(2) = VarPtr(cObject)
                
                hr = DispCallFunc(cFactory, &HC, CC_STDCALL, vbLong, 3, iTypes(0), lList(0), vRet)
                
                If hr >= 0 And vRet >= 0 Then
                    
                    Set m_cFactory = cFactory
                    Set m_cContext = cObject
                    m_hLibrary = hLib
                    InitVbProjContext = True
                    
                    Exit Do
                    
                End If
                
            End If
            
        End If
        
        GetMem4 ByVal pRegInfo, lOfstRegInfo

        pRegInfo = pCOMData + lOfstRegInfo
        
    Loop While lOfstRegInfo > 0
    
CleanUp:
    
    If Not InitVbProjContext Then
    
        If hLib Then
            FreeLibrary hLib
        End If
        
    End If
    
End Function

Private Sub Free()
    
    Set m_cContext = Nothing
    Set m_cFactory = Nothing
    
    If m_hLibrary Then
    
        FreeLibrary m_hLibrary
        m_hLibrary = 0
        
    End If
    
End Sub

Private Sub Class_Terminate()
    Free
End Sub
