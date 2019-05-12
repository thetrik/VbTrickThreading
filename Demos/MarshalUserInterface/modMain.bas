Attribute VB_Name = "modMain"
Option Explicit

Public Type UUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type

Public Declare Function GetMem8 Lib "msvbvm60" ( _
                        ByRef pSrc As Any, _
                        ByRef pDst As Any) As Long

' // {02FAF1A8-5F2E-4849-A8E3-E6B92BC7AE05}
Public Function IID_IUserInterface() As UUID
    GetMem8 520879909525382.3912@, IID_IUserInterface
    GetMem8 40948360675373.1496@, IID_IUserInterface.Data4(0)
End Function

' // {02FAF1A8-5F2E-4849-A8E3-E6B92BC7AE04}
Public Function IID_ILogObject() As UUID
    GetMem8 552843061798759.3661@, IID_ILogObject
    GetMem8 33742601271580.356@, IID_ILogObject.Data4(0)
End Function

