VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IUserInterface

Private mcLogObject As ILogObject
Private mvValue     As Variant

Private Sub Class_Terminate()
    Log "CUserClass::~CUserClass"
End Sub

Private Sub Log( _
            ByRef sText As String)

    If Not mcLogObject Is Nothing Then
        mcLogObject.Log sText & " {ThreadID: 0x" & Hex$(App.ThreadID) & "}"
    End If

End Sub

Private Sub IUserInterface_SetLogObject( _
            ByVal pObj As user_typlib.ILogObject)
            
    Log "CUserClass::SetLogObject"

    Set mcLogObject = pObj

End Sub

Private Sub IUserInterface_ShowMessage( _
            ByVal sMsg As String)
            
    Log "CUserClass::ShowMessage; sMsg: '" & sMsg & "'"

    MsgBox sMsg
    
End Sub

Private Property Get IUserInterface_ThreadID() As Long
    
    Log "CUserClass::ThreadID"
    
    IUserInterface_ThreadID = App.ThreadID
    
End Property

Private Property Let IUserInterface_Value( _
                     ByVal RHS As Variant)
                     
    Log "CUserClass::Value_put"

    If IsObject(RHS) Then
        Set mvValue = RHS
    Else
        mvValue = RHS
    End If
    
End Property

Private Property Get IUserInterface_Value() As Variant

    Log "CUserClass::Value_get"

    If IsObject(mvValue) Then
        Set IUserInterface_Value = mvValue
    Else
        IUserInterface_Value = mvValue
    End If
    
End Property
