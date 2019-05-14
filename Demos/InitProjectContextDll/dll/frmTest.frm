VERSION 5.00
Begin VB.Form frmTest 
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCallbackToExe 
      Caption         =   "Callback"
      Height          =   915
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   2475
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CC_STDCALL As Long = 4

Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
                         ByVal pvInstance As IUnknown, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long
                         
Private m_pfn   As Long

Public Sub SetCallback( _
           ByVal pfn As Long)
    m_pfn = pfn
End Sub

Private Sub cmdCallbackToExe_Click()
    
    If m_pfn = 0 Then Exit Sub
    
    DispCallFunc Nothing, m_pfn, CC_STDCALL, vbEmpty, 0, 0, 0, 0
    
End Sub

Private Sub Form_Load()
    Caption = "From DLL; ThreadID: 0x" & Hex$(App.ThreadID)
End Sub
