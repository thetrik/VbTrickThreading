VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAsynchQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // Queue of dowloaders
' //
Option Explicit

Private m_cQueue    As Collection

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cQueue.[_NewEnum]
End Property

Public Property Get Item( _
                    ByVal vKeyIndex As Variant) As CAsynchDownloader
Attribute Item.VB_UserMemId = 0
    Set Item = m_cQueue(vKeyIndex)
End Property

' // Start downloading
Public Function Add( _
                ByRef sUrl As String, _
                Optional ByVal cCallback As ICallbackEvents, _
                Optional ByVal cLogWindow As TextBox) As CAsynchDownloader
    Dim cObj    As CAsynchDownloader
    
    Set cObj = New CAsynchDownloader
    
    Set cObj.LogWindow = cLogWindow
    Set cObj.Callback = cCallback
    
    cObj.StartDownloading sUrl
    
    m_cQueue.Add cObj
    
    Set Add = cObj
    
End Function

Private Sub Class_Initialize()
    Set m_cQueue = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_cQueue = Nothing
End Sub
