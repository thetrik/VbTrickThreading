VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICallbackEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' // This method is called when downloading process has been completed
Public Sub Complete( _
           ByVal cObj As CAsynchDownloader)
End Sub

' // This mmethod is called when an error occurs
Public Sub Error( _
           ByVal cObj As CAsynchDownloader, _
           ByVal lError As Long)
End Sub

