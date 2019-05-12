Attribute VB_Name = "modCopyProgress"
' //
' // Callback proc for CopyFileEx
' //

Option Explicit

Private Const PROGRESS_CONTINUE As Long = 0

Public Function CopyProgressRoutine( _
                ByVal TotalFileSize As Currency, _
                ByVal TotalBytesTransferred As Currency, _
                ByVal StreamSize As Currency, _
                ByVal StreamBytesTransferred As Currency, _
                ByVal dwStreamNumber As Long, _
                ByVal dwCallbackReason As Long, _
                ByVal hSourceFile As Long, _
                ByVal hDestinationFile As Long, _
                ByRef cNotify As CCopyProgress) As Long
    
    cNotify.ProgressRoutine TotalFileSize, TotalBytesTransferred
    CopyProgressRoutine = PROGRESS_CONTINUE
    
End Function

