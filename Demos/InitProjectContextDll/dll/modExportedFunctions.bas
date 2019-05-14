Attribute VB_Name = "modExportedFunctions"
Option Explicit

Public Function CreateForm() As Object
    Dim frm As frmTest
    
    Set frm = New frmTest
    
    frm.Show
    
    Set CreateForm = frm
    
End Function
