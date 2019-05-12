VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Julia set by The trick"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbThreads 
      Height          =   255
      Left            =   1380
      Max             =   20
      Min             =   1
      TabIndex        =   6
      Top             =   9180
      Value           =   5
      Width           =   7815
   End
   Begin VB.Timer tmrRender 
      Interval        =   100
      Left            =   4800
      Top             =   7740
   End
   Begin VB.HScrollBar hsbImaginaryPart 
      Height          =   255
      Left            =   1380
      Max             =   1000
      Min             =   -1000
      TabIndex        =   2
      Top             =   8820
      Width           =   7815
   End
   Begin VB.HScrollBar hsbRealPart 
      Height          =   255
      Left            =   1380
      Max             =   1000
      Min             =   -1000
      TabIndex        =   1
      Top             =   8460
      Width           =   7815
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00000000&
      Height          =   8235
      Left            =   120
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   0
      Top             =   120
      Width           =   9075
   End
   Begin VB.Label Label3 
      Caption         =   "Threads:"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   9120
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Imaginary part:"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   8760
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Real part:"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   8460
      Width           =   1155
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Julia Set multithreading example
' // By The Trick 2018
' // This example creates several threads to calculate part of Julia set
' //

Option Explicit

Private Declare Function ColorHLSToRGB Lib "SHLWAPI.DLL" ( _
                         ByVal wHue As Integer, _
                         ByVal wLuminance As Integer, _
                         ByVal wSaturation As Integer) As Long

Private Sub Form_Load()
    Dim lIndex  As Long
    
    ' // Initialize multithreading module
    modMultiThreading.Initialize
    
    ' // Initialize palette
    For lIndex = 0 To 255
        glPalette(lIndex) = ColorHLSToRGB(lIndex / 255 * 240, lIndex / 255 * 240, 240)
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' // Wait threads
    WaitForThreadsCompletion
    
    ' // Clean up
    modMultiThreading.Uninitialize
    
End Sub

Private Sub hsbImaginaryPart_Change()
    
    ' // Cancel drawing
    gbCancelFlag = True
    gfImaginary = hsbImaginaryPart.Value / 1000
    
    ' // Run threads
    GenerateJulia picResult, hsbThreads.Value
    
End Sub

Private Sub hsbImaginaryPart_Scroll()
    Dim bIsInIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    If Not bIsInIDE Then hsbImaginaryPart_Change
    
End Sub

Private Sub hsbRealPart_Change()
    
    ' // Cancel drawing
    gbCancelFlag = True
    gfReal = hsbRealPart.Value / 1000
    
    ' // Run threads
    GenerateJulia picResult, hsbThreads.Value
    
End Sub

Private Sub hsbRealPart_Scroll()
    Dim bIsInIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    If Not bIsInIDE Then hsbRealPart_Change
    
End Sub

Private Sub tmrRender_Timer()
    
    ' // Update picture periodically
    Render picResult
    
End Sub
