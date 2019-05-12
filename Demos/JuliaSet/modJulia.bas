Attribute VB_Name = "modJulia"
' //
' // modJulia.bas - module with drawing functions
' //

Option Explicit

Private Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type

' // Complex number
Private Type tComplex
    fR              As Single
    fI              As Single
End Type

' // Thread data
Private Type tThreadData
    lX              As Long     ' // X position to draw
    lY              As Long     ' // Y
    lWidth          As Long     ' // Width of part to draw
    lHeight         As Long     ' // Height of part to draw
    lTotalWidth     As Single   ' // Total canvas width
End Type

Private Declare Function SetDIBitsToDevice Lib "gdi32" ( _
                         ByVal hdc As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal dx As Long, _
                         ByVal dy As Long, _
                         ByVal SrcX As Long, _
                         ByVal SrcY As Long, _
                         ByVal Scan As Long, _
                         ByVal NumScans As Long, _
                         ByRef Bits As Any, _
                         ByRef BitsInfo As BITMAPINFO, _
                         ByVal wUsage As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" ( _
                         ByVal nCount As Long, _
                         ByRef lpHandles As Long, _
                         ByVal bWaitAll As Long, _
                         ByVal dwMilliseconds As Long) As Long

Public gfReal           As Single       ' // C.r
Public gfImaginary      As Single       ' // C.i
Public glPixels()       As Long         ' // Array of pixels
Public gHandles()       As Long         ' // Array of threads handles
Public glThreadCount    As Long         ' // Number of threads
Public gbCancelFlag     As Boolean      ' // If true - cancel drawing
Public glPalette(255)   As Long         ' // Palette


Private mtThreadsData() As tThreadData  ' // Threads data

' // Update picture using pixels array
Public Sub Render( _
           ByVal cPic As PictureBox)
    Dim tBI As BITMAPINFO
    
    If glThreadCount = 0 Then Exit Sub

    With tBI.bmiHeader

    .biBitCount = 32
    .biHeight = cPic.ScaleY(cPic.ScaleHeight, cPic.ScaleMode, vbPixels)
    .biWidth = cPic.ScaleX(cPic.ScaleWidth, cPic.ScaleMode, vbPixels)
    .biPlanes = 1
    .biSize = Len(tBI.bmiHeader)
    .biSizeImage = .biWidth * .biHeight * 4
    
    End With
    
    ' // Draw
    SetDIBitsToDevice cPic.hdc, 0, 0, tBI.bmiHeader.biWidth, tBI.bmiHeader.biHeight, 0, 0, 0, _
                      tBI.bmiHeader.biHeight, glPixels(0, 0), tBI, 0
    
End Sub

' // Generate Julia Set to glPixels() array
' // lThreadCount - number of threads to draw
' // The function divides whole canvas into areas and creates threads which draw these areas
Public Sub GenerateJulia( _
           ByVal cPic As PictureBox, _
           ByVal lThreadCount As Long)
    Dim lIndex      As Long
    Dim lWidth      As Long
    Dim lHeight     As Long
    Dim lPart       As Long
    Dim lCurX       As Long
    Dim fCurX       As Single
    Dim bIsInIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    ' // Wait for threads completion
    WaitForThreadsCompletion

    ReDim mtThreadsData(lThreadCount - 1)
    ReDim gHandles(lThreadCount - 1)
    
    ' // Get sizes
    lWidth = cPic.ScaleX(cPic.ScaleWidth, cPic.ScaleMode, vbPixels)
    lHeight = cPic.ScaleY(cPic.ScaleHeight, cPic.ScaleMode, vbPixels)
    lPart = lWidth \ lThreadCount
    
    ReDim Preserve glPixels(lWidth - 1, lHeight - 1)
    
    glThreadCount = lThreadCount
    gbCancelFlag = False
    
    For lIndex = 0 To lThreadCount - 1
        
        ' // Set parameters
        With mtThreadsData(lIndex)
        
        .lX = lCurX
        .lY = 0
        .lTotalWidth = lWidth
        
        ' // Remaining part for last thread
        If lIndex = lThreadCount - 1 Then
            .lWidth = lWidth - lCurX
        Else
            .lWidth = lPart
        End If
        
        .lHeight = lHeight
        
        End With
        
        lCurX = lCurX + lPart
        
        ' // Create thread
        gHandles(lIndex) = vbCreateThread(0, 0, AddressOf ThreadProc, VarPtr(mtThreadsData(lIndex)), 0, 0)
    
        If bIsInIDE Then DoEvents
        
    Next

End Sub

' // This function calculates the part of Julia Set
Private Function ThreadProc( _
                 ByRef tData As tThreadData) As Long
    Dim fX      As Single
    Dim fY      As Single
    Dim fStepX  As Single
    Dim fStepY  As Single
    Dim lWidth  As Long
    Dim lY      As Long
    Dim lX      As Long
    Dim bIsInIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    ' // Calculate scaled area (-1,-1)-(1, 1)
    fX = tData.lX / tData.lTotalWidth * 2 - 1
    fY = -1
    
    fStepX = 2 / tData.lTotalWidth
    fStepY = 2 / tData.lHeight
    
    For lY = 0 To tData.lHeight - 1
        
        ' // Cancelation
        If gbCancelFlag Then Exit Function
        If bIsInIDE Then DoEvents
        
        For lX = 0 To tData.lWidth - 1
            
            glPixels(lX + tData.lX, lY + tData.lY) = glPalette(Julia(fX, fY))
            fX = fX + fStepX
            
        Next
        
        fX = tData.lX / tData.lTotalWidth * 2 - 1
        fY = fY + fStepY
        
    Next
    
End Function

' // Wait threads
Public Sub WaitForThreadsCompletion()
    
    If glThreadCount = 0 Then Exit Sub
    
    ' // No error checking
    WaitForMultipleObjects glThreadCount, gHandles(0), 1, -1
    
    Do While glThreadCount
        
        glThreadCount = glThreadCount - 1
        CloseHandle gHandles(glThreadCount)
        gHandles(glThreadCount) = 0
        
    Loop
    
End Sub

' // Calculate julia for pixel Z^2 + C
Private Function Julia( _
                 ByVal fX As Single, _
                 ByVal fY As Single) As Long
    Dim tZ          As tComplex
    Dim tC          As tComplex
    Dim fTmp        As Single
    Dim lCounter    As Long
    Dim fRadius     As Single

    tZ.fR = fX:     tZ.fI = fY
    tC.fR = gfReal: tC.fI = gfImaginary

    Do While lCounter < 255 And fRadius < 100
        
        fTmp = tZ.fR
        
        tZ.fR = tZ.fR * tZ.fR - tZ.fI * tZ.fI
        tZ.fI = fTmp * tZ.fI + tZ.fI * fTmp
        
        tZ.fR = tZ.fR + tC.fR
        tZ.fI = tZ.fI + tC.fI
        
        fRadius = tZ.fR * tZ.fR + tZ.fI * tZ.fI

        lCounter = lCounter + 1
        
    Loop
    
    Julia = lCounter
    
End Function

Public Function MakeTrue( _
                ByRef bValue As Boolean) As Boolean
    MakeTrue = True: bValue = True
End Function
