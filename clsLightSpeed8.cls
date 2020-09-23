VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLightSpeed8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'Email: jason_bullen@yahoo.com
'Copyright Jason Bullen November 2003. All right reserved.
'This source code is copyrighted material which may not be published
' in any form without explicit prior permission from the author.


' Set this to TRUE to check initialization and function parameters
#Const DEBUGGING = False


' Remember initialize picturebox
Private mPicBoxDC As Long
' Handles
Private mDC As Long
Private mBitmap As Long
' The format descriptor
Private mBmpInfo As tBitmapInfo
' Frequently used
Private mBufferWidth As Integer
Private mBufferHeight As Integer
Private mBufferPitch As Long
Private mBufferSize As Long
' The RGB data
Private mBitsPointer As Long
' The palette
Private mPalette(0 To 255) As tRgbQuad



' WIN32
'-------------------------------------------------------------------------------------------------
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As tBitmapInfo, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal first As Long, ByVal count As Long, pRGBQuad As tRgbQuad) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal first As Long, ByVal count As Long, pcRGBQuad As tRgbQuad) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Const BI_RGB As Long = 0
Private Const DIB_RGB_COLORS As Long = 0
Private Const DIB_PAL_COLORS As Long = 1

Private Type tBitmapInfoHeader
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type tRgbQuad
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type tBitmapInfo
        bmiHeader As tBitmapInfoHeader
        bmiColors As tRgbQuad
End Type

Private Const NoExtraCheck As Integer = 1


'-------------------------------------------------------------------------------------------------
' INITIALIZERS
'-------------------------------------------------------------------------------------------------
Public Function InitPicture(picBox As PictureBox, grabSource As Boolean) As Boolean
#If DEBUGGING Then
    If mDC Then
        Call Err.Raise(1, "clsLightSpeed8", "CLASS ALREADY INITIALIZED!!!")
    End If
#End If
    mPicBoxDC = picBox.hdc
    mBufferWidth = (picBox.ScaleWidth + 3) And &HFFFFFFFC    'This rounds UP to the next 4 pixels
    mBufferPitch = mBufferWidth
    mBufferHeight = picBox.ScaleHeight
    With mBmpInfo.bmiHeader
        .biSize = Len(mBmpInfo.bmiHeader)
        .biWidth = mBufferWidth
        .biHeight = -mBufferHeight      '-' makes the picture TOP-DOWN in memory
        .biPlanes = 1
        .biBitCount = 8
        .biCompression = BI_RGB
        .biSizeImage = mBufferPitch * -.biHeight
    End With

    mBufferSize = CLng(mBufferWidth) * CLng(mBufferHeight)

    mDC = CreateCompatibleDC(mPicBoxDC)
    If mDC Then
        mBitmap = CreateDIBSection(mDC, mBmpInfo, DIB_RGB_COLORS, VarPtr(mBitsPointer), ByVal 0&, ByVal 0&)
        If mBitmap Then
            Call SelectObject(mDC, mBitmap)
            If grabSource Then
                Call BitBlt(mDC, 0, 0, mBufferWidth, mBufferHeight, mPicBoxDC, 0, 0, vbSrcCopy)
            End If
            InitPicture = False ' no error
            Exit Function
        End If
        DeleteDC mDC
    End If
    mDC = 0
    mBitmap = 0
    mPicBoxDC = 0
    InitPicture = True
#If DEBUGGING Then
    Call Err.Raise(1, "clsLightSpeed8", "InitPicture FAILED")
#End If
End Function


Public Function InitDimensions(ByVal width As Integer, ByVal height As Integer) As Boolean
#If DEBUGGING Then
    If mDC Then
        Call Err.Raise(1, "clsLightSpeed8", "CLASS ALREADY INITIALIZED!!!")
    End If
#End If
    mPicBoxDC = 0
    mBufferWidth = (width + 3) And &HFFFFFFFC    'This rounds UP to the next 4 pixels
    mBufferPitch = mBufferWidth
    mBufferHeight = height
    With mBmpInfo.bmiHeader
        .biSize = Len(mBmpInfo.bmiHeader)
        .biWidth = mBufferWidth
        .biHeight = -mBufferHeight      '-' makes the picture TOP-DOWN in memory
        .biPlanes = 1
        .biBitCount = 8
        .biCompression = BI_RGB
    End With

    mBufferSize = CLng(mBufferWidth) * CLng(mBufferHeight)

    mDC = CreateCompatibleDC(mPicBoxDC)
    If mDC Then
        mBitmap = CreateDIBSection(mDC, mBmpInfo, DIB_RGB_COLORS, VarPtr(mBitsPointer), ByVal 0&, ByVal 0&)
        If mBitmap Then
            Call SelectObject(mDC, mBitmap)
            InitDimensions = False ' no error
            Exit Function
        End If
        DeleteDC mDC
    End If
    mDC = 0
    mBitmap = 0
    InitDimensions = True
#If DEBUGGING Then
    Call Err.Raise(1, "clsLS", "InitDimensions FAILED")
#End If
End Function


'-------------------------------------------------------------------------------------------------
'PROPERTIES
'-------------------------------------------------------------------------------------------------
Public Property Get GetPointer() As Long
    Call IsClassInitializedRaiseError(NoExtraCheck, "GetPointer")
    GetPointer = mBitsPointer
End Property

Public Property Get GetWidth() As Integer
    Call IsClassInitializedRaiseError(NoExtraCheck, "GetWidth")
    GetWidth = mBufferWidth
End Property

Public Property Get GetPitch() As Integer
    Call IsClassInitializedRaiseError(NoExtraCheck, "GetPitch")
    GetPitch = mBufferPitch
End Property

Public Property Get GetHeight() As Integer
    Call IsClassInitializedRaiseError(NoExtraCheck, "GetHeight")
    GetHeight = mBufferHeight
End Property

Public Property Get GetDC() As Long
    Call IsClassInitializedRaiseError(NoExtraCheck, "GetDC")
    GetDC = mDC
End Property

Public Property Set SetPictureBox(newValue As PictureBox)
    mPicBoxDC = newValue.hdc
End Property


'-------------------------------------------------------------------------------------------------
'PALETTE
'-------------------------------------------------------------------------------------------------
Public Sub ReadPaletteFile(fileName As String)
    Dim fileNum As Integer
    Dim readStr As String * 768
    Dim i As Integer, j As Integer, c As Byte
    
#If DEBUGGING Then
    If Dir(fileName) = "" Then
        Call Err.Raise(1, "clsLightSpeed8", "ReadPaletteFile FAILED")
    End If
#End If
    
    fileNum = FreeFile(0)
    Open fileName For Binary Access Read As #fileNum
    readStr = Input(768, fileNum)
    j = 1
    For i = 0 To 255
        c = Asc(Mid(readStr, j, 1))
        mPalette(i).rgbRed = c
        j = j + 1
        c = Asc(Mid(readStr, j, 1))
        mPalette(i).rgbGreen = c
        j = j + 1
        c = Asc(Mid(readStr, j, 1))
        mPalette(i).rgbBlue = c
        j = j + 1
    Next
    Close fileNum
End Sub

Public Sub SetPaletteColor(index As Integer, red As Integer, green As Integer, blue As Integer)
#If DEBUGGING Then
    If index < 0 Or index > 255 Or _
        red < 0 Or red > 255 Or _
        green < 0 Or green > 255 Or _
        blue < 0 Or blue > 255 Then
            Call Err.Raise(1, "clsLightSpeed8", "SetPaletteColor FAILED")
    End If
#End If
    mPalette(index).rgbRed = red
    mPalette(index).rgbGreen = green
    mPalette(index).rgbBlue = blue
End Sub

Public Sub SetPalette(first As Integer, count As Integer)
    Call IsClassInitializedRaiseError(NoExtraCheck, "SetPalette")
    Call SetDIBColorTable(mDC, first, count, mPalette(0))
End Sub


'-------------------------------------------------------------------------------------------------
' MODIFY ARRRAY
'-------------------------------------------------------------------------------------------------
' This function redirects a SINGLE dimension array to access our memory DIB
' ONLY single dimension arrays can be redirected by this function
' This is intentional because 2 and 3 dimension arrays are much slower
' It returns the original memory address which must be recorded and used
' in the call to FixArray() below
Public Function GetArray(ByRef theArray() As Byte) As Long
    Dim pointer As Long
    
    Call IsClassInitializedRaiseError(NoExtraCheck, "GetArray")
    ReDim theArray(0 To 0) As Byte
    ' Get pointer to safearray structure
    Call CopyMemory(ByVal VarPtr(pointer), ByVal VarPtrArray(theArray()), 4)
    ' Record old data pointer
    Call CopyMemory(ByVal VarPtr(GetArray), ByVal pointer + 12, 4)
    ' Offset to data pointer and change to point to memory bitmap
    Call CopyMemory(ByVal pointer + 12, ByVal VarPtr(mBitsPointer), 4)
    ' Offset to array size and change to match memory bitmap size
    Call CopyMemory(ByVal pointer + 16, ByVal VarPtr(mBufferSize), 4)
End Function

' This function fixes modified arrays by redirecting them to the original
' memory address.
' This MUST be done before the App closes or the array is ERASED!!
Public Sub FixArray(ByRef theArray() As Byte, oldPointer As Long)
    Dim pointer As Long, size As Long
    
    Call IsClassInitializedRaiseError(NoExtraCheck, "FixArray")
    Call CopyMemory(ByVal VarPtr(pointer), ByVal VarPtrArray(theArray()), 4)
    Call CopyMemory(ByVal pointer + 12, ByVal VarPtr(oldPointer), 4)
    size = 1
    Call CopyMemory(ByVal pointer + 16, ByVal VarPtr(size), 4)
End Sub


'-------------------------------------------------------------------------------------------------
' GRAB - Copy Image Data into Array
' Copy image data to our memory bitmap (DIB) from various sources
'-------------------------------------------------------------------------------------------------
Public Sub GrabPicture()
    Call IsClassInitializedRaiseError(mPicBoxDC, "GrabPicture")
    Call BitBlt(mDC, 0, 0, mBufferWidth, mBufferHeight, mPicBoxDC, 0, 0, vbSrcCopy)
End Sub

Public Sub GrabPictureArea(dstX As Integer, dstY As Integer, _
                           width As Integer, height As Integer, _
                           srcX As Integer, srcY As Integer)
    
    Call IsClassInitializedRaiseError(mPicBoxDC, "GrabPictureArea")
    Call BitBlt(mDC, dstX, dstY, width, height, mPicBoxDC, srcX, srcY, vbSrcCopy)
End Sub

Public Sub GrabOtherPicture(picBox As PictureBox)
    Dim width As Integer, height As Integer
    
    Call IsClassInitializedRaiseError(picBox.hdc, "GrabOtherPicture")
    width = picBox.ScaleX(picBox.ScaleWidth, picBox.ScaleMode, vbPixels)
    height = picBox.ScaleY(picBox.ScaleHeight, picBox.ScaleMode, vbPixels)
    Call BitBlt(mDC, 0, 0, width, height, picBox.hdc, 0, 0, vbSrcCopy)
End Sub

Public Sub GrabOtherPictureArea(sourceDC As Long, _
                                dstX As Integer, dstY As Integer, _
                                width As Integer, height As Integer, _
                                srcX As Integer, srcY As Integer)
    
    Call IsClassInitializedRaiseError(sourceDC, "GrabOtherPictureArea")
    Call BitBlt(mDC, dstX, dstY, width, height, sourceDC, srcX, srcY, vbSrcCopy)
End Sub


'-------------------------------------------------------------------------------------------------
' PUT BITS
' Copy image data from our memory bitmap (DIB) to various sources
'-------------------------------------------------------------------------------------------------
Public Sub PutPicture()
    Call IsClassInitializedRaiseError(mPicBoxDC, "PutPicture")
    Call BitBlt(mPicBoxDC, 0, 0, mBufferWidth, mBufferHeight, mDC, 0, 0, vbSrcCopy)
End Sub

Public Sub PutPictureArea(dstX As Integer, dstY As Integer, _
                          width As Integer, height As Integer, _
                          srcX As Integer, srcY As Integer)
    
    Call IsClassInitializedRaiseError(mPicBoxDC, "PutPictureArea")
    Call BitBlt(mPicBoxDC, dstX, dstY, width, height, mDC, srcX, srcY, vbSrcCopy)
End Sub

Public Sub PutOtherPicture(picBox As PictureBox)
    Call IsClassInitializedRaiseError(picBox.hdc, "PutOtherPicture")
    Call BitBlt(picBox.hdc, 0, 0, mBufferWidth, mBufferHeight, mDC, 0, 0, vbSrcCopy)
End Sub

Public Sub PutOtherPictureArea(picDC As Long, _
                               dstX As Integer, dstY As Integer, _
                               width As Integer, height As Integer, _
                               srcX As Integer, srcY As Integer)
    Call IsClassInitializedRaiseError(picDC, "PutOtherPictureArea")
    Call BitBlt(picDC, dstX, dstY, width, height, mDC, srcX, srcY, vbSrcCopy)
End Sub


'-------------------------------------------------------------------------------------------------
' TERMINATE
'-------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Call Destroy
End Sub


'-------------------------------------------------------------------------------------------------
' Destroy Objects and handles
'-------------------------------------------------------------------------------------------------
Public Sub Destroy()
    If mDC Then
        Call DeleteDC(mDC)
        mDC = 0
    End If
    If mBitmap Then
        Call DeleteObject(mBitmap)
        mBitmap = 0
    End If
    mPicBoxDC = 0
End Sub



'-------------------------------------------------------------------------------------------------
' CLASS INITIALIZED CHECKS
'-------------------------------------------------------------------------------------------------
Private Sub IsClassInitializedRaiseError(extraCheck As Long, text As String)
#If DEBUGGING Then
    If mDC = 0 Or extraCheck = 0 Then
        Call Err.Raise(1, "clsLightSpeed8", "clsLightSpeed8 Not Initialized!" & vbCrLf & "Function: " & text)
    End If
#End If
End Sub


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Direct Image Manipulation Routines - DLL Interface
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'SHARED PARAMETERS
'-------------------------------------------------------------------------------------------------
'       source                Source LightSpeed class
'       dstBits               Raw image data
'       dstPitch              Distance to next vertical pixel (width+padding)
'       dstX, dstY            Top left destination
'       dstCenX, dstCenY      Offset to destination center (rotation pivot)
'       dstWidth, dstHeight   Size of destination rectangle
'       srcBits               Raw image data
'       srcPitch              Distance to next vertical pixel (width+padding)
'       srcX, srcY            Top left source
'       srcCenX, srcCenY      Offset to source center (rotation pivot)
'       srcWidth, srcHeight   Size of source rectangle
'       angle                 Rotation angle clockwise in Radians
'       zoom                  1.0=normal, 2.0=double scale, 0.5=half scale, etc.
'       colorKey              Ignore the colorKey color (always palette index 0 - for speed)
'       width, height         Size of rectangle


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'FILL the AREA with color index ZERO
Public Sub FillZero()
    Call IsClassInitializedRaiseError(mBitsPointer, "FillZero")
    Call ZeroMemory(ByVal mBitsPointer, mBufferSize)
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'FILL the AREA with the color
Public Sub FillColor(color As Byte)
    Call IsClassInitializedRaiseError(mBitsPointer, "FillColor")
    Call FillMemory(ByVal mBitsPointer, mBufferSize, color)
End Sub


