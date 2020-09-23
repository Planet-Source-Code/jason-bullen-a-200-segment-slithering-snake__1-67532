Attribute VB_Name = "modRotate8"
Option Explicit
Option Base 0

'Email: jason_bullen@yahoo.com
'Copyright Jason Bullen November 2003. All right reserved.
'This source code is copyrighted material which may not be published in any form without explicit prior permission from the author.


Public Sub VB8_ScaleRotate( _
        ByRef dstBits() As Byte, ByVal dstPitch As Integer, _
        ByVal dstX As Integer, ByVal dstY As Integer, _
        ByVal dstCenX As Integer, ByVal dstCenY As Integer, _
        ByVal dstWidth As Integer, ByVal dstHeight As Integer, _
        ByRef srcBits() As Byte, ByVal srcPitch As Integer, _
        ByVal srcX As Integer, ByVal srcY As Integer, _
        ByVal srcCenX As Integer, ByVal srcCenY As Integer, _
        ByVal srcWidth As Integer, ByVal srcHeight As Integer, _
        ByVal angle As Single, ByVal zoom As Single, ByVal colorKey As Boolean)

    Dim RowAdd As Integer
    Dim x As Integer, y As Integer, irx As Long, iry As Long
    Dim xSrcMax As Integer, ySrcMax As Integer
    Dim cosa As Single, sina As Single
    Dim lcosa As Long, lsina As Long
    Dim issx As Long, issy As Long, iss2x As Long, iss2y As Long
    Dim dIndex As Long, color As Byte

    xSrcMax = srcX + srcWidth     ' Set Right and Bottom limits of the source image
    ySrcMax = srcY + srcHeight

    cosa = Cos(angle) / zoom    ' Get direction vector and scale it
    sina = Sin(angle) / zoom
    
    lcosa = cosa * 65536#       ' Convert the direction vector to Fixed Point 16.16 bits
    lsina = sina * 65536#
    
    iss2x = (srcX + srcCenX - dstCenX * cosa - dstCenY * sina) * 65536#    ' Find the rotated top left position in source
    iss2y = (srcY + srcCenY - dstCenY * cosa + dstCenX * sina) * 65536#

    dIndex = CLng(dstY) * CLng(dstPitch) + CLng(dstX)   ' Get the top left position in destination

    RowAdd = dstPitch - dstWidth    ' Get amount to add to destination to move down 1 line

    If colorKey Then
        For y = 0 To dstHeight - 1
            issx = iss2x    ' Set the 'X' scan start position
            issy = iss2y
            For x = 0 To dstWidth - 1
                irx = (issx + 32768) \ 65536              ' Get the rounded integer component of Source Scan X
                If irx >= srcX And irx < xSrcMax Then       ' Skip if outside source rectangle
                    
                    iry = (issy + 32768) \ 65536           ' Get the rounded integer component of Source Scan Y
                    If iry >= srcY And iry < ySrcMax Then   ' Skip if outside source rectangle
                        
                        color = srcBits(iry * srcPitch + irx)
                        If color Then
                            dstBits(dIndex) = color
                        End If
                    End If
                End If
                dIndex = dIndex + 1    ' Move one pixel to the right in destination
                issx = issx + lcosa        ' Add the direction vectors (scan X)
                issy = issy - lsina
            Next                  ' Loop X
            dIndex = dIndex + RowAdd  'Move to one pixel down and left edge of destination
            iss2x = iss2x + lsina         'Add direction vector minus 90 degress (scan Y)
            iss2y = iss2y + lcosa
        Next  ' Loop Y
    Else
        For y = 0 To dstHeight - 1
            issx = iss2x    ' Set the 'X' scan start position
            issy = iss2y
            For x = 0 To dstWidth - 1
                irx = (issx + 32768) \ 65536              ' Get the rounded integer component of Source Scan X
                If irx >= srcX And irx < xSrcMax Then       ' Skip if outside source rectangle
                    
                    iry = (issy + 32768) \ 65536           ' Get the rounded integer component of Source Scan Y
                    If iry >= srcY And iry < ySrcMax Then   ' Skip if outside source rectangle
                        
                        dstBits(dIndex) = srcBits(iry * srcPitch + irx)
                    End If
                End If
                dIndex = dIndex + 1    ' Move one pixel to the right in destination
                issx = issx + lcosa        ' Add the direction vectors (scan X)
                issy = issy - lsina
            Next                  ' Loop X
            dIndex = dIndex + RowAdd  'Move to one pixel down and left edge of destination
            iss2x = iss2x + lsina         'Add direction vector minus 90 degress (scan Y)
            iss2y = iss2y + lcosa
        Next  ' Loop Y
    End If
End Sub



