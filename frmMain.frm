VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serpent"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   498
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRandomMove 
      Caption         =   "Random Movement"
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   2400
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.HScrollBar hscNumSegments 
      Height          =   255
      LargeChange     =   10
      Left            =   7560
      Max             =   200
      Min             =   10
      SmallChange     =   2
      TabIndex        =   10
      Top             =   360
      Value           =   50
      Width           =   2295
   End
   Begin VB.HScrollBar hscShrink 
      Height          =   255
      LargeChange     =   10
      Left            =   7560
      Max             =   100
      TabIndex        =   8
      Top             =   1200
      Value           =   50
      Width           =   2295
   End
   Begin VB.CheckBox chkAngleSpeed 
      Caption         =   "Lock Turns"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.0"
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   600
      Left            =   9000
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   3840
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "START"
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   6840
      Width           =   2295
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   9480
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox picMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   120
      ScaleHeight     =   485
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   485
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Number of Segments"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tail Shrinkage"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblPix 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      TabIndex        =   5
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label lblCogs 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblFPS 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   6240
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'How it works
'mPositions holds an array containing the X and Y position and the rotation of the serpent segments.
'Each position is 1 (SNAKE_SPEED) pixel apart so for the first segment (33 pixels wide) we step
'33 positions through the positions array to find it's position.
'   This step value reduces as the snake segment shrinks in size.

'mHead and mTail are positions in the array. Both scan through the array backwards (decremented).
'   The mHead position is always decremented once after storing the position data each frame.

'The mTail position is different because the user can change the size and length of the snake.
'   mTail is calculated at startup and each time the user changes the length and shrinkage.
'   Then it works the same as mHead - Decremented once each frame.

'When drawing the snake we start at the TAIL and step forwards (increment) through the array to
'get the segment positions.


Const MAX_NUM_SEGMENTS As Integer = 201 'If you change this - change the form's ScrollBar
Const SEGMENT_SPACE As Integer = 33
Const NUM_POSITIONS As Integer = 6700   'Must be more than MAX_NUM_SEGMENTS * SEGMENT_SPACE
Const SNAKE_SPEED As Single = 1#
Const TURN_SPEED_DELTA As Single = 0.00005

'The LIGHT SPEED classes - one for each picturebox
Private mLightSpeed8(0 To 1) As clsLightSpeed8
' frequently used class values
Private mLightSpeedPitch0 As Integer
Private mLightSpeedPitch1 As Integer
'Arrays for direct image access
Private mArray0() As Byte
Private mArray1() As Byte
'To reset arrays to original state before erasing
Private mArray0Pnt As Long
Private mArray1Pnt As Long

'Window size
Private mWinSizeX As Integer, mWinSizeY As Integer
'Cursor pos
Private mCursorX As Integer, mCursorY As Integer

' Speed stats
Private mFrame As Integer
Private mNumVisibleSegs As Integer
Private mNumPix As Long
' Options
Private mRun As Boolean
Private mVB_CODE As Boolean


' Position of segment
Private Type tPosition
    x As Single
    y As Single
    ang As Single
End Type
Private mPositions(0 To NUM_POSITIONS) As tPosition

' Which graphic piece to use and the segment zoom factor
Private Type tSegment
    graphic As Integer
    zoom As Single
End Type
Private mSegments(0 To MAX_NUM_SEGMENTS) As tSegment

Private mNumSegments As Integer
Private mHead As Integer, mTail As Integer
'Turn speed
Private mTurnSpeed As Single, mTurnSpeedTarget As Single


'---------------------------------------------------------------------------------
' The main loop
'---------------------------------------------------------------------------------
Private Sub Animate()
    Dim i As Integer, scan As Integer, piece As Integer
    Dim x As Integer, y As Integer
    Dim directionToCursor As Single
    Dim headX As Single, headY As Single, newAng As Single
    
    mRun = True
    Timer1.Enabled = True

    mTurnSpeed = 0.01
    mTurnSpeedTarget = 0.01
    
    ' set segment pieces, sizes and mTail position
    Call hscShrink_Change
    
    Do While mRun
        Call mLightSpeed8(1).FillZero
        mNumVisibleSegs = 0
        'Draw segments TAIL first
        scan = mTail
        For i = 0 To mNumSegments - 1
            x = mPositions(scan).x
            y = mPositions(scan).y
            Call DrawSegment(x, y, mPositions(scan).ang, mSegments(i).graphic, mSegments(i).zoom)
            scan = scan - SEGMENT_SPACE * mSegments(i).zoom
            If scan < 0 Then scan = scan + NUM_POSITIONS
        Next
        
        'Get next direction
        headX = mPositions(mHead).x
        headY = mPositions(mHead).y
        directionToCursor = Atan2(mCursorY - headY, mCursorX - headX)
        newAng = StepAngle(mPositions(mHead).ang, directionToCursor, mTurnSpeed)
        'Move
        headX = headX + Cos(newAng) * SNAKE_SPEED
        headY = headY + Sin(newAng) * SNAKE_SPEED
        'New Head Pos
        mHead = mHead - 1
        If mHead < 0 Then
            mHead = mHead + NUM_POSITIONS
        End If
        'Save new position
        mPositions(mHead).x = headX
        mPositions(mHead).y = headY
        mPositions(mHead).ang = newAng
            
        'New Tail Pos
        mTail = mTail - 1
        If mTail < 0 Then
            mTail = mTail + NUM_POSITIONS
        End If
        
        'Change the serpent turning speed
        If chkAngleSpeed.value = 1 Then
            mTurnSpeed = 0.01
        Else
            If Abs(mTurnSpeed - mTurnSpeedTarget) < TURN_SPEED_DELTA Then
                mTurnSpeed = mTurnSpeedTarget
            Else
                If mTurnSpeedTarget > mTurnSpeed Then
                    mTurnSpeed = mTurnSpeed + TURN_SPEED_DELTA
                Else
                    mTurnSpeed = mTurnSpeed - TURN_SPEED_DELTA
                End If
            End If
        End If
'        Text1.text = Format(mTurnSpeed, "#0.00000")
        
        lblCogs.Caption = Str(mNumVisibleSegs) & " Segments"
        Call mLightSpeed8(1).PutPicture
        
        mFrame = mFrame + 1
        DoEvents
    Loop
    Timer1.Enabled = False
End Sub




'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Private Sub DrawSegment(centerX As Integer, centerY As Integer, ang As Single, piece As Integer, zoom As Single)
    Dim left As Integer, top As Integer
    Dim right As Integer, bottom As Integer
    Dim width As Integer, height As Integer
    Dim srcLeft As Integer, srcTop As Integer
    Dim srcWidth As Integer, srcHeight As Integer
    Dim srcCenterX As Integer, srcCenterY As Integer
    Dim segRad As Integer
    
    Select Case piece
    Case 0
        srcLeft = 0
        srcTop = 0
        srcWidth = 62
        srcHeight = 66
        srcCenterX = 15
        srcCenterY = 32
        segRad = 44
    Case 1
        srcLeft = 63
        srcTop = 0
        srcWidth = 62
        srcHeight = 66
        srcCenterX = 15
        srcCenterY = 32
        segRad = 44
    Case 2  'Head
        srcLeft = 126
        srcTop = 0
        srcWidth = 130
        srcHeight = 82
        srcCenterX = 15
        srcCenterY = 40
        segRad = 120
    Case 3  'Tail
        srcLeft = 1
        srcTop = 84
        srcWidth = 184
        srcHeight = 66
        srcCenterX = 137
        srcCenterY = 32
        segRad = 140
    End Select
    
    ' Shrink destination area by zoom factor
    segRad = segRad * zoom
    
    ' cull off-screen pieces
    left = centerX - segRad
    top = centerY - segRad
    If left >= mWinSizeX Then Exit Sub
    If top >= mWinSizeY Then Exit Sub
    right = centerX + segRad
    bottom = centerY + segRad
    If right <= 0 Then Exit Sub
    If bottom <= 0 Then Exit Sub
    
    ' ensure we don't write to off screen memory - crash
    If left < 0 Then left = 0
    If top < 0 Then top = 0
    If right > mWinSizeX Then right = mWinSizeX
    If bottom > mWinSizeY Then bottom = mWinSizeY
    
    width = right - left
    height = bottom - top

        Call VB8_ScaleRotate(mArray1(), mLightSpeedPitch1, _
                         left, top, centerX - left, centerY - top, _
                         width, height, mArray0(), mLightSpeedPitch0, _
                         srcLeft, srcTop, srcCenterX, srcCenterY, _
                         srcWidth, srcHeight, _
                         ang, zoom, True)
    
    mNumVisibleSegs = mNumVisibleSegs + 1
    mNumPix = mNumPix + CLng(width) * CLng(height)
End Sub





'---------------------------------------------------------------------------------
' Snake Parameters
'---------------------------------------------------------------------------------
Private Sub hscShrink_Change()
    Dim i As Integer, scan As Integer
    Dim value As Single, subt As Single
    
    'Ensure even number of segments
    mNumSegments = hscNumSegments.value And &HFFFE

    'Use the length and shrinkage values to calcuate tail positon
    'Set zoom factors at the same time
    subt = (hscShrink.value / 100) / mNumSegments
    value = 1#
    scan = mHead
    For i = mNumSegments - 1 To 0 Step -1
        mSegments(i).zoom = value
        value = value - subt
        
        scan = scan + SEGMENT_SPACE * value
        If scan >= NUM_POSITIONS Then scan = scan - NUM_POSITIONS
    Next
    mTail = scan
    
    ' set segment pieces
    For i = 0 To mNumSegments - 1
        mSegments(i).graphic = i And 1
    Next
    'Head and Tail
    mSegments(mNumSegments - 1).graphic = 2
    mSegments(0).graphic = 3
End Sub
Private Sub hscShrink_Scroll()
    Call hscShrink_Change
End Sub
Private Sub hscNumSegments_Change()
    Call hscShrink_Change
End Sub
Private Sub hscNumSegments_Scroll()
    Call hscShrink_Change
End Sub

'---------------------------------------------------------------------------------
' New Target Position
'---------------------------------------------------------------------------------
Private Sub Timer2_Timer()
    Dim x As Single, y As Single
    If chkRandomMove.value Then
        x = Rnd() * mWinSizeX
        y = Rnd() * mWinSizeY
        Call picMain_MouseMove(1, 0, x, y)
        mTurnSpeedTarget = 0.0005 + Rnd() * 0.015
    End If
End Sub



'---------------------------------------------------------------------------------
' UI Stuff
'---------------------------------------------------------------------------------
Private Sub optLanguage_Click(index As Integer)
    mVB_CODE = index
End Sub

Private Sub cmdPause_Click()
    If mRun Then
        mRun = False
        cmdPause.Caption = "START"
    Else
        cmdPause.Caption = "STOP"
        Call Animate
    End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button Then
        mCursorX = x
        mCursorY = y
    End If
End Sub
Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picMain_MouseMove(Button, Shift, x, y)
End Sub



'---------------------------------------------------------------------------------
' Statistics
'---------------------------------------------------------------------------------
Private Sub Timer1_Timer()
    lblFPS.Caption = "FPS:" & Str(mFrame)
    lblPix.Caption = Format(mNumPix / 1000000, "##0.0#") & " million pixels/second"
    mFrame = 0
    mNumPix = 0
End Sub


'---------------------------------------------------------------------------------
' LOAD - UNLOAD
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    Dim i As Integer
    If (App.LogMode <> 1) Then
        Call MsgBox("COMPILE ME!", vbOKOnly)
    End If

    'Position our snake off screen
    For i = 0 To NUM_POSITIONS - 1
        mPositions(i).x = -50
        mPositions(i).y = -50
        mPositions(i).ang = 0
    Next
End Sub

Private Sub Form_Activate()
    picMain.CurrentX = 30
    picMain.CurrentY = 60
    picMain.Print "1. Click 'START' to Begin."
    picMain.CurrentX = 80
    picMain.CurrentY = 250
    picMain.Print "There is NO block copying in this demo!"
    picMain.CurrentX = 45
    picMain.CurrentY = 280
    picMain.Print "Each segment is idividually scaled and rotated."
    
    ' Load our pictures
    picSource.Picture = LoadPicture("serpent.gif")

    mWinSizeX = picMain.ScaleWidth
    mWinSizeY = picMain.ScaleHeight
    mCursorX = mWinSizeX \ 2
    mCursorY = mWinSizeY \ 2
    
    'Create LIGHT SPEED objects
    Call CreateObjects

    'frequently used
    mLightSpeedPitch0 = mLightSpeed8(0).GetPitch()
    mLightSpeedPitch1 = mLightSpeed8(1).GetPitch()
      
    mVB_CODE = True
    mRun = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = mRun
    mRun = False
    cmdPause.Caption = "START"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call DeleteObjects
End Sub


Private Sub CreateObjects()
    Set mLightSpeed8(0) = New clsLightSpeed8
    Set mLightSpeed8(1) = New clsLightSpeed8
    
    'Initialize the source LS object with the size we know the file is
    Call mLightSpeed8(0).InitDimensions(256, 256)
    'Set the palette to convert the GIF file to
    mLightSpeed8(0).ReadPaletteFile ("serpent.pal")
    Call mLightSpeed8(0).SetPalette(0, 256)
    ' Capture the source picturebox to the DIB
    Set mLightSpeed8(0).SetPictureBox = picSource
    Call mLightSpeed8(0).GrabPicture

    'Initialize the Destination LS object and set it's palette to match the Source
    Call mLightSpeed8(1).InitPicture(picMain, False)
    mLightSpeed8(1).ReadPaletteFile ("serpent.pal")
    Call mLightSpeed8(1).SetPalette(0, 256)
    
    'Store original pointer for array recovery
    mArray0Pnt = mLightSpeed8(0).GetArray(mArray0)
    mArray1Pnt = mLightSpeed8(1).GetArray(mArray1)
End Sub


Private Sub DeleteObjects()
    'Recover our arrays and erase them
    If mArray0Pnt Then
        Call mLightSpeed8(0).FixArray(mArray0, mArray0Pnt)
        Erase mArray0
        mArray0Pnt = 0
    End If
    If mArray1Pnt Then
        Call mLightSpeed8(1).FixArray(mArray1, mArray1Pnt)
        Erase mArray1
        mArray1Pnt = 0
    End If
    'Delete objects
    If Not mLightSpeed8(0) Is Nothing Then
        Set mLightSpeed8(0) = Nothing
    End If
    If Not mLightSpeed8(1) Is Nothing Then
        Set mLightSpeed8(1) = Nothing
    End If
End Sub


