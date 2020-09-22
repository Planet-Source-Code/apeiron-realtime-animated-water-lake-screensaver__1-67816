VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Lake Screensaver (Debug Mode)"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picBeachMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2760
      Left            =   3000
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   550
      TabIndex        =   16
      Top             =   5520
      Visible         =   0   'False
      Width           =   8310
   End
   Begin VB.PictureBox picBeach 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2760
      Left            =   2400
      Picture         =   "Form1.frx":54B2
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   550
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   8310
   End
   Begin VB.PictureBox picfishmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   360
      Picture         =   "Form1.frx":12D94
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox picShipMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   4650
      Left            =   8520
      Picture         =   "Form1.frx":1589F
      ScaleHeight     =   306
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.PictureBox picShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4650
      Left            =   6480
      Picture         =   "Form1.frx":185F8
      ScaleHeight     =   306
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.PictureBox picOverlayMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   4440
      Picture         =   "Form1.frx":1D16C
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   10
      Top             =   7680
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.PictureBox picOverlay 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   3840
      Picture         =   "Form1.frx":1F515
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.PictureBox picSpriteMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   2400
      Picture         =   "Form1.frx":2EB6C
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   6060
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   2400
      Picture         =   "Form1.frx":34992
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   6060
   End
   Begin VB.PictureBox picBackMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   1320
      Picture         =   "Form1.frx":3BD34
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5460
      Left            =   360
      Picture         =   "Form1.frx":421BF
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   9660
      Begin VB.PictureBox picFish 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   660
         Left            =   0
         Picture         =   "Form1.frx":7671B
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   80
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.PictureBox picSky 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   360
      Picture         =   "Form1.frx":796A1
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1280
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   19260
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   6840
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H80000009&
      Height          =   4695
      Left            =   360
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   6600
      Picture         =   "Form1.frx":88D17
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   9660
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const NUM_FRAMES As Integer = 12

Dim nImgWidth As Double
Dim nImgHeight As Double
Dim framenumber As Integer

Private Type sprite
    X As Long
    Y As Long
    z As Long           '' Not used 0 closest to background 3 (5) farthest out and lower on the inverted pic (0 is higher on the inverted pic)
    shDc As Long        ' Source hdc of sprite
    speed As Integer    ' How fast it moves
    FramesToWait As Integer ' Pause before new instance is created for a certain number of frames
    frameNum As Integer
    visible As Boolean
    width As Single
    height As Single
End Type

Dim Ship As sprite

Dim ShipWidth As Long
Dim ShipHeight As Long

Dim bf As BLENDFUNCTION, lBF As Long

Private Sub BlendTest()

    With bf
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = waterTransparency
        .AlphaFormat = 0
    End With
    
    RtlMoveMemory lBF, bf, 4
    
End Sub


Private Sub Command1_Click()
        bCancel = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        bCancel = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Init
    'Unload Me
    UnloadAll
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 1, ByVal 0&, 0
    DoEvents
End Sub

Public Sub Init(Optional initString As String)

    Dim WinStyle As Long
    Dim PreviewRect As RECT
        
    If initString = "" Then
        initString = Command$
    End If
    
    Randomize (Now)
      
    ReadIni
    BlendTest
    DisplayMode = Mid$(LCase$(Trim$(initString)), 1, 2) ' /s, /c, /p
    
    nImgHeight = picBackground.ScaleHeight
    nImgWidth = picBackground.ScaleWidth
    picDisplay.height = 2 * picBackground.height
    picDisplay.width = picBackground.width

    picBuffer.width = 3 * picBackground.width
    picBuffer.height = picBackground.height
    
    Select Case DisplayMode
        Case "/p" ' preview
            PreviewHWND = GetHwndFromCmd(Command$)
            frmPreview.Show
            frmPreview.visible = False
            GetClientRect PreviewHWND, PreviewRect
            WinStyle = GetWindowLong(frmPreview.hwnd, GWL_STYLE)
            SetWindowLong frmPreview.hwnd, GWL_STYLE, WinStyle Or WS_CHILD
            SetWindowLong frmPreview.hwnd, GWL_HWNDPARENT, PreviewHWND
            SetParent frmPreview.hwnd, PreviewHWND
            SetWindowPos frmPreview.hwnd, HWND_TOP, 0&, 0&, PreviewRect.Right, PreviewRect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
            SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 0, ByVal 0&, 0
            picDisplay.AutoRedraw = True
            frmPreview.visible = True
        Case "/s" ',"" ' Full screen normal
            SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 0, ByVal 0&, 0
            ShowCursor False
            picDisplay.AutoRedraw = True
            SetStretchBltMode picDisplay.hdc, vbPaletteModeNone
            SetWindowPos frmDisplay.hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            frmDisplay.Show
        'Case "/a"
        Case "/c"
            frmConfig.Show
        Case Else
            Me.Show
            picDisplay.AutoRedraw = False
    End Select

    SetStretchBltMode picBackground.hdc, vbPaletteModeNone
    SetStretchBltMode picDisplay.hdc, vbPaletteModeNone
    SetStretchBltMode picBuffer.hdc, vbPaletteModeNone

    bCancel = False
    DoEvents
    runIt
        
End Sub

Public Sub runIt()
    Dim skyPos As Double 'Long
    Dim bird1 As sprite
    Dim fish As sprite
    Dim tickTime As Long
    Dim theTime As Long
    Dim FPS As Integer
    Dim skyTest As Long
    Dim shipSourceTemp As Double
    Dim shipDestTemp As Double
    
    ShipWidth = picShip.ScaleWidth
    ShipHeight = picShip.ScaleHeight
    
'    With bird1
'        .X = 0
'        .Y = nImgHeight \ 2
        '.shDc = picSprite.hdc
'        .speed = 5
'        .visible = True
'    End With
    
    With Ship
        .X = 50
        .Y = 50
        .width = ShipWidth
        .height = ShipHeight
        .visible = True
    End With
    
'    With fish
'        .X = 50
'        .Y = nImgHeight + 100
'        .shDc = picSprite.hdc
'        .speed = 5
'        .visible = True
'    End With
    
    framenumber = 12
    
    Do While Not bCancel
        tickTime = GetTickCount
        ' Draw Sky
        If skyPos + 640 > picSky.ScaleWidth Then
            BitBlt picBackground.hdc, 0, 0, 640, nImgHeight, picSky.hdc, skyPos, 0, vbSrcCopy
            BitBlt picBackground.hdc, picSky.ScaleWidth - skyPos, 0, 640, nImgHeight, picSky.hdc, 0, 0, vbSrcCopy
        Else
            BitBlt picBackground.hdc, 0, 0, 640, nImgHeight, picSky.hdc, skyPos, 0, vbSrcCopy
        End If
        skyPos = skyPos + 0.25 ' 0.5 ' 1
        If skyPos >= picSky.ScaleWidth Then
            'Stop
            skyPos = 0 'picSky.ScaleWidth - skyPos - 640
        End If
        
        ' Draw Background

        If framenumber >= 12 Then
            framenumber = 0 'NUM_FRAMES
        End If
        
        framenumber = framenumber + 1
        
        ' Draw Sprite
        If birdAnim Then
            If bird1.visible Then
                If bird1.frameNum >= 9 Then bird1.frameNum = 0
                BitBlt picBackground.hdc, bird1.X, bird1.Y, 40, 40, picSpriteMask.hdc, bird1.frameNum * 40, 0, vbSrcAnd
                BitBlt picBackground.hdc, bird1.X, bird1.Y, 40, 40, picSprite.hdc, bird1.frameNum * 40, 0, vbSrcPaint
                bird1.frameNum = bird1.frameNum + 1
                bird1.X = bird1.X + bird1.speed
                If bird1.X > nImgWidth Then
                    bird1.visible = False
                    bird1.FramesToWait = Rnd() * 400
                    bird1.frameNum = 0
                End If
            Else
                If bird1.FramesToWait <= 0 Then
                    ' Start a fresh instance of a bird
                    bird1.X = 0
                    bird1.Y = Rnd() * (nImgHeight * 0.75)
                    bird1.visible = True
                    bird1.speed = Rnd() * 5 + 10
                    bird1.frameNum = 0
                Else
                    ' Wait some more frames to create another bird
                    bird1.FramesToWait = bird1.FramesToWait - 1
                End If
            End If
        End If
        
        If showBack Then
        ' Draw Overlay
            BitBlt picBackground.hdc, 0, 0, 640, 240, picOverlayMask.hdc, 0, 0, vbSrcAnd
            BitBlt picBackground.hdc, 0, 0, 640, 240, picOverlay.hdc, 0, 0, vbSrcPaint
        End If
        
        'createAnimation
        makeWaves framenumber
        
        ' Draw underwater picture
        BitBlt picDisplay.hdc, 0, nImgHeight, 640, 240, picBack.hdc, 0, 0, vbSrcCopy
        
        ' Draw fish
         If fishAnim Then
            If fish.visible Then
                'If fish.frameNum >= 9 Then fish.frameNum = 0
                BitBlt picDisplay.hdc, fish.X, fish.Y, 80, 40, picfishmask.hdc, 0, 0, vbSrcAnd
                BitBlt picDisplay.hdc, fish.X, fish.Y, 80, 40, picFish.hdc, 0, 0, vbSrcPaint
                fish.frameNum = fish.frameNum + 1
                fish.X = fish.X + fish.speed
                fish.Y = fish.Y + (Sin(fish.X) * 10)
                If fish.X > nImgWidth Then
                    fish.visible = False
                    fish.FramesToWait = Rnd() * 100
                    fish.frameNum = 0
                End If
            Else
                If fish.FramesToWait <= 0 Then
                    ' Start a fresh instance of a fish
                    fish.X = 0
                    fish.Y = Rnd() * (nImgHeight * 0.75) + nImgHeight
                    fish.visible = True
                    fish.speed = Rnd() * 10 + 5
                    fish.frameNum = 0
                Else
                    ' Wait some more frames to create another fish
                    fish.FramesToWait = fish.FramesToWait - 1
                End If
            End If
        End If

        ' Draw waves
        AlphaBlend picDisplay.hdc, 0, nImgHeight, nImgWidth, nImgHeight, picBuffer.hdc, 0, 0, nImgWidth, nImgHeight, lBF
        
        If showBack Then
            ' Draw Overlay
            BitBlt picBackground.hdc, 0, 0, 640, 240, picOverlayMask.hdc, 0, 0, vbSrcAnd
            BitBlt picBackground.hdc, 0, 0, 640, 240, picOverlay.hdc, 0, 0, vbSrcPaint
        End If
        
        BitBlt picDisplay.hdc, 0, 0, nImgWidth, nImgHeight, picBackground.hdc, 0, 0, vbSrcCopy

        
        ' Draw ship on waves only if not behind overlay
        ' Ship stuff
        If shipAnim Then
            If Ship.X + ShipWidth < 297 And Ship.visible Then
                StretchBlt picDisplay.hdc, Ship.X, Ship.Y, Ship.width, Ship.height, picShipMask.hdc, 0, 0, ShipWidth, ShipHeight, vbSrcAnd
                StretchBlt picDisplay.hdc, Ship.X, Ship.Y, Ship.width, Ship.height, picShip.hdc, 0, 0, ShipWidth, ShipHeight, vbSrcPaint
            ElseIf Ship.X > nImgWidth Then
                ' Reset Ship
                Ship.X = 50
                Ship.Y = 50
                Ship.width = ShipWidth
                Ship.height = ShipHeight
                Ship.visible = True
            ElseIf Ship.visible Then
                'SetStretchBltMode picDisplay.hdc, vbPaletteModeNone
                StretchBlt picDisplay.hdc, Ship.X, Ship.Y, Ship.width, Ship.height, picShipMask.hdc, 0, 0, ShipWidth, ShipHeight, vbSrcAnd
                StretchBlt picDisplay.hdc, Ship.X, Ship.Y, Ship.width, Ship.height, picShip.hdc, 0, 0, ShipWidth, ShipHeight, vbSrcPaint
                If showBack Then
                    ' Redraw overlay and waves to cover ship going behind
                    BitBlt picDisplay.hdc, 0, 0, 640, 240, picOverlayMask.hdc, 0, 0, vbSrcAnd
                    BitBlt picDisplay.hdc, 0, 0, 640, 240, picOverlay.hdc, 0, 0, vbSrcPaint
                    'BitBlt picDisplay.hdc, 295, nImgHeight, 640, 35, picBuffer.hdc, 295, 0, vbSrcCopy
                    AlphaBlend picDisplay.hdc, 298, nImgHeight, nImgWidth, 35, picBuffer.hdc, 298, 0, nImgWidth, 35, lBF
                End If
            End If
            
            If Ship.height >= 50 Then
                Ship.height = Ship.height - 5
                Ship.Y = Ship.Y + 3
                If Ship.width > 60 Then
                    Ship.width = Ship.width - 1
                End If
            ElseIf Ship.height >= 25 Then
                Ship.height = Ship.height - 3
                Ship.Y = Ship.Y + 1.5
                If Ship.width > 10 Then
                    Ship.width = Ship.width - 4
                End If
            Else
                Ship.visible = False
            End If
            
            Ship.X = Ship.X + 5
        End If
        
        If showBeach Then
            BitBlt picDisplay.hdc, (picOverlay.ScaleWidth) - picBeachMask.ScaleWidth, (picOverlay.ScaleHeight * 2) - picBeachMask.ScaleHeight, 550, 180, picBeachMask.hdc, 0, 0, vbSrcAnd
            BitBlt picDisplay.hdc, (picOverlay.ScaleWidth) - picBeach.ScaleWidth, (picOverlay.ScaleHeight * 2) - picBeach.ScaleHeight, 550, 180, picBeach.hdc, 0, 0, vbSrcPaint
        End If
        
        DoEvents
        
        Select Case DisplayMode
            Case "/p" ' preview
                StretchBlt frmPreview.hdc, 0, 0, frmPreview.ScaleWidth, frmPreview.ScaleHeight, picDisplay.hdc, 0, 0, 640, 480, vbSrcCopy
            Case "/s" ' Full screen normal
                StretchBlt frmDisplay.hdc, 0, 0, Screen.width / Screen.TwipsPerPixelX, Screen.height / Screen.TwipsPerPixelY, picDisplay.hdc, 0, 0, 640, 480, vbSrcCopy
            'Case "/a" ' Password TO DO add this ?
            Case Else ' Just run it normally, probably testing
                'StretchBlt frmDisplay.hdc, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, picDisplay.hdc, 0, 0, 640, 480, vbSrcCopy
        End Select
        
        theTime = (GetTickCount - tickTime)
        
        If theTime > 0 Then
            FPS = (1000 / theTime)
        End If
        
        If FPS > 15 Then
            Sleep (1000 \ 15) - theTime
        End If
        
        'Me.Caption = "Lake Screensaver: FPS" & FPS
    Loop

    UnloadAll
End Sub

Public Sub makeWaves(phase As Integer)
        Dim p1 As Double
        Dim dispx As Integer, dispy As Integer
        Dim i As Double
        Dim nImg14 As Single, refY As Long, refX As Long
        ' Thanks to David Griffiths for the original java lake
        ' whose math I adapted for this.

        ' Flip Buffer
        StretchBlt picBuffer.hdc, 2 * nImgWidth, nImgHeight, nImgWidth, -nImgHeight, picBackground.hdc, 0, 0, nImgWidth, nImgHeight, vbSrcCopy

        ' Draw Ship Reflection on water
        refX = 2 * nImgWidth + Ship.X
        refY = nImgHeight - Ship.Y - Ship.height - 20
        StretchBlt picBuffer.hdc, refX, refY, Ship.width, Ship.height, picShipMask.hdc, 0, 0, ShipWidth, ShipHeight, vbSrcAnd
        StretchBlt picBuffer.hdc, refX, refY, Ship.width, Ship.height, picShip.hdc, 0, 0, ShipWidth, ShipHeight, vbSrcPaint

        If showBack Then ' Cover what is hidden by land
            StretchBlt picBuffer.hdc, 2 * nImgWidth + 300, nImgHeight, nImgWidth, -nImgHeight, picBackground.hdc, 300, 0, nImgWidth, nImgHeight, vbSrcCopy
        End If

        p1 = 2 * 3.14 * phase / NUM_FRAMES ' 3.14=PI

'        Buffer is in reverse order.  Inverted pic is all the way to right
'        Final frame is the one at left
        dispx = 0 '(NUM_FRAMES - phase) * nImgWidth

        For i = 0 To nImgHeight
'          dispy defines the vertical sine displacement. It
'          attenuates higher up the image, for perspective
            nImg14 = nImgHeight / 14
            dispy = (nImg14 * (i + 28#) * Sin((nImg14 * (nImgHeight - i)) / CDbl(i + 1) + p1) / nImgHeight)
    
            If i < -dispy Then
            ' Copy Original line because it falls out of range
                BitBlt picBuffer.hdc, dispx, i - 1, nImgWidth, 2, picBuffer.hdc, 2 * nImgWidth, i, vbSrcCopy
            Else
            'Else copy dithered line.
            'Added two tests.
            'The first is to check if it falls off of the bottom of the
            'picture.  The next is if it is before the beginning of the picture.
                
                If nImgHeight - (i + dispy) <= 0 Then
                    dispy = -dispy
                End If
                    
                If i + dispy <= 0 Then
                    dispy = 1
                End If
                
                ' Displacement all fixed so blt this line.
                BitBlt picBuffer.hdc, 0, i, nImgWidth, 1, picBuffer.hdc, 2 * nImgWidth, i + dispy, vbSrcCopy
            End If
            DoEvents
        Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bCancel = True
    DoEvents
    ShowCursor True
    UnloadAll
End Sub
