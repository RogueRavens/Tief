‘-------------------------------------------------------
‘ AnimateSprite Source Code File
‘-------------------------------------------------------
Option Explicit
Option Base 0

‘Windows API functions and structures
Private Declare Function GetTickCount Lib “kernel32” () As Long

Const C_BLACK As Long = &H0

Const KEY_ESC As Integer = 27
Const KEY_LEFT As Integer = 37
Const KEY_UP As Integer = 38
Const KEY_RIGHT As Integer = 39
Const KEY_DOWN As Integer = 40

‘program constants
Const SCREENWIDTH As Long = 640
Const SCREENHEIGHT As Long = 480
Const FULLSCREEN As Boolean = True
Const DRAGONSPEED As Integer = 4

Dim dragonSpr As TSPRITE
Dim dragonImg As Direct3DTexture8

Dim terrain As Direct3DSurface8
Dim backbuffer As Direct3DSurface8

Private Sub Form_Load()
    ‘set up the main form
    Form1.Caption = “AnimateSprite”
    Form1.KeyPreview = False
    Form1.ScaleMode = 3
    Form1.width = Screen.TwipsPerPixelX * (SCREENWIDTH + 12)
    Form1.height = Screen.TwipsPerPixelY * (SCREENHEIGHT + 30)
    Form1.Show

    ‘initialize Direct3D
     InitDirect3D Me.hwnd, SCREENWIDTH, SCREENHEIGHT, FULLSCREEN

    ‘get the back buffer
    Set backbuffer = d3ddev.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)

    ‘load the background image
    Set terrain = LoadSurface(App.Path & “\terrain.png”, 640, 480)
    If terrain Is Nothing Then
        MsgBox “Error loading png”
        Shutdown
    End If

    ‘load the dragon sprite
    Set dragonImg = LoadTexture(d3ddev, App.Path & “\dragon.png”)

    ‘initialize the dragon sprite
     InitSprite d3ddev, dragonSpr
     With dragonSpr
          .FramesPerRow = 8
          .FrameCount = 8
          .DirX = 0
        .CurrentFrame = 0
        .AnimDelay = 2
         .width = 128
        .height = 128
        .ScaleFactor = 1
        .x = 150
        .y = 100
    End With

    Dim start As Long
    start = GetTickCount()

    ‘start main game loop
    Do While True

        If GetTickCount - start > 25 Then

            ‘draw the background
             DrawSurface terrain, 0, 0

            ‘start rendering
            d3ddev.BeginScene

            ‘move the dragon
            MoveDragon

           ‘animate the dragon
           AnimateDragon

          ‘stop rendering
            d3ddev.EndScene

         ‘draw the back buffer to the screen
          d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0

         start = GetTickCount
         DoEvents
    End If


  Loop
End Sub

Public Sub MoveDragon()
    With dragonSpr
       .x = .x + .SpeedX
       If .x < 0 Then
          .x = 0
          .SpeedX = 0
        End If
        If .x > SCREENWIDTH - .width Then
           .x = SCREENWIDTH - .width - 1
           .SpeedX = 0
        End If

        .y = .y + .SpeedY
        If .y < 0 Then
           .y = 0
           .SpeedX = 0
        End If
        If .y > SCREENHEIGHT - .height Then
           .y = SCREENHEIGHT - .height - 1
           .SpeedX = 0
        End If
    End With

End Sub

Public Sub AnimateDragon()

    With dragonSpr
         ‘increment the animation counter
         .AnimCount = .AnimCount + 1

         ‘has the animation counter waited long enough?
          If .AnimCount > .AnimDelay Then
             .AnimCount = 0

         ‘okay, go to the next frame
          .CurrentFrame = .CurrentFrame + 1
 
         ‘loop through the frames
         If .CurrentFrame > .StartFrame + 7 Then
            .CurrentFrame = .StartFrame
         End If

    End If
End With

‘draw the dragon sprite
DrawSprite dragonImg, dragonSpr, &HFFFFFFFF

End Sub

Public Sub DrawSurface(ByRef source As Direct3DSurface8, _
ByVal x As Long, ByVal y As Long)
     Dim r As DxVBLibA.RECT
     Dim point As DxVBLibA.point
     Dim desc As D3DSURFACE_DESC
    source.GetDesc desc

    ‘set dimensions of the source image
     r.Left = x
     r.Top = y
     r.Right = x + desc.width
     r.Bottom = y + desc.height

    ‘set the destination point
    point.x = 0
    point.y = 0

   ‘draw the scroll window
   d3ddev.CopyRects source, r, 1, backbuffer, point

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shutdown
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case KEY_UP
             dragonSpr.StartFrame = 0
             dragonSpr.CurrentFrame = 0
             dragonSpr.SpeedX = 0
             dragonSpr.SpeedY = -DRAGONSPEED
             
        Case KEY_RIGHT
            dragonSpr.StartFrame = 16
            dragonSpr.CurrentFrame = 16
           dragonSpr.SpeedX = DRAGONSPEED
           dragonSpr.SpeedY = 0

        Case KEY_DOWN
             dragonSpr.StartFrame = 32
             dragonSpr.CurrentFrame = 32
            dragonSpr.SpeedX = 0
            dragonSpr.SpeedY = DRAGONSPEED
        Case KEY_LEFT
             dragonSpr.StartFrame = 48
            dragonSpr.CurrentFrame = 48
            dragonSpr.SpeedX = -DRAGONSPEED
            dragonSpr.SpeedY = 0
        Case KEY_ESC
            Shutdown
    End Select
    
End Sub