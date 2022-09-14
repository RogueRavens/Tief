‘-------------------------------------------------------
‘ WalkAbout program

‘ Requires the following files:
‘ Direct3D, Global, TileScroller,
‘ Sprite, and mt
‘-------------------------------------------------------

Option Explicit
Option Base 0

Const HEROSPEED As Integer = 4
Dim heroSpr As TSPRITE
Dim heroImg As Direct3DTexture8
Dim SuperHero As Boolean

Dim frm As New Form1

Public Sub Main()

    ‘set up the main form
    frm.Caption = “WalkAbout”
    frm.AutoRedraw = False
    frm.BorderStyle = 1
    frm.ClipControls = False
    frm.ScaleMode = 3
    frm.width = Screen.TwipsPerPixelX * (SCREENWIDTH + 12)
    frm.height = Screen.TwipsPerPixelY * (SCREENHEIGHT + 30)
    frm.Show

    ‘initialize Direct3D
     InitDirect3D frm.hwnd

     InitDirectInput
     InitKeyboard frm.hwnd

    ‘get reference to the back buffer
     Set backbuffer = d3ddev.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
     
    ‘load the bitmap file
     Set tiles = LoadSurface(App.Path & “\ireland.bmp”, 1024, 576)

    ‘load the map data from the Mappy export file
    LoadBinaryMap App.Path & “\ireland.mar”, MAPWIDTH, MAPHEIGHT

    ‘load the dragon sprite  
     Set heroImg = LoadTexture(d3ddev, App.Path & “\hero_sword_walk.bmp”)

    ‘initialize the dragon sprite
    InitSprite d3ddev, heroSpr
    With heroSpr
        .FramesPerRow = 9
        .FrameCount = 9
        .CurrentFrame = 0
        .AnimDelay = 2
        .width = 96
        .height = 96
        .ScaleFactor = 1
        .x = (SCREENWIDTH - .width) / 2
        .y = (SCREENHEIGHT - .height) / 2
    End With

    ‘create the small scroll buffer surface
    Set scrollbuffer = d3ddev.CreateImageSurface( _
        SCROLLBUFFERWIDTH, _
        SCROLLBUFFERHEIGHT, _
        dispmode.Format)

    ‘start player in the city of Dubh Linn
     ScrollX = 1342 * TILEWIDTH
     ScrollY = 945 * TILEHEIGHT

    ‘this helps to keep a steady framerate
      Dim start As Long
     start = GetTickCount()

    ‘clear the screen to black 
    d3ddev.Clear 0, ByVal 0, D3DCLEAR_TARGET, C_BLACK, 1, 0

    ‘main loop
    Do While (True)
        ‘poll DirectInput for keyboard input
        Check_Keyboard

        ‘update the scrolling window
        UpdateScrollPosition
        DrawTiles
        DrawScrollWindow
        Scroll 0, 0
        ShowScrollData

       ‘reset scroll speed
        SuperHero = False

       ‘set the screen refresh to about 50 fps
        If GetTickCount - start > 20 Then

            ‘start rendering
            d3ddev.BeginScene

            ‘animate the dragon
            If heroSpr.Animating Then
                AnimateSprite heroSpr, heroImg
            End If

            ‘draw the hero sprite
             DrawSprite heroImg, heroSpr, &HFFFFFFFF

             ‘stop rendering
             d3ddev.EndScene

            d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
            start = GetTickCount
            DoEvents
        End If
    Loop
End Sub

Public Sub ShowScrollData()
    Static old As point
    Dim player As point
    Dim tile As point
    Dim tilenum As Long

    player.x = ScrollX + SCREENWIDTH / 2
    player.y = ScrollY + SCREENHEIGHT / 2
    tile.x = player.x \ TILEWIDTH
    tile.y = player.y \ TILEHEIGHT

    If (tile.x <> old.x) Or (tile.y <> old.y) Then
        old = tile
        tilenum = mapdata(tile.y * MAPWIDTH + tile.x)
        frm.Caption = “WalkAbout - “ & _
            “Scroll=(“ & player.x & “,” & player.y & “) “ & _
            “Tile(“ & tile.x & “,” & tile.y & “)=” & tilenum
    End If
End Sub

‘This is called from DirectInput.bas on keypress events
Public Sub KeyPressed(ByVal key As Long)
    Select Case key
        Case KEY_UP, KEY_NUMPAD8
            heroSpr.AnimSeq = 0
            heroSpr.Animating = True
            Scroll 0, -HEROSPEED

        Case KEY_NUMPAD9
            heroSpr.AnimSeq = 1
            heroSpr.Animating = True
            Scroll HEROSPEED, -HEROSPEED

        Case KEY_RIGHT, KEY_NUMPAD6
            heroSpr.AnimSeq = 2
            heroSpr.Animating = True
            Scroll HEROSPEED, 0

        Case KEY_NUMPAD3
            heroSpr.AnimSeq = 3
            heroSpr.Animating = True
            Scroll HEROSPEED, HEROSPEED
        
        Case KEY_DOWN, KEY_NUMPAD2
            heroSpr.AnimSeq = 4
            heroSpr.Animating = True
            Scroll 0, HEROSPEED

        Case KEY_NUMPAD1
            heroSpr.AnimSeq = 5
            heroSpr.Animating = True
            Scroll -HEROSPEED, HEROSPEED

        Case KEY_LEFT, KEY_NUMPAD4
            heroSpr.AnimSeq = 6
            heroSpr.Animating = True
            Scroll -HEROSPEED, 0

        Case KEY_NUMPAD7
            heroSpr.AnimSeq = 7
            heroSpr.Animating = True
            Scroll -HEROSPEED, -HEROSPEED

        Case KEY_LSHIFT, KEY_RSHIFT
            SuperHero = True

        Case KEY_ESC
        Shutdown

    End Select

    ‘Debug.Print “Key = “ & key
End Sub

Public Sub Scroll(ByVal horiz As Long, ByVal vert As Long)
    SpeedX = horiz
    SpeedY = vert

    If SuperHero Then
        SpeedX = SpeedX * 4
        SpeedY = SpeedY * 4
    End If
End Sub