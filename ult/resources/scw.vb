‘-------------------------------------------------------
‘ ScrollWorld by Eric
‘
‘ Requires: Globals.bas, Direct3D.vb, Ts.vb
‘-------------------------------------------------------

Option Explicit
Option Base 0

Private Sub Form_Load()
    ‘set up the main form
    Form1.Caption = “ScrollWorld”
    Form1.ScaleMode = 3
    Form1.width = Screen.TwipsPerPixelX * (SCREENWIDTH + 12)
    Form1.height = Screen.TwipsPerPixelY * (SCREENHEIGHT + 30)
    Form1.Show

   ‘initialize Direct3D
   InitDirect3D Me.hwnd, SCREENWIDTH, SCREENHEIGHT, FULLSCREEN

   ‘get reference to the back buffer
    Set backbuffer = d3ddev.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)

    ‘load the bitmap file
    Set tiles = LoadSurface(App.Path & “\map1.bmp”, 1024, 640)

    ‘load the map data from the Mappy export file
     LoadMap App.Path & “\map1.txt”

    ‘create the small scroll buffer surface
    Set scrollbuffer = d3ddev.CreateImageSurface(
        SCROLLBUFFERWIDTH, _
        SCROLLBUFFERHEIGHT, _
        dispmode.Format)

    ‘this helps to keep a steady framerate
     Dim start As Long
     start = GetTickCount()

    ‘clear the screen to black
     d3ddev.Clear 0, ByVal 0, D3DCLEAR_TARGET, C_BLACK, 1, 0
    ‘main loop
    Do While (True)
        ‘update the scroll position
        UpdateScrollPosition

        ‘draw tiles onto the scroll buffer
        DrawTiles

        ‘draw the scroll window onto the screen
        DrawScrollWindow

        ‘set the screen refresh to about 40 fps
        If GetTickCount - start > 25 Then
            d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
            start = GetTickCount
            DoEvents
        End If
    Loop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)

    ‘move mouse on left side to scroll left
     If X < SCREENWIDTH / 2 Then SpeedX = -STEP

    ‘move mouse on right side to scroll right
     If X > SCREENWIDTH / 2 Then SpeedX = STEP

     ‘move mouse on top half to scroll up
     If Y < SCREENHEIGHT / 2 Then SpeedY = -STEP

     ‘move mouse on bottom half to scroll down
     If Y > SCREENHEIGHT / 2 Then SpeedY = STEP
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 27 Then Shutdown
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Shutdown
End Sub