‘-------------------------------------------------------
‘ SpriteTest Source Code File
‘-------------------------------------------------------
Option Explicit
Option Base 0

‘Windows API functions and structures
Private Declare Function GetTickCount Lib “kernel32” () As Long

‘program constants
Const SCREENWIDTH As Long = 640
Const SCREENHEIGHT As Long = 480
Const FULLSCREEN As Boolean = False

Dim tavernSprite As TSPRITE
Dim tavernImage As Direct3DTexture8

Dim terrain As Direct3DSurface8
Dim backbuffer As Direct3DSurface8

Private Sub Form_Load()
    Static lStartTime As Long
    Static lCounter As Long
    Static lNewTime As Long

‘set up the main form
Form1.Caption = “DrawSprite”
Form1.KeyPreview = True
Form1.ScaleMode = 3
Form1.width = Screen.TwipsPerPixelX * (SCREENWIDTH + 12)
Form1.height = Screen.TwipsPerPixelY * (SCREENHEIGHT + 30)
Form1.Show

‘initialize Direct3D
InitDirect3D Me.hwnd, SCREENWIDTH, SCREENHEIGHT, FULLSCREEN

Set backbuffer = d3ddev.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)

Set terrain = LoadSurface(App.Path & “\terrain.bmp”, 640, 480)
If terrain Is Nothing Then
    MsgBox “Error loading bitmap”
    Shutdown
End If

Set tavernImage = LoadTexture(d3ddev, App.Path & “\tavern.bmp”)
InitSprite d3ddev, tavernSprite
With tavernSprite
    .width = 220
    .height = 224
    .ScaleFactor = 1
    .x = 200
     .y = 100
End With

Dim start As Long
start = GetTickCount()

‘start main game loop
Do While True

   If GetTickCount - start > 25 Then
       ‘draw the background image
       DrawSurface terrain, 0, 0

       ‘start rendering
       d3ddev.BeginScene

       ‘draw the sprite
       spritedraw tavernImage, tavernSprite

       ‘stop rendering
       d3ddev.EndScene

       ‘draw the back buffer to the screen
        d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0

        start = GetTickCount
         DoEvents
    End If

   Loop
End Sub
Public Sub DrawSurface( _
    ByRef source As Direct3DSurface8, _
    ByVal x As Long, _
    ByVal y As Long)

    Dim r As DxVBLibA.RECT
    Dim point As DxVBLibA.point
    Dim desc As D3DSURFACE_DESC

    ‘get the properties for the surface
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
    If KeyCode = 27 Then Shutdown
End Sub