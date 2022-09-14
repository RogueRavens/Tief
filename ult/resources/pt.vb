‘-------------------------------------------------------
‘ PrintText Program
‘-------------------------------------------------------
Option Explicit
Option Base 0

Const SCREENWIDTH As Long = 800
Const SCREENHEIGHT As Long = 600

Const C_PURPLE As Long = &HFFFF00FF
Const C_RED As Long = &HFFFF0000
Const C_GREEN As Long = &HFF00FF00
Const C_BLUE As Long = &HFF0000FF
Const C_WHITE As Long = &HFFFFFFFF
Const C_BLACK As Long = &H0
Const C_GRAY As Long = &HFFAAAAAA

Dim fontImg As Direct3DTexture8
Dim fontSpr As TSPRITE

Private Sub Form_Load()
    ‘set up the main form
    Me.Caption = “PrintText”
    Me.ScaleMode = 3
    Me.width = Screen.TwipsPerPixelX * (SCREENWIDTH + 12)
    Me.height = Screen.TwipsPerPixelY * (SCREENHEIGHT + 30)
    Me.Show

    ‘initialize Direct3D
    InitDirect3D Me.hWnd

    ‘get reference to the back buffer
    Set backbuffer = d3ddev.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)

    ‘load the bitmap file
    Set fontImg = LoadTexture(d3ddev, App.Path & “\font.bmp”)

    InitSprite d3ddev, fontSpr
    fontSpr.FramesPerRow = 20
    fontSpr.width = 16
    fontSpr.height = 24
    fontSpr.ScaleFactor = 2

   ‘clear the screen to black
    d3ddev.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0, 1, 0

    d3ddev.BeginScene

    PrintText fontImg, fontSpr, 10, 10, C_BLUE, _
        “W e l c o m e T o U L T 2”

    fontSpr.ScaleFactor = 1
    PrintText fontImg, fontSpr, 10, 70, C_WHITE, _
        “abcdefghijklmnopqrstuvwxyz”
    PrintText fontImg, fontSpr, 10, 100, C_GRAY, _
        “ABCDEFGHIJKLMNOPQRSTUVWXYZ”
    PrintText fontImg, fontSpr, 10, 130, C_GREEN, _
        “!””#$%&()*+,-./0123456789:;<=>?@[\]^_{|}~”
    fontSpr.ScaleFactor = 3
    PrintText fontImg, fontSpr, 10, 180, C_RED, _
        “B I G H U N T !”

    fontSpr.ScaleFactor = 0.6
    PrintText fontImg, fontSpr, 10, 260, C_PURPLE, _
        “Tent”

    d3ddev.EndScene
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Shutdown
End Sub

Private Sub Form_Paint()
    d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub