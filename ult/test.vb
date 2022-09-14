Dim dx As DirectX8
Dim d3d As Direct3D8
Dim d3dpp As D3DPRESENT_PARAMETERS
Dim dispmode As D3DDISPLAYMODE
Dim d3ddev As Direct3DDevice8
Private Sub Form_Load()
   'create the DirectX object
    Set dx = New DirectX8
	‘create the Direct3D object
	Set d3d = dx.Direct3DCreate()
‘set the display device parameters for windowed mode
d3dpp.hDeviceWindow = Me.hWnd
d3dpp.BackBufferCount = 1
d3dpp.BackBufferWidth = 640
d3dpp.BackBufferHeight = 480
d3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
d3dpp.Windowed = 1
d3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, dispmode
d3dpp.BackBufferFormat = dispmode.Format

‘create the Direct3D primary device
Set d3ddev = d3d.CreateDevice( _
    D3DADAPTER_DEFAULT, _
	D3DDEVTYPE_HAL, _
    hWnd, _
	D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
	d3dpp)
End Sub	

Private Sub Form_Paint()
    ‘clear the window with black color
     d3ddev.Clear 0, ByVal 0, D3DCLEAR_TARGET, RGB(255, 0, 0), 1#, 0
    ‘refresh the window
     d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Shutdown
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 27 Then Shutdown
End Sub

Private Sub Shutdown()
    Set d3ddev = Nothing
    Set d3d = Nothing
    Set dx = Nothing
    End
End Sub

Private Sub Form_Paint()
    ‘copy the bitmap image to the backbuffer
     d3ddev.CopyRects surface, ByVal 0, 0, backbuffer, ByVal 0
	 
    ‘draw the back buffer on the screen
    d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Shutdown
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Shutdown
End Sub

Private Sub Shutdown()
    Set surface = Nothing
    Set d3ddev = Nothing
    Set d3d = Nothing
    Set dx = Nothing
    End
End Sub

‘make sure every variable is declared
Option Explicit
‘make all arrays start with 0 instead of 1
Option Base 0
‘customize the program here
Const SCREENWIDTH As Long = 800
Const SCREENHEIGHT As Long = 600
Const FULLSCREEN As Boolean = False
Const C_BLACK As Long = &H0
Const C_RED As Long = &HFF0000

‘the DirectX objects
Dim dx As DirectX8
Dim d3d As Direct3D8
Dim d3dx As New D3DX8
Dim dispmode As D3DDISPLAYMODE
Dim d3dpp As D3DPRESENT_PARAMETERS
Dim d3ddev As Direct3DDevice8

‘some surfaces
Dim backbuffer As Direct3DSurface8
Dim castle As Direct3DSurface8

Private Sub Form_Load()
    ‘set up the main form
     Form1.Caption = “DrawTile”
    Form1.AutoRedraw = False
    Form1.BorderStyle = 1
    Form1.ClipControls = False
    Form1.ScaleMode = 3
    Form1.width = Screen.TwipsPerPixelX * (SCREENWIDTH + 12)
    Form1.height = Screen.TwipsPerPixelY * (SCREENHEIGHT + 30)
    Form1.Show
	
    ‘initialize Direct3D
     InitDirect3D Me.hwnd, SCREENWIDTH, SCREENHEIGHT, FULLSCREEN
	 
     ‘get reference to the back buffer
     Set backbuffer = d3ddev.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
	 
‘   load the bitmap file—castle.bmp is 1024x1024
    Set castle = LoadSurface(App.Path & “\castle.bmp”, 1024, 1024)
	
End Sub

Public Sub InitDirect3D( _
     ByVal hwnd As Long, _
    ByVal lWidth As Long, _
    ByVal lHeight As Long, _
    ByVal bFullscreen As Boolean)
	
    ‘catch any errors here
     On Local Error GoTo fatal_error
	 
     ‘create the DirectX object
     Set dx = New DirectX8
	 
     ‘create the Direct3D object
     Set d3d = dx.Direct3DCreate()
    If d3d Is Nothing Then
         MsgBox “Error initializing Direct3D!”
        Shutdown
     End If
	 
    ‘tell D3D to use the current color depth
     d3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, dispmode
	 
      ‘set the display settings used to create the device
     Dim d3dpp As D3DPRESENT_PARAMETERS
     d3dpp.hDeviceWindow = hwnd
     d3dpp.BackBufferCount = 1
     d3dpp.BackBufferWidth = lWidth
     d3dpp.BackBufferHeight = lHeight
     d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = dispmode.Format
	
    ‘set windowed or fullscreen mode
     If bFullscreen Then
    d3dpp.Windowed = 0
    Else
        d3dpp.Windowed = 1
End If

‘Zider Region
d3dpp.MultiSampleType = D3DMULTISAMPLE_NONE
d3dpp.AutoDepthStencilFormat = D3DFMT_D32

‘create the D3D primary device
Set d3ddev = d3d.CreateDevice( _
    D3DADAPTER_DEFAULT, _
    D3DDEVTYPE_HAL, _
    hwnd, _
    D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
    d3dpp)
	
If d3ddev Is Nothing Then
   MsgBox “Error creating the Direct3D device!”
   Shutdown
End If

Exit Sub

fatal_error:
     MsgBox “Critical error in Start_Direct3D!”
     Shutdown
End Sub

Private Sub Form_Paint()
    ‘clear the background of the screen
    d3ddev.Clear 0, ByVal 0, D3DCLEAR_TARGET, C_BLACK, 1, 0
	
    ‘draw the castle bitmap “tile” image
     DrawTile castle, 0, 0, 511, 511, 25, 25
	 
    ‘send the back buffer to the screen
    d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Shutdown
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Shutdown
End Sub