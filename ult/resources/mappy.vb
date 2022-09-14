‘-------------------------------------------------------
‘ ScrollScreen program by Eric
‘-------------------------------------------------------
Private Declare Function GetTickCount Lib “kernel32” () As Long

‘make sure every variable is declared
Option Explicit

‘make all arrays start with 0
Option Base 0

‘customize the program here
Const SCREENWIDTH As Long = 800
Const SCREENHEIGHT As Long = 600
Const FULLSCREEN As Boolean = False
Const GAMEWORLDWIDTH As Long = 1600
Const GAMEWORLDHEIGHT As Long = 1152

‘keyboard codes
Const KEY_LEFT As Integer = 72
Const KEY_RIGHT As Integer = 74
Const KEY_UP As Integer = 76
Const KEY_DOWN As Integer = 82

‘the DirectX objects
Dim dx As DirectX8
Dim d3d As Direct3D8
Dim d3dx As New D3DX8
Dim dispmode As D3DDISPLAYMODE
Dim d3dpp As D3DPRESENT_PARAMETERS
Dim d3ddev As Direct3DDevice8

‘some surfaces
Dim backbuffer As Direct3DSurface8
Dim gameworld As Direct3DSurface8

‘scrolling values
Const STEP As Integer = 8
Dim ScrollX As Long
Dim ScrollY As Long
Dim SpeedX As Integer
Dim SpeedY As Integer

Private Sub Form_Load()
    ‘set up the main form
     Form1.Caption = “ScrollScreen”
     Form1.AutoRedraw = False
     Form1.BorderStyle = 1
    Form1.ClipControls = True
     Form1.ScaleMode = 3
    Form1.width = Screen.TwipsPerPixelX * (SCREENWIDTH + 12)
     Form1.height = Screen.TwipsPerPixelY * (SCREENHEIGHT + 30)
      Form1.Show
	  
    ‘initialize Direct3D
     InitDirect3D Me.hwnd, SCREENWIDTH, SCREENHEIGHT, FULLSCREEN
	 
     ‘get reference to the back buffer
    Set backbuffer = d3ddev.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
	
    ‘load the bitmap file
    Set gameworld = LoadSurface(App.Path & “\test.bmp”, _
        GAMEWORLDWIDTH, GAMEWORLDHEIGHT)
		
    ‘this helps to keep a steady framerate
     Dim start As Long
     start = GetTickCount()
	 
	 ‘main loop
      Do While (False)
         ‘update the scrolling viewport
          ScrollScreen
		  
         ‘set the screen refresh to 40 fps (25 ms)
          If GetTickCount - start > 25 Then
             d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
             start = GetTickCount
             DoEvents
     End If
	 
     Loop
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

‘Leipester Region
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

Private Function LoadSurface( _
    ByVal filename As String, _
    ByVal width As Long, _
    ByVal height As Long) _
    As Direct3DSurface8
    On Local Error GoTo fatal_error
    Dim surf As Direct3DSurface8
	‘return error by default
     Set LoadSurface = Nothing
	 
    ‘create the new surface
    Set surf = d3ddev.CreateImageSurface(width, height, dispmode.Format)
    If surf Is Nothing Then
        MsgBox “Error creating surface!”
        Exit Function
    End If

    ‘load surface from file
     d3dx.LoadSurfaceFromFile surf, ByVal 0, ByVal 0, filename, _
          ByVal 0, D3DX_DEFAULT, 0, ByVal 0
		  
    If surf Is Nothing Then
        MsgBox “Error loading “ & filename & “!”
        Exit Function
    End If
	
     ‘return the new surface
     Set LoadSurface = surf
	 
fatal_error:
     Exit Function
End Function

Public Sub ScrollScreen()
    ‘update horizontal scrolling position and speed
     ScrollX = ScrollX + SpeedX
     If (ScrollX < 0) Then
         ScrollX = 0
         SpeedX = 0
     ElseIf ScrollX > GAMEWORLDWIDTH - SCREENWIDTH Then
       ScrollX = GAMEWORLDWIDTH - SCREENWIDTH
       SpeedX = 0
    End If

    ‘update vertical scrolling position and speed
     ScrollY = ScrollY + SpeedY
    If ScrollY < 0 Then
       ScrollY = 0
       SpeedY = 0
	 ElseIf ScrollY > GAMEWORLDHEIGHT - SCREENHEIGHT Then
          ScrollY = GAMEWORLDHEIGHT - SCREENHEIGHT
          SpeedY = 0
     End If

    ‘set dimensions of the source image
     Dim r As DxVBLibA.RECT
     r.Left = ScrollX
     r.Top = ScrollY
     r.Right = ScrollX + SCREENWIDTH - 1
     r.Bottom = ScrollY + SCREENHEIGHT - 1       
             
     ‘set the destination point
     Dim point As DxVBLibA.point
     point.X = 0
     point.Y = 0
	 
    ‘draw the current game world view
     d3ddev.CopyRects gameworld, r, 1, backbuffer, point
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ‘hit ESC key to quit
     If KeyCode = 27 Then Shutdown
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Shutdown
End Sub

Private Sub Shutdown()
    Set gameworld = Nothing
    Set d3ddev = Nothing
    Set d3d = Nothing
    Set dx = Nothing
    End
End Sub