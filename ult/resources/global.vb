‘-------------------------------------------------------
‘ Globals File
‘-------------------------------------------------------
Option Explicit
Option Base 0

‘Windows API functions
Public Declare Function GetTickCount Lib “kernel32” () As Long

‘colors
Public Const C_BLACK As Long = &H0

‘customize the program here
Public Const FULLSCREEN As Boolean = True
Public Const SCREENWIDTH As Long = 800
Public Const SCREENHEIGHT As Long = 600
Public Const STEP As Integer = 8

‘game world size
Public Const GAMEWORLDWIDTH As Long = 1600
Public Const GAMEWORLDHEIGHT As Long = 1152

‘tile size
Public Const TILEWIDTH As Integer = 64
Public Const TILEHEIGHT As Integer = 64

‘scrolling window size
Public Const WINDOWWIDTH As Integer = (SCREENWIDTH \ TILEWIDTH) * TILEWIDTH
Public Const WINDOWHEIGHT As Integer = (SCREENHEIGHT \ TILEHEIGHT) * TILEHEIGHT

‘scroll buffer size
Public Const SCROLLBUFFERWIDTH As Integer = SCREENWIDTH + TILEWIDTH
Public Const SCROLLBUFFERHEIGHT As Integer = SCREENHEIGHT + TILEHEIGHT