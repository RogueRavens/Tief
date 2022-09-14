‘-------------------------------------------------------
‘Sprite Support File
‘-------------------------------------------------------
Option Explicit
Option Base 0

‘sprite properties
Public Type TSPRITE
    spriteObject As D3DXSprite
    x As Long
    y As Long
    width As Long
    height As Long
    FramesPerRow As Long
    StartFrame As Long
    FrameCount As Long
    CurrentFrame As Long
    Animating As Boolean
    AnimSeq As Long
    AnimDelay As Long
    AnimCount As Long
    SpeedX As Long
    SpeedY As Long
    DirX As Long
    DirY As Long
    ScaleFactor As Single
End Type

*** Public Function LoadTexture ***

*** Public Sub InitSprite ***

*** Public Sub DrawSprite ***