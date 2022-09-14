‘-------------------------------------------------------
‘ MouseTest Program
‘-------------------------------------------------------
Option Explicit
Option Base 0

Implements DirectXEvent8

‘DirectX objects and structures
Public dx As New DirectX8
Public di As DirectInput8
Public diDev As DirectInputDevice8
Dim didevstate As DIMOUSESTATE

‘program constants and variables
Const BufferSize = 10
Public EventHandle As Long
Public Drawing As Boolean
Public Suspended As Boolean

Private Sub Form_Load()
    Me.Caption = “MouseTest”
    Me.Show

    ‘create the DirectInput object
    Set di = dx.DirectInputCreate

    ‘create the mouse object
    Set diDev = di.CreateDevice(“guid_SysMouse”)

    ‘configure DirectInputDevice to support the mouse
    Call diDev.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call diDev.SetCooperativeLevel(Form1.hWnd, DISCL_FOREGROUND Or _
         DISCL_EXCLUSIVE)

    ‘set properties for the mouse device
     Dim diProp As DIPROPLONG
     diProp.lHow = DIPH_DEVICE
     diProp.lObj = 0
     diProp.lData = BufferSize
    Call diDev.SetProperty(“DIPROP_BUFFERSIZE”, diProp)

    ‘create mouse callback event handler
    EventHandle = dx.CreateEvent(Form1)
    diDev.SetEventNotification EventHandle

    ‘acquire the mouse
    diDev.Acquire

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Shutdown
End Sub

Private Sub Shutdown()
    On Local Error Resume Next
    dx.DestroyEvent EventHandle
    Set diDev = Nothing
    Set di = Nothing
    Set dx = Nothing
    End
End Sub

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
    Dim diDeviceData(1 To BufferSize) As DIDEVICEOBJECTDATA
    Static lMouseX As Long
    Static lMouseY As Long
    Static lOldSeq As Long
    Dim n As Long

   ‘loop through events
   For n = 1 To diDev.GetDeviceData(diDeviceData, 0)
       Select Case diDeviceData(n).lOfs
           Case DIMOFS_X
               lMouseX = lMouseX + diDeviceData(n).lData
               If lMouseX < 0 Then lMouseX = 0
               If lMouseX >= Form1.ScaleWidth Then
                   lMouseX = Form1.ScaleWidth - 1
               End If

                If lOldSeq <> diDeviceData(n).lSequence Then
                    Debug.Print “MouseMove: “ & lMouseX & “ x “ & lMouseY
                    lOldSeq = diDeviceData(n).lSequence
                Else
                    lOldSeq = 0
                End If

Case DIMOFS_Y
    lMouseY = lMouseY + diDeviceData(n).lData
    If lMouseY < 0 Then lMouseY = 0
    If lMouseY >= Form1.ScaleHeight Then
        lMouseY = Form1.ScaleHeight - 1
    End If

    If lOldSeq <> diDeviceData(n).lSequence Then
        Debug.Print “Mouse: “ & lMouseX & “ x “ & lMouseY
        lOldSeq = diDeviceData(n).lSequence
    Else
        lOldSeq = 0
    End If

Case DIMOFS_BUTTON0
     If diDeviceData(n).lData > 0 Then   
         Debug.Print “ButtonDown 1”
     Else
        Debug.Print “ButtonUp 1”
     End If

Case DIMOFS_BUTTON1
    If diDeviceData(n).lData > 0 Then
         Debug.Print “ButtonDown 2”
    Else
        Debug.Print “ButtonUp 2”
    End If

    Case DIMOFS_BUTTON2
        If diDeviceData(n).lData > 0 Then
           Debug.Print “ButtonDown 3”
        Else
            Debug.Print “ButtonUp 3”
        End If
        
    End Select
  Next n
End Sub