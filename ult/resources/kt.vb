‘-------------------------------------------------------
‘ KeyboardTest Program
‘-------------------------------------------------------
Option Explicit
Option Base 0

‘Windows API functions and structures
Private Declare Function GetTickCount Lib “kernel32” () As Long

‘program variables
Dim dx As New DirectX8
Dim dinput As DirectInput8
Dim diDevice As DirectInputDevice8
Dim diState As DIKEYBOARDSTATE
Dim sKeyNames(255) As String

Private Sub Form_Load()
    ‘set up the form
    Form1.Caption = “KeyboardTest”
    Form1.Show

    Set dinput = dx.DirectInputCreate()
    If Err.Number <> 0 Then
         MsgBox “Error creating DirectInput object”
         End
    End If

    ‘create an interface to the keyboard
     Set diDevice = dinput.CreateDevice(“GUID_SysKeyboard”)
     diDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
     diDevice.SetCooperativeLevel hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
     diDevice.Acquire

     ‘initialize the keyboard value array
     InitKeyNames

     ‘main game loop
     Do While True
         Check_Keyboard
         DoEvents
     Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Shutdown
End Sub

Public Sub Check_Keyboard()
    Dim n As Long

    ‘get the list of pressed keys
    diDevice.GetDeviceStateKeyboard diState

    ‘scan the entire list for pressed keys
    For n = 0 To 255
        If diState.Key(n) > 0 Then
            Debug.Print n & “ = “ & sKeyNames(n)
        End If
    Next

    ‘check for ESC key
    If diState.Key(1) > 0 Then Shutdown

End Sub


Public Sub Shutdown()
    diDevice.Unacquire
    Set diDevice = Nothing
    Set dinput = Nothing
    Set dx = Nothing
    End
End Sub

Private Sub InitKeyNames()
    sKeyNames(1) = “ESC”
    sKeyNames(2) = “1”
    sKeyNames(3) = “2”
    sKeyNames(4) = “3”
    sKeyNames(5) = “4”
    sKeyNames(6) = “5”
    sKeyNames(7) = “6”
    sKeyNames(8) = “7”
    sKeyNames(9) = “8”
    sKeyNames(10) = “9”
    sKeyNames(11) = “0”
    sKeyNames(12) = “-”
    sKeyNames(13) = “=”
    sKeyNames(14) = “BACKSPACE”
    sKeyNames(15) = “TAB”
    sKeyNames(16) = “Q”
    sKeyNames(17) = “W”
    sKeyNames(18) = “E”
    sKeyNames(19) = “R”
    sKeyNames(20) = “T”
    sKeyNames(21) = “Y”
    sKeyNames(22) = “U”
    sKeyNames(23) = “I”
    sKeyNames(24) = “O”
    sKeyNames(25) = “P”
    sKeyNames(26) = “[“
    sKeyNames(27) = “ ]”
    sKeyNames(28) = “ENTER”
    sKeyNames(29) = “LCTRL”
    sKeyNames(30) = “A”
    sKeyNames(31) = “S”
    sKeyNames(32) = “D”
    sKeyNames(33) = “F”
    sKeyNames(34) = “G”
    sKeyNames(35) = “H”
    sKeyNames(36) = “J”
    sKeyNames(37) = “K”
    sKeyNames(38) = “L”
    sKeyNames(39) = “;”
    sKeyNames(40) = “‘“
    sKeyNames(41) = “`”
    sKeyNames(42) = “LSHIFT”
    sKeyNames(43) = “\”
    sKeyNames(44) = “Z”
    sKeyNames(45) = “X”
    sKeyNames(46) = “C”