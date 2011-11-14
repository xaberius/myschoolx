Attribute VB_Name = "ModuleFade"
Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Sub Fade1(Form As Form)
Dim Alpha As Byte
Dim Delay As Single
Do
    Delay = Timer
    Do While (Timer - Delay) < 0.01
        DoEvents
    Loop
    Alpha = Alpha + 5
    SetLayeredWindowAttributes Form.hWnd, 0, Alpha, LWA_ALPHA
Loop Until Alpha = 255
End Sub

Sub Fade2(Form As Form)
    Dim Ret As Long
    Ret = GetWindowLong(Form.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Form.hWnd, GWL_EXSTYLE, Ret
End Sub
