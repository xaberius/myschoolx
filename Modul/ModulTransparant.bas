Attribute VB_Name = "ModulTransparant"
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function ReleaseCapture Lib "user32" () As Long
    Private Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
    Private Declare Function SetWindowRgn Lib "user32" _
    (ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long

Sub MakeTaskbarTransparent(LhWnd As Long, ByVal bLevel As Byte)
On Error GoTo Salah
    Dim lOldStyle As Long
    
    If (LhWnd <> 0) Then
        lOldStyle = GetWindowLong(LhWnd, (-20))
        SetWindowLong LhWnd, (-20), lOldStyle Or &H80000
        SetLayeredWindowAttributes LhWnd, 0, bLevel, &H2&
    End If
Salah:
End Sub

Sub FadeForm(hWnd As Long, Optional FadeOut As Boolean = True)
Dim I As Integer
If FadeOut Then
   For I = 1 To 15
     DoEvents
     MakeTaskbarTransparent hWnd, I * 17
     Sleep 100
   Next I
Else
    For I = 1 To 15
      DoEvents
      MakeTaskbarTransparent hWnd, Int(255 / I)
      Sleep 50
    Next I
End If
End Sub

Sub CekForm(Form As Form, TxtID As String)
SQL = "select codes from FormX where formName='" & Form.Name & "'"
Set RSFind = DbCon.Execute(SQL)
If RSFind.RecordCount = 0 Then
    Form.Hide
    MsgBox "Form Is Not Valid!!!", vbCritical
    FrmDataForm.Show
    FrmDataForm.CmdAdd_Click
    FrmDataForm.TxtFormID = TxtID
    FrmDataForm.TxtFormID.Locked = True
    FrmDataForm.TxtFormName = Form.Name
    Unload Form
End If
End Sub


Sub Tombol(Form As Form, Stat As Boolean)
With Form
    .CmdAdd.Enabled = Stat
    .CmdEdit.Enabled = Stat
    .CmdDelete.Enabled = Stat
    .CmdSave.Enabled = Not Stat
    .CmdCancel.Enabled = Not Stat
End With
End Sub
