VERSION 5.00
Object = "{9CAA1C67-43C4-4FFF-A005-20037C74BF32}#1.0#0"; "AlphaImageControl.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form FrmLoading 
   BorderStyle     =   0  'None
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   DrawStyle       =   1  'Dash
   Icon            =   "FrmLoading.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin ARProgBarCtrl.ARProgressBar Bar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   661
      Value           =   2
      ForeColor       =   4210752
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseGradient     =   -1  'True
      IniColor        =   8454016
      EndColor        =   49152
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   3720
      Top             =   720
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   4471
      Image           =   "FrmLoading.frx":014A
      Scaler          =   1
      Props           =   5
   End
End
Attribute VB_Name = "FrmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Fade1 Me
End Sub

Private Sub Form_Load()
Me.Caption = "Loading..."
Bar.Value = 0
Me.Height = Me.aicAlphaImage1.Height
Me.Width = Me.aicAlphaImage1.Width
Me.Timer1.Enabled = True
Fade2 Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
FadeForm Me.hWnd, False
End Sub

Private Sub Timer1_Timer()

If Bar.Value = 100 Then
    Timer1.Enabled = False
    Unload Me
    FrmMainAdmin.Show
ElseIf Bar.Value = 50 Then
    Bar.CaptionForeColor = vbWhite
    Bar.Value = Val(Bar.Value) + 1
Else
    Bar.Value = Val(Bar.Value) + 1
End If
End Sub
