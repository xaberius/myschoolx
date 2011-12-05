VERSION 5.00
Object = "{9CAA1C67-43C4-4FFF-A005-20037C74BF32}#1.0#0"; "AlphaImageControl.ocx"
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SmartMenuXP.ocx"
Begin VB.Form FrmMainAdministration 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14280
   Icon            =   "FrmMainAdministration.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin VBSmartXPMenu.SmartMenuXP SmartMenuXP1 
      Height          =   375
      Left            =   0
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   16761024
      BorderStyle     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   9840
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Administration's Menu"
      BeginProperty Font 
         Name            =   "Minion Pro SmBd"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11040
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin AlphaImageControl.aicAlphaImage LogoUtama 
      Height          =   3105
      Left            =   4920
      Top             =   3120
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   5477
      Image           =   "FrmMainAdministration.frx":014A
      Props           =   5
   End
   Begin AlphaImageControl.aicAlphaImage BannerUtama 
      Height          =   690
      Left            =   -240
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   1217
      Image           =   "FrmMainAdministration.frx":58615
      Scaler          =   1
      Angle           =   30
      Props           =   5
   End
   Begin AlphaImageControl.aicAlphaImage BackGroundUtama 
      Height          =   8940
      Left            =   0
      Top             =   600
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   15769
      Image           =   "FrmMainAdministration.frx":91B97
      Scaler          =   1
      Opacity         =   70
      Props           =   5
   End
End
Attribute VB_Name = "FrmMainAdministration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
CekForm Me, TxtID
End Sub

Private Sub Form_Load()
TxtID = "Main Administration"
Me.Width = Me.Width + 1000
Me.Height = Me.Height + 1000
Me.BackGroundUtama.Height = Me.Height
Me.BackGroundUtama.Width = Me.Width
Me.BannerUtama.Width = Me.Width + 1000
Me.LogoUtama.Left = Me.LogoUtama.Left + 100

With SmartMenuXP1.MenuItems
        .Add 0, "mnuServer", , "&Logout   "
        .Add "mnuServer", "mnuLogin", , "&Logout"
        .Add "mnuServer", "mnuExit", , "&Exit"
        
        .Add 0, "mnuMaster", , "&Master   "
        .Add "mnuMaster", "mnuDataStudent", , "&Data Student"
        
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If User.UserType = "0001" Then
    FadeForm Me.hWnd, False
    FrmChoise.Show
Else
    FadeForm Me.hWnd, False
End If
End Sub

Private Sub SmartMenuXP1_Click(ByVal ID As Long)
With SmartMenuXP1.MenuItems
        Select Case .Key(ID)
            Case "mnuLogin": SystemLog Me.Name, "Logout", "User Logout": Unload Me: FrmLogin.Show
            Case "mnuExit": SystemLog Me.Name, "Logout", "User Logout": Unload Me
            Case "mnuDataStudent": FrmDataStudent.Show , FrmMainAdministration
            
            Case "mnuDataUser": FrmUser.Show , FrmMainAdmin
            Case "mnuUserType": FrmUserType.Show , FrmMainAdmin
            Case "mnuDataUser": FrmUser.Show , FrmMainAdmin
            Case "mnuSysLog": FrmLogSystem.Show , FrmMainAdmin
            Case "mnuDataUser": FrmUser.Show , FrmMainAdmin
            
            
        End Select
End With
End Sub


