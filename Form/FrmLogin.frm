VERSION 5.00
Object = "{9CAA1C67-43C4-4FFF-A005-20037C74BF32}#1.0#0"; "AlphaImageControl.ocx"
Object = "{A7960112-5DC4-4575-BFA3-DAD80FEE0438}#33.0#0"; "BasKomponen.ocx"
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form FrmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BasKomponen.BasForm BasForm1 
      Height          =   3000
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5292
      ButtonMax       =   0   'False
      ButtonMin       =   0   'False
      Caption         =   "Login"
      Object.ToolTipText     =   "Login"
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   30
         Left            =   2880
         Top             =   840
      End
      Begin ARProgBarCtrl.ARProgressBar Bar 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Value           =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowCaption     =   0   'False
         UseGradient     =   -1  'True
         EndColor        =   16761024
      End
      Begin BasTombol.vbButton vbButton2 
         Height          =   495
         Left            =   3480
         TabIndex        =   2
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Login"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmLogin.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BasTombol.vbButton vbButton1 
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmLogin.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro R"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   3480
         PasswordChar    =   "X"
         TabIndex        =   1
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox TxtUserID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Kozuka Mincho Pro R"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         Left            =   3480
         TabIndex        =   0
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox TxtID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   3480
         TabIndex        =   7
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Dodger"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Dodger"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Form Login"
         BeginProperty Font 
            Name            =   "Dodger"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000DDFF&
         BorderWidth     =   2
         Height          =   2415
         Left            =   120
         Top             =   480
         Width           =   6735
      End
      Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
         Height          =   7170
         Left            =   -2160
         Top             =   -2520
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   12647
         Image           =   "FrmLogin.frx":0038
         Scaler          =   1
         Opacity         =   90
         Props           =   5
      End
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Try As Integer

Private Sub Form_Activate()
TxtPassword.SetFocus
TxtID = "Login"
CekForm Me, TxtID
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub Form_Load()
Try = 0
Me.Height = Me.BasForm1.Height
Me.Width = Me.BasForm1.Width
TxtUserID = GetSetting(App.Title, "startup", "login")
TxtUserID.SelLength = Len(TxtUser)
End Sub

Private Sub Timer1_Timer()
Dim A As Boolean

If Try = 3 Then
    MsgBox "You Are Not Authorized!!!", vbCritical
    Unload Me
    End
End If

Bar.Value = 0
While Bar.Value < 100
    Bar.Value = Val(Bar.Value) + 2
    If Bar.Value = 50 Then
        SQL = "select Password,UserType from UserX where UserID='" & Trim(TxtUserID.Text) & "' "
        Set RSFind = DbCon.Execute(SQL)

        If RSFind.RecordCount = 0 Then
            MsgBox "User Is Not Registed!", vbCritical
            TxtUserID = ""
            TxtUserID.SetFocus
            Timer1.Enabled = False
            Bar.Value = 0
            Try = Try + 1
            Exit Sub
        ElseIf Trim(Trans.decryp_pass(25, RSFind!Password)) <> Trim(TxtPassword) Then
            MsgBox "Password Is Not Registed!", vbCritical
            TxtPassword = ""
            TxtPassword.SetFocus
            Timer1.Enabled = False
            Bar.Value = 0
            Try = Try + 1
            Exit Sub
        Else
            A = True
        End If
        Bar.CaptionForeColor = vbWhite
        Bar.Value = Val(Bar.Value) + 2
    End If
    
Wend
If Bar.Value = 100 Then
        Timer1.Enabled = False
        If A Then
            SaveSetting App.Title, "startup", "login", Trim(TxtUserID)
            User.UserId = Trim(TxtUserID)
            User.UserType = Trim(RSFind!UserType)
            If Trim(RSFind!UserType) = "0001" Then
                SystemLog Me.Name, "Login", Trim(User.UserId) & " Enter System."
                Unload Me
                FrmChoise.Show
                'FrmMainAdmin.Show
            ElseIf Trim(RSFind!UserType) = "0002" Then
                SystemLog Me.Name, "Login", Trim(User.UserId) & " Enter System."
                Unload Me
                FrmMainAdministration.Show
            Else
                MsgBox "User Type Is Not Valid"
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub TxtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub vbButton1_Click()
Unload Me
End Sub

Private Sub vbButton1_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub vbButton2_Click()
If Trim(TxtUserID) = "" Then
    MsgBox "User Still Blank"
    TxtID.SetFocus
    Exit Sub
ElseIf Trim(TxtPassword) = "" Then
    MsgBox "Password Still Blank"
    TxtPassword.SetFocus
    Exit Sub
End If

    Bar.Visible = True
    Timer1.Enabled = True
End Sub

Private Sub vbButton2_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

