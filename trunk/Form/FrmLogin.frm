VERSION 5.00
Object = "{9CAA1C67-43C4-4FFF-A005-20037C74BF32}#1.0#0"; "AlphaImageControl.ocx"
Object = "{A7960112-5DC4-4575-BFA3-DAD80FEE0438}#33.0#0"; "BasKomponen.ocx"
Begin VB.Form FrmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin BasKomponen.BasForm BasForm1 
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   10583
      ButtonMax       =   0   'False
      ButtonMin       =   0   'False
      Caption         =   "Login"
      Object.ToolTipText     =   "Login"
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4080
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox TxtID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4080
         TabIndex        =   4
         Top             =   840
         Width           =   3135
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
         TabIndex        =   2
         Top             =   840
         Width           =   2175
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
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000DDFF&
         BorderWidth     =   2
         Height          =   5295
         Left            =   120
         Top             =   600
         Width           =   7935
      End
      Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
         Height          =   7170
         Left            =   -240
         Top             =   -720
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   12647
         Image           =   "FrmLogin.frx":0000
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
Private Sub Form_Activate()
TxtID = "Login"
CekForm Me, TxtID
End Sub

