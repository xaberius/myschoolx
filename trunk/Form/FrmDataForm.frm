VERSION 5.00
Object = "{9CAA1C67-43C4-4FFF-A005-20037C74BF32}#1.0#0"; "AlphaImageControl.ocx"
Object = "{A7960112-5DC4-4575-BFA3-DAD80FEE0438}#33.0#0"; "BasKomponen.ocx"
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmDataForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BasKomponen.BasForm BasForm1 
      Height          =   5985
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10557
      ButtonMax       =   0   'False
      ButtonMin       =   0   'False
      Caption         =   ""
      Object.ToolTipText     =   ""
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   3705
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   6535
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Form ID"
         Columns(0).DataField=   "FormID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Form Name"
         Columns(1).DataField=   "FormName"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).PartialRightColumn=   0   'False
         Splits(0).MarqueeStyle=   2
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectorWidth=   529
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   8421376
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5689"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5583"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
         AllowUpdate     =   0   'False
         Appearance      =   2
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   2
         FootLines       =   1
         TabAction       =   2
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         CellTipsWidth   =   0
         GroupByCaption  =   "Keterangan"
         DeadAreaBackColor=   14215660
         RowDividerColor =   8454143
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H7DBBFF&,.bold=-1,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.borderColor=&H80000013&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&HFF8080&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&H80000005&,.fgcolor=&H0&,.bold=0"
         _StyleDefs(26)  =   ":id=13,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(27)  =   ":id=13,.fontname=Verdana"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.namedParent=37,.bgcolor=&H555555&"
         _StyleDefs(30)  =   ":id=14,.fgcolor=&H37D7FF&,.bold=-1,.fontsize=600,.italic=0,.underline=0"
         _StyleDefs(31)  =   ":id=14,.strikethrough=0,.charset=255"
         _StyleDefs(32)  =   ":id=14,.fontname=Terminal"
         _StyleDefs(33)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(34)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(35)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.fgcolor=&HFFFF&,.borderColor=&H80FF80&"
         _StyleDefs(36)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(37)  =   ":id=17,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(38)  =   ":id=17,.fontname=MS Sans Serif"
         _StyleDefs(39)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.fgcolor=&HFFFF&"
         _StyleDefs(40)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF&"
         _StyleDefs(41)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.namedParent=37,.bgcolor=&H80FFFF&"
         _StyleDefs(42)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(43)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0,.bgcolor=&HFF80&,.fgcolor=&HFFFFFF&,.borderColor=&H800040&"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   ":id=35,.wraptext=0,.locked=0"
         _StyleDefs(60)  =   "Named:id=36:Selected"
         _StyleDefs(61)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   ":id=36,.borderColor=&H80000013&"
         _StyleDefs(63)  =   "Named:id=37:Caption"
         _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&H80000009&"
         _StyleDefs(65)  =   "Named:id=38:HighlightRow"
         _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&HA00000&,.borderColor=&H800040&"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33,.bgcolor=&H4000&"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33,.bgcolor=&HFF0000&"
      End
      Begin BasTombol.vbButton CmdQuit 
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Quit"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDataForm.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BasTombol.vbButton CmdCancel 
         Height          =   375
         Left            =   6960
         TabIndex        =   6
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDataForm.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BasTombol.vbButton CmdSave 
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Save"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDataForm.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BasTombol.vbButton CmdDelete 
         Height          =   375
         Left            =   3120
         TabIndex        =   4
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Delete"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDataForm.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BasTombol.vbButton CmdEdit 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Edit"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDataForm.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BasTombol.vbButton CmdAdd 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Add"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDataForm.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtFormName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         TabIndex        =   1
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox TxtFormID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         TabIndex        =   0
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox TxtID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   330
         Left            =   2880
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Form Name"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Form ID"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000DDFF&
         BorderWidth     =   2
         Height          =   5295
         Left            =   120
         Top             =   600
         Width           =   9735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Form Of Forms"
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   5655
      End
      Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
         Height          =   5970
         Left            =   0
         Top             =   0
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   10530
         Image           =   "FrmDataForm.frx":00A8
         Scaler          =   1
         Props           =   5
      End
   End
End
Attribute VB_Name = "FrmDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Sub RefreshData()
Grid.DataSource = Nothing
SQL = "Select * from FormX"
Set Grid.DataSource = DbCon.Execute(SQL)
Grid.Refresh
End Sub

Public Sub CmdAdd_Click()
Edit = False
Form_Load
Tombol Me, False
TxtFormID.Locked = False
TxtFormName.Locked = False
TxtFormID.BackColor = vbWhite
TxtFormID.ForeColor = vbBlack
TxtFormID.SetFocus
End Sub

Private Sub CmdCancel_Click()
Form_Load
End Sub

Private Sub CmdDelete_Click()
Edit = False
Tombol Me, False
TxtFormID = Trim(Grid.Columns(0).Text)
TxtFormName = Trim(Grid.Columns(1).Text)
TxtFormID.Locked = True
TxtFormName.Locked = True

If MsgBox("Are You Sure to Delete This??", vbCritical + vbYesNo) = vbYes Then
    SQL = "delete from FormX where formid='" & Trim(TxtFormID.Text) & "'"
    DbCon.Execute SQL
    MsgBox "Data Deleted"
End If
RefreshData
End Sub

Private Sub CmdEdit_Click()
Grid_DblClick
End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
Dim Code As String
If Trim(TxtFormID.Text) = "" Then
    MsgBox "Form ID Still Blank"
    Exit Sub
ElseIf Trim(TxtFormName.Text) = "" Then
    MsgBox "Form Name Still Blank"
    Exit Sub
End If
    
Code = Trans.encryp_pass(25, Trim(TxtFormName.Text))
If Not Edit Then
    SQL = "insert into formX values('" & Trim(TxtFormID.Text) & "','" & Trim(TxtFormName.Text) & _
    "','" & Code & "')"
    DbCon.Execute SQL
    MsgBox "Data Saved"
    Form_Load
    RefreshData
Else
        SQL = "update FormX Set FormName='" & Trim(TxtFormName.Text) & "' where FormID='" & Trim(TxtFormID.Text) & "'"
        DbCon.Execute SQL
        MsgBox "Data Updated"
        Form_Load
        RefreshData
End If
End Sub

Private Sub Form_Activate()
CekForm Me, TxtID
End Sub

Private Sub Form_Load()
Tombol Me, True
TxtID = "A01-02-01"
Me.Width = Me.BasForm1.Width
Me.Height = Me.BasForm1.Height
Bersih
TxtFormID.Locked = True
TxtFormName.Locked = True
TxtFormID.BackColor = vbWhite
TxtFormID.ForeColor = vbBlack
Edit = False
RefreshData
End Sub

Sub Bersih()
TxtFormID = ""
TxtFormName = ""
End Sub

Private Sub Grid_DblClick()
Edit = True
Tombol Me, False
TxtFormID = Trim(Grid.Columns(0).Text)
TxtFormName = Trim(Grid.Columns(1).Text)
TxtFormID.Locked = True
TxtFormName.Locked = False
TxtFormID.BackColor = vbRed
TxtFormID.ForeColor = vbWhite
End Sub

Private Sub TxtFormID_Change()
SQL = "select * from FormX where FormID='" & Trim(TxtFormID) & "'"
Set RSFind = DbCon.Execute(SQL)
If RSFind.RecordCount > 0 Then
    MsgBox "Menu Already Exist!!", vbCritical
    Exit Sub
End If
End Sub

Private Sub TxtFormID_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtFormID_KeyPress(KeyAscii As Integer)
KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub TxtFormName_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub
