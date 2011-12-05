VERSION 5.00
Object = "{A7960112-5DC4-4575-BFA3-DAD80FEE0438}#33.0#0"; "BasKomponen.ocx"
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmDataStudent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BasKomponen.BasForm BasForm1 
      Height          =   6480
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   11430
      Caption         =   "Student's Data"
      Object.ToolTipText     =   "Student's Data"
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   5265
         Left            =   9720
         TabIndex        =   16
         Top             =   480
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   9287
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ID"
         Columns(0).DataField=   "ID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Student's Name"
         Columns(1).DataField=   "Name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Address"
         Columns(2).DataField=   "Address"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "City"
         Columns(3).DataField=   "city"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
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
         _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(49)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0,.bgcolor=&HFF80&,.fgcolor=&HFFFFFF&,.borderColor=&H800040&"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   ":id=35,.wraptext=0,.locked=0"
         _StyleDefs(68)  =   "Named:id=36:Selected"
         _StyleDefs(69)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(70)  =   ":id=36,.borderColor=&H80000013&"
         _StyleDefs(71)  =   "Named:id=37:Caption"
         _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&H80000009&"
         _StyleDefs(73)  =   "Named:id=38:HighlightRow"
         _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&HA00000&,.borderColor=&H800040&"
         _StyleDefs(75)  =   "Named:id=39:EvenRow"
         _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(77)  =   "Named:id=40:OddRow"
         _StyleDefs(78)  =   ":id=40,.parent=33,.bgcolor=&H4000&"
         _StyleDefs(79)  =   "Named:id=41:RecordSelector"
         _StyleDefs(80)  =   ":id=41,.parent=34"
         _StyleDefs(81)  =   "Named:id=42:FilterBar"
         _StyleDefs(82)  =   ":id=42,.parent=33,.bgcolor=&HFF0000&"
      End
      Begin VB.TextBox TxtPAddress 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3480
         TabIndex        =   8
         Top             =   3480
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog CDI 
         Left            =   8640
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin BasTombol.vbButton vbButton1 
         Height          =   375
         Left            =   7440
         TabIndex        =   9
         Top             =   3240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Picture"
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
         MICON           =   "FrmDataStudent.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtCity 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3480
         TabIndex        =   3
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3480
         TabIndex        =   2
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox TxtPName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3480
         TabIndex        =   7
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox TxtSchool 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3480
         TabIndex        =   6
         Top             =   2760
         Width           =   3255
      End
      Begin TDBDate6Ctl.TDBDate BirthDate 
         Height          =   330
         Left            =   3480
         TabIndex        =   5
         Top             =   2400
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   582
         Calendar        =   "FrmDataStudent.frx":001C
         Caption         =   "FrmDataStudent.frx":0148
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmDataStudent.frx":01B4
         Keys            =   "FrmDataStudent.frx":01D2
         Spin            =   "FrmDataStudent.frx":0230
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "11/25/2011"
         ValidateMode    =   0
         ValueVT         =   1275068423
         Value           =   40872
         CenturyMode     =   0
      End
      Begin VB.TextBox TxtPlace 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3480
         TabIndex        =   4
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3480
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox TxtStudentID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3480
         TabIndex        =   0
         Top             =   600
         Width           =   3255
      End
      Begin BasTombol.vbButton CmdAdd 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   5880
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
         MICON           =   "FrmDataStudent.frx":0258
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
         TabIndex        =   11
         Top             =   5880
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
         MICON           =   "FrmDataStudent.frx":0274
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
         TabIndex        =   12
         Top             =   5880
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
         MICON           =   "FrmDataStudent.frx":0290
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
         Left            =   5640
         TabIndex        =   13
         Top             =   5880
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
         MICON           =   "FrmDataStudent.frx":02AC
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
         Left            =   7080
         TabIndex        =   14
         Top             =   5880
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
         MICON           =   "FrmDataStudent.frx":02C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin BasTombol.vbButton CmdQuit 
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         Top             =   5880
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
         MICON           =   "FrmDataStudent.frx":02E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent's Address"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3480
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   2535
         Left            =   7440
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent's name"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "School Before"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "911 Porscha"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000DDFF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmDataStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean

Private Sub CmdAdd_Click()
Tombol False
Edit = False
End Sub

Function CheckBlank() As Boolean
If Trim(TxtStudentID) = "" Then
    MsgBox "Student's ID Still Blank"
    TxtStudentID.SetFocus
    CheckBlank = False
ElseIf Trim(TxtName) = "" Then
    MsgBox "Student's Name Still Blank"
    TxtName.SetFocus
    CheckBlank = False
ElseIf Trim(TxtAddress) = "" Then
    MsgBox "Student's Addrees Still Blank"
    TxtAddress.SetFocus
    CheckBlank = False
ElseIf Trim(TxtCity) = "" Then
    MsgBox "Student's City Still Blank"
    TxtCity.SetFocus
    CheckBlank = False
ElseIf Year(BirthDate.Value) > Year(Date) - 5 Then
    MsgBox "Birth Date Invalid"
    BirthDate.SetFocus
    CheckBlank = False
ElseIf Trim(TxtPName) = "" Then
    MsgBox "Parent's Name Still Blank"
    TxtPName.SetFocus
    CheckBlank = False
ElseIf Trim(TxtPAddress) = "" Then
    MsgBox "Parent's Address Still Blank"
    TxtPAddress.SetFocus
    CheckBlank = False
ElseIf Trim(CDI.FileName) = "" Then
    MsgBox "Picture Still Empty"
    vbButton1.SetFocus
    CheckBlank = False
Else
    CheckBlank = True
End If
End Function

Private Sub CmdCancel_Click()
Form_Load
End Sub

Private Sub CmdDelete_Click()
If MsgBox("Are You Sure To Delete This??", vbCritical + vbYesNo) = vbYes Then
    FileSystem.Kill App.Path & "\Picture\" & Trim(Grid.Columns(0).Text) & ".jpg"
    SystemLog Me.Name, "Delete", "Delete Student Where Student ID = " & Trim(Grid.Columns(0).Text)
    SQL = "Delete from student where id='" & Trim(Grid.Columns(0).Text) & "'"
    DbCon.Execute SQL
    RefreshData
End If
    
End Sub

Private Sub CmdEdit_Click()
Grid_DblClick
End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
If Not CheckBlank Then Exit Sub
If Edit = False Then
    SQL = "insert into Student values('" & Trim(TxtStudentID) & "','" & Trim(TxtName) & "','" & _
            Trim(TxtAddress) & "','" & Trim(TxtCity) & "','" & Trim(TxtPlace) & _
            "','" & FormatTgl(BirthDate) & "','" & Trim(TxtSchool) & "','" & Trim(TxtPName) & "','" & _
            Trim(TxtPAddress) & "','" & Trim(CDI.FileName) & "')"
    DbCon.Execute SQL
    SavePicture Image1.Picture, App.Path & "\Picture\" & Trim(TxtStudentID) & ".jpg"
    SystemLog Me.Name, "Save", "Save New Student Where Student ID = " & Trim(TxtStudentID)
    MsgBox "Data Saved"
    Form_Load
ElseIf Edit Then
    SQL = "Update Student Set Name='" & Trim(TxtName) & "',address='" & Trim(TxtAddress) & _
        "',city='" & Trim(TxtCity) & "',place='" & Trim(TxtPlace) & "',date='" & FormatTgl(BirthDate) & _
        "',school='" & Trim(TxtSchool) & "',ParentsName='" & Trim(TxtPName) & _
        "', ParentsAddress='" & Trim(TxtPAddress) & "',Picture='" & Trim(CDI.FileName) & _
        "' where id='" & Trim(TxtStudentID) & " '"
    DbCon.Execute SQL
    FileSystem.Kill App.Path & "\Picture\" & Trim(TxtStudentID) & ".jpg"
    SavePicture Image1.Picture, App.Path & "\Picture\" & Trim(TxtStudentID) & ".jpg"
    SystemLog Me.Name, "Update", "Update Student Where Student ID = " & Trim(TxtStudentID)
    MsgBox "Data Updated"
    Form_Load
End If
    
End Sub

Private Sub Form_Load()
Me.Height = Me.BasForm1.Height
Me.Width = Me.BasForm1.Width
Me.Grid.Left = 120
Tombol True
ClearField
RefreshData
End Sub

Sub ClearField()
TxtStudentID = ""
TxtName = ""
TxtAddress = ""
TxtCity = ""
TxtPlace = ""
TxtSchool = ""
BirthDate.Value = Date - 1825
TxtPName = ""
TxtPAddress = ""
CDI.FileName = ""
Image1.Picture = Nothing
End Sub

Sub Tombol(Stat As Boolean)
CmdAdd.Visible = Stat
CmdEdit.Visible = Stat
CmdDelete.Visible = Stat

CmdSave.Visible = Not Stat
CmdCancel.Visible = Not Stat
Grid.Visible = Stat
End Sub

Sub RefreshData()
Grid.DataSource = Nothing
SQL = "Select id,name,address,city,ParentsName,ParentsAddress from Student order by ID"
Set Grid.DataSource = DbCon.Execute(SQL)
Grid.Refresh
End Sub

Private Sub Grid_DblClick()
Edit = True
SQL = "Select Name,address,city,place,date,school,ParentsName,ParentsAddress,picture from Student where id='" & _
        Trim(Grid.Columns(0).Text) & "'"
Set RSFind = DbCon.Execute(SQL)

TxtStudentID = Trim(Grid.Columns(0).Text)
TxtName = Trim(RSFind!Name)
TxtAddress = Trim(RSFind!address)
TxtCity = Trim(RSFind!city)
TxtPlace = Trim(RSFind!place)
BirthDate = RSFind!Date
TxtSchool = Trim(RSFind!School)
TxtPName = Trim(RSFind!ParentsName)
TxtPAddress = Trim(RSFind!parentsAddress)
CDI.FileName = Trim(RSFind!Picture)
Image1.Picture = LoadPicture(CDI.FileName)

Tombol False
End Sub

Private Sub TxtStudentID_Change()
If Len(Trim(TxtStudentID)) < 2 Then
    SQL = "Select ID from Student where id='" & Trim(TxtStudentID) & "'"
    Set RSFind = DbCon.Execute(SQL)
    If RSFind.RecordCount > 0 Then
        MsgBox "Student ID Already Exist. Try Again."
        Exit Sub
    End If
End If

End Sub

Private Sub vbButton1_Click()
With CDI
    .DialogTitle = "Picture"
    .Filter = "JPEG|*.jpg"
    .ShowOpen

    If .FileName <> "" Then
        Set Me.Image1.Picture = Nothing
        Me.Image1.Picture = LoadPicture(.FileName)
    End If
End With
End Sub
