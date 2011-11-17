VERSION 5.00
Object = "{9CAA1C67-43C4-4FFF-A005-20037C74BF32}#1.0#0"; "AlphaImageControl.ocx"
Object = "{A7960112-5DC4-4575-BFA3-DAD80FEE0438}#33.0#0"; "BasKomponen.ocx"
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FrmUser 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BasKomponen.BasForm BasForm1 
      Height          =   6945
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12250
      ButtonMax       =   0   'False
      ButtonMin       =   0   'False
      Caption         =   ""
      Object.ToolTipText     =   ""
      Begin VB.TextBox TxtUserPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "X"
         TabIndex        =   2
         Top             =   1440
         Width           =   3255
      End
      Begin MSAdodcLib.Adodc AdoType 
         Height          =   330
         Left            =   8400
         Top             =   1800
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbPosition 
         Height          =   345
         Left            =   2880
         TabIndex        =   3
         Tag             =   "Kode"
         Top             =   1800
         Width           =   3255
         DataFieldList   =   "Column 0"
         BevelType       =   0
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelColorHighlight=   -2147483634
         BevelColorFace  =   -2147483627
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   8454143
         BackColorOdd    =   65535
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5741
         _ExtentY        =   609
         _StockProps     =   93
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxtUserID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         TabIndex        =   0
         Top             =   720
         Width           =   3255
      End
      Begin BasTombol.vbButton CmdAdd 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   6360
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
         MICON           =   "FrmUser.frx":0000
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
         TabIndex        =   5
         Top             =   6360
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
         MICON           =   "FrmUser.frx":001C
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
         TabIndex        =   6
         Top             =   6360
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
         MICON           =   "FrmUser.frx":0038
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
         TabIndex        =   7
         Top             =   6360
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
         MICON           =   "FrmUser.frx":0054
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
         TabIndex        =   8
         Top             =   6360
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
         MICON           =   "FrmUser.frx":0070
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
         Left            =   8400
         TabIndex        =   9
         Top             =   6360
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
         MICON           =   "FrmUser.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   3705
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   6535
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "User Number"
         Columns(0).DataField=   "UserNumber"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "User ID"
         Columns(1).DataField=   "UserID"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "User Name"
         Columns(2).DataField=   "UserName"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "User Type"
         Columns(3).DataField=   "UserType"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Created Date"
         Columns(4).DataField=   "CreatedDate"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
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
         Splits(0)._ColumnProps(26)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
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
         _StyleDefs(44)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(45)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
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
         _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0,.bgcolor=&HFF80&,.fgcolor=&HFFFFFF&,.borderColor=&H800040&"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   ":id=35,.wraptext=0,.locked=0"
         _StyleDefs(72)  =   "Named:id=36:Selected"
         _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   ":id=36,.borderColor=&H80000013&"
         _StyleDefs(75)  =   "Named:id=37:Caption"
         _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&H80000009&"
         _StyleDefs(77)  =   "Named:id=38:HighlightRow"
         _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&HA00000&,.borderColor=&H800040&"
         _StyleDefs(79)  =   "Named:id=39:EvenRow"
         _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(81)  =   "Named:id=40:OddRow"
         _StyleDefs(82)  =   ":id=40,.parent=33,.bgcolor=&H4000&"
         _StyleDefs(83)  =   "Named:id=41:RecordSelector"
         _StyleDefs(84)  =   ":id=41,.parent=34"
         _StyleDefs(85)  =   "Named:id=42:FilterBar"
         _StyleDefs(86)  =   ":id=42,.parent=33,.bgcolor=&HFF0000&"
      End
      Begin VB.TextBox TxtUserName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         TabIndex        =   1
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox TxtID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   330
         Left            =   2880
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox TxtUserNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   330
         Left            =   2880
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   17
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Form Of User"
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
         TabIndex        =   15
         Top             =   120
         Width           =   5295
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000DDFF&
         BorderWidth     =   2
         Height          =   6255
         Left            =   120
         Top             =   600
         Width           =   9735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         TabIndex        =   13
         Top             =   1080
         Width           =   2295
      End
      Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
         Height          =   7050
         Left            =   0
         Top             =   0
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   12435
         Image           =   "FrmUser.frx":00A8
         Scaler          =   1
         Props           =   5
      End
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean

Private Sub CmbPosition_DropDown()
AdoType.RecordSource = ""
SQL = "Select TypeID,TypeName from UserType order by TypeID "
Set RSFind = DbCon.Execute(SQL)
If RSFind.BOF Then Exit Sub
AdoType.RecordSource = SQL
AdoType.Refresh
With CmbPosition
    .DataSourceList = AdoType
    .DataFieldList = "TypeName"
    .Columns(0).Visible = False
    .Columns(1).Width = 5000
End With
End Sub

Private Sub CmdAdd_Click()
Edit = False
Tombol Me, False
Bersih
TxtUserID.Locked = 0
TxtUserName.Locked = 0
TxtUserPassword.Locked = 0
CmbPosition.Enabled = 1
End Sub
Sub RefreshData()
Set Grid.DataSource = Nothing
SQL = "select a.UserID,a.UserName,a.CreatedDate,b.TypeName as userType,a.userNumber from UserX as a, UserType as b " & _
    "  where a.userType=b.typeid order by UserID"
Set Grid.DataSource = DbCon.Execute(SQL)
Grid.Refresh
End Sub

Private Sub cmdCancel_Click()
Form_Load
End Sub

Private Sub CmdDelete_Click()
If MsgBox("Are You Sure To Delete This?", vbCritical + vbYesNo) = vbYes Then
    SQL = "delete from userX where UserID='" & Trim(Grid.Columns(1).Text) & "'"
    DbCon.Execute SQL
    MsgBox "Data Deleted."
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
If Trim(TxtUserID) = "" Then
    MsgBox "User ID Still Blank"
    Exit Sub
ElseIf Trim(TxtUserName) = "" Then
    MsgBox "User Name Still Blank"
    Exit Sub
ElseIf Trim(TxtUserPassword) = "" Then
    MsgBox "User Password Still Blank"
    Exit Sub
ElseIf Trim(CmbPosition) = "" Or Not CmbPosition.IsItemInList Then
    MsgBox "User Position Still Blank"
    Exit Sub
End If

If Not Edit Then
    SQL = "insert into UserX values('" & Trim(TxtUserID.Text) & "','" & Trim(TxtUserName.Text) & _
        "','" & Trim(Trans.encryp_pass(25, Trim(TxtUserPassword.Text))) & "','" & FormatTgl(Date) & "','" & _
        FormatTgl(Date + 30) & "','" & Trim(CmbPosition.Columns(0).Text) & "')"
    DbCon.Execute SQL
    MsgBox "Data Saved"
    RefreshData
    Form_Load
Else
    SQL = "Update UserX set UserID='" & Trim(TxtUserID.Text) & "',UserName=,'" & Trim(TxtUserName.Text) & _
        "',UserPassword='" & Trim(Trans.encryp_pass(25, Trim(TxtUserPassword.Text))) & _
        "',UserType='" & Trim(CmbPosition.Columns(0).Text) & "'"
    DbCon.Execute SQL
    MsgBox "Data Updated"
    RefreshData
    Form_Load
End If
End Sub

Private Sub Form_Activate()
CekForm Me, TxtID
End Sub

Private Sub Form_Load()
AdoType.ConnectionString = ConDB
TxtID = "A01-03-01"
Me.Height = Me.BasForm1.Height
Me.Width = Me.BasForm1.Width
Bersih
Tombol Me, True
TxtUserID.Locked = 1
TxtUserName.Locked = 1
TxtUserPassword.Locked = 1
CmbPosition.Enabled = 0
RefreshData
End Sub

Sub Bersih()
TxtUserID = ""
TxtUserName = ""
TxtUserPassword = ""
CmbPosition = ""
End Sub

Private Sub Grid_DblClick()
Edit = False
TxtUserNumber = Trim(Grid.Columns(0).Text)
TxtUserID = Trim(Grid.Columns(1).Text)
TxtUserName = Trim(Grid.Columns(2).Text)
CmbPosition = Trim(Grid.Columns(3).Text)
SQL = "select password from userX where userNumber=" & Val(Grid.Columns(0).Text) & ""
Set RSFind = DbCon.Execute(SQL)
TxtUserPassword = Trim(Trans.decryp_pass(25, RSFind!Password))
End Sub

