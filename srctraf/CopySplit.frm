VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CopySplit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6150
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   11610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   11610
   Begin VB.TextBox edcCheckingMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   420
      Left            =   1485
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "Checking that all Multicast Stations selected"
      Top             =   3780
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "&Import"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7290
      TabIndex        =   47
      Top             =   5670
      Width           =   1050
   End
   Begin VB.TextBox edcMetroMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   420
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "Call Counterpoint as this is a paid feature"
      Top             =   2940
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.ListBox lbcFrom 
      Height          =   3375
      Index           =   4
      ItemData        =   "CopySplit.frx":0000
      Left            =   960
      List            =   "CopySplit.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   45
      Top             =   1710
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.CommandButton cmcClear 
      Appearance      =   0  'Flat
      Caption         =   "C&lear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8880
      TabIndex        =   43
      Top             =   5445
      Width           =   1050
   End
   Begin VB.PictureBox pbcExcludeArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7275
      Picture         =   "CopySplit.frx":0004
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2925
      Width           =   180
   End
   Begin VB.PictureBox pbcStationSelection 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   375
      ScaleHeight     =   255
      ScaleWidth      =   3975
      TabIndex        =   36
      Top             =   5130
      Visible         =   0   'False
      Width           =   3975
      Begin VB.OptionButton rbcStationSelection 
         Caption         =   "Single"
         Height          =   195
         Index           =   0
         Left            =   1785
         TabIndex        =   38
         Top             =   0
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton rbcStationSelection 
         Caption         =   "Range"
         Height          =   195
         Index           =   1
         Left            =   2835
         TabIndex        =   37
         Top             =   0
         Width           =   885
      End
      Begin VB.Label lacStationSelection 
         Caption         =   "Station selection"
         Height          =   210
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.ListBox lbcFrom 
      Height          =   3375
      Index           =   3
      ItemData        =   "CopySplit.frx":00DE
      Left            =   855
      List            =   "CopySplit.frx":00E0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   35
      Top             =   1605
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.ListBox lbcFrom 
      Height          =   3375
      Index           =   2
      ItemData        =   "CopySplit.frx":00E2
      Left            =   675
      List            =   "CopySplit.frx":00E4
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.ListBox lbcFrom 
      Height          =   3375
      Index           =   1
      ItemData        =   "CopySplit.frx":00E6
      Left            =   540
      List            =   "CopySplit.frx":00E8
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   1635
      Width           =   3450
   End
   Begin VB.ListBox lbcFrom 
      Height          =   3570
      Index           =   0
      ItemData        =   "CopySplit.frx":00EA
      Left            =   420
      List            =   "CopySplit.frx":00EC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   1470
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.PictureBox pbcIncludeArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7275
      Picture         =   "CopySplit.frx":00EE
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2355
      Width           =   180
   End
   Begin VB.PictureBox pbcUpMove 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6540
      Picture         =   "CopySplit.frx":01C8
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3750
      Width           =   180
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   10515
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "CopySplit.frx":02A2
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   19
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   16
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcPrinting 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   9600
      ScaleHeight     =   1200
      ScaleWidth      =   3825
      TabIndex        =   25
      Top             =   3585
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox pbcNewTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   11580
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   45
      Width           =   60
   End
   Begin VB.CommandButton cmcSpec 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1170
      Picture         =   "CopySplit.frx":30BC
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ckcIncludeDormant 
      Caption         =   "Include Dormant Names"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Width           =   2070
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8895
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5370
      Width           =   45
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   10665
      Top             =   5535
   End
   Begin VB.PictureBox pbcStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1815
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox edcSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   180
      MaxLength       =   10
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   375
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   45
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   10
      Top             =   825
      Width           =   45
   End
   Begin VB.PictureBox pbcSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   375
      Width           =   60
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   7245
      TabIndex        =   2
      Top             =   75
      Width           =   4260
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   22
      Top             =   5670
      Width           =   1050
   End
   Begin VB.CommandButton cmcDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1170
      Picture         =   "CopySplit.frx":31B6
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   180
      MaxLength       =   10
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   8820
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "CopySplit.frx":32B0
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "CopySplit.frx":3F6E
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   90
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5640
      Width           =   75
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4545
      TabIndex        =   21
      Top             =   5670
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3180
      TabIndex        =   20
      Top             =   5670
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpec 
      Height          =   450
      Left            =   2655
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   794
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   2
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTo 
      Height          =   4110
      Left            =   7575
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1245
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   7250
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   12632256
      ForeColorSel    =   -2147483635
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmcMoveFrom 
      Appearance      =   0  'Flat
      Caption         =   "  &Move"
      Height          =   300
      Left            =   6495
      TabIndex        =   30
      Top             =   3690
      Width           =   810
   End
   Begin VB.CommandButton cmcMoveInclude 
      Appearance      =   0  'Flat
      Caption         =   "&Include"
      Height          =   300
      Left            =   6360
      TabIndex        =   31
      Top             =   2295
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStation 
      Height          =   3435
      Left            =   300
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1485
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin ComctlLib.TabStrip tbcCategory 
      Height          =   4500
      Left            =   120
      TabIndex        =   26
      Top             =   945
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   7938
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&DMA Market"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Format"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&MSA Market"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&State"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S&tation"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Zone"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmcMoveExclude 
      Appearance      =   0  'Flat
      Caption         =   "&Exclude"
      Height          =   300
      Left            =   6360
      TabIndex        =   42
      Top             =   2865
      Width           =   1125
   End
   Begin VB.Label lacResult 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   195
      Left            =   7800
      TabIndex        =   44
      Top             =   1020
      Width           =   3450
   End
   Begin VB.Label plcScreen 
      Caption         =   "Region Definition"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   1425
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   165
      Top             =   5610
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "CopySplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CopySplit.frm on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmRafSrchKey2                 tmAttSrchKey2                 tmClfSrchKey4             *
'*  tmSdfSrchKey5                 tmSmf                         tmAvail                   *
'*  tmSpot                        imVefCode                     imVpfIndex                *
'*  lmLLD                         lmFirstAllowedChgDate         smInclExcl                *
'*  imLastColSorted               imLastSort                    smShowOnProp              *
'*  smShowOnOrder                 smShowOnInv                   lmTopRow                  *
'*  imInitNoRows                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CopySplit.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

'Region Area
Dim tmRaf As RAF            'RAF record image
Dim tmRafSrchKey As LONGKEY0  'RAF key record image
Dim hmRaf As Integer        'RAF Handle
Dim imRafRecLen As Integer      'RAF record length

'Split Entity
Dim tmSef() As SEF            'SEF record image
Dim tmSefSrchKey As LONGKEY0  'SEF key record image
Dim tmSefSrchKey1 As SEFKEY1  'SEF key record image
Dim hmSef As Integer        'SEF Handle
Dim imSefRecLen As Integer      'SEF record length
'5882 no longer needed
'IDC Enforced
'Dim tmIef As IEF            'IEF record image
'Dim tmIefSrchKey As LONGKEY0  'IEF key record image
'Dim tmIefSrchKey1 As LONGKEY0  'IEF key record image (GenericCifCode)
'Dim tmIefSrchKey2 As LONGKEY0  'IEF key record image (SplitRafCode)
'Dim hmIef As Integer        'IEF Handle
'Dim imIefRecLen As Integer      'IEF record length

'Agreement
Dim tmAtt As ATT                'ATT record image
Dim imAttRecLen As Integer      'ATT record length
Dim hmAtt As Integer            'Agreement file handle

'ARTT- Get Owner names
Dim tmArtt As ARTT            'ARTT record image

'Market Names
Dim tmMkt As MKT            'MKT record image
Dim tmMet As MET            'MET record image

'Stations
Dim tmSHTT As SHTT
Dim imStationPop As Integer

'Format Names
Dim tmFmt As FMT            'MKT record image

'State
Dim tmSnt As SNT

'Zone
Dim tmTzt As TZT

Dim tmRegionCode() As SORTCODE
Dim smRegionCodeTag As String

'Contract line
Dim hmClf As Integer        'Contract line file handle
Dim tmClf As CLF            'CLF record image
Dim imClfRecLen As Integer

Dim hmSdf As Integer        'Spot detail file handle
Dim tmSdf As SDF            'SDF record image
Dim imSdfRecLen As Integer  'SDF record length

Dim hmSmf As Integer

Dim hmSsf As Integer

Dim hmStf As Integer


'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imInNewTab As Integer
Dim imSettingValue As Integer
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Dim smNowDate As String
Dim lmNowDate As Long
Dim smStatus As String
Dim imSpecChg As Integer
Dim imCountChg As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmSpecEnableRow As Long
Dim lmSpecEnableCol As Long
Dim imSpecCtrlVisible As Integer
Dim lmToRowSelected As Long

Dim imLastStationColSorted As Integer
Dim imLastStationSort As Integer
Dim lmStationRangeRow As Long
Dim imStationMoved As Integer

Dim bmLoadingInfo As Boolean

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

Dim hmUnMatch As Integer
Dim bmMulticastMissing As Boolean
Dim shtt_rst As ADODB.Recordset
Dim att_rst As ADODB.Recordset

Dim imMissingShttCode() As Integer

'Mouse down
Const SPECROW3INDEX = 3
'Const CATEGORYINDEX = 2
Const NAMEINDEX = 2 '4
'Const INCLEXCLINDEX = 6
'Const AUDPCTINDEX = 8
Const STATUSINDEX = 4   '10
'Const SHOWONINDEX = 12
'Const SHOWONPROPINDEX = 14
'Const SHOWONORDERINDEX = 16
'Const SHOWONINVINDEX = 18

Const TOINCLEXCLINDEX = 0
Const TONAMEINDEX = 1
Const TOCATEGORYINDEX = 2
Const TOCODEINDEX = 3

Const STATIONINDEX = 0
Const MARKETINDEX = 1
Const STATEINDEX = 2
Const ZONEINDEX = 3
Const FORMATINDEX = 4
Const SHTTCODEINDEX = 5
Const SORTINDEX = 6
Const SELECTEDINDEX = 7

Const LBCFORMATINDEX = 0
Const LBCDMAMARKETINDEX = 1
Const LBCSTATEINDEX = 2
Const LBCZONEINDEX = 3
Const LBCMSAMARKETINDEX = 4







Private Sub cbcSelect_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box

    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    mSetMousePointer vbHourglass
    bmLoadingInfo = True
    cmcDone.Enabled = False
    cmcCancel.Enabled = False
    cmcSave.Enabled = False
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    mClearCtrlFields
    If (ilRet = 0) And (cbcSelect.ListIndex > 1) Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY, 0) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            If cbcSelect.ListCount > 0 Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.ListIndex = -1
            End If
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        grdSpec.Redraw = False
        mMoveRecToCtrl
        grdSpec.Redraw = True
        mMoveSEFRecToCtrl
        '12/24/08:  Allow changes to region definition- this might be temporary
        'If tmRaf.sAssigned = "N" Then
            cmcMoveInclude.Enabled = True
            cmcMoveExclude.Enabled = True
        'Else
        '    cmcMoveInclude.Enabled = False
        '    cmcMoveExclude.Enabled = False
        'End If
    Else
        imSelectedIndex = cbcSelect.ListIndex
        If (slStr <> "[New]") And (slStr <> "[Model]") Then
            grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = slStr
        End If
        cmcMoveInclude.Enabled = True
        cmcMoveExclude.Enabled = True
    End If
    cmcDone.Enabled = True
    cmcCancel.Enabled = True
    mSetMousePointer vbDefault
    imChgMode = False
    bmLoadingInfo = False
    mSetCommands
    Exit Sub
cbcSelectErr:
    cmcDone.Enabled = True
    cmcCancel.Enabled = True
    On Error GoTo 0
    mSetMousePointer vbDefault
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcSelect_GotFocus()
    mSetMousePointer vbHourglass
    lmSpecEnableRow = -1
    mSpecSetShow
    If cbcSelect.Text = "" Then
        gFindMatch sgUserDefVehicleName, 0, cbcSelect
        If gLastFound(cbcSelect) >= 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        Else
            If cbcSelect.ListCount >= 1 Then
                cbcSelect.ListIndex = 0
            End If
        End If
        imComboBoxIndex = cbcSelect.ListIndex
        imSelectedIndex = imComboBoxIndex
    End If
    imComboBoxIndex = imSelectedIndex
    If imSelectedIndex <= 1 Then
        mClearCtrlFields
    End If
    gCtrlGotFocus cbcSelect
    mSetMousePointer vbDefault
End Sub

Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub ckcIncludeDormant_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                         ilPos                         slStr                     *
'*                                                                                        *
'******************************************************************************************


    smRegionCodeTag = ""
    mPopulate igAdfCode
End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    sgDoneMsg = ""
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSpecSetShow
    gCtrlGotFocus cmcCancel
End Sub


Private Sub cmcClear_Click()
    Dim llFromRow As Long

    For llFromRow = grdTo.Rows - 1 To grdTo.FixedRows Step -1
        grdTo.Row = llFromRow
        grdTo.Col = TOINCLEXCLINDEX
        mXFerToFrom
    Next llFromRow
    If imStationMoved Then
        mResortStations
    End If
    cmcClear.Enabled = False
End Sub

Private Sub cmcDone_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                                                                                 *
'******************************************************************************************

    Dim slMess As String
    Dim ilRet As Integer
    Dim slStr As String

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If imSpecChg Or imCountChg Then
        slStr = grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)
        If imSelectedIndex > 1 Then
            slMess = "Save Changes to " & slStr
        Else
            slMess = "Add " & slStr
        End If
        ilRet = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If ilRet = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
            igSplitChgd = True
        End If
    End If
    sgDoneMsg = grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSpecSetShow
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
            End Select
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub cmcImport_Click()
    Dim ilRet As Integer
    
    igBrowserType = 2   'txt
    Browser.Show vbModal
    If igBrowserReturn = 1 Then
        ilRet = mReadStationFile(sgBrowserFile)
    End If

End Sub

Private Sub cmcMoveExclude_Click()
    grdTo.Redraw = False
    If (Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) = REMOTEEXPORT Then
        mXFerFromTo 5, "Excl", -1
    Else
        mXFerFromTo tbcCategory.SelectedItem.Index, "Excl", -1
    End If
    grdTo.Redraw = True
End Sub

Private Sub cmcMoveExclude_GotFocus()
    mSpecSetShow
End Sub

Private Sub cmcMoveFrom_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRow                         llToRow                       ilShtt                    *
'*  ilMkt                         ilFormat                      slStr                     *
'*                                                                                        *
'******************************************************************************************

    Dim llFromRow As Long

    grdTo.Redraw = False
    For llFromRow = grdTo.Rows - 1 To grdTo.FixedRows Step -1
        grdTo.Row = llFromRow
        grdTo.Col = TOINCLEXCLINDEX
        If grdTo.CellBackColor = GRAY Then
            mXFerToFrom
        End If
    Next llFromRow
    grdTo.Redraw = True
    If imStationMoved Then
        mResortStations
    End If
    imCountChg = True
    mSetCommands
End Sub

Private Sub cmcMoveInclude_Click()
    grdTo.Redraw = False
    If (Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) = REMOTEEXPORT Then
        mXFerFromTo 5, "Incl", -1
    Else
        mXFerFromTo tbcCategory.SelectedItem.Index, "Incl", -1
    End If
    grdTo.Redraw = True
End Sub

Private Sub cmcMoveInclude_GotFocus()
    mSpecSetShow
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ilRet = mSaveRec()
    If Not ilRet Then
        Exit Sub
    End If
    smRegionCodeTag = ""
    mPopulate igAdfCode
    mSetCommands
    igSplitChgd = True
End Sub

Private Sub cmcSave_GotFocus()
    mSpecSetShow
End Sub

Private Sub edcDropDown_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                                                   *
'******************************************************************************************


    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
            End Select
    End Select
    imLbcArrowSetting = False
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub edcDropDown_DblClick()
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
            End Select
    End Select
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
            End Select
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFound                       ilLoop                                                  *
'******************************************************************************************

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
            End Select
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                                                                                *
'******************************************************************************************

    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                End Select
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                End Select
        End Select
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                End Select
        End Select
        imDoubleClickName = False
    End If
End Sub

Private Sub edcSpec_Change()
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case NAMEINDEX
            '    Case AUDPCTINDEX
            End Select
    End Select
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub edcSpec_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSpec_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False

End Sub

Private Sub edcSpec_KeyPress(KeyAscii As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                         slStr                                                   *
'******************************************************************************************


    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case NAMEINDEX
            '    Case AUDPCTINDEX
            '        ilPos = InStr(edcSpec.SelText, ".")
            '        If ilPos = 0 Then
            '            ilPos = InStr(edcSpec.Text, ".")    'Disallow multi-decimal points
            '            If ilPos > 0 Then
            '                If KeyAscii = KEYDECPOINT Then
            '                    Beep
            '                    KeyAscii = 0
            '                    Exit Sub
            '                End If
            '            End If
            '        End If
            '        'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            '        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
            '            Beep
            '            KeyAscii = 0
            '            Exit Sub
            '        End If
            '        slStr = edcSpec.Text
            '        slStr = Left$(slStr, edcSpec.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSpec.SelStart - edcSpec.SelLength)
            '        If gCompNumberStr(slStr, "100.00") > 0 Then
            '            Beep
            '            KeyAscii = 0
            '            Exit Sub
            '        End If

            End Select
    End Select

End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(COPYJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        grdSpec.Enabled = False
        grdStation.Enabled = False
        imUpdateAllowed = False
    Else
        grdSpec.Enabled = True
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        grdStation.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        If lmEnableCol > 0 Then
        '    mEnableBox
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        tmcTerminate.Enabled = True
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmRegionCode
    Erase imMissingShttCode
    
    btrDestroy hmRaf
    btrDestroy hmSef
    '5882 no longer needed
   ' btrDestroy hmIef
    btrDestroy hmSdf
    btrDestroy hmSsf
    btrDestroy hmSmf
    btrDestroy hmStf
    btrDestroy hmClf
    btrDestroy hmAtt
    
    shtt_rst.Close
    att_rst.Close
    
    Set CopySplit = Nothing   'Remove data segment

End Sub

Private Sub grdSpec_EnterCell()
    mSpecSetShow
End Sub

Private Sub grdSpec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lmTopRow = grdSpec.TopRow
    grdSpec.Redraw = False
End Sub

Private Sub grdSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine if in header
'    If y < grdSpec.RowHeight(0) Then
'        mSortCol grdSpec.Col
'        Exit Sub
'    End If
    'Determine row and col mouse up onto
    On Error GoTo grdSpecErr:
    ilCol = grdSpec.MouseCol
    ilRow = grdSpec.MouseRow
    If ilCol < grdSpec.FixedCols Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    If ilRow < grdSpec.FixedRows Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    If ilRow Mod 2 = 0 Then
        ilRow = ilRow + 1
    End If
    If grdSpec.ColWidth(ilCol) <= 15 Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    If grdSpec.RowHeight(ilRow) <= 15 Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    'lmTopRow = grdSpec.TopRow
    DoEvents
    grdSpec.Col = ilCol
    grdSpec.Row = ilRow
    grdSpec.Redraw = True
    mSpecEnableBox
    On Error GoTo 0
    Exit Sub
grdSpecErr:
    On Error GoTo 0
    If (lmSpecEnableRow >= grdSpec.FixedRows) And (lmSpecEnableRow < grdSpec.Rows) Then
        grdSpec.Row = lmSpecEnableRow
        grdSpec.Col = lmSpecEnableCol
        mSpecSetFocus
    End If
    grdSpec.Redraw = False
    grdSpec.Redraw = True
    Exit Sub
End Sub

Private Sub grdStation_Click()
    Dim llRow As Long
    Dim llCol As Long
    Dim llLoop As Long

    llRow = grdStation.Row
    llCol = grdStation.Col
    If rbcStationSelection(1).Value Then
        If lmStationRangeRow >= grdStation.FixedRows Then
            If lmStationRangeRow < llRow Then
                For llLoop = lmStationRangeRow To llRow Step 1
                    If grdStation.TextMatrix(llLoop, STATIONINDEX) <> "" Then
                        grdStation.Row = llLoop
                        grdStation.TextMatrix(llLoop, SELECTEDINDEX) = "T"
                        For llCol = STATIONINDEX To FORMATINDEX Step 1
                            grdStation.Col = llCol
                            grdStation.CellBackColor = GRAY
                            'grdStation.CellForeColor = vbWhite
                        Next llCol
                    End If
                Next llLoop
            Else
                For llLoop = llRow To lmStationRangeRow Step 1
                    If grdStation.TextMatrix(llLoop, STATIONINDEX) <> "" Then
                        grdStation.Row = llLoop
                        grdStation.TextMatrix(llLoop, SELECTEDINDEX) = "T"
                        For llCol = STATIONINDEX To FORMATINDEX Step 1
                            grdStation.Col = llCol
                            grdStation.CellBackColor = GRAY
                            'grdStation.CellForeColor = vbWhite
                        Next llCol
                    End If
                Next llLoop
            End If
            lmStationRangeRow = -1
        Else
            lmStationRangeRow = llRow
            grdStation.Row = llRow
            grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T"
            For llCol = STATIONINDEX To FORMATINDEX Step 1
                grdStation.Col = llCol
                grdStation.CellBackColor = GRAY
                'grdStation.CellForeColor = vbWhite
            Next llCol
        End If
    Else
        lmStationRangeRow = -1
        If llRow >= grdStation.FixedRows Then
            If grdStation.TextMatrix(llRow, STATIONINDEX) <> "" Then
                If grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T" Then
                    grdStation.Row = llRow
                    grdStation.TextMatrix(llRow, SELECTEDINDEX) = "F"
                    For llCol = STATIONINDEX To FORMATINDEX Step 1
                        grdStation.Col = llCol
                        grdStation.CellBackColor = vbWhite
                        'grdStation.CellForeColor = vbBlack
                    Next llCol
                Else
                    grdStation.Row = llRow
                    grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T"
                    For llCol = STATIONINDEX To FORMATINDEX Step 1
                        grdStation.Col = llCol
                        grdStation.CellBackColor = GRAY
                        'grdStation.CellForeColor = vbWhite
                    Next llCol
                End If
            End If
        End If
    End If
    imCountChg = True

    mSetCommands
End Sub

Private Sub grdStation_EnterCell()
    mSpecSetShow
End Sub

Private Sub grdStation_GotFocus()
    mSpecSetShow
End Sub

Private Sub grdStation_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Y < grdStation.RowHeight(0) Then
        mSetMousePointer vbHourglass
        grdStation.Col = grdStation.MouseCol
        mStationSortCol grdStation.Col
        mSetCommands
        mSetMousePointer vbDefault
        Exit Sub
    End If
End Sub

Private Sub grdTo_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llLoop                                                                                *
'******************************************************************************************

    Dim llRow As Long
    Dim llCol As Long

    llRow = grdTo.Row
    llCol = grdTo.Col
    If llRow >= grdTo.FixedRows Then
        If grdTo.TextMatrix(llRow, TOINCLEXCLINDEX) <> "" Then
            If grdTo.CellBackColor = GRAY Then
                grdTo.Row = llRow
                For llCol = TOINCLEXCLINDEX To TOCATEGORYINDEX Step 1
                    grdTo.Col = llCol
                    grdTo.CellBackColor = vbWhite
                Next llCol
            Else
                grdTo.Row = llRow
                For llCol = TOINCLEXCLINDEX To TOCATEGORYINDEX Step 1
                    grdTo.Col = llCol
                    grdTo.CellBackColor = GRAY
                Next llCol
            End If
        End If
    End If
End Sub

Private Sub grdTo_GotFocus()
    mSpecSetShow
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcDropDown.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slName                                                                                *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer

    mSetMousePointer vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    imTerminate = False
    imBypassFocus = False
    imSettingValue = False
    imStartMode = True
    imChgMode = False
    imBSMode = False
    bmLoadingInfo = False
    imLbcArrowSetting = False
    imLbcMouseDown = False
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imCalType = 0   'Standard
    imCtrlVisible = False
    imSpecCtrlVisible = False
    imSpecChg = False
    imCountChg = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmStationRangeRow = -1
    imInNewTab = False
    imStationPop = False
    lmToRowSelected = -1
    imStationMoved = False
    mInitBox

    If Not gRecLengthOk("Raf.btr", Len(tmRaf)) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If
    hmRaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", CopyRegn
    On Error GoTo 0
    imRafRecLen = Len(tmRaf)

    ReDim tmSef(0 To 0) As SEF
    If Not gRecLengthOk("Sef.btr", Len(tmSef(0))) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If
    hmSef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSef, "", sgDBPath & "Sef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sef.Btr)", CopyRegn
    On Error GoTo 0
    imSefRecLen = Len(tmSef(0))
''5882 no longer needed
'    hmIef = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmIef, "", sgDBPath & "Ief.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ief.Btr)", CopyRegn
'    On Error GoTo 0
'    imIefRecLen = Len(tmIef)

    hmAtt = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAtt, "", sgDBPath & "Att.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Att.Mkd)", CopyRegn
    On Error GoTo 0
    imAttRecLen = Len(tmAtt)

    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", CopyRegn
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", CopyRegn
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)

    hmSmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", CopyRegn
    On Error GoTo 0

    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", CopyRegn
    On Error GoTo 0

    hmStf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Stf.Btr)", CopyRegn
    On Error GoTo 0

    If Not gRecLengthOk("Artt.mkd", Len(tmArtt)) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If
    mPopOwners


    If Not gRecLengthOk("Mkt.mkd", Len(tmMkt)) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If
    mPopDMAMarkets
    If Not gRecLengthOk("Met.mkd", Len(tmMet)) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If
    mPopMSAMarkets

    If Not gRecLengthOk("Snt.mkd", Len(tmSnt)) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If
    mPopStates

    If Not gRecLengthOk("Tzt.mkd", Len(tmTzt)) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If
    mPopTimeZones

    If (UBound(tgStates) <= LBound(tgStates)) Or (UBound(tgTimeZones) <= LBound(tgTimeZones)) Then
        MsgBox "Exit Traffic System and Sign-On to the Affiliate System to initialize tables required by Copy Splits", vbCritical + vbOKOnly, "Split Copy"
    End If

    If Not gRecLengthOk("Shtt.mkd", Len(tmSHTT)) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If

    If Not gRecLengthOk("Fmt.mkd", Len(tmFmt)) Then
        imTerminate = True
        mSetMousePointer vbDefault
        Exit Sub
    End If
    mPopFormats

    'Populate only if Station category selected
    'mPopStations

    
    mPopCategory

    mPopulate igAdfCode

    '11/6/14: Hide tab control and show station Grid instead
    If (Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) = REMOTEEXPORT Then
        lbcFrom(LBCDMAMARKETINDEX).Visible = False
        tbcCategory.Visible = False
        mPopStations True
        grdStation.Visible = True
        pbcStationSelection.Visible = True
        cmcImport.Enabled = True
    End If

    'mXFerRecToCtrl
    CopySplit.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone CopySplit
    ' Dan M 9-25-09 adjust look of 'wait' message
    gAdjustScreenMessage Me, pbcPrinting
    mSetMousePointer vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    mSetMousePointer vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  flTextHeight                  ilCol                         llRet                     *
'*                                                                                        *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRow As Integer
    'flTextHeight = pbcDates.TextHeight("1") - 35


    grdSpec.Move 2655, 435
    tbcCategory.Move 120, 945
    lbcFrom(LBCFORMATINDEX).Move tbcCategory.Left + 180, tbcCategory.Top + 525, lbcFrom(LBCFORMATINDEX).Width, lbcFrom(LBCFORMATINDEX).Height
    lbcFrom(LBCDMAMARKETINDEX).Move lbcFrom(LBCFORMATINDEX).Left, lbcFrom(LBCFORMATINDEX).Top, lbcFrom(LBCFORMATINDEX).Width, lbcFrom(LBCFORMATINDEX).Height
    lbcFrom(LBCSTATEINDEX).Move lbcFrom(LBCFORMATINDEX).Left, lbcFrom(LBCFORMATINDEX).Top, lbcFrom(LBCFORMATINDEX).Width, lbcFrom(LBCFORMATINDEX).Height
    lbcFrom(LBCZONEINDEX).Move lbcFrom(LBCFORMATINDEX).Left, lbcFrom(LBCFORMATINDEX).Top, lbcFrom(LBCFORMATINDEX).Width, lbcFrom(LBCFORMATINDEX).Height
    lbcFrom(LBCMSAMARKETINDEX).Move lbcFrom(LBCFORMATINDEX).Left, lbcFrom(LBCFORMATINDEX).Top, lbcFrom(LBCFORMATINDEX).Width, lbcFrom(LBCFORMATINDEX).Height
    'grdTo.Move grdTo.Left, tbcCategory.Top + 60, grdTo.Width, tbcCategory.Height - 120
    pbcStationSelection.Move lbcFrom(LBCFORMATINDEX).Left, lbcFrom(LBCFORMATINDEX).Top + lbcFrom(LBCFORMATINDEX).Height + 60

    mGridSpecLayout
    mGridSpecColumnWidths
    mGridSpecColumns

    mGridStationLayout
    mGridStationColumnWidths
    mGridStationColumns
    ilRow = grdStation.FixedRows
    Do
        If ilRow + 1 > grdStation.Rows Then
            grdStation.AddItem ""
        End If
        grdStation.RowHeight(ilRow) = fgBoxGridH + 15
        ilRow = ilRow + 1
    Loop While grdStation.RowIsVisible(ilRow - 1)
    gGrid_IntegralHeight grdStation, CInt(fgBoxGridH + 30) ' + 15
    grdStation.Height = grdStation.Height + 15 '- 30
    grdStation.Top = lbcFrom(LBCFORMATINDEX).Top
    grdStation.Left = lbcFrom(LBCFORMATINDEX).Left
    'cmcClearStationSelection.Top = pbcStationSelection.Top

    mGridToLayout
    mGridToColumnWidths
    mGridToColumns
    ilRow = grdTo.FixedRows
    Do
        If ilRow + 1 > grdTo.Rows Then
            grdTo.AddItem ""
        End If
        grdTo.RowHeight(ilRow) = fgBoxGridH + 15
        ilRow = ilRow + 1
    Loop While grdTo.RowIsVisible(ilRow - 1)
    gGrid_IntegralHeight grdTo, CInt(fgBoxGridH + 30) ' + 15
    grdTo.Height = grdTo.Height - 30
    grdTo.Left = grdSpec.Left + grdSpec.Width - grdTo.Width
    grdTo.Top = lbcFrom(LBCFORMATINDEX).Top + (lbcFrom(LBCFORMATINDEX).Height) / 2 - (grdTo.Height) / 2 + 60
    lacResult.Top = grdTo.Top - lacResult.Height - 60
    cmcClear.Top = grdTo.Top + grdTo.Height + 60
    cmcClear.Left = grdTo.Left + (grdTo.Width) / 2 - (cmcClear.Width) / 2

    cbcSelect.Left = grdSpec.Left + grdSpec.Width - cbcSelect.Width
    cbcSelect.Top = 60
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop

    pbcPrinting.Left = (CopySplit.Width - pbcPrinting.Width) \ 2
    pbcPrinting.Top = (CopySplit.Height - pbcPrinting.Height) \ 2
    edcMetroMsg.Left = tbcCategory.Left + tbcCategory.Width / 2 - edcMetroMsg.Width \ 2
    edcMetroMsg.Top = tbcCategory.Top + tbcCategory.Height / 2 - edcMetroMsg.Height / 2
    edcCheckingMsg.Left = tbcCategory.Left + tbcCategory.Width / 2 - edcCheckingMsg.Width \ 2
    edcCheckingMsg.Top = tbcCategory.Top + tbcCategory.Height / 2 - edcCheckingMsg.Height / 2
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()


'
'   mTerminate
'   Where:
'

    mSetMousePointer vbDefault
    igManUnload = YES
    Unload CopySplit
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mSpecFieldsOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilValue                                                 *
'******************************************************************************************

'
'   iRet = mSpecFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilError As Integer

    ilError = False
    grdSpec.Row = SPECROW3INDEX
    grdSpec.Col = NAMEINDEX
    grdSpec.CellForeColor = vbBlack
    grdSpec.Col = STATUSINDEX
    grdSpec.CellForeColor = vbBlack
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = NAMEINDEX
        grdSpec.CellForeColor = vbMagenta
    End If
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = STATUSINDEX
        grdSpec.CellForeColor = vbMagenta
    End If
    If ilError Then
        mSpecFieldsOk = False
    Else
        mSpecFieldsOk = True
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

'
'   mXFerRecToCtrl
'   Where:
'

    grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = Trim$(tmRaf.sName)
    If tmRaf.sState = "A" Then
        grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
    ElseIf tmRaf.sState = "D" Then
        grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Dormant"
    End If
End Sub


Private Sub lbcFrom_GotFocus(Index As Integer)
    mSpecSetShow
End Sub



Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDropDown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSpecSetShow
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub pbcNewTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilLoop                        slNameCode                *
'*  slCode                                                                                *
'******************************************************************************************


    If imInNewTab Then
        Exit Sub
    End If
    If imUpdateAllowed = False Then
        cmcCancel.SetFocus
        Exit Sub
    End If

    If imSelectedIndex > 1 Then
        pbcSpecSTab.SetFocus
        Exit Sub
    End If
    If imSelectedIndex = 0 Then
        grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
        pbcSpecSTab.SetFocus
        Exit Sub
    End If
    If UBound(tmSef) = LBound(tmSef) Then
        imInNewTab = True
        igSplitType = 0         'Split copy flag
        igIncludeDormantSplits = False
        If ckcIncludeDormant.Value = vbChecked Then
            igIncludeDormantSplits = True
        End If
        'SplitModel.Show vbModal
        lgSplitModelCodeRaf = 0
        RegionModel.Show vbModal
        DoEvents
        If (igSplitModelReturn = 1) And (lgSplitModelCodeRaf) > 0 Then
            mSetMousePointer vbHourglass
            bmLoadingInfo = True
            grdSpec.Redraw = False
            cmcDone.Enabled = False
            cmcCancel.Enabled = False
            cmcSave.Enabled = False
            If mReadRec(imSelectedIndex, SETFORREADONLY, lgSplitModelCodeRaf) Then
                mMoveRecToCtrl
                mMoveSEFRecToCtrl
                grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = ""
            End If
            grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
            grdSpec.Redraw = True
            cmcDone.Enabled = True
            cmcCancel.Enabled = True
            mSetMousePointer vbDefault
            imCountChg = True
            bmLoadingInfo = False
            mSetCommands
        Else
            grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
        End If
    End If
    
    imInNewTab = False
    pbcSpecSTab.SetFocus

End Sub

Private Sub pbcPrinting_Paint()
    pbcPrinting.CurrentX = (pbcPrinting.Width - pbcPrinting.TextWidth("Resolving Station Conflicts....")) / 2
    pbcPrinting.CurrentY = (pbcPrinting.Height - pbcPrinting.TextHeight("Resolving Station Conflicts....")) / 2 - 30
    pbcPrinting.Print "Resolving Station Conflicts...."
End Sub

Private Sub pbcSpecSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSpecSTab.hWnd Then
        Exit Sub
    End If
    If imSpecCtrlVisible Then
        Do
            ilNext = False
            Select Case grdSpec.Row
                Case SPECROW3INDEX
                    Select Case grdSpec.Col
                        Case NAMEINDEX
                            mSpecSetShow
                            cmcDone.SetFocus
                            Exit Sub
                        Case Else
                            If grdSpec.Col >= NAMEINDEX + 2 Then
                                grdSpec.Col = grdSpec.Col - 2
                            Else
                                mSpecSetShow
                                cmcDone.SetFocus
                                Exit Sub
                            End If
                    End Select
            End Select
            If mSpecColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSpecSetShow
    Else
        grdSpec.Row = SPECROW3INDEX '+1 to bypass title
        grdSpec.Col = grdSpec.FixedCols
        Do
            If mSpecColOk() Then
                Exit Do
            Else
                grdSpec.Col = grdSpec.Col + 2
            End If
        Loop
    End If
    mSpecEnableBox
End Sub

Private Sub pbcSpecTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSpecTab.hWnd Then
        Exit Sub
    End If
    If imSpecCtrlVisible Then
        Do
            ilNext = False
            Select Case grdSpec.Row
                Case SPECROW3INDEX
                    Select Case grdSpec.Col
                        Case STATUSINDEX
                            mSpecSetShow
                            cmcDone.SetFocus
                            Exit Sub
                        Case Else
                            grdSpec.Col = grdSpec.Col + 2
                    End Select
            End Select
            If mSpecColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSpecSetShow
    Else
        grdSpec.Row = grdSpec.Rows - 2
        grdSpec.Col = STATUSINDEX
        Do
            If mSpecColOk() Then
                Exit Do
            Else
                grdSpec.Col = grdSpec.Col - 2
            End If
        Loop
    End If
    mSpecEnableBox
End Sub





Private Sub pbcStatus_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
        smStatus = "Active"
        pbcStatus_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        smStatus = "Dormant"
        pbcStatus_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smStatus = "Active" Then
            smStatus = "Dormant"
            pbcStatus_Paint
        ElseIf smStatus = "Dormant" Then
            smStatus = "Active"
            pbcStatus_Paint
        End If
    End If
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub pbcStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smStatus = "Active" Then
        smStatus = "Dormant"
        pbcStatus_Paint
    Else
        smStatus = "Active"
        pbcStatus_Paint
    End If
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub pbcStatus_Paint()
    pbcStatus.Cls
    pbcStatus.CurrentX = fgBoxInsetX
    pbcStatus.CurrentY = 0 'fgBoxInsetY
    pbcStatus.Print smStatus
End Sub


Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    'Select Case lmEnableCol
                    '    Case AIRTIMEINDEX
                    '        imBypassFocus = True    'Don't change select text
                    '        edcDropDown.SetFocus
                    '        SendKeys slKey
                    'End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub



Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer, llModelRafCode As Long) As Integer
'
'   iRet = mReadRec(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    Dim tlSef As SEF

    If ilSelectIndex > 1 Then
        slNameCode = tmRegionCode(ilSelectIndex - 2).sKey    'lbcCopyRegnCode.List(ilSelectIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mReadRecErr
        gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", CopySplit
        On Error GoTo 0
        tmRafSrchKey.lCode = CLng(slCode)
    Else
        tmRafSrchKey.lCode = llModelRafCode
    End If
    ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Region)", CopySplit
    On Error GoTo 0
    ReDim tmSef(0 To 0) As SEF
    tmSefSrchKey1.lRafCode = tmRaf.lCode
    tmSefSrchKey1.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hmSef, tlSef, imSefRecLen, tmSefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlSef.lRafCode = tmRaf.lCode)
        tmSef(UBound(tmSef)) = tlSef
        ReDim Preserve tmSef(0 To UBound(tmSef) + 1) As SEF
        ilRet = btrGetNext(hmSef, tlSef, imSefRecLen, BTRV_LOCK_NONE, ilForUpdate)
    Loop
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function


Private Sub mClearCtrlFields()
    Dim llRow As Long
    Dim ilCol As Integer


    ReDim tmSef(0 To 0) As SEF
    grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = ""
    mPopFrom "Format"
    mPopFrom "DMA Market"
    mPopFrom "MSA Market"
    mPopFrom "State Name"
    mPopFrom "Time Zone"

    If imStationPop Then
        imStationPop = False
        mPopStations False
    End If
    For llRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
        grdTo.Row = llRow
        For ilCol = TOINCLEXCLINDEX To TOCODEINDEX Step 1
            grdTo.Col = ilCol
            grdTo.CellBackColor = vbWhite
            grdTo.Text = ""
        Next ilCol
    Next llRow
    imSpecChg = False
    imCountChg = False
    lmEnableRow = -1
    lmEnableCol = -1
    lmSpecEnableRow = -1
    lmSpecEnableCol = -1
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

'
'   mSetCommands
'   Where:
'

    If bmLoadingInfo Then
        Exit Sub
    End If
    'Update button set if all mandatory fields have data and any field altered
    If imSpecChg Or imCountChg Then
        cbcSelect.Enabled = False
        ckcIncludeDormant.Enabled = False
    Else
        cbcSelect.Enabled = True
        ckcIncludeDormant.Enabled = True
    End If
    If imSpecChg Or imCountChg Then  'At least one event added
        If imUpdateAllowed Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSpecEnableBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (grdSpec.Row < grdSpec.FixedRows) Or (grdSpec.Row >= grdSpec.Rows) Or (grdSpec.Col < grdSpec.FixedCols) Or (grdSpec.Col >= grdSpec.Cols - 1) Then
        Exit Sub
    End If
    lmSpecEnableRow = grdSpec.Row
    lmSpecEnableCol = grdSpec.Col

    Select Case grdSpec.Row
        Case SPECROW3INDEX
            Select Case grdSpec.Col
                Case NAMEINDEX 'Name
                    edcSpec.MaxLength = 80
                    edcSpec.Text = grdSpec.Text
                Case STATUSINDEX
                    smStatus = grdSpec.Text
                    If (smStatus = "") Or (smStatus = "Missing") Then
                        smStatus = "Active"
                    End If
            End Select
    End Select
    mSpecSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSpecSetShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilNoGames                     ilOrigUpper                   ilLoop                    *
'*  llRow                         ilIndex                       ilCol                     *
'*  ilCatChg                                                                              *
'******************************************************************************************

    Dim slStr As String

    If (lmSpecEnableRow >= grdSpec.FixedRows) And (lmSpecEnableRow < grdSpec.Rows) Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                    Case NAMEINDEX
                        slStr = edcSpec.Text
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> slStr Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = slStr
                        If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX)) = "" Then
                            grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
                        End If
                    Case STATUSINDEX
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> smStatus Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smStatus
                End Select
        End Select
    End If
    lmSpecEnableRow = -1
    lmSpecEnableCol = -1
    imSpecCtrlVisible = False
    edcSpec.Visible = False
    edcDropDown.Visible = False
    cmcDropDown.Visible = False
    'lbcCategory.Visible = False
    'pbcInclExcl.Visible = False
    pbcStatus.Visible = False
    'pbcYN.Visible = False
    mSetCommands
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSpecSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdSpec.Row < grdSpec.FixedRows) Or (grdSpec.Row >= grdSpec.Rows) Or (grdSpec.Col < grdSpec.FixedCols) Or (grdSpec.Col >= grdSpec.Cols - 1) Then
        Exit Sub
    End If
    imSpecCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdSpec.Col - 1 Step 1
        llColPos = llColPos + grdSpec.ColWidth(ilCol)
    Next ilCol
    Select Case grdSpec.Row
        Case SPECROW3INDEX
            Select Case grdSpec.Col
                Case NAMEINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    edcSpec.Visible = True
                    edcSpec.SetFocus
                Case STATUSINDEX
                    pbcStatus.Move grdSpec.Left + llColPos + 45, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 45, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    pbcStatus_Paint
                    pbcStatus.Visible = True
                    pbcStatus.SetFocus
            End Select
    End Select
End Sub



Private Sub mGridSpecLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilCol = 0 To grdSpec.Cols - 1 Step 1
        grdSpec.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
    grdSpec.RowHeight(0) = 15
    grdSpec.RowHeight(1) = 15
    grdSpec.RowHeight(2) = 150
    grdSpec.RowHeight(3) = fgBoxGridH
    grdSpec.RowHeight(4) = 15
    grdSpec.ColWidth(0) = 15
    grdSpec.ColWidth(1) = 15
    grdSpec.ColWidth(3) = 15
    grdSpec.ColWidth(5) = 15

    'Horizontal
    For ilCol = 1 To grdSpec.Cols - 1 Step 1
        grdSpec.Row = 1
        grdSpec.Col = ilCol
        grdSpec.CellBackColor = vbBlue
    Next ilCol
    For ilRow = grdSpec.FixedRows + 2 To grdSpec.Rows - 1 Step 3
        For ilCol = 1 To grdSpec.Cols - 1 Step 1
            grdSpec.Row = ilRow
            grdSpec.Col = ilCol
            grdSpec.CellBackColor = vbBlue
        Next ilCol
    Next ilRow
    'Vertical Line
    For ilRow = 1 To grdSpec.Rows - 1 Step 1
        grdSpec.Row = ilRow
        grdSpec.Col = 1
        grdSpec.CellBackColor = vbBlue
    Next ilRow
    For ilCol = grdSpec.FixedCols + 1 To grdSpec.Cols - 1 Step 2
        For ilRow = 1 To grdSpec.Rows - 1 Step 1
            grdSpec.Row = ilRow
            grdSpec.Col = ilCol
            grdSpec.CellBackColor = vbBlue
        Next ilRow
    Next ilCol
End Sub




Private Sub mGridSpecColumns()
    grdSpec.Row = SPECROW3INDEX - 1
    grdSpec.Col = NAMEINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, NAMEINDEX) = "Name"
    grdSpec.Col = STATUSINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, STATUSINDEX) = "Status"

End Sub

Private Sub mGridSpecColumnWidths()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llMinWidth                    ilColInc                      ilLoop                    *
'*                                                                                        *
'******************************************************************************************

    Dim llWidth As Long
    Dim ilCol As Integer

    grdSpec.ColWidth(STATUSINDEX) = 0.09 * grdSpec.Width
    llWidth = fgPanelAdj
    For ilCol = 0 To grdSpec.Cols - 1 Step 1
        If ilCol <> NAMEINDEX Then
            llWidth = llWidth + grdSpec.ColWidth(ilCol)
        End If
    Next ilCol
    grdSpec.ColWidth(NAMEINDEX) = grdSpec.Width - llWidth - 90
    llWidth = llWidth + grdSpec.ColWidth(NAMEINDEX)
    llWidth = grdSpec.Width - llWidth
    If llWidth >= 15 Then
        Do
            For ilCol = grdSpec.FixedCols To grdSpec.Cols - 1 Step 1
                If grdSpec.ColWidth(ilCol) > 15 Then
                    If ilCol = NAMEINDEX Then
                        grdSpec.ColWidth(ilCol) = grdSpec.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub


Private Sub tbcCategory_Click()
    lbcFrom(LBCFORMATINDEX).Visible = False
    lbcFrom(LBCDMAMARKETINDEX).Visible = False
    lbcFrom(LBCSTATEINDEX).Visible = False
    grdStation.Visible = False
    lbcFrom(LBCZONEINDEX).Visible = False
    pbcStationSelection.Visible = False
    lbcFrom(LBCMSAMARKETINDEX).Visible = False
    edcMetroMsg.Visible = False
    cmcImport.Enabled = False
    '11/6/14:
    If (Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) = REMOTEEXPORT Then
        If tbcCategory.SelectedItem.Index <> 5 Then
            Exit Sub
        End If
    End If
    Select Case tbcCategory.SelectedItem.Index
        Case 1
            lbcFrom(LBCDMAMARKETINDEX).Visible = True
        Case 2
            lbcFrom(LBCFORMATINDEX).Visible = True
        Case 3
            lbcFrom(LBCMSAMARKETINDEX).Visible = True
            If (Asc(tgSpf.sUsingFeatures8) And ALLOWMSASPLITCOPY) <> ALLOWMSASPLITCOPY Then
                edcMetroMsg.Visible = True
            End If
        Case 4
            lbcFrom(LBCSTATEINDEX).Visible = True
        Case 5
            mPopStations True
            grdStation.Visible = True
            pbcStationSelection.Visible = True
            cmcImport.Enabled = True
        Case 6
            lbcFrom(LBCZONEINDEX).Visible = True
    End Select

End Sub


Private Function mSaveRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCount                       ilFound                       llRow                     *
'*                                                                                        *
'******************************************************************************************

    Dim ilError As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilIefRet As Integer
    Dim slMsg As String
    Dim ilSef As Integer
    Dim tlSef As SEF

    ilError = False
    mSetMousePointer vbHourglass
    If mSpecFieldsOk() = False Then
        ilError = True
    End If
    If ilError Then
        mSetMousePointer vbDefault
        MsgBox "One or more fields not defined", vbOKOnly + vbExclamation, "Save"
        Beep
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSetMousePointer vbDefault
        mSaveRec = False
        Exit Function
    End If
    If Trim$(grdTo.TextMatrix(grdTo.FixedRows, TOINCLEXCLINDEX)) = "" Then
        mSetMousePointer vbDefault
        Beep
        MsgBox "No Items placed into the result list", vbOKOnly + vbExclamation, "Save"
        mSaveRec = False
        Exit Function
    End If
    
    If Not mMulticastOk() Then
        mSetMousePointer vbDefault
        Beep
        'MsgBox "Not All Multicast Stations included, see " & sgDBPath & "Messages\" & "MulticastMissing.txt", vbOKOnly + vbExclamation, "Save"
        ilRet = MsgBox("Not All Multicast Stations included, see " & sgDBPath & "Messages\" & "MulticastMissing.txt, Continue with save", vbInformation + vbYesNo, "Warning")
        If ilRet = vbNo Then
            mSaveRec = False
            Exit Function
        End If
    End If
    Do  'Loop until record updated or added
        If imSelectedIndex > 1 Then
            'Reread record in so lastest is obtained
            If Not mReadRec(imSelectedIndex, SETFORWRITE, 0) Then
                mSetMousePointer vbDefault
                MsgBox "Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save"
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec
        If imSelectedIndex <= 1 Then 'New selected
            tmRaf.lCode = 0
            tmRaf.iAdfCode = igAdfCode
            tmRaf.sAssigned = "N"
            slStr = Format$(gNow(), "m/d/yy")
            gPackDate slStr, tmRaf.iDateEntrd(0), tmRaf.iDateEntrd(1)
            slStr = ""
            gPackDate slStr, tmRaf.iDateDormant(0), tmRaf.iDateDormant(1)
            tmRaf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            ilRet = btrInsert(hmRaf, tmRaf, imRafRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Region RAF)"
        Else 'Old record-Update
            tmRaf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            ilRet = btrUpdate(hmRaf, tmRaf, imRafRecLen)
            slMsg = "mSaveRec (btrUpdate: Region RAF)"
        End If
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, CopySplit
        On Error GoTo 0
        'Remove SEF records
        If imSelectedIndex > 1 Then
            For ilSef = 0 To UBound(tmSef) - 1 Step 1
                tmSefSrchKey.lCode = tmSef(ilSef).lCode
                ilRet = btrGetEqual(hmSef, tlSef, imSefRecLen, tmSefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmSef)
                End If
            Next ilSef
        End If
        mMoveSefCtrlToRec
        slMsg = "mSaveRec (btrInsert: Region SEF)"
        For ilSef = 0 To UBound(tmSef) - 1 Step 1
            ilRet = btrInsert(hmSef, tmSef(ilSef), imSefRecLen, INDEXKEY0)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, CopySplit
            On Error GoTo 0
        Next ilSef
        '5882 no longer needed
        'Update IEF to reexport
'        If imSelectedIndex > 1 Then
'            tmIefSrchKey2.lCode = tmRaf.lCode
'            ilIefRet = btrGetEqual(hmIef, tmIef, imIefRecLen, tmIefSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
'            Do While (ilIefRet = BTRV_ERR_NONE) And (tmIef.lSplitRafCode = tmRaf.lCode)
'                If (tmIef.sExportStatus = "E") Then
'                    tmIef.sExportStatus = "R"
'                    ilIefRet = btrUpdate(hmIef, tmIef, imIefRecLen)
'                End If
'                ilIefRet = btrGetNext(hmIef, tmIef, imIefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'            Loop
'        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, CopySplit
    On Error GoTo 0
    imSpecChg = False
    imCountChg = False
    mSaveRec = True
    mSetMousePointer vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    mSetMousePointer vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser regions    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mPopulate(ilAdfCode As Integer)
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    Dim ilIncludeDormant As Integer

    'Repopulate if required- if sales source changed by another user while in this screen
    'imPopReqd = False
    If ckcIncludeDormant.Value = vbChecked Then
        ilIncludeDormant = True
    Else
        ilIncludeDormant = False
    End If
    ilRet = gPopRegionBox(CopySplit, ilAdfCode, "C", ilIncludeDormant, cbcSelect, tmRegionCode(), smRegionCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopRegionBox)", CopyRegn
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        cbcSelect.AddItem "[Model]", 1  'Force as first item on list
        'mPopReqd = True
    End If
    If cbcSelect.ListIndex <> 0 Then
        cbcSelect.ListIndex = 0
    Else
        cbcSelect_Change
    End If
    imCountChg = False
    imSpecChg = False
    'frcCategory.Visible = True
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Private Sub mPopCategory()
    'lbcCategory.Clear
    'lbcCategory.AddItem "Format", 0
    'lbcCategory.ItemData(lbcCategory.NewIndex) = 4
    'lbcCategory.AddItem "Market", 1
    'lbcCategory.ItemData(lbcCategory.NewIndex) = 0
    'lbcCategory.AddItem "State Name", 2
    'lbcCategory.ItemData(lbcCategory.NewIndex) = 1
    'lbcCategory.AddItem "Zip Code", 3
    'lbcCategory.ItemData(lbcCategory.NewIndex) = 2
    'lbcCategory.AddItem "Owner", 4
    'lbcCategory.ItemData(lbcCategory.NewIndex) = 3
    'lbcCategory.AddItem "Station", 5
    'lbcCategory.ItemData(lbcCategory.NewIndex) = 5

End Sub



Private Sub mPopStations(ilSetMouse As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilOwner                       ilVef                         llDropDate                *
'*  llOffAir                      llNowDate                                               *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilShtt As Integer
    Dim ilMkt As Integer
    Dim ilFormat As Integer
    Dim ilAddStation As Integer
    Dim llRow As Long
    Dim slStr As String

    If imStationPop Then
        Exit Sub
    End If
    If ilSetMouse Then
        mSetMousePointer vbHourglass
    End If
    grdStation.Redraw = False
    llRow = grdStation.FixedRows
    ilRet = gObtainStations()
    For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
        ilAddStation = True
        If llRow >= grdStation.Rows Then
            grdStation.AddItem ""
            grdStation.RowHeight(llRow) = fgBoxGridH + 15
        End If


        grdStation.TextMatrix(llRow, STATIONINDEX) = Trim$(tgStations(ilShtt).sCallLetters)
        slStr = ""
        For ilMkt = LBound(tgMarkets) To UBound(tgMarkets) - 1 Step 1
            If tgMarkets(ilMkt).iCode = tgStations(ilShtt).iMktCode Then
                slStr = Trim$(tgMarkets(ilMkt).sName)
                Exit For
            End If
        Next ilMkt
        grdStation.TextMatrix(llRow, MARKETINDEX) = slStr
        grdStation.TextMatrix(llRow, STATEINDEX) = Trim$(tgStations(ilShtt).sState)
        grdStation.TextMatrix(llRow, ZONEINDEX) = Trim$(tgStations(ilShtt).sTimeZone)
        slStr = ""
        For ilFormat = LBound(tgFormats) To UBound(tgFormats) - 1 Step 1
            If tgFormats(ilFormat).iCode = tgStations(ilShtt).iFmtCode Then
                slStr = Trim$(tgFormats(ilFormat).sName)
                Exit For
            End If
        Next ilFormat
        grdStation.TextMatrix(llRow, FORMATINDEX) = slStr

        'Code
        grdStation.TextMatrix(llRow, SHTTCODEINDEX) = Trim$(str$(tgStations(ilShtt).iCode))
        grdStation.TextMatrix(llRow, SORTINDEX) = ""
        grdStation.TextMatrix(llRow, SELECTEDINDEX) = "F"
        llRow = llRow + 1
    Next ilShtt
    imStationPop = True
    imLastStationColSorted = -1
    mStationSortCol 0
    grdStation.Row = 0
    grdStation.Col = SHTTCODEINDEX
    grdStation.Redraw = True
    If ilSetMouse Then
        mSetMousePointer vbDefault
    End If
End Sub

Private Sub mPopFormats()
    Dim ilRet As Integer

    ilRet = gObtainFormats()
    mPopFrom "Format"
End Sub

Private Sub mPopDMAMarkets()
    Dim ilRet As Integer

    ilRet = gObtainMarkets()
    mPopFrom "DMA Market"
End Sub

Private Sub mPopMSAMarkets()
    Dim ilRet As Integer
    
    ilRet = gObtainMSAMarkets()
    mPopFrom "MSA Market"
End Sub

Private Sub mPopOwners()
    'Dim ilRet As Integer
   '
   ' ilRet = gObtainOwners()
End Sub
Private Sub mPopStates()
    Dim ilRet As Integer

    ilRet = gObtainStates()
    mPopFrom "State Name"
End Sub

Private Sub mPopTimeZones()
    Dim ilRet As Integer

    ilRet = gObtainTimeZones()
    mPopFrom "Time Zone"
End Sub



Private Sub mPopFrom(slCategory As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilOwner                       ilShtt                        slState                   *
'*  slZipCode                                                                             *
'******************************************************************************************

    Dim ilMkt As Integer
    Dim ilFmt As Integer
    Dim ilTzt As Integer
    Dim ilSnt As Integer

    If StrComp(slCategory, "Station", vbTextCompare) = 0 Then
        Exit Sub
    End If
    If StrComp(slCategory, "DMA Market", vbTextCompare) = 0 Then
        lbcFrom(LBCDMAMARKETINDEX).Clear
        For ilMkt = LBound(tgMarkets) To UBound(tgMarkets) - 1 Step 1
            lbcFrom(LBCDMAMARKETINDEX).AddItem Trim$(tgMarkets(ilMkt).sName)
            lbcFrom(LBCDMAMARKETINDEX).ItemData(lbcFrom(LBCDMAMARKETINDEX).NewIndex) = tgMarkets(ilMkt).iCode
        Next ilMkt
    ElseIf StrComp(slCategory, "MSA Market", vbTextCompare) = 0 Then
        lbcFrom(LBCMSAMARKETINDEX).Clear
        If (Asc(tgSpf.sUsingFeatures8) And ALLOWMSASPLITCOPY) = ALLOWMSASPLITCOPY Then
            For ilMkt = LBound(tgMSAMarkets) To UBound(tgMSAMarkets) - 1 Step 1
                lbcFrom(LBCMSAMARKETINDEX).AddItem Trim$(tgMSAMarkets(ilMkt).sName)
                lbcFrom(LBCMSAMARKETINDEX).ItemData(lbcFrom(LBCMSAMARKETINDEX).NewIndex) = tgMSAMarkets(ilMkt).iCode
            Next ilMkt
        End If
    ElseIf StrComp(slCategory, "State Name", vbTextCompare) = 0 Then
        'For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
        '    'State
        '    slState = Trim$(tgStations(ilShtt).sState)
        '    If slState <> "" Then
        '        gFindMatch slState, 0, lbcFrom
        '        If gLastFound(lbcFrom) < 0 Then
        '            lbcFrom.AddItem slState
        '            lbcFrom.ItemData(lbcFrom.NewIndex) = tgStations(ilShtt).iCode
        '        End If
        '    End If
        'Next ilShtt
        lbcFrom(LBCSTATEINDEX).Clear
        For ilSnt = LBound(tgStates) To UBound(tgStates) - 1 Step 1
            lbcFrom(LBCSTATEINDEX).AddItem Trim$(tgStates(ilSnt).sPostalName) & " (" & Trim$(tgStates(ilSnt).sName) & ")"
            lbcFrom(LBCSTATEINDEX).ItemData(lbcFrom(LBCSTATEINDEX).NewIndex) = ilSnt
        Next ilSnt
    'ElseIf StrComp(slCategory, "Zip Code", vbTextCompare) = 0 Then
    '    For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
    '        'Zip Code
    '        slZipCode = Trim$(tgStations(ilShtt).sZip)
    '        If slZipCode <> "" Then
    '            gFindMatch slZipCode, 0, lbcFrom
    '            If gLastFound(lbcFrom) < 0 Then
    '                lbcFrom.AddItem slZipCode
    '                lbcFrom.ItemData(lbcFrom.NewIndex) = tgStations(ilShtt).iCode
    '            End If
    '        End If
    '    Next ilShtt
    'ElseIf StrComp(slCategory, "Owner", vbTextCompare) = 0 Then
    '    For ilOwner = LBound(tgOwners) To UBound(tgOwners) - 1 Step 1
    '        lbcFrom.AddItem Trim$(tgOwners(ilOwner).sLastName)
    '        lbcFrom.ItemData(lbcFrom.NewIndex) = tgOwners(ilOwner).iCode
    '    Next ilOwner
    ElseIf StrComp(slCategory, "Format", vbTextCompare) = 0 Then
        lbcFrom(LBCFORMATINDEX).Clear
        For ilFmt = LBound(tgFormats) To UBound(tgFormats) - 1 Step 1
            lbcFrom(LBCFORMATINDEX).AddItem Trim$(tgFormats(ilFmt).sName)
            lbcFrom(LBCFORMATINDEX).ItemData(lbcFrom(LBCFORMATINDEX).NewIndex) = tgFormats(ilFmt).iCode
        Next ilFmt
    ElseIf StrComp(slCategory, "Time Zone", vbTextCompare) = 0 Then
        lbcFrom(LBCZONEINDEX).Clear
        For ilTzt = LBound(tgTimeZones) To UBound(tgTimeZones) - 1 Step 1
            'lbcFrom(LBCZONEINDEX).AddItem Trim$(tgTimeZones(ilTzt).sCSIName) & ", " & Trim$(tgTimeZones(ilTzt).sName)
            Select Case Left$(Trim$(tgTimeZones(ilTzt).sCSIName), 1)
                Case "E"
                    lbcFrom(LBCZONEINDEX).AddItem Trim$(tgTimeZones(ilTzt).sName) & " (ETZ)"
                    lbcFrom(LBCZONEINDEX).ItemData(lbcFrom(LBCZONEINDEX).NewIndex) = tgTimeZones(ilTzt).iCode
                Case "C"
                    lbcFrom(LBCZONEINDEX).AddItem Trim$(tgTimeZones(ilTzt).sName) & " (CTZ)"
                    lbcFrom(LBCZONEINDEX).ItemData(lbcFrom(LBCZONEINDEX).NewIndex) = tgTimeZones(ilTzt).iCode
                Case "M"
                    lbcFrom(LBCZONEINDEX).AddItem Trim$(tgTimeZones(ilTzt).sName) & " (MTZ)"
                    lbcFrom(LBCZONEINDEX).ItemData(lbcFrom(LBCZONEINDEX).NewIndex) = tgTimeZones(ilTzt).iCode
                Case "P"
                    lbcFrom(LBCZONEINDEX).AddItem Trim$(tgTimeZones(ilTzt).sName) & " (PTZ)"
                    lbcFrom(LBCZONEINDEX).ItemData(lbcFrom(LBCZONEINDEX).NewIndex) = tgTimeZones(ilTzt).iCode
            End Select
        Next ilTzt
    End If
End Sub

Private Function mSpecColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    mSpecColOk = True
    If grdSpec.ColWidth(grdSpec.Col) <= 15 Then
        mSpecColOk = False
        Exit Function
    End If
    If grdSpec.CellBackColor = LIGHTYELLOW Then
        mSpecColOk = False
        Exit Function
    End If

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    Dim slStr As String
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)
    gFindMatch slStr, 0, cbcSelect    'Determine if name exist
    If gLastFound(cbcSelect) <> -1 Then   'Name found
        If gLastFound(cbcSelect) <> imSelectedIndex Then
            If Trim$(slStr) = cbcSelect.List(gLastFound(cbcSelect)) Then
                Beep
                MsgBox "Advertiser Region name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                mOKName = False
                Exit Function
            End If
        End If
    End If
    mOKName = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim slStr As String
    tmRaf.sName = grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)
    tmRaf.sAbbr = ""
    tmRaf.sCustom = "N"
    tmRaf.sUnused = ""
    tmRaf.sInclExcl = ""
    tmRaf.sCategory = ""
    tmRaf.iAudPct = 0
    tmRaf.sShowNoProposal = "N"
    tmRaf.sShowOnOrder = "N"
    tmRaf.sShowOnInvoice = "N"


    If Trim$(grdTo.TextMatrix(grdTo.FixedRows + 1, TOINCLEXCLINDEX)) = "" Then
        slStr = Trim$(grdTo.TextMatrix(grdTo.FixedRows, TOINCLEXCLINDEX))
        Select Case UCase$(slStr)
            Case "INCL"
                tmRaf.sInclExcl = "I"
            Case "EXCL"
                tmRaf.sInclExcl = "E"
            Case Else
                tmRaf.sInclExcl = "I"
        End Select

        slStr = grdTo.TextMatrix(grdTo.FixedRows, TOCATEGORYINDEX)
        Select Case UCase$(slStr)
            Case "DMA MARKET"
                tmRaf.sCategory = "M"
            Case "MSA MARKET"
                tmRaf.sCategory = "A"
            Case "STATE"
                tmRaf.sCategory = "N"
            Case "ZONE"
                tmRaf.sCategory = "T"
            Case "FORMAT"
                tmRaf.sCategory = "F"
            Case "STATION"
                tmRaf.sCategory = "S"
        End Select
    End If
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX)
    Select Case UCase$(slStr)
        Case "ACTIVE"
            tmRaf.sState = "A"
        Case "DORMANT"
            If tmRaf.sState <> "D" Then
                slStr = Format$(gNow(), "m/d/yy")
                gPackDate slStr, tmRaf.iDateDormant(0), tmRaf.iDateDormant(1)
            End If
            tmRaf.sState = "D"
        Case Else
            tmRaf.sState = "A"
    End Select
    tmRaf.lRegionCode = 0
    tmRaf.sType = "C"       'Split Copy

    'tmRaf.sUnused = ""

    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveSefCtrlToRec               *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveSefCtrlToRec()
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim llRow As Long
    Dim ilSeqNo As Integer
    Dim slInclExcl As String
    Dim ilPos As Integer

    ReDim tmSef(0 To 0) As SEF
    ilSeqNo = 1
    For llRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
        slStr = Trim$(grdTo.TextMatrix(llRow, TOCATEGORYINDEX))
        If slStr <> "" Then
            slInclExcl = Left$(grdTo.TextMatrix(llRow, TOINCLEXCLINDEX), 1)
            Select Case UCase$(slStr)
                Case "DMA MARKET"
                    ilUpper = UBound(tmSef)
                    tmSef(ilUpper).iIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                    tmSef(ilUpper).lLongCode = 0
                    tmSef(ilUpper).sName = ""
                    tmSef(ilUpper).sCategory = "M"
                    tmSef(ilUpper).sInclExcl = slInclExcl
                    tmSef(ilUpper).iSeqNo = ilSeqNo
                    ilSeqNo = ilSeqNo + 1
                    ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
                Case "MSA MARKET"
                    ilUpper = UBound(tmSef)
                    tmSef(ilUpper).iIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                    tmSef(ilUpper).lLongCode = 0
                    tmSef(ilUpper).sName = ""
                    tmSef(ilUpper).sCategory = "A"
                    tmSef(ilUpper).sInclExcl = slInclExcl
                    tmSef(ilUpper).iSeqNo = ilSeqNo
                    ilSeqNo = ilSeqNo + 1
                    ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
                Case "STATE"
                    ilUpper = UBound(tmSef)
                    slStr = grdTo.TextMatrix(llRow, TONAMEINDEX)
                    ilPos = InStr(1, slStr, "(", vbTextCompare)
                    If ilPos >= 1 Then
                        tmSef(ilUpper).sName = Trim$(Left(slStr, ilPos - 1))
                    Else
                        tmSef(ilUpper).sName = Trim$(grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX))
                    End If
                    tmSef(ilUpper).iIntCode = 0
                    tmSef(ilUpper).lLongCode = 0
                    tmSef(ilUpper).sCategory = "N"
                    tmSef(ilUpper).sInclExcl = slInclExcl
                    tmSef(ilUpper).iSeqNo = ilSeqNo
                    ilSeqNo = ilSeqNo + 1
                    ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
                Case "ZONE"
                    ilUpper = UBound(tmSef)
                    tmSef(ilUpper).iIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                    tmSef(ilUpper).lLongCode = 0
                    tmSef(ilUpper).sName = ""
                    tmSef(ilUpper).sCategory = "T"
                    tmSef(ilUpper).sInclExcl = slInclExcl
                    tmSef(ilUpper).iSeqNo = ilSeqNo
                    ilSeqNo = ilSeqNo + 1
                    ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
                Case "FORMAT"
                    ilUpper = UBound(tmSef)
                    tmSef(ilUpper).iIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                    tmSef(ilUpper).lLongCode = 0
                    tmSef(ilUpper).sName = ""
                    tmSef(ilUpper).sCategory = "F"
                    tmSef(ilUpper).sInclExcl = slInclExcl
                    tmSef(ilUpper).iSeqNo = ilSeqNo
                    ilSeqNo = ilSeqNo + 1
                    ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
                Case "STATION"
                    ilUpper = UBound(tmSef)
                    tmSef(ilUpper).iIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                    tmSef(ilUpper).lLongCode = 0
                    tmSef(ilUpper).sName = ""
                    tmSef(ilUpper).sCategory = "S"
                    tmSef(ilUpper).sInclExcl = slInclExcl
                    tmSef(ilUpper).iSeqNo = ilSeqNo
                    ilSeqNo = ilSeqNo + 1
                    ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
            End Select
        End If
    Next llRow
    For ilLoop = 0 To UBound(tmSef) - 1 Step 1
        tmSef(ilLoop).lCode = 0
        tmSef(ilLoop).lRafCode = tmRaf.lCode
        tmSef(ilLoop).sUnused = ""
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveSefRecToCtrl               *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record to Controls        *
'*                                                     *
'*******************************************************
Sub mMoveSEFRecToCtrl()

    Dim ilSef As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llRow As Long
    Dim ilCol As Integer
    Dim slCategory As String
    Dim slInclExcl As String
    Dim ilPos As Integer

    grdTo.Redraw = False
    For ilSef = 0 To UBound(tmSef) - 1 Step 1
        DoEvents
        slCategory = UCase$(tmRaf.sCategory)
        slInclExcl = tmRaf.sInclExcl
        If Trim$(tmSef(ilSef).sCategory) <> "" Then
            slCategory = tmSef(ilSef).sCategory
            slInclExcl = tmSef(ilSef).sInclExcl
        End If
        If slInclExcl = "E" Then
            slInclExcl = "Excl"
        Else
            slInclExcl = "Incl"
        End If
        Select Case UCase$(slCategory)
            Case "M"
                For ilLoop = 0 To lbcFrom(LBCDMAMARKETINDEX).ListCount - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcFrom(LBCDMAMARKETINDEX).ItemData(ilLoop)) Then
                        lbcFrom(LBCDMAMARKETINDEX).Selected(ilLoop) = True
                        mXFerFromTo 1, slInclExcl, -1
                        Exit For
                    End If
                Next ilLoop
            Case "A"
                For ilLoop = 0 To lbcFrom(LBCMSAMARKETINDEX).ListCount - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcFrom(LBCMSAMARKETINDEX).ItemData(ilLoop)) Then
                        lbcFrom(LBCMSAMARKETINDEX).Selected(ilLoop) = True
                        mXFerFromTo 3, slInclExcl, -1
                        Exit For
                    End If
                Next ilLoop
            Case "N"
                For ilLoop = 0 To lbcFrom(LBCSTATEINDEX).ListCount - 1 Step 1
                    slStr = Trim$(lbcFrom(LBCSTATEINDEX).List(ilLoop))
                    ilPos = InStr(1, slStr, "(", vbTextCompare)
                    If ilPos > 0 Then
                        slStr = Trim$(Left$(slStr, ilPos - 1))
                    End If
                    If StrComp(Trim$(tmSef(ilSef).sName), slStr, vbTextCompare) = 0 Then
                        lbcFrom(LBCSTATEINDEX).Selected(ilLoop) = True
                        mXFerFromTo 4, slInclExcl, -1
                        Exit For
                    End If
                Next ilLoop
            Case "T"
                For ilLoop = 0 To lbcFrom(LBCZONEINDEX).ListCount - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcFrom(LBCZONEINDEX).ItemData(ilLoop)) Then
                        lbcFrom(LBCZONEINDEX).Selected(ilLoop) = True
                        mXFerFromTo 6, slInclExcl, -1
                        Exit For
                    End If
                Next ilLoop
            Case "F"
                For ilLoop = 0 To lbcFrom(LBCFORMATINDEX).ListCount - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcFrom(LBCFORMATINDEX).ItemData(ilLoop)) Then
                        lbcFrom(LBCFORMATINDEX).Selected(ilLoop) = True
                        mXFerFromTo 2, slInclExcl, -1
                        Exit For
                    End If
                Next ilLoop
            Case "S"
                mPopStations False
                For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(grdStation.TextMatrix(llRow, SHTTCODEINDEX)) Then
                        grdStation.Row = llRow
                        For ilCol = STATIONINDEX To FORMATINDEX Step 1
                            grdStation.Col = ilCol
                            grdStation.CellBackColor = GRAY
                        Next ilCol
                        mXFerFromTo 5, slInclExcl, llRow
                        Exit For
                    End If
                Next llRow
        End Select
    Next ilSef
    imCountChg = False
    imSpecChg = False
    mSetCommands
    grdTo.Redraw = True
End Sub

Private Sub mStationSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
        slStr = Trim$(grdStation.TextMatrix(llRow, STATIONINDEX))
        If slStr <> "" Then
            slSort = UCase$(Trim$(grdStation.TextMatrix(llRow, ilCol)))
            If slSort = "" Then
                slSort = "~"
            End If
            slStr = grdStation.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastStationColSorted) Or ((ilCol = imLastStationColSorted) And (imLastStationSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdStation.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdStation.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastStationColSorted Then
        imLastStationColSorted = SORTINDEX
    Else
        imLastStationColSorted = -1
        imLastStationSort = -1
    End If
    gGrid_SortByCol grdStation, STATIONINDEX, SORTINDEX, imLastStationColSorted, imLastStationSort
    imLastStationColSorted = ilCol
End Sub
Private Sub mGridStationLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdStation.Rows - 1 Step 1
        grdStation.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdStation.Cols - 1 Step 1
        grdStation.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridStationColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdStation.Row = grdStation.FixedRows - 1
    grdStation.Col = STATIONINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Station"
    grdStation.Col = MARKETINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "DMA Market"
    grdStation.Col = STATEINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "State"
    grdStation.Col = ZONEINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Zone"
    grdStation.Col = FORMATINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Format"
    grdStation.Col = SHTTCODEINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Sef Code"
    grdStation.Col = SORTINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Sort"
    grdStation.Col = SELECTEDINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Selected"

End Sub

Private Sub mGridStationColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdStation.ColWidth(SHTTCODEINDEX) = 0
    grdStation.ColWidth(SORTINDEX) = 0
    grdStation.ColWidth(SELECTEDINDEX) = 0
    grdStation.ColWidth(STATIONINDEX) = 0.12 * grdStation.Width
    grdStation.ColWidth(MARKETINDEX) = 0.2 * grdStation.Width
    grdStation.ColWidth(STATEINDEX) = 0.072 * grdStation.Width
    grdStation.ColWidth(ZONEINDEX) = 0.1 * grdStation.Width
    grdStation.ColWidth(FORMATINDEX) = 0.2 * grdStation.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdStation.Width
    For ilCol = 0 To grdStation.Cols - 1 Step 1
        llWidth = llWidth + grdStation.ColWidth(ilCol)
        If (grdStation.ColWidth(ilCol) > 15) And (grdStation.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdStation.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdStation.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdStation.Width
            For ilCol = 0 To grdStation.Cols - 1 Step 1
                If (grdStation.ColWidth(ilCol) > 15) And (grdStation.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdStation.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdStation.FixedCols To grdStation.Cols - 1 Step 1
                If grdStation.ColWidth(ilCol) > 15 Then
                    ilColInc = grdStation.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdStation.ColWidth(ilCol) = grdStation.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub





Private Sub mGridToColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdTo.Row = grdTo.FixedRows - 1
    grdTo.Col = TOINCLEXCLINDEX
    grdTo.CellFontBold = False
    grdTo.CellFontName = "Arial"
    grdTo.CellFontSize = 6.75
    grdTo.CellForeColor = vbBlue
    grdTo.CellBackColor = vbWhite   'LIGHTBLUE
    grdTo.TextMatrix(grdTo.Row, grdTo.Col) = "I/E"
    grdTo.Col = TONAMEINDEX
    grdTo.CellFontBold = False
    grdTo.CellFontName = "Arial"
    grdTo.CellFontSize = 6.75
    grdTo.CellForeColor = vbBlue
    grdTo.CellBackColor = vbWhite   'LIGHTBLUE
    grdTo.TextMatrix(grdTo.Row, grdTo.Col) = "Name"
    grdTo.Col = TOCATEGORYINDEX
    grdTo.CellFontBold = False
    grdTo.CellFontName = "Arial"
    grdTo.CellFontSize = 6.75
    grdTo.CellForeColor = vbBlue
    grdTo.CellBackColor = vbWhite   'LIGHTBLUE
    grdTo.TextMatrix(grdTo.Row, grdTo.Col) = "Category"

End Sub

Private Sub mGridToColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdTo.ColWidth(TOCODEINDEX) = 0
    grdTo.ColWidth(TOINCLEXCLINDEX) = 0.12 * grdTo.Width
    grdTo.ColWidth(TONAMEINDEX) = 0.6 * grdTo.Width
    grdTo.ColWidth(TOCATEGORYINDEX) = 0.2 * grdTo.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdTo.Width
    For ilCol = 0 To grdTo.Cols - 1 Step 1
        llWidth = llWidth + grdTo.ColWidth(ilCol)
        If (grdTo.ColWidth(ilCol) > 15) And (grdTo.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdTo.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdTo.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdTo.Width
            For ilCol = 0 To grdTo.Cols - 1 Step 1
                If (grdTo.ColWidth(ilCol) > 15) And (grdTo.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdTo.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdTo.FixedCols To grdTo.Cols - 1 Step 1
                If grdTo.ColWidth(ilCol) > 15 Then
                    ilColInc = grdTo.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdTo.ColWidth(ilCol) = grdTo.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub

Private Sub mGridToLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdTo.Rows - 1 Step 1
        grdTo.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdTo.Cols - 1 Step 1
        grdTo.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub


Private Sub mXFerFromTo(ilCategoryIndex As Integer, slInclExcl As String, llFromRow As Long)
    Dim llRow As Long
    Dim llToRow As Long
    Dim llLoopRow As Long
    Dim ilIndex As Integer
    Dim slStr As String
    Dim ilPos As Integer

    If (ilCategoryIndex <> 5) Then
        If ilCategoryIndex = 1 Then
            ilIndex = LBCDMAMARKETINDEX
        ElseIf ilCategoryIndex = 2 Then
            ilIndex = LBCFORMATINDEX
        ElseIf ilCategoryIndex = 3 Then
            ilIndex = LBCMSAMARKETINDEX
        ElseIf ilCategoryIndex = 4 Then
            ilIndex = LBCSTATEINDEX
        ElseIf ilCategoryIndex = 6 Then
            ilIndex = LBCZONEINDEX
        Else
            Exit Sub
        End If
        '5/19/18: retain list order
        'For llLoopRow = lbcFrom(ilIndex).ListCount - 1 To 0 Step -1
        For llLoopRow = 0 To lbcFrom(ilIndex).ListCount - 1 Step 1
            If lbcFrom(ilIndex).Selected(llLoopRow) Then
                llToRow = -1
                For llRow = grdTo.Rows - 1 To grdTo.FixedRows Step -1
                    If Trim$(grdTo.TextMatrix(llRow, TOINCLEXCLINDEX)) = "" Then
                        llToRow = llRow
                    End If
                Next llRow
                If llToRow = -1 Then
                    grdTo.AddItem ""
                    llToRow = grdTo.Rows - 1
                    grdTo.RowHeight(llToRow) = fgBoxGridH + 15
                End If
                grdTo.TextMatrix(llToRow, TOINCLEXCLINDEX) = slInclExcl
                grdTo.TextMatrix(llToRow, TONAMEINDEX) = lbcFrom(ilIndex).List(llLoopRow)
                Select Case ilCategoryIndex
                    Case 1
                        grdTo.TextMatrix(llToRow, TOCATEGORYINDEX) = "DMA Market"
                    Case 2
                        grdTo.TextMatrix(llToRow, TOCATEGORYINDEX) = "Format"
                    Case 3
                        grdTo.TextMatrix(llToRow, TOCATEGORYINDEX) = "MSA Market"
                    Case 4
                        grdTo.TextMatrix(llToRow, TOCATEGORYINDEX) = "State"
                    Case 6
                        grdTo.TextMatrix(llToRow, TOCATEGORYINDEX) = "Zone"
                End Select
                grdTo.TextMatrix(llToRow, TOCODEINDEX) = lbcFrom(ilIndex).ItemData(llLoopRow)
                '5/19/18: retain list order
                'lbcFrom(ilIndex).RemoveItem llLoopRow
            End If
        Next llLoopRow
        '5/19/18: retain list order
        For llLoopRow = lbcFrom(ilIndex).ListCount - 1 To 0 Step -1
            If lbcFrom(ilIndex).Selected(llLoopRow) Then
                lbcFrom(ilIndex).RemoveItem llLoopRow
            End If
        Next llLoopRow

    Else
        '12/15/14: Speed-up inital showing of previously selected stations.
        'For llLoopRow = grdStation.Rows - 1 To grdStation.FixedRows Step -1
        If llFromRow = -1 Then
            '5/19/18: retain list order
            'llLoopRow = grdStation.Rows - 1
            llLoopRow = grdStation.FixedRows
        Else
            llLoopRow = llFromRow
        End If
        '5/19/18: retain list order
        'Do While llLoopRow >= grdStation.FixedRows
        Do While llLoopRow <= grdStation.Rows - 1
            grdStation.Row = llLoopRow
            grdStation.Col = STATIONINDEX
            If grdStation.CellBackColor = GRAY Then
                llToRow = -1
                For llRow = grdTo.Rows - 1 To grdTo.FixedRows Step -1
                    If Trim$(grdTo.TextMatrix(llRow, TOINCLEXCLINDEX)) = "" Then
                        llToRow = llRow
                    End If
                Next llRow
                If llToRow = -1 Then
                    grdTo.AddItem ""
                    llToRow = grdTo.Rows - 1
                    grdTo.RowHeight(llToRow) = fgBoxGridH + 15
                End If
                grdTo.TextMatrix(llToRow, TOINCLEXCLINDEX) = slInclExcl
                grdTo.TextMatrix(llToRow, TONAMEINDEX) = grdStation.TextMatrix(llLoopRow, STATIONINDEX)
                grdTo.TextMatrix(llToRow, TOCATEGORYINDEX) = "Station"
                grdTo.TextMatrix(llToRow, TOCODEINDEX) = grdStation.TextMatrix(llLoopRow, SHTTCODEINDEX)
                '5/19/18: retain list order
                'grdStation.RemoveItem llLoopRow
            End If
            '12/15/14: Speed-up inital transfer of stations
            If llFromRow <> -1 Then
                Exit Do
            End If
            '5/19/18: retain list order
            'llLoopRow = llLoopRow - 1
            llLoopRow = llLoopRow + 1
        'Next llLoopRow
        Loop
        '5/19/18: retain list order
        If llFromRow = -1 Then
            llLoopRow = grdStation.Rows - 1
        Else
            llLoopRow = llFromRow
        End If
        Do While llLoopRow >= grdStation.FixedRows
            grdStation.Row = llLoopRow
            grdStation.Col = STATIONINDEX
            If grdStation.CellBackColor = GRAY Then
                grdStation.RemoveItem llLoopRow
            End If
            '12/15/14: Speed-up inital transfer of stations
            If llFromRow <> -1 Then
                Exit Do
            End If
            llLoopRow = llLoopRow - 1
        'Next llLoopRow
        Loop
        
    End If
    If Trim$(grdTo.TextMatrix(grdTo.FixedRows, TOINCLEXCLINDEX)) <> "" Then
        cmcClear.Enabled = True
    Else
        cmcClear.Enabled = False
    End If
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)) = "" Then
        If Trim$(grdTo.TextMatrix(grdTo.FixedRows + 1, TOINCLEXCLINDEX)) = "" Then
            Select Case grdTo.TextMatrix(grdTo.FixedRows, TOCATEGORYINDEX)
                Case "DMA Market"
                    grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX)
                Case "Format"
                    grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX)
                Case "MSA Market"
                    grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX)
                Case "State"
                    slStr = grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX)
                    ilPos = InStr(1, slStr, ",", vbTextCompare)
                    If ilPos >= 1 Then
                        grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = Trim$(Mid$(slStr, ilPos + 1))
                    Else
                        grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX)
                    End If
                Case "Zone"
                    slStr = grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX)
                    ilPos = InStr(1, slStr, ",", vbTextCompare)
                    If ilPos >= 1 Then
                        grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = Trim$(Mid$(slStr, ilPos + 1))
                    Else
                        grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX)
                    End If
                Case "Station"
                    grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX)
            End Select
            grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
        End If
    End If
    imCountChg = True
    mSetCommands
End Sub

Private Sub mXFerToFrom()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                                                                                 *
'******************************************************************************************

    Dim llRow As Long
    Dim llFromRow As Long
    Dim llToRow As Long
    Dim ilShtt As Integer
    Dim ilMkt As Integer
    Dim ilFormat As Integer
    Dim slStr As String

    llFromRow = grdTo.Row
    Select Case grdTo.TextMatrix(llFromRow, TOCATEGORYINDEX)
        Case "DMA Market"
            lbcFrom(LBCDMAMARKETINDEX).AddItem grdTo.TextMatrix(llFromRow, TONAMEINDEX)
            lbcFrom(LBCDMAMARKETINDEX).ItemData(lbcFrom(LBCDMAMARKETINDEX).NewIndex) = Val(grdTo.TextMatrix(llFromRow, TOCODEINDEX))
            grdTo.RemoveItem llFromRow
            grdTo.AddItem ""
        Case "Format"
            lbcFrom(LBCFORMATINDEX).AddItem grdTo.TextMatrix(llFromRow, TONAMEINDEX)
            lbcFrom(LBCFORMATINDEX).ItemData(lbcFrom(LBCFORMATINDEX).NewIndex) = Val(grdTo.TextMatrix(llFromRow, TOCODEINDEX))
            grdTo.RemoveItem llFromRow
            grdTo.AddItem ""
        Case "MSA Market"
            lbcFrom(LBCMSAMARKETINDEX).AddItem grdTo.TextMatrix(llFromRow, TONAMEINDEX)
            lbcFrom(LBCMSAMARKETINDEX).ItemData(lbcFrom(LBCMSAMARKETINDEX).NewIndex) = Val(grdTo.TextMatrix(llFromRow, TOCODEINDEX))
            grdTo.RemoveItem llFromRow
            grdTo.AddItem ""
        Case "State"
            lbcFrom(LBCSTATEINDEX).AddItem grdTo.TextMatrix(llFromRow, TONAMEINDEX)
            lbcFrom(LBCSTATEINDEX).ItemData(lbcFrom(LBCSTATEINDEX).NewIndex) = Val(grdTo.TextMatrix(llFromRow, TOCODEINDEX))
            grdTo.RemoveItem llFromRow
            grdTo.AddItem ""
        Case "Station"
            llToRow = -1
            For llRow = grdStation.Rows - 1 To grdStation.FixedRows Step -1
                If Trim$(grdStation.TextMatrix(llRow, TOINCLEXCLINDEX)) = "" Then
                    llToRow = llRow
                End If
            Next llRow
            If llToRow = -1 Then
                grdStation.AddItem ""
                llToRow = grdStation.Rows - 1
                grdStation.RowHeight(llToRow) = fgBoxGridH + 15
            End If
            For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
                If Val(grdTo.TextMatrix(llFromRow, TOCODEINDEX)) = tgStations(ilShtt).iCode Then
                    grdStation.TextMatrix(llToRow, STATIONINDEX) = Trim$(tgStations(ilShtt).sCallLetters)
                    slStr = ""
                    For ilMkt = LBound(tgMarkets) To UBound(tgMarkets) - 1 Step 1
                        If tgMarkets(ilMkt).iCode = tgStations(ilShtt).iMktCode Then
                            slStr = Trim$(tgMarkets(ilMkt).sName)
                            Exit For
                        End If
                    Next ilMkt
                    grdStation.TextMatrix(llToRow, MARKETINDEX) = slStr
                    grdStation.TextMatrix(llToRow, STATEINDEX) = Trim$(tgStations(ilShtt).sState)
                    slStr = ""
                    For ilFormat = LBound(tgFormats) To UBound(tgFormats) - 1 Step 1
                        If tgFormats(ilFormat).iCode = tgStations(ilShtt).iFmtCode Then
                            slStr = Trim$(tgFormats(ilFormat).sName)
                            Exit For
                        End If
                    Next ilFormat
                    grdStation.TextMatrix(llToRow, FORMATINDEX) = slStr
                    grdStation.TextMatrix(llToRow, SHTTCODEINDEX) = Val(grdTo.TextMatrix(llFromRow, TOCODEINDEX))
                    grdTo.RemoveItem llFromRow
                    grdTo.AddItem ""
                    imStationMoved = True
                    Exit For
                End If
            Next ilShtt
        Case "Zone"
            lbcFrom(LBCZONEINDEX).AddItem grdTo.TextMatrix(llFromRow, TONAMEINDEX)
            lbcFrom(LBCZONEINDEX).ItemData(lbcFrom(LBCZONEINDEX).NewIndex) = Val(grdTo.TextMatrix(llFromRow, TOCODEINDEX))
            grdTo.RemoveItem llFromRow
            grdTo.AddItem ""
    End Select
    If Trim$(grdTo.TextMatrix(grdTo.FixedRows, TOINCLEXCLINDEX)) <> "" Then
        cmcClear.Enabled = True
    Else
        cmcClear.Enabled = False
    End If
End Sub

Private Sub mResortStations()
    If imLastStationSort = flexSortStringNoCaseAscending Then
        imLastStationSort = flexSortStringNoCaseDescending
    Else
        imLastStationSort = flexSortStringNoCaseAscending
    End If
    mStationSortCol imLastStationColSorted
    imStationMoved = False
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    cmcCancel_Click
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFile                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*
'*      1-6-05 add flag to tell whether to intialize the
'*             list box of data because multiple files
'*             selection is allowed.
'*******************************************************
Private Function mReadStationFile(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim blFound As Boolean
    Dim hlFrom As Integer
    Dim blMissingStationsLogged As Boolean
    
    mSetMousePointer vbHourglass
    blMissingStationsLogged = False
    ilRet = 0
    'On Error GoTo mReadFileErr:
    'hlFrom = FreeFile
    'Open slFromFile For Input Access Read As hlFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        mSetMousePointer vbDefault
        Close hlFrom
        MsgBox "Open " & slFromFile & ", Error #" & str$(ilRet), vbExclamation, "Open Error"
        cmcCancel.SetFocus
        mReadStationFile = False
        Exit Function
    End If
    Err.Clear
    Do
        'On Error GoTo mReadFileErr:
        If EOF(hlFrom) Then
            Exit Do
        End If
        Line Input #hlFrom, slLine
        On Error GoTo 0
        ilRet = Err.Number
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                slLine = UCase(Trim$(slLine))
                blFound = False
                For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
                    If grdStation.TextMatrix(llRow, STATIONINDEX) <> "" Then
                        If UCase(Trim$(grdStation.TextMatrix(llRow, STATIONINDEX))) = slLine Then
                            grdStation.Row = llRow
                            grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T"
                            For llCol = STATIONINDEX To FORMATINDEX Step 1
                                grdStation.Col = llCol
                                grdStation.CellBackColor = GRAY
                            Next llCol
                            blFound = True
                            Exit For
                        End If
                    End If
                Next llRow
                If Not blFound Then
                    For llRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
                        If grdTo.TextMatrix(llRow, TONAMEINDEX) <> "" Then
                            If UCase(Trim$(grdTo.TextMatrix(llRow, TONAMEINDEX))) = slLine Then
                                blFound = True
                                Exit For
                            End If
                        End If
                    Next llRow
                End If
                If Not blFound Then
                    If Not blMissingStationsLogged Then
                        blMissingStationsLogged = True
                        gLogMsg slLine, "StationsNotFound.Txt", True
                    Else
                        gLogMsg slLine, "StationsNotFound.Txt", False
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    Close hlFrom
    mSetMousePointer vbDefault
    If blMissingStationsLogged Then
        gMsgBox "See " & sgDBPath & "Messages\" & "StationsNotFound.txt", vbOKOnly + vbInformation, "Stations Not Found"
    End If
    mReadStationFile = True
    Exit Function
'mReadFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function

Private Sub mSetMousePointer(ilPointer As Integer)
    gSetMousePointer grdSpec, grdStation, ilPointer
    gSetMousePointer grdTo, grdTo, ilPointer
    Screen.MousePointer = ilPointer  'Wait
End Sub

Private Function mMulticastOk() As Boolean
    Dim llRow As Long
    Dim slStr As String
    Dim slInclExcl As String
    Dim ilPass As Integer
    Dim ilIntCode As Integer
    Dim slName As String
    Dim ilPos As Integer
    Dim ilIndex As Integer
    Dim ilShttCode As Integer
    Dim slSQLQuery As String
    Dim ilShtt As Integer
    
    bmMulticastMissing = False
    edcCheckingMsg.Visible = True
    ReDim imMissingShttCode(0 To 0) As Integer
    
    For ilPass = 0 To 1 Step 1
        For llRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
            slStr = Trim$(grdTo.TextMatrix(llRow, TOCATEGORYINDEX))
            If slStr <> "" Then
                slInclExcl = UCase(Left$(grdTo.TextMatrix(llRow, TOINCLEXCLINDEX), 1))
                If ((ilPass = 0) And (slInclExcl = "I")) Or ((ilPass = 0) And (slInclExcl = "")) Or ((ilPass = 1) And (slInclExcl = "E")) Then
                    Select Case UCase$(slStr)
                        Case "DMA MARKET"
                            ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                            For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                                If tgStations(ilShtt).iMktCode = ilIntCode Then
                                    'If tgStations(ilShtt).lMultiCastGroupID > 0 Then
                                        'Check all stations in MulticastGroup
                                        slSQLQuery = "Select shttCode, shttMktCode as CheckCode From shtt Where shttMultiCastGroupID = " & tgStations(ilShtt).lMultiCastGroupID
                                        mCheckCode tgStations(ilShtt).iCode, tgStations(ilShtt).lMultiCastGroupID, ilPass, slSQLQuery, ilIntCode
                                    'End If
                                End If
                            Next ilShtt
                        Case "MSA MARKET"
                            ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                            For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                                If tgStations(ilShtt).iMetCode = ilIntCode Then
                                    'If tgStations(ilShtt).lMultiCastGroupID > 0 Then
                                        'Check all stations in MulticastGroup
                                        slSQLQuery = "Select shttCode, shttMetCode as CheckCode From shtt Where shttMultiCastGroupID = " & tgStations(ilShtt).lMultiCastGroupID
                                        mCheckCode tgStations(ilShtt).iCode, tgStations(ilShtt).lMultiCastGroupID, ilPass, slSQLQuery, ilIntCode
                                    'End If
                                End If
                            Next ilShtt
                        Case "STATE"
                            slStr = grdTo.TextMatrix(llRow, TONAMEINDEX)
                            ilPos = InStr(1, slStr, "(", vbTextCompare)
                            If ilPos >= 1 Then
                                slName = Trim$(Left(slStr, ilPos - 1))
                            Else
                                slName = Trim$(grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX))
                            End If
                            For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                                If Trim$(tgStations(ilShtt).sState) = slName Then
                                    'If tgStations(ilShtt).lMultiCastGroupID > 0 Then
                                        'Check all stations in MulticastGroup
                                        slSQLQuery = "Select shttCode as CheckCode From shtt Where shttMultiCastGroupID = " & tgStations(ilShtt).lMultiCastGroupID
                                        mCheckName tgStations(ilShtt).iCode, tgStations(ilShtt).lMultiCastGroupID, ilPass, slSQLQuery, slName
                                    'End If
                                End If
                            Next ilShtt
                        Case "ZONE"
                            ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                            For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                                If tgStations(ilShtt).iTztCode = ilIntCode Then
                                    'If tgStations(ilShtt).lMultiCastGroupID > 0 Then
                                        'Check all stations in MulticastGroup
                                        slSQLQuery = "Select shttCode, shttTztCode as CheckCode From shtt Where shttMultiCastGroupID = " & tgStations(ilShtt).lMultiCastGroupID
                                        mCheckCode tgStations(ilShtt).iCode, tgStations(ilShtt).lMultiCastGroupID, ilPass, slSQLQuery, ilIntCode
                                    'End If
                                End If
                            Next ilShtt
                        Case "FORMAT"
                            ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                            For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                                If tgStations(ilShtt).iFmtCode = ilIntCode Then
                                    'If tgStations(ilShtt).lMultiCastGroupID > 0 Then
                                        'Check all stations in MulticastGroup
                                        slSQLQuery = "Select shttCode, shttFmtCode as CheckCode From shtt Where shttMultiCastGroupID = " & tgStations(ilShtt).lMultiCastGroupID
                                        mCheckCode tgStations(ilShtt).iCode, tgStations(ilShtt).lMultiCastGroupID, ilPass, slSQLQuery, ilIntCode
                                    'End If
                                End If
                            Next ilShtt
                        Case "STATION"
                            ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                            For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                                If tgStations(ilShtt).iCode = ilIntCode Then
                                    'If tgStations(ilShtt).lMultiCastGroupID > 0 Then
                                        'Check all stations in MulticastGroup
                                        slSQLQuery = "Select shttCode, shttCode as CheckCode From shtt Where shttMultiCastGroupID = " & tgStations(ilShtt).lMultiCastGroupID
                                        mCheckCode tgStations(ilShtt).iCode, tgStations(ilShtt).lMultiCastGroupID, ilPass, slSQLQuery, ilIntCode
                                    'End If
                                End If
                            Next ilShtt
                    End Select
                End If
            End If
        Next llRow
    Next ilPass
    If bmMulticastMissing Then
        gLogMsgWODT "C", hmUnMatch, ""
    End If
    mMulticastOk = Not bmMulticastMissing
    edcCheckingMsg.Visible = False
End Function

Private Function mStationIncluded(ilPass As Integer, ilShttCode As Integer) As Boolean
    Dim llRow As Long
    Dim slStr As String
    Dim slInclExcl As String
    Dim ilIntCode As Integer
    Dim slName As String
    Dim ilPos As Integer
    Dim ilShtt As Integer
    
    mStationIncluded = False
    For llRow = grdTo.FixedRows To grdTo.Rows - 1 Step 1
        slStr = Trim$(grdTo.TextMatrix(llRow, TOCATEGORYINDEX))
        If slStr <> "" Then
            slInclExcl = UCase(Left$(grdTo.TextMatrix(llRow, TOINCLEXCLINDEX), 1))
            If ((ilPass = 0) And (slInclExcl = "I")) Or ((ilPass = 0) And (slInclExcl = "")) Or ((ilPass = 1) And (slInclExcl = "E")) Then
                Select Case UCase$(slStr)
                    Case "DMA MARKET"
                        ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                        For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                            If tgStations(ilShtt).iCode = ilShttCode Then
                                If tgStations(ilShtt).iMktCode = ilIntCode Then
                                    mStationIncluded = True
                                    Exit Function
                                End If
                            End If
                        Next ilShtt
                    Case "MSA MARKET"
                        ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                        For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                            If tgStations(ilShtt).iCode = ilShttCode Then
                                If tgStations(ilShtt).iMetCode = ilIntCode Then
                                    mStationIncluded = True
                                    Exit Function
                                End If
                            End If
                        Next ilShtt
                    Case "STATE"
                        slStr = grdTo.TextMatrix(llRow, TONAMEINDEX)
                        ilPos = InStr(1, slStr, "(", vbTextCompare)
                        If ilPos >= 1 Then
                            slName = Trim$(Left(slStr, ilPos - 1))
                        Else
                            slName = Trim$(grdTo.TextMatrix(grdTo.FixedRows, TONAMEINDEX))
                        End If
                        For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                            If tgStations(ilShtt).iCode = ilShttCode Then
                                If tgStations(ilShtt).sState = slName Then
                                    mStationIncluded = True
                                    Exit Function
                                End If
                            End If
                        Next ilShtt
                    Case "ZONE"
                        ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                        For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                            If tgStations(ilShtt).iCode = ilShttCode Then
                                If tgStations(ilShtt).iTztCode = ilIntCode Then
                                    mStationIncluded = True
                                    Exit Function
                                End If
                            End If
                        Next ilShtt
                    Case "FORMAT"
                        ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                        For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                            If tgStations(ilShtt).iCode = ilShttCode Then
                                If tgStations(ilShtt).iFmtCode = ilIntCode Then
                                    mStationIncluded = True
                                    Exit Function
                                End If
                            End If
                        Next ilShtt
                    Case "STATION"
                        ilIntCode = Val(grdTo.TextMatrix(llRow, TOCODEINDEX))
                        For ilShtt = 0 To UBound(tgStations) - 1 Step 1
                            If tgStations(ilShtt).iCode = ilShttCode Then
                                If tgStations(ilShtt).iCode = ilIntCode Then
                                    mStationIncluded = True
                                    Exit Function
                                End If
                            End If
                        Next ilShtt
                End Select
            End If
        End If
    Next llRow
End Function

Private Sub mCheckCode(ilSourceShttCode As Integer, llSourceMulticastGroupId As Long, ilPass As Integer, slInSQLQuery As String, ilMatchValue As Integer)
    Dim slSQLQuery As String
    Dim ilShttCode As Integer
    Dim ilShtt As Integer
    Dim slMsg As String
    Dim ilLoop As Integer
    Dim blFd As Boolean
    
    If llSourceMulticastGroupId <= 0 Then
        Exit Sub
    End If
    slSQLQuery = "Select Count(1) as attCount From ATT Where attShfCode = " & ilSourceShttCode
    slSQLQuery = slSQLQuery & " And attOffAir >= '" & Format(Now(), sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " And attDropDate >= '" & Format(Now(), sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " And attMulticast = 'Y'"
    Set att_rst = gSQLSelectCall(slSQLQuery)
    If att_rst.EOF Then
        Exit Sub
    End If
    If att_rst!attCount <= 0 Then
        Exit Sub
    End If
    Set shtt_rst = gSQLSelectCall(slInSQLQuery)
    Do While Not shtt_rst.EOF
        If shtt_rst!CheckCode <> ilMatchValue Then
            ilShttCode = shtt_rst!shttCode
            slSQLQuery = "Select attMulticast From ATT Where attShfCode = " & ilShttCode
            slSQLQuery = slSQLQuery & " And attOffAir >= '" & Format(Now(), sgSQLDateForm) & "'"
            slSQLQuery = slSQLQuery & " And attDropDate >= '" & Format(Now(), sgSQLDateForm) & "'"
            Set att_rst = gSQLSelectCall(slSQLQuery)
            Do While Not att_rst.EOF
                If att_rst!attMulticast = "Y" Then
                    'Check if selected
                    If Not mStationIncluded(ilPass, ilShttCode) Then
                        'Output message
                        ilShtt = gBinarySearchStation(ilShttCode)
                        If ilShtt <> -1 Then
                            slMsg = Trim$(tgStations(ilShtt).sCallLetters)
                        Else
                            slMsg = "Station Code " & ilShttCode
                        End If
                        If Not bmMulticastMissing Then
                            gLogMsgWODT "ON", hmUnMatch, sgDBPath & "Messages\" & "MulticastMissing.txt"
                        End If
                        bmMulticastMissing = True
                        blFd = False
                        For ilLoop = 0 To UBound(imMissingShttCode) - 1 Step 1
                            If ilShttCode = imMissingShttCode(ilLoop) Then
                                blFd = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not blFd Then
                            slMsg = "Multicast station " & slMsg & " not included in Region Definition"
                            gLogMsgWODT "W", hmUnMatch, slMsg
                            imMissingShttCode(UBound(imMissingShttCode)) = ilShttCode
                            ReDim Preserve imMissingShttCode(0 To UBound(imMissingShttCode) + 1) As Integer
                        End If
                        Exit Do
                    End If
                End If
                att_rst.MoveNext
            Loop
        End If
        shtt_rst.MoveNext
    Loop
End Sub

Private Sub mCheckName(ilSourceShttCode As Integer, llSourceMulticastGroupId As Long, ilPass As Integer, slInSQLQuery As String, slName As String)
    Dim slSQLQuery As String
    Dim ilShttCode As Integer
    Dim ilShtt As Integer
    Dim slMsg As String
    Dim ilLoop As Integer
    
    If llSourceMulticastGroupId <= 0 Then
        Exit Sub
    End If
    slSQLQuery = "Select Count(1) as attCount From ATT Where attShfCode = " & ilSourceShttCode
    slSQLQuery = slSQLQuery & " And attOffAir >= '" & Format(Now(), sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " And attDropDate >= '" & Format(Now(), sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " And attMulticast = 'Y'"
    Set att_rst = gSQLSelectCall(slSQLQuery)
    If att_rst.EOF Then
        Exit Sub
    End If
    If att_rst!attCount <= 0 Then
        Exit Sub
    End If
    Set shtt_rst = gSQLSelectCall(slInSQLQuery)
    Do While Not shtt_rst.EOF
        ilShttCode = shtt_rst!CheckCode
        For ilLoop = 0 To UBound(tgStations) - 1 Step 1
            If (tgStations(ilLoop).iCode = ilShttCode) And (tgStations(ilLoop).lMultiCastGroupID = llSourceMulticastGroupId) Then
                If Trim$(tgStations(ilLoop).sState) <> slName Then
                    slSQLQuery = "Select attMulticast From ATT Where attShfCode = " & ilShttCode
                    slSQLQuery = slSQLQuery & " And attOffAir >= '" & Format(Now(), sgSQLDateForm) & "'"
                    slSQLQuery = slSQLQuery & " And attDropDate >= '" & Format(Now(), sgSQLDateForm) & "'"
                    Set att_rst = gSQLSelectCall(slSQLQuery)
                    Do While Not att_rst.EOF
                        If att_rst!attMulticast = "Y" Then
                            'Check if selected
                            If Not mStationIncluded(ilPass, ilShttCode) Then
                                'Output message
                                ilShtt = gBinarySearchStation(ilShttCode)
                                If ilShtt <> -1 Then
                                    slMsg = Trim$(tgStations(ilShtt).sCallLetters)
                                Else
                                    slMsg = "Station Code " & ilShttCode
                                End If
                                If Not bmMulticastMissing Then
                                    gLogMsgWODT "ON", hmUnMatch, sgDBPath & "Messages\" & "MulticastMissing.txt"
                                End If
                                bmMulticastMissing = True
                                slMsg = "Multicast station " & slMsg & " not included in Region Definition"
                                gLogMsgWODT "W", hmUnMatch, slMsg
                                Exit Do
                            End If
                        End If
                        att_rst.MoveNext
                    Loop
                End If
            End If
        Next ilLoop
        shtt_rst.MoveNext
    Loop
End Sub

