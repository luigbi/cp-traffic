VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ExptReRate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7815
   ClientLeft      =   825
   ClientTop       =   2400
   ClientWidth     =   11550
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7815
   ScaleWidth      =   11550
   Begin VB.Frame frcProcessing 
      Caption         =   "Generating ReRate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   2640
      TabIndex        =   93
      Top             =   840
      Visible         =   0   'False
      Width           =   6615
      Begin ComctlLib.ProgressBar prgProcessing 
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblProcessing 
         Caption         =   "Preparing to Generate ReRate"
         Height          =   375
         Left            =   240
         TabIndex        =   95
         Top             =   480
         Width           =   6135
      End
   End
   Begin V81TrafficReports.CSI_ComboBoxList cbcDemo 
      Height          =   285
      Left            =   7800
      TabIndex        =   77
      Top             =   3585
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin VB.Frame frcMQ 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5445
      TabIndex        =   28
      Top             =   7605
      Visible         =   0   'False
      Width           =   2985
      Begin VB.TextBox edcYear 
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   2220
         TabIndex        =   12
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox edcStart 
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   855
         TabIndex        =   10
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lacTitle 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1605
         TabIndex        =   11
         Top             =   30
         Width           =   600
      End
      Begin VB.Label lacTitle 
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   9
         Top             =   30
         Width           =   750
      End
   End
   Begin V81TrafficReports.CSI_Calendar edcDate 
      Height          =   240
      Index           =   2
      Left            =   4365
      TabIndex        =   16
      Top             =   7590
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   423
      Text            =   "6/13/24"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin V81TrafficReports.CSI_Calendar edcDate 
      Height          =   240
      Index           =   1
      Left            =   4035
      TabIndex        =   14
      Top             =   7605
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   423
      Text            =   "6/13/24"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   -1  'True
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin V81TrafficReports.CSI_ComboBoxList cbcRevision 
      Height          =   165
      Left            =   10455
      TabIndex        =   19
      Top             =   60
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   291
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   0
   End
   Begin V81TrafficReports.CSI_Calendar edcDate 
      Height          =   240
      Index           =   0
      Left            =   8625
      TabIndex        =   8
      Top             =   45
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   423
      Text            =   "6/13/24"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   -1  'True
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCntr 
      Height          =   2310
      Left            =   4995
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   435
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   4075
      _Version        =   393216
      Rows            =   3
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.Frame frcResearchOptions 
      Caption         =   "Research Options"
      ForeColor       =   &H8000000D&
      Height          =   2000
      Left            =   120
      TabIndex        =   53
      Top             =   3060
      Width           =   11295
      Begin VB.OptionButton rbcReRateBookByLine 
         Height          =   195
         Left            =   9240
         TabIndex        =   89
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ckcHideBonus 
         Caption         =   "Hidden Bonus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6450
         TabIndex        =   87
         Top             =   810
         Width           =   1680
      End
      Begin VB.CheckBox ckcInvBonus 
         Caption         =   "Invoice  Bonus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4755
         TabIndex        =   86
         Top             =   810
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox ckcMG 
         Caption         =   "Spot MG/Outside"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2580
         TabIndex        =   85
         Top             =   810
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox ckcTreatMissedAsOrdered 
         Caption         =   "Treat Missed Spots as Ordered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5730
         TabIndex        =   80
         Top             =   1665
         Width           =   3015
      End
      Begin VB.CheckBox ckcTreatMGOsAsOrdered 
         Caption         =   "Treat MG/Outsides as Ordered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2580
         TabIndex        =   79
         Top             =   1665
         Width           =   2895
      End
      Begin VB.Frame frcIndex 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   1380
         Width           =   4395
         Begin VB.OptionButton rbcIndex 
            Caption         =   "Gimp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2460
            TabIndex        =   74
            Top             =   0
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton rbcIndex 
            Caption         =   "GRP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   3690
            TabIndex        =   73
            Top             =   0
            Width           =   780
         End
         Begin VB.Label lacIndex 
            Caption         =   "Compute Index By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   1560
         End
      End
      Begin VB.Frame frcDemo 
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
         Height          =   240
         Left            =   120
         TabIndex        =   65
         Top             =   525
         Width           =   7935
         Begin VB.OptionButton rbcDemo 
            Caption         =   "Primary"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2460
            TabIndex        =   70
            Top             =   0
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton rbcDemo 
            Caption         =   "2nd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3690
            TabIndex        =   69
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton rbcDemo 
            Caption         =   "3rd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   4635
            TabIndex        =   68
            Top             =   0
            Width           =   750
         End
         Begin VB.OptionButton rbcDemo 
            Caption         =   "4th"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   5490
            TabIndex        =   67
            Top             =   0
            Width           =   645
         End
         Begin VB.OptionButton rbcDemo 
            Caption         =   "Override:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   6330
            TabIndex        =   66
            Top             =   0
            Width           =   1140
         End
         Begin VB.Label lacDemo 
            Caption         =   "Purchase Contract Demo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   71
            Top             =   0
            Width           =   2070
         End
      End
      Begin VB.Frame frcBonus 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   120
         TabIndex        =   61
         Top             =   1095
         Width           =   7005
         Begin VB.OptionButton rbcBonus 
            Caption         =   "Vehicle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2460
            TabIndex        =   63
            Top             =   0
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton rbcBonus 
            Caption         =   "Vehicle, Daypart"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   3690
            TabIndex        =   62
            Top             =   0
            Width           =   2010
         End
         Begin VB.Label lacBonus 
            Caption         =   "Bonus by"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   64
            Top             =   -15
            Width           =   855
         End
      End
      Begin VB.CommandButton cmcSetBook 
         Caption         =   "Set Book by Line"
         Enabled         =   0   'False
         Height          =   255
         Left            =   9480
         TabIndex        =   60
         Top             =   210
         Width           =   1650
      End
      Begin VB.Frame frcReRateBook 
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
         Height          =   270
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   8805
         Begin VB.OptionButton rbcReRateBook 
            Caption         =   "Closest to Air Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   4320
            TabIndex        =   58
            Top             =   0
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton rbcReRateBook 
            Caption         =   "Vehicle Default"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   2460
            TabIndex        =   57
            Top             =   0
            Width           =   1740
         End
         Begin VB.OptionButton rbcReRateBook 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   7950
            TabIndex        =   56
            Top             =   0
            Width           =   840
         End
         Begin VB.OptionButton rbcReRateBook 
            Caption         =   "Contract Line"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   6330
            TabIndex        =   55
            Top             =   0
            Width           =   1545
         End
         Begin VB.Label lacTitle 
            Caption         =   "Research Book Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.Label lacShow 
         Caption         =   "Rows to Include"
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
         Left            =   120
         TabIndex        =   88
         Top             =   810
         Width           =   1395
      End
      Begin VB.Label lacResearchOptions 
         Caption         =   "Exception"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1665
         Width           =   2295
      End
   End
   Begin VB.Frame frcOutputOptions 
      Caption         =   "Output Options"
      ForeColor       =   &H8000000D&
      Height          =   2100
      Left            =   120
      TabIndex        =   35
      Top             =   5115
      Width           =   11295
      Begin VB.CheckBox ckcComment 
         Caption         =   "Comment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9480
         TabIndex        =   92
         Top             =   1095
         Width           =   1245
      End
      Begin VB.CheckBox ckcPriceType 
         Caption         =   "Price Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8160
         TabIndex        =   91
         Top             =   1095
         Width           =   1245
      End
      Begin VB.CheckBox ckcAudioType 
         Caption         =   "Audio Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5600
         TabIndex        =   90
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Frame frcLayout 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   270
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   10935
         Begin VB.OptionButton rbcLayout 
            Caption         =   "Merged with each research row"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   4635
            TabIndex        =   84
            Top             =   0
            Width           =   3135
         End
         Begin VB.OptionButton rbcLayout 
            Caption         =   "Separate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2460
            TabIndex        =   83
            Top             =   0
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Header Layout"
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
            Left            =   0
            TabIndex        =   82
            Top             =   0
            Width           =   1710
         End
      End
      Begin VB.CheckBox ckcCsv 
         Caption         =   "Export to CSV"
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
         Left            =   120
         TabIndex        =   52
         Top             =   1665
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox edcCSV 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2585
         TabIndex        =   51
         Top             =   1665
         Visible         =   0   'False
         Width           =   5715
      End
      Begin VB.Frame frcColumnLayout 
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
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1380
         Width           =   8790
         Begin VB.OptionButton rbcColumnLayout 
            Caption         =   "Separate Columns (PPP...RRR...)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2460
            TabIndex        =   49
            Top             =   0
            Value           =   -1  'True
            Width           =   3465
         End
         Begin VB.OptionButton rbcColumnLayout 
            Caption         =   "Pair Columns (PR PR PR)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   6300
            TabIndex        =   48
            Top             =   0
            Width           =   2625
         End
         Begin VB.Label lacColumnLayout 
            Caption         =   "Purchase(P) and ReRate(R)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   0
            TabIndex        =   50
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.Frame frcShow 
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
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   810
         Width           =   8805
         Begin VB.OptionButton rbcShow 
            Caption         =   "Hidden Only"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   4635
            TabIndex        =   46
            Top             =   0
            Width           =   1620
         End
         Begin VB.OptionButton rbcShow 
            Caption         =   "Package Only"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2460
            TabIndex        =   45
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton rbcShow 
            Caption         =   "Package and Hidden"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   6300
            TabIndex        =   44
            Top             =   0
            Width           =   2490
         End
      End
      Begin VB.CheckBox ckcSummary 
         Caption         =   "Summary Only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2580
         TabIndex        =   42
         Top             =   525
         Width           =   1590
      End
      Begin VB.CheckBox ckcIncludeInactive 
         Caption         =   "Inactive Lines"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4755
         TabIndex        =   41
         Top             =   525
         Width           =   1530
      End
      Begin VB.CheckBox ckcCost 
         Caption         =   "Cost"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2580
         TabIndex        =   39
         Top             =   1095
         Width           =   765
      End
      Begin VB.CheckBox ckcRating 
         Caption         =   "Rating"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3435
         TabIndex        =   38
         Top             =   1095
         Width           =   855
      End
      Begin VB.CheckBox ckcCPM 
         Caption         =   "CPM/CPP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4395
         TabIndex        =   37
         Top             =   1095
         Width           =   1125
      End
      Begin VB.CheckBox ckcACT1Lineup 
         Caption         =   "Lineup #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6930
         TabIndex        =   36
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lbcShow 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   76
         Top             =   525
         Width           =   2055
      End
      Begin VB.Label lacInclude 
         Caption         =   "Columns to Include"
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
         Left            =   120
         TabIndex        =   40
         Top             =   1095
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmcClosestBook 
      Appearance      =   0  'Flat
      Caption         =   "Closest Books"
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
      Left            =   6825
      TabIndex        =   34
      Top             =   7305
      Width           =   1650
   End
   Begin VB.Frame frcRevNo 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   360
      Left            =   6615
      TabIndex        =   30
      Top             =   2685
      Width           =   3780
      Begin VB.OptionButton rbcRevNo 
         Caption         =   "Latest"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   33
         Top             =   105
         Width           =   870
      End
      Begin VB.OptionButton rbcRevNo 
         Caption         =   "Original"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   32
         Top             =   90
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.Label lacRevNo 
         Caption         =   "Purchase Rev#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   105
         Width           =   1425
      End
   End
   Begin VB.ListBox lbcBookNames 
      Height          =   255
      ItemData        =   "ExportReRate.frx":0000
      Left            =   10695
      List            =   "ExportReRate.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   29
      Top             =   7605
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9720
      Top             =   7560
   End
   Begin VB.CommandButton cmcReturn 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
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
      Left            =   8820
      TabIndex        =   25
      Top             =   7305
      Width           =   2115
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "Press to Generate Spreadsheet and Transfer To Excel"
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
      Left            =   420
      TabIndex        =   23
      Top             =   7305
      Width           =   4725
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9360
      Top             =   7560
   End
   Begin VB.ListBox lbcCntrCode 
      Height          =   255
      ItemData        =   "ExportReRate.frx":0004
      Left            =   11040
      List            =   "ExportReRate.frx":0006
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   7560
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CheckBox ckcAllSpotLens 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10305
      TabIndex        =   22
      Top             =   2400
      Width           =   1515
   End
   Begin VB.CheckBox ckcAllCntr 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5010
      TabIndex        =   20
      Top             =   2805
      Width           =   1515
   End
   Begin VB.ListBox lbcSpotLens 
      Height          =   1815
      ItemData        =   "ExportReRate.frx":0008
      Left            =   10305
      List            =   "ExportReRate.frx":000A
      MultiSelect     =   2  'Extended
      TabIndex        =   21
      Top             =   435
      Width           =   1125
   End
   Begin VB.ListBox lbcAdvertiser 
      Height          =   2595
      ItemData        =   "ExportReRate.frx":000C
      Left            =   150
      List            =   "ExportReRate.frx":000E
      TabIndex        =   17
      Top             =   420
      Width           =   4740
   End
   Begin VB.Frame frcInstallMethod 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   6825
      Begin VB.OptionButton rbcDatesBy 
         Caption         =   "Range"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   5445
         TabIndex        =   6
         Top             =   0
         Width           =   990
      End
      Begin VB.OptionButton rbcDatesBy 
         Caption         =   "Contract"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4275
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton rbcDatesBy 
         Caption         =   "Quarter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3210
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton rbcDatesBy 
         Caption         =   "Week"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.OptionButton rbcDatesBy 
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lacTitle 
         Caption         =   "Dates by"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   -15
         Width           =   825
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
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
      Left            =   5445
      TabIndex        =   24
      Top             =   7305
      Width           =   1050
   End
   Begin VB.Label lacTitle 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   3510
      TabIndex        =   15
      Top             =   7605
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lacTitle 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   2685
      TabIndex        =   13
      Top             =   7575
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lacTitle 
      Caption         =   "Active on or After"
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
      Index           =   0
      Left            =   6960
      TabIndex        =   7
      Top             =   75
      Width           =   1665
   End
   Begin VB.Label lacMonth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12510
      TabIndex        =   27
      Top             =   -60
      Width           =   540
   End
End
Attribute VB_Name = "ExptReRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ExptReRate.frm on Fri 3/12/10 @ 11:00 AM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmToCSV As Integer   'From file hanle
Dim hmToExcel As Integer   'From file hanle
Dim smToCSV As String
Dim smToExcel As String
Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control

Dim lmEnableCol As Integer   'Current Column number
Dim lmEnableRow As Integer   'Current Row number

Dim imCbcDemoBottom As Integer

Dim smBypassCtrlNames(0 To 16) As String

Dim imChgMode As Integer
Dim tmChfAdvtExt() As CHFADVTEXT
Dim sm1or2PlaceRating As String
Dim imAdfCode As Integer

Dim imReRateDnfCode As Integer  'DnfCode for regular line spots, not used for MG/Outsides and Bonus spots
'rbcReRateBook (general book assignment):
'    Default Vehicle
'    Closest
'    Contract Line (Each line can have a diiferent setting)
'         ReRateLineBook
'             Default Default Vehicle
'             Closest
'             Purchased (only in sign on as csi)
'             Specified book
'    None
'
Dim lmReRatePop As Long             'For each contract, the population for ReRate computations will be obtained from the first schedule
                                    'spot and be used throught out contract.
                                    
Dim imMGDetailDnfCode As Integer '-1=Not assigned; -2=Mixture; >0 book assigned code
Dim imBonusDetailDnfCode As Integer '-1=Not assigned; -2=Mixture; >0 book assigned code
Dim imBonusTotalDnfCode As Integer '-1=Not assigned; -2=Mixture; >0 book assigned code
Dim imCntrLnDetailDnfCode As Integer '-1=Not assigned; -2=Mixture; >0 book assigned code

'Dim imReRatePopDnfCode As Integer
Dim lmWklyspots() As Long       'sched lines weekly # spots
Dim lmWklyAvgAud() As Long             'sched lines weekly avg aud
Dim lmWklyRates() As Long           'sched lines weekly rates
Dim lmWklyPopEst() As Long
Dim lmWklyMoDate() As Long

'Dim imMGVefCode() As Integer
'Dim imBonusVefCode() As Integer
Private Type MGBONUSINFO
    iVefCode As Integer
    iDnfCode As Integer
    iRdfCode As Integer
    iSpotLen As Integer 'TTP 10123
    sAudioTypes As String 'TTP 10144
End Type
Dim tmBonusInfo() As MGBONUSINFO
Dim tmMGInfo() As MGBONUSINFO

Private Type RATECARDSORT
    sKey As String * 6
    lDate As Long
    iRcfIndex As Integer
    iRcfCode As Integer
End Type
Dim tmRateCardSort() As RATECARDSORT

Dim lmCost() As Long
Dim imRtg() As Integer
Dim lmGrimp() As Long
Dim lmGRP() As Long

Dim lmAdvtOrderCost() As Long
Dim imAdvtOrderRtg() As Integer
Dim lmAdvtOrderGrimp() As Long
Dim lmAdvtOrderGRP() As Long
Dim lmAdvtOrderTotalSpots As Long
Dim lmAdvtOrderPop As Long
Dim imAdvtOrderDnfCode As Integer

Dim lmAdvtReRateCost() As Long
Dim imAdvtReRateRtg() As Integer
Dim lmAdvtReRateGrimp() As Long
Dim lmAdvtReRateGRP() As Long
Dim lmAdvtReRateTotalSpots As Long
Dim lmAdvtReRatePop As Long
Dim imAdvrReRateDnfCode As Integer


Dim lmAdvtBonusReRateCost() As Long
Dim imAdvtBonusReRateRtg() As Integer
Dim lmAdvtBonusReRateGrimp() As Long
Dim lmAdvtBonusReRateGRP() As Long
Dim lmAdvtBonusReRateTotalSpots As Long
Dim lmAdvtBonusReRatePop As Long
Dim imAdvtBonusReRateDnfCode As Integer


Dim lmAdvtPlusBonusReRateCost() As Long
Dim imAdvtPlusBonusReRateRtg() As Integer
Dim lmAdvtPlusBonusReRateGrimp() As Long
Dim lmAdvtPlusBonusReRateGRP() As Long
Dim lmAdvtPlusBonusReRateTotalSpots As Long
Dim lmAdvtPlusBonusReRatePop As Long
Dim imAdvtPlusBonusReRateDnfCode As Integer

Dim tmBookName() As SORTCODE
Dim smBookNameTag As String

Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT

Dim hmAdf As Integer 'Advertiser file handle
Dim tmAdf As ADF        'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer
Dim tmAgf As AGF
Dim imAgfRecLen As Integer
Dim hmPrf As Integer        'Prf Handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer      'Prf record length
Dim tmPrfSrchKey As LONGKEY0  'Prf key record image
Dim tmSrchKey As INTKEY0

Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF
Dim tmMnfList() As MNFLIST        'array of mnf codes for Missed reasons and billing rules

Dim hmCHF As Integer            'Contract header file handle
Dim tmChf As CHF
Dim imCHFRecLen As Integer
Dim tmChfSrchKey0 As LONGKEY0    'Key record image
Dim tmChfSrchKey1 As CHFKEY1
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClfP As CLF
Dim tmClfR As CLF
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmDrf As Integer        'Rating book file handle
Dim tmDrf As DRF            'DRF record image
Dim imDrfRecLen As Integer  'DRF record length
Dim hmDpf As Integer        'Demo Plus file handle
Dim tmDpf As DPF            'DPF record image
Dim imDpfRecLen As Integer  'DPF record length
'  Research Estimates
Dim hmDef As Integer
Dim hmRaf As Integer        'RAF file handle

Dim hmCbf As Integer            'Contract BR file handle
Dim imCbfRecLen As Integer      'CBF record length
Dim tmCbf As CBF

Dim tmChfPurchase As CHF            'CHF record image
Dim tmClfPurchase() As CLFLIST      'CLF record image
Dim tmCffPurchase() As CFFLIST      'CFF record image

Dim tmChfReRate As CHF            'CHF record image
Dim tmClfReRate() As CLFLIST      'CLF record image
Dim tmCffReRate() As CFFLIST      'CFF record image

Dim hmDnf As Integer            'Multiname file handle
Dim imDnfRecLen As Integer      'MNF record length
Dim tmDnfSrchKey0 As INTKEY0
Dim tmDnf As DNF
Dim bmBookByLine As Boolean

Private Type RERATEINFO              'array of vehicles or entire contract's spots per week , quarter at a time
    lSeqNo As Long
    lChfCode(0 To 1) As Long
    lClfCode(0 To 1) As Long
    iVefCode As Integer
    iRdfCode As Integer
    sType As String * 1           'S = std, O = order, a = air, h = hidden
    sSubType As String * 1      'M=MG; B=Bonus; Blank for all other
    sProduct As String * 35
    lCntrNo As Long
    iLineNo As Integer
    iPkLineNo As Integer         'associated package line # reference (if stype = H)
    sAudioType As String * 64    'L=Live, P=PreRec, R=Recorded (Or any combination like LPR)
    iLen As Integer
    sACT1LineupCode As String * 11
    '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
    sACT1StoredTime As String * 1
    sACT1StoredSpots As String * 1
    sACT1StoreClearPct As String * 1
    sACT1DaypartFilter As String * 1
    sPriceType As String * 1
    sLineComment As String * 1000
    iDnfCode(0 To 1) As Integer
    lPop(0 To 1) As Long
    lSatelliteEst(0 To 1) As Long       '6-1-04
    'Index:0=Ordered; 1=Aired; 2=MG's; 3=Bonus
    lTotalSpots(0 To 1) As Long               'total spots per this vehicles qtr
    'lTotalCost(0 To 1) As Long
    dTotalCost(0 To 1) As Double 'TTP 10439 - Rerate 21,000,000
    iTotalAvgRating(0 To 1) As Integer
    lTotalAvgAud(0 To 1) As Long
    lTotalGrimps(0 To 1) As Long
    lTotalGRP(0 To 1) As Long
    lTotalCPP(0 To 1) As Long
    lTotalCPM(0 To 1) As Long
    sCBS(0 To 1) As String * 1
End Type
Dim tmReRate() As RERATEINFO

Private Type FORMULAINFO
    sRowType As String * 2     'SL=Regular Line; PL=Package; HL=Hidden Line; BL=Bonus; CT (Contract Total);CB (Contract Bonus);CS (Contract Total Plus Bonus);AT;AB;AS
    iExcelRow As Integer
    lPop As Long
    'lExtTotal As Long
    dExtTotal As Double 'TTP 10439 - Rerate 21,000,000
End Type
Dim tmFormulaInfo() As FORMULAINFO
Dim imCurrentFormulaIndex As Integer
Dim imByCntrLnDnfCode As Integer

Dim omBook As Object
Dim omSheet As Object
Dim imExcelRow As Integer
Dim imCostColumn(0 To 2) As Integer
Dim imCPMCPPColumn(0 To 3) As Integer
Dim imRatingColumn(0 To 1) As Integer
Dim imAQHColumn(0 To 1) As Integer
Dim imRightAlignColumn(0 To 19) As Integer
Dim imReRateColumn As Integer
'Dim imACT1LineupCodeColumn As Integer
'Dim imAudioTypeColumn As Integer
Dim imPurchasedColumn As Integer
Dim imAutoFitSkipColumn(0 To 3) As Integer
Dim imAutoFitSummaryColumn(0 To 3) As Integer
Dim imCtrlKey As Integer
Dim bmInGrid As Boolean
Dim bmAllClicked As Boolean
Dim bmSetAll As Boolean
Dim lmScrollTop As Long
Dim lmLastClickedRow As Long
Dim imMnfDemo As Integer
Dim imNumberDecPlaces As Integer
Dim imAdjDecPlaces As Integer

Dim bmInSummaryMode As Boolean
Dim smSummaryRecords() As String
Dim lmContractPopulation As Long
Dim imNoCntr As Integer

'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates for MG/OS
Dim hlSmf As Integer        'Smf handle
Dim ilSmfRecLen As Integer     'Record length
Dim tlSmf As SMF
Dim tlSmfSrchKey As SMFKEY0
Dim tlSmfSrchKey2 As LONGKEY0

Dim smColumnLetter(0 To 50) As String   'Left most column number is one

Private Type RERATEBOOKDNFCODES
    lChfCode As Long
    sType As String * 1 'S=Standard;B=Bonus;M=MG
    iLineNo As Integer  'Type = S
    iVefCode As Integer 'Type = B and M
    iDrfCode As Integer 'Type = B
    iDnfCode As Integer
End Type
Dim tmReRateBookDnfCodes() As RERATEBOOKDNFCODES

'Const CNTRNOINDEX = 0
'Const PRODUCTINDEX = 1
'Const VEHICLENAMEINDEX = 2
'Const DAYPARTINDEX = 3
'Const LENGTHINDEX = 4
'Const DEMOINDEX = 5
'Const CNTRBOOKINDEX = 6
'Const RERATEBOOKINDEX = 7
'Const SELECTEDINDEX = 8
'Const PURCHASECHFCODEINDEX = 9

Const GENINDEX = 0
Const PRODUCTINDEX = 1
Const CNTRNOINDEX = 2
Const VERSIONINDEX = 3
Const SELECTEDINDEX = 4
Const PURCHASECHFCODEINDEX = 5
Const STARTREVNOINDEX = 6
Const RERATECHFCODEINDEX = 7
Const ENDREVNOINDEX = 8
Const NODEMOSINDEX = 9

'Color &HBBGGRR
' BackColor, ForeColor, FillColor (standard RGB colors: form, controls)
Const PURCHASECOLOR = &HCCF2FF
Const RERATECOLOR = &H99E6FF

'Excel Column numbers
Dim LINEEXCEL As Integer
Dim VEHICLEEXCEL As Integer
Dim DAYPARTEXCEL As Integer
Dim AUDIOTYPEEXCEL As Integer
Dim LINEUPEXCEL As Integer
Dim LENEXCEL As Integer
Dim PRICETYPEEXCEL As Integer
Dim RATEEXCEL As Integer
Dim LINECOMMENTEXCEL As Integer
Dim LASTSTATICCOLEXCEL As Integer

Dim PEXTTOTALEXCEL As Integer
Dim PUNITEXCEL As Integer
Dim PAQHEXCEL As Integer
Dim PRTGEXCEL As Integer
Dim PCPMEXCEL As Integer
Dim PCPPEXCEL As Integer
Dim PGIMPEXCEL As Integer
Dim PGRPEXCEL As Integer
Dim PBOOKEXCEL As Integer

Dim REXTTOTALEXCEL As Integer
Dim RUNITEXCEL As Integer
Dim RAQHEXCEL As Integer
Dim RRTGEXCEL As Integer
Dim RCPMEXCEL As Integer
Dim RCPPEXCEL As Integer
Dim RGIMPEXCEL As Integer
Dim RGRPEXCEL As Integer
Dim RBOOKEXCEL As Integer
Dim INDEXEXCEL As Integer
Dim imSortColumn As Integer
Dim imSortDir As Integer
Dim imMGCheckState As Integer
Dim imRatingCheckState As Integer
Dim imAudioTypeCheckState As Integer
Dim imPriceTypeCheckState As Integer
Dim imCommentCheckState As Integer
Dim imACT1LineupCheckState As Integer
Dim tmReRateHeader As ReRateHeaderInfo
Dim bmExportedHeader As Boolean
Dim imReRateLastBookMode As Integer 'TTP 10140, show When [Set Book by Line] method is enabled - Keep track of Last "Research Book Name" option (rbcReRateBook)

'TTP 10193 - Add Line Comment
Dim hmCxf As Integer            'Comments file handle
Dim tmCxf As CXF               'CXF record image
Dim tmCxfSrchKey As LONGKEY0     'CXF key record image
Dim imCxfRecLen As Integer         'CXF record length

Dim smDelimiter As String 'TTP 10309: ReRate: a comma in the agency name causes data to not line up with column headers when running the report with the "merged" option; use a upper ASCII delimeter


Private Sub cbcDemo_GotFocus()
    mSetShow
End Sub

Private Sub cbcDemo_OnChange()
    mSetCommands
End Sub

Private Sub cbcDemo_ReSetLoc()
    cbcDemo.Top = imCbcDemoBottom - cbcDemo.Height
End Sub

Private Sub cbcRevision_LostFocus()
    mSetShow
End Sub

Private Sub ckcACT1Lineup_Click()
    mSetCommands
End Sub

Private Sub ckcACT1Lineup_GotFocus()
    mSetShow
End Sub

Private Sub ckcAllCntr_Click()
    Dim slSelect As String
'    Dim blSelect As Boolean
    Dim llRow As Long
'    Dim llLastRow As Long
'
    If bmInGrid Then
        Exit Sub
    End If
    If ckcAllCntr.Value = vbChecked Then
        slSelect = "1"
'        blSelect = vbTrue
    Else
        slSelect = ""
'        blSelect = vbFalse
    End If
'    For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
'        If grdCntr.TextMatrix(llRow, PRODUCTINDEX) <> "" Then
'            llLastRow = llRow
'            grdCntr.TextMatrix(llRow, SELECTEDINDEX) = slSelect
'        End If
'    Next llRow
'    If blSelect Then
'        grdCntr.Col = 0
'        grdCntr.Row = grdCntr.FixedRows
'        grdCntr.RowSel = llLastRow
'        grdCntr.ColSel = grdCntr.Cols - 1
'    Else
'        grdCntr.Col = SELECTEDINDEX
'        grdCntr.Row = 0
'    End If
    For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        'If grdCntr.TextMatrix(llRow, PRODUCTINDEX) <> "" Then
        If grdCntr.TextMatrix(llRow, CNTRNOINDEX) <> "" Then
            grdCntr.TextMatrix(llRow, SELECTEDINDEX) = slSelect
            mPaintRowColor llRow
        End If
    Next llRow
    mSetDemo
    grdCntr.Row = 0
    grdCntr.Col = SELECTEDINDEX
    mSetCommands
End Sub

Private Sub ckcAllCntr_GotFocus()
    mSetShow
End Sub

Private Sub ckcAllSpotLens_Click()
    Dim Value As Integer
    Value = False
    If ckcAllSpotLens.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If bmSetAll Then
        bmAllClicked = True
        llRg = CLng(lbcSpotLens.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSpotLens.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        bmAllClicked = False
    End If
    mSetCommands

End Sub

Private Sub ckcAllSpotLens_GotFocus()
    mSetShow
End Sub

Private Sub ckcCost_Click()
    mSetCommands
End Sub

Private Sub ckcCost_GotFocus()
    mSetShow
End Sub

Private Sub ckcCPM_Click()
    mSetCommands
End Sub

Private Sub ckcCPM_GotFocus()
    mSetShow
End Sub

Private Sub ckcCsv_Click()
    'TTP 10258: ReRate - make it work without requiring Office
    'JW 10/14/21 - Fix v81 TTP 10258 issues per Jason call 10/14/21 - no PRPRPR mode for CSV
    If ckcCsv.Value = 0 Then
        cmcExport.Caption = "Press to Generate Spreadsheet and Transfer To Excel"
        rbcColumnLayout(1).Enabled = True
    Else
        cmcExport.Caption = "Press to Generate and Export ReRate CSV file"
        rbcColumnLayout(1).Enabled = False
        rbcColumnLayout(0).Value = True
    End If
End Sub

Private Sub ckcCsv_GotFocus()
    mSetShow
End Sub

Private Sub ckcHideBonus_Click()
    If ckcInvBonus.Value = vbUnchecked And ckcHideBonus.Value = vbUnchecked Then
        rbcBonus(0).Enabled = False
        rbcBonus(1).Enabled = False
    Else
        rbcBonus(0).Enabled = True
        rbcBonus(1).Enabled = True
    End If
    
    mSetCommands
End Sub

Private Sub ckcHideBonus_GotFocus()
    mSetShow
End Sub

Private Sub ckcInvBonus_Click()
    If ckcInvBonus.Value = vbUnchecked And ckcHideBonus.Value = vbUnchecked Then
        rbcBonus(0).Enabled = False
        rbcBonus(1).Enabled = False
    Else
        rbcBonus(0).Enabled = True
        rbcBonus(1).Enabled = True
    End If
    
    mSetCommands
End Sub

Private Sub ckcInvBonus_GotFocus()
    mSetShow
End Sub

Private Sub ckcMG_Click()
    mSetCommands
End Sub

Private Sub ckcMG_GotFocus()
    mSetShow
End Sub

Private Sub ckcRating_Click()
    mSetCommands
End Sub

Private Sub ckcRating_GotFocus()
    mSetShow
End Sub

Private Sub ckcSummary_Click()
    If ckcSummary.Value = vbChecked Then
        frcShow.Enabled = False
        ckcRating.Enabled = False
        ckcRating.Value = vbUnchecked
        ckcAudioType.Enabled = False
        ckcAudioType.Value = vbUnchecked
        ckcPriceType.Enabled = False
        ckcPriceType.Value = vbUnchecked
        ckcComment.Enabled = False
        ckcComment.Value = vbUnchecked
        ckcACT1Lineup.Enabled = False
        ckcACT1Lineup.Value = vbUnchecked
    Else
        frcShow.Enabled = True
        ckcRating.Enabled = True
        ckcRating.Value = imRatingCheckState
        ckcAudioType.Enabled = True
        ckcAudioType.Value = imAudioTypeCheckState
        ckcPriceType.Enabled = True
        ckcPriceType.Value = imPriceTypeCheckState
        ckcComment.Enabled = True
        ckcComment.Value = imCommentCheckState
        If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then
            ckcACT1Lineup.Enabled = True
            ckcACT1Lineup.Value = imACT1LineupCheckState
        Else
            ckcACT1Lineup.Enabled = False
            ckcACT1Lineup.Value = vbUnchecked
        End If
    End If
    mSetCommands
    mGetCSVFilename
End Sub

Private Sub ckcSummary_GotFocus()
    mSetShow
End Sub

Private Sub ckcTreatMGOsAsOrdered_Click()
    'TTP 9922 - ReRate - treat makegoodsoutsides as aired
    If ckcTreatMGOsAsOrdered.Value = vbChecked Then
        'if "Treat MG/Outside as aired" is checked on, then the Rows to Include "Spot MG/Outside" checkbox should probably be unchecked and grayed out.
        ckcMG.Value = vbUnchecked
        ckcMG.Enabled = False
    Else
        'if "Treat MG/Out" is unchecked, then "Rows to include spot MG/Outside" should be available to be checked.
        ckcMG.Enabled = True
        ckcMG.Value = imMGCheckState
    End If
    mSetCommands
End Sub

Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate False
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcClosestBook_Click()
    ClosestBooks.Show vbModal
    mPopExcludeBooks
End Sub

Private Sub cmcExport_Click()
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim ilCol As Integer
    Dim slLogoExtension As String
    Dim illoop As Integer
    Dim blSkipColumn As Boolean
    Dim slWidth As String
    '12/1/2020 - TTP 9765 Object Required error message when driving crazy
    cmcExport.Enabled = False
    bmExportedHeader = False 'TTP 10082 - merge header into columns
'    If imExporting Then
'        Exit Sub
'    End If
    ReDim smSummaryRecords(0 To 0) As String
    ReDim tmFormulaInfo(0 To 0) As FORMULAINFO
    imCurrentFormulaIndex = -1
    ilRet = 0
    'smToExcel = gFileNameFilterNotPath(Trim$(edcExcel.Text))
    'If gFileExist(smToExcel) = 0 Then
    '    Kill smToExcel
    'End If
    'If (InStr(smToExcel, ":") = 0) And (Left$(smToExcel, 2) <> "\\") Then
    '    smToExcel = sgExportPath & smToExcel
    'End If
    
    'TTP 10258: ReRate - make it work without requiring Office
    If ckcCsv.Value = vbChecked Then
        ilRet = 0
        smToCSV = gFileNameFilterNotPath(Trim$(edcCSV.Text))
        If gFileExist(smToCSV) = 0 Then
            err = 0
            On Error Resume Next
            Kill smToCSV
            If err <> 0 Then
                If err = 70 Then
                    MsgBox "Error Creating " & smToCSV & ", Error #" & str$(err) & " - " & Error(err) & vbCrLf & "File may be open or in use by another user.", vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                Else
                    MsgBox "Error Creating " & smToCSV & ", Error #" & str$(err) & " - " & Error(err), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                End If
                cmcExport.Enabled = True
                Exit Sub
            End If
            On Error GoTo 0 'Clear Error Handler
        End If
        If (InStr(smToCSV, ":") = 0) And (Left$(smToCSV, 2) <> "\\") Then
            smToCSV = sgExportPath & smToCSV
        End If
    End If

    'If ilRet = 0 Then
    '    ilRet = gFileOpen(smToExcel, "Append", hmToExcel)
    '    If ilRet <> 0 Then
    '        MsgBox "Open " & smToExcel & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
    '        Exit Sub
    '    End If
    ' Else
    '    ilRet = 0
    '    ilRet = gFileOpen(smToExcel, "Output", hmToExcel)
    '    If ilRet <> 0 Then
    '        MsgBox "Open " & smToExcel & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
    '        Exit Sub
    '    End If
    ' End If


    ilRet = 0
    ilRet = gFileExist(smToCSV)
    If ilRet = 0 Then
        ilRet = gFileOpen(smToCSV, "Append", hmToCSV)
        If ilRet <> 0 Then
            MsgBox "Open " & smToCSV & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
             Exit Sub
         End If
     Else
        ilRet = 0
        ilRet = gFileOpen(smToCSV, "Output", hmToCSV)
        If ilRet <> 0 Then
            MsgBox "Open " & smToCSV & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            Exit Sub
        End If
     End If

    Screen.MousePointer = vbHourglass
    imExporting = True
    imExcelRow = 3  '1
    
    mSetExcelColumns
    mDefineAlignColumns

    imCostColumn(0) = RATEEXCEL
    imCostColumn(1) = PEXTTOTALEXCEL
    imCostColumn(2) = REXTTOTALEXCEL
    imCPMCPPColumn(0) = PCPMEXCEL
    imCPMCPPColumn(1) = PCPPEXCEL
    imCPMCPPColumn(2) = RCPMEXCEL
    imCPMCPPColumn(3) = RCPPEXCEL
    imRatingColumn(0) = PRTGEXCEL
    imRatingColumn(1) = RRTGEXCEL
    imAQHColumn(0) = PAQHEXCEL
    imAQHColumn(1) = RAQHEXCEL
    imReRateColumn = REXTTOTALEXCEL '17
    imPurchasedColumn = PEXTTOTALEXCEL  '8
    'imACT1LineupCodeColumn = LINEUPEXCEL
    'imAudioTypeColumn = AUDIOTYPEEXCEL
    
    imAutoFitSkipColumn(0) = 1
    imAutoFitSkipColumn(3) = LINEUPEXCEL
    'Don't autofit column with title Purchase and ReRate
    'Ignored for Paired columns
    If ckcCost.Value = vbChecked Then
        'Title above extended total
        imAutoFitSkipColumn(1) = PEXTTOTALEXCEL '8
        imAutoFitSkipColumn(2) = REXTTOTALEXCEL '17
    Else
        'Title above units
        imAutoFitSkipColumn(1) = PEXTTOTALEXCEL + 1 '9
        imAutoFitSkipColumn(2) = REXTTOTALEXCEL + 1 '18
    End If
    imAutoFitSummaryColumn(0) = 2   'total
    imAutoFitSummaryColumn(1) = 4   'Product
    imAutoFitSummaryColumn(2) = PBOOKEXCEL  '16   'Book name
    imAutoFitSummaryColumn(3) = RBOOKEXCEL  '25   'Book name
    If rbcColumnLayout(1).Value Then
        'imCostColumn(0) = 7
        'imCostColumn(1) = 8
        'imCostColumn(2) = 9
        'imCPMCPPColumn(0) = 16
        'imCPMCPPColumn(1) = 17
        'imCPMCPPColumn(2) = 18
        'imCPMCPPColumn(3) = 19
        'imRatingColumn(0) = 14
        'imRatingColumn(1) = 15
        imReRateColumn = REXTTOTALEXCEL + 10 '27
        imAutoFitSkipColumn(1) = 1
        imAutoFitSkipColumn(2) = 1
        'imAQHColumn(0) = 12
        'imAQHColumn(1) = 13
        'imAutoFitSummaryColumn(0) = 2   'total
        'imAutoFitSummaryColumn(1) = 4   'Product
        'imAutoFitSummaryColumn(2) = 24   'Book name
        'imAutoFitSummaryColumn(3) = 25   'Book name
    'ElseIf rbcForm(2).Value = True Then
    '    imCostColumn(0) = 7
    '    imCostColumn(1) = 9
    '    imCostColumn(2) = 9
    '    imCPMCPPColumn(0) = 13
    '    imCPMCPPColumn(1) = 14
    '    imCPMCPPColumn(2) = 13
    '    imCPMCPPColumn(3) = 14
    '    imRatingColumn(0) = 12
    '    imRatingColumn(1) = 12
    '    imReRateColumn = 10
    '    imAutoFitSkipColumn(2) = imAutoFitSkipColumn(1)
    End If
    
    'TTP 10258: ReRate - make it work without requiring Office - if not using OFFICE - dont check if Open workbook function works
    If ExptReRate.ckcCsv.Value = vbUnchecked Then
        ilRet = gExcelOutputGeneration("O", omBook, omSheet, 1)
    Else
        ilRet = True
    End If
    '12/1/2020 - TTP 9765 Object Required error message when driving crazy
    If ilRet = False Then
        MsgBox "Unable to Generate ReRate, Excel in use." & vbCrLf & "Please Close Excel and try again", vbCritical + vbOKOnly, "Generate ReRate Error"
        ilRet = 429 'Component can't create object or return reference to this object
        GoTo mErrorSkipHere
    End If
    'Add picture
    'ActiveSheet.Shapes.AddPicture("full path of your file with extension", linktofile:=msoFalse, savewithdocument:=msoCTrue, Left:=50, Top:=50, Width:=-1; Height:=-1)
    'msoFalse = 0; msoTrue = -1;
    slLogoExtension = ""
    If gFileExist(sgLogoPath & "RptLogo.jpg") = FILEEXISTS Then
        slLogoExtension = ".jpg"
    ElseIf gFileExist(sgLogoPath & "RptLogo.gif") = FILEEXISTS Then
        slLogoExtension = ".gif"
    ElseIf gFileExist(sgLogoPath & "RptLogo.Bmp") = FILEEXISTS Then
        slLogoExtension = ".bmp"
    End If
    'TTP 10604 - ReRate Report Error When Excel Not Installed and Exporting to CSV
    If ExptReRate.ckcCsv.Value = vbUnchecked Then
        If slLogoExtension <> "" Then
            omSheet.Shapes.AddPicture sgLogoPath & "RptLogo" & slLogoExtension, 0, -1, omSheet.Range("A1").Left, omSheet.Range("A1").Top, -1, -1
        'Else
        '    If gFileExist(sgLogoPath & "CustLogo.jpg") = FILEEXISTS Then
        '        slLogoExtension = ".jpg"
        '    ElseIf gFileExist(sgLogoPath & "CustLogo.gif") = FILEEXISTS Then
        '        slLogoExtension = ".gif"
        '    ElseIf gFileExist(sgLogoPath & "CustLogo.Bmp") = FILEEXISTS Then
        '        slLogoExtension = ".bmp"
        '    End If
        '    If slLogoExtension <> "" Then
        '        omSheet.Shapes.AddPicture sgLogoPath & "CustLogo" & slLogoExtension, 0, -1, omSheet.Range("A1").Left, omSheet.Range("A1").Top, -1, -1
        '    End If
        End If
    End If
    
    If ExptReRate.ckcCsv.Value = vbUnchecked Then mSetCellRule
    
    'Print #hmToCSV, Trim$(tgSpf.sGClient) & " ReRate"
    If slLogoExtension <> "" Then
        imExcelRow = 3  '1
        mPrint "ReRate" & IIF(ckcSummary.Value = vbChecked, " Summary", "")
    Else
        imExcelRow = 1
        mPrint Trim$(tgSpf.sGClient) & " ReRate" & IIF(ckcSummary.Value = vbChecked, " Summary", ""), smDelimiter
    End If
    If ExptReRate.ckcCsv.Value = vbUnchecked Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1)

    mExport
            
    If ExptReRate.ckcCsv.Value = vbUnchecked Then
        If imNumberDecPlaces = 1 Then
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.0", -1, PAQHEXCEL)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.0", -1, PGIMPEXCEL)
        ElseIf imNumberDecPlaces = 2 Then
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.00", -1, PAQHEXCEL)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.00", -1, PGIMPEXCEL)
        ElseIf imNumberDecPlaces = 3 Then
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.000", -1, PAQHEXCEL)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.000", -1, PGIMPEXCEL)
        Else
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0", -1, PAQHEXCEL)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0", -1, PGIMPEXCEL)
        End If
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.00", -1, PCPMEXCEL)
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0", -1, PCPPEXCEL)
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.0", -1, PGRPEXCEL)
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.0", -1, PRTGEXCEL)
        If imNumberDecPlaces = 1 Then
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.0", -1, RAQHEXCEL)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.0", -1, RGIMPEXCEL)
        ElseIf imNumberDecPlaces = 2 Then
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.00", -1, RAQHEXCEL)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.00", -1, RGIMPEXCEL)
        ElseIf imNumberDecPlaces = 3 Then
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.000", -1, RAQHEXCEL)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.000", -1, RGIMPEXCEL)
        Else
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0", -1, RAQHEXCEL)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0", -1, RGIMPEXCEL)
        End If
        
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.00", -1, RCPMEXCEL)
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0", -1, RCPPEXCEL)
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.0", -1, RGRPEXCEL)
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.0", -1, RRTGEXCEL)
        ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "#0.00", -1, INDEXEXCEL)
        
        'Set alignment
        For ilCol = 0 To UBound(imRightAlignColumn) Step 1
            If ((ckcSummary.Value = vbChecked) And (imRightAlignColumn(ilCol) <> LINEEXCEL)) Or (ckcSummary.Value = vbUnchecked) Then
                ilRet = gExcelOutputGeneration("HA", omBook, omSheet, , str(xlRight), -1, imRightAlignColumn(ilCol))
            End If
        Next ilCol
    End If
    
    If ckcSummary.Value = vbUnchecked Then
        If ExptReRate.ckcCsv.Value = vbUnchecked Then
            'Auto Fit (Not in Summary mode)
            For ilCol = 1 To omSheet.UsedRange.Columns.Count Step 1
                blSkipColumn = False
                For illoop = 0 To UBound(imAutoFitSkipColumn) Step 1
                    If imAutoFitSkipColumn(illoop) = ilCol Then
                        blSkipColumn = True
                        Exit For
                    End If
                Next illoop
                If Not blSkipColumn Then
                    'slAction, , olSheet, , , ,ilColumn
                    ilRet = gExcelOutputGeneration("AF", omBook, omSheet, , , , ilCol)
                End If
                Select Case ilCol
                    Case AUDIOTYPEEXCEL
                        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 11.45, , ilCol)
                    Case PRICETYPEEXCEL
                        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 10.43, , ilCol)
                    Case PEXTTOTALEXCEL
                        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 9.45, , ilCol)
                    Case REXTTOTALEXCEL
                        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 9.45, , ilCol)
                    Case LINECOMMENTEXCEL
                        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 14.43, , ilCol)
                End Select
            Next ilCol
            
            'Column with (column 1)
            If rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , "2", , 1)
            
            'TTP 10082 - merge header into columns
            If rbcLayout(1).Value = True Then
                ilRet = gExcelOutputGeneration("AF", omBook, omSheet, , , , LINEEXCEL)
                ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , 11) '11 is a Blank , just before Line#
            End If
        End If
    Else
        mOutputSummary
        'For ilLoop = 0 To UBound(imAutoFitSummaryColumn) Step 1
        '    ilRet = gExcelOutputGeneration("AF", omBook, omSheet, , , , imAutoFitSummaryColumn(ilLoop))
        'Next ilLoop
        If ExptReRate.ckcCsv.Value = vbUnchecked Then
            For ilCol = 1 To omSheet.UsedRange.Columns.Count Step 1
                blSkipColumn = False
                For illoop = 0 To UBound(imAutoFitSkipColumn) Step 1
                    If imAutoFitSkipColumn(illoop) = ilCol Then
                        blSkipColumn = True
                        Exit For
                    End If
                Next illoop
                If Not blSkipColumn Then
                    'slAction, , olSheet, , , ,ilColumn
                    ilRet = gExcelOutputGeneration("AF", omBook, omSheet, , , , ilCol)
                End If
                Select Case ilCol
                    Case PEXTTOTALEXCEL
                        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 9.45, , ilCol)
                    Case REXTTOTALEXCEL
                        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 9.45, , ilCol)
                End Select
            Next ilCol
            'Column width (Column 1)
            ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , "2", , 1)
        End If
    End If
    
    If ExptReRate.ckcCsv.Value = vbUnchecked Then mSendFormulaToExcel
    
    If ExptReRate.ckcCsv.Value = vbUnchecked Then
        'Hide columns must be after the auto fit
        If ckcCost.Value = vbUnchecked Then 'Columns to Include: Cost
            For ilCol = LBound(imCostColumn) To UBound(imCostColumn) Step 1
                ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , imCostColumn(ilCol))
            Next ilCol
        End If
        If ckcCPM.Value = vbUnchecked Then 'Columns to Include: Include CPM
            For ilCol = LBound(imCPMCPPColumn) To UBound(imCPMCPPColumn) Step 1
                ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , imCPMCPPColumn(ilCol))
            Next ilCol
        End If
        If ckcRating.Value = vbUnchecked Then 'Columns to Include: Include Rating
            For ilCol = LBound(imRatingColumn) To UBound(imRatingColumn) Step 1
                ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , imRatingColumn(ilCol))
            Next ilCol
        End If
        If ckcACT1Lineup.Value = vbUnchecked Then 'Columns to Include: Include Act1 Line #
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , LINEUPEXCEL)
        End If
        'TTP 10156: make audio type optional
        If ckcAudioType.Value = vbUnchecked Then 'Columns to Include: Include Audio Type
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , AUDIOTYPEEXCEL)
        End If
        'TTP 10192 - Price Type
        If ckcPriceType.Value = vbUnchecked Then 'Columns to Include: Include Price Type
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , PRICETYPEEXCEL)
        End If
        'TTP 10193 - Line Comment
        If ckcComment.Value = vbUnchecked Then 'Columns to Include: Include Line Comment
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , LINECOMMENTEXCEL)
        End If
        
        If ckcSummary.Value = vbChecked Then 'Summary Only
            ilRet = gExcelOutputGeneration("AF", omBook, omSheet, , , , LINEEXCEL)    'LINE #
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , LINEUPEXCEL)    'Act1 Line #
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , AUDIOTYPEEXCEL) 'Audio Type
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , LENEXCEL)       'Length
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , PRICETYPEEXCEL) 'Price Type
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , RATEEXCEL)      'Rate
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , LINECOMMENTEXCEL)   'Line Comment
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , PBOOKEXCEL)      'Purch Book
            ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , RBOOKEXCEL)      'ReRate Book
            
            If rbcReRateBook(2).Value = False Then
                For ilCol = LBound(imAQHColumn) To UBound(imAQHColumn) Step 1
                    ilRet = gExcelOutputGeneration("H", omBook, omSheet, , "True", , imAQHColumn(ilCol))
                Next ilCol
            End If
        End If
        If rbcReRateBook(2).Value Then 'Research Book Name by: None
            'Since columns have no values, the AutoFit reduces the width.  Set so that ReRate width matched Purchase width
            slWidth = str(omSheet.Columns(PAQHEXCEL).ColumnWidth)
            ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , slWidth, , RAQHEXCEL)
            slWidth = str(omSheet.Columns(PRTGEXCEL).ColumnWidth)
            ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , slWidth, , RRTGEXCEL)
            slWidth = str(omSheet.Columns(PCPMEXCEL).ColumnWidth)
            ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , slWidth, , RCPMEXCEL)
            slWidth = str(omSheet.Columns(PCPPEXCEL).ColumnWidth)
            ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , slWidth, , RCPPEXCEL)
            slWidth = str(omSheet.Columns(PGIMPEXCEL).ColumnWidth)
            ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , slWidth, , RGIMPEXCEL)
            slWidth = str(omSheet.Columns(PGRPEXCEL).ColumnWidth)
            ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , slWidth, , RGRPEXCEL)
        End If
        'Font Size
        For ilCol = 1 To omSheet.UsedRange.Columns.Count Step 1
            'slAction, , olSheet, , , ,ilColumn
            ilRet = gExcelOutputGeneration("FS", omBook, omSheet, , "12", , ilCol)
        Next ilCol
    End If
    
    'TTP 10258: ReRate - make it work without requiring Office
    If ExptReRate.ckcCsv.Value = vbUnchecked Then
        ilRet = gCreateFormControlFile(ExptReRate, "ReRate", smBypassCtrlNames(), False)
        Screen.MousePointer = vbDefault
        'Save fails sometimes
        'ilRet = MsgBox("Focus will be sent to Excel so that you can View and Save the generated file", vbOKOnly + vbApplicationModal + vbInformation, "Excel")
        ilRet = MsgBox("This report will be sent to Excel for you to review and save", vbOKOnly + vbApplicationModal + vbInformation, "ReRate")
        ilRet = gExcelOutputGeneration("V")
        'Save as
    Else
        'CSV Only - close file handle before prompting user
        Close hmToCSV
        ilRet = MsgBox("Export Successfully Completed, Exported File " & smToCSV, vbOKOnly + vbApplicationModal + vbInformation, "ReRate")
    End If
    
    'ilRet = gExcelOutputGeneration("S", omBook, omSheet, 1, smToExcel)
    
    'Screen.MousePointer = vbDefault
    
    'ilRet = MsgBox("The Excel file saved as " & sgCR & sgLF & sgExportPath & smToExcel & sgCR & sgLF & " would you like to open the file using Excel", vbYesNo + vbQuestion + vbApplicationModal, "Excel File")
    'If ilRet = vbYes Then
    '    ilRet = gExcelOutputGeneration("V")
    'End If
    cmcExport.Enabled = True 'All Done Exporting
    
'12/1/2020 - TTP 9765 Object Required error message when driving crazy
mErrorSkipHere:
    cmcExport.Enabled = True
    Screen.MousePointer = vbDefault
    
    Close hmToCSV
    Close hmToExcel
    cmcCancel.Caption = "&Done"
    
    If ilRet > -1 Then       'error will be an error code (as long as ilRet isn't 0, but often times ilRet=False is returned from functions, and False=0)
        'Int(False) = 0
        'lacInfo(0).Caption = "Export Failed"
        gLogMsg "Export failed: #" & Trim$(str$(ilRet)), "ExportReRate.txt", False
    Else
        'Int(True) = -1
        'lacInfo(0).Caption = "Export Successfully Completed"
        If ckcCsv.Value = vbChecked Then
            gLogMsg "Export Successfully Completed, Exported File " & smToCSV, "ExportReRate.txt", False
        Else
            gLogMsg "Export Successfully Completed, Exported File " & "to Excel", "ExportReRate.txt", False
        End If
    End If
    'lacInfo(1).Caption = "Export Files: " & smToExcel1 & " and " & smToExcel12
    
    'lacInfo(0).Visible = True
    'lacInfo(1).Visible = True
    'cmcCancel.SetFocus
    'cmcExport.Enabled = False
    imExporting = False
    Exit Sub

End Sub

Private Sub cmcExport_GotFocus()
    mSetShow
End Sub

Private Sub cmcReturn_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate True
  
End Sub

Private Sub cmcReturn_GotFocus()
    mSetShow
End Sub

Private Sub cmcSetBook_Click()
    Dim slStartDate As String               'Contract start date
    Dim slEndDate As String                 'contract end date
    Dim ilBook As Integer
    cmcSetBook.Enabled = False 'TTP 10242 - ReRate error when double-clicking "set books by line
    'Add/Remove Contract
    mDetermineDateRange slStartDate, slEndDate
    If rbcDatesBy(3).Value And edcDate(0).Text <> "" Then   'By Contract
        slStartDate = "1/1/1970"
    End If
    lgReRateStartDate = gDateValue(slStartDate)
    lgReRateEndDate = gDateValue(slEndDate)
    mAddRemoveCntrByLine
    mGetAllowedLengths
    ReRateLineBook.Show vbModal
    If igTerminateReturn = 1 Then
        For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
            '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
            If tgBookByLineAssigned(ilBook).iReRateDnfCode <> 0 Then
                'TTP 10140, show When [Set Book by Line] method is enabled - Keep track of Last "Research Book Name" option (rbcReRateBook)
                bmBookByLine = True
                rbcReRateBookByLine.Value = True
                rbcReRateBookByLine.Enabled = True
                If rbcReRateBook(0).Value Then rbcReRateBook(0).Value = False: imReRateLastBookMode = 0
                If rbcReRateBook(1).Value Then rbcReRateBook(1).Value = False: imReRateLastBookMode = 1
                If rbcReRateBook(2).Value Then rbcReRateBook(2).Value = False: imReRateLastBookMode = 2
                If rbcReRateBook(3).Value Then rbcReRateBook(3).Value = False: imReRateLastBookMode = 3
                'rbcReRateBook(1).Value = False
                'rbcReRateBook(2).Value = False
                'rbcReRateBook(3).Value = False
                Exit For
            End If
        Next ilBook
    Else
        If bmBookByLine = True And UBound(tgBookByLineAssigned) > 0 Then
            'there's some saved, they hit cancel, but donsnt mean they wanted to clear the books
        Else
            mClearBookByLine
        End If
    End If
    mSetCommands
    cmcSetBook.Enabled = True 'TTP 10242 - ReRate error when double-clicking "set books by line
End Sub

Private Sub edcCSV_Change()
    mSetCommands
End Sub

Private Sub edcCSV_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDate_CalendarChanged(Index As Integer)
    'mPopCntr
    'mSetCommands
End Sub

Private Sub edcDate_Change(Index As Integer)
    tmcDelay.Enabled = False
    'mPopCntr
    'mSetCommands
    tmcDelay.Enabled = True
End Sub

Private Sub edcDate_GotFocus(Index As Integer)
    tmcDelay.Enabled = False
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDate_LostFocus(Index As Integer)
    tmcDelay.Enabled = False
    mPopCntr
    mSetCommands
End Sub

Private Sub edcStart_Change()
    mPopCntr
    mSetCommands
End Sub

Private Sub edcStart_gotfocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStart_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcStart.Text
    If (Trim$(slStr) = "") And (KeyAscii = KEY0) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = Left$(slStr, edcStart.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcStart.SelStart - edcStart.SelLength)
    If rbcDatesBy(1) Then
        If gCompNumberStr(slStr, "12") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Else
        If gCompNumberStr(slStr, "4") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcYear_Change()
    mPopCntr
    mSetCommands
End Sub

Private Sub edcYear_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    Me.Refresh
    'TTP 10140, show When [Set Book by Line] method is enabled - Keep track of Last "Research Book Name" option (rbcReRateBook)
    'Disable/Unset this Radio...
    'rbcReRateBookByLine.Enabled = False
    rbcReRateBookByLine.Value = False
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    mSetShow
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width      'move off the screen so screen won't flash
    End If
    imMGCheckState = ckcMG.Value
    imRatingCheckState = ckcRating.Value
    imAudioTypeCheckState = ckcAudioType.Value
    imPriceTypeCheckState = ckcPriceType.Value
    imCommentCheckState = ckcComment.Value
    imACT1LineupCheckState = ckcACT1Lineup.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    Set omSheet = Nothing
    Set omBook = Nothing
    ilRet = gExcelOutputGeneration("Q")
    Set ogExcel = Nothing
    
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    
    ilRet = btrClose(hmDnf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmCbf)
    ilRet = btrClose(hmDpf)
    ilRet = btrClose(hmDef)
    ilRet = btrClose(hmDrf)
    btrDestroy hmRaf
    btrDestroy hmDnf
    btrDestroy hmMnf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmCHF
    btrDestroy hmCbf
    btrDestroy hmDpf
    btrDestroy hmDef
    btrDestroy hmDrf
    'TTP 10193 - Add Line Comment
    ilRet = btrClose(hmCxf)
    btrDestroy hmCHF
    
    'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates for MG/OS
    ilRet = btrClose(hlSmf)
    btrDestroy hlSmf

    Erase lmAdvtOrderCost, imAdvtOrderRtg, lmAdvtOrderGrimp, lmAdvtOrderGRP
    Erase lmAdvtReRateCost, imAdvtReRateRtg, lmAdvtReRateGrimp, lmAdvtReRateGRP
    Erase lmAdvtPlusBonusReRateCost, imAdvtPlusBonusReRateRtg, lmAdvtPlusBonusReRateGrimp, lmAdvtPlusBonusReRateGRP
    Erase lmAdvtBonusReRateCost, imAdvtBonusReRateRtg, lmAdvtBonusReRateGrimp, lmAdvtBonusReRateGRP

    Erase tmClfPurchase, tmCffPurchase, tmChfAdvtExt
    Erase tmClfReRate, tmCffReRate
    Erase imRtg, lmGrimp, lmGRP, lmCost
    
    Erase lmWklyspots, lmWklyRates, lmWklyAvgAud, lmWklyPopEst, lmWklyMoDate
    Erase tmReRate
    Erase tmMGInfo, tmBonusInfo  'imBonusVefCode
    Erase smSummaryRecords, tmFormulaInfo
    Erase tgBookByLineCntr, tgBookByLineAssigned, igReRateAllowedLengths, igExcludeDnfCode
    Erase tgBookInfo, tgBookVehicle
    Erase tmRateCardSort, lmCost, imRtg, lmGrimp, lmGRP
    Erase tmBookName, tmSdfExtSort, tmSdfExt, tmMnfList, tmReRateBookDnfCodes
    
    Set ExptReRate = Nothing   'Remove data segment

End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim slStdDate As String
    Dim illoop As Integer
    
    'TTP 10309: ReRate: a comma in the agency name causes data to not line up with column headers when running the report with the "merged" option
    smDelimiter = Chr$(30) 'ASCII 30 is defined as a "Record Separator" - https://www.asciitable.com/
    
    imTerminate = False
    imFirstActivate = True
    Screen.MousePointer = vbHourglass
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    imAdfCode = -1
    imCtrlKey = False
    bmInGrid = False
    bmAllClicked = False
    bmSetAll = True
    lmEnableRow = -1
    lmEnableCol = -1

    ReDim tmClfPurchase(0 To 0) As CLFLIST
    tmClfPurchase(0).iStatus = -1 'Not Used
    tmClfPurchase(0).lRecPos = 0
    tmClfPurchase(0).iFirstCff = -1
    ReDim tmCffPurchase(0 To 0) As CFFLIST
    tmCffPurchase(0).iStatus = -1 'Not Used
    tmCffPurchase(0).lRecPos = 0
    tmCffPurchase(0).iNextCff = -1

    ReDim tmClfReRate(0 To 0) As CLFLIST
    tmClfReRate(0).iStatus = -1 'Not Used
    tmClfReRate(0).lRecPos = 0
    tmClfReRate(0).iFirstCff = -1
    ReDim tmCffReRate(0 To 0) As CFFLIST
    tmCffReRate(0).iStatus = -1 'Not Used
    tmCffReRate(0).lRecPos = 0
    tmCffReRate(0).iNextCff = -1

    'ReDim tgBookByLineCntr(0 To 0) As BOOKBYLINECNTR
    'ReDim tgBookByLineAssigned(0 To 0) As BOOKBYLINEASSIGNED
    'bmBookByLine = False
    'rbcReRateBookByLine.Value = False
    'rbcReRateBookByLine.Enabled = True
    mClearBookByLine
    
    lbcAdvertiser.Clear
    lbcSpotLens.Clear
    lbcCntrCode.Clear
    mClearGrid
    cbcRevision.BackColor = &HFFFF00
    frcMQ.Move lacTitle(0).Left, 15
    'If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
    '    ckcCsv.Visible = True
    '    edcCSV.Visible = True
    '    rbcReRateBook(3).Visible = True
    'Else
    '    ckcCsv.Visible = False
    '    edcCSV.Visible = False
    '    rbcReRateBook(3).Visible = False
    '    cmcExport.Top = frcIndex.Top + 2 * frcIndex.Height
    '    cmcCancel.Top = cmcExport.Top
    '    cmcReturn.Top = cmcExport.Top
    '    ExptReRate.Height = cmcExport.Top + 2 * cmcExport.Height
    'End If
    If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) <> ACT1CODES Then
        ckcACT1Lineup.Enabled = False
        ckcACT1Lineup.Value = vbUnchecked
    End If
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "AGF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: AGF.Btr)", ExptReRate
    imAgfRecLen = Len(tmAgf)
    
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "ADF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: ADF.Btr)", ExptReRate
    imAdfRecLen = Len(tmAdf)
    
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imPrfRecLen = Len(tmPrf)
        
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imCHFRecLen = Len(tmChf)
    
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imClfRecLen = Len(tmClfPurchase(0).ClfRec)
    
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imCffRecLen = Len(tmCffPurchase(0).CffRec)
    
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imMnfRecLen = Len(tmMnf)
    
    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imDnfRecLen = Len(tmDnf)
        
    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imCbfRecLen = Len(tmCbf)
        
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imDrfRecLen = Len(tmDrf)
   
    hmDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()     '7-23-01
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    imDpfRecLen = Len(tmDpf)
    
    hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()     '7-23-01
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()     '7-23-01
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'TTP 10193 - Add Line Comment
    hmCxf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCxf, "", sgDBPath & "CXF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCxfRecLen = Len(tmCxf)
    
         
    'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates for MG/OS
    hlSmf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExptReRate
    ilSmfRecLen = Len(tlSmf)
    
    edcDate(0).Text = ""
    
    mPopAdvt
    ilRet = gObtainRcfRifRdf()
    mSortRcf
    
    mPopSpotLength
    
    mPopBooks
    
    mPopExcludeBooks
    
    mPopDemo
    If rbcDemo(4).Value = True Then
        cbcDemo.Enabled True
    Else
        cbcDemo.Enabled False
    End If
    
    mGridColumnWidths
    mGridColumnTitles
    
    imCostColumn(0) = 7
    imCostColumn(1) = 8
    imCostColumn(2) = 17
    imCPMCPPColumn(0) = 12
    imCPMCPPColumn(1) = 13
    imCPMCPPColumn(2) = 21
    imCPMCPPColumn(3) = 22
    imRatingColumn(0) = 11
    imRatingColumn(1) = 20
    imReRateColumn = REXTTOTALEXCEL '17
    imPurchasedColumn = PEXTTOTALEXCEL '8
    
    For illoop = 1 To UBound(smColumnLetter) Step 1
        smColumnLetter(illoop) = gExcelColumnToLetter(illoop)
    Next illoop
    
    smBypassCtrlNames(0) = "lbcAdvertiser"
    smBypassCtrlNames(1) = "grdCntr"
    smBypassCtrlNames(2) = "ckcAllCntr"
    smBypassCtrlNames(3) = "lbcSpotLens"
    smBypassCtrlNames(4) = "ckcAllSpotLens"
    smBypassCtrlNames(5) = "edcDate(0)"
    smBypassCtrlNames(6) = "cbcRevision"
    smBypassCtrlNames(7) = "lacTitle(3)"
    smBypassCtrlNames(8) = "lacTitle(4)"
    smBypassCtrlNames(9) = "edcStart"
    smBypassCtrlNames(10) = "edcYear"
    smBypassCtrlNames(11) = "lacTitle(5)"
    smBypassCtrlNames(12) = "lacTitle(6)"
    smBypassCtrlNames(13) = "edcDate(1)"
    smBypassCtrlNames(14) = "edcDate(2)"
    smBypassCtrlNames(15) = "lbcBookNames"
    smBypassCtrlNames(16) = "lbcCntrCode"
    
    If tgSpf.sSAudData = "H" Then
        imNumberDecPlaces = 1
        imAdjDecPlaces = 10
    ElseIf tgSpf.sSAudData = "N" Then
        imNumberDecPlaces = 2
        imAdjDecPlaces = 100
    ElseIf tgSpf.sSAudData = "U" Then
        imNumberDecPlaces = 3
        imAdjDecPlaces = 1000
    Else
        imNumberDecPlaces = 0
        imAdjDecPlaces = 1
    End If
    
    gSetFormCtrls ExptReRate, "ReRate"
    
    'TTP 10258: ReRate - make it work without requiring Office (CSV Option always Visible)
    'If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
    ckcCsv.Visible = True
    edcCSV.Visible = True
    'rbcReRateBook(3).Visible = True
    'Else
    '    ckcCsv.Visible = False
    '    edcCSV.Visible = False
    '    frcOutputOptions.Height = frcOutputOptions.Height - 285
    '    frcResearchOptions.Top = frcResearchOptions.Top + 60
    '    frcOutputOptions.Top = frcOutputOptions.Top + 120
    'End If
    
    gCenterStdAlone ExptReRate
    
    frcMQ.Move lacTitle(0).Left, lacTitle(0).Top
    lacTitle(5).Move lacTitle(0).Left, lacTitle(0).Top
    edcDate(1).Move lacTitle(5).Left + lacTitle(5).Width, edcDate(0).Top
    lacTitle(6).Move edcDate(1).Left + edcDate(1).Width + 120, lacTitle(0).Top
    edcDate(2).Move lacTitle(6).Left + lacTitle(6).Width, edcDate(0).Top
    Screen.MousePointer = vbDefault

    'Sortable Color
    grdCntr.Col = 1
    grdCntr.Row = 0
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.Col = 2
    grdCntr.Row = 0
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.Col = 1
    grdCntr.Row = 1
    grdCntr.CellBackColor = LIGHTBLUE
    grdCntr.Col = 2
    grdCntr.Row = 1
    grdCntr.CellBackColor = LIGHTBLUE
    
    Exit Sub

mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mPopAdvt()
'   mAdvtPop
'   Where:
'       RptForm as Form
'       lbcSelection as control
'
'
    Dim ilRet As Integer
    ilRet = gPopAdvtBox(ExptReRate, lbcAdvertiser, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mPopAdvt (gPopAdvtBox)", ExptReRate
        On Error GoTo 0
    End If
    Exit Sub
mAdvtPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCntrPop                        *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                    :7/10/96 -Use new contract status*
'                                                      *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopCntr()
'
'   mCntrPop
'   Where:
'       slcntrStatus(I)- O; H; W; C; I; D or blank for all
'       ilHOState(I) - 1 only get cnt (w/o revision) H & O only
'                      2 combo - get latest orders includ revisions (H O G or N) if G or N, show instead of the H or O
'                      3 everything - revision & orders (HOGNWCI) if GNWCI, show over the H or O
'
'
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String  'Name and code
    Dim slCode As String    'Code number
    Dim ilCurrent As Integer
    Dim illoop As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim llCntrNo As Long
    Dim ilShow As Integer
    Dim slCntrType As String
    Dim ilAdfCode As Integer
    Dim llLen As Long
    Dim ilErr As Integer
    Dim slShow As String
    Dim ilRevNo As Integer
    Dim ilVerNo As Integer
    Dim ilExtRevNo As Integer
    Dim slRevNo As String
    Dim llRow As Long
    Dim slProduct As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilDemo As Integer
    Dim slCntrNo As String
    
    'TTP 10140, show When [Set Book by Line] method is enabled - Keep track of Last "Research Book Name" option (rbcReRateBook)
    'bmBookByLine = False
    'rbcReRateBookByLine.Value = False
    'rbcReRateBookByLine.Enabled = False
    'rbcReRateBook(imReRateLastBookMode).Value = True
    mClearBookByLine
    
    llLen = 0
    ilErr = False
    'Clear Grid
    grdCntr.Redraw = False
    mClearGrid
    lbcCntrCode.Clear
    If imAdfCode <= 0 Then
        grdCntr.Redraw = True
        Exit Sub
    End If
    mDetermineDateRange slStartDate, slEndDate
    If slStartDate = "1/1/1970" Then
        grdCntr.Redraw = True
        Exit Sub
    End If
    slCntrType = "C"
    If tgUrf(0).sResvType <> "H" Then
        slCntrType = slCntrType & "V"
    End If
    If tgUrf(0).sRemType <> "H" Then
        slCntrType = slCntrType & "T"
    End If
    If tgUrf(0).sDRType <> "H" Then
        slCntrType = slCntrType & "R"
    End If
    If tgUrf(0).sPIType <> "H" Then
        slCntrType = slCntrType & "Q"
    End If
    'If tgUrf(0).sPSAType <> "H" Then
    '    slCntrType = slCntrType & "S"
    'End If
    'If tgUrf(0).sPromoType <> "H" Then
    '    slCntrType = slCntrType & "M"
    'End If
    If slCntrType = "CVTRQSM" Then
        slCntrType = ""
    End If
    'ilShow = 1
    ilShow = 7  '5                  'show # and advt name
    ilCurrent = 1
    'ilRet = gPopCntrForAASBox(ExptReRate, 0, imAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, tgMultiCntrCodeAD(), sgMultiCntrCodeTagAD)
    ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
    sgCntrForDateStamp = ""
    'ilRet = gObtainCntrForDate(ExptReRate, slStartDate, slEndDate, "HO", slCntrType, 4, tmChfAdvtExt())
    ilRet = gCntrForActiveOHD(ExptReRate, slStartDate, slEndDate, "", "", "HO", slCntrType, 5, tmChfAdvtExt(), imAdfCode)
    If ilRet <> BTRV_ERR_NONE Then
        grdCntr.Redraw = True
        On Error GoTo mCntrPopErr
        gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", ExptReRate
        On Error GoTo 0
    End If
    'Sort by (Product) __contract number__
    For ilIndex = 0 To UBound(tmChfAdvtExt) - 1 Step 1 'lbcMultiCntrCode.ListCount - 1 Step 1
        If (tmChfAdvtExt(ilIndex).iAdfCode = imAdfCode) And (tmChfAdvtExt(ilIndex).iMnfDemo0 > 0) Then
            ilRet = True
            If tmChfAdvtExt(ilIndex).sAdServerDefined = "Y" Then
                'check if this is a Ad Server Only conrtact...
                ilRet = gExistClf(tmChfAdvtExt(ilIndex).lCode)          'False if No CLF records found
            End If
            If ilRet Then
                slCntrNo = tmChfAdvtExt(ilIndex).lCntrNo
                Do While Len(slCntrNo) < 6
                    slCntrNo = "0" & slCntrNo
                Loop
                lbcCntrCode.AddItem tmChfAdvtExt(ilIndex).sProduct & "|" & slCntrNo
                lbcCntrCode.ItemData(lbcCntrCode.NewIndex) = ilIndex
            End If
        End If
    Next ilIndex
    llRow = grdCntr.FixedRows
    grdCntr.Rows = 3 '3/4/21 - TTP 10088: clear extra rows for Sort Feature to not sort Blank lines.
    For illoop = 0 To lbcCntrCode.ListCount - 1 Step 1
        ilIndex = lbcCntrCode.ItemData(illoop)
        If llRow >= grdCntr.Rows Then
            grdCntr.AddItem ""
            grdCntr.RowHeight(llRow) = fgFlexGridRowH
        End If
        grdCntr.Row = llRow
        tmChfSrchKey0.lCode = tmChfAdvtExt(ilIndex).lCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
        
            If (tmChfAdvtExt(ilIndex).iExtRevNo <> 0) Or (rbcRevNo(1).Value) Then
                tmChfSrchKey1.lCntrNo = tmChfAdvtExt(ilIndex).lCntrNo
                If (rbcRevNo(0).Value) Then
                    tmChfSrchKey1.iCntRevNo = tmChfAdvtExt(ilIndex).iCntRevNo
                Else
                    tmChfSrchKey1.iCntRevNo = 32000
                End If
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmChfAdvtExt(ilIndex).lCntrNo)
                    If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
                        tmChfAdvtExt(ilIndex).sProduct = tmChf.sProduct
                        tmChfAdvtExt(ilIndex).iExtRevNo = tmChf.iExtRevNo
                        tmChfAdvtExt(ilIndex).iCntRevNo = tmChf.iCntRevNo
                        tmChfAdvtExt(ilIndex).lCode = tmChf.lCode
                        If rbcRevNo(1).Value Then
                            Exit Do
                        End If
                    End If
                    ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            grdCntr.TextMatrix(llRow, PRODUCTINDEX) = Trim$(tmChfAdvtExt(ilIndex).sProduct)
            grdCntr.TextMatrix(llRow, CNTRNOINDEX) = tmChfAdvtExt(ilIndex).lCntrNo
            If tmChfAdvtExt(ilIndex).iExtRevNo = 0 Then
                grdCntr.TextMatrix(llRow, VERSIONINDEX) = "Original"
            Else
                grdCntr.TextMatrix(llRow, VERSIONINDEX) = "R" & tmChfAdvtExt(ilIndex).iCntRevNo & "-" & tmChfAdvtExt(ilIndex).iExtRevNo
            End If
            grdCntr.TextMatrix(llRow, PURCHASECHFCODEINDEX) = tmChfAdvtExt(ilIndex).lCode
            grdCntr.TextMatrix(llRow, STARTREVNOINDEX) = tmChfAdvtExt(ilIndex).iExtRevNo
            grdCntr.TextMatrix(llRow, NODEMOSINDEX) = 0
            For ilDemo = 0 To UBound(tmChf.iMnfDemo) Step 1
                If tmChf.iMnfDemo(ilDemo) > 0 Then
                    grdCntr.TextMatrix(llRow, NODEMOSINDEX) = ilDemo + 1
                End If
            Next ilDemo
            grdCntr.TextMatrix(llRow, ENDREVNOINDEX) = -1
            tmChfSrchKey1.lCntrNo = tmChfAdvtExt(ilIndex).lCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmChfAdvtExt(ilIndex).lCntrNo)
                If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
                    grdCntr.TextMatrix(llRow, RERATECHFCODEINDEX) = tmChf.lCode
                    grdCntr.TextMatrix(llRow, ENDREVNOINDEX) = tmChf.iExtRevNo
                    If tmChf.sCBSOrder = "C" Then
                        grdCntr.Col = PRODUCTINDEX
                        grdCntr.CellForeColor = vbRed
                        grdCntr.Col = CNTRNOINDEX
                        grdCntr.CellForeColor = vbRed
                    End If
                    Exit Do
                End If
                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            llRow = llRow + 1
        End If
    Next illoop
    For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        'If grdCntr.TextMatrix(llRow, PRODUCTINDEX) <> "" Then
        If grdCntr.TextMatrix(llRow, CNTRNOINDEX) <> "" Then
            mPaintRowColor llRow
        End If
    Next llRow
    'grdCntr.Row = grdCntr.FixedRows
    gGrid_AlignAllColsLeft grdCntr
    grdCntr.Redraw = True
    '3/4/21 - TTP 10088: Retain last sort, by performing the last Sort by Product or Contract (ASC/DESC); using a Negative Column #
    If imSortColumn = 0 Then imSortColumn = 1: imSortDir = flexSortGenericAscending 'Product
    mSortByColumn -imSortColumn
    Screen.MousePointer = vbDefault
    Exit Sub
mCntrPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Public Function mPopBooks() As Integer
    Dim ilRet As Integer
    Dim ilVefCode As Integer
    Dim ilSort As Integer
    Dim ilShow As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim illoop As Integer
    Dim llLastIndex As Long
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    ReDim imRBDnfCode(0 To 0) As Integer
    
    ilVefCode = 0
    ilSort = 1  'sort by date, then book name
    ilShow = 1  'show book name, then date
    ilRet = gPopBookNameBox(ExptReRate, 0, 0, ilVefCode, ilSort, ilShow, lbcBookNames, tmBookName(), smBookNameTag)
    ReDim tgBookInfo(0 To lbcBookNames.ListCount) As BOOKINFO
    ReDim Preserve tgBookVehicle(0 To 0) As BOOKVEHICLE
    For illoop = 0 To lbcBookNames.ListCount - 1 Step 1
        slNameCode = tmBookName(illoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        'lbcBookNames.ItemData(ilLoop) = Val(slCode)
        tmDnfSrchKey0.iCode = Val(slCode)
        ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tgBookInfo(illoop).iDnfCode = tmDnf.iCode
            tgBookInfo(illoop).sName = tmDnf.sBookName
            gUnpackDateLong tmDnf.iBookDate(0), tmDnf.iBookDate(1), tgBookInfo(illoop).lBookDate
            tgBookInfo(illoop).lFirst = -1
            slSQLQuery = "Select Distinct drfVefCode from DRF_Demo_Rsrch_Data where drfDnfCode = " & tmDnf.iCode
            slSQLQuery = slSQLQuery & " And drfDemoDataType <> 'P' And drfMnfSocEco = 0"
            Set tmp_rst = gSQLSelectCall(slSQLQuery)
            Do While Not tmp_rst.EOF
                If tgBookInfo(illoop).lFirst = -1 Then
                    tgBookInfo(illoop).lFirst = UBound(tgBookVehicle)
                    tgBookVehicle(UBound(tgBookVehicle)).iVefCode = tmp_rst!drfVefCode
                    tgBookVehicle(UBound(tgBookVehicle)).lNext = -1
                    llLastIndex = UBound(tgBookVehicle)
                    ReDim Preserve tgBookVehicle(0 To UBound(tgBookVehicle) + 1) As BOOKVEHICLE
                Else
                    tgBookVehicle(llLastIndex).lNext = UBound(tgBookVehicle)
                    tgBookVehicle(UBound(tgBookVehicle)).iVefCode = tmp_rst!drfVefCode
                    tgBookVehicle(UBound(tgBookVehicle)).lNext = -1
                    llLastIndex = UBound(tgBookVehicle)
                    ReDim Preserve tgBookVehicle(0 To UBound(tgBookVehicle) + 1) As BOOKVEHICLE
                End If
                tmp_rst.MoveNext
            Loop
        End If
    Next illoop
    
    'lbcBookNames.ListIndex = -1
    lbcBookNames.Clear
End Function

Function mFindClosestBook(llDate As Long, llSdf As Long) As Integer
    Dim ilDnf As Integer
    Dim llNext As Long
    Dim ilVefCode As Integer
    Dim illoop As Integer
    Dim blFound As Boolean
    
    'Books are in descending date order
    ilVefCode = tmSdfExt(llSdf).iVefCode
    For ilDnf = 0 To UBound(tgBookInfo) - 1 Step 1
        If tgBookInfo(ilDnf).lBookDate <= llDate Then
            If gBinarySearchExcludeBook(tgBookInfo(ilDnf).iDnfCode) = -1 Then
                llNext = tgBookInfo(ilDnf).lFirst
                Do While llNext <> -1
                    If tgBookVehicle(llNext).iVefCode = ilVefCode Then
                        If (tmSdfExt(llSdf).sSchStatus = "S") And (tmSdfExt(llSdf).sSpotType <> "X") And (tmSdfExt(llSdf).lMdDate = 0) Then
                            ''blFound = False
                            ''For ilLoop = 0 To UBound(tmReRateBookDnfCodes) - 1 Step 1
                            ''    If (tmReRateBookDnfCodes(ilLoop).lChfCode = tmSdfExt(llSdf).lChfCode) And (tmReRateBookDnfCodes(ilLoop).iLineNo = tmSdfExt(llSdf).iLineNo) And (tmReRateBookDnfCodes(ilLoop).iDnfCode = tgBookInfo(ilDnf).iDnfCode) Then
                            ''        blFound = True
                            ''        Exit For
                            ''    End If
                            ''Next ilLoop
                            ''If Not blFound Then
                            ''    tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).lChfCode = tmSdfExt(llSdf).lChfCode
                            ''    tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).iLineNo = tmSdfExt(llSdf).iLineNo
                            ''    tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).iDnfCode = tgBookInfo(ilDnf).iDnfCode
                            ''    ReDim Preserve tmReRateBookDnfCodes(0 To UBound(tmReRateBookDnfCodes) + 1) As RERATEBOOKDNFCODES
                            ''End If
                            'mSaveReRateBookByLine llSdf, tgBookInfo(ilDnf).iDnfCode
                        End If
                        mFindClosestBook = tgBookInfo(ilDnf).iDnfCode
                        Exit Function
                    End If
                    llNext = tgBookVehicle(llNext).lNext
                Loop
            End If
        End If
    Next ilDnf
    mFindClosestBook = 0
End Function

Sub mExport()
    Dim ilRet As Integer
    Dim ilRet2 As Integer
    Dim ilChf As Integer            'number of contracts processed so far
    Dim illoop As Integer                   'temp loop variable
    Dim ilLoop2 As Integer                  'temp loop variable
    Dim ilLoop3 As Integer                  'temp loop variable
    Dim llContrCode As Long                 'Contr ID to process
    Dim slStartDate As String               'Contract start date
    Dim slEndDate As String                 'contract end date
    Dim llStartDate As Long                 'quarter requested start date
    Dim llEndDate As Long                   'quarter reqested end date
    Dim ilClfP As Integer                    'loop for lines
    Dim ilCff As Integer                    'loop for flights
    Dim ilClfR As Integer                    'loop for lines
    Dim slStr As String                     'temp string for conversions
    Dim llDate As Long                      'temp serial date
    Dim llDate2 As Long
    Dim ilDay As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slCntrType As String                'valid contract types (per inq, direct respon, etc) to retrieve
    Dim slCntrStatus As String              'valid contr status (working, complete, etc) to retrieve
    Dim ilHOState As Integer                'which type of Holds Orders to retrieved (internally WCI)
    Dim llPop As Long                       'population obtained per schedule line
    Dim llAvgAud As Long                    'avg audience obtained per flight
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim ilSpots As Integer
    'Dim llOvStartTime As Long
    'Dim llOvEndTime As Long
    ReDim ilInputDays(0 To 6) As Integer    'valid days of the week for audience retrieval
    Dim ilUpperWk As Integer
    Dim llTemp As Long
    Dim ilReRate As Integer
    Dim llTotalCPP As Long
    Dim llTotalCPM As Long
    Dim llTotalGrImp As Long
    Dim llTotalGRP As Long
    'Dim llTotalCost As Long
    Dim dlTotalCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim llTotalAvgAud As Long
    Dim ilTotalAvgRtg As Integer
    Dim llLnSpots As Long
    Dim ilTotLnSpts As Integer
    Dim llPopEst As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim ilSpotType As Integer   '0=Order; 1=Aired; 2=MG; 3=Bonus
    Dim ilDnfCode As Integer
    Dim ilPackageDnfCode As Integer     '-1=Not set; -2=Mixture; > 0 matching across hidden lines
    Dim ilCntrTotalDnfCode As Integer
    Dim ilBonusDnfCode As Integer
    Dim llLnStartDate As Long
    Dim llLnEndDate As Long
    Dim slCBS As String
   
    Dim ilNoWks As Integer
    Dim ilHiddenLines As Integer
    Dim ilLine As Integer
    Dim llTotalCntrSpots As Long
    
    Dim llSeqNo As Long
    Dim ilVef As Integer
    Dim ilVefIndex As Integer
    Dim ilRow As Integer
    Dim ilAdvtOrderRow As Integer
    Dim ilAdvtReRateRow As Integer
    Dim ilAdvtBonusReRateRow As Integer
    Dim ilAdvtPlusBonusReRateRow As Integer
    Dim blReRateAdded As Boolean
    Dim llPurchaseLineAud As Long
    
    Dim tlContractOrderReRate As RERATEINFO
    Dim tlAdvtOrderReRate As RERATEINFO
    Dim ilFind As Integer
    Dim tlClf As CLF
    Dim ilExtRevNo As Integer
    Dim ilClf As Integer
    Dim ilBook As Integer
    Dim ilPlusDnfCode As Integer
    Dim ilSdf As Integer
    Dim iProcessedCount  As Integer
    Dim iCntrCount  As Integer
    bmInSummaryMode = False
    
    '7-23-01 setup global variable for Demo Plus file (to see if any exists)
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If

    mDetermineDateRange slStartDate, slEndDate
    ilNoWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
    llStartDate = gDateValue(slStartDate)           'convert string date to long for date comparisons
    llEndDate = gDateValue(slEndDate)               'convert string end date to long for date comparisons

    ilSpotType = 0
    ReDim lmAdvtOrderCost(0 To 0) As Long
    ReDim imAdvtOrderRtg(0 To 0) As Integer
    ReDim lmAdvtOrderGrimp(0 To 0) As Long
    ReDim lmAdvtOrderGRP(0 To 0) As Long
    lmAdvtOrderTotalSpots = 0
    lmAdvtOrderPop = -1
    imAdvtOrderDnfCode = -1
    
    ReDim lmAdvtReRateCost(0 To 0) As Long
    ReDim imAdvtReRateRtg(0 To 0) As Integer
    ReDim lmAdvtReRateGrimp(0 To 0) As Long
    ReDim lmAdvtReRateGRP(0 To 0) As Long
    lmAdvtReRateTotalSpots = 0
    lmAdvtReRatePop = -1
    imAdvrReRateDnfCode = -1
    
    ReDim lmAdvtBonusReRateCost(0 To 0) As Long
    ReDim imAdvtBonusReRateRtg(0 To 0) As Integer
    ReDim lmAdvtBonusReRateGrimp(0 To 0) As Long
    ReDim lmAdvtBonusReRateGRP(0 To 0) As Long
    lmAdvtBonusReRateTotalSpots = 0
    lmAdvtBonusReRatePop = -1
    imAdvtBonusReRateDnfCode = -1
    
    ReDim lmAdvtPlusBonusReRateCost(0 To 0) As Long
    ReDim imAdvtPlusBonusReRateRtg(0 To 0) As Integer
    ReDim lmAdvtPlusBonusReRateGrimp(0 To 0) As Long
    ReDim lmAdvtPlusBonusReRateGRP(0 To 0) As Long
    lmAdvtPlusBonusReRateTotalSpots = 0
    lmAdvtPlusBonusReRatePop = -1
    imAdvtPlusBonusReRateDnfCode = -1
    
    ilAdvtOrderRow = 0
    ilAdvtReRateRow = 0
    ilAdvtBonusReRateRow = 0
    ilAdvtPlusBonusReRateRow = 0
    imNoCntr = 0
    For ilChf = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        'If grdCntr.TextMatrix(ilChf, PRODUCTINDEX) <> "" And grdCntr.TextMatrix(ilChf, SELECTEDINDEX) = "1" Then
        If grdCntr.TextMatrix(ilChf, CNTRNOINDEX) <> "" And grdCntr.TextMatrix(ilChf, SELECTEDINDEX) = "1" Then
            imNoCntr = imNoCntr + 1
        End If
    Next ilChf
    'For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1                                           'loop while llCurrentRecd < llRecsRemaining
''    For ilChf = 0 To lbcCntrCode.ListCount - 1 Step 1
''        If lbcCntrCode.Selected(ilChf) Then
'    For ilChf = 0 To lbcMultiCntr.ListCount - 1 Step 1
'        If lbcMultiCntr.Selected(ilChf) Then
    'JW Bonus Improvement
    iCntrCount = 0
    iProcessedCount = 0
    prgProcessing.Value = 0
    prgProcessing.Max = 100
    prgProcessing.Visible = True
    lblProcessing.Caption = "Processing"
    frcProcessing.Top = 840 'Keep it out of the control area because other control will show ontop
    frcProcessing.Left = (Me.ScaleWidth - frcProcessing.Width) / 2
    For ilChf = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        If grdCntr.TextMatrix(ilChf, CNTRNOINDEX) <> "" And grdCntr.TextMatrix(ilChf, SELECTEDINDEX) = "1" Then
            iCntrCount = iCntrCount + 1
        End If
    Next ilChf
    If iCntrCount > 1 Then frcProcessing.Visible = True
    
    ReDim tmReRateBookDnfCodes(0 To 0) As RERATEBOOKDNFCODES
    For ilChf = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        'If grdCntr.TextMatrix(ilChf, PRODUCTINDEX) <> "" And grdCntr.TextMatrix(ilChf, SELECTEDINDEX) = "1" Then
        If grdCntr.TextMatrix(ilChf, CNTRNOINDEX) <> "" And grdCntr.TextMatrix(ilChf, SELECTEDINDEX) = "1" Then
            'imNoCntr = imNoCntr + 1
            ReDim tmReRate(0 To 0) As RERATEINFO
            ReDim tmClfPurchase(0 To 0) As CLFLIST
            tmClfPurchase(0).iStatus = -1 'Not Used
            tmClfPurchase(0).lRecPos = 0
            tmClfPurchase(0).iFirstCff = -1
            ReDim tmCffPurchase(0 To 0) As CFFLIST
            tmCffPurchase(0).iStatus = -1 'Not Used
            tmCffPurchase(0).lRecPos = 0
            tmCffPurchase(0).iNextCff = -1
            
            ReDim tmClfReRate(0 To 0) As CLFLIST
            tmClfReRate(0).iStatus = -1 'Not Used
            tmClfReRate(0).lRecPos = 0
            tmClfReRate(0).iFirstCff = -1
            ReDim tmCffReRate(0 To 0) As CFFLIST
            tmCffReRate(0).iStatus = -1 'Not Used
            tmCffReRate(0).lRecPos = 0
            tmCffReRate(0).iNextCff = -1
            
            llTotalCntrSpots = 0
            
            'ReDim imBonusVefCode(0 To 0) As Integer
            ReDim tmBonusInfo(0 To 0) As MGBONUSINFO
            
            '**************************************
            'Obtain the ReRate population from the first scheduled spotand use for all ReRate computations
            '**************************************
            lmReRatePop = -1
            imBonusTotalDnfCode = -1
            'If grdCntr.TextMatrix(ilChf, VERSIONINDEX) = "Original" Then
                slCode = grdCntr.TextMatrix(ilChf, PURCHASECHFCODEINDEX)
                llContrCode = Val(slCode)
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tmChfPurchase, tmClfPurchase(), tmCffPurchase(), False)  '8-28-12 do not sort by special dp order
                lblProcessing.Caption = "Processing Contract " & iProcessedCount + 1 & " of " & iCntrCount & " (Contract # " & tmChfPurchase.lCntrNo & ")..."
            'Else
            '    tmChfSrchKey1.lCntrNo = grdCntr.TextMatrix(ilChf, CNTRNOINDEX)
            '    tmChfSrchKey1.iCntRevNo = 32000
            '    tmChfSrchKey1.iPropVer = 32000
            '    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            '    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = grdCntr.TextMatrix(ilChf, CNTRNOINDEX))
            '        If grdCntr.TextMatrix(ilChf, VERSIONINDEX) = "Original" Then
            '            ilExtRevNo = 0
            '        Else
            '            ilExtRevNo = Val(grdCntr.TextMatrix(ilChf, VERSIONINDEX))
            '        End If
            '        If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.iExtRevNo = ilExtRevNo) Then
            '            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, tmChf.lCode, False, tmChfPurchase, tmClfPurchase(), tmCffPurchase(), False)  '8-28-12 do not sort by special dp order
            '            Exit Do
            '        End If
            '        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            '    Loop
            'End If
            mFilterLinesBySpotLength tmClfPurchase
            imMnfDemo = tmChfPurchase.iMnfDemo(0)
            If rbcDemo(1).Value Then
                imMnfDemo = tmChfPurchase.iMnfDemo(1)
            ElseIf rbcDemo(2).Value Then
                imMnfDemo = tmChfPurchase.iMnfDemo(2)
            ElseIf rbcDemo(3).Value Then
                imMnfDemo = tmChfPurchase.iMnfDemo(3)
            ElseIf rbcDemo(4).Value Then
                slNameCode = tgDemoCode(cbcDemo.ListIndex).sKey
                ilRet2 = gParseItem(slNameCode, 2, "\", slCode)
                imMnfDemo = CInt(slCode)
            End If
            If ilRet Then
                slCode = grdCntr.TextMatrix(ilChf, RERATECHFCODEINDEX)
                llContrCode = Val(slCode)
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tmChfReRate, tmClfReRate(), tmCffReRate(), False)  '8-28-12 do not sort by special dp order
                mFilterLinesBySpotLength tmClfReRate
                mMergeMissingLines
                mHandleHiddenLinesMoved
            End If
            'If (ilRet And tmChfPurchase.iPctTrade <> 100) Then                                  'get a contract and test for printables,
            If ilRet Then                                   'get a contract and test for printables,
                If rbcDatesBy(3).Value Then   'By Contract
                    gUnpackDate tmChfPurchase.iStartDate(0), tmChfPurchase.iStartDate(1), slStartDate
                    gUnpackDate tmChfPurchase.iEndDate(0), tmChfPurchase.iEndDate(1), slEndDate
                    ilNoWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
                    llStartDate = gDateValue(slStartDate)           'convert string date to long for date comparisons
                    llEndDate = gDateValue(slEndDate)               'convert string end date to long for date comparisons
                End If
                sm1or2PlaceRating = gSet1or2PlaceRating(tmChfPurchase.iAgfCode)
                ilTotLnSpts = 0                 'init total # spots per line
                            'search for all spots by for this contract
                ReDim tmSdfExtSort(0 To 0) As SDFEXTSORT
                ReDim tmSdfExt(0 To 0) As SDFEXT
                ilRet = gObtainCntrSpot(-1, False, tmChfReRate.lCode, -1, "S", slStartDate, slEndDate, tmSdfExtSort(), tmSdfExt(), 0, False, True) 'search for spots between requested user dates
                
                ''ReDim tmReRate(1 To 1) As RESEARCHINFO           'list of Research totals for cnt
                'ReDim tmReRate(0 To 0) As RESEARCHINFO           'list of Research totals for cnt
                For ilClfP = LBound(tmClfPurchase) To UBound(tmClfPurchase) - 1 Step 1
                    tmClfP = tmClfPurchase(ilClfP).ClfRec
                    lmReRatePop = -1
                    If tmClfP.sType = "H" Or tmClfP.sType = "S" Then   '3-6-01 process on hidden & std lines (no packages)
                        ReDim tmMGInfo(0 To 0) As MGBONUSINFO
                        llPurchaseLineAud = -1
                        ilSpotType = 0
                        ilTotLnSpts = 0                 'init total # spots per line
                        ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, tmClfP.iDnfCode, 0, imMnfDemo, llPop)
                        
                        ReDim lmWklyspots(0 To ilNoWks) As Long       'sched lines weekly # spots
                        ReDim lmWklyAvgAud(0 To ilNoWks) As Long             'sched lines weekly avg aud
                        ReDim lmWklyRates(0 To ilNoWks) As Long           'sched lines weekly rates
                        ReDim lmWklyPopEst(0 To ilNoWks) As Long
                        ReDim lmWklyMoDate(0 To ilNoWks) As Long
                        
                        mGetPurchasedSpotCount ilClfP, llStartDate, llEndDate, imMnfDemo, ilTotLnSpts, llTotalCntrSpots, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode

                        'Finished all flights, calculat the lines research values
                        ilUpperWk = ilNoWks + 1 '14
                        ReDim imRtg(0 To ilUpperWk)                   'setup arrays for return values from audtolnresearch
                        ReDim lmGrimp(0 To ilUpperWk)
                        ReDim lmGRP(0 To ilUpperWk)
                        ReDim lmCost(0 To ilUpperWk)                    'setup arrays for return values from audtolnresearch
                        blReRateAdded = False
                        'Schedule line complete, get its avg aud data for the line
                        'If ilTotLnSpts > 0 And llPop > 0 Then
                            'gAvgAudToLnResearch sm1or2PlaceRating, False, llPop, lmWklyPopEst(), lmWklyspots(), lmWklyRates(), lmWklyAvgAud(), llTotalCost, llTotalAvgAud, imRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                            gAvgAudToLnResearch sm1or2PlaceRating, False, llPop, lmWklyPopEst(), lmWklyspots(), lmWklyRates(), lmWklyAvgAud(), dlTotalCost, llTotalAvgAud, imRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                            'Build totals by line
                            ilReRate = UBound(tmReRate)
                            tmReRate(ilReRate).lChfCode(ilSpotType) = tmClfP.lChfCode
                            tmReRate(ilReRate).lClfCode(ilSpotType) = tmClfP.lCode
                            tmReRate(ilReRate).iVefCode = tmClfP.iVefCode
                            tmReRate(ilReRate).iRdfCode = tmClfP.iRdfCode
                            tmReRate(ilReRate).sType = tmClfP.sType
                            tmReRate(ilReRate).sSubType = ""
                            tmReRate(ilReRate).lCntrNo = tmChfPurchase.lCntrNo
                            tmReRate(ilReRate).sProduct = tmChfPurchase.sProduct
                            tmReRate(ilReRate).iLineNo = tmClfP.iLine
                            tmReRate(ilReRate).iPkLineNo = tmClfP.iPkLineNo
                            tmReRate(ilReRate).iLen = tmClfP.iLen
                            tmReRate(ilReRate).sAudioType = tmClfP.sLiveCopy 'TTP 10144
                            tmReRate(ilReRate).sACT1LineupCode = tmClfP.sACT1LineupCode
                            '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                            tmReRate(ilReRate).sACT1StoredTime = tmClfP.sACT1StoredTime
                            tmReRate(ilReRate).sACT1StoredSpots = tmClfP.sACT1StoredSpots
                            tmReRate(ilReRate).sACT1StoreClearPct = tmClfP.sACT1StoreClearPct
                            tmReRate(ilReRate).sACT1DaypartFilter = tmClfP.sACT1DaypartFilter
                            tmReRate(ilReRate).iDnfCode(ilSpotType) = tmClfP.iDnfCode
                            tmReRate(ilReRate).lPop(ilSpotType) = llPop
                            tmReRate(ilReRate).lSatelliteEst(ilSpotType) = llPopEst          '6-4-04
                            'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                            tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                            tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                            tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                            tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                            tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                            tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                            tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                            tmReRate(ilReRate).lTotalSpots(ilSpotType) = ilTotLnSpts
                            'tmReRate(ilReRate).sCBS = "N"
                            gUnpackDateLong tmClfP.iStartDate(0), tmClfP.iStartDate(1), llLnStartDate
                            gUnpackDateLong tmClfP.iEndDate(0), tmClfP.iEndDate(1), llLnEndDate
                            If llLnEndDate < llLnStartDate Then
                                tmReRate(ilReRate).sCBS(ilSpotType) = "Y"
                            Else
                                tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                            End If
                            If Trim(gStripChr0(tmReRate(ilReRate).sLineComment)) = "" Then
                                tmReRate(ilReRate).sLineComment = mGetcxfComment(tmClfP.lCxfCode) 'Purchased Comment
                            End If
                            tmReRate(ilReRate).sPriceType = tmCff.sPriceType
                            llPurchaseLineAud = llTotalAvgAud
                            blReRateAdded = True
                        'End If
                        'Get aired
                        ilSpotType = 1
                        ilTotLnSpts = 0                 'init total # spots per line
                         'Find matching line
                        ilClfR = -1
                        ilTotLnSpts = 0
                        For ilFind = 0 To UBound(tmClfReRate) - 1 Step 1
                            If tmClfP.iLine = tmClfReRate(ilFind).ClfRec.iLine Then
                                ilClfR = ilFind
                                tmClfR = tmClfReRate(ilClfR).ClfRec
                                'mGetAiredSpotCount ilClfR, llStartDate, llEndDate, ilMnfDemo, ilTotLnSpts, llTotalCntrSpots, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode
                                Exit For
                            End If
                        Next ilFind
                        'Get Research Book
                        If ilClfR <> -1 Then
                            imByCntrLnDnfCode = 0
                            'If rbcReRateBook(3).Value = True Then   'Contract Line
                            If bmBookByLine Then
                                'Contract Line:
                                'imReRateDnfCode = tmClfR.iDnfCode
                                'Look-up in tgBookByLineAssigned
                                For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                                    If tgBookByLineAssigned(ilBook).lChfCode = tmClfR.lChfCode Then
                                        If tgBookByLineAssigned(ilBook).iLineNo = tmClfR.iLine And tgBookByLineAssigned(ilBook).sType = tmClfR.sType Then
                                            imReRateDnfCode = mGetDnfByContractLine(ilClf, -1)
                                            Exit For
                                        End If
                                    End If
                                Next ilBook
                            ElseIf rbcReRateBook(1).Value = True Then   'Closest
                                'Get the closest
                                imReRateDnfCode = 0
                            ElseIf rbcReRateBook(2).Value = True Then   'None
                                'None
                                imReRateDnfCode = tmClfR.iDnfCode   'Use the line book so that unit totals will be computed and values removed in the outpit was 0
                            Else
                                'Use the default vehicle book
                                ilVef = gBinarySearchVef(tmClfR.iVefCode)
                                If ilVef <> -1 Then
                                    imReRateDnfCode = tgMVef(ilVef).iDnfCode
                                Else
                                    imReRateDnfCode = 0
                                End If
                            End If
                            'Note: DnfCode = 0 results in llPop = 0
                            'Population moved to mGetAiredSpotCount
                            'ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, imReRateDnfCode, 0, ilMnfDemo, llPop)
                            ReDim lmWklyspots(0 To ilNoWks) As Long       'sched lines weekly # spots
                            ReDim lmWklyAvgAud(0 To ilNoWks) As Long             'sched lines weekly avg aud
                            ReDim lmWklyRates(0 To ilNoWks) As Long           'sched lines weekly rates
                            ReDim lmWklyPopEst(0 To ilNoWks) As Long
                            ReDim lmWklyMoDate(0 To ilNoWks) As Long
                            mGetAiredSpotCount ilClfR, llStartDate, llEndDate, imMnfDemo, ilTotLnSpts, llTotalCntrSpots, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode
                            'If ilTotLnSpts > 0 And lmReRatePop > 0 Then 'And llPop > 0 Then
                                tmClfR = tmClfReRate(ilClfR).ClfRec
                                'gAvgAudToLnResearch sm1or2PlaceRating, False, lmReRatePop, lmWklyPopEst(), lmWklyspots(), lmWklyRates(), lmWklyAvgAud(), llTotalCost, llTotalAvgAud, imRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                                gAvgAudToLnResearch sm1or2PlaceRating, False, lmReRatePop, lmWklyPopEst(), lmWklyspots(), lmWklyRates(), lmWklyAvgAud(), dlTotalCost, llTotalAvgAud, imRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                                'Build totals by line
                                ilReRate = UBound(tmReRate)
                                If Not blReRateAdded Then
                                    tmReRate(ilReRate).lChfCode(0) = tmClfP.lChfCode
                                    tmReRate(ilReRate).lClfCode(0) = tmClfP.lCode
                                    tmReRate(ilReRate).iVefCode = tmClfP.iVefCode
                                    tmReRate(ilReRate).iRdfCode = tmClfP.iRdfCode
                                    tmReRate(ilReRate).sType = tmClfP.sType
                                    tmReRate(ilReRate).sSubType = ""
                                    tmReRate(ilReRate).lCntrNo = tmChfPurchase.lCntrNo
                                    tmReRate(ilReRate).sProduct = tmChfPurchase.sProduct
                                    tmReRate(ilReRate).iLineNo = tmClfP.iLine
                                    tmReRate(ilReRate).iPkLineNo = tmClfP.iPkLineNo
                                    tmReRate(ilReRate).iLen = tmClfP.iLen
                                    tmReRate(ilReRate).sAudioType = tmClfR.sLiveCopy 'TTP 10144
                                    tmReRate(ilReRate).sACT1LineupCode = tmClfR.sACT1LineupCode
                                    '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                                    tmReRate(ilReRate).sACT1StoredTime = tmClfR.sACT1StoredTime
                                    tmReRate(ilReRate).sACT1StoredSpots = tmClfR.sACT1StoredSpots
                                    tmReRate(ilReRate).sACT1StoreClearPct = tmClfR.sACT1StoreClearPct
                                    tmReRate(ilReRate).sACT1DaypartFilter = tmClfR.sACT1DaypartFilter
                                End If
                                If ckcMG.Value = vbChecked Or ckcTreatMGOsAsOrdered.Value = vbChecked Or imMGDetailDnfCode = -1 Then
                                    tmReRate(ilReRate).iDnfCode(ilSpotType) = imCntrLnDetailDnfCode 'imReRateDnfCode
                                Else
                                    If imMGDetailDnfCode = imCntrLnDetailDnfCode Then   'imReRateDnfCode Then
                                        tmReRate(ilReRate).iDnfCode(ilSpotType) = imReRateDnfCode
                                    Else
                                        tmReRate(ilReRate).iDnfCode(ilSpotType) = -2
                                    End If
                                End If
                                tmReRate(ilReRate).lChfCode(ilSpotType) = tmClfR.lChfCode
                                tmReRate(ilReRate).lClfCode(ilSpotType) = tmClfR.lCode
                                tmReRate(ilReRate).lPop(ilSpotType) = lmReRatePop
                                tmReRate(ilReRate).lSatelliteEst(ilSpotType) = llPopEst          '6-4-04
                                'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                                tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                                tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                                tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                                tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                                tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                                tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                                tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                                tmReRate(ilReRate).lTotalSpots(ilSpotType) = ilTotLnSpts
                                gUnpackDateLong tmClfR.iStartDate(0), tmClfR.iStartDate(1), llLnStartDate
                                gUnpackDateLong tmClfR.iEndDate(0), tmClfR.iEndDate(1), llLnEndDate
                                If llLnEndDate < llLnStartDate Then
                                    tmReRate(ilReRate).sCBS(ilSpotType) = "Y"
                                Else
                                    tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                                End If
                                If Trim(gStripChr0(tmReRate(ilReRate).sLineComment)) = "" Then
                                    tmReRate(ilReRate).sLineComment = mGetcxfComment(tmClfR.lCxfCode) ' "ReRate Comment"
                                End If
                                tmReRate(ilReRate).sPriceType = tmCff.sPriceType
                                blReRateAdded = True
                            'End If
                            If blReRateAdded Then
                                ReDim Preserve tmReRate(0 To ilReRate + 1)
                            End If
                            'Get MG's
                            For ilVef = 0 To UBound(tmMGInfo) - 1 Step 1
                                ReDim lmWklyspots(0 To ilNoWks) As Long       'sched lines weekly # spots
                                ReDim lmWklyAvgAud(0 To ilNoWks) As Long             'sched lines weekly avg aud
                                ReDim lmWklyRates(0 To ilNoWks) As Long           'sched lines weekly rates
                                ReDim lmWklyPopEst(0 To ilNoWks) As Long
                                ReDim lmWklyMoDate(0 To ilNoWks) As Long
                                                        
                                ilSpotType = 1
                                ilTotLnSpts = 0                 'init total # spots per line
                                                
                                mGetMGSpotCount tmMGInfo(ilVef).iVefCode, tmMGInfo(ilVef).iDnfCode, tmMGInfo(ilVef).iRdfCode, ilClfR, llStartDate, llEndDate, imMnfDemo, ilTotLnSpts, llTotalCntrSpots, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode
                                'If ilTotLnSpts > 0 And lmReRatePop > 0 Then 'llPop > 0 Then
                                    'gAvgAudToLnResearch sm1or2PlaceRating, False, lmReRatePop, lmWklyPopEst(), lmWklyspots(), lmWklyRates(), lmWklyAvgAud(), llTotalCost, llTotalAvgAud, imRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                                    gAvgAudToLnResearch sm1or2PlaceRating, False, lmReRatePop, lmWklyPopEst(), lmWklyspots(), lmWklyRates(), lmWklyAvgAud(), dlTotalCost, llTotalAvgAud, imRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                                    'Build totals by line
                                    ilReRate = UBound(tmReRate)
                                    tmReRate(ilReRate).lChfCode(0) = tmClfR.lChfCode
                                    tmReRate(ilReRate).lChfCode(1) = tmClfR.lChfCode
                                    tmReRate(ilReRate).lClfCode(0) = tmClfR.lCode
                                    tmReRate(ilReRate).lClfCode(1) = tmClfR.lCode
                                    tmReRate(ilReRate).iVefCode = tmMGInfo(ilVef).iVefCode
                                    tmReRate(ilReRate).iRdfCode = tmMGInfo(ilVef).iRdfCode
                                    tmReRate(ilReRate).sType = tmClfR.sType
                                    tmReRate(ilReRate).sSubType = "M"
                                    tmReRate(ilReRate).lCntrNo = tmChfPurchase.lCntrNo
                                    tmReRate(ilReRate).sProduct = tmChfPurchase.sProduct
                                    tmReRate(ilReRate).iLineNo = tmClfR.iLine
                                    tmReRate(ilReRate).iPkLineNo = tmClfR.iPkLineNo
                                    tmReRate(ilReRate).iLen = tmClfR.iLen
                                    tmReRate(ilReRate).sAudioType = tmClfR.sLiveCopy 'TTP 10144
                                    tmReRate(ilReRate).sACT1LineupCode = tmClfP.sACT1LineupCode
                                    '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                                    tmReRate(ilReRate).sACT1StoredTime = tmClfP.sACT1StoredTime
                                    tmReRate(ilReRate).sACT1StoredSpots = tmClfP.sACT1StoredSpots
                                    tmReRate(ilReRate).sACT1StoreClearPct = tmClfP.sACT1StoreClearPct
                                    tmReRate(ilReRate).sACT1DaypartFilter = tmClfP.sACT1DaypartFilter
                                    tmReRate(ilReRate).iDnfCode(ilSpotType) = imMGDetailDnfCode 'tmMGInfo(ilVef).iDnfCode  'imReRateDnfCode
                                    tmReRate(ilReRate).lPop(ilSpotType) = lmReRatePop   'llPop
                                    tmReRate(ilReRate).lSatelliteEst(ilSpotType) = llPopEst          '6-4-04
                                    'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                                    tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                                    tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                                    tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                                    tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                                    tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                                    tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                                    tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                                    tmReRate(ilReRate).lTotalSpots(ilSpotType) = ilTotLnSpts
                                    tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                                    If Trim(gStripChr0(tmReRate(ilReRate).sLineComment)) = "" Then
                                        tmReRate(ilReRate).sLineComment = mGetcxfComment(tmClf.lCxfCode) '"Make Good Comment"
                                    End If
                                    tmReRate(ilReRate).sPriceType = tmCff.sPriceType
                                    'save the Ordered Audince so that Index can be computed
                                    If llPurchaseLineAud <> -1 Then
                                        tmReRate(ilReRate).lTotalAvgAud(0) = llPurchaseLineAud
                                    End If
                                    ReDim Preserve tmReRate(0 To ilReRate + 1)
                                'End If
                            
                            Next ilVef
                        Else
                            ilRet = ilRet
                        End If
                    Else
                        'this is a package, we will deal with these later
                    End If
                Next ilClfP                                    'get next line
                'all line totals completed, get total contract research values
                'If UBound(tmReRate) > 1 Then       'found at least 1 line with research totals
                llLnSpots = 0
                If UBound(tmReRate) > LBound(tmReRate) Then       'found at least 1 line with research totals
                    'Compute package #'s
                    For ilClfP = LBound(tmClfPurchase) To UBound(tmClfPurchase) - 1 Step 1
                        tmClfP = tmClfPurchase(ilClfP).ClfRec
                        If tmClfP.sType <> "H" And tmClfP.sType <> "S" Then
                            blReRateAdded = False
                            For ilSpotType = 0 To 1 Step 1
                                slCBS = "N"
                                ilPackageDnfCode = -1
                                If ilSpotType = 0 Then
                                    tlClf = tmClfP
                                Else
                                    ilClfR = -1
                                    For ilFind = 0 To UBound(tmClfReRate) - 1 Step 1
                                        If tmClfP.iLine = tmClfReRate(ilFind).ClfRec.iLine Then
                                            ilClfR = ilFind
                                            tlClf = tmClfReRate(ilFind).ClfRec
                                            gUnpackDateLong tlClf.iStartDate(0), tlClf.iStartDate(1), llLnStartDate
                                            gUnpackDateLong tlClf.iEndDate(0), tlClf.iEndDate(1), llLnEndDate
                                            If llLnEndDate < llLnStartDate Then
                                                slCBS = "Y"
                                            Else
                                                slCBS = "N"
                                            End If
                                            Exit For
                                        End If
                                    Next ilFind
                                End If
                                If slCBS <> "Y" Then
                                    llPop = -1
                                    ilHiddenLines = 0
                                    ReDim imRtg(0 To 0)                   'setup arrays for return values from audtolnresearch
                                    ReDim lmGrimp(0 To 0)
                                    ReDim lmGRP(0 To 0)
                                    ReDim lmCost(0 To 0)                    'setup arrays for return values from audtolnresearch
                                    For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                                        If tlClf.iLine = tmReRate(illoop).iPkLineNo Then
                                            If (ilSpotType = 0) Or ((ilSpotType = 1) And (tmReRate(illoop).sCBS(ilSpotType) <> "Y")) Then
                                                ReDim Preserve lmCost(0 To ilHiddenLines) As Long
                                                ReDim Preserve lmGrimp(0 To ilHiddenLines) As Long
                                                ReDim Preserve lmGRP(0 To ilHiddenLines) As Long
                                                'lmCost(ilHiddenLines) = tmReRate(illoop).lTotalCost(ilSpotType)
                                                lmCost(ilHiddenLines) = tmReRate(illoop).dTotalCost(ilSpotType) 'TTP 10439 - Rerate 21,000,000
                                                lmGrimp(ilHiddenLines) = tmReRate(illoop).lTotalGrimps(ilSpotType)
                                                lmGRP(ilHiddenLines) = tmReRate(illoop).lTotalGRP(ilSpotType)
                                                'determine if varying populations across the weeks (demo estimates) or across the lines
                                                If tmReRate(illoop).iDnfCode(ilSpotType) > 0 Then
                                                    If tgSpf.sDemoEstAllowed = "Y" Then             '6-4-04
                                                        If llPop < 0 Then
                                                            llPop = tmReRate(illoop).lSatelliteEst(ilSpotType)
                                                        ElseIf llPop <> tmReRate(illoop).lSatelliteEst(ilSpotType) And tmReRate(illoop).lSatelliteEst(ilSpotType) <> 0 Then
                                                            llPop = 0
                                                        End If
                                                    Else
                                                        If llPop < 0 Then
                                                            llPop = tmReRate(illoop).lPop(ilSpotType)
                                                        ElseIf llPop <> tmReRate(illoop).lPop(ilSpotType) And tmReRate(illoop).lPop(ilSpotType) <> 0 Then
                                                            llPop = 0
                                                        End If
                                                    End If
                                                End If
                                                If tmReRate(illoop).lTotalSpots(ilSpotType) > 0 Then
                                                    If tmReRate(illoop).iDnfCode(ilSpotType) > 0 Then
                                                        If ilSpotType = 0 Then
                                                            If tmReRate(illoop).sSubType <> "M" Then
                                                                If ilPackageDnfCode = -1 Then
                                                                   ilPackageDnfCode = tmReRate(illoop).iDnfCode(ilSpotType)
                                                                Else
                                                                   If (ilPackageDnfCode <> tmReRate(illoop).iDnfCode(ilSpotType)) And (ilPackageDnfCode <> -2) Then
                                                                       ilPackageDnfCode = -2
                                                                   End If
                                                                End If
                                                            End If
                                                        Else
                                                            If ilPackageDnfCode = -1 Then
                                                                ilPackageDnfCode = tmReRate(illoop).iDnfCode(ilSpotType)
                                                            Else
                                                                If (ilPackageDnfCode <> tmReRate(illoop).iDnfCode(ilSpotType)) And (ilPackageDnfCode <> -2) Then
                                                                    ilPackageDnfCode = -2
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                ilHiddenLines = ilHiddenLines + 1
                                            End If
                                        End If
                                    Next illoop
                                    'Get total package spots, cost his ok as each week of hidden lines must match the package line cost
                                    llLnSpots = 0
                                    If ilSpotType = 0 Then
                                        ilCff = tmClfPurchase(ilClfP).iFirstCff
                                    Else
                                        ilCff = tmClfReRate(ilClfR).iFirstCff
                                    End If
                                    Do While ilCff <> -1
                                        If ilSpotType = 0 Then
                                            tmCff = tmCffPurchase(ilCff).CffRec
                                        Else
                                            tmCff = tmCffReRate(ilCff).CffRec
                                        End If
                                        For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                                            ilInputDays(illoop) = False
                                        Next illoop
                
                                        gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                                        llFltStart = gDateValue(slStr)
                                        'backup start date to Monday
                                        illoop = gWeekDayLong(llFltStart)
                                        Do While illoop <> 0
                                            llFltStart = llFltStart - 1
                                            illoop = gWeekDayLong(llFltStart)
                                        Loop
                                        gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                                        llFltEnd = gDateValue(slStr)
                                        '
                                        'Loop thru the flight by week and build the number of spots for each week
                                        '
                                        For llDate2 = llFltStart To llFltEnd Step 7
                                            If llDate2 >= llStartDate And llDate2 <= llEndDate Then
                                                If tmCff.sDyWk = "W" Then            'weekly
                                                    ilSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                                                Else                                        'daily
                                                     If illoop + 6 < llFltEnd Then           'we have a whole week
                                                        ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                                                     Else                                    'do partial week
                                                        For llDate = llDate2 To llFltEnd Step 1
                                                            ilDay = gWeekDayLong(llDate)
                                                            ilSpots = ilSpots + tmCff.iDay(ilDay)
                                                        Next llDate
                                                    End If
                                                End If
                                                llLnSpots = llLnSpots + ilSpots
                                            End If                  'if llDate2 >= llStartDate and llDate2 <= llEndDAte
                                        Next llDate2
                                        If ilSpotType = 0 Then
                                            ilCff = tmCffPurchase(ilCff).iNextCff               'get next flight record from mem
                                        Else
                                            ilCff = tmCffReRate(ilCff).iNextCff               'get next flight record from mem
                                        End If
                                    Loop                                            'while ilcff <> -1
                                    'Obtain package research totals
                                    'gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llLnSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
                                    gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llLnSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
                                    'If llTotalCPP <> 0 Or llTotalCPM <> 0 Then
                                        'maintain research contract totals for all contracts to be able to total them by demo
                                        'include all contracts types (with remnants, DR, & PI)
                                        ilReRate = UBound(tmReRate)
                                        tmReRate(ilReRate).lChfCode(0) = tlClf.lChfCode
                                        tmReRate(ilReRate).lChfCode(1) = tlClf.lChfCode
                                        tmReRate(ilReRate).lClfCode(0) = tlClf.lCode
                                        tmReRate(ilReRate).lClfCode(1) = tlClf.lCode
                                        tmReRate(ilReRate).iVefCode = tlClf.iVefCode
                                        tmReRate(ilReRate).iRdfCode = tlClf.iRdfCode
                                        tmReRate(ilReRate).sType = tlClf.sType
                                        tmReRate(ilReRate).sSubType = ""
                                        tmReRate(ilReRate).lCntrNo = tmChfPurchase.lCntrNo
                                        tmReRate(ilReRate).sProduct = tmChfPurchase.sProduct
                                        tmReRate(ilReRate).iLineNo = tlClf.iLine
                                        tmReRate(ilReRate).iPkLineNo = 0
                                        tmReRate(ilReRate).iLen = tlClf.iLen
                                        tmReRate(ilReRate).sAudioType = tlClf.sLiveCopy 'TTP 10144
                                        tmReRate(ilReRate).sACT1LineupCode = tlClf.sACT1LineupCode
                                        '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                                        tmReRate(ilReRate).sACT1StoredTime = tlClf.sACT1StoredTime
                                        tmReRate(ilReRate).sACT1StoredSpots = tlClf.sACT1StoredSpots
                                        tmReRate(ilReRate).sACT1StoreClearPct = tlClf.sACT1StoreClearPct
                                        tmReRate(ilReRate).sACT1DaypartFilter = tlClf.sACT1DaypartFilter
                                        tmReRate(ilReRate).iDnfCode(ilSpotType) = ilPackageDnfCode  'tlClf.iDnfCode
                                        tmReRate(ilReRate).lPop(ilSpotType) = llPop
                                        tmReRate(ilReRate).lSatelliteEst(ilSpotType) = llPop            '6-4-04
                                        'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                                        tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                                        tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                                        tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                                        tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                                        tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                                        tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                                        tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                                        tmReRate(ilReRate).lTotalSpots(ilSpotType) = llLnSpots
                                        If Trim(gStripChr0(tmReRate(ilReRate).sLineComment)) = "" Then
                                            tmReRate(ilReRate).sLineComment = mGetcxfComment(tlClf.lCxfCode) ' "ReRate Package Comment"
                                        End If
                                        tmReRate(ilReRate).sPriceType = tmCff.sPriceType
                                        gUnpackDateLong tlClf.iStartDate(0), tlClf.iStartDate(1), llLnStartDate
                                        gUnpackDateLong tlClf.iEndDate(0), tlClf.iEndDate(1), llLnEndDate
                                        If llLnEndDate < llLnStartDate Then
                                            tmReRate(ilReRate).sCBS(ilSpotType) = "Y"
                                        Else
                                            tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                                        End If
                                        blReRateAdded = True
                                        'ReDim Preserve tmReRate(0 To ilReRate + 1)
                                    'End If                  'cpp or cpm are zero
                                End If
                            Next ilSpotType
                            If blReRateAdded Then
                                ReDim Preserve tmReRate(0 To ilReRate + 1)
                            End If
                        End If
                    Next ilClfP
                    'Build total for contract and advertiser
                    blReRateAdded = False
                    For ilSpotType = 0 To 1 Step 1
                        llPop = -1
                        ilCntrTotalDnfCode = -1
                        llTotalCntrSpots = 0
                        ilRow = 0
                        ReDim imRtg(0 To 0)                   'setup arrays for return values from audtolnresearch
                        ReDim lmGrimp(0 To 0)
                        ReDim lmGRP(0 To 0)
                        ReDim lmCost(0 To 0)                    'setup arrays for return values from audtolnresearch
                        For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                            If (ilSpotType = 0) Or ((ilSpotType = 1) And (tmReRate(illoop).sCBS(ilSpotType) <> "Y")) Then
                                If tmReRate(illoop).sType = "S" Or tmReRate(illoop).sType = "H" Then
                                    ReDim Preserve lmCost(0 To ilRow) As Long
                                    ReDim Preserve lmGrimp(0 To ilRow) As Long
                                    ReDim Preserve lmGRP(0 To ilRow) As Long
                                    'lmCost(ilRow) = tmReRate(illoop).lTotalCost(ilSpotType)
                                    lmCost(ilRow) = tmReRate(illoop).dTotalCost(ilSpotType) 'TTP 10439 - Rerate 21,000,000
                                    lmGrimp(ilRow) = tmReRate(illoop).lTotalGrimps(ilSpotType)
                                    lmGRP(ilRow) = tmReRate(illoop).lTotalGRP(ilSpotType)
                                    ilRow = ilRow + 1
                                    If ilSpotType = 0 Then
                                        ReDim Preserve lmAdvtOrderCost(0 To ilAdvtOrderRow) As Long
                                        ReDim Preserve lmAdvtOrderGrimp(0 To ilAdvtOrderRow) As Long
                                        ReDim Preserve lmAdvtOrderGRP(0 To ilAdvtOrderRow) As Long
                                        'lmAdvtOrderCost(ilAdvtOrderRow) = tmReRate(illoop).lTotalCost(ilSpotType)
                                        lmAdvtOrderCost(ilAdvtOrderRow) = tmReRate(illoop).dTotalCost(ilSpotType) 'TTP 10439 - Rerate 21,000,000
                                        lmAdvtOrderGrimp(ilAdvtOrderRow) = tmReRate(illoop).lTotalGrimps(ilSpotType)
                                        lmAdvtOrderGRP(ilAdvtOrderRow) = tmReRate(illoop).lTotalGRP(ilSpotType)
                                        ilAdvtOrderRow = ilAdvtOrderRow + 1
                                    Else
                                        ReDim Preserve lmAdvtReRateCost(0 To ilAdvtReRateRow) As Long
                                        ReDim Preserve lmAdvtReRateGrimp(0 To ilAdvtReRateRow) As Long
                                        ReDim Preserve lmAdvtReRateGRP(0 To ilAdvtReRateRow) As Long
                                        'lmAdvtReRateCost(ilAdvtReRateRow) = tmReRate(illoop).lTotalCost(ilSpotType)
                                        lmAdvtReRateCost(ilAdvtReRateRow) = tmReRate(illoop).dTotalCost(ilSpotType) 'TTP 10439 - Rerate 21,000,000
                                        lmAdvtReRateGrimp(ilAdvtReRateRow) = tmReRate(illoop).lTotalGrimps(ilSpotType)
                                        lmAdvtReRateGRP(ilAdvtReRateRow) = tmReRate(illoop).lTotalGRP(ilSpotType)
                                        ilAdvtReRateRow = ilAdvtReRateRow + 1
                                    End If
                                    
                                    If tmReRate(illoop).sType = "S" Then
                                        llTotalCntrSpots = llTotalCntrSpots + tmReRate(illoop).lTotalSpots(ilSpotType)
                                    End If
                                    'determine if varying populations across the weeks (demo estimates) or across the lines
                                    If (tmReRate(illoop).iDnfCode(ilSpotType) > 0) Or (tmReRate(illoop).iDnfCode(ilSpotType) = -2) Then
                                        If tgSpf.sDemoEstAllowed = "Y" Then             '6-4-04
                                            If llPop < 0 Then
                                                llPop = tmReRate(illoop).lSatelliteEst(ilSpotType)
                                            ElseIf llPop <> tmReRate(illoop).lSatelliteEst(ilSpotType) And tmReRate(illoop).lSatelliteEst(ilSpotType) <> 0 Then
                                                llPop = 0
                                            End If
                    
                                        Else
                                            If llPop < 0 Then
                                                llPop = tmReRate(illoop).lPop(ilSpotType)
                                            ElseIf llPop <> tmReRate(illoop).lPop(ilSpotType) And tmReRate(illoop).lPop(ilSpotType) <> 0 Then
                                                llPop = 0
                                            End If
                                        End If
                                    End If
                                    If (tmReRate(illoop).lTotalSpots(ilSpotType) > 0) And (tmReRate(illoop).iDnfCode(ilSpotType) > 0) Then
                                        If ilSpotType = 0 Then
                                            If tmReRate(illoop).sSubType <> "M" Then
                                                If ilCntrTotalDnfCode = -1 Then
                                                   ilCntrTotalDnfCode = tmReRate(illoop).iDnfCode(ilSpotType)
                                                Else
                                                   If (ilCntrTotalDnfCode <> tmReRate(illoop).iDnfCode(ilSpotType)) And (ilCntrTotalDnfCode <> -2) Then
                                                       ilCntrTotalDnfCode = -2
                                                   End If
                                                End If
                                            End If
                                        Else
                                            If ilCntrTotalDnfCode = -1 Then
                                                ilCntrTotalDnfCode = tmReRate(illoop).iDnfCode(ilSpotType)
                                            Else
                                                If (ilCntrTotalDnfCode <> tmReRate(illoop).iDnfCode(ilSpotType)) And (ilCntrTotalDnfCode <> -2) Then
                                                    ilCntrTotalDnfCode = -2
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    llTotalCntrSpots = llTotalCntrSpots + tmReRate(illoop).lTotalSpots(ilSpotType)
                                End If
                            End If
                        Next illoop
                        If ilSpotType = 0 Then
                            If lmAdvtOrderPop < 0 Then
                                lmAdvtOrderPop = llPop
                            ElseIf lmAdvtOrderPop <> llPop And llPop <> 0 Then
                                lmAdvtOrderPop = 0
                            End If
                            lmAdvtOrderTotalSpots = lmAdvtOrderTotalSpots + llTotalCntrSpots
                            If imAdvtOrderDnfCode = -1 Then
                                imAdvtOrderDnfCode = ilCntrTotalDnfCode
                            Else
                                If (ilCntrTotalDnfCode <> imAdvtOrderDnfCode) And (imAdvtOrderDnfCode <> -2) Then
                                    imAdvtOrderDnfCode = -2
                                End If
                            End If
                        Else
                            If lmAdvtReRatePop < 0 Then
                                lmAdvtReRatePop = llPop
                            ElseIf lmAdvtReRatePop <> llPop And llPop <> 0 Then
                                lmAdvtReRatePop = 0
                            End If
                            lmAdvtReRateTotalSpots = lmAdvtReRateTotalSpots + llTotalCntrSpots
                            If imAdvrReRateDnfCode = -1 Then
                                imAdvrReRateDnfCode = ilCntrTotalDnfCode
                            Else
                                If (ilCntrTotalDnfCode <> imAdvrReRateDnfCode) And (imAdvrReRateDnfCode <> -2) Then
                                    imAdvrReRateDnfCode = -2
                                End If
                            End If
                        End If
                        'gResearchTotals False, llPop, lmCost(), imRtg(), lmGrimp(), lmGRP(), llTotalCost, ilTotalAvgRtg, llTotalGrimp, llTotalGRP, llTotalCPP, llTotalCPM
                        '4/7/99
                        'Obtain contract research totals
                        'gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llTotalCntrSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
                        gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llTotalCntrSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
                        'If llTotalCPP <> 0 Or llTotalCPM <> 0 Then   '10/13/20 - TTP 9953 - Commented Out due to Zero Dollar contracts would retain last Contract Totals (tlContractOrderReRate)
                            'maintain research contract totals for all contracts to be able to total them by demo
                            'include all contracts types (with remnants, DR, & PI)
                            ilReRate = UBound(tmReRate)
                            tmReRate(ilReRate).lChfCode(0) = tmChfPurchase.lCode
                            tmReRate(ilReRate).lChfCode(1) = tmChfReRate.lCode
                            tmReRate(ilReRate).lClfCode(0) = 0
                            tmReRate(ilReRate).lClfCode(1) = 0
                            tmReRate(ilReRate).iVefCode = 0
                            tmReRate(ilReRate).iRdfCode = 0
                            tmReRate(ilReRate).sType = "T"
                            tmReRate(ilReRate).sSubType = "C"
                            tmReRate(ilReRate).lCntrNo = tmChfPurchase.lCntrNo
                            tmReRate(ilReRate).sProduct = tmChfPurchase.sProduct
                            tmReRate(ilReRate).iLineNo = 0
                            tmReRate(ilReRate).iPkLineNo = 0
                            tmReRate(ilReRate).iLen = 0
                            tmReRate(ilReRate).sAudioType = ""
                            tmReRate(ilReRate).sACT1LineupCode = ""
                            '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                            tmReRate(ilReRate).sACT1StoredTime = ""
                            tmReRate(ilReRate).sACT1StoredSpots = ""
                            tmReRate(ilReRate).sACT1StoreClearPct = ""
                            tmReRate(ilReRate).sACT1DaypartFilter = ""
                            tmReRate(ilReRate).iDnfCode(ilSpotType) = ilCntrTotalDnfCode
                            tmReRate(ilReRate).lPop(ilSpotType) = llPop
                            tmReRate(ilReRate).lSatelliteEst(ilSpotType) = llPop            '6-4-04
                            'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                            tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                            tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                            tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                            tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                            tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                            tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                            tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                            tmReRate(ilReRate).lTotalSpots(ilSpotType) = llTotalCntrSpots
                            tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                            tmReRate(ilReRate).sLineComment = "" ' "Line Comment research totals"
                            tmReRate(ilReRate).sPriceType = ""
                            If ilSpotType = 0 Then
                                tlContractOrderReRate = tmReRate(ilReRate)
                            End If
                            blReRateAdded = True
                        'End If  '10/13/20 - TTP 9953 - Commented Out
                    Next ilSpotType
                    If blReRateAdded Then
                        ReDim Preserve tmReRate(0 To ilReRate + 1)
                    End If
                End If                      'found at least 1 line with research totals
            End If                          'ilfoundcnt = true
            'Export ReRate Contract Information
            'Output Lines in Line Number Order except hidden lines within package lines
            'Create array of line numbers for std and packages
'            ReDim ilLineNo(0 To 0) As Integer
'            For ilLoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
'                If tmReRate(ilLoop).sType <> "H" Then
'                    ilLineNo(UBound(ilLineNo)) = tmReRate(ilLoop).iLineNo
'                    ReDim Preserve ilLineNo(0 To UBound(ilLineNo) + 1) As Integer
'                End If
'            Next ilLoop
'            If UBound(ilLineNo) > 1 Then
'                ArraySortTyp fnAV(ilLineNo(), 0), UBound(ilLineNo) - 1, 0, LenB(ilLineNo(0)), 0, -1, 0
'            End If
            For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                tmReRate(illoop).lSeqNo = 0
            Next illoop
            llSeqNo = 1
            For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                If tmReRate(illoop).sType <> "B" Then
                    If (tmReRate(illoop).sType <> "S") And (tmReRate(illoop).sType <> "H") And (tmReRate(illoop).lSeqNo = 0) Then
                        tmReRate(illoop).lSeqNo = llSeqNo
                        llSeqNo = llSeqNo + 1
                        If (tmReRate(illoop).sType <> "T") Then
                            For ilRow = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                                If (tmReRate(ilRow).iPkLineNo = tmReRate(illoop).iLineNo) And (tmReRate(ilRow).lCntrNo = tmReRate(illoop).lCntrNo) Then
                                    tmReRate(ilRow).lSeqNo = llSeqNo
                                    llSeqNo = llSeqNo + 1
                                End If
                            Next ilRow
                            For ilRow = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                                If (tmReRate(ilRow).sType = "T") And (tmReRate(ilRow).lCntrNo = tmReRate(illoop).lCntrNo) Then
                                    tmReRate(ilRow).lSeqNo = llSeqNo
                                    llSeqNo = llSeqNo + 1
                                End If
                            Next ilRow
                        End If
                    ElseIf tmReRate(illoop).sType = "S" Then
                        tmReRate(illoop).lSeqNo = llSeqNo
                        llSeqNo = llSeqNo + 1
                    End If
                End If
            Next illoop
            
            'Get Bonus
            'For ilVef = 0 To UBound(imBonusVefCode) - 1 Step 1
            imBonusTotalDnfCode = -1
            For ilVef = 0 To UBound(tmBonusInfo) - 1 Step 1
            
                ReDim lmWklyspots(0 To 0) As Long       'sched lines weekly # spots
                ReDim lmWklyAvgAud(0 To 0) As Long             'sched lines weekly avg aud
                ReDim lmWklyRates(0 To 0) As Long           'sched lines weekly rates
                ReDim lmWklyPopEst(0 To 0) As Long
                ReDim lmWklyMoDate(0 To 0) As Long
                ReDim imRtg(0 To 0) As Integer                   'setup arrays for return values from audtolnresearch
                ReDim lmGrimp(0 To 0) As Long
                ReDim lmGRP(0 To 0) As Long
                ReDim lmCost(0 To 0) As Long                    'setup arrays for return values from audtolnresearch
                
                ilSpotType = 1
                ilTotLnSpts = 0                 'init total # spots per line
                                        
                'If rbcReRateBook(3).Value = True Then
                '    imReRateDnfCode = tmBonusInfo(ilVef).iDnfCode
                'Else
                '    ilVefIndex = gBinarySearchVef(tmBonusInfo(ilVef).iVefCode) 'imBonusVefCode(ilVef))
                '    If ilVefIndex <> -1 Then
                '        imReRateDnfCode = tgMVef(ilVefIndex).iDnfCode
                '    Else
                '    End If
                'End If
                'imReRateDnfCode = tmBonusInfo(ilVef).iDnfCode
                'Using first aired spot to determine Population (lmReRatePop)
                'ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, imReRateDnfCode, 0, ilMnfDemo, llPop)
                                        
                'mGetBonusSpotCount imBonusVefCode(ilVef), llStartDate, llEndDate, ilMnfDemo, ilTotLnSpts, llTotalCntrSpots, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode
                mGetBonusSpotCount tmBonusInfo(ilVef).iVefCode, tmBonusInfo(ilVef).iDnfCode, tmBonusInfo(ilVef).iRdfCode, llStartDate, llEndDate, imMnfDemo, ilTotLnSpts, llTotalCntrSpots, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode, tmBonusInfo(ilVef).iSpotLen
                ReDim Preserve imRtg(0 To UBound(lmWklyspots)) As Integer
                ReDim Preserve lmCost(0 To UBound(lmWklyspots)) As Long
                ReDim Preserve lmGrimp(0 To UBound(lmWklyspots)) As Long
                ReDim Preserve lmGRP(0 To UBound(lmWklyspots)) As Long

                If ilTotLnSpts > 0 And lmReRatePop > 0 Then 'llPop > 0 Then
                    If imBonusTotalDnfCode = -1 Then
                        imBonusTotalDnfCode = imBonusDetailDnfCode
                    ElseIf imBonusTotalDnfCode <> imBonusDetailDnfCode And imBonusTotalDnfCode <> -2 Then
                        imBonusTotalDnfCode = -2
                    End If
                    'gAvgAudToLnResearch sm1or2PlaceRating, False, lmReRatePop, lmWklyPopEst(), lmWklyspots(), lmWklyRates(), lmWklyAvgAud(), llTotalCost, llTotalAvgAud, imRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                    gAvgAudToLnResearch sm1or2PlaceRating, False, lmReRatePop, lmWklyPopEst(), lmWklyspots(), lmWklyRates(), lmWklyAvgAud(), dlTotalCost, llTotalAvgAud, imRtg(), ilTotalAvgRtg, lmGrimp(), llTotalGrImp, lmGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                    'Build totals by line
                    ilReRate = UBound(tmReRate)
                    tmReRate(ilReRate).lChfCode(0) = tmChfPurchase.lCode
                    tmReRate(ilReRate).lChfCode(1) = tmChfReRate.lCode
                    tmReRate(ilReRate).lClfCode(0) = 0 'tmClfPurchase.lCode
                    tmReRate(ilReRate).lClfCode(1) = 0
                    tmReRate(ilReRate).iVefCode = tmBonusInfo(ilVef).iVefCode   'imBonusVefCode(ilVef)
                    tmReRate(ilReRate).iRdfCode = 0 'tmClfPurchase.iRdfCode
                    If rbcBonus(1).Value Then
                        tmReRate(ilReRate).iRdfCode = tmBonusInfo(ilVef).iRdfCode
                    End If
                    tmReRate(ilReRate).sType = "B"  'tmClfPurchase.sType
                    tmReRate(ilReRate).sSubType = "B"
                    tmReRate(ilReRate).lCntrNo = tmChfPurchase.lCntrNo
                    tmReRate(ilReRate).sProduct = tmChfPurchase.sProduct
                    tmReRate(ilReRate).iLineNo = 0  'tmClfPurchase.iLine
                    tmReRate(ilReRate).iPkLineNo = 0    'tmClfPurchase.iPkLineNo
                    tmReRate(ilReRate).iLen = tmBonusInfo(ilVef).iSpotLen 'TTP 10123
                    tmReRate(ilReRate).sAudioType = tmBonusInfo(ilVef).sAudioTypes 'TTP 10144
                    tmReRate(ilReRate).sACT1LineupCode = ""
                    '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                    tmReRate(ilReRate).sACT1StoredTime = ""
                    tmReRate(ilReRate).sACT1StoredSpots = ""
                    tmReRate(ilReRate).sACT1StoreClearPct = ""
                    tmReRate(ilReRate).sACT1DaypartFilter = ""
                    tmReRate(ilReRate).iDnfCode(ilSpotType) = imBonusDetailDnfCode  'imReRateDnfCode
                    tmReRate(ilReRate).lPop(ilSpotType) = lmReRatePop
                    tmReRate(ilReRate).lSatelliteEst(ilSpotType) = llPopEst          '6-4-04
                    'tmReRate(ilReRate).lTotalCost(ilSpotType) = 0
                    tmReRate(ilReRate).dTotalCost(ilSpotType) = 0 'TTP 10439 - Rerate 21,000,000
                    tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                    tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                    tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                    tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                    tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                    tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                    tmReRate(ilReRate).lTotalSpots(ilSpotType) = ilTotLnSpts
                    tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                    If Trim(gStripChr0(tmReRate(ilReRate).sLineComment)) = "" Then
                        tmReRate(ilReRate).sLineComment = mGetcxfComment(tmClfR.lCxfCode)  '"Bonus Comment"
                    End If
                    tmReRate(ilReRate).sPriceType = tmCff.sPriceType
                    ReDim Preserve tmReRate(0 To ilReRate + 1)
                End If
            
            Next ilVef
            
            If UBound(tmReRate) > LBound(tmReRate) Then       'found at least 1 line with research totals
                ilRow = 0
                llPop = -1
                ilSpotType = 1
                llLnSpots = 0
                ReDim imRtg(0 To 0)                   'setup arrays for return values from audtolnresearch
                ReDim lmGrimp(0 To 0)
                ReDim lmGRP(0 To 0)
                ReDim lmCost(0 To 0)                    'setup arrays for return values from audtolnresearch
                ilBonusDnfCode = -1
                For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                    If tmReRate(illoop).sType = "B" Then
                        ReDim Preserve lmCost(0 To ilRow) As Long
                        ReDim Preserve lmGrimp(0 To ilRow) As Long
                        ReDim Preserve lmGRP(0 To ilRow) As Long
                        'lmCost(ilRow) = tmReRate(illoop).lTotalCost(ilSpotType)
                        lmCost(ilRow) = tmReRate(illoop).dTotalCost(ilSpotType) 'TTP 10439 - Rerate 21,000,000
                        lmGrimp(ilRow) = tmReRate(illoop).lTotalGrimps(ilSpotType)
                        lmGRP(ilRow) = tmReRate(illoop).lTotalGRP(ilSpotType)
                        llLnSpots = llLnSpots + tmReRate(illoop).lTotalSpots(ilSpotType)
                        'determine if varying populations across the weeks (demo estimates) or across the lines
                        If tmReRate(illoop).iDnfCode(ilSpotType) > 0 Then
                            If tgSpf.sDemoEstAllowed = "Y" Then             '6-4-04
                                If llPop < 0 Then
                                    llPop = tmReRate(illoop).lSatelliteEst(ilSpotType)
                                ElseIf llPop <> tmReRate(illoop).lSatelliteEst(ilSpotType) And tmReRate(illoop).lSatelliteEst(ilSpotType) <> 0 Then
                                    llPop = 0
                                End If
                            Else
                                If llPop < 0 Then
                                    llPop = tmReRate(illoop).lPop(ilSpotType)
                                ElseIf llPop <> tmReRate(illoop).lPop(ilSpotType) And tmReRate(illoop).lPop(ilSpotType) <> 0 Then
                                    llPop = 0
                                End If
                            End If
                        End If
                        If (tmReRate(illoop).lTotalSpots(ilSpotType) > 0) And (tmReRate(illoop).iDnfCode(ilSpotType) > 0) Then
                            If ilBonusDnfCode = -1 Then
                                ilBonusDnfCode = tmReRate(illoop).iDnfCode(ilSpotType)
                            Else
                                If (ilBonusDnfCode <> tmReRate(illoop).iDnfCode(ilSpotType)) And (ilBonusDnfCode <> -2) Then
                                    ilBonusDnfCode = -2
                                End If
                            End If
                        End If
                        
                        ilRow = ilRow + 1
                    End If
                Next illoop
                'Obtain Bonus totals
                'gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llLnSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
                gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llLnSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
                If llLnSpots > 0 Then
                    ilReRate = UBound(tmReRate)
                    tmReRate(ilReRate).lChfCode(0) = tmChfPurchase.lCode
                    tmReRate(ilReRate).lChfCode(1) = tmChfReRate.lCode
                    tmReRate(ilReRate).lClfCode(0) = 0
                    tmReRate(ilReRate).lClfCode(1) = 0
                    tmReRate(ilReRate).iVefCode = 0
                    tmReRate(ilReRate).iRdfCode = 0
                    tmReRate(ilReRate).sType = "B"      'Bonus total
                    tmReRate(ilReRate).sSubType = "T"
                    tmReRate(ilReRate).lCntrNo = tmChfPurchase.lCntrNo
                    tmReRate(ilReRate).sProduct = tmChfPurchase.sProduct
                    tmReRate(ilReRate).iLineNo = 0
                    tmReRate(ilReRate).iPkLineNo = 0
                    tmReRate(ilReRate).iLen = 0
                    tmReRate(ilReRate).sAudioType = ""
                    tmReRate(ilReRate).sACT1LineupCode = ""
                    '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                    tmReRate(ilReRate).sACT1StoredTime = ""
                    tmReRate(ilReRate).sACT1StoredSpots = ""
                    tmReRate(ilReRate).sACT1StoreClearPct = ""
                    tmReRate(ilReRate).sACT1DaypartFilter = ""
                    tmReRate(ilReRate).iDnfCode(ilSpotType) = imBonusTotalDnfCode   'ilBonusDnfCode    'imReRateDnfCode
                    tmReRate(ilReRate).lPop(ilSpotType) = llPop
                    tmReRate(ilReRate).lSatelliteEst(ilSpotType) = llPop            '6-4-04
                    'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                    tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                    tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                    tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                    tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                    tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                    tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                    tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                    tmReRate(ilReRate).lTotalSpots(ilSpotType) = llLnSpots
                    tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                    tmReRate(ilReRate).sLineComment = "" '"Bonus total Comment"
                    tmReRate(ilReRate).sPriceType = "" 'tmCff.sPriceType
                    blReRateAdded = True
                    ReDim Preserve tmReRate(0 To ilReRate + 1)

                End If
            End If
            
            'Sort bonus after Contract total
            For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                If tmReRate(illoop).sType = "B" Then
                    tmReRate(illoop).lSeqNo = llSeqNo
                    llSeqNo = llSeqNo + 1
                End If
            Next illoop
            If UBound(tmReRate) > 1 Then
                ArraySortTyp fnAV(tmReRate(), 0), UBound(tmReRate), 0, LenB(tmReRate(0)), 0, -2, 0
            End If
            mOutputCntrHeader llStartDate, llEndDate
            For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                'Generate output
                mOutputResearch tmReRate(illoop)
            Next illoop
            
            'Compute contract total plus Bonus
            llPop = -1
            ilRow = 0
            ilSpotType = 1
            llLnSpots = 0
            If lmAdvtPlusBonusReRatePop = -1 Then
                lmAdvtPlusBonusReRatePop = lmAdvtReRatePop
            Else
                If (lmAdvtPlusBonusReRatePop <> lmAdvtReRatePop) And (lmAdvtPlusBonusReRatePop <> 0) Then
                    lmAdvtPlusBonusReRatePop = 0
                End If
            End If
            ReDim lmCost(0 To 0) As Long
            ReDim lmGrimp(0 To 0) As Long
            ReDim lmGRP(0 To 0) As Long
            For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                If ((tmReRate(illoop).sType = "B") And (tmReRate(illoop).sSubType = "B")) Or ((tmReRate(illoop).sType = "T") And (tmReRate(illoop).sSubType = "C")) Then
                    ReDim Preserve lmCost(0 To ilRow) As Long
                    ReDim Preserve lmGrimp(0 To ilRow) As Long
                    ReDim Preserve lmGRP(0 To ilRow) As Long
                    'lmCost(ilRow) = tmReRate(illoop).lTotalCost(ilSpotType)
                    lmCost(ilRow) = tmReRate(illoop).dTotalCost(ilSpotType) 'TTP 10439 - Rerate 21,000,000
                    lmGrimp(ilRow) = tmReRate(illoop).lTotalGrimps(ilSpotType)
                    lmGRP(ilRow) = tmReRate(illoop).lTotalGRP(ilSpotType)
                    
                    If ((tmReRate(illoop).sType = "B") And (tmReRate(illoop).sSubType = "B")) Then
                        ReDim Preserve lmAdvtBonusReRateCost(0 To ilAdvtBonusReRateRow) As Long
                        ReDim Preserve lmAdvtBonusReRateGrimp(0 To ilAdvtBonusReRateRow) As Long
                        ReDim Preserve lmAdvtBonusReRateGRP(0 To ilAdvtBonusReRateRow) As Long
                        'lmAdvtBonusReRateCost(ilAdvtBonusReRateRow) = tmReRate(illoop).lTotalCost(ilSpotType)
                        lmAdvtBonusReRateCost(ilAdvtBonusReRateRow) = tmReRate(illoop).dTotalCost(ilSpotType) 'TTP 10439 - Rerate 21,000,000
                        lmAdvtBonusReRateGrimp(ilAdvtBonusReRateRow) = tmReRate(illoop).lTotalGrimps(ilSpotType)
                        lmAdvtBonusReRateGRP(ilAdvtBonusReRateRow) = tmReRate(illoop).lTotalGRP(ilSpotType)
                        ilAdvtBonusReRateRow = ilAdvtBonusReRateRow + 1
                        lmAdvtBonusReRateTotalSpots = lmAdvtBonusReRateTotalSpots + tmReRate(illoop).lTotalSpots(ilSpotType)
                        If imAdvtBonusReRateDnfCode = -1 Then
                            imAdvtBonusReRateDnfCode = imBonusTotalDnfCode
                        Else
                            If (imAdvtBonusReRateDnfCode <> imBonusTotalDnfCode) And (imAdvtBonusReRateDnfCode <> -2) Then
                                imAdvtBonusReRateDnfCode = -2
                            End If
                        End If
                    End If
                    
                    ReDim Preserve lmAdvtPlusBonusReRateCost(0 To ilAdvtPlusBonusReRateRow) As Long
                    ReDim Preserve lmAdvtPlusBonusReRateGrimp(0 To ilAdvtPlusBonusReRateRow) As Long
                    ReDim Preserve lmAdvtPlusBonusReRateGRP(0 To ilAdvtPlusBonusReRateRow) As Long
                    'lmAdvtPlusBonusReRateCost(ilAdvtPlusBonusReRateRow) = tmReRate(illoop).lTotalCost(ilSpotType)
                    lmAdvtPlusBonusReRateCost(ilAdvtPlusBonusReRateRow) = tmReRate(illoop).dTotalCost(ilSpotType) 'TTP 10439 - Rerate 21,000,000
                    lmAdvtPlusBonusReRateGrimp(ilAdvtPlusBonusReRateRow) = tmReRate(illoop).lTotalGrimps(ilSpotType)
                    lmAdvtPlusBonusReRateGRP(ilAdvtPlusBonusReRateRow) = tmReRate(illoop).lTotalGRP(ilSpotType)
                    ilAdvtPlusBonusReRateRow = ilAdvtPlusBonusReRateRow + 1
                    lmAdvtPlusBonusReRateTotalSpots = lmAdvtPlusBonusReRateTotalSpots + tmReRate(illoop).lTotalSpots(ilSpotType)
                    'determine if varying populations across the weeks (demo estimates) or across the lines
                    If tmReRate(illoop).iDnfCode(ilSpotType) > 0 Then
                        If tgSpf.sDemoEstAllowed = "Y" Then             '6-4-04
                            If llPop < 0 Then
                                llPop = tmReRate(illoop).lSatelliteEst(ilSpotType)
                            ElseIf llPop <> tmReRate(illoop).lSatelliteEst(ilSpotType) And tmReRate(illoop).lSatelliteEst(ilSpotType) <> 0 Then
                                llPop = 0
                            End If
    
                        Else
                            If llPop < 0 Then
                                llPop = tmReRate(illoop).lPop(ilSpotType)
                            ElseIf llPop <> tmReRate(illoop).lPop(ilSpotType) And tmReRate(illoop).lPop(ilSpotType) <> 0 Then
                                llPop = 0
                            End If
                        End If
                        If ilCntrTotalDnfCode <> imBonusTotalDnfCode Then
                            If imBonusTotalDnfCode <> -1 Then
                                ilPlusDnfCode = -2
                            Else
                                ilPlusDnfCode = ilCntrTotalDnfCode
                            End If
                        Else
                            ilPlusDnfCode = imBonusTotalDnfCode
                        End If
                        If imAdvtPlusBonusReRateDnfCode = -1 Then
                           imAdvtPlusBonusReRateDnfCode = ilPlusDnfCode 'imBonusTotalDnfCode
                        Else
                           'If (imAdvtPlusBonusReRateDnfCode <> imBonusTotalDnfCode) And (imAdvtPlusBonusReRateDnfCode <> -2) Then
                           If (imAdvtPlusBonusReRateDnfCode <> ilPlusDnfCode) And (imAdvtPlusBonusReRateDnfCode <> -2) Then
                               imAdvtPlusBonusReRateDnfCode = -2
                           End If
                        End If
                        If ((tmReRate(illoop).sType = "B") And (tmReRate(illoop).sSubType = "B")) Then
                            If lmAdvtBonusReRatePop = -1 Then
                                lmAdvtBonusReRatePop = llPop
                            Else
                                If lmAdvtBonusReRatePop <> llPop And (lmAdvtBonusReRatePop <> 0) Then
                                    lmAdvtBonusReRatePop = 0
                                End If
                            End If
                            If lmAdvtPlusBonusReRatePop = -1 Then
                                lmAdvtPlusBonusReRatePop = llPop
                            Else
                                If lmAdvtPlusBonusReRatePop <> llPop And (lmAdvtPlusBonusReRatePop <> 0) Then
                                    lmAdvtPlusBonusReRatePop = 0
                                End If
                            End If
                        End If
                    End If
                    llLnSpots = llLnSpots + tmReRate(illoop).lTotalSpots(ilSpotType)
                    ilRow = ilRow + 1
                End If
            Next illoop
            'Obtain totals
            If ilRow > 1 Then
                ReDim tmReRate(0 To 0) As RERATEINFO
                tmReRate(0) = tlContractOrderReRate
                'gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llLnSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
                gResearchTotals sm1or2PlaceRating, False, llPop, lmCost(), lmGrimp(), lmGRP(), llLnSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
                'If llTotalCPP <> 0 Or llTotalCPM <> 0 Then   '10/13/20 - TTP 9953 - Commented Out due to Zero Dollar contracts would retain last Contract Totals (tlContractOrderReRate)
                    'maintain research contract totals for all contracts to be able to total them by demo
                    'include all contracts types (with remnants, DR, & PI)
                    ilReRate = UBound(tmReRate)
                    tmReRate(ilReRate).lChfCode(0) = tmChfPurchase.lCode
                    tmReRate(ilReRate).lChfCode(1) = tmChfReRate.lCode
                    tmReRate(ilReRate).lClfCode(0) = 0
                    tmReRate(ilReRate).lClfCode(1) = 0
                    tmReRate(ilReRate).iVefCode = 0
                    tmReRate(ilReRate).iRdfCode = 0
                    tmReRate(ilReRate).sType = "T"      'Contract total plus bonus
                    tmReRate(ilReRate).sSubType = "B"
                    tmReRate(ilReRate).lCntrNo = tmChfPurchase.lCntrNo
                    tmReRate(ilReRate).sProduct = tmChfPurchase.sProduct
                    tmReRate(ilReRate).iLineNo = 0
                    tmReRate(ilReRate).iPkLineNo = 0
                    tmReRate(ilReRate).iLen = 0
                    tmReRate(ilReRate).sAudioType = ""
                    tmReRate(ilReRate).sACT1LineupCode = ""
                    '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                    tmReRate(ilReRate).sACT1StoredTime = ""
                    tmReRate(ilReRate).sACT1StoredSpots = ""
                    tmReRate(ilReRate).sACT1StoreClearPct = ""
                    tmReRate(ilReRate).sACT1DaypartFilter = ""
                    If imBonusTotalDnfCode <> ilCntrTotalDnfCode Then
                        tmReRate(ilReRate).iDnfCode(ilSpotType) = -2
                    Else
                        tmReRate(ilReRate).iDnfCode(ilSpotType) = imBonusTotalDnfCode   'imReRateDnfCode
                    End If
                    tmReRate(ilReRate).lPop(ilSpotType) = llPop
                    tmReRate(ilReRate).lSatelliteEst(ilSpotType) = llPop            '6-4-04
                    'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                    tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                    tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                    tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                    tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                    tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                    tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                    tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                    tmReRate(ilReRate).lTotalSpots(ilSpotType) = llLnSpots
                    tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                    tmReRate(ilReRate).sLineComment = "" ' "Obtain Totals Comment"
                    tmReRate(ilReRate).sPriceType = "" 'tmCff.sPriceType
                    blReRateAdded = True
                    ReDim Preserve tmReRate(0 To ilReRate + 1)
                'End If  '10/13/20 - TTP 9953 - Commented Out 'cpp or cpm are zero
                'Print #hmToCSV, ""
                'mPrint ""
                mOutputResearch tmReRate(0)
            End If
        
            iProcessedCount = iProcessedCount + 1
            If iProcessedCount / iCntrCount * 100 < 100 Then
                prgProcessing.Value = iProcessedCount / iCntrCount * 100
            End If
        End If
    Next ilChf                  'get another cnt
    prgProcessing.Value = 100
    lblProcessing.Caption = "Generating Excel ..."
    'Create Advertiser Total
    If imNoCntr > 1 Then
        ilSpotType = -1
        ReDim tmReRate(0 To 0) As RERATEINFO
        'gResearchTotals sm1or2PlaceRating, False, lmAdvtOrderPop, lmAdvtOrderCost(), lmAdvtOrderGrimp(), lmAdvtOrderGRP(), lmAdvtOrderTotalSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
        gResearchTotals sm1or2PlaceRating, False, lmAdvtOrderPop, lmAdvtOrderCost(), lmAdvtOrderGrimp(), lmAdvtOrderGRP(), lmAdvtOrderTotalSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
        'If llTotalCPP <> 0 Or llTotalCPM <> 0 Then   '10/13/20 - TTP 9953 - Commented Out due to Zero Dollar contracts would retain last Contract Totals (tlContractOrderReRate)
            'maintain research contract totals for all contracts to be able to total them by demo
            'include all contracts types (with remnants, DR, & PI)
            ilSpotType = 0
            ilReRate = UBound(tmReRate)
            tmReRate(ilReRate).lChfCode(0) = 0
            tmReRate(ilReRate).lChfCode(1) = 0
            tmReRate(ilReRate).lClfCode(0) = 0
            tmReRate(ilReRate).lClfCode(1) = 0
            tmReRate(ilReRate).iVefCode = 0
            tmReRate(ilReRate).iRdfCode = 0
            tmReRate(ilReRate).sType = "T"
            tmReRate(ilReRate).sSubType = "A"
            tmReRate(ilReRate).lCntrNo = -1
            tmReRate(ilReRate).sProduct = ""
            tmReRate(ilReRate).iLineNo = 0
            tmReRate(ilReRate).iPkLineNo = 0
            tmReRate(ilReRate).iLen = 0
            tmReRate(ilReRate).sAudioType = ""
            tmReRate(ilReRate).sACT1LineupCode = ""
            '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
            tmReRate(ilReRate).sACT1StoredTime = ""
            tmReRate(ilReRate).sACT1StoredSpots = ""
            tmReRate(ilReRate).sACT1StoreClearPct = ""
            tmReRate(ilReRate).sACT1DaypartFilter = ""
            tmReRate(ilReRate).iDnfCode(ilSpotType) = imAdvtOrderDnfCode
            tmReRate(ilReRate).lPop(ilSpotType) = lmAdvtOrderPop
            tmReRate(ilReRate).lSatelliteEst(ilSpotType) = lmAdvtOrderPop            '6-4-04
            'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
            tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
            tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
            tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
            tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
            tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
            tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
            tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
            tmReRate(ilReRate).lTotalSpots(ilSpotType) = lmAdvtOrderTotalSpots
            tmReRate(ilReRate).sCBS(ilSpotType) = "N"
            tmReRate(ilReRate).sLineComment = "" ' "Adv Total Comment"
            tmReRate(ilReRate).sPriceType = "" 'tmCff.sPriceType
            tlAdvtOrderReRate = tmReRate(ilReRate)
        'End If  '10/13/20 - TTP 9953 - Commented Out
        'gResearchTotals sm1or2PlaceRating, False, lmAdvtReRatePop, lmAdvtReRateCost(), lmAdvtReRateGrimp(), lmAdvtReRateGRP(), lmAdvtReRateTotalSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
        gResearchTotals sm1or2PlaceRating, False, lmAdvtReRatePop, lmAdvtReRateCost(), lmAdvtReRateGrimp(), lmAdvtReRateGRP(), lmAdvtReRateTotalSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
        'If llTotalCPP <> 0 Or llTotalCPM <> 0 Then   '10/13/20 - TTP 9953 - Commented Out due to Zero Dollar contracts would retain last Contract Totals (tlContractOrderReRate)
            'maintain research contract totals for all contracts to be able to total them by demo
            'include all contracts types (with remnants, DR, & PI)
            ilSpotType = 1
            ilReRate = UBound(tmReRate)
            tmReRate(ilReRate).iDnfCode(ilSpotType) = imAdvrReRateDnfCode
            tmReRate(ilReRate).lPop(ilSpotType) = lmAdvtReRatePop
            tmReRate(ilReRate).lSatelliteEst(ilSpotType) = lmAdvtReRatePop            '6-4-04
            'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
            tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
            tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
            tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
            tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
            tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
            tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
            tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
            tmReRate(ilReRate).lTotalSpots(ilSpotType) = lmAdvtReRateTotalSpots
            tmReRate(ilReRate).sCBS(ilSpotType) = "N"
            tmReRate(ilReRate).sLineComment = "" '"Another Adv Comment"
            tmReRate(ilReRate).sPriceType = "" 'tmCff.sPriceType

        'End If  '10/13/20 - TTP 9953 - Commented Out
        If ilSpotType <> -1 Then
            'Print #hmToCSV, ""
            'TTP 10082 - merge header into columns
            If rbcLayout(0).Value = True Then mPrint ""
            mOutputResearch tmReRate(0)
            
            'Output Bonus total
            If ilAdvtBonusReRateRow > 0 Then
                ilSpotType = 1
                ReDim tmReRate(0 To 0) As RERATEINFO
                'gResearchTotals sm1or2PlaceRating, False, lmAdvtBonusReRatePop, lmAdvtBonusReRateCost(), lmAdvtBonusReRateGrimp(), lmAdvtBonusReRateGRP(), lmAdvtBonusReRateTotalSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
                gResearchTotals sm1or2PlaceRating, False, lmAdvtBonusReRatePop, lmAdvtBonusReRateCost(), lmAdvtBonusReRateGrimp(), lmAdvtBonusReRateGRP(), lmAdvtBonusReRateTotalSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
                'If llTotalCPP <> 0 Or llTotalCPM <> 0 Then
                If lmAdvtBonusReRateTotalSpots > 0 Then
                    'maintain research contract totals for all contracts to be able to total them by demo
                    'include all contracts types (with remnants, DR, & PI)
                    ilReRate = UBound(tmReRate)
                    tmReRate(ilReRate).lChfCode(0) = 0
                    tmReRate(ilReRate).lChfCode(1) = 0
                    tmReRate(ilReRate).lClfCode(0) = 0
                    tmReRate(ilReRate).lClfCode(1) = 0
                    tmReRate(ilReRate).iVefCode = 0
                    tmReRate(ilReRate).iRdfCode = 0
                    tmReRate(ilReRate).sType = "B"
                    tmReRate(ilReRate).sSubType = "N"
                    tmReRate(ilReRate).lCntrNo = -1
                    tmReRate(ilReRate).sProduct = ""
                    tmReRate(ilReRate).iLineNo = 0
                    tmReRate(ilReRate).iPkLineNo = 0
                    tmReRate(ilReRate).iLen = 0
                    tmReRate(ilReRate).sAudioType = ""
                    tmReRate(ilReRate).sACT1LineupCode = ""
                    '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
                    tmReRate(ilReRate).sACT1StoredTime = ""
                    tmReRate(ilReRate).sACT1StoredSpots = ""
                    tmReRate(ilReRate).sACT1StoreClearPct = ""
                    tmReRate(ilReRate).sACT1DaypartFilter = ""
                    tmReRate(ilReRate).iDnfCode(ilSpotType) = imAdvtBonusReRateDnfCode
                    tmReRate(ilReRate).lPop(ilSpotType) = lmAdvtBonusReRatePop
                    tmReRate(ilReRate).lSatelliteEst(ilSpotType) = lmAdvtReRatePop            '6-4-04
                    'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                    tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                    tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                    tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                    tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                    tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                    tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                    tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                    tmReRate(ilReRate).lTotalSpots(ilSpotType) = lmAdvtBonusReRateTotalSpots
                    tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                    tmReRate(ilReRate).sLineComment = "" '"Bonus Total Comment"
                    tmReRate(ilReRate).sPriceType = "" 'tmCff.sPriceType
                    'Print #hmToCSV, ""
                    'mPrint ""
                    mOutputResearch tmReRate(0)
                End If
                ''Print #hmToCSV, ""
                'mPrint ""
                'mOutputResearch tmReRate(0)
            End If
            
            'Output Advertiser Total plus Bonus
            If ilAdvtPlusBonusReRateRow > 0 Then
                ilSpotType = 1
                ReDim tmReRate(0 To 0) As RERATEINFO
                tmReRate(0) = tlAdvtOrderReRate
                'gResearchTotals sm1or2PlaceRating, False, lmAdvtPlusBonusReRatePop, lmAdvtPlusBonusReRateCost(), lmAdvtPlusBonusReRateGrimp(), lmAdvtPlusBonusReRateGRP(), lmAdvtPlusBonusReRateTotalSpots, llTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
                gResearchTotals sm1or2PlaceRating, False, lmAdvtPlusBonusReRatePop, lmAdvtPlusBonusReRateCost(), lmAdvtPlusBonusReRateGrimp(), lmAdvtPlusBonusReRateGRP(), lmAdvtPlusBonusReRateTotalSpots, dlTotalCost, ilTotalAvgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud 'TTP 10439 - Rerate 21,000,000
                'If llTotalCPP <> 0 Or llTotalCPM <> 0 Then   '10/13/20 - TTP 9953 - Commented Out due to Zero Dollar contracts would retain last Contract Totals (tlContractOrderReRate)
                    'maintain research contract totals for all contracts to be able to total them by demo
                    'include all contracts types (with remnants, DR, & PI)
                    ilReRate = UBound(tmReRate)
                    tmReRate(ilReRate).sSubType = "G"
                    tmReRate(ilReRate).iDnfCode(ilSpotType) = imAdvtPlusBonusReRateDnfCode
                    tmReRate(ilReRate).lPop(ilSpotType) = lmAdvtPlusBonusReRatePop
                    tmReRate(ilReRate).lSatelliteEst(ilSpotType) = lmAdvtReRatePop            '6-4-04
                    'tmReRate(ilReRate).lTotalCost(ilSpotType) = llTotalCost
                    tmReRate(ilReRate).dTotalCost(ilSpotType) = dlTotalCost 'TTP 10439 - Rerate 21,000,000
                    tmReRate(ilReRate).lTotalCPP(ilSpotType) = llTotalCPP
                    tmReRate(ilReRate).lTotalCPM(ilSpotType) = llTotalCPM
                    tmReRate(ilReRate).lTotalGrimps(ilSpotType) = llTotalGrImp
                    tmReRate(ilReRate).lTotalGRP(ilSpotType) = llTotalGRP
                    tmReRate(ilReRate).iTotalAvgRating(ilSpotType) = ilTotalAvgRtg
                    tmReRate(ilReRate).lTotalAvgAud(ilSpotType) = llTotalAvgAud
                    tmReRate(ilReRate).lTotalSpots(ilSpotType) = lmAdvtPlusBonusReRateTotalSpots
                    tmReRate(ilReRate).sCBS(ilSpotType) = "N"
                    tmReRate(ilReRate).sLineComment = "" ' "Adv Plus Bonus Comment"
                    tmReRate(ilReRate).sPriceType = "" 'tmCff.sPriceType
                'End If  '10/13/20 - TTP 9953 - Commented Out
                'Print #hmToCSV, ""
                'mPrint ""
                mOutputResearch tmReRate(0)
            End If
        End If
    End If
 'process next demo
    frcProcessing.Visible = False
End Sub

Private Sub frcColumnLayout_Click()
    mSetShow
End Sub

Private Sub frcDemo_Click()
    mSetShow
End Sub

Private Sub frcIndex_Click()
    mSetShow
End Sub

Private Sub frcInstallMethod_Click()
    mSetShow
End Sub

Private Sub frcMQ_Click()
    mSetShow
End Sub

Private Sub frcReRateBook_Click()
    mSetShow
End Sub

Private Sub frcShow_Click()
    mSetShow
End Sub

Private Sub grdCntr_Click()
'    Dim llRow As Long
'
'    bmInGrid = True
'    If grdCntr.Row >= grdCntr.FixedRows Then
'        llRow = grdCntr.MouseRow
'        If grdCntr.TextMatrix(grdCntr.Row, PRODUCTINDEX) <> "" Then
'            If grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
'                If imCtrlKey Then
'                    grdCntr.TextMatrix(llRow, SELECTEDINDEX) = ""
'                End If
'            Else
'                grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1"
'            End If
'        Else
'            grdCntr.TextMatrix(llRow, SELECTEDINDEX) = ""
'        End If
'    End If
'    ckcAllCntr.Value = vbUnchecked
'    bmInGrid = False
'    mSetCommands
    '3/4/21 - TTP 10088: Sort by Product or Contract (ASC/DESC)
    If grdCntr.MouseRow < grdCntr.FixedRows Then
        mSortByColumn grdCntr.MouseCol
    End If
    mSetCommands
End Sub

Private Sub grdCntr_GotFocus()
    mSetShow
End Sub

Private Sub grdCntr_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If

End Sub

Private Sub grdCntr_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdCntr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdCntr.RowHeight(0) Then
        'grdCntr.Col = grdCntr.MouseCol
        'mVehSortCol grdCntr.Col
        'grdCntr.Row = 0
        'grdCntr.Col = PRODUCTINDEX
        Exit Sub
    End If
    '4/7/21 JW Prevent Click in Grey area past last Row from Selecting the Last row
    If Y > ((grdCntr.CellHeight * (grdCntr.Rows - (grdCntr.TopRow - grdCntr.FixedRows))) + (grdCntr.GridLineWidth * Screen.TwipsPerPixelY * (grdCntr.Rows - 1))) Then
        Exit Sub
    End If
    'D.S. 07-28-17
    'ilFound = gGrid_GetRowCol(grdCntr, X, Y, llCurrentRow, llCol)
    bmInGrid = True
    llCurrentRow = grdCntr.MouseRow
    llCol = grdCntr.MouseCol
    If llCurrentRow < grdCntr.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdCntr.FixedRows Then
        'If grdCntr.TextMatrix(llCurrentRow, PRODUCTINDEX) <> "" Then
        If grdCntr.TextMatrix(llCurrentRow, CNTRNOINDEX) <> "" Then
            If llCol = VERSIONINDEX Then
                grdCntr.TextMatrix(grdCntr.Row, SELECTEDINDEX) = "1"
                mPaintRowColor grdCntr.Row
                lmLastClickedRow = llCurrentRow
                grdCntr.Row = llCurrentRow
                grdCntr.Col = llCol
            Else
                'grdCntr.TopRow = lmScrollTop
                llTopRow = grdCntr.TopRow
                'If (Shift And CTRLMASK) > 0 Then
                '    If grdCntr.TextMatrix(grdCntr.Row, PURCHASECHFCODEINDEX) <> "" Then
                '        If grdCntr.TextMatrix(grdCntr.Row, SELECTEDINDEX) <> "1" Then
                '            grdCntr.TextMatrix(grdCntr.Row, SELECTEDINDEX) = "1"
                '        Else
                '            grdCntr.TextMatrix(grdCntr.Row, SELECTEDINDEX) = ""
                '        End If
                '        mPaintRowColor grdCntr.Row
                '    End If
                'Else
                '    For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
                '        If grdCntr.TextMatrix(llRow, PRODUCTINDEX) <> "" Then
                '            grdCntr.TextMatrix(llRow, SELECTEDINDEX) = ""
                '            If grdCntr.TextMatrix(llRow, PURCHASECHFCODEINDEX) <> "" Then
                '                If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                '                    If llRow = llCurrentRow Then
                '                        grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '                    Else
                '                        grdCntr.TextMatrix(llRow, SELECTEDINDEX) = ""
                '                    End If
                '                ElseIf lmLastClickedRow < llCurrentRow Then
                '                    If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                '                        grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '                    End If
                '                Else
                '                    If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                '                        grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1"
                '                    End If
                '                End If
                '                mPaintRowColor llRow
                '            End If
                '        End If
                '    Next llRow
                '    grdCntr.TopRow = llTopRow
                '    grdCntr.Row = llCurrentRow
                'End If
                If grdCntr.TextMatrix(grdCntr.Row, PURCHASECHFCODEINDEX) <> "" Then
                    If grdCntr.TextMatrix(grdCntr.Row, SELECTEDINDEX) <> "1" Then
                        grdCntr.TextMatrix(grdCntr.Row, SELECTEDINDEX) = "1"
                    Else
                        grdCntr.TextMatrix(grdCntr.Row, SELECTEDINDEX) = ""
                    End If
                    mPaintRowColor grdCntr.Row
                    
                    'TTP 10140, show When [Set Book by Line] method is enabled - Keep track of Last "Research Book Name" option (rbcReRateBook)
                    'Reset the Line by Book because we made a Contract Change
                    If UBound(tgBookByLineAssigned) = 0 Then
                        bmBookByLine = False
                        rbcReRateBookByLine.Value = False
                        'rbcReRateBookByLine.Enabled = False
                        rbcReRateBook(imReRateLastBookMode).Value = True
                    End If
                End If
                grdCntr.TopRow = llTopRow
                grdCntr.Row = llCurrentRow
                lmLastClickedRow = llCurrentRow
            End If
        End If
    End If
    mSetDemo
    'smGridTypeAhead = ""
    If llCol = VERSIONINDEX Then
        mEnableBox
    Else
        grdCntr.Row = 0
        grdCntr.Col = SELECTEDINDEX
        'bmInGrid = False
    End If
    ckcAllCntr.Value = vbUnchecked
    bmInGrid = False
    
    'TTP 10172 - 7/1/21 - JW - list of purchased books doesn't refresh when selecting a smaller set of contracts
    mClearBookByLine
    
End Sub

Private Sub grdCntr_Scroll()
    mSetShow
End Sub

Private Sub lacDemo_Click()
    mSetShow
End Sub

Private Sub lacInclude_Click()
    mSetShow
End Sub

Private Sub lacIndex_Click()
    mSetShow
End Sub

Private Sub lacMonth_Click()
    mSetShow
End Sub

Private Sub lacShow_Click()
    mSetShow
End Sub

Private Sub lacTitle_Click(Index As Integer)
    mSetShow
End Sub

Private Sub lbcAdvertiser_Click()
    Dim slCntrStatus As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    
    If lbcAdvertiser.ListIndex >= 0 Then
        slNameCode = tgAdvertiser(lbcAdvertiser.ListIndex).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
        imAdfCode = Val(slCode)
        mPopCntr
        'smToExcel = sgExportPath & "ReRate_" & Trim$(lbcAdvertiser.Text) & ".xlsx"
        'edcExcel.Text = smToExcel
        
        
        
        'smToCSV = sgExportPath & "ReRateExport_" & Trim$(lbcAdvertiser.Text) & ".CSV"
        'edcCSV.Text = smToCSV
    Else
        imAdfCode = -1
        
        'smToExcel = ""
        'edcExcel.Text = ""
            
        'smToCSV = ""
        'edcCSV.Text = ""
        mClearGrid
    End If
    
    'TTP 10258: ReRate - make it work without requiring Office
    mGetCSVFilename
    
    'ReDim tgBookByLineCntr(0 To 0) As BOOKBYLINECNTR
    'ReDim tgBookByLineAssigned(0 To 0) As BOOKBYLINEASSIGNED
    'reset "Research Book Name" option, set last option back
    'bmBookByLine = False
    'rbcReRateBookByLine.Value = False
    'rbcReRateBookByLine.Enabled = False
    'rbcReRateBook(imReRateLastBookMode).Value = True
    mClearBookByLine
    mSetCommands
End Sub

Private Sub lbcAdvertiser_GotFocus()
    mSetShow
End Sub

Private Sub lbcSpotLens_Click()
    If Not bmAllClicked Then
        bmSetAll = False
        ckcAllSpotLens.Value = vbUnchecked
        bmSetAll = True
    End If
    'ReDim tgBookByLineCntr(0 To 0) As BOOKBYLINECNTR
    'ReDim tgBookByLineAssigned(0 To 0) As BOOKBYLINEASSIGNED
    ''reset "Research Book Name" option, set last option back
    'bmBookByLine = False
    'rbcReRateBookByLine.Value = False
    ''rbcReRateBookByLine.Enabled = False
    'rbcReRateBook(imReRateLastBookMode).Value = True
    mClearBookByLine
    mSetCommands
End Sub

Private Sub lbcSpotLens_GotFocus()
    mSetShow
End Sub

Private Sub rbcColumnLayout_Click(Index As Integer)
    If rbcColumnLayout(Index).Value Then
        mSetCommands
    End If
End Sub

Private Sub rbcColumnLayout_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Sub rbcDatesBy_Click(Index As Integer)
    If rbcDatesBy(Index).Value Then
        Select Case Index
            Case 0  'Week
                frcMQ.Visible = False
                edcDate(0).Visible = False
                edcDate(1).Visible = False
                edcDate(2).Visible = False
                lacTitle(5).Visible = False
                lacTitle(6).Visible = False
                lacTitle(0).Caption = "Week Start Date"
                lacTitle(0).Visible = True
                edcDate(0).Text = ""
                edcDate(0).Visible = True
            Case 1  'Month
                edcDate(0).Visible = False
                edcDate(1).Visible = False
                edcDate(2).Visible = False
                lacTitle(0).Visible = False
                lacTitle(5).Visible = False
                lacTitle(6).Visible = False
                lacTitle(0).Visible = False
                lacTitle(3).Caption = "Month"
                edcStart.Text = ""
                edcYear.Text = ""
                frcMQ.Visible = True
            Case 2  'Quarter
                edcDate(0).Visible = False
                edcDate(1).Visible = False
                edcDate(2).Visible = False
                lacTitle(0).Visible = False
                lacTitle(5).Visible = False
                lacTitle(6).Visible = False
                lacTitle(0).Visible = False
                lacTitle(3).Caption = "Quarter"
                edcStart.Text = ""
                edcYear.Text = ""
                frcMQ.Visible = True
            Case 3  'Contract
                frcMQ.Visible = False
                edcDate(1).Visible = False
                edcDate(2).Visible = False
                lacTitle(5).Visible = False
                lacTitle(6).Visible = False
                lacTitle(0).Caption = "Active on or After"
                lacTitle(0).Visible = True
                edcDate(0).Text = ""
                edcDate(0).Visible = True
            Case 4  'Rangle
                frcMQ.Visible = False
                lacTitle(0).Visible = False
                edcDate(0).Visible = False
                edcDate(1).Text = ""
                edcDate(2).Text = ""
                edcDate(1).Visible = True
                edcDate(2).Visible = True
                lacTitle(5).Visible = True
                lacTitle(6).Visible = True
        End Select
        'ReDim tgBookByLineCntr(0 To 0) As BOOKBYLINECNTR
        'ReDim tgBookByLineAssigned(0 To 0) As BOOKBYLINEASSIGNED
        ''reset "Research Book Name" option, set last option back
        'bmBookByLine = False
        'rbcReRateBookByLine.Value = False
        ''rbcReRateBookByLine.Enabled = False
        'rbcReRateBook(imReRateLastBookMode).Value = True
        mClearBookByLine
        mPopCntr
        mSetCommands
    End If
End Sub

Private Sub rbcDatesBy_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Sub rbcDemo_Click(Index As Integer)
    If rbcDemo(Index).Value Then
        If Index = 4 Then
            cbcDemo.Enabled True
        Else
            cbcDemo.Enabled False
        End If
        mSetCommands
    End If
End Sub

Private Sub rbcDemo_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Sub rbcIndex_Click(Index As Integer)
    If rbcIndex(Index).Value Then
        mSetCommands
    End If
End Sub

Private Sub rbcIndex_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Sub rbcLayout_Click(Index As Integer)
    'TTP 10082 - merge header into columns (0=Original Header/Detail Layout, 1=merged layout)
    If Index = 0 Then 'Separate
        ckcSummary.Enabled = True
    ElseIf Index = 1 Then 'Merged with each research row
        ckcSummary.Enabled = False
        ckcSummary.Value = vbUnchecked
    End If
End Sub

Private Sub rbcReRateBook_Click(Index As Integer)
    If rbcReRateBook(Index).Value Then
        'ReDim tgBookByLineCntr(0 To 0) As BOOKBYLINECNTR
        'ReDim tgBookByLineAssigned(0 To 0) As BOOKBYLINEASSIGNED
        'tgBookByLineAssigned(0).iNext = -1
        'bmBookByLine = False
        'rbcReRateBookByLine.Value = False
        ''rbcReRateBookByLine.Enabled = False
        mClearBookByLine (False) 'Dont reset the last Used option when Clearing the BookByLine cache
        ckcCsv.Enabled = True
        Select Case Index
            Case 0  'Vehicle Default
            Case 1  'Closet to Air Date
            Case 2  'None
                ckcCost.Value = vbChecked
                ckcRating.Value = vbChecked
                ckcCPM.Value = vbChecked
                ckcCsv.Value = vbUnchecked: ckcCsv.Enabled = False
            Case 3  'Contract Line
        End Select
        mSetCommands
    End If
    imReRateLastBookMode = Index 'Keep track of Last "Research Book Name" option (rbcReRateBook)
End Sub

Private Sub rbcReRateBook_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Sub rbcReRateBookByLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBook As Integer
    bmBookByLine = False
    For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
        '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
        If tgBookByLineAssigned(ilBook).iReRateDnfCode <> 0 Then
            bmBookByLine = True
            Exit For
        End If
    Next ilBook
    If bmBookByLine Then
        rbcReRateBookByLine.Value = True
    Else
        rbcReRateBookByLine.Value = False
        If cmcSetBook.Enabled = True Then cmcSetBook_Click
    End If
End Sub

Private Sub rbcRevNo_Click(Index As Integer)
    Dim llRow As Long
    If rbcRevNo(Index).Value Then
        For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
            If grdCntr.TextMatrix(llRow, CNTRNOINDEX) <> "" Then
                mPopRevisions grdCntr.TextMatrix(llRow, CNTRNOINDEX)
                If Index = 0 Then
                    grdCntr.TextMatrix(llRow, VERSIONINDEX) = "Original"
                    grdCntr.TextMatrix(llRow, PURCHASECHFCODEINDEX) = cbcRevision.GetItemData(0)
                Else
                    grdCntr.TextMatrix(llRow, VERSIONINDEX) = cbcRevision.GetName(cbcRevision.ListCount - 1)
                    grdCntr.TextMatrix(llRow, PURCHASECHFCODEINDEX) = cbcRevision.GetItemData(cbcRevision.ListCount - 1)
                End If
            End If
        Next llRow
    End If
End Sub

Private Sub rbcShow_Click(Index As Integer)
    If rbcShow(Index).Value Then
        mSetCommands
    End If
End Sub

Private Sub rbcShow_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub

Private Sub mOutputResearch(tlReRate As RERATEINFO)
    Dim ilAdf As Integer
    Dim ilRdf As Integer
    Dim ilVef As Integer
    Dim ilDnf As Integer
    Dim illoop As Integer
    Dim slStr As String
    Dim ilSpotType As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilRet As Integer
    Dim ilCol As Integer
    Dim ilColorCode As Integer '0=no; 1=red; 2=green
    Dim blClearTotalFields As Boolean ' Rate; Rtging; AQH
    'Dim blIncludeFormula As Boolean
    'Dim llFormulaSpots As Long
    'Dim llFormulaAQH As Long
    'Dim llFormulaPop As Long
    'Dim llFormulaExtTotal As Long
    Dim slRowType As String
    Dim slBook As String
    Dim ilBook As Integer
    Dim blUsingClosest As Boolean
    Dim ilClf As Integer
    Dim blBuildBooks As Boolean
    Dim slRecord As String
    'TTP 9912 - ReRate report when using the show package only or show package and hidden options, if there's a mix of purchased books
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    
    blClearTotalFields = False
    'Print #hmToCSV, ",Line#,Package Line#,Vehicle,Daypart,Lineup #,Spot Length,Book Name,# Spots,Dollars,AQH,Population,Gross Impressions,Rating,GRP,CPM,CPP,Book Name,# Spots,Dollars,AQH,Population,Gross Impressions,Rating,GRP,CPM,CPP,Index  "
    slRecord = ""
    'TTP 10082 - merge header into columns (Add Contract Header details to colums)
    If rbcLayout(1).Value = True Then
        slRecord = tmReRateHeader.Agency & smDelimiter & tmReRateHeader.Advertiser & smDelimiter & tmReRateHeader.Product & smDelimiter & tmReRateHeader.OrderNo & smDelimiter & tmReRateHeader.PurchaseRevision & smDelimiter & tmReRateHeader.ReRateRevision & smDelimiter & tmReRateHeader.Demo & smDelimiter & tmReRateHeader.ContractPopulation & smDelimiter & tmReRateHeader.ReRatePopulation & smDelimiter & tmReRateHeader.Period & smDelimiter
    End If

    'blIncludeFormula = False
    'llFormulaSpots = -1
    'llFormulaAQH = -1
    'llFormulaPop = -1
    'llFormulaExtTotal = -1
    'Ignore hidden lines if showing package numbers
    If rbcShow(1).Value = True And (tlReRate.sType = "H") Then
        Exit Sub
    End If
    If (tlReRate.sType <> "T") And (tlReRate.sType <> "B") And (tlReRate.sSubType <> "M") Then
        If ckcIncludeInactive.Value = vbUnchecked Then
            If (tlReRate.lTotalSpots(0) <= 0) And (tlReRate.lTotalSpots(1) <= 0) Then
                Exit Sub
            End If
        End If
        If (tlReRate.sCBS(0) = "Y") And (tlReRate.sCBS(1) = "Y") Then
            Exit Sub
        End If
    End If
    If (tlReRate.sType <> "T") And (tlReRate.sType <> "B") And (tlReRate.sSubType <> "M") Then
        slRowType = tlReRate.sType
        If slRowType = "A" Or slRowType = "O" Then
            slRowType = "PL"
        End If
        If slRowType = "H" Then
            slRowType = "HL"
        End If
        If slRowType = "S" Then
            slRowType = "SL"
        End If
        '11/13/2020 - TTP 9993 - ReRate Gimps and Grps lines twice, when unpackage or package on a later revision.  Don't show Negative Line Numbers on output
        'slRecord = slRecord & "," & tlReRate.iLineNo
        slRecord = slRecord & smDelimiter & Abs(Val(tlReRate.iLineNo))
        'If tlReRate.sType = "H" Then
        '    slRecord = slRecord & "," & tlReRate.iPkLineNo
        'Else
        '    slRecord = slRecord & ","
        'End If
        ilVef = gBinarySearchVef(tlReRate.iVefCode) 'Line#
        If tlReRate.sType = "H" Then 'Vehicle
            If ilVef <> -1 Then
                slRecord = slRecord & smDelimiter & "    " & Trim$(tgMVef(ilVef).sName)
            Else
                slRecord = slRecord & smDelimiter
            End If
        Else 'Vehicle
            If ilVef <> -1 Then
                slRecord = slRecord & smDelimiter & Trim$(tgMVef(ilVef).sName)
            Else
                slRecord = slRecord & smDelimiter
            End If
        End If
        ilRdf = gBinarySearchRdf(tlReRate.iRdfCode)
        If ilRdf <> -1 Then 'Daypart
            slRecord = slRecord & smDelimiter & Trim$(tgMRdf(ilRdf).sName)
        Else
            slRecord = slRecord & smDelimiter
        End If
        '6/18/21 - JW - Task 7: Modify the ReRate report so that it includes the lineup settings when run with the "Lineup #" option checked on.
        slRecord = slRecord & smDelimiter & "ACT1Code=" & Trim$(tlReRate.sACT1LineupCode)
        'ACT 1 Settings
        slRecord = slRecord & " ACT1stored="
        If Trim$(tlReRate.sACT1StoredTime) <> "" Then slRecord = slRecord & "T"
        If Trim$(tlReRate.sACT1StoredSpots) <> "" Then slRecord = slRecord & "S"
        If Trim$(tlReRate.sACT1StoreClearPct) <> "" Then slRecord = slRecord & "C"
        If Trim$(tlReRate.sACT1DaypartFilter) <> "" Then slRecord = slRecord & "F"
    
        slRecord = slRecord & smDelimiter & mGetAudioTypes(tlReRate.sAudioType) 'TTP 10144 Audio Type
        slRecord = slRecord & smDelimiter & tlReRate.iLen 'Len
        
        'TTP 10192 Price Type
        If ckcPriceType.Value Then
            Select Case tlReRate.sPriceType
                Case "N": slRecord = slRecord & smDelimiter & "N/C"
                Case "M": slRecord = slRecord & smDelimiter & "MG"
                Case "B": slRecord = slRecord & smDelimiter & "Bonus"
                Case "S": slRecord = slRecord & smDelimiter & "Spinoff"
                Case "R": slRecord = slRecord & smDelimiter & "Recap"
                Case "P": slRecord = slRecord & smDelimiter & "Package"
                Case "A": slRecord = slRecord & smDelimiter & "ADU"
                Case "T": slRecord = slRecord & smDelimiter & "Paid" 'T=True, meaning it was Paid
                Case Else: slRecord = slRecord & smDelimiter
            End Select
        Else
            slRecord = slRecord & smDelimiter
        End If
        
    ElseIf (tlReRate.sType = "B" And tlReRate.sSubType = "B") Then
        slRowType = "BL"
        slRecord = slRecord & smDelimiter & "Bonus"  'Line#
        'slRecord = slRecord & ","       'Pkg Line #
        ilVef = gBinarySearchVef(tlReRate.iVefCode)
        If ilVef <> -1 Then
            slRecord = slRecord & smDelimiter & Trim$(tgMVef(ilVef).sName)
        Else
            slRecord = slRecord & smDelimiter
        End If
        If rbcBonus(0).Value Then
            slRecord = slRecord & smDelimiter       'Daypart
        Else
            ilRdf = gBinarySearchRdf(tlReRate.iRdfCode)
            If ilRdf <> -1 Then
                slRecord = slRecord & smDelimiter & Trim$(tgMRdf(ilRdf).sName)
            Else
                slRecord = slRecord & smDelimiter
            End If
        End If
        slRecord = slRecord & smDelimiter       'Act1 Lineup code
        slRecord = slRecord & smDelimiter & mGetAudioTypes(tlReRate.sAudioType) 'TTP 10144 Audio Type
        slRecord = slRecord & smDelimiter & tlReRate.iLen 'Len
        slRecord = slRecord & smDelimiter ''TTP 10192 Price Type (Bonus)
        
    ElseIf tlReRate.sSubType = "M" Then
        slRowType = tlReRate.sType & "L"
        slRecord = slRecord & smDelimiter & tlReRate.iLineNo  'Line#
        'If tlReRate.sType = "H" Then
        '    slRecord = slRecord & "," & tlReRate.iPkLineNo
        'Else
        '    slRecord = slRecord & ","
        'End If
        ilVef = gBinarySearchVef(tlReRate.iVefCode)
        If ilVef <> -1 Then
            slRecord = slRecord & smDelimiter & "        MG: " & Trim$(tgMVef(ilVef).sName)
        Else
            slRecord = slRecord & smDelimiter
        End If
        'slRecord = slRecord & ","       'Daypart
        ilRdf = gBinarySearchRdf(tlReRate.iRdfCode)
        If ilRdf <> -1 Then
            slRecord = slRecord & smDelimiter & Trim$(tgMRdf(ilRdf).sName)
        Else
            slRecord = slRecord & smDelimiter
        End If
        slRecord = slRecord & smDelimiter       'ACT1 Lineup Code
        slRecord = slRecord & smDelimiter & mGetAudioTypes(tlReRate.sAudioType) 'TTP 10144 Audio Type
        slRecord = slRecord & smDelimiter & tlReRate.iLen 'Len
        slRecord = slRecord & smDelimiter 'TTP 10192 Price Type

    Else
        slRecord = slRecord & smDelimiter   'Line #
        'Vehicle column
        blClearTotalFields = True
        If tlReRate.sSubType = "C" Then
            slRecord = slRecord & smDelimiter & "Contract Total"       'Line#
            slRowType = "CT"
        ElseIf tlReRate.sSubType = "T" Then
            slRecord = slRecord & smDelimiter & "Contract Bonus Total"       'Line#
            slRowType = "CB"
        ElseIf tlReRate.sSubType = "B" Then
            slRecord = slRecord & smDelimiter & "Contract Total with Bonus"       'Line#
            slRowType = "CS"
        ElseIf tlReRate.sSubType = "A" Then
            slRecord = slRecord & smDelimiter & "Advertiser Total"       'Line#
            slRowType = "AT"
        ElseIf tlReRate.sSubType = "N" Then
            slRecord = slRecord & smDelimiter & "Advertiser Bonus Total"       'Line#
            slRowType = "AB"
        ElseIf tlReRate.sSubType = "G" Then
            slRecord = slRecord & smDelimiter & "Advertiser Total with Bonus"       'Line#
            slRowType = "AS"
        Else
            blClearTotalFields = False
            slRecord = slRecord & smDelimiter
        End If
        slRecord = slRecord & smDelimiter       'Daypart
        slRecord = slRecord & smDelimiter       'ACT1 Lineup Code
        slRecord = slRecord & smDelimiter & mGetAudioTypes(gStripChr0(tlReRate.sAudioType)) 'TTP 10144 Audio Type
        slRecord = slRecord & smDelimiter & tlReRate.iLen 'Len
        slRecord = slRecord & smDelimiter ''TTP 10192 Price Type
    End If
        
    ilStart = 0
    ilEnd = 1
    If tlReRate.sType = "T" Or tlReRate.sType = "B" Or tlReRate.sSubType = "M" Then
        If (tlReRate.sType = "T") Or (tlReRate.sSubType = "C") Then
        Else
            ilStart = 1
            slRecord = slRecord & smDelimiter   'Rate
            slRecord = slRecord & smDelimiter 'TTP 10193 Line Comment
            slRecord = slRecord & smDelimiter   'Extended
            slRecord = slRecord & smDelimiter   'Units
            slRecord = slRecord & smDelimiter   'AQH
            slRecord = slRecord & smDelimiter   'Rtg
            slRecord = slRecord & smDelimiter   'CPM
            slRecord = slRecord & smDelimiter   'CPP
            slRecord = slRecord & smDelimiter   'Grimp
            slRecord = slRecord & smDelimiter   'GRP
            slRecord = slRecord & smDelimiter   'Book Name
        End If
    End If
    ilColorCode = 0
    For ilSpotType = ilStart To ilEnd Step 1
        If (rbcShow(0).Value = True) And ((tlReRate.sType = "O") Or (tlReRate.sType = "A") Or (tlReRate.sType = "E")) Then
            slRecord = slRecord & smDelimiter   'Rate
            slRecord = slRecord & smDelimiter 'TTP 10193 Line Comment
            slRecord = slRecord & smDelimiter   'Extended
            slRecord = slRecord & smDelimiter   'Units
            slRecord = slRecord & smDelimiter   'AQH
            slRecord = slRecord & smDelimiter   'Rtg
            slRecord = slRecord & smDelimiter   'CPM
            slRecord = slRecord & smDelimiter   'CPP
            slRecord = slRecord & smDelimiter   'Grimp
            slRecord = slRecord & smDelimiter   'GRP
            slRecord = slRecord & smDelimiter   'Book Name
        Else
            If ilSpotType = 0 Then
                'Rate
                If (Not blClearTotalFields) And (tlReRate.lTotalSpots(ilSpotType) > 0) Then
                    'slRecord = slRecord & smDelimiter & gLongToStrDec(tlReRate.lTotalCost(ilSpotType) / tlReRate.lTotalSpots(ilSpotType), 2)
                    slRecord = slRecord & smDelimiter & gLongToStrDec(tlReRate.dTotalCost(ilSpotType) / tlReRate.lTotalSpots(ilSpotType), 2) 'TTP 10439 - Rerate 21,000,000
                Else
                    slRecord = slRecord & smDelimiter
                End If
                If tlReRate.sType = "H" Or tlReRate.sType = "S" Or tlReRate.sType = "O" Or tlReRate.sType = "A" Or tlReRate.sType = "E" Then
                    slRecord = slRecord & smDelimiter & Trim(gStripChr0(tlReRate.sLineComment)) 'TTP 10193 Line Comment
                Else
                    slRecord = slRecord & smDelimiter 'TTP 10193 Line Comment
                End If
            End If
            'Extended Totals
            'slRecord = slRecord & smDelimiter & gLongToStrDec(tlReRate.lTotalCost(ilSpotType), 2)
            slRecord = slRecord & smDelimiter & gDblToStrDec(tlReRate.dTotalCost(ilSpotType), 2) 'TTP 10439 - Rerate 21,000,000
            'Units
            slRecord = slRecord & smDelimiter & tlReRate.lTotalSpots(ilSpotType)
            If ilSpotType = 1 And rbcReRateBook(2).Value = True Then
                slRecord = slRecord & smDelimiter   'AQH
                slRecord = slRecord & smDelimiter   'Rtg
                slRecord = slRecord & smDelimiter   'CPM
                slRecord = slRecord & smDelimiter   'CPP
                slRecord = slRecord & smDelimiter   'Grimp
                slRecord = slRecord & smDelimiter   'GRP
                slRecord = slRecord & smDelimiter   'Book Name
                slRecord = slRecord & smDelimiter   'Index
                'blIncludeFormula = True
                'llFormulaSpots = tlReRate.lTotalSpots(ilSpotType)
                'llFormulaAQH = tlReRate.lTotalAvgAud(ilSpotType)
                'llFormulaPop = tlReRate.lPop(ilSpotType)
                'llFormulaExtTotal = tlReRate.lTotalCost(ilSpotType)
            Else
                'AQH
                If Not blClearTotalFields Then
                    'slRecord = slRecord & "," & tlReRate.lTotalAvgAud(ilSpotType)
                    slRecord = slRecord & smDelimiter & gLongToStrDec(tlReRate.lTotalAvgAud(ilSpotType), imNumberDecPlaces)
                Else
                    slRecord = slRecord & smDelimiter
                End If
                'Rating
                If Not blClearTotalFields Then
                    'slRecord = slRecord & "," & tlReRate.iTotalAvgRating(ilSpotType)
                    If sm1or2PlaceRating <> "2" Then
                        slRecord = slRecord & smDelimiter & gIntToStrDec(tlReRate.iTotalAvgRating(ilSpotType), 1)
                    Else
                        'JW 10/14/21 - Fix v81 TTP 10258 issues per Jason Email: Thu 10/14/21 10:13 AM (#3)
                        slRecord = slRecord & smDelimiter & gIntToStrDec(tlReRate.iTotalAvgRating(ilSpotType), 2)
                    End If
                Else
                    slRecord = slRecord & smDelimiter
                End If
                'CPM
                slRecord = slRecord & smDelimiter & gLongToStrDec(tlReRate.lTotalCPM(ilSpotType), 2)
                'CPP
                slRecord = slRecord & smDelimiter & tlReRate.lTotalCPP(ilSpotType)
                'slRecord = slRecord & "," & tlReRate.lPop(ilSpotType)
                'Grimp
                ''slRecord = slRecord & "," & tlReRate.lTotalGrimps(ilSpotType)
                'If tgSpf.sSAudData = "H" Then
                '    slRecord = slRecord & "," & gLongToStrDec(tlReRate.lTotalGrimps(ilSpotType), 1)
                'ElseIf tgSpf.sSAudData = "N" Then
                '    slRecord = slRecord & "," & gLongToStrDec(tlReRate.lTotalGrimps(ilSpotType), 2)
                'ElseIf tgSpf.sSAudData = "U" Then
                '    slRecord = slRecord & "," & gLongToStrDec(tlReRate.lTotalGrimps(ilSpotType), 3)
                'Else
                '    slRecord = slRecord & "," & Trim$(str$(tlReRate.lTotalGrimps(ilSpotType)))
                'End If
                slRecord = slRecord & smDelimiter & gLongToStrDec(tlReRate.lTotalGrimps(ilSpotType), imNumberDecPlaces)
            
                slRecord = slRecord & smDelimiter & gLongToStrDec(tlReRate.lTotalGRP(ilSpotType), 1)
                'Book Name
                ilDnf = -1
                'For ilLoop = 0 To lbcBookNames.ListCount - 1 Step 1
                '    If tlReRate.iDnfCode(ilSpotType) = lbcBookNames.ItemData(ilLoop) Then
                '        ilDnf = ilLoop
                '        Exit For
                '    End If
                'Next ilLoop
                'If slRowType = "CT" Or slRowType = "CS" Then
                If (tlReRate.sType <> "T") Then
                    If (ilSpotType = 1) And ((tlReRate.sType = "O") Or (tlReRate.sType = "A") Or (tlReRate.sType = "E")) Then
                    Else
                        For illoop = 0 To UBound(tgBookInfo) - 1 Step 1
                            If tlReRate.iDnfCode(ilSpotType) = tgBookInfo(illoop).iDnfCode Then
                                ilDnf = illoop
                                Exit For
                            End If
                        Next illoop
                    End If
                End If
                If ilDnf <> -1 Then
                    'slRecord = slRecord & "," & Trim$(lbcBookNames.List(ilDnf))
                    slRecord = slRecord & smDelimiter & Trim$(tgBookInfo(ilDnf).sName)
                Else
                    'Add: Get book name
                    'If tlReRate.iDnfCode(ilSpotType) = -2 Then
                    If ilSpotType = 1 Then
                        'blUsingClosest = False
                        'If bmBookByLine Then
                        '    For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                        '        If tgBookByLineAssigned(ilBook).lClfCode = tlReRate.lClfCode(1) Then
                        '            If tgBookByLineAssigned(ilBook).lReRateDnfCode = -2 Then 'Closest
                        '                blUsingClosest = True
                        '            End If
                        '            Exit For
                        '        End If
                        '    Next ilBook
                        'Else
                        '    If rbcReRateBook(1).Value Then
                        '        blUsingClosest = True
                        '    End If
                        'End If
                        'If ilSpotType = 0 Then
                        '    blUsingClosest = False
                        'End If
                        'If ((tlReRate.sType = "O") Or (tlReRate.sType = "A") Or (tlReRate.sType = "E")) And (blUsingClosest = False) Then
                        If ((tlReRate.sType = "O") Or (tlReRate.sType = "A") Or (tlReRate.sType = "E")) Then
                            For ilBook = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
                                If (tmReRate(ilBook).iPkLineNo = tlReRate.iLineNo) And (tmReRate(ilBook).lCntrNo = tlReRate.lCntrNo) And (tmReRate(ilBook).sType = "H") Then
                                    'If tmReRate(ilBook).iDnfCode(ilSpotType) = -2 Then
                                        'slBook = ""
                                        'Exit For
                                        For ilClf = 0 To UBound(tmReRateBookDnfCodes) - 1 Step 1
                                            If (tmReRateBookDnfCodes(ilClf).lChfCode = tmReRate(ilBook).lChfCode(ilSpotType)) And (tmReRateBookDnfCodes(ilClf).iLineNo = tmReRate(ilBook).iLineNo) Then
                                                ilDnf = -1
                                                For illoop = 0 To UBound(tgBookInfo) - 1 Step 1
                                                    If tmReRateBookDnfCodes(ilClf).iDnfCode = tgBookInfo(illoop).iDnfCode Then
                                                        ilDnf = illoop
                                                        Exit For
                                                    End If
                                                Next illoop
                                                If ilDnf <> -1 Then
                                                    If slBook = "" Then
                                                        slBook = ":" & Trim$(tgBookInfo(ilDnf).sName) & ":"
                                                    Else
                                                        If InStr(1, slBook, ":" & Trim$(tgBookInfo(ilDnf).sName) & ":") = 0 Then
                                                            slBook = slBook & Trim$(tgBookInfo(ilDnf).sName) & ":"
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next ilClf
                                    'Else
                                    '    For ilClf = 0 To UBound(tmReRateBookDnfCodes) - 1 Step 1
                                    '        ilDnf = -1
                                    '        If (tmReRateBookDnfCodes(ilClf).lChfCode = tmReRate(ilBook).lChfCode) And (tmReRateBookDnfCodes(ilClf).iLineNo = tmReRate(ilBook).iLineNo) Then
                                    '            For ilLoop = 0 To UBound(tgBookInfo) - 1 Step 1
                                    '                If tmReRate(ilBook).iDnfCode(ilSpotType) = tgBookInfo(ilLoop).iDnfCode Then
                                    '                    ilDnf = ilLoop
                                    '                    Exit For
                                    '                End If
                                    '            Next ilLoop
                                    '            If ilDnf <> -1 Then
                                    '                If slBook = "" Then
                                    '                    slBook = ":" & Trim$(tgBookInfo(ilDnf).sName) & ":"
                                    '                Else
                                    '                    If InStr(1, slBook, ":" & Trim$(tgBookInfo(ilDnf).sName) & ":") = 0 Then
                                    '                        slBook = slBook & Trim$(tgBookInfo(ilDnf).sName) & ":"
                                    '                    End If
                                    '                End If
                                    '            End If
                                    '        End If
                                    '    Next ilClf
                                    'End If
                                End If
                            Next ilBook
                            If slBook <> "" Then
                                slBook = Mid$(slBook, 2)
                                slBook = Left$(slBook, Len(slBook) - 1)
                                slRecord = slRecord & smDelimiter & slBook
                            Else
                                'slRecord = slRecord & "," & "Mixture"
                                slRecord = slRecord & smDelimiter & ""
                            End If
                        Else
                            'If blUsingClosest Then
                                slBook = ""
                                For ilClf = 0 To UBound(tmReRateBookDnfCodes) - 1 Step 1
                                    blBuildBooks = False
                                    If tlReRate.sType <> "T" Then
                                        If (tmReRateBookDnfCodes(ilClf).lChfCode = tlReRate.lChfCode(ilSpotType)) And (tmReRateBookDnfCodes(ilClf).iLineNo = tlReRate.iLineNo) And (tmReRateBookDnfCodes(ilClf).sType = "S") Then
                                            blBuildBooks = True
                                        End If
                                    Else
                                        If slRowType = "CT" Or slRowType = "CS" Then
                                            If (tmReRateBookDnfCodes(ilClf).lChfCode = tlReRate.lChfCode(ilSpotType)) Then
                                                blBuildBooks = True
                                            End If
                                        ElseIf slRowType = "AT" Or slRowType = "AS" Then
                                            blBuildBooks = True
                                        End If
                                    End If
                                    If blBuildBooks Then
                                        ilDnf = -1
                                        For illoop = 0 To UBound(tgBookInfo) - 1 Step 1
                                            If tmReRateBookDnfCodes(ilClf).iDnfCode = tgBookInfo(illoop).iDnfCode Then
                                                ilDnf = illoop
                                                Exit For
                                            End If
                                        Next illoop
                                        If ilDnf <> -1 Then
                                            If slBook = "" Then
                                                slBook = ":" & Trim$(tgBookInfo(ilDnf).sName) & ":"
                                            Else
                                                If InStr(1, slBook, ":" & Trim$(tgBookInfo(ilDnf).sName) & ":") = 0 Then
                                                    slBook = slBook & Trim$(tgBookInfo(ilDnf).sName) & ":"
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ilClf
                                If slBook <> "" And tlReRate.sType = "T" Then
                                    If InStr(1, slBook, ":") > 0 Then
                                        slBook = ""
                                    End If
                                End If
                                If slBook <> "" Then
                                    slBook = Mid$(slBook, 2)
                                    slBook = Left$(slBook, Len(slBook) - 1)
                                    slRecord = slRecord & smDelimiter & slBook
                                Else
                                    'slRecord = slRecord & "," & "Mixture"
                                    slRecord = slRecord & smDelimiter & ""
                                End If
                            'Else
                            '    slRecord = slRecord & "," & "Mixture"
                            'End If
                        End If
                    Else
                        'ilSpotType <> 1
                        'TTP 9912 - ReRate report when using the show package only or show package and hidden options, if there's a mix of purchased books
                        If ilSpotType = 0 And tlReRate.sType = "O" Then 'get purchased package book(s) for Purchase "O" Ordered Lines (Packages)
                            'Get purchase book(s) - from the CLF record, make ":" separated list of each Packgage vehicle's Purchase Book Names
                            slSQLQuery = ""
                            slSQLQuery = slSQLQuery & "SELECT dnfBookName "
                            slSQLQuery = slSQLQuery & " FROM ""DNF_Demo_Rsrch_Names"" "
                            slSQLQuery = slSQLQuery & " WHERE dnfCode IN ("
                            slSQLQuery = slSQLQuery & "     SELECT distinct(clfdnfCode) "
                            slSQLQuery = slSQLQuery & "     FROM ""CLF_Contract_Line"" "
                            slSQLQuery = slSQLQuery & "     WHERE clfChfCode = " & tlReRate.lChfCode(0)
                            slSQLQuery = slSQLQuery & "           AND clfPkLineNo = " & tlReRate.iLineNo
                            slSQLQuery = slSQLQuery & " )"
                            
                            Set tmp_rst = gSQLSelectCall(slSQLQuery)
                            Do While Not tmp_rst.EOF
                                If slBook = "" Then
                                    slBook = ":" & Trim(tmp_rst!dnfBookName) & ":"
                                Else
                                    slBook = slBook & Trim(tmp_rst!dnfBookName) & ":"
                                End If
                                tmp_rst.MoveNext
                            Loop
                        End If
                        If slBook <> "" Then
                            slRecord = slRecord & smDelimiter & Mid(slBook, 2, Len(slBook) - 2) 'book Name
                        Else
                            slRecord = slRecord & smDelimiter
                        End If
                        slBook = ""
                    End If
                End If
        
                'Compute Index from Gross Impressions or GRP
                If ilStart = 0 And ilSpotType = 1 Then
                    If rbcIndex(1).Value Then
                        slStr = Format(gDivStr(gMulStr("100", str(tlReRate.lTotalGRP(1)) & ".00"), str(tlReRate.lTotalGRP(0)) & ".00"), "#0.00")
                    Else
                        slStr = Format(gDivStr(gMulStr("100", str(tlReRate.lTotalGrimps(1)) & ".00"), str(tlReRate.lTotalGrimps(0)) & ".00"), "#0.00")
                    End If
                    If Val(slStr) < 100 Then
                        'slStr = gRoundStr(slStr, ".01", 1)
                        'gFormatStr slStr, 0, 1, slStr
                        gFormatStr slStr, 0, 2, slStr
                        ilColorCode = 1
                    ElseIf Val(slStr) >= 100 Then
                        gFormatStr slStr, 0, 2, slStr
                        ilColorCode = 2
                    End If
                    slRecord = slRecord & smDelimiter & slStr
                ElseIf tlReRate.sSubType = "M" And ilSpotType = 1 Then
                    slStr = Format(gDivStr(gMulStr("100", str(tlReRate.lTotalAvgAud(1)) & ".00"), str(tlReRate.lTotalAvgAud(0)) & ".00"), "#0.00")
                    slRecord = slRecord & smDelimiter & slStr
                    If Val(slStr) < 100 Then
                        ilColorCode = 1
                    ElseIf Val(slStr) >= 100 Then
                        ilColorCode = 2
                    End If
                End If
            End If
        End If
    Next ilSpotType
    'Print #hmToCSV, slRecord
    
    'mPrint slRecord, smDelimiter, slRowType, tlReRate.lPop(1), tlReRate.lTotalCost(1)
    mPrint slRecord, smDelimiter, slRowType, tlReRate.lPop(1), tlReRate.dTotalCost(1) 'TTP 10439 - Rerate 21,000,000
    
    'If rbcReRateBook(2).Value And blIncludeFormula Then
    '    'Add formula in cells not set
    '    If llFormulaSpots > 0 Then
    '        ''ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & smColumnLetter(RUNITEXCEL) & imExcelRow - 1 & "*" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1, imExcelRow - 1, RGIMPEXCEL) ', slDelimiter)
    '        'ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & smColumnLetter(RUNITEXCEL) & imExcelRow - 1 & "*" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "," & """" & """" & ")", imExcelRow - 1, RGIMPEXCEL) ', slDelimiter)
    '        'If rbcIndex(0).Value Then
    '        '    'Index by GImp
    '        '    'ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & "((" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0)" & " AND " & "(" & smColumnLetter(PAQHEXCEL) & imExcelRow - 1 & "> 0))" & "," & "100*" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "/" & smColumnLetter(PGIMPEXCEL) & imExcelRow - 1 & "," & """" & """" & ")", imExcelRow - 1, INDEXEXCEL) ', slDelimiter)
    '        '    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "IF(" & smColumnLetter(PAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "100*" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "/" & smColumnLetter(PGIMPEXCEL) & imExcelRow - 1 & "," & """" & """" & ")" & "," & """" & """" & ")", imExcelRow - 1, INDEXEXCEL) ', slDelimiter)
    '        'End If
    '        'If llFormulaPop > 0 Then
    '        '    'ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "*" & "100" & "/" & llFormulaPop, imExcelRow - 1, RGRPEXCEL) ', slDelimiter)
    '        '    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "*" & "100" & "/" & llFormulaPop & "," & """" & """" & ")", imExcelRow - 1, RGRPEXCEL) ', slDelimiter)
    '        '    If rbcIndex(1).Value Then
    '        '        'Index by GRG
    '        '        'ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "100*" & smColumnLetter(RGRPEXCEL) & imExcelRow - 1 & "/" & smColumnLetter(PGRPEXCEL) & imExcelRow - 1 & "," & """" & """" & ")", imExcelRow - 1, INDEXEXCEL) ', slDelimiter)
    '        '        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "IF(" & smColumnLetter(PAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "100*" & smColumnLetter(RGRPEXCEL) & imExcelRow - 1 & "/" & smColumnLetter(PGRPEXCEL) & imExcelRow - 1 & "," & """" & """" & ")" & "," & """" & """" & ")", imExcelRow - 1, INDEXEXCEL) ', slDelimiter)
    '        '    End If
    '        '    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "*" & "100" & "/" & llFormulaPop & "," & """" & """" & ")", imExcelRow - 1, RRTGEXCEL) ', slDelimiter)
    '        '    If llFormulaExtTotal >= 0 And ckcCPM.Value = vbChecked Then
    '        '        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & llFormulaExtTotal & "*" & llFormulaPop & "/" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "/" & "10000" & "," & """" & """" & ")", imExcelRow - 1, RCPPEXCEL) ', slDelimiter)
    '        '        'ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & llFormulaExtTotal & "/" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "/" & "100", imExcelRow - 1, RCPMEXCEL) ', slDelimiter)
    '        '        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & llFormulaExtTotal & "/" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "/" & "100" & "," & """" & """" & ")", imExcelRow - 1, RCPMEXCEL) ', slDelimiter)
    '        '    End If
    '        'End If
    '        mExcelFormulaSetting llFormulaSpots, llFormulaPop, llFormulaExtTotal
    '    End If
    'End If
    If ExptReRate.ckcCsv.Value = vbUnchecked Then
        If (tlReRate.sType = "T") Or ((tlReRate.sType = "B") And ((tlReRate.sSubType = "T") Or (tlReRate.sSubType = "N"))) Then
            ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1)
        End If
        If tlReRate.sSubType = "M" Then
            ilRet = gExcelOutputGeneration("FI", omBook, omSheet, , "True", imExcelRow - 1)
        End If
        'Set Index column bold
        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, omSheet.UsedRange.Columns.Count)
        If ilColorCode = 1 Then
            ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , str(Red), imExcelRow - 1, omSheet.UsedRange.Columns.Count)
        ElseIf ilColorCode = 2 Then
            ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , str(DARKGREEN), imExcelRow - 1, omSheet.UsedRange.Columns.Count)
        End If
    End If
    ''Set Foreground color:
    'If rbcColor(1).Value Then
    '    If ilColorCode = 1 Then
    '        For ilCol = imReRateColumn To omSheet.UsedRange.Columns.Count - 1 Step 1
    '            ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , str(RED), imExcelRow - 1, ilCol)
    '        Next ilCol
    '    ElseIf ilColorCode = 2 Then
    '        For ilCol = imReRateColumn To omSheet.UsedRange.Columns.Count - 1 Step 1
    '            ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , str(DARKGREEN), imExcelRow - 1, ilCol)
    '        Next ilCol
    '    End If
    'End If
    
End Sub

Private Function mGetAudioTypes(slAudioTypeCodes As String) As String
    'TTP 10144 get Audio Type  Bonuses could have Types(s)
    Dim illoop As Integer
    Dim slString As String
    mGetAudioTypes = ""
    'If ckcAudioType.Value = False Then Exit Function
    'Could be Multiple types
    For illoop = 1 To Len(Trim(slAudioTypeCodes))
        Select Case Mid(Trim(slAudioTypeCodes), illoop, 1)
            Case "R": slString = "RC"
            Case "M": slString = "LP"
            Case "S": slString = "RP"
            Case "P": slString = "PC"
            Case "Q": slString = "PP"
            Case "L": slString = "LC"
            Case Else: slString = Mid(Trim(slAudioTypeCodes), illoop, 1)
        End Select
        If slString <> "" Then
            If mGetAudioTypes <> "" Then mGetAudioTypes = mGetAudioTypes & ":"
            mGetAudioTypes = mGetAudioTypes & slString
        End If
    Next illoop
End Function

Private Sub mGetPurchasedSpotCount(ilClf As Integer, llStartDate As Long, llEndDate As Long, ilMnfDemo As Integer, ilTotLnSpts As Integer, llTotalCntrSpots As Long, llAvgAud As Long, llPopEst As Long, ilAudFromSource As Integer, llAudFromCode As Long)
    Dim ilCff As Integer
    Dim illoop As Integer
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim slStr As String
    Dim llDate As Long
    Dim llDate2 As Long
    Dim ilDay As Integer
    Dim ilSpots As Integer
    Dim ilUpperWk As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    ReDim ilInputDays(0 To 6) As Integer
    Dim ilRet As Integer
        
    If tmClfP.iStartTime(0) = 1 And tmClfP.iStartTime(1) = 0 Then
        llOvStartTime = 0
        llOvEndTime = 0
    Else
        'override times exist
        gUnpackTimeLong tmClfP.iStartTime(0), tmClfP.iStartTime(1), False, llOvStartTime
        gUnpackTimeLong tmClfP.iEndTime(0), tmClfP.iEndTime(1), True, llOvEndTime
    End If
    
    ilCff = tmClfPurchase(ilClf).iFirstCff
    Do While ilCff <> -1
        tmCff = tmCffPurchase(ilCff).CffRec
        For illoop = 0 To 6                 'init all days to not airing, setup for research results later
            ilInputDays(illoop) = False
        Next illoop
    
        gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
        llFltStart = gDateValue(slStr)
        'backup start date to Monday
        illoop = gWeekDayLong(llFltStart)
        Do While illoop <> 0
            llFltStart = llFltStart - 1
            illoop = gWeekDayLong(llFltStart)
        Loop
        gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
        llFltEnd = gDateValue(slStr)
        '
        'Loop thru the flight by week and build the number of spots for each week
        '
        For llDate2 = llFltStart To llFltEnd Step 7
            If llDate2 >= llStartDate And llDate2 <= llEndDate Then
                If tmCff.sDyWk = "W" Then            'weekly
                    ilSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                     For ilDay = 0 To 6 Step 1
                        If (llDate2 + ilDay >= llFltStart) And (llDate2 + ilDay <= llFltEnd) Then
                            If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                                ilInputDays(ilDay) = True
                            End If
                        End If
                     Next ilDay
                Else                                        'daily
                     If illoop + 6 < llFltEnd Then           'we have a whole week
                        ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                        For ilDay = 0 To 6 Step 1
                            If tmCff.iDay(ilDay) > 0 Then
                                ilInputDays(ilDay) = True
                            End If
                        Next ilDay
                     Else                                    'do partial week
                        For llDate = llDate2 To llFltEnd Step 1
                            ilDay = gWeekDayLong(llDate)
                            ilSpots = ilSpots + tmCff.iDay(ilDay)
                            If tmCff.iDay(ilDay) > 0 Then
                                ilInputDays(ilDay) = True
                            End If
                        Next llDate
                    End If
                End If

'For ilDay = 0 To 6 Step 1
'    Debug.Print "  --- mGetPurchasedSpotCount - " & IIF(ilInputDays(ilDay) = True, "ilDay" & ilDay & " = True", "ilDay" & ilDay & " = False")
'Next ilDay
                    
'Debug.Print "  getDemo1; RDFCode=" & tmClfP.iRdfCode
                ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmClfP.iDnfCode, tmClfP.iVefCode, 0, ilMnfDemo, llDate2, llDate2, tmClfP.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClfP.sType, tmClfP.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                ''Loop and build avg aud, spots, & spots per week
                ''calc week index
                'ilUpperWk = (llDate2 - llStartDate) / 7 + 1
                ''ilUpperWk = UBound(lmWklyspots)
                'lmWklyspots(ilUpperWk - 1) = ilSpots
                ilTotLnSpts = ilTotLnSpts + ilSpots
                If tmClfP.sType = "S" Then
                    llTotalCntrSpots = llTotalCntrSpots + ilSpots
                End If
                'lmWklyRates(ilUpperWk - 1) = tmCff.lActPrice
                'lmWklyAvgAud(ilUpperWk - 1) = llAvgAud
                'lmWklyPopEst(ilUpperWk - 1) = llPopEst
                'ilUpperWk = ilUpperWk + 1
                mAddSpotsToWeekArray llDate2, llStartDate, ilSpots, llAvgAud, tmCff.lActPrice, llPopEst, False

            End If                  'if llDate2 >= llStartDate and llDate2 <= llEndDAte
        Next llDate2
        ilCff = tmCffPurchase(ilCff).iNextCff               'get next flight record from mem
    Loop
End Sub
Private Sub mGetAiredSpotCount(ilClf As Integer, llStartDate As Long, llEndDate As Long, ilMnfDemo As Integer, ilTotLnSpts As Integer, llTotalCntrSpots As Long, llAvgAud As Long, llPopEst As Long, ilAudFromSource As Integer, llAudFromCode As Long)
'Debug.Print "mGetAiredSpotCount: " & Format(llStartDate, "ddddd") & " - " & Format(llEndDate, "ddddd") & " " & IIF(ckcTreatMGOsAsOrdered.Value = vbChecked, " TreatMG as Ordered", "")

    Dim llSdf As Long
    Dim ilCff As Integer
    Dim illoop As Integer
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim slStr As String
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilSpots As Integer
    Dim ilUpperWk As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llMoDate As Long
    ReDim ilInputDays(0 To 6) As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim blVefFd As Boolean
    Dim ilDnfCode As Integer
    Dim llTime As Long
    Dim ilRdfCode As Integer
    Dim ilBook As Integer
    Dim ilRdfIndex As Integer
    If tmClfR.iStartTime(0) = 1 And tmClfR.iStartTime(1) = 0 Then
        llOvStartTime = 0
        llOvEndTime = 0
    Else
        'override times exist
        gUnpackTimeLong tmClfR.iStartTime(0), tmClfR.iStartTime(1), False, llOvStartTime
        gUnpackTimeLong tmClfR.iEndTime(0), tmClfR.iEndTime(1), True, llOvEndTime
    End If
    
    ilSpots = 1
    imMGDetailDnfCode = -1
    imCntrLnDetailDnfCode = -1

    '--------------------
    'TTP 9922 treat makegoods/outsides as ordered, TTP 9954 - treat missed spots to be included as ordered
    If ckcTreatMGOsAsOrdered.Value = vbChecked Or ckcTreatMissedAsOrdered.Value = vbChecked Then
        For llSdf = LBound(tmSdfExt) To UBound(tmSdfExt) - 1 Step 1
            If ckcTreatMGOsAsOrdered.Value = vbChecked Then
                'TTP 10157 - "treate MGs as ordered" when checked on can affect bonus spots, resulting in bonuses being shown without dayparts and with no AQH
                'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates for MG/OS (Fix Parens around IF statement)
                If (tmSdfExt(llSdf).sSchStatus = "G" Or tmSdfExt(llSdf).sSchStatus = "O") And tmSdfExt(llSdf).sSpotType <> "X" Then
                    'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates AND VEHICLE for MG/OS
                    ''Subst Veh with Ordered Vehicle (Because this could have been Made Good with another Vehicle.   Put it back to the scheduled vehicle for the Treat as Ordered/Aired research
                    'For ilLoop = LBound(tmClfReRate) To UBound(tmClfReRate)
                    '    'TTP 10154 - subscript error when running report for contract with lines without spots with "treat MG/Outsides as ordered" and "treat missed spots as ordered" checked on
                    '    If UBound(tmSdfExt) >= ilLoop Then 'TTP 10253 - no harm in letting it get to end of array
                    '        If tmSdfExt(ilLoop).sSchStatus = "S" And tmSdfExt(ilLoop).iLineNo = tmSdfExt(llSdf).iLineNo Then
                    '            'Here's the Scheduled for this MG or Outside
                    '            tmSdfExt(llSdf).iVefCode = tmSdfExt(ilLoop).iVefCode
                    '            Exit For
                    '        End If
                    '    End If
                    'Next ilLoop
                    
                    'Find the Time from the Research Book
                    ilRdfIndex = gBinarySearchRdf(tmClfReRate(ilClf).ClfRec.iRdfCode)
                    If ilRdfIndex <> -1 Then
                        tmSdfExt(llSdf).iTime(0) = tgMRdf(ilRdfIndex).iStartTime(0, 6)
                        tmSdfExt(llSdf).iTime(1) = tgMRdf(ilRdfIndex).iStartTime(1, 6)
                    End If
                    
                    'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates for MG/OS
                    tlSmfSrchKey2.lCode = tmSdfExt(llSdf).lCode
                    ilRet = btrGetEqual(hlSmf, tlSmf, ilSmfRecLen, tlSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        'ilFound = True
                    Else
                        tmSdfExt(llSdf).iDate(0) = tlSmf.iMissedDate(0)
                        tmSdfExt(llSdf).iDate(1) = tlSmf.iMissedDate(1)
                        
                        'TTP 10329, Subst Veh with Ordered Vehicle using the SMF (Because this could have been Made Good with another Vehicle.   Put it back to the scheduled vehicle for the Treat as Ordered/Aired research
                        tmSdfExt(llSdf).iVefCode = tlSmf.iOrigSchVef
                    End If

                    'Flag as Scheduled
                    tmSdfExt(llSdf).sSchStatus = "S"
                   
                    'Clear Make good date
                    tmSdfExt(llSdf).lMdDate = 0
                End If
            End If
            If ckcTreatMissedAsOrdered.Value = vbChecked Then
                'TTP 10157 - "treate MGs as ordered" when checked on can affect bonus spots, resulting in bonuses being shown without dayparts and with no AQH
                If tmSdfExt(llSdf).sSchStatus = "M" And tmSdfExt(llSdf).sSpotType <> "X" Then
                    tmSdfExt(llSdf).sSchStatus = "S"
                    'Find the Time from the Research Book
                    ilRdfIndex = gBinarySearchRdf(tmClfReRate(ilClf).ClfRec.iRdfCode)
                    If ilRdfIndex <> -1 Then
                        tmSdfExt(llSdf).iTime(0) = tgMRdf(ilRdfIndex).iStartTime(0, 6)
                        tmSdfExt(llSdf).iTime(1) = tgMRdf(ilRdfIndex).iStartTime(1, 6)
                    End If
                End If
            End If
        Next llSdf
    End If
    
    
    
    For llSdf = LBound(tmSdfExt) To UBound(tmSdfExt) - 1 Step 1
        'TTP 10165 - 5/12/21
        'If tmSdfExt(llSdf).iDate(0) = -1 And tmSdfExt(llSdf).iDate(1) = -1 Then
        '    llDate = -1 'set Date to epoch's eve :-) -- so that the checks below will allow this "outside/Makegood" to be qualified
        'Else
        gUnpackDateLong tmSdfExt(llSdf).iDate(0), tmSdfExt(llSdf).iDate(1), llDate
        'End If
        '-----------------------------------------------------------
        'Get Aired Spots
        If (llDate >= llStartDate) And (llDate <= llEndDate) Then
            If lmReRatePop = -1 Then
                'get Population
                If (tmSdfExt(llSdf).sSchStatus = "S") And (tmSdfExt(llSdf).sSpotType <> "X") And (tmSdfExt(llSdf).lMdDate = 0) And (tmSdfExt(llSdf).iLineNo = tmClfR.iLine) Then
                    If rbcReRateBook(1).Value And Not bmBookByLine Then  'Or (rbcReRateBook(3).Value And imByCntrLnDnfCode = -2) Then 'Closest
                        ilDnfCode = mFindClosestBook(llDate, llSdf) 'tmSdfExt(llSdf).iVefCode)
'Debug.Print "  - Closest Book:" & ilDnfCode
                        ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, 0, ilMnfDemo, lmReRatePop)
                    'ElseIf rbcReRateBook(3).Value Then  'Contract Line
                    End If
                    If bmBookByLine Then
                        ilDnfCode = 0
                        For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                            If tgBookByLineAssigned(ilBook).lChfCode = tmSdfExt(llSdf).lChfCode Then
                                If (tgBookByLineAssigned(ilBook).iLineNo = tmSdfExt(llSdf).iLineNo) And (tgBookByLineAssigned(ilBook).iVefCode = tmSdfExt(llSdf).iVefCode) Then
                                    ilDnfCode = mGetDnfByContractLine(ilBook, llSdf)
'Debug.Print "  - Line Book:" & ilDnfCode
                                    ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, 0, ilMnfDemo, lmReRatePop)
                                    Exit For
                                End If
                            End If
                        Next ilBook
                    Else
                        If Not rbcReRateBook(1).Value Then
                            ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, imReRateDnfCode, 0, ilMnfDemo, lmReRatePop)
                        End If
                    End If
                End If
            End If
            
            If (tmSdfExt(llSdf).iLineNo = tmClfR.iLine) Then
                'get Book
                ilDnfCode = imReRateDnfCode
'Debug.Print "  - Get Book:" & ilDnfCode
                If rbcReRateBook(1).Value And Not bmBookByLine Then  'Or (rbcReRateBook(3).Value And imByCntrLnDnfCode = -2) Then 'Closest
                    ilDnfCode = mFindClosestBook(llDate, llSdf) 'tmSdfExt(llSdf).iVefCode)
                    'imReRateDnfCode = ilDnfCode
                End If
                'If rbcReRateBook(3).Value = True Then   'Contract Line
                If bmBookByLine Then
                    For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                        If tgBookByLineAssigned(ilBook).lChfCode = tmClfR.lChfCode Then
                            If tgBookByLineAssigned(ilBook).iLineNo = tmClfR.iLine And tgBookByLineAssigned(ilBook).sType = tmClfR.sType Then
                                ilDnfCode = mGetDnfByContractLine(ilBook, llSdf)
'Debug.Print "  - Get Book:" & ilDnfCode
                                Exit For
                            End If
                        End If
                    Next ilBook
                End If
            End If
        End If
        
        '----------------------------
        'Get Demo Avg Aud, Spot Count (For Scheduled spots)
        If (llDate = -1 And tmSdfExt(llSdf).iLineNo = tmClfR.iLine) Or ((llDate >= llStartDate) And (llDate <= llEndDate) And (tmSdfExt(llSdf).sSchStatus = "S") And (tmSdfExt(llSdf).sSpotType <> "X") And (tmSdfExt(llSdf).lMdDate = 0) And (tmSdfExt(llSdf).iLineNo = tmClfR.iLine)) Then
            llMoDate = llDate
            '---------------------------------------------------------------------------
            'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates for MG/OS
            ''TTP 10253 / FIX TTP 10165 - ReRate - when "treat MG/Outsides as ordered" - restore the StartDate (that was -1'd)
            'If ckcTreatMGOsAsOrdered.Value = vbChecked Then
            '    gUnpackDateLong tmClfReRate(ilClf).ClfRec.iStartDate(0), tmClfReRate(ilClf).ClfRec.iStartDate(1), llMoDate
            'End If

            illoop = gWeekDayLong(llMoDate)
            Do While illoop <> 0
                llMoDate = llMoDate - 1
                illoop = gWeekDayLong(llMoDate)
            Loop
                       
'Debug.Print "  - Get Demo Avg Aud, Spot Count (For Scheduled spots)"
            ilCff = tmClfReRate(ilClf).iFirstCff
            Do While ilCff <> -1
                tmCff = tmCffReRate(ilCff).CffRec
                For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                    ilInputDays(illoop) = False
                Next illoop
            
                gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                llFltStart = gDateValue(slStr)
                'backup start date to Monday
                illoop = gWeekDayLong(llFltStart)
                Do While illoop <> 0
                    llFltStart = llFltStart - 1
                    illoop = gWeekDayLong(llFltStart)
                Loop
                gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                llFltEnd = gDateValue(slStr)
                
                'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates for MG/OS
                If llDate >= llFltStart And llDate <= llFltEnd Then
                    If tmCff.sDyWk = "W" Then            'weekly
                        For ilDay = 0 To 6 Step 1
                           If (llMoDate + ilDay >= llFltStart) And (llMoDate + ilDay <= llFltEnd) Then
                               If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                                   ilInputDays(ilDay) = True
                                   
                               End If
                           End If
                        Next ilDay
                    Else                                        'daily
                         If illoop + 6 < llFltEnd Then           'we have a whole week
                            For ilDay = 0 To 6 Step 1
                                If tmCff.iDay(ilDay) > 0 Then
                                    ilInputDays(ilDay) = True
                                End If
                            Next ilDay
                         Else                                    'do partial week
                            For llDate = llDate To llFltEnd Step 1
                                ilDay = gWeekDayLong(llDate)
                                If tmCff.iDay(ilDay) > 0 Then
                                    ilInputDays(ilDay) = True
                                End If
                            Next llDate
                        End If
                    End If
                    If imCntrLnDetailDnfCode = -1 Then
                        imCntrLnDetailDnfCode = ilDnfCode
                    Else
                        If imCntrLnDetailDnfCode <> ilDnfCode And imCntrLnDetailDnfCode <> -2 Then
                            imCntrLnDetailDnfCode = -2
                        End If
                    End If
                    
                    '---------------------------------------------------------------------------
                    'TTP 10329, TTP 10330 - JW - 10/29/21; Add SMF_Spot_MG_Specs, to lookup Missed Dates for MG/OS
                    ''TTP 10253 / FIX TTP 10165 - ReRate - when "treat MG/Outsides as ordered" - restore the StartDate (that was -1'd)
                    'If ckcTreatMGOsAsOrdered.Value = vbChecked Then
                        'gUnpackDateLong tmClfReRate(ilClf).ClfRec.iStartDate(0), tmClfReRate(ilClf).ClfRec.iStartDate(1), llDate
                        'ilInputDays(gWeekDayLong(llDate)) = True
                    'End If
'Debug.Print "  - mSaveReRateBookByLine:" & ilDnfCode

                    mSaveReRateBookByLine llSdf, ilDnfCode

'Debug.Print "  getDemo2; RDFCode=" & tmClfP.iRdfCode & " - Date:" & llDate


                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, tmClfR.iVefCode, 0, ilMnfDemo, llDate, llDate, tmClfR.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClfR.sType, tmClfR.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                    ''Loop and build avg aud, spots, & spots per week
                    ''calc week index
                    'ilUpperWk = (llDate - llStartDate) / 7 + 1
                    ''ilUpperWk = UBound(lmWklyspots)
                    'If ilUpperWk - 1 > UBound(lmWklyspots) Then
                    '    ReDim Preserve lmWklyspots(0 To ilUpperWk) As Integer
                    '    ReDim Preserve lmWklyAvgAud(0 To ilUpperWk) As Long
                    '    ReDim Preserve lmWklyPopEst(0 To ilUpperWk) As Long
                    'End If
                    'lmWklyspots(ilUpperWk - 1) = lmWklyspots(ilUpperWk - 1) + ilSpots
                    ilTotLnSpts = ilTotLnSpts + ilSpots
                    If tmClfR.sType = "S" Then
                        llTotalCntrSpots = llTotalCntrSpots + ilSpots
                    End If
                    'lmWklyRates(ilUpperWk - 1) = tmCff.lActPrice
                    'If (ckcMG.Value = vbChecked) Then
                    '    lmWklyAvgAud(ilUpperWk - 1) = llAvgAud
                    'Else
                    '    lmWklyAvgAud(ilUpperWk - 1) = lmWklyAvgAud(ilUpperWk - 1) + llAvgAud
                    'End If
                    'lmWklyPopEst(ilUpperWk - 1) = llPopEst
                    If (ckcMG.Value = vbChecked Or ckcTreatMGOsAsOrdered.Value = vbChecked) Then
                        mAddSpotsToWeekArray llDate, llStartDate, ilSpots, llAvgAud, tmCff.lActPrice, llPopEst, False
                    Else
                        mAddSpotsToWeekArray llDate, llStartDate, ilSpots, llAvgAud, tmCff.lActPrice, llPopEst, True
                    End If
                    Exit Do
                End If
                ilCff = tmCffReRate(ilCff).iNextCff
            Loop
        Else
            

            '----------------------------------
            'Bonus?
            blVefFd = False
            'If (llDate >= llStartDate) And (llDate <= llEndDate) And (tmSdfExt(llSdf).iLineNo = tmClfReRate.iLine) Then
            If (((llDate >= llStartDate) And (llDate <= llEndDate)) Or (rbcDatesBy(3).Value = True)) And (tmSdfExt(llSdf).iLineNo = tmClfR.iLine) Then
                If (tmSdfExt(llSdf).sSpotType = "X") And (mTestShowFill(tmSdfExt(llSdf).sPriceType) = "Y") Then 'Bonus
'Debug.Print "  - Bonus"
                    If rbcReRateBook(0).Value Then  'Default vehicle
                        ilVef = gBinarySearchVef(tmSdfExt(llSdf).iVefCode)
                        If ilVef <> -1 Then
                            ilDnfCode = tgMVef(ilVef).iDnfCode
                        Else
                            ilDnfCode = 0
                        End If
                    End If
                    'If rbcReRateBook(3).Value Then  'Contract Line
                    If bmBookByLine Then
                        ilDnfCode = 0
                        For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                            If tgBookByLineAssigned(ilBook).lChfCode = tmSdfExt(llSdf).lChfCode Then
                                If (tgBookByLineAssigned(ilBook).iLineNo = tmSdfExt(llSdf).iLineNo) And (tgBookByLineAssigned(ilBook).iVefCode = tmSdfExt(llSdf).iVefCode) And (tgBookByLineAssigned(ilBook).iBonusCount > 0) Then
                                    ilDnfCode = mGetDnfByContractLine(ilBook, llSdf)
                                    Exit For
                                End If
                            End If
                        Next ilBook
                    End If
                    blVefFd = False
                    ilRdfCode = 0
                    If (rbcBonus(1).Value = True) Then
                        gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llTime
                        ilRdfCode = mFindDaypart(tmSdfExt(llSdf).iVefCode, llDate, llTime)
                    End If
                    For ilVef = 0 To UBound(tmBonusInfo) - 1 Step 1
                        If (rbcBonus(0).Value = True) Then
                            'TTP 10123
                            If (tmSdfExt(llSdf).iVefCode = tmBonusInfo(ilVef).iVefCode) And (ilDnfCode = tmBonusInfo(ilVef).iDnfCode) And (tmSdfExt(llSdf).iLen = tmBonusInfo(ilVef).iSpotLen) Then
                                blVefFd = True
                                Exit For
                            End If
                        Else
                            If (tmSdfExt(llSdf).iVefCode = tmBonusInfo(ilVef).iVefCode) And (ilDnfCode = tmBonusInfo(ilVef).iDnfCode) And (ilRdfCode = tmBonusInfo(ilVef).iRdfCode) And (tmSdfExt(llSdf).iLen = tmBonusInfo(ilVef).iSpotLen) Then
                                blVefFd = True
                                'TTP 10144
                                If InStr(1, tmBonusInfo(ilVef).sAudioTypes, tmClfR.sLiveCopy) = 0 Then
                                    tmBonusInfo(ilVef).sAudioTypes = Trim(tmBonusInfo(ilVef).sAudioTypes) & tmClfR.sLiveCopy
                                End If
                                
                                Exit For
                            End If
                        End If
                    Next ilVef
                    If Not blVefFd Then
                        tmBonusInfo(UBound(tmBonusInfo)).iVefCode = tmSdfExt(llSdf).iVefCode
                        tmBonusInfo(UBound(tmBonusInfo)).iDnfCode = ilDnfCode
                        tmBonusInfo(UBound(tmBonusInfo)).iRdfCode = ilRdfCode
                        tmBonusInfo(UBound(tmBonusInfo)).iSpotLen = tmSdfExt(llSdf).iLen 'TTP 10123
                        tmBonusInfo(UBound(tmBonusInfo)).sAudioTypes = tmClfR.sLiveCopy 'TTP 10144
                        ReDim Preserve tmBonusInfo(0 To UBound(tmBonusInfo) + 1) As MGBONUSINFO
                    End If
                'ElseIf (tmSdfExt(llSdf).sSchStatus = "G" Or tmSdfExt(llSdf).sSchStatus = "O") And (ckcMG.Value = vbChecked) Then
                '------------------------------------
                'Make Good / Outside?
                ElseIf (tmSdfExt(llSdf).sSpotType <> "X") And ((tmSdfExt(llSdf).sSchStatus = "G" Or tmSdfExt(llSdf).sSchStatus = "O")) Then

'Debug.Print "  - Make Good / Outside"
                    If (ckcMG.Value = vbChecked Or ckcTreatMGOsAsOrdered.Value = vbChecked) Then
'Debug.Print "  -> MG/OS as Ordered Checked"
                    
                        gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llTime
                        ilRdfCode = mFindDaypart(tmSdfExt(llSdf).iVefCode, llDate, llTime)
'Debug.Print "  --> TIME:" & llTime & " , RdfCode:" & ilRdfCode
                        If rbcReRateBook(0).Value Then  'Default vehicle
                            ilVef = gBinarySearchVef(tmSdfExt(llSdf).iVefCode)
                            If ilVef <> -1 Then
                                ilDnfCode = tgMVef(ilVef).iDnfCode
                            Else
                                ilDnfCode = 0
                            End If
                        End If
                        'If rbcReRateBook(3).Value Then  'Contract Line
                        If bmBookByLine Then
                            For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                                If tgBookByLineAssigned(ilBook).lChfCode = tmSdfExt(llSdf).lChfCode Then
                                    If (tgBookByLineAssigned(ilBook).iLineNo = tmSdfExt(llSdf).iLineNo) And (tgBookByLineAssigned(ilBook).iVefCode = tmSdfExt(llSdf).iVefCode) And ((tgBookByLineAssigned(ilBook).iMGCount > 0) Or (tgBookByLineAssigned(ilBook).iOutsideCount > 0)) Then
                                        ''ilDnfCode = tgBookByLineAssigned(ilBook).lReRateDnfCode
                                        'If tgBookByLineAssigned(ilBook).lReRateDnfCode = -1 Then 'Vehicle default
                                        '    ilVef = gBinarySearchVef(tmSdfExt(llSdf).iVefCode)
                                        '    If ilVef <> -1 Then
                                        '        ilDnfCode = tgMVef(ilVef).iDnfCode
                                        '    Else
                                        '        ilDnfCode = 0
                                        '    End If
                                        'ElseIf tgBookByLineAssigned(ilBook).lReRateDnfCode = -2 Then 'Closest
                                        '    'set at top of this routine
                                        'ElseIf tgBookByLineAssigned(ilBook).lReRateDnfCode = -3 Then 'Purchase
                                        '    ilDnfCode = tmClfP.iDnfCode
                                        'Else
                                        '    ilDnfCode = tgBookByLineAssigned(ilBook).lReRateDnfCode
                                        'End If
                                        ''imReRateDnfCode = ilDnfCode
                                        ilDnfCode = mGetDnfByContractLine(ilBook, llSdf)
'Debug.Print "  --> DnfCode :" & ilDnfCode
                                        Exit For
                                    End If
                                End If
                            Next ilBook
                        End If
                        For ilVef = 0 To UBound(tmMGInfo) - 1 Step 1
                            If (tmSdfExt(llSdf).iVefCode = tmMGInfo(ilVef).iVefCode) And (ilDnfCode = tmMGInfo(ilVef).iDnfCode) And (ilRdfCode = tmMGInfo(ilVef).iRdfCode) Then
                                blVefFd = True
                                Exit For
                            End If
                        Next ilVef
                        If Not blVefFd Then
                            tmMGInfo(UBound(tmMGInfo)).iVefCode = tmSdfExt(llSdf).iVefCode
                            tmMGInfo(UBound(tmMGInfo)).iDnfCode = ilDnfCode
                            tmMGInfo(UBound(tmMGInfo)).iRdfCode = ilRdfCode
                            ReDim Preserve tmMGInfo(0 To UBound(tmMGInfo) + 1) As MGBONUSINFO
                        End If
                    Else
'Debug.Print "  -> MG/OS as Ordered NOT Checked"
                        'ilTotLnSpts = ilTotLnSpts + 1
                        'If tmClfReRate.sType = "S" Then
                        '    llTotalCntrSpots = llTotalCntrSpots + 1
                        'End If
                        If rbcReRateBook(0).Value Then  'Default vehicle
                            ilVef = gBinarySearchVef(tmSdfExt(llSdf).iVefCode)
                            If ilVef <> -1 Then
                                ilDnfCode = tgMVef(ilVef).iDnfCode
                            Else
                                ilDnfCode = 0
                            End If
                        End If
                        'If rbcReRateBook(3).Value Then  'Contract Line
                        If bmBookByLine Then
                            For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                                If tgBookByLineAssigned(ilBook).lChfCode = tmSdfExt(llSdf).lChfCode Then
                                    If (tgBookByLineAssigned(ilBook).iLineNo = tmSdfExt(llSdf).iLineNo) And (tgBookByLineAssigned(ilBook).iVefCode = tmSdfExt(llSdf).iVefCode) And ((tgBookByLineAssigned(ilBook).iMGCount > 0) Or (tgBookByLineAssigned(ilBook).iOutsideCount > 0) Or (tgBookByLineAssigned(ilBook).iBonusCount > 0)) Then
                                        ''ilDnfCode = tgBookByLineAssigned(ilBook).lReRateDnfCode
                                        'If tgBookByLineAssigned(ilBook).lReRateDnfCode = -1 Then 'Vehicle default
                                        '    ilVef = gBinarySearchVef(tmSdfExt(llSdf).iVefCode)
                                        '    If ilVef <> -1 Then
                                        '        ilDnfCode = tgMVef(ilVef).iDnfCode
                                        '    Else
                                        '        ilDnfCode = 0
                                        '    End If
                                        'ElseIf tgBookByLineAssigned(ilBook).lReRateDnfCode = -2 Then 'Closest
                                        '    'set at top of this routine
                                        'ElseIf tgBookByLineAssigned(ilBook).lReRateDnfCode = -3 Then 'Purchase
                                        '    ilDnfCode = tmClfP.iDnfCode
                                        'Else
                                        '    ilDnfCode = tgBookByLineAssigned(ilBook).lReRateDnfCode
                                        'End If
                                        'imReRateDnfCode = ilDnfCode
                                        ilDnfCode = mGetDnfByContractLine(ilBook, llSdf)
'Debug.Print "  -> ilDnfCode: " & ilDnfCode
                                        Exit For
                                    End If
                                End If
                            Next ilBook
                        End If
                        mSaveReRateBookByLine llSdf, ilDnfCode
'Debug.Print "  -> mGetMGAud1...."
                        mGetMGAud llSdf, tmSdfExt(llSdf).iVefCode, ilDnfCode, ilClf, ilMnfDemo, llStartDate, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode, ilTotLnSpts, llTotalCntrSpots
                    End If
                End If
            End If
        End If
    Next llSdf
    If (ckcMG.Value = vbUnchecked And ckcTreatMGOsAsOrdered.Value = vbUnchecked) Then
        For illoop = 0 To UBound(lmWklyAvgAud) Step 1
            If lmWklyspots(illoop) > 0 Then
                lmWklyAvgAud(illoop) = lmWklyAvgAud(illoop) / lmWklyspots(illoop)
            End If
        Next illoop
    End If
End Sub

Private Sub mGetMGSpotCount(ilVefCode As Integer, ilDnfCode As Integer, ilRdfCode As Integer, ilClf As Integer, llStartDate As Long, llEndDate As Long, ilMnfDemo As Integer, ilTotLnSpts As Integer, llTotalCntrSpots As Long, llAvgAud As Long, llPopEst As Long, ilAudFromSource As Integer, llAudFromCode As Long)
    Dim llSdf As Long
    Dim ilCff As Integer
    Dim illoop As Integer
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim slStr As String
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilSpots As Integer
    Dim ilUpperWk As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llMoDate As Long
    ReDim ilInputDays(0 To 6) As Integer
    Dim ilRet As Integer
    Dim llTime As Long
    Dim ilSdfRdfCode As Integer
    Dim ilSdfDnfCode As Integer
    Dim ilBook As Integer
    
    'llOvStartTime = 0
    'llOvEndTime = 86400
    ilSpots = 1
    imMGDetailDnfCode = -1
    For llSdf = LBound(tmSdfExt) To UBound(tmSdfExt) - 1 Step 1
        If (tmSdfExt(llSdf).iVefCode = ilVefCode) Then
            gUnpackDateLong tmSdfExt(llSdf).iDate(0), tmSdfExt(llSdf).iDate(1), llDate
            If (((llDate >= llStartDate) And (llDate <= llEndDate)) Or (rbcDatesBy(3).Value = True)) And (tmSdfExt(llSdf).iLineNo = tmClfR.iLine) And ((tmSdfExt(llSdf).sSchStatus = "G") Or (tmSdfExt(llSdf).sSchStatus = "O")) And (tmSdfExt(llSdf).sSpotType <> "X") And (tmSdfExt(llSdf).lMdDate <> 0) Then
                'For ilLoop = 0 To 6                 'init all days to not airing, setup for research results later
                '    ilInputDays(ilLoop) = False
                'Next ilLoop
            
                'ilInputDays(gWeekDayLong(llDate)) = True
                'gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llOvStartTime
                'gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llOvEndTime
                'llOvEndTime = llOvEndTime + tmSdfExt(llSdf).iLen

                ''ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, imReRateDnfCode, ilVefCode, 0, ilMnfDemo, llDate, llDate, tmClfReRate.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClfReRate.sType, tmClfReRate.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                'ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, imReRateDnfCode, ilVefCode, 0, ilMnfDemo, llDate, llDate, 0, llOvStartTime, llOvEndTime, ilInputDays(), "S", 0, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                ''Loop and build avg aud, spots, & spots per week
                ''calc week index
                'ilUpperWk = (llDate - llStartDate) / 7 + 1
                ''ilUpperWk = UBound(lmWklyspots)
                'If ilUpperWk - 1 > UBound(lmWklyspots) Then
                '    ReDim Preserve lmWklyspots(0 To ilUpperWk) As Integer
                '    ReDim Preserve lmWklyAvgAud(0 To ilUpperWk) As Long
                '    ReDim Preserve lmWklyPopEst(0 To ilUpperWk) As Long
                'End If
                'lmWklyspots(ilUpperWk - 1) = lmWklyspots(ilUpperWk - 1) + ilSpots
                'ilTotLnSpts = ilTotLnSpts + ilSpots
                'If tmClfReRate.sType = "S" Then
                '    llTotalCntrSpots = llTotalCntrSpots + ilSpots
                'End If
                ''Obtain price from Missed spot
                'lmWklyRates(ilUpperWk - 1) = 0
                'ilCff = tmClfReRate(ilClf).iFirstCff
                'Do While ilCff <> -1
                '    tmCff = tmCffReRate(ilCff).CffRec
                '    gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llFltStart
                '    gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llFltEnd
                '    If (tmSdfExt(llSdf).lMdDate >= llFltStart) And (tmSdfExt(llSdf).lMdDate <= llFltEnd) Then
                '        lmWklyRates(ilUpperWk - 1) = tmCff.lActPrice
                '        Exit Do
                '    End If
                '    ilCff = tmCffReRate(ilCff).iNextCff
                'Loop
                ''lmWklyAvgAud(ilUpperWk - 1) = llAvgAud
                'lmWklyAvgAud(ilUpperWk - 1) = lmWklyAvgAud(ilUpperWk - 1) + llAvgAud
                'lmWklyPopEst(ilUpperWk - 1) = llPopEst
                gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llTime
                ilSdfRdfCode = mFindDaypart(tmSdfExt(llSdf).iVefCode, llDate, llTime)
                ilSdfDnfCode = imReRateDnfCode
                If rbcReRateBook(1).Value And Not bmBookByLine Then 'Or (rbcReRateBook(3).Value And imByCntrLnDnfCode = -2) Then   'Closest
                    ilDnfCode = mFindClosestBook(llDate, llSdf) 'tmSdfExt(llSdf).iVefCode)
                End If
                'If rbcReRateBook(3).Value Then  'Contract Line
                If bmBookByLine Then
                    For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                        If tgBookByLineAssigned(ilBook).lChfCode = tmSdfExt(llSdf).lChfCode Then
                            If (tgBookByLineAssigned(ilBook).iLineNo = tmSdfExt(llSdf).iLineNo) And (tgBookByLineAssigned(ilBook).iVefCode = tmSdfExt(llSdf).iVefCode) And ((tgBookByLineAssigned(ilBook).iMGCount > 0) Or (tgBookByLineAssigned(ilBook).iOutsideCount > 0)) Then
                                ''ilDnfCode = tgBookByLineAssigned(ilBook).lReRateDnfCode
                                'If tgBookByLineAssigned(ilBook).lReRateDnfCode = -1 Then 'Vehicle default
                                '    ilVef = gBinarySearchVef(tmSdfExt(llSdf).iVefCode)
                                '    If ilVef <> -1 Then
                                '        ilDnfCode = tgMVef(ilVef).iDnfCode
                                '    Else
                                '        ilDnfCode = 0
                                '    End If
                                'ElseIf tgBookByLineAssigned(ilBook).lReRateDnfCode = -2 Then 'Closest
                                '    'set at top of this routine
                                'ElseIf tgBookByLineAssigned(ilBook).lReRateDnfCode = -3 Then 'Purchase
                                '    ilDnfCode = tmClfP.iDnfCode
                                'Else
                                '    ilDnfCode = tgBookByLineAssigned(ilBook).lReRateDnfCode
                                'End If
                                ilDnfCode = mGetDnfByContractLine(ilBook, llSdf)
                                Exit For
                            End If
                        End If
                    Next ilBook
                End If
                'If (ilDnfCode = ilSdfDnfCode) And (ilRdfCode = ilSdfRdfCode) Then
                'If rbcReRateBook(1).Value Or rbcReRateBook(3).Value Then
                If rbcReRateBook(1).Value Or bmBookByLine Then
                    If lmReRatePop = -1 Then
                        ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, 0, ilMnfDemo, lmReRatePop)
                    End If
'Debug.Print "  -> mGetMGAud2...."
                    mGetMGAud llSdf, ilVefCode, ilDnfCode, ilClf, ilMnfDemo, llStartDate, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode, ilTotLnSpts, llTotalCntrSpots
                Else
                    If lmReRatePop = -1 Then
                        ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilSdfDnfCode, 0, ilMnfDemo, lmReRatePop)
                    End If
'Debug.Print "  -> mGetMGAud3...."
                    mGetMGAud llSdf, ilVefCode, ilSdfDnfCode, ilClf, ilMnfDemo, llStartDate, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode, ilTotLnSpts, llTotalCntrSpots
                End If
                'End If
            End If
        End If
    Next llSdf
    For illoop = 0 To UBound(lmWklyAvgAud) Step 1
        If lmWklyspots(illoop) > 0 Then
            lmWklyAvgAud(illoop) = lmWklyAvgAud(illoop) / lmWklyspots(illoop)
        End If
    Next illoop
End Sub
Private Sub mGetBonusSpotCount(ilVefCode As Integer, ilDnfCode As Integer, ilRdfCode As Integer, llStartDate As Long, llEndDate As Long, ilMnfDemo As Integer, ilTotLnSpts As Integer, llTotalCntrSpots As Long, llAvgAud As Long, llPopEst As Long, ilAudFromSource As Integer, llAudFromCode As Long, Optional ilSpotLength As Integer = 0)
    Dim llSdf As Long
    Dim ilCff As Integer
    Dim illoop As Integer
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim slStr As String
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilSpots As Integer
    Dim ilUpperWk As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llMoDate As Long
    ReDim ilInputDays(0 To 6) As Integer
    Dim ilRet As Integer
    Dim ilSdfDnfCode As Integer
    Dim ilSdfRdfCode As Integer
    Dim ilVef As Integer
    Dim ilClf As Integer
    Dim llTime As Long
    Dim ilBook As Integer
    
    
    ilSpots = 1
    imBonusDetailDnfCode = -1
    For llSdf = LBound(tmSdfExt) To UBound(tmSdfExt) - 1 Step 1
        gUnpackDateLong tmSdfExt(llSdf).iDate(0), tmSdfExt(llSdf).iDate(1), llDate
        If (((llDate >= llStartDate) And (llDate <= llEndDate)) Or (rbcDatesBy(3).Value = True)) And (tmSdfExt(llSdf).iVefCode = ilVefCode) And (tmSdfExt(llSdf).sSpotType = "X") And (ilSpotLength = 0 Or ilSpotLength = tmSdfExt(llSdf).iLen) Then
            'If pop not computed determined, get it
            If lmReRatePop = -1 Then
                If rbcReRateBook(1).Value Then  'Or (rbcReRateBook(3).Value And imByCntrLnDnfCode = -2) Then 'Closest
                    ilDnfCode = mFindClosestBook(llDate, llSdf) 'tmSdfExt(llSdf).iVefCode)
                    ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, 0, ilMnfDemo, lmReRatePop)
                'ElseIf rbcReRateBook(3).Value Then  'Contract Line
                ElseIf bmBookByLine Then
                    ilDnfCode = 0
                    For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                        If tgBookByLineAssigned(ilBook).lChfCode = tmSdfExt(llSdf).lChfCode Then
                            If (tgBookByLineAssigned(ilBook).iLineNo = tmSdfExt(llSdf).iLineNo) And (tgBookByLineAssigned(ilBook).iVefCode = tmSdfExt(llSdf).iVefCode) Then
                                ilDnfCode = mGetDnfByContractLine(ilBook, llSdf)
                                ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, 0, ilMnfDemo, lmReRatePop)
                                Exit For
                            End If
                        End If
                    Next ilBook
                Else
                    ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, imReRateDnfCode, 0, ilMnfDemo, lmReRatePop)
                End If
            End If
            If (rbcBonus(1).Value = True) Then
                gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llTime
                ilSdfRdfCode = mFindDaypart(tmSdfExt(llSdf).iVefCode, llDate, llTime)
            Else
                ilSdfRdfCode = ilRdfCode
            End If
            If ilSdfRdfCode = ilRdfCode Then
                If mTestShowFill(tmSdfExt(llSdf).sPriceType) = "Y" Then
                    ilSdfDnfCode = ilDnfCode    'imReRateDnfCode
                    If rbcReRateBook(1).Value Then  'Or (rbcReRateBook(3).Value And imByCntrLnDnfCode = -2) Then  'Closest
                        ilSdfDnfCode = mFindClosestBook(llDate, llSdf)  'tmSdfExt(llSdf).iVefCode)
                    End If
                    If rbcReRateBook(0).Value Then  'Default vehicle
                        ilVef = gBinarySearchVef(tmSdfExt(llSdf).iVefCode)
                        If ilVef <> -1 Then
                            ilSdfDnfCode = tgMVef(ilVef).iDnfCode
                        Else
                            ilSdfDnfCode = 0
                        End If
                    End If
                    'If rbcReRateBook(3).Value Then  'Contract Line
                    If bmBookByLine Then
                        For ilBook = 0 To UBound(tgBookByLineAssigned) - 1 Step 1
                            If tgBookByLineAssigned(ilBook).lChfCode = tmSdfExt(llSdf).lChfCode Then
                                If (tgBookByLineAssigned(ilBook).iLineNo = tmSdfExt(llSdf).iLineNo) And (tgBookByLineAssigned(ilBook).iVefCode = tmSdfExt(llSdf).iVefCode) And (tgBookByLineAssigned(ilBook).iBonusCount > 0) Then
                                    'ilSdfDnfCode = tgBookByLineAssigned(ilClf).lReRateDnfCode
                                    ilSdfDnfCode = mGetDnfByContractLine(ilBook, llSdf)
                                    Exit For
                                End If
                            End If
                        Next ilBook
                    End If
                    If imBonusDetailDnfCode = -1 Then
                        imBonusDetailDnfCode = ilDnfCode
                    ElseIf imBonusDetailDnfCode <> ilDnfCode And imBonusDetailDnfCode <> -2 Then
                        imBonusDetailDnfCode = -2
                    End If
                    If ilSdfDnfCode = ilDnfCode Then
                        For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                            ilInputDays(illoop) = False
                        Next illoop
                        
                        ilInputDays(gWeekDayLong(llDate)) = True
                        gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llOvStartTime
                        gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llOvEndTime
                        llOvEndTime = llOvEndTime + tmSdfExt(llSdf).iLen
                        'ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, imReRateDnfCode, ilVefCode, 0, ilMnfDemo, llDate, llDate, tmClfReRate.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClfReRate.sType, tmClfReRate.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
'Debug.Print "  getDemo3; RDFCode=0000"
                        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, 0, ilMnfDemo, llDate, llDate, 0, llOvStartTime, llOvEndTime, ilInputDays(), "S", 0, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                        ''Loop and build avg aud, spots, & spots per week
                        ''calc week index
                        'ilUpperWk = (llDate - llStartDate) / 7 + 1
                        ''ilUpperWk = UBound(lmWklyspots)
                        'If ilUpperWk - 1 > UBound(lmWklyspots) Then
                        '    ReDim Preserve lmWklyspots(0 To ilUpperWk) As Integer
                        '    ReDim Preserve lmWklyAvgAud(0 To ilUpperWk) As Long
                        '    ReDim Preserve lmWklyPopEst(0 To ilUpperWk) As Long
                        'End If
                        'lmWklyspots(ilUpperWk - 1) = lmWklyspots(ilUpperWk - 1) + ilSpots
                        ilTotLnSpts = ilTotLnSpts + ilSpots
                        ''If tmClfReRate.sType = "S" Then
                        ''    llTotalCntrSpots = llTotalCntrSpots + ilSpots
                        ''End If
                        'lmWklyRates(ilUpperWk - 1) = 0
                        'lmWklyAvgAud(ilUpperWk - 1) = lmWklyAvgAud(ilUpperWk - 1) + llAvgAud
                        'lmWklyPopEst(ilUpperWk - 1) = llPopEst
                        mAddSpotsToWeekArray llDate, llStartDate, ilSpots, llAvgAud, 0, llPopEst, True
                    End If
                End If
            End If
        End If
    Next llSdf
    For illoop = 0 To UBound(lmWklyAvgAud) Step 1
        If lmWklyspots(illoop) > 0 Then
            lmWklyAvgAud(illoop) = lmWklyAvgAud(illoop) / lmWklyspots(illoop)
        End If
    Next illoop
End Sub
Private Sub mOutputCntrHeader(llStartDate As Long, llEndDate As Long)
    Dim ilAgf As Integer
    Dim ilAdf As Integer
    Dim ilRet As Integer
    Dim llCntrPop As Long
    Dim llReRatePop As Long
    Dim illoop As Integer
    Dim slIndexBy As String
    Dim ilChf As Integer
    Dim slDemo As String
    Dim slTempHeader As String
    
    If rbcIndex(1).Value Then
        slIndexBy = "GRP"
    Else
        slIndexBy = "Gimp"
    End If

    'TTP 10082 - merge header into columns
    If rbcLayout(0).Value = True Then mPrint ""
    If tmChfPurchase.iAgfCode > 0 Then
        ilAgf = gBinarySearchAgf(tmChfPurchase.iAgfCode)
        If ilAgf <> -1 Then
            'Print #hmToCSV, "Agency: " & tgCommAgf(ilAgf).sName
            If rbcLayout(0).Value = True Then mPrint "Agency: " & tgCommAgf(ilAgf).sName, smDelimiter
            tmReRateHeader.Agency = Trim(tgCommAgf(ilAgf).sName)
        Else
            'Print #hmToCSV, "Agency Code: " & tmChfPurchase.iAgfCode
            If rbcLayout(0).Value = True Then mPrint "Agency Code: " & tmChfPurchase.iAgfCode, smDelimiter
            tmReRateHeader.Agency = tmChfPurchase.iAgfCode
        End If
        If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    End If
    ilAdf = gBinarySearchAdf(imAdfCode)
    If ilAdf <> -1 Then
        'Print #hmToCSV, "Advertiser: " & Trim$(tgCommAdf(ilAdf).sName)
        If rbcLayout(0).Value = True Then mPrint "Advertiser: " & Trim$(tgCommAdf(ilAdf).sName), smDelimiter
        tmReRateHeader.Advertiser = Trim$(tgCommAdf(ilAdf).sName)
    Else
        'Print #hmToCSV, "Advertiser Code: " & imAdfCode
        If rbcLayout(0).Value = True Then mPrint "Advertiser Code: " & imAdfCode, smDelimiter
        tmReRateHeader.Advertiser = imAdfCode
    End If
    If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    
    'Print #hmToCSV, "Product: " & Trim$(tmChfPurchase.sProduct)
    If rbcLayout(0).Value = True Then mPrint "Product: " & Trim$(tmChfPurchase.sProduct), smDelimiter
    tmReRateHeader.Product = Trim$(tmChfPurchase.sProduct)
    If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    
    'Print #hmToCSV, "Order#: " & Trim$(tmChfPurchase.lCntrNo)
    If rbcLayout(0).Value = True Then mPrint "Order#: " & Trim$(tmChfPurchase.lCntrNo), smDelimiter
    tmReRateHeader.OrderNo = Trim$(tmChfPurchase.lCntrNo)
    If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    
    If rbcLayout(0).Value = True Then mPrint "Purchase Revision#: " & "R" & tmChfPurchase.iCntRevNo & "-" & tmChfPurchase.iExtRevNo, smDelimiter
    tmReRateHeader.PurchaseRevision = "R" & tmChfPurchase.iCntRevNo & "-" & tmChfPurchase.iExtRevNo
    If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    
    If rbcLayout(0).Value = True Then mPrint "ReRate Revision#: " & "R" & tmChfReRate.iCntRevNo & "-" & tmChfReRate.iExtRevNo, smDelimiter
    tmReRateHeader.ReRateRevision = "R" & tmChfReRate.iCntRevNo & "-" & tmChfReRate.iExtRevNo
    If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    
    'Print #hmToCSV, "Demo: " & Trim$(cbcDemoNames.Text)
    slDemo = mGetDemo(imMnfDemo)
    If rbcLayout(0).Value = True Then mPrint "Demo: " & slDemo, smDelimiter
    tmReRateHeader.Demo = slDemo
    If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    
    llCntrPop = -1
    llReRatePop = -1
    For illoop = LBound(tmReRate) To UBound(tmReRate) - 1 Step 1
        If ((tmReRate(illoop).sType = "S") Or (tmReRate(illoop).sType = "H")) And (Trim$(tmReRate(illoop).sSubType) = "") Then
            If llCntrPop = -1 Then
                llCntrPop = tmReRate(illoop).lPop(0)
            Else
                If llCntrPop <> tmReRate(illoop).lPop(0) Then
                    llCntrPop = -2
                End If
            End If
            If llReRatePop = -1 Then
                llReRatePop = tmReRate(illoop).lPop(1)
            Else
                If llReRatePop <> tmReRate(illoop).lPop(1) Then
                    llReRatePop = -2
                End If
            End If
        End If
    Next illoop
    If llCntrPop >= 0 Then
        If rbcLayout(0).Value = True Then mPrint "Contract Population: " & gLongToStrDec(llCntrPop, imNumberDecPlaces), smDelimiter
        tmReRateHeader.ContractPopulation = gLongToStrDec(llCntrPop, imNumberDecPlaces)
    Else
        If rbcLayout(0).Value = True Then mPrint "Contract Population: " & "Population varies Across Books", smDelimiter '"Varies"
        tmReRateHeader.ContractPopulation = "Population varies Across Books"
    End If
    If ExptReRate.ckcCsv.Value = vbUnchecked Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    
    If llReRatePop >= 0 Then
        If rbcLayout(0).Value = True Then mPrint "ReRate Population: " & gLongToStrDec(llReRatePop, imNumberDecPlaces), smDelimiter
        tmReRateHeader.ReRatePopulation = gLongToStrDec(llReRatePop, imNumberDecPlaces)
    Else
        If rbcLayout(0).Value = True Then mPrint "ReRate Population: " & "Population varies Across Books", smDelimiter '"Varies"
        tmReRateHeader.ReRatePopulation = "Population varies Across Books"
    End If
    If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    
    If rbcDatesBy(0).Value Then   'Weekly
        If rbcLayout(0).Value = True Then mPrint "Period: Week Start Date " & Format(edcDate(0).Text, "mm/dd/yy"), smDelimiter
        tmReRateHeader.Period = "Week " & Format(edcDate(0).Text, "mm/dd/yy")
    ElseIf rbcDatesBy(1).Value Then   'Month
        If rbcLayout(0).Value = True Then mPrint "Period: Month " & edcStart.Text & " Year " & edcYear.Text, smDelimiter
        tmReRateHeader.Period = "Month " & edcStart.Text & " Year " & edcYear.Text
    ElseIf rbcDatesBy(2).Value Then   'Quarter
        If rbcLayout(0).Value = True Then mPrint "Period: Quarter " & edcStart.Text & " Year " & edcYear.Text, smDelimiter
        tmReRateHeader.Period = "Quarter " & edcStart.Text & " Year " & edcYear.Text
    ElseIf rbcDatesBy(3).Value Then   'Contract
        If ckcSummary.Value = vbUnchecked Then
            If rbcLayout(0).Value = True Then mPrint "Period: " & Format(llStartDate, "mm/dd/yy") & "-" & Format(llEndDate, "mm/dd/yy"), smDelimiter
            tmReRateHeader.Period = Format(llStartDate, "mm/dd/yy") & "-" & Format(llEndDate, "mm/dd/yy")
        Else
            If imNoCntr > 1 Then
                If rbcLayout(0).Value = True Then mPrint "Period: " & "Contract Date Span", smDelimiter
                tmReRateHeader.Period = "Contract Date Span"
            Else
                If rbcLayout(0).Value = True Then mPrint "Period: " & Format(llStartDate, "mm/dd/yy") & "-" & Format(llEndDate, "mm/dd/yy"), smDelimiter
                tmReRateHeader.Period = Format(llStartDate, "mm/dd/yy") & "-" & Format(llEndDate, "mm/dd/yy")
            End If
        End If
    ElseIf rbcDatesBy(4).Value Then   'Range
        If rbcLayout(0).Value = True Then mPrint "Period: " & Format(edcDate(1).Text, "mm/dd/yy") & "-" & Format(edcDate(2).Text, "mm/dd/yy"), smDelimiter
        tmReRateHeader.Period = Format(edcDate(1).Text, "mm/dd/yy") & "-" & Format(edcDate(2).Text, "mm/dd/yy")
    End If
    If ExptReRate.ckcCsv.Value = vbUnchecked And rbcLayout(0).Value = True Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
    If rbcLayout(0).Value = True Then mPrint ""
    
    'TTP 10082 - merge header into columns, Only print the Header row once
    If rbcLayout(1).Value = True And bmExportedHeader = True Then Exit Sub
    
    '--------------------------------------------------------
    'TTP 10082 - merge header into columns - Column Headers1
    If rbcColumnLayout(1).Value = True Then
        If rbcLayout(0).Value = True Then slTempHeader = ",,,,,,,,,,Purch,Purch,Purch,Purch,Purch,Purch,Purch,Purch,Purch,ReRate,ReRate,ReRate,ReRate,ReRate,ReRate,ReRate,ReRate,ReRate," & slIndexBy & ","
        If rbcLayout(1).Value = True Then slTempHeader = ",,,,,,,,,,,,,,,,,,,,Purch,Purch,Purch,Purch,Purch,Purch,Purch,Purch,Purch,ReRate,ReRate,ReRate,ReRate,ReRate,ReRate,ReRate,ReRate,ReRate," & slIndexBy & ","
        slTempHeader = Replace(slTempHeader, ",", smDelimiter)
        mPrint slTempHeader, smDelimiter
    Else
        If ckcCost.Value = vbChecked Then
            If rbcLayout(0).Value = True Then slTempHeader = ",,,,,,,,,,Purchased,,,,,,,,,ReRate,,,,,,,,," & slIndexBy & ","
            If rbcLayout(1).Value = True Then slTempHeader = ",,,,,,,,,,,,,,,,,,,,Purchased,,,,,,,,,ReRate,,,,,,,,," & slIndexBy & ","
            slTempHeader = Replace(slTempHeader, ",", smDelimiter)
            mPrint slTempHeader, smDelimiter
        Else
            'Cost columns hidden
            If rbcLayout(0).Value = True Then slTempHeader = ",,,,,,,,,,Purchased,Purchased,,,,,,,,ReRate,ReRate,,,,,,,," & slIndexBy & ","
            If rbcLayout(1).Value = True Then slTempHeader = ",,,,,,,,,,,,,,,,,,,,Purchased,Purchased,,,,,,,,ReRate,ReRate,,,,,,,," & slIndexBy & ","
            slTempHeader = Replace(slTempHeader, ",", smDelimiter)
            mPrint slTempHeader, smDelimiter
        End If
    End If
    If ExptReRate.ckcCsv.Value = vbUnchecked Then
        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1)
        ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , vbBlue, imExcelRow - 1)
    End If
    
    '--------------------------------------------------------
    'TTP 10082 - merge header into columns - Column Headers2
    'TTP 10144 Added Audio Type Column
    If rbcLayout(0).Value = True Then slTempHeader = ",Line#,Vehicle,Daypart,Lineup #,Audio Type,Len,Price Type,Rate,Line Comment,Ext Totals,Units,AQH,Rtg,CPM,CPP,GIMPs,GRPs,Book,Ext Totals,Units,AQH,Rtg,CPM,CPP,GIMPs,GRPs,Book,Index"
    If rbcLayout(1).Value = True Then slTempHeader = "Agency,Advertiser,Product,Order#,Purchase Rev#,ReRate Rev#,Demo,Contract Population,ReRate Population,Period,,Line#,Vehicle,Daypart,Lineup #,Audio Type,Len,Price Type,Rate,Line Comment,Ext Totals,Units,AQH,Rtg,CPM,CPP,GIMPs,GRPs,Book,Ext Totals,Units,AQH,Rtg,CPM,CPP,GIMPs,GRPs,Book,Index"
    slTempHeader = Replace(slTempHeader, ",", smDelimiter)
    mPrint slTempHeader, smDelimiter
            
    If ExptReRate.ckcCsv.Value = vbUnchecked Then
        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1)
        ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , vbBlue, imExcelRow - 1)
    End If
    bmExportedHeader = True 'TTP 10082 - merge contract header into columns..  flag to not repeat printing headers between contracts.
End Sub

Private Sub mSendToExcel()
'   Excel Object Library
'   https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview?view=vs-2017
'   https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range?redirectedfrom=MSDN&view=excel-pia
'   https://support.microsoft.com/en-us/help/219151/how-to-automate-microsoft-excel-from-visual-basic

    Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object

   'Start a new workbook in Excel
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add


   'Add data to cells of the first worksheet in the new workbook
   Set oSheet = oBook.Worksheets(1)
   oSheet.Range("A1").Value = "Last Name"
   oSheet.Range("B1").Value = "First Name"
   oSheet.Range("A1:B1").Font.Bold = True
   oSheet.Range("A2").Value = "Doe"
   oSheet.Range("B2").Value = "John"

   'Save the Workbook and Quit Excel
   oBook.SaveAs "C:csi\Book1.xlsx"
    
    ' This makes Excel visible
    oExcel.Visible = True
    
    'The Quit will remove the excel image
    
   'oExcel.Quit
   
   
End Sub

'Private Sub mPrint(slInRecord As String, Optional slDelimiter As String = "", Optional slRowType As String = "", Optional llPop As Long = 0, Optional llExtTotal As Long = 0)
Private Sub mPrint(slInRecord As String, Optional slDelimiter As String = "", Optional slRowType As String = "", Optional llPop As Long = 0, Optional dlExtTotal As Double = 0) 'TTP 10439 - Rerate 21,000,000
    Dim ilColumn As Integer
    Dim ilCell As Integer
    Dim ilCol As Integer
    Dim ilRet As Integer
    Dim slField As String
    Dim blPRPRColorSwitch As Boolean
    Dim blSkipColor As Boolean
    Dim slStr As String
    Dim slVehicle As String
    Dim slType As String

    Dim slRecord As String
    Dim slCSVRecord As String
    Dim slRecordsArray() As String
    Dim ilCSV As Integer
    Dim slCSV() As String
    Dim blNeedComma As Boolean
    Dim blSkipColumn As Boolean
    ilColumn = 1
    slRecord = slInRecord
    If Not bmInSummaryMode Then
        If slRecord <> "" Then
            If slDelimiter = "" Then
'Debug.Print "adding to summary: " & Trim(slRecord)
                smSummaryRecords(UBound(smSummaryRecords)) = Trim(slRecord)
                ReDim Preserve smSummaryRecords(0 To UBound(smSummaryRecords) + 1) As String
            Else
                'look in 2nd Column For a Number - if it is Numeric; then it is a Detail Line, and we dont want to keep that.
                ilRet = gParseItem(slRecord, 2, slDelimiter, slStr)
                If IsNumeric(slStr) = False Then
                    'Keep non Detail Lines (Keep Header and total lines only)
'Debug.Print "adding to summary: " & Trim(slRecord)
                    smSummaryRecords(UBound(smSummaryRecords)) = Trim(slRecord)
                    ReDim Preserve smSummaryRecords(0 To UBound(smSummaryRecords) + 1) As String
                End If
            End If
        End If

        If ckcSummary.Value = vbChecked Then
            Exit Sub
        End If
    End If
    
    '----------------------------------------------------------------
    'TTP 10258: ReRate - make it work without requiring Office (CSV)
    If ExptReRate.ckcCsv.Value = vbChecked Then
        slCSVRecord = ""
        blNeedComma = False
        slCSV = Split(slRecord, slDelimiter)
        For ilCSV = 0 To UBound(slCSV)
            'Skip exporting some columns
            blSkipColumn = False
            'Cost
            If ckcCost.Value = vbUnchecked Then
                If ilCSV + 1 = RATEEXCEL Then blSkipColumn = True 'RATE
                If ilCSV + 1 = REXTTOTALEXCEL Then blSkipColumn = True 'REXTTOTAL
                If ilCSV + 1 = PEXTTOTALEXCEL Then blSkipColumn = True 'PEXTTOTAL
            End If
            'Rating
            If ckcRating.Value = vbUnchecked Then
                If ilCSV + 1 = PRTGEXCEL Then blSkipColumn = True 'PRTG
                If ilCSV + 1 = RRTGEXCEL Then blSkipColumn = True 'RRTG
            End If
            'CPM/CPP
            If ckcCPM.Value = vbUnchecked Then
                If ilCSV + 1 = PCPMEXCEL Then blSkipColumn = True 'PCPM
                If ilCSV + 1 = PCPPEXCEL Then blSkipColumn = True 'PCPP
                If ilCSV + 1 = RCPMEXCEL Then blSkipColumn = True 'RCPM
                If ilCSV + 1 = RCPPEXCEL Then blSkipColumn = True 'RCPP
            End If
            'AudioType
            If ckcAudioType.Value = vbUnchecked Then
                If ilCSV + 1 = AUDIOTYPEEXCEL Then blSkipColumn = True 'AUDIOTYPE
            End If
            'LineUp# (Act1)
            If ckcACT1Lineup.Value = vbUnchecked Then
                If ilCSV + 1 = LINEUPEXCEL Then blSkipColumn = True 'LINEUP
            End If
            'PriceType
            If ckcPriceType.Value = vbUnchecked Then
                If ilCSV + 1 = PRICETYPEEXCEL Then blSkipColumn = True 'PRICETYPE
            End If
            'Comment
            If ckcComment.Value = vbUnchecked Then
                If ilCSV + 1 = LINECOMMENTEXCEL Then blSkipColumn = True 'LINECOMMENT
            End If
            'If Summary Mode
            If ckcSummary.Value = vbChecked Then
                If ilCSV + 1 = LINEUPEXCEL Then blSkipColumn = True 'LINEUP
                If ilCSV + 1 = AUDIOTYPEEXCEL Then blSkipColumn = True 'AUDIOTYPE
                If ilCSV + 1 = LENEXCEL Then blSkipColumn = True 'LEN
                If ilCSV + 1 = PRICETYPEEXCEL Then blSkipColumn = True 'PRICETYPE
                If ilCSV + 1 = RATEEXCEL Then blSkipColumn = True 'RATE
                If ilCSV + 1 = LINECOMMENTEXCEL Then blSkipColumn = True 'LINECOMMENT
                If ilCSV + 1 = PAQHEXCEL Then blSkipColumn = True 'PAQH
                If ilCSV + 1 = RAQHEXCEL Then blSkipColumn = True 'RAQH
                If ilCSV + 1 = PBOOKEXCEL Then blSkipColumn = True 'PBOOKEXCEL
                If ilCSV + 1 = RBOOKEXCEL Then blSkipColumn = True 'RBOOKEXCEL
            End If
            
            If blSkipColumn = False Then
                If blNeedComma Then slCSVRecord = slCSVRecord & ",": blNeedComma = False
                If IsNumeric(slCSV(ilCSV)) Then
                    slCSVRecord = slCSVRecord & Trim(slCSV(ilCSV))
                Else
                    If Trim(slCSV(ilCSV)) <> "" Then slCSVRecord = slCSVRecord & """" & Trim(slCSV(ilCSV)) & """"
                End If
                blNeedComma = True
            End If
        Next ilCSV
        Print #hmToCSV, slCSVRecord
    End If
    
    If (slRowType <> "") Then
        tmFormulaInfo(UBound(tmFormulaInfo)).sRowType = slRowType
        tmFormulaInfo(UBound(tmFormulaInfo)).iExcelRow = imExcelRow
        tmFormulaInfo(UBound(tmFormulaInfo)).lPop = llPop
        tmFormulaInfo(UBound(tmFormulaInfo)).dExtTotal = dlExtTotal 'TTP 10439 - Rerate 21,000,000
        ReDim Preserve tmFormulaInfo(0 To UBound(tmFormulaInfo) + 1) As FORMULAINFO
    End If
    If (rbcColumnLayout(1).Value = True) And (slDelimiter <> "") Then 'Columns: PR, PR, PR
        slRecordsArray = Split(slInRecord, smDelimiter)
        If Not IsArray(slRecordsArray) Then
            Exit Sub
        End If
        slRecord = ""
        'TTP 10082 - merge header into columns
        If rbcLayout(0).Value = True Then
            'Original header/Detail layout
            If UBound(slRecordsArray) = 0 Then
                slRecord = slRecord & slRecordsArray(0)
            Else
                For ilCell = 0 To LASTSTATICCOLEXCEL - 1 Step 1 '0 - 7
                    slRecord = slRecord & slRecordsArray(ilCell) & smDelimiter
                Next ilCell
            End If
            If UBound(slRecordsArray) >= PBOOKEXCEL Then
                For ilCell = LASTSTATICCOLEXCEL To PBOOKEXCEL - 10 Step 1 '8 - 16
                    slRecord = slRecord & slRecordsArray(ilCell) & smDelimiter
                    slRecord = slRecord & slRecordsArray(ilCell + 9) & smDelimiter
                Next ilCell
                'Get Index
                If UBound(slRecordsArray) >= INDEXEXCEL - 1 Then
                    slRecord = slRecord & slRecordsArray(INDEXEXCEL - 1)
                End If
            End If
        ElseIf rbcLayout(1).Value = True Then
            'contract header merged into columns
            If UBound(slRecordsArray) = 0 Then
                slRecord = slRecord & slRecordsArray(0)
            Else
                For ilCell = 0 To (LASTSTATICCOLEXCEL - 1) + 10 Step 1
                    slRecord = slRecord & slRecordsArray(ilCell) & smDelimiter
                Next ilCell
            End If
            If UBound(slRecordsArray) >= RBOOKEXCEL Then
                For ilCell = (LASTSTATICCOLEXCEL) + 10 To RBOOKEXCEL - 10 Step 1
                    slRecord = slRecord & slRecordsArray(ilCell) & smDelimiter
                    slRecord = slRecord & slRecordsArray(ilCell + 9) & smDelimiter
                Next ilCell
                'Get Index
                If UBound(slRecordsArray) >= INDEXEXCEL - 1 Then
                    slRecord = slRecord & slRecordsArray(INDEXEXCEL - 1)
                End If
            End If
        End If
                
        'JW 10/14/21 - Fix v81 TTP 10258 issues per Jason Email: Thu 10/14/21 10:13 AM (#1,2)
        If ExptReRate.ckcCsv.Value = vbUnchecked Then
            ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, imExcelRow, ilColumn, slDelimiter)
            If (Not bmInSummaryMode) And (InStr(1, slInRecord, "Line#") <= 0) And (InStr(1, slInRecord, "Contract Total") <= 0) And (InStr(1, slInRecord, "Advertiser Total") <= 0) And (InStr(1, slInRecord, "Purchased") <= 0) And (InStr(1, slInRecord, "Bonus Total") <= 0) And (InStr(1, slInRecord, "Purch") <= 0) Then
                ilRet = gParseItem(slInRecord, imPurchasedColumn, smDelimiter, slField)
                If slField <> "" Or (InStr(1, slInRecord, "MG:") > 0) Or (InStr(1, slInRecord, "Bonus") > 0) Then
                    blPRPRColorSwitch = True
                    For ilCell = imPurchasedColumn To omSheet.UsedRange.Columns.Count - 1 Step 2
                        blSkipColor = False
                        If ckcCost.Value = vbUnchecked Then
                            For ilCol = LBound(imCostColumn) To UBound(imCostColumn) Step 1
                                If imCostColumn(ilCol) = ilCell Then
                                    blSkipColor = True
                                End If
                            Next ilCol
                        End If
                        If ckcCPM.Value = vbUnchecked Then
                            For ilCol = LBound(imCPMCPPColumn) To UBound(imCPMCPPColumn) Step 1
                                If imCPMCPPColumn(ilCol) = ilCell Then
                                    blSkipColor = True
                                End If
                            Next ilCol
                        End If
                        If ckcRating.Value = vbUnchecked Then
                            For ilCol = LBound(imRatingColumn) To UBound(imRatingColumn) Step 1
                                If imRatingColumn(ilCol) = ilCell Then
                                    blSkipColor = True
                                End If
                            Next ilCol
                        End If
                        If Not blSkipColor Then
                            If blPRPRColorSwitch Then
                                blPRPRColorSwitch = False
                                If ExptReRate.ckcCsv.Value = vbUnchecked Then
                                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow, ilCell)
                                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow, ilCell + 1)
                                End If
                            Else
                                blPRPRColorSwitch = True
                                If ExptReRate.ckcCsv.Value = vbUnchecked Then
                                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow, ilCell)
                                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow, ilCell + 1)
                                End If
                            End If
                        End If
                    Next ilCell
                End If
            End If
        End If
    Else
        'Columns: PPP, RRR
        'JW 10/14/21 - Fix v81 TTP 10258 issues per Jason Email: Thu 10/14/21 10:13 AM (#1,2)
        If ExptReRate.ckcCsv.Value = vbUnchecked Then
            ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, imExcelRow, ilColumn, slDelimiter)
            If slDelimiter <> "" Then
                If (Not bmInSummaryMode) And (InStr(1, slInRecord, "Line#") <= 0) And (InStr(1, slInRecord, "Contract Total") <= 0) And (InStr(1, slInRecord, "Advertiser Total") <= 0) And (InStr(1, slInRecord, "Purchased") <= 0) And (InStr(1, slInRecord, "Bonus Total") <= 0) Then
                    If ExptReRate.ckcCsv.Value = vbUnchecked Then
                        ilRet = gParseItem(slInRecord, imPurchasedColumn, slDelimiter, slField)
                        If slField <> "" Or (InStr(1, slInRecord, "MG:") > 0) Or (InStr(1, slInRecord, "Bonus") > 0) Then
                            For ilCol = imPurchasedColumn To imReRateColumn - 1 Step 1
                                ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow, ilCol)
                            Next ilCol
                        End If
                        ilRet = gParseItem(slInRecord, imReRateColumn, slDelimiter, slField)
                        If slField <> "" Then
                            For ilCol = imReRateColumn To omSheet.UsedRange.Columns.Count - 1 Step 1
                                If ExptReRate.ckcCsv.Value = vbUnchecked Then ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow, ilCol)
                            Next ilCol
                        End If
                    End If
                End If
            End If
        End If
    End If
    imExcelRow = imExcelRow + 1

End Sub

Private Sub mPopSpotLength()
    Dim ilRif As Integer
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilLen As Integer
    Dim ilTest As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim blFd As Boolean
    ReDim ilLength(0 To 0) As Integer
    
    ilMin = 9999
    ilMax = -1
    lbcSpotLens.Clear
    For ilRif = 0 To UBound(tgMRif) - 1 Step 1
        ilVefCode = tgMRif(ilRif).iVefCode
        ilVpfIndex = gBinarySearchVpf(ilVefCode)
        If ilVpfIndex <> -1 Then
            For ilLen = 0 To 9 Step 1
                If tgVpf(ilVpfIndex).iSLen(ilLen) > 0 Then
                    blFd = False
                    For ilTest = 0 To UBound(ilLength) - 1 Step 1
                        If tgVpf(ilVpfIndex).iSLen(ilLen) = ilLength(ilTest) Then
                            blFd = True
                            Exit For
                        End If
                    Next ilTest
                    If Not blFd Then
                        If tgVpf(ilVpfIndex).iSLen(ilLen) < ilMin Then
                            ilMin = tgVpf(ilVpfIndex).iSLen(ilLen)
                        End If
                        If tgVpf(ilVpfIndex).iSLen(ilLen) > ilMax Then
                            ilMax = tgVpf(ilVpfIndex).iSLen(ilLen)
                        End If
                        ilLength(UBound(ilLength)) = tgVpf(ilVpfIndex).iSLen(ilLen)
                        ReDim Preserve ilLength(0 To UBound(ilLength) + 1) As Integer
                    End If
                End If
            Next ilLen
        End If
    Next ilRif
    For ilLen = ilMin To ilMax Step 1
        For ilTest = 0 To UBound(ilLength) - 1 Step 1
            If ilLen = ilLength(ilTest) Then
                lbcSpotLens.AddItem Trim$(str$(ilLength(ilTest)))
                lbcSpotLens.ItemData(lbcSpotLens.NewIndex) = ilLength(ilTest)
                Exit For
            End If
        Next ilTest
    Next ilLen
    '7/1/21 - JW - Bonus improvement, ok'd per Jason; There's really never a reason to run it for a single spot length
    ckcAllSpotLens.Value = vbChecked
End Sub

Private Sub mSetCommands()
    Dim illoop As Integer
    Dim blOk As Boolean
    
    cmcCancel.Caption = "Cancel"
    cmcExport.Enabled = False
    cmcSetBook.Enabled = False
    If lbcAdvertiser.ListIndex < 0 Then
        Exit Sub
    End If
    If Not mAnyGridRowSelected() Then
        Exit Sub
    End If
    blOk = False
    For illoop = 0 To lbcSpotLens.ListCount - 1 Step 1
        If lbcSpotLens.Selected(illoop) Then
            blOk = True
            Exit For
        End If
    Next illoop
    If Not blOk Then
        Exit Sub
    End If
    'TTP 10140, show When [Set Book by Line] method is enabled - Keep track of Last "Research Book Name" option (rbcReRateBook)
    If rbcReRateBook(2).Value = False Then
        cmcSetBook.Enabled = True
    End If
    If rbcDatesBy(0).Value Then 'By Week
        If edcDate(0).Text = "" Then
            Exit Sub
        End If
    ElseIf rbcDatesBy(1).Value Then   'By Month
        If edcStart.Text = "" Then
            Exit Sub
        End If
    ElseIf rbcDatesBy(2).Value Then   'By Quarter
        If edcYear.Text = "" Then
            Exit Sub
        End If
    ElseIf rbcDatesBy(3).Value Then   'By Contract
        If edcDate(0).Text = "" Then
            Exit Sub
        End If
    ElseIf rbcDatesBy(4).Value Then   'By Range
        If (edcDate(1).Text = "") Or (edcDate(2).Text = "") Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    blOk = False
    For illoop = 0 To rbcDemo.UBound Step 1
        If rbcDemo(illoop).Value Then
            blOk = True
            Exit For
        End If
    Next illoop
    If (blOk = True) And (rbcDemo(4).Value = True) Then
        If cbcDemo.ListIndex < 0 Then
            blOk = False
        End If
    End If
    If Not blOk Then
        Exit Sub
    End If
    blOk = False
    For illoop = 0 To rbcShow.UBound Step 1
        If rbcShow(illoop).Value Then
            blOk = True
            Exit For
        End If
    Next illoop
    If Not blOk Then
        Exit Sub
    End If
    blOk = False
    For illoop = 0 To rbcColumnLayout.UBound Step 1
        If rbcColumnLayout(illoop).Value Then
            blOk = True
            Exit For
        End If
    Next illoop
    If Not blOk Then
        Exit Sub
    End If
    blOk = False
    For illoop = 0 To rbcReRateBook.UBound Step 1
        If rbcReRateBook(illoop).Value Then
            blOk = True
            Exit For
        End If
    Next illoop
    'TTP 10140, show When [Set Book by Line] method is enabled - Keep track of Last "Research Book Name" option (rbcReRateBook)
    If rbcReRateBookByLine.Value = True Then
        blOk = True
    End If

    If Not blOk Then
        Exit Sub
    End If
    blOk = False
    For illoop = 0 To rbcIndex.UBound Step 1
        If rbcIndex(illoop).Value Then
            blOk = True
            Exit For
        End If
    Next illoop
    If Not blOk Then
        Exit Sub
    End If
    'if not specified, line book is used for the population
    'If cbcPopBookNames.ListIndex < 0 Then
    '    Exit Sub
    'End If
    
    'JW 10/14/21 - Fix v81 TTP 10258 issues per Jason call 10/14/21 - no PRPRPR mode for CSV
    If rbcColumnLayout(1).Value = True Then
        ckcCsv.Enabled = False
    Else
        ckcCsv.Enabled = True
    End If
    If ckcCsv.Value = vbChecked Then
        rbcColumnLayout(1).Enabled = False
    Else
        rbcColumnLayout(1).Enabled = True
    End If
    
    cmcExport.Enabled = True
End Sub

Private Sub mGridColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim illoop As Integer

    grdCntr.ColWidth(NODEMOSINDEX) = 0
    grdCntr.ColWidth(ENDREVNOINDEX) = 0
    grdCntr.ColWidth(STARTREVNOINDEX) = 0
    grdCntr.ColWidth(PURCHASECHFCODEINDEX) = 0
    grdCntr.ColWidth(SELECTEDINDEX) = 0
    grdCntr.ColWidth(GENINDEX) = 0.1 * grdCntr.Width
    grdCntr.ColWidth(PRODUCTINDEX) = 0.53 * grdCntr.Width
    grdCntr.ColWidth(CNTRNOINDEX) = 0.15 * grdCntr.Width
    grdCntr.ColWidth(VERSIONINDEX) = 0.22 * grdCntr.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdCntr.Width
    For ilCol = 0 To grdCntr.cols - 1 Step 1
        llWidth = llWidth + grdCntr.ColWidth(ilCol)
        If (grdCntr.ColWidth(ilCol) > 15) And (grdCntr.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdCntr.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdCntr.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdCntr.Width
            For ilCol = 0 To grdCntr.cols - 1 Step 1
                If (grdCntr.ColWidth(ilCol) > 15) And (grdCntr.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdCntr.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdCntr.FixedCols To grdCntr.cols - 1 Step 1
                If grdCntr.ColWidth(ilCol) > 15 Then
                    ilColInc = grdCntr.ColWidth(ilCol) / llMinWidth
                    For illoop = 1 To ilColInc Step 1
                        grdCntr.ColWidth(ilCol) = grdCntr.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next illoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub

Private Sub mGridColumnTitles()

    grdCntr.Row = 0
    grdCntr.Col = GENINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    'grdCntr.CellBackColor = vbWhite
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Gen"
    grdCntr.Col = PRODUCTINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    'grdCntr.CellBackColor = vbWhite
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Product"
    grdCntr.Col = CNTRNOINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    'grdCntr.CellBackColor = vbWhite
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Contract"
    grdCntr.Col = VERSIONINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    'grdCntr.CellBackColor = vbWhite
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Purchase"
    
    grdCntr.Row = 1
    grdCntr.Col = GENINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    'grdCntr.CellBackColor = vbWhite
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = ""
    grdCntr.Col = PRODUCTINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    'grdCntr.CellBackColor = vbWhite
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Name"
    grdCntr.Col = CNTRNOINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    'grdCntr.CellBackColor = vbWhite
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Number"
    grdCntr.Col = VERSIONINDEX
    grdCntr.CellFontBold = False
    grdCntr.CellFontName = "Arial"
    grdCntr.CellFontSize = 6.75
    grdCntr.CellForeColor = vbBlue
    'grdCntr.CellBackColor = vbWhite
    grdCntr.TextMatrix(grdCntr.Row, grdCntr.Col) = "Revision #"
    grdCntr.Col = PURCHASECHFCODEINDEX
End Sub

Function mTestShowFill(slPriceType As String) As String

    Dim ilLoopOnAdvt As Integer
    
    If slPriceType <> "-" And slPriceType <> "+" Then     'neither a - or +, fill wasnt overridden then use advt to determine how to show
        ilLoopOnAdvt = gBinarySearchAdf(imAdfCode)
        If ilLoopOnAdvt <> -1 Then
            If tgCommAdf(ilLoopOnAdvt).sBonusOnInv = "N" Then
                If ckcHideBonus.Value = vbChecked Then
                    mTestShowFill = "Y"
                Else
                    mTestShowFill = "N"
                End If
            Else
                If ckcInvBonus.Value = vbChecked Then
                    mTestShowFill = "Y"
                Else
                    mTestShowFill = "N"
                End If
            End If
        Else
            If ckcHideBonus.Value = vbChecked Then
                mTestShowFill = "Y"
            Else
                mTestShowFill = "N"
            End If
        End If
    Else                'was overrriden in fill screen, use spot
        If slPriceType = "-" Then
            If ckcHideBonus.Value = vbChecked Then
                mTestShowFill = "Y"
            Else
                mTestShowFill = "N"
            End If
        Else
            If ckcInvBonus.Value = vbChecked Then
                mTestShowFill = "Y"
            Else
                mTestShowFill = "N"
            End If
        End If
    End If

   Exit Function
End Function

Private Sub mGetMGAud(llSdf As Long, ilVefCode As Integer, ilDnfCode As Integer, ilClf As Integer, ilMnfDemo As Integer, llStartDate As Long, llAvgAud As Long, llPopEst As Long, ilAudFromSource As Integer, llAudFromCode As Long, ilTotLnSpts As Integer, llTotalCntrSpots As Long)
    Dim ilRet As Integer
    Dim ilCff As Integer
    Dim illoop As Integer
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim llDate As Long
    Dim ilUpperWk As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilSpots As Integer
    Dim llPrice As Long
    ReDim ilInputDays(0 To 6) As Integer
    
    If imMGDetailDnfCode = -1 Then
        imMGDetailDnfCode = ilDnfCode
    ElseIf imMGDetailDnfCode <> ilDnfCode And imMGDetailDnfCode <> -2 Then
        imMGDetailDnfCode = -2
    End If
    ilSpots = 1
    For illoop = 0 To 6                 'init all days to not airing, setup for research results later
        ilInputDays(illoop) = False
    Next illoop
    llPrice = 0
    gUnpackDateLong tmSdfExt(llSdf).iDate(0), tmSdfExt(llSdf).iDate(1), llDate
    ilInputDays(gWeekDayLong(llDate)) = True
    gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llOvStartTime
    gUnpackTimeLong tmSdfExt(llSdf).iTime(0), tmSdfExt(llSdf).iTime(1), False, llOvEndTime
    llOvEndTime = llOvEndTime + tmSdfExt(llSdf).iLen
    ''ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, imReRateDnfCode, ilVefCode, 0, ilMnfDemo, llDate, llDate, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
    'ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, imReRateDnfCode, ilVefCode, 0, ilMnfDemo, llDate, llDate, 0, llOvStartTime, llOvEndTime, ilInputDays(), "S", 0, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, 0, ilMnfDemo, llDate, llDate, 0, llOvStartTime, llOvEndTime, ilInputDays(), "S", 0, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
    ''Loop and build avg aud, spots, & spots per week
    ''calc week index
    ''ilUpperWk = (llDate - llStartDate) / 7 + 1
    'ilUpperWk = UBound(lmWklyspots)
    'If ilUpperWk - 1 > UBound(lmWklyspots) Then
    '    ReDim Preserve lmWklyspots(0 To ilUpperWk) As Integer
    '    ReDim Preserve lmWklyAvgAud(0 To ilUpperWk) As Long
    '    ReDim Preserve lmWklyPopEst(0 To ilUpperWk) As Long
    'End If
    'lmWklyspots(ilUpperWk - 1) = lmWklyspots(ilUpperWk - 1) + ilSpots
    ilTotLnSpts = ilTotLnSpts + ilSpots
    If tmClfR.sType = "S" Then
        llTotalCntrSpots = llTotalCntrSpots + ilSpots
    End If
    'Obtain price from Missed spot
    'lmWklyRates(ilUpperWk - 1) = 0
    ilCff = tmClfReRate(ilClf).iFirstCff
    Do While ilCff <> -1
        tmCff = tmCffReRate(ilCff).CffRec
        gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llFltStart
        gUnpackDateLong tmCff.iEndDate(0), tmCff.iEndDate(1), llFltEnd
        If (tmSdfExt(llSdf).lMdDate >= llFltStart) And (tmSdfExt(llSdf).lMdDate <= llFltEnd) Then
            'lmWklyRates(ilUpperWk - 1) = tmCff.lActPrice
            llPrice = tmCff.lActPrice
            Exit Do
        End If
        ilCff = tmCffReRate(ilCff).iNextCff
    Loop
    ''lmWklyAvgAud(ilUpperWk - 1) = llAvgAud
    'lmWklyAvgAud(ilUpperWk - 1) = lmWklyAvgAud(ilUpperWk - 1) + llAvgAud
    'lmWklyPopEst(ilUpperWk - 1) = llPopEst
    mAddSpotsToWeekArray llDate, llStartDate, ilSpots, llAvgAud, tmCff.lActPrice, llPopEst, True
End Sub

Private Sub mAddSpotsToWeekArray(llDate As Long, llStartDate As Long, ilSpots As Integer, llAvgAud As Long, llPrice As Long, llPop As Long, blSumAvg As Boolean)
    Dim ilUpperWk As Integer
    Dim ilWkIndex As Integer
    Dim ilWk As Integer
    Dim blFd As Boolean
    Dim ilPurchaseUpper As Integer
    Dim llMoDate As Long
    Dim llUB As Long
    
    llMoDate = gObtainPrevMondayLong(llDate)
    ilUpperWk = (llDate - llStartDate) / 7 + 1
    ilWkIndex = ilUpperWk - 1
'Debug.Print "mAddSpotsToWeekArray(llDate=" & llDate & ", llStartDate=" & llStartDate & ",ilSpots=" & ilSpots & ",llAvgAud=" & llAvgAud&; ",llPrice=" & llPrice&; ",llPop=" & llPop&; ",blSumAvg=" & blSumAvg & ")"
    
    'ilUpperWk = UBound(lmWklyspots)
    ilPurchaseUpper = UBound(lmWklyspots)
    If ilWkIndex > UBound(lmWklyspots) Then
        ReDim Preserve lmWklyspots(0 To ilWkIndex) As Long
        ReDim Preserve lmWklyRates(0 To ilWkIndex) As Long
        ReDim Preserve lmWklyAvgAud(0 To ilWkIndex) As Long
        ReDim Preserve lmWklyPopEst(0 To ilWkIndex) As Long
        ReDim Preserve lmWklyMoDate(0 To ilWkIndex) As Long
        For ilWk = ilPurchaseUpper + 1 To ilWkIndex Step 1
            lmWklyspots(ilWk) = 0
            lmWklyRates(ilWk) = 0
            lmWklyAvgAud(ilWk) = 0
            lmWklyPopEst(ilWk) = 0
            lmWklyMoDate(ilWk) = 0
        Next ilWk
        ReDim Preserve imRtg(0 To ilWkIndex)                   'setup arrays for return values from audtolnresearch
        ReDim Preserve lmGrimp(0 To ilWkIndex)
        ReDim Preserve lmGRP(0 To ilWkIndex)
        ReDim Preserve lmCost(0 To ilWkIndex)                    'setup arrays for return values from audtolnresearch
    Else
        blFd = False
        For ilWk = 0 To UBound(lmWklyspots) Step 1
            If (lmWklyRates(ilWk) = 0) And (lmWklyAvgAud(ilWk) = 0) And (lmWklyPopEst(ilWk) = 0) Then
                ilWkIndex = ilWk
                blFd = True
                Exit For
            End If
            If blSumAvg = True And lmWklyspots(ilWk) > 0 Then
                If (lmWklyRates(ilWk) = llPrice) And (lmWklyAvgAud(ilWk) = (llAvgAud * lmWklyspots(ilWk))) And (lmWklyPopEst(ilWk) = llPop) And (lmWklyMoDate(ilWk) = llMoDate) Then
                    ilWkIndex = ilWk
                    blFd = True
                    Exit For
                End If
            Else
                If (lmWklyRates(ilWk) = llPrice) And (lmWklyAvgAud(ilWk) = llAvgAud) And (lmWklyPopEst(ilWk) = llPop) And (lmWklyMoDate(ilWk) = llMoDate) Then
                    ilWkIndex = ilWk
                    blFd = True
                    Exit For
                End If
            End If
        Next ilWk
        If Not blFd Then
            llUB = UBound(lmWklyspots)
            ReDim Preserve lmWklyspots(0 To llUB + 1) As Long
            ReDim Preserve lmWklyRates(0 To llUB + 1) As Long
            ReDim Preserve lmWklyAvgAud(0 To llUB + 1) As Long
            ReDim Preserve lmWklyPopEst(0 To llUB + 1) As Long
            ReDim Preserve lmWklyMoDate(0 To llUB + 1) As Long
            ilWkIndex = UBound(lmWklyspots)
            For ilWk = ilPurchaseUpper + 1 To ilWkIndex Step 1
                lmWklyspots(ilWk) = 0
                lmWklyRates(ilWk) = 0
                lmWklyAvgAud(ilWk) = 0
                lmWklyPopEst(ilWk) = 0
                lmWklyMoDate(ilWk) = 0
            Next ilWk
            ReDim Preserve imRtg(0 To UBound(lmWklyspots))                   'setup arrays for return values from audtolnresearch
            ReDim Preserve lmGrimp(0 To UBound(lmWklyspots))
            ReDim Preserve lmGRP(0 To UBound(lmWklyspots))
            ReDim Preserve lmCost(0 To UBound(lmWklyspots))                    'setup arrays for return values from audtolnresearch
        End If
    End If
    lmWklyspots(ilWkIndex) = lmWklyspots(ilWkIndex) + ilSpots
    'ilTotLnSpts = ilTotLnSpts + ilSpots
    'If tmClf.sType = "S" Then
    '    llTotalCntrSpots = llTotalCntrSpots + ilSpots
    'End If
    lmWklyRates(ilWkIndex) = llPrice
    'If (ckcMG.Value = vbChecked) Then
    If Not blSumAvg Then
        lmWklyAvgAud(ilWkIndex) = llAvgAud
    Else
        lmWklyAvgAud(ilWkIndex) = lmWklyAvgAud(ilWkIndex) + llAvgAud
    End If
    lmWklyPopEst(ilWkIndex) = llPop
    lmWklyMoDate(ilWkIndex) = llMoDate
End Sub

Private Function mGetDemo(ilDemoCode As Integer) As String
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    slSQLQuery = "Select mnfName from MNF_Multi_Names where mnfCode = " & ilDemoCode
    Set tmp_rst = gSQLSelectCall(slSQLQuery)
    If Not tmp_rst.EOF Then
        mGetDemo = Trim$(tmp_rst!mnfName)
    Else
        mGetDemo = ""
    End If
End Function
Private Sub mClearGrid()
    Dim llRow As Long
    Dim ilCol As Integer
    bmInGrid = True
    ckcAllCntr.Value = vbUnchecked
    bmInGrid = False
    grdCntr.Rows = 3
    gGrid_IntegralHeight grdCntr, fgBoxGridH + 15
    gGrid_FillWithRows grdCntr, fgBoxGridH + 15
    grdCntr.Height = grdCntr.Height + 60
    For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        grdCntr.RowHeight(llRow) = fgFlexGridRowH
        grdCntr.Row = llRow
        For ilCol = 0 To grdCntr.cols - 1 Step 1
            If (ilCol = PURCHASECHFCODEINDEX) Then
                grdCntr.TextMatrix(llRow, ilCol) = 0
            Else
                grdCntr.TextMatrix(llRow, ilCol) = ""
            End If
            grdCntr.Col = ilCol
            grdCntr.CellBackColor = WHITE
            'TTP 10380 - ReRate report: contract list does not clear red font color that indicates CBS contract when switching date period
            grdCntr.CellForeColor = vbBlack
        Next ilCol
        grdCntr.CellBackColor = WHITE
    Next llRow
    lmScrollTop = grdCntr.FixedRows
    mSetDemo
    lmLastClickedRow = -1
    grdCntr.Row = 0
    grdCntr.Col = PURCHASECHFCODEINDEX
End Sub

Private Function mAnyGridRowSelected() As Boolean
    Dim llRow As Long
    
    mAnyGridRowSelected = False
    For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        'If grdCntr.TextMatrix(llRow, PRODUCTINDEX) <> "" And grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
        If grdCntr.TextMatrix(llRow, CNTRNOINDEX) <> "" And grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
            mAnyGridRowSelected = True
            Exit For
        End If
    Next llRow
End Function
Private Sub mTerminate(ilFromCancel As Integer)
'
'   mTerminate
'   Where:
'
    If ilFromCancel Then
        igRptReturn = True  'Show list
    Else
        igRptReturn = False
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload ExptReRate
    igManUnload = NO
End Sub
Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    
    grdCntr.Row = llRow
    For llCol = GENINDEX To VERSIONINDEX Step 1
        grdCntr.Col = llCol
        If grdCntr.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            If llCol = GENINDEX Then
                grdCntr.CellFontName = "Monotype Sorts"
                grdCntr.TextMatrix(llRow, GENINDEX) = ""
            End If
            'grdCntr.CellBackColor = vbWhite
            'grdCntr.CellForeColor = vbWindowText
        Else
            If llCol = GENINDEX Then
                grdCntr.CellFontName = "Monotype Sorts"
                grdCntr.TextMatrix(llRow, GENINDEX) = "4"
            End If
            'grdCntr.CellBackColor = vbHighlight
            'grdCntr.CellForeColor = vbWhite
        End If
        If llCol >= GENINDEX + 1 And llCol < VERSIONINDEX Then
            'grdCntr.CellBackColor = LIGHTBLUE
            'If grdCntr.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            '    grdCntr.CellBackColor = vbWhite
            '    grdCntr.CellForeColor = vbWindowText
            'Else
            '    grdCntr.CellBackColor = vbHighlight
            '    grdCntr.CellForeColor = vbWhite
            'End If
            grdCntr.CellBackColor = LIGHTYELLOW
        End If
    Next llCol
End Sub


Private Sub mDetermineDateRange(slStartDate As String, slEndDate As String)
    Dim slStr As String

    If rbcDatesBy(0).Value And edcDate(0).Text <> "" Then 'By Week
        slStartDate = Format$(edcDate(0).Text, "m/d/yy")               'reformat date to insure year is there
        slEndDate = DateAdd("d", 6, slStartDate)
    ElseIf rbcDatesBy(1).Value And (edcStart.Text <> "") And (edcYear.Text <> "") Then  'By Month
        slStr = edcStart.Text & "/15/" & edcYear.Text
        slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
        slEndDate = gObtainEndStd(slStartDate)
    ElseIf rbcDatesBy(2).Value And (edcStart.Text <> "") And (edcYear.Text <> "") Then    'By Quarter
        If edcStart.Text = 2 Then
            slStr = "4/15/" & edcYear.Text
            slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
            slStr = "6/15/" & edcYear.Text
            slEndDate = gObtainEndStd(slStr)
        ElseIf edcStart.Text = 3 Then
            slStr = "7/15/" & edcYear.Text
            slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
            slStr = "9/15/" & edcYear.Text
            slEndDate = gObtainEndStd(slStr)
        ElseIf edcStart.Text = 4 Then
            slStr = "10/15/" & edcYear.Text
            slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
            slStr = "12/15/" & edcYear.Text
            slEndDate = gObtainEndStd(slStr)
        Else
            slStr = "1/15/" & edcYear.Text
            slStartDate = gObtainStartStd(slStr)               'reformat date to insure year is there
            slStr = "3/15/" & edcYear.Text
            slEndDate = gObtainEndStd(slStr)
        End If
    ElseIf rbcDatesBy(3).Value And edcDate(0).Text <> "" Then   'By Contract
        slStartDate = Format$(edcDate(0).Text, "m/d/yy")
        slEndDate = "12/31/2069"
    ElseIf rbcDatesBy(4).Value And (edcDate(1).Text <> "") And edcDate(2).Text <> "" Then   'By Range
        slStartDate = Format$(edcDate(1).Text, "m/d/yy")
        slEndDate = Format$(edcDate(2).Text, "m/d/yy")
    Else
        slStartDate = "1/1/1970"
        slEndDate = "12/31/2069"
    End If

End Sub

Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    mPopCntr
    mSetCommands
End Sub

Private Sub mSetDemo()
    Dim llRow As Long
    Dim ilMaxAllowedDemo As Integer
    Dim ilDemo As Integer
    
    ilMaxAllowedDemo = -1
    For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
        'If grdCntr.TextMatrix(llRow, PRODUCTINDEX) <> "" Then
        If grdCntr.TextMatrix(llRow, CNTRNOINDEX) <> "" Then
            If grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                If ilMaxAllowedDemo = -1 Then
                    ilMaxAllowedDemo = Val(grdCntr.TextMatrix(llRow, NODEMOSINDEX))
                Else
                    If Val(grdCntr.TextMatrix(llRow, NODEMOSINDEX)) < ilMaxAllowedDemo Then
                        ilMaxAllowedDemo = Val(grdCntr.TextMatrix(llRow, NODEMOSINDEX))
                    End If
                End If
            End If
        End If
    Next llRow
    For ilDemo = 1 To 4 Step 1
        If ilDemo <= ilMaxAllowedDemo Then
            rbcDemo(ilDemo - 1).Enabled = True
        Else
            rbcDemo(ilDemo - 1).Enabled = False
            If ilMaxAllowedDemo = -1 Then
                rbcDemo(ilDemo - 1).Value = False
                cbcDemo.Enabled False
            End If
        End If
    Next ilDemo
    If (rbcDemo(0).Enabled) And (rbcDemo(0).Value = False) And (rbcDemo(1).Value = False) And (rbcDemo(2).Value = False) And (rbcDemo(3).Value = False) And (rbcDemo(4).Value = False) Then
        rbcDemo(0).Value = True
    End If
End Sub

'Private Function mObtainReRatePopulation() As Long
'    Dim ilSdf As Integer
'
'
'End Function

Private Sub mSortRcf()
    Dim ilRcf As Integer
    Dim slKey As String
    ReDim tmRateCardSort(0 To UBound(tgMRcf)) As RATECARDSORT
    
    For ilRcf = 0 To UBound(tgMRcf) - 1 Step 1
        gUnpackDateForSort tgMRcf(ilRcf).iStartDate(0), tgMRcf(ilRcf).iStartDate(1), slKey
        tmRateCardSort(ilRcf).sKey = slKey
        gUnpackDateLong tgMRcf(ilRcf).iStartDate(0), tgMRcf(ilRcf).iStartDate(1), tmRateCardSort(ilRcf).lDate
        tmRateCardSort(ilRcf).iRcfIndex = ilRcf
        tmRateCardSort(ilRcf).iRcfCode = tgMRcf(ilRcf).iCode
    Next ilRcf
    If UBound(tmRateCardSort) > 0 Then
        'Sort in descending order
        ArraySortTyp fnAV(tmRateCardSort(), 0), UBound(tmRateCardSort), 1, LenB(tmRateCardSort(0)), 0, LenB(tmRateCardSort(0).sKey), 0
    End If
    
End Sub

Private Function mFindDaypart(ilVefCode As Integer, llDate As Long, llTime As Long) As Integer
    Dim ilSort As Integer
    Dim ilRcf As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilIndex As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDay As Integer
    Dim llBFTime As Long
    Dim ilBFDays As Integer
    Dim ilBFRdfIndex As Integer
    Dim ilDateDay As Integer
    Dim ilDayCount As Integer
    
    ilDateDay = gWeekDayLong(llDate)
    ilBFRdfIndex = -1

    For ilSort = 0 To UBound(tmRateCardSort) - 1 Step 1
        If llDate >= tmRateCardSort(ilSort).lDate Then
            For llRif = 0 To UBound(tgMRif) - 1 Step 1
                If (tgMRif(llRif).iRcfCode = tmRateCardSort(ilSort).iRcfCode) And (tgMRif(llRif).iVefCode = ilVefCode) Then
                    ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                    If ilRdf <> -1 Then
                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDateDay) = "Y") Then
                                    gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llStartTime
                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llEndTime
                                    If (llTime >= llStartTime) And (llTime <= llEndTime) Then
                                        ilDayCount = 0
                                        For ilDay = 0 To 6 Step 1
                                            If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y") Then
                                                ilDayCount = ilDayCount + 1
                                            End If
                                        Next ilDay
                                        If ilBFRdfIndex = -1 Then
                                            llBFTime = llEndTime - llStartTime
                                            ilBFDays = ilDayCount
                                            ilBFRdfIndex = ilRdf
                                        Else
                                            If (Asc(tgSpf.sUsingFeatures) And BESTFITWEIGHT) <> BESTFITWEIGHT Then  'Test which is best fit
                                                If llEndTime - llStartTime < llBFTime Then
                                                    llBFTime = llEndTime - llStartTime
                                                    ilBFDays = ilDayCount
                                                    ilBFRdfIndex = ilRdf
                                                ElseIf (llEndTime - llStartTime = llBFTime) And (ilDayCount < ilBFDays) Then
                                                    llBFTime = llEndTime - llStartTime
                                                    ilBFDays = ilDayCount
                                                    ilBFRdfIndex = ilRdf
                                                End If
                                            Else
                                                'days first
                                                If ilDayCount < ilBFDays Then
                                                    llBFTime = llEndTime - llStartTime
                                                    ilBFDays = ilDayCount
                                                    ilBFRdfIndex = ilRdf
                                                ElseIf (ilDayCount = ilBFDays) And (llEndTime - llStartTime < llBFTime) Then
                                                    llBFTime = llEndTime - llStartTime
                                                    ilBFDays = ilDayCount
                                                    ilBFRdfIndex = ilRdf
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next ilIndex
                    End If
                End If
            Next llRif
            If ilBFRdfIndex <> -1 Then
                Exit For
            End If
        End If
    Next ilSort
    If ilBFRdfIndex = -1 Then
        mFindDaypart = 0
    Else
        mFindDaypart = tgMRdf(ilBFRdfIndex).iCode
    End If
End Function

Private Sub mFilterLinesBySpotLength(tlClf() As CLFLIST)
    Dim blAllSelected As Boolean
    Dim illoop As Integer
    Dim ilClf As Integer
    Dim ilPkg As Integer
    Dim ilIndex As Integer
    Dim blAnyRemoved As Boolean
    
    blAllSelected = True
    For illoop = 0 To lbcSpotLens.ListCount - 1 Step 1
        If Not lbcSpotLens.Selected(illoop) Then
            blAllSelected = False
            Exit For
        End If
    Next illoop
    If blAllSelected Then
        Exit Sub
    End If
    blAnyRemoved = False
    For ilClf = 0 To UBound(tlClf) - 1 Step 1
        For illoop = 0 To lbcSpotLens.ListCount - 1 Step 1
            If tlClf(ilClf).ClfRec.iLen = lbcSpotLens.List(illoop) Then
                If lbcSpotLens.Selected(illoop) = False Then
                    tlClf(ilClf).ClfRec.lCode = -tlClf(ilClf).ClfRec.lCode
                    blAnyRemoved = True
                    If (tlClf(ilClf).ClfRec.sType = "O") Or (tlClf(ilClf).ClfRec.sType = "A") Then
                        For ilPkg = 0 To UBound(tlClf) - 1 Step 1
                            If tlClf(ilClf).ClfRec.iLine = tlClf(ilPkg).ClfRec.iPkLineNo Then
                                tlClf(ilPkg).ClfRec.lCode = -tlClf(ilPkg).ClfRec.lCode
                            End If
                        Next ilPkg
                    End If
                End If
            End If
        Next illoop
    Next ilClf
    If blAnyRemoved Then
        'Remove lines
        ReDim tlTemp(0 To UBound(tlClf)) As CLFLIST
        For ilClf = 0 To UBound(tlClf) Step 1
            tlTemp(ilClf) = tlClf(ilClf)
        Next ilClf
        ilIndex = 0
        For ilClf = 0 To UBound(tlTemp) - 1 Step 1
            If tlTemp(ilClf).ClfRec.lCode > 0 Then
                tlClf(ilIndex) = tlTemp(ilClf)
                ilIndex = ilIndex + 1
            End If
        Next ilClf
        ReDim Preserve tlClf(0 To ilIndex) As CLFLIST
    End If
End Sub

Private Sub mMergeMissingLines()
    'Merge ReRate lines missing in Purchase into purchase
    Dim ilClfP As Integer
    Dim ilClfR As Integer
    Dim blFound As Boolean
    Dim ilClfUpper As Integer
    Dim ilCffUpper As Integer
    Dim ilPrevCffUpper As Integer
    Dim ilNext As Integer
    
    For ilClfR = 0 To UBound(tmClfReRate) - 1 Step 1
        blFound = False
        For ilClfP = 0 To UBound(tmClfPurchase) - 1 Step 1
            If tmClfReRate(ilClfR).ClfRec.iLine = tmClfPurchase(ilClfP).ClfRec.iLine Then
                blFound = True
                Exit For
            End If
        Next ilClfP
        If Not blFound Then
            ''Add the line and set flight as either not defined or one week with zero spots
            ''Test if will be shown
            'ilClfUpper = UBound(tmClfPurchase)
            'tmClfPurchase(ilClfUpper) = tmClfReRate(ilClfR)
            'tmClfPurchase(ilClfUpper).ClfRec.iDnfCode = 0
            'tmClfPurchase(ilClfUpper).iFirstCff = -1
            'ilNext = tmClfReRate(ilClfR).iFirstCff
            'Do While ilNext <> -1
            '    ilCffUpper = UBound(tmCffPurchase)
            '    tmCffPurchase(ilCffUpper) = tmCffReRate(ilNext)
            '    tmCffPurchase(ilCffUpper).CffRec.sDyWk = "W"
            '    tmCffPurchase(ilCffUpper).CffRec.iSpotsWk = 0
            '    tmCffPurchase(ilCffUpper).CffRec.lPropPrice = 0
            '    tmCffPurchase(ilCffUpper).CffRec.lActPrice = 0
            '    If tmClfPurchase(ilClfUpper).iFirstCff = -1 Then
            '        tmClfPurchase(ilClfUpper).iFirstCff = ilCffUpper
            '    Else
            '        tmCffPurchase(ilPrevCffUpper).iNextCff = ilCffUpper
            '    End If
            '    ilPrevCffUpper = ilCffUpper
            '    ReDim Preserve tmCffPurchase(0 To ilCffUpper + 1) As CFFLIST
            '    ilNext = tmCffReRate(ilNext).iNextCff
            'Loop
            'ReDim Preserve tmClfPurchase(0 To ilClfUpper + 1) As CLFLIST
            mCopyLines tmClfReRate(ilClfR), tmClfPurchase(), tmCffReRate(), tmCffPurchase(), tmClfReRate(ilClfR).ClfRec.iPkLineNo
        End If
    Next ilClfR
    
End Sub

Private Sub mOutputSummary()
    Dim slDelimiter As String
    Dim slRecord As String
    Dim ilRec As Integer
    Dim ilRet As Integer
    Dim slOrder As String
    Dim slProduct As String
    Dim ilFieldNo As Integer
    Dim ilPos As Integer
    Dim blColumnOneShow As Boolean
    Dim blTitle1Shown As Boolean
    Dim blTitle2Shown As Boolean
    Dim slStr As String * 1
    Dim slBonus As String
    Dim slIndex As String
    Dim slAgency As String
    Dim slRowType As String
    Dim slExtTotal As String
    Dim slPopulation As String
    Dim blAddSpaceLine As Boolean
    Dim ilCell As Integer
    Dim ilCol As Integer
    Dim blPRPRColorSwitch As Boolean
    Dim blNextRowAdvt As Boolean
    Dim ilAdvtTotalRec As Integer
    Dim blPurchPopVaries As Boolean
    Dim blReRatePopVaries As Boolean
    Dim ilPurchPopRec As Integer
    Dim ilReRatePopRec As Integer
    
    bmInSummaryMode = True
    blTitle1Shown = False
    blTitle2Shown = False
    blColumnOneShow = False
    blAddSpaceLine = False
    
    slAgency = ""
    For ilRec = 0 To UBound(smSummaryRecords) - 1 Step 1
        If InStr(1, smSummaryRecords(ilRec), "Agency:") > 0 Then
            If slAgency = "" Then
                slAgency = smSummaryRecords(ilRec)
            Else
                If slAgency <> smSummaryRecords(ilRec) Then
                    slAgency = "Agency: Varies Across Contracts"
                    Exit For
                End If
            End If
        End If
    Next ilRec
    For ilRec = 0 To UBound(smSummaryRecords) - 1 Step 1
        If InStr(1, smSummaryRecords(ilRec), "Contract Total") > 0 Or InStr(1, smSummaryRecords(ilRec), "Contract Bonus Total") Then
            ilAdvtTotalRec = ilRec
        End If
    Next ilRec

    If imNoCntr > 1 Then
        blPurchPopVaries = False
        blReRatePopVaries = False
        ilPurchPopRec = -1
        ilReRatePopRec = -1
        For ilRec = 0 To UBound(smSummaryRecords) - 1 Step 1
            If InStr(1, smSummaryRecords(ilRec), "Contract Population:") > 0 Then
                If ilPurchPopRec = -1 Then
                    ilPurchPopRec = ilRec
                Else
                    If smSummaryRecords(ilRec) <> smSummaryRecords(ilPurchPopRec) Then
                        'JW 10/14/21 - Fix v81 TTP 10258 issues per Jason Email: Thu 10/14/21 10:13 AM (#2)
                        smSummaryRecords(ilPurchPopRec) = "Contract Population: " & "Population varies Across Contracts"
                    End If
                End If
            End If
            If InStr(1, smSummaryRecords(ilRec), "ReRate Population:") > 0 Then
                If ilReRatePopRec = -1 Then
                    ilReRatePopRec = ilRec
                    ilRet = gParseItem(smSummaryRecords(ilRec), 2, ":", slPopulation)
                    If InStr(1, smSummaryRecords(ilRec), "Varies") > 0 Then
                        lmContractPopulation = 0
                    Else
                        lmContractPopulation = Val(slPopulation)
                    End If
                Else
                    If smSummaryRecords(ilRec) <> smSummaryRecords(ilReRatePopRec) Then
                        'JW 10/14/21 - Fix v81 TTP 10258 issues per Jason Email: Thu 10/14/21 10:13 AM (#2)
                        smSummaryRecords(ilReRatePopRec) = "ReRate Population: " & "Population varies Across Contracts"
                        lmContractPopulation = 0
                    End If
                End If
            End If
        Next ilRec
    End If

    blNextRowAdvt = False
    For ilRec = 0 To UBound(smSummaryRecords) - 1 Step 1
        If (ilRec >= ilAdvtTotalRec) Then
            blNextRowAdvt = True
        End If
        If InStr(1, smSummaryRecords(ilRec), "Product:") > 0 Then
            ilRet = gParseItem(smSummaryRecords(ilRec), 2, ":", slProduct)
            slProduct = Left(slProduct, Len(slProduct))
        ElseIf InStr(1, smSummaryRecords(ilRec), "Order#:") > 0 Then
            ilRet = gParseItem(smSummaryRecords(ilRec), 2, ":", slOrder)
            slOrder = Left(slOrder, Len(slOrder))
        Else
            If InStr(1, smSummaryRecords(ilRec), "Agency:") > 0 Then
                smSummaryRecords(ilRec) = slAgency
            End If
            ilRet = gParseItem(smSummaryRecords(ilRec), 2, smDelimiter, slRecord)
            slRecord = smSummaryRecords(ilRec)
            'ilRet = gParseItem(smSummaryRecords(ilRec), 2, "|", slDelimiter)
            If InStr(1, slRecord, "Line#") Then
                slRecord = Replace(slRecord, "Line#", "")
                slRecord = Replace(slRecord, "Vehicle", "Order")
                slRecord = Replace(slRecord, "Daypart", "Product")
                slRecord = Replace(slRecord, "Lineup #", "")
                slRecord = Replace(slRecord, "Audio Type", "")
                slRecord = Replace(slRecord, "Len", "")
                slRecord = Replace(slRecord, "Price Type", "")
                slRecord = Replace(slRecord, "Rate", "")
                slRecord = Replace(slRecord, "Line Comment", "")
                slRecord = Replace(slRecord, "Book", "")
            End If
            If (InStr(1, slRecord, ":") > 0) And (Left(slRecord, 1) <> smDelimiter) Then
                If Not blColumnOneShow Then
                    mPrint slRecord, smDelimiter, ""
                    If ExptReRate.ckcCsv.Value = vbUnchecked Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1, 1)
                    If (InStr(1, slRecord, "Period:") > 0) Then
                        blColumnOneShow = True
                    End If
                End If
            ElseIf InStr(1, slRecord, smDelimiter & "Purch") > 0 And InStr(1, slRecord, smDelimiter & "ReRate") > 0 Then
                If Not blTitle1Shown Then
                    blTitle1Shown = True
                    mPrint slRecord, smDelimiter, ""
                    If ExptReRate.ckcCsv.Value = vbUnchecked Then
                        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1)
                        ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , vbBlue, imExcelRow - 1)
                    End If
                End If
            ElseIf InStr(1, smSummaryRecords(ilRec), smDelimiter & "Line") > 0 And InStr(1, smSummaryRecords(ilRec), smDelimiter & "Vehicle") > 0 Then
                If Not blTitle2Shown Then
                    blTitle2Shown = True
                    mPrint slRecord, smDelimiter, ""
                    If ExptReRate.ckcCsv.Value = vbUnchecked Then
                        ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1)
                        ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , vbBlue, imExcelRow - 1)
                    End If
                End If
            Else
                ilRet = gParseItem(smSummaryRecords(ilRec), 2, smDelimiter, slBonus)
                If slBonus <> "Bonus" Then
                    If (InStr(1, slRecord, "Advertiser Total") > 0) Or (InStr(1, slRecord, "Advertiser Bonus Total") > 0) Then
                        slRecord = Mid(slRecord, 2)
                        slRecord = Replace(slRecord, smDelimiter & smDelimiter, smDelimiter & smDelimiter & smDelimiter, 1, 1)
                        slRowType = "AT"
                        If InStr(1, slRecord, "Advertiser Bonus Total") > 0 Then
                            slRowType = "AB"
                        ElseIf InStr(1, slRecord, "Advertiser Total with Bonus") > 0 Then
                            slRowType = "AS"
                        End If
                    ElseIf InStr(1, slRecord, "Contract Total") > 0 Or InStr(1, slRecord, "Bonus Total") > 0 Then
                        'column 4 replace with Order and column 5 with Product
                        ilPos = 1
                        ilFieldNo = 0
                        Do
                            slStr = Mid(slRecord, ilPos, 1)
                            If slStr = smDelimiter Then
                                ilFieldNo = ilFieldNo + 1
                                If ilFieldNo = 3 Then
                                    'Insert into field 4 and 5
                                    slRecord = Left(slRecord, ilPos) & slOrder & smDelimiter & slProduct & smDelimiter & Mid(slRecord, ilPos + 2)
                                    slRecord = Mid(slRecord, 2)
                                    Exit Do
                                End If
                            End If
                            ilPos = ilPos + 1
                        Loop While ilPos < Len(slRecord)
                        slRowType = "CT"
                        If InStr(1, slRecord, "Contract Bonus Total") > 0 Then
                            slRowType = "CB"
                        ElseIf InStr(1, slRecord, "Contract Total with Bonus") > 0 Then
                            slRowType = "CS"
                        End If
                    End If
                    If slRowType = "CT" Or slRowType = "AT" Then
                        If blAddSpaceLine Then
                            If rbcLayout(0).Value = True Then mPrint ""
                        Else
                            blAddSpaceLine = True
                        End If
                    End If
                    ilRet = gParseItem(smSummaryRecords(ilRec), REXTTOTALEXCEL, smDelimiter, slExtTotal)
                    mPrint slRecord, smDelimiter, slRowType, lmContractPopulation, gStrDecToLong(slExtTotal, 2)
                    If ExptReRate.ckcCsv.Value = vbUnchecked Then ilRet = gExcelOutputGeneration("FB", omBook, omSheet, , "True", imExcelRow - 1)
                    If InStr(1, slRecord, "Contract Total") > 0 Or InStr(1, slRecord, "Contract Bonus") > 0 Then
                        If rbcColumnLayout(1).Value Then
                            blPRPRColorSwitch = True
                            For ilCell = imPurchasedColumn To omSheet.UsedRange.Columns.Count - 1 Step 2
                                If blPRPRColorSwitch Then
                                    blPRPRColorSwitch = False
                                    If ExptReRate.ckcCsv.Value = vbUnchecked Then
                                        ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow - 1, ilCell)
                                        ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow - 1, ilCell + 1)
                                        If Not blNextRowAdvt Then
                                            ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow, ilCell)
                                            ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow, ilCell + 1)
                                        End If
                                    End If
                                Else
                                    blPRPRColorSwitch = True
                                    If ExptReRate.ckcCsv.Value = vbUnchecked Then
                                        ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow - 1, ilCell)
                                        ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow - 1, ilCell + 1)
                                        If Not blNextRowAdvt Then
                                            ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow, ilCell)
                                            ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow, ilCell + 1)
                                        End If
                                    End If
                                End If
                            Next ilCell
                        Else
                            If ExptReRate.ckcCsv.Value = vbUnchecked Then
                                For ilCol = imPurchasedColumn To imReRateColumn - 1 Step 1
                                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow - 1, ilCol)
                                    If Not blNextRowAdvt Then
                                        ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PURCHASECOLOR), imExcelRow, ilCol)
                                    End If
                                Next ilCol
                                For ilCol = imReRateColumn To omSheet.UsedRange.Columns.Count - 1 Step 1
                                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow - 1, ilCol)
                                    If Not blNextRowAdvt Then
                                        ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(RERATECOLOR), imExcelRow, ilCol)
                                    End If
                                Next ilCol
                            End If
                        End If
                    End If
                    If InStr(1, slRecord, "Contract Total") > 0 Or InStr(1, slRecord, "Bonus Total") > 0 Or InStr(1, slRecord, "Advertiser Total") > 0 Then
                        ilRet = gParseItem(smSummaryRecords(ilRec), 27, smDelimiter, slIndex)
                        If Val(slIndex) < 100 Then
                            If ExptReRate.ckcCsv.Value = vbUnchecked Then ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , str(Red), imExcelRow - 1, omSheet.UsedRange.Columns.Count)
                        Else
                            If ExptReRate.ckcCsv.Value = vbUnchecked Then ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , str(DARKGREEN), imExcelRow - 1, omSheet.UsedRange.Columns.Count)
                        End If
                    End If
                End If
            End If
        End If
    Next ilRec
End Sub

Private Sub mPopRevisions(slCntrNo As String)
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    Dim blFirstNo As Boolean
    
    cbcRevision.Clear
    blFirstNo = True
    slSQLQuery = "Select Distinct chfCntRevNo, chfExtRevNo, chfCode from chf_Contract_Header where ((chfSchStatus = 'F') or (chfSchStatus = 'M')) And chfCntrNo = " & slCntrNo
    slSQLQuery = slSQLQuery & " Order By chfCntRevNo"
    Set tmp_rst = gSQLSelectCall(slSQLQuery)
    Do While Not tmp_rst.EOF
        If blFirstNo Then
            cbcRevision.AddItem ("[Original]")
            blFirstNo = False
        Else
            cbcRevision.AddItem ("R" & tmp_rst!chfCntRevNo & "-" & tmp_rst!chfExtRevNo)
        End If
        cbcRevision.SetItemData = tmp_rst!chfCode
        tmp_rst.MoveNext
    Loop
End Sub

Private Sub mEnableBox()
'
'   mInitParameters ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    
    If (grdCntr.Row >= grdCntr.FixedRows) And (grdCntr.Row < grdCntr.Rows) And (grdCntr.Col >= VERSIONINDEX) And (grdCntr.Col < grdCntr.cols - 1) Then
        lmEnableRow = grdCntr.Row
        lmEnableCol = grdCntr.Col
        'imShowGridBox = True

        Select Case grdCntr.Col
            Case VERSIONINDEX
                mPopRevisions grdCntr.TextMatrix(lmEnableRow, CNTRNOINDEX)
                cbcRevision.Move grdCntr.Left + grdCntr.ColPos(grdCntr.Col) + 30, grdCntr.Top + grdCntr.RowPos(grdCntr.Row) + 15, grdCntr.ColWidth(grdCntr.Col) - 30, grdCntr.RowHeight(grdCntr.Row) - 15
                slStr = grdCntr.TextMatrix(lmEnableRow, lmEnableCol)
                cbcRevision.ZOrder vbBringToFront
                cbcRevision.Visible = True  'Set visibility
                cbcRevision.SelText (Trim(slStr))
                cbcRevision.SetFocus
        End Select
    End If
End Sub

Private Sub mSetShow()
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (lmEnableRow >= grdCntr.FixedRows) And (lmEnableRow < grdCntr.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case VERSIONINDEX
                cbcRevision.Visible = False  'Set visibility
                slStr = cbcRevision.Text
                If slStr = "" Then
                    If cbcRevision.ListCount > 0 Then
                        cbcRevision.Text = cbcRevision.GetName(0)
                        slStr = cbcRevision.Text
                    End If
                End If
                If slStr = "[Original]" Then
                    grdCntr.TextMatrix(lmEnableRow, lmEnableCol) = "Original"
                Else
                    grdCntr.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
                If cbcRevision.ListIndex >= 0 Then
                    grdCntr.TextMatrix(lmEnableRow, PURCHASECHFCODEINDEX) = cbcRevision.GetItemData(cbcRevision.ListIndex)
                Else
                    grdCntr.TextMatrix(lmEnableRow, PURCHASECHFCODEINDEX) = 0
                End If
        End Select
        lmEnableCol = -1
        lmEnableRow = -1
        bmInGrid = False
    End If
End Sub

Private Sub mHandleHiddenLinesMoved()
    Dim ilClfP As Integer
    Dim ilClfR As Integer
    Dim blFound As Boolean
    Dim blRemoved As Boolean
    For ilClfR = 0 To UBound(tmClfReRate) - 1 Step 1
        blFound = False
        blRemoved = False
        For ilClfP = 0 To UBound(tmClfPurchase) - 1 Step 1
            If tmClfReRate(ilClfR).ClfRec.iLine = tmClfPurchase(ilClfP).ClfRec.iLine Then
                blFound = True
                If tmClfReRate(ilClfR).ClfRec.iPkLineNo <> tmClfPurchase(ilClfP).ClfRec.iPkLineNo Then
                    If tmClfReRate(ilClfR).ClfRec.iPkLineNo = 0 Then
                        'Moved out of the package
                        mCopyLines tmClfPurchase(ilClfP), tmClfReRate(), tmCffPurchase(), tmCffReRate(), tmClfPurchase(ilClfP).ClfRec.iPkLineNo
                        mCopyLines tmClfReRate(ilClfR), tmClfPurchase(), tmCffReRate(), tmCffPurchase(), 0
                        '11/13/2020 - TTP 9993 - ReRate Gimps and Grps lines twice, when unpackage or package on a later revision. negate package Line # of lines that were moved out of package
                        tmClfPurchase(ilClfR).ClfRec.iLine = -tmClfPurchase(ilClfR).ClfRec.iLine
                        tmClfReRate(UBound(tmClfReRate) - 1).ClfRec.iLine = -tmClfReRate(UBound(tmClfReRate) - 1).ClfRec.iLine
                    ElseIf tmClfPurchase(ilClfP).ClfRec.iPkLineNo = 0 Then
                        'Moved into the package
                        mCopyLines tmClfReRate(ilClfR), tmClfPurchase(), tmCffReRate(), tmCffPurchase(), tmClfReRate(ilClfR).ClfRec.iPkLineNo
                        mCopyLines tmClfPurchase(ilClfP), tmClfReRate(), tmCffPurchase(), tmCffReRate(), 0
                        '11/13/2020 - TTP 9993 - ReRate Gimps and Grps lines twice, when unpackage or package on a later revision.  negate Line # of standard lines that were moved in to package
                        tmClfPurchase(ilClfP).ClfRec.iLine = -tmClfPurchase(ilClfP).ClfRec.iLine
                        tmClfReRate(UBound(tmClfReRate) - 1).ClfRec.iLine = -tmClfReRate(UBound(tmClfReRate) - 1).ClfRec.iLine
                    Else
                       'Moved to different package
                       mCopyLines tmClfPurchase(ilClfP), tmClfReRate(), tmCffPurchase(), tmCffReRate(), tmClfPurchase(ilClfP).ClfRec.iPkLineNo
                       mCopyLines tmClfReRate(ilClfR), tmClfPurchase(), tmCffReRate(), tmCffPurchase(), tmClfReRate(ilClfR).ClfRec.iPkLineNo
                    End If
                End If
                Exit For
            End If
        Next ilClfP
    Next ilClfR
End Sub

Private Sub mCopyLines(tlFromClf As CLFLIST, tlToClf() As CLFLIST, tlFromCff() As CFFLIST, tlToCff() As CFFLIST, ilPkLineNo As Integer)
    Dim ilClfUpper As Integer
    Dim ilNext As Integer
    Dim ilCffUpper As Integer
    Dim ilPrevCffUpper As Integer
    
    ilClfUpper = UBound(tlToClf)
    tlToClf(ilClfUpper) = tlFromClf
    tlToClf(ilClfUpper).ClfRec.iPkLineNo = ilPkLineNo
    tlToClf(ilClfUpper).ClfRec.iDnfCode = 0
    tlToClf(ilClfUpper).iFirstCff = -1
    ilNext = tlFromClf.iFirstCff
    Do While ilNext <> -1
        ilCffUpper = UBound(tlToCff)
        tlToCff(ilCffUpper) = tlFromCff(ilNext)
        tlToCff(ilCffUpper).CffRec.sDyWk = "W"
        tlToCff(ilCffUpper).CffRec.iSpotsWk = 0
        tlToCff(ilCffUpper).CffRec.lPropPrice = 0
        tlToCff(ilCffUpper).CffRec.lActPrice = 0
        If tlToClf(ilClfUpper).iFirstCff = -1 Then
            tlToClf(ilClfUpper).iFirstCff = ilCffUpper
        Else
            tlToCff(ilPrevCffUpper).iNextCff = ilCffUpper
        End If
        ilPrevCffUpper = ilCffUpper
        ReDim Preserve tlToCff(0 To ilCffUpper + 1) As CFFLIST
        ilNext = tlFromCff(ilNext).iNextCff
    Loop
    ReDim Preserve tlToClf(0 To ilClfUpper + 1) As CLFLIST

End Sub

Private Sub mSetExcelColumns()
    LINEEXCEL = 2
    VEHICLEEXCEL = LINEEXCEL + 1 '3
    DAYPARTEXCEL = VEHICLEEXCEL + 1 '4
    LINEUPEXCEL = DAYPARTEXCEL + 1 '5
    AUDIOTYPEEXCEL = LINEUPEXCEL + 1 '6
    LENEXCEL = AUDIOTYPEEXCEL + 1 '7
    PRICETYPEEXCEL = LENEXCEL + 1 '8
    RATEEXCEL = PRICETYPEEXCEL + 1 '9
    LINECOMMENTEXCEL = RATEEXCEL + 1 '10
    LASTSTATICCOLEXCEL = LINECOMMENTEXCEL 'Number of columns before "Purchased Total" column
    INDEXEXCEL = 29 'Last column is the GIMP/GRP
    
    If rbcColumnLayout(0).Value Then
        'PPP RRR
        PEXTTOTALEXCEL = LASTSTATICCOLEXCEL + 1 '9 -> 11
        PUNITEXCEL = PEXTTOTALEXCEL + 1 '10 -> 12
        PAQHEXCEL = PUNITEXCEL + 1 '11 -> 13
        PRTGEXCEL = PAQHEXCEL + 1 '12 -> 14
        PCPMEXCEL = PRTGEXCEL + 1 '13 -> 15
        PCPPEXCEL = PCPMEXCEL + 1 '14 -> 16
        PGIMPEXCEL = PCPPEXCEL + 1 '15 -> 17
        PGRPEXCEL = PGIMPEXCEL + 1 '16 -> 18
        PBOOKEXCEL = PGRPEXCEL + 1 '17 -> 19
        REXTTOTALEXCEL = PBOOKEXCEL + 1 '18 -> 20
        RUNITEXCEL = REXTTOTALEXCEL + 1 '19 -> 21
        RAQHEXCEL = RUNITEXCEL + 1 '20 -> 22
        RRTGEXCEL = RAQHEXCEL + 1 '21 -> 23
        RCPMEXCEL = RRTGEXCEL + 1 '22 -> 24
        RCPPEXCEL = RCPMEXCEL + 1 '23 -> 25
        RGIMPEXCEL = RCPPEXCEL + 1 ' 24 -> 26
        RGRPEXCEL = RGIMPEXCEL + 1 '25 -> 27
        RBOOKEXCEL = RGRPEXCEL + 1 '26 -> 28
    Else
        'PR,PR,PR
        PEXTTOTALEXCEL = LASTSTATICCOLEXCEL + 1 '9
        PUNITEXCEL = PEXTTOTALEXCEL + 2 '11
        PAQHEXCEL = PUNITEXCEL + 2 '13
        PRTGEXCEL = PAQHEXCEL + 2 '15
        PCPMEXCEL = PRTGEXCEL + 2 '17
        PCPPEXCEL = PCPMEXCEL + 2 '19
        PGIMPEXCEL = PCPPEXCEL + 2 '21
        PGRPEXCEL = PGIMPEXCEL + 2 '23
        PBOOKEXCEL = PGRPEXCEL + 2 '25
        
        REXTTOTALEXCEL = PEXTTOTALEXCEL + 1 '10
        RUNITEXCEL = REXTTOTALEXCEL + 2 '12
        RAQHEXCEL = RUNITEXCEL + 2 '14
        RRTGEXCEL = RAQHEXCEL + 2 '16
        RCPMEXCEL = RRTGEXCEL + 2 '18
        RCPPEXCEL = RCPMEXCEL + 2 '20
        RGIMPEXCEL = RCPPEXCEL + 2 '22
        RGRPEXCEL = RGIMPEXCEL + 2 '24
        RBOOKEXCEL = RGRPEXCEL + 2 '26
    End If
    'TTP 10082 - merge header into columns (Shift everything over by 10 columns because 1st 10 will be contract details)
    If rbcLayout(1).Value = True Then
        'we have 10 contract summary columns before the other columns, so shift everything over by 10
        LINEEXCEL = LINEEXCEL + 10
        VEHICLEEXCEL = VEHICLEEXCEL + 10
        DAYPARTEXCEL = DAYPARTEXCEL + 10
        AUDIOTYPEEXCEL = AUDIOTYPEEXCEL + 10
        LINEUPEXCEL = LINEUPEXCEL + 10
        LENEXCEL = LENEXCEL + 10
        PRICETYPEEXCEL = PRICETYPEEXCEL + 10
        RATEEXCEL = RATEEXCEL + 10
        LINECOMMENTEXCEL = LINECOMMENTEXCEL + 10
        PEXTTOTALEXCEL = PEXTTOTALEXCEL + 10
        PUNITEXCEL = PUNITEXCEL + 10
        PAQHEXCEL = PAQHEXCEL + 10
        PRTGEXCEL = PRTGEXCEL + 10
        PCPMEXCEL = PCPMEXCEL + 10
        PCPPEXCEL = PCPPEXCEL + 10
        PGIMPEXCEL = PGIMPEXCEL + 10
        PGRPEXCEL = PGRPEXCEL + 10
        PBOOKEXCEL = PBOOKEXCEL + 10
        REXTTOTALEXCEL = REXTTOTALEXCEL + 10
        RUNITEXCEL = RUNITEXCEL + 10
        RAQHEXCEL = RAQHEXCEL + 10
        RRTGEXCEL = RRTGEXCEL + 10
        RCPMEXCEL = RCPMEXCEL + 10
        RCPPEXCEL = RCPPEXCEL + 10
        RGIMPEXCEL = RGIMPEXCEL + 10
        RGRPEXCEL = RGRPEXCEL + 10
        RBOOKEXCEL = RBOOKEXCEL + 10
        INDEXEXCEL = INDEXEXCEL + 10
    End If
End Sub

Private Sub mExcelFormulaSetting(llFormulaSpots As Long, llFormulaPop As Long, llFormulaExtTotal As Long)
    Dim ilRet As Integer
    
    'Add formula in cells not set
    If llFormulaSpots > 0 Then
        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & smColumnLetter(RUNITEXCEL) & imExcelRow - 1 & "*" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "," & """" & """" & ")", imExcelRow - 1, RGIMPEXCEL) ', slDelimiter)
        If rbcIndex(0).Value Then
            'Index by GImp
            ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "IF(" & smColumnLetter(PAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "100*" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "/" & smColumnLetter(PGIMPEXCEL) & imExcelRow - 1 & "," & """" & """" & ")" & "," & """" & """" & ")", imExcelRow - 1, INDEXEXCEL) ', slDelimiter)
        End If
        If llFormulaPop > 0 Then
            ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "*" & "100" & "/" & llFormulaPop & "," & """" & """" & ")", imExcelRow - 1, RGRPEXCEL) ', slDelimiter)
            If rbcIndex(1).Value Then
                'Index by GRG
                ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "IF(" & smColumnLetter(PAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & "100*" & smColumnLetter(RGRPEXCEL) & imExcelRow - 1 & "/" & smColumnLetter(PGRPEXCEL) & imExcelRow - 1 & "," & """" & """" & ")" & "," & """" & """" & ")", imExcelRow - 1, INDEXEXCEL) ', slDelimiter)
            End If
            ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "*" & "100" & "/" & llFormulaPop & "," & """" & """" & ")", imExcelRow - 1, RRTGEXCEL) ', slDelimiter)
            If llFormulaExtTotal >= 0 And ckcCPM.Value = vbChecked Then
                ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & llFormulaExtTotal & "*" & llFormulaPop & "/" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "/" & "10000" & "," & """" & """" & ")", imExcelRow - 1, RCPPEXCEL) ', slDelimiter)
                ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , "=" & "IF(" & smColumnLetter(RAQHEXCEL) & imExcelRow - 1 & "> 0" & "," & llFormulaExtTotal & "/" & smColumnLetter(RGIMPEXCEL) & imExcelRow - 1 & "/" & "100" & "," & """" & """" & ")", imExcelRow - 1, RCPMEXCEL) ', slDelimiter)
            End If
        End If
    End If

End Sub

Sub mSetCellRule()
'search: vb6 excel formatconditions
'Example 1
'    'https://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/
'    Dim rg As Range
'    Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
'    Set rg = Range("A2", Range("A2").End(xlDown))
'
'    'clear any existing conditional formatting
'    rg.FormatConditions.Delete
'
'    'define the rule for each conditional format
'    Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=$a$1")
'    Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=$a$1")
'    Set cond3 = rg.FormatConditions.Add(xlCellValue, xlEqual, "=$a$1")
'
'    'define the format applied for each conditional format
'    With cond1
'    .Interior.Color = vbGreen
'    .Font.Color = vbWhite
'    End With
'
'    With cond2
'    .Interior.Color = vbRed
'    .Font.Color = vbWhite
'    End With
'
'    With cond3
'    .Interior.Color = vbYellow
'    .Font.Color = vbRed
'    End With
 
'Example 2
'    Dim rng As Range
'    Dim i As Integer
'    Set rng = Range("A1:C4")
'
'    ' Clear all existing formats
'
'    For i = rng.FormatConditions.Count To 1 Step -1
'
'    rng.FormatConditions(i).Delete
'    Next
'
'    With rng
'    .FormatConditions.Add xlCellValue, xlBetween, 0, 10
'    .FormatConditions(1).Interior.Color = RGB(196, 196, 196)
'    .FormatConditions.Add xlCellValue, xlNotBetween, 0, 10
'    .FormatConditions(2).Interior.Color = RGB(255, 255, 255)
'    End With

'Name    Value   Description
'xlBetween   1   Between. Can be used only if two formulas are provided.
'xlEqual 3   Equal.
'xlGreater   5   Greater than.
'xlGreaterEqual  7   Greater than or equal to.
'xlLess  6   Less than.
'xlLessEqual 8   Less than or equal to.
'xlNotBetween    2   Not between. Can be used only if two formulas are provided.
'xlNotEqual  4   Not equal.


    Dim rng As Range
    Dim i As Integer
    Dim slRange As String
    
    If rbcReRateBook(2).Value = False Then
        Exit Sub
    End If
    slRange = smColumnLetter(INDEXEXCEL) & ":" & smColumnLetter(INDEXEXCEL)
    Set rng = omSheet.Range(slRange)
    ' Clear all existing formats
    For i = rng.FormatConditions.Count To 1 Step -1

        rng.FormatConditions(i).Delete
    Next

    With rng
    .FormatConditions.Add xlCellValue, xlEqual, "Index"
    .FormatConditions(1).Font.Color = BLUE
    If rbcIndex(1).Value Then
        .FormatConditions.Add xlCellValue, xlEqual, "GRP"
    Else
        .FormatConditions.Add xlCellValue, xlEqual, "Gimp"
    End If
    .FormatConditions(2).Font.Color = BLUE
    .FormatConditions.Add xlCellValue, xlGreaterEqual, 100
    '.FormatConditions.Add xlExpression, , "IF(((Z)<>" & """" & "Index" & """" & ") AND (Z)) >= 100,True, False)"
    .FormatConditions(3).Font.Color = DARKGREEN
    .FormatConditions.Add xlCellValue, xlLess, 100
    '.FormatConditions.Add xlExpression, , "ISNUMERIC(Z1:Z100) AND (Z1:Z100) < 100"
    .FormatConditions(4).Font.Color = Red
    End With
End Sub

Private Sub mDefineAlignColumns()
    imRightAlignColumn(0) = LINEEXCEL
    imRightAlignColumn(1) = LENEXCEL
    imRightAlignColumn(2) = RATEEXCEL
    imRightAlignColumn(3) = PEXTTOTALEXCEL
    imRightAlignColumn(4) = PAQHEXCEL
    imRightAlignColumn(5) = PRTGEXCEL
    imRightAlignColumn(6) = PCPMEXCEL
    imRightAlignColumn(7) = PCPPEXCEL
    imRightAlignColumn(8) = PGIMPEXCEL
    imRightAlignColumn(9) = PGRPEXCEL
    imRightAlignColumn(10) = REXTTOTALEXCEL
    imRightAlignColumn(11) = RUNITEXCEL
    imRightAlignColumn(12) = RAQHEXCEL
    imRightAlignColumn(13) = RRTGEXCEL
    imRightAlignColumn(14) = RCPMEXCEL
    imRightAlignColumn(15) = RCPPEXCEL
    imRightAlignColumn(16) = RGIMPEXCEL
    imRightAlignColumn(17) = RGRPEXCEL
    imRightAlignColumn(18) = INDEXEXCEL
    imRightAlignColumn(19) = PUNITEXCEL
End Sub

Private Function mFormula_GrImpFromAQH(ilRow As Integer) As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Units * AQH
    mFormula_GrImpFromAQH = "=" & "IF(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RUNITEXCEL) & ilRow & "*" & smColumnLetter(RAQHEXCEL) & ilRow & "," & """" & """" & ")"
End Function

Private Function mFormula_GrImpFromSumForPL(ilStartRow As Integer, ilEndRow As Integer) As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Sum of GrImp cells
    Dim ilRow As Integer
    Dim slGrImp As String
    Dim slDelimiter As String
    slGrImp = ""
    slDelimiter = ""
    For ilRow = ilStartRow To ilEndRow Step 1
        If tmFormulaInfo(ilRow).sRowType = "BL" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        End If
    Next ilRow
    'mFormula_GrImpFromSum = "=" & "Sum(" & slGrImp & ")"
    mFormula_GrImpFromSumForPL = "=" & "IF(" & "Sum(" & slGrImp & ")" & ">0" & "," & "Sum(" & slGrImp & ")" & "," & """" & """" & ")"
End Function
Private Function mFormula_GrImp_Index(ilRow As Integer) As String
    'ReRate GrImp / Purchase GrImp
    'mFormula_GrImp_Index = "=" & "IF(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & "IF(" & smColumnLetter(PAQHEXCEL) & ilRow & "> 0" & "," & "100*" & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & smColumnLetter(PGIMPEXCEL) & ilRow & "," & """" & """" & ")" & "," & """" & """" & ")"
    'To avoid Excel shoulding the cell as in error, include the cell that GrImp is referencing (AQH)
    'mFormula_GrImp_Index = "=" & "IF(" & "AND(" & smColumnLetter(PGIMPEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "<> " & """" & """" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & ")" & "," & "100*" & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & smColumnLetter(PGIMPEXCEL) & ilRow & "," & """" & """" & ")"
    mFormula_GrImp_Index = "=" & "IF(" & "AND(" & smColumnLetter(PGIMPEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "<> " & """" & """" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & ")" & "," & "100*" & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & smColumnLetter(PGIMPEXCEL) & ilRow & "," & """" & """" & ")"
End Function
Private Function mFormula_GRPFromGrImp(ilRow As Integer, llPop As Long) As String
    'Formula form: =Sum(Cell:Cell)     Sums the range
    'Formula form: =Sum(Cell, Cell)    Sums each cell
    'GrImp / Population
    ''mFormula_GRPFromGrImp = "=" & "IF(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "*" & "100" & "/" & llPop & "," & """" & """" & ")"
    'mFormula_GRPFromGrImp = "=" & "IF(" & smColumnLetter(RGIMPEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "*" & "100" & "/" & llPop & "," & """" & """" & ")"
    'To avoid Excel shoulding the cell as in error, include the cell that GrImp is referencing (AQH)
    mFormula_GRPFromGrImp = "=" & "IF(" & "AND(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "<> " & """" & """" & "," & llPop & "> 0" & ")" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "*" & "100" & "*" & imAdjDecPlaces & "/" & llPop & "," & """" & """" & ")"
End Function
Private Function mFormula_GRP_Index(ilRow As Integer) As String
    'ReRate GRP / Purchase GRP
    'mFormula_GRP_Index = "=" & "IF(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & "IF(" & smColumnLetter(PAQHEXCEL) & ilRow & "> 0" & "," & "100*" & smColumnLetter(RGRPEXCEL) & ilRow & "/" & smColumnLetter(PGRPEXCEL) & ilRow & "," & """" & """" & ")" & "," & """" & """" & ")"
    mFormula_GRP_Index = "=" & "IF(" & "AND(" & smColumnLetter(PGRPEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "<> " & """" & """" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & ")" & "," & "100*" & smColumnLetter(RGRPEXCEL) & ilRow & "/" & smColumnLetter(PGRPEXCEL) & ilRow & "," & """" & """" & ")"
End Function
'Private Function mFormula_CPPFromGrImp(ilRow As Integer, llExtTotal As Long, llPop As Long) As String
Private Function mFormula_CPPFromGrImp(ilRow As Integer, dlExtTotal As Double, llPop As Long) As String 'TTP 10439 - Rerate 21,000,000
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Cost * Population / GrImp
    'mFormula_CPPFromGrImp = "=" & "IF(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & llExtTotal & "*" & llPop & "/" & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & "10000" & "," & """" & """" & ")"
    'To avoid Excel shoulding the cell as in error, include the cell that GrImp is referencing (AQH)
    mFormula_CPPFromGrImp = "=" & "IF(" & "AND(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "<> " & """" & """" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "> 0" & ")" & "," & dlExtTotal & "*" & llPop & "/" & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & "10000" & "/" & imAdjDecPlaces & "," & """" & """" & ")"
End Function
'Private Function mFormula_CPMFromGrImp(ilRow As Integer, llExtTotal As Long) As String
Private Function mFormula_CPMFromGrImp(ilRow As Integer, dlExtTotal As Double) As String 'TTP 10439 - Rerate 21,000,000
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Cost * Population / GrImp
    'mFormula_CPMFromGrImp = "=" & "IF(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & llExtTotal & "/" & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & "100" & "," & """" & """" & ")"
    'To avoid Excel shoulding the cell as in error, include the cell that GrImp is referencing (AQH)
    mFormula_CPMFromGrImp = "=" & "IF(" & "AND(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "<> " & """" & """" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "> 0" & ")" & "," & dlExtTotal & "/" & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & "100" & "," & """" & """" & ")"
End Function
Private Function mFormula_RatingFromAQH(ilRow As Integer, llPop As Long) As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'AQH / Population
    'mFormula_RatingFromAQH = "=" & "IF(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "*" & "100" & "/" & llPop & "," & """" & """" & ")"
    'To avoid Excel shoulding the cell as in error, include the cell that GrImp is referencing (AQH)
    mFormula_RatingFromAQH = "=" & "IF(" & "AND(" & smColumnLetter(RAQHEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "<> " & """" & """" & "," & llPop & "> 0" & ")" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "*" & "100" & "*" & imAdjDecPlaces & "/" & llPop & "," & """" & """" & ")"
End Function
Private Function mFormula_AQHFromGrImp(ilRow As Integer) As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Units * AQH
    mFormula_AQHFromGrImp = "=" & "IF(" & "AND(" & smColumnLetter(RUNITEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "<> " & """" & """" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "> 0" & ")" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & smColumnLetter(RUNITEXCEL) & ilRow & "," & """" & """" & ")"
    'mFormula_AQHFromGrImp = "=" & "IF(" & "AND(" & smColumnLetter(RUNITEXCEL) & ilRow & "> 0" & "," & smColumnLetter(RAQHEXCEL) & ilRow & "<> " & """" & """" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "> 0" & ")" & "," & smColumnLetter(RGIMPEXCEL) & ilRow & "/" & smColumnLetter(RUNITEXCEL) & ilRow & "," & """" & """" & ")"
End Function
Private Function mFormula_GrImpFromSumForCT(ilStartRow As Integer, ilEndRow As Integer) As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Sum of GrImp cells
    Dim ilRow As Integer
    Dim slGrImp As String
    Dim slDelimiter As String
    slGrImp = ""
    slDelimiter = ""
    For ilRow = ilStartRow To ilEndRow Step 1
        If tmFormulaInfo(ilRow).sRowType = "SL" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        ElseIf tmFormulaInfo(ilRow).sRowType = "PL" Then
            If rbcShow(2).Value Then
                'If Package plus hidden, set package
            ElseIf rbcShow(1).Value Then
                'If Package only, set package
                slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
                slDelimiter = ","
            Else
                'Hidden line, show package but can't edit it
            End If
        ElseIf tmFormulaInfo(ilRow).sRowType = "HL" Then
            If rbcShow(2).Value Then
                'If Package plus hidden, set package
                slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
                slDelimiter = ","
            ElseIf rbcShow(1).Value Then
                'If Package only, set package
            Else
                'Hidden line, show package but can't edit it
                slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
                slDelimiter = ","
            End If
        End If
    Next ilRow
    'mFormula_GrImpFromSum = "=" & "Sum(" & slGrImp & ")"
    mFormula_GrImpFromSumForCT = "=" & "IF(" & "Sum(" & slGrImp & ")" & ">0" & "," & "Sum(" & slGrImp & ")" & "," & """" & """" & ")"
End Function
Private Function mFormula_GrImpFromSumForCB(ilStartRow As Integer, ilEndRow As Integer) As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Sum of GrImp cells
    Dim ilRow As Integer
    Dim slGrImp As String
    Dim slDelimiter As String
    slGrImp = ""
    slDelimiter = ""
    For ilRow = ilStartRow To ilEndRow Step 1
        If tmFormulaInfo(ilRow).sRowType = "BL" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        End If
    Next ilRow
    'mFormula_GrImpFromSum = "=" & "Sum(" & slGrImp & ")"
    mFormula_GrImpFromSumForCB = "=" & "IF(" & "Sum(" & slGrImp & ")" & ">0" & "," & "Sum(" & slGrImp & ")" & "," & """" & """" & ")"
End Function
Private Function mFormula_GrImpFromSumForCS(ilStartRow As Integer, ilEndRow As Integer) As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Sum of GrImp cells
    Dim ilRow As Integer
    Dim slGrImp As String
    Dim slDelimiter As String
    slGrImp = ""
    slDelimiter = ""
    For ilRow = ilStartRow To ilEndRow Step 1
        If tmFormulaInfo(ilRow).sRowType = "CT" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        ElseIf tmFormulaInfo(ilRow).sRowType = "CB" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        End If
    Next ilRow
    'mFormula_GrImpFromSum = "=" & "Sum(" & slGrImp & ")"
    mFormula_GrImpFromSumForCS = "=" & "IF(" & "Sum(" & slGrImp & ")" & ">0" & "," & "Sum(" & slGrImp & ")" & "," & """" & """" & ")"
End Function
Private Function mFormula_GrImpFromSumForAT() As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Sum of GrImp cells
    Dim ilRow As Integer
    Dim slGrImp As String
    Dim slDelimiter As String
    slGrImp = ""
    slDelimiter = ""
    For ilRow = 0 To UBound(tmFormulaInfo) - 1 Step 1
        If tmFormulaInfo(ilRow).sRowType = "CT" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        End If
    Next ilRow
    'mFormula_GrImpFromSum = "=" & "Sum(" & slGrImp & ")"
    mFormula_GrImpFromSumForAT = "=" & "IF(" & "Sum(" & slGrImp & ")" & ">0" & "," & "Sum(" & slGrImp & ")" & "," & """" & """" & ")"
End Function
Private Function mFormula_GrImpFromSumForAB() As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Sum of GrImp cells
    Dim ilRow As Integer
    Dim slGrImp As String
    Dim slDelimiter As String
    slGrImp = ""
    slDelimiter = ""
    For ilRow = 0 To UBound(tmFormulaInfo) - 1 Step 1
        If tmFormulaInfo(ilRow).sRowType = "CB" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        End If
    Next ilRow
    'mFormula_GrImpFromSum = "=" & "Sum(" & slGrImp & ")"
    mFormula_GrImpFromSumForAB = "=" & "IF(" & "Sum(" & slGrImp & ")" & ">0" & "," & "Sum(" & slGrImp & ")" & "," & """" & """" & ")"
End Function
Private Function mFormula_GrImpFromSumForAS() As String
    'Formula form: =IF(Test,TrueValue,FalseValue)
    'Sum of GrImp cells
    Dim ilRow As Integer
    Dim slGrImp As String
    Dim slDelimiter As String
    slGrImp = ""
    slDelimiter = ""
    For ilRow = 0 To UBound(tmFormulaInfo) - 1 Step 1
        If tmFormulaInfo(ilRow).sRowType = "CT" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        ElseIf tmFormulaInfo(ilRow).sRowType = "CB" Then
            slGrImp = slGrImp & slDelimiter & smColumnLetter(RGIMPEXCEL) & tmFormulaInfo(ilRow).iExcelRow
            slDelimiter = ","
        End If
    Next ilRow
    'mFormula_GrImpFromSum = "=" & "Sum(" & slGrImp & ")"
    mFormula_GrImpFromSumForAS = "=" & "IF(" & "Sum(" & slGrImp & ")" & ">0" & "," & "Sum(" & slGrImp & ")" & "," & """" & """" & ")"
End Function
Public Sub mSendFormulaToExcel()
    Dim ilRet As Integer
    Dim ilContractRow As Integer
    Dim ilPackageRow As Integer
    Dim ilStartHiddenRow As Integer
    Dim ilEndHiddenRow As Integer
    Dim ilRowOuter As Integer
    Dim ilRowInner As Integer
    Dim slFormula As String
    Dim blSetCells As Boolean
    Dim blCTFound As Boolean
    
    If rbcReRateBook(2).Value = False Then
        Exit Sub
    End If
    If ckcSummary.Value = vbUnchecked Then
        ilRowOuter = 0
        ilContractRow = ilRowOuter
        blCTFound = False
        Do While ilRowOuter < UBound(tmFormulaInfo)
            blSetCells = False
            If blCTFound Then
                blCTFound = False
                If tmFormulaInfo(ilRowOuter).sRowType <> "BL" And tmFormulaInfo(ilRowOuter).sRowType <> "CB" And tmFormulaInfo(ilRowOuter).sRowType <> "CS" Then
                    ilContractRow = ilRowOuter
                End If
            End If
            Select Case tmFormulaInfo(ilRowOuter).sRowType
                Case "SL"   'Standard Line
                    blSetCells = True
                    slFormula = mFormula_GrImpFromAQH(tmFormulaInfo(ilRowOuter).iExcelRow)
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                Case "PL"   'Package
                    ilPackageRow = ilRowOuter
                    'Scan and generate each hidden line
                    ilRowInner = ilRowOuter + 1
                    ilStartHiddenRow = -1
                    Do While ilRowInner < UBound(tmFormulaInfo)
                        If tmFormulaInfo(ilRowInner).sRowType <> "HL" Then
                            If ilStartHiddenRow = -1 Then
                                ilStartHiddenRow = 1
                                ilEndHiddenRow = 0
                            End If
                            Exit Do
                        End If
                        If ilStartHiddenRow = -1 Then
                            ilStartHiddenRow = ilRowInner   'tmFormulaInfo(ilRowInner).iExcelRow
                        End If
                        ilEndHiddenRow = ilRowInner 'tmFormulaInfo(ilRowInner).iExcelRow
                        slFormula = mFormula_GrImpFromAQH(tmFormulaInfo(ilRowInner).iExcelRow)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowInner).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                        slFormula = mFormula_GRPFromGrImp(tmFormulaInfo(ilRowInner).iExcelRow, tmFormulaInfo(ilRowInner).lPop)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowInner).iExcelRow, RGRPEXCEL) ', slDelimiter)
                        slFormula = mFormula_CPPFromGrImp(tmFormulaInfo(ilRowInner).iExcelRow, tmFormulaInfo(ilRowInner).dExtTotal, tmFormulaInfo(ilRowInner).lPop) 'TTP 10439 - Rerate 21,000,000
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowInner).iExcelRow, RCPPEXCEL)  ', slDelimiter)
                        slFormula = mFormula_CPMFromGrImp(tmFormulaInfo(ilRowInner).iExcelRow, tmFormulaInfo(ilRowInner).dExtTotal) 'TTP 10439 - Rerate 21,000,000
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowInner).iExcelRow, RCPMEXCEL) ', slDelimiter)
                        slFormula = mFormula_RatingFromAQH(tmFormulaInfo(ilRowInner).iExcelRow, tmFormulaInfo(ilRowInner).lPop)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowInner).iExcelRow, RRTGEXCEL) ', slDelimiter)
                        If rbcIndex(1).Value Then
                            slFormula = mFormula_GRP_Index(tmFormulaInfo(ilRowInner).iExcelRow)
                        Else
                            slFormula = mFormula_GrImp_Index(tmFormulaInfo(ilRowInner).iExcelRow)
                        End If
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowInner).iExcelRow, INDEXEXCEL) ', slDelimiter)
                        ilRowInner = ilRowInner + 1
                    Loop
                    If rbcShow(2).Value Then
                        'If Package plus hidden, set package
                        blSetCells = True
                        slFormula = mFormula_GrImpFromSumForPL(ilStartHiddenRow, ilEndHiddenRow)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                        slFormula = mFormula_AQHFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RAQHEXCEL) ', slDelimiter)
                    ElseIf rbcShow(1).Value Then
                        'If Package only, set package
                        blSetCells = True
                        slFormula = mFormula_GrImpFromAQH(tmFormulaInfo(ilRowOuter).iExcelRow)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                    Else
                        'Hidden, show package but no computations
                        blSetCells = False
                    End If
                Case "BL"   'Bonus
                    blSetCells = True
                    slFormula = mFormula_GrImpFromAQH(tmFormulaInfo(ilRowOuter).iExcelRow)
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                Case "CT"   'Contract Total
                    'If ilRowOuter < UBound(tmFormulaInfo) Then
                    '    If tmFormulaInfo(ilRowOuter + 1).sRowType = "BL" Then
                    '        blCTFound = False
                    '    Else
                    '        blCTFound = True
                    '    End If
                    'Else
                        blCTFound = True
                    'End If
                    blSetCells = True
                    If ckcSummary.Value = vbUnchecked Then
                        slFormula = mFormula_GrImpFromSumForCT(ilContractRow, ilRowOuter)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                        slFormula = mFormula_AQHFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RAQHEXCEL) ', slDelimiter)
                    End If
                Case "CB"   'Contract Bonus
                    'If ilRowOuter < UBound(tmFormulaInfo) Then
                    '    If tmFormulaInfo(ilRowOuter + 1).sRowType = "CS" Then
                    '        blCTFound = False
                    '    Else
                    '        blCTFound = True
                    '    End If
                    'Else
                        blCTFound = True
                    'End If
                    blSetCells = True
                    If ckcSummary.Value = vbUnchecked Then
                        slFormula = mFormula_GrImpFromSumForCB(ilContractRow, ilRowOuter)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                        slFormula = mFormula_AQHFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow)
                        ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RAQHEXCEL) ', slDelimiter)
                    End If
                Case "CS"   'Contract plus bonus
                    blCTFound = True
                    blSetCells = True
                    slFormula = mFormula_GrImpFromSumForCS(ilContractRow, ilRowOuter)
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                    slFormula = mFormula_AQHFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow)
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RAQHEXCEL) ', slDelimiter)
                Case "AT"   'Advertiser Total
                    blSetCells = True
                    slFormula = mFormula_GrImpFromSumForAT()
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                    slFormula = mFormula_AQHFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow)
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RAQHEXCEL) ', slDelimiter)
                Case "AB"   'Advertiser Bonus
                    blSetCells = True
                    slFormula = mFormula_GrImpFromSumForAB()
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                    slFormula = mFormula_AQHFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow)
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RAQHEXCEL) ', slDelimiter)
                Case "AS"   'Advertiser plus bonus
                    blSetCells = True
                    slFormula = mFormula_GrImpFromSumForAS()
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGIMPEXCEL) ', slDelimiter)
                    slFormula = mFormula_AQHFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow)
                    ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RAQHEXCEL) ', slDelimiter)
            End Select
            If blSetCells Then
                slFormula = mFormula_GRPFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow, tmFormulaInfo(ilRowOuter).lPop)
                ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RGRPEXCEL) ', slDelimiter)
                slFormula = mFormula_CPPFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow, tmFormulaInfo(ilRowOuter).dExtTotal, tmFormulaInfo(ilRowOuter).lPop) 'TTP 10439 - Rerate 21,000,000

                ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RCPPEXCEL)  ', slDelimiter)
                slFormula = mFormula_CPMFromGrImp(tmFormulaInfo(ilRowOuter).iExcelRow, tmFormulaInfo(ilRowOuter).dExtTotal) 'TTP 10439 - Rerate 21,000,000
                ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RCPMEXCEL) ', slDelimiter)
                slFormula = mFormula_RatingFromAQH(tmFormulaInfo(ilRowOuter).iExcelRow, tmFormulaInfo(ilRowOuter).lPop)
                ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, RRTGEXCEL) ', slDelimiter)
                If rbcIndex(1).Value Then
                    slFormula = mFormula_GRP_Index(tmFormulaInfo(ilRowOuter).iExcelRow)
                Else
                    slFormula = mFormula_GrImp_Index(tmFormulaInfo(ilRowOuter).iExcelRow)
                End If
                ilRet = gExcelOutputGeneration("FV", omBook, omSheet, , slFormula, tmFormulaInfo(ilRowOuter).iExcelRow, INDEXEXCEL) ', slDelimiter)
                If tmFormulaInfo(ilRowOuter).sRowType <> "PL" Then
                    ilRowOuter = ilRowOuter + 1
                Else
                    ilRowOuter = ilRowOuter + ilEndHiddenRow - ilStartHiddenRow + 2
                End If
            Else
                If tmFormulaInfo(ilRowOuter).sRowType <> "PL" Then
                    ilRowOuter = ilRowOuter + 1
                Else
                    ilRowOuter = ilRowOuter + ilEndHiddenRow - ilStartHiddenRow + 2
                End If
            End If
        Loop
    Else
        'Summary Only
    End If
End Sub

Private Sub mAddRemoveCntrByLine()
    Dim llRow As Long
    Dim blFound As Boolean
    Dim ilChf As Integer
    
    'Add/Remove contracts contracts
    'If UBound(tgBookByLineCntr) > LBound(tgBookByLineCntr) Then
        For llRow = grdCntr.FixedRows To grdCntr.Rows - 1 Step 1
            'TTP 10141 - switching contracts causes Book by line to not load
            blFound = False
            'If grdCntr.TextMatrix(llRow, PRODUCTINDEX) <> "" And grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
            If grdCntr.TextMatrix(llRow, CNTRNOINDEX) <> "" And grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                blFound = False
                For ilChf = 0 To UBound(tgBookByLineCntr) - 1 Step 1
                    If grdCntr.TextMatrix(llRow, CNTRNOINDEX) = tgBookByLineCntr(ilChf).lCntrNo Then
                        'TTP 10141
                        tgBookByLineCntr(ilChf).sSelected = "1"
                        blFound = True
                        Exit For
                    End If
                Next ilChf
                
                If Not blFound Then
                    ilChf = UBound(tgBookByLineCntr)
                    tgBookByLineCntr(ilChf).lChfCode = grdCntr.TextMatrix(llRow, RERATECHFCODEINDEX)
                    tgBookByLineCntr(ilChf).lCntrNo = grdCntr.TextMatrix(llRow, CNTRNOINDEX)
                    tgBookByLineCntr(ilChf).sSelected = "1"
                    tgBookByLineCntr(ilChf).iFirst = -1
                    ReDim Preserve tgBookByLineCntr(0 To ilChf + 1) As BOOKBYLINECNTR
                    blFound = True
                End If
            'ElseIf grdCntr.TextMatrix(llRow, PRODUCTINDEX) <> "" And grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "" Then
            ElseIf grdCntr.TextMatrix(llRow, CNTRNOINDEX) <> "" And grdCntr.TextMatrix(llRow, SELECTEDINDEX) = "" Then
                blFound = False
                For ilChf = 0 To UBound(tgBookByLineCntr) - 1 Step 1
                    If grdCntr.TextMatrix(llRow, CNTRNOINDEX) = tgBookByLineCntr(ilChf).lCntrNo Then
                        'TTP 10141
                        blFound = True
                        tgBookByLineCntr(ilChf).sSelected = ""
                        Exit For
                    End If
                Next ilChf

            End If
            'TTP 10141 - switching contracts causes Book by line to not load
            If blFound = False Then
                tgBookByLineCntr(ilChf).sSelected = ""
            End If
        Next llRow
    'End If

End Sub

Private Sub mGetAllowedLengths()
    Dim illoop As Integer
    ReDim igReRateAllowedLengths(0 To 0) As Integer
    For illoop = 0 To lbcSpotLens.ListCount - 1 Step 1
        If lbcSpotLens.Selected(illoop) Then
            igReRateAllowedLengths(UBound(igReRateAllowedLengths)) = lbcSpotLens.List(illoop)
            ReDim Preserve igReRateAllowedLengths(0 To UBound(igReRateAllowedLengths) + 1) As Integer
        End If
    Next illoop
End Sub

Public Function mGetDnfByContractLine(ilBook As Integer, llSdf As Long) As Integer
    Dim ilDnfCode As Integer
    Dim ilVef As Integer
    Dim llDate As Long
    
    ilDnfCode = imReRateDnfCode
    '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
    If tgBookByLineAssigned(ilBook).iReRateDnfCode = -1 Then 'Vehicle default
        If llSdf <> -1 Then
            ilVef = gBinarySearchVef(tmSdfExt(llSdf).iVefCode)
            If ilVef <> -1 Then
                ilDnfCode = tgMVef(ilVef).iDnfCode
            Else
                ilDnfCode = 0
            End If
        Else
            ilVef = gBinarySearchVef(tmClfR.iVefCode)
            If ilVef <> -1 Then
                ilDnfCode = tgMVef(ilVef).iDnfCode
            Else
                ilDnfCode = 0
            End If
        End If
    ElseIf tgBookByLineAssigned(ilBook).iReRateDnfCode = -2 Then 'Closest
        If llSdf <> -1 Then
            gUnpackDateLong tmSdfExt(llSdf).iDate(0), tmSdfExt(llSdf).iDate(1), llDate
            ilDnfCode = mFindClosestBook(llDate, llSdf) 'tmSdfExt(llSdf).iVefCode)
        Else
            ilDnfCode = 0
        End If
    ElseIf tgBookByLineAssigned(ilBook).iReRateDnfCode = -3 Then 'Purchase
        ilDnfCode = tmClfP.iDnfCode
    Else
        '1/25/22 - JW - TTP 10385 - ReRate overflow error / possible overflow errors
        ilDnfCode = tgBookByLineAssigned(ilBook).iReRateDnfCode
    End If
    mGetDnfByContractLine = ilDnfCode
End Function

Private Sub mPopExcludeBooks()
    Dim slSQLQuery As String
    Dim rst_eff As ADODB.Recordset

    ReDim igExcludeDnfCode(0 To 0) As Integer
    slSQLQuery = "Select effLong1 from EFF_Extra_Fields"
    slSQLQuery = slSQLQuery & " Where effType = 'E'"
    slSQLQuery = slSQLQuery & " Order by effLong1"
    Set rst_eff = gSQLSelectCall(slSQLQuery)
    Do While Not rst_eff.EOF
        igExcludeDnfCode(UBound(igExcludeDnfCode)) = rst_eff!effLong1
        ReDim Preserve igExcludeDnfCode(0 To UBound(igExcludeDnfCode) + 1) As Integer
        rst_eff.MoveNext
    Loop
    
End Sub

Private Sub mSaveReRateBookByLine(llSdf As Long, ilDnfCode As Integer)
    Dim blFound As Boolean
    Dim illoop As Integer
    
    blFound = False
    For illoop = 0 To UBound(tmReRateBookDnfCodes) - 1 Step 1
        If (tmReRateBookDnfCodes(illoop).lChfCode = tmSdfExt(llSdf).lChfCode) And (tmReRateBookDnfCodes(illoop).iLineNo = tmSdfExt(llSdf).iLineNo) And (tmReRateBookDnfCodes(illoop).iDnfCode = ilDnfCode) And (tmReRateBookDnfCodes(illoop).sType = "S") Then
            blFound = True
            Exit For
        End If
    Next illoop
    If Not blFound Then
        tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).lChfCode = tmSdfExt(llSdf).lChfCode
        tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).sType = "S"
        tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).iLineNo = tmSdfExt(llSdf).iLineNo
        tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).iVefCode = tmSdfExt(llSdf).iVefCode
        tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).iDrfCode = 0
        tmReRateBookDnfCodes(UBound(tmReRateBookDnfCodes)).iDnfCode = ilDnfCode
        ReDim Preserve tmReRateBookDnfCodes(0 To UBound(tmReRateBookDnfCodes) + 1) As RERATEBOOKDNFCODES
    End If

End Sub
Public Sub mPopDemo()
'
'   mDemoPop
'   Where:
'
    Dim ilRet As Integer
    Dim illoop As Integer
    imCbcDemoBottom = cbcDemo.Top + cbcDemo.Height
    ilRet = gPopMnfPlusFieldsBox(ExptReRate, cbcDemo, tgDemoCode(), sgDemoCodeTag, "D")
    cbcDemo.SetDropDownWidth (cbcDemo.Width)
    'cbcDemo.SetDropDownNumRows (4)
    cbcDemo.PopUpListDirection "B"
    cbcDemo.ZOrder vbBringToFront
End Sub

'3/4/21 - TTP 10088: Sort by Product or Contract (ASC/DESC)
'mSortByColumn Sort by the indicated column.
'Sorts the Contract Grid based on ilSortColumn.
'If Sorting by a new Column, Sort it ASC
'If sorting by SAME Column, Swap Sort Order (ASC/DESC)
'If ilSortColumn is Negative, Re-Sort column, without swapping Order (Used grid Reloads)
Private Sub mSortByColumn(ByVal ilSortColumn As Integer)
    If imSortDir = 0 Then imSortDir = flexSortGenericDescending
    If ilSortColumn = 0 Or ilSortColumn = 3 Then Exit Sub 'Ignore sorting on Gen and Revision
    ' Hide the FlexGrid.
    grdCntr.Visible = False
    If ilSortColumn < 0 Then ' Use Negative column Number to "Refresh Column Sorting", (Dont Change Direction)
        ilSortColumn = Abs(ilSortColumn)
        If imSortDir = flexSortGenericAscending Then
            imSortDir = flexSortGenericDescending
        Else
            imSortDir = flexSortGenericAscending
        End If
    End If
    'Sort using the column specified in ilSortColumn.
    grdCntr.Col = ilSortColumn
    grdCntr.ColSel = ilSortColumn
    grdCntr.Row = 0
    grdCntr.RowSel = 0
    ' If this is a new sort column, sort ascending.
    ' Otherwise switch which sort order we use.
    If imSortColumn <> ilSortColumn Then
        grdCntr.Sort = flexSortGenericAscending
        imSortDir = flexSortGenericAscending
    ElseIf imSortDir = flexSortGenericAscending Then
        grdCntr.Sort = flexSortGenericDescending
        imSortDir = flexSortGenericDescending
    Else
        grdCntr.Sort = flexSortGenericAscending
        imSortDir = flexSortGenericAscending
    End If
    ' Display the FlexGrid.
    imSortColumn = ilSortColumn
    grdCntr.Refresh
    grdCntr.Visible = True
End Sub

Function mGetcxfComment(llCxfCode As Long) As String
    Dim ilRet As Integer
    mGetcxfComment = ""
    If ckcComment.Value = False Then Exit Function 'Comments off, dont spend time looking it up
    tmCxfSrchKey.lCode = llCxfCode
    If tmCxfSrchKey.lCode <> 0 Then
        imCxfRecLen = Len(tmCxf) '5027
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            mGetcxfComment = gStripChr0(Left$(tmCxf.sComment, 40))
        End If
    End If
End Function

'TTP 10172 - 7/1/21 - JW - function Clear BookByLine Array - to make it consistant
Sub mClearBookByLine(Optional blResetLastOption As Boolean = True)
    'reset "Research Book Name" caches
    ReDim tgBookByLineCntr(0 To 0) As BOOKBYLINECNTR
    ReDim tgBookByLineAssigned(0 To 0) As BOOKBYLINEASSIGNED
    'disable "Research Book Name" mode
    bmBookByLine = False
    rbcReRateBookByLine.Value = False
    
    'set last used option back
    If blResetLastOption Then rbcReRateBook(imReRateLastBookMode).Value = True
    tgBookByLineAssigned(0).iReRateDnfCode = 0
End Sub

'TTP 10258: ReRate - make it work without requiring Office
Sub mGetCSVFilename()
    If lbcAdvertiser.Text = "" Then
        edcCSV.Text = ""
    Else
        'has a Advertiser Selected
        If ckcSummary.Value = vbChecked Then
            'in Summary Mode
            'TTP 10994 - Rerate report: Error #76 when generating the Excel version of the report when it's for a direct advertiser
            'smToCSV = sgExportPath & "ReRateSummaryExport_" & Trim$(lbcAdvertiser.Text) & ".CSV"
            smToCSV = sgExportPath & "ReRateSummaryExport_" & gRemoveIllegalPastedChar(Trim$(lbcAdvertiser.Text)) & ".CSV"
            edcCSV.Text = smToCSV
        Else
            'Not in Summary Mode
            'TTP 10994 - Rerate report: Error #76 when generating the Excel version of the report when it's for a direct advertiser
            'smToCSV = sgExportPath & "ReRateExport_" & Trim$(lbcAdvertiser.Text) & ".CSV"
            smToCSV = sgExportPath & "ReRateExport_" & gRemoveIllegalPastedChar(Trim$(lbcAdvertiser.Text)) & ".CSV"
            edcCSV.Text = smToCSV
        End If
    End If
End Sub
