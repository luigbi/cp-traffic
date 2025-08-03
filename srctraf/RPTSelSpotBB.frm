VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelSpotBB 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spot Business Booked"
   ClientHeight    =   6120
   ClientLeft      =   165
   ClientTop       =   1305
   ClientWidth     =   9270
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   9270
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "RPTSelSpotBB.frx":0000
      Left            =   6000
      List            =   "RPTSelSpotBB.frx":0002
      TabIndex        =   95
      Top             =   480
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   71
      Top             =   615
      Width           =   2055
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6675
      Top             =   -180
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7215
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   -15
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7575
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   165
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   30
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4320
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Txt"
      Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Frame frcCopies 
      Caption         =   "Printing"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox edcCopies 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   73
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   78
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox cbcFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   75
         Top             =   270
         Width           =   2925
      End
      Begin VB.TextBox edcFileName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   77
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   74
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   76
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Revenue on the Books Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4755
      Left            =   75
      TabIndex        =   79
      Top             =   1365
      Width           =   9090
      Begin VB.PictureBox pbcSelC 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Height          =   4515
         Left            =   75
         ScaleHeight     =   4515
         ScaleWidth      =   4710
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   4710
         Begin V81TrafficReports.CSI_Calendar calStart 
            Height          =   285
            Left            =   675
            TabIndex        =   91
            Top             =   345
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Text            =   "6/17/2011"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   0   'False
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
         Begin VB.PictureBox plcAdj 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   4095
            TabIndex        =   97
            Top             =   3600
            Width           =   4095
            Begin VB.CheckBox ckcAdj 
               Caption         =   "Air Time"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2160
               TabIndex        =   99
               Top             =   0
               Width           =   1185
            End
            Begin VB.CheckBox ckcAdj 
               Caption         =   "Rep "
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1320
               TabIndex        =   98
               Top             =   0
               Width           =   705
            End
         End
         Begin VB.CheckBox ckcUseSS 
            Caption         =   "Use Sales Src as Major Sort"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2040
            TabIndex        =   96
            Top             =   3900
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.PictureBox plcGameSubTotal 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   3135
            TabIndex        =   92
            Top             =   4200
            Width           =   3135
            Begin VB.OptionButton rbcGameSubTotal 
               Caption         =   "Separate"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1920
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton rbcGameSubTotal 
               Caption         =   "Combine"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   720
               TabIndex        =   93
               Top             =   0
               Value           =   -1  'True
               Width           =   1110
            End
         End
         Begin VB.PictureBox plcPerType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   4455
            TabIndex        =   8
            Top             =   30
            Width           =   4455
            Begin VB.OptionButton rbcPerType 
               Caption         =   "Cal"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   3450
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   690
            End
            Begin VB.OptionButton rbcPerType 
               Caption         =   "Week"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   765
               TabIndex        =   9
               Top             =   0
               Width           =   870
            End
            Begin VB.OptionButton rbcPerType 
               Caption         =   "Std"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1740
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   0
               Width           =   690
            End
            Begin VB.OptionButton rbcPerType 
               Caption         =   "Corp"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2565
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   0
               Width           =   765
            End
         End
         Begin VB.ComboBox cbcSortVG 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2385
            TabIndex        =   29
            Top             =   1230
            Width           =   1500
         End
         Begin VB.ComboBox cbcSort3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   27
            Top             =   1230
            Width           =   1500
         End
         Begin VB.ComboBox cbcSort2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2340
            TabIndex        =   23
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox cbcSort1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   20
            Top             =   735
            Width           =   1500
         End
         Begin VB.CheckBox ckcSkipVG 
            Caption         =   "Skip"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   3930
            TabIndex        =   30
            Top             =   1230
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.CheckBox ckcSkip3 
            Caption         =   "Skip"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1650
            TabIndex        =   28
            Top             =   1230
            Width           =   660
         End
         Begin VB.CheckBox ckcSkip2 
            Caption         =   "Skip"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   3930
            TabIndex        =   24
            Top             =   705
            Width           =   720
         End
         Begin VB.CheckBox ckcSkip1 
            Caption         =   "Skip"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1650
            TabIndex        =   21
            Top             =   735
            Width           =   660
         End
         Begin VB.PictureBox plcTotalsBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   4365
            TabIndex        =   31
            Top             =   1605
            Width           =   4365
            Begin VB.OptionButton rbcTotalsBy 
               Caption         =   "Summary"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   2940
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   0
               Width           =   1185
            End
            Begin VB.OptionButton rbcTotalsBy 
               Caption         =   "Advt"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2115
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   0
               Width           =   765
            End
            Begin VB.OptionButton rbcTotalsBy 
               Caption         =   "Contract"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   990
               TabIndex        =   32
               Top             =   0
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin VB.CheckBox ckcShowSlspSplit 
            Caption         =   "Show Slsp Splits"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   61
            Top             =   3900
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox edcPeriods 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3990
            MaxLength       =   2
            TabIndex        =   18
            Top             =   270
            Width           =   420
         End
         Begin VB.PictureBox plcGrossNet 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   2775
            TabIndex        =   35
            Top             =   1845
            Width           =   2775
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Spots"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   1920
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   440
               TabIndex        =   36
               Top             =   0
               Value           =   -1  'True
               Width           =   870
            End
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1380
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   0
               Width           =   630
            End
         End
         Begin VB.PictureBox plcAllTypes 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1470
            Left            =   90
            ScaleHeight     =   1470
            ScaleWidth      =   4575
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   2145
            Width           =   4575
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Cash"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   21
               Left            =   3480
               TabIndex        =   51
               Top             =   480
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "OCBB"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   20
               Left            =   3780
               TabIndex        =   90
               Top             =   1170
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   750
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   19
               Left            =   3135
               TabIndex        =   89
               Top             =   1170
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3960
               TabIndex        =   43
               Top             =   -30
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   18
               Left            =   1905
               TabIndex        =   60
               Top             =   1170
               Value           =   1  'Checked
               Width           =   1155
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   17
               Left            =   795
               TabIndex        =   59
               Top             =   1170
               Value           =   1  'Checked
               Width           =   1005
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Non-Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   16
               Left            =   3240
               TabIndex        =   58
               Top             =   930
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   15
               Left            =   2370
               TabIndex        =   57
               Top             =   930
               Value           =   1  'Checked
               Width           =   765
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "H/C"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   14
               Left            =   1575
               TabIndex        =   56
               Top             =   930
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "NTR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   13
               Left            =   795
               TabIndex        =   55
               Top             =   930
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Rep"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   12
               Left            =   2865
               TabIndex        =   54
               Top             =   690
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "AirTime"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   11
               Left            =   1710
               TabIndex        =   53
               Top             =   690
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Trade"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   795
               TabIndex        =   52
               Top             =   690
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   2400
               TabIndex        =   50
               Top             =   450
               Width           =   945
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   1560
               TabIndex        =   49
               Top             =   450
               Width           =   765
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "PI"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   795
               TabIndex        =   48
               Top             =   450
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   3240
               TabIndex        =   47
               Top             =   240
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   2040
               TabIndex        =   46
               Top             =   210
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   795
               TabIndex        =   45
               Top             =   210
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2760
               TabIndex        =   44
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1710
               TabIndex        =   42
               Top             =   -30
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcAllTypes 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   795
               TabIndex        =   41
               Top             =   -30
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.TextBox edcMonth 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2610
            MaxLength       =   3
            TabIndex        =   16
            Top             =   240
            Width           =   420
         End
         Begin VB.TextBox edcStart 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   675
            MaxLength       =   4
            TabIndex        =   14
            Top             =   240
            Width           =   600
         End
         Begin VB.TextBox edcContract 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3540
            MaxLength       =   9
            TabIndex        =   39
            Top             =   1845
            Width           =   1125
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contr#"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2895
            TabIndex        =   38
            Top             =   1845
            Width           =   645
         End
         Begin VB.Label lacSort2 
            Appearance      =   0  'Flat
            Caption         =   "Sort #2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2385
            TabIndex        =   22
            Top             =   540
            Width           =   675
         End
         Begin VB.Label lacSort1 
            Appearance      =   0  'Flat
            Caption         =   "Sort #1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   90
            TabIndex        =   19
            Top             =   540
            Width           =   675
         End
         Begin VB.Label lacStart 
            Appearance      =   0  'Flat
            Caption         =   "Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   90
            TabIndex        =   13
            Top             =   285
            Width           =   465
         End
         Begin VB.Label lacPeriods 
            Appearance      =   0  'Flat
            Caption         =   "#Periods"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3180
            TabIndex        =   17
            Top             =   300
            Width           =   810
         End
         Begin VB.Label lacSort3 
            Appearance      =   0  'Flat
            Caption         =   "Sort #3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   90
            TabIndex        =   25
            Top             =   1065
            Width           =   675
         End
         Begin VB.Label lacSortVG 
            Appearance      =   0  'Flat
            Caption         =   "Vehicle Group"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2370
            TabIndex        =   26
            Top             =   1065
            Width           =   1290
         End
         Begin VB.Label lacMonth 
            Appearance      =   0  'Flat
            Caption         =   "Start Month"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1515
            TabIndex        =   15
            Top             =   300
            Width           =   1110
         End
      End
      Begin VB.PictureBox pbcOption 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Height          =   4020
         Left            =   4785
         ScaleHeight     =   4020
         ScaleWidth      =   4230
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   180
         Width           =   4230
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   6
            ItemData        =   "RPTSelSpotBB.frx":0004
            Left            =   2295
            List            =   "RPTSelSpotBB.frx":0006
            MultiSelect     =   2  'Extended
            TabIndex        =   87
            Top             =   2550
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   5
            ItemData        =   "RPTSelSpotBB.frx":0008
            Left            =   2280
            List            =   "RPTSelSpotBB.frx":000F
            MultiSelect     =   2  'Extended
            TabIndex        =   86
            Top             =   2445
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   4
            ItemData        =   "RPTSelSpotBB.frx":0016
            Left            =   2280
            List            =   "RPTSelSpotBB.frx":0018
            MultiSelect     =   2  'Extended
            TabIndex        =   85
            Top             =   2340
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.CheckBox ckcAllSort3 
            Caption         =   "All Sort #3"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   1950
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CheckBox ckcAllSort2 
            Caption         =   "All Sort #2"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2265
            TabIndex        =   64
            Top             =   0
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   3
            ItemData        =   "RPTSelSpotBB.frx":001A
            Left            =   2265
            List            =   "RPTSelSpotBB.frx":001C
            MultiSelect     =   2  'Extended
            TabIndex        =   65
            Top             =   300
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.CheckBox ckcAllSortVG 
            Caption         =   "All Group Items"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2265
            TabIndex        =   68
            Top             =   1950
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   2
            ItemData        =   "RPTSelSpotBB.frx":001E
            Left            =   2265
            List            =   "RPTSelSpotBB.frx":0020
            MultiSelect     =   2  'Extended
            TabIndex        =   69
            Top             =   2235
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   1
            ItemData        =   "RPTSelSpotBB.frx":0022
            Left            =   120
            List            =   "RPTSelSpotBB.frx":0024
            MultiSelect     =   2  'Extended
            TabIndex        =   67
            Top             =   2235
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   0
            ItemData        =   "RPTSelSpotBB.frx":0026
            Left            =   120
            List            =   "RPTSelSpotBB.frx":0028
            MultiSelect     =   2  'Extended
            TabIndex        =   63
            Top             =   300
            Width           =   1950
         End
         Begin VB.CheckBox ckcAllSort1 
            Caption         =   "All Sort #1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   0
            Width           =   1905
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   72
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   70
      Top             =   105
      Width           =   2805
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   1395
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   555
         Width           =   750
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   1035
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RptSelSpotBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptSelSpotBB.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  ckcAvails_Click                                                                       *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelSpotBB.Frm - Quarterly Booked Report
'                           Show Spot Counts & Avails by Advt
'
' Release: 4.3
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAllSort1 As Integer 'True=Set list box; False= don't change list box
Dim imSetAllSort2 As Integer
Dim imSetAllSort3 As Integer
Dim imSetAllSortVG As Integer

Dim imAllSort1Clicked As Integer
Dim imAllSort2Clicked As Integer
Dim imAllSort3Clicked As Integer
Dim imAllSortVGClicked As Integer

Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name

Dim imCodes() As Integer
Dim smLogUserCode As String
'Import contract report
'Spot week Dump
Dim imTerminate As Integer
Dim imSort1Index As Integer
Dim imSort2Index As Integer
Dim imSort3Index As Integer
Dim imSort1PrevIndex As Integer
Dim imSort2PrevIndex As Integer
Dim imSort3PrevIndex As Integer
'
'       mFillSortOptions - Populate the combo boxes with the list of Sort Options
'       <input>  cbcCombo - combo control
'                ilShowNone - true if NONE allowed as an option
'
Public Sub mFillSortOptions(cbcCombo As Control, ilShowNone As Integer)

        cbcCombo.Clear

        If ilShowNone Then
            cbcCombo.AddItem "None"
        End If
        cbcCombo.AddItem "Advertiser"
        cbcCombo.AddItem "Agency"
        cbcCombo.AddItem "Bus Category"
        cbcCombo.AddItem "Prod Protection"
        cbcCombo.AddItem "Salesperson"
        cbcCombo.AddItem "Vehicle"
        cbcCombo.ListIndex = 0

        Exit Sub
End Sub

'Dim tmSRec As LPOPREC
'Rate Card
Private Sub cbcFileType_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcFileType.Text <> "" Then
            gManLookAhead cbcFileType, imBSMode, imComboBoxIndex
        End If
        imFTSelectedIndex = cbcFileType.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcFileType_Click()
    imComboBoxIndex = cbcFileType.ListIndex
    imFTSelectedIndex = cbcFileType.ListIndex
    mSetCommands
End Sub
Private Sub cbcFileType_GotFocus()
    If cbcFileType.Text = "" Then
        cbcFileType.ListIndex = 0
    End If
    imComboBoxIndex = cbcFileType.ListIndex
    gCtrlGotFocus cbcFileType
End Sub
Private Sub cbcFileType_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcFileType_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcFileType.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcSort1_Click()

    imSort1Index = cbcSort1.ListIndex
    If imSort1Index + 1 = imSort2Index Or imSort1Index + 1 = imSort3Index Then
        'error, cannot have the same sort parameter defined for 2 sort fields
        cbcSort1.ListIndex = imSort1PrevIndex
        MsgBox "Same sort selection as Sort #2 or Sort #3; select another"
    Else
        mHidePrevSelection imSort1PrevIndex + 1, 1
        mShowCurrSelection imSort1Index + 1, 1
        imSort1PrevIndex = imSort1Index
        'mIsVehicleSelected imSort1Index + 1, 1
        mIsVehicleSelected 1            '8-23-11
        imSetAllSort1 = False
        ckcAllSort1.Value = vbUnchecked
        imSetAllSort1 = True
    End If
    mSetCommands
End Sub

Private Sub cbcSort2_Click()

    imSort2Index = cbcSort2.ListIndex
    If (imSort2Index = imSort1Index + 1 Or imSort2Index = imSort3Index) And (imSort2Index <> 0) Then
        'error, cannot have the same sort parameter defined for 2 sort fields
         MsgBox "Same sort selection as Sort #1 or Sort #3; select another"
         cbcSort2.ListIndex = 0
   Else
        mHidePrevSelection imSort2PrevIndex, 2
        mShowCurrSelection imSort2Index, 2
        imSort2PrevIndex = imSort2Index
       ' mIsVehicleSelected imSort2Index, 2
        mIsVehicleSelected 2            '8-23-11
        imSetAllSort2 = False
        ckcAllSort2.Value = vbUnchecked
        imSetAllSort2 = True

    End If
    mSetCommands
End Sub

Private Sub cbcSort3_Click()
    imSort3Index = cbcSort3.ListIndex
    If (imSort3Index = imSort1Index + 1 Or imSort3Index = imSort2Index) And (imSort3Index <> 0) Then
        'error, cannot have the same sort parameter defined for 2 sort fields
        MsgBox "Same sort selection as Sort #1 or Sort #2; select another"
         cbcSort3.ListIndex = 0         '8-23-11
    Else
        mHidePrevSelection imSort3PrevIndex, 3
        mShowCurrSelection imSort3Index, 3
        imSort3PrevIndex = imSort3Index
        'mIsVehicleSelected imSort3Index, 3
        mIsVehicleSelected 3            '8-23-11
        imSetAllSort3 = False
        ckcAllSort3.Value = vbUnchecked
        imSetAllSort3 = True
    End If
    mSetCommands
End Sub

Private Sub cbcSortVG_click()
Dim ilLoop As Integer
Dim ilSetIndex As Integer
Dim ilRet As Integer

        ilLoop = cbcSortVG.ListIndex
        ilSetIndex = gFindVehGroupInx(ilLoop, tgVehicleSets1())
        If ilSetIndex > 0 Then
            smVehGp5CodeTag = ""
            ilRet = gPopMnfPlusFieldsBox(RptSelSpotBB, lbcSelection(6), tgSOCode(), smVehGp5CodeTag, "H" & Trim$(str$(ilSetIndex)))
            lbcSelection(6).Visible = True
            ckcAllSortVG.Visible = True
            ckcSkipVG.Visible = True
        Else
            lbcSelection(6).Visible = False
            ckcAllSortVG.Visible = False
            ckcSkipVG.Visible = False
            ckcSkipVG.Value = vbUnchecked

        End If
        mSetCommands
End Sub
Private Sub ckcAllSort1_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllSort1.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllSort1 Then
        imAllSort1Clicked = True
        llRg = CLng(lbcSelection(imSort1Index).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(imSort1Index).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllSort1Clicked = False
    End If
    mSetCommands
End Sub
Private Sub ckcAllSort1_GotFocus()
    gCtrlGotFocus ckcAllSort1
End Sub

Private Sub ckcAllSort2_click()
 'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllSort2.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllSort2 Then
        imAllSort2Clicked = True
        llRg = CLng(lbcSelection(imSort2Index - 1).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(imSort2Index - 1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllSort2Clicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAllSort3_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllSort3.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllSort3 Then
        imAllSort3Clicked = True
        llRg = CLng(lbcSelection(imSort3Index - 1).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(imSort3Index - 1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllSort3Clicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAllSort3_GotFocus()
    gCtrlGotFocus ckcAllSort3
End Sub

Private Sub ckcAllSortVG_Click()
    Dim Value As Integer
    Value = False
    If ckcAllSortVG.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllSortVG Then
        imAllSortVGClicked = True
        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllSort3Clicked = False
    End If
    mSetCommands

End Sub

Private Sub ckcAllTypes_Click(Index As Integer)
Dim ilRet As Integer
    If Index = 12 Then      'rep
        If ckcAllTypes(Index).Value = vbUnchecked Then
            lbcSelection(5).Clear
            ilRet = gRptVehPop(RptSelSpotBB, lbcSelection(5), VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH)
        Else
            lbcSelection(5).Clear
            ilRet = gRptVehPop(RptSelSpotBB, lbcSelection(5), VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHREP_W_CLUSTER + VEHREP_WO_CLUSTER + ACTIVEVEH)
        End If
        If cbcSort1.ListIndex = 5 Then            'vehicle?
            ckcAllSort1.Value = vbUnchecked
        ElseIf cbcSort2.ListIndex = 6 Then
            ckcAllSort2.Value = vbUnchecked
        ElseIf cbcSort3.ListIndex = 6 Then
            ckcAllSort3.Value = vbUnchecked
        End If

    End If
    mSetCommands
End Sub

Private Sub cmcBrowse_Click()
    gAdjustCDCFilter imFTSelectedIndex, cdcSetup
    cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    cdcSetup.fileName = edcFileName.Text
    cdcSetup.InitDir = Left$(sgRptSavePath, Len(sgRptSavePath) - 1)
    cdcSetup.Action = 2    'DLG_FILE_SAVE
    edcFileName.Text = cdcSetup.fileName
    mSetCommands
    gChDrDir        '3-25-03
    'ChDrive Left$(sgCurDir, 2)  'Set the default drive
    'ChDir sgCurDir              'set the default path
End Sub
Private Sub cmcBrowse_GotFocus()
    gCtrlGotFocus cmcBrowse
End Sub
Private Sub cmcCancel_Click()
    If igGenRpt Then
        Exit Sub
    End If
    'mTerminate True
    mTerminate False
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcGen_Click()
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    igOutput = frcOutput.Enabled
    igCopies = frcCopies.Enabled
    'igWhen = frcWhen.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    'igReportType = frcRptType.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    'frcWhen.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False
    'frcRptType.Enabled = False

    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportSpotBB() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If

        ilRet = gCmcGenSpotBB()
        If ilRet = -1 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            'mTerminate
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        ElseIf ilRet = 0 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If

        Screen.MousePointer = vbHourglass

        gGenSpotBB
        Screen.MousePointer = vbDefault
        If rbcOutput(0).Value Then
            DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
            igDestination = 0
            Report.Show vbModal
        ElseIf rbcOutput(1).Value Then
            ilCopies = Val(edcCopies.Text)
            ilRet = gOutputToPrinter(ilCopies)
        Else
            slFileName = edcFileName.Text
            'ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
        End If
    Next ilJobs
    
    sgVehicleTag = ""
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCRGrfClear
    Screen.MousePointer = vbDefault
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    Exit Sub
End Sub
Private Sub cmcGen_GotFocus()
    gCtrlGotFocus cmcGen
End Sub
Private Sub cmcGen_KeyDown(KeyCode As Integer, Shift As Integer)
    imGenShiftKey = Shift
End Sub
Private Sub cmcList_Click()
    If igGenRpt Then
        Exit Sub
    End If
    mTerminate True
End Sub
Private Sub cmcSetup_Click()
    'cdcSetup.Flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub
Private Sub edcCopies_Change()
    mSetCommands
End Sub
Private Sub edcCopies_GotFocus()
    gCtrlGotFocus edcCopies
End Sub
Private Sub edcCopies_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcFileName_Change()
    mSetCommands
End Sub
Private Sub edcFileName_GotFocus()
    gCtrlGotFocus edcFileName
End Sub
Private Sub edcFileName_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer

    ilPos = InStr(edcFileName.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcFileName.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    If ((KeyAscii <> KEYBACKSPACE) And (KeyAscii <= 32)) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KEYDOWN) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcMonth_Change()
    mSetCommands
End Sub

Private Sub edcMonth_GotFocus()
    gCtrlGotFocus edcMonth
End Sub
Private Sub edcPeriods_Change()
    mSetCommands
End Sub

Private Sub edcPeriods_gotfocus()
    gCtrlGotFocus edcPeriods
End Sub

Private Sub edcStart_Change()
    mSetCommands
End Sub
Private Sub edcStart_gotfocus()
    gCtrlGotFocus edcStart
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    RptSelSpotBB.Refresh
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    imTerminate = False
    igGenRpt = False
    mParseCmmdLine
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mInit
    If imTerminate Then 'Used for print only
        mTerminate True
        Exit Sub
    End If
    'RptSelSpotBB.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgMNFCodeRpt
    Erase imCodes
    PECloseEngine
    
    Set RptSelSpotBB = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcRptType_Click()
Dim ilListIndex As Integer

        ilListIndex = lbcRptType.ListIndex
        If ilListIndex = REV_ON_BOOKS Then
        Else            'SPOTREV_REGISTER
            frcOption.Caption = "Spot Revenue Register Selection"
            lacPeriods.Visible = False
            edcPeriods.Visible = False
            edcPeriods.Text = "1"           'Spot Revenue Register is always 1 month or week
            plcGrossNet.Visible = False     'Spot Revenue Register will show both gross & net for the period
            ckcUseSS.Move plcGrossNet.Left, plcGrossNet.Top - 30
            ckcUseSS.Visible = True
        End If
End Sub

'       item selected within one of the 7 list boxes
Private Sub lbcSelection_Click(Index As Integer)
    
    If Index = 0 Then           'single advt selected
        If imSort1Index = 0 Then
            imSetAllSort1 = False
            ckcAllSort1.Value = vbUnchecked
            imSetAllSort1 = True
        ElseIf imSort2Index - 1 = 0 Then
            imSetAllSort2 = False
            ckcAllSort2.Value = vbUnchecked
            imSetAllSort2 = True
        ElseIf imSort3Index - 1 = 0 Then
            imSetAllSort3 = False
            ckcAllSort3.Value = vbUnchecked
            imSetAllSort3 = True
        End If
    ElseIf Index = 1 Then
        If imSort1Index = 1 Then
            imSetAllSort1 = False
            ckcAllSort1.Value = vbUnchecked
            imSetAllSort1 = True
        ElseIf imSort2Index - 1 = 1 Then
            imSetAllSort2 = False
            ckcAllSort2.Value = vbUnchecked
            imSetAllSort2 = True
        ElseIf imSort3Index - 1 = 1 Then
            imSetAllSort3 = False
            ckcAllSort3.Value = vbUnchecked
            imSetAllSort3 = True
        End If
    ElseIf Index = 2 Then
        If imSort1Index = 2 Then
            imSetAllSort1 = False
            ckcAllSort1.Value = vbUnchecked
            imSetAllSort1 = True
        ElseIf imSort2Index - 1 = 2 Then
            imSetAllSort2 = False
            ckcAllSort2.Value = vbUnchecked
            imSetAllSort2 = True
        ElseIf imSort3Index - 1 = 2 Then
            imSetAllSort3 = False
            ckcAllSort3.Value = vbUnchecked
            imSetAllSort3 = True
        End If
    ElseIf Index = 3 Then
        If imSort1Index = 3 Then
            imSetAllSort1 = False
            ckcAllSort1.Value = vbUnchecked
            imSetAllSort1 = True
        ElseIf imSort2Index - 1 = 3 Then
            imSetAllSort2 = False
            ckcAllSort2.Value = vbUnchecked
            imSetAllSort2 = True
        ElseIf imSort3Index - 1 = 3 Then
            imSetAllSort3 = False
            ckcAllSort3.Value = vbUnchecked
            imSetAllSort3 = True
        End If
    ElseIf Index = 4 Then
        If imSort1Index = 4 Then
            imSetAllSort1 = False
            ckcAllSort1.Value = vbUnchecked
            imSetAllSort1 = True
        ElseIf imSort2Index - 1 = 4 Then
            imSetAllSort2 = False
            ckcAllSort2.Value = vbUnchecked
            imSetAllSort2 = True
        ElseIf imSort3Index - 1 = 4 Then
            imSetAllSort3 = False
            ckcAllSort3.Value = vbUnchecked
            imSetAllSort3 = True
        End If
    ElseIf Index = 5 Then
        If imSort1Index = 5 Then
            imSetAllSort1 = False
            ckcAllSort1.Value = vbUnchecked
            imSetAllSort1 = True
        ElseIf imSort2Index - 1 = 5 Then
            imSetAllSort2 = False
            ckcAllSort2.Value = vbUnchecked
            imSetAllSort2 = True
        ElseIf imSort3Index - 1 = 5 Then
            imSetAllSort3 = False
            ckcAllSort3.Value = vbUnchecked
            imSetAllSort3 = True
        End If
    ElseIf Index = 6 Then
        imSetAllSortVG = False
        ckcAllSortVG.Value = vbUnchecked
        imSetAllSortVG = True
    End If
    


    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/16/93       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*            Place focus before populating all lists  *                                                   *
'*******************************************************
Private Sub mInit()
Dim ilRet As Integer
Dim ilLoop As Integer
Dim slStr As String
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If
    'Set options for report generate
    'hdJob = rpcRpt.hJob
    'ilMultiTable = True
    'ilDummy = LlSetOption(hdJob, LL_OPTION_SORTVARIABLES, True)
    'ilDummy = LlSetOption(hdJob, LL_OPTION_ONLYONETABLE, Not ilMultiTable)

    RptSelSpotBB.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllSort1Clicked = False
    imSetAllSort1 = True
    imSetAllSort2 = True
    imSetAllSort3 = True
    imSetAllSortVG = True
    
    imAllSort1Clicked = False
    imAllSort2Clicked = False
    imAllSort3Clicked = False
    imAllSortVGClicked = False
    
    imSort1Index = -1       'sort selection #1 index
    imSort2Index = -1
    imSort3Index = -1
    imSort1PrevIndex = -1       'sort selection #1 index
    imSort2PrevIndex = -1
    imSort3PrevIndex = -1
    
    If tgSpf.sRUseCorpCal <> "Y" Then       'disable corp calendar if not defined
        rbcPerType(2).Enabled = False
    End If

    lacStart.Move 90, plcPerType.Top + plcPerType.Height + 30      'start year caption
    edcStart.MaxLength = 4                                          'year length (i.e. 1970,2011)
    edcStart.Move 600, plcPerType.Top + plcPerType.Height, 600
    lacMonth.Move 1320, lacStart.Top, 1110                               'start month caption
    edcMonth.Move 2400, edcStart.Top
    lacPeriods.Caption = "# Periods"
    lacPeriods.Move 3000, lacStart.Top, 960                          '# periods caption
    edcPeriods.Move 3820, edcStart.Top
    
    lacSort1.Move 90, edcStart.Top + edcStart.Height                        'Sort #1 caption
    cbcSort1.Move 90, lacSort1.Top + lacSort1.Height - 30                  'sort #1 list
    ckcSkip1.Move cbcSort1.Left + cbcSort1.Width, cbcSort1.Top             'skip Sort #1
    cbcSort2.Move ckcSkip1.Left + ckcSkip1.Width + 120, cbcSort1.Top        'sort #2 list
    ckcSkip2.Move cbcSort2.Left + cbcSort2.Width, cbcSort2.Top              'skip Sort #2
    lacSort2.Move cbcSort2.Left, lacSort1.Top                               'Sort #2 caption
    
    lacSort3.Move 90, cbcSort1.Top + cbcSort1.Height + 30                   'Sort #3 caption
    cbcSort3.Move 90, lacSort3.Top + lacSort3.Height - 30                   'sort #3 list
    ckcSkip3.Move cbcSort3.Left + cbcSort3.Width, cbcSort3.Top              'skip Sort #1
    cbcSortVG.Move ckcSkip3.Left + ckcSkip3.Width + 120, cbcSort3.Top       'sort #2 list
    ckcSkipVG.Move cbcSortVG.Left + cbcSortVG.Width, cbcSortVG.Top          'skip Sort #2
    lacSortVG.Move cbcSortVG.Left, lacSort3.Top                             'Sort VG caption
    
    plcTotalsBy.Move 90, cbcSort3.Top + cbcSort3.Height + 60                'totals by contract, advt, summary
    plcGrossNet.Move 90, plcTotalsBy.Top + plcTotalsBy.Height + 30          'gross or net
    
    lacContract.Move 2895, plcGrossNet.Top                       'single contract selection caption
    edcContract.Move 3555, plcGrossNet.Top - 30                  'single contract selection
    
    plcAllTypes.Move 90, plcGrossNet.Top + plcGrossNet.Height + 30        'all contract types

    plcAdj.Move 90, plcAllTypes.Top + plcAllTypes.Height                'include adjustments
    '12-1-16  move controls for new control for A/T adj
    'ckcShowSlspSplit.Move 2430, ckcAdj.Top                              'show slsp splits
    ckcShowSlspSplit.Move 90, plcAdj.Top + plcAdj.Height
    'plcGameSubTotal.Move 90, ckcAdj.Top + ckcAdj.Height
    plcGameSubTotal.Move 90, ckcShowSlspSplit.Top + ckcShowSlspSplit.Height
   
    lbcSelection(6).Move 2265, 2235         'this is for the vehicle group items, it always resides in the same loc.
    gCenterStdAlone RptSelSpotBB
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Private Sub mInitReport()
Dim ilRet As Integer
    'cbcWhenDay.AddItem "One Time"
    'cbcWhenDay.AddItem "Every M-F"
    'cbcWhenDay.AddItem "Every M-Sa"
    'cbcWhenDay.AddItem "Every M-Su"
    'cbcWhenDay.AddItem "Every Monday"
    'cbcWhenDay.AddItem "Every Tuesday"
    'cbcWhenDay.AddItem "Every Wednesday"
    'cbcWhenDay.AddItem "Every Thursday"
    'cbcWhenDay.AddItem "Every Friday"
    'cbcWhenDay.AddItem "Every Saturday"
    'cbcWhenDay.AddItem "Every Sunday"
    'cbcWhenDay.AddItem "Cal Month End+1"
    'cbcWhenDay.AddItem "Cal Month End+2"
    'cbcWhenDay.AddItem "Cal Month End+3"
    'cbcWhenDay.AddItem "Cal Month End+4"
    'cbcWhenDay.AddItem "Cal Month End+5"
    'cbcWhenDay.AddItem "Std Month End+1"
    'cbcWhenDay.AddItem "Std Month End+2"
    'cbcWhenDay.AddItem "Std Month End+3"
    'cbcWhenDay.AddItem "Std Month End+4"
    'cbcWhenDay.AddItem "Std Month End+5"
    'cbcWhenDay.ListIndex = 0
    'cbcWhenTime.AddItem "Right Now"
    'cbcWhenTime.AddItem "at 10PM"
    'cbcWhenTime.AddItem "at 12AM"
    'cbcWhenTime.AddItem "at 2AM"
    'cbcWhenTime.AddItem "at 4AM"
    'cbcWhenTime.AddItem "at 6AM"
    'cbcWhenTime.ListIndex = 0
    'Setup report output types
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    gPopExportTypes cbcFileType     '10-20-01
    pbcSelC.Visible = False
    'lbcRptType.Clear
    'lbcRptType.AddItem smSelectedRptName

    Screen.MousePointer = vbHourglass

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    lbcSelection(1).Clear
    lbcSelection(1).Tag = ""
    lbcSelection(2).Clear
    lbcSelection(2).Tag = ""
    lbcSelection(3).Clear
    lbcSelection(3).Tag = ""
    sgVehicleTag = ""               'populate vehicles when entering report
    lbcSelection(4).Clear
    lbcSelection(4).Tag = ""
    sgSalespersonTag = ""
    
    mFillSortOptions cbcSort1, False
    mFillSortOptions cbcSort2, True
    mFillSortOptions cbcSort3, True
    gPopVehicleGroups RptSelSpotBB!cbcSortVG, tgVehicleSets1(), True
    
    ilRet = gRptAdvtPop(RptSelSpotBB, lbcSelection(0))      'populate advt
    ilRet = gRptAgencyPop(RptSelSpotBB, lbcSelection(1))       'populate agencies
    ilRet = gRptMnfPop(RptSelSpotBB, "B", lbcSelection(2), tgMnfCodeCT(), sgMNFCodeTagCT)       'bus category
    ilRet = gRptMnfPop(RptSelSpotBB, "C", lbcSelection(3), tgMNFCodeRpt(), sgMNFCodeTagRpt)       'product protection
    ilRet = gRptSPersonPop(RptSelSpotBB, lbcSelection(4))   'populate salespeople
    'ilRet = gRptSellConvVehPop(RptSelSpotBB, lbcSelection(5))
    ilRet = gRptVehPop(RptSelSpotBB, lbcSelection(5), VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHREP_W_CLUSTER + VEHREP_WO_CLUSTER + ACTIVEVEH)
    'No office selection
    'ilRet = gPopOfficeSourceBox(RptSelSpotBB, lbcSelection(6), tgSOCodeCT(), sgSOCodeTagCT)

    'cbcSort1_Click
    frcOption.Enabled = True
    pbcSelC.Visible = True
    pbcOption.Visible = True
    ckcAllSort1.Visible = True
    lbcSelection(0).Visible = True                  'show budget name list box (base budget)
    pbcOption.Visible = True
    pbcOption.Enabled = True
    
    lbcRptType.AddItem "Revenue on the Books", REV_ON_BOOKS
    lbcRptType.AddItem "Spot Revenue Register", SPOTREV_REGISTER

    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
    End If
    
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slRptListCmmd As String

    slCommand = sgCommandStr    'Command$
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False  'True 'False
    '    imShowHelpmsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
    '    imShowHelpmsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpmsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelSpotBB, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Quarterly Booked Spots"
    '    igRptCallType = -1  'unused in standalone exe, CONTRACTSJOB 'SLSPCOMMSJOB   'LOGSJOB 'CONTRACTSJOB 'COPYJOB 'COLLECTIONSJOB'CONTRACTSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
    '    igRptType = -1  'unused in standalone exe   '3 'Log '1   'Contract    '0   'Summary '3 Program  '1  links
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)      'Function ID (what function calling this report if )
        End If
    'End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    Dim ilEnable As Integer
    Dim ilLoop As Integer
    Dim ilBoxSelected As Integer
    Dim ilBoxesToCheck As Integer
    
    ilEnable = True
    
    ilBoxSelected = True
    ilBoxesToCheck = 6      'default to check the vehicle group list
    If cbcSortVG.ListIndex = 0 Then
        ilBoxesToCheck = 5
    End If
    For ilLoop = 0 To ilBoxesToCheck     'check at least 1 entry in each shown list box has been selected
        If lbcSelection(ilLoop).Visible = True Then
            If lbcSelection(ilLoop).SelCount <= 0 Then
                ilBoxSelected = False
                ilEnable = False
                Exit For
            End If
        End If
    Next ilLoop
    
    If ilBoxSelected Then
        If ((edcStart.Text = "" Or edcPeriods.Text = "") And (rbcPerType(0).Value <> True)) Or (rbcPerType(0).Value = True And edcPeriods.Text = "") Then
            ilEnable = False
        End If
        If Not rbcPerType(0).Value = True Then      'weekly, doesnt have month input.  Test for month input on std, cal, corp
            If edcMonth.Text = "" Then
                ilEnable = False
            End If
        End If
    End If
    
    If ilEnable Then
        If rbcOutput(0).Value Then  'Display
            ilEnable = True
        ElseIf rbcOutput(1).Value Then  'Print
            If edcCopies.Text <> "" Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        Else    'Save As
            If (imFTSelectedIndex >= 0) And (edcFileName.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        End If
    End If
    cmcGen.Enabled = ilEnable
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate(ilFromCancel As Integer)
'
'   mTerminate
'   Where:
'

    If ilFromCancel Then
        igRptReturn = True
    Else
        igRptReturn = False
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RptSelSpotBB
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub



Private Sub plcAdj_Paint()
    plcAdj.Cls
    plcAdj.CurrentX = 0
    plcAdj.CurrentY = 0
    plcAdj.Print "Adjustments"
End Sub

Private Sub plcAllTypes_Paint()
    plcAllTypes.Cls
    plcAllTypes.CurrentX = 0
    plcAllTypes.CurrentY = 0
    plcAllTypes.Print "Select"
End Sub

Private Sub plcGameSubTotal_Paint()
    plcGameSubTotal.Cls
    plcGameSubTotal.CurrentX = 0
    plcGameSubTotal.CurrentY = 0
    plcGameSubTotal.Print "Events"
End Sub

Private Sub plcGrossNet_Paint()
    plcGrossNet.Cls
    plcGrossNet.CurrentX = 0
    plcGrossNet.CurrentY = 0
    plcGrossNet.Print "For"
End Sub

Private Sub plcPerType_Paint()
    plcPerType.Cls
    plcPerType.CurrentX = 0
    plcPerType.CurrentY = 0
    plcPerType.Print "Periods"
End Sub
Private Sub plcTotalsBy_Paint()
    plcTotalsBy.Cls
    plcTotalsBy.CurrentX = 0
    plcTotalsBy.CurrentY = 0
    plcTotalsBy.Print "Totals by"
End Sub

Private Sub rbcGrossNet_Click(Index As Integer)
    If Index = 2 Then
        ckcAllTypes(19).Visible = True          'spot option:  show fills and open/close BB types
        ckcAllTypes(20).Visible = True
    Else
        ckcAllTypes(19).Visible = False          '$ option:  always ignore fills and open/close BB types ( no $)
        ckcAllTypes(20).Visible = False
    End If
End Sub

Private Sub rbcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    'End of coded added
    If rbcOutput(Index).Value Then
        Select Case Index
            Case 0  'Display
                frcFile.Enabled = False
                frcCopies.Visible = False   'Print Box
                frcFile.Visible = False     'Save to File Box
                frcCopies.Enabled = False
                'frcWhen.Enabled = False
                'pbcWhen.Visible = False
            Case 1  'Print
                frcFile.Visible = False
                frcFile.Enabled = False
                frcCopies.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcCopies.Visible = True
            Case 2  'File
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcFile.Visible = True
        End Select
    End If
    mSetCommands
End Sub
Private Sub rbcOutput_GotFocus(Index As Integer)
    If imFirstTime Then
        mInitReport
        If imTerminate Then 'Used for print only
            mTerminate True
            Exit Sub
        End If
        imFirstTime = False
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub


Private Sub rbcPerType_Click(Index As Integer)
    If Index = 0 Then           'weekly
        lacStart.Caption = "Start Monday"
        edcStart.MaxLength = 10
        lacStart.Move 90, plcPerType.Top + plcPerType.Height + 30, 1680
        edcStart.Move 1290, plcPerType.Top + plcPerType.Height, 1200
        lacMonth.Visible = False
        edcMonth.Visible = False
        edcStart.Visible = False
        lacStart.Visible = True
        calStart.Move 1290, plcPerType.Top + plcPerType.Height, 1200
        calStart.Visible = True
        lacPeriods.Move 2730, lacStart.Top, 960
        edcPeriods.Move 3690, edcStart.Top
    Else
        lacStart.Caption = "Year"
        lacStart.Move 90, plcPerType.Top + plcPerType.Height + 30, 940      'start year caption
        edcStart.MaxLength = 4                                          'year length (i.e. 1970,2011)
        edcStart.Move 600, plcPerType.Top + plcPerType.Height, 600
        lacMonth.Move 1320, lacStart.Top, 1110                               'start month caption
        edcMonth.Move 2400, edcStart.Top
        lacMonth.Caption = "Start Month"
        lacPeriods.Caption = "# Periods"
        lacPeriods.Move 3000, lacStart.Top, 960                          '# periods caption
        edcPeriods.Move 3820, edcStart.Top
        calStart.Move 1710, 495, 615
        calStart.Visible = False
        edcStart.Visible = True
        lacMonth.Visible = True
        edcMonth.Visible = True
    End If
    
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
'           mHidePrevSelection - hide the previous list box selection based on
'           sort field #1, #2 or #3, new one should be shown
'           <input>  ilprevSortIndex - sort index ( none, adv, agy,bus cat, prod prot, slsp, vehicle ) that was changed
'                    ilWhichSort - sort field #1, 2, or 3
'
Public Sub mHidePrevSelection(ilPrevSortIndex As Integer, ilWhichSort As Integer)
        
        'sort field #1 must have a selection
        If ilWhichSort = 2 Then             'sort field 2
            ckcAllSort2.Visible = False
        ElseIf ilWhichSort = 3 Then         'sort field 3
            ckcAllSort3.Visible = False
        End If
        
        If ilPrevSortIndex = 1 Then            'advt
            lbcSelection(0).Visible = False
        ElseIf ilPrevSortIndex = 2 Then        'agy
            lbcSelection(1).Visible = False
        ElseIf ilPrevSortIndex = 3 Then        'bus cat
            lbcSelection(2).Visible = False
        ElseIf ilPrevSortIndex = 4 Then        'prod prot
            lbcSelection(3).Visible = False
        ElseIf ilPrevSortIndex = 5 Then        'slsp
            lbcSelection(4).Visible = False
            ckcShowSlspSplit.Visible = False
            ckcShowSlspSplit.Value = vbUnchecked
        ElseIf ilPrevSortIndex = 6 Then        'vehicle
            lbcSelection(5).Visible = False
            cbcSortVG.Enabled = False
            lbcSelection(6).Visible = False     'vehicle group items
            ckcAllSortVG.Visible = False
            ckcSkipVG.Visible = False
            ckcSkipVG.Value = vbUnchecked
        Else                                    'none for either sort #2 or sort #3
                                                'sort #1 cant be 0 as the index has been incremented
                                                'from the calling rtn
            
        End If
        Exit Sub
End Sub
'           mShowCurrSelection - show the list box of the sort selected
'           <input>  ilCurrSortIndex - sort index ( none, adv, agy,bus cat, prod prot, slsp, vehicle) selected
'                    ilWhichSort - sort field #1, 2, or 3
Public Sub mShowCurrSelection(ilCurrSortIndex As Integer, ilWhichSort As Integer)
Dim ilLeft As Integer
Dim ilTop As Integer
Dim slSortCaption As String

        If ilCurrSortIndex < 0 Then
            Exit Sub
        End If
        If ilWhichSort = 1 Then         'sort field 1
            ilLeft = 120
            ilTop = 300
        ElseIf ilWhichSort = 2 Then
            ilLeft = 2265
            ilTop = 300
            If ilCurrSortIndex > 0 Then
                ckcAllSort2.Visible = True
            End If
        ElseIf ilWhichSort = 3 Then
            ilLeft = 120
            ilTop = 2235
            If ilCurrSortIndex > 0 Then
                ckcAllSort3.Visible = True
            End If
        End If
        
        slSortCaption = ""
        If ilCurrSortIndex = 1 Then              'advt
            lbcSelection(0).Move ilLeft, ilTop
            lbcSelection(0).Visible = True
            slSortCaption = "All Advertisers"
        ElseIf ilCurrSortIndex = 2 Then          'agy
            lbcSelection(1).Move ilLeft, ilTop
            lbcSelection(1).Visible = True
            slSortCaption = "All Agencies"
        ElseIf ilCurrSortIndex = 3 Then          'Bus Category
            lbcSelection(2).Move ilLeft, ilTop
            lbcSelection(2).Visible = True
            slSortCaption = "All Bus Categories"
        ElseIf ilCurrSortIndex = 4 Then          'prod prot
            lbcSelection(3).Move ilLeft, ilTop
            lbcSelection(3).Visible = True
            slSortCaption = "All Prod Prot"
        ElseIf ilCurrSortIndex = 5 Then          'slsp
            lbcSelection(4).Move ilLeft, ilTop
            lbcSelection(4).Visible = True
            slSortCaption = "All Salespeople"
            ckcShowSlspSplit.Visible = True
            ckcShowSlspSplit.Value = vbChecked
        ElseIf ilCurrSortIndex = 6 Then          'vehicle
            lbcSelection(5).Move ilLeft, ilTop
            lbcSelection(5).Visible = True
            slSortCaption = "All Vehicles"
            cbcSortVG.Enabled = True
        Else                'if 0, it has to be either sort 2 or sort3 as
                            'the index for sort 1 has been incremented from calling rtn, and
                            'the index has to be non-zero
            
        End If
        
        If ilWhichSort = 1 Then
            ckcAllSort1.Caption = slSortCaption
        ElseIf ilWhichSort = 2 Then
            ckcAllSort2.Caption = slSortCaption
        ElseIf ilWhichSort = 3 Then
            ckcAllSort3.Caption = slSortCaption
        End If
        

End Sub
'           mIsVehicleSelected - determine if vehicle in any of the sorts were selected
'           if so, need to show the vehicle groups option
'           <input>   ilwhichsort - sort fields #1, 2 or 3
Public Sub mIsVehicleSelected(ilWhichSort As Integer)
'   8-23-11 remove the sort index being processed, test all sorts to see if any are by vehicle
'Public Sub mIsVehicleSelected(ilSortIndex As Integer, ilWhichSort As Integer)
    'If ilSortIndex = 6 Then
    If imSort1Index + 1 = 6 Or imSort2Index = 6 Or imSort3Index = 6 Then
        cbcSortVG.Visible = True
        'cbcSortVG.ListIndex = 0     'default to none
        lacSortVG.Visible = True
        lbcSelection(6).Visible = True
        ckcAllSortVG.Visible = True
        ckcSkipVG.Visible = True
        If imSort1Index + 1 = 6 Then
            plcGameSubTotal.Visible = True
        End If
    Else                'first sort is not by vehicle
        cbcSortVG.Visible = False
        lacSortVG.Visible = False
        ckcAllSortVG.Value = vbUnchecked
        ckcSkipVG.Visible = False
        ckcSkipVG.Value = vbUnchecked
        plcGameSubTotal.Visible = False
        rbcGameSubTotal(0).Value = True     'default to combine event subtotals for game vehicles
    End If
    Exit Sub
End Sub
