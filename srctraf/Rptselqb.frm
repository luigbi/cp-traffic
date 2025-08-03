VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelQB 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5535
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
   ScaleHeight     =   5535
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   17
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
      TabIndex        =   23
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
      TabIndex        =   24
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4080
      Top             =   4680
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
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   13
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
         TabIndex        =   10
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
         TabIndex        =   12
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Vehicle Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3810
      Left            =   45
      TabIndex        =   14
      Top             =   1680
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
         Height          =   3480
         Left            =   60
         ScaleHeight     =   3480
         ScaleWidth      =   4530
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   4530
         Begin VB.OptionButton rbcVersion 
            Caption         =   "Spot plus Digital Lines"
            Height          =   210
            Index           =   1
            Left            =   1920
            TabIndex        =   69
            Top             =   80
            Width           =   2175
         End
         Begin VB.OptionButton rbcVersion 
            Caption         =   "Spot"
            Height          =   210
            Index           =   0
            Left            =   960
            TabIndex        =   68
            Top             =   80
            Value           =   -1  'True
            Width           =   735
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   960
            TabIndex        =   26
            Top             =   390
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   450
            Text            =   "7/15/23"
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
            CSI_ForceMondaySelectionOnly=   0   'False
            CSI_AllowBlankDate=   0   'False
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin VB.PictureBox plcGrossNet 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   60
            ScaleHeight     =   270
            ScaleWidth      =   3180
            TabIndex        =   64
            Top             =   3000
            Width           =   3180
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   900
               TabIndex        =   66
               Top             =   15
               Value           =   -1  'True
               Width           =   870
            End
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   2040
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   15
               Width           =   720
            End
         End
         Begin VB.CheckBox ckcSelC4 
            Caption         =   "Suppress Selling Vehicle"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   1920
            TabIndex        =   63
            Top             =   2190
            Width           =   2490
         End
         Begin VB.CheckBox ckcSelC3 
            Caption         =   "Suppress Rates"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   2190
            Width           =   1815
         End
         Begin VB.PictureBox plc30sOrUnits 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   60
            ScaleHeight     =   270
            ScaleWidth      =   3540
            TabIndex        =   49
            Top             =   2760
            Width           =   3540
            Begin VB.OptionButton rbc30sOrUnits 
               Caption         =   "Units"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2040
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   15
               Width           =   1560
            End
            Begin VB.OptionButton rbc30sOrUnits 
               Caption         =   "30"" Units"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   900
               TabIndex        =   54
               Top             =   15
               Value           =   -1  'True
               Width           =   1110
            End
         End
         Begin VB.PictureBox plcCntrFeed 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            ScaleHeight     =   240
            ScaleWidth      =   3330
            TabIndex        =   50
            Top             =   3225
            Width           =   3330
            Begin VB.CheckBox ckcCntrFeed 
               Caption         =   "Contract Spots"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   900
               TabIndex        =   51
               Top             =   0
               Value           =   1  'Checked
               Width           =   1605
            End
            Begin VB.CheckBox ckcCntrFeed 
               Caption         =   "Feed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2640
               TabIndex        =   52
               Top             =   0
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.PictureBox plcStart 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   60
            ScaleHeight     =   270
            ScaleWidth      =   3900
            TabIndex        =   46
            Top             =   2520
            Width           =   3900
            Begin VB.OptionButton rbcStart 
               Caption         =   "Standard Quarter"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   900
               TabIndex        =   47
               Top             =   15
               Value           =   -1  'True
               Width           =   1830
            End
            Begin VB.OptionButton rbcStart 
               Caption         =   "Start Date"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   2760
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   15
               Width           =   1320
            End
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   60
            ScaleHeight     =   975
            ScaleWidth      =   4425
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   750
            Width           =   4425
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   12
               Left            =   3600
               TabIndex        =   41
               Top             =   720
               Value           =   1  'Checked
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "N/C"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   11
               Left            =   2880
               TabIndex        =   40
               Top             =   720
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   1800
               TabIndex        =   39
               Top             =   720
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Trade"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   900
               TabIndex        =   38
               Top             =   720
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   3240
               TabIndex        =   37
               Top             =   480
               Width           =   990
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2520
               TabIndex        =   36
               Top             =   480
               Width           =   720
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "PI"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   1815
               TabIndex        =   35
               Top             =   465
               Value           =   1  'Checked
               Width           =   555
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   900
               TabIndex        =   34
               Top             =   480
               Value           =   1  'Checked
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   3240
               TabIndex        =   33
               Top             =   240
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2040
               TabIndex        =   32
               Top             =   240
               Value           =   1  'Checked
               Width           =   1125
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   900
               TabIndex        =   31
               Top             =   240
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1830
               TabIndex        =   30
               Top             =   0
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcSelC1 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   900
               TabIndex        =   29
               Top             =   0
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   60
            ScaleHeight     =   435
            ScaleWidth      =   3780
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1740
            Width           =   3780
            Begin VB.OptionButton rbcSelC2 
               Caption         =   "Exclude"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   900
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   210
               Width           =   1155
            End
            Begin VB.OptionButton rbcSelC2 
               Caption         =   "Show separately"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1860
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   45
               Width           =   1800
            End
            Begin VB.OptionButton rbcSelC2 
               Caption         =   "Hide"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   900
               TabIndex        =   42
               Top             =   15
               Value           =   -1  'True
               Width           =   675
            End
         End
         Begin VB.TextBox edcSelCFrom1 
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
            Left            =   3240
            MaxLength       =   3
            TabIndex        =   22
            Top             =   390
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label lblVersion 
            Caption         =   "Version"
            Height          =   255
            Left            =   0
            TabIndex        =   67
            Top             =   75
            Width           =   735
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   25
            Top             =   420
            Width           =   885
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# Weeks"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2400
            TabIndex        =   21
            Top             =   435
            Visible         =   0   'False
            Width           =   810
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
         Height          =   3420
         Left            =   4440
         ScaleHeight     =   3420
         ScaleWidth      =   4575
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CheckBox ckcAllAvails 
            Caption         =   "All Missed Avail Names"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2400
            TabIndex        =   61
            Top             =   0
            Width           =   2250
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   3
            ItemData        =   "Rptselqb.frx":0000
            Left            =   2400
            List            =   "Rptselqb.frx":0007
            MultiSelect     =   2  'Extended
            TabIndex        =   60
            Top             =   300
            Width           =   2115
         End
         Begin VB.CheckBox ckcAllSS 
            Caption         =   "All Sales Sources"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2400
            TabIndex        =   59
            Top             =   1860
            Width           =   2025
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1080
            Index           =   2
            ItemData        =   "Rptselqb.frx":000E
            Left            =   2400
            List            =   "Rptselqb.frx":0015
            MultiSelect     =   2  'Extended
            TabIndex        =   58
            Top             =   2160
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1080
            Index           =   1
            ItemData        =   "Rptselqb.frx":001C
            Left            =   240
            List            =   "Rptselqb.frx":0023
            TabIndex        =   57
            Top             =   2160
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   0
            ItemData        =   "Rptselqb.frx":002A
            Left            =   240
            List            =   "Rptselqb.frx":0031
            MultiSelect     =   2  'Extended
            TabIndex        =   56
            Top             =   300
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   0
            Width           =   1425
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Rate Cards"
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
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   1860
            Visible         =   0   'False
            Width           =   1005
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   18
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   16
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
Attribute VB_Name = "RptSelQB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselqb.frm on Wed 6/17/09 @ 12:56 P
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
' File Name: RptSelQB.Frm - Quarterly Booked Report
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
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imSetAllSS As Integer   'true to set All Sales Source
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imAllSSClicked As Integer   'All Sales Sources
Dim imSetAllAvails As Integer   '11-20-08
Dim imAllAvailsClicked As Integer   '11-20-08
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
'Vehicle link file- used to obtain start date
'Delivery file- used to obtain start date
'Vehicle conflict file- used to obtain start date
'Spot projection- used to obtain date status
'Library calendar file- used to obtain post log date status
'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
'Log
Dim imCodes() As Integer
Dim smLogUserCode As String
'Import contract report
'Spot week Dump
Dim imTerminate As Integer
'Dim tmSRec As LPOPREC
'Rate Card
Dim smRateCardTag As String
Dim smPlcCntrFeed As String

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

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllAvails_click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllAvails.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    ilValue = Value
    If imSetAllAvails Then
        imAllAvailsClicked = True
        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllAvailsClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAllSS_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllSS.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    ilValue = Value
    If imSetAllSS Then
        imAllSSClicked = True
        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllSSClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAllSS_GotFocus()
    gCtrlGotFocus ckcAllSS
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
        If rbcVersion(0).Value = True Then
            'Spot Version
            If Not gOpenPrtJob("QtrSpots.Rpt") Then
                igGenRpt = False
                frcOutput.Enabled = igOutput
                frcCopies.Enabled = igCopies
                'frcWhen.Enabled = igWhen
                frcFile.Enabled = igFile
                frcOption.Enabled = igOption
                'frcRptType.Enabled = igReportType
                Exit Sub
            End If
        End If
        If rbcVersion(1).Value = True Then
            'Spot + Digital Version - TTP 10729 - Quarterly Booked Spots report: add digital lines
            If Not gOpenPrtJob("QtrSptCb.Rpt") Then
                igGenRpt = False
                frcOutput.Enabled = igOutput
                frcCopies.Enabled = igCopies
                'frcWhen.Enabled = igWhen
                frcFile.Enabled = igFile
                frcOption.Enabled = igOption
                'frcRptType.Enabled = igReportType
                Exit Sub
            End If
        End If
        
        ilRet = gCmcGenQB(imGenShiftKey, smLogUserCode)
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

        gCRQtrlyBookSpots
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
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCRAvrClear
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

Private Sub CSI_CalFrom_CalendarChanged()
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer

    slDate = CSI_CalFrom.Text
    slDate = gObtainStartStd(slDate)
    llDate = gDateValue(slDate)

    'populate Rate Cards and bring in Rcf, Rif, and Rdf
    ilRet = gPopRateCardBox(RptSelQB, llDate, RptSelQB!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
    mSetCommands
End Sub

Private Sub CSI_CalFrom_Change()
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer

    slDate = CSI_CalFrom.Text
    slDate = gObtainStartStd(slDate)
    llDate = gDateValue(slDate)

    'populate Rate Cards and bring in Rcf, Rif, and Rdf
    ilRet = gPopRateCardBox(RptSelQB, llDate, RptSelQB!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
    mSetCommands
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
    If ((KeyAscii <> KEYBACKSPACE) And (KeyAscii <= 32)) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KeyDown) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    RptSelQB.Refresh
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
    'RptSelQB.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode, tgMNFCodeRpt
    Erase imCodes
    PECloseEngine
    
    Set RptSelQB = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Not imAllClicked Then
        If Index = 0 Then           'vehicle list box
            imSetAll = False
            ckcAll.Value = vbUnchecked  'False
            imSetAll = True
        End If
    End If
    If Not imAllSSClicked Then
        If Index = 2 Then           'sales source list box
            imSetAllSS = False
            ckcAllSS.Value = vbUnchecked  'False
            imSetAllSS = True
        End If
    End If
    If Not imAllAvailsClicked Then
        If Index = 3 Then           'named avails list box
            imSetAllAvails = False
            ckcAllAvails.Value = vbUnchecked  'False
            imSetAllAvails = True
        End If
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
    Dim illoop As Integer
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

    RptSelQB.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    illoop = InStr(slStr, "&")
    If illoop > 0 Then
        slStr = Left$(slStr, illoop - 1) & "&&" & Mid$(slStr, illoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False        'all vehicles
    imSetAll = True             'all vehicles
    imAllSSClicked = False      'all Sales sources
    imSetAllSS = True           'All Sales sources
    imAllAvailsClicked = False   '11-20-08
    imSetAllAvails = True
    If tgSpf.sSystemType = "R" Then     'radio vs network/syndicator
        ckcCntrFeed(0).Value = vbChecked
        ckcCntrFeed(1).Value = vbChecked
        plcCntrFeed.Visible = True
        ckcCntrFeed(1).Visible = True
        smPlcCntrFeed = "Include"
    Else
        ckcCntrFeed(0).Value = vbChecked
        ckcCntrFeed(1).Value = vbUnchecked
        ckcCntrFeed(0).Visible = False
        ckcCntrFeed(1).Visible = False
        smPlcCntrFeed = ""
    End If
    plcStart.Visible = True

    'List boxes are placed on form, no need to move them
    pbcSelC.Move 90, 255, 4515, 3360

    gCenterStdAlone RptSelQB
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
    gPopExportTypes cbcFileType     '10-20-01
    pbcSelC.Visible = False

    Screen.MousePointer = vbHourglass

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    mSellConvVirtVehPop 0, False
    'retrieve the Sales Sources
    ilRet = gPopMnfPlusFieldsBox(RptSelQB, RptSelQB!lbcSelection(2), tgMNFCodeRpt(), sgMNFCodeTagRpt, "S")
    ilRet = gAvailsPop(RptSelQB, lbcSelection(3), tgNamedAvail())       'show the named avails for selectivity
    ckcAllAvails.Value = vbChecked          'default to use all avails
    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True
    lacSelCFrom.Visible = True
'    edcSelCFrom.Visible = True
    ckcAll.Visible = True
    ckcAllSS.Visible = True
    'edcSelCFrom.Move 1350, edcSelCFrom.Top, 945
    lacSelCFrom.Visible = True
    lacSelCFrom1.Visible = True
    edcSelCFrom1.Visible = True
    lbcSelection(0).Visible = True                  'show budget name list box (base budget)
    lbcSelection(1).Visible = True                 'split budgets
    lbcSelection(2).Visible = True                  'sales sources
    lbcSelection(3).Visible = True
    laclbcName(0).Visible = True
    pbcOption.Visible = True
    pbcOption.Enabled = True
    
    mSetCommands
    Screen.MousePointer = vbDefault
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
    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
        
    ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
    If (ilRet = CP_MSG_NONE) Then
        ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
        ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
        igRptCallType = Val(slStr)      'Function ID (what function calling this report if )
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSellConvVirtVehPop             *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSellConvVirtVehPop(ilIndex As Integer, ilUselbcVehicle As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        ilRet = gPopUserVehicleBox(RptSelQB, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelQB, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelQB
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVirtVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim illoop As Integer

    ilEnable = True
    igRCSelectedIndex = 0
    'TTP 10729 - Quarterly Booked Spots report: add digital lines
    'If (CSI_CalFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
    If (CSI_CalFrom.Text <> "") And (edcSelCFrom1.Text <> "" Or rbcVersion(1).Value = True) Then
        'atleast one budget must be selected
        If ilEnable Then
            ilEnable = False
             For illoop = 0 To lbcSelection(1).ListCount - 1 Step 1      'get rate card selected
                If lbcSelection(1).Selected(illoop) Then
                    igRCSelectedIndex = illoop
                    ilEnable = True
                    Exit For
                End If
            Next illoop
            If lbcSelection(0).SelCount > 0 And lbcSelection(1).SelCount > 0 And lbcSelection(2).SelCount > 0 And lbcSelection(3).SelCount > 0 Then        'at least 1 vehicle selection
                ilEnable = True
            Else
                ilEnable = False
            End If
        End If
    Else
        ilEnable = False
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
    Unload RptSelQB
    igManUnload = NO
End Sub

Private Sub rbcVersion_Click(Index As Integer)
    'TTP 10729 - Quarterly Booked Spots report: add digital lines
    Select Case Index
        Case 0 'Spot
            lacSelCFrom1.Visible = True
            edcSelCFrom1.Visible = True
            ckcSelC4(0).Enabled = True
            rbcStart(0).Enabled = True
            rbcStart(1).Enabled = True
            
        Case 1 'Spot plus Digital
            lacSelCFrom1.Visible = False
            edcSelCFrom1.Visible = False 'The “# Weeks” field will not be available. This version of the report always runs for one broadcast month, which is 4 or 5 weeks.
            ckcSelC4(0).Enabled = False  'The “Suppress selling vehicle” checkbox will be grayed out, as the selling vehicle field is not shown on this new version of the report output.
            ckcSelC4(0).Value = vbUnchecked
            rbcStart(0).Enabled = False  'The “Standard Quarter” and “Start Date” option will be grayed out, as this new version of the report always runs for a maximum of five weeks and always starts from the start of the selected broadcast month.
            rbcStart(1).Enabled = False  'The “Standard Quarter” and “Start Date” option will be grayed out, as this new version of the report always runs for a maximum of five weeks and always starts from the start of the selected broadcast month.
            
    End Select
    mSetCommands
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plc30sOrUnits_Paint()
    plc30sOrUnits.CurrentX = 0
    plc30sOrUnits.CurrentY = 0
    plc30sOrUnits.Print "Show"

End Sub

Private Sub plcCntrFeed_Paint()
    plcCntrFeed.Cls
    plcCntrFeed.CurrentX = 0
    plcCntrFeed.CurrentY = 0
    plcCntrFeed.Print smPlcCntrFeed
End Sub

Private Sub plcGrossNet_Paint()
    plcGrossNet.CurrentX = 0
    plcGrossNet.CurrentY = 0
    plcGrossNet.Print "By"
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

Private Sub rbcSelC2_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC2(Index).Value
    'End of coded added
    If rbcSelC2(0).Value Then            'Hide reservations
        ckcSelC1(3).Value = vbChecked   'True           'disallow to be selected
        ckcSelC1(3).Enabled = False
    ElseIf rbcSelC2(1).Value Then               'show separately
        ckcSelC1(3).Value = vbChecked   'True
        ckcSelC1(3).Enabled = True
    Else                                 'exclude
        ckcSelC1(3).Value = vbUnchecked 'False
        ckcSelC1(3).Enabled = False
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub

Private Sub plcStart_Paint()
    plcStart.CurrentX = 0
    plcStart.CurrentY = 0
    plcStart.Print "Use"
End Sub

Private Sub plcSelC1_Paint()
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    plcSelC1.Print "Select"
End Sub

Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print "Reserved"
End Sub
