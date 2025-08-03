VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelDS 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Spot Selection"
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
      Left            =   6600
      TabIndex        =   68
      Top             =   585
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
      TabIndex        =   73
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
      TabIndex        =   74
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
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4260
      Top             =   4890
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
      Left            =   2055
      TabIndex        =   53
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
         TabIndex        =   55
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   56
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   54
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
      TabIndex        =   57
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   62
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
         TabIndex        =   59
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
         TabIndex        =   61
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   58
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   60
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Oversold Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4080
      Left            =   60
      TabIndex        =   63
      Top             =   1365
      Width           =   9180
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
         Height          =   3855
         Left            =   30
         ScaleHeight     =   3855
         ScaleWidth      =   4785
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   4785
         Begin V81TrafficReports.CSI_Calendar csi_CalTo 
            Height          =   285
            Left            =   3000
            TabIndex        =   1
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            Text            =   "6/17/2011"
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
         Begin V81TrafficReports.CSI_Calendar csi_CalFrom 
            Height          =   285
            Left            =   1200
            TabIndex        =   0
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            Text            =   "6/17/2011"
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
            CSI_DefaultDateType=   2
         End
         Begin VB.TextBox edcETime 
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
            Left            =   3345
            MaxLength       =   10
            TabIndex        =   11
            Text            =   "12M"
            Top             =   810
            Width           =   1170
         End
         Begin VB.TextBox edcSTime 
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
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "12M"
            Top             =   810
            Width           =   1170
         End
         Begin VB.PictureBox plcShow 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   120
            ScaleHeight     =   420
            ScaleWidth      =   4275
            TabIndex        =   83
            Top             =   3390
            Width           =   4275
            Begin VB.OptionButton rbcShow 
               Caption         =   "Advertiser"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1215
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   0
               Width           =   1185
            End
            Begin VB.OptionButton rbcShow 
               Caption         =   "All"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   585
               TabIndex        =   43
               Top             =   0
               Value           =   -1  'True
               Width           =   510
            End
            Begin VB.OptionButton rbcShow 
               Caption         =   "Agency"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2445
               TabIndex        =   45
               Top             =   0
               Width           =   1005
            End
            Begin VB.OptionButton rbcShow 
               Caption         =   "Prod Protection"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   585
               TabIndex        =   46
               Top             =   210
               Width           =   1590
            End
            Begin VB.OptionButton rbcShow 
               Caption         =   "Salesperson"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   2190
               TabIndex        =   47
               Top             =   210
               Width           =   1560
            End
         End
         Begin VB.PictureBox plcEvents 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   105
            ScaleHeight     =   240
            ScaleWidth      =   4740
            TabIndex        =   82
            Top             =   3150
            Width           =   4740
            Begin VB.CheckBox ckcEvents 
               Caption         =   "Programs"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   585
               TabIndex        =   40
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1140
            End
            Begin VB.CheckBox ckcEvents 
               Caption         =   "Comments"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1740
               TabIndex        =   41
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1230
            End
            Begin VB.CheckBox ckcEvents 
               Caption         =   "Open Avails only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2985
               TabIndex        =   42
               Top             =   -30
               Width           =   1695
            End
         End
         Begin VB.TextBox edcContr 
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
            Left            =   3330
            MaxLength       =   9
            TabIndex        =   14
            Top             =   1125
            Width           =   1170
         End
         Begin VB.PictureBox plcRanks 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4575
            TabIndex        =   80
            Top             =   2910
            Width           =   4575
            Begin VB.CheckBox ckcRank 
               Caption         =   "ROS"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   3525
               TabIndex        =   39
               Top             =   -30
               Value           =   1  'Checked
               Width           =   660
            End
            Begin VB.CheckBox ckcRank 
               Caption         =   "DP"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2940
               TabIndex        =   38
               Top             =   -30
               Value           =   1  'Checked
               Width           =   525
            End
            Begin VB.CheckBox ckcRank 
               Caption         =   "Sponsor"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1860
               TabIndex        =   37
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1035
            End
            Begin VB.CheckBox ckcRank 
               Caption         =   "Fixed Time"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   600
               TabIndex        =   36
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1215
            End
         End
         Begin VB.PictureBox plcSpots 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   4380
            TabIndex        =   79
            Top             =   2190
            Width           =   4380
            Begin VB.CheckBox ckcSpots 
               Caption         =   "MG"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1950
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   420
               Value           =   1  'Checked
               Width           =   675
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Charge"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1035
               TabIndex        =   27
               Top             =   -30
               Value           =   1  'Checked
               Width           =   930
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2010
               TabIndex        =   28
               Top             =   -30
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "ADU"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2715
               TabIndex        =   29
               Top             =   -30
               Value           =   1  'Checked
               Width           =   720
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Bonus"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   3465
               TabIndex        =   30
               Top             =   -30
               Value           =   1  'Checked
               Width           =   885
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "+ Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   840
               TabIndex        =   31
               Top             =   195
               Value           =   1  'Checked
               Width           =   735
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "-Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   1680
               TabIndex        =   32
               Top             =   195
               Value           =   1  'Checked
               Width           =   645
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "N/C"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2400
               TabIndex        =   33
               Top             =   195
               Value           =   1  'Checked
               Width           =   600
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Recapturable"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   3060
               TabIndex        =   34
               Top             =   195
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Spinoff"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   840
               TabIndex        =   35
               Top             =   420
               Value           =   1  'Checked
               Width           =   900
            End
         End
         Begin VB.TextBox edcLength 
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
            Index           =   1
            Left            =   1530
            MaxLength       =   3
            TabIndex        =   13
            Top             =   1140
            Width           =   495
         End
         Begin VB.TextBox edcLength 
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
            Index           =   0
            Left            =   945
            MaxLength       =   3
            TabIndex        =   12
            Top             =   1140
            Width           =   495
         End
         Begin VB.PictureBox plcDays 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   135
            ScaleHeight     =   510
            ScaleWidth      =   4590
            TabIndex        =   77
            Top             =   345
            Width           =   4590
            Begin VB.CheckBox ckcDays 
               Caption         =   "Skip to New Page Each Day"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   480
               TabIndex        =   9
               Top             =   225
               Width           =   2565
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Mo"
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   465
               TabIndex        =   2
               Top             =   -30
               Value           =   1  'Checked
               Width           =   560
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Tu"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1080
               TabIndex        =   3
               Top             =   -15
               Value           =   1  'Checked
               Width           =   525
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "We"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   2
               Left            =   1700
               TabIndex        =   4
               Top             =   -30
               Value           =   1  'Checked
               Width           =   555
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Th"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2310
               TabIndex        =   5
               Top             =   -30
               Value           =   1  'Checked
               Width           =   525
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Fr"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   2895
               TabIndex        =   6
               Top             =   -30
               Value           =   1  'Checked
               Width           =   480
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Sa"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   3435
               TabIndex        =   7
               Top             =   -30
               Value           =   1  'Checked
               Width           =   525
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Su"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   4020
               TabIndex        =   8
               Top             =   -30
               Value           =   1  'Checked
               Width           =   525
            End
         End
         Begin VB.PictureBox plcCTypes 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   4455
            TabIndex        =   76
            Top             =   1470
            Width           =   4455
            Begin VB.CheckBox ckcCType 
               Caption         =   "Trades"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   3360
               TabIndex        =   25
               Top             =   435
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   2385
               TabIndex        =   24
               Top             =   435
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   1695
               TabIndex        =   23
               Top             =   435
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Per Inquiry"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   480
               TabIndex        =   22
               Top             =   435
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   3840
               TabIndex        =   21
               Top             =   210
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   2760
               TabIndex        =   20
               Top             =   210
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   1560
               TabIndex        =   19
               Top             =   210
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   480
               TabIndex        =   18
               Top             =   210
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Feed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3240
               TabIndex        =   17
               Top             =   -30
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2280
               TabIndex        =   16
               Top             =   -30
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1440
               TabIndex        =   15
               Top             =   -30
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.Label lacETime 
            Appearance      =   0  'Flat
            Caption         =   "End Time"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2460
            TabIndex        =   89
            Top             =   855
            Width           =   840
         End
         Begin VB.Label lacSTime 
            Appearance      =   0  'Flat
            Caption         =   "Start Time"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
            TabIndex        =   88
            Top             =   855
            Width           =   990
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contract #"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2295
            TabIndex        =   81
            Top             =   1170
            Width           =   960
         End
         Begin VB.Label lacLength 
            Appearance      =   0  'Flat
            Caption         =   "Lengths"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
            TabIndex        =   78
            Top             =   1170
            Width           =   780
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Dates-Start"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   75
            Top             =   60
            Width           =   960
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "End"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2520
            TabIndex        =   72
            Top             =   60
            Width           =   450
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
         Height          =   3810
         Left            =   4755
         ScaleHeight     =   3810
         ScaleWidth      =   4335
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   150
         Width           =   4335
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   4
            ItemData        =   "Rptselds.frx":0000
            Left            =   105
            List            =   "Rptselds.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   87
            Top             =   2295
            Visible         =   0   'False
            Width           =   4200
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   3
            ItemData        =   "Rptselds.frx":0004
            Left            =   120
            List            =   "Rptselds.frx":0006
            MultiSelect     =   2  'Extended
            TabIndex        =   86
            Top             =   2310
            Visible         =   0   'False
            Width           =   4200
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   2
            ItemData        =   "Rptselds.frx":0008
            Left            =   120
            List            =   "Rptselds.frx":000A
            MultiSelect     =   2  'Extended
            TabIndex        =   85
            Top             =   2310
            Visible         =   0   'False
            Width           =   4200
         End
         Begin VB.CheckBox ckcAllOthers 
            Caption         =   "All Others"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   2010
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   1
            ItemData        =   "Rptselds.frx":000C
            Left            =   120
            List            =   "Rptselds.frx":000E
            MultiSelect     =   2  'Extended
            TabIndex        =   66
            Top             =   2295
            Visible         =   0   'False
            Width           =   4200
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   0
            ItemData        =   "Rptselds.frx":0010
            Left            =   120
            List            =   "Rptselds.frx":0012
            MultiSelect     =   2  'Extended
            TabIndex        =   65
            Top             =   300
            Width           =   4200
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   0
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6600
      TabIndex        =   69
      Top             =   1035
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   67
      Top             =   105
      Width           =   2805
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   75
      TabIndex        =   49
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   240
         Value           =   -1  'True
         Width           =   1410
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
Attribute VB_Name = "RptSelDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselds.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelDS.Frm - Daily Spot Report
'
'
' Release: 4.7  10/9/00
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
Dim imSetAllOther As Integer
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imAllOtherClicked As Integer
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
Private Sub ckcAllOthers_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllOthers.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim ilIndex As Integer
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllOther Then
        imAllOtherClicked = True
        If RptSelDS!rbcShow(1).Value Then   'advt selection
            ilIndex = 1
        ElseIf RptSelDS!rbcShow(2).Value Then   'agy selection
            ilIndex = 2
        ElseIf RptSelDS!rbcShow(3).Value Then   'product protection selection
            ilIndex = 3
        Else
            ilIndex = 4
        End If

        llRg = CLng(lbcSelection(ilIndex).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(ilIndex).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllOtherClicked = False
    End If
    mSetCommands
End Sub
Private Sub cmcBrowse_Click()
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
        If Not gOpenPrtJob("DailySpt.Rpt") Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenDS(imGenShiftKey, smLogUserCode)
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

        gCreateDS
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
    gCrCbfClear
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
Private Sub edcContr_GotFocus()
    gCtrlGotFocus edcContr
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
Private Sub edcETime_GotFocus()
    gCtrlGotFocus edcETime
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
Private Sub edcLength_GotFocus(Index As Integer)
    gCtrlGotFocus edcLength(Index)
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
'Private Sub edcSelCFrom_Change()
'    mSetCommands
'End Sub
'Private Sub edcSelCFrom_GotFocus()
'    gCtrlGotFocus edcSelCFrom
'End Sub
'Private Sub edcSelCFrom1_Change()
'    mSetCommands
'End Sub
'Private Sub edcSelCFrom1_GotFocus()
'    gCtrlGotFocus edcSelCFrom1
'End Sub
Private Sub edcSTime_GotFocus()
   gCtrlGotFocus edcSTime
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
    RptSelDS.Refresh
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
    'RptSelDS.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgRptNameCode
    Erase tgRptSalespersonCode
    Erase tgRptAgencyCode
    Erase tgRptAdvertiserCode
    Erase tgRptNameCode
    Erase imCodes
    PECloseEngine
    
    Set RptSelDS = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    If Index = 0 Then
        If Not imAllClicked Then
            'If index = 0 Then           'vehicle list box
                imSetAll = False
                ckcAll.Value = vbUnchecked  'False
                imSetAll = True
            'End If
        End If
    Else
        If Not imAllOtherClicked Then
            'If index = 0 Then           'vehicle list box
                imSetAllOther = False
                ckcAllOthers.Value = vbUnchecked    'False
                imSetAllOther = True
            'End If
        End If
    End If
    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop(lbcSelection As Control)
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required-
    ilRet = gPopAdvtBox(RptSelDS, lbcSelection, tgRptAdvertiserCode(), sgRptAdvertiserCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptSelDS
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
'*      Procedure Name:mAgencyPop                      *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Agency list box       *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAgencyPop(lbcSelection As Control)
'
'   mAgencyPop
'   Where:
'
    Dim ilRet As Integer
   ilRet = gPopAgyBox(RptSelDS, lbcSelection, tgRptAgencyCode(), sgRptAgencyCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gPopAgyBox)", RptSelDS
        On Error GoTo 0
    End If
    Exit Sub
mAgencyPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSel
    'Set RptSel = Nothing   'Remove data segment
    Exit Sub
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

    RptSelDS.Caption = smSelectedRptName '& " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imAllOtherClicked = False
    imSetAllOther = True
    ckcAll.Move 120, 0
    lbcSelection(0).Move 120, ckcAll.Height + 30, 4200, 3300

    'lacRC.Move ckcAll.Left, lbcSelection(0).Top + lbcSelection(0).Height + 30
    'lbcSelection(1).Move lbcSelection(0).Left, lacRC.Top + lacRC.Height + 30, 4200, 1380

    lbcSelection(0).Visible = True
    gCenterStdAlone RptSelDS
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
    gPopExportTypes cbcFileType     '10-20-01
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    pbcSelC.Visible = False

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    mSellConvVirtVehPop 0, False
    mSPersonPop lbcSelection(4)         'populate slsp list box
    mAdvtPop lbcSelection(1)            'populate advt list box
    mAgencyPop lbcSelection(2)          'populate agency list box
    mMnfPop "C", RptSelDS!lbcSelection(3), tgRptNameCode(), sgRptNameCodeTag    'Traffic!lbcSalesperson

    Screen.MousePointer = vbHourglass

    'Set the top parametrs of controls
'    edcSelCFrom.Move 1080, 30
'    edcSelCFrom1.Move 3240, 30
    plcDays.Move 120, 330
    edcSTime.Move 1125, 810
    edcETime.Move 3330, 810
    edcLength(0).Move 945, 1140
    edcLength(1).Move 1530, 1140
    plcCTypes.Move 120, 1470
    If tgSpf.sSystemType = "R" Then
        ckcCType(2).Visible = True
    End If
    plcSpots.Move 120, 2190
    plcRanks.Move 120, 2910
    plcEvents.Move 120, 3150
    plcShow.Move 120, 3390
    lbcSelection(0).Move 120, 300
    lbcSelection(1).Move 120, 2295
    lbcSelection(2).Move 120, 2295
    lbcSelection(3).Move 120, 2295
    lbcSelection(4).Move 120, 2295
    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True
    lacSelCFrom.Visible = True
'    edcSelCFrom.Visible = True
    ckcAll.Visible = True   '9-12-02 vbChecked    'True
    pbcOption.Visible = True
    pbcOption.Enabled = True
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
End Sub
'
'                   mMnfPop - Populate list box with MNF records
'                           slType = Mnf type to match (i.e. "H", "A")
'                           lbcLocal  - local list box to fill
'                           lbcMster - master list box with codes
'                   Created: DH 9/12/96
'
Private Sub mMnfPop(slType As String, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) 'lbcMster As Control)
ReDim ilfilter(0) As Integer
ReDim slFilter(0) As String
ReDim ilOffSet(0) As Integer
Dim ilRet As Integer
    ilfilter(0) = CHARFILTER
    slFilter(0) = slType
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType")

    'ilRet = gIMoveListBox(RptSelCt, lbcLocal, lbcMster, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(RptSelDS, lbcLocal, tlSortCode(), slSortCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMnfPopErr
        gCPErrorMsg ilRet, "mMnfPop (gImoveListBox)", RptSelDS
        On Error GoTo 0
    End If
    Exit Sub
mMnfPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    'gInitStdAlone RptSelDS, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Daily Spot"
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
        ilRet = gPopUserVehicleBox(RptSelDS, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelDS, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgVehicleTag)    'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelDS
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
    Dim ilLoop As Integer
    Dim ilIndex As Integer

    ilEnable = False
'    If (edcSelCFrom.Text <> "") Then
    If (CSI_CalFrom.Text <> "") Then
        ilEnable = True
        'atleast one vehicle must be selected
        If ilEnable Then
            ilEnable = False
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
            ilEnable = False
            'Check all other selections (advt, agy, product protection, slsp)
            If RptSelDS!rbcShow(0).Value Then        'show everything, no selection
                ilEnable = True
            Else
                If RptSelDS!rbcShow(1).Value Then   'advt selection
                    ilIndex = 1
                ElseIf RptSelDS!rbcShow(2).Value Then   'agy selection
                    ilIndex = 2
                ElseIf RptSelDS!rbcShow(3).Value Then   'product protection selection
                    ilIndex = 3
                Else
                    ilIndex = 4
                End If
                For ilLoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1      '
                    If lbcSelection(ilIndex).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
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
'*      Procedure Name:mSPersonPop                     *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Salesperson  list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSPersonPop(lbcSelection As Control)
'
'   mSPersonPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required-
    ilRet = gPopSalespersonBox(RptSelDS, 0, True, True, lbcSelection, tgRptSalespersonCode(), sgRptSalespersonCodeTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gPopSalespersonBox)", RptSelDS
        On Error GoTo 0
    End If
    Exit Sub
mSPersonPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Unload RptSelDS
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
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
Private Sub rbcShow_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcShow(Index).Value
    'End of coded added
    lbcSelection(1).Visible = False
    lbcSelection(2).Visible = False
    lbcSelection(3).Visible = False
    lbcSelection(4).Visible = False
    ckcAllOthers.Value = vbUnchecked
    If Index = 0 Then           'show all
          lbcSelection(0).Height = 3300
          ckcAllOthers.Visible = False
          ckcAllOthers.Value = vbChecked    'True

    ElseIf Index = 1 Then       'selective advertisrs
        lbcSelection(1).Visible = True
        lbcSelection(0).Height = 1650
        ckcAllOthers.Visible = True
        ckcAllOthers.Caption = "All Advertisers"
    ElseIf Index = 2 Then       'selective agencies
        lbcSelection(2).Visible = True
        lbcSelection(0).Height = 1650
        ckcAllOthers.Visible = True
        ckcAllOthers.Caption = "All Agencies"
    ElseIf Index = 3 Then       'selective product protection
        lbcSelection(3).Visible = True
        lbcSelection(0).Height = 1650
        ckcAllOthers.Visible = True
        ckcAllOthers.Caption = "All Product Protection"
    Else                        'selective salespeople
        lbcSelection(4).Visible = True
        lbcSelection(0).Height = 1650
        ckcAllOthers.Visible = True
        ckcAllOthers.Caption = "All Salespeople"
    End If
End Sub
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcShow_Paint()
    plcShow.CurrentX = 0
    plcShow.CurrentY = 0
    plcShow.Print "Show"
End Sub
Private Sub plcEvents_Paint()
    plcEvents.CurrentX = 0
    plcEvents.CurrentY = 0
    plcEvents.Print "Events"
End Sub
Private Sub plcRanks_Paint()
    plcRanks.CurrentX = 0
    plcRanks.CurrentY = 0
    plcRanks.Print "Ranks"
End Sub
Private Sub plcSpots_Paint()
    plcSpots.CurrentX = 0
    plcSpots.CurrentY = 0
    plcSpots.Print "Spot Types"
End Sub
Private Sub plcDays_Paint()
    plcDays.CurrentX = 0
    plcDays.CurrentY = 0
    plcDays.Print "Days"
End Sub
Private Sub plcCTypes_Paint()
    plcCTypes.CurrentX = 0
    plcCTypes.CurrentY = 0
    plcCTypes.Print "Contract Types"
End Sub
