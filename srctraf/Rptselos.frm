VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelOS 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Oversold Report Selection"
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
      TabIndex        =   65
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
      TabIndex        =   70
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
      TabIndex        =   71
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
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   3960
      Top             =   4920
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
      TabIndex        =   50
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
         TabIndex        =   52
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   53
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   51
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
      TabIndex        =   54
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   59
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
         TabIndex        =   56
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
         TabIndex        =   58
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   55
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   57
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Oversold Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4035
      Left            =   75
      TabIndex        =   60
      Top             =   1440
      Width           =   9210
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
         Height          =   3735
         Left            =   105
         ScaleHeight     =   3735
         ScaleWidth      =   4710
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   4710
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   1080
            TabIndex        =   83
            Top             =   0
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   450
            Text            =   "9/4/2019"
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
         Begin VB.ComboBox cbcGroup 
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
            Left            =   1470
            TabIndex        =   47
            Top             =   3360
            Width           =   1500
         End
         Begin VB.PictureBox plcTotals 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4500
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   3120
            Width           =   4500
            Begin VB.OptionButton rbcTotals 
               Caption         =   "Week"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1680
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   0
               Width           =   885
            End
            Begin VB.OptionButton rbcTotals 
               Caption         =   "Day"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   960
               TabIndex        =   45
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin VB.PictureBox plcSort 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4500
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   2880
            Width           =   4500
            Begin VB.OptionButton rbcSort 
               Caption         =   "Time within Week"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   480
               TabIndex        =   43
               Top             =   0
               Value           =   -1  'True
               Width           =   1800
            End
            Begin VB.OptionButton rbcSort 
               Caption         =   "Week within Time"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2400
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   0
               Width           =   1800
            End
         End
         Begin VB.PictureBox plcRanks 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4575
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   2385
            Width           =   4575
            Begin VB.CheckBox ckcRank 
               Caption         =   "ROS"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   3555
               TabIndex        =   40
               Top             =   -30
               Value           =   1  'Checked
               Width           =   645
            End
            Begin VB.CheckBox ckcRank 
               Caption         =   "DP"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2985
               TabIndex        =   39
               Top             =   -30
               Value           =   1  'Checked
               Width           =   510
            End
            Begin VB.CheckBox ckcRank 
               Caption         =   "Sponsor"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1890
               TabIndex        =   38
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
               TabIndex        =   37
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
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   1680
            Width           =   4380
            Begin VB.CheckBox ckcSpots 
               Caption         =   "MG"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   2985
               TabIndex        =   82
               Top             =   450
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   27
               Top             =   -30
               Value           =   1  'Checked
               Width           =   930
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Charge"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2040
               TabIndex        =   28
               Top             =   -30
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3000
               TabIndex        =   29
               Top             =   -30
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "ADU"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   3720
               TabIndex        =   30
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Bonus"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   480
               TabIndex        =   31
               Top             =   210
               Value           =   1  'Checked
               Width           =   870
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "+Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   1440
               TabIndex        =   32
               Top             =   210
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "-Fill"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2280
               TabIndex        =   33
               Top             =   210
               Value           =   1  'Checked
               Width           =   645
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "N/C"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   3000
               TabIndex        =   34
               Top             =   210
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Recapturable"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   480
               TabIndex        =   35
               Top             =   435
               Value           =   1  'Checked
               Width           =   1440
            End
            Begin VB.CheckBox ckcSpots 
               Caption         =   "Spinoff"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   1920
               TabIndex        =   36
               Top             =   435
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
            Index           =   3
            Left            =   3720
            MaxLength       =   3
            TabIndex        =   15
            Top             =   630
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
            Index           =   2
            Left            =   3120
            MaxLength       =   3
            TabIndex        =   14
            Top             =   630
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
            Index           =   1
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   13
            Text            =   "10"
            Top             =   630
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
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   12
            Text            =   "30"
            Top             =   630
            Width           =   495
         End
         Begin VB.PictureBox plcDays 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4545
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   360
            Width           =   4545
            Begin VB.CheckBox ckcDays 
               Caption         =   "Mo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   480
               TabIndex        =   5
               Top             =   -30
               Value           =   1  'Checked
               Width           =   555
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Tu"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1065
               TabIndex        =   6
               Top             =   -30
               Value           =   1  'Checked
               Width           =   555
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "We"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1605
               TabIndex        =   7
               Top             =   -30
               Value           =   1  'Checked
               Width           =   555
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Th"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2175
               TabIndex        =   8
               Top             =   -30
               Value           =   1  'Checked
               Width           =   555
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Fr"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   2760
               TabIndex        =   9
               Top             =   -30
               Value           =   1  'Checked
               Width           =   570
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Sa"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   3330
               TabIndex        =   10
               Top             =   -30
               Value           =   1  'Checked
               Width           =   540
            End
            Begin VB.CheckBox ckcDays 
               Caption         =   "Su"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   3930
               TabIndex        =   11
               Top             =   -30
               Value           =   1  'Checked
               Width           =   555
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
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   960
            Width           =   4455
            Begin VB.CheckBox ckcCType 
               Caption         =   "Trades"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   3360
               TabIndex        =   26
               Top             =   435
               Width           =   900
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   2400
               TabIndex        =   25
               Top             =   435
               Width           =   870
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   1695
               TabIndex        =   24
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
               TabIndex        =   23
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
               TabIndex        =   22
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
               TabIndex        =   21
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
               TabIndex        =   20
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
               TabIndex        =   19
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
               TabIndex        =   18
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2280
               TabIndex        =   17
               Top             =   -30
               Value           =   1  'Checked
               Width           =   8870
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1440
               TabIndex        =   16
               Top             =   -30
               Value           =   1  'Checked
               Width           =   840
            End
         End
         Begin VB.PictureBox plcShow 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4500
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   2640
            Width           =   4500
            Begin VB.OptionButton rbcShow 
               Caption         =   "Booked/Unsold"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2160
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   0
               Width           =   1800
            End
            Begin VB.OptionButton rbcShow 
               Caption         =   "Sold/Available"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   600
               TabIndex        =   41
               Top             =   0
               Value           =   -1  'True
               Width           =   1695
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
            MaxLength       =   2
            TabIndex        =   4
            Top             =   30
            Width           =   420
         End
         Begin VB.Label lacGroup 
            Appearance      =   0  'Flat
            Caption         =   "Vehicle Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   80
            Top             =   3420
            Width           =   1335
         End
         Begin VB.Label lacLength 
            Appearance      =   0  'Flat
            Caption         =   "Lengths to Highlight"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   76
            Top             =   690
            Width           =   1815
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   72
            Top             =   60
            Width           =   840
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# Weeks"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2400
            TabIndex        =   69
            Top             =   60
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
         Height          =   3825
         Left            =   4710
         ScaleHeight     =   3825
         ScaleWidth      =   4350
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   165
         Width           =   4350
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   1
            Left            =   120
            TabIndex        =   63
            Top             =   2055
            Width           =   4000
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   0
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   62
            Top             =   240
            Width           =   4000
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   -30
            Width           =   2475
         End
         Begin VB.Label lacRC 
            Appearance      =   0  'Flat
            Caption         =   "Rate Card"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   81
            Top             =   1800
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   66
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   64
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
         TabIndex        =   2
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
         TabIndex        =   1
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
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1065
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
Attribute VB_Name = "RptSelOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselos.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelOS.Frm - Oversold Report
'
'
' Release: 4.7  9/25/00
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
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
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
        If Not gOpenPrtJob("Oversold.Rpt") Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenOS(imGenShiftKey, smLogUserCode)
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

        gCreateOS
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

Private Sub CSI_CalFrom_CalendarChanged()
Dim sldate As String
Dim lldate As Long
Dim ilRet As Integer

    sldate = CSI_CalFrom.Text           'retrieve jan thru dec year
    sldate = gObtainStartStd(sldate)
    lldate = gDateValue(sldate)

    'populate Rate Cards and bring in Rcf, Rif, and Rdf
    ilRet = gPopRateCardBox(RptSelOS, lldate, RptSelOS!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
    mSetCommands
End Sub

Private Sub CSI_CalFrom_Change()
Dim sldate As String
Dim lldate As Long
Dim ilRet As Integer

    sldate = CSI_CalFrom.Text           'retrieve jan thru dec year
    sldate = gObtainStartStd(sldate)
    lldate = gDateValue(sldate)

    'populate Rate Cards and bring in Rcf, Rif, and Rdf
    ilRet = gPopRateCardBox(RptSelOS, lldate, RptSelOS!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
    mSetCommands

End Sub

Private Sub CSI_CalFrom_GotFocus()
    gCtrlGotFocus CSI_CalFrom
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
'Private Sub edcSelCFrom_Change()
'Dim sldate As String
'Dim lldate As Long
'Dim ilRet As Integer
'Dim ilLen As Integer
'    ilLen = Len(edcSelCFrom)
'    If ilLen >= 4 Then
'        sldate = edcSelCFrom           'retrieve jan thru dec year
'        sldate = gObtainStartStd(sldate)
'        lldate = gDateValue(sldate)
'
'        'populate Rate Cards and bring in Rcf, Rif, and Rdf
'        ilRet = gPopRateCardBox(RptSelOS, lldate, RptSelOS!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
'    End If
'    mSetCommands
'End Sub
'Private Sub edcSelCFrom_GotFocus()
'    gCtrlGotFocus edcSelCFrom
'End Sub
Private Sub edcSelCFrom1_Change()
    mSetCommands
End Sub
Private Sub edcSelCFrom1_GotFocus()
    gCtrlGotFocus edcSelCFrom1
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
    'RptSelOS.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase imCodes
    PECloseEngine
    
    Set RptSelOS = Nothing   'Remove data segment
    
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
            ckcAll.Value = vbUnchecked  '9-12-02 False
            imSetAll = True
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

    RptSelOS.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    ckcAll.Move 120, 0
    lbcSelection(0).Move 120, ckcAll.Height + 30, 4200, 1650

    lacRC.Move ckcAll.Left, lbcSelection(0).Top + lbcSelection(0).Height + 30
    lbcSelection(1).Move lbcSelection(0).Left, lacRC.Top + lacRC.Height + 30, 4200, 1380

    'pbcSelC.Move 90, 255, 4515, 3360
    lbcSelection(0).Visible = True
    lbcSelection(1).Visible = True
    gCenterStdAlone RptSelOS
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
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    gPopExportTypes cbcFileType     '10-20-01
    pbcSelC.Visible = False

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    mSellConvVirtVehPop 0, False
    gPopVehicleGroups RptSelOS!cbcGroup, tgVehicleSets1(), True

    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True
    lacSelCFrom.Visible = True
'    edcSelCFrom.Visible = True
    ckcAll.Visible = True
    pbcOption.Visible = True
    pbcOption.Enabled = True

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
    'gInitStdAlone RptSelOS, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Oversold Report"
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
        ilRet = gPopUserVehicleBox(RptSelOS, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelOS, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelOS
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

    ilEnable = False
'    If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
    If (CSI_CalFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
        ilEnable = True
        'atleast one budget must be selected
        If ilEnable Then
            ilEnable = False
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1      'budget entry must be selected
                If lbcSelection(0).Selected(ilLoop) Then
                    igBSelectedIndex = ilLoop
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
            ilEnable = False
            'Check Forecast selection
            For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1      'vehicle entry must be selected
                If lbcSelection(1).Selected(ilLoop) Then
                    igRCSelectedIndex = ilLoop
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
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
    Unload RptSelOS
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
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcTotals_Paint()
    plcTotals.CurrentX = 0
    plcTotals.CurrentY = 0
    plcTotals.Print "Totals by"
End Sub
Private Sub plcSort_Paint()
    plcSort.CurrentX = 0
    plcSort.CurrentY = 0
    plcSort.Print "Sort"
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
Private Sub plcShow_Paint()
    plcShow.CurrentX = 0
    plcShow.CurrentY = 0
    plcShow.Print "Show"
End Sub
