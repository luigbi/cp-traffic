VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelRR 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Research Revenue"
   ClientHeight    =   5535
   ClientLeft      =   180
   ClientTop       =   1485
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      Left            =   3360
      Top             =   4800
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
      Caption         =   "Research Revenue Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3930
      Left            =   45
      TabIndex        =   14
      Top             =   1545
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
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   4455
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
         Begin VB.PictureBox plcSelC12 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   -360
            ScaleHeight     =   360
            ScaleWidth      =   4380
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   3240
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Use Primary Only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   54
               Top             =   -30
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1110
               TabIndex        =   53
               Top             =   -30
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1695
               TabIndex        =   52
               Top             =   -15
               Visible         =   0   'False
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSelC6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1800
            ScaleHeight     =   240
            ScaleWidth      =   4620
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   4620
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "EST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   73
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "CST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1110
               TabIndex        =   72
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "MST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1695
               TabIndex        =   71
               Top             =   -390
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2295
               TabIndex        =   70
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   360
               TabIndex        =   69
               Top             =   480
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   1320
               TabIndex        =   68
               Top             =   480
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2520
               TabIndex        =   67
               Top             =   480
               Visible         =   0   'False
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSelC5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   0
            ScaleHeight     =   600
            ScaleWidth      =   4380
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2445
               TabIndex        =   65
               Top             =   195
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   1350
               TabIndex        =   64
               Top             =   195
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   465
               TabIndex        =   63
               Top             =   195
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2295
               TabIndex        =   62
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "MST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1695
               TabIndex        =   61
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "CST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1095
               TabIndex        =   60
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "EST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   59
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2580
               TabIndex        =   58
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   2865
               TabIndex        =   57
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   3120
               TabIndex        =   56
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   720
            TabIndex        =   21
            Top             =   360
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   450
            Text            =   "9/9/2019"
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
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   2475
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2760
            Width           =   2475
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   480
               TabIndex        =   50
               Top             =   0
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1440
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.PictureBox plcDemos 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   500
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   4260
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1800
            Width           =   4260
            Begin VB.CheckBox ckcDemo1 
               Caption         =   "Primary Demo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   31
               Top             =   0
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox ckcDemo234 
               Caption         =   "2nd, 3rd, 4th Demo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2400
               TabIndex        =   32
               Top             =   0
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.CheckBox ckcDemoBase 
               Caption         =   "Base Demo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   720
               TabIndex        =   33
               Top             =   240
               Value           =   1  'Checked
               Width           =   1575
            End
         End
         Begin VB.ComboBox cbcBook 
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
            Left            =   240
            TabIndex        =   41
            Top             =   3120
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.PictureBox plcBook 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   4275
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   2280
            Width           =   4275
            Begin VB.OptionButton rbcBook 
               Caption         =   "Schedule line book"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   480
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   240
               Width           =   2055
            End
            Begin VB.OptionButton rbcBook 
               Caption         =   "Default book"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2880
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   0
               Width           =   1335
            End
            Begin VB.OptionButton rbcBook 
               Caption         =   "Closest book to air dates"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   480
               TabIndex        =   35
               Top             =   0
               Value           =   -1  'True
               Width           =   2415
            End
         End
         Begin VB.PictureBox plcSortBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   120
            ScaleHeight     =   480
            ScaleWidth      =   4260
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   720
            Width           =   4260
            Begin VB.OptionButton rbcSortby 
               Caption         =   "Bus Category"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2640
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton rbcSortby 
               Caption         =   "Product Protection"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   720
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   240
               Width           =   1935
            End
            Begin VB.OptionButton rbcSortby 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   2040
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton rbcSortby 
               Caption         =   "Advertiser"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   1395
            End
         End
         Begin VB.TextBox edcSelCTo 
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
            Left            =   2910
            MaxLength       =   2
            TabIndex        =   23
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lbcActive 
            Appearance      =   0  'Flat
            Caption         =   "Active Contract Dates (35 days maximum)"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   3615
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Start"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   26
            Top             =   420
            Width           =   600
         End
         Begin VB.Label lacSelCTo 
            Appearance      =   0  'Flat
            Caption         =   "# Days"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2250
            TabIndex        =   22
            Top             =   420
            Width           =   645
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
         Height          =   4095
         Left            =   4605
         ScaleHeight     =   4095
         ScaleWidth      =   4455
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lbcMultiCntr 
            Appearance      =   0  'Flat
            Height          =   2970
            Left            =   0
            MultiSelect     =   2  'Extended
            TabIndex        =   43
            Top             =   720
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.ListBox lbcCntrCode 
            Appearance      =   0  'Flat
            Height          =   2970
            Left            =   0
            MultiSelect     =   2  'Extended
            TabIndex        =   44
            Top             =   600
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   1
            Left            =   2190
            MultiSelect     =   2  'Extended
            TabIndex        =   47
            Top             =   360
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.CheckBox ckcAllCntrs 
            Caption         =   "All Contracts"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2190
            TabIndex        =   46
            Top             =   0
            Width           =   2295
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   0
            Left            =   0
            MultiSelect     =   2  'Extended
            TabIndex        =   45
            Top             =   375
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   1935
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
      Top             =   150
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
         Width           =   1485
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   1035
      Width           =   360
   End
End
Attribute VB_Name = "RptSelRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptselrr.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelRR.Frm           Revenue Research
'
' Release: 5.1
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imSetAllCntr As Integer 'true=set all contr list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imAllClickedCntr As Integer
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smLogUserCode As String
Dim imBookInx As Integer        'index to book selected for specific book
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
Dim imTerminate As Integer

Dim smPaintCaption5 As String    'caption for panel
Dim smPaintCaption6 As String    'caption for panel
Dim smPaintCaption12 As String    'caption for panel
Private Sub cbcBook_Change()
    imBookInx = cbcBook.ListIndex
End Sub
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
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
        ckcAllCntrs.Visible = False
        lbcSelection(0).Visible = False
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
Private Sub ckcAllCntrs_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllCntrs.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAllCntr Then
        imAllClickedCntr = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClickedCntr = False
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
        If Not gOpenPrtJob("ResrchRv.rpt") Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenRR(imGenShiftKey, smLogUserCode)
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

        gCrResearchRev
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
    mSetCommands
End Sub

Private Sub CSI_CalFrom_Change()
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
'    mSetCommands
'End Sub
'Private Sub edcSelCFrom_GotFocus()
'    gCtrlGotFocus edcSelCFrom
'End Sub
Private Sub edcSelCTo_Change()
    mSetCommands
End Sub
Private Sub edcSelCTo_GotFocus()
    mSetCommands
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
    RptSelRR.Refresh
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
    'RptSelRR.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgMultiCntrCodeAD
    Erase tgClfAD
    Erase tgCffAD
    PECloseEngine
    
    Set RptSelRR = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSelection_Click(Index As Integer)
Dim slCntrStatus As String
Dim ilHOState  As Integer
Dim ilHowManyDefined As Integer
    'If Not imAllClicked Then
        If Index = 0 Then               'contr seletivity
            If Not imAllClickedCntr Then
                imSetAllCntr = False
                ckcAllCntrs.Value = vbUnchecked
                imSetAllCntr = True
            End If
        Else
            If Not imAllClicked Then
                Screen.MousePointer = vbHourglass
                slCntrStatus = "HO"              'default to holds and orders
                ilHOState = 3                   'in addition to sch holds & orders, show latest GN if applicable,
                                                'plus the revised orders turned proposals (WCI)
                ckcAll.Enabled = True
                imSetAll = False
                ckcAll.Value = vbUnchecked
                ckcAll.Visible = True
                imSetAll = True
                'imSetAllCntr = True
                ckcAllCntrs.Value = vbUnchecked
                ckcAllCntrs.Enabled = True
                ckcAllCntrs.Visible = True
                'imSetAllCntr = False
                lbcSelection(0).Visible = True
    
                mCntrPop slCntrStatus, ilHOState
    
    
                ilHowManyDefined = lbcSelection(0).ListCount
                If ilHowManyDefined = 1 Then
                    ckcAllCntrs.Value = vbChecked
                End If
            End If
        End If
   ' End If
    Screen.MousePointer = vbDefault
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mSetCommands
End Sub
Private Sub mAdvtPop(RptForm As Form, lbcSelection As Control)
'
'   mAdvtPop
'   Where:
'       RptForm as Form
'       lbcSelection as control
'
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(RptSelCt, lbcSelection, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(RptForm, lbcSelection, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptForm
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
Private Sub mCntrPop(slCntrStatus As String, ilHOState As Integer)
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
    Dim ilLoop As Integer
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
    llLen = 0
    ilErr = False
    lbcSelection(0).Clear
    lbcCntrCode.Clear
    For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
        If lbcSelection(1).Selected(ilLoop) Then
            sgMultiCntrCodeTag = ""             'init the date stamp so the box will be populated
            ReDim tgMultiCntrCodeAD(0 To 0) As SORTCODE
            lbcMultiCntr.Clear
            slNameCode = tgAdvertiser(ilLoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            'ilCurrent = 1   '0=Current; 1=All
            'ilFilter = Val(slCode)   'by contract #; -101=by advertiser
            'ilVehCode = -1  'All vehicles
            'ilRet = gPopCntrBox(RptSelCt, ilCntrType, ilFilter, ilCurrent, 0, ilVehCode, lbcMultiCntr, lbcMultiCntrCode, True, False, False, False)
            'slCntrType = ""                                 'all Types
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
            If tgUrf(0).sPSAType <> "H" Then
                slCntrType = slCntrType & "S"
            End If
            If tgUrf(0).sPromoType <> "H" Then
                slCntrType = slCntrType & "M"
            End If
            If slCntrType = "CVTRQSM" Then
                slCntrType = ""
            End If
            'ilShow = 1
            ilShow = 5                  'show # and advt name
            ilCurrent = 1
            ilAdfCode = Val(slCode)
            'load up list box with contracts with matching adv
            'ilRet = gPopCntrForAASBox(RptSelCt, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, lbcMultiCntrCode)
            ilRet = gPopCntrForAASBox(RptSelRR, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, tgMultiCntrCodeAD(), sgMultiCntrCodeTagAD)
            If ilRet <> CP_MSG_NOPOPREQ Then
                On Error GoTo mCntrPopErr
                gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", RptSelRR
                On Error GoTo 0
            End If
            For ilIndex = 0 To UBound(tgMultiCntrCodeAD) - 1 Step 1 'lbcMultiCntrCode.ListCount - 1 Step 1
                slName = Trim$(tgMultiCntrCodeAD(ilIndex).sKey)  'lbcMultiCntrCode.List(ilIndex)
                If Not gOkAddStrToListBox(slName, llLen, True) Then
                    ilErr = True
                    Exit For
                End If
                lbcCntrCode.AddItem slName  'lbcMultiCntrCode.List(ilIndex)
            Next ilIndex
            If ilErr Then
                Exit For
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To lbcCntrCode.ListCount - 1 Step 1
        slNameCode = lbcCntrCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 1, "|", slCode)
        llCntrNo = 99999999 - CLng(slCode)
        slShow = Trim$(str$(llCntrNo))
        ilRet = gParseItem(slName, 2, "|", slCode)
        ilRet = gParseItem(slCode, 1, "-", slRevNo)
        ilRevNo = 999 - CLng(slRevNo)
        ilRet = gParseItem(slCode, 2, "-", slRevNo)
        ilExtRevNo = 999 - CLng(slRevNo)
        ilRet = gParseItem(slName, 4, "|", slCode)
        ilVerNo = 999 - CLng(slCode)
        ilRet = gParseItem(slName, 5, "|", slCode)
        If (slCode = "W") Or (slCode = "C") Or (slCode = "I") Or (slCode = "D") Then
            If (ilRevNo > 0) Then
                slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
            Else
                slShow = slShow & " V" & Trim$(str$(ilVerNo))
            End If
        Else
            slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
        End If
        ilRet = gParseItem(slName, 6, "|", slCode)
        slShow = slShow & " " & slCode
        lbcSelection(0).AddItem Trim$(slShow)  'Add ID to list box
    Next ilLoop
    sgMultiCntrCodeTagAD = ""       '5-19-03 init so next time thru the contracts are populated when advt selected
    Screen.MousePointer = vbDefault
    Exit Sub
mCntrPopErr:
 On Error GoTo 0
    imTerminate = True
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

    RptSelRR.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imAllClickedCntr = False
    imSetAllCntr = True
    pbcSelC.Move 90, 255, 4515, 3360

    gCenterStdAlone RptSelRR
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
    Dim ilVefCode As Integer
    Dim ilShow As Integer
    Dim ilSort As Integer
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
    'lbcRptType.Clear
    'lbcRptType.AddItem smSelectedRptName

    Screen.MousePointer = vbHourglass
    lbcSelection(0).Clear
    lbcSelection(1).Clear
    mAdvtPop RptSelRR, lbcSelection(1)
    lbcSelection(0).Tag = ""
    lbcSelection(1).Tag = ""
    ilVefCode = 0
    ilSort = 0      '0=sort by book name, or 1= date then book name
    ilShow = 1      '0=Show book name only, 1=show book name & date
    ilRet = gPopBookNameBox(RptSelRR, 0, 0, ilVefCode, ilSort, ilShow, cbcBook, tgBookNameCode(), sgBookNameCodeTag)
    cbcBook.Visible = False         'use to retrieve the books for internal array
    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height + 60
    pbcSelC.Visible = True
    pbcOption.Visible = True
    lbcSelection(1).Visible = True                  'show advt name list box
    lbcSelection(1).Move 15, ckcAll.Top + ckcAll.Height + 30, 2135, 3000
    lbcSelection(0).Visible = True
    lbcSelection(0).Move lbcSelection(1).Width + 60, ckcAll.Top + ckcAll.Height + 30, 2135, 3000

    pbcOption.Visible = True
    pbcOption.Enabled = True

    'If lbcRptType.ListCount > 0 Then
    '    gFindMatch smSelectedRptName, 0, lbcRptType
    '    If gLastFound(lbcRptType) < 0 Then
    '        MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
    '        imTerminate = True
    '        Exit Sub
    '    End If
     '   lbcRptType.ListIndex = gLastFound(lbcRptType)
    'End If
    
    'Date: 4/2/2020 added all contract types
    'Standard, Reserved, Remnant, DR. Per Inquiry, PSA, Promo
    plcSortBy.Left = 120
    plcSelC5.Move plcSortBy.Left, plcSortBy.Top + plcSortBy.Height, 4260
    plcSelC5.Height = 700   '  440
    'smPaintCaption5 = "Include"
    'plcSelC5_Paint
    
    ckcSelC5(0).Move 730, ckcSelC5(7).Top + ckcSelC5(7).Height - 30, 1080
    ckcSelC5(0).Caption = "Standard"
    ckcSelC5(0).Value = vbChecked   'True
    ckcSelC5(0).Visible = True
    
    ckcSelC5(1).Move 1800, ckcSelC5(7).Top + ckcSelC5(7).Height - 30, 1200
    ckcSelC5(1).Caption = "Reserved"
    ckcSelC5(1).Value = vbChecked   'True
    ckcSelC5(1).Visible = True
    ckcSelC5(1).Enabled = True
    
    ckcSelC5(2).Move 3000, ckcSelC5(7).Top + ckcSelC5(7).Height - 30, 1080
    ckcSelC5(2).Caption = "Remnant"
    ckcSelC5(2).Value = vbChecked   'True
    ckcSelC5(2).Visible = True
    
    ckcSelC5(3).Move 730, ckcSelC5(0).Top + ckcSelC5(0).Height - 30, 600
    ckcSelC5(3).Caption = "DR"
    ckcSelC5(3).Value = vbChecked   'True
    ckcSelC5(3).Visible = True
    
    ckcSelC5(4).Move 1325, ckcSelC5(0).Top + ckcSelC5(0).Height - 30, 1320
    ckcSelC5(4).Caption = "Per Inquiry"
    ckcSelC5(4).Value = vbChecked   'True
    ckcSelC5(4).Visible = True
    
    ckcSelC5(5).Move 2580, ckcSelC5(0).Top + ckcSelC5(0).Height - 30, 720
    ckcSelC5(5).Caption = "PSA"
    ckcSelC5(5).Value = vbUnchecked 'False
    ckcSelC5(5).Visible = True  '9-12-02 vbChecked 'True
    
    ckcSelC5(6).Move 3300, ckcSelC5(0).Top + ckcSelC5(0).Height - 30, 900
    ckcSelC5(6).Caption = "Promo"
    ckcSelC5(6).Value = vbUnchecked 'False
    ckcSelC5(6).Visible = True
    
    'Date: 4/7/2020 added Hold and Orders to include/exclude
    ckcSelC5(7).Caption = "Hold"
    ckcSelC5(7).Visible = True
    ckcSelC5(7).Value = vbChecked   'True
    ckcSelC5(7).Move 730, -30, 1080
    
    ckcSelC5(8).Caption = "Orders"
    ckcSelC5(8).Visible = True
    ckcSelC5(8).Value = vbChecked   'True
    ckcSelC5(8).Move 1800, ckcSelC5(7).Top, 1080
    
    plcSelC5.Visible = True
    
    CSI_CalFrom.ZOrder 0
    
    plcDemos.Move 120, (plcSelC5.Top + plcSelC5.Height)  ' - 100
    plcBook.Move plcSelC5.Left, plcDemos.Top + plcDemos.Height + 10
    plcGrossNet.Move plcSelC5.Left, plcBook.Top + plcBook.Height + 10
    plcSelC5_Paint
    plcSelC5.Print "Include"
    plcSortBy.Print "Sort by"
    plcBook.Print "Use"
    plcGrossNet.Print "By"
    mSetCommands
    
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
End Sub

Private Sub plcSelC5_Paint()
    plcSelC5.Cls
    plcSelC5.CurrentX = 0
    plcSelC5.CurrentY = 0
'    plcSelC5.Print "Include"
End Sub

Private Sub plcSelC6_Paint()
    plcSelC6.Cls
    plcSelC6.CurrentX = 0
    plcSelC6.CurrentY = 0
    plcSelC6.Print smPaintCaption6
End Sub

Private Sub plcSelC12_Paint()
    plcSelC12.Cls
    plcSelC12.CurrentX = 0
    plcSelC12.CurrentY = 0
    plcSelC12.Print smPaintCaption12
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
    'gInitStdAlone RptSelRR, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Audience Delivery"
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
    ilEnable = False
    'If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
    '    'atleast one advertiser must be selected
    '    For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1      'vehicle entry must be selected
    '        If lbcSelection(1).Selected(ilLoop) Then
    '            ilEnable = True
    '            Exit For
    '        End If
    '    Next ilLoop
    'End If

    'at least one advt must be selected
    'If ilEnable And (Not ckcall) Then        'selective advt, see if contract has been selected


    'If All advertisers are selected, a date span is required.
    'If selective advt/contrs, a date is not required and all contracts selected will be processed.
    'if selective advt/contrs and a date is entered, only those contracts active for the entered span will be processed
    If ckcAll.Value = vbChecked Then                   'all advt/cntrs, date required
'        If edcSelCFrom.Text <> "" And edcSelCTo.Text <> "" Then
        If CSI_CalFrom.Text <> "" And edcSelCTo.Text <> "" Then     '12-16-19 change to use csi calendar control
            ilEnable = True
        End If
    Else                        'not all selected , seeif at least one advt selected
        ilEnable = False
        If CSI_CalFrom.Text <> "" And edcSelCTo.Text <> "" Then     '12-16-19 change to use csi calendar control
            'atleast one advertiser must be selected
            For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1      'vehicle entry must be selected
                If lbcSelection(1).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
            If ilEnable Then                        'now see if at least one contract has been selected
                ilEnable = False
                For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1      'vehicle entry must be selected
                    If lbcSelection(0).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
            End If
        End If
    End If
    'End If
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
    Unload RptSelRR
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub plcGrossNet_Paint()
    plcGrossNet.CurrentX = 0
    plcGrossNet.CurrentY = 0
    plcGrossNet.Print "By"
End Sub

Private Sub rbcBook_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcBook(Index).Value
    'End of coded added
    mSetCommands
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
    'Terminate False
End Sub
Private Sub plcBook_Paint()
    plcBook.CurrentX = 0
    plcBook.CurrentY = 0
    plcBook.Print "Use"
End Sub
Private Sub plcSortBy_Paint()
    plcSortBy.CurrentX = 0
    plcSortBy.CurrentY = 0
    plcSortBy.Print "Sort by"
End Sub
