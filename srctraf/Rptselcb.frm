VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelCb 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facility Report Selection"
   ClientHeight    =   6480
   ClientLeft      =   2835
   ClientTop       =   2430
   ClientWidth     =   9915
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
   ScaleHeight     =   6480
   ScaleWidth      =   9915
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   18
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8145
      TabIndex        =   8
      Top             =   -15
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcMultiCntr 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox lbcLnCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcCntrCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   960
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   240
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
         ForeColor       =   &H00FFFF00&
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
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   14
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
         TabIndex        =   11
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
         TabIndex        =   13
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Vehicle Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4935
      Left            =   75
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   9570
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
         Height          =   4650
         Left            =   120
         ScaleHeight     =   4650
         ScaleMode       =   0  'User
         ScaleWidth      =   4695
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   165
         Visible         =   0   'False
         Width           =   4695
         Begin VB.PictureBox plcSelC15 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   3240
            ScaleHeight     =   975
            ScaleWidth      =   4575
            TabIndex        =   143
            Top             =   4320
            Visible         =   0   'False
            Width           =   4575
            Begin VB.CheckBox chkIncludeAdjustments 
               Caption         =   "Adjustments"
               Height          =   210
               Left            =   1010
               TabIndex        =   149
               Top             =   480
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.CheckBox ckcSelDigitalComments 
               Caption         =   "Digital Line Comments"
               Enabled         =   0   'False
               Height          =   210
               Left            =   1010
               TabIndex        =   147
               Top             =   720
               Width           =   2295
            End
            Begin VB.CheckBox ckcSelDigital 
               Caption         =   "Digital Lines"
               Height          =   210
               Left            =   720
               TabIndex        =   146
               Top             =   240
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.CheckBox ckcSelSpots 
               Caption         =   "Spots"
               Height          =   210
               Left            =   720
               TabIndex        =   145
               Top             =   0
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Include"
               Height          =   255
               Left            =   0
               TabIndex        =   144
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.PictureBox plcSelC11 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   139
            Top             =   2640
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC11 
               Caption         =   "30"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1560
               TabIndex        =   142
               Top             =   0
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.OptionButton rbcSelC11 
               Caption         =   "Units"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   141
               Top             =   0
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC11 
               Caption         =   "30/60"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2400
               TabIndex        =   140
               Top             =   15
               Visible         =   0   'False
               Width           =   1005
            End
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalTo2 
            Height          =   255
            Left            =   3120
            TabIndex        =   31
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "06/11/2024"
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
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom2 
            Height          =   255
            Left            =   960
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "06/11/2024"
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
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   1080
            TabIndex        =   23
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "06/11/2024"
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
         Begin VB.CheckBox ckcIncludeISCI 
            Caption         =   "Include ISCI/Creative Title"
            Height          =   210
            Left            =   1800
            TabIndex        =   124
            Top             =   4080
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.CheckBox ckcSelC15 
            Caption         =   "Summary Only"
            Height          =   210
            Left            =   240
            TabIndex        =   123
            Top             =   4080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.PictureBox plcSelC14 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   3675
            TabIndex        =   120
            Top             =   3840
            Visible         =   0   'False
            Width           =   3675
            Begin VB.OptionButton rbcSelC14 
               Caption         =   "Airing (Log)"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1920
               TabIndex        =   121
               Top             =   0
               Visible         =   0   'False
               Width           =   1365
            End
            Begin VB.OptionButton rbcSelC14 
               Caption         =   "Selling"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   720
               TabIndex        =   119
               Top             =   0
               Visible         =   0   'False
               Width           =   1065
            End
         End
         Begin VB.PictureBox plcSelC13 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   210
            ScaleHeight     =   240
            ScaleWidth      =   4260
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   3615
            Visible         =   0   'False
            Width           =   4260
            Begin VB.CheckBox ckcSelC13 
               Caption         =   "Weekly"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1695
               TabIndex        =   118
               Top             =   -15
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC13 
               Caption         =   "Std"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1110
               TabIndex        =   117
               Top             =   -30
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC13 
               Caption         =   "Cal"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   116
               Top             =   -30
               Visible         =   0   'False
               Width           =   1200
            End
         End
         Begin VB.ListBox lbcAgyAdvtCode 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   3210
            Sorted          =   -1  'True
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   3240
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.PictureBox plcSelC12 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   105
            ScaleHeight     =   240
            ScaleWidth      =   4260
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   3390
            Visible         =   0   'False
            Width           =   4260
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Local"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   111
               Top             =   -30
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Feed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1110
               TabIndex        =   113
               Top             =   -30
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1935
               TabIndex        =   114
               Top             =   -15
               Visible         =   0   'False
               Width           =   1020
            End
         End
         Begin VB.TextBox edcSet3 
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
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   126
            Top             =   4440
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.PictureBox plcSelC10 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   106
            Top             =   3120
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Use Primary Only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   107
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1200
               TabIndex        =   108
               Top             =   0
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2400
               TabIndex        =   109
               Top             =   0
               Visible         =   0   'False
               Width           =   1140
            End
         End
         Begin VB.TextBox edcSet2 
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
            Left            =   720
            TabIndex        =   45
            Text            =   "Minor Set #"
            Top             =   2880
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox edcSet1 
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
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Text            =   "Major Set #"
            Top             =   2880
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cbcSet2 
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
            Left            =   2880
            TabIndex        =   102
            Top             =   2880
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.ComboBox cbcSet1 
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
            Left            =   2040
            TabIndex        =   100
            Top             =   2880
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox plcSelC9 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   101
            Top             =   2400
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC9 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2040
               TabIndex        =   104
               Top             =   0
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.OptionButton rbcSelC9 
               Caption         =   "Corporate"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   103
               Top             =   0
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.OptionButton rbcSelC9 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3240
               TabIndex        =   105
               Top             =   0
               Visible         =   0   'False
               Width           =   1005
            End
         End
         Begin VB.PictureBox plcSelC8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4620
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   4620
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   3480
               TabIndex        =   61
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   3000
               TabIndex        =   60
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   2640
               TabIndex        =   59
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2160
               TabIndex        =   58
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1695
               TabIndex        =   57
               Top             =   -30
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1110
               TabIndex        =   56
               Top             =   -30
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "EST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   55
               Top             =   -30
               Visible         =   0   'False
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSelC7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4380
            TabIndex        =   96
            Top             =   2220
            Visible         =   0   'False
            Width           =   4380
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3240
               TabIndex        =   99
               Top             =   0
               Width           =   1005
            End
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1200
               TabIndex        =   97
               Top             =   0
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   2400
               TabIndex        =   98
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.PictureBox plcSelC6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4620
            TabIndex        =   77
            Top             =   1800
            Visible         =   0   'False
            Width           =   4620
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Hid"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   3960
               TabIndex        =   89
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Canc"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   3360
               TabIndex        =   88
               Top             =   0
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Miss"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   2520
               TabIndex        =   87
               Top             =   0
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   1920
               TabIndex        =   86
               Top             =   0
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "MST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1320
               TabIndex        =   85
               Top             =   0
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "CST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   84
               Top             =   0
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "EST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   83
               Top             =   0
               Visible         =   0   'False
               Width           =   720
            End
         End
         Begin VB.PictureBox plcSelC5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4500
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1440
            Visible         =   0   'False
            Width           =   4500
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "BB"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   3600
               TabIndex        =   76
               Top             =   0
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   3120
               TabIndex        =   75
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
               TabIndex        =   73
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
               TabIndex        =   71
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "EST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   66
               Top             =   0
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "CST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   67
               Top             =   0
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "MST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1695
               TabIndex        =   68
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2295
               TabIndex        =   69
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   465
               TabIndex        =   70
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
               TabIndex        =   72
               Top             =   195
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2445
               TabIndex        =   74
               Top             =   195
               Visible         =   0   'False
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSelC3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   -15
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   735
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2880
               TabIndex        =   42
               Top             =   -30
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   3330
               TabIndex        =   43
               Top             =   -30
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2445
               TabIndex        =   64
               Top             =   195
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   1350
               TabIndex        =   63
               Top             =   195
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   465
               TabIndex        =   62
               Top             =   195
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2295
               TabIndex        =   41
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "MST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1695
               TabIndex        =   40
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "CST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1095
               TabIndex        =   39
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "EST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   38
               Top             =   -45
               Visible         =   0   'False
               Width           =   600
            End
         End
         Begin VB.TextBox edcSelCTo1 
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
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   945
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
            Left            =   3360
            MaxLength       =   3
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   420
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
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   30
            Top             =   240
            Width           =   1170
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "incl4"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   3600
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "incl3"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2640
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "New/Changed"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   1785
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "All"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   91
               Top             =   0
               Value           =   -1  'True
               Width           =   510
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "New/Changed"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   840
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   0
               Width           =   840
            End
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   30
            ScaleHeight     =   240
            ScaleWidth      =   4620
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   975
            Visible         =   0   'False
            Width           =   4620
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Veh"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   3840
               TabIndex        =   53
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Veh"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   3210
               TabIndex        =   52
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Advt"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2640
               TabIndex        =   51
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Advt"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   600
               TabIndex        =   48
               Top             =   0
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Agy"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1290
               TabIndex        =   49
               Top             =   0
               Width           =   765
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Slsp"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   1920
               TabIndex        =   50
               Top             =   0
               Width           =   810
            End
         End
         Begin VB.TextBox edcSelCFrom 
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
            Left            =   360
            MaxLength       =   10
            TabIndex        =   24
            Top             =   0
            Width           =   1170
         End
         Begin VB.PictureBox plcSelC4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1650
               TabIndex        =   35
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   34
               Top             =   0
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2400
               TabIndex        =   36
               Top             =   0
               Width           =   1005
            End
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalTo 
            Height          =   255
            Left            =   3600
            TabIndex        =   26
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "06/11/2024"
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
         Begin VB.Label lacBarterVehiclesOnly 
            Caption         =   "*Only Barter Vehicles Will Be Included"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   4200
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contract #"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   240
            TabIndex        =   125
            Top             =   4440
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Active Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
            TabIndex        =   22
            Top             =   120
            Width           =   1365
         End
         Begin VB.Label lacSelCTo1 
            Appearance      =   0  'Flat
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2325
            TabIndex        =   46
            Top             =   375
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# of Days"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2280
            TabIndex        =   25
            Top             =   0
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lacSelCTo 
            Appearance      =   0  'Flat
            Caption         =   "Active End Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   28
            Top             =   270
            Width           =   1380
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
         Height          =   4650
         Left            =   4920
         ScaleHeight     =   4576.481
         ScaleMode       =   0  'User
         ScaleWidth      =   4455
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   165
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox CkcAllVeh 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   259
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.CheckBox ckcAllAAS 
            Caption         =   "All "
            ForeColor       =   &H80000008&
            Height          =   259
            Left            =   15
            TabIndex        =   128
            Top             =   -15
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   8
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   138
            Top             =   305
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   7
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   137
            Top             =   305
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   6
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   136
            Top             =   305
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   5
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   135
            Top             =   305
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   4
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   134
            Top             =   305
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   3
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   133
            Top             =   305
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   2
            Left            =   60
            MultiSelect     =   2  'Extended
            TabIndex        =   132
            Top             =   305
            Visible         =   0   'False
            Width           =   4350
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   1
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   131
            Top             =   305
            Visible         =   0   'False
            Width           =   4395
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   4020
            Index           =   0
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   130
            Top             =   305
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   259
            Left            =   255
            TabIndex        =   129
            Top             =   0
            Width           =   2145
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   19
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   17
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
         Caption         =   "Export"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   120
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1395
      End
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
         Width           =   1455
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
Attribute VB_Name = "RptSelCb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselcb.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmChfAdvtExt                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelCb.Frm  (duplicated from rptselct: to contain bridge reports 3-10-00)
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
'M for N Tracer record
Dim hmMtf As Integer        'Spot Detail
'Dim tmMtf As MTF
'Dim imMtfRecLen As Integer
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllAAS As Integer 'True=Set list box; False= don't change list box
Dim imAllClickedAAS As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllVeh As Integer
Dim imAllClickedVeh As Integer
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
'Spot projection- used to obtain date status
'Library calendar file- used to obtain post log date status
'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
Dim hmSpf As Integer            'Site file handle
Dim tmSpf As SPF                'SPF record image
Dim imSpfRecLen As Integer        'SPF record length
'Log
Dim imCodes() As Integer
Dim smLogUserCode As String
'Import contract report
'Spot week Dump
Dim imVefCode As Integer
Dim smVehName As String
Dim lmNoRecCreated As Long
Dim imTerminate As Integer
Dim ilAASCodes()  As Integer
Dim imHideProposalPrice As Integer
Dim smVehGp5CodeTag As String
Dim smMnfCodeTag As String
Dim smPlcSelC1P As String
Dim smPlcSelC2P As String
Dim smPlcSelC3P As String
Dim smPlcSelC4P As String
Dim smPlcSelC5P As String
Dim smPlcSelC6P As String
Dim smPlcSelC7P As String
Dim smPlcSelC8P As String
Dim smPlcSelC9P As String
Dim smPlcSelC10P As String
Dim smPlcSelC11P As String
Dim smPlcSelC12P As String
Dim smPlcSelC13P As String
Dim smPlcSelC14P As String
'
'           Selectivity for Hi-Lo Spot Rate Report
'
Public Sub mHiLoSelectivity()
    Dim slDate As String
    Dim llDate As Long
    Dim slThruDate As String
'   Reserve ckcSelC8 for days of the week; not shown but used in common routine mObtainSDF
    lbcSelection(6).Visible = True      'vehicles
    ckcAll.Caption = "All Vehicles"
    ckcAll.Visible = True

    lacSelCFrom.Caption = "Most recent date to include"
    lacSelCFrom.Move 120, 30, 2520
    lacSelCFrom.Visible = True
'            edcSelCFrom.Move 2520, 0, 960
    CSI_CalFrom.Move 2520, 0                    '9-11-19 use csi calendar control vs edit box
    slDate = Format$(gNow(), "m/d/yy")    'todays date
    llDate = gDateValue(slDate)
    llDate = llDate - 45                 'process 45 days total as default (from yesterday)
    slThruDate = Format(llDate, "m/d/yy")  'show the earliest date to include
    slDate = gDecOneDay(slDate)         'default to yesterdays date
'            edcSelCFrom.Text = slDate           'default to todays date
'           edcSelCFrom.Visible = True
    CSI_CalFrom.Text = slDate
    CSI_CalFrom.Visible = True
    lacSelCTo.Caption = "# Days to Include"
'            lacSelCTo.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30, 1440
    lacSelCTo.Move 120, CSI_CalFrom.Top + CSI_CalFrom.Height + 30, 1440
    lacSelCTo1.Caption = " thru " & Trim$(slThruDate)
    lacSelCTo1.Move 2280, lacSelCTo.Top, 1560
    lacSelCTo1.Visible = True
    edcSelCTo.Move 1680, lacSelCTo.Top - 30, 480
    edcSelCTo.MaxLength = 3
    edcSelCTo.Text = "45"
    lacSelCTo.Visible = True
    edcSelCTo.Visible = True
    
    rbcSelC4(1).Value = False
    rbcSelC4(0).Visible = True
    rbcSelC4(1).Visible = True
    rbcSelC4(2).Visible = False
    smPlcSelC4P = "By"
    rbcSelC4(0).Caption = "Detail"
    rbcSelC4(1).Caption = "Summary"
    rbcSelC4(0).Left = 360
    rbcSelC4(1).Left = 1200
    rbcSelC4(1).Width = 1160
    If rbcSelC4(0).Value = vbChecked Then
        rbcSelC4_click 0
    Else
        rbcSelC4(0).Value = vbChecked   'True
    End If
    plcSelC4.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
    plcSelC4.Visible = True
    
    mAskContractTypesCkcSelC3 plcSelC4.Top + plcSelC4.Height
    'plcSelC3.Move 120, plcSelC4.Top + plcSelC4.Height, 4260, 660
    plcSelC3_Paint
    
    '12-28-17 make selection subroutine
    plcSelC10.Move 120, plcSelC3.Top + plcSelC3.Height + 30, 3120
    ckcSelC10(0).Caption = "Use Daypart with Overrides"
    ckcSelC10(0).Move 0, 0, 3120
    ckcSelC10(0).Visible = True
    plcSelC10.Visible = True
    
    'vehicle group selection
    gPopVehicleGroups RptSelCb!cbcSet1, tgVehicleSets1(), True
    lacSelCFrom1.Caption = "Vehicle Group"
    lacSelCFrom1.Move 120, plcSelC10.Top + plcSelC10.Height + 60, 1680
    cbcSet1.Move 1440, lacSelCFrom1.Top - 15, 1500
    lacSelCFrom1.Visible = True
    cbcSet1.Visible = True
    
    lacContract.Caption = "Contract #"
    lacContract.Move 120, cbcSet1.Top + cbcSet1.Height + 60, 975
    lacContract.Visible = True
    edcSet3.Move 1080, cbcSet1.Height + cbcSet1.Top + 30, 960
    edcSet3.Visible = True
    edcSet3.MaxLength = 8

    pbcSelC.Visible = True
    Exit Sub
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

Private Sub cbcSet1_Click()
    Dim ilListIndex As Integer
    Dim ilSetIndex As Integer
    Dim illoop As Integer
    Dim ilRet As Integer
    ilSetIndex = cbcSet1.ListIndex
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            If ilListIndex = CNT_ACCRUEDEFER Then
                illoop = cbcSet1.ListIndex
                ilSetIndex = gFindVehGroupInx(illoop, tgVehicleSets1())
                If ilSetIndex > 0 Then
                    smVehGp5CodeTag = ""
                    ilRet = gPopMnfPlusFieldsBox(RptSelCb, lbcSelection(7), tgMNFCodeRpt(), sgMNFCodeTagRpt, "H" & Trim$(str$(ilSetIndex)))
                    lbcSelection(7).Visible = True
                    If ilSetIndex = 1 Then              'participants vehicle sets
                        CkcAllveh.Caption = "All Participants"
                    ElseIf ilSetIndex = 2 Then          'subtotals vehicle sets
                        CkcAllveh.Caption = "All Sub-totals"
                    ElseIf ilSetIndex = 3 Then          'market vehicle sets
                        CkcAllveh.Caption = "All Markets"
                    ElseIf ilSetIndex = 4 Then          'format vehicle sets
                        CkcAllveh.Caption = "All Formats"
                    ElseIf ilSetIndex = 5 Then          'research vehicle sets
                       CkcAllveh.Caption = "All Research"
                    ElseIf ilSetIndex = 6 Then          'sub-company vehicle sets
                        CkcAllveh.Caption = "All Sub-Companies"
                    End If
                    CkcAllveh.Visible = True
                Else
                    lbcSelection(7).Visible = False
                    CkcAllveh.Value = vbUnchecked   '9-12-02 False
                    CkcAllveh.Visible = False
                End If

            End If

    End Select
    mSetCommands
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
    Dim ilIndex As Integer
    Dim ilListIndex As Integer
    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAll Then
        imAllClicked = True

        If igRptCallType = CONTRACTSJOB Then
            ilListIndex = lbcRptType.ListIndex
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            Select Case ilListIndex
                Case CNT_SPTSBYADVT, CNT_MGREVENUE, CNT_SPTCOMBO 'spots by advt
                    If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_SPTCOMBO Then
                        plcSelC12.Visible = True
                        ckcSelC12(0).Value = vbChecked          'default to contracts spots on
                        ckcSelC12(1).Value = vbChecked          'default to feed spots on
                    Else
                        plcSelC12.Visible = False
                    End If
                    If rbcSelCSelect(0).Value Then 'Advt
                        If Value Then
                            lbcSelection(0).Visible = False
                            lbcSelection(5).Visible = False
                        Else
                            lbcSelection(0).Visible = True
                            lbcSelection(5).Visible = True
                        End If
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        
                        'TTP 10674: Select / Deselect all Advertisers
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    Else
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                    
                Case CNT_SPTSBYDATETIME         '10-27-15       CNT_MISSED  'Spots by times; Missed Spots
                    If rbcSelC14(0).Value = True Then
                        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, CLng(ilValue), llRg)
                        llRet = llRet
                    Else
                        llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, CLng(ilValue), llRg)
                    End If
                
                Case CNT_MISSED             '10-27-15 do not combine with Spots by Date & Time as Missed doesn't have the Selling/Airing vehicle option
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, CLng(ilValue), llRg)
                
                Case 5  'Recap
                
                Case CNT_PLACEMENT, CNT_DISCREP  'Placement; Discrepancy
                    If Value Then
                        lbcSelection(0).Visible = False
                        lbcSelection(5).Visible = False
                        If ilListIndex = 7 Then
                            'If lgOrigCntrNo > 0 Then
                                lacSelCTo.Visible = True
                                edcSelCTo.Visible = True
                            'Else
                            '    lacSelCTo.Visible = False
                            '    edcSelCTo.Visible = False
                            'End If
                        End If
                    Else
                        lbcSelection(0).Visible = True
                        lbcSelection(5).Visible = True
                        If ilListIndex = 7 Then
                            lacSelCTo.Visible = False
                            edcSelCTo.Visible = False
                        End If
                    End If
                
                Case CNT_MG  'MG
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                
                Case 9  'Sales Spot Tracking
                
                Case 10, 12 'Commercial Change, Affiliate Spot Tracking
                
                Case CNT_HISTORY 'History
                    llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    If Value Then
                        lbcSelection(0).Visible = False
                        lbcSelection(5).Visible = False
                    Else
                        lbcSelection(0).Visible = True
                        lbcSelection(5).Visible = True
                    End If
                
                Case CNT_SPOTSALES 'Spot Sales
                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)

                Case CNT_ACCRUEDEFER
                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                
                Case CNT_HILORATE
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            End Select
        End If
    'Else
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllAAS_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllAAS.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim llRg As Long
    Dim llRet As Long
    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAllAAS Then
        imAllClickedAAS = True
        If igRptCallType = CONTRACTSJOB Then
            If (igRptType = 0) And (ilIndex > 1) Then
                ilIndex = ilIndex + 1
            End If
            If ilIndex = CNT_SPTSBYADVT Or ilIndex = CNT_MGREVENUE Or ilIndex = CNT_SPTCOMBO Then     '6-16-00
                llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            
            ElseIf ilIndex = CNT_MISSED Then
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            
            ElseIf ilIndex = CNT_ACCRUEDEFER Then
                llRg = CLng(lbcSelection(4).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(4).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        End If
     imAllClickedAAS = False
    End If
    mSetCommands
End Sub

Private Sub CkcAllVeh_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim ilValue As Integer
    Dim ilIndex As Integer

    Value = False
    If CkcAllveh.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAllVeh Then
        imAllClickedVeh = True
        If igRptCallType = CONTRACTSJOB Then
            If ilIndex = CNT_ACCRUEDEFER Then
                'deselect/select all vehicle group items
                llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        End If
        imAllClickedVeh = False
    End If
    mSetCommands
End Sub

Private Sub ckcSelC10_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC10(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If

    End Select
    mSetCommands
End Sub

Private Sub ckcSelC3_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC3(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
    End Select
    mSetCommands
End Sub

Private Sub ckcSelC5_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC5(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If

    End Select
    mSetCommands
End Sub

Private Sub ckcSelC6_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC6(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
    End Select
    mSetCommands
End Sub

Private Sub ckcSelC8_Click(Index As Integer)
 'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC8(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            
            If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_SPTCOMBO Then
                If Value Then
                    RptSelCb!lacBarterVehiclesOnly.Move 120, RptSelCb!cbcSet1.Top + cbcSet1.Height + 240
                    RptSelCb!lacBarterVehiclesOnly.Visible = True
                Else
                    RptSelCb!lacBarterVehiclesOnly.Visible = False
                End If
            End If
    End Select
    mSetCommands
End Sub

Private Sub ckcSelDigital_Click()
    If ckcSelDigital.Value = vbChecked Then
        chkIncludeAdjustments.Enabled = True
    Else
        chkIncludeAdjustments.Enabled = False
    End If

    If ckcSelDigital.Value = vbChecked And rbcOutput(3).Value = True Then
        ckcSelDigitalComments.Enabled = True
    Else
        ckcSelDigitalComments.Enabled = False
        ckcSelDigitalComments.Value = vbUnchecked
    End If
    mSetCommands
End Sub

Private Sub ckcSelSpots_Click()
    Dim illoop As Integer
    'Disable/Enable all the Spot Include checkboxes
    For illoop = 0 To 10
        ckcSelC5(illoop).Enabled = ckcSelSpots.Value = vbChecked
    Next illoop
    
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
    mTerminate False
End Sub

Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub

Private Sub cmcGen_Click()
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilListIndex As Integer
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    
    If igGenRpt Then
        Exit Sub
    End If
    
    'TTP 10674: Prevent BAD Date Range
    If CSI_CalTo.Text <> "" And CSI_CalTo.Visible Then
        If IsDate(CSI_CalTo.Text) = False Then
            CSI_CalTo.SetFocus
            Beep
            Exit Sub
        End If
    End If
    If CSI_CalFrom.Text <> "" And CSI_CalFrom.Visible Then
        If IsDate(CSI_CalFrom.Text) = False Then
            CSI_CalFrom.SetFocus
            Beep
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    
    igGenRpt = True
    igOutput = frcOutput.Enabled
    igCopies = frcCopies.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False
    ilListIndex = lbcRptType.ListIndex
    
    'LOGSJOB: igRptType = 0 or 2 => Log format; 1 or 3 => Delivery
    'If (igRptCallType = LOGSJOB) And ((igRptType = 0) Or (igRptType = 2)) And ((ilListIndex = 1) Or (ilListIndex = 3)) Then
    If (igRptCallType = CONTRACTSJOB) Then
        If (igRptType = 0) And (ilListIndex > 1) Then
            ilListIndex = ilListIndex + 1
        End If
        If ilListIndex = 0 And rbcSelCInclude(2).Value Then 'Contract Report
            igUsingCrystal = False
        ElseIf ilListIndex = 10 Then 'Commercial Changes
            igUsingCrystal = False
        ElseIf ilListIndex = 11 Then 'Contract History
            igUsingCrystal = False
        ElseIf ilListIndex = 12 Then 'Tracking for affiliate
            igUsingCrystal = False
        Else
            igUsingCrystal = True
        End If
    Else
        igUsingCrystal = True
    End If
    If (igRptCallType = CONTRACTSJOB) Then
        ilNoJobs = 1
        ilStartJobNo = 1
    Else
        ilNoJobs = 1
        ilStartJobNo = 1
    End If
    
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportCb() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        ilRet = gCmcGenCb(ilListIndex, imGenShiftKey, smLogUserCode)
        '-1 is a Crystal failure of gSetSelection or gSEtFormula
        If ilRet = -1 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf ilRet = 0 Then   '0 = invalid input data, stay in
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf ilRet = 2 Then           'successful return from bridge reports
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        '1 falls thru - successful crystal report
        'If contract spot projection or quarterly avails- create records
        If (igRptCallType = CONTRACTSJOB) Then
            If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_MGREVENUE Or ilListIndex = CNT_SPTCOMBO Then     'spots by advt (agy orslsp)
                Screen.MousePointer = vbHourglass
                gSpotAdvtRpt
                Screen.MousePointer = vbDefault
            
            ElseIf ilListIndex = CNT_SPTSBYDATETIME Then    'Spots by date and time
                Screen.MousePointer = vbHourglass
                gSpotDateRpt 3
                Screen.MousePointer = vbDefault
            
            ElseIf ilListIndex = CNT_MISSED Then    'Spots by date and time
                Screen.MousePointer = vbHourglass
                gSpotDateRpt 2
                Screen.MousePointer = vbDefault
            
            ElseIf ilListIndex = CNT_ACCRUEDEFER Then
                Screen.MousePointer = vbHourglass
                gGenAccrueDefer
                Screen.MousePointer = vbDefault
            
            ElseIf ilListIndex = CNT_HILORATE Then
                Screen.MousePointer = vbHourglass
                gGenHiLoRate
                Screen.MousePointer = vbDefault
            
            ElseIf ilListIndex = CNT_DISCREP_SUM Then           '6-22-16
                Screen.MousePointer = vbHourglass
                gSpotDispSumRpt
                Screen.MousePointer = vbDefault
            End If
        End If
        
        Screen.MousePointer = vbDefault
        
        If rbcOutput(0).Value Then
            DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
            igDestination = 0
            Report.Show vbModal
        
        ElseIf rbcOutput(1).Value Then
            ilCopies = Val(edcCopies.Text)
            ilRet = gOutputToPrinter(ilCopies)
        
        Else
            '-------------------------------
            'TTP 10674 - Spot and Digital Line combo Export or report?
            If RptSelCb.rbcOutput(3) Then
                'Don't Open Crystal, we are Exporing to CSV
                MsgBox "Export Complete" & vbCrLf & "Export Stored in- " & sgExportPath & slFileName, vbInformation, "Spot and Digital Line Combo - Export to CSV"
            Else
                slFileName = edcFileName.Text
                ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
            End If
        End If
    Next ilJobs
    imGenShiftKey = 0

    If (igRptCallType = CONTRACTSJOB) Then
        If (ilListIndex = CNT_SPOTSALES) Then     'Spot Sales by Date and Adv
            Screen.MousePointer = vbHourglass
            If RptSelCb!rbcSelCSelect(0).Value Or RptSelCb!rbcSelCSelect(1).Value Then
                gCRGrfClear
            Else
                gCrCbfClear
            End If
            Screen.MousePointer = vbDefault
        End If
        'DS
        If (ilListIndex = CNT_PLACEMENT) Or (ilListIndex = CNT_DISCREP) Or (ilListIndex = CNT_DISCREP_SUM) Then     'Spot Placement and Spot Discrepancy & spot discr summary by month
            Screen.MousePointer = vbHourglass
            gCrCbfClear
            Screen.MousePointer = vbDefault
        End If
        If ilListIndex = CNT_SPTSBYDATETIME Or ilListIndex = CNT_MISSED Or ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_MGREVENUE Or ilListIndex = CNT_ACCRUEDEFER Or ilListIndex = CNT_HILORATE Or ilListIndex = CNT_SPTCOMBO Then
            Screen.MousePointer = vbHourglass
            gCRGrfClear
            Screen.MousePointer = vbDefault
        End If
    End If
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

Private Sub CSI_CalFrom_GotFocus()
    CSI_CalFrom.ZOrder vbBringToFront
End Sub

Private Sub CSI_CalFrom2_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalFrom2_GotFocus()
    CSI_CalFrom2.ZOrder vbBringToFront
End Sub

Private Sub CSI_CalTo_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalTo_GotFocus()
    CSI_CalTo.ZOrder vbBringToFront
End Sub

Private Sub CSI_CalTo2_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalTo2_GotFocus()
    CSI_CalTo2.ZOrder vbBringToFront
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

Private Sub edcSelCFrom_Change()
    Dim ilLen As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim slThruDate As String
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            ElseIf ilListIndex = CNT_HILORATE Then
                ilLen = Len(edcSelCFrom.Text)
                If ilLen >= 3 Then
                    slDate = edcSelCFrom.Text          '
                    llDate = gDateValue(slDate)
                    slDate = Format(llDate, "m/d/yy")
                    llDate = llDate - (Val(edcSelCTo.Text)) + 1
                    lacSelCTo1.Caption = "thru " & Format$(llDate, "m/d/yy")
                End If
            End If

    End Select
    mSetCommands
End Sub

Private Sub edcSelCFrom_GotFocus()
    gCtrlGotFocus edcSelCFrom
End Sub

Private Sub edcSelCFrom_KeyPress(KeyAscii As Integer)
    Exit Sub
End Sub

Private Sub edcSelCFrom1_Change()
    mSetCommands
End Sub

Private Sub edcSelCFrom1_GotFocus()
    gCtrlGotFocus edcSelCFrom1
End Sub

Private Sub edcSelCFrom1_KeyPress(KeyAscii As Integer)
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If

    End Select
End Sub

Private Sub edcSelCTo_Change()
    Dim ilLen As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim slThruDate As String
    Dim ilListIndex As Integer
    Dim ilDays As Integer

    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            ElseIf ilListIndex = CNT_HILORATE Then
                ilLen = Len(edcSelCTo.Text)
                If ilLen >= 1 Then
                    ilDays = Val(edcSelCTo.Text)
'                    slDate = edcSelCFrom.Text
                    slDate = CSI_CalFrom.Text           '9-11-19 use csi calendar control vs edit box
                    llDate = gDateValue(slDate)
                    slDate = Format(llDate, "m/d/yy")
                    llDate = llDate - ilDays + 1
                    lacSelCTo1.Caption = "thru " & Format$(llDate, "m/d/yy")
                End If
            End If

    End Select
    mSetCommands
End Sub

Private Sub edcSelCTo_GotFocus()
    gCtrlGotFocus edcSelCTo
End Sub

Private Sub edcSelCTo_KeyPress(KeyAscii As Integer)
    Exit Sub
End Sub

Private Sub edcSelCTo1_Change()
    mSetCommands
End Sub

Private Sub edcSelCTo1_GotFocus()
    gCtrlGotFocus edcSelCTo1
End Sub

Private Sub edcSet1_GotFocus()
    gCtrlGotFocus edcSet1
End Sub

Private Sub edcSet2_GotFocus()
    gCtrlGotFocus edcSet2
End Sub

Private Sub edcSet3_GotFocus()
    gCtrlGotFocus edcSet3
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
    RptSelCb.Refresh
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
    'RptSelCb.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tgCSVNameCode
    'Erase tgSellNameCode
    Erase tgRptSelSalespersonCode
    Erase tgRptSelAgencyCode
    Erase tgRptSelAdvertiserCodeCb
    'Erase tgRptSelNameCode
    Erase tgRptSelBudgetCodeCB
    Erase tgMultiCntrCodeCB
    Erase tgManyCntCodeCB
    Erase tgRptSelDemoCodeCB
    Erase tgSOCodeCB
    Erase tgMnfCodeCB
    Erase lgPrintedCnts
    Erase tgClfCB
    Erase tgCffCB
    Erase imCodes
    PECloseEngine
    Set RptSelCb = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcRptType_Click()
    ReDim ilAASCodes(0 To 1) As Integer
                                            'update as ordered (update aired), bill as aired
    rbcSelCInclude(2).Visible = False
    Select Case igRptCallType

        Case CONTRACTSJOB
            mMorelbcRptType
    End Select
    mSetCommands
End Sub

'           5-7-98 If selective advertiser and only one contract exists,
'           force "All contracts" selected on
'
Private Sub lbcSelection_Click(Index As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilUpper                       slCode                    *
'*  ilRet                         ilHOState                     ilHowManyDefined          *
'*                                                                                        *
'******************************************************************************************
    Dim ilListIndex As Integer
    Dim slNameCode As String
    ReDim ilAASCodes(0 To 1) As Integer
    Dim slCntrStatus As String
    If Not imAllClicked Then
        ilListIndex = lbcRptType.ListIndex
        If igRptCallType = CONTRACTSJOB Then
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            Select Case ilListIndex
                Case CNT_SPTSBYADVT, CNT_MGREVENUE, CNT_SPTCOMBO    'summary/spots by advt
                    If rbcSelCSelect(0).Value Then
                        If Index = 5 Then
                            sgMultiCntrCodeTagCB = ""       '1-3-03 clear tag to refresh to cnt list
                            slCntrStatus = "HO"
                            mCntrPop slCntrStatus, 1        'get only orders (w/o revisions)

                            If ilListIndex = CNT_SPTSBYADVT And tgSpf.sSystemType = "R" Then
                                slNameCode = "99999999|999-999||999||[Feed Spots]\0"
                                lbcCntrCode.AddItem slNameCode, 0
                                lbcSelection(0).AddItem "[Feed Spots]", 0 'Add ID to list box

                                'selective adv, turn off generic contract & feed spot selectivty
                                plcSelC12.Visible = False
                                ckcSelC12(0).Value = vbChecked     'default contracts spots on
                                ckcSelC12(1).Value = vbChecked  'default feed spots on
                            End If

                            If imTerminate Then
                                cmcCancel_Click
                                Exit Sub
                            End If
                        Else
                            If Index = 6 Then
                                mSetAllAASUnchecked

                            End If
                        End If
                    Else
                        If Index = 6 Then
                            mSetAllAASUnchecked

                        Else
                            mSetAllUnchecked
                        End If
                    End If
                
                Case CNT_SPTSBYDATETIME, CNT_MISSED  'Spots by times; Missed Spots
                    If Index = 2 Then           '2-16-06
                        mSetAllAASUnchecked

                    Else
                        mSetAllUnchecked
                    End If
                
                Case 6, 7  'Placement; Discrepancy
                    If Index = 5 Then
                        slCntrStatus = "HO"
                        mCntrPop slCntrStatus, 1                'get only orders (no revisions)
                        If imTerminate Then
                            cmcCancel_Click
                            Exit Sub
                        End If
                    End If
                
                Case 8  'MG
                    mSetAllUnchecked
                
                Case 9  'Sales Spot Tracking
                
                Case 10, 12 'Commercial Change, Affiliate Spot Tracking
                
                'Case 11 'History
                '    If index = 5 Then
                 '       mCntrPop igRptType
                 '   End If
                
                Case 13                      'Spot Sales
                    mSetAllUnchecked

                Case CNT_ACCRUEDEFER
                    If Index = 3 Then               'vehicles
                        mSetAllUnchecked
                    ElseIf Index = 4 Then           'sales sources
                        mSetAllAASUnchecked
                    ElseIf Index = 7 Then           'vehicle group items
                        mSetAllVehUnchecked
                    End If
                
                Case CNT_HILORATE
                    mSetAllUnchecked
                    
            End Select
        End If
    Else
        'imSetAll = False
        'ckcAll.Value = False
        'imSetAll = True
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
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(RptSelCb, lbcSelection, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(RptSelCb, lbcSelection, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptSelCb
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
    'ilRet = gPopAgyBox(RptSelCb, lbcSelection, Traffic!lbcAgency)
    ilRet = gPopAgyCollectBox(RptSelCb, "A", lbcSelection, lbcAgyAdvtCode)
    'ilRet = gPopAgyBox(RptSelCb, lbcSelection, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gPopAgyBox)", RptSelCb
        On Error GoTo 0
    End If
    Exit Sub
mAgencyPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'                       sub mChooseSpotType - Select the type
'                       of spots to report:
'                       Charged, 0.00, ADU, Bonus, Extra, Fill,
'                       No Charge, N/C MG, Recapturable, &
'                       Spinoff
Private Sub mChooseSpotType()
    plcSelC5.Visible = True
    'plcSelC5.Caption = "Include"
    smPlcSelC5P = "Include"
    plcSelC5.Visible = True                            'Type check box: trades, DP, PI, psa, etc.

    ckcSelC5(0).Caption = "Charge"
    ckcSelC5(0).Move 720, -30, 910
    ckcSelC5(0).Visible = True
    ckcSelC5(0).Value = vbChecked ' = True
    ckcSelC5(1).Caption = "0.00"
    ckcSelC5(1).Move 1670, -30, 630
    ckcSelC5(1).Visible = True
    ckcSelC5(1).Value = vbChecked   'True
    ckcSelC5(2).Caption = "ADU"
    ckcSelC5(2).Move 2330, -30, 630
    ckcSelC5(2).Visible = True
    ckcSelC5(2).Value = vbChecked 'True
    ckcSelC5(3).Value = vbChecked  'True
    ckcSelC5(3).Caption = "Bonus"
    ckcSelC5(3).Move 3000, -30, 840
    ckcSelC5(3).Visible = True
    ckcSelC5(4).Value = vbChecked 'True
    ckcSelC5(4).Caption = "+Fill"
    ckcSelC5(4).Move 720, 210, 750
    ckcSelC5(4).Visible = True
    ckcSelC5(4).Value = vbChecked
    ckcSelC5(5).Value = vbChecked   'True
    ckcSelC5(5).Caption = "-Fill"
    ckcSelC5(5).Move 1500, 210, 630
    ckcSelC5(5).Visible = True
    ckcSelC5(6).Value = vbChecked
    ckcSelC5(6).Caption = "No Charge"
    ckcSelC5(6).Move 2200, 210, 1350
    ckcSelC5(6).Visible = True
    ckcSelC5(7).Value = vbChecked ' True
    ckcSelC5(7).Caption = "MG"
    ckcSelC5(7).Move 3480, 210, 870
    ckcSelC5(7).Visible = True      '10-18-10 show      'Hide for now, always include MG
    ckcSelC5(8).Value = vbChecked   'True
    ckcSelC5(8).Caption = "Recapturable"
    ckcSelC5(8).Move 720, 420, 1470
    ckcSelC5(8).Visible = True
    ckcSelC5(8).Value = vbChecked   'True
    ckcSelC5(9).Value = vbChecked   'True
    ckcSelC5(9).Caption = "Spinoff"
    ckcSelC5(9).Move 2160, 420, 1000
    ckcSelC5(9).Visible = True  '9-12-02 vbChecked   'True

    '5-25-05 if spts by date & time or spots by advt, allow BB to be excluded
    If RptSelCb!lbcRptType.ListIndex = CNT_SPTSBYADVT Or RptSelCb!lbcRptType.ListIndex = CNT_SPTSBYDATETIME Or RptSelCb!lbcRptType.ListIndex = CNT_SPTCOMBO Then
        ckcSelC5(10).Value = vbUnchecked
        ckcSelC5(10).Caption = "BB"
        ckcSelC5(10).Move 3160, 420, 600
        ckcSelC5(10).Visible = True

    ElseIf RptSelCb!lbcRptType.ListIndex = CNT_MGREVENUE Then
        'makegood revenue doesnt make sense to ask for extras or fills
        'move nocharge to the left, and move recapturable and spinoff to same line as no charge option
        ckcSelC5(4).Visible = False
        ckcSelC5(5).Visible = False
        ckcSelC5(6).Move 720, 195, 1320
        ckcSelC5(7).Move 1965, 195, 600
        ckcSelC5(8).Move 2685, 195, 1440
        ckcSelC5(9).Move 720, 420, 900
        ckcSelC5(6).Value = vbUnchecked   'False       'assume not to include no charge
        ckcSelC5(7).Value = vbUnchecked                 'mg rates (not mg/out spots)
        ckcSelC5(8).Value = vbUnchecked   'False       'assume not to include recapturable
        ckcSelC5(9).Value = vbUnchecked   'False       'assume not to include spinoff
        ckcSelC5(10).Value = vbUnchecked   '5-25-05
        ckcSelC5(1).Value = vbUnchecked   'False       'assume not to include 0.00
        ckcSelC5(2).Value = vbUnchecked   'False       'assume not to include adu
        ckcSelC5(3).Value = vbUnchecked   'False       'assume not to include bonus
    
    ElseIf RptSelCb!lbcRptType.ListIndex = CNT_ACCRUEDEFER Then
        'ignore +/- fills : ckcselc5(4) & ckcselc5(5)
        ckcSelC5(4).Visible = False
        ckcSelC5(5).Visible = False
        ckcSelC5(6).Move 720, 195, 1320
        ckcSelC5(7).Move 1965, 195, 600
        ckcSelC5(8).Move 2685, 195, 1440
        ckcSelC5(9).Move 720, 420, 900
        ckcSelC5(6).Value = vbChecked
        ckcSelC5(7).Value = vbChecked   'include mg/outsides
        ckcSelC5(8).Value = vbChecked   'include recapturable
        ckcSelC5(9).Value = vbChecked   'include spinoff
        ckcSelC5(10).Value = vbUnchecked
        ckcSelC5(1).Value = vbChecked   'include 0.00
        ckcSelC5(2).Value = vbChecked   'include adu
        ckcSelC5(3).Value = vbChecked   'include bonus
        '8-16-12 add option to include/exclude billing methods
        plcSelC13.Move 120, plcSelC12.Top + plcSelC12.Height - 60
        ckcSelC13(0).Move 720, -30, 600
        ckcSelC13(1).Move 1440, -30, 600
        ckcSelC13(2).Move 2160, -30, 960
        
        ckcSelC13(0).Visible = True
        ckcSelC13(1).Visible = True
        ckcSelC13(2).Visible = True
        ckcSelC13(0).Value = vbChecked       'default cal on
        ckcSelC13(1).Value = vbChecked      'default std on
        ckcSelC13(2).Value = vbChecked      'default weekly on
        plcSelC13.Visible = True
    Else
        ckcSelC5(10).Visible = False         'bb
        ckcSelC5(10).Value = vbUnchecked     'insure not selected
    End If
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
    llLen = 0
    ilErr = False
    lbcSelection(0).Clear
    lbcCntrCode.Clear
Screen.MousePointer = vbHourglass
    For illoop = 0 To lbcSelection(5).ListCount - 1 Step 1
        If lbcSelection(5).Selected(illoop) Then
            sgMultiCntrCodeTag = ""             'init the date stamp so the box will be populated
            ReDim tgMultiCntrCodeCB(0 To 0) As SORTCODE
            lbcMultiCntr.Clear
            slNameCode = tgAdvertiser(illoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            'ilCurrent = 1   '0=Current; 1=All
            'ilFilter = Val(slCode)   'by contract #; -101=by advertiser
            'ilVehCode = -1  'All vehicles
            'ilRet = gPopCntrBox(RptSelCb, ilCntrType, ilFilter, ilCurrent, 0, ilVehCode, lbcMultiCntr, lbcMultiCntrCode, True, False, False, False)
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
            'ilRet = gPopCntrForAASBox(RptSelCb, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, lbcMultiCntrCode)
            ilRet = gPopCntrForAASBox(RptSelCb, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, tgMultiCntrCodeCB(), sgMultiCntrCodeTagCB)
            sgMultiCntrCodeTagCB = ""
            If ilRet <> CP_MSG_NOPOPREQ Then
                On Error GoTo mCntrPopErr
                gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", RptSelCb
                On Error GoTo 0
            End If
            For ilIndex = 0 To UBound(tgMultiCntrCodeCB) - 1 Step 1 'lbcMultiCntrCode.ListCount - 1 Step 1
                slName = Trim$(tgMultiCntrCodeCB(ilIndex).sKey)  'lbcMultiCntrCode.List(ilIndex)
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
    Next illoop
    For illoop = 0 To lbcCntrCode.ListCount - 1 Step 1
        slNameCode = lbcCntrCode.List(illoop)
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
    Next illoop
    Screen.MousePointer = vbDefault
    Exit Sub
mCntrPopErr:
 On Error GoTo 0
    Screen.MousePointer = vbDefault
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
    Dim ilMultiTable As Integer
    Dim illoop As Integer
    Dim slStr As String
    imFirstActivate = True
    hmMtf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMtf, "", sgDBPath & "Mtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        lgMtfNoRecs = btrRecords(hmMtf)
        btrDestroy hmMtf
    Else
        lgMtfNoRecs = 0
    End If
    Screen.MousePointer = vbHourglass
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If
    'Set options for report generate
'VB6**    hdJob = rpcRpt.hJob
    ilMultiTable = True
    'dummy = LlSetOption(hdJob, LL_OPTION_HELPAVAILABLE, False)
'VB6**    ilDummy = LlSetOption(hdJob, LL_OPTION_SORTVARIABLES, True)
'VB6**    ilDummy = LlSetOption(hdJob, LL_OPTION_ONLYONETABLE, Not ilMultiTable)

    RptSelCb.Caption = smSelectedRptName & " Report"
    'frcOption.Caption = smSelectedRptName & " Selection"
    slStr = Trim$(smSelectedRptName)
    illoop = InStr(slStr, "&")
    If illoop > 0 Then
        slStr = Left$(slStr, illoop - 1) & "&&" & Mid$(slStr, illoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imSetAllAAS = True
    imAllClickedAAS = False
    imAllClickedVeh = False
    imSetAllVeh = True
'    cbcSel.Move 120, 30
    plcSelC3.Height = 240
    lacSelCFrom.Move 120, 75
    lacSelCTo.Move 120, 390
    lacSelCFrom1.Move 2400, 75
    lacSelCTo1.Move 2400, 390
    edcSelCFrom.Move 1500, 30
    edcSelCFrom1.Move 3240, 30
    edcSelCTo.Move 1500, 345
    edcSelCTo1.Move 2715, 345
    plcSelC1.Move 120, 675
    'plcSelC1.Caption = "Select"
    smPlcSelC1P = "Select"
    rbcSelCSelect(0).Move 600, 0
    rbcSelCSelect(0).Caption = "Advt"
    rbcSelCSelect(1).Move 1290, 0
    rbcSelCSelect(1).Caption = "Agency"
    rbcSelCSelect(2).Move 2020, 0
    rbcSelCSelect(2).Caption = "Salesperson"
    plcSelC2.Move 120, 885
    'plcSelC2.Caption = "Include"
    smPlcSelC2P = "Include"
    rbcSelCInclude(0).Move 705, 0
    rbcSelCInclude(0).Caption = "All"
    rbcSelCInclude(1).Move 1245, 0
    rbcSelCInclude(2).Move 2655, 0
    plcSelC3.Move 120, 675
    'plcSelC3.Caption = "Zone"
    smPlcSelC3P = "Zone"
    ckcSelC3(0).Move 465, -30
    ckcSelC3(0).Caption = "EST"
    ckcSelC3(1).Move 1065, -30
    ckcSelC3(1).Caption = "CST"
    ckcSelC3(2).Move 1710, -30
    ckcSelC3(2).Caption = "MST"
    ckcSelC3(3).Move 2355, -30
    ckcSelC3(3).Caption = "PST"
    ckcSelC10(0).Visible = False
    plcSelC10.Visible = False
    edcSet1.Visible = False
    edcSet2.Visible = False
    cbcSet1.Visible = False
    cbcSet2.Visible = False

    plcSelC4.Move 120, 360
    plcSelC5.Move 120, 1095
    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3270
    lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    'pbcSelC.Move 90, 255, 4515, 3360
'    pbcSelC.Move 90, 255, 4515, 3930
    pbcSelC.Move 90, 255, 4695, 4690
    gCenterStdAlone RptSelCb
End Sub

'           mInitControls - set controls to proper positions, sizes
'                   hidden, shown, etc.
'
'           Created :  11/28/98 D Hosaka
'
Private Sub mInitControls()
    lbcSelection(0).Move 15, ckcAll.Height + 30, 4365, 4200 '4050 '3270
    lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(2).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(3).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(4).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(5).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(5).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width / 2 - 30, lbcSelection(5).Height  '1110
    lbcSelection(6).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(8).Move lbcSelection(6).Left, lbcSelection(6).Top, lbcSelection(6).Width, lbcSelection(6).Height       'conv & airing
    lbcSelection(0).Move lbcSelection(5).Left + lbcSelection(5).Width + 60, lbcSelection(0).Top, lbcSelection(0).Width / 2 - 30, lbcSelection(0).Height '840
    lbcSelection(7).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width, lbcSelection(5).Height       'advt
    lbcSelection(0).Visible = False
    lbcSelection(1).Visible = False
    lbcSelection(2).Visible = False
    lbcSelection(3).Visible = False
    lbcSelection(4).Visible = False
    lbcSelection(5).Visible = False
    lbcSelection(6).Visible = False
    lbcSelection(7).Visible = False
    lbcSelection(8).Visible = False
    plcSelC6.Visible = False
    lacSelCFrom.Visible = False
    edcSelCFrom.Visible = False
    lacSelCFrom1.Visible = False
    edcSelCFrom1.Visible = False
    lacSelCFrom1.Width = 810
    lacSelCTo.Visible = False
    edcSelCTo.Visible = False
    lacSelCTo1.Visible = False
    edcSelCTo1.Visible = False
    plcSelC1.Visible = False
    rbcSelCSelect(3).Visible = False
    rbcSelC4(2).Enabled = True
    ckcSelC6(0).Value = vbUnchecked   'False
    ckcSelC6(1).Visible = False
    ckcSelC6(1).Value = vbUnchecked    'False
    ckcSelC6(2).Visible = False
    ckcSelC6(3).Visible = False
    ckcSelC6(4).Visible = False
    plcSelC2.Enabled = True
    plcSelC1.Visible = False
    plcSelC2.Visible = False
    plcSelC4.Visible = False
    plcSelC5.Visible = False
    plcSelC6.Visible = False
    plcSelC7.Visible = False
    plcSelC8.Visible = False
'   cbcSel.Visible = False
    lacSelCFrom.Move 120, 75, 1380
    'lacSelCFrom.Width = 1380
    'edcSelCFrom.Move 1500, edcSelCFrom.Top, 1350
    edcSelCFrom.Move 1500, 30, 1350
    lacSelCTo.Move 120, 390, 1380
    lacSelCTo1.Move 2400, 390
    edcSelCTo.Move 1500, 345, 1350
    edcSelCTo.MaxLength = 10    '8 5/27/99 changed for short form date m/d/yyyy
    edcSelCTo1.Move 2715, 345
    edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
    edcSelCFrom1.Move 3240, 30
    lacSelCFrom1.Move 2340, 75
    edcSelCFrom.Text = ""
    edcSelCFrom1.Text = ""
    edcSelCTo.Text = ""
    edcSelCTo1.Text = ""
    plcSelC1.Top = 675
    plcSelC2.Top = 885
    plcSelC2.Height = 240
    plcSelC1.Left = 120
    plcSelC2.Left = 120
    plcSelC5.Height = 240
    rbcSelCInclude(2).Left = 2655
    rbcSelCInclude(2).Top = 0
    rbcSelCInclude(3).Visible = False
    rbcSelCInclude(4).Visible = False
    rbcSelCSelect(2).Top = 0
    rbcSelCSelect(3).Top = 0
    rbcSelC4(2).Value = False
    plcSelC1.Height = 240
    plcSelC3.Height = 240
    ckcSelC3(3).Top = ckcSelC3(0).Top
    ckcSelC3(4).Top = 210
    ckcSelC3(5).Top = 210
    ckcSelC3(6).Top = 210
    ckcSelC3(2).Visible = False
    ckcSelC3(3).Visible = False
    ckcSelC3(4).Visible = False
    ckcSelC3(5).Visible = False
    ckcSelC3(6).Visible = False
    ckcSelC5(2).Visible = False
    ckcSelC5(3).Visible = False
    ckcSelC5(4).Visible = False
    ckcSelC5(5).Visible = False
    ckcSelC5(6).Visible = False
    ckcSelC5(7).Visible = False
    ckcSelC5(8).Visible = False
    ckcSelC5(9).Visible = False
    ckcSelC8(0).Visible = False
    ckcSelC8(1).Visible = False
    ckcSelC8(2).Visible = False
    ckcSelC5(1).Enabled = True
    plcSelC3.Visible = False
    ckcSelC6(0).Enabled = True
    edcSelCTo.Text = ""
    ckcAll.Move lbcSelection(1).Left            'readjust 'Check All' location to be above left most list box
    ckcAllAAS.Move ckcAll.Left, ckcAll.Top
    ckcAll.Enabled = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'
'             5/5/98 - Set frcOption.Visible to false
'             at developement, at run time set to
'             True when all questions on screen have
'             been formatted.  This way the screen comes
'             up all at once, rather than pieces.
'
'       6-16-00 Remove all references to Contract "BR"
'               and Insertion Orders (reports are coded
'               in rptselct)
'
'*                                                     *
'*******************************************************
Private Sub mInitReport()
    Dim ilIndex As Integer
    Dim ilRet As Integer
    gPopExportTypes cbcFileType     '10-20-01
    pbcSelC.Visible = False
    lbcRptType.Clear

    Select Case igRptCallType
        Case CONTRACTSJOB
            lgOrigCntrNo = 0
            hmSpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                imSpfRecLen = Len(tmSpf)
                ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    lgOrigCntrNo = tmSpf.lDiscCurrCntrNo
                End If
                ilRet = btrClose(hmSpf)
                btrDestroy hmSpf
            End If
            'RptSelCb.Caption = "Contract Report Selection"
            mAdvtPop lbcSelection(5)    'Called to initialize Traffic!Advertiser required be mCntrPop
            If imTerminate Then
                Exit Sub
            End If
            lbcSelection(0).Clear
            lbcSelection(0).Tag = ""
            Screen.MousePointer = vbHourglass
            mAgencyPop lbcSelection(1)
            If imTerminate Then
                Exit Sub
            End If
            'mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                                                'populate when needed
            mSellConvVehPop 3
            If imTerminate Then
                Exit Sub
            End If
            'mSellConvVirtVehPop 6, False
            'lbcselection(11) used for demos (cpp/cpm report) and single select budgets (tieout report), populate when needed
            'ilRet = gPopMnfPlusFieldsBox(RptSelCb, lbcSelection(11), lbcDemoCode, "D")
            lbcRptType.AddItem "Proposals/Contracts", 0                         '0=proposal
            lbcRptType.AddItem "Paperwork Summary", 1                           '1=paperwork summary (contract summaries)

            'If tgUrf(0).islfCode = 0 Then           'its a slsp thats is asking for this report,
                                                    'don't allow them to exclude reserves
                ilIndex = 2
                If igRptType = 0 Then   'Proposal
                    'rbcRptType(2).Visible = False
                Else    'Contract
                    'rbcRptType(2).Caption = "Spots by Advt"
                    lbcRptType.AddItem "Spots by Advertiser", ilIndex           '2=spots by advt
                    ilIndex = ilIndex + 1
                End If
                lbcRptType.AddItem "Spots by Date & Time", ilIndex              '3=spots by date & time
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Business Booked by Contract", ilIndex       '4=projection (named changed to Business Booked)
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Contract Recap", ilIndex                    '5=contr recap
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot Placements", ilIndex                   '6=Spot placements
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot Discrepancies", ilIndex                '7=spot discrepancies
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "MG's", ilIndex                              '8=makegood
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Spot Tracking", ilIndex               '9=sales spot traking
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Commercial Changes", ilIndex                '10=coml changes
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Contract History", ilIndex                  '11 Contract history
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Affiliate Spot Tracking", ilIndex           '12 affil spot traking
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot Sales", ilIndex                        '13=spot sales
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Missed Spots", ilIndex                      '14=missed spots
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot Business Booked", ilIndex              '15=spot projection (name changed to Business Booked)
                ilIndex = ilIndex + 1
                'spot reprints - used
                lbcRptType.AddItem "Business Booked by Spot Reprint", ilIndex   '16= Business booked reprint
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Avails", ilIndex                            '17=quarterly summary & detail avails
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Average Spot Prices", ilIndex               '18=avg spot prices
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Advertiser Units Ordered", ilIndex          '19=advt units ordered
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Analysis by CPP & CPM", ilIndex       '20=sales analysis by cpp & cpm
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Average Rate", ilIndex                      '21=Average Rate
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Tie-Out", ilIndex                           '22=Detail Tie Out
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Billed and Booked", ilIndex                 '23=Billed & booked by advt, Slsp, owner, vehicle
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Weekly Sales Activity by Quarter", ilIndex  '24=Sales Activity
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Comparison", ilIndex                  '25=Sales Comparison by Advt, Slsp, Agy, comp code, Bus code
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Weekly Sales Activity by Month", ilIndex    '26=Cumulative Activity Report (pacing)
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Average Prices to Make Plan", ilIndex       '27=Avg Prices needed to make plan
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "CPP/CPM by Vehicle", ilIndex                '28=Curent cpp/cpm by vehicle
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Analysis Summary", ilIndex            '29=Sales Analysis Summary
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Insertion Orders", ilIndex                  '30=
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Makegood Revenue", ilIndex                  '31=
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Accrual/Deferral", ilIndex                  '32=12-20-06
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Hi-Lo Spot Rate", ilIndex                   '33=6-2-10
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot Discrepancy Summary by Month", ilIndex '34=6-21-16
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot and Digital Line Combo", ilIndex '35=TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines

            'End If
            'frcOption.Caption = "Contract Selection"
            ckcAll.Caption = "All Contracts"
            frcOption.Enabled = True
            pbcSelC.Height = pbcSelC.Height - 60
            'lbcSelection(0).Move lbcSelection(0).Left, pbcSelC.Height + 150, lbcSelection(0).Width, ckcAll.Top - pbcSelC.Height
            'rbcSelCSelect(0).Value = True   'Advertiser/Contract #
            'lbcRptType.ListIndex = 0
            pbcSelC.Visible = True
            pbcOption.Visible = True
            lacSelCFrom.Visible = False
            lacSelCTo.Visible = False
            edcSelCFrom.Visible = False
            edcSelCTo.Visible = False
            ckcAll.Visible = False
    End Select
    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
        'RptSelCb.Caption = smSelectedRptName & " Report"
        'frcOption.Caption = smSelectedRptName & " Selection"
        'slStr = Trim$(smSelectedRptName)
        'ilLoop = InStr(slStr, "&")
        'If ilLoop > 0 Then
        '    slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
        'End If
        'frcOption.Caption = slStr & " Selection"
    End If
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
End Sub

'***********************************************************************
'*                                                                     *
'*      Procedure Name:mMorelbcRptType                                 *
'*                                                                     *
'*             Created:5/17/93       By:D. LeVine                      *
'*            Modified:              By:D. Smith                       *
'*                                                                     *
'*            Comments: D.S. 4/20/00 Converted Spots by Date and Time  *
'*                      from Bridge to Crystal. Added Start and End    *
'*                      Time to Missed Spots and Spots by Day and Time *
'*
'       6-16-00 Remove all references to Contract "BR"
'               and Insertion Orders (reports are coded
'               in rptselct)
'*      12-20-06 Accrual/Deferral report
'***********************************************************************
Private Sub mMorelbcRptType()
    Dim ilListIndex As Integer
    Dim ilRet As Integer
    Dim ilValue As Integer
    Dim ilTop As Integer
    ReDim ilAASCodes(0 To 1) As Integer

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    rbcSelCInclude(2).Visible = False
    If tgSpf.sRUseCorpCal = "Y" Then       'if Using corp cal,retain in memory
        ilRet = gObtainCorpCal()
    End If
    'ilRet = gObtainVef()
    'Select Case igRptCallType
        'Case CONTRACTSJOB
        ilListIndex = lbcRptType.ListIndex
        If (igRptType = 0) And (ilListIndex > 1) Then
            ilListIndex = ilListIndex + 1
        End If
        mInitControls           'set controls to proper positions, widths, hidden, shown, etc.
        mSellConvVirtVehPop 6, False
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        Select Case ilListIndex
            Case CNT_SPTSBYADVT, CNT_SPTCOMBO    'spots by advt, Spot and Digital Line Combo Report
                mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                                                   'populate when needed
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
                ckcAll.Visible = True
                ckcAll.Enabled = True
                '9-12-19 use csi calendar controls vs editbox
                lacSelCFrom.Move 120, 60, 1320
                If ilListIndex = CNT_SPTCOMBO Then
                    lacSelCFrom.Caption = "Report Dates"
                Else
                    lacSelCFrom.Caption = "Spots-From"
                End If
                lacSelCFrom.Visible = True
                CSI_CalFrom.Move 1320, 30
                CSI_CalFrom.Visible = True
                lacSelCFrom1.Move 2535, lacSelCFrom.Top, 250
                lacSelCFrom1.Caption = "To"
                lacSelCFrom1.Visible = True
                CSI_CalTo.Move 2895, CSI_CalFrom.Top
                CSI_CalTo.Visible = True
                CSI_CalFrom.CSI_AllowBlankDate = True
                CSI_CalTo.CSI_AllowBlankDate = True
               
                lacSelCTo.Caption = "Entered-From"
                If ilListIndex = CNT_SPTCOMBO Then
                    CSI_CalFrom2.Visible = False
                    CSI_CalTo2.Visible = False
                    lacSelCTo.Visible = False
                Else
                    lacSelCTo.Move lacSelCFrom.Left, CSI_CalFrom.Top + CSI_CalFrom.Height + 90, 1440
                    lacSelCTo.Visible = True
                    CSI_CalFrom2.Move CSI_CalFrom.Left, CSI_CalFrom.Top + CSI_CalFrom.Height + 60
                    CSI_CalFrom2.Visible = True
                    lacSelCTo1.Move lacSelCFrom1.Left, lacSelCTo.Top, 240
                    lacSelCTo1.Visible = True
                    CSI_CalTo2.Move CSI_CalTo.Left, CSI_CalFrom2.Top
                    CSI_CalTo2.Visible = True
                    CSI_CalFrom2.CSI_AllowBlankDate = True
                    CSI_CalTo2.CSI_AllowBlankDate = True
                End If
                
                plcSelC3.Visible = False
                smPlcSelC1P = "Select"
                If ilListIndex = CNT_SPTCOMBO Then
                    rbcSelCSelect(0).Caption = "Advertiser"
                    rbcSelCSelect(0).Left = 720
                    rbcSelCSelect(0).Width = 1275
                    rbcSelCSelect(1).Caption = "Agency"
                    rbcSelCSelect(1).Left = 1990
                    rbcSelCSelect(1).Width = 980
                    rbcSelCSelect(1).Visible = True
                    rbcSelCSelect(1).Enabled = True
                Else
                    rbcSelCSelect(0).Caption = "Advt"
                    rbcSelCSelect(0).Left = 600
                    rbcSelCSelect(0).Width = 675
                    rbcSelCSelect(1).Caption = "Agency"
                    rbcSelCSelect(1).Left = 1290
                    rbcSelCSelect(1).Width = 980
                    rbcSelCSelect(1).Visible = True
                    rbcSelCSelect(1).Enabled = True
                End If
                
                If ilListIndex = CNT_SPTCOMBO Then
                    rbcSelCSelect(2).Enabled = False
                    rbcSelCSelect(2).Visible = False
                Else
                    rbcSelCSelect(2).Caption = "Salesperson"
                    rbcSelCSelect(2).Left = 2250
                    rbcSelCSelect(2).Width = 1370
                    rbcSelCSelect(2).Enabled = True
                    rbcSelCSelect(2).Visible = True
                End If
                plcSelC1.Visible = True
                lbcSelection(3).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(5).Visible = False
                lbcSelection(0).Visible = False
                lbcSelection(4).Visible = False
                lbcSelection(6).Visible = False
                If rbcSelCSelect(0).Value Then
                    lbcSelection(1).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(5).Visible = True
                    lbcSelection(0).Visible = True
                    ckcAll.Caption = "All Advertisers"
                ElseIf rbcSelCSelect(1).Value Then
                    lbcSelection(0).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(5).Visible = False
                    lbcSelection(1).Visible = True
                    ckcAll.Caption = "All Agencies"
                ElseIf rbcSelCSelect(2).Value Then
                    lbcSelection(0).Visible = False
                    lbcSelection(1).Visible = False
                    lbcSelection(5).Visible = False
                    lbcSelection(2).Visible = True
                    ckcAll.Caption = "All Salespeople"
                End If
                
                If ilListIndex = CNT_SPTCOMBO Then
                    plcSelC1.Move 120, 0, 3005
                    lacSelCFrom.Top = plcSelC1.Top + plcSelC1.Height + 30
                    CSI_CalFrom.Top = plcSelC1.Top + plcSelC1.Height + 30
                    lacSelCFrom1.Top = plcSelC1.Top + plcSelC1.Height + 30
                    CSI_CalTo.Top = plcSelC1.Top + plcSelC1.Height + 30
                Else
                    smPlcSelC2P = "Show Spot Prices"
                    rbcSelCInclude(0).Caption = "Yes"
                    rbcSelCInclude(0).Left = 1590
                        rbcSelCInclude(0).Width = 615
                    If rbcSelCInclude(0).Value Then
                        rbcSelCInclude_Click 0
                    Else
                        rbcSelCInclude(0).Value = vbChecked 'True
                    End If
                    rbcSelCInclude(1).Caption = "No"
                    rbcSelCInclude(1).Left = 2310
                    rbcSelCInclude(1).Width = 615
                    rbcSelCInclude(1).Enabled = True
                    plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height + 30, 2805
                    plcSelC2.Visible = True
                End If
                rbcSelCSelect(0).Value = True           'force default to advt
                If gUsingBarters() = True Then          '11-5-15
                    plcSelC8.Move plcSelC2.Left + plcSelC2.Width, plcSelC2.Top, 1560
                    ckcSelC8(0).Width = 1560
                    ckcSelC8(0).Caption = "Use Acq*"
                    ckcSelC8(0).Value = vbUnchecked
                    ckcSelC8(0).Visible = True
                    plcSelC8.Visible = True
                End If
                
                '12-28-17
                mAskContractTypesCkcSelC3 plcSelC2.Top + plcSelC2.Height
                plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height + 30
                plcSelC5.Height = 655                   '3 lines of check boxes
                mChooseSpotType
                
                plcSelC4.Move 120, plcSelC5.Top + plcSelC5.Height, 3480, 495
                'plcSelC4.Caption = "Sort by Advt, Contract -"
                If ilListIndex = CNT_SPTCOMBO Then
                    smPlcSelC4P = "Sort by"
                    rbcSelC4(0).Caption = "Advertiser"
                    rbcSelC4(1).Caption = "Vehicle"
                    rbcSelC4(0).Move 720, 0, 1680
                    rbcSelC4(1).Move 2160, 0, 1680, 220
                    'TTP 10892 - Spot and Digital Line Combo report: new option to exclude invoice adjustments
                    plcSelC15.Move 120, 665, 4480, 975  'Include Spots,Digital,Digital Comments
                    plcSelC15.Visible = True
                    plcSelC3.Top = plcSelC15.Top + plcSelC15.Height + 60  '1st include list (Contract types)
                    plcSelC5.Top = plcSelC3.Top + plcSelC3.Height + 80 '2nd include list (Spot types)
                    plcSelC4.Top = plcSelC4.Top + plcSelC4.Height + 80 'Sort by
                    plcSelC15.Visible = True
                    plcSelC4.Move 120, plcSelC5.Top + plcSelC5.Height + 80, 3480, 495
                Else
                    plcSelC15.Visible = False
                    smPlcSelC4P = "Sort by Advt, Contract -"
                    rbcSelC4(0).Caption = "Vehicle, Date"
                    rbcSelC4(1).Caption = "Date, Vehicle"
                    rbcSelC4(0).Move 720, 195, 1680
                    rbcSelC4(1).Move 2160, rbcSelC4(0).Top, 1680
                End If
                plcSelC4.Visible = True
                rbcSelC4(0).Visible = True
                rbcSelC4(1).Visible = True
                If rbcSelC4(0).Value Then
                    rbcSelC4_click 0
                Else
                    rbcSelC4(0).Value = True
                End If
                rbcSelC4(2).Visible = False
                    
                If ilListIndex = CNT_SPTCOMBO Then
                    'Show: Gross or net?
                    'TTP 10892 - Spot and Digital Line Combo report: new option to exclude invoice adjustments
                    plcSelC7.Move 120, plcSelC4.Top + 220
                    RptSelCb.plcSelC7.ZOrder (0)
                    'plcSelC7.Top = 3300
                    rbcSelC7(1).Left = 720
                    rbcSelC7(0).Left = 1660
                    plcSelC7.Visible = True
                    rbcSelC7(0).Visible = True
                    rbcSelC7(1).Visible = True
                    rbcSelC7(2).Visible = False
                    rbcSelC7(1).Value = True
                    smPlcSelC7P = "Show"
                Else
                    plcSelC7.Move 120, plcSelC4.Top + plcSelC4.Height
                    'plcSelC7.Caption = "Show Status Column"
                    smPlcSelC7P = "Show Status Column"
                    rbcSelC7(0).Caption = "Yes"
                    rbcSelC7(1).Caption = "No"
                    rbcSelC7(0).Left = 1800
                    rbcSelC7(1).Left = 2460
                    plcSelC7.Visible = True
                    rbcSelC7(0).Visible = True
                    rbcSelC7(1).Visible = True
                    If rbcSelC7(0).Value Then
                        rbcSelC7_click 0
                    Else
                        rbcSelC7(0).Value = True
                    End If
                    rbcSelC7(2).Visible = False
                    '2-1-01 Detail or Summary
                    If rbcSelC9(0).Value Then
                        rbcSelC9_click 0
                    Else
                        rbcSelC9(0).Value = True
                    End If
                    rbcSelC9(1).Value = False
                    rbcSelC9(0).Visible = True
                    rbcSelC9(1).Visible = True
                    rbcSelC9(2).Visible = False
                    smPlcSelC9P = "By"
                    rbcSelC9(0).Caption = "Detail"
                    rbcSelC9(1).Caption = "Summary"
                    rbcSelC9(0).Left = 360
                    rbcSelC9(1).Left = 1200
                    rbcSelC9(1).Width = 1160
                    plcSelC9.Move plcSelC7.Left, plcSelC7.Top + plcSelC7.Height
                    plcSelC9.Visible = True
                End If
                
                '5-25-06 feature to calc gross or net
                If ilListIndex = CNT_SPTCOMBO Then
                    
                Else
                    plcSelC11.Move 120, plcSelC9.Top + plcSelC9.Height
                    smPlcSelC11P = "Show"
                    rbcSelC11(0).Caption = "Gross"
                    rbcSelC11(1).Caption = "Net"
                    rbcSelC11(0).Move 600, 0, 960
                    rbcSelC11(1).Move 1560, 0, 1920
                    rbcSelC11(0).Visible = True
                    rbcSelC11(1).Visible = True
                    rbcSelC11(2).Visible = False
                    rbcSelC11(0).Value = True
                    plcSelC11.Visible = True
                    
                    cbcSet1.Move 120, plcSelC11.Top + plcSelC11.Height + 30, 2000
                    cbcSet1.AddItem "All Audio Types"
                    cbcSet1.AddItem "Live Coml"
                    cbcSet1.AddItem "Live Promo"
                    cbcSet1.AddItem "Pre-Recorded Coml"
                    cbcSet1.AddItem "Pre-Recorded Promo"
                    cbcSet1.AddItem "Recorded Coml"
                    cbcSet1.AddItem "Recorded Promo"
                    cbcSet1.ListIndex = 0
                    cbcSet1.Visible = True
                    
                    mAskCntrFeed cbcSet1.Top + cbcSet1.Height + 30   'ask contracts spots and feed spots selectiviy
                    pbcSelC.Height = 4430
                    pbcSelC.Visible = True
                    pbcOption.Visible = True
                    
                    ckcIncludeISCI.Move 120, cbcSet1.Top + cbcSet1.Height + 30
                    ckcIncludeISCI.Height = 254
                    ckcIncludeISCI.Visible = True
                    ckcIncludeISCI.Value = False
                End If
                'TTP 11089 - Spots by Advertiser: Audio type selection box overlaps text
                If ilListIndex <> 2 Then
                    lacContract.Caption = "Contract #"
                    lacContract.Move 120, 4000, 975
                    lacContract.Visible = True
                    edcSet3.Move lacContract.Left + lacContract.Width, 4000, 960
                    edcSet3.Visible = True
                    edcSet3.MaxLength = 8
                End If
                
            Case CNT_MGREVENUE                          '6-16-00
                'Prepass code uses same as Spots by Advertiser.  Entered Date selectivity is not used in this report.
                'The fields must be set to "all Dates" before processing
                mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                                                    'populate when needed
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
                ckcAll.Visible = True   '9-12-02 vbChecked  'True
                ckcAll.Enabled = True
                
                '9-11-19  change to use csi calendar vs edit box for date input
                lacSelCFrom.Move 120, 60, 1320
                lacSelCFrom.Caption = "Missed-From"
                lacSelCFrom.Visible = True
                CSI_CalFrom.Move 1320, 30
                CSI_CalFrom.Visible = True
                lacSelCFrom1.Move 2535, lacSelCFrom.Top, 240
                lacSelCFrom1.Caption = "To"
                lacSelCFrom1.Visible = True
                CSI_CalTo.Move 2895, CSI_CalFrom.Top
                CSI_CalTo.Visible = True
                
                'mg selectivity
                lacSelCTo.Move lacSelCFrom.Left, CSI_CalFrom.Top + CSI_CalFrom.Height + 90, 1440
                lacSelCTo.Caption = "MG-From"
                lacSelCTo.Visible = True
                CSI_CalFrom2.Move CSI_CalFrom.Left, CSI_CalFrom.Top + CSI_CalFrom.Height + 60
                CSI_CalFrom2.Visible = True
                lacSelCTo1.Move lacSelCFrom1.Left, lacSelCTo.Top, 240
                lacSelCTo1.Visible = True
                CSI_CalTo2.Move CSI_CalTo.Left, CSI_CalFrom2.Top
                CSI_CalTo2.Visible = True
                 
                plcSelC1.Move 120, CSI_CalFrom2.Top + CSI_CalFrom.Height + 60
                smPlcSelC1P = "Select"
                rbcSelCSelect(0).Caption = "Advt"
                rbcSelCSelect(0).Left = 600
                rbcSelCSelect(0).Width = 675
                rbcSelCSelect(1).Caption = "Agency"
                rbcSelCSelect(1).Left = 1290
                rbcSelCSelect(1).Width = 945
                rbcSelCSelect(1).Visible = True
                rbcSelCSelect(2).Caption = "Salesperson"
                rbcSelCSelect(2).Left = 2220
                rbcSelCSelect(2).Width = 1380
                rbcSelCSelect(1).Enabled = True
                rbcSelCSelect(2).Enabled = True
                rbcSelCSelect(2).Visible = True
                plcSelC1.Visible = True
                lbcSelection(3).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(5).Visible = False
                lbcSelection(0).Visible = False
                lbcSelection(4).Visible = False
                lbcSelection(6).Visible = False
                If rbcSelCSelect(0).Value Then
                    lbcSelection(1).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(5).Visible = True
                    lbcSelection(0).Visible = True
                    ckcAll.Caption = "All Advertisers"
                    'ckcAll.Visible = True
                ElseIf rbcSelCSelect(1).Value Then
                    lbcSelection(0).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(5).Visible = False
                    lbcSelection(1).Visible = True
                    ckcAll.Caption = "All Agencies"
                    'ckcAll.Visible = True
                ElseIf rbcSelCSelect(2).Value Then
                    lbcSelection(0).Visible = False
                    lbcSelection(1).Visible = False
                    lbcSelection(5).Visible = False
                    lbcSelection(2).Visible = True
                    ckcAll.Caption = "All Salespeople"
                    'ckcAll.Visible = True
                End If
                plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height
                'plcSelC2.Caption = "Show Spot Prices"
                smPlcSelC2P = "Show Spot Prices"
                rbcSelCInclude(0).Caption = "Yes"
                rbcSelCInclude(0).Left = 1470
                rbcSelCInclude(0).Width = 585
                If rbcSelCInclude(0).Value Then
                    rbcSelCInclude_Click 0
                Else
                    rbcSelCInclude(0).Value = True
                End If
                rbcSelCSelect(0).Value = True           'force default to advt
                rbcSelCInclude(1).Caption = "No"
                rbcSelCInclude(1).Left = 2070
                rbcSelCInclude(1).Width = 615
                rbcSelCInclude(1).Enabled = True
                plcSelC2.Visible = True
                plcSelC5.Move 120, plcSelC2.Top + plcSelC2.Height
                plcSelC5.Height = 655                   '3 lines of check boxes
                mChooseSpotType
                'Currently make spot types not visible,
                'Make Show Spot prices not visible
                plcSelC2.Visible = False                'show spot prices
                'plcSelC5.Visible = False                'Type check box: trades, DP, PI, psa, etc.
                plcSelC6.Move plcSelC5.Left, plcSelC5.Top + plcSelC5.Height + 30
                'plcSelC6.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height
                'plcSelC6.Caption = "Include"
                smPlcSelC6P = "Include"
                plcSelC6.Height = 720       '945 chged 3-27-11
                ckcSelC6(0).Caption = "Makegood"
                ckcSelC6(0).Move 720, -30, 1200
                ckcSelC6(0).Value = vbChecked   ' True
                ckcSelC6(0).Visible = True
                ckcSelC6(1).Caption = "Outside"
                ckcSelC6(1).Move 1920, -30, 1200
                ckcSelC6(1).Value = vbChecked   ' True
                ckcSelC6(1).Visible = True
                ckcSelC6(2).Caption = "Missed"
                ckcSelC6(2).Move 720, 225, 1200
                ckcSelC6(2).Value = vbChecked   'True
                ckcSelC6(2).Visible = True
                ckcSelC6(3).Caption = "Cancelled"
                ckcSelC6(3).Move 1920, 225, 1200
                ckcSelC6(3).Value = vbChecked   ' True
                ckcSelC6(3).Visible = True
                ckcSelC6(4).Caption = "For MG/Outside- show switched vehicles only "
                ckcSelC6(4).Move 0, 480, 4530, 255

                ckcSelC6(4).Value = vbChecked   ' True
                ckcSelC6(4).Visible = True
                plcSelC6.Visible = True

                'Set the default to show Status column, status will always show,
                '****NOTE Dont use plcselc7 for any new options.  this is tested in prepass
                plcSelC7.Move 120, plcSelC5.Top + plcSelC5.Height
                'plcSelC7.Caption = "Show Status Column"
                smPlcSelC7P = "Show Status Column"
                rbcSelC7(0).Caption = "Yes"
                rbcSelC7(1).Caption = "No"
                rbcSelC7(0).Left = 1800
                rbcSelC7(1).Left = 2460
                plcSelC7.Visible = True
                rbcSelC7(0).Visible = True
                rbcSelC7(1).Visible = True
                If rbcSelC7(0).Value Then
                    rbcSelC7_click 0
                Else
                    rbcSelC7(0).Value = True
                End If
                rbcSelC7(2).Visible = False
                plcSelC7.Visible = False
                pbcSelC.Visible = True
                pbcOption.Visible = True
                '7-20-04 Force local spots to be included, network (feed) to be excluded.
                'MG Revenue report uses same subroutine as Spots by Advertiser (gSpotAdvtRpt)
                ckcSelC12(0).Value = vbChecked          'include feed spots
                ckcSelC12(1).Value = vbUnchecked        'exclude feed spots
                '3-27-11 ask billed, unbilled, or both
                plcSelC9.Move 120, plcSelC6.Top + plcSelC6.Height + 30
                rbcSelC9(0).Caption = "Billed"
                rbcSelC9(0).Move 600, 0, 960
                rbcSelC9(1).Caption = "Unbilled"
                rbcSelC9(1).Move 1560, 0, 1080
                rbcSelC9(2).Caption = "Both"
                rbcSelC9(2).Move 2640, 0, 720
                rbcSelC9(2).Value = True
                rbcSelC9(0).Visible = True
                rbcSelC9(1).Visible = True
                rbcSelC9(2).Visible = True
                plcSelC9.Visible = True
                smPlcSelC9P = "Show"
            
            Case CNT_SPTSBYDATETIME, CNT_MISSED          'Spots by times; Missed Spots ****
                mAirConvVehPop 8
                plcSelC2.Visible = False
                plcSelC3.Visible = False
                lbcSelection(6).Visible = True
                lbcSelection(3).Visible = False
                lbcSelection(0).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(5).Visible = False
                ckcAll.Caption = "All Vehicles"
                ckcAll.Visible = True
                '9-12-19 use csi calendar controls vs edit box
                lacSelCFrom.Move 120, 60, 1320
                lacSelCFrom.Caption = "Dates-Start"
                lacSelCFrom.Visible = True
                CSI_CalFrom.Move 1200, 30
                CSI_CalFrom.Visible = True
                lacSelCTo.Move 2415, lacSelCFrom.Top, 360
                lacSelCTo.Caption = "End"
                lacSelCTo.Visible = True
                CSI_CalTo.Move 2895, CSI_CalFrom.Top
                CSI_CalTo.Visible = True
                CSI_CalFrom.CSI_AllowBlankDate = True
                CSI_CalTo.CSI_AllowBlankDate = True

                lacSelCFrom1.Caption = "Times-Start"
                lacSelCFrom1.Move lacSelCFrom.Left, lacSelCFrom.Top + edcSelCFrom.Height + 30, 1200
                lacSelCFrom1.Visible = True
                edcSet1.Text = "12M"
                edcSet1.Move CSI_CalFrom.Left, CSI_CalFrom.Top + CSI_CalFrom.Height + 60, 960
                edcSet1.Visible = True
                edcSet1.MaxLength = 10

                lacSelCTo1.Caption = "End"
                lacSelCTo1.Move lacSelCTo.Left, lacSelCFrom1.Top, 360
                lacSelCTo1.Visible = True
                edcSet2.Text = "12M"
                edcSet2.Move CSI_CalTo.Left, edcSet1.Top, 960
                edcSet2.Visible = True
                edcSet2.MaxLength = 10

                lacContract.Caption = "Contract #"
                lacContract.Move lacSelCFrom.Left, lacSelCFrom1.Top + edcSelCFrom1.Height + 15, 975
                lacContract.Visible = True
                edcSet3.Move CSI_CalFrom.Left, edcSelCFrom1.Top + edcSelCFrom1.Height + 365, 960
                edcSet3.Visible = True
                edcSet3.MaxLength = 8

                'plcSelC1.Caption = "Show Spot Prices"
                smPlcSelC1P = "Show Spot Prices"
                plcSelC1.Move lacSelCFrom.Left, lacContract.Top + lacContract.Height + 115
                rbcSelCSelect(0).Caption = "Yes"
                rbcSelCSelect(0).Left = 1470
                rbcSelCSelect(0).Width = 625
                If rbcSelCSelect(0).Value Then
                    rbcSelCSelect_click 0
                Else
                    rbcSelCSelect(0).Value = True
                End If
                rbcSelCSelect(1).Caption = "No"
                rbcSelCSelect(1).Left = 2070
                rbcSelCSelect(1).Width = 615
                rbcSelCSelect(1).Enabled = True
                rbcSelCSelect(2).Enabled = True
                rbcSelCSelect(2).Visible = False

                If ilListIndex = CNT_MISSED Then
                    mSPersonPop lbcSelection(2)         '2-16-06
                    
                    '4-6-20 contract type selectivity
                    plcSelC10.Move 120, plcSelC1.Top + plcSelC1.Height, 4260
                    smPlcSelC10P = "Include"
                    ckcSelC10(0).Caption = "Holds"
                    ckcSelC10(0).Move 720, -30, 840
                    ckcSelC10(0).Value = vbChecked   'True
                    If ckcSelC10(0).Value = vbChecked Then
                        ckcSelC10_click 0
                    Else
                        ckcSelC10(0).Value = vbChecked   'True
                    End If
                    ckcSelC10(0).Visible = True

                    ckcSelC10(1).Value = vbChecked   'True
                    ckcSelC10(1).Caption = "Orders"
                    ckcSelC10(1).Move 1620, -30, 900
                    If ckcSelC10(1).Value = vbChecked Then
                        ckcSelC10_click 1
                    Else
                        ckcSelC10(1).Value = vbChecked   'True
                    End If
                    ckcSelC10(1).Visible = True
                    plcSelC10.Visible = True
                    plcSelC10_Paint
                    '12-28-17
                    mAskContractTypesCkcSelC6 plcSelC10.Top + plcSelC10.Height
                    plcSelC3.Left = 120
                    plcSelC3.Top = plcSelC6.Top
                    smPlcSelC3P = "Include"
                    plcSelC3.Height = 240
                    plcSelC3.Move 120, plcSelC6.Top + plcSelC6.Height + 30, 4000
                    ckcSelC3(0).Left = 720  '675
                    ckcSelC3(0).Width = 1020
                    ckcSelC3(0).Caption = "Missed"
                    If ckcSelC3(0).Value = vbChecked Then
                        ckcSelC3_click 0
                    Else
                        ckcSelC3(0).Value = vbChecked ' = True
                    End If
                    ckcSelC3(0).Visible = True
                    ckcSelC3(1).Left = 1815 '1695
                    ckcSelC3(1).Width = 1180
                    ckcSelC3(1).Caption = "Cancelled"
                   If ckcSelC3(1).Value = vbChecked Then
                        ckcSelC3_click 1
                    Else
                        ckcSelC3(1).Value = vbChecked ' = True
                    End If
                    ckcSelC3(1).Visible = True
                    ckcSelC3(2).Left = 3075 '2835
                    ckcSelC3(2).Width = 1020
                    ckcSelC3(2).Caption = "Hidden"
                    If ckcSelC3(2).Value = vbChecked Then
                       ckcSelC3_click 2
                   Else
                       ckcSelC3(2).Value = vbChecked ' = True
                  End If
                  
                    ckcSelC3(2).Visible = True
                    ckcSelC3(3).Visible = False
                    ckcSelC3(4).Visible = False
                    ckcSelC3(5).Visible = False
                    ckcSelC3(6).Visible = False
                    plcSelC3.Visible = True
                    plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height
                    plcSelC5.Height = 655                   '3 lines of check boxes
                    mChooseSpotType

                    plcSelC7.Move 120, plcSelC5.Top + plcSelC5.Height + 30
                    smPlcSelC7P = "Sort by"
                    rbcSelC7(0).Caption = "Vehicle"
                    rbcSelC7(1).Caption = "Salesperson"
                    rbcSelC7(0).Move 720, 0, 960
                    rbcSelC7(1).Move 1680, 0, 1440
                    rbcSelC7(0).Visible = True
                    rbcSelC7(1).Visible = True
                    rbcSelC7(2).Visible = False
                    rbcSelC7(0).Value = True

                    plcSelC7.Visible = True

                    plcSelC9.Move plcSelC7.Left, plcSelC7.Top + plcSelC7.Height
                    smPlcSelC9P = "Show"
                    rbcSelC9(0).Caption = "Gross"
                    rbcSelC9(1).Caption = "Net"
                    rbcSelC9(0).Move 600, 0, 960
                    rbcSelC9(1).Move 1560, 0, 1920
                    rbcSelC9(0).Visible = True
                    rbcSelC9(1).Visible = True
                    rbcSelC9(2).Visible = False
                    rbcSelC9(0).Value = True
                    plcSelC9.Visible = True

                    '4-6-20 Use different control, ckcSelC10 needed for hold/order selection
                    ckcSelC15.Move 120, plcSelC9.Top + plcSelC9.Height
                    ckcSelC15.Caption = "Summary Only"
                    ckcSelC15.Visible = True

                    mAskCntrFeed plcSelC10.Top + plcSelC10.Height     '2-16-06 chg from plc 5 to 11 ask contracts spots and feed spots selectiviy
                    If rbcSelC14(0).Value Then               'selling vs airing option not enabled with Missed spots, default to selling
                        rbcSelC14_Click 0
                    Else
                        rbcSelC14(0).Value = True
                    End If
                Else                     'spots by date & time
                    '8-24-01 if showing spot prices, option to show rates using the full spot rate or just cash portion for trades
                    'plcSelC4.Caption = "For trades- "
                    smPlcSelC4P = "For trades- "
                    plcSelC4.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height + 80
                    rbcSelC4(0).Caption = "Show full price"
                    rbcSelC4(0).Move 960, 0, 1800
                    rbcSelC4(1).Caption = "Cash only portion"
                    rbcSelC4(1).Move 2520, 0, 2520
                    rbcSelC4(0).Visible = True
                    rbcSelC4(1).Visible = True
                    rbcSelC4(2).Visible = False
                    rbcSelC4(0).Enabled = True
                    rbcSelC4(1).Enabled = True
                    rbcSelC4(2).Enabled = False
                    If rbcSelC4(0).Value Then               'set default to show full price
                        rbcSelC4_click 0
                    Else
                        rbcSelC4(0).Value = True
                    End If
                    plcSelC4.Visible = True
                    
                    '12-28-17 contract type selectivity
                    plcSelC10.Move 120, plcSelC4.Top + plcSelC4.Height, 4260
                    smPlcSelC10P = "Include"
                    ckcSelC10(0).Caption = "Holds"
                    ckcSelC10(0).Move 720, -30, 840
                    ckcSelC10(0).Value = vbChecked   'True
                    If ckcSelC10(0).Value = vbChecked Then
                        ckcSelC10_click 0
                    Else
                        ckcSelC10(0).Value = vbChecked   'True
                    End If
                    ckcSelC10(0).Visible = True
                    
                    ckcSelC10(1).Value = vbChecked   'True
                    ckcSelC10(1).Caption = "Orders"
                    ckcSelC10(1).Move 1620, -30, 900
                    If ckcSelC10(1).Value = vbChecked Then
                        ckcSelC10_click 1
                    Else
                        ckcSelC10(1).Value = vbChecked   'True
                    End If
                    ckcSelC10(1).Visible = True
                    plcSelC10.Visible = True
                    plcSelC10_Paint
                    '12-28-17
                    mAskContractTypesCkcSelC6 plcSelC10.Top + plcSelC10.Height
                    'plcSelC5.Move 120, plcSelC4.Top + plcSelC4.Height
                    plcSelC5.Move 120, plcSelC6.Top + plcSelC6.Height
                    plcSelC5.Height = 655                   '3 lines of check boxes
                    mChooseSpotType
                    mAskCntrFeed plcSelC5.Top + plcSelC5.Height     'ask contracts spots and feed spots selectiviy

                    '5-25-06 feature to calc gross or net
                    ilTop = plcSelC5.Top + plcSelC5.Height
                    If tgSpf.sSystemType = "R" Then     'if station, extra question asked
                        ilTop = plcSelC12.Top + plcSelC12 + Height
                    End If
                    plcSelC9.Move 120, ilTop + 30
                    smPlcSelC9P = "Show"
                    rbcSelC9(0).Caption = "Gross"
                    rbcSelC9(1).Caption = "Net"
                    rbcSelC9(0).Move 600, 0, 960
                    rbcSelC9(1).Move 1560, 0, 840
                    rbcSelC9(0).Visible = True
                    rbcSelC9(1).Visible = True
                    rbcSelC9(2).Visible = False
                    rbcSelC9(0).Value = True
                    plcSelC9.Visible = True
                    
                    ckcSelC15.Move 2640, plcSelC9.Top
                    '1-3-18 summary only option
                    ckcSelC15.Visible = True

                    If imHideProposalPrice Then
                        rbcSelC11(0).Value = True
                        ilTop = plcSelC9.Top + plcSelC9.Height + 30

                    Else
                        plcSelC11.Move plcSelC9.Left, plcSelC9.Top + plcSelC9.Height + 30
                        rbcSelC11(0).Move 960, 0, 960
                        rbcSelC11(0).Caption = "Actual"
                        rbcSelC11(1).Move 1920, 0, 1200
                        rbcSelC11(1).Caption = "Rate Card"
                        rbcSelC11(0).Value = True
                        rbcSelC11(0).Visible = True
                        rbcSelC11(1).Visible = True
                        smPlcSelC11P = "Use Price"
                        plcSelC11.Visible = True
                        ilTop = plcSelC11.Top + plcSelC11.Height + 30
                    End If
                    
                    '4-20-06 If sports used, ask to skip to newpage each game
                    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
                    If (ilValue And &H1) = USINGSPORTS Then 'Using Sports
                        'ilTop = plcSelC5.Top + plcSelC5.Height
                        'If tgSpf.sSystemType = "R" Then     'if station, extra question asked
                        '    ilTop = plcSelC12.Top + plcSelC12 + Height
                        'End If
                        If imHideProposalPrice Then
                            plcSelC13.Move 120, plcSelC9.Top + plcSelC9.Height '
                        Else                'ok to show proposal price question :  guide user
                            plcSelC13.Move 120, plcSelC11.Top + plcSelC11.Height
                        End If
                        ckcSelC13(0).Caption = "For sport vehicles, skip to new page each event"
                        ckcSelC13(0).Move 0, 0, 4300
                        ckcSelC13(0).Visible = True
                        ckcSelC13(1).Visible = False
                        ckcSelC13(2).Visible = False
                        plcSelC13.Visible = True
                        smPlcSelC13P = ""
                        plcSelC13_Paint
                        ilTop = plcSelC13.Top + plcSelC13.Height + 30
                    End If

                    '8-31-15 in addition to standard vehicles (conventional, game), use selling or airing vehicles for spot data
                    plcSelC14.Top = ilTop
                    smPlcSelC14P = "Use"
                    rbcSelC14(0).Caption = "Selling"
                    rbcSelC14(0).Move 480, 0, 960
                    rbcSelC14(1).Caption = "Airing (Log)"
                    rbcSelC14(1).Move 1440, 0, 1800
                    rbcSelC14(0).Visible = True
                    rbcSelC14(0).Value = True
                    rbcSelC14(1).Visible = True
                    plcSelC14.Visible = True
                    cbcSet1.Move 120, plcSelC14.Top + plcSelC14.Height + 30, 2000
                    cbcSet1.AddItem "All Audio Types"
                    cbcSet1.AddItem "Live Coml"
                    cbcSet1.AddItem "Live Promo"
                    cbcSet1.AddItem "Pre-Recorded Coml"
                    cbcSet1.AddItem "Pre-Recorded Promo"
                    cbcSet1.AddItem "Recorded Coml"
                    cbcSet1.AddItem "Recorded Promo"
                    cbcSet1.ListIndex = 0
                    cbcSet1.Visible = True
                End If
                plcSelC1.Visible = True
                pbcSelC.Height = 4430           '12-28-17
                pbcSelC.Visible = True
                pbcOption.Visible = True
                
            Case CNT_PLACEMENT  'Placement
                plcSelC2.Visible = False
                plcSelC1.Visible = False
                plcSelC3.Visible = False
                lacSelCTo.Visible = False
                edcSelCTo.Visible = False
                lbcSelection(3).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(0).Visible = True
                lbcSelection(5).Visible = True
                ckcAll.Caption = "All Advertisers"
                ckcAll.Visible = True
                lacSelCFrom.Caption = "Dates-Start"
'                edcSelCFrom.Move 1110, edcSelCFrom.Top, 945
                CSI_CalFrom.Move 1170, edcSelCFrom.Top
                lacSelCFrom.Visible = True
'                edcSelCFrom.Visible = True
'                edcSelCFrom.MaxLength = 10  '8   5/27/99 changed for short form date m/d/yyyy
                CSI_CalFrom.Visible = True      '9-11-19 use csi calendar control vs edit box
                CSI_CalFrom.CSI_AllowBlankDate = True
                lacSelCFrom1.Caption = "End"
                lacSelCFrom1.Move 2445, lacSelCFrom.Top, 360
'                edcSelCFrom1.Move 3120, edcSelCFrom1.Top, 945       '3120
                CSI_CalTo.Move 2880, CSI_CalFrom.Top
                CSI_CalTo.CSI_AllowBlankDate = True
                lacSelCFrom1.Visible = True
'                edcSelCFrom1.Visible = True
'                edcSelCFrom1.MaxLength = 10 '8   5/27/99 changed for short form date m/d/yyyy
                CSI_CalTo.Visible = True
'                plcSelC3.Move 0, edcSelCFrom.Top + edcSelCFrom.Height + 30, 4380, 410
                plcSelC3.Move 0, CSI_CalFrom.Top + CSI_CalFrom.Height + 30, 4380, 480
                smPlcSelC3P = ""
                ckcSelC3(0).Caption = "Show spots outside of date range w/ date error"
                ckcSelC3(0).Move 120, 0, 4380
                ckcSelC3(0).Visible = True
                ckcSelC3(1).Caption = "Include Fill Spots"
                ckcSelC3(1).Move 120, 270, 2400
                ckcSelC3(1).Value = vbChecked       'default to include all spots
                ckcSelC3(1).Visible = True
                plcSelC3.Visible = True
                pbcSelC.Visible = True
                pbcOption.Visible = True
                
            Case CNT_DISCREP  'Discrepancies
                plcSelC2.Visible = False
                plcSelC1.Visible = False
                plcSelC3.Visible = False
                lbcSelection(3).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(0).Visible = True
                lbcSelection(5).Visible = True
                ckcAll.Caption = "All Advertisers"
                ckcAll.Visible = True
                If ckcAll.Value = vbChecked Then
                    ckcAll_Click
                Else
                    ckcAll.Value = vbChecked
                End If
                lacSelCFrom.Caption = "Dates-Start"
                CSI_CalFrom.Move 1170, edcSelCFrom.Top
                lacSelCFrom.Visible = True
'                edcSelCFrom.Visible = True
'                edcSelCFrom.MaxLength = 10  '8   5/27/99 changed for short form date m/d/yyyy
                CSI_CalFrom.Visible = True      '9-11-19 use csi calendar control vs edit box
                CSI_CalFrom.CSI_AllowBlankDate = True
                lacSelCFrom1.Caption = "End"
                lacSelCFrom1.Move 2445, lacSelCFrom.Top, 360
'                edcSelCFrom1.Move 3120, edcSelCFrom1.Top, 945       '3120
                CSI_CalTo.Move 2880, CSI_CalFrom.Top
                CSI_CalTo.CSI_AllowBlankDate = True
                lacSelCFrom1.Visible = True
'                edcSelCFrom1.Visible = True
'                edcSelCFrom1.MaxLength = 10 '8   5/27/99 changed for short form date m/d/yyyy
                CSI_CalTo.Visible = True
'                    edcSelCTo.Move 990, edcSelCTo.Top, 945
                lacSelCTo.Move 120, CSI_CalFrom.Top + CSI_CalFrom.Height + 120
                edcSelCTo.Move CSI_CalFrom.Left, lacSelCTo.Top, 945
                lacSelCTo.Caption = "Starting #"
                If lgOrigCntrNo > 0 Then
                    edcSelCTo.Text = Trim$(str$(lgOrigCntrNo))
                Else
                    edcSelCTo.Text = ""
                End If
                If ckcAll.Value = vbChecked Then
                    lacSelCTo.Visible = True
                    edcSelCTo.Visible = True
                Else
                    lacSelCTo.Visible = False
                    edcSelCTo.Visible = False
                End If
                plcSelC3.Move 0, edcSelCTo.Top + edcSelCTo.Height + 60, 4380
                smPlcSelC3P = ""
                ckcSelC3(0).Caption = "Show spots outside of date range w/ date error"
                ckcSelC3(0).Move 120, 0, 4380
                ckcSelC3(0).Visible = True
                plcSelC3.Visible = True
                pbcSelC.Visible = True
                pbcOption.Visible = True
                
            Case CNT_SPOTSALES  'Spot sales by vehicle or advertiser
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(0).Visible = False
                lbcSelection(5).Visible = False
                ckcAll.Caption = "All Vehicles"
                ckcAll.Visible = True
                '9-11-19 use csi calendar control vs edit box
                lacSelCFrom.Caption = "Dates-Start"
'                edcSelCFrom.Move 990, edcSelCFrom.Top, 945
                CSI_CalFrom.Move 1170, edcSelCFrom.Top
                lacSelCFrom.Visible = True
'                edcSelCFrom.Visible = True
                CSI_CalFrom.Visible = True
'                edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
                lacSelCFrom1.Caption = "End"
'                lacSelCFrom1.Move 2280, edcSelCFrom.Top + 30, 945
                lacSelCFrom1.Move 2445, CSI_CalFrom.Top + 30, 360
'                edcSelCFrom1.Move 3120, edcSelCFrom.Top, 945
                CSI_CalTo.Move 2880, CSI_CalFrom.Top
                lacSelCFrom1.Visible = True
'                edcSelCFrom1.Visible = True
                CSI_CalTo.Visible = True
'                edcSelCFrom1.MaxLength = 10 '8  5/27/99 changed for short form date m/d/yyyy
                edcSelCTo.Text = "12M"
                edcSelCTo1.Text = "12M"
                lacSelCTo.Caption = "Times-Start"
                lacSelCTo.Move lacSelCFrom.Left, CSI_CalFrom.Top + edcSelCFrom.Height + 60, 1200
                edcSelCTo.Move CSI_CalFrom.Left, CSI_CalFrom.Top + CSI_CalFrom.Height + 90, 945
                lacSelCTo.Visible = True
                edcSelCTo.Visible = True
                edcSelCTo.MaxLength = 10    '8 5/27/99 changed for short form date m/d/yyyy
                lacSelCTo1.Caption = "End"
                lacSelCTo1.Move lacSelCFrom1.Left, edcSelCTo.Top + 30, 945
                edcSelCTo1.Move CSI_CalTo.Left, edcSelCTo.Top, 945
                lacSelCTo1.Visible = True
                edcSelCTo1.Visible = True
                edcSelCTo1.MaxLength = 10   '8 5/27/99 changed for short form date m/d/yyyy
                plcSelC8.Top = edcSelCTo.Top + edcSelCTo.Height + 60
                mAskDaysOfWk
                'plcSelC1.Top = edcSelCTo.Top + edcSelCTo.Height + 60
                plcSelC1.Top = plcSelC8.Top + plcSelC8.Height + 60
                smPlcSelC1P = "Subtotals"
                rbcSelCSelect(0).Caption = "None"
                rbcSelCSelect(0).Left = 840
                rbcSelCSelect(0).Width = 735
                plcSelC1.Height = 435           '2 lines of subtotal options
                If rbcSelCSelect(0).Value Then
                    rbcSelCSelect_click 0
                Else
                    rbcSelCSelect(0).Value = True
                End If
                rbcSelCSelect(1).Caption = "Date"
                rbcSelCSelect(1).Left = 1590
                rbcSelCSelect(1).Width = 705
                rbcSelCSelect(1).Visible = True
                rbcSelCSelect(1).Enabled = True
                rbcSelCSelect(2).Caption = "Advertiser"
                rbcSelCSelect(2).Left = 2300
                rbcSelCSelect(2).Width = 1300
                rbcSelCSelect(2).Visible = True
                rbcSelCSelect(2).Enabled = True
                rbcSelCSelect(3).Caption = "Sales Source"
                rbcSelCSelect(3).Move 840, 195, 2540
                rbcSelCSelect(3).Visible = True
                'plcSelC2.Top = edcSelCFrom.Top + edcSelCFrom.Height + 120
                plcSelC2.Top = plcSelC1.Top + plcSelC1.Height
                'plcSelC2.Caption = "Show"
                smPlcSelC2P = "Show"
                rbcSelCInclude(0).Caption = "Net"
                rbcSelCInclude(0).Left = 840
                rbcSelCInclude(0).Width = 705
                If rbcSelCInclude(0).Value Then
                    rbcSelCInclude_Click 0
                Else
                    rbcSelCInclude(0).Value = True
                End If
                rbcSelCInclude(1).Caption = "Net-Net"
                rbcSelCInclude(1).Left = 1590
                rbcSelCInclude(1).Width = 1100
                rbcSelCInclude(1).Visible = True
                rbcSelCInclude(1).Enabled = True
                rbcSelCInclude(2).Visible = False
                plcSelC2.Visible = True
                plcSelC3.Left = 120
                'plcSelC3.Top = plcSelC1.Top + plcSelC1.Height
                plcSelC3.Top = plcSelC2.Top + plcSelC2.Height
                plcSelC3.Height = 240
                'plcSelC3.Caption = "Include"
                smPlcSelC3P = "Include"
                plcSelC7.Top = plcSelC2.Top + plcSelC2.Height
                plcSelC7.Left = 120
                plcSelC7.Visible = True                         'as aired or ordered
                'plcSelC7.Caption = "For"
                smPlcSelC7P = "For"
                rbcSelC7(0).Visible = True
                rbcSelC7(0).Caption = "Ordered"
                rbcSelC7(0).Left = 360
                rbcSelC7(0).Width = 1200
                rbcSelC7(1).Visible = True
                rbcSelC7(1).Caption = "Aired"
                rbcSelC7(1).Left = 1400
                rbcSelC7(1).Width = 960
                rbcSelC7(2).Caption = "As Aired/Pkg Ordered"
                rbcSelC7(2).Move 2180, 0, 2280
                rbcSelC7(2).Visible = True
                If tgSpf.sInvAirOrder = "A" Then            'default as aired
                    If rbcSelC7(1).Value Then
                        rbcSelC7_click 1
                    Else
                        rbcSelC7(1).Value = True
                    End If
                Else
                    If rbcSelC7(0).Value Then
                        rbcSelC7_click 0
                    Else
                        rbcSelC7(0).Value = True
                    End If
                End If
                'Contract Type selection
                '12-28-17 ask contract types, change to subroutine
                smPlcSelC6P = "Include"
                mAskContractTypesCkcSelC6 plcSelC7.Top + plcSelC7.Height
                plcSelC3.Left = 120                'plcSelC3.Top = plcSelC1.Top + plcSelC1.Height
                plcSelC3.Top = plcSelC6.Top + plcSelC6.Height
                plcSelC3.Width = 4400
                plcSelC3.Height = 240
                'plcSelC3.Caption = "Include"
                'plcSelC3.Caption = ""
                smPlcSelC3P = ""
                ckcSelC3(0).Value = vbChecked
                ckcSelC3(0).Left = 720
                ckcSelC3(0).Width = 1020
                ckcSelC3(0).Caption = "Missed"
                If ckcSelC3(0).Value = vbUnchecked Then
                    ckcSelC3_click 0
                Else
                    ckcSelC3(0).Value = vbUnchecked
                End If
                'ckcSelC3(0).Value = False
                ckcSelC3(0).Visible = True
                ckcSelC3(1).Left = 1695
                ckcSelC3(1).Width = 1180
                ckcSelC3(1).Caption = "Cancelled"
                'ckcSelC3(1).Value = False
                ckcSelC3(1).Visible = True
                ckcSelC3(2).Left = 2835
                ckcSelC3(2).Width = 1280
                ckcSelC3(2).Caption = "Hidden"
                'ckcSelC3(2).Value = False
                ckcSelC3(2).Visible = True
                ckcSelC3(3).Visible = False
                ckcSelC3(4).Visible = False
                ckcSelC3(5).Visible = False
                ckcSelC3(6).Visible = False
                plcSelC3.Visible = True
                plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height
                plcSelC5.Height = 655                   '3 lines of check boxes
                mChooseSpotType

                plcSelC4.Move 120, plcSelC5.Top + plcSelC5.Height
                'plcSelC4.Caption = "Show"
                smPlcSelC4P = "Show"
                rbcSelC4(0).Left = 720
                rbcSelC4(0).Caption = "Spot counts"
                rbcSelC4(0).Width = 1440
                rbcSelC4(0).Visible = True
                rbcSelC4(1).Left = 2160
                rbcSelC4(1).Width = 1440
                rbcSelC4(1).Caption = "Unit counts"
                rbcSelC4(1).Visible = True
                If rbcSelC4(1).Value Then
                    rbcSelC4_click 1
                Else
                    rbcSelC4(1).Value = True
                End If
                rbcSelC4(2).Visible = False
                plcSelC4.Visible = True        'True
                lacContract.Move 120, plcSelC4.Top + plcSelC4.Height + 30
                edcSet1.Move 1200, plcSelC4.Top + plcSelC4.Height
                edcSet1.Text = ""
                lacContract.Visible = True
                edcSet1.Visible = True
                ckcSelC15.Move lacContract.Left, edcSet1.Top + edcSet1.Height + 30, 4000
                ckcSelC15.Caption = "Use Prop Price (vs Actual Price)"            '12-5-18
                ckcSelC15.Visible = True
                lbcSelection(3).Visible = True
                plcSelC1.Visible = True
                pbcSelC.Height = 4300
                pbcSelC.Visible = True
                pbcOption.Visible = True
                
            Case CNT_ACCRUEDEFER            '12-20-06 Accrual/Deferral
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(0).Visible = False
                lbcSelection(5).Visible = False
                mSellConvRepVehPop 3

                lacSelCFrom.Caption = "Dates-Start"
                CSI_CalFrom.Move 1110, edcSelCFrom.Top
                lacSelCFrom.Visible = True
'                edcSelCFrom.Visible = True
'                edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
                CSI_CalFrom.Visible = True
                lacSelCFrom1.Caption = "End"
                lacSelCFrom1.Move 2280, edcSelCFrom.Top + 30, 945
'                edcSelCFrom1.Move 3120, edcSelCFrom.Top, 945
                CSI_CalTo.Move 2760, CSI_CalFrom.Top
                lacSelCFrom1.Visible = True
'                edcSelCFrom1.Visible = True
'                edcSelCFrom1.MaxLength = 10 '8  5/27/99 changed for short form date m/d/yyyy
                CSI_CalTo.Visible = True
'                plcSelC8.Top = edcSelCFrom.Top + edcSelCFrom.Height + 60
                plcSelC8.Top = CSI_CalFrom.Top + CSI_CalFrom.Height + 60
                mAskDaysOfWk

                'plcSelC1.Top = edcSelCTo.Top + edcSelCTo.Height + 60
                plcSelC1.Top = plcSelC8.Top + plcSelC8.Height + 60
                smPlcSelC1P = "Sort by- "
                rbcSelCSelect(0).Caption = "Sales Source"
                rbcSelCSelect(0).Left = 960
                rbcSelCSelect(0).Width = 1440
                plcSelC1.Height = 435           '2 lines of subtotal options
                If rbcSelCSelect(0).Value Then
                    rbcSelCSelect_click 0
                Else
                    rbcSelCSelect(0).Value = True
                End If
                rbcSelCSelect(1).Caption = "Sales Origin"
                rbcSelCSelect(1).Move 2520, 0, 1680

                rbcSelCSelect(1).Visible = True
                rbcSelCSelect(1).Enabled = True
                rbcSelCSelect(2).Caption = "Vehicle"
                rbcSelCSelect(2).Move 960, 195, 960
                rbcSelCSelect(2).Visible = True
                rbcSelCSelect(2).Enabled = True
                rbcSelCSelect(3).Caption = "Vehicle"        'unused
                rbcSelCSelect(3).Move 840, 195, 2540
                rbcSelCSelect(3).Visible = False
                'vehicle group selection
                gPopVehicleGroups RptSelCb!cbcSet1, tgVehicleSets1(), True
                lacSelCTo.Caption = "Vehicle Group"
                lacSelCTo.Move 120, plcSelC1.Top + plcSelC1.Height + 30, 1680
                cbcSet1.Move 1440, lacSelCTo.Top - 15, 1140
                lacSelCTo.Visible = True
                cbcSet1.Visible = True
                lacContract.Move 2700, lacSelCTo.Top
                lacContract.Visible = True
                edcSelCTo.Move 3660, cbcSet1.Top, 810           'selective contract #
                edcSelCTo.Visible = True

                plcSelC10.Move 120, cbcSet1.Top + cbcSet1.Height
                ckcSelC10(0).Caption = "Summary Only"
                ckcSelC10(0).Move 0, 0, 2160
                ckcSelC10(0).Visible = True
                plcSelC10.Visible = True

                'Contract Type selection
                plcSelC6.Move 120, plcSelC10.Top + plcSelC10.Height, 4260
                plcSelC6.Height = 440
                smPlcSelC6P = "Include"
                ckcSelC6(0).Move 720, -30, 1080
                ckcSelC6(0).Caption = "Standard"
                If ckcSelC6(0).Value = vbChecked Then
                    ckcSelC6_click 0
                Else
                    ckcSelC6(0).Value = vbChecked
                End If
                ckcSelC6(0).Visible = True
                ckcSelC6(1).Move 1800, -30, 1200
                ckcSelC6(1).Caption = "Reserved"
                If ckcSelC6(1).Value = vbChecked Then
                    ckcSelC6_click 1
                Else
                    ckcSelC6(1).Value = vbChecked
                End If
                ckcSelC6(1).Visible = True
                If tgUrf(0).iSlfCode > 0 Then           'its a slsp thats is asking for this report,
                                                        'don't allow them to exclude reserves
                    ckcSelC6(1).Enabled = False
                Else
                    ckcSelC6(1).Enabled = True
                End If
                ckcSelC6(2).Move 3000, -30, 1080
                ckcSelC6(2).Caption = "Remnant"
                If ckcSelC6(2).Value = vbChecked Then
                    ckcSelC6_click 2
                Else
                    ckcSelC6(2).Value = vbChecked
                End If
                ckcSelC6(2).Visible = True
                ckcSelC6(3).Move 720, 195, 600
                ckcSelC6(3).Caption = "DR"
                If ckcSelC6(3).Value = vbChecked Then
                    ckcSelC6_click 3
                Else
                    ckcSelC6(3).Value = vbChecked   'True
                End If
                ckcSelC6(3).Visible = True
                ckcSelC6(4).Move 1380, 195, 1320
                ckcSelC6(4).Caption = "Per Inquiry"
                If ckcSelC6(4).Value = vbChecked Then
                    ckcSelC6_click 4
                Else
                    ckcSelC6(4).Value = vbChecked
                End If
                ckcSelC6(4).Visible = True
                ckcSelC6(5).Move 2700, 195, 720
                ckcSelC6(5).Caption = "PSA"
                ckcSelC6(5).Visible = True
                plcSelC6.Visible = True
                ckcSelC6(6).Move 3420, 195, 960
                ckcSelC6(6).Caption = "Promo"
                ckcSelC6(6).Visible = True

                plcSelC3.Move 120, plcSelC6.Top + plcSelC6.Height, 4400, 435
                smPlcSelC3P = ""
                ckcSelC3(0).Value = vbChecked
                ckcSelC3(0).Move 720, -30, 1020
                ckcSelC3(0).Caption = "Air Time"
                ckcSelC3(0).Value = vbChecked

                ckcSelC3(0).Visible = True
                ckcSelC3(1).Move 1740, -30, 720
                ckcSelC3(1).Caption = "Rep"
                ckcSelC3(1).Visible = True
                ckcSelC3(1).Value = vbChecked

                ckcSelC3(2).Move 2460, -30, 720
                ckcSelC3(2).Caption = "NTR"
                ckcSelC3(2).Visible = True
                ckcSelC3(2).Value = vbChecked

                ckcSelC3(3).Move 3180, -30, 1200
                ckcSelC3(3).Caption = "Hardcost"
                ckcSelC3(3).Visible = True

                ckcSelC3(4).Move 720, 195, 720
                ckcSelC3(4).Caption = "Polit"
                ckcSelC3(4).Visible = True
                ckcSelC3(4).Value = vbChecked

                ckcSelC3(5).Move 1440, 195, 1320
                ckcSelC3(5).Caption = "Non-Polit"
                ckcSelC3(5).Visible = True
                ckcSelC3(5).Value = vbChecked
                ckcSelC3(6).Visible = False
                plcSelC3.Visible = True

                plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height + 30
                plcSelC5.Height = 655   'chg from 2 to 3 lines of check boxes- 435  2 lines of check boxes
                mChooseSpotType                 'charge, n/c, recapturable, etc

                plcSelC4.Move 120, plcSelC5.Top + plcSelC5.Height
                smPlcSelC4P = "Show"
                rbcSelC4(0).Left = 720
                rbcSelC4(0).Caption = "Spot counts"
                rbcSelC4(0).Width = 1440
                rbcSelC4(0).Visible = True
                rbcSelC4(1).Left = 2160
                rbcSelC4(1).Width = 1440
                rbcSelC4(1).Caption = "Unit counts"
                rbcSelC4(1).Visible = True
                If rbcSelC4(0).Value Then
                    rbcSelC4_click 0
                Else
                    rbcSelC4(0).Value = True
                End If
                rbcSelC4(2).Visible = False
                plcSelC4.Visible = False            'take out option to allow unit counts, always default to spot counts

                plcSelC12.Move 120, plcSelC5.Top + plcSelC5.Height
                ckcSelC12(0).Caption = "Show spot counts"       'default to NO
                ckcSelC12(0).Visible = True
                ckcSelC12(0).Move 0, 0, 1920
                plcSelC12.Visible = True
                'lbcselection(3) = vehicles
                'lbcselection(4) = sales sources
                'lbcselection(7) = vehicle group items
                ckcAll.Caption = "All Vehicles"
                lbcSelection(3).Height = (lbcSelection(3).Height / 2) '- 240
                ckcAll.Visible = True
                ckcAllAAS.Caption = "All Sales Sources"
                ckcAllAAS.Visible = True
                ckcAllAAS.Move 0, lbcSelection(3).Top + lbcSelection(3).Height + 30
                lbcSelection(4).Clear
                ilRet = gPopMnfPlusFieldsBox(RptSelCb, lbcSelection(4), tgMnfCodeCB(), smMnfCodeTag, "S")

'                lbcSelection(4).Move lbcSelection(3).Left, ckcAllAAS.Top + ckcAllAAS.Height, lbcSelection(4).Width / 2 - 240, 1500
                lbcSelection(4).Move lbcSelection(3).Left, ckcAllAAS.Top + ckcAllAAS.Height, lbcSelection(4).Width / 2 - 120, lbcSelection(3).Height
                lbcSelection(7).Move lbcSelection(4).Width + 240, lbcSelection(4).Top, lbcSelection(4).Width, lbcSelection(3).Height
                CkcAllveh.Move lbcSelection(7).Left, ckcAllAAS.Top
                lbcSelection(3).Visible = True  'vehicles
                lbcSelection(4).Visible = True  'sales sources
                plcSelC1.Visible = True
                pbcSelC.Visible = True
                pbcOption.Visible = True
                
            Case CNT_HILORATE
                mHiLoSelectivity
                
            Case CNT_DISCREP_SUM                        '6-21-17
                mDiscrepSumSelectivity
                
        End Select


        frcOption.Visible = True
   ' End Select
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

    imHideProposalPrice = True
    '1-3-18 chg to also test for Guide sign in to be able to see Proposal price on spots by date and time
    If Trim$(slStr) = "CSI" Or Trim$(slStr) = "Guide" Then          'allow to see proposal in spots by date & time report
        imHideProposalPrice = False
    End If
    
    If igRptCallType <> GENERICBUTTON Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)
    Else
        igRptType = 0
    End If

    ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
    If (ilRet = CP_MSG_NONE) Then
        ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
        ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
        igRptCallType = Val(slStr)
    End If
    
    If (igRptCallType = CONTRACTSJOB) And (igRptType = 3) Then
        ilRet = gParseItem(slCommand, 5, "\", smLogUserCode)
        ilRet = gParseItem(slCommand, 6, "\", slStr)
        imVefCode = Val(slStr)
        ilRet = gParseItem(slCommand, 7, "\", smVehName)
        ilRet = gParseItem(slCommand, 8, "\", slStr)
        lmNoRecCreated = Val(slStr)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSellConvVehPop                 *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSellConvVehPop(ilIndex As Integer)
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVehPopErr
        gCPErrorMsg ilRet, "mSellConvVehPop (gPopUserVehicleBox: Vehicle)", RptSelCb
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mAirConvVehPop(ilIndex As Integer)
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVehPopErr
        gCPErrorMsg ilRet, "mSellConvVehPop (gPopUserVehicleBox: Vehicle)", RptSelCb
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
        'ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)   'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelCb
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
'*      Procedure Name:mSellConvVVPkgPop               *
'*                                                     *
'*             Created:10/31/96      By:D. Hosaka      *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box with conventional, virtual *
'*                      veh, package & airing          *
'*******************************************************
Private Sub mSellConvVVPkgPop(ilIndex As Integer, ilUselbcVehicle As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        'ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)   'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVVPkgPopErr
        gCPErrorMsg ilRet, "mSellConvVVPkgPop (gPopUserVehicleBox: Vehicle)", RptSelCb
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVVPkgPopErr:
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
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = CONTRACTSJOB Then
        If (igRptType = 0) And (ilListIndex > 1) Then
            ilListIndex = ilListIndex + 1
        End If
        If (ilListIndex = CNT_MGREVENUE Or ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_SPTCOMBO) Then
            If ckcAll.Value = vbChecked Then
                ilEnable = True
                If lbcSelection(6).SelCount <= 0 Then
                    ilEnable = False
                End If
            Else
                If rbcSelCSelect(0).Value Then                  'advt, get selective cnts
                    'Can't use SelCount as property does not exist for ListBoxbox
                    If ckcAll.Value = vbChecked Then
                        ilEnable = True
                    Else
                        For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                            If lbcSelection(0).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    End If
                End If
                'something must be selected for vehicles
                If ilEnable And Not (ckcAllAAS.Value = vbChecked) Then
                    ilEnable = False
                    'at least one vehicle must be selected
                    For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                        If lbcSelection(6).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
                If rbcSelCSelect(1).Value Then                    'agy
                    For illoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                        If lbcSelection(1).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                ElseIf rbcSelCSelect(2).Value Then               'slsp
                    For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                        If lbcSelection(2).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                ElseIf rbcSelCSelect(3).Value Then              'agy
                    For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                        If lbcSelection(6).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
            End If
            
        'ElseIf rbcRptType(3).Value Then
        ElseIf (ilListIndex = CNT_SPTSBYDATETIME) Then     '3=Spot by Time
            If RptSelCb!rbcSelC14(0).Value = True Then
                If lbcSelection(6).SelCount > 0 Then
                    ilEnable = True
                End If
            Else
                If lbcSelection(8).SelCount > 0 Then
                    ilEnable = True
                End If
            End If
        ElseIf (ilListIndex = CNT_MISSED) Then     '14=Missed Spots
            For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                If lbcSelection(6).Selected(illoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next illoop
        'ElseIf rbcRptType(4).Value Then
        ElseIf ilListIndex = CNT_PLACEMENT Then 'Placement
            If ckcAll.Value = vbChecked Then
                ilEnable = True
            Else
                For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                    If lbcSelection(0).Selected(illoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next illoop
            End If
        ElseIf ilListIndex = CNT_DISCREP Then 'Discrepanies
            If ckcAll.Value = vbChecked Then
                ilEnable = True
            Else
                For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                    If lbcSelection(0).Selected(illoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next illoop
            End If
        ElseIf ilListIndex = CNT_SPOTSALES Then 'Spot sales by vehicle or advertiser
'            If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
            If (CSI_CalFrom.Text <> "") And (CSI_CalTo.Text <> "") Then         '9-11-19 use csi calendar controls vs edit box
                ilEnable = False
                For illoop = 0 To lbcSelection(3).ListCount - 1 Step 1
                    If lbcSelection(3).Selected(illoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next illoop
            Else
                ilEnable = False
            End If
        ElseIf ilListIndex = CNT_ACCRUEDEFER Then       '12-20-06
            ilEnable = True
'            If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
            If (CSI_CalFrom.Text <> "") And (CSI_CalTo.Text <> "") Then         '9-11-19 use csi calendar controls vs edit box
                'at least one vehicle must be selected
                If (lbcSelection(3).SelCount = 0) Or (lbcSelection(4).SelCount = 0) Or (cbcSet1.ListIndex > 0 And lbcSelection(7).SelCount = 0) Then
                    ilEnable = False
                Else
                    'see if a vehicle group was selected

                End If
            Else
                ilEnable = False
            End If
         ElseIf ilListIndex = CNT_HILORATE Then      'Hi-lo Spot Rate 6-1-10
            ilEnable = True
'            If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
            If (CSI_CalFrom.Text <> "") And (edcSelCTo.Text <> "") Then     'use csi calendar control for date vs edit box
                'at least one vehicle must be selected
                If (lbcSelection(6).SelCount = 0) Then
                    ilEnable = False
                End If
            Else
                ilEnable = False
            End If
        ElseIf ilListIndex = CNT_DISCREP_SUM Then    'Spot Discrepancy Summary by Month 6-21-16
            ilEnable = True
            If edcSelCTo.Text = "" Then
                ilEnable = False
            End If
        Else
            ilEnable = False
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
        ElseIf rbcOutput(2).Value Then  'Save As
            If (imFTSelectedIndex >= 0) And (edcFileName.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        Else 'Export
            ilEnable = True
        End If
    End If
    If ilListIndex = CNT_SPTCOMBO Then    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
        rbcOutput(0).Top = 230
        rbcOutput(1).Top = 450
        rbcOutput(2).Top = 660
        rbcOutput(3).Top = 880
        rbcOutput(3).Visible = True
        If ckcSelSpots.Value = vbUnchecked And ckcSelDigital.Value = vbUnchecked Then
            ilEnable = False
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
'*            Comments: Populate Sales office list box *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mSPersonPop(lbcSelection As Control)
'
'   mSPersonPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopSalespersonBox(RptSelCb, 0, True, True, lbcSelection, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(RptSelCb, 0, True, True, lbcSelection, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gPopSalespersonBox)", RptSelCb
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
    Unload RptSelCb
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcSelC12_Paint()
    plcSelC12.CurrentX = 0
    plcSelC12.CurrentY = 0
    plcSelC12.Print smPlcSelC12P
End Sub

Private Sub plcSelC13_Paint()
    plcSelC13.CurrentX = 0
    plcSelC13.CurrentY = 0
    plcSelC13.Print "Billed"
End Sub

Private Sub plcSelC14_Paint()
    plcSelC14.CurrentX = 0
    plcSelC14.CurrentY = 0
    plcSelC14.Print smPlcSelC14P
End Sub

Private Sub rbcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    'End of coded added
    rbcSelC7(0).Enabled = True 'Gross
    rbcSelC7(1).Enabled = True 'Net
    If rbcOutput(Index).Value Then
        Select Case Index
            Case 0  'Display
                frcFile.Enabled = False
                frcCopies.Visible = False   'Print Box
                frcFile.Visible = False     'Save to File Box
                frcCopies.Enabled = False
                'frcWhen.Enabled = False
                'pbcWhen.Visible = False
                ckcSelDigitalComments.Enabled = False
                ckcSelDigitalComments.Value = vbUnchecked
                rbcSelC4(0).Enabled = True
                rbcSelC4(1).Enabled = True
            
            Case 1  'Print
                frcFile.Visible = False
                frcFile.Enabled = False
                frcCopies.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcCopies.Visible = True
                ckcSelDigitalComments.Enabled = False
                ckcSelDigitalComments.Value = vbUnchecked
                rbcSelC4(0).Enabled = True
                rbcSelC4(1).Enabled = True
            
            Case 2  'File
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcFile.Visible = True
                ckcSelDigitalComments.Enabled = False
                ckcSelDigitalComments.Value = vbUnchecked
                rbcSelC4(0).Enabled = True
                rbcSelC4(1).Enabled = True
            
            Case 3 ' Export  - TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
                If ckcSelDigital.Value = vbChecked Then
                    ckcSelDigitalComments.Enabled = True
                    chkIncludeAdjustments.Enabled = True
                Else
                    ckcSelDigitalComments.Enabled = False
                    ckcSelDigitalComments.Value = vbUnchecked
                    chkIncludeAdjustments.Enabled = False
                End If
                rbcSelC4(0).Enabled = False
                rbcSelC4(1).Enabled = False
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = False
                frcFile.Visible = False
                rbcSelC7(0).Enabled = False 'Gross
                rbcSelC7(1).Enabled = False 'Net
                
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

Private Sub rbcSelC11_Click(Index As Integer)
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB

    End Select
End Sub

Private Sub rbcSelC14_Click(Index As Integer)
    If Index = 0 Then
        lbcSelection(6).Visible = True  'conv & selling
        lbcSelection(8).Visible = False 'conv & airing
    Else
        lbcSelection(6).Visible = False  'conv & selling
        lbcSelection(8).Visible = True 'conv & airing
    End If
End Sub

Private Sub rbcSelC4_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC4(Index).Value
    'End of coded added
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If

    End Select
End Sub

Private Sub rbcSelC7_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC7(Index).Value
    'End of coded added
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If Value Then
        Select Case igRptCallType
            Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            If ilListIndex = CNT_SPOTSALES Then
                If Index = 0 Then                                'as ordered
                    ckcSelC3(0).Enabled = True
                    ckcSelC3(1).Enabled = True
                    ckcSelC3(2).Enabled = True
                    ckcSelC3(0).Value = vbChecked ' = True                'always assume missed is included
                    ckcSelC3(1).Value = vbUnchecked ' = False               'always assume cancel is excluded
                    ckcSelC3(2).Value = vbUnchecked ' = False               'always asume hidden is excluded
                Else
                    ckcSelC3(0).Enabled = True
                    ckcSelC3(1).Enabled = True
                    ckcSelC3(2).Enabled = True
                    ckcSelC3(0).Value = vbUnchecked ' = False                'always assume missed is excluded
                    ckcSelC3(1).Value = vbUnchecked ' = False               'always assume cancel is excluded
                    ckcSelC3(2).Value = vbUnchecked ' = False               'always asume hidden is excluded
                End If
            ElseIf ilListIndex = CNT_MISSED Then
                If Index = 0 Then               'vehicle option
                    lbcSelection(2).Visible = False
                    ckcAllAAS.Visible = False
                    lbcSelection(6).Height = 3270
                Else                            'slsp option
                    lbcSelection(6).Height = 1605
                    lbcSelection(2).Height = 1605
                    lbcSelection(2).Move 15, lbcSelection(6).Top + lbcSelection(6).Height + 315 'cut vehicle box horizontally in half
                    ckcAllAAS.Caption = "All Salespeople"
                    ckcAllAAS.Move 15, lbcSelection(2).Top - 270
                    ckcAllAAS.Visible = True
                    lbcSelection(2).Visible = True
                End If
            End If
        End Select
    End If
End Sub

Private Sub rbcSelC9_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC9(Index).Value
    'End of coded added
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If

    End Select
End Sub

Private Sub rbcSelCInclude_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelCInclude(Index).Value
    'End of coded added
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If Value Then
        Select Case igRptCallType
            Case CONTRACTSJOB
                If (igRptType = 0) And (ilListIndex > 1) Then
                    ilListIndex = ilListIndex + 1
                End If

                If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_SPTCOMBO Then
                    If Index = 1 Or Not gUsingBarters() Then           'do not show rates or not using barters, disable the acq feature
                        plcSelC8.Enabled = False
                        ckcSelC8(0).Enabled = False
                        ckcSelC8(0).Value = vbUnchecked
                    Else
                        plcSelC8.Enabled = True
                        ckcSelC8(0).Enabled = True
                    End If
                End If
        End Select
        mSetCommands
    End If
End Sub

Private Sub rbcSelCSelect_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelCSelect(Index).Value
    'End of coded added
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    If Value Then
        Select Case igRptCallType
            Case CONTRACTSJOB
                If (igRptType = 0) And (ilListIndex > 1) Then
                    ilListIndex = ilListIndex + 1
                End If
                If ilListIndex = 1 Or ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_MGREVENUE Or ilListIndex = CNT_SPTCOMBO Then 'spots by advt
                    Select Case Index
                        Case 0  'Advertiser/Contract #
                            lbcSelection(1).Visible = False
                            lbcSelection(2).Visible = False
                            lbcSelection(5).Visible = True
                            ckcAll.Caption = "All Advertisers"
                        
                        Case 1  'Agency
                            lbcSelection(0).Visible = False
                            lbcSelection(2).Visible = False
                            lbcSelection(5).Visible = False
                            lbcSelection(1).Visible = True
                            ckcAll.Caption = "All Agencies"
                        
                        Case 2  'Salesperson
                            lbcSelection(0).Visible = False
                            lbcSelection(1).Visible = False
                            lbcSelection(5).Visible = False
                            lbcSelection(2).Visible = True
                            ckcAll.Caption = "All Salespeople"
                        
                        Case 3  'vehicles
                            lbcSelection(0).Visible = False
                            lbcSelection(1).Visible = False
                            lbcSelection(5).Visible = False
                            lbcSelection(6).Visible = True
                            ckcAll.Caption = "All Vehicles"
                    End Select
                    
                    If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_MGREVENUE Or ilListIndex = CNT_SPTCOMBO Then
                        lbcSelection(0).Height = lbcSelection(3).Height / 2 '1605           'cnt list box
                        lbcSelection(5).Height = lbcSelection(3).Height / 2 '1605           'advt list box
                        lbcSelection(2).Height = lbcSelection(3).Height / 2 '1605           'slsp list box
                        lbcSelection(1).Height = lbcSelection(3).Height / 2 '1605           'agy list box
                        lbcSelection(6).Height = lbcSelection(3).Height / 2 '1560           'vehicle list box
                        lbcSelection(6).Visible = True          'show Vehicle list box
                        lbcSelection(6).Move 15, lbcSelection(5).Top + lbcSelection(5).Height + 315 'cut vehicle box horizontally in half
                        ckcAllAAS.Caption = "All Vehicles"
                        ckcAllAAS.Move 15, lbcSelection(6).Top - 270
                        ckcAllAAS.Visible = True
                        Select Case Index
                            Case 0  'Advertiser/Contract #
                                lbcSelection(0).Visible = True
                            Case 1  'Agency
                            Case 2  'Salesperson
                            Case 3  'vehicles
                                rbcSelCInclude(1).Value = True          'force contract/line option to line
                        End Select

                    End If

                ElseIf ilListIndex = CNT_BOB Then
                    If Index = 0 Then                       'all vehicles including pkg
                        mSellConvVVPkgPop 6, False                    'lbcselection(6), vehicles
                    Else                                    'show all vehicles excl hidden
                        mSellConvVirtVehPop 6, False
                    End If

                ElseIf ilListIndex = CNT_SPTSBYDATETIME Then     '8-24-01
                    If Index = 0 Then
                        plcSelC4.Enabled = True
                        rbcSelC4(0).Enabled = True
                        rbcSelC4(1).Enabled = True
                        rbcSelC4(0).Value = True        'default to show full price
                    Else            'no spot rates
                        plcSelC4.Enabled = False
                        rbcSelC4(0).Enabled = False
                        rbcSelC4(1).Enabled = False
                    End If
                End If
        End Select
        mSetCommands
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub

Private Sub plcSelC11_Paint()
    plcSelC11.CurrentX = 0
    plcSelC11.CurrentY = 0
    'plcSelC11.Print "Counts by"
    plcSelC11.Print smPlcSelC11P
End Sub

Private Sub plcSelC10_Paint()
    plcSelC10.CurrentX = 0
    plcSelC10.CurrentY = 0
    'plcSelC10.Print "Slsp"
    plcSelC10.Print smPlcSelC10P
End Sub

Private Sub plcSelC9_Paint()
    plcSelC9.CurrentX = 0
    plcSelC9.CurrentY = 0
    'plcSelC9.Print "Month"
    plcSelC9.Print smPlcSelC9P
End Sub

Private Sub plcSelC8_Paint()
    plcSelC8.CurrentX = 0
    plcSelC8.CurrentY = 0
    'plcSelC8.Print "For"
    plcSelC8.Print smPlcSelC8P
End Sub

Private Sub plcSelC7_Paint()
    plcSelC7.CurrentX = 0
    plcSelC7.CurrentY = 0
    'plcSelC7.Print "By"
    plcSelC7.Print smPlcSelC7P
End Sub

Private Sub plcSelC6_Paint()
    plcSelC6.CurrentX = 0
    plcSelC6.CurrentY = 0
    'plcSelC6.Print "Discrep"
    plcSelC6.Print smPlcSelC6P
End Sub

Private Sub plcSelC5_Paint()
    plcSelC5.CurrentX = 0
    plcSelC5.CurrentY = 0
    'plcSelC5.Print "Contract"
    plcSelC5.Print smPlcSelC5P
End Sub

Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    'plcSelC3.Print "Zone"
    plcSelC3.Print smPlcSelC3P
End Sub

Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    'plcSelC2.Print "Include"
    plcSelC2.Print smPlcSelC2P
End Sub

Private Sub plcSelC1_Paint()
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    'plcSelC1.Print smPlcSelC1P
    plcSelC1.Print smPlcSelC1P
End Sub

Private Sub plcSelC4_Paint()
    plcSelC4.CurrentX = 0
    plcSelC4.CurrentY = 0
    'plcSelC4.Print "Option"
    plcSelC4.Print smPlcSelC4P
End Sub

Public Sub mAskCntrFeed(ilTop As Integer)
    plcSelC12.Move 120, ilTop, 4000
    ckcSelC12(0).Move 720, 0, 1680       'local
    ckcSelC12(1).Move 2400, 0, 1440      'feed
    ckcSelC12(0).Value = vbChecked
    ckcSelC12(1).Value = vbChecked
    If tgSpf.sSystemType = "R" Then         'radio vs network/syndicator
        ckcSelC12(0).Visible = True
        ckcSelC12(0).Caption = "Contract spots"
        ckcSelC12(1).Visible = True
        ckcSelC12(1).Caption = "Feed spots"
        plcSelC12.Visible = True
        smPlcSelC12P = "Include"
        plcSelC12_Paint
    End If
End Sub

Public Sub mSetAllAASUnchecked()
    imSetAllAAS = False
    ckcAllAAS.Value = vbUnchecked 'False
    imSetAllAAS = True
End Sub

Public Sub mSetAllUnchecked()
    imSetAll = False
    ckcAll.Value = vbUnchecked  'False
    imSetAll = True
End Sub

Public Sub mSetAllVehUnchecked()
    imSetAllVeh = False
    CkcAllveh.Value = vbUnchecked 'False
    imSetAllVeh = True
End Sub

Public Sub mAskDaysOfWk()
    ckcSelC8(0).Caption = "Mo"
    ckcSelC8(1).Caption = "Tu"
    ckcSelC8(2).Caption = "We"
    ckcSelC8(3).Caption = "Th"
    ckcSelC8(4).Caption = "Fr"
    ckcSelC8(5).Caption = "Sa"
    ckcSelC8(6).Caption = "Su"
    ckcSelC8(0).Move 120, 0, 600
    ckcSelC8(1).Move 720, 0, 600
    ckcSelC8(2).Move 1320, 0, 600
    ckcSelC8(3).Move 1920, 0, 600
    ckcSelC8(4).Move 2520, 0, 600
    ckcSelC8(5).Move 3120, 0, 600
    ckcSelC8(6).Move 3720, 0, 600

    ckcSelC8(0).Value = vbChecked
    ckcSelC8(1).Value = vbChecked
    ckcSelC8(2).Value = vbChecked
    ckcSelC8(3).Value = vbChecked
    ckcSelC8(4).Value = vbChecked
    ckcSelC8(5).Value = vbChecked
    ckcSelC8(6).Value = vbChecked

    ckcSelC8(0).Visible = True
    ckcSelC8(1).Visible = True
    ckcSelC8(2).Visible = True
    ckcSelC8(3).Visible = True
    ckcSelC8(4).Visible = True
    ckcSelC8(5).Visible = True
    ckcSelC8(6).Visible = True
    plcSelC8.Visible = True
    Exit Sub
End Sub

Public Sub mSellConvRepVehPop(ilIndex As Integer)
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(RptSelCb, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvRepVehPopErr
        gCPErrorMsg ilRet, "mSellConvRepVehPop (gPopUserVehicleBox: Vehicle)", RptSelCb
        On Error GoTo 0
    End If
    Exit Sub
mSellConvRepVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'               Spot Discrepancy Summary by Month Selectivity
'
Public Sub mDiscrepSumSelectivity()
    Dim llStdLastInvDate As Long
    Dim slDate As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    
    lacSelCFrom.Move 120, 180
    lacSelCFrom.Caption = "Month"
    cbcSet1.Move 840, 120, 1320
    cbcSet1.Clear
    cbcSet1.AddItem "January"
    cbcSet1.AddItem "February"
    cbcSet1.AddItem "March"
    cbcSet1.AddItem "April"
    cbcSet1.AddItem "May"
    cbcSet1.AddItem "June"
    cbcSet1.AddItem "July"
    cbcSet1.AddItem "August"
    cbcSet1.AddItem "September"
    cbcSet1.AddItem "October"
    cbcSet1.AddItem "November"
    cbcSet1.AddItem "December"
    'obtain site, and get last month billed.  Default to next month to be billed
    gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llStdLastInvDate
    llStdLastInvDate = llStdLastInvDate + 1
    'get the month & Year of next invoice period
    slDate = gObtainEndStd(Format$(llStdLastInvDate, "m/d/yy"))
    gObtainMonthYear 0, slDate, ilMonth, ilYear
    edcSelCTo.Text = Trim$(str(ilYear))
    cbcSet1.ListIndex = ilMonth - 1
    
    lacSelCTo.Move 2520, 180, 480
    lacSelCTo.Caption = "Year"
    edcSelCTo.Move lacSelCTo.Width + lacSelCTo.Left + 140, cbcSet1.Top, 600
    edcSelCTo.MaxLength = 4
    
    lacSelCFrom1.Caption = "Contract #"
    lacSelCFrom1.Move 120, cbcSet1.Top + cbcSet1.Height + 240, 1200
    edcSelCFrom1.Move 1200, lacSelCFrom1.Top - 60, 1200
    edcSelCFrom1.MaxLength = 9
    
    cbcSet1.Visible = True
    lacSelCFrom.Visible = True
    edcSelCFrom1.Visible = True
    lacSelCFrom1.Visible = True
    edcSelCTo.Visible = True
    lacSelCTo.Visible = True
    
    pbcSelC.Visible = True
    pbcOption.Visible = True
End Sub

'       mAskContractTypesCkcSelC6 - ask contract types using control CkcSelC6 array
'       <input>  iltop = top position for first line of selection
Public Sub mAskContractTypesCkcSelC6(ilTop As Integer)
    'plcSelC6.Move 120, plcSelC7.Top + plcSelC7.Height, 4260
    plcSelC6.Move 120, ilTop, 4260, 440
    ckcSelC6(0).Move 720, -30, 1080
    ckcSelC6(0).Caption = "Standard"
    If ckcSelC6(0).Value = vbChecked Then
        ckcSelC6_click 0
    Else
        ckcSelC6(0).Value = vbChecked
    End If
    ckcSelC6(0).Visible = True
    ckcSelC6(1).Move 1800, -30, 1200
    ckcSelC6(1).Caption = "Reserved"
    If ckcSelC6(1).Value = vbChecked Then
        ckcSelC6_click 1
    Else
        ckcSelC6(1).Value = vbChecked
    End If
    ckcSelC6(1).Visible = True
    If tgUrf(0).iSlfCode > 0 Then           'its a slsp thats is asking for this report,
                                            'don't allow them to exclude reserves
        ckcSelC6(1).Enabled = False
    Else
        ckcSelC6(1).Enabled = True
    End If
    ckcSelC6(2).Move 3000, -30, 1080
    ckcSelC6(2).Caption = "Remnant"
    If ckcSelC6(2).Value = vbChecked Then
        ckcSelC6_click 2
    Else
        ckcSelC6(2).Value = vbChecked
    End If
    ckcSelC6(2).Visible = True
    ckcSelC6(3).Move 720, 195, 600
    ckcSelC6(3).Caption = "DR"
    If ckcSelC6(3).Value = vbChecked Then
        ckcSelC6_click 3
    Else
        ckcSelC6(3).Value = vbChecked   'True
    End If
    ckcSelC6(3).Visible = True
    ckcSelC6(4).Move 1260, 195, 1320
    ckcSelC6(4).Caption = "Per Inquiry"
    If ckcSelC6(4).Value = vbChecked Then
        ckcSelC6_click 4
    Else
        ckcSelC6(4).Value = vbChecked
    End If
    ckcSelC6(4).Visible = True
    
                
    ckcSelC6(5).Move 2580, 195, 720
    ckcSelC6(5).Caption = "PSA"
    ckcSelC6(5).Value = vbUnchecked 'False
    ckcSelC6(5).Visible = True
    
    ckcSelC6(6).Move 3300, 195, 900
    ckcSelC6(6).Caption = "Promo"
    ckcSelC6(6).Value = vbUnchecked 'False
    ckcSelC6(6).Visible = True
    
    plcSelC6.Visible = True
    Exit Sub
End Sub

'       '12-28-17
'       mAskContractTypesCkcSelC3 - ask contract types using control CkcSelC3 array
'       <input>  iltop = top position for first line of selectionPublic
Sub mAskContractTypesCkcSelC3(ilTop As Integer)
    plcSelC3.Move 120, ilTop, 4260, 660
    
    ckcSelC3(0).Caption = "Holds"
    ckcSelC3(0).Move 720, -30, 840
    ckcSelC3(0).Value = vbChecked   'True
    If ckcSelC3(0).Value = vbChecked Then
        ckcSelC3_click 0
    Else
        ckcSelC3(0).Value = vbChecked   'True
    End If
    ckcSelC3(0).Visible = True
    
    ckcSelC3(1).Value = vbChecked   'True
    ckcSelC3(1).Caption = "Orders"
    ckcSelC3(1).Move 1500, -30, 900
    If ckcSelC3(1).Value = vbChecked Then
        ckcSelC3_click 1
    Else
        ckcSelC3(1).Value = vbChecked   'True
    End If
    ckcSelC3(1).Visible = True
    
    ckcSelC3(2).Move 720, 210, 1080
    ckcSelC3(2).Caption = "Standard"
    If ckcSelC3(2).Value = vbChecked Then
        ckcSelC3_click 2
    Else
        ckcSelC3(2).Value = vbChecked   'True
    End If
    ckcSelC3(2).Visible = True
    
    ckcSelC3(3).Move 1800, 210, 1200
    ckcSelC3(3).Caption = "Reserved"
    If ckcSelC3(3).Value = vbChecked Then
        ckcSelC3_click 3
    Else
        ckcSelC3(3).Value = vbChecked   'True
    End If
    ckcSelC3(3).Visible = True
    
    If tgUrf(0).iSlfCode > 0 Then           'its a slsp thats is asking for this report,
                                            'don't allow them to exclude reserves
        ckcSelC3(4).Enabled = False
    Else
        ckcSelC3(4).Enabled = True
    End If
    ckcSelC3(4).Move 3000, 210, 1080
    ckcSelC3(4).Caption = "Remnant"
    If ckcSelC3(4).Value = vbChecked Then
        ckcSelC3_click 4
    Else
        ckcSelC3(4).Value = vbChecked   'True
    End If
    ckcSelC3(4).Visible = True
    
    ckcSelC3(5).Move 720, 440, 600
    ckcSelC3(5).Caption = "DR"
    If ckcSelC3(5).Value = vbChecked Then
        ckcSelC3_click 5
    Else
        ckcSelC3(5).Value = vbChecked   'True
    End If
    ckcSelC3(5).Visible = True
    
    ckcSelC3(6).Move 1260, 440, 1320
    ckcSelC3(6).Caption = "Per Inquiry"
    If ckcSelC3(6).Value = vbChecked Then
        ckcSelC3_click 6
    Else
        ckcSelC3(6).Value = vbChecked   'True
    End If
    ckcSelC3(6).Visible = True
    
    ckcSelC3(7).Move 2580, 440, 720
    ckcSelC3(7).Caption = "PSA"
    ckcSelC3(7).Value = vbUnchecked 'False
    ckcSelC3(7).Visible = True
    
    ckcSelC3(8).Move 3300, 440, 900
    ckcSelC3(8).Caption = "Promo"
    ckcSelC3(8).Value = vbUnchecked 'False
    ckcSelC3(8).Visible = True
    smPlcSelC3P = "Include"
    plcSelC3.Visible = True
    Exit Sub
End Sub
