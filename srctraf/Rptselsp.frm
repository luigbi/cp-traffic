VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelSP 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5925
   ClientLeft      =   210
   ClientTop       =   1860
   ClientWidth     =   11475
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
   ScaleHeight     =   5925
   ScaleWidth      =   11475
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   8295
      TabIndex        =   75
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   5280
      Top             =   1680
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
      Height          =   4440
      Left            =   45
      TabIndex        =   14
      Top             =   1440
      Width           =   10890
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
         Height          =   4155
         Left            =   210
         ScaleHeight     =   4155
         ScaleWidth      =   6135
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   6135
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   1320
            TabIndex        =   23
            Top             =   30
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            Text            =   "9/6/2019"
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
         Begin VB.TextBox edcIndex 
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
            Index           =   9
            Left            =   5400
            MaxLength       =   4
            TabIndex        =   68
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcIndex 
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
            Index           =   8
            Left            =   4800
            MaxLength       =   4
            TabIndex        =   67
            Text            =   "4.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcIndex 
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
            Index           =   7
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   66
            Text            =   "3.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcIndex 
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
            Index           =   6
            Left            =   3600
            MaxLength       =   4
            TabIndex        =   65
            Text            =   "2.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcIndex 
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
            Index           =   5
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   64
            Text            =   "2.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcIndex 
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
            Index           =   4
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   63
            Text            =   "1.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcIndex 
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
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   62
            Text            =   "1.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcIndex 
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
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   61
            Text            =   "1.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcIndex 
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
            Left            =   600
            MaxLength       =   4
            TabIndex        =   60
            Text            =   "1.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox lacLengths 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   0
            TabIndex        =   59
            TabStop         =   0   'False
            Text            =   "Enterthe index associated with the spot length  (.50, 1.00, 2.00) "
            Top             =   3240
            Width           =   4815
         End
         Begin VB.TextBox lacLengths 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   0
            TabIndex        =   58
            TabStop         =   0   'False
            Text            =   "Enter up to 10 Spot Lengths  "
            Top             =   2640
            Width           =   4575
         End
         Begin VB.TextBox edcIndex 
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
            Left            =   0
            MaxLength       =   4
            TabIndex        =   57
            Text            =   "1.00"
            Top             =   3480
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Index           =   9
            Left            =   5400
            MaxLength       =   3
            TabIndex        =   56
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Index           =   8
            Left            =   4800
            MaxLength       =   3
            TabIndex        =   55
            Text            =   "120"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Index           =   7
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   54
            Text            =   "90"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Index           =   6
            Left            =   3600
            MaxLength       =   3
            TabIndex        =   53
            Text            =   "60"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Index           =   5
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   52
            Text            =   "45"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Index           =   4
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   51
            Text            =   "30"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   50
            Text            =   "20"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   49
            Text            =   "15"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Left            =   600
            MaxLength       =   3
            TabIndex        =   48
            Text            =   "10"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcLen 
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
            Left            =   0
            MaxLength       =   3
            TabIndex        =   47
            Text            =   "5"
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox edcEstPct 
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
            Left            =   2700
            MaxLength       =   3
            TabIndex        =   43
            Text            =   "100"
            Top             =   1920
            Width           =   570
         End
         Begin VB.TextBox edcPctChg 
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
            TabIndex        =   41
            Text            =   "100"
            Top             =   1590
            Width           =   570
         End
         Begin VB.PictureBox plcValRate 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1320
            Width           =   4380
            Begin VB.OptionButton rbcValRate 
               Caption         =   "Avg 30' Rate"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   2760
               TabIndex        =   39
               Top             =   0
               Width           =   1485
            End
            Begin VB.OptionButton rbcValRate 
               Caption         =   "Rate Card"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   1560
               TabIndex        =   38
               Top             =   0
               Value           =   -1  'True
               Width           =   1215
            End
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
            Left            =   4200
            TabIndex        =   46
            Top             =   2280
            Width           =   1500
         End
         Begin VB.TextBox edcSet2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   3000
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "Minor Set #"
            Top             =   2310
            Width           =   1215
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
            Left            =   1200
            TabIndex        =   45
            Top             =   2280
            Width           =   1500
         End
         Begin VB.TextBox edcSet1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   0
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "Major Set #"
            Top             =   2310
            Width           =   1215
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4140
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1020
            Width           =   4140
            Begin VB.OptionButton rbcSelC2 
               Caption         =   "Both"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2820
               TabIndex        =   34
               Top             =   15
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton rbcSelC2 
               Caption         =   "Summary"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1545
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   0
               Width           =   1110
            End
            Begin VB.OptionButton rbcSelC2 
               Caption         =   "Detail"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   0
               Width           =   795
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
            Left            =   1995
            MaxLength       =   1
            TabIndex        =   27
            Top             =   360
            Width           =   345
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
            Left            =   720
            MaxLength       =   4
            TabIndex        =   25
            Top             =   360
            Width           =   570
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   3780
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   750
            Visible         =   0   'False
            Width           =   3780
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Corporate"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   600
               TabIndex        =   20
               Top             =   0
               Width           =   1155
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1920
               TabIndex        =   21
               Top             =   0
               Width           =   1365
            End
         End
         Begin VB.Label lacPctSellout 
            Caption         =   "Est. % Sellout of Unsold Avails"
            Height          =   255
            Left            =   0
            TabIndex        =   42
            Top             =   1950
            Width           =   2655
         End
         Begin VB.Label lacPctChg 
            Caption         =   "+/- (ie -5 or 5) % change of Unsold Spot Prices"
            Height          =   255
            Left            =   0
            TabIndex        =   40
            Top             =   1620
            Width           =   3255
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   22
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lacSelCTo1 
            Appearance      =   0  'Flat
            Caption         =   "Qtr"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1485
            TabIndex        =   19
            Top             =   405
            Width           =   285
         End
         Begin VB.Label lacSelCTo 
            Appearance      =   0  'Flat
            Caption         =   "Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   18
            Top             =   405
            Width           =   660
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
         Height          =   3540
         Left            =   6120
         ScaleHeight     =   3540
         ScaleWidth      =   4530
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   4530
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   3
            ItemData        =   "Rptselsp.frx":0000
            Left            =   2400
            List            =   "Rptselsp.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   71
            Top             =   300
            Width           =   1995
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1500
            Index           =   2
            ItemData        =   "Rptselsp.frx":0004
            Left            =   240
            List            =   "Rptselsp.frx":0006
            MultiSelect     =   2  'Extended
            TabIndex        =   70
            Top             =   300
            Width           =   1995
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   1
            ItemData        =   "Rptselsp.frx":0008
            Left            =   2400
            List            =   "Rptselsp.frx":000A
            TabIndex        =   73
            Top             =   2160
            Width           =   1995
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   0
            ItemData        =   "Rptselsp.frx":000C
            Left            =   240
            List            =   "Rptselsp.frx":000E
            TabIndex        =   72
            Top             =   2160
            Width           =   1995
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   255
            TabIndex        =   69
            Top             =   0
            Width           =   1425
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Avail Names for Sports"
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
            Index           =   2
            Left            =   2400
            TabIndex        =   44
            Top             =   0
            Width           =   2085
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Rate Card"
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
            Index           =   1
            Left            =   2400
            TabIndex        =   24
            Top             =   1920
            Width           =   1365
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Budget Names"
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
            TabIndex        =   26
            Top             =   1920
            Width           =   1365
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   8280
      TabIndex        =   76
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   7920
      TabIndex        =   74
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
         Width           =   1095
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   10680
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RptSelSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselsp.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelSP.Frm   Sales vs Plan
'
' Release: 1.0
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
        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    imAllClicked = False
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
    If rbcSelC2(2).Value Then           'both
        ilStartJobNo = 1
        ilNoJobs = 2
    ElseIf rbcSelC2(0).Value Then       'detail only
        ilStartJobNo = 1
        ilNoJobs = 1
    Else
        ilStartJobNo = 2                'summary
        ilNoJobs = 2
    End If
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If (ilJobs = 1) Then
            'ilRet = gOpenPrtJob("SlsvsPln.rpt")        'replaced by slvspldt (new detail version)
             ilRet = gOpenPrtJob("SlvsPlDt.rpt")
        Else
            ilRet = gOpenPrtJob("SlvsPlsm.rpt")
        End If
        If Not ilRet Then
            'If Not gOpenPrtJob("SlsvsPln.Rpt") Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenSP(imGenShiftKey, smLogUserCode)
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

        If (igJobRptNo = 2 And rbcSelC2(1).Value) Or (igJobRptNo = 1) Then  'summary only or detail (with/without detail) requested
            Screen.MousePointer = vbHourglass
            gCRSalesPlan
            Screen.MousePointer = vbDefault
        End If

        'Screen.MousePointer = vbHourGlass
        'gCRSalesPlan
        'Screen.MousePointer = vbDefault
        If rbcOutput(0).Value Then
            DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
            igDestination = 0
            Report.Show vbModal
        ElseIf rbcOutput(1).Value Then
            ilCopies = Val(edcCopies.Text)
            ilRet = gOutputToPrinter(ilCopies)
        Else
            slFileName = edcFileName.Text
           ' ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-10
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

Private Sub edcEstPct_GotFocus()
    gCtrlGotFocus edcEstPct
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

Private Sub edcIndex_GotFocus(Index As Integer)
    gCtrlGotFocus edcIndex(Index)
End Sub

Private Sub edcLen_GotFocus(Index As Integer)
    gCtrlGotFocus edcLen(Index)
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcPctChg_GotFocus()
    gCtrlGotFocus edcPctChg
End Sub
'
'Private Sub edcSelCFrom_Change()
'    mSetCommands
'End Sub
'Private Sub edcSelCFrom_GotFocus()
'    gCtrlGotFocus edcSelCFrom
'End Sub

Private Sub edcSelCTo_Change()
Dim ilLen As Integer
Dim slDate As String
Dim llDate As Long
Dim ilRet As Integer
    ilLen = Len(edcSelCTo)
    If ilLen = 4 Then
        'if using Corporate calendar, need to show all rate cards applicable, and they must pick 2 years
            If rbcSelCSelect(0).Value Then          'corp, need to adjust the date for the rate cards to bring in
            'if asking for corp year 1997, need to bring in 1996 and 1997 since rates cards are input as jan-dec
            ilLen = Val(edcSelCTo) - 1
            ilRet = gGetCorpCalIndex(ilLen)
            If ilRet < 1 Then
                slDate = "1/15/" & Trim$(edcSelCTo)           'retrieve jan thru dec year
            Else                                                'no error
                slDate = Trim$(str$(tgMCof(ilRet).iStartMnthNo)) & "/15/" & Trim$(str$(Val((edcSelCTo) - 1)))      'retrieve corp year
            End If
            llDate = gDateValue(slDate)
            'populate Rate Cards and bring in Rcf, Rif, and Rdf
            ilRet = gPopRateCardBox(RptSelSP, llDate, RptSelSP!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
            lbcSelection(1).Visible = True
        Else
            slDate = "1/15/" & Trim$(edcSelCTo)           'retrieve jan thru dec year
            slDate = gObtainStartStd(slDate)
            llDate = gDateValue(slDate)

            'populate Rate Cards and bring in Rcf, Rif, and Rdf
            ilRet = gPopRateCardBox(RptSelSP, llDate, RptSelSP!lbcSelection(1), tgRateCardCode(), smRateCardTag, -1)
            lbcSelection(1).Visible = True
        End If
    End If
    mSetCommands
End Sub
Private Sub edcSelCTo_GotFocus()
    gCtrlGotFocus edcSelCTo
End Sub
Private Sub edcSelCTo1_Change()
    mSetCommands
End Sub
Private Sub edcSelCTo1_GotFocus()
    gCtrlGotFocus edcSelCTo1
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
    RptSelSP.Refresh
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
    'RptSelSP.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgRptSelBudgetCodeSP
    Erase tgClfSP
    Erase tgCffSP
    Erase imCodes
    PECloseEngine
    
    Set RptSelSP = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Not imAllClicked Then
        If Index = 2 Then
            imSetAll = False
            ckcAll.Value = vbUnchecked  'False
            imSetAll = True
        End If
    End If
    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub
'**********************************************************************************
'
'                   mAskCorpOrStd - Ask on Report input screen:
'                           Month  o Corp   o Std
'                   Set the proper properties to this control
'                   Created: DH 9/9/96
'*********************************************************************************
'
Private Sub mAskCorpOrStd()
    Dim ilRet As Integer

    If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
        rbcSelCSelect(0).Enabled = False
        rbcSelCSelect(1).Value = True
       ' ilRet = gObtainCorpCal()            'Retain corporate calendar in memeory if using
    Else
        rbcSelCSelect(0).Value = True
        ilRet = gObtainCorpCal()            'Retain corporate calendar in memeory if using
    End If
    plcSelC1.Visible = True
End Sub
'
'
'                   mAskEffDate - Ask Effective Date, Start Year
'                                 and Quarter
'
'                   6/7/97
'
'
Private Sub mAskEffDate()
'    lacSelCFrom.Left = 120
'    edcSelCFrom.Move 1350, edcSelCFrom.Top, 945
'    edcSelCFrom.MaxLength = 10  '8   5/28/99 allow 10 char input date mm/dd/yyyy
'    lacSelCFrom.Caption = "Effective Date"
'    lacSelCFrom.Top = 75
'    lacSelCFrom.Visible = True
'    edcSelCFrom.Visible = True
'    lacSelCTo.Caption = "Year"
'    lacSelCTo.Visible = True
'    lacSelCTo.Left = 120
'    lacSelCTo.Top = edcSelCFrom.Top + edcSelCFrom.Height + 75
'    lacSelCTo1.Left = 1580
'    lacSelCTo1.Caption = "Quarter"
'    lacSelCTo1.Width = 810
'    lacSelCTo1.Top = edcSelCFrom.Top + edcSelCFrom.Height + 75
'    lacSelCTo1.Visible = True
'    edcSelCTo.Move 600, edcSelCFrom.Top + edcSelCFrom.Height + 30, 600
'    edcSelCTo1.Move 2340, edcSelCFrom.Top + edcSelCFrom.Height + 30, 300
'    edcSelCTo.MaxLength = 4
'    edcSelCTo1.MaxLength = 1
'    edcSelCTo.Visible = True
'    edcSelCTo1.Visible = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBudgetPop                      *
'*             Created:              By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Budget name  list     *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mBudgetPop()
'
'   mBudgetPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gPopVehBudgetBox(RptSelSP, 2, 0, 1, lbcSelection(0), tgRptSelBudgetCodeSP(), sgRptSelBudgetCodeTagSP)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mBudgetPopErr
        gCPErrorMsg ilRet, "mBudgetPopErr (gPopVehBudgetBox)", RptSelSP
        On Error GoTo 0
    End If
    Exit Sub
mBudgetPopErr:
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

    RptSelSP.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
'    lacSelCFrom.Move 120, 75
'    lacSelCTo.Move 120, 390
'    lacSelCTo1.Move 2400, 390
'    edcSelCFrom.Move 1500, 30
'    edcSelCTo.Move 1500, 345
'    edcSelCTo1.Move 2715, 345
'    plcSelC1.Move 120, 675
'    pbcSelC.Move 90, 255, 4515, 3360

    gCenterStdAlone RptSelSP
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
    'ReDim tgMMnf(1 To 1) As MNF
    ReDim tgMMnf(0 To 0) As MNF
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

    'pbcSelC.Visible = False
    'lbcRptType.Clear
    'lbcRptType.AddItem smSelectedRptName

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    'mSellConvVirtVehPop 2, False
    mSellConvVVActDormPop 2, False
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    lacSelCFrom.Visible = True
    lacSelCTo.Visible = True
'    edcSelCFrom.Visible = True
    edcSelCTo.Visible = True
    ckcAll.Visible = True
    mAskEffDate             'ask effective date, year & qtr
    mAskCorpOrStd
 
    'Detail, summary or both
'    rbcSelC2(0).Left = 600
'    plcSelC2.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height
'    rbcSelC2(1).Left = rbcSelC2(0).Left + rbcSelC2(0).Width
'    rbcSelC2(2).Left = rbcSelC2(1).Left + rbcSelC2(1).Width
    sgMnfVehGrpTag = ""
    gPopVehicleGroups RptSelSP!cbcSet1, tgVehicleSets1(), False
    gPopVehicleGroups RptSelSP!cbcSet2, tgVehicleSets2(), True
    cbcSet1.ListIndex = 0
    cbcSet2.ListIndex = 0
'    edcSet1.Move 120, plcSelC2.Top + plcSelC2.Height + 90
'    cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 30
'    edcSet2.Move 120, cbcSet1.Top + cbcSet1.Height + 90
'    cbcSet2.Move cbcSet1.Left, edcSet2.Top - 30


    ' ****** Temporarily patch of the ability to use sets 1 and 2 for sorting
    ' ******
    '
    'plcSelC2.Visible = False
    'cbcSet1.Visible = False
    'cbcSet2.Visible = False
    'edcSet1.Visible = False
    'edcSet2.Visible = False
    'rbcSelC2(0).Value = True
    mBudgetPop              'lbcSelection(0), one budget only
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    
    mPopAvailNames
    
'   ckcAll.Move 15, 15
'    lbcSelection(0).Visible = True                  'show budget name list box (base budget)
'    lbcSelection(1).Visible = True                 'split budgets
'    lbcSelection(2).Visible = True                  'vehicle selection
'    laclbcName(0).Visible = True
'    laclbcName(0).Caption = "Budget Names"
'    laclbcName(1).Visible = True
'    laclbcName(1).Caption = "Rate Cards"
'    lbcSelection(2).Move 15, ckcAll.Height + 30, 4380, 1500
'    lbcSelection(0).Move lbcSelection(2).Left, lbcSelection(2).Top + lbcSelection(2).Height + 300, lbcSelection(2).Width / 2, lbcSelection(2).Height - 15
'    lbcSelection(1).Move lbcSelection(2).Left + lbcSelection(2).Width / 2 + 60, lbcSelection(2).Top + lbcSelection(2).Height + 300, lbcSelection(2).Width / 2, lbcSelection(2).Height - 15
'    laclbcName(0).Move lbcSelection(2).Left, lbcSelection(0).Top - laclbcName(0).Height - 30, 2205
'    laclbcName(1).Move lbcSelection(2).Left + lbcSelection(2).Width / 2 + 60, lbcSelection(1).Top - laclbcName(0).Height - 30, 2205
    pbcSelC.Visible = True
    pbcOption.Visible = True
    If tgSpf.sRUseCorpCal = "Y" Then       'if Using corp cal,retain in memory
        ilRet = gObtainCorpCal()
    End If

    'If lbcRptType.ListCount > 0 Then
    '    gFindMatch smSelectedRptName, 0, lbcRptType
    '    If gLastFound(lbcRptType) < 0 Then
    '        MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
    '        imTerminate = True
    '        Exit Sub
    '    End If
     '   lbcRptType.ListIndex = gLastFound(lbcRptType)
    'End If
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
    'gInitStdAlone RptSelSP, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Sales vs Plan"
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
'*      Procedure Name:mSellConvVVActDormPop             *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSellConvVVActDormPop(ilIndex As Integer, ilUselbcVehicle As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        ilRet = gPopUserVehicleBox(RptSelSP, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelSP, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)   'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVVActDormPop (gPopUserVehicleBox: Vehicle)", RptSelSP
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
    If (CSI_CalFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
'    If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
        ilEnable = True
        If Not ckcAll.Value = vbChecked Then
            ilEnable = False
            For ilLoop = 0 To lbcSelection(2).ListCount - 1 Step 1      'vehicle entry must be selected
                If lbcSelection(2).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        End If
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
    Unload RptSelSP
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcValRate_Paint()
    plcValRate.CurrentX = 0
    plcValRate.CurrentY = 0
    plcValRate.Print "Inv. Valuation use"
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
Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print "Month"
End Sub
Private Sub plcSelC1_Paint()
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    plcSelC1.Print "Select"
End Sub
'
'           mPopAvailNames - populate list box of all avail names and
'           default the selection based on anf property (anfRptDefault)
Public Sub mPopAvailNames()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilLoopOnListBox As Integer
    Dim slName As String
    Dim slCode As String
    Dim slNameCode As String
    Dim ilAnfCode As Integer

        ilRet = gAvailsPop(RptSelSP, lbcSelection(3), tgNamedAvail())       'show the named avails for selectivity
        'default the selection on if user named avails flag is set
        For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1
            If tgAvailAnf(ilLoop).sRptDefault = "Y" Then                'set as selected
                For ilLoopOnListBox = 0 To lbcSelection(3).ListCount
                    slNameCode = tgNamedAvail(ilLoopOnListBox).sKey
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slName, 3, "|", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilAnfCode = Val(slCode)
                    If ilAnfCode = tgAvailAnf(ilLoop).iCode Then        'the named avail is set to show on the Combo aVails, default selected
                        lbcSelection(3).Selected(ilLoopOnListBox) = True
                        Exit For
                    End If
                Next ilLoopOnListBox
            End If
        Next ilLoop

End Sub
