VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSel30 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CPP CPM 30"" Unit Report Selection"
   ClientHeight    =   6405
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
   ScaleHeight     =   6405
   ScaleWidth      =   11475
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   8295
      TabIndex        =   100
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   5760
      Top             =   2520
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
         TabIndex        =   20
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
         TabIndex        =   12
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
         TabIndex        =   16
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
         TabIndex        =   14
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Vehicle Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4920
      Left            =   45
      TabIndex        =   23
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
         Height          =   4635
         Left            =   90
         ScaleHeight     =   4635
         ScaleWidth      =   6225
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   6225
         Begin VB.PictureBox plcGrossNet 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3000
            ScaleHeight     =   240
            ScaleWidth      =   1980
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   3975
            Width           =   1980
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   113
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1320
               TabIndex        =   112
               Top             =   0
               Width           =   645
            End
         End
         Begin VB.CheckBox ckcDetail 
            Caption         =   "Include Line Detail by Daypart"
            Height          =   375
            Left            =   0
            TabIndex        =   110
            Top             =   3900
            Width           =   3015
         End
         Begin VB.PictureBox plcDP 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   3300
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   4200
            Visible         =   0   'False
            Width           =   3300
            Begin VB.OptionButton rbcDP 
               Caption         =   "Std DP"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   2520
               TabIndex        =   109
               Top             =   0
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.OptionButton rbcDP 
               Caption         =   "Ovverride DP"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   600
               TabIndex        =   108
               Top             =   0
               Visible         =   0   'False
               Width           =   1935
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
            Left            =   5040
            TabIndex        =   106
            Top             =   3960
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.PictureBox plcCTypes 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   5535
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   1260
            Width           =   5535
            Begin VB.CheckBox ckcCType 
               Caption         =   "Non-Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   11
               Left            =   4080
               TabIndex        =   47
               Top             =   480
               Value           =   1  'Checked
               Width           =   1380
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Holds"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   36
               Top             =   -30
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Orders"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1920
               TabIndex        =   37
               Top             =   -30
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3120
               TabIndex        =   46
               Top             =   480
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   3000
               TabIndex        =   38
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1200
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Reserved"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   4320
               TabIndex        =   39
               Top             =   -30
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Remnant"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   960
               TabIndex        =   40
               Top             =   240
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "DR"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2160
               TabIndex        =   41
               Top             =   240
               Value           =   1  'Checked
               Width           =   735
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Per Inquiry"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2880
               TabIndex        =   42
               Top             =   240
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "PSA"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   4320
               TabIndex        =   43
               Top             =   240
               Width           =   825
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Promo"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   960
               TabIndex        =   44
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox ckcCType 
               Caption         =   "Trades"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   2040
               TabIndex        =   45
               Top             =   480
               Value           =   1  'Checked
               Width           =   1020
            End
         End
         Begin VB.PictureBox plcSpotTypes 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   120
            ScaleHeight     =   555
            ScaleWidth      =   5820
            TabIndex        =   95
            Top             =   2040
            Width           =   5820
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "Spinoff"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   3720
               TabIndex        =   55
               Top             =   240
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "unuse"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   11
               Left            =   3240
               TabIndex        =   59
               Top             =   480
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1320
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "unuse"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   2040
               TabIndex        =   58
               Top             =   480
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "unuse"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   960
               TabIndex        =   57
               Top             =   480
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "Unuse"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   4800
               TabIndex        =   56
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "Charge"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   48
               Top             =   0
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2160
               TabIndex        =   49
               Top             =   0
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "ADU"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3120
               TabIndex        =   50
               Top             =   0
               Value           =   1  'Checked
               Width           =   840
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "Bonus"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   4080
               TabIndex        =   51
               Top             =   0
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "N/C"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   960
               TabIndex        =   52
               Top             =   240
               Value           =   1  'Checked
               Width           =   780
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "MG"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   1800
               TabIndex        =   53
               Top             =   240
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.CheckBox ckcSpotType 
               Caption         =   "Recap"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2640
               TabIndex        =   54
               Top             =   240
               Value           =   1  'Checked
               Width           =   1035
            End
         End
         Begin VB.PictureBox plcBook 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4620
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   990
            Width           =   4620
            Begin VB.OptionButton rbcBook 
               Caption         =   "Closest to Airing"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   2760
               TabIndex        =   33
               Top             =   0
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.OptionButton rbcBook 
               Caption         =   "Schd Line"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   480
               TabIndex        =   31
               Top             =   0
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton rbcBook 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1680
               TabIndex        =   32
               Top             =   0
               Width           =   1005
            End
         End
         Begin VB.TextBox edcYear 
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
            Left            =   600
            MaxLength       =   4
            TabIndex        =   9
            Top             =   60
            Width           =   570
         End
         Begin VB.PictureBox plcMonthType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2400
            ScaleHeight     =   240
            ScaleWidth      =   3900
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   4320
            Visible         =   0   'False
            Width           =   3900
            Begin VB.OptionButton rbcMonthType 
               Caption         =   "Corporate"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1320
               TabIndex        =   18
               Top             =   0
               Width           =   1275
            End
            Begin VB.OptionButton rbcMonthType 
               Caption         =   "Calendar"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   17
               Top             =   0
               Width           =   1035
            End
            Begin VB.OptionButton rbcMonthType 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   2640
               TabIndex        =   19
               Top             =   0
               Value           =   -1  'True
               Width           =   1365
            End
         End
         Begin VB.TextBox edcStartMonth 
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
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   11
            Top             =   60
            Width           =   330
         End
         Begin VB.TextBox edcNoMonths 
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
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   13
            Top             =   60
            Width           =   345
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   2340
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   420
            Width           =   2340
            Begin VB.OptionButton rbcByCPPCPM 
               Caption         =   "CPP"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   480
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   795
            End
            Begin VB.OptionButton rbcByCPPCPM 
               Caption         =   "CPM"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1320
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   0
               Width           =   990
            End
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
            Left            =   4200
            TabIndex        =   30
            Top             =   660
            Width           =   1500
         End
         Begin VB.PictureBox plcDemo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2640
            ScaleHeight     =   240
            ScaleWidth      =   3420
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   420
            Width           =   3420
            Begin VB.OptionButton rbcDemo 
               Caption         =   "Primary"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   24
               Top             =   0
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton rbcDemo 
               Caption         =   "Select"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1920
               TabIndex        =   25
               Top             =   0
               Width           =   1125
            End
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
            Left            =   4680
            MaxLength       =   9
            TabIndex        =   15
            Top             =   60
            Width           =   1170
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
            TabIndex        =   60
            Text            =   "5"
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
            TabIndex        =   61
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
            Index           =   2
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   62
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
            Index           =   3
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   63
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
            Index           =   4
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   64
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
            Index           =   5
            Left            =   3000
            MaxLength       =   3
            TabIndex        =   65
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
            Index           =   6
            Left            =   3600
            MaxLength       =   3
            TabIndex        =   66
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
            Index           =   7
            Left            =   4200
            MaxLength       =   3
            TabIndex        =   67
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
            Index           =   8
            Left            =   4800
            MaxLength       =   3
            TabIndex        =   68
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
            Index           =   9
            Left            =   5400
            MaxLength       =   3
            TabIndex        =   69
            Top             =   2880
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
            Index           =   0
            Left            =   0
            MaxLength       =   4
            TabIndex        =   70
            Text            =   "1.0"
            Top             =   3540
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
            Index           =   0
            Left            =   0
            TabIndex        =   85
            TabStop         =   0   'False
            Text            =   "Enter up to 10 Spot Lengths  "
            Top             =   2640
            Width           =   4575
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
            TabIndex        =   84
            TabStop         =   0   'False
            Text            =   "Enterthe index associated with the spot length  (.50, 1.00, 2.00) "
            Top             =   3300
            Width           =   4815
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
            TabIndex        =   71
            Text            =   "1.0"
            Top             =   3540
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
            TabIndex        =   72
            Text            =   "1.0"
            Top             =   3540
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
            TabIndex        =   73
            Text            =   "1.0"
            Top             =   3540
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
            TabIndex        =   74
            Text            =   "1.0"
            Top             =   3540
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
            TabIndex        =   75
            Text            =   "2.0"
            Top             =   3540
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
            TabIndex        =   76
            Text            =   "2.0"
            Top             =   3540
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
            TabIndex        =   77
            Text            =   "3.0"
            Top             =   3540
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
            TabIndex        =   78
            Text            =   "4.0"
            Top             =   3540
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
            Index           =   9
            Left            =   5400
            MaxLength       =   4
            TabIndex        =   79
            Top             =   3540
            Width           =   465
         End
         Begin VB.PictureBox plcSortBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   3180
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   690
            Width           =   3180
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Vehicle, DP"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1800
               TabIndex        =   28
               Top             =   0
               Value           =   -1  'True
               Width           =   1365
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Advertiser"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   480
               TabIndex        =   27
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.Label lacContract 
            Caption         =   "Cnt#"
            Height          =   255
            Left            =   4200
            TabIndex        =   97
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lacYear 
            Appearance      =   0  'Flat
            Caption         =   "Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   93
            Top             =   120
            Width           =   420
         End
         Begin VB.Label lacNoMonths 
            Appearance      =   0  'Flat
            Caption         =   "# Mos."
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2880
            TabIndex        =   91
            Top             =   120
            Width           =   645
         End
         Begin VB.Label lacStartMonth 
            Caption         =   "Start Month"
            Height          =   210
            Left            =   1320
            TabIndex        =   90
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label lacGroup 
            Caption         =   "Group"
            Height          =   210
            Left            =   3480
            TabIndex        =   89
            Top             =   690
            Width           =   735
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
         Height          =   4620
         Left            =   6120
         ScaleHeight     =   4620
         ScaleWidth      =   4650
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   4650
         Begin VB.CheckBox ckcAllGroupItems 
            Caption         =   "All Group Items"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2400
            TabIndex        =   104
            Top             =   2340
            Width           =   1905
         End
         Begin VB.CheckBox ckcAllAdvt 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   120
            Width           =   1905
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1920
            Index           =   3
            ItemData        =   "RptSelCP30Unit.frx":0000
            Left            =   2400
            List            =   "RptSelCP30Unit.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   102
            Top             =   2640
            Width           =   2220
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1920
            Index           =   2
            ItemData        =   "RptSelCP30Unit.frx":0004
            Left            =   240
            List            =   "RptSelCP30Unit.frx":0006
            MultiSelect     =   2  'Extended
            TabIndex        =   82
            Top             =   2640
            Width           =   4395
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1710
            Index           =   1
            ItemData        =   "RptSelCP30Unit.frx":0008
            Left            =   2400
            List            =   "RptSelCP30Unit.frx":000A
            TabIndex        =   98
            Top             =   480
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1710
            Index           =   0
            ItemData        =   "RptSelCP30Unit.frx":000C
            Left            =   240
            List            =   "RptSelCP30Unit.frx":000E
            MultiSelect     =   2  'Extended
            TabIndex        =   92
            Top             =   480
            Width           =   1995
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   2340
            Width           =   1425
         End
         Begin VB.Label lacDemo 
            Caption         =   "Demo"
            Height          =   255
            Left            =   2400
            TabIndex        =   105
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   8280
      TabIndex        =   101
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   7920
      TabIndex        =   99
      Top             =   120
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
Attribute VB_Name = "RptSel30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptSel30.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSel30.Frm   CPP/CPM by 30"Unit
'
' Release: 6.0
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
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection(2))
Dim imSetAllAdvt As Integer 'True=Set list box; False= don't change list box
Dim imAllClickedAdvt As Integer  'True=All box clicked (don't call ckcAll within lbcSelection(0))
Dim imSetAllGroupItems As Integer 'True=Set list box; False= don't change list box
Dim imAllClickedGroupItems As Integer  'True=All box clicked (don't call ckcAll within lbcSelection(3))
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

Private Sub cbcSet1_Click()
Dim ilLoop As Integer
Dim ilSetIndex As Integer
Dim ilRet As Integer

    ilLoop = cbcSet1.ListIndex
    ilSetIndex = gFindVehGroupInx(ilLoop, tgVehicleSets1())
    If ilSetIndex > 0 Then
        smVehGp5CodeTag = ""
        ilRet = gPopMnfPlusFieldsBox(RptSel30, lbcSelection(3), tgSOCode(), smVehGp5CodeTag, "H" & Trim$(str$(ilSetIndex)))
        lbcSelection(3).Visible = True
        ckcAllGroupItems.Visible = True
        lbcSelection(2).Width = 1995
    Else
        lbcSelection(3).Visible = False
        ckcAllGroupItems.Visible = False
        lbcSelection(2).Width = 4395
        ckcAllGroupItems.Value = vbUnchecked
    End If
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
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    imAllClicked = False
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllAdvt_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllAdvt.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllAdvt Then
        imAllClickedAdvt = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    imAllClickedAdvt = False
    mSetCommands
End Sub

Private Sub ckcAllGroupItems_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllGroupItems.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllGroupItems Then
        imAllClickedGroupItems = True
        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(3).hWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    imAllClickedGroupItems = False
    mSetCommands
End Sub

Private Sub ckcCType_Click(Index As Integer)
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
        If Not gGenReport30() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If

        ilRet = gCmcGen30(imGenShiftKey, smLogUserCode)
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
        If rbcSortBy(0).Value = True Then           'by Advt
            gCrCP30UnitAdv
        Else
            gCrCP30UnitVehicle
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

Private Sub edcContract_GotFocus()
    gCtrlGotFocus edcContract
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

Private Sub edcIndex_GotFocus(Index As Integer)
    gCtrlGotFocus edcIndex(Index)
End Sub

Private Sub edcLen_Change(Index As Integer)
    mSetCommands
End Sub
Private Sub edcLen_GotFocus(Index As Integer)
    gCtrlGotFocus edcLen(Index)
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcNoMonths_Change()
    mSetCommands
End Sub

Private Sub edcNoMonths_GotFocus()
    gCtrlGotFocus edcNoMonths
End Sub

Private Sub edcStartMonth_Change()
    mSetCommands
End Sub

Private Sub edcStartMonth_GotFocus()
    gCtrlGotFocus edcStartMonth
End Sub

Private Sub edcYear_Change()
    mSetCommands
End Sub
Private Sub edcYear_GotFocus()
    gCtrlGotFocus edcYear
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
    RptSel30.Refresh
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
    'RptSel30.Show
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
    
    Set RptSel30 = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Not imAllClicked Then
        If Index = 2 Then           'vehicles
            imSetAll = False
            ckcAll.Value = vbUnchecked  'False
            imSetAll = True
        ElseIf Index = 0 Then          'advt
            imSetAllAdvt = False
            ckcAllAdvt.Value = vbUnchecked  'False
            imSetAllAdvt = True
        ElseIf Index = 3 Then           'group items
            imSetAllGroupItems = False
            ckcAllGroupItems.Value = vbUnchecked  'False
            imSetAllGroupItems = True
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

'   Currently all month types except standard are disabled
'*********************************************************************************
'
Private Sub mAskCorpOrStd()
    Dim ilRet As Integer

'    If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
'        rbcMonthType(1).Enabled = False
'        rbcMonthType(2).Value = True
'     Else
'        rbcMonthType(1).Value = True
'        ilRet = gObtainCorpCal()            'Retain corporate calendar in memeory if using
'    End If
    
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

    RptSel30.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imAllClickedGroupItems = False
    imSetAllGroupItems = True
    imAllClickedAdvt = False
    imSetAllAdvt = True
'    lacSelCFrom.Move 120, 75
'    lacSelCTo.Move 120, 390
'    lacSelCTo1.Move 2400, 390
'    edcSelCFrom.Move 1500, 30
'    edcSelCTo.Move 1500, 345
'    edcSelCTo1.Move 2715, 345
'    plcSelC1.Move 120, 675
'    pbcSelC.Move 90, 255, 4515, 3360
    lbcSelection(2).Width = 4395            'default vehicles to width of list box area, unless group items required
    lbcSelection(0).Width = 4395            'default advt to width of list box area, unless demo required

    gCenterStdAlone RptSel30
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
    Dim ilShow As Integer
    Dim ilSort As Integer
    Dim ilVefCode As Integer
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
    ilRet = gPopAdvtBox(RptSel30, lbcSelection(0), tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptSel30
        On Error GoTo 0
    End If
    ilRet = gPopMnfPlusFieldsBox(RptSel30, lbcSelection(1), tgRptSelDemoCodeCP(), sgRptSelDemoCodeTagCP, "D")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopErr
        gCPErrorMsg ilRet, "RptSel30: mInitReport (gPopMnfPlusFieldsBox)", RptSel30
        On Error GoTo 0
    End If
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    ckcAll.Visible = True
    mAskEffDate             'ask effective date, year & qtr
    mAskCorpOrStd
 
    'Detail, summary or both
'    rbcSelC2(0).Left = 600
'    plcSelC2.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height
'    rbcSelC2(1).Left = rbcSelC2(0).Left + rbcSelC2(0).Width
'    rbcSelC2(2).Left = rbcSelC2(1).Left + rbcSelC2(1).Width
    sgMnfVehGrpTag = ""
    gPopVehicleGroups RptSel30!cbcSet1, tgVehicleSets1(), True
    cbcSet1.ListIndex = 0
    
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
    
    'build all the books to search but dont show them, for internal use only
    ilSort = 0      '0=sort by book name, or 1= date then book name
    ilShow = 1      '0=Show book name only, 1=show book name & date
    ilVefCode = 0
    ilRet = gPopBookNameBox(RptSelAD, 0, 0, ilVefCode, ilSort, ilShow, cbcBook, tgBookNameCode(), sgBookNameCodeTag)

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
    Exit Sub
mPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    'gInitStdAlone RptSel30, slStr, ilTestSystem
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
    
    ilRet = gPopUserVehicleBox(RptSel30, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)   'lbcCSVNameCode)
 
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVVActDormPop (gPopUserVehicleBox: Vehicle)", RptSel30
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
    If (edcYear.Text <> "") And (edcStartMonth.Text <> "") And (edcNoMonths.Text <> "") Then
        ilEnable = False
        If lbcSelection(0).SelCount > 0 And lbcSelection(2).SelCount > 0 Then        'at leat 1 advt & 1 vehicle must be selected
            ilEnable = True
            If RptSel30!rbcDemo(1).Value = True Then            'by user demo, see if one has been selected
                If lbcSelection(1).SelCount <= 0 Then           'none selected
                    ilEnable = False
                Else                                            'determine if a vehicle group has been selected
                    If cbcSet1.ListIndex > 0 Then
                        If lbcSelection(3).SelCount <= 0 Then
                            ilEnable = False
                        End If
                    End If
                End If
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
    Unload RptSel30
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcBook_Paint()
    plcBook.CurrentX = 0
    plcBook.CurrentY = 0
    plcBook.Print "Book"
End Sub

Private Sub plcCTypes_Paint()
    plcCTypes.CurrentX = 0
    plcCTypes.CurrentY = 0
    plcCTypes.Print "Include"
End Sub

Private Sub plcDemo_Paint()
    plcDemo.CurrentX = 0
    plcDemo.CurrentY = 0
    plcDemo.Print "Demo"
End Sub

Private Sub plcDP_Paint()
    plcDP.CurrentX = 0
    plcDP.CurrentY = 0
    plcDP.Print "Show by"
End Sub

Private Sub plcGrossNet_Paint()
    plcGrossNet.CurrentX = 0
    plcGrossNet.CurrentY = 0
    plcGrossNet.Print "By"
End Sub

Private Sub plcMonthType_Paint()
    plcMonthType.CurrentX = 0
    plcMonthType.CurrentY = 0
    plcMonthType.Print "Use Month"
End Sub

Private Sub plcSortBy_Paint()
    plcSortBy.CurrentX = 0
    plcSortBy.CurrentY = 0
    plcSortBy.Print "Sort"
End Sub

Private Sub plcSpotTypes_Paint()
    plcSpotTypes.CurrentX = 0
    plcSpotTypes.CurrentY = 0
    plcSpotTypes.Print "Include"
End Sub

Private Sub rbcDemo_Click(Index As Integer)
    If Index = 0 Then
        lbcSelection(0).Width = 4395            'use primary demo, show advt list box width of area
        lbcSelection(1).Visible = False
        lacDemo.Visible = False
    Else
        lbcSelection(0).Width = 1995            'demo list box required
        lbcSelection(1).Visible = True
        lacDemo.Visible = True
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

Private Sub rbcSortBy_Click(Index As Integer)
    If Index = 0 Then           'advt option doesnt use the vehicle groups
        cbcSet1.Enabled = False
        cbcSet1.ListIndex = 0               'default to No vehicle group selected
        lbcSelection(3).Visible = False
        ckcAllGroupItems.Visible = False
        lbcSelection(2).Width = 4395
        ckcAllGroupItems.Value = vbUnchecked
        rbcDP(0).Enabled = False
        rbcDP(1).Enabled = False
        ckcDetail.Value = vbUnchecked
        ckcDetail.Enabled = False
    Else
        cbcSet1.Enabled = True  'allow vehicle group to be selected
        rbcDP(0).Enabled = True
        rbcDP(1).Enabled = True
        ckcDetail.Enabled = True
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print "By"
End Sub


