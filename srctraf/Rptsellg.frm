VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelLg 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   720
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
   ScaleHeight     =   5505
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   36
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6690
      TabIndex        =   95
      Top             =   -90
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8835
      Top             =   -150
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
      Left            =   8025
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   -75
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
      Left            =   8310
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   -90
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcBudgetCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   255
      Sorted          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   4905
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.FileListBox lbcFileName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3060
      Pattern         =   "*.Dal"
      TabIndex        =   78
      Top             =   4935
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSMask.MaskEdBox mkcPhone 
      Height          =   210
      Left            =   1110
      TabIndex        =   77
      Tag             =   "The number and extension of the buyer."
      Top             =   4410
      Visible         =   0   'False
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16776960
      ForeColor       =   0
      MaxLength       =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(AAA) AAA-AAAA Ext(AAAA)"
      PromptChar      =   "_"
   End
   Begin VB.ListBox lbcSort 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2325
      Sorted          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   5085
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcAgyAdvtCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2175
      Sorted          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   4665
      Visible         =   0   'False
      Width           =   1455
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
      ScaleWidth      =   90
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4245
      Width           =   90
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   855
      Top             =   4770
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
         Left            =   1095
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "1"
         Top             =   315
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   165
         TabIndex        =   7
         Top             =   810
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   345
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
      Height          =   3690
      Left            =   90
      TabIndex        =   14
      Top             =   1755
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
         Height          =   3360
         Left            =   90
         ScaleHeight     =   3360
         ScaleWidth      =   4530
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   4530
         Begin VB.PictureBox plcSelC3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   -15
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   61
            Top             =   735
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2880
               TabIndex        =   69
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
               TabIndex        =   70
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   62
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
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   49
            Top             =   30
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
            Left            =   3240
            MaxLength       =   3
            TabIndex        =   45
            Top             =   30
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
            Left            =   510
            MaxLength       =   10
            TabIndex        =   47
            Top             =   45
            Width           =   1170
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   15
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   55
            Top             =   1200
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "incl3"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2985
               TabIndex        =   92
               Top             =   0
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "New/Changed"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2655
               TabIndex        =   58
               Top             =   0
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "All"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   56
               Top             =   0
               Width           =   510
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "New/Changed"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1245
               TabIndex        =   57
               Top             =   0
               Width           =   1560
            End
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   30
            ScaleHeight     =   240
            ScaleWidth      =   4140
            TabIndex        =   50
            Top             =   975
            Visible         =   0   'False
            Width           =   4140
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   3495
               TabIndex        =   94
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   3210
               TabIndex        =   93
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Advt"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2775
               TabIndex        =   54
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
               TabIndex        =   51
               Top             =   0
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Agency"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1290
               TabIndex        =   52
               Top             =   0
               Width           =   900
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Salesperson"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   2220
               TabIndex        =   53
               Top             =   0
               Width           =   1290
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
            Left            =   90
            MaxLength       =   10
            TabIndex        =   43
            Top             =   2415
            Width           =   1170
         End
         Begin VB.ComboBox cbcSel 
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
            Left            =   300
            TabIndex        =   59
            Top             =   135
            Visible         =   0   'False
            Width           =   4305
         End
         Begin VB.PictureBox plcSelC4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   71
            Top             =   480
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1650
               TabIndex        =   73
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   72
               Top             =   0
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2400
               TabIndex        =   74
               Top             =   0
               Width           =   1005
            End
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Active Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   60
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lacSelCTo1 
            Appearance      =   0  'Flat
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2325
            TabIndex        =   48
            Top             =   375
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# of Days"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2400
            TabIndex        =   44
            Top             =   75
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lacSelCTo 
            Appearance      =   0  'Flat
            Caption         =   "Active End Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   46
            Top             =   270
            Width           =   1380
         End
      End
      Begin VB.PictureBox pbcSelA 
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
         Height          =   960
         Left            =   120
         ScaleHeight     =   960
         ScaleWidth      =   4455
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox ckcADate 
            Caption         =   "Check3D1"
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
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   165
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox edcSelA 
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
            Left            =   630
            MaxLength       =   8
            TabIndex        =   18
            Top             =   105
            Width           =   1170
         End
         Begin VB.PictureBox plcSel1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   285
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   3555
            TabIndex        =   19
            Top             =   420
            Visible         =   0   'False
            Width           =   3585
            Begin VB.CheckBox ckcSel1 
               Caption         =   "Pending"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1635
               TabIndex        =   21
               Top             =   0
               Width           =   975
            End
            Begin VB.CheckBox ckcSel1 
               Caption         =   "Fed Only Events"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1620
            End
         End
         Begin VB.PictureBox plcSel2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   285
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   3255
            TabIndex        =   22
            Top             =   660
            Visible         =   0   'False
            Width           =   3285
            Begin VB.CheckBox ckcSel2 
               Caption         =   "M-F"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   570
            End
            Begin VB.CheckBox ckcSel2 
               Caption         =   "Sat"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   570
               TabIndex        =   24
               Top             =   0
               Width           =   570
            End
            Begin VB.CheckBox ckcSel2 
               Caption         =   "Sun"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1140
               TabIndex        =   25
               Top             =   0
               Width           =   600
            End
         End
         Begin VB.Label lacFromA 
            Appearance      =   0  'Flat
            Caption         =   "Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   17
            Top             =   165
            Width           =   420
         End
      End
      Begin VB.PictureBox pbcSelB 
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
         Height          =   990
         Left            =   135
         ScaleHeight     =   990
         ScaleWidth      =   4425
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   330
         Visible         =   0   'False
         Width           =   4425
         Begin VB.ComboBox cbcTo 
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
            TabIndex        =   34
            Top             =   555
            Width           =   1080
         End
         Begin VB.ComboBox cbcFrom 
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
            Left            =   765
            TabIndex        =   32
            Top             =   105
            Width           =   1080
         End
         Begin VB.Label lacTo 
            Appearance      =   0  'Flat
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   33
            Top             =   615
            Width           =   570
         End
         Begin VB.Label lacFrom 
            Appearance      =   0  'Flat
            Caption         =   "From"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   31
            Top             =   165
            Width           =   570
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
         Left            =   4605
         ScaleHeight     =   3420
         ScaleWidth      =   4455
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   12
            Left            =   0
            TabIndex        =   91
            Top             =   0
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   11
            Left            =   15
            TabIndex        =   90
            Top             =   0
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   10
            Left            =   15
            TabIndex        =   89
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   9
            Left            =   15
            TabIndex        =   88
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   8
            Left            =   15
            TabIndex        =   87
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   7
            Left            =   15
            TabIndex        =   86
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   6
            Left            =   15
            TabIndex        =   80
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   5
            Left            =   30
            TabIndex        =   79
            Top             =   45
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   4
            Left            =   30
            TabIndex        =   41
            Top             =   45
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   3
            Left            =   30
            TabIndex        =   40
            Top             =   45
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   2
            Left            =   60
            TabIndex        =   39
            Top             =   45
            Visible         =   0   'False
            Width           =   4350
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   1
            Left            =   15
            TabIndex        =   28
            Top             =   45
            Visible         =   0   'False
            Width           =   4395
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   0
            Left            =   15
            TabIndex        =   27
            Top             =   45
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   255
            TabIndex        =   29
            Top             =   0
            Width           =   3945
         End
         Begin VB.Label laclbcName 
            Appearance      =   0  'Flat
            Caption         =   "Compare To"
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
            Left            =   2300
            TabIndex        =   82
            Top             =   3060
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   2055
            TabIndex        =   83
            Top             =   3120
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
      TabIndex        =   37
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   35
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
         Width           =   1275
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
         Width           =   900
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
Attribute VB_Name = "RptSelLg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptsellg.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelLg.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imFTSelectedIndex As Integer
Dim imFromSelectedIndex As Integer
Dim imToSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
'Vehicle link file- used to obtain start date
'Delivery file- used to obtain start date
'Vehicle conflict file- used to obtain start date
'Spot projection- used to obtain date status
'Library calendar file- used to obtain post log date status
'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
'Import contract report
'Spot week Dump
Dim imTerminate As Integer
Dim ilAASCodes()  As Integer
Dim smRegionName As String
Dim lmRegionCode As Long

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
Private Sub cbcFrom_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcFrom.Text <> "" Then
            gManLookAhead cbcFrom, imBSMode, imComboBoxIndex
        End If
        imFromSelectedIndex = cbcFrom.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcFrom_Click()
    imComboBoxIndex = cbcFrom.ListIndex
    imFromSelectedIndex = cbcFrom.ListIndex
    mSetCommands
End Sub
Private Sub cbcFrom_GotFocus()
    If cbcFrom.Text = "" Then
        cbcFrom.ListIndex = 0
    End If
    imComboBoxIndex = cbcFrom.ListIndex
    gCtrlGotFocus cbcFrom
End Sub
Private Sub cbcFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcFrom_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcFrom.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcSel_Change()
    mSetCommands
End Sub
Private Sub cbcSel_Click()
    mSetCommands
End Sub
Private Sub cbcTo_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcTo.Text <> "" Then
            gManLookAhead cbcTo, imBSMode, imComboBoxIndex
        End If
        imToSelectedIndex = cbcTo.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcTo_Click()
    imComboBoxIndex = cbcTo.ListIndex
    imToSelectedIndex = cbcTo.ListIndex
    mSetCommands
End Sub
Private Sub cbcTo_GotFocus()
    If cbcTo.Text = "" Then
        cbcTo.ListIndex = cbcTo.ListCount - 1
    End If
    imComboBoxIndex = cbcTo.ListIndex
    gCtrlGotFocus cbcTo
End Sub
Private Sub cbcTo_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcTo_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcTo.SelLength <> 0 Then    'avoid deleting two characters
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
    Dim ilValue As Integer
    Dim ilIndex As Integer
    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAll Then
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
Private Sub ckcSel1_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSel1(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    mSetCommands
End Sub
Private Sub ckcSel2_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSel2(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
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
    mTerminate False
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
'
'
'       8-23-01 dh allow logs to use customized logos
'
Private Sub cmcGen_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilListIndex                                                                           *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim ilLoop As Integer
    Dim ilVpfIndex As Integer
    Dim slVehicleLogo As String
    Dim slSaveRptLogoName As String
    Dim slSavePath As String
    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    igOutput = frcOutput.Enabled
    igCopies = frcCopies.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False
    'If (sgRnfRptName = "L07") Or (sgRnfRptName = "L08") Or (sgRnfRptName = "L27") Then   'comml summary & 2 versions of comml schedule
    If (sgRnfRptName = "L07") Or (sgRnfRptName = "L27") Then    'comml summary & 2 versions of comml schedule; 1-12-12 L08 is a valid crystal report
    'L08 converted to Crystal L36 11-3-99
        igUsingCrystal = False
    Else
        igUsingCrystal = True
    End If
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        'Setup correct logo to print for this vehicles log or CP
        'Rename existing rptlogo.bmp to vehicles logo; then rename back later
        ilVpfIndex = -1
        'For ilLoop = 0 To UBound(tgVpf) Step 1
        '    If igcodes(0) = tgVpf(ilLoop).iVefKCode Then
            ilLoop = gBinarySearchVpf(igcodes(0))
            If ilLoop <> -1 Then
                ilVpfIndex = ilLoop
        '        Exit For
            End If
        'Next ilLoop
        'ilRet = MsgBox(Str$(ilVpfIndex) & ", " & Trim$(tgVpf(ilVpfIndex).sCPLogo) & ", " & Trim$(Str$(tgVpf(ilVpfIndex).irnfCertCode)))



'12-17-14 no longer required to save rptlogo with a different name.  send the customized name of logo to formula "LogoLocation"
'        If ilVpfIndex >= 0 Then     '8-23-01 dont test for CPs, And InStr(sgRnfRptName, "C") = 1 Then
'            If tgVpf(ilVpfIndex).sCPLogo <> "   " Then
'                'slSavePath = Trim$(sgRptPath)
'                'sgRptPath = sgRootDrive & "csi\"
''                sgRptPath = "c:\csi\"       'force the link to bmp to c:\csi
'                'Rename the original rptlogo.bmp to a saved name, then name the vehicle logo to rptlogo.bmp for crystal reporting
'                'slSaveRptLogoName = Trim$(sgRptPath) & Trim$("rptlogo.bmp")
'                slSaveRptLogoName = Trim$(sgLogoPath) & Trim$("rptlogo.") & sgRptLogoExt
'
'                On Error GoTo RptLogoErr:
'                'Name slSaveRptLogoName As Trim$(sgRptPath) & "savelogo.bmp"
'                Name slSaveRptLogoName As Trim$(sgLogoPath) & "savelogo." & sgRptLogoExt
'                'slVehicleLogo = Trim$(sgRptPath) & "G" & Trim$(tgVpf(ilVpfIndex).sCPLogo) & Trim$(".bmp")
'                slVehicleLogo = Trim$(sgLogoPath) & "G" & Trim$(tgVpf(ilVpfIndex).sCPLogo) & Trim$(".bmp")
'
'                On Error GoTo CPLogoErr:
'                'Name Trim$(slVehicleLogo) As Trim$(sgRptPath) & Trim$("rptlogo.bmp")
'                Name Trim$(slVehicleLogo) As Trim$(sgLogoPath) & Trim$("rptlogo.bmp")   'the customized logos are always .bmp for now
'                'sgRptPath = slSavePath
'            End If
'        End If
        sgReportListName = sgRnfRptName
        If Not gGenReportLg() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If
    
'        If tgVpf(ilVpfIndex).sCPLogo <> "   " Then          'only re-establish logo if customized
'            If Not gSetFormula("LogoLocation", "'" & Trim$(sgLogoPath) & Trim$("rptlogo.bmp") & "'") Then
'                ilRet = ilRet
'            End If
'        End If
'12-17-14    Send the logo name to the formula LogoLocation rather than having it hardcoded to c:\csi\rptlogo
        If Trim$(tgVpf(ilVpfIndex).sCPLogo) = "" Then          'if not using customized logs, assume rptlogo
            If Not gSetFormula("LogoLocation", "'" & Trim$(sgLogoPath) & Trim$("rptlogo." & sgRptLogoExt) & "'") Then
                ilRet = ilRet
            End If
        Else
            If Not gSetFormula("LogoLocation", "'" & Trim$(sgLogoPath) & "g" & Trim$(tgVpf(ilVpfIndex).sCPLogo) & Trim$(".bmp") & "'") Then
                ilRet = ilRet
            End If
        End If

        ilRet = gCmcGenLg(lmRegionCode, smRegionName)
        '-1 is a Crystal failure of gSetSelection or gSEtFormula
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
        ElseIf ilRet = 0 Then           '0 = invalid input data, stay in
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        ElseIf ilRet = 2 Then           'successful from Bridgereport
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
       End If
       '1 falls thru - successful crystal report
       'L43 required VOF information
        If sgRnfRptName = "L09" Or sgRnfRptName = "L17" Or sgRnfRptName = "L41" Then          'Copy play list by vehicle (l17 = plylist showing reel # field as cart #)
            Screen.MousePointer = vbHourglass
            gCRPlayListGenLg False    'no VOF information
            Screen.MousePointer = vbDefault

        '9-11-00 ElseIf sgRnfRptName = "L43" Then          '
        ElseIf sgRnfRptName = "L70" Then          '9-11-00
            Screen.MousePointer = vbHourglass
            gCRPlayListGenLg True     'get VOF information
            Screen.MousePointer = vbDefault
            '12-18-02 Create C82 = customized version of L10
            '9-20-16 C89 copy of L10, excludes Politicals.  clients runs 2 "L" logs, have to make this a "C" log to get 3 logs printed
            'L89 (copy of L10 with columns removed), added 2/22/19
        ElseIf sgRnfRptName = "L10" Or sgRnfRptName = "L28" Or sgRnfRptName = "L34" Or sgRnfRptName = "L40" Or sgRnfRptName = "C82" Or sgRnfRptName = "C89" Or sgRnfRptName = "L89" Then          'Commercial Schedule (AMFM version)
            Screen.MousePointer = vbHourglass
            gCreate7Day
            Screen.MousePointer = vbDefault
        '9-11-00 ElseIf sgRnfRptName = "L11" Or sgRnfRptName = "L32" Or sgRnfRptName = "L35" Or sgRnfRptName = "C22" Or sgRnfRptName = "C23" Then             'Commercial Summary (AMFM version-15 DP), AMFM or Jones , AMFM version 18DP
        ElseIf sgRnfRptName = "L11" Or sgRnfRptName = "L32" Or sgRnfRptName = "L35" Or sgRnfRptName = "C72" Or sgRnfRptName = "C73" Or sgRnfRptName = "C84" Then              '9-11-00 Commercial Summary (AMFM version-15 DP), AMFM or Jones , AMFM version 18DP
            Screen.MousePointer = vbHourglass
            gCmlSum15DP
            Screen.MousePointer = vbDefault
        ElseIf sgRnfRptName = "L14" Then          'log M-F, 2 columns across with page skips by user defined
                                                  'all times for all 5 days are printed on a page (most likely by daypart)
                                                   'ie:  vehicle runs M-5 6a-8P with page skips at 12n
                                                   'One page is printed containing Mo 6a-12n, T6a-12n....thru Fri 6a-12n
                                                   'Next page is Mo 12n-8p, tu 12n-8p....thru fri 12n-8p
            Screen.MousePointer = vbHourglass
            gL14PageSkips 5                      'pass flag to process 5 days
            Screen.MousePointer = vbDefault
        ElseIf sgRnfRptName = "L21" Or sgRnfRptName = "L26" Then   'l21 sorts & skips by named avail, l26 does not use named avail
        'log 7 days, 2 columns across with page skips indicated per day.
                                                'Each days page skip should be at the same time.  One day
                                                'processed at a time (This log is the duplicate of L14 except
                                                'it processes M-Su instead of M-f
            Screen.MousePointer = vbHourglass
            gL14PageSkips 7                     'pass flag to process 7 days
            Screen.MousePointer = vbDefault
       '9-11-00  ElseIf sgRnfRptName = "L29" Or sgRnfRptName = "C14" Or sgRnfRptName = "C19" Or sgRnfRptName = "L44" Then   'l44 needs the DP descriptions from Interface table
        ElseIf sgRnfRptName = "L29" Or sgRnfRptName = "C14" Or sgRnfRptName = "C19" Or sgRnfRptName = "L71" Or sgRnfRptName = "L74" Then   '9-11-00 l71 needs the DP descriptions from Interface table, or named avails comments
            'Sheridan log (setup major sort, all back-to back breaks get same seq #)
                                                'or CP to setup incremental Comml IDs from AVail comment field
                                                'L71 doesnt care about the sequence #s, wont be used for sorts in that log
            Screen.MousePointer = vbHourglass
            gSetSeqL29
            Screen.MousePointer = vbDefault
        End If
        '9-11-00 ElseIf sgRnfRptName = "C20" Or sgRnfRptName = "C21" Then        'update the header and footer comments into the ODF,
        '12-6-00 need to update the header/footer pointers into the ODF for customized CPs
        If sgRnfRptName = "C01" Or sgRnfRptName = "C70" Or sgRnfRptName = "C71" Or sgRnfRptName = "C74" Or sgRnfRptName = "C75" Or sgRnfRptName = "C76" Or sgRnfRptName = "C77" Or sgRnfRptName = "C78" Or sgRnfRptName = "C79" Or sgRnfRptName = "C81" Or sgRnfRptName = "C83" Then           '1-5-04 update the header and footer comments into the ODF,
                                                                        'fixup the avail times to be running times for m-f /m-su logs (days across)
            gFixAirTimes
        ElseIf sgRnfRptName = "C85" Or sgRnfRptName = "C86" Or sgRnfRptName = "C87" Then        '5-8-12 add c87
            gFixAirTimes
        'ElseIf sgRnfRptName = "L70" Or sgRnfRptName = "L71" Or sgRnfRptName = "L72" Or sgRnfRptName = "L73" Or sgRnfRptName = "L74" Or sgRnfRptName = "L75" Or sgRnfRptName = "L76" Or sgRnfRptName = "L77" Or sgRnfRptName = "L78" or sgRnfRptname = "L79" Then
         ElseIf ((Val(Mid(sgRnfRptName, 2, 2)) >= 70) And (Mid(sgRnfRptName, 1, 1) = "L")) Then
            gFixAirTimes
        ElseIf sgRnfRptName = "L04" Then
            gFixAirTimes
        End If
        
        '8-10-16 take
'        ElseIf sgRnfRptName = "L36" Or sgRnfRptName = "L38" Or sgRnfRptName = "L08" Or sgRnfRptName = "L88" Then              '11-3-99 another version of Comml summary (l08 converted).  L36=save to file version, l38 = printed version
         If sgRnfRptName = "L36" Or sgRnfRptName = "L38" Or sgRnfRptName = "L08" Or sgRnfRptName = "L88" Then              '11-3-99 another version of Comml summary (l08 converted).  L36=save to file version, l38 = printed version
                                               'or CP to setup incremental Comml IDs from AVail comment field
            Screen.MousePointer = vbHourglass
            gL36ComlSmry
            Screen.MousePointer = vbDefault
        ElseIf sgRnfRptName = "L37" Or sgRnfRptName = "L39" Or sgRnfRptName = "C80" Then    '5-17-01 C80 , 11-4-99 another version of Comml schedule (l27 converted) . l37=save to file veresion, l39 = printed version
            Screen.MousePointer = vbHourglass
            gL37ComlSch (Val(sgLogUserCode))
            Screen.MousePointer = vbDefault
        End If

        If sgRnfRptName = "L87" Then
            Screen.MousePointer = vbHourglass
            gGenL87Master
            Screen.MousePointer = vbDefault
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
            'ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-19-01
        End If
        '
        
        '12-17-14 no longer required to saved and restore rptlogo with a different name.  customized name of logo was sent to formula "LogoLocation"

        'Rename the vehicles logo back to rptlogo.bmp
        'Only rename back to rptlog.bmp if a valid vehicle option found
'        If ilVpfIndex >= 0 Then     '5-25-01 avoid subscript out of range if there wasn't a valid vpf index (i.e. -1)
'            If (ilVpfIndex >= 0) And (tgVpf(ilVpfIndex).sCPLogo <> "   ") Then  '8-23-01 dont test for cps, And InStr(sgRnfRptName, "C") = 1 Then
'                'slSavePath = Trim$(sgRptPath)
'                slSavePath = Trim$(sgLogoPath)
'                '5676
'                'sgRptPath = sgRootDrive & "csi\"
'                'sgRptPath = "c:\csi\"
'                'ilRet = MsgBox(Trim$(sgRptPath) & "rptlogo.bmp" & ", " & Trim$(sgRptPath) & "G" & Trim$(tgVpf(ilVpfIndex).scplogo) & Trim$(".bmp"))
'                On Error GoTo RptLogoErr:
'                'Name Trim$(sgRptPath) & "rptlogo.bmp" As Trim$(sgRptPath) & "G" & Trim$(tgVpf(ilVpfIndex).sCPLogo) & Trim$(".bmp")
'                Name Trim$(sgLogoPath) & "rptlogo.bmp" As Trim$(sgLogoPath) & "G" & Trim$(tgVpf(ilVpfIndex).sCPLogo) & Trim$(".bmp")
'                On Error GoTo SaveLogoErr:
'                'Name Trim$(sgRptPath) & "savelogo.bmp" As Trim$(sgRptPath) & Trim$("rptlogo.bmp")
'                Name Trim$(sgLogoPath) & "savelogo." & sgRptLogoExt As Trim$(sgLogoPath) & Trim$("rptlogo.") & sgRptLogoExt
'                'sgRptPath = slSavePath
'            End If
'        End If
    Next ilJobs
    imGenShiftKey = 0
    '9-11-00 If sgRnfRptName = "L09" Or sgRnfRptName = "L17" Or sgRnfRptName = "L41" Or sgRnfRptName = "L43" Then      'copy playlist by vehicle (l17 = plylist showing reel # field as cart #)
    If sgRnfRptName = "L09" Or sgRnfRptName = "L17" Or sgRnfRptName = "L41" Or sgRnfRptName = "L70" Then      '9-11-00 copy playlist by vehicle (l17 = plylist showing reel # field as cart #)
        Screen.MousePointer = vbHourglass
        gCRPlayListClear
        Screen.MousePointer = vbDefault
        '9-20-16 C89 copy of L10, excludes Politicals.  clients runs 2 "L" logs, have to make this a "C" log to get 3 logs printed
        'L89 (copy of L10 with columns removed), added 2/22/19
    ElseIf sgRnfRptName = "L10" Or sgRnfRptName = "L28" Or sgRnfRptName = "L34" Or sgRnfRptName = "L37" Or sgRnfRptName = "L39" Or sgRnfRptName = "L40" Or sgRnfRptName = "C80" Or sgRnfRptName = "C82" Or sgRnfRptName = "C89" Or sgRnfRptName = "L89" Then        '7Day commercial schedule (AMFM version)
        Screen.MousePointer = vbHourglass
        gClearSvr
        Screen.MousePointer = vbDefault
    '9-11-00 ElseIf sgRnfRptName = "L11" Or sgRnfRptName = "L32" Or sgRnfRptName = "L35" Or sgRnfRptName = "L36" Or sgRnfRptName = "L38" Or sgRnfRptName = "C22" Or sgRnfRptName = "C23" Then     'Comml Summary (AMFM version-15DP), Jones version or AMFM version (18DP)
    ElseIf sgRnfRptName = "L11" Or sgRnfRptName = "L32" Or sgRnfRptName = "L35" Or sgRnfRptName = "L36" Or sgRnfRptName = "L38" Or sgRnfRptName = "C72" Or sgRnfRptName = "C73" Or sgRnfRptName = "C84" Or sgRnfRptName = "L08" Or sgRnfRptName = "L88" Then      '5-4-04 Comml Summary (AMFM version-15DP), Jones version or AMFM version (18DP)
        Screen.MousePointer = vbHourglass
        'gClearGrf
        gCRGrfClear         '8-20-13 use only 1 common grf clear rtn which changes the way records are removed
        Screen.MousePointer = vbDefault
    ElseIf sgRnfRptName = "L87" Then
        Screen.MousePointer = vbHourglass
        gCrCbfClear
        Screen.MousePointer = vbDefault
    End If
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    'frcWhen.Enabled = igWhen
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    Exit Sub
RptLogoErr:
    On Error GoTo 0
    Resume Next
CPLogoErr:
    On Error GoTo 0
    '5676
   ' MsgBox slVehicleLogo & "  does not exist in " & sgRootDrive & "CSI; Copy into " & sgRootDrive & "CSI"
    MsgBox "G" & Trim$(tgVpf(ilVpfIndex).sCPLogo) & Trim$(".bmp") & " does not exist in " & sgLogoPath & "; Copy into " & sgLogoPath
    Resume Next
SaveLogoErr:
    On Error GoTo 0
    'MsgBox "SaveLogo.bmp does not exist in C:\CSI"
    Resume Next
'cmcGenErr:
'    ilDDFSet = True
'    Resume Next
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
    If (KeyAscii <= 32) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KEYDOWN) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcSelA_Change()
    mSetCommands
End Sub
Private Sub edcSelA_GotFocus()
    gCtrlGotFocus edcSelA
End Sub
Private Sub edcSelCFrom_Change()
    mSetCommands
End Sub
Private Sub edcSelCFrom_GotFocus()
    gCtrlGotFocus edcSelCFrom
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
End Sub
Private Sub edcSelCTo_Change()
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
    If imTerminate = -99 Then
        Exit Sub
    End If
    If imTerminate Then 'Used for print only
        'mTerminate
        cmcCancel_Click
        Exit Sub
    End If
    'RptSelLg.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Erase tgAirNameCodeLg
    Erase tgCSVNameCode
    Erase tgSellNameCodeLg
    Erase tgRptSelLgSalespersonCode
    Erase tgRptSelLgAgencyCode
    Erase tgRptSelLgAdvertiserCode
    Erase tgRptSelLgNameCode
    Erase tgRptSelLgBudgetCode
    'Erase tgMultiCntrCode
    'Erase tgManyCntCode
    Erase tgRptSelLgDemoCode
    'Erase tgSOCode
    Erase igcodes
    PECloseEngine
    
    Set RptSelLg = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcRptType_Click()
    ReDim ilAASCodes(0 To 1) As Integer
    rbcSelCInclude(2).Visible = False
    mSetCommands
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    Dim ilListIndex As Integer
    ReDim ilAASCodes(0 To 1) As Integer
    If Not imAllClicked Then
        ilListIndex = lbcRptType.ListIndex
            imSetAll = False
            ckcAll.Value = vbUnchecked  'false
            imSetAll = True
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
Private Sub lbcSort_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  'False
        imSetAll = True
    End If
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Private Sub mInit()
    Dim ilRet As Integer
    Dim ilMultiTable As Integer

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        imTerminate = -99
        Exit Sub
    End If
    'Set options for report generate
'VB6**    hdJob = rpcRpt.hJob
    ilMultiTable = True
    'dummy = LlSetOption(hdJob, LL_OPTION_HELPAVAILABLE, False)
'VB6**    ilDummy = LlSetOption(hdJob, LL_OPTION_SORTVARIABLES, True)
'VB6**    ilDummy = LlSetOption(hdJob, LL_OPTION_ONLYONETABLE, Not ilMultiTable)

    imAllClicked = False
    imSetAll = True
    'gCenterStdAlone RptSelLg
    'RptSelLg.Move -90, -90, 30, 30      'make form small and out of the way so its not seen
    RptSelLg.Move -330, -330, 30, 30      'make form small and out of the way so its not seen
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
    pbcSelA.Visible = False
    pbcSelB.Visible = False
    pbcSelC.Visible = False
    lbcRptType.Clear

    Screen.MousePointer = vbHourglass
    'If (igRptType = 0) Or (igRptType = 1) Or (igRptType = 2) Then
    If igOutputTo = 0 Then
        rbcOutput(0).Value = True
    ElseIf igOutputTo = 2 Then
        rbcOutput(2).Value = True
    Else
        rbcOutput(1).Value = True          'always print these automatically generated reports
    'rbcOutput(0).Value = True           'display -- for test purposes only
    End If

    ckcSelC3(0).Value = vbUnchecked 'False
    ckcSelC3(1).Value = vbUnchecked 'False
    ckcSelC3(2).Value = vbUnchecked 'False
    ckcSelC3(3).Value = vbUnchecked 'False
    If igZones = 0 Then
        ckcSelC3(0).Value = vbChecked   'True
        ckcSelC3(1).Value = vbChecked   'True
        ckcSelC3(2).Value = vbChecked   'True
        ckcSelC3(3).Value = vbChecked   'True
    ElseIf igZones = 1 Then
        ckcSelC3(0).Value = vbChecked   'True
    ElseIf igZones = 2 Then
        ckcSelC3(1).Value = vbChecked   'True
    ElseIf igZones = 3 Then
        ckcSelC3(2).Value = vbChecked   'True
    Else
        ckcSelC3(3).Value = vbChecked   'True
    End If
    cmcGen_Click
    imTerminate = True
    Exit Sub
    'End If
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSelLg
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*
'*            Special mParseCmmdLine for "Log" process *
'*            Assumes that only the LOG reports come
'*            thru here
'*            10/19/99 D Levine:  Add 13/14 parameter for
'             affiliate system
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVpfIndex As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim hlRnf As Integer
    Dim ilRnfRecLen As Integer
    Dim tlRnf As RNF
    Dim tlRnfSrchKey As INTKEY0
    Dim ilFound As Integer
    'RNF name   RNF Code     Description
    'L01        83           Log
    'L02        88           Short form Log (time option)
    'L03        198          Short form log (break option)
    'L04        89           Long form log (time option)
    'L05        199          Long form log (break option)
    'L06        90           Log4
    'L07        84           commercial schedule
    'L08        207          Commercial Summary
    'L09        79           PlayList by Vehicle
    'L10        200          Commercial Schedule (AMFM version)
    'L11        208          Commerical Summary (AMFM Version - 15DP across)
    'L12        224          Shadow 5 days across (m-f)
    'L13        225          Shadow 2 days across (sa-su)
    'L14        227          Shadow 5 days vertical
    'l15
    'l16
    'l17                     Global Copy Playlist (l17 = plylist showing reel # field as cart #)
    'L27        243          ABC Comml Schdule w/new event IDs
    'L28        245          Copy of L10 (coml schedule) without spots ordered column
    'L29        246          Sheridan log, each line of spot data is given a unique sort seq #
                             'any avail back to back is given same sort seq # so that these spots
                             'all print on the same line
    'L30        249          Sheridan (include anncr & spots & page ejects)
    'L31        251          Jones log
    'L32        252          Jones log
    'l33        255          copy of L04 (AMFM)
    'L34        256          copy of L10 (AMFM)
    'L35        257          copy of L11 (AMFM, 18 DP)
    'L36        261          L08 (commercial summary) converted to Crystal  for export 11-3-99
    'l37        262          L07/L27 converted to Crystal for export 11-5-99 (coml sched)
    'l38        263          L38 (coml summary) converted to Crystsal for printing 11-11-99
    'l39        264          l39 (coml sched) converted to crystal for printing 11-11-99
    'l40        266          l40 (comml sched for amfm 1-3-00 Show predefined dayparts from vehicle options table for every spot
    'l41        269          copy playlist by isci (no description in header)
    'l42        272          Sports byline
    'l43        278          Copy Playlist by isci/advt with vof options
    'l43 changed to l70 9-11-00
    'l44        279          Comml schedule with vof options
    'l44 changed to l71 9-11-00
    'l70        279
    'l71        280
    'l72        295           copy of L05, made customized
    'l73        297           copy of l01, made customized
    'L80        367          log (copy of C78) with spot counts & daypart notations (added 3-9-11)
    'l83                     duplicate of L73 (removed comments and shows game info) 5-29-13
    'l85                     duplicate of l75 (removed comments and shows game info) 5-29-13.  7-15-13 originally l84, changed to l85 (jf), then chg from l84 to c88
    'L88                     duplicate of L38 (include psa avails; L38 excludes all avail names that do no start with "N")
    'C01        86           Short Form CP Daily
    'C02        201          Short form CP 5-day
    'C03        202          short form CP 7-day
    'c04        87           Long form CP daily
    'C05        203          Long form CP 5-day
    'C06        205          Long form CP 7-day
    'C07        204          CP w/Contract #
    'c08        222
    'C09        223
    'C13        244
    'C14        247     'Sheridan CP (prepass sets up incrmental COml ID # from Avails comment field)
    'C15        248     'Sheridan CP Show break #, length and advt
    'C16        250     'Sheridcan CP
    'c17                'Media America CP
    'c18        267     'CP with copy script
    'c19        268     'cp with copy script & dayparts from vehicle options interface table
    'c20        275     'MAI m-f CP with VOF options
    'c20 changed to c70 9-11-00
    'c21        276     'MAI m-su CP with VOF options
    'c21 changed to c71 9-11-00
    'C22        277     'MAI m-f CP (like coml summary) with VOF options
    'c22 changed to c72 9-11-00
    'C23        278     'MAI m-s CP (like coml summary) with VOF options
    'c23 changed to c73 9-11-00
    'c24        281     'MAI CP with vof options
    'c24 changed to c74 9-11-00
    'c75        286     'MAI CP copy of C74 WITHOUT the day showing insubheader caption (meant for ONE DAY LOG ONLY)
    'c76        287     'CP with bb using avail comments
    'c77        288     'cp with bb using avail comments (no dates for each new date)
    'c78        289     'cp with spot counts & daypart notations
    'c79        291     'same as c77 but with smaller font and no top/bottom margins
    'c80        296     'copy of l39 but as a CP and using customized options
    'C88                     New sports log with break #; using new event type for local break definitions and running break #s
    slCommand = sgCommandStr    'Command$
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    ''igStdAloneMode defined as "Debug" mode
    'igStdAloneMode = True 'Switch from/to stand alone mode-No DDE
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    sgCallAppName = ""
    '    ilTestSystem = False  'True 'False
    '    'Change these following parameters to debug
    '    slCommStartDate = "12/20/99"
    '    slCommDays = "7"
    '    'slCommVehCode = "91" 'Gold Mine
    '    slCommVehCode = "20"   'ESPN
    '    slRnfCode = "296"      'l29
    '    slGenDate = "10/10/01"
    '    slGenTime = "1:01:07P"
    '    slLogType = "L"
    '    'parms: Logs^Test (or Prod)\ user name\jobcode(igrptcalltype)\Rnfcode (igrpttype)\usercode\Start Date\#days\STartTime\EndTime\VehCode\Zones\DisplPrint\format type(save to file)\filename (save tofile) ,GenDate,GenTime
    '    'Zones: 0 = all, 1 = EST, 2 = CST, 3= MST, 4=PST
    '    'DisplPrint : 0=display, 1 = print, 2 save to file
    '    'for save to file option only : if negative its affiliate (reserve sign and subtract 1 from index)
    '    'for save to file option only:  file name to save file to
    '    'slGenDate - generation date of ODF, used as key for filtering
    '    'slGenTime - generation time of ODF used as key for filtering
    '    'sgLogType l=log, c=cp, o=other
    '    'lmRegionCode - if split networks, region code, else 0 (5-16-08 currently L78)
    '    'smRegionName - if split networks, region name, else blank (5-16-08 currently L78)
    '    slCommand = "Logs^Test\Guide\6\" & slRnfCode & "\1\" & slCommStartDate & "\" & slCommDays & "\12M\12M\" & slCommVehCode & "\0" & "\0" & "\6" & "\l38.rtf\" & slGenDate & "\" & slGenTime & "\" & slLogType
    '    imShowHelpmsg = False
    'Else
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
    'gInitStdAlone RptSelLg, slStr, ilTestSystem

    'ilRet = gParseItem(slCommand, 2, "\", smSelectedRptName)    ' report name not used anymore
    ilRet = gParseItem(slCommand, 3, "\", slStr)                'report call type (always to be a 6)
    igRptCallType = Val(slStr)      'always going to be LOGSJOB
    ilRet = gParseItem(slCommand, 4, "\", slStr)        'rnf code
    igRnfCode = Val(slStr)
    ilRet = gParseItem(slCommand, 5, "\", sgLogUserCode)        'log user code
    ilRet = gParseItem(slCommand, 6, "\", sgLogStartDate)
    ilRet = gParseItem(slCommand, 7, "\", sgLogNoDays)
    ilRet = gParseItem(slCommand, 8, "\", sgLogStartTime)
    ilRet = gParseItem(slCommand, 9, "\", sgLogEndTime)
    ilRet = gParseItem(slCommand, 10, "\", slStr)               'vehicle code

    igNoCodes = 1                                               'force to only one vehicle at a time
    'imNoCodes = Val(slStr)
    ReDim igcodes(0 To igNoCodes) As Integer
    If igNoCodes > 0 Then
        For ilLoop = 0 To igNoCodes - 1 Step 1
            igcodes(ilLoop) = Val(slStr)
        Next ilLoop
    Else
        igcodes(0) = -1
    End If
    hlRnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlRnf, "", sgDBPath & "Rnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlRnf)
        btrDestroy hlRnf
        Exit Sub
    End If
    ilRnfRecLen = Len(tlRnf)
    tlRnfSrchKey.iCode = igRnfCode
    ilRet = btrGetEqual(hlRnf, tlRnf, ilRnfRecLen, tlRnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        sgRnfRptName = Trim$(tlRnf.sName)
        ilRet = btrClose(hlRnf)
        btrDestroy hlRnf
    Else
        ilRet = btrClose(hlRnf)
        btrDestroy hlRnf
        Exit Sub
    End If
    ilRet = gParseItem(slCommand, 11, "\", slStr)               'time zones (0=all, 1=est, 2=cst, 3=mst, 4=pst)
    igZones = Val(slStr)

    ilVpfIndex = -1
    'For ilLoop = 0 To UBound(tgVpf) Step 1
    '    If igcodes(0) = tgVpf(ilLoop).iVefKCode Then
        ilLoop = gBinarySearchVpf(igcodes(0))
        If ilLoop <> -1 Then
            ilVpfIndex = ilLoop
    '        Exit For
        End If
    'Next ilLoop

    'Determine the index of the time zone based on the vehicle options table
    If igZones > 0 And ilVpfIndex > 0 Then                         'getting selective time zone
        If igZones = 1 Then
            slStr = "EST"
        ElseIf igZones = 2 Then
            slStr = "CST"
        ElseIf igZones = 3 Then
            slStr = "MST"
        ElseIf igZones = 4 Then
            slStr = "PST"
        Else
            igZones = -1
            Exit Sub
        End If
        ilFound = False
        'For ilLoop = 1 To 4         'make sure valid zone in table
        For ilLoop = 0 To 3         'make sure valid zone in table
            If slStr = Trim$(tgVpf(ilVpfIndex).sGZone(ilLoop)) Then
                ilFound = True
                Exit For
            ElseIf slStr = Trim$(tgVpf(ilVpfIndex).sGZone(ilLoop)) Then
                ilFound = True
                Exit For
            ElseIf slStr = Trim$(tgVpf(ilVpfIndex).sGZone(ilLoop)) Then
                ilFound = True
                Exit For
            ElseIf slStr = Trim$(tgVpf(ilVpfIndex).sGZone(ilLoop)) Then
                ilFound = True
                Exit For
            End If
        Next ilLoop



        'slStr = Trim$(tgVpf(ilVpfIndex).sGZone(igZones))
        'If slStr = "EST" Then
        '    igZones = 1
        'ElseIf slStr = "CST" Then
        '    igZones = 2
        'ElseIf slStr = "MST" Then
        '    igZones = 3
        'ElseIf slStr = "PST" Then
        '    igZones = 4
        'Else
        '    igZones = 0                  'force all, time zone not found in time zone table
        'End If
    End If

    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    ilRet = gPopExportTypes(cbcFileType)         '10-19-01

    ilRet = gParseItem(slCommand, 12, "\", slStr)               'Display or Print flag
    igOutputTo = Val(slStr)
    If igOutputTo = 0 Then                          'display
        rbcOutput(0).Value = True
        rbcOutput(1).Value = False
    ElseIf igOutputTo = 2 Then              '10/19/99
        rbcOutput(2).Value = True
        rbcOutput(0).Value = False
        rbcOutput(1).Value = False
        ilRet = gParseItem(slCommand, 13, "\", slStr)               'Save File Index
        imFTSelectedIndex = Val(slStr)
        ilRet = gParseItem(slCommand, 14, "\", slStr)               'Save File Name
        edcFileName.Text = Trim$(slStr)
    Else
        rbcOutput(1).Value = True
        rbcOutput(0).Value = False
    End If
    'gen date & time for affiliate system only
    ilRet = gParseItem(slCommand, 15, "\", slStr)
    gPackDate slStr, igNowDate(0), igNowDate(1)
    ilRet = gParseItem(slCommand, 16, "\", slStr)
    gPackTime slStr, igNowTime(0), igNowTime(1)
    '5-25-01
    'igNowDate & igNowTime is used directly to retrieve ODF data when going direct to Crystal
    'igODFGenDate & igODFGenTime is used to gather ODF records to create prepass (igNowDate & igNowTime
    'are destroyed before prepass is created)
    igODFGenDate(0) = igNowDate(0)
    igODFGenDate(1) = igNowDate(1)
    igODFGenTime(0) = igNowTime(0)
    igODFGenTime(1) = igNowTime(1)
    '10-10-01 get generation time of ODF
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgGenTime
    ilRet = gParseItem(slCommand, 17, "\", sgLogType)       '6-19-00 l=log, c=cp, o=other
    ilRet = gParseItem(slCommand, 18, "\", slStr)           'region code
    lmRegionCode = Val(slStr)
    ilRet = gParseItem(slCommand, 19, "\", slStr)            'region name
    smRegionName = Trim$(slStr)                               '5-16-08 region name, blank if full network

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
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RptSelLg
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
        imFirstTime = False
        mInitReport
        If imTerminate Then 'Used for print only
            'mTerminate
            cmcCancel_Click
            Exit Sub
        End If
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub
Private Sub rbcSelCInclude_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelCInclude(Index).Value
    'End of coded added

    If Value Then
        mSetCommands
    End If
End Sub
Private Sub rbcSelCSelect_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelCSelect(Index).Value
    'End of coded added

    If Value Then
        mSetCommands
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    mTerminate False
End Sub
Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    plcSelC3.Print "Zone"
End Sub
Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print "Include"
End Sub
Private Sub plcSelC1_Paint()
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    plcSelC1.Print "Select"
End Sub
Private Sub plcSelC4_Paint()
    plcSelC4.CurrentX = 0
    plcSelC4.CurrentY = 0
    plcSelC4.Print "Option"
End Sub
