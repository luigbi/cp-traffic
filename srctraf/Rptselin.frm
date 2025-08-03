VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelIn 
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
            MaxLength       =   8
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
            MaxLength       =   8
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
Attribute VB_Name = "RptSelIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselin.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelIn.Frm
'            Invoice Print program
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
'Import contract report
'Spot week Dump
Dim imTerminate As Integer
Dim ilAASCodes()  As Integer
'10016
Dim myPDFEmailLogger As CLogger

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

Private Sub cmcGen_Click()
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilListIndex As Integer
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim slFullSelection As String
    Dim ilCounter As Integer
    Dim vlArray As Variant
    Dim myTempReport As CRAXDRT.Report
    Dim ilFound As Integer 'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
    Dim ilLoop As Integer
    
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
    igUsingCrystal = True
    ilNoJobs = 2                'default to 2 pass report (Affdavit detail & affidavit summary)
    ilStartJobNo = 1
    '11-2-01 igInvoiceType = 0 Previous Bridge report (ordered,aired,reconciled form)
    'Dan multiple reports as one 11/19/08.  Multiple reports are not just 'generated' from inside rptSelIn, but also from outside--Invoice.frm.-- which then calls here multiple times.
    'therefore, report may already be set.
    If Not bgComingFromInvoice Then
        ReDim igPreviousTimes(0 To 1, 0 To 0) As Integer
        ReDim igPreviousDates(0 To 1, 0 To 0) As Integer
    End If
    If ogReport Is Nothing Then
        Set ogReport = New CReportHelper
    End If
    If igInvoiceType >= 0 And igInvoiceType <= 7 Then       '01-18-07 6 = combined Air Time & NTR, 3-8-12 3-col aired
        ilNoJobs = 1
        ilStartJobNo = 1
    'added by Dan for multiple reports
    Else
        ogReport.iLastPrintJob = 2
    End If
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportIn() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If
        ilRet = gCmcGenIn(ilListIndex, imGenShiftKey, sgLogUserCode)
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


        'Setup correct logo to print for this vehicles log or CP
        'Rename existing rptlogo.bmp to vehicles logo; then rename back later
        'ilVpfIndex = -1
        'For ilLoop = 0 To UBound(tgVpf) Step 1
        ''    If igCodes(0) = tgVpf(ilLoop).iVefKCode Then
        '        ilVpfIndex = ilLoop
        '        Exit For
        '    End If
        'Next ilLoop
        'If ilVpfIndex >= 0 Then
        '    If tgvpf(ilVpfIndex).scplogo <> "   " Then
                'Rename the original rptlogo.bmp to a saved name, then name the vehicle logo to rptlogo.bmp for crystal reporting
                'slSaveRptLogoName = Trim$(sgRptPath) & Trim$("rptlogo.bmp")
                'Name slSaveRptLogoName As Trim$(sgRptPath) & "savelogo.bmp"
                'slVehicleLogo = Trim$(sgRptPath) & "G" & Trim$(tgvpf(ilVpfIndex).scplogo) & Trim$(".bmp")
                'Name Trim$(slVehicleLogo) As Trim$(sgRptPath) & Trim$("rptlogo.bmp")
        '    End If
        'End If
        Screen.MousePointer = vbDefault
        'dan 10-31-08 for multiple report jobs: don't call form until last job.
'         If Not ogReport Is Nothing Then
            'dan multiple reports, if last job, then go ahead and output.
            'Dan 08-17-09 Commented out below to stop 'mulitple reports'.  Problem with data being incorrect.

            If ogReport.iLastPrintJob = ilJobs Or ogReport.TreatAsLastReport Then
'                If Not gSetSelection(sgSelection & sgSelectionToAdd) Then
'                    Exit Sub
'                End If
                If Not Invoice!rbcType(INVGEN_Archive).Value Then            'archive invoices to PDF, Output method doesnt apply since all will be exported to pdf
                    If rbcOutput(0).Value Then
                        DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
                        igDestination = 0
                        Report.Show vbModal
                    ElseIf rbcOutput(1).Value Then
                        ilCopies = Val(edcCopies.Text)
                        'User is either printing or printing to pdf.  If there are invoices to be sent to agency, need to keep the report open
'                        If UBound(tgInvPDF_Info) > 0 Then
                        '3-1-17 always keep report open
                            ilRet = gOutputToPrinter(ilCopies, True)        'keep report open
'                        Else
'                            ilRet = gOutputToPrinter(ilCopies)
'                        End If
                    Else
                        slFileName = edcFileName.Text
                        'ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
                        ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
                    End If
                End If
                
                ilCounter = ilCounter
                vlArray = ogReport.Reports.Items
                For ilCounter = 0 To ogReport.Reports.Count - 1
                    ilCounter = ilCounter
                    Set myTempReport = vlArray(ilCounter)
                    myTempReport.RecordSelectionFormula = Trim$(sgSetSelectionForAll(ilCounter))
                    ilCounter = ilCounter
                    myTempReport.DiscardSavedData
                Next ilCounter
                
                '11-15-16 create archived file if finals or separate archive operation, or reprints
                If Invoice!rbcType(INVGEN_Archive).Value Or Invoice!ckcArchive.Value = vbChecked Or Invoice!rbcType(INVGEN_Reprint).Value Then    '12-18-16
                    mCreateArchivePdfs
                End If
                '11-18-16 create Export pdf if finals or reprint and OK to send PDF to agency
                '6-22-17 Feature to use PDF email must be set in Invoice Site to continue
                '6-22-17 Using Feature - for finals continue to create pdf:  may or may not send automatically based on addl feature
                '6-22-17 using feature for reprint - continue to create pdf only if user indicates ok to send.  if ok to send, may or may not send automatically (based on addl feature)
                'If Invoice!rbcType(INVGEN_Final).Value Or ((Invoice!rbcType(INVGEN_Reprint).Value) And (bgSendReprintPDF)) Then
                If ((Invoice!rbcType(INVGEN_Final).Value) Or ((Invoice!rbcType(INVGEN_Reprint).Value) And (bgSendPDF))) And ((Asc(tgSaf(0).sFeatures3) And INVEMAILINDEX) = INVEMAILINDEX) Then
                    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
                    If bgSendSelevtivePDF = True Then
                        ilFound = 0
                        For ilLoop = 0 To UBound(tmRPInfo) - 1
                            If tmRPInfo(ilLoop).iSelectiveEmail = 1 Then 'this inv is selected to email
                                ilFound = 1
                                Exit For
                            End If
                        Next ilLoop
                        If ilFound = 1 Then mCreateInvoiceEmailPdfs
                    Else
                        mCreateInvoiceEmailPdfs
                    End If
                End If

'                'delete prepasses
                gIvrClear
                gIMRClear
                Set myTempReport = Nothing

                On Error Resume Next
                If UBound(igPreviousDates) > 0 Then
                    mDeletePreviousTimeDate
                End If
                On Error GoTo 0
            Else    'this is here because ivr is cleared before report has really been created.
                mSavePreviousTimeDate
            End If
    Next ilJobs
    
    'delete prepasses
'    gIvrClear
 '   gIMRClear

    Set ogReport = Nothing
    If Not bgComingFromInvoice Then
        Erase igPreviousTimes
        Erase igPreviousDates
    End If

    '
    'Rename the vehicles logo back to rptlogo.bmp
    'Only rename back to rptlog.bmp if a valid vehicle option found
    'If (ilVpfIndex >= 0) And (tgvpf(ilVpfIndex).scplogo <> "   ") Then
        'Name Trim$(slVehicleLogo) As Trim$(sgRptPath) & "G" & Trim$(tgvpf(ilVpfIndex).scplogo) & Trim$(".bmp")
        'Name Trim$(slSaveRptLogoName) As Trim$(sgRptPath) & Trim$("rptlogo.bmp")
    'End If
    imGenShiftKey = 0

    'Screen.MousePointer = vbHourglass
    'Dan M 8-14-09 moved for cr2008--prepass deleted after all reports written as opposed to one at a time
'    gIvrClear
'    gIMRClear
    Screen.MousePointer = vbDefault
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    'frcWhen.Enabled = igWhen
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    Exit Sub
'cmcGenErr:
'    ilDDFSet = True
'    Resume Next
End Sub

Private Sub mSavePreviousTimeDate()
    Dim ilIndex As Integer
    ilIndex = UBound(igPreviousTimes, 2)
    ReDim Preserve igPreviousDates(1, ilIndex + 1) As Integer
    igPreviousDates(0, ilIndex) = igNowDate(0)
    igPreviousDates(1, ilIndex) = igNowDate(1)
    igPreviousTimes(0, ilIndex) = igNowTime(0)
    igPreviousTimes(1, ilIndex) = igNowTime(1)
End Sub

Private Sub mDeletePreviousTimeDate()
    Dim ilIndex As Integer
    For ilIndex = 0 To (UBound(igPreviousTimes) - 1)
        igNowDate(0) = igPreviousDates(0, ilIndex)
        igNowDate(1) = igPreviousDates(1, ilIndex)
        igNowTime(0) = igPreviousTimes(0, ilIndex)
        igNowTime(1) = igPreviousTimes(1, ilIndex)
        gIvrClear
        gIMRClear
    Next ilIndex
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
    If ((KeyAscii <> KEYBACKSPACE) And (KeyAscii <= 32)) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KeyDown) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
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
    mInit
    If imTerminate = -99 Then
        Exit Sub
    End If
    If imTerminate Then 'Used for print only
        'mTerminate
        cmcCancel_Click
        Exit Sub
    End If
    'RptSelIn.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    'Erase tgSellNameCode
    Erase tgRptSelInSalespersonCode
    Erase tgRptSelInAgencyCode
    Erase tgRptSelInAdvertiserCode
    Erase tgRptSelInNameCode
    Erase tgRptSelInBudgetCode
    'Erase tgMultiCntrCode
    'Erase tgManyCntCode
    Erase tgRptSelInDemoCode
    'Erase tgSOCode
    Erase igcodes
    PECloseEngine
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set RptSelIn = Nothing   'Remove data segment
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
            ckcAll.Value = vbUnchecked  'False
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
    'gCenterStdAlone RptSelIn
    'RptSelIn.Move -90, -90, 30, 30      'make form small and out of the way so its not seen
    RptSelIn.Move -330, -330, 30, 30      'make form small and out of the way so its not seen
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
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    gPopExportTypes cbcFileType     '10-20-01
    pbcSelA.Visible = False
    pbcSelB.Visible = False
    pbcSelC.Visible = False
    lbcRptType.Clear

    Screen.MousePointer = vbHourglass
    'If (igRptType = 0) Or (igRptType = 1) Or (igRptType = 2) Then
    If igOutputTo = 0 Then
        rbcOutput(0).Value = True
    Else
        rbcOutput(1).Value = True          'always print these automatically generated reports
    'rbcOutput(0).Value = True           'display -- for test purposes only
    End If

    ckcSelC3(0).Value = vbUnchecked 'False
    ckcSelC3(1).Value = vbUnchecked 'False
    ckcSelC3(2).Value = vbUnchecked 'False
    ckcSelC3(3).Value = vbUnchecked 'False
    If igZones = 0 Then
        ckcSelC3(0).Value = vbChecked
        ckcSelC3(1).Value = vbChecked
        ckcSelC3(2).Value = vbChecked
        ckcSelC3(3).Value = vbChecked
    ElseIf igZones = 1 Then
        ckcSelC3(0).Value = vbChecked
    ElseIf igZones = 2 Then
        ckcSelC3(1).Value = vbChecked
    ElseIf igZones = 3 Then
        ckcSelC3(2).Value = vbChecked
    Else
        ckcSelC3(3).Value = vbChecked
    End If
    cmcGen_Click
    imTerminate = True
    Exit Sub
    'End If
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSelIn
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
'*
'*          12-13-02 Add option to print NTR invoices
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer

    slCommand = sgCommandStr    'Command$
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    ''igStdAloneMode defined as "Debug" mode
    ''igStdAloneMode = True 'Switch from/to stand alone mode-No DDE
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True
    '    sgCallAppName = ""
    '    ilTestSystem = False  'True 'False
    '    'Change these following parameters to debug
    '    slGenDate = "11/4/98"
    '    slGenTime = "4:26:53P"
    '    slDisPlayPrint = "D"    'D= display, P= print
    '    slInvoiceType = "2"     '1 = invoice (no spots) as ordered:  invport1.rpt
    '                            '2 = 2 passes: affidavit (list spots as aired):  invaff.rpt  &
    '                            '              affidavit summary (total spots by length & station):  invaffsm.rpt
    '                            '3 = as aired, inv & affidavit combined :  invport3.rpt
    '                            '4 = NTR invoice (12-13-02)
    '                            '5 = REP, show ordered and aired spots & $
    '                            '6 = 01-18-07 combined air time & NTR
    '    'Mandatory parms are 1st and 2nd :
    '    '1st parm:  function coming from (Logs^Test (or Prod or NoHelp)
    '    '2nd parm:  user name
    '   11-16-16 add another parameters indicating month & year generated for pdf files and email info
    '    'parms: Logs^Test (or Prod)\ user name\jobcode(igrptcalltype)\(igrpttype)\gendate\gentime\DisplPrint\InvoiceType\InvMonthYr
    '    slCommand = "Logs^Test\Guide\0\0\" & slGenDate & "\" & slGenTime & "\" & slDisPlayPrint & "\" & slInvoiceType & "\" & slInvMonthYr

    '    imShowHelpmsg = False
    'Else
    '    igStdAloneMode = False
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            End
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
    'gInitStdAlone RptSelIn, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)                'Call Type
    ilRet = gParseItem(slCommand, 4, "\", slStr)                'Report type
    ilRet = gParseItem(slCommand, 5, "\", slStr)                'Gen Date
    gPackDate slStr, igNowDate(0), igNowDate(1)
    ilRet = gParseItem(slCommand, 6, "\", slStr)                'Gen Time
    gPackTime slStr, igNowTime(0), igNowTime(1)
    ilRet = gParseItem(slCommand, 7, "\", slStr)            'diplay or print
Debug.Print slCommand

    If Trim$(slStr) = "D" Then                           'display
        rbcOutput(0).Value = True
        rbcOutput(1).Value = False
        igOutputTo = 0
    Else
        rbcOutput(1).Value = True
        rbcOutput(0).Value = False
        igOutputTo = 1
    End If
    ilRet = gParseItem(slCommand, 8, "\", slStr)            'Invoice Type
    igInvoiceType = Val(slStr)
    
    ilRet = gParseItem(slCommand, 9, "\", sgInvMonthYear)             '11-16-16 get the month/year generated
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RptSelIn
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

'               Loop thru Email_Pdf array and create a seprate pdf for each unique agency and/or direct advertisr
'               mCreateInvoiceEmailPdfs
'               <input> tgInvPdf_Info global array
'
Private Sub mCreateInvoiceEmailPdfs()
    Dim slDate As String
    Dim slStr As String
    Dim slTime As String
    Dim ilLoopOnDiff As Integer
    Dim slPDFFileName As String
    Dim ilRet As Integer
    Dim blRet As Boolean
    Dim slPDFPathName As String
    Dim slClientName As String
    Dim tlSite As SITE
    Dim slFromAddress As String
    Dim tlSrchKey0 As INTKEY0
    Dim hlSite As Integer
    Dim tlMnf As MNF
    Dim hlMnf As Integer
    Dim slRecipient As String
    Dim slEmailErrMsg As String
    
    Dim hlInvPdfList As Integer         '7-6-17 file containing list of pdf  invoices created / sent
    Dim slInvPDFList As String          '7-6-17 name of file containing list of pdf invoices created/sent
    Dim ilLoopOnInv As Integer
    Dim slTempStr As String
    Dim slReprint As String * 1
    
    Dim ilLoop As Integer 'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
    Dim slSelectedInvoices As String
    Dim ilFound As Integer
    Dim blHasNTR As Boolean
    
'    slStorePath = ogReportn.PDFPathName
'    ogReport.PDFPathName = slNewPath
'    ogReport.export(....)
'    ogReport.PDFPathName = slStorePath

    If UBound(tgInvPDF_Info) > 0 Then
        hlSite = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hlSite, "", sgDBPath & "Site.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlSite)
            btrDestroy hlSite
            Exit Sub
        End If
        tlSrchKey0.iCode = 1
        ilRet = btrGetEqual(hlSite, tlSite, Len(tlSite), tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            slFromAddress = Trim$(tlSite.sEmailAcctName)
            ilRet = btrClose(hlSite)
            btrDestroy hlSite
        Else
            slFromAddress = "emailsender@counterpoint.net"
            ilRet = btrClose(hlSite)
            btrDestroy hlSite
            Exit Sub
        End If

        mFormatGenDateTime slDate, slTime
        slPDFPathName = ogReport.PDFPathName
        
        slReprint = ""
        If Invoice!rbcType(INVGEN_Reprint).Value = True Then                     'test for reprint and note it in filename
            slReprint = "R"
        End If
        slTempStr = Trim$(tgInvPDF_Info(0).sPDFExportPath)               'get the export path from the first entry, each one will be the same
        If right(slTempStr, 1) <> "\" Then
            slTempStr = slTempStr & "\"
        End If
        slInvPDFList = Trim$(slTempStr) & "Inv PDF Emails " & sgInvMonthYear & Trim$(slReprint) & " " & Trim$(slDate) & "_" & Trim$(slTime) & ".csv"
        ilRet = 0
        On Error GoTo mOpenInvPDFFileErr:
        hlInvPdfList = FreeFile
        Open slInvPDFList For Output As hlInvPdfList
        If ilRet <> 0 Then
            Close #hlInvPdfList
            Screen.MousePointer = vbDefault
            MsgBox "Open Error #" & str$(err.Number) & slInvPDFList, vbOKOnly, "Open Error"
            Exit Sub
        Else
            Print #hlInvPdfList, "Agency,Advertiser,Contract #,Invoice #,Month/Year,Date/Time Sent"
            If ilRet <> 0 Then
                Close #hlInvPdfList
                Screen.MousePointer = vbDefault
                MsgBox "Open Error #" & str$(err.Number) & slInvPDFList, vbOKOnly, "Writing Inv PDF Email List Error"
                Exit Sub
            End If
        End If

        ilRet = 0
        On Error GoTo mOpenInvPDFFileErr:
                   
        'open a file to list all the invoices and contracts sent
        'Generate all the agency PDFs before trying to email them out, in case email issues
        For ilLoopOnDiff = LBound(tgInvPDF_Info) To UBound(tgInvPDF_Info) - 1
            'new selection based on unique vehicle & contract
            If tgInvPDF_Info(ilLoopOnDiff).iAgfCode > 0 Then
               ' slStr = Trim$(sgSelection) & " and  ({CHF_Contract_Header.chfagfCode} = " & tgInvPDF_Info(ilLoopOnDiff).iAgfCode & ")"
                ogReport.AddToSelection = " and ({CHF_Contract_Header.chfagfCode} = " & tgInvPDF_Info(ilLoopOnDiff).iAgfCode & ")"
            Else
               ' slStr = Trim$(sgSelection) & " and  ({CHF_Contract_Header.chfadfCode} = " & tgInvPDF_Info(ilLoopOnDiff).iAdfCode & ")"
                ogReport.AddToSelection = " and ({CHF_Contract_Header.chfadfCode} = " & tgInvPDF_Info(ilLoopOnDiff).iAdfCode & ")"
            End If
            'determine the selective invoices to print
'            If Not gSetSelection(slStr) Then
'                Return
'            End If
            'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
            slSelectedInvoices = ""
            blHasNTR = True
            If bgSendSelevtivePDF = True Then
                'Determine which Invoices are selected for this Agy/Adv
                For ilLoop = 0 To UBound(tmRPInfo) - 1
                    If tmRPInfo(ilLoop).iSelectiveEmail = 1 Then 'this inv is selected to email
                        If tmRPInfo(ilLoop).iAgfCode = tgInvPDF_Info(ilLoopOnDiff).iAgfCode Or (tmRPInfo(ilLoop).iAgfCode = tgInvPDF_Info(ilLoopOnDiff).iAgfCode = 0 And tmRPInfo(ilLoop).iAdfCode = tgInvPDF_Info(ilLoopOnDiff).iAdfCode) Then
                            If slSelectedInvoices <> "" Then slSelectedInvoices = slSelectedInvoices & ","
                            slSelectedInvoices = slSelectedInvoices & tmRPInfo(ilLoop).lCntrNo
                        End If
                    End If
                Next ilLoop
                
                'Fix TTP 10826 / TTP 10813 -- Reset Selection, then add to selection to INCLUDE only Invoices being PDF Emailed
                If bgSendSelevtivePDF = True And bsSelectedEmailInvoices <> "" Then
                    'Fix v81 TTP 10826 - NTR reprint issue per Jason Email Thu 1/25/24 8:32 AM
                    If imCombineAirAndNTR Then
                        gSetSelection (sgSetSelectionForAll(0))
                    Else
                        gSetSelection (sgSetSelectionForAll(1))
                    End If
                End If
                ogReport.AddToSelection = ogReport.AddToSelection & " and ({CHF_Contract_Header.chfCntrNo} in [" & slSelectedInvoices & "])"
                blHasNTR = mDoesPayeeHaveNTR(IIF(tgInvPDF_Info(ilLoopOnDiff).iAgfCode = 0, tgInvPDF_Info(ilLoopOnDiff).iAdfCode, tgInvPDF_Info(ilLoopOnDiff).iAgfCode), slSelectedInvoices)
                Debug.Print ogReport.AddToSelection
                If slSelectedInvoices = "" Then
                    'Clear this Adv from Emailing as No Invoices for this one were selected
                    tgInvPDF_Info(ilLoopOnDiff).sPDFExportPath = ""
                    GoTo skipGenerating
                End If
            End If
            
            'set blKeepReportOpen to TRUE,
            'other Error message Object Variable or with block not set occurs
            slStr = Trim$(tgInvPDF_Info(ilLoopOnDiff).sPayeeName) & "_" & Trim$(slDate) & "_" & Trim$(slTime)
            'filename:  Vehicle, Cnt #, AdvtName,Revision #, CurrentDate Genned,Current Time Genned
            slPDFFileName = gStripCntrlChars(slStr)
            '6/25/18
            slPDFFileName = gFileNameFilter(slPDFFileName)
            
            ogReport.PDFPathName = Trim$(tgInvPDF_Info(ilLoopOnDiff).sPDFExportPath) & "\"
            ilRet = gExportCRW(slPDFFileName, 0, True, "", blHasNTR)
            ogReport.DiscardSavedData = True
            
'                '7-6-17 find all the matching agencies/advertisers and list their contract/inv #s
            For ilLoopOnInv = LBound(tgInvPDF_DetailInfo) To UBound(tgInvPDF_DetailInfo) - 1
                If tgInvPDF_DetailInfo(ilLoopOnInv).iAgfCode = tgInvPDF_Info(ilLoopOnDiff).iAgfCode Then
                    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
                    ilFound = 0
                    If bgSendSelevtivePDF = True Then
                        'look for Contract in tmRPInfo selected for email
                        For ilLoop = 0 To UBound(tmRPInfo) - 1
                            If tmRPInfo(ilLoop).lCntrNo = tgInvPDF_DetailInfo(ilLoopOnInv).lCntrNo Then
                                If tmRPInfo(ilLoop).iSelectiveEmail = 1 Then 'this inv is selected to email
                                    ilFound = 1
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                    Else
                        ilFound = 1
                    End If
                    If ilFound = 1 Then
                        slStr = Trim$(tgInvPDF_Info(ilLoopOnDiff).sPayeeName) & "," & Trim$(tgInvPDF_DetailInfo(ilLoopOnInv).sAdvtName) & "," & str$(tgInvPDF_DetailInfo(ilLoopOnInv).lCntrNo) & "," & str$(tgInvPDF_DetailInfo(ilLoopOnInv).lInvNo) & ",""" & Trim$(sgInvMonthYear) & """,""" & Trim$(slDate) & " " & Trim$(slTime) & """"
                        Print #hlInvPdfList, slStr
                    End If
                End If
            Next ilLoopOnInv
skipGenerating:
        Next ilLoopOnDiff
        Close #hlInvPdfList
        ogReport.PDFPathName = slPDFPathName
        PEClosePrintJob     'export was keeping report open (gexportcrw with true parameter); need to close , all done
        'ogReport.DiscardSavedData = True
        
        'use short client name vs long client name
        slClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hlMnf)
                btrDestroy hlMnf
            Else
                tlSrchKey0.iCode = tgSpf.iMnfClientAbbr
                ilRet = btrGetEqual(hlMnf, tlMnf, Len(tlMnf), tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slClientName = Trim$(tlMnf.sName)
                End If
            End If
        End If
        
        If Not igTestSystem Then            '6-22-17 must be in Production to send invoices; otherwise while testing a real client DB, the invoice will be sent
            'If (bgSendPDF) And (((Asc(tgSaf(0).sFeatures3) And INVSENDEMAILINDEX) = INVSENDEMAILINDEX)) Then         '6-22-17 continue to send PDF email out (its final and auto send, or reprint and OK to send)
            'Fix TTP 10826 / TTP 10813 - INVSENDEMAILINDEX was Automatic setting not Email Enabled setting
            If (bgSendPDF) And (((Asc(tgSaf(0).sFeatures3) And INVSENDEMAILINDEX) = INVSENDEMAILINDEX) Or ((tgSpfx.iInvExpFeature And INVEXP_SELECTIVEEMAIL) = INVEXP_SELECTIVEEMAIL And bgSendSelevtivePDF = True)) Then           '6-22-17 continue to send PDF email out (its final and auto send, or reprint and OK to send)
                slClientName = slClientName & " " & sgInvMonthYear & " Invoices"
                'all pdfs for agencies created, email them out
                If mPDFEmailStart() Then
                    For ilLoopOnDiff = LBound(tgInvPDF_Info) To UBound(tgInvPDF_Info) - 1
                        If Trim(tgInvPDF_Info(ilLoopOnDiff).sPDFExportPath) <> "" Then 'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist (ExportPath cleared if No pdf Generated for this Agency)
                            '10016 logging moved from above
                            For ilLoopOnInv = LBound(tgInvPDF_DetailInfo) To UBound(tgInvPDF_DetailInfo) - 1
                                If tgInvPDF_DetailInfo(ilLoopOnInv).iAgfCode = tgInvPDF_Info(ilLoopOnDiff).iAgfCode Then
                                    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
                                    If bgSendSelevtivePDF = True Then
                                        'Make sure the contract is Selected for Email
                                        ilFound = 0
                                        For ilLoop = 0 To UBound(tmRPInfo) - 1
                                            If tmRPInfo(ilLoop).lCntrNo = tgInvPDF_DetailInfo(ilLoopOnInv).lCntrNo Then
                                                If tmRPInfo(ilLoop).iSelectiveEmail = 1 Then
                                                    ilFound = 1
                                                    Exit For
                                                End If
                                            End If
                                        Next ilLoop
                                        If ilFound = 1 Then
                                            slStr = "Agency: " & Trim$(tgInvPDF_Info(ilLoopOnDiff).sPayeeName) & " Advertiser: " & Trim$(tgInvPDF_DetailInfo(ilLoopOnInv).sAdvtName) & " Contract#: " & str$(tgInvPDF_DetailInfo(ilLoopOnInv).lCntrNo) & " Invoice #: " & str$(tgInvPDF_DetailInfo(ilLoopOnInv).lInvNo) & " Month/Year: " & Trim$(sgInvMonthYear) & " Date/Time sent: " & Trim$(slDate) & " " & Trim$(slTime) & """"
                                            myPDFEmailLogger.WriteFacts slStr
                                        End If
                                    Else
                                        slStr = "Agency: " & Trim$(tgInvPDF_Info(ilLoopOnDiff).sPayeeName) & " Advertiser: " & Trim$(tgInvPDF_DetailInfo(ilLoopOnInv).sAdvtName) & " Contract#: " & str$(tgInvPDF_DetailInfo(ilLoopOnInv).lCntrNo) & " Invoice #: " & str$(tgInvPDF_DetailInfo(ilLoopOnInv).lInvNo) & " Month/Year: " & Trim$(sgInvMonthYear) & " Date/Time sent: " & Trim$(slDate) & " " & Trim$(slTime) & """"
                                        myPDFEmailLogger.WriteFacts slStr
                                    End If
                                End If
                            Next ilLoopOnInv
                            
                            slStr = Trim$(tgInvPDF_Info(ilLoopOnDiff).sPayeeName) & "_" & Trim$(slDate) & "_" & Trim$(slTime)
                            'filename:  Vehicle, Cnt #, AdvtName,Revision #, CurrentDate Genned,Current Time Genned
                            slPDFFileName = gStripCntrlChars(slStr)
                            '6/25/18
                            slPDFFileName = gFileNameFilter(slPDFFileName)
                            slRecipient = Trim$(tgInvPDF_Info(ilLoopOnDiff).sPDFEmailAddress)
                            If mSendPDFEmail(slRecipient, tgSpf.sGClient, slFromAddress, slClientName, "Attached are Invoices from " & Trim$(tgSpf.sGClient), Trim$(tgInvPDF_Info(ilLoopOnDiff).sPDFExportPath) & "\" & slPDFFileName & ".pdf", slEmailErrMsg, tgInvPDF_Info(ilLoopOnDiff).sPayeeName) Then
                                'all ok
                                ilRet = ilRet
                                '12-15-16
                                If Len(slEmailErrMsg) > 0 Then          'at least 1 was bad in syntax.  coming thru here at least one seemed to be good
                                    tgInvPDFEmailer_ErrMsg(UBound(tgInvPDFEmailer_ErrMsg)).sMsg = Trim$(tgInvPDF_Info(ilLoopOnDiff).sPayeeName) & ": " & ogEmailer.ErrorMessage & " " & slEmailErrMsg
                                    ReDim Preserve tgInvPDFEmailer_ErrMsg(LBound(tgInvPDFEmailer_ErrMsg) To UBound(tgInvPDFEmailer_ErrMsg) + 1) As InvPDFEmailer_ErrMsg
                                End If
                            Else
                                tgInvPDFEmailer_ErrMsg(UBound(tgInvPDFEmailer_ErrMsg)).sMsg = Trim$(tgInvPDF_Info(ilLoopOnDiff).sPayeeName) & ": " & ogEmailer.ErrorMessage & " " & slEmailErrMsg   'ttp #8943- no err msg shown
                                ReDim Preserve tgInvPDFEmailer_ErrMsg(LBound(tgInvPDFEmailer_ErrMsg) To UBound(tgInvPDFEmailer_ErrMsg) + 1) As InvPDFEmailer_ErrMsg
                            End If
                        End If
                    Next ilLoopOnDiff
                Else
                    'emailer not available
                End If
                gPDFEmailEnd
            End If
        End If                      'Not igTestSystem
    End If
    Exit Sub
mOpenInvPDFFileErr:
    ilRet = 1
    Resume Next
End Sub

Private Sub mCreateArchivePdfs()
    Dim slDate As String
    Dim slStr As String
    Dim slTime As String
    Dim slPDFFileName As String
    Dim ilRet As Integer
    Dim blRet As Boolean
    Dim slPDFPathName As String
    Dim slTempPath As String

'    slStorePath = ogReport.PDFPathName
'    ogReport.PDFPathName = slNewPath
'    ogReport.export(....)
'    ogReport.PDFPathName = slStorePath
    
    mFormatGenDateTime slDate, slTime
    
    slPDFPathName = ogReport.PDFPathName
    
    ogReport.DiscardSavedData = True
    
    'selection either has filter to exclude PDF and EDI invoices (final archiving) or includes all because its reprint (or archive operation)
    '            ilRet = ogReport.SetSelection(sgSelection)
    'determine if slash should be appended to end of string
    slTempPath = Trim$(sgExportPath)
    If right(slTempPath, 1) <> "\" Then
        slTempPath = slTempPath & "\"
    End If
    slTempPath = gStripCntrlChars(slTempPath)
    
    'set blKeepReportOpen to TRUE,
    'other Error message Object Variable or with block not set occurs
    'determine all or selective invoices for the filename description
    slStr = "Archive "
    '12-20-16:  3 different ways to get archive- 1) archive as separate operation (rbctype(5), using all or selective advt (rbcAdvt(0) 2) archive (ckcArchive) with finals, all or selective cnts, 3)reprint (rbctype(2) with all or select advt
    '            If Invoice!rbcCntr(INVCNTR_All).Value Then                '12-15-16
    If ((Invoice!rbcType(INVGEN_Archive).Value Or Invoice!rbcType(INVGEN_Reprint).Value) And (Invoice!rbcAdvt(0).Value)) Or (Invoice!ckcArchive.Value = vbChecked And Invoice!rbcCntr(INVCNTR_All).Value) Then
        slStr = slStr & "All "
    End If
        
    slPDFFileName = slPDFFileName & slStr & sgInvMonthYear & " Invoices " & Trim$(slDate) & "_" & Trim$(slTime)
    'filename:  "Archive " ,  CurrentDate Genned,Current Time Genned
    
    ogReport.PDFPathName = slTempPath
    ilRet = gExportCRW(slPDFFileName, 0, True)
    
    ogReport.DiscardSavedData = True
    ogReport.PDFPathName = slPDFPathName        'restore the original path
    Exit Sub
End Sub

'
'               convert generation date and time to string for Email PDF filename
'               Remove slash and replace with dashes in date, remove colon in time
'           mFormatGenDateTime()
'           <input>  global date variable:  igNowDate(0 to 1)
'                    global time variable:  lgNowTime
'           <output>  slDate
'                     slTime
Public Sub mFormatGenDateTime(slDate As String, slTime As String)
    Dim slTemp As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llDate As Long

    gUnpackDateLong igNowDate(0), igNowDate(1), llDate
     slStr = Format$(llDate, "m/d/yy")               'Now date as string
    'replace slash with dash in date
     slDate = ""
     For ilLoop = 1 To Len(slStr) Step 1
         slTemp = Mid$(slStr, ilLoop, 1)
         If slTemp <> "/" Then
             slDate = Trim$(slDate) & Trim$(slTemp)
         Else
             slDate = Trim$(slDate) & "-"
         End If
     Next ilLoop
     
     slStr = gFormatTimeLong(lgNowTime, "A", "1")
     'remove colons from time
     slTime = ""
     For ilLoop = 1 To Len(slStr) Step 1
         slTemp = Mid$(slStr, ilLoop, 1)
         If slTemp <> ":" Then
             slTime = Trim$(slTime) & Trim$(slTemp)
         End If
     Next ilLoop
     Do While Len(slTime) < 8
         slTime = "0" & slTime
     Loop
     Exit Sub
End Sub

'10016 moved here
'Dan M 11/7/16 email for darlene's invoicing 8245 10016 added optional test
Private Function mPDFEmailStart() As Boolean
    Dim blRet As Boolean
    Dim slLogName As String
    blRet = True
    Set myPDFEmailLogger = New CLogger
    myPDFEmailLogger.BlockUserName = True
    slLogName = myPDFEmailLogger.CreateLogName(sgDBPath & "Messages\" & "PDFEmail" & sgUserName & ".Txt")
    myPDFEmailLogger.LogPath = slLogName
    myPDFEmailLogger.WriteFacts ""
    myPDFEmailLogger.WriteFacts "*************************"
    myPDFEmailLogger.WriteFacts "Starting sending of emails", True
    '10016
    If bgPDFEmailTestMode Then
        myPDFEmailLogger.WriteWarning "In test mode!  Not sending emails", True
    End If
    Set ogEmailer = New CEmail
    If Len(ogEmailer.ErrorMessage) > 0 Then
        blRet = False
        gLogMsg "Email could not be set up in RptSelIn-mPDFEmailStart: " & ogEmailer.ErrorMessage, "TrafficErrors.Txt", False
    End If
    mPDFEmailStart = blRet
End Function

Private Sub gPDFEmailEnd()
    myPDFEmailLogger.WriteFacts "Ending sending emails", True
    myPDFEmailLogger.CleanFolder myPDFEmailLogger.MessageFolder
    Set ogEmailer = Nothing
End Sub

Private Function mSendPDFEmail(slRecipient As String, slFromName As String, slFromAddress As String, slSubject As String, slMessage As String, slAttachment As String, Optional slErrorMessage As String = "", Optional slAgencyName As String = "") As Boolean
    'note: ogemailer must already be created.  I: slRecipient may be single email or multiple separated by ;  slFromName,slMessage,slSubject,slAttachment may be blank  slFromAddress and slRecipient wil be tested for validity
    ' slAttachment should be complete path to file and will be tested that it actually exists.  O: true if sent ok.
    'errors written to trafficErrors.txt, but not shown as msg box
    '10016, added the payee for error messages
    Dim blRet As Boolean
    Dim slTo() As String
    Dim c As Integer
    Dim blAtLeastOne As Boolean
On Error GoTo ERRBOX
    blRet = True
    slErrorMessage = ""
    blAtLeastOne = False
    If InStr(slRecipient, ";") > 0 Then
        slTo = Split(slRecipient, ";")
    Else
        ReDim slTo(0)
        slTo(0) = slRecipient
    End If
    'now build and send
    If blRet Then
        If Not ogEmailer Is Nothing Then
            With ogEmailer
                'because previous error message wasn't erased
                .Clear False, True
                For c = 0 To UBound(slTo)
                    If Len(slTo(c)) > 0 Then
                        If .TestAddress(slTo(c)) Then
                            blAtLeastOne = True
                            .AddTOAddress slTo(c)
                            '10016
                            myPDFEmailLogger.WriteFacts " To: " & slTo(c)
                           ' myPDFEmailLogger.WriteFacts " Payee: " & slPayee & " To: " & slTo(c)
                            If Len(.ErrorMessage) > 0 Then
                                blRet = False
                                slErrorMessage = .ErrorMessage
                                myPDFEmailLogger.WriteError .ErrorMessage, False, False
                                GoTo CONTINUE
                            End If
                        Else
                           ' slErrorMessage = slBadAddress & slTo(c) & ","
                            slErrorMessage = slErrorMessage & slTo(c) & " " & .ErrorMessage & ","
                            myPDFEmailLogger.WriteWarning slTo(c) & " " & .ErrorMessage
'                            myPDFEmailLogger.WriteWarning " Payee: " & slPayee & ": " & slTo(c) & " " & .ErrorMessage
                        End If
                    Else
                        myPDFEmailLogger.WriteWarning "Missing email address."
'                        myPDFEmailLogger.WriteWarning " Payee " & slPayee & " missing email address."
                        slErrorMessage = slErrorMessage & " missing email address,"
                    End If
                Next c
                If Len(slErrorMessage) > 0 Then
                    slErrorMessage = "invalid emails: " & mLoseLastLetterIfComma(slErrorMessage)
                    'myPDFEmailLogger.WriteWarning slErrorMessage
                End If
                If blAtLeastOne Then
                    .FromAddress = slFromAddress
                    .FromName = slFromName
                    .Subject = slSubject
                    .Message = slMessage
                    .Attachment = slAttachment
                    '10016
                    If bgPDFEmailTestMode = False Then
                        If Not .Send() Then
                            blRet = False
                            'the send fails if there's an error message.  Let's not repeat it.
                            If Len(slErrorMessage) = 0 Then
                                slErrorMessage = .ErrorMessage
                                '10224 flip these statements
'                                myPDFEmailLogger.WriteError "Send failed.  See warnings above.", True, False
                                myPDFEmailLogger.WriteError "Send failed.  " & .ErrorMessage, True, False

                            Else
'                                myPDFEmailLogger.WriteError "Send failed.  " & .ErrorMessage, True, False
                                myPDFEmailLogger.WriteError "Send failed.  See warnings above.", True, False
                            End If
                        End If
                    Else
                        'mimics the real 'send'
                        If Len(slErrorMessage) = 0 Then
                            myPDFEmailLogger.WriteWarning "Send skipped", True
                        Else
                            myPDFEmailLogger.WriteError "Send failed.  See warnings above.", True, False
                        End If
                    End If
                Else
                    blRet = False
                    slErrorMessage = "no valid 'to' address. Did not send. " & slErrorMessage
                    myPDFEmailLogger.WriteError "no valid 'to' address. Did not send.", True, False
                End If
                '10016
                If blRet Then
                    myPDFEmailLogger.WriteFacts "Sent from " & slFromAddress & " Subject: " & slSubject & " Message: " & slMessage
                    myPDFEmailLogger.WriteFacts " -> Attachment: " & slAttachment
                    'Add Success to the "Error" list, so that a list of Successful sends will appear as a Confirmation
                    If slAgencyName <> "" Then
                        tgInvPDFEmailer_ErrMsg(UBound(tgInvPDFEmailer_ErrMsg)).sMsg = "Success sending: " & Trim$(slAgencyName)
                        ReDim Preserve tgInvPDFEmailer_ErrMsg(LBound(tgInvPDFEmailer_ErrMsg) To UBound(tgInvPDFEmailer_ErrMsg) + 1) As InvPDFEmailer_ErrMsg
                    End If
                End If
                myPDFEmailLogger.WriteFacts "-------------------------"
            End With
        Else
            blRet = False
            slErrorMessage = "ogEmailer does not exist"
        End If
    End If
CONTINUE:
    mSendPDFEmail = blRet
    Exit Function
ERRBOX:
    mSendPDFEmail = False
    slErrorMessage = err.Description
End Function

Private Function mLoseLastLetterIfComma(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String
    Dim llLastLetter As Long
    
    llLength = Len(slInput)
    llLastLetter = InStrRev(slInput, ",")
    If llLength > 0 And llLastLetter = llLength Then
        slNewString = Mid(slInput, 1, llLength - 1)
    Else
        slNewString = slInput
    End If
    mLoseLastLetterIfComma = slNewString
End Function

Private Function mDoesPayeeHaveNTR(ilPayeeCode As Integer, slSelectedContracts) As Boolean
    Dim slContracts() As String
    Dim ilLoop As Integer
    Dim ilLoop2 As Integer
    slContracts = Split(slSelectedContracts, ",")
    mDoesPayeeHaveNTR = False
    For ilLoop = 0 To UBound(slContracts)
        For ilLoop2 = 0 To UBound(tmInvAirNTRStatus)
            If tmInvAirNTRStatus(ilLoop2).lCntrNo = Val(slContracts(ilLoop)) Then
                If tmInvAirNTRStatus(ilLoop2).iPayeeCode = ilPayeeCode Then
                    mDoesPayeeHaveNTR = tmInvAirNTRStatus(ilLoop2).bHasNTR
                    Exit For
                End If
            End If
            If mDoesPayeeHaveNTR = True Then Exit For
        Next ilLoop2
    Next ilLoop
End Function
