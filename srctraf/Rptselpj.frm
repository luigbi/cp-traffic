VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelPJ 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   1545
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
      TabIndex        =   36
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6690
      TabIndex        =   82
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
      TabIndex        =   78
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
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   -90
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
               TabIndex        =   79
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
               TabIndex        =   81
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
               TabIndex        =   80
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
            Left            =   195
            MaxLength       =   10
            TabIndex        =   43
            Top             =   1440
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
            MaxLength       =   10
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
         Top             =   165
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   6
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   76
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   5
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   75
            Top             =   45
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   4
            Left            =   30
            MultiSelect     =   2  'Extended
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
            MultiSelect     =   2  'Extended
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
            MultiSelect     =   2  'Extended
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
            MultiSelect     =   2  'Extended
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
            MultiSelect     =   2  'Extended
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
         Width           =   1095
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
Attribute VB_Name = "RptSelPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselpj.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: rptselpj.Frm
'            Salesperson Projection reports (except Proj Scenario)
'            5/28/99 Allow 10 character input date (from 8) : m/d/yyyy
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
Dim smSelectedRptName As String 'Passed selected report name
'Vehicle link file- used to obtain start date
'Delivery file- used to obtain start date
'Vehicle conflict file- used to obtain start date
'Spot projection- used to obtain date status
'Library calendar file- used to obtain post log date status
'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
'Log
Dim imNoCodes As Integer
Dim imCodes() As Integer
Dim smLogStartDate As String
Dim smLogNoDays As String
Dim smLogUserCode As String
Dim smLogStartTime As String
Dim smLogEndTime As String
'Import contract report
Dim smChfConvName As String
Dim smChfConvDate As String
Dim smChfConvTime As String
'Spot week Dump
Dim imTerminate As Integer
Dim ilAASCodes()  As Integer
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
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilIndex As Integer
    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAll Then
        imAllClicked = True

        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(2).hwnd, LB_SELITEMRANGE, ilValue, llRg)
        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(6).hwnd, LB_SELITEMRANGE, ilValue, llRg)
    'Else
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
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
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
    'If (igRptCallType = GENERICBUTTON) Or ((igRptCallType = LOGSJOB) And ((igRptType = 0) Or (igRptType = 1) Or (igRptType = 2))) Then
        mTerminate False
    'Else
    '    mTerminate True
    'End If
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
    ilListIndex = lbcRptType.ListIndex

   igUsingCrystal = True

   ilNoJobs = 1

   ilNoJobs = 1
   ilStartJobNo = 1

    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportPj() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenPjct(ilListIndex, imGenShiftKey, smLogUserCode)
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

        If ilListIndex = PRJ_SALESPERSON Or ilListIndex = PRJ_VEHICLE Or ilListIndex = PRJ_OFFICE Or ilListIndex = PRJ_CATEGORY Then
            Screen.MousePointer = vbHourglass
            gCrProj
            Screen.MousePointer = vbDefault
        End If

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

    If ilListIndex = PRJ_SALESPERSON Or ilListIndex = PRJ_VEHICLE Or ilListIndex = PRJ_OFFICE Or ilListIndex = PRJ_CATEGORY Then
        Screen.MousePointer = vbHourglass
        gCRGrfClear
        Screen.MousePointer = vbDefault
    End If

    If igUsingCrystal Then          'close and re-open to clean up resources
        PECloseEngine
        ilRet = PEOpenEngine()      're-open since its closed in terminate routine again
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
End Sub
Private Sub edcSelCTo_Change()
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

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    RptSelPJ.Refresh
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
    'RptSelPJ.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgAirNameCodePJ
    Erase tgCSVNameCode
    Erase tgSellNameCode
    Erase tgRptSelPjSalespersonCode
    Erase tgRptSelPjAgencyCode
    Erase tgRptSelPjAdvertiserCode
    Erase tgRptSelPjNameCode
    Erase tgRptSelPjBudgetCode
    Erase tgRptSelPjDemoCode
    Erase imCodes
    PECloseEngine
    
    Set RptSelPJ = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub


Private Sub lbcRptType_Click()
    Dim ilListIndex As Integer
    rbcSelCInclude(2).Visible = False

    ilListIndex = lbcRptType.ListIndex
    ckcAll.Visible = False
    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3000
    lbcSelection(6).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(2).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(2).Visible = False
    lbcSelection(6).Visible = False
    edcSelCTo.Visible = False
    lacSelCTo.Visible = False
    edcSelCFrom.MaxLength = 0
    edcSelCFrom.Width = 1170
    lacSelCFrom.Width = 1500
    lacSelCFrom.Caption = "Report Date"
    lacSelCFrom.Visible = True
    edcSelCFrom.Visible = True
    plcSelC3.Visible = False
    plcSelC1.Visible = False
    plcSelC2.Visible = False
    rbcSelCSelect(0).Value = True
    rbcSelCInclude(0).Value = True
    If ilListIndex = PRJ_SALESPERSON Or ilListIndex = PRJ_VEHICLE Or ilListIndex = PRJ_OFFICE Or ilListIndex = PRJ_CATEGORY Then
        plcSelC2.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30
        plcSelC2.Visible = True
        'plcSelC2.Caption = "For"
        rbcSelCInclude(0).Left = 360
        rbcSelCInclude(0).Visible = True
        rbcSelCInclude(1).Visible = True
        rbcSelCInclude(0).Width = 990
        rbcSelCInclude(0).Caption = "Current"
        rbcSelCInclude(1).Left = 1380
        rbcSelCInclude(1).Width = 660
        rbcSelCInclude(1).Caption = "Past"
        rbcSelCInclude(0).Value = True          'default to current
        rbcSelCInclude(2).Visible = False
        plcSelC1.Move 120, plcSelC2.Top + plcSelC2.Height
        plcSelC1.Visible = True
        'plcSelC1.Caption = "Month"
        rbcSelCSelect(0).Left = 660
        rbcSelCSelect(0).Visible = True
        rbcSelCSelect(1).Visible = True
        rbcSelCSelect(0).Width = 1140
        rbcSelCSelect(0).Caption = "Corporate"
        rbcSelCSelect(1).Left = 1840
        rbcSelCSelect(1).Width = 1200
        rbcSelCSelect(1).Caption = "Standard"
        rbcSelCSelect(1).Value = True
        If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
            rbcSelCSelect(0).Enabled = False
        Else
            rbcSelCSelect(0).Value = True
        End If
        rbcSelCSelect(2).Visible = False
        plcSelC3.Move 120, plcSelC1.Top + plcSelC1.Height
        plcSelC3.Visible = True
        'plcSelC3.Caption = ""
        ckcSelC3(0).Left = 0
        ckcSelC3(0).Visible = True
        ckcSelC3(0).Caption = "Show Differences"
        ckcSelC3(0).Width = 1920
    End If
    If ilListIndex = PRJ_POTENTIAL Then
        edcSelCTo.Visible = False
        lacSelCTo.Visible = False
        lacSelCFrom.Visible = False
        edcSelCFrom.Visible = False
        edcSelCFrom.MaxLength = 0
        edcSelCFrom.Width = 1170
        lacSelCFrom.Width = 1500
        lacSelCFrom.Caption = "Report Date"
        plcSelC3.Visible = False
        plcSelC1.Visible = False
        plcSelC2.Visible = False
        rbcSelCSelect(0).Value = True
        rbcSelCInclude(0).Value = True

        plcSelC2.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30
        plcSelC2.Visible = True
        'plcSelC2.Caption = "For"
        rbcSelCInclude(0).Left = 360
        rbcSelCInclude(0).Visible = True
        rbcSelCInclude(1).Visible = True
        rbcSelCInclude(0).Width = 960
        rbcSelCInclude(0).Caption = "Current"
        rbcSelCInclude(1).Left = 1320
        rbcSelCInclude(1).Width = 720
        rbcSelCInclude(1).Caption = "Past"
        rbcSelCInclude(0).Value = True          'default to current
        rbcSelCInclude(2).Visible = False
        ckcAll.Visible = False
        rbcSelCSelect(1).Value = True           'force to "standard" only , altho not really used since all data is gathered for this report,
    End If                                          'not by a period
    If ilListIndex = PRJ_SALESPERSON Then
        ckcAll.Caption = "All Salespeople"
        lbcSelection(2).Visible = True
        ckcAll.Visible = True
    ElseIf ilListIndex = PRJ_VEHICLE Then
        ckcAll.Caption = "All Vehicles"
        lbcSelection(6).Visible = True
        ckcAll.Visible = True
    ElseIf ilListIndex = PRJ_OFFICE Or ilListIndex = PRJ_CATEGORY Then
        ckcAll.Value = vbChecked    'True                 'this option doesnt have office choice, set
                                            'to fall thru on general code in gCmcGenPjMore
    End If
    mSetCommands
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    Dim ilListIndex As Integer
    ReDim ilAASCodes(0 To 1) As Integer
    If Not imAllClicked Then
        ilListIndex = lbcRptType.ListIndex

              If (ilListIndex = 0 Or ilListIndex = 1) Then
                imSetAll = False
                ckcAll.Value = vbUnchecked  'False
                imSetAll = True
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
    RptSelPJ.Caption = smSelectedRptName & " Report"
    frcOption.Caption = smSelectedRptName & " Selection"

    gPopExportTypes cbcFileType '10-20-01

    imAllClicked = False
    imSetAll = True
    cbcSel.Move 120, 30
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
    rbcSelCSelect(0).Move 600, 0
    rbcSelCSelect(0).Caption = "Advt"
    rbcSelCSelect(1).Move 1290, 0
    rbcSelCSelect(1).Caption = "Agency"
    rbcSelCSelect(2).Move 2020, 0
    rbcSelCSelect(2).Caption = "Salesperson"
    plcSelC2.Move 120, 885
    'plcSelC2.Caption = "Include"
    rbcSelCInclude(0).Move 705, 0
    rbcSelCInclude(0).Caption = "All"
    rbcSelCInclude(1).Move 1245, 0
    rbcSelCInclude(2).Move 2655, 0
    plcSelC3.Move 120, 675
    'plcSelC3.Caption = "Zone"
    ckcSelC3(0).Move 465, -30
    ckcSelC3(0).Caption = "EST"
    ckcSelC3(1).Move 1065, -30
    ckcSelC3(1).Caption = "CST"
    ckcSelC3(2).Move 1710, -30
    ckcSelC3(2).Caption = "MST"
    ckcSelC3(3).Move 2355, -30
    ckcSelC3(3).Caption = "PST"

    plcSelC4.Move 120, 360
    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3000
    lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    pbcSelA.Move 90, 255, 4515, 3360
    pbcSelB.Move 90, 255, 4515, 3360
    pbcSelC.Move 90, 255, 4515, 3360
    gCenterStdAlone RptSelPJ
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:8/14/97       By:W. Bjerke      *
'*                                                     *
'*            Comments: Initialize report screen       *
'*            Modified: for rptselpjPJ only              *
'*******************************************************
Private Sub mInitReport()
    Dim ilRet As Integer
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    pbcSelA.Visible = False
    pbcSelB.Visible = False
    pbcSelC.Visible = False

    lbcRptType.Clear


    Screen.MousePointer = vbHourglass
    mSPersonPop lbcSelection(2)
    If imTerminate Then
        Exit Sub
    End If
    mSellConvVirtVehPop 6, False
    If imTerminate Then
        Exit Sub
    End If
    If tgSpf.sRUseCorpCal = "Y" Then       'if Using corp cal,retain in memory
        ilRet = gObtainCorpCal()
    End If
    lbcRptType.AddItem "Salesperson Projection", PRJ_SALESPERSON
    lbcRptType.AddItem "Vehicle Projection", PRJ_VEHICLE
    lbcRptType.AddItem "Sales Office Projection", PRJ_OFFICE
    lbcRptType.AddItem "Category Projection", PRJ_CATEGORY
    lbcRptType.AddItem "Office Projection by Potential", PRJ_POTENTIAL
    'lbcRptType.AddItem "Projection Scenarios", PRJ_SCENARIO
    'lbcRptType.ListIndex = 0                          'set default
    pbcOption.Visible = True
    pbcSelC.Visible = True
    frcOption.Enabled = True


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
    Dim ilLoop As Integer
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
    'gInitStdAlone RptSelPJ, slStr, ilTestSystem
    ''ilRet = gParseItem(slCommand, 3, "\", slStr)
    ''igRptCallType = Val(slStr)
    If igRptCallType <> GENERICBUTTON Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)
    Else
        igRptType = 0
    End If
    ' "Salesperson Projection", PRJ_SALESPERSON
    ' "Vehicle Projection", PRJ_VEHICLE
    ' "Sales Office Projection", PRJ_OFFICE
    ' "Category Projection", PRJ_CATEGORY
    ' "Office Projection by Potential", PRJ_POTENTIAL
    'See rptselPs "Projection Scenarios", PRJ_SCENARIO
    'If igStdAloneMode Then
    '    smSelectedRptName = "Office Projection by Potential"
    '    igRptCallType = PROPOSALPROJECTION 'NYFEED  'COLLECTIONSJOB 'SLSPCOMMSJOB   'LOGSJOB 'COPYJOB 'COLLECTIONSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
    '    igRptType = 3   '3 'Log     '0   'Summary '3 Program  '1  links
    '    slCommand = "x\x\x\x\2\2/6/95\7\12M\12M\1\26" '"" '"CONT0802.ASC\11/20/94\10:11:0 AM" '"x\x\x\x\2"
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)
        End If
    'End If
    If igRptCallType = LOGSJOB Then
        ilRet = gParseItem(slCommand, 5, "\", smLogUserCode)
        ilRet = gParseItem(slCommand, 6, "\", smLogStartDate)
        ilRet = gParseItem(slCommand, 7, "\", smLogNoDays)
        ilRet = gParseItem(slCommand, 8, "\", smLogStartTime)
        ilRet = gParseItem(slCommand, 9, "\", smLogEndTime)
        ilRet = gParseItem(slCommand, 10, "\", slStr)
        imNoCodes = Val(slStr)
        ReDim imCodes(0 To imNoCodes) As Integer
        If imNoCodes > 0 Then
            For ilLoop = 0 To imNoCodes - 1 Step 1
                ilRet = gParseItem(slCommand, 11 + ilLoop, "\", slStr)
                imCodes(ilLoop) = Val(slStr)
            Next ilLoop
        Else
            imCodes(0) = -1
        End If
        'If (igRptType = 0) Or (igRptType = 1) Or (igRptType = 2) Then
        '    igStdAloneMode = True 'Switch from/to stand alone mode-No DDE
        'End If
    End If
    If igRptCallType = CHFCONVMENU Then
        ilRet = gParseItem(slCommand, 5, "\", smChfConvName)
        ilRet = gParseItem(slCommand, 6, "\", smChfConvDate)
        ilRet = gParseItem(slCommand, 7, "\", smChfConvTime)
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
        'ilRet = gPopUserVehicleBox(rptselpj, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(rptselpj, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        ilRet = gPopUserVehicleBox(RptSelPJ, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(rptselpj, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelPJ, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)   'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelPJ
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVirtVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload rptselpj
    'Set rptselpj = Nothing   'Remove data segment
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/21/93       By:D. LeVine      *
'*            Modified:8/14/97       By:W. Bjerke      *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*            Modified: for RptSelPJ only              *
'*******************************************************
Private Sub mSetCommands()

    Dim ilEnable As Integer
    Dim ilLoop As Integer
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex

    If ilListIndex = PRJ_SALESPERSON Then
        If ckcAll.Value = vbChecked Then
            ilEnable = True
        Else
            For ilLoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                If lbcSelection(2).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        End If
        If (edcSelCFrom.Text = "" And RptSelPJ!rbcSelCInclude(1).Value) Then    'for past projections, date required
            ilEnable = False
        End If
    ElseIf ilListIndex = PRJ_VEHICLE Then
        If ckcAll.Value = vbChecked Then
            ilEnable = True
        Else
            For ilLoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                If lbcSelection(6).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        End If
        If (edcSelCFrom.Text = "" And RptSelPJ!rbcSelCInclude(1).Value) Then      'for past projection, date required
            ilEnable = False
        End If
    ElseIf ilListIndex = PRJ_OFFICE Or ilListIndex = PRJ_VEHICLE Or ilListIndex = PRJ_CATEGORY Or ilListIndex = PRJ_POTENTIAL Then
        ilEnable = True
        If (edcSelCFrom.Text = "" And RptSelPJ!rbcSelCInclude(1).Value) Then      'for past projection, date required
            ilEnable = False
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
    'ilRet = gPopSalespersonBox(rptselpj, 0, True, True, lbcSelection, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(RptSelPJ, 0, True, True, lbcSelection, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gPopSalespersonBox)", RptSelPJ
        On Error GoTo 0
    End If
    Exit Sub
mSPersonPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload rptselpj
    'Set rptselpj = Nothing   'Remove data segment
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
    Unload RptSelPJ
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
Private Sub rbcSelC4_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC4(Index).Value
    'End of coded added
Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
End Sub
Private Sub rbcSelCInclude_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelCInclude(Index).Value
    'End of coded added
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex

    If ilListIndex = PRJ_SALESPERSON Or ilListIndex = PRJ_OFFICE Or ilListIndex = PRJ_VEHICLE Or ilListIndex = PRJ_CATEGORY Or ilListIndex = PRJ_POTENTIAL Then
        If rbcSelCInclude(0).Value Then             'requested current slsp projections, which means no rollover date applies
            lacSelCFrom.Visible = False
            edcSelCFrom.Visible = False
            edcSelCFrom.Text = ""                   'test for 0 rollover dates to get curent stuff
            ckcSelC3(0).Enabled = False              'disallow diff for current report
        Else
            ckcSelC3(0).Enabled = True               'diff allow, requsting past (rollover stuff)
            lacSelCFrom.Visible = True               'turn report date back on
            edcSelCFrom.Visible = True
            ckcSelC3(0).Value = vbUnchecked 'False
        End If
    End If

    mSetCommands

End Sub
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    plcSelC3.Print ""
End Sub
Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print "For"
End Sub
Private Sub plcSelC1_Paint()
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    plcSelC1.Print "Month"
End Sub
Private Sub plcSelC4_Paint()
    plcSelC4.CurrentX = 0
    plcSelC4.CurrentY = 0
    plcSelC4.Print "Option"
End Sub
