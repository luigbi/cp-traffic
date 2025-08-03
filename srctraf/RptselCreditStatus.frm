VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelCreditStatus 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5685
   ClientLeft      =   585
   ClientTop       =   2775
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
   ScaleHeight     =   5685
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   6075
      Top             =   810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   37
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6690
      TabIndex        =   96
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
      TabIndex        =   86
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
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   -90
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcBudgetCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3435
      Sorted          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   4830
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.FileListBox lbcFileName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3615
      Pattern         =   "*.Dal"
      TabIndex        =   79
      Top             =   4815
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSMask.MaskEdBox mkcPhone 
      Height          =   210
      Left            =   3000
      TabIndex        =   78
      Tag             =   "The number and extension of the buyer."
      Top             =   4545
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
      Left            =   3795
      Sorted          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   4965
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcAgyAdvtCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3780
      Sorted          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   5205
      Visible         =   0   'False
      Width           =   945
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
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4245
      Width           =   90
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
      Height          =   4155
      Left            =   120
      TabIndex        =   14
      Top             =   1470
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
         Height          =   3930
         Left            =   840
         ScaleHeight     =   3930
         ScaleWidth      =   4530
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   4530
         Begin VB.TextBox edcText2 
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
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   98
            Top             =   3600
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox edcText1 
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
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   97
            Top             =   3600
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.CheckBox ckcOption 
            Caption         =   "ckcOption"
            Height          =   210
            Left            =   210
            TabIndex        =   150
            Top             =   3660
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.ComboBox cbcSet3 
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
            Left            =   3270
            TabIndex        =   148
            Top             =   2415
            Visible         =   0   'False
            Width           =   1110
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
            Left            =   3165
            TabIndex        =   147
            Top             =   2490
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox plcSelC12 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   3435
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC12 
               Caption         =   "Extra"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2790
               TabIndex        =   142
               Top             =   0
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.OptionButton rbcSelC12 
               Caption         =   "Billing"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   141
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.OptionButton rbcSelC12 
               Caption         =   "Revenue"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1770
               TabIndex        =   140
               Top             =   0
               Visible         =   0   'False
               Width           =   1020
            End
         End
         Begin VB.PictureBox plcSelC11 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            ScaleHeight     =   240
            ScaleWidth      =   4500
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   3300
            Visible         =   0   'False
            Width           =   4500
            Begin VB.CheckBox ckcSelC11 
               Caption         =   "Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   138
               Top             =   -30
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CheckBox ckcSelC11 
               Caption         =   "Non-Polit"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1410
               TabIndex        =   137
               Top             =   -30
               Visible         =   0   'False
               Width           =   1290
            End
         End
         Begin VB.PictureBox plcSelC10 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            ScaleHeight     =   240
            ScaleWidth      =   4260
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   3120
            Visible         =   0   'False
            Width           =   4260
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2415
               TabIndex        =   130
               Top             =   -15
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Feed Spots"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1410
               TabIndex        =   129
               Top             =   -30
               Visible         =   0   'False
               Width           =   1290
            End
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Contract Spots"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   128
               Top             =   -30
               Visible         =   0   'False
               Width           =   1605
            End
         End
         Begin VB.PictureBox plcSelC9 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   15
            ScaleHeight     =   270
            ScaleWidth      =   4380
            TabIndex        =   125
            Top             =   2865
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcTrans 
               Caption         =   "Show transaction comments"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   495
               TabIndex        =   126
               Top             =   45
               Visible         =   0   'False
               Width           =   2865
            End
         End
         Begin VB.TextBox edcCheck 
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
            Left            =   1545
            MaxLength       =   10
            TabIndex        =   117
            Top             =   2640
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.PictureBox plcSelC8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   2325
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC8 
               Caption         =   "History"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2040
               TabIndex        =   119
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton rbcSelC8 
               Caption         =   "Receivables"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   118
               Top             =   0
               Width           =   1305
            End
            Begin VB.OptionButton rbcSelC8 
               Caption         =   "Both"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3240
               TabIndex        =   115
               Top             =   0
               Width           =   1005
            End
         End
         Begin VB.PictureBox plcSelC7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   255
            ScaleHeight     =   270
            ScaleWidth      =   4380
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   2190
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC7 
               Caption         =   "Include Fill Spots"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   1080
               TabIndex        =   113
               Top             =   0
               Visible         =   0   'False
               Width           =   2265
            End
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
            Left            =   1155
            TabIndex        =   107
            Text            =   "Major Set #"
            Top             =   2520
            Visible         =   0   'False
            Width           =   1455
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
            Left            =   2505
            TabIndex        =   108
            Top             =   2445
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox plcSelC6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1200
            ScaleHeight     =   255
            ScaleWidth      =   3675
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   3675
            Begin VB.CheckBox ckcSelC6Add 
               Caption         =   "Check1"
               Height          =   210
               Index           =   0
               Left            =   345
               TabIndex        =   135
               Top             =   15
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.OptionButton rbcSelC6 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2310
               TabIndex        =   106
               Top             =   0
               Width           =   1005
            End
            Begin VB.OptionButton rbcSelC6 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   105
               Top             =   -15
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC6 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1650
               TabIndex        =   104
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.PictureBox plcSelC5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   15
            ScaleHeight     =   270
            ScaleWidth      =   4380
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Include Fill Spots"
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   0
               Left            =   -60
               TabIndex        =   101
               Top             =   0
               Visible         =   0   'False
               Width           =   2265
            End
         End
         Begin VB.PictureBox plcSelC3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   30
            ScaleHeight     =   525
            ScaleWidth      =   4380
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1230
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2880
               TabIndex        =   70
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
               TabIndex        =   71
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   50
            Top             =   360
            Visible         =   0   'False
            Width           =   960
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
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   46
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
            Left            =   480
            MaxLength       =   10
            TabIndex        =   48
            Top             =   360
            Width           =   960
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   990
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "HC"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   3480
               TabIndex        =   132
               Top             =   0
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "incl3"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2760
               TabIndex        =   93
               Top             =   0
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "New/Changed"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   1920
               TabIndex        =   59
               Top             =   0
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "All"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   57
               Top             =   0
               Width           =   510
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "New/Changed"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1200
               TabIndex        =   58
               Top             =   0
               Width           =   1680
            End
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4410
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   4410
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Sales Origin"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   9
               Left            =   120
               TabIndex        =   133
               Top             =   0
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "NTR"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   8
               Left            =   480
               TabIndex        =   111
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "NTR"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   7
               Left            =   3720
               TabIndex        =   110
               Top             =   0
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Office/Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   6
               Left            =   3360
               TabIndex        =   102
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   3240
               TabIndex        =   95
               Top             =   0
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   2880
               TabIndex        =   94
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Advt"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2520
               TabIndex        =   55
               Top             =   0
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Advt"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   3735
               TabIndex        =   52
               Top             =   -15
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Agency"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   885
               TabIndex        =   53
               Top             =   0
               Width           =   1170
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Salesperson"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   1560
               TabIndex        =   54
               Top             =   0
               Width           =   1935
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
            Left            =   480
            MaxLength       =   10
            TabIndex        =   44
            Top             =   0
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
            Height          =   315
            Left            =   540
            TabIndex        =   60
            Top             =   120
            Visible         =   0   'False
            Width           =   4305
         End
         Begin VB.PictureBox plcSelC4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1650
               TabIndex        =   74
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   73
               Top             =   0
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2400
               TabIndex        =   75
               Top             =   0
               Width           =   1005
            End
         End
         Begin VB.Label lacCaption 
            Caption         =   "Sort by"
            Height          =   300
            Left            =   3330
            TabIndex        =   149
            Top             =   2715
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lacSort 
            Caption         =   "Sort by"
            Height          =   300
            Left            =   2610
            TabIndex        =   134
            Top             =   2700
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lacCheck 
            Appearance      =   0  'Flat
            Caption         =   "Check #"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   165
            TabIndex        =   116
            Top             =   2700
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Active Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   61
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lacSelCTo1 
            Appearance      =   0  'Flat
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1440
            TabIndex        =   49
            Top             =   480
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# of Days"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2400
            TabIndex        =   45
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
            TabIndex        =   47
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
         Height          =   1200
         Left            =   120
         ScaleHeight     =   1200
         ScaleWidth      =   4455
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox ckcDelinquentOnly 
            Caption         =   "Delinquent Only"
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
            Height          =   255
            Left            =   2685
            TabIndex        =   23
            Top             =   345
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.CheckBox ckcInclCommentsA 
            Caption         =   "FindIT"
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
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   840
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CheckBox ckcADate 
            Caption         =   "FindIT"
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
            Height          =   255
            Left            =   150
            TabIndex        =   16
            Top             =   30
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
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   18
            Top             =   0
            Width           =   1170
         End
         Begin VB.PictureBox plcSel1 
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
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   4275
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   330
            Visible         =   0   'False
            Width           =   4275
            Begin VB.CheckBox ckcSel1 
               Caption         =   "Pending"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1635
               TabIndex        =   21
               Top             =   0
               Width           =   1000
            End
            Begin VB.CheckBox ckcSel1 
               Caption         =   "Fed Only Events"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   20
               Top             =   15
               Width           =   1620
            End
         End
         Begin VB.PictureBox plcSel2 
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
            Height          =   285
            Left            =   90
            ScaleHeight     =   285
            ScaleWidth      =   3285
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   570
            Visible         =   0   'False
            Width           =   3285
            Begin VB.CheckBox ckcSel2 
               Caption         =   "M-F"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   15
               TabIndex        =   24
               Top             =   0
               Width           =   600
            End
            Begin VB.CheckBox ckcSel2 
               Caption         =   "Sat"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   570
               TabIndex        =   25
               Top             =   0
               Width           =   570
            End
            Begin VB.CheckBox ckcSel2 
               Caption         =   "Sun"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   1140
               TabIndex        =   26
               Top             =   0
               Width           =   630
            End
         End
         Begin VB.Label lacInclude 
            Appearance      =   0  'Flat
            Caption         =   "Include"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1095
            TabIndex        =   143
            Top             =   825
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lacFromA 
            Appearance      =   0  'Flat
            Caption         =   "Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   510
            TabIndex        =   17
            Top             =   30
            Width           =   660
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
         Height          =   1785
         Left            =   135
         ScaleHeight     =   1785
         ScaleWidth      =   4425
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   4425
         Begin VB.PictureBox plcType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   2385
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   15
            Width           =   2385
            Begin VB.OptionButton rbcType 
               Caption         =   "Detail"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   146
               Top             =   0
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton rbcType 
               Caption         =   "Summary"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   990
               TabIndex        =   145
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.PictureBox plcRepInv 
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
            Height          =   285
            Left            =   120
            ScaleHeight     =   285
            ScaleWidth      =   3585
            TabIndex        =   121
            Top             =   1440
            Visible         =   0   'False
            Width           =   3585
            Begin VB.CheckBox ckcRepInv 
               Caption         =   "Internal"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2160
               TabIndex        =   123
               Top             =   0
               Value           =   1  'Checked
               Width           =   1065
            End
            Begin VB.CheckBox ckcRepInv 
               Caption         =   "External"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1080
               TabIndex        =   122
               Top             =   0
               Value           =   1  'Checked
               Width           =   1000
            End
         End
         Begin VB.TextBox edcAsOfDate 
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
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   120
            Top             =   1080
            Visible         =   0   'False
            Width           =   1170
         End
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
            Height          =   315
            Left            =   750
            TabIndex        =   35
            Top             =   675
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
            Height          =   315
            Left            =   750
            TabIndex        =   33
            Top             =   270
            Width           =   1080
         End
         Begin VB.Label lacAsOfDate 
            Caption         =   "Entered on or After"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label lacTo 
            Appearance      =   0  'Flat
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   34
            Top             =   750
            Width           =   570
         End
         Begin VB.Label lacFrom 
            Appearance      =   0  'Flat
            Caption         =   "From"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   32
            Top             =   330
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
         Height          =   3825
         Left            =   4605
         ScaleHeight     =   3825
         ScaleWidth      =   4455
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox ckcAllGroups 
            Caption         =   "All Vehicle Groups"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1155
            TabIndex        =   109
            Top             =   0
            Visible         =   0   'False
            Width           =   3060
         End
         Begin VB.CheckBox ckcAllRC 
            Caption         =   "All Rate Cards"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Visible         =   0   'False
            Width           =   3945
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   12
            ItemData        =   "RptselCreditStatus.frx":0000
            Left            =   75
            List            =   "RptselCreditStatus.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   92
            Top             =   255
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   11
            Left            =   60
            MultiSelect     =   2  'Extended
            TabIndex        =   91
            Top             =   285
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   10
            Left            =   90
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   90
            Top             =   330
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   9
            Left            =   75
            MultiSelect     =   2  'Extended
            TabIndex        =   89
            Top             =   270
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   8
            Left            =   75
            MultiSelect     =   2  'Extended
            TabIndex        =   88
            Top             =   285
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   7
            Left            =   75
            MultiSelect     =   2  'Extended
            TabIndex        =   87
            Top             =   270
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   6
            Left            =   45
            MultiSelect     =   2  'Extended
            TabIndex        =   81
            Top             =   240
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   5
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   80
            Top             =   270
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   4
            Left            =   45
            MultiSelect     =   2  'Extended
            TabIndex        =   42
            Top             =   270
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   3
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   41
            Top             =   255
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   2
            Left            =   60
            MultiSelect     =   2  'Extended
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   4350
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   1
            Left            =   60
            MultiSelect     =   2  'Extended
            TabIndex        =   29
            Top             =   315
            Visible         =   0   'False
            Width           =   4395
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   0
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   28
            Top             =   240
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   30
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
            TabIndex        =   83
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
            TabIndex        =   84
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
      TabIndex        =   38
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   36
      Top             =   135
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
Attribute VB_Name = "RptSelCreditStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptSelCreditStatus.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelCreditStatus.Frm
'
'  Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllGroup As Integer 'True=Set list box; False= don't change list box
Dim imAllGroupClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
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

'Copy inventory to set printables flag
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length

'Spot week Dump
Dim imFirstActivate As Integer
Dim imTerminate As Integer
Dim ilAASCodes()  As Integer
Dim smPlcSel1P As String
Dim smPlcSel2P As String
Dim smPlcSelC1P As String
Dim smPlcSelC2P As String
Dim smPlcSelC3P As String
Dim smPlcSelC4P As String
Dim smPlcSelC6P As String
Dim smPlcSelC8P As String
Dim smPlcSelC9P As String
Dim smPlcSelC10P As String
Dim smPlcSelC11P As String
Dim smPlcSelC12P As String
Dim imAutoReport As Integer         'run report without asking questions (called from another module:  ShoCredit)

Const ChooseRadio = 0
Const ChooseCheck = 1

'*******************************************************
'*                                                     *
'*      Procedure Name:mSellConvStdPkgPop             *
'*                                                     *
'*             Created:11/21/02      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box with conventional, selling *
'*                      and std package vehicles for
'*                      rate card report
'*******************************************************
Private Sub mSellConvStdPkgPop(ilIndex As Integer, ilUselbcVehicle As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHSTDPKG + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHSTDPKG + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvStdPkgPopErr
        gCPErrorMsg ilRet, "mSellConvStdPkgPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mSellConvStdPkgPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
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

Private Sub cbcSet1_Click()
Dim ilSetIndex As Integer
Dim ilRet As Integer
Dim ilLoop As Integer
Dim ilListIndex As Integer

ilListIndex = lbcRptType.ListIndex

    If (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_STATEMENT) Then      '
        'Statements only
        'defaulted to Use Assigned, do nothing
    ElseIf (igRptCallType = USERLIST) Then
        'user report, do nothing
    Else
        ilLoop = cbcSet1.ListIndex
        ilSetIndex = gFindVehGroupInx(ilLoop, tgVehicleSets1())
        If ilSetIndex > 0 Then
            'smVehGp5CodeTag = ""       '8-25-03 move to after report type test; common arrays destroyed here (tgsocode)
            'ilRet = gPopMnfPlusFieldsBox(RptSelCreditStatus, lbcSelection(7), tgSOCode(), smVehGp5CodeTag, "H" & Trim$(Str$(ilSetIndex)))
    
            'If (igRptCallType <> COLLECTIONSJOB And ilListIndex <> COLL_CASH) And (igRptCallType <> COLLECTIONSJOB And ilListIndex <> COLL_CASHSUM) Then
            If (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_CASH) Or (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_CASHSUM) Or (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_POAPPLY) Then
            Else
                smVehGp5CodeTag = ""
                ilRet = gPopMnfPlusFieldsBox(RptSelCreditStatus, lbcSelection(7), tgSOCode(), smVehGp5CodeTag, "H" & Trim$(str$(ilSetIndex)))
    
                ckcAllGroups.Move 15, 1840
                lbcSelection(7).Move 15, ckcAllGroups.Height + ckcAllGroups.Top, 4380, 1600
                lbcSelection(2).Height = 1600       'slsp
                lbcSelection(6).Height = 1600       'vehicles
                lbcSelection(9).Height = 1600       'sales source
                lbcSelection(5).Height = 1600       'advt
                If ilSetIndex = 1 Then              'participants vehicle sets
                    lbcSelection(7).Visible = True
                    ckcAllGroups.Caption = "All Participants"
                ElseIf ilSetIndex = 2 Then          'subtotals vehicle sets
                    lbcSelection(7).Visible = True
                    ckcAllGroups.Caption = "All Sub-totals"
                ElseIf ilSetIndex = 3 Then          'market vehicle sets
                    lbcSelection(7).Visible = True
                    ckcAllGroups.Caption = "All Markets"
                ElseIf ilSetIndex = 4 Then          'format vehicle sets
                    lbcSelection(7).Visible = True
                    ckcAllGroups.Caption = "All Formats"
                ElseIf ilSetIndex = 5 Then          'research vehicle sets
                    lbcSelection(7).Visible = True
                    ckcAllGroups.Caption = "All Research"
                ElseIf ilSetIndex = 6 Then          'Sub-company vehicle sets
                    lbcSelection(7).Visible = True
                    ckcAllGroups.Caption = "All Sub-companies"
                End If
                ckcAllGroups.Visible = True
                ckcAllGroups.Value = vbChecked  '9-12-02False
    
                If (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_AGEVEHICLE) Then
                    ckcSelC3(0).Enabled = True
                End If
            End If
        Else
            lbcSelection(7).Visible = False
            ckcAllGroups.Value = vbUnchecked    '9-12-02 False
            ckcAllGroups.Visible = False
            lbcSelection(2).Move 15, ckcAll.Height + 30, 4380, 3330     'slsp
            lbcSelection(6).Move 15, ckcAll.Height + 30, 4380, 3330     'vehicles
            lbcSelection(9).Move 15, ckcAll.Height + 30, 4380, 3330     'sales source
            lbcSelection(5).Move 15, ckcAll.Height + 30, 4380, 3330     'advertiser
            If (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_AGEVEHICLE) Then
                ckcSelC3(0).Enabled = False
                ckcSelC3(0).Value = vbUnchecked
            End If
        End If
    End If
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
        If igRptCallType = RATECARDSJOB Then
            If lbcSelection(0).ListCount > 0 Then       'select all vehicles
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        ElseIf igRptCallType = BUDGETSJOB Then
            If lbcSelection(0).ListCount > 0 Then       'select all vehicles
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
            If lbcSelection(1).ListCount > 0 Then       'select all offices
                llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        ElseIf igRptCallType = PROGRAMMINGJOB Then
            If igRptType = 3 Then           'reports (vs links)
                If lbcSelection(0).ListCount > 0 Then
                    llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If                          'lbselection(0).listcount > 0
            Else                                'links
                'If (ilListIndex = 0) Or (ilListIndex = 1) Or (ilListIndex = 2) Then
                    If lbcSelection(0).ListCount > 0 Then
                        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                    If lbcSelection(1).ListCount > 0 Then
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                'ElseIf (ilListIndex = 3) Or (ilListIndex = 4) Then
                    If lbcSelection(2).ListCount > 0 Then
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                    If lbcSelection(3).ListCount > 0 Then
                        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(3).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                'End If
            End If                      'program report (vs links)
        ElseIf igRptCallType = CHFCONVMENU Then
            If lbcSelection(0).ListCount > 0 Then
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        ElseIf igRptCallType = COPYJOB Then
            If ilIndex = COPY_ROT Then
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                lbcSelection(5).Visible = False
            ElseIf ilIndex = COPY_INVBYSTARTDATE Or ilIndex = COPY_INVPRODUCER Then
                llRg = CLng(lbcSelection(10).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(10).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = COPY_SPLITROT Then         '1-30-09
                'advertiser list
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            Else
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
            If ilIndex = 1 Then             'copy status by advt
                plcSelC10.Visible = True
                lbcSelection(5).Visible = False
            End If
        ElseIf igRptCallType = POSTLOGSJOB Then
            llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        ElseIf igRptCallType = BULKCOPY Then
            llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(3).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        ElseIf igRptCallType = CMMLCHG Then
            llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        ElseIf igRptCallType = INVOICESJOB Then
            If ilIndex = INV_REGISTER Then
                 If rbcSelCSelect(0).Value = True Then       'invoice sort
                 ElseIf rbcSelCSelect(1).Value = True Then    'advt
                     llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                     llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                 ElseIf rbcSelCSelect(2).Value = True Then       'agy
                     llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                     llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                 ElseIf rbcSelCSelect(3).Value = True Then       'slsp
                     llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                     llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                 ElseIf rbcSelCSelect(4).Value = True Or rbcSelCSelect(5).Value = True Or rbcSelCSelect(6).Value = True Then 'bill vehicle, air vehicle, office/vehicle
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                     llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                 ElseIf rbcSelCSelect(7).Value = True Then           'ntr
                     llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                     llRet = SendMessageByNum(lbcSelection(8).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                 ElseIf rbcSelCSelect(8).Value = True Then       'sales source
                    llRg = CLng(lbcSelection(9).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(9).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                Else                                            'sales origin
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
            Else            'not invoice register
                llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(3).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(8).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(9).ListCount - 1) * &H10000 Or 0   'Sales source
                llRet = SendMessageByNum(lbcSelection(9).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                '10-8-03
                llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0   'Sales source
                llRet = SendMessageByNum(lbcSelection(7).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If

        ElseIf igRptCallType = COLLECTIONSJOB Then
            'If rbcRptType(3).Value Then 'Statements
            If ilIndex = COLL_AGEPAYEE Or ilIndex = COLL_AGESLSP Or ilIndex = COLL_AGEVEHICLE Or ilIndex = COLL_AGEOWNER Or ilIndex = COLL_AGESS Or ilIndex = COLL_AGEPRODUCER Then '2-10-00
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(3).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = COLL_DISTRIBUTE Then
                llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(3).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = COLL_CASH Then
                llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = 5 Then
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = COLL_POAPPLY Then
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            'TTP 9893
            ElseIf ilIndex = COLL_CREDITSTATUS Then
                llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            Else
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        ElseIf (igRptCallType = AGENCIESLIST And ilIndex = 2) Then
            llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        ElseIf (igRptCallType = USERLIST And ilIndex = USER_ACTIVITY) Then
            llRg = CLng(lbcSelection(10).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(10).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        Else
            'If (rbcRptType(1).Value) Or (igRptCallType = EVENTNAMESLIST) Then
            If (ilIndex = 1) Or (igRptCallType = EVENTNAMESLIST) Or (ilIndex = 0 Or ilIndex = 5 And igRptCallType = VEHICLESLIST) Then
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
'                If igRptCallType = SALESPEOPLELIST Then
'                    llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
'                    llRet = SendMessageByNum(lbcSelection(1).hwnd, LB_SELITEMRANGE, ilValue, llRg)
'                End If
            Else    'Salesperson
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(1).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        End If
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllGroups_Click()

Dim Value As Integer
Value = False
If ckcAllGroups.Value = vbChecked Then
    Value = True
End If
Dim ilIndex As Integer
Dim ilValue As Integer
Dim llRg As Long
Dim llRet As Long

    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAllGroup Then
        imAllGroupClicked = True
        If igRptCallType = COLLECTIONSJOB Then
            'TTP 9893
            If ilIndex = COLL_CREDITSTATUS Then
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            Else
                If lbcSelection(7).ListCount > 0 Then       'select all vehicles
                    llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(7).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
                If ilIndex = COLL_AGESS Or ilIndex = COLL_AGEOWNER Or ilIndex = COLL_AGEPRODUCER Then
                    llRg = CLng(lbcSelection(9).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(9).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
                If ilIndex = COLL_CASH Or ilIndex = COLL_SALESCOMM_COLL Then
                    llRg = CLng(lbcSelection(4).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(4).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
            End If
        ElseIf igRptCallType = INVOICESJOB Then
            If ilIndex = INV_REGISTER Then
                 If rbcSelCSelect(2).Value = True Or rbcSelCSelect(3).Value = True Then       'agy or slsp
                     llRg = CLng(lbcSelection(9).ListCount - 1) * &H10000 Or 0
                     llRet = SendMessageByNum(lbcSelection(9).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                 ElseIf rbcSelCSelect(1).Value = True Or rbcSelCSelect(5).Value = True Or rbcSelCSelect(8).Value = True Then       'agy, air vehicle or s/s
                     llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                     llRet = SendMessageByNum(lbcSelection(7).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
            Else
                If lbcSelection(9).ListCount > 0 Then       'select all vehicles
                    llRg = CLng(lbcSelection(9).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(9).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
            End If
        ElseIf igRptCallType = COPYJOB Then
            llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        ElseIf igRptCallType = PROGRAMMINGJOB Then
            llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0               'all avail names
            llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
        End If
    End If
    imAllGroupClicked = False
    mSetCommands
End Sub

Private Sub ckcAllRC_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllRC.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
Dim ilIndex As Integer
Dim ilValue As Integer
Dim llRg As Long
Dim llRet As Long
    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If igRptCallType = RATECARDSJOB Then
            If lbcSelection(11).ListCount > 0 Then       'select all vehicles
                llRg = CLng(lbcSelection(11).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(11).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        End If
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcDelinquentOnly_Click()
    If ckcDelinquentOnly.Value = vbChecked Then 'if delinquent only, doesnt make sense
                                                'to allow user to ask for unrestricted or zero balance
        ckcSel1(0).Value = vbUnchecked
        ckcSel1(1).Value = vbUnchecked
        ckcSel1(0).Enabled = False
        ckcSel1(1).Enabled = False
    Else
        ckcSel1(0).Enabled = True
        ckcSel1(1).Enabled = True
    End If

End Sub

Private Sub ckcInclCommentsA_Click()

    If igRptCallType = PROGRAMMINGJOB Then
        If lbcRptType.ListIndex = PRG_AIRING_INV Then
            If ckcInclCommentsA.Value = vbChecked Then          'show selling inv?
                ckcADate.Enabled = True
            Else
                ckcADate.Enabled = False
                ckcADate.Value = vbUnchecked
            End If
        End If
    Else
        If ckcInclCommentsA.Value = vbChecked Then
            lacFromA.Move 120, ckcInclCommentsA.Top + ckcInclCommentsA.Height + 30, 2400
            lacFromA.Caption = "Comment Entered as of"
            edcSelA.Move 2420, lacFromA.Top - 30
            edcSelA.Visible = True
            lacFromA.Visible = True
        Else
            lacFromA.Visible = False
            edcSelA.Visible = False
        End If
    End If

    mSetCommands
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

Private Sub ckcSelC10_click(Index As Integer)
    Dim ilListIndex As Integer
    Dim Value As Integer

    Value = False
    If ckcSelC10(Index).Value = vbChecked Then
        Value = True
    End If

    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = COLLECTIONSJOB Then
        If ilListIndex = COLL_AGEVEHICLE Then
            If Index = 0 And Value = False Then
                ckcSelC10(1).Value = vbUnchecked
            End If
            If Index = 1 And Value = True Then
                ckcSelC10(0).Value = vbChecked
            End If
        End If
    End If
End Sub

Private Sub ckcSelC3_click(Index As Integer)
Dim llVehTypes As Long
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcSelC3(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = COLLECTIONSJOB Then
        If ilListIndex = COLL_AGEVEHICLE Then       '1-22-03 Ageing by vehicle; if selecting to show owners share,
                                                    'default to participant vehicle group unless another group has been selected
            If Value Then
                If cbcSet1.ListIndex = 0 Then
                    cbcSet1.ListIndex = 1
                End If
            Else
                cbcSet1.ListIndex = 0
            End If
        End If
    ElseIf igRptCallType = AGENCIESLIST Then
        If ilListIndex = 2 And rbcSelC8(1).Value = True Then                 'mailing labels for vehicles
            ckcAll.Value = vbUnchecked
            'determine the vehicles to populate
            llVehTypes = 0


            If ckcSelC3(0).Value = vbChecked Then               'airing vehicles
                llVehTypes = llVehTypes + VEHAIRING
            End If
            If ckcSelC3(1).Value = vbChecked Then               'conventional
                llVehTypes = llVehTypes + VEHCONV_WO_FEED + VEHCONV_W_FEED
            End If
            If ckcSelC3(2).Value = vbChecked Then
                llVehTypes = llVehTypes + VEHLOG + VEHLOGVEHICLE    'LOG vehicles
            End If
            If ckcSelC3(3).Value = vbChecked Then               'NTR
                llVehTypes = llVehTypes + VEHNTR
            End If
            If ckcSelC3(4).Value = vbChecked Then               'rep
                llVehTypes = llVehTypes + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER
            End If
            If ckcSelC3(5).Value = vbChecked Then               'simulcast
                llVehTypes = llVehTypes + VEHSIMUL
            End If
            mVehLabelsPop llVehTypes
        End If
    End If
    mSetCommands
End Sub



Private Sub ckcSelC6Add_Click(Index As Integer)
mSetCommands
End Sub

Private Sub ckcSelC7_Click()
Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = INVOICESJOB Then
        If ilListIndex = INV_REGISTER Then           '3-17-05 if Invoices register & hard cost only,
                                                    'dont allow user to get air time too
            If ckcSelC7.Value = vbChecked Then
                rbcSelC6(1).Value = True
                rbcSelC6(0).Enabled = False
                rbcSelC6(1).Enabled = False
                rbcSelC6(2).Enabled = False
            Else
                rbcSelC6(0).Enabled = True
                rbcSelC6(1).Enabled = True
                rbcSelC6(2).Enabled = True
            End If
        End If
    End If
End Sub



Private Sub cmcBrowse_Click()
    'dan M 8/18/2010
    gAdjustCDCFilter imFTSelectedIndex, cdcSetup
    '9-23-02 uncomment out the next 5 instructions to make Browse button operate
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
'
'
'           5-2-01 dh gCRPlayListClear was commented out, prepass file never intialized
'                     upon completion of Copy Status by Advt or Date
'
Private Sub cmcGen_Click()
    Dim slDate As String
    Dim slTime As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilListIndex As Integer
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim slInputStartDate As String
    Dim slInputEndDate As String

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
    'LOGSJOB: igRptType = 0 or 2 => Log format; 1 or 3 => Delivery
    'If (igRptCallType = LOGSJOB) And ((igRptType = 0) Or (igRptType = 2)) And ((ilListIndex = 1) Or (ilListIndex = 3)) Then
    If (igRptCallType = LOGSJOB) And (ilListIndex = 1 Or ilListIndex = 2 Or ilListIndex = 5) Then
        'If (imGenShiftKey And vbCtrlMask) = CTRLMASK Then
        '    igUsingCrystal = True
        'Else
            igUsingCrystal = False
        'End If
    ElseIf (igRptCallType = POSTLOGSJOB) Then
        igUsingCrystal = True
    ElseIf (igRptCallType = COPYJOB) Then
        'If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then        '7-1-04
        '    ilListIndex = ilListIndex + 1
        'End If
        'If (ilListIndex = 0) Or (ilListIndex = 1) Or (ilListIndex = 2) Then 'Contracts Missing Copy
        '12-7-99 Contracts missing copy converted to Crystal
        'If (ilListIndex = 0) Or (ilListIndex = 1) Then
        igUsingCrystal = True
        'Else
        '    igUsingCrystal = False
        'End If
        'ilListIndex = lbcRptType.ListIndex
   ElseIf (igRptCallType = INVOICESJOB) Then
        If ilListIndex = 0 Or ilListIndex = 1 Then
            igUsingCrystal = True
        'ElseIf ilListIndex = 1 Then
        '    igUsingCrystal = False
        End If
    ElseIf igRptCallType = PROGRAMMINGJOB Then
'        If igRptType = 3 Then           'reports (vs links)
'            If lbcRptType.ListIndex = 0 Then        'program library
'                igUsingCrystal = True
'            End If
'        End If
        igUsingCrystal = True
    ElseIf igRptCallType = DALLASFEED Then      'everything uses crystal
        'igUsingCrystal = False
        'If ilListIndex = 1 Then
            igUsingCrystal = True
        'End If
    ElseIf igRptCallType = NYFEED Then
        'If ilListIndex = 0 Or ilListIndex = 1 Or ilListIndex = 2 Or ilListIndex = 3 Then
            igUsingCrystal = True
        'Else
        '    igUsingCrystal = False
        'End If
    ElseIf igRptCallType = PHOENIXFEED Then
        'igUsingCrystal = False
        igUsingCrystal = True
    ElseIf igRptCallType = CMMLCHG Then
        'igUsingCrystal = False
        igUsingCrystal = True
    ElseIf igRptCallType = EXPORTAFFSPOTS Then
        'igUsingCrystal = False
        igUsingCrystal = True
    ElseIf igRptCallType = BULKCOPY Then
        'If (ilListIndex = 0) Or (ilListIndex = 1) Then
        '    igUsingCrystal = False
        'Else
            igUsingCrystal = True
        'End If
    Else
        igUsingCrystal = True
    End If
    If (igRptCallType = COLLECTIONSJOB) And (ilListIndex = 7) Then  'Generate two reports
        'TTP 9893
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        ckcAll.Value = 1
        ckcAllGroups.Value = 1
        gGenCreditStatusGRF  'CreditAg.rpt and .rpt will now use GRF to support the Adv (lbcSelection(0)) and and Agcy's (lbcSelection(1)) pick lists
        
        If ckcSel2(0).Value = vbChecked Then
            ilStartJobNo = 1
            If ckcSel2(1).Value = vbChecked Then
                ilNoJobs = 2
            Else
                ilNoJobs = 1
            End If
        Else
            ilStartJobNo = 2
            ilNoJobs = 2
        End If
    Else
        ilNoJobs = 1
        ilStartJobNo = 1
    End If
    'Dan multi  credit reports needs to change here 12/17/08
    Set ogReport = New CReportHelper
    ogReport.iLastPrintJob = ilNoJobs
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReport(slYear, slMonth, slDay, slTime) Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        
            ilRet = gCmcGen(ilListIndex, imGenShiftKey, smLogUserCode, slYear, slMonth, slDay, slTime)
   
'        End If  'user wants to quit.
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
        
       End If
       
        If ilJobs >= ogReport.iLastPrintJob Then     'Dan only go through once.
            If rbcOutput(0).Value Then
                DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
                igDestination = 0
                ' Dan add rollback to 8.5 for copy book. removed 9/03/09
'                If Not bgRollback Then
'                    Report.Show vbModal
'                Else
'                    RollBackReport.Show vbModal
'                End If
                Report.Show vbModal
            ElseIf rbcOutput(1).Value Then
                ilCopies = Val(edcCopies.Text)
                ilRet = gOutputToPrinter(ilCopies)
            Else
                slFileName = edcFileName.Text
                'ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
                ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-19-01
            End If
        End If 'Dan lastPrintjob
    Next ilJobs
    Set ogReport = Nothing
    imGenShiftKey = 0
    
    'TTP 9893
    If ilListIndex = COLL_CREDITSTATUS Then
        Screen.MousePointer = vbHourglass
        gCRGrfClear
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
    cdcSetup.flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.ShowPrinter
End Sub
Private Sub edcAsOfDate_GotFocus()
    gCtrlGotFocus edcAsOfDate
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
Dim ilLen As Integer
Dim llDate As Long
Dim ilListIndex As Integer
Dim ilRet As Integer
Dim slStr As String
    ilListIndex = lbcRptType.ListIndex

    If igRptCallType = RATECARDSJOB Then
        If ilListIndex = 0 Then                 'rate card report
            ilLen = Len(edcSelCFrom)
            If ilLen = 4 Then
                slStr = gObtainEndStd("01/15/" & edcSelCFrom)
                llDate = gDateValue(slStr)
                ilRet = gPopRateCardBox(RptSelCreditStatus, llDate, lbcSelection(11), tgRateCardCode(), smRCTag, -1)
                lbcSelection(11).Visible = True
                ckcAllRC.Visible = True
            End If
        End If
    End If
    mSetCommands
End Sub
Private Sub edcSelCFrom_GotFocus()
    gCtrlGotFocus edcSelCFrom
End Sub
Private Sub edcSelCFrom_KeyPress(KeyAscii As Integer)
    Dim ilListIndex As Integer
    If igRptCallType = COPYJOB Then
        ilListIndex = lbcRptType.ListIndex
        'If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then        '7-1-04
        '    ilListIndex = ilListIndex + 1
        'End If
        If ilListIndex = 4 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
    If igRptCallType = BULKCOPY Then
        ilListIndex = lbcRptType.ListIndex
        If ilListIndex = 2 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
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
    Dim ilListIndex As Integer
    If igRptCallType = COPYJOB Then
        ilListIndex = lbcRptType.ListIndex
        'If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then        '7-1-04
        '    ilListIndex = ilListIndex + 1
        'End If
        If ilListIndex = 4 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
    If igRptCallType = BULKCOPY Then
        ilListIndex = lbcRptType.ListIndex
        If ilListIndex = 2 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
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
    RptSelCreditStatus.Refresh
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
    'RptSelCreditStatus.Show
    imFirstTime = True

    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tgAirNameCode
    Erase tgCSVNameCode
    Erase tgSellNameCode
    'Erase tgRptSelCreditStatusSalespersonCode
    'Erase tgRptSelCreditStatusAgencyCode
    'Erase tgRptSelCreditStatusAdvertiserCode
    Erase tgRptSelCreditStatusNameCode
    Erase tgRptSelCreditStatusBudgetCode
    Erase tgMultiCntrCode
    Erase tgMNFCodeRpt

    'Erase tgManyCntCode
    'Erase tgRptSelCreditStatusDemoCode
    'Erase tgSOCode
    Erase imCodes
    sgUserSortCodeTag = ""          'clear to repopulate for rentry

    PECloseEngine

    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set RptSelCreditStatus = Nothing   'Remove data segment

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcRptType_Click()
    Dim slStr As String
    Dim ilListIndex As Integer
    Dim ilTop As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim slCaption As String
    Dim ilLoopOnListBox As Integer
    Dim ilLoop As Integer
    Dim slCode As String
    Dim slName As String

    ReDim ilAASCodes(0 To 1) As Integer
    rbcSelCInclude(2).Visible = False
    Select Case igRptCallType
        Case VEHICLESLIST
            lbcSelection(0).Visible = True
'            If rbcRptType(Index).Value Then
'                Select Case Index
                Select Case lbcRptType.ListIndex
                    Case 0  'Summary
                        frcOption.Enabled = True    '5-25-01
                        pbcOption.Visible = True    '5-25-01
                        ckcSelC3(0).Move 0, 0, 2880
                        ckcSelC3(0).Caption = "Include dormant vehicles"
                        ckcSelC3(0).Value = vbUnchecked
                        ckcSelC3(0).Visible = True
                        plcSelC3.Move 120, 0
                        plcSelC3.Visible = True
                        pbcSelC.Visible = True
                    Case 1  'Options
                        frcOption.Enabled = True
                        pbcOption.Visible = True
                    Case 2   'virtual vehicles
                        frcOption.Enabled = False
                        pbcOption.Visible = False
                    Case 5  'Participant
                        frcOption.Enabled = True    '5-25-01
                        pbcOption.Visible = True    '5-25-01
                        ckcSelC3(0).Move 0, 0, 2880
                        ckcSelC3(0).Caption = "Include dormant vehicles"
                        ckcSelC3(0).Value = vbUnchecked
                        ckcSelC3(0).Visible = True
                        plcSelC3.Move 0, 520
                        plcSelC3.Visible = True
                        pbcSelC.Visible = True
                        edcCheck.Move 1400, 120, 900, 65
                        edcCheck.Visible = True
                        lacCheck.Move 0, 120, 1200, 210
                        lacCheck.Caption = "Effective Date"
                        lacCheck.Visible = True
                        edcAsOfDate.Visible = True
                        plcSelC4.Visible = True
                        plcSelC4.Move 700, plcSelC3.Top + plcSelC3.Height + 50, 3000
                        plcSelC4.ZOrder
                        rbcSelC4(0).Visible = True
                        rbcSelC4(0).Caption = "Type"
                        rbcSelC4(0).Value = True
                        rbcSelC4(0).Move 0, 0, 800
                        rbcSelC4(1).Visible = True
                        rbcSelC4(1).Caption = "Owner"
                        rbcSelC4(1).Move 800, 0, 900
                        rbcSelC4(2).Visible = False
                        lacSort.Visible = True
                        lacSort.Move 0, plcSelC3.Top + plcSelC3.Height + 50

                End Select
 '           End If
        Case ADVERTISERSLIST, AGENCIESLIST

            lbcSelection(0).Visible = True
            Select Case lbcRptType.ListIndex
                Case 0  'Summary
                    plcType.Visible = True      'detail or summary
                    rbcType(0).Visible = True
                    rbcType(1).Visible = True
                    pbcOption.Visible = True    'vehicle list
                    
                    If rbcType(0).Value Then             'default to SLSP
                        rbcType_Click 0
                    Else
                        rbcType(0).Value = True
                    End If
                    
'                    lacAsOfDate.Visible = True
'                    edcAsOfDate.Visible = True
'                    plcRepInv.Visible = True
'                    frcOption.Enabled = True
'                    pbcOption.Visible = False
'                    pbcSelB.Visible = True
                Case 1  'Credit
                    frcOption.Enabled = True
                    pbcSelB.Visible = False
                    pbcOption.Visible = True
                Case 2  'mailing labels
                    lacSelCFrom.Visible = False
                    edcSelCFrom.Visible = False
                    lacSelCTo.Visible = False
                    edcSelCTo.Visible = False
                    mAgyAdvtPop lbcSelection(2)   'Called to initialize agy and direct advertiser (statements)
                    If imTerminate Then
                        Exit Sub
                    End If
                    '2-18-05 option to output by payee (agy or direct advt) or vehicle labels
                    plcSelC8.Move 120, 0
                    rbcSelC8(0).Move 840, 0, 1200
                    rbcSelC8(0).Caption = "Payee"
                    rbcSelC8(0).Visible = True
                    rbcSelC8(0).Value = True
                    rbcSelC8(1).Move 1800, 0, 1080
                    rbcSelC8(1).Caption = "Vehicle"
                    rbcSelC8(1).Visible = True
                    If rbcSelC8(0).Value = True Then
                        rbcSelC8_Click 0
                    Else
                        rbcSelC8(0).Value = True
                    End If
                    rbcSelC8(2).Visible = False
                    smPlcSelC8P = "By"
                    plcSelC1.Move 120, plcSelC8.Top + plcSelC8.Height
                    plcSelC1.Height = 480
                    smPlcSelC1P = "Print"
                    rbcSelCSelect(0).Caption = "2 across (1" & """" & " x 4" & """" & ")"
                    rbcSelCSelect(1).Caption = "3 across (1" & """" & " x 2 5/8" & """" & ")"

                    rbcSelCSelect(0).Move 840, 0, 2280
                    rbcSelCSelect(1).Move 840, 240, 2280
                    rbcSelCSelect(2).Visible = False
                    rbcSelCSelect(0).Visible = True
                    rbcSelCSelect(1).Visible = True
                    If rbcSelCSelect(0).Value Then
                        rbcSelCSelect_click 0
                    Else
                        rbcSelCSelect(0).Value = True
                    End If

                    plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height
                    smPlcSelC2P = "Address"
                    rbcSelCInclude(0).Caption = "Contract"
                    rbcSelCInclude(0).Move 840, 0, 1200
                    rbcSelCInclude(1).Caption = "Billing"
                    rbcSelCInclude(1).Move 2040, 0, 1080

                    rbcSelCInclude(0).Visible = True
                    rbcSelCInclude(1).Visible = True
                    rbcSelCInclude(2).Visible = False
                    If rbcSelCInclude(1).Value Then
                        rbcSelCInclude_Click 1
                    Else
                        rbcSelCInclude(1).Value = True
                    End If

                    plcSelC4.Move 120, plcSelC2.Top + plcSelC2.Height
                    smPlcSelC4P = "Contact"
                    rbcSelC4(0).Caption = "Buyer"
                    rbcSelC4(0).Move 840, 0, 840
                    rbcSelC4(1).Caption = "Payables"
                    rbcSelC4(1).Move 1680, 0, 1080
                    rbcSelC4(2).Caption = "None"
                    rbcSelC4(2).Move 2760, 0, 720
                    rbcSelC4(0).Visible = True
                    rbcSelC4(1).Visible = True
                    rbcSelC4(2).Visible = True
                    If rbcSelC4(2).Value Then
                        rbcSelC4_click 2
                    Else
                        rbcSelC4(2).Value = True
                    End If
                    plcSelC1.Visible = True
                    plcSelC2.Visible = True
                    plcSelC4.Visible = True
                    plcSelC8.Visible = True
                    lbcSelection(0).Visible = False
                    lbcSelection(2).Visible = True
                    ckcAll.Caption = "All Agencies and Advertisers"
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                    pbcSelC.Visible = True
            End Select
        Case SALESPEOPLELIST
            'lbcSelection(0).Visible = True
'            If rbcRptType(Index).Value Then
'                Select Case Index
                Select Case lbcRptType.ListIndex
                    Case 0  'Summary
                        lbcSelection(1).Visible = False         'sales office
                        lbcSelection(0).Visible = True          'salespeople
                        frcOption.Enabled = True
                        pbcOption.Visible = True
                        pbcSelC.Visible = True    '5-25-01
                        ckcSelC7.Move 0, 0, 2880
                        ckcSelC7.Caption = "Include Dormant Salespeople"
                        ckcSelC7.Value = vbUnchecked
                        ckcSelC7.Visible = True
                        plcSelC7.Move 120, 120
                        plcSelC7.Visible = True

'                    Case 1  'Options
'                        lbcSelection(0).Visible = True
'                        lbcSelection(1).Visible = False
'                        frcOption.Enabled = True
'                        pbcOption.Visible = True
                End Select
'            End If
        Case EVENTNAMESLIST
        Case USERLIST           '9-28-09
            Select Case lbcRptType.ListIndex
                Case USER_OPTIONS      'user report
                    frcOption.Enabled = True    '5-25-01
                    pbcSelC.Visible = True    '5-25-01
                    ckcSelC7.Move 0, 0, 2880
                    ckcSelC7.Caption = "Include dormant users"
                    ckcSelC7.Value = vbUnchecked
                    ckcSelC7.Visible = True
                    plcSelC7.Move 120, 120
                    plcSelC7.Visible = True
                Case USER_ACTIVITY     '5-6-11 user activity log
                    slStr = Format(gNow, "m/d/yy")
                    'Start/enddates
                    llDate = gDateValue(DateAdd("d", -(tgSaf(0).iNoDaysRetainUAF + 1), slStr))
                    edcSelCFrom.Text = Format$(llDate, "m/d/yy")     'default to # days to retain earliest date
                    edcSelCFrom.Move 1920, 0
                    edcSelCFrom.MaxLength = 10
                    edcSelCFrom1.Move 1920, edcSelCFrom.Top + edcSelCFrom.Height + 30, edcSelCFrom.Width
                    edcSelCFrom1.MaxLength = 10     'xx/xx/xxxx
                    edcSelCFrom1.Text = slStr

                    edcSelCFrom.Visible = True
                    edcSelCFrom1.Visible = True
                    lacSelCFrom.Caption = "Activity Dates - Start"
                    lacSelCFrom1.Caption = "End"
                    lacSelCFrom.Move 120, edcSelCFrom.Top + 30, 1920
                    lacSelCFrom1.Move 1380, edcSelCFrom1.Top + 30, 480
                    lacSelCFrom.Visible = True
                    lacSelCFrom1.Visible = True
                    
                    lacSelCTo.Caption = "Activity Times - Start"
                    lacSelCTo.Move 120, edcSelCFrom1.Top + edcSelCFrom1.Height + 60, 1920
                    edcSelCTo.Move 1920, edcSelCFrom1.Top + edcSelCFrom1.Height + 30, edcSelCFrom.Width
                    edcSelCTo.Text = "12M"
                    edcSelCTo.MaxLength = 10
                    lacSelCTo1.Caption = "End"
                    lacSelCTo1.Move 1380, edcSelCTo.Top + edcSelCTo.Height + 60, 480
                    edcSelCTo1.Move 1920, edcSelCTo.Top + edcSelCTo.Height + 30, edcSelCTo.Width
                    edcSelCTo1.Text = "12M"
                    edcSelCTo.MaxLength = 10
                    
                    lacSelCTo.Visible = True
                    lacSelCTo1.Visible = True
                    edcSelCTo.Visible = True
                    edcSelCTo1.Visible = True
                    
                     'Major/Minor sort parameters
                    lacCheck.Move 120, edcSelCTo1.Top + edcSelCTo1.Height + 60, 840
                    lacCheck.Caption = "Sort #1"
                    cbcSet1.Move 960, edcSelCTo1.Top + edcSelCTo1.Height + 30, 1300
                    
                    mFillSortList cbcSet1, 4, False
                                          
                    lacCheck.Visible = True
                    cbcSet1.Visible = True
                    
                    lacSort.Move 120, cbcSet1.Top + cbcSet1.Height + 60, 840
                    lacSort.Caption = "Sort #2"
                    cbcSet2.Move 960, cbcSet1.Top + cbcSet1.Height + 30, 1300
                    lacSort.Visible = True
                    cbcSet2.Visible = True
                    mFillSortList cbcSet2, 0, True

                    lacCaption.Move 120, cbcSet2.Top + cbcSet2.Height + 60, 840
                    lacCaption.Caption = "Sort #3"
                    cbcSet3.Move 960, cbcSet2.Top + cbcSet2.Height + 30, 1300
                    lacCaption.Visible = True
                    mFillSortList cbcSet3, 0, True
                    cbcSet3.Visible = True
                    
                    plcSelC10.Move 2380, cbcSet1.Top, 1300, (cbcSet3.Top - cbcSet1.Top) + cbcSet3.Height
                    ckcSelC10(0).Caption = "Skip"
                    ckcSelC10(0).Visible = True
                    ckcSelC10(0).Move 240, 30, 1300
                    ckcSelC10(1).Caption = "Skip"
                    ckcSelC10(1).Visible = True
                    ckcSelC10(1).Move 240, cbcSet1.Height + 60, 1300
                    ckcSelC10(2).Move 240, ckcSelC10(1).Top + ckcSelC10(1).Height + 60, 1300
                    ckcSelC10(2).Caption = "Skip"
                    ckcSelC10(2).Visible = False        'true   hide for now, uncessary to skip for minor sort
                    plcSelC10.Visible = True
                    
                    pbcSelC.Visible = True
                    ckcAll.Caption = "All Users"
                    ckcAll.Visible = True
                    ilRet = gPopAllUsers(RptSelCreditStatus, 0, lbcSelection(10), tgUserSortCode(), sgUserSortCodeTag)
                    lbcSelection(10).Height = 3180
                    lbcSelection(10).Visible = True
                    lbcSelection(10).Enabled = True
                    pbcOption.Visible = True
                    pbcOption.Enabled = True

                    frcOption.Enabled = True
            End Select
        Case RATECARDSJOB
            'set all defaults so nothing is showing
            lbcSelection(0).Visible = False
            plcSelC2.Visible = False
            ckcAll.Visible = False
            plcSelC4.Visible = False
            lacSelCFrom.Visible = False
            edcSelCFrom.Visible = False
            lacSelCTo.Visible = False
            edcSelCTo.Visible = False
            Select Case lbcRptType.ListIndex
                Case RC_RCITEMS                  'rate card flights report
                    lbcSelection(0).Visible = True
                    lbcSelection(0).Move 15, ckcAll.Top + ckcAll.Height + 30, 4380, 1440    'vehicle list box
                    lbcSelection(11).Move 15, lbcSelection(0).Top + lbcSelection(0).Height + 360, 4380, 1440
                    ckcAllRC.Move 0, lbcSelection(0).Top + lbcSelection(0).Height + 60
                    ckcAll.Caption = "All Vehicles"
                    ckcAll.Left = 0
                    ckcAll.Visible = True
                    lbcSelection(0).Visible = True      'show vehicle list box
                    laclbcName(0).Visible = False       'hide labels in list box
                    laclbcName(1).Visible = False
                    pbcSelC.Visible = True              'show selectivity box
                    'show all rate cards existing
                    'ilRet = gPopRateCardBox(RptSelCreditStatus, 0, lbcSelection(11), tgRateCardCode(), smRateCardTag, -1)
                    'lbcSelection(11).Visible = True
                    'ckcAllRC.Visible = True

                    'rbcSelCInclude_click 1, True        'force to show Year box and/or start date box
                    plcSelC2.Move lacSelCFrom.Left, edcSelCFrom.Top + edcSelCFrom.Height + 30
                    'plcSelC2.Move 120, 30
                    'plcSelC2.Caption = "Show"
                    smPlcSelC2P = "Show"
                    rbcSelCInclude(0).Caption = "Quarter"
                    rbcSelCInclude(0).Left = 600
                    rbcSelCInclude(0).Width = 980
                    rbcSelCInclude(1).Caption = "Month"
                    rbcSelCInclude(1).Left = 1605
                    rbcSelCInclude(1).Width = 900
                    rbcSelCInclude(2).Caption = "Week"
                    rbcSelCInclude(2).Left = 2520
                    rbcSelCInclude(2).Width = 765
                    rbcSelCInclude(0).Enabled = True
                    rbcSelCInclude(1).Enabled = True
                    If rbcSelCInclude(1).Value Then
                        rbcSelCInclude_Click 1
                    Else
                        rbcSelCInclude(1).Value = True
                    End If
                    rbcSelCInclude(2).Enabled = True
                    rbcSelCInclude(2).Visible = True
                    plcSelC2.Visible = True
                    rbcSelCInclude_Click 1  ', True        'force to show Year box and/or start date box
                    plcSelC4.Move plcSelC2.Left, plcSelC2.Top + plcSelC2.Height
                    'plcSelC4.Caption = " "
                    smPlcSelC4P = " "
                    rbcSelC4(0).Caption = "Corporate"
                    rbcSelC4(0).Left = 600
                    rbcSelC4(0).Width = 1200
                    rbcSelC4(1).Caption = "Standard"
                    rbcSelC4(1).Left = 1800
                    rbcSelC4(1).Width = 1200
                    rbcSelC4(0).Enabled = True
                    rbcSelC4(1).Enabled = True
                    rbcSelC4(2).Visible = False
                    rbcSelC4(1).Value = True            'default to standard
                    If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
                        rbcSelC4(0).Enabled = False
                    Else
                        rbcSelC4(0).Value = True
                    End If
                    plcSelC4.Visible = True
                Case RC_DAYPARTS
            End Select
        Case BUDGETSJOB
            Select Case lbcRptType.ListIndex
                Case 0, 1          'budgets (office & vehicle)   or comparisons
                    plcSelC1.Top = 30
                    plcSelC1.Left = 30
                    plcSelC1.Height = 240
                    'plcSelC1.Caption = "Option"
                    smPlcSelC1P = "Option"
                    rbcSelCSelect(0).Caption = "Office"
                    rbcSelCSelect(0).Left = 600
                    rbcSelCSelect(0).Width = 775
                    rbcSelCSelect(1).Caption = "Vehicle"
                    rbcSelCSelect(1).Left = 1360
                    rbcSelCSelect(1).Width = 980
                    rbcSelCSelect(0).Enabled = True
                    rbcSelCSelect(1).Enabled = True
                    rbcSelCSelect(2).Visible = False
                    If rbcSelCSelect(1).Value Then                  'is option by vehicle already set?
                        rbcSelCSelect_click 1   ', True                        'force click event of radio button
                    Else
                        rbcSelCSelect(1).Value = True               'default to Office
                    End If
                    plcSelC1.Visible = True                         'quarter, month or week options
                    plcSelC2.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height + 30, plcSelC1.Width, plcSelC1.Height
                    'plcSelC2.Caption = "Show"
                    smPlcSelC2P = "Show"
                    rbcSelCInclude(0).Caption = "Quarter"
                    rbcSelCInclude(0).Left = 600
                    rbcSelCInclude(0).Width = 960
                    rbcSelCInclude(1).Caption = "Month"
                    rbcSelCInclude(1).Left = 1545
                    rbcSelCInclude(1).Width = 980
                    rbcSelCInclude(2).Caption = "Week"
                    rbcSelCInclude(2).Left = 2400
                    rbcSelCInclude(2).Width = 765
                    rbcSelCInclude(0).Enabled = True
                    rbcSelCInclude(1).Enabled = True
                    If rbcSelCInclude(1).Value Then
                        rbcSelCInclude_Click 1  ', True
                    Else
                        rbcSelCInclude(1).Value = True
                    End If
                    rbcSelCInclude(2).Enabled = True
                    rbcSelCInclude(2).Visible = True
                    plcSelC2.Visible = True
                    plcSelC4.Move plcSelC1.Left, plcSelC2.Top + plcSelC2.Height + 30, plcSelC1.Width, plcSelC1.Height
                    'plcSelC4.Caption = " "
                    smPlcSelC4P = " "
                    rbcSelC4(0).Caption = "Corporate"
                    rbcSelC4(0).Left = 600
                    rbcSelC4(0).Width = 1200
                    rbcSelC4(1).Caption = "Standard"
                    rbcSelC4(1).Left = 1800
                    rbcSelC4(1).Width = 1200
                    rbcSelC4(2).Visible = False
                    If tgSpf.sRUseCorpCal = "Y" Then       'if Using corp cal, dfault it; otherwise disable it
                        rbcSelC4(1).Enabled = False
                        'rbcSelC4(0).Value = True
                        If rbcSelC4(0).Value Then
                            rbcSelC4_click 0    ', True
                        Else
                            rbcSelC4(0).Value = True   'corp
                        End If
                    Else
                        rbcSelC4(0).Value = False
                        'rbcSelC4(1).Value = True
                        If rbcSelC4(1).Value Then
                            rbcSelC4_click 1    ', True
                        Else
                            rbcSelC4(1).Value = True
                        End If
                    End If
                    plcSelC4.Visible = True
                    plcSelC3.Move 120, plcSelC4.Top + plcSelC4.Height
                    'plcSelC3.Caption = ""
                    smPlcSelC3P = ""
                    ckcSelC3(0).Move 0, 0, 1680
                    ckcSelC3(0).Caption = "Summary Only"
                    ckcSelC3(0).Value = vbChecked   'True
                    ckcSelC3(0).Visible = True
                    plcSelC3.Visible = True
                    'office and vehicle list box in rbcSelCSelect
                    lbcSelection(4).Visible = True      'show budget name list box
                    laclbcName(0).Visible = True
                    laclbcName(0).Caption = "Budget Names"
                    If lbcRptType.ListIndex = 0 Then    'budget rept (vs comparisons)
                        lbcSelection(4).Move lbcSelection(0).Left, lbcSelection(0).Top + lbcSelection(0).Height + 300, lbcSelection(0).Width, lbcSelection(0).Height - 120
                        'setup label for Budget list box
                        lacSelCFrom.Visible = False         'hide label
                        edcSelCFrom.Visible = False         'hide from date
                        laclbcName(1).Visible = False
                    Else             'comparisons option
                        lbcSelection(2).Visible = True          'budget names comparison list
                        laclbcName(1).Visible = True
                        laclbcName(1).Caption = "Compare To (4 Maximum)"
                        lbcSelection(4).Move lbcSelection(0).Left, lbcSelection(0).Top + lbcSelection(0).Height + 300, lbcSelection(0).Width / 2, lbcSelection(0).Height
                        lbcSelection(2).Move lbcSelection(0).Left + lbcSelection(0).Width / 2 + 60, lbcSelection(0).Top + lbcSelection(0).Height + 300, lbcSelection(0).Width / 2, lbcSelection(0).Height - 300
                        'setup label for comparison list box
                        laclbcName(1).Move lbcSelection(0).Left + lbcSelection(0).Width / 2 + 60, lbcSelection(4).Top - laclbcName(1).Height - 30, 3420 '1710
                    End If
                    laclbcName(0).Move lbcSelection(0).Left, lbcSelection(4).Top - laclbcName(0).Height - 30, 1605
                    pbcSelC.Visible = True              'make Selectivity box visible
                    pbcOption.Visible = True            'make list boxes visible
                    edcSelCFrom1.Visible = False        'hide date from
                    lacSelCFrom1.Visible = False
                    edcSelCTo.Visible = False           'hide date to
                    edcSelCTo1.Visible = False
                    lacSelCTo1.Visible = False
                    lacSelCTo.Visible = False           'hide date to label
                    lacSelCFrom.Visible = False
                    edcSelCFrom.Visible = False
            End Select
        Case PROGRAMMINGJOB
            If igRptType = 3 Then           'reports (vs links)
                Select Case lbcRptType.ListIndex
                Case 0
                    plcSelC4.Move 120, 0
                    rbcSelC4(0).Move 0, 0, 840
                    rbcSelC4(1).Move 960, 0, 960
                    rbcSelC4(0).Caption = "Active"
                    rbcSelC4(1).Caption = "Expired"
                    rbcSelC4(1).Value = True
                    plcSelC4.Visible = True
                    rbcSelC4(0).Visible = True
                    rbcSelC4(1).Visible = True
                    rbcSelC4(2).Visible = False
                    'Start/enddates
                    edcSelCFrom.Move 1320, plcSelC4.Top + plcSelC4.Height
                    edcSelCFrom1.Move 1320, edcSelCFrom.Top + edcSelCFrom.Height + 30, edcSelCFrom.Width
                    edcSelCFrom1.MaxLength = 10     'xx/xx/xxxx
                    edcSelCFrom.Visible = True
                    edcSelCFrom1.Visible = True
                    lacSelCFrom.Caption = "Dates - Start"
                    lacSelCFrom1.Caption = "End"
                    lacSelCFrom.Move 120, edcSelCFrom.Top + 30, 1560
                    lacSelCFrom1.Move 780, edcSelCFrom1.Top + 30, 480
                    lacSelCFrom.Visible = True
                    lacSelCFrom1.Visible = True
                    
                    pbcSelC.Visible = True
                    'lbcSelection(0).Move lbcSelection(0).Left, pbcSelC.Height + 60, lbcSelection(0).Width, ckcAll.Top - pbcSelC.Height
                    ckcAll.Caption = "All Vehicles"
                    lbcSelection(0).Visible = True
                    lbcSelection(1).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(3).Visible = False
                    ilListIndex = lbcRptType.ListIndex
                End Select
            Else                        'links
                Select Case lbcRptType.ListIndex
                    Case 0  'selling vehicles
                        ckcAll.Caption = "All Vehicles"
                        ckcSel1(0).Caption = "Current"
                        ckcSel1(1).Left = ckcSel2(2).Left '1140
                        ckcSel1(0).Value = vbChecked
                        ckcSel1(1).Caption = "Pending"
                        lbcSelection(1).Visible = False
                        lbcSelection(2).Visible = False
                        lbcSelection(3).Visible = False
                        lbcSelection(0).Visible = True
                        ckcInclCommentsA.Move plcSel2.Left, plcSel2.Top + plcSel2.Height + 30, 4000     '7-24-14 option to show avail length defined
                        ckcInclCommentsA.Caption = "Include Selling Avail Lengths"
                        plcSel1.Visible = True
                        plcSel2.Visible = True
                        ckcInclCommentsA.Visible = True
                    Case 1  'Airing
                        ckcAll.Caption = "All Vehicles"
                        ckcSel1(0).Caption = "Current"
                        ckcSel1(0).Value = vbChecked
                        ckcSel1(1).Left = ckcSel2(2).Left '1140
                        ckcSel1(1).Caption = "Pending"
                        ckcSel1(1).Visible = True
                        lbcSelection(1).Visible = True
                        lbcSelection(0).Visible = False
                        lbcSelection(2).Visible = False
                        lbcSelection(3).Visible = False
                        ckcInclCommentsA.Move plcSel2.Left, plcSel2.Top + plcSel2.Height + 30, 4000 '7-24-14 option to show avail length defined
                        ckcInclCommentsA.Caption = "Include Selling Avail Lengths"
                        plcSel1.Visible = True
                        plcSel2.Visible = True
                        ckcInclCommentsA.Visible = True
                    Case 2  'Conflict
                        ckcAll.Caption = "All Vehicles"
                        plcSel1.Visible = False
                        lbcSelection(1).Visible = False 'Airing
                        lbcSelection(2).Visible = False
                        lbcSelection(3).Visible = False
                        lbcSelection(0).Visible = True  'Selling vehicle
                        plcSel2.Visible = True
                    Case 3, 5  'Airing/Conventional Vehicles
                        ckcAll.Caption = "All Vehicles"
                        ckcSel1(0).Caption = "Feed"
                        ckcSel1(1).Left = ckcSel2(2).Left - 420 '1140
                        ckcSel1(1).Caption = "Subfeed"
                        ckcSel1(1).Visible = True
                        lbcSelection(0).Visible = False
                        lbcSelection(1).Visible = False
                        lbcSelection(2).Visible = False
                        lbcSelection(3).Visible = True
                        If ckcSel1(0).Value = vbChecked Then
                            ckcSel1_click 0
                        Else
                            ckcSel1(0).Value = vbChecked    'True
                        End If
                        ckcSel1(1).Value = vbUnchecked  'False
                        plcSel1.Visible = True
                        plcSel2.Visible = True
                    Case 4, 6  'Feed
                        ckcAll.Caption = "All Feeds"
                        ckcSel1(0).Caption = "Feed"
                        ckcSel1(1).Left = ckcSel2(2).Left - 420 '1140
                        ckcSel1(1).Caption = "Subfeed"
                        ckcSel1(1).Visible = True
                        lbcSelection(0).Visible = False
                        lbcSelection(1).Visible = False
                        lbcSelection(2).Visible = True
                        lbcSelection(3).Visible = False
                        If ckcSel1(0).Value = vbChecked Then
                            ckcSel1_click 0
                        Else
                            ckcSel1(0).Value = vbChecked    'True
                        End If
                        ckcSel1(1).Value = vbUnchecked  'False
                        plcSel1.Visible = True
                        plcSel2.Visible = True
                    Case PRG_AIRING_INV                     '3-31-15
                        'default m-f, sat & sun to be obtained
                        ckcSel2(0).Value = vbChecked
                        ckcSel2(1).Value = vbChecked
                        ckcSel2(2).Value = vbChecked
                        ckcAll.Caption = "All Vehicles"
                        'current vs pending defaulted to current and disabled
                        ckcSel1(0).Caption = "Current"
                        ckcSel1(0).Value = vbChecked
                        ckcSel1(1).Left = ckcSel2(2).Left '1140
                        ckcSel1(1).Caption = "Pending"
                        ckcSel1(1).Visible = False
                        ckcSel1(0).Visible = False      'do not show current vs pending, default to current
                        plcSel2.Move 120, edcSelA.Top + edcSelA.Height + 60         'm-f, sa, su

                        lbcSelection(1).Visible = True
                        lbcSelection(1).Height = 1500
                        lbcSelection(0).Visible = False
                        lbcSelection(2).Visible = False
                        lbcSelection(3).Visible = False

                        ckcInclCommentsA.Move plcSel2.Left, plcSel2.Top + plcSel2.Height + 30, 4000 '7-24-14 option to show avail length defined
                        ckcInclCommentsA.Caption = "Include Selling Inventory"
                        ckcADate.Visible = True
                        
                        ckcADate.Caption = "Discrepancy Only"
                        ckcADate.Move ckcInclCommentsA.Left, ckcInclCommentsA.Top + ckcInclCommentsA.Height + 60, 2000
                        ckcADate.Enabled = False
                        
                        ilRet = gAvailsPop(RptSelCreditStatus, lbcSelection(5), tgNamedAvail())       'show the named avails for selectivity
                        'default the selection on if user named avails flag is set
                        For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1
                            'If tgAvailAnf(ilLoop).sRptDefault = "Y" Then                'set as selected
                                For ilLoopOnListBox = 0 To lbcSelection(5).ListCount
                                    slStr = tgNamedAvail(ilLoopOnListBox).sKey
                                    ilRet = gParseItem(slStr, 1, "\", slName)
                                    ilRet = gParseItem(slName, 3, "|", slName)
                                    ilRet = gParseItem(slStr, 2, "\", slCode)
                                Next ilLoopOnListBox
                            'End If
                        Next ilLoop
                        ckcAllGroups.Caption = "All Avail Names"
                        ckcAllGroups.Move lbcSelection(1).Left, lbcSelection(1).Top + lbcSelection(1).Height + 120
                        ckcAllGroups.Visible = True
                        lbcSelection(5).Move lbcSelection(1).Left, ckcAllGroups.Top + ckcAllGroups.Height, lbcSelection(1).Width, lbcSelection(1).Height
                        lbcSelection(5).Visible = True
                        
                        plcSel1.Visible = False         'hide current vs previous, default to current
                        plcSel2.Visible = True
                        ckcInclCommentsA.Visible = True
                End Select
            End If
        Case COLLECTIONSJOB
            mCollectionSelectivity
        Case COPYJOB
            ilListIndex = lbcRptType.ListIndex
            'If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then    '7-1-04
            '    ilListIndex = ilListIndex + 1
            'End If
            lbcSelection(0).Width = lbcSelection(6).Width       'insure advt & cntr list boxes are full width
            lbcSelection(5).Width = lbcSelection(6).Width
            lbcSelection(6).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(5).Visible = False
            edcSelCTo.MaxLength = 10    '8 5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
            edcSelCFrom.MaxLength = 10  '8
            edcSelCTo.Width = 1170
            edcSelCFrom.Width = 1170
            plcSelC3.Visible = False
            If (ilListIndex = 0) Or (ilListIndex = 1) Or (ilListIndex = 2) Or (ilListIndex = 11) Or (ilListIndex = 13) Or (ilListIndex = 14) Then    'copy playlist
                pbcOption.Visible = False
'                lacSelCFrom.Width = 900
'                lacSelCFrom.Caption = "Start Date"
'                lacSelCFrom.Visible = True
'                edcSelCFrom.Text = ""
'                edcSelCFrom.Left = 1050
'                edcSelCFrom.Visible = True
'                lacSelCTo.Width = 900
'                lacSelCTo.Caption = "End Date"
'                lacSelCTo.Visible = True
'                edcSelCTo.Text = ""
'                edcSelCTo.Left = 1050
'                edcSelCTo.Visible = True

                mAskDates False, False      '4-9-12 replace with common rtn to ask start/end dates; do not default to todays date

                If (ilListIndex = 0) Or (ilListIndex = 1) Then      'copy status by date or advt
                    lacSelCFrom1.Visible = False
                    edcSelCFrom1.Visible = False
                    'plcSelC1.Caption = "Spots"
                    smPlcSelC1P = "Spots"
                    rbcSelCSelect(2).Visible = False
                    rbcSelCSelect(0).Left = 570
                    rbcSelCSelect(0).Width = 540
                    rbcSelCSelect(0).Caption = "All"
                    rbcSelCSelect(0).Visible = True
                    rbcSelCSelect(1).Left = 1100
                    rbcSelCSelect(1).Width = 1190
                    rbcSelCSelect(1).Caption = "With Copy"
                    rbcSelCSelect(1).Visible = True
                    rbcSelCSelect(2).Caption = "Without Copy"
                    rbcSelCSelect(2).Left = 2300
                    rbcSelCSelect(2).Width = 1580
                    rbcSelCSelect(2).Visible = True
                    rbcSelCSelect(2).Enabled = True
                    If rbcSelCSelect(0).Value Then
                        rbcSelCSelect_click 0
                    Else
                        rbcSelCSelect(0).Value = True    'All
                    End If
                    'plcSelC2.Caption = "Include Unassigned"
                    smPlcSelC2P = "Include Unassigned"
                    rbcSelCInclude(0).Left = 1770
                    rbcSelCInclude(0).Width = 650
                    rbcSelCInclude(0).Caption = "Yes"
                    rbcSelCInclude(0).Visible = True
                    rbcSelCInclude(1).Left = 2385
                    rbcSelCInclude(1).Width = 510
                    rbcSelCInclude(1).Caption = "No"
                    rbcSelCInclude(1).Visible = True
                    If rbcSelCInclude(0).Value Then
                        rbcSelCInclude_Click 0
                    Else
                        rbcSelCInclude(0).Value = True   'Yes for All
                    End If
                    plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height
                    plcSelC2.Enabled = False
                    plcSelC2.Visible = True
                    lbcSelection(2).Visible = False
                    plcSelC1.Visible = True

                    plcSelC10.Move 120, plcSelC2.Top + plcSelC2.Height, 4000
                    ckcSelC10(0).Move 720, 0, 1680       'local
                    ckcSelC10(1).Move 2400, 0, 1440      'feed
                    ckcSelC10(0).Value = vbChecked
                    ckcSelC10(1).Value = vbChecked
                    If tgSpf.sSystemType = "R" Then         'radio vs network/syndicator
                        ckcSelC10(0).Visible = True
                        ckcSelC10(0).Caption = "Contract spots"
                        ckcSelC10(1).Visible = True
                        ckcSelC10(1).Caption = "Feed spots"
                        plcSelC10.Visible = True
                        smPlcSelC10P = "Include"
                        plcSelC10_Paint
                        plcSelC10.Visible = True
                    End If

                    If ilListIndex = 0 Then
                        mSellConvVirtVehPop 1, True     'w/o package vehicles
                        If imTerminate Then
                            cmcCancel_Click
                            Exit Sub
                        End If
                        '***lbcSelection(1).Move lbcSelection(0).Left, pbcSelC.Height + 60, lbcSelection(0).Width, ckcAll.Top ' pbcSelC.Height
                        lbcSelection(0).Visible = False
                        lbcSelection(1).Visible = True
                        ckcAll.Caption = "All Vehicles"
                        ckcAll.Visible = True
                        pbcOption.Visible = True
                    Else
                        '***lbcSelection(0).Move lbcSelection(0).Left, pbcSelC.Height + 60, lbcSelection(0).Width, ckcAll.Top '- pbcSelC.Height
                        mSellConvVirtVehPop 6, True     'w/o package vehicles
                        If imTerminate Then
                            cmcCancel_Click
                            Exit Sub
                        End If
                        lbcSelection(0).Width = lbcSelection(0).Width / 2   '
                        lbcSelection(0).Visible = True
                        lbcSelection(1).Visible = False
                        ckcAll.Caption = "All Advertisers"
                        ckcAll.Visible = True
                        pbcOption.Visible = True
                    End If
                ElseIf ilListIndex = 2 Then             'contracts missing copy
                    lacSelCFrom1.Visible = False
                    edcSelCFrom1.Visible = False
                    'mSellConvVirtVehPop 6, True       'ignore package vehicles
                    mSellConvAirPop 6, True       'selling, airing & conventional vehicles
                    If imTerminate Then
                        cmcCancel_Click
                        Exit Sub
                    End If
                    mRemoveAirVeh lbcSelection(6), tgVehicle()

                    rbcSelCSelect(0).Caption = "Advertiser"
                    rbcSelCSelect(1).Caption = "Vehicle"
                    rbcSelCSelect(0).Value = True
                    smPlcSelC1P = "Sort by"
                    plcSelC1.Visible = True
                    rbcSelCSelect(0).Move 735, 0, 1200
                    rbcSelCSelect(1).Move 2055, 0, 960

                    rbcSelCSelect(0).Visible = True
                    rbcSelCSelect(1).Visible = True
                    rbcSelCSelect(2).Visible = False

                    plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height
                    smPlcSelC2P = "Show by "
                    rbcSelCInclude(0).Caption = "Contract"
                    rbcSelCInclude(1).Caption = "Line"
                    rbcSelCInclude(0).Value = True
                    rbcSelCInclude(0).Move 840, 0, 1200
                    rbcSelCInclude(1).Move 2040, 0, 720
                    plcSelC2.Visible = True
                    rbcSelCInclude(1).Visible = True
                    'plcSelC1.Visible = False
                    'plcSelC2.Visible = False
                    plcSelC3.Move 120, plcSelC2.Top + plcSelC2.Height
                    'plcSelC3.Move 120, edcSelCTo.Top + edcSelCTo.Height + 60

                    smPlcSelC3P = "Include"
                    ckcSelC3(0).Left = 735
                    ckcSelC3(0).Width = 1280
                    ckcSelC3(0).Caption = "Unassigned"
                    ckcSelC3(1).Left = 2025
                    ckcSelC3(1).Width = 1380
                    ckcSelC3(1).Caption = "To Reassign"
                    '5-8-05 remove the option to exclude Missing copy since thats the intent of this report
                    'ckcSelC3(2).Visible = False
                    '6-17-05 missed option reinstated.  this is for missed spots, not missing copy
                    ckcSelC3(2).Visible = True
                    ckcSelC3(2).Caption = "Missed"
                    ckcSelC3(2).Move 3375, -30, 920
                    ckcSelC3(3).Visible = False

                    If Not (ckcSelC3(0).Value = vbChecked) Then
                        ckcSelC3_click 0
                    Else
                        ckcSelC3(0).Value = vbUnchecked 'False
                    End If

                    'DS ?????
                    If ckcSelC3(1).Value = vbChecked Then
                        ckcSelC3_click 1
                    Else
                        ckcSelC3(1).Value = vbUnchecked 'False
                    End If

                    If ckcSelC3(2).Value = vbChecked Then
                        ckcSelC3_click 1
                    Else
                        ckcSelC3(2).Value = vbChecked   'True
                    End If
                    ckcSelC3(0).Visible = True
                    ckcSelC3(1).Visible = True
                    plcSelC3.Visible = True


                    plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height
                    ckcSelC5(0).Top = 0
                    ckcSelC5(0).Left = 0
                    plcSelC5.Visible = True
                    ckcSelC5(0).Value = vbChecked   'True
                    ckcSelC5(0).Visible = True

                    plcSelC9.Move 120, plcSelC5.Top + plcSelC5.Height
                    ckcTrans.Top = 0
                    ckcTrans.Left = 0
                    ckcTrans.Caption = "Show Daypart Name"
                    plcSelC9.Visible = True
                    ckcTrans.Value = vbUnchecked   'True
                    ckcTrans.Visible = True

                    '11-16-05 selective contract # (mainly for testing)
                    lacCheck.Caption = "Contract #"
                    lacCheck.Move 120, plcSelC9.Top + plcSelC9.Height + 30
                    edcCheck.Move 1200, lacCheck.Top - 30
                    lacCheck.Visible = True
                    edcCheck.Visible = True

                    'plcSelC10.Move 120, plcSelC5.Top + plcSelC5.Height - 30
                    plcSelC10.Move 120, edcCheck.Top + edcCheck.Height + 30
                    ckcSelC10(0).Move 720, 0, 1680       'local
                    ckcSelC10(1).Move 2400, 0, 1440      'feed
                    ckcSelC10(0).Value = vbChecked
                    ckcSelC10(1).Value = vbChecked
                    If tgSpf.sSystemType = "R" Then         'radio vs network/syndicator
                        ckcSelC10(0).Visible = True
                        ckcSelC10(0).Caption = "Contract spots"
                        ckcSelC10(1).Visible = True
                        ckcSelC10(1).Caption = "Feed spots"
                        plcSelC10.Visible = True
                        smPlcSelC10P = "Include"
                        plcSelC10_Paint
                        plcSelC7.Move 120, plcSelC10.Top + plcSelC10.Height
                    Else
                        plcSelC7.Move 120, plcSelC10.Top
                    End If
                    ckcSelC7.Move 0, 0, 4000
                    ckcSelC7.Caption = "Skip to new page each new vehicle"
                    plcSelC7.Visible = False        'dont allow skip to new page on advertiser sort
                    ckcSelC7.Visible = False
                    ckcSelC7.Value = False
                    '***lbcSelection(1).Move lbcSelection(0).Left, pbcSelC.Height + 60, lbcSelection(0).Width, ckcAll.Top - pbcSelC.Height
                    lbcSelection(0).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(1).Visible = False
                    lbcSelection(6).Visible = True
                    ckcAll.Caption = "All Vehicles"
                    ckcAll.Visible = True
                    pbcOption.Visible = True
                ElseIf ilListIndex = 11 Or ilListIndex = 13 Or ilListIndex = 14 Then     'Play List by ISCI (11), vehicle (13), Advt (14)
                    'plcSelC1.Caption = "Show by"
                    smPlcSelC1P = "Show by"
                    rbcSelCSelect(2).Visible = False
                    rbcSelCSelect(0).Left = 870
                    rbcSelCSelect(0).Width = 1050
                    rbcSelCSelect(0).Caption = "Vehicle"
                    rbcSelCSelect(1).Left = 2000
                    rbcSelCSelect(1).Width = 1110
                    rbcSelCSelect(1).Caption = "ISCI"

                    'no idea why this is setting the radio button.
                    'possibly used to be choices as to the type of playlist?
                    If ilListIndex = 11 Then
                        If rbcSelCSelect(1).Value Then  'default is ISCI
                            rbcSelCSelect_click 1   ', True
                        Else
                            rbcSelCSelect(1).Value = True
                        End If
                    ElseIf ilListIndex = 13 Then
                        If rbcSelCSelect(0).Value Then  'default is ISCI
                            rbcSelCSelect_click 0   ', True
                        Else
                            rbcSelCSelect(0).Value = True
                        End If
                    ElseIf ilListIndex = 14 Then
                        If rbcSelCSelect(2).Value Then  'default is Advertiser
                            rbcSelCSelect_click 2   ', True
                        Else
                            rbcSelCSelect(2).Value = True
                        End If
                    End If

                    'plcSelC1.Visible = True
                    plcSelC1.Visible = False
                    plcSelC2.Visible = False
                    'turn on in rbcselect
                    'lbcselection(0).Visible = False
                    'lbcselection(1).Visible = False
                    pbcOption.Visible = True
                End If
                pbcSelC.Visible = True
            ElseIf ilListIndex = COPY_ROT Then 'Rotation by Advertiser
                lacSelCFrom1.Visible = False
                edcSelCFrom1.Visible = False
                plcSelC1.Visible = False
                plcSelC2.Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                
                '8-3-10 add span of rotations vs just effective date
                lacSelCFrom.Visible = True
                lacSelCFrom.Caption = "Rotations Active-   Start"
                lacSelCFrom.Move 120, 120, 2040
                
                edcSelCFrom.Move 2280, 120, 1140
                edcSelCFrom.Visible = True
                edcSelCFrom.Text = ""
                               
                edcSelCFrom1.Move 2280, edcSelCFrom.Top + edcSelCFrom.Height + 30, 1140
                edcSelCFrom1.Visible = True
                edcSelCFrom1.Text = ""
                
                lacSelCFrom1.Move 1680, edcSelCFrom1.Top + 30
                lacSelCFrom1.Caption = "End"
                lacSelCFrom1.Visible = True
                
                '10-30-10 add span of date entered
                lacSelCTo.Visible = True
                lacSelCTo.Caption = "Rotations Entered- Start"
                lacSelCTo.Move 120, edcSelCFrom1.Top + edcSelCFrom1.Height + 30, 2040
                
                edcSelCTo.Move 2280, lacSelCTo.Top, 1140
                edcSelCTo.Visible = True
                edcSelCTo.Text = ""
                
                edcSelCTo1.Move 2280, edcSelCTo.Top + edcSelCTo.Height + 30, 1140
                edcSelCTo1.Visible = True
                edcSelCTo1.Text = ""
                
                lacSelCTo1.Move 1740, edcSelCTo1.Top + 30
                lacSelCTo1.Caption = "End"
                lacSelCTo1.Visible = True
                edcSelCFrom1.MaxLength = 10
                edcSelCTo1.MaxLength = 10

                                
                plcSelC5.Move 120, edcSelCTo1.Top + edcSelCTo1.Height
                ckcSelC5(0).Caption = "Show Inventory"
                ckcSelC5(0).Move 0, 0
                ckcSelC5(0).Value = vbChecked
                ckcSelC5(0).Visible = True
                plcSelC5.Visible = True

                plcSelC7.Move 120, plcSelC5.Top + plcSelC5.Height
                ckcSelC7.Caption = "Show Assign Dates"
                ckcSelC7.Move 0, 0
                ckcSelC7.Value = vbChecked
                ckcSelC7.Visible = True
                plcSelC7.Visible = True
                
                '8-3-10 add option to show rotation comments
                plcSelC9.Move 120, plcSelC7.Top + plcSelC7.Height
                ckcTrans.Caption = "Include Comments"
                ckcTrans.Move 0, 0
                ckcTrans.Value = vbChecked
                ckcTrans.Visible = True
                plcSelC9.Visible = True
                
                '8-11-10 option to include dormant rotations
                plcSelC11.Move plcSelC9.Left, plcSelC9.Top + plcSelC9.Height
                ckcSelC11(0).Move 0, 0, 3120
                ckcSelC11(0).Caption = "Include Dormant Rotations"
                plcSelC11.Visible = True
                ckcSelC11(0).Visible = True
                
                ckcOption.Move 120, plcSelC11.Top + plcSelC11.Height + 60, 4000, 400
                ckcOption.Caption = "For selected vehicles on hidden lines, check+show pkg rotations"
                ckcOption.Visible = True
                
                pbcSelC.Visible = True
                '***lbcSelection(0).Move lbcSelection(0).Left, pbcSelC.Height + 60, lbcSelection(0).Width, ckcAll.Top - pbcSelC.Height
                lbcSelection(0).Visible = True
                lbcSelection(0).Width = lbcSelection(0).Width / 2   '
                lbcSelection(0).Height = 1605
                lbcSelection(5).Height = 1605
                lbcSelection(6).Height = 1530
                ckcAllGroups.Move lbcSelection(0).Left, lbcSelection(0).Top + lbcSelection(0).Height + 120
                ckcAllGroups.Caption = "All Vehicles"
                lbcSelection(6).Move lbcSelection(0).Left, ckcAllGroups.Top + ckcAllGroups.Height
                lbcSelection(6).Visible = True

                ckcAllGroups.Visible = True
                ckcAll.Caption = "All Advertisers"
                ckcAll.Visible = True
                pbcOption.Visible = True
            ElseIf ilListIndex = COPY_SPLITROT Then     '1-30-09
                'lacSelCFrom1.Visible = False
                'edcSelCFrom1.Visible = False
                plcSelC1.Visible = False
                plcSelC2.Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                'lacSelCFrom.Visible = False    'True
                'edcSelCFrom.Visible = False 'True
                lacSelCFrom.Width = 2280
                lacSelCFrom.Caption = "Active Rotations-Start"
                lacSelCFrom.Top = 120
                lacSelCFrom.Visible = True
                edcSelCFrom.Move 2040, 120, 960
                edcSelCFrom.MaxLength = 10
                edcSelCFrom.Visible = True

                lacSelCTo.Move 1560, edcSelCFrom.Top + edcSelCFrom.Height + 60, 480
                lacSelCTo.Visible = True
                lacSelCTo.Caption = "End"
                edcSelCTo.Move edcSelCFrom.Left, edcSelCFrom.Top + edcSelCFrom.Height + 60, 960
                edcSelCTo.Visible = True
                edcSelCTo.MaxLength = 10

                plcSelC1.Move 120, edcSelCTo.Top + edcSelCTo.Height + 60
                rbcSelCSelect(0).Move 600, 0, 1200
                rbcSelCSelect(1).Move 1800, 0, 1080
                rbcSelCSelect(2).Move 2880, 0, 720
                rbcSelCSelect(0).Caption = "Split Copy"
                rbcSelCSelect(1).Caption = "Blackout"
                rbcSelCSelect(2).Caption = "Both"
                rbcSelCSelect(2).Value = True
                rbcSelCSelect(0).Visible = True
                rbcSelCSelect(1).Visible = True
                rbcSelCSelect(2).Visible = True
                plcSelC1.Visible = True
                plcSelC7.Move 120, plcSelC1.Top + plcSelC1.Height
                ckcSelC7.Move 0, 0, 3000
                ckcSelC7.Caption = "Include Dormant Rotations"
                ckcSelC7.Visible = True
                plcSelC7.Visible = True

                pbcSelC.Visible = True
                lbcSelection(0).Visible = True
                lbcSelection(0).Width = lbcSelection(0).Width
                lbcSelection(0).Height = 1605
                lbcSelection(5).Height = 1605
                lbcSelection(6).Height = 1530
                ckcAllGroups.Move lbcSelection(0).Left, lbcSelection(0).Top + lbcSelection(0).Height + 120
                ckcAllGroups.Caption = "All Vehicles"
                lbcSelection(6).Move lbcSelection(0).Left, ckcAllGroups.Top + ckcAllGroups.Height
                lbcSelection(6).Visible = True

                ckcAllGroups.Visible = True
                ckcAll.Caption = "All Advertisers"
                ckcAll.Visible = True
                pbcOption.Visible = True
            ElseIf ilListIndex = 4 Then 'Inventory by numbers
                lacSelCFrom1.Visible = False
                edcSelCFrom1.Visible = False
                plcSelC1.Visible = False
                plcSelC2.Visible = False
                pbcOption.Visible = False
                edcSelCTo.MaxLength = 12
                edcSelCFrom.MaxLength = 12
                lacSelCFrom.Width = 1300
                lacSelCFrom.Caption = "Lowest Cart #"
                lacSelCFrom.Visible = True
                edcSelCFrom.Text = ""
                edcSelCFrom.Left = 1350
                edcSelCFrom.Visible = True
                lacSelCTo.Width = 1300
                lacSelCTo.Caption = "Highest Cart #"
                lacSelCTo.Visible = True
                edcSelCTo.Text = ""
                edcSelCTo.Left = 1350
                edcSelCTo.Visible = True
                pbcSelC.Visible = True
            ElseIf ilListIndex = 5 Then 'Inventory by ISCI
                lacSelCFrom1.Visible = False
                edcSelCFrom1.Visible = False
                plcSelC1.Visible = False
                plcSelC2.Visible = False
                pbcOption.Visible = False
                lacSelCFrom.Width = 1300
                lacSelCFrom.Caption = "Lowest ISCI #"
                lacSelCFrom.Visible = True
                edcSelCFrom.Text = ""
                edcSelCFrom.Left = 1350
                edcSelCFrom.MaxLength = 20
                edcSelCFrom.Width = 2340
                edcSelCFrom.Visible = True
                lacSelCTo.Width = 1300
                lacSelCTo.Caption = "Highest ISCI #"
                lacSelCTo.Visible = True
                edcSelCTo.Text = ""
                edcSelCTo.Left = 1350
                edcSelCTo.MaxLength = 20
                edcSelCTo.Width = 2340
                edcSelCTo.Visible = True
                pbcSelC.Visible = True
            ElseIf ilListIndex = 6 Then  'Inventory by Advertiser
                lacSelCFrom1.Visible = False
                edcSelCFrom1.Visible = False
                pbcSelC.Visible = False
                '**lbcSelection(0).Move lbcSelection(0).Left, 90, lbcSelection(0).Width, ckcAll.Top - pbcSelC.Top - 90
                lbcSelection(2).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(0).Visible = True
                ckcAll.Caption = "All Advertisers"
                ckcAll.Visible = True
                pbcOption.Visible = True
            ElseIf (ilListIndex = 7) Or (ilListIndex = 8) Or (ilListIndex = 9) Or (ilListIndex = 10) Then 'Inventory by Start Date; Expiration Date; Purge Date; Entry Date
                If ilListIndex = 7 Then
                    slCaption = ""
                ElseIf ilListIndex = 8 Then
                    slCaption = "Expired- "
                ElseIf ilListIndex = 9 Then
                    slCaption = "Purge- "
                ElseIf ilListIndex = 10 Then
                    slCaption = "Entry- "
                End If
                mAskDates False, True, slCaption                      '4-9-12 ask start/end dates, default end date to todays date
                pbcOption.Visible = False

                plcSelC5.Visible = False
                ckcSelC5(0).Visible = False
                plcSelC6.Visible = False
                If ilListIndex = COPY_INVBYSTARTDATE Or ilListIndex = 10 Then            '4-13-05 inventory by entry date, ask printables only
                    If ilListIndex = 10 Then        'by Entry date
                        '4-9-13 Ask span of dates sent
                        '4-10-13 hide and create new report instead.
'                        lacSelCFrom1.Move lacSelCTo.Left, edcSelCTo.Top + edcSelCTo.Height + 30, 1800
'                        edcText1.Move edcSelCTo.Left, edcSelCTo.Top + edcSelCTo.Height + 30, edcSelCTo.Width
'                        lacSelCFrom1.Caption = "Sent- Start Date"
'                        lacSelCFrom1.Visible = True
'                        edcText1.Visible = True
'
'                        lacSelCTo1.Move lacSelCTo.Left, edcText1.Top + edcText1.Height + 30, 1800
'                        edcText2.Move edcSelCTo.Left, edcText1.Top + edcText1.Height + 30, edcSelCTo.Width
'                        lacSelCTo1.Caption = "Sent- End Date"
'                        lacSelCTo1.Visible = True
'                        edcText2.Visible = True
'                        plcSelC5.Move 120, edcText2.Top + edcText2.Height + 30

                        plcSelC5.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
                        ckcSelC5(0).Caption = "Printables Only"
                        ckcSelC5(0).Move 0, 0
                        ckcSelC5(0).Visible = True
                        plcSelC5.Visible = True
                        ilTop = plcSelC5.Top + plcSelC5.Height
                    Else                    'by start date
                        ilTop = edcSelCTo.Top + edcSelCTo.Height + 30
                        'inventory by start date, allow selectivity by Media code
                        ilRet = gPopMCFBox(RptSelCreditStatus, lbcSelection(10), tgMcfCode(), sgMcfTagRpt)
                        ckcAll.Visible = True
                        ckcAll.Caption = "All Media Codes"
                        If tgSpf.sUseCartNo = "N" Then
                            lbcSelection(10).Visible = False
                            ckcAll.Visible = False
                            ckcAll.Value = vbChecked
                        Else
                            lbcSelection(10).Visible = True
                            pbcOption.Visible = True
                        End If
                                        ' added option to include salesperson 6-04-08 Dan M
                        ckcSelC10(0).Caption = "Salesperson"
                        ckcSelC10(0).Move 0, 0, 1500
                        ckcSelC10(0).Value = 0
                        ckcSelC10(0).Visible = True
                        lacCheck.Caption = "Include"
                        lacCheck.Move 120, ilTop + plcSelC6.Height, 700
                        lacCheck.Visible = True
                        plcSelC10.Move lacCheck.Width + 120, ilTop + plcSelC6.Height
                        plcSelC10.Visible = True

                    End If

                    If tgSpf.sTapeShowForm = "C" Then
                        plcSelC6.Move 120, ilTop
                        rbcSelC6(0).Caption = "Carted"
                        rbcSelC6(1).Caption = "Uncarted"
                        rbcSelC6(2).Caption = "Both"
                        rbcSelC6(0).Move 720, 0, 960
                        rbcSelC6(1).Move 1740, 0, 1080
                        rbcSelC6(2).Move 2940, 0, 720
                        rbcSelC6(2).Value = True

                        rbcSelC6(0).Visible = True
                        rbcSelC6(1).Visible = True
                        rbcSelC6(2).Visible = True
                        smPlcSelC6P = "Include"
                        plcSelC6.Visible = True
                    Else
                        rbcSelC6(2).Value = True        'not using the carting feature, force to show all types for cifCleared field
                    End If

                End If
                pbcSelC.Visible = True
            ElseIf ilListIndex = COPY_INVPRODUCER Then      '4-10-13
                mAskDates True, False, "Rotation- "                     '4-9-12 ask start/end dates, default end date to todays date"
                
                plcSelC3.Move lacSelCTo.Left, edcSelCTo.Top + edcSelCTo.Height + 30, 4380, 525

                smPlcSelC3P = "Action Status"
                ckcSelC3(0).Caption = "Not Sent"
                ckcSelC3(1).Caption = "Sent"
                ckcSelC3(2).Caption = "Produced"
                ckcSelC3(3).Caption = "Held"
                ckcSelC3(0).Value = vbUnchecked
                ckcSelC3(1).Value = vbChecked
                ckcSelC3(2).Value = vbUnchecked
                ckcSelC3(3).Value = vbUnchecked
                ckcSelC3(0).Move 1200, 0, 1200
                ckcSelC3(1).Move 2400, 0, 720
                ckcSelC3(2).Move 1200, 255, 1200
                ckcSelC3(3).Move 2400, 255, 720
                ckcSelC3(0).Visible = True
                ckcSelC3(1).Visible = True
                ckcSelC3(2).Visible = True
                ckcSelC3(3).Visible = True
                plcSelC3.Visible = True
                If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
                    plcSelC7.Move 120, plcSelC3.Top + plcSelC3.Height, 3000
                    ckcSelC7.Move 0, 0, 3000
                    ckcSelC7.Caption = "Retrieve vCreative Statuses"
                    ckcSelC7.Value = vbChecked
                    ckcSelC7.Visible = True
                    plcSelC7.Visible = True
                Else
                    ckcSelC7.Value = vbUnchecked
                End If

                ilRet = gPopMCFBox(RptSelCreditStatus, lbcSelection(10), tgMcfCode(), sgMcfTagRpt)
                ckcAll.Visible = True
                ckcAll.Caption = "All Media Codes"
                lbcSelection(10).Visible = True
                pbcOption.Visible = True
                pbcSelC.Visible = True

            ElseIf ilListIndex = 12 Then 'unapproved copy
                lacSelCFrom1.Visible = False
                edcSelCFrom1.Visible = False
                plcSelC1.Visible = False
                plcSelC2.Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lacSelCFrom.Visible = False
                edcSelCFrom.Visible = False
                lacSelCTo.Width = 1800
                lacSelCTo.Caption = "Received on or after"
                lacSelCTo.Visible = True
                edcSelCTo.Text = ""
                edcSelCTo.Left = 1890
                edcSelCTo.Visible = True
                pbcSelC.Visible = True
                'lbcSelection(0).Move lbcSelection(0).Left, pbcSelC.Height + 60, lbcSelection(0).Width, ckcAll.Top - pbcSelC.Height
                'lbcSelection(0).Visible = True
                'ckcAll.Caption = "All Vehicles"
                'ckcAll.Visible = True
                pbcOption.Visible = False
            'copy regions merged into RptSelCreditStatussr (copy regions & split regions)
'            ElseIf ilListIndex = COPY_REGIONS Then     '7-18-00Copy regions by advt or regions
'                lacSelCFrom.Width = 1920
'                lacSelCFrom.Caption = "Creation Start Date"
'                lacSelCFrom.Visible = True
'                edcSelCFrom.Text = ""
'                edcSelCFrom.Left = 1920
'                edcSelCFrom.Visible = True
'                lacSelCTo.Width = 1920
'                lacSelCTo.Caption = "Creation End Date"
'                lacSelCTo.Visible = True
'                edcSelCTo.Text = ""
'                edcSelCTo.Left = 1920
'                edcSelCTo.Visible = True
'                'plcSelC1.Caption = "Sort by"
'                smPlcSelC1P = "Sort by"
'                rbcSelCSelect(0).Caption = "Advertiser"
'                rbcSelCSelect(0).Visible = True
'                rbcSelCSelect(0).Move 720, 0, 1200
'                rbcSelCSelect(1).Caption = "Region"
'                rbcSelCSelect(1).Visible = True
'                rbcSelCSelect(1).Move 2010
'                rbcSelCSelect(2).Visible = False
'                If rbcSelCSelect(0).Value Then                  'is option by advt already set?
'                    rbcSelCSelect_click 0   ', True                        'force click event of radio button
'                Else
'                    rbcSelCSelect(0).Value = True               'default to advt
'                End If
'                plcSelC1.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
'                plcSelC1.Visible = True
'                pbcSelC.Visible = True
'                lbcSelection(0).Visible = True
'                lbcSelection(0).Width = lbcSelection(0).Width
'                ckcAll.Caption = "All Advertisers"
'                ckcAll.Visible = True
'                pbcOption.Visible = True
'                pbcOption.Visible = True
            ElseIf (ilListIndex = COPY_BOOK) Then           '8-30-05
'                lacSelCFrom.Width = 900
'                lacSelCFrom.Caption = "Start Date"
'                lacSelCFrom.Visible = True
'                edcSelCFrom.Text = ""
'                edcSelCFrom.Left = 1050
'                edcSelCFrom.Visible = True
'                lacSelCTo.Width = 900
'                lacSelCTo.Caption = "End Date"
'                lacSelCTo.Visible = True
'                edcSelCTo.Text = ""
'                edcSelCTo.Left = 1050
'                edcSelCTo.Visible = True

                mAskDates False, False                      'ask start/end dates, do not default the dates to todays date
                ckcAll.Caption = "All Vehicles"
                ckcAll.Visible = True
                lbcSelection(2).Visible = True
                pbcOption.Visible = True
                pbcSelC.Visible = True
        
            ElseIf ilListIndex = COPY_SCRIPTAFFS Then          '4-9-12
                pbcSelC.Visible = True                          'this shows the start/end dates
                lbcSelection(2).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(0).Visible = True
                ckcAll.Caption = "All Advertisers"
                ckcAll.Visible = True
                pbcOption.Visible = True

                mAskDates False, False                      'ask start/end dates, do not default the dates to todays date
                plcSelC7.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
                ckcSelC7.Move 0, 0, 2880
                ckcSelC7.Caption = "Show Notarization Text"
                ckcSelC7.Value = vbChecked
                ckcSelC7.Visible = True
                plcSelC7.Visible = True
                plcSelC9.Move 120, plcSelC7.Top + plcSelC7.Height + 30
                ckcTrans.Caption = "Show Inventory Detail"
                ckcTrans.Visible = True
                ckcTrans.Move 0, 0, 2280
                plcSelC9.Visible = True
            End If
        Case POSTLOGSJOB
            ilListIndex = lbcRptType.ListIndex
            If ilListIndex = 0 Then     'Post Log Status
                frcOption.Enabled = True

                lacSelCFrom.Caption = "From Date"
                lacSelCFrom.Visible = True
                lacSelCTo1.Width = 700
                lacSelCTo1.Move 120, 720
                lacSelCTo1.Caption = "Contr #"
                lacSelCTo1.Visible = True

                lacSelCTo.Caption = "To Date"
                lacSelCTo.Visible = True

                edcSelCFrom.Visible = True     'From Date   Tab 43
                edcSelCFrom.Width = 1170
                edcSelCTo.Visible = True       'To Date     Tab 47
                edcSelCTo.Width = 1170

                edcSelCTo1.Move 1500, 665, 1170
                edcSelCTo1.Visible = True      'Contr #     Tab 49



                plcSelC1.Visible = False
                plcSelC2.Visible = False
                pbcSelC.Visible = True

                plcSelC3.Left = 120
                plcSelC3.Top = 1020
                plcSelC3.Height = 480
                'plcSelC3.Caption = "Include"
                smPlcSelC3P = "Include"
                ckcSelC3(0).Left = 750
                ckcSelC3(0).Width = 930
                ckcSelC3(0).Caption = "Billed"
                If ckcSelC3(0).Value = vbChecked Then
                    ckcSelC3_click 0
                Else
                    ckcSelC3(0).Value = vbChecked   'True
                End If
                ckcSelC3(1).Left = 1680
                ckcSelC3(1).Width = 990
                ckcSelC3(1).Caption = "Unbilled"
                If ckcSelC3(1).Value = vbChecked Then
                    ckcSelC3_click 1
                Else
                    ckcSelC3(1).Value = vbChecked   'True
                End If
                'ckcSelC3(2).Left = 2265
                'ckcSelC3(2).Width = 930
                'ckcSelC3(2).Caption = "Missed"
                'ckcSelC3(2).Value = True
                ckcSelC3(3).Left = 2760 '3165
                ckcSelC3(3).Width = 1270
                ckcSelC3(3).Caption = "PSA/Promo"
                ckcSelC3(3).Value = vbUnchecked 'False
                ckcSelC3(4).Left = 750
                ckcSelC3(4).Width = 980
                ckcSelC3(4).Caption = "Missed"
                ckcSelC3(4).Value = vbUnchecked 'False
                ckcSelC3(5).Left = 1790
                ckcSelC3(5).Width = 1160
                ckcSelC3(5).Caption = "Cancelled"
                ckcSelC3(5).Value = vbUnchecked 'False
                ckcSelC3(6).Left = 3010
                ckcSelC3(6).Width = 980
                ckcSelC3(6).Caption = "Hidden"
                ckcSelC3(6).Value = vbUnchecked 'False

                ckcSelC3(0).Visible = True
                ckcSelC3(1).Visible = True
                ckcSelC3(2).Visible = False 'True
                ckcSelC3(3).Visible = True
                ckcSelC3(4).Visible = True
                ckcSelC3(5).Visible = True
                ckcSelC3(6).Visible = True
                '5-27-05 option for +fill / -fill, not necessary for P.L report, but  code is the same routine
                ckcSelC3(7).Visible = False
                ckcSelC3(8).Visible = False
                ckcSelC3(7).Value = vbChecked
                ckcSelC3(8).Value = vbChecked

                plcSelC3.Visible = True
                mAskCntrFeed plcSelC3.Top + plcSelC3.Height     'ask contracts spots and feed spots selectiviy

                '4-28-11 option to show spot detail vs day is complete summary
                If tgSpf.sSystemType = "R" Then         'adjust for summary/detail question
                    ilTop = plcSelC10.Top + plcSelC10.Height
                Else
                    ilTop = plcSelC3.Top + plcSelC3.Height
                End If
                
                plcSelC8.Move 120, ilTop
                rbcSelC8(0).Caption = "Spot Detail"
                rbcSelC8(0).Move 0, 0, 1200
                rbcSelC8(1).Caption = "Day is Complete Summary"
                rbcSelC8(1).Move 1200, 0, 2760
                rbcSelC8(0).Value = True            'default to detail
                rbcSelC8(0).Visible = True          'detail
                rbcSelC8(1).Visible = True          'summary
                'using live log or day is complete can get summary version of Days Incomplete
                'otherwise, dont give user option to get the day is complete summary
                'If (Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) = USINGLIVELOG Or (tgSpf.sBActDayCompl = "Y") Then
                    plcSelC8.Visible = True
                    plcSelC9.Move 120, plcSelC8.Top + plcSelC8.Height
                    ckcTrans.Caption = "Day is Complete Discrepancies Only"
                    ckcTrans.Move 0, 0, 4000
                    plcSelC9.Visible = False
                'End If
                
                ckcAll.Caption = "All Vehicles"
                lbcSelection(0).Visible = True
                pbcOption.Visible = True
            ElseIf ilListIndex = 1 Then    'Missing ISCI
                frcOption.Enabled = True
                lacSelCFrom.Caption = "From Date"
                lacSelCTo.Caption = "To Date"
                lacSelCFrom.Visible = True
                edcSelCFrom.Visible = True
                edcSelCFrom.Width = 1170
                lacSelCTo.Visible = True
                edcSelCTo.Visible = True
                edcSelCTo.Width = 1170
                plcSelC1.Visible = False
                plcSelC2.Visible = False
                pbcSelC.Visible = True
                plcSelC3.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
                ckcSelC3(7).Caption = "+Fill"
                ckcSelC3(7).Move 750, 0, 720
                ckcSelC3(7).Visible = True
                ckcSelC3(8).Caption = "-Fill"
                ckcSelC3(8).Move 1470, 0, 720
                ckcSelC3(8).Visible = True
                smPlcSelC3P = "Include"
                plcSelC3.Visible = True
                mAskCntrFeed plcSelC3.Top + plcSelC3.Height 'ask contracts spots and feed spots selectiviy

                'the Missing ISCI report is the same as the Post Log Status report.
                'default to all spots, and never a Discrep only for the Posting status
                rbcSelC8(0).Value = True
                ckcTrans.Value = vbUnchecked

                ckcAll.Caption = "All Vehicles"
                lbcSelection(0).Visible = True
                pbcOption.Visible = True
            ElseIf ilListIndex = PL_LIVELOG Then       '12-8-05
                lbcSelection(0).Clear
                'mConvVehPop 0                   'populate conventional vehicles only
                mLiveLogVehiclesPop 0           '12-21-12 get only live log vehicles
                frcOption.Enabled = True
                lacSelCFrom.Caption = "From Date"
                lacSelCFrom.Visible = True
                lacSelCTo.Caption = "To Date"
                lacSelCTo.Visible = True
                edcSelCFrom.Visible = True     'From Date   Tab 43
                edcSelCTo.Visible = True       'To Date     Tab 47
                pbcSelC.Visible = True
                ckcAll.Caption = "All Vehicles"
                ckcAll.Visible = True
                lbcSelection(0).Visible = True
                pbcOption.Visible = True
            End If
        Case INVOICESJOB
            pbcOption.Visible = False
            plcSelC1.Visible = False
            plcSelC3.Visible = False
            plcSelC2.Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(5).Visible = False
            lbcSelection(6).Visible = False
            lbcSelection(3).Visible = False
            pbcSelC.Visible = False
            ckcAll.Visible = False
            Select Case lbcRptType.ListIndex
                Case INV_REGISTER  '0 =Invoice Register
                    pbcOption.Visible = True
                    lbcSelection(0).Visible = False
                    lacSelCFrom.Caption = "Start Date"
                    lacSelCTo.Caption = "End Date"
                    lacSelCFrom.Visible = True
                    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr
                    If Trim$(slStr) <> "" Then
                        edcSelCFrom.Text = gIncOneDay(slStr)
                    Else
                        edcSelCFrom.Text = ""
                    End If
                    edcSelCFrom.Visible = True
                    lacSelCTo.Visible = True
                    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
                    If Trim$(slStr) <> "" Then
                        edcSelCTo.Text = slStr
                    Else
                        edcSelCTo.Text = ""
                    End If
                    edcSelCTo.Visible = True
                    edcSelCFrom.Left = 960
                    lacSelCTo.Move edcSelCFrom.Left + edcSelCFrom.Width + 120, lacSelCFrom.Top
                    edcSelCTo.Move lacSelCTo.Left + 840, edcSelCFrom.Top
                    plcSelC1.Top = edcSelCTo.Top + edcSelCTo.Height + 30
                    plcSelC1.Height = 840   '630
                    'plcSelC1.Caption = "By"
                    smPlcSelC1P = "By"
                    rbcSelCSelect(0).Caption = "Invoice"
                    rbcSelCSelect(0).Left = 300
                    rbcSelCSelect(0).Width = 980

                    rbcSelCSelect(1).Caption = "Advertiser"
                    rbcSelCSelect(1).Left = 1320
                    rbcSelCSelect(1).Width = 1190

                    rbcSelCSelect(2).Caption = "Agency"
                    rbcSelCSelect(2).Left = 2565    '300
                    rbcSelCSelect(2).Top = 0    '195

                    rbcSelCSelect(3).Top = 195
                    rbcSelCSelect(3).Caption = "Salesperson"
                    rbcSelCSelect(3).Left = 300 '1320
                    rbcSelCSelect(3).Width = 1580

                    rbcSelCSelect(4).Top = 195  '390
                    rbcSelCSelect(4).Caption = "Bill Vehicle"
                    rbcSelCSelect(4).Left = 1760    '300
                    rbcSelCSelect(4).Width = 2000

                    rbcSelCSelect(5).Top = 195  '390
                    rbcSelCSelect(5).Caption = "Air Vehicle"
                    rbcSelCSelect(5).Left = 3060    '1800
                    rbcSelCSelect(5).Width = 2000

                    rbcSelCSelect(6).Top = 390  '585                  '10-29-99
                    rbcSelCSelect(6).Caption = "Office/Vehicle"
                    rbcSelCSelect(6).Left = 300
                    rbcSelCSelect(6).Width = 1920

                    rbcSelCSelect(7).Caption = "NTR"
                    rbcSelCSelect(7).Move 1860, 390

                    rbcSelCSelect(8).Caption = "Sales Source"   '10-18-02
                    rbcSelCSelect(8).Move 2565, 390, 1680

                    rbcSelCSelect(9).Move 300, 585, 1680
                    rbcSelCSelect(9).Caption = "Sales Origin"

                    rbcSelCSelect(0).Visible = True
                    rbcSelCSelect(1).Visible = True
                    rbcSelCSelect(2).Visible = True
                    rbcSelCSelect(3).Visible = True
                    rbcSelCSelect(4).Visible = True
                    rbcSelCSelect(5).Visible = True
                    rbcSelCSelect(6).Visible = True         '10-22-99
                    rbcSelCSelect(7).Visible = True         '9-16-02   NTR
                    rbcSelCSelect(8).Visible = True         '10-17-02 Sales Source
                    rbcSelCSelect(9).Visible = True
                    If rbcSelCSelect(0).Value Then
                        rbcSelCSelect_click 0   ', True
                    Else
                        rbcSelCSelect(0).Value = True
                    End If

                    plcSelC1.Visible = True
                    pbcSelC.Visible = True
                    'pbcOption.Visible = True
                Case 1  'View Export
                    'pbcSelC.Visible = False
                    mPopInvoiceExportFileNames
                    lbcSelection(0).Visible = True
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case INV_DISTRIBUTE                 'Billing Distribution
                    pbcOption.Visible = True
                    lbcSelection(3).Visible = True
                    ckcAll.Caption = "All Participants"
                    ckcAll.Visible = True
                    pbcSelC.Visible = True
                    lbcSelection(3).Visible = True

                    mAskStartEndDates
                    plcSelC3.Left = 120
                    plcSelC3.Top = edcSelCFrom.Top + edcSelCFrom.Height + 30
                    mInvAskTypes                    'Include I & A types, detail vs summary
                    
                    plcSelC9.Move 120, plcSelC2.Top + plcSelC2.Height + 60, 3600       '1-22-15 option to skip each vehicle; in addition to total participant
                    ckcTrans.Move 0, 0, 3600
                    ckcTrans.Caption = "Skip to a new page each vehicle"
                    ckcTrans.Visible = True
                    plcSelC9.Visible = True
                    
                    lacCheck.Caption = "Contract #"
                    lacCheck.Move 120, plcSelC9.Top + plcSelC9.Height + 60
                    edcCheck.Move 1200, plcSelC9.Top + plcSelC9.Height + 30
                    lacCheck.Visible = True
                    edcCheck.Visible = True
                Case INV_CREDITMEMO
                    mAskCreditMemo
                Case INV_SUMMARY                    '6-28-05
                    lacSelCTo.Move 120, 60, 600
                    edcSelCTo.Move 800, 0, 480
                    lacSelCTo1.Move 1400, 60, 600
                    edcSelCTo1.Move 2000, edcSelCTo.Top, 600
                    edcSelCTo.MaxLength = 3    'month in char or number (1-12)
                    edcSelCTo1.MaxLength = 4
                    lacSelCTo.Caption = "Month"
                    lacSelCTo1.Caption = "Year"
                    lacSelCTo.Visible = True
                    lacSelCTo1.Visible = True
                    edcSelCTo.Visible = True
                    edcSelCTo1.Visible = True

                    plcSelC1.Move 120, edcSelCTo.Top + edcSelCTo.Height + 60
                    rbcSelCSelect(0).Caption = "Advt"
                    rbcSelCSelect(0).Move 600, 0, 675
                    rbcSelCSelect(1).Caption = "Agency"
                    rbcSelCSelect(1).Move 1290, 0, 980
                    rbcSelCSelect(1).Visible = True
                    rbcSelCSelect(2).Caption = "Salesperson"
                    rbcSelCSelect(2).Move 2250, 0, 1350   '2220
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
                    If rbcSelCSelect(0).Value Then
                        lbcSelection(1).Visible = False
                        lbcSelection(2).Visible = False
                        lbcSelection(5).Visible = True
                        lbcSelection(0).Visible = True
                        ckcAll.Caption = "All Advertisers"
                        ckcAll.Visible = True
                    ElseIf rbcSelCSelect(1).Value Then
                        lbcSelection(0).Visible = False
                        lbcSelection(2).Visible = False
                        lbcSelection(5).Visible = False
                        lbcSelection(1).Visible = True
                        ckcAll.Caption = "All Agencies"
                        ckcAll.Visible = True
                    ElseIf rbcSelCSelect(2).Value Then
                        lbcSelection(0).Visible = False
                        lbcSelection(1).Visible = False
                        lbcSelection(5).Visible = False
                        lbcSelection(2).Visible = True
                        ckcAll.Caption = "All Salespeople"
                        ckcAll.Visible = True
                    End If
                    If rbcSelCSelect(0).Value = True Then
                        rbcSelCSelect_click 0
                    Else
                        rbcSelCSelect(0).Value = True
                    End If
                    rbcSelCInclude(0).Caption = "Detail"
                    rbcSelCInclude(1).Caption = "Summary"
                    rbcSelCInclude(0).Move 720, 0, 840
                    rbcSelCInclude(1).Move 1560, 0, 1200
                    If rbcSelCInclude(1).Value = True Then
                        rbcSelCInclude_Click 1
                    Else
                        rbcSelCInclude(1).Value = True
                    End If
                    rbcSelCInclude(0).Visible = True
                    rbcSelCInclude(1).Visible = True
                    smPlcSelC2P = "Show"
                    plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height
                    plcSelC2.Visible = True

                    plcSelC7.Move 120, plcSelC2.Top + plcSelC2.Height
                    ckcSelC7.Caption = "Include Invoice Adjustments (AN)"
                    ckcSelC7.Move 0, 0, 4200
                    ckcSelC7.Visible = True
                    ckcSelC7.Enabled = True
                    plcSelC7.Visible = True

                    lacCheck.Caption = "Contract #"
                    lacCheck.Move 120, plcSelC7.Top + plcSelC7.Height + 60
                    edcCheck.Move 1200, plcSelC7.Top + plcSelC7.Height + 30
                    lacCheck.Visible = True
                    edcCheck.Visible = True

                    pbcSelC.Visible = True
                    pbcOption.Visible = True
                Case INV_TAXREGISTER
                    mAskStartEndDates
                    mAskContract edcSelCFrom.Top + edcSelCFrom.Height, 120
                    plcSelC3.Move plcSelC3.Left, edcSelCTo1.Top + edcSelCTo1.Height + 60, 3705, 450
                    mAskTranTypes
                    ckcSelC3(1).Value = vbUnchecked     'default to all types off except invoices
                    ckcSelC3(2).Value = vbUnchecked
                    ckcSelC3(3).Value = vbUnchecked
                    plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height, 3600
                    ckcSelC5(0).Move 0, 0, 3600
                    ckcSelC5(0).Caption = "Include non-taxable transactions"
                    ckcSelC5(0).Value = vbChecked
                    plcSelC5.Visible = True
                    ckcSelC5(0).Visible = True
                    lbcSelection(6).Height = 3240
                    lbcSelection(6).Visible = True
                    ckcAll.Caption = "All Vehicles"
                    ckcAll.Visible = True

                    pbcSelC.Visible = True
                    pbcOption.Visible = True
                Case INV_RECONCILE          '11-30-07
                    ckcAll.Visible = False
                    lacSelCFrom.Caption = "Active Dates: Start"
                    lacSelCTo.Caption = "End"
                    lacSelCFrom.Visible = True
                    edcSelCFrom.Text = ""
                    edcSelCFrom.Visible = True
                    lacSelCTo.Visible = True
                    edcSelCTo.Text = ""
                    edcSelCTo.Visible = True
                    edcSelCFrom.Left = 1440
                    lacSelCFrom.Width = 1200
                    edcSelCFrom.Move 1320, 30, 1200      'from date text box
                    lacSelCTo.Move edcSelCFrom.Left + edcSelCFrom.Width + 120, lacSelCFrom.Top, 1200
                    edcSelCTo.Move lacSelCTo.Left + 480, edcSelCFrom.Top

                    mAskContract edcSelCTo.Top + edcSelCTo.Height + 60, 120
                     'Discrep only, showing only those months that don't balance from installment to billed
                    plcSelC5.Move 120, edcSelCTo1.Top + edcSelCTo1.Height + 60
                    ckcSelC5(0).Caption = "Discrepancy Only"
                    ckcSelC5(0).Top = 0
                    ckcSelC5(0).Left = 0
                    plcSelC5.Visible = True
                    ckcSelC5(0).Value = vbUnchecked   'True
                    ckcSelC5(0).Visible = True
                    pbcSelC.Visible = True
                    pbcOption.Visible = True
            End Select
        Case CHFCONVMENU
        Case GENERICBUTTON
        Case DALLASFEED
            Select Case lbcRptType.ListIndex
                Case 0  'Dump
                    lbcSelection(1).Visible = False
                    lbcSelection(0).Visible = True
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case 1  'Studio Log
                    lbcSelection(1).Visible = True
                    lbcSelection(0).Visible = False
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case 2  'Error Log
                    frcOption.Enabled = False
                    pbcOption.Visible = False
            End Select
        Case NYFEED
            Select Case lbcRptType.ListIndex
                Case 0  'Feed
                    pbcSelC.Visible = False
                    lbcSelection(0).Visible = True
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case 1  'Error Log
                    pbcSelC.Visible = False
                    frcOption.Enabled = False
                    pbcOption.Visible = False
                Case 2  'Suppression
                frcOption.Enabled = True
                pbcOption.Visible = False
                ckcAll.Visible = False
                pbcSelC.Visible = True
                edcSelCTo.Visible = False
                lacSelCTo.Visible = False
                edcSelCFrom.Visible = True
                edcSelCFrom.MaxLength = 10  '8  5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
                edcSelCFrom.Width = 1170
                lacSelCFrom.Width = 1500
                lacSelCFrom.Caption = "Active Date"
                lacSelCFrom.Visible = True
                plcSelC3.Visible = False
                plcSelC1.Visible = False
                plcSelC2.Visible = False
                rbcSelCSelect(0).Value = False
                rbcSelCInclude(0).Value = False
                Case 3  'Replacement
                frcOption.Enabled = True
                pbcOption.Visible = False
                ckcAll.Visible = False
                pbcSelC.Visible = True
                edcSelCTo.Visible = False
                lacSelCTo.Visible = False
                edcSelCFrom.Enabled = True
                edcSelCFrom.MaxLength = 10  '8  5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
                edcSelCFrom.Width = 1170
                lacSelCFrom.Width = 1500
                lacSelCFrom.Caption = "Active Date"
                lacSelCFrom.Visible = True
                edcSelCFrom.Visible = True
                plcSelC3.Visible = False
                plcSelC1.Visible = False
                plcSelC2.Visible = False
                rbcSelCSelect(0).Value = False
                rbcSelCInclude(0).Value = False
            End Select
        Case PHOENIXFEED
            Select Case lbcRptType.ListIndex
                Case 0  'Studio Log
                    lbcSelection(0).Visible = True
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case 1  'Error Log
                    frcOption.Enabled = False
                    pbcOption.Visible = False
            End Select
        Case CMMLCHG
            Select Case lbcRptType.ListIndex
                Case 0  'Commercial Change Export
                    lbcSelection(0).Visible = True
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case 1  'Commercial change
                    plcSelC1.Visible = False
                    plcSelC2.Visible = False
                    plcSelC3.Visible = False
                    lbcSelection(0).Visible = False
                    ckcAll.Visible = False
                    lacSelCFrom.Left = 120
                    lacSelCFrom1.Left = 2385
                    edcSelCFrom.Move 1340, edcSelCFrom.Top, 945
                    edcSelCFrom1.Move 2700, edcSelCFrom.Top, 945
                    lacSelCTo.Left = 120
                    edcSelCTo.Move 1340, edcSelCTo.Top, 945
                    lacSelCTo1.Left = 2385
                    edcSelCTo1.Move 2700, edcSelCTo1.Top, 945
                    edcSelCTo.MaxLength = 10    ' 8   5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
                    edcSelCTo1.MaxLength = 10   '8
                    edcSelCFrom.MaxLength = 10  '8
                    edcSelCFrom1.MaxLength = 10 '8
                    lacSelCFrom.Caption = "Created: From"
                    lacSelCFrom1.Caption = "To"
                    lacSelCFrom.Visible = True
                    lacSelCFrom1.Visible = True
                    lacSelCTo.Caption = "Aired: From"
                    lacSelCTo1.Caption = "To"
                    'plcSelC1.Caption = "Show"
                    smPlcSelC1P = "Show"
                    rbcSelCSelect(0).Caption = "New"
                    rbcSelCSelect(0).Left = 600
                    rbcSelCSelect(0).Width = 675
                    rbcSelCSelect(1).Caption = "All Spots"
                    rbcSelCSelect(1).Left = 1290
                    rbcSelCSelect(1).Width = 1000
                    rbcSelCSelect(1).Enabled = True
                    rbcSelCSelect(2).Visible = False
                    lacSelCTo.Visible = True
                    lacSelCTo1.Visible = True
                    edcSelCFrom.Visible = True
                    edcSelCFrom1.Visible = True
                    edcSelCTo.Visible = True
                    edcSelCTo1.Visible = True
                    pbcSelC.Visible = True
                    plcSelC1.Visible = True
                    pbcOption.Visible = True
            End Select
        Case EXPORTAFFSPOTS
            Select Case lbcRptType.ListIndex
                Case 0  'Export
                    lbcSelection(0).Visible = True
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case 1  'Error Log
                    frcOption.Enabled = False
                    pbcOption.Visible = False
            End Select
        Case BULKCOPY
            mBulkCopySelectivity
            
        End Select
    mSetCommands
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    Dim ilListIndex As Integer
    Dim slNameCode As String
    Dim ilHOState As Integer
    ReDim ilAASCodes(0 To 1) As Integer
    Dim slCntrStatus As String
    If Not imAllClicked Then
        ilListIndex = lbcRptType.ListIndex

        If igRptCallType = BUDGETSJOB Then
              If (ilListIndex = 0 Or ilListIndex = 1) And (Index <> 4 And Index <> 2) Then
                imSetAll = False
                ckcAll.Value = vbUnchecked  ' False
                imSetAll = True
              End If
        ElseIf igRptCallType = RATECARDSJOB Then
              If (ilListIndex = 0) Then
                If Index = 0 Then
                    imSetAll = False
                    ckcAll.Value = vbUnchecked  'False
                    imSetAll = True
                ElseIf Index = 11 Then
                    imSetAll = False
                    ckcAllRC.Value = vbUnchecked    'False
                    imSetAll = True
                End If
              End If
        ElseIf igRptCallType = COPYJOB Then
            If (ilListIndex = 1 Or ilListIndex = COPY_ROT) Then           'Copy Status by Advt  or Rot rpt

                If Index = 0 Then               'selecting on advt list
                    ckcAll.Enabled = True
                    imSetAll = False
                    ckcAll.Value = vbUnchecked  ' False
                    ckcAll.Visible = True
                    imSetAll = True
                    slCntrStatus = "HO"
                    ilHOState = 1
                    mCntrPop slCntrStatus, ilHOState
                    If tgSpf.sSystemType = "R" Then             'radio station has feed spots
                        slNameCode = "99999999|999-999||999||[Feed Spots]\0"
                        lbcSelection(3).AddItem slNameCode, 0
                        lbcSelection(5).AddItem "[Feed Spots]", 0 'Add ID to list box
                    End If
                    'selective adv, turn off generic contract & feed spot selectivty
                    plcSelC10.Visible = False
                    ckcSelC10(0).Value = vbChecked     'default contracts spots on
                    ckcSelC10(1).Value = vbChecked  'default feed spots on

                    If imTerminate Then
                        cmcCancel_Click
                        Exit Sub
                    End If
                    lbcSelection(5).Visible = True
                    lbcSelection(5).Move lbcSelection(0).Left + lbcSelection(0).Width + 60, lbcSelection(0).Top, lbcSelection(0).Width
                ElseIf Index = 2 Then           'clicked on vehicle list box
                    imSetAllGroup = False
                    ckcAllGroups.Value = vbUnchecked
                    imSetAllGroup = True
                ElseIf Index = 6 Then           'clicked on vehicle list box
                    imSetAllGroup = False
                    ckcAllGroups.Value = vbUnchecked
                    imSetAllGroup = True

                End If

                If ilListIndex = 1 Then             'copy status by advt
                    lbcSelection(0).Height = 2970
                    lbcSelection(5).Height = 2970
                Else            'make advt & cntr list boxes half the size.  Make room for vehicle selectivity
                    lbcSelection(0).Height = 1605
                    lbcSelection(5).Height = 1605
                    lbcSelection(2).Height = 1530

                End If

            Else
                If ilListIndex = COPY_SPLITROT Then       '1-30-09
                    If Index = 6 Then           'vehicles
                        imSetAllGroup = False
                        ckcAllGroups.Value = vbUnchecked
                        imSetAllGroup = True
                    Else
                        imSetAll = False
                        ckcAll.Value = vbUnchecked  'False
                        imSetAll = True
                    End If
                Else
                    imSetAll = False
                    ckcAll.Value = vbUnchecked  'False
                    imSetAll = True
                End If
            End If
        ElseIf igRptCallType = COLLECTIONSJOB Then
            If ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE Or ilListIndex = COLL_AGESLSP Or ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_CASH Then
                If Index = 7 Or Index = 9 Then                       'vehicle groups selection or Sales Source
                    ckcAllGroups.Enabled = True
                    imSetAllGroup = False
                    ckcAllGroups.Value = vbUnchecked    '9-12-02 False
                    ckcAllGroups.Visible = True
                    imSetAllGroup = True
                'TTP 9893
                 ElseIf Index = 1 And ilListIndex = COLL_CREDITSTATUS Then
                    imSetAll = False
                    ckcAll.Value = vbUnchecked  ' False
                    imSetAll = True
                 ElseIf Index = 0 And ilListIndex = COLL_CREDITSTATUS Then
                    imSetAllGroup = False
                    ckcAllGroups.Value = vbUnchecked  ' False
                    imSetAllGroup = True
                Else
                    ckcAll.Enabled = True
                    imSetAll = False
                    ckcAll.Value = vbUnchecked  '9-12-02 False
                    ckcAll.Visible = True
                    imSetAll = True
                End If
            ElseIf ilListIndex = COLL_DISTRIBUTE Or ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGEPRODUCER Or ilListIndex = COLL_AGESS Then
                If Index = 9 Then
                    imSetAllGroup = False
                    ckcAllGroups.Value = vbUnchecked
                    imSetAllGroup = True
                Else
                    imSetAll = False
                    ckcAll.Value = vbUnchecked  ' False
                    imSetAll = True
                End If
            ElseIf ilListIndex = COLL_CASHSUM Or ilListIndex = COLL_CASH Or ilListIndex = COLL_PAYHISTORY Or ilListIndex = COLL_POAPPLY Or ilListIndex = COLL_ACCTHIST Then
                imSetAll = False
                ckcAll.Value = vbUnchecked  ' False
                imSetAll = True
            ElseIf ilListIndex = COLL_STATEMENT Then
                imSetAll = False
                ckcAll.Value = vbUnchecked  ' False
                imSetAll = True
            End If
        ElseIf igRptCallType = INVOICESJOB Then
            If ilListIndex = INV_REGISTER Then
                'index = lbcselection list box
                If Index = 5 Or Index = 1 Or Index = 2 Or Index = 6 Or Index = 8 Or (Index = 9 And rbcSelCSelect(2).Value = False) Then         '5 = advt list box, 1 = agy, 2 = slsp, 6 = vehicle, 8 = ntr list, 9 = s/s
                    ckcAll.Enabled = True
                    imSetAll = False
                    ckcAll.Value = vbUnchecked
                    ckcAll.Visible = True
                    imSetAll = True
                Else          'secondary list box for one of the options
                    If rbcSelCSelect(2).Value = True Or rbcSelCSelect(3).Value = True Or rbcSelCSelect(1).Value = True Or rbcSelCSelect(5).Value = True Or rbcSelCSelect(8).Value = True Then     'rbcSelCSelect(2)& (3) = s/s, all others are vg list boxes
                        ckcAllGroups.Enabled = True
                        imSetAllGroup = False
                        ckcAllGroups.Value = vbUnchecked
                        ckcAllGroups.Visible = True
                        imSetAllGroup = True
                    End If
                End If
            ElseIf ilListIndex = INV_SUMMARY Or ilListIndex = INV_TAXREGISTER Then     '6-28-05
                If Index = 7 Then
                    ckcAllGroups.Enabled = True
                    imSetAllGroup = False
                    ckcAllGroups.Value = vbUnchecked
                    ckcAllGroups.Visible = True
                    imSetAllGroup = True
                Else
                    'ckcAllGroups.Visible = False
                    ckcAll.Enabled = True
                    imSetAll = False
                    ckcAll.Value = vbUnchecked
                    ckcAll.Visible = True
                    imSetAll = True

                End If
            ElseIf ilListIndex = INV_DISTRIBUTE Or ilListIndex = INV_CREDITMEMO Then
                imSetAll = False
                ckcAll.Value = vbUnchecked  ' False
                imSetAll = True
            End If
        ElseIf igRptCallType = USERLIST Then
            If ilListIndex = USER_ACTIVITY Then
                imSetAll = False
                ckcAll.Value = vbUnchecked  ' False
                imSetAll = True
            End If
        Else
            imSetAll = False
            ckcAll.Value = vbUnchecked  '9-12-02 False
            imSetAll = True
        End If
    Else
        imSetAll = False
        'ckcAll.value = vbUnchecked   'False
        'imSetAll = True
    End If
    'imSetAllGroup = False
    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub
Private Sub lbcSort_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked   'False
        imSetAll = True
    End If
    mSetCommands
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
Private Sub mAdvtPop(lbcSelection As control)
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(RptSelCreditStatus, lbcSelection, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(RptSelCreditStatus, lbcSelection, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
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
Private Sub mAgencyPop(lbcSelection As control)
'
'   mAgencyPop
'   Where:
'
    Dim ilRet As Integer
    'ilRet = gPopAgyBox(RptSelCreditStatus, lbcSelection, Traffic!lbcAgency)
    ilRet = gPopAgyBox(RptSelCreditStatus, lbcSelection, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gPopAgyBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mAgencyPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
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
Private Sub mAgyAdvtPop(lbcSelection As control)
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    ilRet = gPopAgyCollectBox(RptSelCreditStatus, "A", lbcSelection, lbcAgyAdvtCode)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgyAdvtPopErr
        gCPErrorMsg ilRet, "mAgyAdvtPop (gPopAgyCollectBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mAgyAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAirVehPop                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mAirVehPop()
    Dim ilRet As Integer
    'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHAIRING + ACTIVEVEH, lbcSelection(1), lbcAirNameCode)
    ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHAIRING + ACTIVEVEH, lbcSelection(1), tgAirNameCode(), sgAirNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAirVehPopErr
        gCPErrorMsg ilRet, "mAirVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mAirVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAllConvAirVehPop               *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mAllConvAirVehPop(ilIndex As Integer, ilUselbcVehicle As Integer, Optional blTestMerge As Boolean = False)
    Dim ilRet As Integer
    'ilRet = gPopUserVehicleBox(RptSelCreditStatus, 7, lbcSelection(0), Traffic!lbcVehicle)
    If Not blTestMerge Then
        If ilUselbcVehicle Then
            'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
            ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
        Else
            'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcSelection(ilIndex), lbcAirNameCode)
            ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcSelection(ilIndex), tgAirNameCode(), sgAirNameCodeTag)
        End If
    Else
        If ilUselbcVehicle Then
            'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
            ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHTESTLOGMERGE, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
        Else
            'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH, lbcSelection(ilIndex), lbcAirNameCode)
            ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + ACTIVEVEH + VEHTESTLOGMERGE, lbcSelection(ilIndex), tgAirNameCode(), sgAirNameCodeTag)
        End If
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAllConvAirVehPopErr
        gCPErrorMsg ilRet, "mAllConvAirVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mAllConvAirVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAllVehExLogPop                 *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mAllVehExLogPop(ilIndex As Integer)
    Dim ilRet As Integer
    'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAllVehExLogPopErr
        gCPErrorMsg ilRet, "mAllVehExLogPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mAllVehExLogPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAllVehPop                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mAllVehPop()
    Dim ilRet As Integer
    Dim llVehicles As Long
    'llVehicles = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHLOGVEHICLE + ACTIVEVEH + DORMANTVEH
    'ilRet = gPopUserVehicleBox(RptSelCreditStatus, &H7FFF, lbcSelection(0), Traffic!lbcVehicle)
    llVehicles = VEHALLTYPES + DORMANTVEH + ACTIVEVEH
    'llVehicles = &H7FFFFFFF And (Not VEHBYPASSWEGENER_OLA)
    ilRet = gPopUserVehicleBox(RptSelCreditStatus, llVehicles, lbcSelection(0), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAllVehPopErr
        gCPErrorMsg ilRet, "mAllVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mAllVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
Private Sub mPartVehPop()
    Dim ilRet As Integer
    Dim llVehicles As Long
    llVehicles = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + VEHSPORT + ACTIVEVEH + DORMANTVEH
    ilRet = gPopUserVehicleBox(RptSelCreditStatus, llVehicles, lbcSelection(0), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPartVehPopErr
        gCPErrorMsg ilRet, "mPartVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mPartVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'**************************************************************************
'
'                   mCashTrMercProm - Ask selectivity:
'                               Cash, Trade, Merchandise, Promotions
'
'                   Created:  4/97 D.Hosaka
'           <input> ilTop - top location of options within container
'                   ilAskHardCost - hide or allow hard cost option
'**************************************************************************
Private Sub mAskCashTrMercProm(ilTop As Integer, ilAskhardCost As Integer)
    plcSelC2.Top = ilTop    '     900
    plcSelC2.Height = 450
    'plcSelC2.Caption = "Only"
    smPlcSelC2P = "Only"
    rbcSelCInclude(0).Caption = "Cash"
    rbcSelCInclude(0).Width = 720
    rbcSelCInclude(0).Left = 450
    rbcSelCInclude(1).Caption = "Trade"
    rbcSelCInclude(1).Left = 1230
    rbcSelCInclude(0).Enabled = True
    If rbcSelCInclude(0).Value Then
        rbcSelCInclude_Click 0  ', True
    Else
        rbcSelCInclude(0).Value = True
    End If
    rbcSelCInclude(2).Caption = "Merchandise"
    'rbcSelCInclude(2).Move 450, rbcSelCInclude(0).Top + 195, 1380
    rbcSelCInclude(2).Move 2070, rbcSelCInclude(0).Top, 1380
    rbcSelCInclude(2).Visible = True
    rbcSelCInclude(3).Caption = "Promotions"
    'rbcSelCInclude(3).Move 1890, rbcSelCInclude(0).Top + 195, 1320
    rbcSelCInclude(3).Move 450, rbcSelCInclude(0).Top + 195, 1320
    rbcSelCInclude(4).Caption = "Hard Cost"
    rbcSelCInclude(4).Move 1890, rbcSelCInclude(0).Top + 195, 1200
    rbcSelCInclude(3).Visible = True
    If ilAskhardCost Then
        rbcSelCInclude(4).Visible = True       'hard cost
    Else
        rbcSelCInclude(4).Visible = False
    End If
End Sub
'
'***********************************************************************
'
'               mAskEntryDates - Ask STart Date,  End Date
'                                   Default dates to end of
'                                   previous reconciled and currnet
'                                   reconciling period
'***********************************************************************
Private Sub mAskEntryDates()
Dim slStr As String

        lacSelCFrom.Caption = "Entry Dates-Start"
        lacSelCTo.Caption = "End"
        lacSelCFrom.Visible = True
        lacSelCFrom.Width = 2400
        gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr
        If Trim$(slStr) <> "" Then
            edcSelCFrom.Text = gIncOneDay(slStr)
        Else
            edcSelCFrom.Text = ""
        End If
        edcSelCFrom.Visible = True
        lacSelCTo.Visible = True
        gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
        If Trim$(slStr) <> "" Then
            edcSelCTo.Text = slStr
        Else
            edcSelCTo.Text = ""
        End If
        edcSelCTo.Visible = True
        edcSelCFrom.Left = 1680
        lacSelCTo.Move edcSelCFrom.Left + edcSelCFrom.Width, lacSelCFrom.Top
        edcSelCTo.Move lacSelCTo.Left + 480, edcSelCFrom.Top
        edcSelCFrom.Width = 960
        edcSelCTo.Width = 960

End Sub
'
'***********************************************************************
'
'               mAskStartEndDates - Ask STart Date,  End Date
'                                   Default dates to end of
'                                   previous reconciled and currnet
'                                   reconciling period
'***********************************************************************
Private Sub mAskStartEndDates()
Dim slStr As String

        lacSelCFrom.Caption = "Start Date"
        lacSelCTo.Caption = "End Date"
        lacSelCFrom.Visible = True
        gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr
        If Trim$(slStr) <> "" Then
            edcSelCFrom.Text = gIncOneDay(slStr)
        Else
            edcSelCFrom.Text = ""
        End If
        edcSelCFrom.Visible = True
        lacSelCTo.Visible = True
        gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
        If Trim$(slStr) <> "" Then
            edcSelCTo.Text = slStr
        Else
            edcSelCTo.Text = ""
        End If
        edcSelCTo.Visible = True
        edcSelCFrom.Left = 1020
        lacSelCTo.Move edcSelCFrom.Left + edcSelCFrom.Width + 120, lacSelCFrom.Top
        edcSelCTo.Move lacSelCTo.Left + 840, edcSelCFrom.Top
        edcSelCFrom.Width = 960
        edcSelCTo.Width = 960
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBudgetPop                      *
'*                                                     *
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
    'ilRet = gPopVehBudgetBox(RptSelCreditStatus, 0, 1, lbcSelection(4), lbcBudgetCode)
    ilRet = gPopVehBudgetBox(RptSelCreditStatus, 2, 0, 1, lbcSelection(4), tgRptSelCreditStatusBudgetCode(), sgRptSelCreditStatusBudgetCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mBudgetPopErr
        gCPErrorMsg ilRet, "mBudgetPopErr (gPopVehBudgetBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mBudgetPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mChfConvPop                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mChfConvPop()
'
'   mChfConvPop
'   Where:
'       lbcSort (O)- contains list of converted file names
'
    Dim hlIcf As Integer        'Slf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlIcf As ICF
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilOffSet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slSortDate As String
    Dim slSortTime As String
    Dim ilIndex As Integer
    Dim slName As String
    Dim ilPos As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'type field record
    Dim llDate As Long

    lbcSort.Clear
    hlIcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlIcf, "", sgDBPath & "Icf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlIcf)
        btrDestroy hlIcf
        Exit Sub
    End If
    ilRecLen = Len(tlIcf) 'btrRecordLength(hlAdf)  'Get and save record length
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hlIcf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlIcf, tlIcf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlIcf)
        btrDestroy hlIcf
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlIcf)
            btrDestroy hlIcf
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hlIcf, llNoRec, -1, "UC", "ICF", "") 'Set extract limits (all records)
    slDate = Format$(gNow(), "m/d/yy")
    llDate = gDateValue(slDate) - 30 * 6    'view only last  6 months
    slDate = Format$(llDate, "m/d/yy")
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = 0
    ilRet = btrExtAddLogicConst(hlIcf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    tlCharTypeBuff.sType = "0"
    ilOffSet = 10 'gFieldOffset("Icf", "IcfType")
    ilRet = btrExtAddLogicConst(hlIcf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilRet = btrExtAddField(hlIcf, 0, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlIcf)
        btrDestroy hlIcf
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hlIcf)    'Extract record
    ilRet = btrExtGetNext(hlIcf, tlIcf, ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            ilRet = btrClose(hlIcf)
            btrDestroy hlIcf
            Exit Sub
        End If
    End If
    ilIndex = -1
    ilRecLen = Len(tlIcf)       'for first get extend oper need to reset recd length
                                'the very first extend get next seems to set it to 0
    'ilRet = btrExtGetFirst(hlIcf, tlIcf, ilRecLen, llRecPos)
    Do While ilRet = BTRV_ERR_REJECT_COUNT
        ilRet = btrExtGetNext(hlIcf, tlIcf, ilRecLen, llRecPos)
    Loop
    Do While ilRet = BTRV_ERR_NONE
        gUnpackDate tlIcf.iDate(0), tlIcf.iDate(1), slDate
        slSortDate = Trim$(str$(9999999 - gDateValue(slDate)))
        Do While Len(slSortDate) < 7
            slSortDate = "0" & slSortDate
        Loop
        gUnpackTime tlIcf.iTime(0), tlIcf.iTime(1), "A", "1", slTime
        slSortTime = Trim$(str$(9999999 - CLng(gTimeToCurrency(slTime, False))))
        Do While Len(slSortTime) < 7
            slSortTime = "0" & slSortTime
        Loop
        lbcSort.AddItem slSortDate & "|" & slSortTime & "\" & Trim$(tlIcf.sAdvtName) & " on " & slDate & " at " & slTime
        ilRet = btrExtGetNext(hlIcf, tlIcf, ilRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlIcf, tlIcf, ilRecLen, llRecPos)
        Loop
    Loop
    For ilLoop = 0 To lbcSort.ListCount - 1 Step 1
        slName = lbcSort.List(ilLoop)
        ilRet = gParseItem(slName, 2, "\", slName)    'Get application name
        lbcSelection(0).AddItem slName
    Next ilLoop
    For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
        slName = lbcSelection(0).List(ilLoop)
        ilPos = InStr(slName, " On")
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1)
            If StrComp(Trim$(smChfConvName), Trim$(slName), 1) = 0 Then
                lbcSort.Selected(ilLoop) = True
                Exit For
            End If
        End If
    Next ilLoop
    ilRet = btrClose(hlIcf)
    btrDestroy hlIcf
    Exit Sub
End Sub
'*******************************************************
Private Sub mCntrPop(slCntrStatus As String, ilHOState As Integer)
'
'   mCntrPop
'   Where:
'       slcntrStatus(I)- O; H; W; C; I; D or blank for all
'       ilHOState(I) - 1 only get cnt (w/o revision)
'                      2 combo - get latest orders includ revisions
'                      3 everything - revision & orders
'                      4 only cnts which are revisions (internally WCI
'
'   5-13-05 show cntr # and product name instead of contr # and advt, since the
'               advt was the first selection
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
    lbcSelection(5).Clear           'init contract list box
    lbcSelection(3).Clear
    For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
        If lbcSelection(0).Selected(ilLoop) Then
            sgMultiCntrCodeTag = ""                 'init time stamp so all cnt get populated for all advt selected
            'lbcMultiCntrCode.Clear
            ReDim tgMultiCntrCode(0 To 0) As SORTCODE
            'lbcMultiCntr.Clear
            lbcSelection(4).Clear
            slNameCode = tgAdvertiser(ilLoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            'ilCurrent = 1   '0=Current; 1=All
            'ilFilter = Val(slCode)   'by contract #; -101=by advertiser
            'ilVehCode = -1  'All vehicles
            'ilRet = gPopCntrBox(RptSelCreditStatus, ilCntrType, ilFilter, ilCurrent, 0, ilVehCode, lbcMultiCntr, lbcMultiCntrCode, True, False, False, False)
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
            'ilShow = 5                  'show # and advt name
            ilShow = 7                  '5-13-05 show cntr # and product name
            ilCurrent = 1
            ilAdfCode = Val(slCode)
            'load up list box with contracts with matching adv
            'ilRet = gPopCntrForAASBox(RptSelCreditStatus, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, lbcMultiCntrCode)
            ilRet = gPopCntrForAASBox(RptSelCreditStatus, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcSelection(4), tgMultiCntrCode(), sgMultiCntrCodeTag)
            If ilRet <> CP_MSG_NOPOPREQ Then
                On Error GoTo mCntrPopErr
                gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", RptSelCreditStatus
                On Error GoTo 0
            End If
            For ilIndex = 0 To UBound(tgMultiCntrCode) - 1 Step 1 'lbcMultiCntrCode.ListCount - 1 Step 1
                slName = Trim$(tgMultiCntrCode(ilIndex).sKey)  'lbcMultiCntrCode.List(ilIndex)
                If Not gOkAddStrToListBox(slName, llLen, True) Then
                    ilErr = True
                    Exit For
                End If
                'lbcCntrCode.AddItem slName  'lbcMultiCntrCode.List(ilIndex)
               lbcSelection(3).AddItem slName
            Next ilIndex

            If ilErr Then
                Exit For
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To lbcSelection(3).ListCount - 1 Step 1
        slNameCode = lbcSelection(3).List(ilLoop)
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

        If Trim$(slCode) = "[Feed Spots]" Then
            slShow = slCode
        Else
            slShow = slShow & " " & slCode
        End If

        lbcSelection(5).AddItem Trim$(slShow)  'Add ID to list box
    Next ilLoop
    Screen.MousePointer = vbDefault
    Exit Sub
mCntrPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvAirVehPop                  *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mConvAirVehPop()
    Dim ilRet As Integer
    'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcSelection(3), Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcSelection(3), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mConvAirVehPopErr
        gCPErrorMsg ilRet, "mConvAirVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mConvAirVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEvtNmPop                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate event name list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mEvtNmPop()
'
'   mEvtNmPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopEvtNmByTypeBox(RptSelCreditStatus, True, True, lbcSelection(0), lbcNameCode)
    ilRet = gPopEvtNmByTypeBox(RptSelCreditStatus, True, True, lbcSelection(0), tgRptSelCreditStatusNameCode(), sgRptSelCreditStatusNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mEvtNmPopErr
        gCPErrorMsg ilRet, "mEvtNmPop (gPopEvtNmBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mEvtNmPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFeedPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Feed list             *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mFeedPop(lbcCtrl As control)
'
'   mFeedPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopMnfPlusFieldsBox(RptSelCreditStatus, lbcCtrl, lbcNameCode, "NNS")
    ilRet = gPopMnfPlusFieldsBox(RptSelCreditStatus, lbcCtrl, tgRptSelCreditStatusNameCode(), sgRptSelCreditStatusNameCodeTag, "NNS")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mFeedPopErr
        gCPErrorMsg ilRet, "mFeedPop (gPopMnfPlusFieldsBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mFeedPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
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
    RptSelCreditStatus.Caption = smSelectedRptName & " Report"
    frcOption.Caption = smSelectedRptName & " Selection"
'VB6**    hdJob = rpcRpt.hJob
    ilMultiTable = True
    'dummy = LlSetOption(hdJob, LL_OPTION_HELPAVAILABLE, False)
'VB6**    ilDummy = LlSetOption(hdJob, LL_OPTION_SORTVARIABLES, True)
'VB6**    ilDummy = LlSetOption(hdJob, LL_OPTION_ONLYONETABLE, Not ilMultiTable)

    imAllClicked = False
    imSetAll = True
    imAllGroupClicked = False
    imSetAllGroup = True
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
    lacSelCFrom.Visible = False
    lacSelCFrom1.Visible = False
    lacSelCTo.Visible = False
    lacSelCTo1.Visible = False
    edcSelCFrom.Visible = False
    edcSelCFrom1.Visible = False
    edcSelCTo.Visible = False
    edcSelCTo1.Visible = False
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

    plcSelC4.Move 120, 360
    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3000
    lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    pbcSelA.Move 90, 255, 4515, 3360
    pbcSelB.Move 90, 255, 4515, 3360
    pbcSelC.Move 90, 255, 4515, 3720

    '3/30/13
    ilRet = gVffRead()

    gCenterStdAlone RptSelCreditStatus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*
'*          6/19/98 Added CopyPlayList by vehicle back
'*          in as a separate report
'*                                                     *
'*******************************************************
Private Sub mInitReport()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    Dim ilRet As Integer
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

    '10-19-01
    'Dan for rollback to 8.5 on copybook 5/11/09 removed 9-03-09
'    If smSelectedRptName = "Copy Book" Then
'        gRollPopExportTypes cbcFileType
'    Else
'        gPopExportTypes cbcFileType
'    End If
    gPopExportTypes cbcFileType
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
    sgPhoneImage = mkcPhone.Text
    lbcRptType.Clear
    If tgSpf.sRUseCorpCal = "Y" Then       'if Using corp cal, dfault it; otherwise disable it
        ilRet = gObtainCorpCal()                  'read the entire corporate calendar into memory
    End If
    ilRet = gObtainAdvt()           '3-23-03
    If Not ilRet Then ' this is set by mVehPop if error occurs
        imTerminate = True
        Exit Sub
    End If
    Select Case igRptCallType
        
        Case COLLECTIONSJOB
            mAdvtPop lbcSelection(0)    'Called to initialize Traffic!Advertiser required be mCntrPop
            If imTerminate Then
                Exit Sub
            End If
            mAgencyPop lbcSelection(1)
            If imTerminate Then
                Exit Sub
            End If
            mAgyAdvtPop lbcSelection(2)    'Called to initialize agy and direct advertiser (statements)
            If imTerminate Then
                Exit Sub
            End If
            mSPersonPop lbcSelection(5)
            If imTerminate Then
                Exit Sub
            End If
           ' mMnfPop "H", RptSelCreditStatus!lbcSelection(3), tgVehicle(), sgVehicleTag    'Traffic!lbcVehicle         'owners groups
            '5-22-02 only show the participants, not other vehicle groups
            ilRet = gPopMnfPlusFieldsBox(RptSelCreditStatus, lbcSelection(3), tgVehicle(), sgVehicleTag, "H1")


            mSellConvVVPkgPop 6, False, True        '3/30/99 use dormant vehicles
            If imTerminate Then
                Exit Sub
            End If

            '1-20-06 Sales Source
            ilRet = gPopMnfPlusFieldsBox(RptSelCreditStatus, RptSelCreditStatus!lbcSelection(9), tgMNFCodeRpt(), sgMNFCodeTagRpt, "S")

            lbcRptType.AddItem "Cash Receipts or Usage", 0
            lbcRptType.AddItem "Ageing by Payee", 1
            lbcRptType.AddItem "Ageing by Salesperson", 2
            lbcRptType.AddItem "Ageing by Vehicle", 3
            lbcRptType.AddItem "Delinquent Accounts", 4
            lbcRptType.AddItem "Statements", 5
            lbcRptType.AddItem "Cash Payment or Usage History", 6
            lbcRptType.AddItem "Advertiser and Agency Credit Status", 7
            lbcRptType.AddItem "Cash Distribution", COLL_DISTRIBUTE
            lbcRptType.AddItem "Cash Summary", COLL_CASHSUM
            lbcRptType.AddItem "Account History", COLL_ACCTHIST
            lbcRptType.AddItem "Merchandising/Promotions", COLL_MERCHANT
            lbcRptType.AddItem "Merchandising/Promotions Recap"
            lbcRptType.AddItem "Ageing by Participant"
            lbcRptType.AddItem "Ageing by Sales source"
            lbcRptType.AddItem "Ageing by Producer"
            lbcRptType.AddItem "On-Account Cash Applied", COLL_POAPPLY
            frcOption.Enabled = True
            rbcSelCSelect(2).Visible = False
            lbcSelection(0).Move 15, ckcAll.Height + 30
            lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
            lbcSelection(2).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
            lbcSelection(6).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
            lbcSelection(5).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height       'owners
            lbcSelection(3).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
            lbcSelection(10).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height      '4-27-07 media codes

            pbcSelC.Visible = True
            pbcOption.Visible = True
        
        Case GENERICBUTTON
            cmcList.Visible = False
            lbcRptType.Visible = False
            lacFromA.Caption = "Report File"
            lacFromA.Width = 1065
            lacFromA.Left = 120
            edcSelA.Move 1095, 0
            edcSelA.MaxLength = 12
            pbcSelA.Visible = True
            plcSel1.Visible = False
            plcSel2.Visible = False
            frcOption.Enabled = True
            pbcOption.Visible = False
        Case COLLECTIONSJOB
            'TTP 9893
            mAdvtPop lbcSelection(0)    'Called to initialize Traffic!Advertiser required be mCntrPop
            If imTerminate Then
                Exit Sub
            End If
            mAgencyPop lbcSelection(1)
            If imTerminate Then
                Exit Sub
            End If
            
            'pbcOption.Visible = False
            pbcOption.Visible = True
            ckcAll.Value = 1
            ckcAllGroups.Value = 1
        
    End Select
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
'    gCenterModalForm RptSelCreditStatus
End Sub
'***************************************************************************'
'
'                       Ask Invoice types (Invoices vs Adjustments)
'                       for the Invoice Register & Billing Distribution
'                       reports.
'
'***************************************************************************
'
Private Sub mInvAskTypes()
    'plcSelC3.Caption = "Include"
    smPlcSelC3P = "Include"
    ckcSelC3(0).Left = 765
    ckcSelC3(0).Width = 1530
    ckcSelC3(0).Caption = "Invoices (IN)"
    If ckcSelC3(0).Value = vbChecked Then
        ckcSelC3_click 0
    Else
        ckcSelC3(0).Value = vbChecked   'True
    End If
    ckcSelC3(1).Left = 2175
    ckcSelC3(1).Width = 1880

    ckcSelC3(1).Caption = "Adjustments(AN)"
    If ckcSelC3(1).Value = vbChecked Then
        ckcSelC3_click 1
    Else
        ckcSelC3(1).Value = vbChecked   'True
    End If
    ckcSelC3(2).Move 765, ckcSelC3(0).Top + 210, 1530
    ckcSelC3(2).Caption = "History (HI)"
    If ckcSelC3(2).Value = vbChecked Then
        ckcSelC3_click 2
    Else
        ckcSelC3(2).Value = vbChecked   'True
    End If
    plcSelC3.Height = 450
    ckcSelC3(0).Visible = True
    ckcSelC3(1).Visible = True
    ckcSelC3(2).Visible = True
    ckcSelC3(3).Visible = False

    'plcSelC2.Caption = "By"
    smPlcSelC2P = "By"
    If lbcRptType.ListIndex = INV_REGISTER And rbcSelCSelect(0).Value Then     'Inv register by invoice #?
        rbcSelCInclude(0).Caption = "Airing Vehicle"
        rbcSelCInclude(0).Left = 300   '300    '860
        rbcSelCInclude(0).Width = 1460
        rbcSelCInclude(1).Caption = "Billing Vehicle"
        rbcSelCInclude(1).Left = 1740       '2340 '1140
        rbcSelCInclude(1).Width = 1520
        rbcSelCInclude(2).Caption = "Summary"
        rbcSelCInclude(2).Left = 3180
        rbcSelCInclude(2).Width = 1160
        rbcSelCInclude(2).Visible = True
        plcSelC2.Width = 5580
    Else
        rbcSelCInclude(0).Caption = "Detail"
        rbcSelCInclude(0).Left = 300    '900   '300
        rbcSelCInclude(0).Width = 840
        rbcSelCInclude(1).Caption = "Summary"
        rbcSelCInclude(1).Left = 1140   '1740   '1140
        rbcSelCInclude(1).Width = 1160
        rbcSelCInclude(2).Visible = False
        plcSelC2.Width = 2820
    End If
    rbcSelCInclude(0).Value = True
    rbcSelCInclude(0).Enabled = True
    rbcSelCInclude(1).Enabled = True
    rbcSelCInclude(0).Visible = True
    rbcSelCInclude(1).Visible = True
    plcSelC2.Move plcSelC3.Left, plcSelC3.Top + plcSelC3.Height
    plcSelC2.Visible = True
    plcSelC3.Visible = True
End Sub
'
'                   mMnfPop - Populate list box with MNF records
'                           slType = Mnf type to match (i.e. "H", "A")
'                           lbcLocal  - local list box to fill
'                           lbcMster - master list box with codes
'                   Created: DH 9/12/96
'
Private Sub mMnfPop(slType As String, lbcLocal As control, tlSortCode() As SORTCODE, slSortCodeTag As String) 'lbcMster As Control)
ReDim ilfilter(0) As Integer
ReDim slFilter(0) As String
ReDim ilOffSet(0) As Integer
Dim ilRet As Integer
    ilfilter(0) = CHARFILTER
    slFilter(0) = slType
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType")

    'ilRet = gIMoveListBox(RptSelCreditStatus, lbcLocal, lbcMster, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(RptSelCreditStatus, lbcLocal, tlSortCode(), slSortCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMnfPopErr
        gCPErrorMsg ilRet, "mMnfPop (gImoveListBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mMnfPopErr:
    On Error GoTo 0
    Unload RptSelCreditStatus
    Set RptSelCreditStatus = Nothing   'Remove data segment
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

    imAutoReport = False                    '2-24-05 assume not running an automatically run report, (report list used)
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

        '2-24-05            auto run the adv & agy credit status
        If StrComp(sgCallAppName, "ShoCrdit", 1) = 0 Then
            imAutoReport = True
        End If

    '    imShowHelpmsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpmsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelCreditStatus, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    'igRptCallType = Val(slStr)
    If igRptCallType <> GENERICBUTTON Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)
    Else
        igRptType = 0
    End If


    'D.S. 10/01 Reports for this module with their associated constants and values
        'AGENCIESLIST
            '"Agency Summary"                    =  0
            '"Agency Credit"                     =  1

        'ADVERTISERSLIST
            '"Advertiser Credit"                 =  1
            '"Advertiser Summary"                =  0

        'BUDGETSJOB
            '"Budgets"                           =  0
            '"Budget Comparisons"                =  1

        'BULKCOPY
            '"Bulk Copy Feed"                    =  0
            '"Bulk Copy Cross Reference"         =  1
            '"Affiliate Bulk Feed by Cart"       =  2
            '"Affiliate Bulk Feed by Vehicle"    =  3
            '"Affiliate Bulk Feed by Date"       =  4
            '"Affiliate Bulk Feed by Advertiser" =  5

        'CMMLCHG
            '"Commercial Change Export"          =  0
            '"Commercial Change"                 =  1

        'COLLECTIONSJOB
           '"Cash Receipts or Usage"                COLL_CASH         =  0
           '"Ageing by Payee"                       COLL_AGEPAYEE     =  1
           '"Ageing by Salesperson"                 COLL_AGESLSP      =  2
           '"Ageing by Vehicle"                     COLL_AGEVEHICLE   =  3
           '"Delinquent Accounts"                   COLL_DELINQUENT   =  4
           '"Statements"                            COLL_STATEMENT    =  5
           '"Payment or Usage History"              COLL_PAYHISTORY   =  6
           '"Advertiser and Agency Credit Status"   COLL_CREDITSTATUS =  7
           '"Cash Distribution"                     COLL_DISTRIBUTE   =  8
           '"Cash Summary"                          COLL_CASHSUM      =  9
           '"Account History"                       COLL_ACCTHIST     = 10
           '"Merchandising/Promotions"              COLL_MERCHANT     = 11
           '"Merchandising/Promotions Recap"        COLL_MERCHRECAP   = 12
           '"Ageing by Participant"
           '"Ageing by Sales source"                COLL_AGESS        = 14
           '"Ageing by Producer"                    COLL_AGEPRODUCER  = 15

        'COPYJOB
            '"Copy Status by Date"               =  0
            '"Copy Status by Advertiser"         =  1
            '"Contracts Missing Copy"            =  2
            '"Copy Rotations by Advertiser"      =  3
            '"Copy Inventory by Number"          =  4
            '"Copy Inventory by ISCI"            =  5
            '"Copy Inventory by Advertiser"      =  6
            '"Copy Inventory by Start Date"      =  7
            '"Copy Inventory by Expiration Date" =  8
            '"Copy Inventory by Purge Date"      =  9
            '"Copy Inventory by Entry Date"      = 10
            '"Copy Play List by ISCI"            = 11
            '"Unapproved Copy"                   = 12
            '"Copy Play List by Vehicle"         = 13
            '"Copy Play List by Advertiser"      = 14
            '"Copy Regions"                      = 15

        'DALLASFEED
            '"Dallas Feed"                       =  0
            '"Dallas Studio Log"                 =  1
            '"Dallas Error Log"                  =  2

        'EXPORTAFFSPOTS
            '"Affiliate Spots Export"            =  0
            '"Affiliate Spots Error Log"         =  1

        'INVOICESJOB
            '"Invoice Register"                  =  0
            '"View Invoice Export"               =  1
            '"Billing Distribution"                 INV_DISTRIBUTE

        'NYFEED
            '"New York Feed"                     =  0
            '"New York Error Log"                =  1
            '"Blackout Suppression"              =  2
            '"Blackout Replacement"              =  3

        'PHOENIXFEED
            '"Phoenix Studio Log"                =  0
            '"Phoenix Error Log"                 =  1

        'POSTLOGSJOB
            '"Log Posting Status"                =  0
            '"Missing ISCI Codes"                =  1

        'PROGRAMMINGJOB
            '"Program Libraries"                 =  0   igRptType = 3
            '"Selling to Airing Vehicles"        =  0
            '"Airing to Selling Vehicles"        =  1
            '"Vehicle Avail Conflicts"           =  2
            '"Delivery by Vehicle"               =  3
            '"Delivery by Feed"                  =  4
            '"Engineering by Vehicle"            =  5
            '"Engineering by Feed"               =  6

        'RATECARDSJOB
            '"Rate Card"                            RC_RCITEMS
            '"Dayparts"                             RC_DAYPARTS

        'SALESPEOPLELIST
            '"Salespeople Summary"               =  0
            '"Salespeople Budgets"               =  1

        'VEHICLESLIST
            '"Vehicle Summary"                   =  0
            '"Vehicle Options"                   =  1
            '"Virtual Vehicles"                  =  2


    'If igStdAloneMode Then
    '    smSelectedRptName = "Rate Card"
    '    igRptCallType = RATECARDSJOB
    '    igRptType = 1 'Log     '0   'Summary '3 Program  '1  links
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
        If (igRptType = 0) Or (igRptType = 1) Or (igRptType = 2) Then
            igStdAloneMode = True 'Switch from/to stand alone mode-No DDE
        End If
    End If
    If igRptCallType = CHFCONVMENU Then
        ilRet = gParseItem(slCommand, 5, "\", smChfConvName)
        ilRet = gParseItem(slCommand, 6, "\", smChfConvDate)
        ilRet = gParseItem(slCommand, 7, "\", smChfConvTime)
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopAffSpotsFileNames           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain files with *.AFE        *
'*                                                     *
'*******************************************************
Private Sub mPopAffSpotsFileNames()
    Dim ilLoop As Integer
    Dim slStr As String
    lbcFileName.Path = Left$(sgExportPath, Len(sgExportPath) - 1)
    lbcFileName.Pattern = "*.AF?"
    'Move name to lbcSelection(0)-Dump and lbcSelection(1)-Log
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        'All names moved into lbcSelection(0)
        slStr = lbcFileName.List(ilLoop)
        lbcSelection(0).AddItem slStr
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopBulkCopyFileNames           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain files with *.Txt        *
'*                                                     *
'*******************************************************
Private Sub mPopBulkCopyFileNames()
    Dim ilLoop As Integer
    Dim slStr As String
    lbcFileName.Path = Left$(sgExportPath, Len(sgExportPath) - 1)
    lbcFileName.Pattern = "*.MSG"
    'Move name to lbcSelection(0)-Dump and lbcSelection(1)-Log
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        slStr = lbcFileName.List(ilLoop)
        'If (UCase$(Mid$(slStr, 1, 1)) >= "A") And (UCase$(Mid$(slStr, 1, 1)) <= "Z") And (UCase$(Mid$(slStr, 3, 1)) >= "0") And (UCase$(Mid$(slStr, 3, 1)) <= "9") Then
        If (UCase$(Mid$(slStr, 3, 1)) >= "0") And (UCase$(Mid$(slStr, 3, 1)) <= "9") Then
            lbcSelection(0).AddItem slStr
        End If
    Next ilLoop
    lbcFileName.Pattern = "X???????.ASC"
    'Move name to lbcSelection(0)-Dump and lbcSelection(1)-Log
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        slStr = lbcFileName.List(ilLoop)
        If (UCase$(Mid$(slStr, 2, 1)) >= "0") And (UCase$(Mid$(slStr, 2, 1)) <= "9") Then
            lbcSelection(1).AddItem slStr
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopCmmlChgFileNames            *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain files with *.CCF        *
'*                                                     *
'*******************************************************
Private Sub mPopCmmlChgFileNames()
    Dim ilLoop As Integer
    Dim slStr As String
    lbcFileName.Path = Left$(sgExportPath, Len(sgExportPath) - 1)
    lbcFileName.Pattern = "*.CCF"
    'Move name to lbcSelection(0)-Dump and lbcSelection(1)-Log
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        slStr = lbcFileName.List(ilLoop)
        lbcSelection(0).AddItem slStr
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopDallasFileNames             *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain files with *.DAL        *
'*                                                     *
'*******************************************************
Private Sub mPopDallasFileNames()
    Dim ilLoop As Integer
    Dim slStr As String
    lbcFileName.Path = Left$(sgExportPath, Len(sgExportPath) - 1)
    lbcFileName.Pattern = "*.DAL"
    'Move name to lbcSelection(0)-Dump and lbcSelection(1)-Log
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        'All names moved into lbcSelection(0)
        'Only names of log images moved into lbcSelection(1)
        slStr = lbcFileName.List(ilLoop)
        lbcSelection(0).AddItem slStr
        'If (UCase$(Mid$(slStr, 7, 1)) >= "A") And (UCase$(Mid$(slStr, 7, 1)) <= "Z") Then
            lbcSelection(1).AddItem slStr
        'End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopInvoiceExportFileNames      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain files with *.EDI        *
'*                                                     *
'*******************************************************
Private Sub mPopInvoiceExportFileNames()
    Dim ilLoop As Integer
    Dim slStr As String
    lbcFileName.Path = Left$(sgExportPath, Len(sgExportPath) - 1)
    lbcFileName.Pattern = "*.EDI"
    'Move name to lbcSelection(0)-Dump and lbcSelection(1)-Log
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        'All names moved into lbcSelection(0)
        slStr = lbcFileName.List(ilLoop)
        lbcSelection(0).AddItem slStr
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopNYFileNames                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain files with *.NY         *
'*                                                     *
'*******************************************************
Private Sub mPopNYFileNames()
    Dim ilLoop As Integer
    Dim slStr As String
    lbcFileName.Path = Left$(sgExportPath, Len(sgExportPath) - 1)
    lbcFileName.Pattern = "??????CS.TXT"
    'Move name to lbcSelection(0)-Dump and lbcSelection(1)-Log
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        'All names moved into lbcSelection(0)
        slStr = lbcFileName.List(ilLoop)
        lbcSelection(0).AddItem slStr
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopNYFileNames                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain files with *.NY         *
'*                                                     *
'*******************************************************
Private Sub mPopPhoenixFileNames()
    Dim ilLoop As Integer
    Dim slStr As String
    lbcFileName.Path = Left$(sgExportPath, Len(sgExportPath) - 1)
    lbcFileName.Pattern = "??????CS.PHX"
    'Move name to lbcSelection(0)-Dump and lbcSelection(1)-Log
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        'All names moved into lbcSelection(0)
        slStr = lbcFileName.List(ilLoop)
        lbcSelection(0).AddItem slStr
    Next ilLoop
End Sub
'
'
'               Selling, conventional and airing vehicles have
'               been populated in list box. If vehicle is an airing
'               vehicle, see if option is set to use airing vehicle
'               for copy.  If not, remove it from list box for user
'               to select.
'
'               <input> lbcLocal - list box containing vehicles populated
'                       tlSortCode - list of vehicles populated with codes, names
'               <output> lbcLocal - list box with airing vehicles removed
'                       tlSortCode - list box of vehicles with codes, names
'                       and airing vehicles removed
Private Sub mRemoveAirVeh(lbcLocal As control, tlSortCode() As SORTCODE)
ReDim tlTempSortCode(0 To 0) As SORTCODE
Dim ilLoop As Integer
Dim slNameCode As String
Dim ilRet As Integer
Dim slCode As String
Dim ilUpper As Integer
Dim ilVehicleOk As Integer
Dim ilVehLoop As Integer
Dim ilVpfCode As Integer
    ilUpper = 0
    For ilLoop = UBound(tlSortCode) - 1 To 0 Step -1
        slNameCode = tlSortCode(ilLoop).sKey    'pick up vehicle code
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        'find the matching vehicle and see if its an airing vehicle.
        'if so, see what its option is for copy
        ilVehicleOk = True
        'For ilVehLoop = 1 To UBound(tgMVef) - 1 Step 1
        For ilVehLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If tgMVef(ilVehLoop).iCode = Val(slCode) Then
                ilVpfCode = gVpfFindIndex(Val(slCode))
                If ilVpfCode >= 0 Then
                    If tgMVef(ilVehLoop).sType = "A" And tgVpf(ilVpfCode).sCopyOnAir <> "Y" Then
                        lbcLocal.RemoveItem ilLoop        'remove item from list box
                        ilVehicleOk = False
                    End If
                End If
                Exit For
            End If
        Next ilVehLoop
        'If the vehicle was found to be removed, do not add it to the name/code array
        If ilVehicleOk Then
            tlTempSortCode(ilUpper).sKey = tlSortCode(ilLoop).sKey
            ReDim Preserve tlTempSortCode(ilUpper + 1) As SORTCODE
            ilUpper = ilUpper + 1
        End If
    Next ilLoop
    'The airing vehicles that dont use copy on airing vehicles have been removed,
    'now put them back into the original array
    ReDim tlSortCode(0 To 0) As SORTCODE
    ilUpper = 0
    For ilLoop = UBound(tlTempSortCode) - 1 To 0 Step -1
        tlSortCode(ilUpper) = tlTempSortCode(ilLoop)
        ReDim Preserve tlSortCode(ilUpper + 1) As SORTCODE
        ilUpper = ilUpper + 1
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSalesOfficePop                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSalesOfficePop(lbcSelection As control)
    Dim ilRet As Integer
    'ilRet = gPopOfficeSourceBox(RptSelCreditStatus, lbcSelection, lbcSOCode)
    ilRet = gPopOfficeSourceBox(RptSelCreditStatus, lbcSelection, tgSOCode(), sgSOCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSalesOfficePopErr
        gCPErrorMsg ilRet, "mSalesOfficePop (gPopOfficeSourceBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mSalesOfficePopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSellConvAirPop                 *
'*                                                     *
'*             Created:8/23/99       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box with Selling, Conventional,*
'*                      and Airing Vehicles            *
'*                                                     *
'*******************************************************
Private Sub mSellConvAirPop(ilIndex As Integer, ilUselbcVehicle As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvAirPopErr
        gCPErrorMsg ilRet, "mSellConvAirPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mSellConvAirPopErr:
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
        'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVirtVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
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
'       3/30/99 add option to include/exclude dormant
'               vehicles
'*******************************************************

Private Sub mSellConvVVPkgPop(ilIndex As Integer, ilUselbcVehicle As Integer, ilUseDormant As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        If ilUseDormant Then
            ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
        Else
            ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
        End If
    Else
        'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        If ilUseDormant Then
            ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)   'lbcCSVNameCode)
        Else
            ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)   'lbcCSVNameCode)
        End If
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVVPkgPopErr
        gCPErrorMsg ilRet, "mSellConvVVPkgPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVVPkgPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSellVehPop                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSellVehPop()
    Dim ilRet As Integer
    'ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHSELLING + ACTIVEVEH, lbcSelection(0), lbcSellNameCode)
    ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHSELLING + ACTIVEVEH, lbcSelection(0), tgSellNameCode(), sgSellNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellVehPopErr
        gCPErrorMsg ilRet, "mSellVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mSellVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
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
    Dim ilListIndex As Integer
    Dim ilIndex As Integer
    Dim ilSelect As Integer

    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = BUDGETSJOB Then
        If ckcAll.Value = vbChecked Then                   ' check for the correct selection list box here
            ilEnable = True
        Else
            If rbcSelCSelect(0).Value Then          'budget office option
                ilSelect = 1
            Else                                    'vehicle option
                ilSelect = 0
            End If
            For ilLoop = 0 To lbcSelection(ilSelect).ListCount - 1 Step 1
                If lbcSelection(ilSelect).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        End If
        If (ilEnable) Then                      'vehicle or office selected, check on budget names
            ilEnable = False                    'reset to test budget name
            For ilLoop = 0 To lbcSelection(4).ListCount - 1 Step 1
                If lbcSelection(4).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        End If
        If ilListIndex = 1 Then                     'comparisons option
            If (ilEnable) Then                      'vehicle or office selected, check on budget names
                ilEnable = False                    'reset to test budget name comparisons
                For ilLoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                    If lbcSelection(2).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
                If edcSelCTo <> "" And RptSelCreditStatus!rbcSelCSelect(2).Value Then
                    ilEnable = False
                End If
            End If
        End If
    ElseIf igRptCallType = RATECARDSJOB Then
        If ilListIndex = RC_RCITEMS Then
            If edcSelCFrom <> "" Then   'Or (edcSelCTo <> "" And RptSelCreditStatus!rbcSelCInclude(2).Value) Then
                If (edcSelCTo = "" And RptSelCreditStatus!rbcSelCInclude(2).Value) Then
                    ilEnable = False
                Else
                    ilEnable = False
                    If ckcAll.Value = vbChecked Then                   ' check for the correct selection list box here
                        ilEnable = True
                    Else
                        For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                            If lbcSelection(0).Selected(ilLoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If ilEnable Then
                        ilEnable = False
                        If ckcAllRC.Value = vbChecked Then                   ' check for the correct selection list box here
                            ilEnable = True
                        Else
                            For ilLoop = 0 To lbcSelection(11).ListCount - 1 Step 1
                                If lbcSelection(11).Selected(ilLoop) Then
                                    ilEnable = True
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    End If
                End If
            End If
        Else
            ilEnable = True
        End If
    ElseIf igRptCallType = USERLIST Then         '9-28-09
        If ilListIndex = USER_OPTIONS Then
            ilEnable = True
        ElseIf ilListIndex = USER_ACTIVITY Then         '5-6-11
            ilEnable = False
            If lbcSelection(10).SelCount > 0 Then
                ilEnable = True
            End If
        End If
    ElseIf igRptCallType = PROGRAMMINGJOB Then
        If igRptType = 3 Then               'program reports (vs links)
            ilEnable = False                'first assume nothing answered
            If ckcAll.Value = vbChecked Then
                ilEnable = True
            Else
                For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                    If lbcSelection(0).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                     End If
                Next ilLoop
             End If
        Else                            'links
            If (ilListIndex = 0) Or (ilListIndex = 1) Or (ilListIndex = 2) Or (ilListIndex = PRG_AIRING_INV) Then   'Selling to Airing
                ilEnable = False
    '            If (rbcRptType(0).Value) Or (rbcRptType(2).Value) Then
                If (ilListIndex = 0) Or (ilListIndex = 2) Then
                    ilSelect = 0
                Else
                    ilSelect = 1
                End If
                If lbcSelection(ilSelect).SelCount > 0 Then         'vehicle selectivity
                    If ilListIndex = PRG_AIRING_INV Then            '3-31-15 addl list box to test to see if at least 1 avail name selected
                        If lbcSelection(5).SelCount > 0 Then
                            ilEnable = True
                        End If
                    Else
                        ilEnable = True
                    End If
                End If
                
                If ilEnable Then
    '                If rbcRptType(0).Value Or rbcRptType(1).Value Then
                    If (ilListIndex = 0) Or (ilListIndex = 1) Then
                        If (ckcSel1(0).Value = vbChecked) Or (ckcSel1(1).Value = vbChecked) Then
                            ilEnable = True
                        Else
                            ilEnable = False
                        End If
                    End If
                End If
                If ilEnable Then
                    If (ckcSel2(0).Value = vbChecked) Or (ckcSel2(1).Value = vbChecked) Or (ckcSel2(2).Value = vbChecked) Then
                        ilEnable = True
                    Else
                        ilEnable = False
                    End If
                End If
            ElseIf (ilListIndex = 3) Or (ilListIndex = 4) Then   'Delivery
                ilEnable = False
    '            If rbcRptType(0).Value Then
                If ilListIndex = 3 Then
                    ilSelect = 3
                Else
                    ilSelect = 2
                End If
                For ilLoop = 0 To lbcSelection(ilSelect).ListCount - 1 Step 1
                    If lbcSelection(ilSelect).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
                If ilEnable Then
                    If (ckcSel1(0).Value = vbChecked) Or (ckcSel1(1).Value = vbChecked) Then
                        ilEnable = True
                    Else
                        ilEnable = False
                    End If
                End If
                If ilEnable Then
                    If (ckcSel2(0).Value = vbChecked) Or (ckcSel2(1).Value = vbChecked) Or (ckcSel2(2).Value = vbChecked) Then
                        ilEnable = True
                    Else
                        ilEnable = False
                    End If
                End If
            ElseIf (ilListIndex = 5) Or (ilListIndex = 6) Then   'Engineering
                ilEnable = False
    '            If rbcRptType(0).Value Then
                If ilListIndex = 5 Then
                    ilSelect = 3
                Else
                    ilSelect = 2
                End If
                For ilLoop = 0 To lbcSelection(ilSelect).ListCount - 1 Step 1
                    If lbcSelection(ilSelect).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
                If ilEnable Then
                    If (ckcSel1(0).Value = vbChecked) Or (ckcSel1(1).Value = vbChecked) Then
                        ilEnable = True
                    Else
                        ilEnable = False
                    End If
                End If
                If ilEnable Then
                    If (ckcSel2(0).Value = vbChecked) Or (ckcSel2(1).Value = vbChecked) Or (ckcSel2(2).Value = vbChecked) Then
                        ilEnable = True
                    Else
                        ilEnable = False
                    End If
                End If
            End If
        End If                  'program reports (vs links)
    ElseIf igRptCallType = CHFCONVMENU Then
        ilEnable = False
        For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
            If lbcSelection(0).Selected(ilLoop) Then
                ilEnable = True
                Exit For
            End If
        Next ilLoop
    ElseIf igRptCallType = GENERICBUTTON Then
        If Trim$(edcSelA.Text) <> "" Then
            ilEnable = True
        Else
            ilEnable = False
        End If
    ElseIf igRptCallType = DALLASFEED Then
        ilEnable = False
        If (ilListIndex = 0) Then
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        ElseIf (ilListIndex = 1) Then
            For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                If lbcSelection(1).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        Else
            ilEnable = True
        End If
    ElseIf igRptCallType = NYFEED Then
        ilEnable = False
        If (ilListIndex = 0) Then
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        ElseIf (ilListIndex = 1) Then
            ilEnable = True
        ElseIf (ilListIndex = 2) Or (ilListIndex = 3) Then
            ilEnable = True     'Blank field allows all records to be shown
        End If
    ElseIf igRptCallType = CMMLCHG Then
        ilEnable = False
        If (ilListIndex = 0) Then
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        Else
            ilEnable = True
        End If
    ElseIf igRptCallType = EXPORTAFFSPOTS Then
        ilEnable = False
        If (ilListIndex = 0) Then
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        Else
            ilEnable = True
        End If
    ElseIf igRptCallType = PHOENIXFEED Then
        ilEnable = False
        If (ilListIndex = 0) Then
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        Else
            ilEnable = True
        End If
    ElseIf igRptCallType = BULKCOPY Then
        ilEnable = False
        If (ilListIndex = 0) Then
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        ElseIf (ilListIndex = 1) Then   'Cross Ref
            For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                If lbcSelection(1).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        ElseIf (ilListIndex = 2) Then   'Affiliate BF by Cart
                ilEnable = True
        ElseIf (ilListIndex = 3) Then   'Affiliate BF by Vehicle
            For ilLoop = 0 To lbcSelection(3).ListCount - 1 Step 1
                If lbcSelection(3).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        ElseIf (ilListIndex = 4) Then   'Affiliate BF by Feed Date
            ilEnable = True
        ElseIf (ilListIndex = 5) Then   'Affiliate BF by Advertiser
            ilEnable = True
        End If
    ElseIf igRptCallType = POSTLOGSJOB Then
        ilEnable = False
        If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    If ilListIndex = 0 Then         'post log status
                        If (ckcSelC3(0).Value = vbChecked) Or (ckcSelC3(1).Value = vbChecked) Or (ckcSelC3(4).Value = vbChecked) Or (ckcSelC3(5).Value = vbChecked) Or (ckcSelC3(6).Value = vbChecked) Then
                            ilEnable = True
                        End If
                    Else                            'missing isci codes
                        ilEnable = True
                    End If
                    Exit For
                End If
            Next ilLoop
        Else
            ilEnable = False
        End If
    ElseIf igRptCallType = COPYJOB Then
        'Copy Status by date; Copy Status by Advertiser; Contracts Missing Copy; Play List
        'If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then        '7-1-04
        '    ilListIndex = ilListIndex + 1
        'End If
        If (ilListIndex = 0) Or (ilListIndex = 1) Or (ilListIndex = 2) Or (ilListIndex = 11) Or (ilListIndex = 13) Or (ilListIndex = 14) Or (ilListIndex = COPY_BOOK) Then
            If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
                If ilListIndex = 0 Then
                    ilIndex = 1
                ElseIf ilListIndex = 1 Then
                    ilIndex = 0
                ElseIf ilListIndex = 2 Then
                    ilIndex = 6        'missing copy
                ElseIf ilListIndex = COPY_BOOK Then
                    ilIndex = 2
                Else
                    If RptSelCreditStatus!rbcSelCSelect(1).Value = True Then    'copy by isci
                        If edcSelCTo1.Text <> "" Then
                            ilIndex = 2
                        Else
                            ilEnable = False
                        End If
                    ElseIf RptSelCreditStatus!rbcSelCSelect(0).Value Then
                        ilIndex = 2
                    Else
                        ilIndex = 0
                    End If
                End If
                ilEnable = False
                For ilLoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                    If lbcSelection(ilIndex).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
                If ilListIndex = 1 And Not ckcAll.Value = vbChecked Then          'status by advt, check if selective cntrs
                    ilEnable = False
                    For ilLoop = 0 To lbcSelection(5).ListCount - 1 Step 1  'if selective, must have at least chosen
                        If lbcSelection(5).Selected(ilLoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next ilLoop
                End If

            Else
                ilEnable = False
            End If
        End If
        'Rotation by Advertiser
        If ilListIndex = COPY_ROT Then
            ilEnable = False
            'blank date is allowed
            'if both ALL buttons checked, ok to generate
            If ckcAll.Value = vbChecked And ckcAllGroups.Value = vbChecked Then                 'all advt
                ilEnable = True
            Else        'check selectivity on either advt or vehicle
                If ckcAll.Value = vbChecked Then        'all advt
                    ilEnable = True
                Else
                    If lbcSelection(0).SelCount > 0 Then        'selected at least one advt
                        'see if contract selected
                        If lbcSelection(5).SelCount > 0 Then    'selected at least one cnt
                            ilEnable = True
                        End If
                    End If
                End If
                If ilEnable = True Then         'now test the vehicles
                    ilEnable = False
                    If lbcSelection(6).SelCount > 0 Then        '12-8-05 chg from lbc(2) to lbc(6)
                        ilEnable = True
                    End If
                End If
            End If

        End If

        'Split copy/Blackout rotation
        If ilListIndex = COPY_SPLITROT Then
            ilEnable = False
            'blank date is allowed
            If lbcSelection(0).SelCount > 0 And lbcSelection(6).SelCount > 0 Then        'selected at least one advt
               ilEnable = True
            End If

        End If
        'Inventory by Number or ISCI
        If (ilListIndex = 4) Or (ilListIndex = 5) Then
            'If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
                ilEnable = True
            'Else
            '    ilEnable = False
            'End If
        End If
        'Inventory by Advertiser or Copy Regions
        If ilListIndex = 6 Then     '2-12-09 ove copy regions to splitregion list 'Or ilListIndex = COPY_REGIONS Then
            ilEnable = False
            For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        End If

        If ilListIndex = COPY_INVBYSTARTDATE Then
            If (edcSelCTo.Text <> "") Then
                ilEnable = True
                If lbcSelection(10).SelCount <= 0 And tgSpf.sUseCartNo <> "N" Then
                    ilEnable = False
                End If
            Else
                ilEnable = False
            End If
        End If

        'Inventory by  Expiration Date; Purge Date; Entry Date
        If (ilListIndex = 8) Or (ilListIndex = 9) Or (ilListIndex = 10) Then
            If (edcSelCTo.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        End If

        If ilListIndex = 12 Then        'unapproved copy
            ilEnable = False
           If (edcSelCTo.Text <> "") Then
                ilEnable = True
           End If
        End If
        
         'COPY_INVPRODUCER blank dates ok, assume everything
        If ilListIndex = COPY_INVPRODUCER Then          '4-10-13  start/end dates must be entered
            ilEnable = True                             'dates are optional
            If lbcSelection(10).SelCount <= 0 Then
                ilEnable = False
            End If
        End If
        
        '4-9-12 Script Affidavits
        If ilListIndex = COPY_SCRIPTAFFS Then
            If lbcSelection(0).SelCount > 0 Then
                If edcSelCFrom.Text <> "" And edcSelCTo.Text <> "" Then
                    ilEnable = True
                End If
            End If
        End If
    ElseIf igRptCallType = COLLECTIONSJOB Then
        'If rbcRptType(4).Value Then  'History
        ilEnable = False
        If ilListIndex = COLL_CASH Then   'cash receipts by entry date must have both dates entered
            ilEnable = True
            If rbcSelCSelect(1).Value Then
                If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
                    ilEnable = True
                End If
            End If
            If rbcSelC4(1).Value Then               'sort by slsp
                ilEnable = False
                If ckcAll.Value = vbChecked Then
                    ilEnable = True
                Else
                    For ilLoop = 0 To lbcSelection(5).ListCount - 1 Step 1
                        If lbcSelection(5).Selected(ilLoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next ilLoop
                End If
            End If
        'D.S. 8/15/01 added 13 lines below to enable and disable generate button
        ElseIf ilListIndex = COLL_CASHSUM Then
            If RptSelCreditStatus!rbcSelCSelect(0).Value = True Then  'sort by vehicle
                ilEnable = True
            Else
                ilEnable = False                          'sort by sales office
                For ilLoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                    If lbcSelection(ilIndex).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
            End If
        ElseIf ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGESLSP Or ilListIndex = COLL_AGEVEHICLE Or ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPRODUCER Then '2-10-00
            If ckcAll.Value = vbChecked Then
                ilEnable = True
            Else
                If ilListIndex = COLL_AGEPAYEE Then
                    ilIndex = 2
                ElseIf ilListIndex = COLL_AGESLSP Then
                    ilIndex = 5
                ElseIf ilListIndex = COLL_AGEVEHICLE Then
                    ilIndex = 6
                Else                'participant or sales source
                    ilIndex = 3
                End If
                For ilLoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                    If lbcSelection(ilIndex).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
            End If
                '6-04-08 added checkboxes. make sure at least one selected.
                If ilEnable = True Then
                    Dim clCheckObject As CheckBox
                    Dim ilChecked As Integer
                    ilChecked = 0
                    For Each clCheckObject In ckcSelC6Add
                    If clCheckObject.Value = 1 Then
                        ilChecked = ilChecked + 1
                    End If
                    Next clCheckObject
                    If ilChecked >= 1 Then
                        ilEnable = True
                    Else
                        ilEnable = False
                    End If
                End If
        ElseIf ilListIndex = COLL_PAYHISTORY Then
            If rbcSelCSelect(0).Value Then  'by agency
                ilIndex = 2
            ElseIf rbcSelCSelect(1).Value Then  'by advt
                ilIndex = 0
            'Else                                'vehicle
            '    ilIndex = 6
            End If
            For ilLoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                If lbcSelection(ilIndex).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        ElseIf ilListIndex = 4 Then  'Delinquent
            If (edcSelCTo.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        'ElseIf rbcRptType(3).Value Then  'Statements
        ElseIf ilListIndex = 5 Or ilListIndex = COLL_POAPPLY Then  'Statements and POs Applied
            For ilLoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                If lbcSelection(2).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
            Next ilLoop
        ElseIf ilListIndex = COLL_CREDITSTATUS Then  'Credit Status
            'TTP 9893
            'TTP 9893
            ilEnable = False
            If (ckcSel2(0).Value = vbChecked) Or (ckcSel2(1).Value = vbChecked) Then
                'show Adv and/or Agv lists
                'ckcSel2(0) = "Agency"
                'ckcSel2(1) = "Advertiser"
                pbcOption.Visible = True
                If ckcSel2(0).Value = vbChecked And ckcSel2(1).Value <> vbChecked Then
                    ckcAll.Visible = True
                    lbcSelection(1).Visible = True          'Agency
                    lbcSelection(1).Height = 4000
                    lbcSelection(1).Top = 285
                    lbcSelection(0).Visible = False         'Advertiser disabled
                    ckcAllGroups.Visible = False
                    'make sure a Agcy Selected
                    For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                        If lbcSelection(1).Selected(ilLoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next ilLoop
                End If
                If ckcSel2(0).Value <> vbChecked And ckcSel2(1).Value = vbChecked Then
                    ckcAll.Visible = False
                    lbcSelection(1).Visible = False         'Agency
                    lbcSelection(0).Top = 285
                    ckcAllGroups.Top = ckcAll.Top
                    lbcSelection(0).Visible = True          'Advertiser disabled
                    lbcSelection(0).Height = 4000
                    ckcAllGroups.Visible = True
                    'make sure a Adv Selected
                    For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                        If lbcSelection(0).Selected(ilLoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next ilLoop
                End If
                If ckcSel2(0).Value = vbChecked And ckcSel2(1).Value = vbChecked Then
                    ckcAll.Visible = True
                    lbcSelection(1).Visible = True          'Agency
                    lbcSelection(1).Height = 1885
                    lbcSelection(1).Top = 285
                    
                    ckcAllGroups.Visible = True
                    ckcAllGroups.Top = lbcSelection(1).Top + lbcSelection(1).Height + 105
                    lbcSelection(0).Visible = True          'Advertiser
                    lbcSelection(0).Top = lbcSelection(1).Top + lbcSelection(1).Height + 405
                    lbcSelection(0).Height = 1885
                    
                    'make sure a Adv Selected
                    For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                        If lbcSelection(0).Selected(ilLoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next ilLoop
                    If ilEnable = True Then
                        ilEnable = False
                        'also make sure a Agcy Selected
                        For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                            If lbcSelection(1).Selected(ilLoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
                
                'ilEnable = True
            Else
                'hide Adv and Agv lists
                pbcOption.Visible = False
                ilEnable = False
            End If
        ElseIf ilListIndex = COLL_DISTRIBUTE Then       'cash distribution
            If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
                ilEnable = True
                If rbcSelCSelect(2).Value Then          'participants (vs invoice #)
                    If Not ckcAll.Value = vbChecked Then
                        ilEnable = False
                        For ilLoop = 0 To lbcSelection(3).ListCount - 1 Step 1
                            If lbcSelection(3).Selected(ilLoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
            End If
        ElseIf ilListIndex = COLL_ACCTHIST Then
            If (ckcSelC3(0).Value = vbChecked) Or (ckcSelC3(1).Value = vbChecked) Or (ckcSelC3(2).Value = vbChecked) Or (ckcSelC3(3).Value = vbChecked) Then
                ilEnable = True
                If rbcSelCSelect(0).Value = True Then
                    ilIndex = 0     'adv
                Else
                    ilIndex = 2     'direct & agy
                End If
                If lbcSelection(ilIndex).SelCount <= 0 Then
                    ilEnable = False
                End If
            Else
                ilEnable = False
            End If

        ElseIf ilListIndex = COLL_MERCHANT Then
            If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
                ilEnable = True
                If Not ckcAll.Value = vbChecked Then
                    ilEnable = False
                    If rbcSelCSelect(0).Value Then      'vehicle selection
                        ilSelect = 6
                    Else
                        ilSelect = 0                    'advertiser selection
                    End If
                    For ilLoop = 0 To lbcSelection(ilSelect).ListCount - 1 Step 1      'check at least 1 vehicle selected
                        If lbcSelection(ilSelect).Selected(ilLoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next ilLoop
                End If
            End If
        Else
            ilEnable = True
        End If
    ElseIf igRptCallType = INVOICESJOB Then
        If ilListIndex = 0 Then                 'invoice option
            If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
                ilEnable = True
                If Not ckcAll.Value = vbChecked Then
                    ilEnable = False

                    ilIndex = -1            'Determine if selection in list boxes have been requested 9-16-02
                    If rbcSelCSelect(1).Value Then          'advt
                        ilIndex = 5
                    ElseIf rbcSelCSelect(2).Value Then      'agy
                        ilIndex = 1
                    ElseIf rbcSelCSelect(3).Value Then      'slsp
                        ilIndex = 2
                    ElseIf rbcSelCSelect(4).Value Or rbcSelCSelect(5).Value Or rbcSelCSelect(6).Value Then    'billing or airing vehicle
                        ilIndex = 6
                    ElseIf rbcSelCSelect(7).Value Then      'ntr
                        ilIndex = 8
                    Else            'list boxes dont apply
                       ilEnable = True
                    End If
                    If ilIndex > 0 Then
                        For ilLoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                            If lbcSelection(ilIndex).Selected(ilLoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If

            Else
                ilEnable = False
            End If
        ElseIf ilListIndex = INV_DISTRIBUTE Then                'billing distribution
            If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
                ilEnable = True
                If Not ckcAll.Value = vbChecked Then
                    ilEnable = gSetGenCommand(RptSelCreditStatus!lbcSelection(3))   '10-8-03
                End If
            End If
        ElseIf ilListIndex = 1 Then
            ilEnable = gSetGenCommand(RptSelCreditStatus!lbcSelection(0))   '10-8-03
        ElseIf ilListIndex = INV_CREDITMEMO Then                '10-8-03
            ilEnable = gSetGenCommand(RptSelCreditStatus!lbcSelection(7))
        ElseIf ilListIndex = INV_SUMMARY Then                       '6-28-05
            If (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
                ilEnable = True
                If Not ckcAll.Value = vbChecked Then
                    ilEnable = False

                    ilIndex = -1            'Determine if selection in list boxes have been requested 9-16-02
                    If rbcSelCSelect(0).Value Then          'advt
                        ilIndex = 5
                    ElseIf rbcSelCSelect(1).Value Then      'agy
                        ilIndex = 1
                    ElseIf rbcSelCSelect(2).Value Then      'slsp
                        ilIndex = 2
                    Else            'list boxes dont apply
                       ilEnable = True
                    End If
                    If ilIndex > 0 Then
                        For ilLoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                            If lbcSelection(ilIndex).Selected(ilLoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
            Else
                ilEnable = False
            End If
         ElseIf ilListIndex = INV_TAXREGISTER Then
            If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
                ilEnable = True
                If Not ckcAll.Value = vbChecked Then
                    ilEnable = gSetGenCommand(RptSelCreditStatus!lbcSelection(6))   'at least 1 vehicle must be selected
                End If
            End If
        ElseIf ilListIndex = INV_RECONCILE Then         'installment reconcil
            If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") Then
                ilEnable = True
            End If
        End If
    
    Else
        'If (rbcRptType(1).Value) Or (rbcRptType(0).Value And (igRptCallType = SALESPEOPLELIST)) Or (igRptCallType = EVENTNAMESLIST) Then
        If (ilListIndex = 1) Or ((ilListIndex = 0) And (igRptCallType = SALESPEOPLELIST)) Or (igRptCallType = EVENTNAMESLIST) Or ((ilListIndex = 0 Or ilListIndex = 5) And igRptCallType = VEHICLESLIST) Then
            ilEnable = False
            If (ilListIndex = 1) Or (igRptCallType = EVENTNAMESLIST) Or ((ilListIndex = 0 Or ilListIndex = 5) And igRptCallType = VEHICLESLIST) Then
                For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                    If lbcSelection(0).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
            Else
'                For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1
'                    If lbcSelection(1).Selected(ilLoop) Then
'                        ilEnable = True
'                        Exit For
'                    End If
'                Next ilLoop
                If lbcSelection(0).SelCount > 0 Then            'for slsp list, test selective slsp
                    ilEnable = True
                End If
            End If
        Else
            ilEnable = True
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
Private Sub mSPersonPop(lbcSelection As control)
'
'   mSPersonPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopSalespersonBox(RptSelCreditStatus, 0, True, True, lbcSelection, Traffic!lbcSalesperson, igSlfFirstNameFirst)
    ilRet = gPopSalespersonBox(RptSelCreditStatus, 0, True, True, lbcSelection, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gPopSalespersonBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mSPersonPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
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
Private Sub mSSourcePop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopSSourceByGroupBox(RptSelCreditStatus, lbcSelection(1), lbcNameCode)
    ilRet = gPopSSourceByGroupBox(RptSelCreditStatus, lbcSelection(1), tgRptSelCreditStatusNameCode(), sgRptSelCreditStatusNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSSourcePopErr
        gCPErrorMsg ilRet, "mSSourcePop (gPopSSourceByGroupBox)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mSSourcePopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
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
    Unload RptSelCreditStatus
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcRepInv_Paint()
    plcRepInv.Cls
    plcRepInv.CurrentX = 0
    plcRepInv.CurrentY = 0
    plcRepInv.Print "Rep Billing"
End Sub


Private Sub plcSelC10_Paint()
    plcSelC10.Cls
    plcSelC10.CurrentX = 0
    plcSelC10.CurrentY = 0
    plcSelC10.Print smPlcSelC10P
End Sub

Private Sub plcSelC11_Paint()
    plcSelC11.Cls
    plcSelC11.CurrentX = 0
    plcSelC11.CurrentY = 0
    plcSelC11.Print smPlcSelC11P
End Sub

Private Sub plcSelC12_Paint()
    plcSelC12.Cls
    plcSelC12.CurrentX = 0
    plcSelC12.CurrentY = 0
    plcSelC12.Print smPlcSelC12P
End Sub

Private Sub plcSelC6_Paint()
    plcSelC6.Cls
    plcSelC6.CurrentX = 0
    plcSelC6.CurrentY = 0
    plcSelC6.Print smPlcSelC6P
End Sub

Private Sub plcSelC8_Paint()
    plcSelC8.Cls
    plcSelC8.CurrentX = 0
    plcSelC8.CurrentY = 0
    plcSelC8.Print smPlcSelC8P
End Sub

Private Sub plcSelC9_Click()
    plcSelC9.Cls
    plcSelC9.CurrentX = 0
    plcSelC9.CurrentY = 0
    plcSelC9.Print smPlcSelC9P
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
        If imAutoReport Then            '2-24-05 if coming from auto report (Show Credit screen), then
                                        'force report to display without asking any questions
            cmcGen_Click
            cmcCancel_Click
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub

Private Sub rbcSelC12_Click(Index As Integer)
    Dim Value As Integer
    Value = rbcSelC12(Index).Value
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex

End Sub

Private Sub rbcSelC4_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC4(Index).Value
    'End of coded added
Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = BUDGETSJOB Then
        If Index = 0 Then
            rbcSelC4(1).Enabled = False
        Else
            rbcSelC4(0).Enabled = False
        End If
    ElseIf igRptCallType = COLLECTIONSJOB Then
        If ilListIndex = COLL_CASH Then
            plcSelC12.Visible = False
            If Index = 0 Then
                ckcAll.Visible = False
                lbcSelection(5).Visible = False
                edcSet1.Text = "Vehicle Group Subsort"
            ElseIf Index = 1 Then
                ckcAll.Visible = True
                ckcAll.Caption = "All Salespeople"
                ckcAll.Left = 0
                lbcSelection(5).Height = 1500
                lbcSelection(7).Height = 1500
                ckcAllGroups.Caption = "All Sales Offices"
                ckcAllGroups.Move ckcAll.Left, lbcSelection(5).Top + lbcSelection(5).Height + 60
                ckcAllGroups.Visible = True
                lbcSelection(7).Move lbcSelection(5).Left, ckcAllGroups.Top + ckcAllGroups.Height + 60, 4365, 1500
                lbcSelection(7).Visible = True
                lbcSelection(5).Visible = True
                edcSet1.Text = "Vehicle Group Subsort"
                plcSelC12.Move 120, plcSelC8.Top + plcSelC8.Height
                rbcSelC12(0).Caption = "Check #"
                rbcSelC12(0).Value = True
                rbcSelC12(1).Caption = "Advertiser"
                rbcSelC12(0).Move 1200, 0, 960
                rbcSelC12(1).Move 2240, 0, 1200
                rbcSelC12(0).Visible = True
                rbcSelC12(1).Visible = True
                plcSelC12.Visible = True
                smPlcSelC12P = "Subtotals by"

            Else
                ckcAll.Visible = False
                lbcSelection(5).Visible = False
                edcSet1.Text = "Vehicle Group Sort"
            End If
        End If
    End If
    mSetCommands
End Sub

Private Sub rbcSelC6_Click(Index As Integer)
Dim Value As Integer
Dim ilListIndex As Integer

    Value = rbcSelC8(Index).Value
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = INVOICESJOB Then
        If ilListIndex = INV_REGISTER Then
            If Index = 0 Then              'air time
                ckcSelC7.Value = vbUnchecked
            End If
        End If
    End If
End Sub

Private Sub rbcSelC8_Click(Index As Integer)
   Dim Value As Integer
    Dim ilListIndex As Integer

    Value = rbcSelC8(Index).Value
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = AGENCIESLIST Then
        If ilListIndex = 2 Then             'mailing labels
            ckcAll.Value = vbUnchecked
            If Index = 0 Then               'payee
                ckcAll.Caption = "All Agencies and Advertisers"
                lbcSelection(2).Visible = True
                lbcSelection(0).Visible = False
                plcSelC2.Visible = True
                plcSelC4.Visible = True
                lacSelCFrom.Visible = False
                plcSelC3.Visible = False
            Else                            'vehicle
                plcSelC2.Visible = False
                plcSelC4.Visible = False
                mVehLabelsPop VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHLOGVEHICLE + VEHNTR + VEHSIMUL                 'populate the vehicles in lbcselection(0)
                ckcAll.Caption = "All Vehicles"
                lacSelCFrom.Move 120, plcSelC1.Top + plcSelC1.Height, 4000
                lacSelCFrom.Caption = "Vehicles without address will be excluded"
                ckcSelC3(0).Value = vbChecked        'airing
                ckcSelC3(1).Value = vbChecked        'conventional
                ckcSelC3(2).Value = vbChecked        'log
                ckcSelC3(3).Value = vbChecked        'ntr
                ckcSelC3(4).Value = vbChecked        'rep
                ckcSelC3(5).Value = vbChecked        'simulcast
                plcSelC3.Move 120, lacSelCFrom.Top + lacSelCFrom.Height
                plcSelC3.Height = 675
                ckcSelC3(0).Caption = "Airing"
                ckcSelC3(1).Caption = "Conventional"
                ckcSelC3(2).Caption = "Log"
                ckcSelC3(3).Caption = "NTR"
                ckcSelC3(4).Caption = "Rep"
                ckcSelC3(5).Caption = "Simulcast"
                ckcSelC3(0).Move 1440, 0, 840
                ckcSelC3(1).Move 2280, 0, 1440
                ckcSelC3(2).Move 1440, 210, 840
                ckcSelC3(3).Move 2280, 210, 840
                ckcSelC3(4).Move 3120, 210, 720
                ckcSelC3(5).Move 1440, 420, 1200
                ckcSelC3(0).Visible = True
                ckcSelC3(1).Visible = True
                ckcSelC3(2).Visible = True
                ckcSelC3(3).Visible = True
                ckcSelC3(4).Visible = True
                ckcSelC3(5).Visible = True

                lacSelCFrom.Visible = True
                smPlcSelC3P = "Vehicle Types"
                plcSelC3.Visible = True
                lbcSelection(0).Visible = True
                lbcSelection(2).Visible = False

            End If
        End If
    ElseIf igRptCallType = INVOICESJOB Then
        If ilListIndex = INV_REGISTER Then
            If rbcSelCSelect(9).Value Then               'Sales Origin option
                    If Index = 0 Then                        'no vehicle totals
                        ckcAll.Visible = False
                        lbcSelection(6).Visible = False
                        ckcAll.Value = vbChecked
                    Else
                        ckcAll.Visible = True
                        lbcSelection(6).Visible = True
                        ckcAll.Value = vbUnchecked
                        ckcAll.Caption = "All Vehicles"
                        smPlcSelC9P = ""
                        plcSelC9.Move 120, plcSelC8.Top + plcSelC8.Height
                        plcSelC9.Visible = True
                        ckcTrans.Left = 0
                        ckcTrans.Visible = True
                        ckcTrans.Caption = "New page each vehicle"
                    End If
            End If
        End If
    ElseIf igRptCallType = POSTLOGSJOB Then
        If ilListIndex = 0 Then         'post log report
            If Index = 0 Then           'detail
                plcSelC9.Visible = False    'hide the option to see day incomplete flags only
                ckcTrans.Value = vbUnchecked
            Else
                plcSelC9.Visible = True
                ckcTrans.Visible = True     'day is incomplete only option
            End If
        End If

    End If

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
            Case RATECARDSJOB
                If rbcSelCInclude(2).Value Then             'week option
                    lacSelCTo.Visible = True
                    edcSelCTo.Visible = True
                    lacSelCTo.Caption = "Start Date"
                    lacSelCTo.Move plcSelC4.Left, plcSelC4.Top + plcSelC4.Height + 75, 960
                    edcSelCTo.MaxLength = 10    '8  5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
                    edcSelCTo.Move plcSelC4.Left + 960, plcSelC4.Top + plcSelC4.Height + 30

                Else
                    lacSelCTo.Visible = False
                    edcSelCTo.Visible = False
                    lacSelCFrom.Visible = True          'show Year label
                    lacSelCFrom.Caption = "Year"
                    edcSelCFrom.Visible = True          'show year edit box
                    edcSelCFrom.Width = 600
                    edcSelCFrom.MaxLength = 4           'year length (i.e.1996, 2000)
                    lacSelCFrom.Move plcSelC2.Left, 75, 510
                    edcSelCFrom.Move lacSelCFrom.Left + lacSelCFrom.Width + 30, 30    'Year edit box
                    plcSelC1.Visible = False                    'don't show first set of radio buttons

                End If

            Case BUDGETSJOB
                If rbcSelCInclude(2).Value Then             'week option, ask start date
                    lacSelCTo.Visible = True
                    edcSelCTo.Visible = True
                    lacSelCTo.Caption = "Start Date"
                    lacSelCTo.Move plcSelC3.Left, plcSelC3.Top + plcSelC3.Height + 75, 960
                    edcSelCTo.MaxLength = 10    '8  5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
                    edcSelCTo.Move plcSelC3.Left + 960, plcSelC3.Top + plcSelC3.Height + 30
                Else
                    lacSelCTo.Visible = False
                    edcSelCTo.Visible = False
                End If
            Case COLLECTIONSJOB
                Select Case ilListIndex
                ' dan    rbcselc6 no longer works on the following: COLL_AGEPAYEE, COLL_AGESLSP, COLL_AGEVEHICLE, COLL_AGEOWNER, COLL_AGESS, COLL_AGEPRODUCER,
                    Case COLL_CASH, COLL_CASHSUM
                        If Index = 0 Or Index = 1 Then       'cash or trade
                            rbcSelC6(0).Value = True    'force to Airtime
                            rbcSelC6(0).Enabled = True
                            rbcSelC6(1).Enabled = True
                            rbcSelC6(2).Enabled = True
                        ElseIf Index = 2 Or Index = 3 Then      'merch or promotions
                            rbcSelC6(0).Value = True
                            rbcSelC6(0).Enabled = False
                            rbcSelC6(1).Enabled = False
                            rbcSelC6(2).Enabled = False
                        Else                                    'ntr hard-cost only
                            rbcSelC6(1).Value = True            'force to ntr
                            rbcSelC6(0).Enabled = False
                            rbcSelC6(1).Enabled = False
                            rbcSelC6(2).Enabled = False
                        End If
                End Select
            Case INVOICESJOB
                Select Case ilListIndex
                    Case INV_SUMMARY
                        'Invoice summary, if detail ask to hide/show transaction amount
                        If Index = 0 Then
                            plcSelC10.Move 120, edcCheck.Top + edcCheck.Height, 3360
                            plcSelC10.Visible = True
                            ckcSelC10(0).Caption = "Hide transaction net amount"
                            ckcSelC10(0).Visible = True
                            ckcSelC10(0).Move 0, 0, 3360
                        Else
                            plcSelC10.Visible = False
                        End If
                    Case INV_REGISTER
                        ckcOption.Enabled = True            'assume OK to ask to show comments on AN
                        If rbcSelCSelect(1).Value = True Then        'by advt
                            If Index = 1 Then           'advt option, with summary selected:  disallow vehicle groups
                                cbcSet1.Visible = False
                                edcSet1.Visible = False
                                lacCheck.Visible = False
                                cbcSet1.ListIndex = 0
                                ckcOption.Value = False
                                ckcOption.Enabled = False
                            Else                        'detail option
                                cbcSet1.Visible = True
                                edcSet1.Visible = True
                                lacCheck.Visible = True
                                cbcSet1.ListIndex = 0
                            End If
                        ElseIf rbcSelCSelect(0).Value = True And Index = 2 Then            'by invoice, summary, disallow AN to show comments
                            ckcOption.Value = False
                            ckcOption.Enabled = False
                            
                        ElseIf rbcSelCSelect(0).Value = False And Index = 1 Then           'all other options, summary disallow AN to show comments
                            ckcOption.Value = False
                            ckcOption.Enabled = False
                        End If
                End Select
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
    Dim slStr As String
    Dim llTodayDate As Long
    Dim llRg As Long
    Dim llRet As Long
    Dim ilTop As Integer
    Dim ilHeight As Integer

    ilListIndex = lbcRptType.ListIndex
    If Value Then
        Select Case igRptCallType
            Case BUDGETSJOB
                    If rbcSelCSelect(0).Value Then      'office
                        lbcSelection(1).Visible = True
                        lbcSelection(1).Move 15, ckcAll.Top + ckcAll.Height + 30, 4380, 1500 'office list box
                        lbcSelection(0).Visible = False
                        ckcAll.Caption = "All Offices"
                    Else                                'vehicle
                        lbcSelection(0).Visible = True
                        lbcSelection(0).Move 15, ckcAll.Top + ckcAll.Height + 30, 4380, 1500 'vehicle list box
                        lbcSelection(1).Visible = False
                        ckcAll.Caption = "All Vehicles"
                    End If
            Case INVOICESJOB
                If (ilListIndex = INV_REGISTER) Then
                    lbcSelection(1).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(3).Visible = False
                    lbcSelection(4).Visible = False     '10-29-99
                    lbcSelection(5).Visible = False
                    lbcSelection(6).Visible = False
                    lbcSelection(7).Visible = False
                    lbcSelection(8).Visible = False     '9-16-02
                    lbcSelection(9).Visible = False     '10-17-02 Sales Source
                    lacCheck.Visible = False
                    ckcAll.Visible = True
                    ckcAll.Value = vbUnchecked   'False
                    edcSet1.Visible = False
                    cbcSet1.Visible = False
                    ckcAllGroups.Visible = False
                    ckcAllGroups.Value = False
                    
                    plcSelC4.Left = 120
                    plcSelC4.Top = plcSelC1.Top + plcSelC1.Height
                    
                    rbcSelC4(0).Caption = "Cash"
                    rbcSelC4(0).Move 480, 0, 720
                    rbcSelC4(0).Value = True
                    rbcSelC4(1).Caption = "Trade"
                    rbcSelC4(1).Move 1420, 0, 960
                    rbcSelC4(2).Caption = "Both"
                    rbcSelC4(2).Move 2380, 0, 720
                    plcSelC4.Visible = True
                    If rbcSelC4(2).Value Then          'default to billing
                        rbcSelC4_click 2
                    Else
                        rbcSelC4(2).Value = True
                    End If
                    smPlcSelC4P = "For"

                    plcSelC3.Left = 120
                    plcSelC3.Top = plcSelC4.Top + plcSelC4.Height
                    
                    mInvAskTypes
                    mAskAirTimeNTR plcSelC2.Top + 30 + plcSelC2.Height  '2-2-03
                    rbcSelC6(0).Enabled = True
                    rbcSelC6(1).Enabled = True
                    rbcSelC6(2).Enabled = True                  ' mAskRecHistBoth plcSelC6.Top + plcSelC6.Height      '3-19-03 option to include receivables, history or both
                   ' rbcSelC8(2).Value = True                            'Include both rceivables & History
                    plcSelC7.Move 120, plcSelC6.Top + plcSelC6.Height        '3-17-05 include hard cost
                    ckcSelC7.Caption = "Hard Cost Only"
                    ckcSelC7.Move 0, 0
                    ckcSelC7.Visible = True
                    ckcSelC7.Enabled = True
                    plcSelC7.Visible = True
                    plcSelC8.Visible = False                'major totals by vehicle for sales origin option
                    plcSelC10.Visible = False
                    plcSelC11.Move plcSelC7.Left, plcSelC7.Top + plcSelC7.Height
                    smPlcSelC11P = "Include"
                    ckcSelC11(0).Move 840, 0, 720
                    ckcSelC11(0).Value = vbChecked
                    ckcSelC11(1).Move 1680, 0, 1320
                    ckcSelC11(1).Value = vbChecked
                    plcSelC11.Visible = True
                    ckcSelC11(0).Visible = True
                    ckcSelC11(1).Visible = True
                    ilTop = plcSelC11.Top
                    ilHeight = plcSelC11.Height
                    
                    '1-14-12 option to show AN comments
                    ckcOption.Caption = "Show Adjustment (AN) comments"
                    ckcOption.Move 120, ilTop + ilHeight + 30, 3600
                    ckcOption.Visible = True
                    ilTop = ckcOption.Top
                    ilHeight = ckcOption.Height + 30
                    
                    '1-14-12 If using installment and Billing different than revenue, which register does user want
                    If rbcSelC12(0).Value Then          'default to billing
                        rbcSelC12_Click 0
                    Else
                        rbcSelC12(0).Value = True
                    End If

                    If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) = INSTALLMENT And (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED Then        'using installment with billing and revenue separate
                        plcSelC12.Move plcSelC11.Left, ilTop + ilHeight
                        plcSelC12.Visible = True
                        smPlcSelC12P = "Use"
                        rbcSelC12(0).Move 480, 0, 840       'billing
                        rbcSelC12(1).Move 1440, 0, 1200      'revenue
                        rbcSelC12(0).Visible = True
                        rbcSelC12(1).Visible = True
                        ilTop = plcSelC12.Top
                        ilHeight = plcSelC12.Height
                    'keep which type of register to get hidden if not using installment or its billing = revenue; as theres only one type
                    End If
                    ckcTrans.Visible = False                'New page each vehicle for Sales Origin and sort by either bill/air vehicle
                    ckcTrans.Value = vbUnchecked
                    Select Case Index
                        Case 0                              'Invoice
                            ckcAll.Visible = False
                            rbcSelCInclude(2).Visible = True
                        Case 1                              'advt
                            lbcSelection(5).Height = 3240
                            lbcSelection(5).Visible = True
                            ckcAll.Caption = "All Advertisers"
                            '5-13-11 user wants a vehicle group subsort within advt
                            gPopVehicleGroups RptSelCreditStatus!cbcSet1, tgVehicleSets1(), True
                            edcSet1.Text = "Vehicle Group"
                            cbcSet1.ListIndex = 0
                            'edcSet1.Move 120, plcSelC12.Top + 30 + plcSelC12.Height   '3-17-05 adjust top to insert hard cost option
                            edcSet1.Move 120, ilTop + 30 + ilHeight
                            cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 45
                            edcSet1.Visible = True
                            cbcSet1.Visible = True
                            lacCheck.Caption = "(subsort only)"
                            lacCheck.FontName = "arial"
                            lacCheck.FontSize = 8
                            lacCheck.Move cbcSet1.Left + edcSet1.Width + 120, edcSet1.Top, 1800
                            lacCheck.Visible = True

                        Case 2                              'agy, 1-29-15 option to add S/S selectivity with agency sort
                            lbcSelection(1).Height = 1600   '3240
                            lbcSelection(1).Visible = True
                            ckcAll.Caption = "All Agencies"
                            ckcAllGroups.Move 15, lbcSelection(1).Top + lbcSelection(1).Height + 60
                            ckcAllGroups.Caption = "All Sales Sources"
                            ckcAllGroups.Visible = True
                            lbcSelection(9).Move 15, ckcAllGroups.Top + ckcAllGroups.Height + 15, 4380, 1600
                            lbcSelection(9).Visible = True
                        Case 3                              'slsp
                            lbcSelection(2).Height = 1600   '3240
                            lbcSelection(2).Visible = True
                            ckcAll.Caption = "All Salespeople"
                            ckcAllGroups.Move 15, lbcSelection(2).Top + lbcSelection(2).Height + 60
                            ckcAllGroups.Caption = "All Sales Sources"
                            ckcAllGroups.Visible = True
                            lbcSelection(9).Move 15, ckcAllGroups.Top + ckcAllGroups.Height + 15, 4380, 1600
                            lbcSelection(9).Visible = True
                        Case 4, 5, 6                           'billing & selling veh , & office/vehicles
                            lbcSelection(6).Height = 3240
                            lbcSelection(6).Visible = True
                            ckcAll.Caption = "All Vehicles"
                            If Index = 5 Then       'airing vehicle only for dual cash posting balancing purposes
                                gPopVehicleGroups RptSelCreditStatus!cbcSet1, tgVehicleSets1(), True
                                edcSet1.Text = "Vehicle Group"
                                cbcSet1.ListIndex = 0
                                'edcSet1.Move 120, plcSelC12.Top + 30 + plcSelC12.Height   '3-17-05 adjust top to insert hard cost option
                                edcSet1.Move 120, ilTop + 30 + ilHeight
                                cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 45
                                edcSet1.Visible = True
                                cbcSet1.Visible = True
                            End If
                        Case 7  '9-16-02
                            'plcSelC6.Visible = False    '2-2-03 turn of air time, ntr or both
                            rbcSelC6(0).Enabled = False
                            rbcSelC6(1).Enabled = False
                            rbcSelC6(2).Enabled = False    '3-17-05 dont hide air time, ntr or both; just disable
                            rbcSelC6(1).Value = True    'default to NTR only
                            lbcSelection(8).Height = 3240
                            lbcSelection(8).Visible = True
                            ckcAll.Caption = "All Item Types"
                        Case 8  '10-17-02  Sales Source
                            lbcSelection(9).Height = 3240
                            ckcAll.Caption = "All Sales Sources"
                            lbcSelection(9).Visible = True
                            gPopVehicleGroups RptSelCreditStatus!cbcSet1, tgVehicleSets1(), True
                            edcSet1.Text = "Vehicle Group"
                            cbcSet1.ListIndex = 0
                            'edcSet1.Move 120, plcSelC12.Top + 30 + plcSelC12.Height       '3-17-05 adjust top to insert the hard cost option
                            edcSet1.Move 120, ilTop + 30 + ilHeight
                            cbcSet1.Move 360 + edcSet1.Width, edcSet1.Top - 45
                            edcSet1.Visible = True
                            cbcSet1.Visible = True
                        Case 9  '11-24-06 Sales Origin
                            'plcSelC8.Move 120, plcSelC12.Top + plcSelC12.Height, 4400, 395
                            plcSelC8.Move 120, ilTop + ilHeight, 4400, 395      '10-28-08 fix feature not shown when inv reg by billing or revenue was implemented
                            smPlcSelC8P = "Include Major Vehicle Totals by -"
                            rbcSelC8(0).Caption = "None"
                            rbcSelC8(0).Move 300, 195, 720
                            rbcSelC8(0).Visible = True
                            If rbcSelC8(0).Value Then
                                rbcSelC8_Click 0
                            Else
                                rbcSelC8(0).Value = True
                            End If
                            rbcSelC8(1).Move 1140, 195, 1250
                            rbcSelC8(1).Caption = "Bill Vehicle"
                            rbcSelC8(1).Visible = True

                            rbcSelC8(2).Move 2460, 195, 1250
                            rbcSelC8(2).Caption = "Air Vehicle"
                            rbcSelC8(2).Visible = True

                            plcSelC8.Visible = True

                            '
                            '   Sales Origins hidden for now, always include all of them
                            '
                            'plcSelC10.Move 120, plcSelC8.Height + plcSelC8.Top + 30
                            plcSelC10.Move 120, ilTop + 30 + ilHeight

                            smPlcSelC10P = "Sales Origins"
                            ckcSelC10(0).Caption = "Local"
                            ckcSelC10(0).Move 1320, 0, 860
                            ckcSelC10(0).Visible = True
                            ckcSelC10(0).Value = vbChecked
                            ckcSelC10(1).Caption = "Natl"
                            ckcSelC10(1).Move 2160, 0, 600
                            ckcSelC10(1).Visible = True
                            ckcSelC10(1).Value = vbChecked
                            ckcSelC10(2).Caption = "Regional"
                            ckcSelC10(2).Move 2940, 0, 1200
                            ckcSelC10(2).Visible = True
                            ckcSelC10(2).Value = vbChecked
                            plcSelC10.Visible = False       '11-25-06 hide for now, include all sales origins

                    End Select
                ElseIf ilListIndex = INV_SUMMARY Then               '6-28-05
                    lbcSelection(1).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(3).Visible = False
                    lbcSelection(4).Visible = False     '10-29-99
                    lbcSelection(5).Visible = False
                    lbcSelection(6).Visible = False
                    lbcSelection(7).Visible = False
                    lbcSelection(8).Visible = False     '9-16-02
                    lbcSelection(9).Visible = False     '10-17-02 Sales Source
                    ckcAll.Visible = True
                    ckcAll.Value = vbUnchecked   'False
                    Select Case Index

                        Case 0                              'advt
                            lbcSelection(5).Height = 3240
                            lbcSelection(5).Visible = True
                            ckcAll.Caption = "All Advertisers"
                        Case 1                              'agy
                            lbcSelection(1).Height = 3240
                            lbcSelection(1).Visible = True
                            ckcAll.Caption = "All Agencies"
                        Case 2                              'slsp
                            lbcSelection(2).Height = 3240
                            lbcSelection(2).Visible = True
                            ckcAll.Caption = "All Salespeople"
                    End Select

                End If
            Case COLLECTIONSJOB
                Select Case ilListIndex
                    Case COLL_AGEPAYEE, COLL_AGESLSP, COLL_AGEVEHICLE  'Ageing
                        If Index = 0 Or Index = 1 Then           'detail/trans type version allows transaction comments to be show
                            plcSelC9.Visible = True
                            ckcTrans.Enabled = True
                        Else
                            ckcTrans.Value = False
                            ckcTrans.Enabled = False
                        End If
                    Case COLL_AGEOWNER, COLL_AGESS, COLL_AGEPRODUCER   'Ageing
                        If Index = 0 Then           'detail/trans type version allows transaction comments to be show
                            plcSelC9.Visible = True
                            ckcTrans.Enabled = True
                        Else
                            ckcTrans.Value = False
                            ckcTrans.Enabled = False
                        End If
                    Case COLL_CASH
                'If (ilListIndex = COLL_CASH) Then        'cash receipts
                    If Index = 0 Then                       'cash receipts by deposit date (default to last month of final inv.)
                        gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr
                        If Trim$(slStr) <> "" Then
                            edcSelCFrom.Text = gIncOneDay(slStr)
                        Else
                            edcSelCFrom.Text = ""
                        End If
                        gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
                        If Trim$(slStr) <> "" Then
                            edcSelCTo.Text = slStr
                        Else
                            edcSelCTo.Text = ""
                        End If
                    Else                                    'cash receipts by entry date   (leave start date blank, the date entred should be one day past last time run
                                                            'default end date to todays date)
                        edcSelCFrom.Text = ""                 'leave start dat blank
                        llTodayDate = gDateValue(gNow())
                        slStr = Format(llTodayDate, "m/d/yy")
                        edcSelCTo.Text = slStr
                        edcSelCFrom.Text = slStr            '2-24-04
                    End If

                    Case COLL_STATEMENT     '9-17-03 removed code (no questions referenced rbcselcselct for statements
                'ElseIf (ilListIndex = 5) Then

                    Case COLL_PAYHISTORY
                'ElseIf ilListIndex = COLL_PAYHISTORY Then       '9-11-03 use all direct adv & agencies (not just agencies)
                    Select Case Index
                        Case 0  'Agency
                            lbcSelection(0).Visible = False
                            lbcSelection(2).Visible = True
                            lbcSelection(6).Visible = False
                            ckcAll.Caption = "All Agencies and Advertisers"
                        Case 1  'Advertiser
                            lbcSelection(2).Visible = False
                            lbcSelection(0).Visible = True
                            lbcSelection(6).Visible = False
                            ckcAll.Caption = "All Advertisers"
                        End Select

                 Case COLL_DISTRIBUTE
                'ElseIf ilListIndex = COLL_DISTRIBUTE Then
                    If Index = 2 Then
                        lbcSelection(3).Visible = True
                        ckcAll.Caption = "All Participants"
                        ckcAll.Visible = True
                        ckcAll.Value = vbUnchecked   'False
                        plcSelC9.Visible = True                 '2-5-15 option to skip page each new vehicle within participant
                        ckcTrans.Visible = True
                        ckcTrans.Value = vbUnchecked
                    Else
                        lbcSelection(3).Visible = False
                        ckcAll.Visible = False
                        ckcAll.Value = vbChecked   'True
                        plcSelC9.Visible = False
                        ckcTrans.Visible = False
                        ckcTrans.Value = vbUnchecked
                    End If

                 Case COLL_MERCHANT
                'ElseIf ilListIndex = COLL_MERCHANT Then
                    If Index = 0 Then
                        lbcSelection(6).Visible = True
                        lbcSelection(0).Visible = False
                        ckcAll.Caption = "All Vehicles"
                    Else
                        lbcSelection(0).Visible = True
                        lbcSelection(6).Visible = False
                        ckcAll.Caption = "All Advertisers"
                    End If
                    ckcAll.Visible = True
                    ckcAll.Value = vbUnchecked   'False

                 Case COLL_CASHSUM
                'D.S. 08/14/01
                'ElseIf ilListIndex = COLL_CASHSUM Then
                    ilListIndex = ilListIndex
                    If rbcSelCSelect(1).Value = True Then
                        sgSOCodeTag = ""                'init, the list box is also used for advt, and coming back into this report shows advt
                                                    'advt instead of the sales office.  Need to repopulate
                        ckcAll.Caption = "All Sales Offices"
                        ckcAll.Visible = True
                        mSalesOfficePop lbcSelection(0)
                        lbcSelection(6).Visible = False
                        lbcSelection(0).Visible = True
                    Else
                        ckcAll.Visible = True
                        ckcAll.Caption = "All Vehicles"
                        lbcSelection(0).Visible = False
                        lbcSelection(6).Visible = True
                    End If

                 Case COLL_ACCTHIST
                'ElseIf (ilListIndex = COLL_ACCTHIST) Then
                    ckcAll.Visible = True
                    Select Case Index
                        Case 0  'Advertiser
                            lbcSelection(0).Visible = True
                            lbcSelection(2).Visible = False
                            ckcAll.Caption = "All Advertisers"
                        Case 1  'Agency
                            lbcSelection(0).Visible = False
                            lbcSelection(2).Visible = True
                            ckcAll.Caption = "All Agencies"
                        End Select
                'End If
                End Select

            Case COPYJOB
                'If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then        '7-1-04
                '    ilListIndex = ilListIndex + 1
                'End If
                If (ilListIndex = 0) Or (ilListIndex = 1) Then
                    If Index = 0 Then
                        rbcSelCInclude(0).Value = True   'Yes for All
                        plcSelC2.Enabled = False
                    Else
                        plcSelC2.Enabled = True
                        rbcSelCInclude(1).Value = True   'No
                    End If
                End If
                If ilListIndex = 2 Then                         'contracts missing copy
                    If rbcSelCSelect(1).Value = True Then
                        plcSelC7.Visible = True
                        ckcSelC7.Visible = True
                    Else
                        plcSelC7.Visible = False
                        ckcSelC7.Visible = False
                        ckcSelC7.Value = False
                    End If
                End If
                If ilListIndex = 11 Or ilListIndex = 13 Then    'copy play list by isci or vehicle
                    'If rbcSelCSelect(0).Value = True Then
                    If ilListIndex = 13 Then                    'vehicle
                        lacSelCTo1.Visible = False
                        edcSelCTo1.Visible = False
                        ckcAll.Caption = "All Vehicles"
                        ckcAll.Visible = True
                        lbcSelection(2).Visible = True
                        rbcSelC8(1).Value = True                'default to generic + split, question isnt asked
                    Else                                        'playlist by ISCI
                        ckcAll.Visible = True
                        lbcSelection(2).Visible = True
                        'lacSelCTo1.Top = 1050              'changed to cover Vehicle vs ISCI option since
                                                            'the vehicle option has been removed
                        lacSelCTo1.Top = plcSelC1.Top
                        lacSelCTo1.Width = 2000
                        lacSelCTo1.Left = 140
                        lacSelCTo1.Caption = "Descrip."
                        lacSelCTo1.Visible = True
                        edcSelCTo1.Text = ""
                        edcSelCTo1.MaxLength = 32
                        'edcSelCTo1.Top = 1050
                        edcSelCTo1.Top = plcSelC1.Top
                        edcSelCTo1.Left = 1050
                        edcSelCTo1.Width = 3400
                        edcSelCTo1.Visible = True
                        
                        plcSelC6.Move 120, edcSelCTo1.Top + edcSelCTo1.Height + 60, 1920
                        ckcSelC6Add(0).Move 0, 0, 1920
                        ckcSelC6Add(0).Caption = "Show Live Script"
                        ckcSelC6Add(0).Visible = True
                        plcSelC6.Visible = True
                                               
                        plcSelC8.Move 120, plcSelC6.Top + plcSelC6.Height, 3840, 480
                        smPlcSelC8P = "Show"
                        rbcSelC8(0).Caption = "Generic"
                        rbcSelC8(0).Move 600, 0, 960
                        rbcSelC8(0).Value = True
                        rbcSelC8(0).Visible = True
                        rbcSelC8(1).Caption = "Generic + Split Copy"
                        rbcSelC8(1).Move 1680, 0, 2400
                        rbcSelC8(1).Visible = True
                        rbcSelC8(2).Caption = "Split Copy Only"
                        rbcSelC8(2).Move 600, 240, 2040
                        rbcSelC8(2).Visible = True
                        If (Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY Then      'split copy feature on
                            plcSelC8.Visible = True
                            plcSelC7.Move 120, plcSelC8.Top + plcSelC8.Height + 30, 3000

                        Else
                            plcSelC8.Visible = False
                            rbcSelC8(0).Visible = False
                            rbcSelC8(1).Visible = False
                            rbcSelC8(2).Visible = False
                            plcSelC7.Move 120, plcSelC6.Top + plcSelC6.Height + 30, 3000
                        End If
                        
                        'plcSelC7.Move 120, plcSelC8.Top + plcSelC8.Height + 30, 3000
                        ckcSelC7.Move 0, 0, 3000
                        ckcSelC7.Caption = "Include Vehicles Ordered"           '7-23-12
                        ckcSelC7.Visible = True
                        plcSelC7.Visible = True
                        ckcSelC7.Value = vbUnchecked
                        '6-27-13 option to show Rotation dates
                        plcSelC9.Move 120, plcSelC7.Top + plcSelC7.Height + 30, 3000
                        ckcTrans.Move 0, 0, 3000
                        ckcTrans.Caption = "Include Rotation Dates"           '7-23-12
                        ckcTrans.Visible = True
                        plcSelC9.Visible = True
                        ckcTrans.Value = vbChecked
                    End If
                End If
                If ilListIndex = 14 Then        'copy playlist by advt
                    ckcAll.Caption = "All Advertisers"
                    ckcAll.Visible = True
                    lbcSelection(0).Visible = True
                    'force all vehicles to be selected
                    llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, True, llRg)
                    rbcSelC8(1).Value = True                'default to generic + split, question isnt asked
                End If
        End Select
        mSetCommands
    End If
End Sub
Private Sub rbcType_Click(Index As Integer)
    If Index = 0 Then       'detail
        lbcSelection(0).Visible = True
        lacAsOfDate.Visible = False
        edcAsOfDate.Visible = False
        lacFrom.Visible = False
        lacTo.Visible = False
        plcRepInv.Visible = False
        frcOption.Enabled = True
        pbcOption.Visible = True
        pbcSelB.Visible = True
        cbcFrom.Visible = False
        cbcTo.Visible = False
    Else                    'summary
        cbcFrom.Visible = True
        cbcTo.Visible = True
        lacAsOfDate.Visible = True
        edcAsOfDate.Visible = True
        lacFrom.Visible = True
        lacTo.Visible = True
        plcRepInv.Visible = True
        frcOption.Enabled = True
        pbcOption.Visible = False
        pbcSelB.Visible = True
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    'plcSelC3.Print "Zone"
    plcSelC3.Print smPlcSelC3P
End Sub
Private Sub plcSel1_Paint()
    plcSel1.CurrentX = 0
    plcSel1.CurrentY = 0
    'plcSelC3.Print "Zone"
    plcSel1.Print smPlcSel1P
End Sub
Private Sub plcSel2_Paint()
    plcSel2.CurrentX = 0
    plcSel2.CurrentY = 0
    'plcSel2.Print "Zone"
    plcSel2.Print smPlcSel2P
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
    'plcSelC1.Print "Select"
    plcSelC1.Print smPlcSelC1P
End Sub
Private Sub plcSelC4_Paint()
    plcSelC4.CurrentX = 0
    plcSelC4.CurrentY = 0
    'plcSelC4.Print "Option"
    plcSelC4.Print smPlcSelC4P
End Sub
'
'
'       For most Collection Job reports, ask Air Time, NTR or BOTH
'       9-17-02
Public Sub mAskAirTimeNTR(ilTop As Integer, Optional RadioOrCheck As Integer = ChooseRadio)
'6/02/08 new option to make check boxes instead of radio buttons
Dim ilHideRadio As Integer
If RadioOrCheck = ChooseRadio Then
    rbcSelC6(0).Caption = "Air Time"
    rbcSelC6(0).Move 480, 0, 1080
    rbcSelC6(0).Value = True
    rbcSelC6(1).Caption = "NTR"
    rbcSelC6(1).Move 1560, 0, 700
    rbcSelC6(2).Caption = "Both"
    rbcSelC6(2).Move 2260, 0, 720
Else
    For ilHideRadio = rbcSelC6.LBound To rbcSelC6.UBound Step 1
        rbcSelC6(ilHideRadio).Visible = False
    Next ilHideRadio
    Load ckcSelC6Add(Airtime)
    ckcSelC6Add(Airtime).Caption = "Air Time"
    ckcSelC6Add(Airtime).Move 480, 0, 1080
    ckcSelC6Add(Airtime).Value = 1
    ckcSelC6Add(Airtime).Visible = True
    Load ckcSelC6Add(NTR)
    ckcSelC6Add(NTR).Caption = "NTR"
    ckcSelC6Add(NTR).Move 1560, 0, 700
    ckcSelC6Add(NTR).Value = 1
    ckcSelC6Add(NTR).Visible = True
    ckcSelC6Add(HardCost).Caption = "Hard Cost"
    ckcSelC6Add(HardCost).Move 2260, 0, 1280
    ckcSelC6Add(HardCost).Value = 0
    ckcSelC6Add(HardCost).Visible = True
End If
plcSelC6.Move 120, ilTop
smPlcSelC6P = "For"
plcSelC6.Visible = True

End Sub




Public Sub mAskPOApply()

    lacSelCFrom.Caption = "Date Entered -Start"
    lacSelCFrom1.Caption = "End"
    lacSelCTo.Caption = "Trans Date -Start"
    lacSelCTo1.Caption = "End"
    lacSelCFrom.Move 30, 160, 1680      'date entered- start
    edcSelCFrom.Move 1680, 120, 1080
    lacSelCFrom1.Move 2880, 160, 360    'date entered- end
    edcSelCFrom1.Move 3240, 120, 1080

    lacSelCTo.Move 30, edcSelCFrom.Top + edcSelCFrom.Height + 60, 1680    'tran date - start
    edcSelCTo.Move 1680, edcSelCFrom.Top + edcSelCFrom.Height + 30, 1080
    lacSelCTo1.Move 2880, lacSelCTo.Top, 360    'tran date - end
    edcSelCTo1.Move 3240, edcSelCTo.Top, 1080

    ckcAll.Caption = "All Agencies and Advertisers"
    gPopVehicleGroups RptSelCreditStatus!cbcSet1, tgVehicleSets1(), True
    edcSet1.Text = "Vehicle Group"
    cbcSet1.ListIndex = 0
    edcSet1.Move 30, edcSelCTo.Top + edcSelCTo.Height + 60
    cbcSet1.Move 1680, edcSet1.Top - 45
    edcSet1.Visible = True
    cbcSet1.Visible = True

    lacSelCFrom.Visible = True
    lacSelCFrom1.Visible = True
    lacSelCTo.Visible = True
    lacSelCTo1.Visible = True
    edcSelCFrom.Visible = True
    edcSelCFrom1.Visible = True
    edcSelCTo.Visible = True
    edcSelCTo1.Visible = True
    lbcSelection(2).Visible = True
End Sub

Public Sub mAskMMYY(DateIn As control)
Dim slStr As String
Dim slMonth As String
Dim slYear As String

       gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
       If Trim$(slStr) <> "" Then
           DateIn.Text = slStr
       Else
           DateIn.Text = ""
       End If

       If Trim$(slStr) <> "" Then
           If tgSpf.sRRP = "C" Then    'Calendar
               slStr = gObtainEndCal(slStr)
               slMonth = Month(gDateValue(slStr))
               slYear = right$(Year(gDateValue(slStr)), 2)
           ElseIf tgSpf.sRRP = "F" Then 'Corporate
               slStr = Format$(gDateValue(slStr) - 15, "m/d/yy")
               slMonth = Month(gDateValue(slStr))
               slYear = right$(Year(gDateValue(slStr)), 2)
           Else
               slStr = gObtainEndStd(slStr)
               slMonth = Month(gDateValue(slStr))
               slYear = right$(Year(gDateValue(slStr)), 2)
           End If
           DateIn.Text = slMonth & "/" & slYear
       Else
           DateIn.Text = ""
       End If
End Sub
'
'       Ask to Show Transaction Comments
'
'       <input> Picture box control
'               check box control
'               ilTopLoc - top location of control
Public Sub mAskShowComments(plcPicture As control, ckcCheck As control, ilTopLoc As Integer)
    plcPicture.Move 120, ilTopLoc, 3360
    smPlcSelC9P = "Show transaction comments"
    ckcCheck.Move 0, 0, 3360
    plcPicture.Visible = True
    ckcCheck.Visible = True
End Sub
'
'
'       Selectivity for Credit/Debit Memo report
'
'       Creation (entered date)
'       Invoice Date (which is the same as trans date)
'       Single Contract
'       Payees
'       10-8-03
Public Sub mAskCreditMemo()
    mAgyAdvtPop lbcSelection(7)    'Called to initialize agy and direct advertiser (statements)
    If imTerminate Then
        Exit Sub
    End If

    lacSelCFrom.Caption = "Date Entered -Start"
    lacSelCFrom1.Caption = "End"
    lacSelCTo.Caption = "Invoice Date -Start"
    lacSelCTo1.Caption = "End"
    lacSelCFrom.Move 30, 160, 1680      'date entered- start
    edcSelCFrom.Move 1680, 120, 1080
    lacSelCFrom1.Move 2880, 160, 360    'date entered- end
    edcSelCFrom1.Move 3240, 120, 1080

    lacSelCTo.Move 30, edcSelCFrom.Top + edcSelCFrom.Height + 60, 1680    'tran date - start
    edcSelCTo.Move 1680, edcSelCFrom.Top + edcSelCFrom.Height + 30, 1080
    lacSelCTo1.Move 2880, lacSelCTo.Top, 360    'tran date - end
    edcSelCTo1.Move 3240, edcSelCTo.Top, 1080
    edcSelCTo.MaxLength = 8
    edcSelCTo1.MaxLength = 8
    edcSelCFrom.MaxLength = 8
    edcSelCFrom1.MaxLength = 8
    lacCheck.Caption = "Contract #"
    lacCheck.Move 30, edcSelCTo.Top + edcSelCTo.Height + 60
    edcCheck.Move 1200, edcSelCTo.Top + edcSelCTo.Height + 30
    lacCheck.Visible = True
    edcCheck.Visible = True
    ckcAll.Caption = "All Agencies and Advertisers"
    lbcSelection(7).Visible = True
    lacSelCFrom.Visible = True
    lacSelCFrom1.Visible = True
    lacSelCTo.Visible = True
    lacSelCTo1.Visible = True
    edcSelCFrom.Visible = True
    edcSelCFrom1.Visible = True
    edcSelCTo.Visible = True
    edcSelCTo1.Visible = True
    'lbcSelection(2).Visible = True
    ckcAll.Visible = True
    pbcSelC.Visible = True
    pbcOption.Visible = True
End Sub

'
'
'           mCollectionSelectivity - Format all Collection selectivity screens
'           Created: 10-21-03
'
Public Sub mCollectionSelectivity()
Dim ilListIndex As Integer
Dim slStr As String
    ilListIndex = lbcRptType.ListIndex
    edcSelCTo.MaxLength = 10    '8   5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
    edcSelCFrom.MaxLength = 10 '8    5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
    edcSelCFrom1.MaxLength = 10
    edcSelCTo1.MaxLength = 10
    edcSelCTo.Width = 1170
    edcSelCFrom.Width = 1170
    lacSelCFrom1.Visible = False
    edcSelCFrom1.Visible = False
    pbcSelA.Visible = False
    plcSelC3.Visible = False
    plcSelC2.Visible = False
    plcSelC1.Visible = False
    rbcSelCInclude(2).Visible = False
    rbcSelCInclude(3).Visible = False
    edcSelCFrom.Move 1050, 30
    lacSelCFrom.Move 120, 75
    edcSelCTo.Move 1050, 345
    lacSelCTo.Move 120, 390
    lacSelCTo1.Visible = False
    edcSelCTo1.Visible = False
    lbcSelection(0).Visible = False         'advt
    lbcSelection(1).Visible = False         'agy
    lbcSelection(2).Visible = False         'advt/agy
    lbcSelection(3).Visible = False         'owners
    lbcSelection(5).Visible = False         'slsp
    lbcSelection(6).Visible = False         'vehicle
    ckcAll.Visible = True
    Select Case ilListIndex
        Case COLL_PAYHISTORY  'payment History
            lacSelCFrom.Width = 900
            lacSelCFrom.Caption = "From Date"
            lacSelCFrom.Visible = True
            edcSelCFrom.Left = 1050
            gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr
            If Trim$(slStr) <> "" Then
                edcSelCFrom.Text = gIncOneDay(slStr)
            Else
                edcSelCFrom.Text = ""
            End If
            edcSelCFrom.Visible = True
            lacSelCTo.Width = 900
            lacSelCTo.Caption = "To Date"
            lacSelCTo.Visible = True
            edcSelCTo.Left = 1050
            gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
            If Trim$(slStr) <> "" Then
                edcSelCTo.Text = slStr
            Else
                edcSelCTo.Text = ""
            End If
            edcSelCTo.Visible = True
            'plcSelC1.Visible = False
            plcSelC1.Move lacSelCTo.Left, edcSelCTo.Top + edcSelCTo.Height
            'plcSelC1.Caption = "Select"
            smPlcSelC1P = "Select"
            rbcSelCSelect(0).Left = 630
            rbcSelCSelect(0).Width = 980
            rbcSelCSelect(0).Caption = "Agency"
            rbcSelCSelect(0).Visible = True
            rbcSelCSelect(1).Left = 1630
            rbcSelCSelect(1).Width = 1220
            rbcSelCSelect(1).Caption = "Advertiser"
            rbcSelCSelect(1).Visible = True

            rbcSelCSelect(2).Visible = False        'was vehicle, no longer required
            If rbcSelCSelect(0).Value Then
                rbcSelCSelect_click 0   ', True
            Else
                rbcSelCSelect(0).Value = True   'Agency
            End If
            mAskCashTrMercProm 900, False               'dont show hard cost option
            mAskAirTimeNTR plcSelC2.Top + 30 + plcSelC2.Height

            '4-28-03 Determine Detail or Summary
            plcSelC8.Move 120, plcSelC6.Top + plcSelC6.Height
            smPlcSelC8P = "Totals by "
            rbcSelC8(0).Move 960, 0, 840
            rbcSelC8(1).Move 1800, 0, 1200
            rbcSelC8(0).Caption = "Detail"
            rbcSelC8(1).Caption = "Summary"
            rbcSelC8(0).Value = True    'default to detail
            rbcSelC8(2).Visible = False
            plcSelC8.Visible = True

            plcSelC7.Move plcSelC8.Left, plcSelC8.Top + plcSelC8.Height, 3360
            ckcSelC7.Caption = "Show transaction comments"
            ckcSelC7.Move 0, 0, 3360
            plcSelC7.Visible = True
            ckcSelC7.Visible = True

            plcSelC1.Visible = True
            plcSelC2.Visible = True
            pbcSelC.Visible = True
            pbcOption.Visible = True
       ' Case 1, 2, 3, 13, 14    'Ageing    '2-10-00
        Case COLL_AGEPAYEE, COLL_AGESLSP, COLL_AGEVEHICLE, COLL_AGEOWNER, COLL_AGESS, COLL_AGEPRODUCER   'Ageing
            ckcAll.Visible = True
            lacSelCTo.Visible = False
            edcSelCTo.Visible = False
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            lacSelCFrom.Width = 3420
            lacSelCFrom.Caption = "Latest transaction date to include"
            lacSelCFrom.Visible = True
            edcSelCFrom.Left = 3000
            gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
            If Trim$(slStr) <> "" Then
                edcSelCFrom.Text = slStr
            Else
                edcSelCFrom.Text = ""
            End If
            edcSelCFrom.Visible = True
            lacSelCFrom1.Move 120, 345, 3000
            lacSelCFrom1.Caption = "MM/YY to use as current month"
            lacSelCFrom1.Visible = True

            mAskMMYY edcSelCFrom1

            edcSelCFrom1.Move 3000, 345, 720
            lacSelCFrom1.Top = edcSelCFrom1.Top + 30
            'edcSelCTo.Visible = True
            edcSelCFrom1.Visible = True
            'plcSelC1.Move 120, edcSelCTo.Top + edcSelCTo.Height

            lacSelCTo.Caption = "Include Ageing MM/YY- Earliest"
            lacSelCTo.Move 120, edcSelCFrom1.Top + edcSelCFrom1.Height + 60, 3240
            edcSelCTo.Move 3000, lacSelCTo.Top - 30, 720
            lacSelCTo1.Caption = "Latest"
            lacSelCTo1.Move 2070, lacSelCTo.Top + lacSelCTo.Height + 90, 600
            edcSelCTo1.Move 3000, lacSelCTo1.Top - 15, 720
            lacSelCTo.Visible = True
            edcSelCTo.Visible = True
            lacSelCTo1.Visible = True
            edcSelCTo1.Visible = True
            'plcSelC1.Move 120, edcSelCFrom1.Top + edcSelCFrom1.Height
            plcSelC1.Move 120, edcSelCTo1.Top + edcSelCTo1.Height
            plcSelC1.Height = 240
            'plcSelC1.Caption = "Totals by "
            smPlcSelC1P = "Totals by "
            rbcSelCSelect(2).Visible = False
            rbcSelCSelect(0).Left = 900    '330
            rbcSelCSelect(0).Width = 780    '840
            rbcSelCSelect(0).Caption = "Detail"
            rbcSelCSelect(0).Visible = True
            If ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGESLSP Or ilListIndex = COLL_AGEVEHICLE Then '7-11-02
                rbcSelCSelect(1).Left = 1740    '1190
                rbcSelCSelect(1).Width = 1160
                rbcSelCSelect(1).Caption = "Tran Type"
                rbcSelCSelect(1).Visible = True
                rbcSelCSelect(2).Left = 2880    '1190
                rbcSelCSelect(2).Width = 1560
                rbcSelCSelect(2).Caption = "Invoice"
                rbcSelCSelect(2).Visible = True
                rbcSelCSelect(3).Caption = "Advertiser"
                rbcSelCSelect(3).Visible = True
                rbcSelCSelect(3).Move 900, 195
                rbcSelCSelect(3).Width = 1200
                plcSelC1.Height = 450

            Else
                rbcSelCSelect(1).Left = 1740    '1190
                rbcSelCSelect(1).Width = 1480
                rbcSelCSelect(1).Caption = "Summary"
                rbcSelCSelect(1).Visible = True
            End If
            If rbcSelCSelect(0).Value Then
                rbcSelCSelect_click 0   ', True
            Else
                rbcSelCSelect(0).Value = True   'Detail
            End If
            plcSelC1.Visible = True
            ' dan M changed to radio buttons 6-03-08
            mAskCashTrMercProm plcSelC1.Top + plcSelC1.Height, False    'hide hard cost option with false

            mAskAirTimeNTR plcSelC2.Top + 30 + plcSelC2.Height, ChooseCheck 'choosecheck or chooseradio

            'mAskCashTrMercProm 1095
            If ilListIndex = COLL_AGEPAYEE Then             'payee
                lbcSelection(2).Visible = True
                ckcAll.Caption = "All Agencies and Advertisers"
                 '7-16-02 add vehicle group sorting
                gPopVehicleGroups RptSelCreditStatus!cbcSet1, tgVehicleSets1(), True
                edcSet1.Text = "Vehicle Group"
                cbcSet1.ListIndex = 0
                edcSet1.Move 120, plcSelC6.Top + 30 + plcSelC6.Height
                cbcSet1.Move 360 + edcSet1.Width, edcSet1.Top - 45
                edcSet1.Visible = True
                cbcSet1.Visible = True
                plcSelC5.Move 120, edcSet1.Top + edcSet1.Height + 60, 3840, 270
                ckcSelC5(0).Value = vbUnchecked
                ckcSelC5(0).Caption = "Segregate 'In Collections'"
                ckcSelC5(0).Move 0, 0, 3840
                ckcSelC5(0).Visible = True
                plcSelC5.Visible = True

                plcSelC7.Move 120, plcSelC5.Top + plcSelC5.Height, 3840, 270
                ckcSelC7.Value = vbUnchecked
                ckcSelC7.Caption = "Separate by Sales Source as major"
                ckcSelC7.Move 0, 0, 4440
                ckcSelC7.Visible = True
                plcSelC7.Visible = True
                mAskShowComments plcSelC9, ckcTrans, plcSelC7.Top + plcSelC7.Height '9-17-03
            ElseIf ilListIndex = COLL_AGESLSP Then         'slsp
                mSalesOfficePop lbcSelection(7)         'sales office for sales option
                ckcAll.Visible = True
                ckcAll.Caption = "All Salespeople"
                ckcAll.Left = 0
                lbcSelection(5).Height = 1500
                lbcSelection(7).Height = 1500
                ckcAllGroups.Caption = "All Sales Offices"
                ckcAllGroups.Move ckcAll.Left, lbcSelection(5).Top + lbcSelection(5).Height + 60
                ckcAllGroups.Visible = True
                lbcSelection(7).Move lbcSelection(5).Left, ckcAllGroups.Top + ckcAllGroups.Height + 60, 4365, 1500
                lbcSelection(7).Visible = True
                lbcSelection(5).Visible = True
                plcSelC5.Move 120, plcSelC6.Top + plcSelC6.Height, 3840
                ckcSelC5(0).Value = vbUnchecked
                ckcSelC5(0).Caption = "Segregate 'In Collections'"
                'ckcSelC5(0).Width = 3840
                ckcSelC5(0).Move 0, 0, 3840
                ckcSelC5(0).Visible = True
                plcSelC5.Visible = True
                mAskShowComments plcSelC9, ckcTrans, plcSelC5.Top + plcSelC5.Height

            ElseIf ilListIndex = COLL_AGEVEHICLE Then                               'vehicle
                lbcSelection(6).Visible = True
                ckcAll.Caption = "All Vehicles"

                 '7-16-02 add vehicle group sorting
                gPopVehicleGroups RptSelCreditStatus!cbcSet1, tgVehicleSets1(), True
                edcSet1.Text = "Vehicle Group"
                cbcSet1.ListIndex = 0
                edcSet1.Move 120, plcSelC6.Top + 30 + plcSelC6.Height
                cbcSet1.Move edcSet1.Width + 120, edcSet1.Top - 45
                edcSet1.Visible = True
                cbcSet1.Visible = True
                '5-20-09 change location/wording for Vehicle Group share.  Need to
                'make room for new question to generate 1-yr ageing
                ckcSelC3(0).Caption = "Share"

'                ckcSelC3(0).Caption = "Include vehicle group share"
                ckcSelC3(0).Visible = True
                ckcSelC3(0).Enabled = False
                ckcSelC3(0).Value = vbUnchecked 'False
                ckcSelC3(0).Move 0, 0, 3360
'                plcSelC3.Move 120, edcSet1.Top + edcSet1.Height, 3000
                plcSelC3.Move cbcSet1.Left + cbcSet1.Width + 240, edcSet1.Top, 840

                smPlcSelC3P = ""
                plcSelC3.Visible = True

                mAskShowComments plcSelC9, ckcTrans, plcSelC3.Top + plcSelC3.Height
                'mAskShowComments plcSelC9, ckcTrans, edcSet1.Top + edcSet1.Height
                plcSelC10.Move 120, plcSelC9.Top + plcSelC9.Height - 60
                ckcSelC10(0).Move 0, 0, 3000
                ckcSelC10(0).Caption = "Include Salespeople Sub-totals"
                ckcSelC10(1).Caption = "New Page"
                ckcSelC10(1).Move 3000, 0, 1200
                ckcSelC10(1).Visible = True
                ckcSelC10(0).Visible = True
                plcSelC10.Visible = True

                plcSelC11.Move 120, plcSelC10.Top + plcSelC10.Height
                ckcSelC11(0).Move 0, 0, 3000
                ckcSelC11(0).Caption = "Extended Ageing columns"
                ckcSelC11(0).Visible = True
                smPlcSelC11P = ""
                plcSelC11.Visible = True


            ElseIf ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPRODUCER Then            'participants or sales source option
                lbcSelection(3).Visible = True
                ckcAll.Caption = "All Participants"

                '1-20-06
                lbcSelection(9).Visible = True
                ckcAllGroups.Caption = "All Sales Sources"
                ckcAllGroups.Visible = True
                lbcSelection(3).Height = 1500   'lbcSelection(3).Height / 2
                ckcAllGroups.Move ckcAll.Left, lbcSelection(3).Top + lbcSelection(3).Height + 60
                lbcSelection(9).Move lbcSelection(3).Left, ckcAllGroups.Top + ckcAllGroups.Height + 30, lbcSelection(9).Width, 1500 'lbcSelection(3).Height

                lbcSelection(9).Visible = True

                mAskShowComments plcSelC9, ckcTrans, plcSelC6.Top + plcSelC6.Height

            End If

            pbcOption.Visible = True
            pbcSelC.Visible = True
            plcSelC2.Visible = True
        Case COLL_DELINQUENT  'Delinguent
            pbcOption.Visible = False
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            lacSelCTo.Visible = False
            edcSelCTo.Visible = False
            lacSelCFrom.Width = 1830
            lacSelCFrom.Caption = "Latest Cash Date"
            lacSelCFrom.Visible = True
            gUnpackDate tgSpf.iRLastPay(0), tgSpf.iRLastPay(1), slStr
            If slStr <> "" Then
                edcSelCFrom.Text = slStr
            Else
                edcSelCFrom.Text = Format$(gNow(), "m/d/yy")
            End If
            edcSelCFrom.Left = 1950
            edcSelCFrom.Visible = True
            lacSelCTo.Width = 1830
            lacSelCTo.Caption = "End of Current Period"
            lacSelCTo.Visible = True
            gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
            If slStr <> "" Then
                edcSelCTo.Text = slStr
            Else
                edcSelCTo.Text = "" 'Format$(Now, "m/d/yy")
            End If
            edcSelCTo.Left = 1950
            edcSelCTo.Visible = True
            mAskAirTimeNTR edcSelCTo.Top + edcSelCTo.Height + 30

            pbcSelC.Visible = True
        Case COLL_STATEMENT  'Statements

            plcSelC1.Visible = False
            plcSelC2.Visible = False
            lacSelCFrom.Width = 1560
            lacSelCFrom.Caption = "Latest Cash Date"
            lacSelCFrom.Visible = True
            gUnpackDate tgSpf.iRLastPay(0), tgSpf.iRLastPay(1), slStr
            If slStr <> "" Then
                edcSelCFrom.Text = slStr
            Else
                edcSelCFrom.Text = Format$(gNow(), "m/d/yy")
            End If
            edcSelCFrom.Left = 1710
            edcSelCFrom.Visible = True
            lacSelCTo.Width = 2760  '1560
            lacSelCTo.Caption = "Latest Billing Date to Include"
            lacSelCTo.Visible = True
            edcSelCTo.Left = 2640       '1710
            gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr
            If Trim$(slStr) <> "" Then
                edcSelCTo.Text = slStr
            Else
                edcSelCTo.Text = ""
            End If
            edcSelCTo.Visible = True
            rbcSelCSelect(2).Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = True
            lbcSelection(6).Visible = False
            ckcAll.Caption = "All Agencies and Advertisers"


            lacSelCTo1.Move 120, lacSelCTo.Top + lacSelCTo.Height + 60, 3000
            lacSelCTo1.Caption = "MM/YY to use as current month"
            lacSelCTo1.Visible = True
            mAskMMYY edcSelCTo1

            edcSelCTo1.Move 3000, lacSelCTo1.Top, 720
            lacSelCTo1.Top = edcSelCTo1.Top + 30
            edcSelCTo1.Visible = True


            '8-2-00  Option for detail or summary (1 line pertran type/date/inv#)
            '6-6-02 change question to use plcSelCInclude (vs plcSelC4)

            plcSelC2.Top = edcSelCTo1.Top + edcSelCTo1.Height + 30
            'plcSelC2.Caption = "Show"
            rbcSelCInclude(0).Caption = "Detail"
            rbcSelCInclude(0).Move 750, 0, 840
            rbcSelCInclude(1).Caption = "Tran Type"
            rbcSelCInclude(1).Move 1590, 0, 1200
            If rbcSelCInclude(1).Value Then
                rbcSelCInclude_Click 1
            Else
                rbcSelCInclude(1).Value = True
            End If
            rbcSelCInclude(2).Caption = "Invoice"
            rbcSelCInclude(2).Visible = True
            rbcSelCInclude(2).Move 2790, 0, 960
            rbcSelCInclude(3).Visible = False
            plcSelC2.Visible = True

            pbcSelC.Visible = True
            pbcOption.Visible = True

            '5-14-08  option to see tax as gross or not
            ckcSelC10(0).Visible = True
            ckcSelC10(0).Move 0, 0, 3000
            ckcSelC10(0).Caption = "Embed Taxes in Detail"
            plcSelC10.Visible = True
            plcSelC10.Top = plcSelC2.Top + plcSelC2.Height + 30
            
            '3-18-10 option to use an override Term notation on Statement vs the
            'assigned one with each advt or agy
            lacCheck.Caption = "Select [Use Assigned] to use Terms assigned to each agency, or select a new Term for this report"
            lacCheck.Move 120, plcSelC10.Top + plcSelC10.Height + 60, 4200, 420
            lacCheck.Font = "Arial"
            lacCheck.FontSize = 8
        
            lacCheck.Visible = True
            gMnfTermsPop RptSelCreditStatus, cbcSet1
            cbcSet1.Move 120, lacCheck.Top + lacCheck.Height, 2000
            cbcSet1.Visible = True


        Case COLL_CASH  'Cash Receipts
            mSalesOfficePop lbcSelection(7)         'sales office for sales option
            plcSelC1.Visible = False
            lacSelCFrom.Width = 1020
            lacSelCFrom.Caption = "Dates- From"
            lacSelCFrom.Visible = True
            edcSelCFrom.Text = ""
            edcSelCFrom.Left = 1320
            edcSelCFrom.Visible = True

            lacSelCTo.Width = 360
            lacSelCTo.Caption = "To"
            lacSelCTo.Move 2760, lacSelCFrom.Top
            lacSelCTo.Visible = True
            edcSelCTo.Text = ""

            edcSelCTo.Move 3120, edcSelCFrom.Top   '1050
            edcSelCTo.Visible = True
            ckcAll.Visible = False

            plcSelC1.Visible = True
            smPlcSelC1P = "Use"
            plcSelC1.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30
            rbcSelCSelect(0).Caption = "Deposit Date"
            rbcSelCSelect(0).Move 450, 0, 1400
            rbcSelCSelect(0).Visible = True
            If rbcSelCSelect(0).Value Then
                rbcSelCSelect_click 0   ', True
            Else
                rbcSelCSelect(0).Value = True
            End If
            rbcSelCSelect(1).Caption = "Entry Date"
            rbcSelCSelect(1).Move 1830, 0, 1290
            rbcSelCSelect(1).Visible = True
            rbcSelCSelect(2).Visible = False

            plcSelC2.Visible = True
            mAskCashTrMercProm plcSelC1.Top + plcSelC1.Height, False    'hide hard cost option

            'Allow selective transaction types
            plcSelC3.Move 120, plcSelC2.Top + plcSelC2.Height
            plcSelC3.Height = 435
            smPlcSelC3P = "Include"
            ckcSelC3(0).Left = 765
            ckcSelC3(0).Width = 2415
            ckcSelC3(0).Caption = "Payments (PI)"
            If ckcSelC3(0).Value = vbChecked Then
                ckcSelC3_click 0
            Else
                ckcSelC3(0).Value = vbChecked   'True
            End If
            ckcSelC3(1).Left = 2415
            ckcSelC3(1).Width = 2415
            ckcSelC3(1).Caption = "Payments (PO)"
            If ckcSelC3(1).Value = vbChecked Then
                ckcSelC3_click 1
            Else
                ckcSelC3(1).Value = vbChecked   'True
            End If
            ckcSelC3(2).Caption = "Journal Entries"
            ckcSelC3(2).Move 765, 195, 2415
            ckcSelC3(2).Value = vbChecked
            ckcSelC3(0).Visible = True
            ckcSelC3(1).Visible = True
            ckcSelC3(2).Visible = True
            ckcSelC3(3).Visible = False
            plcSelC3.Visible = True

            plcSelC4.Move plcSelC3.Left, plcSelC3.Top + plcSelC3.Height
            smPlcSelC4P = "Sort by"
            rbcSelC4(0).Caption = "Date"
            rbcSelC4(1).Caption = "Salesperson"
            rbcSelC4(2).Caption = "Vehicle Group"       '11-17-05
            rbcSelC4(0).Move 720, 0, 680
            rbcSelC4(1).Move 1440, 0, 1400
            rbcSelC4(2).Move 2840, 0, 1740
            If rbcSelC4(0).Value Then
                rbcSelC4_click 0    ', True
            Else
                rbcSelC4(0).Value = True
            End If
            rbcSelC4(0).Visible = True
            rbcSelC4(1).Visible = True
            rbcSelC4(2).Visible = True
            plcSelC4.Visible = True

            mAskAirTimeNTR plcSelC4.Top + plcSelC4.Height

            '7-9-03
            gPopVehicleGroups RptSelCreditStatus!cbcSet1, tgVehicleSets1(), True
            edcSet1.Text = "Vehicle Group Subsort"
            cbcSet1.ListIndex = 0
            edcSet1.Move 120, plcSelC6.Top + 30 + plcSelC6.Height, 2160
            cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 45
            edcSet1.Visible = True
            cbcSet1.Visible = True

            plcSelC7.Move edcSet1.Left, cbcSet1.Top + cbcSet1.Height + 30, 3360
            ckcSelC7.Caption = "Show transaction comments"
            ckcSelC7.Move 0, 0, 3360
            plcSelC7.Visible = True
            ckcSelC7.Visible = True

            '6-5-02 option to get a single check # printed
            lacCheck.Move plcSelC7.Left, plcSelC7.Top + plcSelC7.Height + 30    'chg to plcselc6 if reinstating the question to ask history, recv or both
            lacCheck.Caption = "Check #"
            lacCheck.Visible = True
            edcCheck.Move 900, lacCheck.Top - 30, 1170
            edcCheck.MaxLength = 10
            edcCheck.Visible = True

            pbcSelC.Visible = True
            pbcOption.Visible = True

            '8-4-08 Dan M option to include history, receivables or both
            'plcSelC8.Move 120, plcSelC7.Top + plcSelC7.Height + 30
            smPlcSelC8P = "From"
            plcSelC8.Move 120, edcCheck.Top + edcCheck.Height + 30
            rbcSelC8(0).Move 470, 0, 900
            rbcSelC8(1).Move 1400, 0, 1400
            rbcSelC8(2).Move 2770, 0, 1200
            rbcSelC8(0).Caption = "History"
            rbcSelC8(1).Caption = "Receivables"
            rbcSelC8(2).Caption = "Both"
            rbcSelC8(2).Value = True    'default to both
            plcSelC8.Visible = True

        Case COLL_CREDITSTATUS  'Credit Status option to print only cash in advance overdue 11-5-99
            'TTP 9893
            pbcSelC.Visible = False
            pbcOption.Visible = True
            ckcAllGroups.Left = 0
            ckcAllGroups.Caption = "All Advertisers"
            ckcAll.Caption = "All Agencies"
            
            plcSel2.Move 120, 0
            ckcSel2(0).Move 0, 0, 980
            ckcSel2(0).Caption = "Agency"
            ckcSel2(1).Move 1020, 0, 2200
            ckcSel2(1).Caption = "Advertiser"

            ckcDelinquentOnly.Move 120, plcSel2.Top + plcSel2.Height
            ckcDelinquentOnly.Visible = True
            lacInclude.Move 120, ckcDelinquentOnly.Top + ckcDelinquentOnly.Height + 60
            lacInclude.Visible = True
            ckcADate.Move 960, lacInclude.Top, 4605
            ckcADate.Value = vbUnchecked    'False  '11-15-99 True
            ckcADate.Caption = "No New Orders"
            'ckcADate.Left = 120
            'ckcADate.Top = 180
            'ckcADate.Width = 4605

            plcSel1.Move 120, ckcADate.Top + ckcADate.Height
            smPlcSel1P = ""

            ckcSel1(0).Left = 840   '690
            ckcSel1(0).Width = 1335
            ckcSel1(0).Caption = "Unrestricted"
            ckcSel1(1).Left = 2280  '2040
            ckcSel1(1).Width = 1710
            ckcSel1(1).Caption = "Zero Balances"
            ckcSel2(2).Visible = False

            ckcSel1(0).Value = vbUnchecked  'False
            ckcSel1(1).Value = vbUnchecked  'False
            If imAutoReport Then                'coming from Shocredit
                ckcSel1(1).Value = vbUnchecked  '6-12-07 chg to have check mark off to include zero_balances
                ckcDelinquentOnly.Value = vbChecked   '2-13-09 force from Alerts screen to show delinquents only (overdue)
            End If
            ckcSel2(0).Value = vbChecked    'True
            ckcSel2(1).Value = vbChecked    'True

            ckcInclCommentsA.Move 120, plcSel1.Top + plcSel1.Height, 2400
            ckcInclCommentsA.Caption = "Include Comments"

            ckcInclCommentsA.Visible = True

            lacFromA.Move 375, ckcInclCommentsA.Top + ckcInclCommentsA.Height + 60, 2040
            'lacFromA.Left = 375
            'lacFromA.Width = 2040
            lacFromA.Caption = "Comment Entered as of"
            'edcSelA.Left = 2400
            edcSelA.Move 2400, lacFromA.Top + 60

            lacFromA.Visible = False
            edcSelA.Visible = False

            ckcADate.Visible = True
            'lacFromA.Visible = False    'True
            'edcSelA.Visible = False 'True
            plcSel1.Visible = True
            plcSel2.Visible = True

            pbcSelA.Visible = True
        Case COLL_DISTRIBUTE                'Cash Distribution
            'mAskStartEndDates               'get start & end dates to retrieve
            lbcSelection(3).Visible = True
            ckcAll.Caption = "All Participants"
            ckcAll.Visible = True
            pbcSelC.Visible = True
            lbcSelection(3).Visible = True
            mAskEntryDates
            plcSelC1.Top = lacFrom.Top + lacFrom.Height
            'plcSelC1.Caption = "Select"
            smPlcSelC1P = "Select"
            rbcSelCSelect(0).Left = 630
            rbcSelCSelect(0).Width = 900
            rbcSelCSelect(0).Caption = "Invoice"
            rbcSelCSelect(0).Visible = True
            rbcSelCSelect(1).Left = 1530
            rbcSelCSelect(1).Width = 960
            rbcSelCSelect(1).Caption = "Check #"
            rbcSelCSelect(1).Visible = True
            rbcSelCSelect(2).Caption = "Participant"
            rbcSelCSelect(2).Left = 2550
            rbcSelCSelect(2).Width = 1200
            rbcSelCSelect(2).Visible = True
            If rbcSelCSelect(0).Value Then      'check #
                rbcSelCSelect_click 0   ', True
            Else
                rbcSelCSelect(0).Value = True
            End If
            plcSelC1.Visible = True
            smPlcSelC9P = ""
            plcSelC9.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height, 4000           '2-5-15 option to skip to new page each vehicle within participant
            ckcTrans.Left = 0
            ckcTrans.Width = 4000
            ckcTrans.Caption = "Skip to new page each new vehicle"
            plcSelC9.Visible = False
            ckcTrans.Visible = False
            plcSelC2.Visible = False
            pbcOption.Visible = True
        Case COLL_CASHSUM                  'Cash Summary

            lacSelCFrom.Caption = "Start Date"
            lacSelCTo.Caption = "End Date"
            lacSelCFrom.Visible = True
            edcSelCFrom.Text = Format$(gNow(), "m/d/yy")
            edcSelCFrom.Visible = True
            lacSelCTo.Visible = True
            edcSelCTo.Text = Format$(gNow(), "m/d/yy")
            edcSelCTo.Visible = True
            edcSelCFrom.Left = 960
            lacSelCTo.Move edcSelCFrom.Left + edcSelCFrom.Width + 120, lacSelCFrom.Top
            edcSelCTo.Move lacSelCTo.Left + 840, edcSelCFrom.Top

            plcSelC1.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30
            plcSelC1.Visible = True
            'plcSelC1.Caption = "Sort by"
            smPlcSelC1P = "Sort by"

            rbcSelCSelect(0).Width = 700
            rbcSelCSelect(0).Caption = "Vehicle"
            rbcSelCSelect(0).Move 720, 0, 980
            rbcSelCSelect(0).Value = True
            rbcSelCSelect(0).Visible = True

            rbcSelCSelect(1).Visible = True
            rbcSelCSelect(1).Width = 1280
            rbcSelCSelect(1).Caption = "Sales Office"
            rbcSelCSelect(1).Move 1700
            rbcSelCSelect(1).Visible = True
            If rbcSelCSelect(0).Value Then      'vehicle
                rbcSelCSelect_click 0   ', True
            Else
                rbcSelCSelect(0).Value = True
            End If

            mAskCashTrMercProm plcSelC1.Top + plcSelC1.Height, False    'hide hard cost option
            mAskAirTimeNTR plcSelC2.Top + 30 + plcSelC2.Height

             'Allow selective transaction types
            plcSelC3.Move 120, plcSelC6.Top + plcSelC6.Height + 60
            smPlcSelC3P = "Include"
            ckcSelC3(0).Left = 765
            ckcSelC3(0).Width = 1170
            ckcSelC3(0).Caption = "Payments"
            If ckcSelC3(0).Value = vbChecked Then
                ckcSelC3_click 0
            Else
                ckcSelC3(0).Value = vbChecked   'True
            End If
            ckcSelC3(1).Left = 1935
            ckcSelC3(1).Width = 2800
            ckcSelC3(1).Caption = "Journal Entries"
            If ckcSelC3(1).Value = vbChecked Then
                ckcSelC3_click 1
            Else
                ckcSelC3(1).Value = vbChecked   'True
            End If
            ckcSelC3(0).Visible = True
            ckcSelC3(1).Visible = True
            ckcSelC3(2).Visible = False
            ckcSelC3(3).Visible = False
            plcSelC3.Visible = True

            '7-10-03
            gPopVehicleGroups RptSelCreditStatus!cbcSet1, tgVehicleSets1(), True
            edcSet1.Text = "Vehicle Group"
            cbcSet1.ListIndex = 0
            edcSet1.Move 120, plcSelC3.Top + 30 + plcSelC3.Height
            cbcSet1.Move 360 + edcSet1.Width, edcSet1.Top - 45
            edcSet1.Visible = True
            cbcSet1.Visible = True
            plcSelC2.Visible = True
            pbcSelC.Visible = True

        Case COLL_ACCTHIST                  'Account History
            ckcAll.Visible = False

            lacSelCFrom.Caption = "Start Date"
            lacSelCTo.Caption = "End Date"
            lacSelCFrom.Visible = True
            edcSelCFrom.Text = ""
            edcSelCFrom.Visible = True
            lacSelCTo.Visible = True
            edcSelCTo.Text = ""
            edcSelCTo.Visible = True
            edcSelCFrom.Left = 960
            lacSelCFrom.Width = 1200
            lacSelCTo.Move edcSelCFrom.Left + edcSelCFrom.Width + 120, lacSelCFrom.Top, 1200
            edcSelCTo.Move lacSelCTo.Left + 840, edcSelCFrom.Top

            mAskContract edcSelCFrom.Top + edcSelCFrom.Height, 120

            'mAskCashTrMercProm                             'for now, include all types
            'plcSelC2.Visible = True

            plcSelC1.Move lacSelCTo1.Left, edcSelCTo1.Top + edcSelCTo1.Height + 60, 3705, 255
            smPlcSelC1P = "Select"
            rbcSelCSelect(0).Caption = "Advertiser"
            rbcSelCSelect(1).Caption = "Agency"
            rbcSelCSelect(0).Move 840, 0, 1200
            rbcSelCSelect(1).Move 2280, 0, 960
            plcSelC1.Visible = True
            rbcSelCSelect(0).Visible = True
            rbcSelCSelect(1).Visible = True
            If rbcSelCSelect(0).Value = True Then       'default to advertiser
                rbcSelCSelect_click 0
            Else
                rbcSelCSelect(0).Value = True
            End If

            plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height, 4000, 255
            rbcSelCInclude(0).Caption = "Billing"
            rbcSelCInclude(1).Caption = "Earned"
            rbcSelCInclude(2).Caption = "Both"
            rbcSelCInclude(0).Move 840, 0, 960
            rbcSelCInclude(1).Move 1800, 0, 960
            rbcSelCInclude(2).Move 2760, 0, 720
            rbcSelCInclude(2).Visible = True
            rbcSelCInclude(1).Visible = True

            If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) <> INSTALLMENT Then
                rbcSelCInclude(2).Value = True
            Else
                rbcSelCInclude(0).Value = True
            End If
            plcSelC2.Visible = True
            smPlcSelC2P = "Include"
            plcSelC3.Move plcSelC3.Left, plcSelC2.Top + plcSelC2.Height, 3705, 450
            mAskTranTypes
            plcSelC3.Visible = True
            smPlcSelC3P = "Include"
            ckcSelC3(0).Value = vbChecked
            ckcSelC3(0).Left = 765
            ckcSelC3(0).Width = 1050
            ckcSelC3(0).Caption = "Invoices"

            If ckcSelC3(0).Value = vbUnchecked Then
                ckcSelC3_click 0
            Else
                ckcSelC3(0).Value = vbChecked   'True
            End If
            ckcSelC3(1).Left = 2025
            ckcSelC3(1).Width = 1680
            ckcSelC3(1).Caption = "Adjustments"
            If ckcSelC3(1).Value = vbChecked Then
                ckcSelC3_click 1
            Else
                ckcSelC3(1).Value = vbChecked   'True
            End If
            ckcSelC3(2).Move 765, ckcSelC3(0).Top + 210, 1130
            ckcSelC3(2).Caption = "Payments"
            If ckcSelC3(2).Value = vbChecked Then
                ckcSelC3_click 2
            Else
                ckcSelC3(2).Value = vbChecked   'True
            End If
            ckcSelC3(3).Move 2025, ckcSelC3(0).Top + 210, 1440
            ckcSelC3(3).Caption = "Write-Offs"
            If ckcSelC3(3).Value = vbChecked Then
                ckcSelC3_click 3
            Else
                ckcSelC3(3).Value = vbChecked   'True
            End If
            ckcSelC3(0).Visible = True
            ckcSelC3(1).Visible = True
            ckcSelC3(2).Visible = True
            ckcSelC3(3).Visible = True

            plcSelC5.Move plcSelC3.Left, plcSelC3.Top + plcSelC3.Height, 3360
            ckcSelC5(0).Caption = "Show transaction comments"
            ckcSelC5(0).Move 0, 0, 3360
            plcSelC5.Visible = True
            ckcSelC5(0).Visible = True
            ckcSelC7.Caption = "Show Sales Source/Participant"
            ckcSelC7.Move 0, 0, 3360
            plcSelC7.Move 120, plcSelC5.Top + plcSelC5.Height, 3360
            plcSelC7.Visible = True
            ckcSelC7.Visible = True

            plcSelC10.Move 120, plcSelC7.Top + plcSelC7.Height, 3360
            ckcSelC10(0).Caption = "Skip to new page each group"
            ckcSelC10(0).Move 0, 0, 3360
            ckcSelC10(0).Visible = True
            ckcSelC10(1).Visible = False
            ckcSelC10(2).Visible = False
            ckcSelC10(0).Value = vbUnchecked
            plcSelC10.Visible = True

            '5-18-09  Add option to only see history, receivables
            plcSelC12.Move 120, plcSelC10.Top + plcSelC10.Height + 30
            rbcSelC12(2).Value = True
            rbcSelC12(0).Caption = "History"
            rbcSelC12(1).Caption = "Receivables"
            rbcSelC12(2).Caption = "Both"
            rbcSelC12(0).Move 0, 0, 960
            rbcSelC12(1).Move 960, 0, 1320
            rbcSelC12(2).Move 2400, 0, 720
            rbcSelC12(0).Visible = True
            rbcSelC12(1).Visible = True
            rbcSelC12(2).Visible = True
            plcSelC12.Visible = True

        Case COLL_MERCHANT
            lacSelCFrom.Caption = "Year"
            lacSelCFrom.Visible = True
            lacSelCFrom.Move 120, 75, 600
            lacSelCFrom1.Move 1460, 75, 810
            lacSelCFrom1.Caption = "Quarter"
            lacSelCFrom1.Visible = True
            edcSelCFrom.Move 600, lacSelCFrom.Top - 30, 600
            edcSelCFrom1.Move 2220, lacSelCFrom.Top - 30, 300
            edcSelCFrom.MaxLength = 4
            edcSelCFrom1.MaxLength = 1
            edcSelCFrom.Text = ""
            edcSelCFrom.Visible = True
            edcSelCFrom1.Text = ""
            edcSelCFrom1.Visible = True
            lacSelCTo.Visible = False
            edcSelCTo.Visible = False

            lacSelCTo.Caption = "Percent From "
            lacSelCTo.Visible = True
            lacSelCTo.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 60, 1320
            edcSelCTo.Move 1440, lacSelCTo.Top - 30, 780
            lacSelCTo1.Move 2280, edcSelCFrom.Top + edcSelCFrom.Height + 60, 600
            lacSelCTo1.Caption = "To"
            lacSelCTo1.Visible = True
            edcSelCTo1.Move 2640, lacSelCTo.Top - 30, 780
            edcSelCTo.MaxLength = 6
            edcSelCTo1.MaxLength = 6
            edcSelCTo.Text = ""
            edcSelCTo.Visible = True
            edcSelCTo1.Text = ""
            edcSelCTo1.Visible = True
            lacSelCTo.Visible = True
            edcSelCTo.Visible = True

            plcSelC4.Visible = True
            plcSelC4.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
            'plcSelC4.Caption = "Select"
            smPlcSelC4P = "Select"
            rbcSelC4(0).Caption = "Merchandising"
            rbcSelC4(0).Move 840, 0, 1560
            rbcSelC4(0).Visible = True
            rbcSelC4(1).Caption = "Promotions"
            rbcSelC4(1).Move 2400, 0, 1560
            rbcSelC4(1).Visible = True
            rbcSelC4(2).Visible = False
            If rbcSelC4(0).Value Then
                rbcSelC4_click 0
            Else
                rbcSelC4(0).Value = True
            End If

            plcSelC1.Visible = True
            plcSelC1.Move 120, plcSelC4.Top + plcSelC4.Height
            'plcSelC1.Caption = "Show by"
            smPlcSelC1P = "Show by"
            rbcSelCSelect(0).Caption = "Vehicle"
            rbcSelCSelect(0).Move 840, 0, 960
            rbcSelCSelect(0).Visible = True
            rbcSelCSelect(1).Caption = "Advertiser"
            rbcSelCSelect(1).Move 1860, 0, 1200
            rbcSelCSelect(1).Visible = True
            rbcSelCSelect(2).Visible = False
            If rbcSelCSelect(0).Value Then
                rbcSelCSelect_click 0
            Else
                rbcSelCSelect(0).Value = True
            End If
            plcSelC2.Visible = True
            plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height
            'plcSelC2.Caption = "By"
            smPlcSelC2P = "By"
            rbcSelCInclude(0).Caption = "Detail"
            rbcSelCInclude(0).Move 360, 0, 820
            rbcSelCInclude(0).Visible = True
            rbcSelCSelect(1).Width = 1200
            rbcSelCInclude(1).Caption = "Summary"
            rbcSelCInclude(1).Move 1260, 0, 1300
            rbcSelCInclude(1).Visible = True
            If rbcSelCInclude(0).Value Then
                rbcSelCInclude_Click 0
            Else
                rbcSelCInclude(0).Value = True
            End If

            'selective contract
            lacCheck.Caption = "Contract #"
            lacCheck.Move 120, plcSelC2.Top + plcSelC2.Height + 30
            edcCheck.Move 1080, plcSelC2.Top + plcSelC2.Height
            edcCheck.MaxLength = 9
            lacCheck.Visible = True
            edcCheck.Visible = True
        Case COLL_MERCHRECAP
            lacSelCFrom.Move 120, 75, 1320
            lacSelCFrom1.Move 2325, 75
            edcSelCFrom.Move 1290, edcSelCFrom.Top, 945
            edcSelCFrom1.Move 2700, edcSelCFrom.Top, 945
            lacSelCTo.Left = 120
            edcSelCTo.Move 1290, edcSelCTo.Top, 945
            lacSelCTo1.Left = 2325
            edcSelCTo1.Move 2700, edcSelCTo.Top, 945
            edcSelCTo.MaxLength = 10    '8 5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
            edcSelCTo.Text = ""
            edcSelCTo1.MaxLength = 10   '8
            edcSelCTo1.Text = ""
            edcSelCFrom.MaxLength = 10  '8
            edcSelCFrom1.MaxLength = 10 '8
            edcSelCFrom.Text = ""
            edcSelCFrom1.Text = ""
            lacSelCFrom.Caption = "Active From"
            lacSelCFrom1.Caption = "To"
            lacSelCFrom.Visible = True
            lacSelCFrom1.Visible = True
            lacSelCTo.Caption = "Entered: From"
            lacSelCTo1.Caption = "To"
            lacSelCTo.Visible = False       'True
            lacSelCTo1.Visible = False      'True
            edcSelCFrom.Visible = True
            edcSelCFrom1.Visible = True
            edcSelCTo.Visible = False       'True
            edcSelCTo1.Visible = False      'True
            ckcAll.Visible = False
        Case COLL_POAPPLY
            mAskPOApply
        End Select
    Exit Sub
End Sub
'
'           mAskCntrFeed - ask selectivity Contracts Spots & Feed Spots
'
'           8-9-04 DH
'
'           <input> top location to place control
'
Public Sub mAskCntrFeed(ilTop As Integer)
    plcSelC10.Move 120, ilTop, 4000
    ckcSelC10(0).Move 720, -30, 1680       'local
    ckcSelC10(1).Move 2400, -30, 1440      'feed
    ckcSelC10(0).Value = vbChecked
    ckcSelC10(1).Value = vbChecked
    If tgSpf.sSystemType = "R" Then         'radio vs network/syndicator
        ckcSelC10(0).Visible = True
        ckcSelC10(0).Caption = "Contract spots"
        ckcSelC10(1).Visible = True
        ckcSelC10(1).Caption = "Feed spots"
        plcSelC10.Visible = True
        smPlcSelC10P = "Include"
        plcSelC10_Paint
    End If
End Sub
'
'       mVehLabelsPop - populate all the vehicles possible to
'       have mailing labels
'       2-18-05
'
Public Sub mVehLabelsPop(llVehicletypes As Long)
 Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(RptSelCreditStatus, llVehicletypes + ACTIVEVEH, lbcSelection(0), tgAirNameCode(), sgAirNameCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehLabelsErr
        gCPErrorMsg ilRet, "mVehLabelsPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mVehLabelsErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'
'           mSetPrintForEntryDate - If Printables requested for
'           Copy Inventory by Entry Date, set the flag that it has
'           been printed
'           4-12-05
Public Sub mSetPrintForEntryDate()
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    ReDim tlCifCode(0 To 0) As Long

    slStartDate = RptSelCreditStatus!edcSelCFrom.Text   'earliest  date
    If slStartDate = "" Then                'if blank, set earliest date
        slStartDate = "1/1/1970"
    End If
    llStartDate = gDateValue(slStartDate)
    slStartDate = Format$(slStartDate, "m/d/yy") 'insure year is on date

    slEndDate = RptSelCreditStatus!edcSelCTo.Text   'Latest date
    If slEndDate = "" Then
        slEndDate = "12/31/2069"
    End If
    llEndDate = gDateValue(slEndDate)
    slEndDate = Format$(slEndDate, "m/d/yy") 'insure year is on date

    hmCif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCifRecLen = Len(tmCif)

    btrExtClear hmCif   'Clear any previous extend operation
    ilExtLen = Len(tmCif)  'Extract operation record size

    ilRet = btrGetFirst(hmCif, tmCif, imCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hmCif, llNoRec, -1, "UC", "Cif", "") '"EG") 'Set extract limits (all records)

        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Cif", "CIFENTRYDATE")
        ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Cif", "CIFENTRYDATE")
        ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        '4-13-05 include carted, uncarted, or both when Site TapeShowForm is set to "C" (vs "A" for approved)
        If RptSelCreditStatus!rbcSelC6(2).Value = False Then            'include both carted &  uncarted
            If RptSelCreditStatus!rbcSelC6(0).Value = True Then         'carted
                tlCharTypeBuff.sType = "Y"
            ElseIf RptSelCreditStatus!rbcSelC6(1).Value = True Then     'uncarted
                tlCharTypeBuff.sType = "N"
            End If
            ilOffSet = gFieldOffset("Cif", "CifCleared")
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        End If

        tlCharTypeBuff.sType = "P"          'get only items not already set
        ilOffSet = gFieldOffset("Cif", "CifPrint")
        ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)

        tlCharTypeBuff.sType = "A"          'get only active inventory items
        ilOffSet = gFieldOffset("Cif", "CifPurged")
        ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)

        ilRet = btrExtAddField(hmCif, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainCifErr
        gBtrvErrorMsg ilRet, "gObtainCif (btrExtAddField):" & "Cif.Btr", RptSelCreditStatus
        On Error GoTo 0
        ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainCifErr
            gBtrvErrorMsg ilRet, "gObtainCif (btrExtGetNextExt):" & "Cif.Btr", RptSelCreditStatus
            On Error GoTo 0
            ilExtLen = Len(tmCif)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlCifCode(UBound(tlCifCode)) = tmCif.lCode           'save entire record
                ReDim Preserve tlCifCode(0 To UBound(tlCifCode) + 1) As Long
                ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If

    'the array of printable inventory codes has been created, now go and set the printable flag
    For llNoRec = LBound(tlCifCode) To UBound(tlCifCode) - 1
        Do
            tmCifSrchKey.lCode = tlCifCode(llNoRec)
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmCif.sPrint = "P"
                ilRet = btrUpdate(hmCif, tmCif, imCifRecLen)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
    Next llNoRec

    Erase tlCifCode
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    Exit Sub
mObtainCifErr:
    On Error GoTo 0
    MsgBox "mSetPrintForEntryDate: gObtainCif error", vbCritical + vbOKOnly, "Cif I/O Error"
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAllConvAirSellPkgVehPop        *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                 with conventional, selling airing,  *
'*                 package vehicles
'*                                                     *
'*******************************************************
Public Sub mAllConvAirSellPkgVehPop(ilIndex As Integer)
Dim ilRet As Integer
         ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHSELLING + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)

    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAllConvAirSellPkgVehPopErr
        gCPErrorMsg ilRet, "mAllConvAirSellPkgVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mAllConvAirSellPkgVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
'
'
'               populate list box with conventional vehicles only
'
'           <input> index to list box
Public Sub mConvVehPop(ilIndex As Integer)
Dim ilRet As Integer
        ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHCONV_WO_FEED + VEHCONV_W_FEED + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mConvVehPopErr
        gCPErrorMsg ilRet, "mConvVehPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mConvVehPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
Public Sub mAskContract(ilTop As Integer, ilLeft As Integer)
        lacSelCTo1.Caption = "Contract #"
        edcSelCTo1.Text = ""
        lacSelCTo1.Move ilLeft, ilTop + 90, 1200
        edcSelCTo1.Move 1080, ilTop + 30, 960
        edcSelCTo1.MaxLength = 9                '1-27-06
        lacSelCTo1.Visible = True
        edcSelCTo1.Visible = True
    Exit Sub
End Sub
Public Sub mAskTranTypes()
        plcSelC3.Visible = True
        smPlcSelC3P = "Include"
        ckcSelC3(0).Value = vbChecked
        ckcSelC3(0).Left = 765
        ckcSelC3(0).Width = 1050
        ckcSelC3(0).Caption = "Invoices"

        If ckcSelC3(0).Value = vbUnchecked Then
            ckcSelC3_click 0
        Else
            ckcSelC3(0).Value = vbChecked   'True
        End If
        ckcSelC3(1).Left = 2025
        ckcSelC3(1).Width = 1680
        ckcSelC3(1).Caption = "Adjustments"
        If ckcSelC3(1).Value = vbChecked Then
            ckcSelC3_click 1
        Else
            ckcSelC3(1).Value = vbChecked   'True
        End If
        ckcSelC3(2).Move 765, ckcSelC3(0).Top + 210, 1130
        ckcSelC3(2).Caption = "Payments"
        If ckcSelC3(2).Value = vbChecked Then
            ckcSelC3_click 2
        Else
            ckcSelC3(2).Value = vbChecked   'True
        End If
        ckcSelC3(3).Move 2025, ckcSelC3(0).Top + 210, 1440
        ckcSelC3(3).Caption = "Write-Offs"
        If ckcSelC3(3).Value = vbChecked Then
            ckcSelC3_click 3
        Else
            ckcSelC3(3).Value = vbChecked   'True
        End If
        ckcSelC3(0).Visible = True
        ckcSelC3(1).Visible = True
        ckcSelC3(2).Visible = True
        ckcSelC3(3).Visible = True
    Exit Sub
End Sub
'
'               fill list box with selectivity choices
'               for User Activity Log
'               <input>  cbcList -control to fill
'                        ilListIndex = vbDefault selection
'                        ilIncludeNone - true to include "None" as a choice
Public Sub mFillSortList(cbcList As control, ilListIndex As Integer, ilIncludeNone As Integer)
        If ilIncludeNone Then
            cbcList.AddItem "None"
        End If
        cbcList.AddItem "Activity"
        cbcList.AddItem "Date"
        cbcList.AddItem "System Type"
        cbcList.AddItem "Time"
        cbcList.AddItem "User"
        cbcList.ListIndex = ilListIndex
        Exit Sub
End Sub
'
'               Ask Start date and End Date
'               <input> ilDefaultStartDate - true to default start date to todays date
'                       ilDefaultEndDate  - true to default end date to todays date
Public Sub mAskDates(ilDefaultSTartDate As Boolean, ilDefaultEndDate As Boolean, Optional slCaption As String = "")
Dim ilLeft As Integer
Dim ilWidth As Integer
        
        ilLeft = 1050
        ilWidth = 900
        If slCaption <> "" Then
            ilLeft = 1920
            ilWidth = 1800
        End If
        edcSelCFrom1.Visible = False
        plcSelC1.Visible = False
        plcSelC2.Visible = False
        lacSelCFrom.Width = ilWidth        '900
        lacSelCFrom.Caption = slCaption & "Start Date"
        lacSelCFrom.Visible = True
        edcSelCFrom.Text = ""
        If ilDefaultSTartDate Then
            edcSelCFrom.Text = Format$(gNow(), "m/d/yy")
        End If
        edcSelCFrom.Left = ilLeft
        edcSelCFrom.Visible = True
        lacSelCTo.Width = ilWidth
        lacSelCTo.Caption = slCaption & "End Date"
        lacSelCTo.Visible = True
        edcSelCTo.Text = ""
        If ilDefaultEndDate Then
            edcSelCTo.Text = Format$(gNow(), "m/d/yy")
        End If
        edcSelCTo.Left = ilLeft   '1050
        edcSelCTo.Visible = True
        Exit Sub
End Sub
'
'
'               populate list box with All vehicles that are flagged as live log
'
'           <input> index to list box
Public Sub mLiveLogVehiclesPop(ilIndex As Integer)
Dim ilRet As Integer
        ilRet = gPopUserVehicleBox(RptSelCreditStatus, VEHALLTYPES + VEHLIVELOG + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLiveLogVehiclesPopErr
        gCPErrorMsg ilRet, "mLiveLogVehiclesPop (gPopUserVehicleBox: Vehicle)", RptSelCreditStatus
        On Error GoTo 0
    End If
    Exit Sub
mLiveLogVehiclesPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSelCreditStatus
    'Set RptSelCreditStatus = Nothing   'Remove data segment
    Exit Sub
End Sub
Public Sub mBulkCopySelectivity()
        Select Case lbcRptType.ListIndex
                Case 0  'Feed
                    ckcAll.Visible = False
                    plcSelC1.Visible = False
                    plcSelC2.Visible = False
                    plcSelC3.Visible = False
                    pbcSelC.Visible = False
                    lbcSelection(1).Visible = False
                    lbcSelection(3).Visible = False
                    lbcSelection(0).Visible = True
                    lbcSelection(5).Visible = False
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case 1  'Cross Reference
                    ckcAll.Visible = False
                    plcSelC1.Visible = False
                    plcSelC2.Visible = False
                    plcSelC3.Visible = False
                    pbcSelC.Visible = False
                    lbcSelection(1).Visible = True
                    lbcSelection(3).Visible = False
                    lbcSelection(0).Visible = False
                    lbcSelection(5).Visible = False
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                Case 2  'Affiliate BF by Cart
                    lbcSelection(1).Visible = False
                    lbcSelection(3).Visible = False
                    lbcSelection(0).Visible = False
                    lbcSelection(5).Visible = False
                    pbcOption.Visible = False
                    ckcAll.Visible = False
                    plcSelC1.Visible = False
                    plcSelC2.Visible = False
                    pbcSelC.Visible = False
                    pbcOption.Visible = False
                    edcSelCTo.MaxLength = 12
                    edcSelCFrom.MaxLength = 12
                    lacSelCFrom.Width = 1300
                    lacSelCFrom.Caption = "Lowest Cart #"
                    lacSelCFrom.Visible = True
                    edcSelCFrom.Text = ""
                    edcSelCFrom.Left = 1350
                    edcSelCFrom.Visible = True
                    lacSelCTo.Width = 1300
                    lacSelCTo.Caption = "Highest Cart #"
                    lacSelCTo.Visible = True
                    edcSelCTo.Text = ""
                    edcSelCTo.Left = 1350
                    edcSelCTo.Visible = True
                    pbcSelC.Visible = True
                    plcSelC3.Left = 120
                    plcSelC3.Top = 885
                    'plcSelC3.Caption = "Include"
                    smPlcSelC3P = "Include"
                    ckcSelC3(0).Left = 705
                    ckcSelC3(0).Width = 1200
                    ckcSelC3(0).Caption = "Active"
                    If ckcSelC3(0).Value = vbChecked Then
                        ckcSelC3_click 0
                    Else
                        ckcSelC3(0).Value = vbChecked   'True
                    End If
                    ckcSelC3(0).Visible = True
                    ckcSelC3(1).Left = 1575
                    ckcSelC3(1).Width = 1380
                    ckcSelC3(1).Caption = "Purged"
                    ckcSelC3(1).Visible = True
                    ckcSelC3(1).Value = vbUnchecked 'False
                    ckcSelC3(2).Visible = False
                    ckcSelC3(3).Visible = False
                    plcSelC3.Visible = True
                    frcOption.Enabled = True
                Case 3  'Affiliate BF by Vehicle
                    plcSelC1.Visible = False
                    plcSelC2.Visible = False
                    pbcSelC.Visible = False
                    pbcOption.Visible = False
                    edcSelCFrom.MaxLength = 10  '8  5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
                    lacSelCFrom.Width = 1540
                    lacSelCFrom.Caption = "Active On or After"
                    lacSelCFrom.Visible = True
                    edcSelCFrom.Text = ""
                    edcSelCFrom.Left = 1640
                    edcSelCFrom.Visible = True
                    lacSelCTo.Visible = False
                    edcSelCTo.Visible = False
                    lbcSelection(3).Visible = True
                    lbcSelection(0).Visible = False
                    lbcSelection(1).Visible = False
                    lbcSelection(5).Visible = False
                    ckcAll.Caption = "All Vehicles"
                    pbcSelC.Visible = True
                    ckcAll.Visible = True
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                    plcSelC3.Left = 120
                    plcSelC3.Top = 885
                    'plcSelC3.Caption = "Include"
                    smPlcSelC3P = "Include"
                    ckcSelC3(0).Left = 705
                    ckcSelC3(0).Width = 1200
                    ckcSelC3(0).Caption = "Active"
                    If ckcSelC3(0).Value = vbChecked Then
                        ckcSelC3_click 0
                    Else
                        ckcSelC3(0).Value = vbChecked   'True
                    End If
                    ckcSelC3(0).Visible = True
                    ckcSelC3(1).Left = 1575
                    ckcSelC3(1).Width = 1380
                    ckcSelC3(1).Caption = "Purged"
                    ckcSelC3(1).Visible = True
                    ckcSelC3(1).Value = vbUnchecked 'False
                    ckcSelC3(2).Visible = False
                    ckcSelC3(3).Visible = False
                    plcSelC3.Visible = True
                Case 4  'Affiliate BF by Feed Date
                    lbcSelection(1).Visible = False
                    lbcSelection(3).Visible = False
                    lbcSelection(0).Visible = False
                    lbcSelection(5).Visible = False
                    pbcOption.Visible = False
                    edcSelCTo.MaxLength = 10    '8 5/27/99 chged from 10 to 8 for short form m/d/yyyy date input
                    edcSelCFrom.MaxLength = 10  '8
                    lacSelCFrom.Width = 1300
                    lacSelCFrom.Caption = "Feed From Date"
                    lacSelCFrom.Visible = True
                    edcSelCFrom.Text = ""
                    edcSelCFrom.Left = 1440
                    edcSelCFrom.Visible = True
                    lacSelCTo.Width = 1300
                    lacSelCTo.Caption = "Feed To Date"
                    lacSelCTo.Visible = True
                    edcSelCTo.Text = ""
                    edcSelCTo.Left = 1440
                    edcSelCTo.Visible = True
                    pbcSelC.Visible = True
                    plcSelC3.Left = 120
                    plcSelC3.Top = 885
                    'plcSelC3.Caption = "Include"
                    smPlcSelC3P = "Include"
                    ckcSelC3(0).Left = 705
                    ckcSelC3(0).Width = 1200
                    ckcSelC3(0).Caption = "Active"
                    If ckcSelC3(0).Value = vbChecked Then
                        ckcSelC3_click 0
                    Else
                        ckcSelC3(0).Value = vbChecked   'True
                    End If
                    ckcSelC3(0).Visible = True
                    ckcSelC3(1).Left = 1575
                    ckcSelC3(1).Width = 1380
                    ckcSelC3(1).Caption = "Purged"
                    ckcSelC3(1).Visible = True
                    ckcSelC3(1).Value = vbUnchecked 'False
                    ckcSelC3(2).Visible = False
                    ckcSelC3(3).Visible = False
                    plcSelC3.Visible = True
                    frcOption.Enabled = True
                Case 5  'Affiliate BF by Advertiser
                    plcSelC1.Visible = False
                    plcSelC2.Visible = False
                    lacSelCFrom.Visible = False
                    edcSelCFrom.Visible = False
                    lacSelCTo.Visible = False
                    edcSelCTo.Visible = False
                    'pbcSelC.Visible = False
                    pbcOption.Visible = False
                    lbcSelection(5).Visible = True
                    lbcSelection(3).Visible = False
                    lbcSelection(0).Visible = False
                    lbcSelection(1).Visible = False
                    ckcAll.Caption = "All Advertisers"
                    ckcAll.Visible = True
                    frcOption.Enabled = True
                    pbcOption.Visible = True
                    plcSelC3.Left = 120
                    plcSelC3.Top = 885
                    'plcSelC3.Caption = "Include"
                    smPlcSelC3P = "Include"
                    ckcSelC3(0).Left = 705
                    ckcSelC3(0).Width = 1200
                    ckcSelC3(0).Caption = "Active"
                    If ckcSelC3(0).Value = vbChecked Then
                        ckcSelC3_click 0
                    Else
                        ckcSelC3(0).Value = vbChecked   'True
                    End If
                    ckcSelC3(0).Visible = True
                    ckcSelC3(1).Left = 1575
                    ckcSelC3(1).Width = 1380
                    ckcSelC3(1).Caption = "Purged"
                    ckcSelC3(1).Visible = True
                    ckcSelC3(1).Value = vbUnchecked 'False
                    ckcSelC3(2).Visible = False
                    ckcSelC3(3).Visible = False
                    plcSelC3.Visible = True
                    pbcSelC.Visible = True
            End Select
        Exit Sub
End Sub

'       12/8/2020 - TTP 9893 - add Adv/Agcy filters to Credit Status Reports
'       gGenCreditStatusGRF - Generate prepass file for
'                   Collection report.
'           Uses list box Selection of Agencies
'                   and/or Advertisers, Inserts a GRF Record
'                   Used as a filter for report(s): CreditAg.rpt, CreditAd.Rpt
'
Sub gGenCreditStatusGRF()
    Dim tmGrf As GRF
    Dim hmGrf As Integer
    Dim imGrfRecLen As Integer        'GRF record length
    Dim ilTemp As Long
    Dim ilRet As Integer
    Dim ilError As Integer
    Dim slCode As String
    Dim slDate As String
    Dim slTime As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slNameCode As String
    
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imGrfRecLen = Len(tmGrf)

    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Sub
    End If

    
    If ckcSel2(1).Value = 1 Then                                'select advt
        For ilTemp = 0 To lbcSelection(0).ListCount - 1 Step 1
            If lbcSelection(0).Selected(ilTemp) Then
                slNameCode = tgAdvertiser(ilTemp).sKey          'pick up agcy code
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If slCode <> "" Then
                    tmGrf.iGenDate(0) = igNowDate(0)            'todays date used for removal of records
                    tmGrf.iGenDate(1) = igNowDate(1)
                    tmGrf.lGenTime = lgNowTime
                    tmGrf.iAdfCode = slCode                     'AdfCode
                    tmGrf.iPerGenl(0) = 0
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        GoTo mGenErr
                    End If
                End If
            End If
        Next ilTemp
    End If

    If ckcSel2(0).Value = 1 Then                                'select agcy
        For ilTemp = 0 To lbcSelection(1).ListCount - 1 Step 1
            If lbcSelection(1).Selected(ilTemp) Then
                slNameCode = tgAgency(ilTemp).sKey              'pick up agcy code
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If slCode <> "" Then
                    tmGrf.iGenDate(0) = igNowDate(0)            'todays date used for removal of records
                    tmGrf.iGenDate(1) = igNowDate(1)
                    tmGrf.lGenTime = lgNowTime
                    tmGrf.iAdfCode = 0
                    tmGrf.iPerGenl(0) = slCode                  'AgfCode
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        GoTo mGenErr
                    End If
                End If
            End If
        Next ilTemp
    End If


mGenErr:
    'Cleanup
    ilRet = btrClose(hmGrf)
    btrDestroy hmGrf

End Sub


