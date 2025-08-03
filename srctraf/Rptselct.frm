VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelCt 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facility Report Selection"
   ClientHeight    =   7005
   ClientLeft      =   1110
   ClientTop       =   2175
   ClientWidth     =   11490
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7005
   ScaleWidth      =   11490
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   2070
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   4020
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   2640
         TabIndex        =   15
         Top             =   1200
         Width           =   1005
      End
      Begin VB.CheckBox ckcSeparateFile 
         Caption         =   "Separate files per Vehicle"
         Height          =   255
         Left            =   120
         TabIndex        =   174
         Top             =   1200
         Visible         =   0   'False
         Width           =   2535
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
         Height          =   315
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
         TabIndex        =   14
         Top             =   720
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   645
      End
   End
   Begin V81TrafficReports.CSI_ComboBoxList cbcEMailContent 
      Height          =   330
      Left            =   3045
      TabIndex        =   151
      Top             =   720
      Visible         =   0   'False
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   582
      BorderStyle     =   1
   End
   Begin VB.Frame frcEMail 
      Caption         =   "E-Mail"
      Height          =   1635
      Left            =   2070
      TabIndex        =   147
      Top             =   60
      Visible         =   0   'False
      Width           =   6915
      Begin VB.PictureBox pbcECTab 
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   6600
         ScaleHeight     =   180
         ScaleWidth      =   135
         TabIndex        =   152
         Top             =   1320
         Width           =   135
      End
      Begin VB.TextBox edcSendTo 
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
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   149
         Text            =   "Station"
         Top             =   195
         Width           =   1005
      End
      Begin VB.TextBox edcResponse 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   154
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lacTo 
         Appearance      =   0  'Flat
         Caption         =   "Send To"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lacContent 
         Appearance      =   0  'Flat
         Caption         =   "Content"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   150
         Top             =   660
         Width           =   855
      End
      Begin VB.Label lacResponse 
         Appearance      =   0  'Flat
         Caption         =   "Response By"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   1140
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   9120
      TabIndex        =   26
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   9720
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10680
      Top             =   1320
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
      Left            =   10560
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   1680
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
      Left            =   10560
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcNoSort 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcMultiCntr 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.FileListBox lbcFileName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3960
      Pattern         =   "*.Dal"
      TabIndex        =   83
      Top             =   5520
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSMask.MaskEdBox mkcPhone 
      Height          =   210
      Left            =   4680
      TabIndex        =   82
      Tag             =   "The number and extension of the buyer."
      Top             =   4800
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
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
   Begin VB.ListBox lbcLnCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcCntrCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4080
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".pdf"
      Filter          =   $"Rptselct.frx":0000
      FilterIndex     =   2
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Frame frcCopies 
      Caption         =   "Printing"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   2070
      TabIndex        =   5
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
         TabIndex        =   7
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Vehicle Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5220
      Left            =   75
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   9810
      Begin VB.PictureBox pbcSelC 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
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
         Height          =   4845
         Left            =   60
         ScaleHeight     =   4845
         ScaleMode       =   0  'User
         ScaleWidth      =   4875
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   195
         Visible         =   0   'False
         Width           =   4875
         Begin VB.CheckBox ckcSuppressNTRDetails 
            Caption         =   "Suppress NTR Details"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   240
            TabIndex        =   177
            Top             =   4570
            Width           =   2160
         End
         Begin VB.CheckBox ckcShowACT1 
            Caption         =   "Show ACT1 Codes and Settings"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   240
            TabIndex        =   175
            Top             =   4340
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.PictureBox plcSelC14 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   3015
            TabIndex        =   168
            Top             =   3600
            Visible         =   0   'False
            Width           =   3015
            Begin VB.OptionButton rbcSelC14 
               Caption         =   "Split"
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   170
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton rbcSelC14 
               Caption         =   "Grouped"
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   169
               Top             =   0
               Value           =   -1  'True
               Width           =   1095
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
            Left            =   405
            MaxLength       =   3
            TabIndex        =   35
            Top             =   45
            Visible         =   0   'False
            Width           =   420
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
            Left            =   2745
            MaxLength       =   10
            TabIndex        =   33
            Top             =   90
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.PictureBox plcSelC9 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   4500
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   3120
            Visible         =   0   'False
            Width           =   4500
            Begin VB.OptionButton rbcSelC9 
               Caption         =   "Bill Method"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   3480
               TabIndex        =   171
               Top             =   0
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.ComboBox cbcSort2 
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
               Left            =   3000
               TabIndex        =   158
               Top             =   480
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.ComboBox cbcSort1 
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
               Left            =   720
               TabIndex        =   157
               Top             =   480
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.OptionButton rbcSelC9 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   3360
               TabIndex        =   133
               Top             =   15
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.OptionButton rbcSelC9 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1800
               TabIndex        =   119
               Top             =   0
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.OptionButton rbcSelC9 
               Caption         =   "Corporate"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   118
               Top             =   0
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.OptionButton rbcSelC9 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2640
               TabIndex        =   117
               Top             =   0
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label lacSort2 
               Appearance      =   0  'Flat
               Caption         =   "Sort #2"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   2400
               TabIndex        =   160
               Top             =   525
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.Label lacSort1 
               Appearance      =   0  'Flat
               Caption         =   "Sort #1"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   75
               TabIndex        =   159
               Top             =   525
               Visible         =   0   'False
               Width           =   645
            End
         End
         Begin VB.ListBox lbcAgyAdvtCode 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   2520
            Sorted          =   -1  'True
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   3960
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.CheckBox ckcInclRevAdj 
            Caption         =   "Include Rev Adj"
            Height          =   255
            Left            =   1800
            TabIndex        =   155
            Top             =   4080
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.CheckBox ckcInclZero 
            Caption         =   "Include $0 contracts"
            Height          =   255
            Left            =   240
            TabIndex        =   146
            Top             =   4080
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox lacText 
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
            Left            =   120
            TabIndex        =   142
            Text            =   "Effec Pacing Date"
            Top             =   3840
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.TextBox edcText 
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
            MaxLength       =   10
            TabIndex        =   40
            Top             =   3360
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.PictureBox plcSelC13 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            FillColor       =   &H80000008&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   4380
            Begin VB.OptionButton rbcShow 
               Caption         =   "Ext Contract #"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   3240
               TabIndex        =   176
               TabStop         =   0   'False
               Top             =   360
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.TextBox lacShow 
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
               TabIndex        =   163
               Text            =   "Show"
               Top             =   360
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.OptionButton rbcShow 
               Caption         =   "Order Prop Date"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1560
               TabIndex        =   162
               TabStop         =   0   'False
               Top             =   360
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.OptionButton rbcShow 
               Caption         =   "GRP"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   840
               TabIndex        =   161
               TabStop         =   0   'False
               Top             =   360
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CheckBox ckcSelC13 
               Caption         =   "Page"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   1620
               TabIndex        =   145
               Top             =   30
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CheckBox ckcSelC13 
               Caption         =   "Use Primary Only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   960
               TabIndex        =   144
               Top             =   0
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CheckBox ckcSelC13 
               Caption         =   "Use Primary Only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1800
               TabIndex        =   143
               Top             =   0
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CheckBox ckcSelC13 
               Caption         =   "Use Primary Only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   139
               Top             =   -30
               Visible         =   0   'False
               Width           =   1200
            End
         End
         Begin VB.PictureBox plcSelC12 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   45
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   3360
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1695
               TabIndex        =   137
               Top             =   -15
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1110
               TabIndex        =   136
               Top             =   -30
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.CheckBox ckcSelC12 
               Caption         =   "Use Primary Only"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   135
               Top             =   -30
               Visible         =   0   'False
               Width           =   1200
            End
         End
         Begin VB.TextBox edcTopHowMany 
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
            Left            =   3480
            TabIndex        =   132
            Text            =   "Major Set #"
            Top             =   3240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.PictureBox plcSelC11 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   2880
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC11 
               Caption         =   "30/60"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2400
               TabIndex        =   129
               Top             =   15
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.OptionButton rbcSelC11 
               Caption         =   "Units"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   128
               Top             =   0
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC11 
               Caption         =   "30"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1650
               TabIndex        =   127
               Top             =   0
               Visible         =   0   'False
               Width           =   660
            End
         End
         Begin VB.PictureBox plcSelC10 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   240
            ScaleHeight     =   240
            ScaleWidth      =   4185
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   2640
            Visible         =   0   'False
            Width           =   4185
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Show Splits"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   125
               Top             =   -30
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1320
               TabIndex        =   124
               Top             =   0
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.CheckBox ckcSelC10 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2520
               TabIndex        =   123
               Top             =   0
               Visible         =   0   'False
               Width           =   1020
            End
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
            Left            =   3480
            TabIndex        =   121
            Text            =   "Minor Set #"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
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
            Left            =   3480
            TabIndex        =   120
            Text            =   "Major Set #"
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
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
            TabIndex        =   50
            Tag             =   "48"
            Top             =   3840
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
            Left            =   1680
            TabIndex        =   49
            Top             =   3840
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox plcSelC8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   3885
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   3885
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   2280
               TabIndex        =   113
               Top             =   -30
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1110
               TabIndex        =   112
               Top             =   -30
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.CheckBox ckcSelC8 
               Caption         =   "EST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   111
               Top             =   -30
               Visible         =   0   'False
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSelC7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   4380
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2400
               TabIndex        =   109
               Top             =   0
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   705
               TabIndex        =   108
               Top             =   0
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC7 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   1650
               TabIndex        =   107
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.PictureBox plcSelC6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   360
            ScaleHeight     =   240
            ScaleWidth      =   4620
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   4620
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2520
               TabIndex        =   104
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
               TabIndex        =   103
               Top             =   480
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "Missed"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   4
               Left            =   360
               TabIndex        =   102
               Top             =   480
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2295
               TabIndex        =   101
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
               TabIndex        =   100
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
               TabIndex        =   99
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.CheckBox ckcSelC6 
               Caption         =   "EST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   98
               Top             =   -30
               Visible         =   0   'False
               Width           =   600
            End
         End
         Begin VB.PictureBox plcSelC5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   30
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1560
            Visible         =   0   'False
            Width           =   4380
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "PST"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   3120
               TabIndex        =   18
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
               TabIndex        =   23
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
               TabIndex        =   24
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
               TabIndex        =   73
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
               TabIndex        =   74
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
               TabIndex        =   75
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
               TabIndex        =   76
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
               TabIndex        =   77
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
               TabIndex        =   78
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
               TabIndex        =   79
               Top             =   195
               Visible         =   0   'False
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSelC3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   0
            ScaleHeight     =   435
            ScaleWidth      =   4860
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1080
            Visible         =   0   'False
            Width           =   4860
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Digital"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   4320
               TabIndex        =   178
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   3840
               TabIndex        =   141
               Top             =   0
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   7
               Left            =   2880
               TabIndex        =   66
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
               TabIndex        =   67
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   62
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
               TabIndex        =   61
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
               TabIndex        =   60
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
               TabIndex        =   59
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
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   39
            Top             =   60
            Visible         =   0   'False
            Width           =   945
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
            Left            =   915
            MaxLength       =   10
            TabIndex        =   37
            Top             =   60
            Width           =   1170
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   4515
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   4515
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "Vehicle Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   6
               Left            =   840
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "Agency"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   240
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "incl4"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   2775
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   15
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "incl3"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2325
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "New/Changed"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   1785
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "All"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   3510
               TabIndex        =   46
               Top             =   0
               Value           =   -1  'True
               Width           =   510
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "New/Changed"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1245
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   0
               Width           =   1560
            End
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   240
            ScaleWidth      =   4140
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   4140
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   3495
               TabIndex        =   16
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
               TabIndex        =   115
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
               TabIndex        =   44
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
               TabIndex        =   55
               Top             =   0
               Width           =   675
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Agency"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1290
               TabIndex        =   42
               Top             =   0
               Width           =   900
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Salesperson"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2220
               TabIndex        =   43
               Top             =   0
               Width           =   900
            End
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
            Left            =   60
            TabIndex        =   56
            Top             =   60
            Visible         =   0   'False
            Width           =   4305
         End
         Begin VB.PictureBox plcSelC4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   15
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1650
               TabIndex        =   70
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   705
               TabIndex        =   69
               Top             =   0
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Net-Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   2400
               TabIndex        =   71
               Top             =   0
               Width           =   1005
            End
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   0
            TabIndex        =   164
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "5/15/24"
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
            CSI_AllowBlankDate=   -1  'True
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin V81TrafficReports.CSI_Calendar CSI_CalTo 
            Height          =   255
            Left            =   1200
            TabIndex        =   165
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "5/15/24"
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
            CSI_AllowBlankDate=   -1  'True
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin V81TrafficReports.CSI_Calendar CSI_From1 
            Height          =   255
            Left            =   0
            TabIndex        =   166
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "5/15/24"
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
            CSI_AllowBlankDate=   -1  'True
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin V81TrafficReports.CSI_Calendar CSI_To1 
            Height          =   255
            Left            =   0
            TabIndex        =   167
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "5/15/24"
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
            CSI_AllowBlankDate=   -1  'True
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# of Days"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2400
            TabIndex        =   34
            Top             =   75
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Active Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
            TabIndex        =   57
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lacPeriods 
            Appearance      =   0  'Flat
            Caption         =   "Active End Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   0
            TabIndex        =   140
            Top             =   0
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label lacTopDown 
            Appearance      =   0  'Flat
            Caption         =   "Top How Many"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   131
            Top             =   2880
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lacSelCTo1 
            Appearance      =   0  'Flat
            Caption         =   "To"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2325
            TabIndex        =   38
            Top             =   375
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lacSelCTo 
            Appearance      =   0  'Flat
            Caption         =   "Active End Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   36
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
         Height          =   4365
         Left            =   5040
         ScaleHeight     =   4295.987
         ScaleMode       =   0  'User
         ScaleWidth      =   4575
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CheckBox CkcAllVeh 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   259
            Left            =   0
            TabIndex        =   130
            Top             =   0
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.CheckBox ckcAllAAS 
            Caption         =   "All "
            ForeColor       =   &H80000008&
            Height          =   259
            Left            =   15
            TabIndex        =   114
            Top             =   -15
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   12
            Left            =   0
            TabIndex        =   105
            Top             =   0
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   11
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   96
            Top             =   0
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   10
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   95
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   9
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   94
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   8
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   93
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   7
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   92
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   6
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   85
            Top             =   30
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   5
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   84
            Top             =   45
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   4
            ItemData        =   "Rptselct.frx":008D
            Left            =   30
            List            =   "Rptselct.frx":008F
            MultiSelect     =   1  'Simple
            TabIndex        =   31
            Top             =   45
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   3
            Left            =   30
            MultiSelect     =   2  'Extended
            TabIndex        =   30
            Top             =   45
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   2
            Left            =   60
            MultiSelect     =   2  'Extended
            TabIndex        =   29
            Top             =   45
            Visible         =   0   'False
            Width           =   4350
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   1
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   21
            Top             =   45
            Visible         =   0   'False
            Width           =   4395
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   3390
            Index           =   0
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   20
            Top             =   45
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   259
            Left            =   255
            TabIndex        =   22
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
            TabIndex        =   88
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
            TabIndex        =   89
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
      Left            =   9120
      TabIndex        =   27
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   9120
      TabIndex        =   25
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Export"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   1380
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "E-Mail"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1095
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   810
         Width           =   1395
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   525
         Width           =   750
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Label lacExport 
      Height          =   195
      Left            =   2160
      TabIndex        =   172
      Top             =   1440
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   10920
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RptSelCt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselct.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSmf                                                                                 *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  mAskBobYrQtrPeriods                                                                  *
'******************************************************************************************

'  Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelCt.Frm
'
' Release: 5.0
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
Dim smPaintCaption1 As String    'caption for panel
Dim smPaintCaption2 As String    'caption for panel
Dim smPaintCaption3 As String    'caption for panel
Dim smPaintCaption4 As String    'caption for panel
Dim smPaintCaption5 As String    'caption for panel
Dim smPaintCaption6 As String    'caption for panel
Dim smPaintCaption7 As String    'caption for panel
Dim smPaintCaption8 As String    'caption for panel
Dim smPaintCaption9 As String    'caption for panel
Dim smPaintCaption10 As String    'caption for panel
Dim smPaintCaption11 As String    'caption for panel
Dim smPaintCaption12 As String    'caption for panel
Dim smPaintCaption13 As String
Dim smPaintCaption14 As String

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
Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer
Dim tmChfSrchKey As LONGKEY0
Dim tmChfAdvtExt() As CHFADVTEXT
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
'Dim tmSRec As LPOPREC
'Rate Card
Dim smRateCardTag As String
Dim imMajorSort As Integer          'major sort selection index for Sales Comparison
Dim imMinorSort As Integer          'minor sort selection index for sales comparison
Dim imPrevMinorIndex As Integer
Dim imPrevMajorIndex As Integer

Dim imDoubleClickName As Integer
Dim imEMailContentSelectedIndex As Integer
Dim tmEMailContentCode() As SORTCODE
Dim smEMailContentCodeTag As String
Dim tmEmail_Info() As EMAILINFO                 '7-2-15

'Date: 9/10/2018 added drop down list for sorting   FYM
Dim imSort1Index As Integer
Dim imSort2Index As Integer
Dim imSort1PrevIndex As Integer
Dim imSort2PrevIndex As Integer
Dim imSetAllSort1 As Integer
Dim imSetAllSort2 As Integer

Dim tmBillCycle As BILLCYCLE                '1-13-21   If pulling B & B by Bill method, then an extra set of dates needs to be maintained
'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
Dim tmVehicleList() As Integer
Public bmHasRecords As Boolean
'Dim smClientName As String
Dim tmMnfSrchKey As INTKEY0
Dim tmMnfList() As MNFLIST      'array of mnf codes for Missed reasons and billing rules
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF

'
'       mFillSortOptions - Populate the combo boxes with the list of Sort Options
'       <input>  cbcCombo - combo control
'                ilShowNone - true if NONE allowed as an option
'
Public Sub mFillSortOptions(cbcCombo As Control, ilShowNone As Integer)
    cbcCombo.Clear

    If ilShowNone Then
        cbcCombo.AddItem "None"
    End If
    cbcCombo.AddItem "Active Start Date"
    cbcCombo.AddItem "Advertiser"
    cbcCombo.AddItem "Agency"
    cbcCombo.AddItem "Salesperson"
    cbcCombo.AddItem "Vehicle"
    cbcCombo.ListIndex = 0

    Exit Sub
End Sub

'       mAskTypeOfMonth - determine corporate, standard or calendar month
'       <input> control for options
Public Sub mAskTypeOfMonth(slCaption As String, rbcCorp As Control, rbcStd As Control, rbcCal As Control)
    slCaption = "Month"
    rbcCorp.Caption = "Corp"
    rbcCorp.Left = 660
    rbcCorp.Width = 720
    rbcStd.Caption = "Std"
    rbcStd.Left = 1440
    rbcStd.Width = 600
    rbcCal.Caption = "Cal"
    rbcCal.Move 2080, 0, 600
    rbcCorp.Visible = True
    rbcStd.Visible = True
    rbcCal.Visible = True
    If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
        rbcCorp.Enabled = False
        rbcCorp.Value = False
        rbcStd.Value = True
    Else
        rbcCorp.Value = True
    End If
End Sub

Private Sub mSellConvNoNTRPop(ilIndex As Integer)
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_W_CLUSTER + VEHREP_WO_CLUSTER + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)     'lbcCSVNameCode)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvNoNTRPopErr
        gCPErrorMsg ilRet, "mSellConvNoNTRPop (gPopUserVehicleBox: Vehicle)", RptSelCt
        On Error GoTo 0
    End If
    Exit Sub

mSellConvNoNTRPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'
'   Ask Gross, Net or TNet
'   <input> ilTop: top location of Gross/Net/TNet question
'
Private Sub mAskGrossNetTNet(ilTop As Integer)
    smPaintCaption9 = "By"
    plcSelC9_Paint
    rbcSelC9(0).Move 420, 0, 840    'gross button,
    rbcSelC9(0).Caption = "Gross"
    rbcSelC9(0).Value = True
    rbcSelC9(0).Visible = True
    rbcSelC9(1).Move 1380, 0, 660   'net button
    rbcSelC9(1).Caption = "Net"
    rbcSelC9(1).Visible = True
    rbcSelC9(2).Caption = "T-Net"
    rbcSelC9(2).Move 2040, 0, 960
    rbcSelC9(2).Visible = True
    plcSelC9.Move 120, ilTop, 4400
    plcSelC9.Visible = True
    Exit Sub
End Sub

Private Sub mSetSortOptions(iIndex As Integer)
    Dim Value As Integer
    Dim ilListIndex As Integer
    'Value = rbcSelC9(Index).Value
    Value = iIndex
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            If ilListIndex = CNT_BOB Then                   'Bill & Booked
                If Value Then
                     If iIndex = 0 Then                      'Corp (vs Std)
                        ckcSelC6(0).Value = vbUnchecked     'False           'default trades off
                        ckcSelC8(1).Value = vbUnchecked 'False           'default show mgs where they air off
                        ckcSelC8(0).Visible = False
                        ckcSelC8(1).Visible = False
                        ckcSelC8(0).Value = vbUnchecked 'False       'ignore missed
                        ckcSelC8(1).Value = vbUnchecked 'False       'ignore mgs
                        edcText.Enabled = True
                        lacText.Enabled = True
                        ckcInclRevAdj.Value = vbChecked
                     ElseIf iIndex = 3 Then                          '2-6-15 cal spots always shows spots where scheduled
                        ckcSelC8(0).Visible = True
                        'ckcSelC8(1).Visible = False                '7-29-16 reinstate option
                        ckcInclRevAdj.Value = vbChecked             '7-29-16 default to incl Rev Adjustments
                        edcText.Enabled = False                     '8-23-17 disable ability to pace on cal spots
                        lacText.Enabled = False
                        edcText.Text = ""                           '8-24-17 ensure the pacing date is blank, a previous report could have been pacing
                        '12-4-17 show selection to show mg where they air if Aired Billing; otherwise for As Ordered always include the spot as ordered
                        ckcSelC8(1).Visible = False
                        If tgSpf.sInvAirOrder = "A" Then
                            ckcSelC8(1).Visible = True
                        End If
                     Else                       'std or cal (spots)
                        ckcSelC6(0).Value = vbChecked   'True
                        ckcSelC8(1).Value = vbChecked   'True            'set default to show mgs whre they air
                        ckcSelC8(0).Visible = True
                        ckcSelC8(1).Visible = True
                        ckcSelC8(0).Value = vbUnchecked 'False       'ignore missed
                        ckcSelC8(1).Value = vbChecked   'True       'ignore mgs
                        ckcInclRevAdj.Value = vbChecked
                        If iIndex = 3 Then                           '1-09-08if cal by spots, disallow the pacing feature
                            edcText.Enabled = False
                            lacText.Enabled = False
                            edcText.Text = ""
                        Else
                            edcText.Enabled = True          'pacing date
                            lacText.Enabled = True
                            '12-8-17 disable sub unresolved missed and show mg where air if billing is as ordered, update ordered for std option
                            If tgSpf.sInvAirOrder = "S" Then             'bill as ordered, update as ordered; don't ask adjustment qustions
                                If iIndex = 2 Then               'calendar by contract
                                    ckcSelC8(0).Visible = True
                                    ckcSelC8(1).Visible = True
                                Else            'calendar by spot with as ordered , update ordered billing.  never show subtr unresolved missed and count mg where air.
                                    ckcSelC8(0).Visible = False
                                    ckcSelC8(1).Visible = False
                                End If
                            Else                                'as aired
                                ckcSelC8(0).Visible = True  '9-12-02 vbChecked 'True
                                ckcSelC8(1).Visible = True  '9-12-02 vbChecked 'True
                            End If
                        End If
                     End If
                End If
            ElseIf ilListIndex = CNT_PAPERWORK Then
                'If rbcSelC9(3).Value Then                  'by vehicle, contract or line options not available
                If cbcSort1.ListIndex = 4 Or cbcSort2.ListIndex = 5 Then              'by vehicle, contract or line options not available
                    plcSelC2.Visible = False                'disallow summary if by vehicle
                    rbcSelCInclude(1).Value = True          'default to line report
                    cbcSet1.Visible = False
                    edcSet1.Visible = False
                    ckcSelC5(7).Value = vbUnchecked         'disallow NTR with any vehicle/line option
                    ckcSelC5(7).Enabled = False
                Else                                '2-28-05 not vehicle sort, allow report by contract summary or line
                    rbcSelCInclude(0).Value = True      'default to show by contract
                    plcSelC2.Visible = True
                    cbcSet1.Visible = True
                    edcSet1.Visible = True
                    ckcSelC5(7).Value = vbChecked         'turn NTR option back on and default to show on report
                    ckcSelC5(7).Enabled = True
                    If rbcSelCSelect(3).Value = True Then
                        plcSelC2.Visible = False                'disallow summary if by vehicle
                        rbcSelCInclude(1).Value = True          'default to line report
                        cbcSet1.Visible = False
                        edcSet1.Visible = False
                        ckcSelC5(7).Value = vbUnchecked         'disallow NTR with any vehicle/line option
                        ckcSelC5(7).Enabled = False
                    End If
                    If rbcSelC7(2).Value Then                   'show acq cost only
                        rbcSelCInclude(1).Value = True          'force to show by lines to see acq cost
                    End If
                End If
            ElseIf ilListIndex = CNT_SALESACTIVITY_SS Or ilListIndex = CNT_SALESPLACEMENT Then      '2-18-03 allow option to include slsp subtotals
                If iIndex = 0 Or iIndex = 1 Then          'by market, source, office or source, office, mkt
            
                    plcSelC12.Height = 440
                    ckcSelC12(1).Move 0, 210, 2640
                    ckcSelC12(1).Caption = "Include Slsp Sub-totals"
                    ckcSelC12(1).Visible = True
                    If ilListIndex = CNT_SALESACTIVITY_SS Then
                        '4-1-11 option to split the slsp
                        ckcSelC12(2).Move 2520, 210, 1920
                        ckcSelC12(2).Caption = "Show Slsp Splits"
                        ckcSelC12(2).Visible = True
                    End If
                Else
                    plcSelC12.Height = 240
                    ckcSelC12(1).Visible = False
                    ckcSelC12(1).Value = vbUnchecked        'force to exclude slsp subtotals
                    ckcSelC12(2).Value = vbUnchecked        '4-1-11 no slsp, no splits
                    ckcSelC12(2).Visible = False
            
                End If
                If ilListIndex = CNT_SALESACTIVITY_SS Then
                    plcSelC13.Move 0, plcSelC12.Top + plcSelC12.Height, 4000
                    ckcSelC13(0).Caption = "Air Time"
                    ckcSelC13(1).Caption = "NTR"
                    ckcSelC13(2).Caption = "Hard Cost"
                    smPaintCaption13 = "Include"
                    plcSelC13_Paint
                    ckcSelC13(0).Value = vbChecked
                    ckcSelC13(1).Value = vbUnchecked
                    ckcSelC13(2).Value = vbUnchecked
                    ckcSelC13(0).Move 840, 0, 1080
                    ckcSelC13(1).Move 1920, 0, 720
                    ckcSelC13(2).Move 2640, 0, 1200
                    ckcSelC13(0).Visible = True
                    ckcSelC13(1).Visible = True
                    ckcSelC13(2).Visible = True
                    plcSelC13.Visible = True
                End If
            End If
        Case SLSPCOMMSJOB
            If ilListIndex = COMM_SALESCOMM Then
                If iIndex = 0 Then
                    If Value = True Then
                        ckcSelC10(0).Value = vbUnchecked
                        ckcSelC10(0).Enabled = False
                    End If
                Else
                    ckcSelC10(0).Enabled = True
                End If
            End If
    End Select
End Sub

Private Sub cbcEMailContent_DblClick()
    imDoubleClickName = True
    If Not mEMailContentBranch() Then
        edcResponse.SetFocus
    End If
End Sub

Private Sub cbcEMailContent_GotFocus()
    If cbcEMailContent.ListIndex < 0 Then
        If cbcEMailContent.ListCount > 1 Then
            cbcEMailContent.SetListIndex = 1
        Else
            cbcEMailContent.SetListIndex = 0
        End If
    End If
End Sub

Private Sub cbcEMailContent_OnChange()
    If imChgMode = False Then
        imChgMode = True
        imEMailContentSelectedIndex = cbcEMailContent.ListIndex
        imChgMode = False
    End If
    mSetCommands
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

Private Sub cbcSel_Change()
    mSetCommands
End Sub

Private Sub cbcSel_Click()
    mSetCommands
End Sub

Private Sub cbcSet1_Click()
    Dim ilListIndex As Integer
    Dim ilSetIndex As Integer
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim ilTemp As Integer
    Dim llRg As Long
    Dim llRet As Long
    ilSetIndex = cbcSet1.ListIndex
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            If ilListIndex = CNT_BOB Then
                illoop = cbcSet1.ListIndex
                ilSetIndex = gFindVehGroupInx(illoop, tgVehicleSets1())

                If ilSetIndex > 0 Then
                    ckcSelC10(0).Enabled = True         '2-8-16
                    ckcSelC10(1).Enabled = True
                    smVehGp5CodeTag = ""
                    ilRet = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(7), tgMnfCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(ilSetIndex)))
                    If ilSetIndex = 1 Then              'participants vehicle sets
                        If rbcSelCInclude(4).Value Then      'vehicle/participant option, selecting participant for vehicle group is redundant
                            MsgBox "Vehicle Group is already by Participant, select None or a different group"
                            cbcSet1.ListIndex = 0
                            mSetCommands
                            Exit Sub
                        End If

                        lbcSelection(7).Visible = True
                        CkcAllveh.Caption = "All Participants"
                        If rbcSelCInclude(2).Value = True Then          'if vehicle sort and participants selected, disable ability to do Slsp subsort/splits
                            ckcSelC10(0).Value = Unchecked              '2-8-16 disallow another set of splits (slsp) if splitting by participants
                            ckcSelC10(1).Value = Unchecked
                            ckcSelC10(0).Enabled = False
                            ckcSelC10(1).Enabled = False
                        End If
                    ElseIf ilSetIndex = 2 Then          'subtotals vehicle sets
                        lbcSelection(7).Visible = True
                        CkcAllveh.Caption = "All Sub-totals"
                    ElseIf ilSetIndex = 3 Then          'market vehicle sets
                        lbcSelection(7).Visible = True
                        CkcAllveh.Caption = "All Markets"
                    ElseIf ilSetIndex = 4 Then          'format vehicle sets
                        lbcSelection(7).Visible = True
                        CkcAllveh.Caption = "All Formats"
                    ElseIf ilSetIndex = 5 Then          'research vehicle sets
                        lbcSelection(7).Visible = True
                        CkcAllveh.Caption = "All Research"
                    ElseIf ilSetIndex = 6 Then          'sub-company vehicle sets
                        lbcSelection(7).Visible = True
                        CkcAllveh.Caption = "All Sub-Companies"
                    End If
                    'use advt list box for left, top
                    lbcSelection(6).Move lbcSelection(5).Left, lbcSelection(5).Top, 4380, 1500  '1740
                    If rbcSelCInclude(4).Value Then      'vehicle/participant option
                        'use vehicle list box and owner list box for locations
                        lbcSelection(2).Width = lbcSelection(6).Width \ 2
                        lbcSelection(7).Move lbcSelection(6).Left + lbcSelection(6).Width / 2 + 90, lbcSelection(2).Top, lbcSelection(2).Width - 90, 1500
                        CkcAllveh.Move lbcSelection(7).Left, ckcAllAAS.Top
                        CkcAllveh.Visible = True
                    Else        'vehicle or vehicle gross/net option
                        lbcSelection(7).Move lbcSelection(6).Left, lbcSelection(6).Top + lbcSelection(6).Height + 375, 4380, 1500
                        CkcAllveh.Visible = True
                        CkcAllveh.Move lbcSelection(7).Left, lbcSelection(7).Top - CkcAllveh.Height
                    End If
                Else            'no group selected, or changed to No group, re-enable if Vehicle or Slsp options
                    lbcSelection(7).Visible = False
                    CkcAllveh.Value = vbUnchecked   '9-12-02 False
                    CkcAllveh.Visible = False
                    lbcSelection(6).Height = lbcSelection(2).Height 'ensure the height of vehicle list box is correct
                    lbcSelection(2).Width = lbcSelection(6).Width   'ensure the width of particpant box is correct
                    If rbcSelCInclude(2).Value = True Then          '5-20-20 vehicle option, no group selected.  enabled the options to show slsp
                        ckcSelC10(0).Enabled = True
                        ckcSelC10(1).Enabled = True
                    End If
                End If
            ElseIf ilListIndex = CNT_QTRLY_AVAILS Then
                illoop = cbcSet1.ListIndex
                ilSetIndex = gFindVehGroupInx(illoop, tgVehicleSets1())

                If ilSetIndex > 0 Then
                    smVehGp5CodeTag = ""
                    ilRet = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(2), tgSOCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(ilSetIndex)))
                    If ilSetIndex = 1 Then              'participants vehicle sets
                        lbcSelection(2).Visible = True
                        ckcAllAAS.Caption = "All Participants"
                    ElseIf ilSetIndex = 2 Then          'subtotals vehicle sets
                        lbcSelection(2).Visible = True
                        ckcAllAAS.Caption = "All Sub-totals"
                    ElseIf ilSetIndex = 3 Then          'market vehicle sets
                        lbcSelection(2).Visible = True
                        ckcAllAAS.Caption = "All Markets"
                    ElseIf ilSetIndex = 4 Then          'format vehicle sets
                        lbcSelection(2).Visible = True
                        ckcAllAAS.Caption = "All Formats"
                    ElseIf ilSetIndex = 5 Then          'research vehicle sets
                        lbcSelection(2).Visible = True
                        ckcAllAAS.Caption = "All Research"
                    ElseIf ilSetIndex = 6 Then          'sub-company vehicle sets
                        lbcSelection(2).Visible = True
                        ckcAllAAS.Caption = "All Sub-Companies"
                    End If
                    ckcAllAAS.Visible = True
                Else
                    lbcSelection(2).Visible = False
                    ckcAllAAS.Value = vbUnchecked   '9-12-02 False
                    ckcAllAAS.Visible = False
                End If
            ElseIf ilListIndex = CNT_SALESACTIVITY_SS Or ilListIndex = CNT_SALESPLACEMENT Then  '8-2-02
                illoop = cbcSet1.ListIndex
                ilSetIndex = gFindVehGroupInx(illoop, tgVehicleSets1())
                If ilSetIndex > 0 Then
                    smVehGp5CodeTag = ""
                    ilRet = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(6), tgMnfCodeCT(), sgMNFCodeTagCT, "H" & Trim$(str$(ilSetIndex)))
                End If
                If ilSetIndex = 1 Then          'Participants
                    rbcSelC9(0).Caption = "Participants, Source, Office"
                    rbcSelC9(1).Caption = "Source, Office, Participants, Advt"
                    rbcSelC9(2).Caption = "Advt, Participants, Contract"
                    CkcAllveh.Caption = "All Participants"
                    lbcSelection(6).Visible = True
                    CkcAllveh.Visible = True
                ElseIf ilSetIndex = 2 Then      'Sub-totals
                    rbcSelC9(0).Caption = "Sub-totals, Source, Office"
                    rbcSelC9(1).Caption = "Source, Office, Sub-totals, Advt"
                    rbcSelC9(2).Caption = "Advt, Sub-totals, Contract"
                    CkcAllveh.Caption = "All Sub-totals"
                    lbcSelection(6).Visible = True
                    CkcAllveh.Visible = True
                ElseIf ilSetIndex = 3 Then      'markets
                    rbcSelC9(0).Caption = "Market, Source, Office"
                    rbcSelC9(1).Caption = "Source, Office, Markets, Advt"
                    rbcSelC9(2).Caption = "Advt, Market, Contract"
                    CkcAllveh.Caption = "All Markets"
                    lbcSelection(6).Visible = True
                    CkcAllveh.Visible = True
                ElseIf ilSetIndex = 4 Then      'formats
                    rbcSelC9(0).Caption = "Format, Source, Office"
                    rbcSelC9(1).Caption = "Source, Office, Formats, Advt"
                    rbcSelC9(2).Caption = "Advt, Formats, Contract"
                    CkcAllveh.Caption = "All Formats"
                    lbcSelection(6).Visible = True
                    CkcAllveh.Visible = True
                ElseIf ilSetIndex = 5 Then     'research
                    rbcSelC9(0).Caption = "Research, Source, Office"
                    rbcSelC9(1).Caption = "Source, Office, Research, Advt"
                    rbcSelC9(2).Caption = "Advt, Research, Contract"
                    CkcAllveh.Caption = "All Research"
                    lbcSelection(6).Visible = True
                    CkcAllveh.Visible = True
                ElseIf ilSetIndex = 6 Then     'research
                    rbcSelC9(0).Caption = "Sub-company, Source, Office"
                    rbcSelC9(1).Caption = "Source, Office, Sub-company, Advt"
                    rbcSelC9(2).Caption = "Advt, Sub-company, Contract"
                    CkcAllveh.Caption = "All Sub-companies"
                    lbcSelection(6).Visible = True
                    CkcAllveh.Visible = True
                Else                            'None
                    rbcSelC9(0).Caption = "Source, Office"
                    rbcSelC9(1).Caption = "Source, Office, Advt"
                    rbcSelC9(2).Caption = "Advt, Contract"
                    lbcSelection(6).Visible = False
                    CkcAllveh.Visible = False
                End If
            ElseIf ilListIndex = CNT_SALESCOMPARE Or ilListIndex = CNT_ADVT_UNITS Or ilListIndex = CNT_AVG_PRICES Then
                'sort selection  has been changed from radio buttons to combo box.
                'set the radio button to the new method so that all the other code
                'works that previously tested the radio buttons
                
                'Date:  8/27/2019 added major/minor sorts to ADVERTISER UNITS
                '       11/6/2019 added major/minor sorts to AVG SPOTS PRICE
                imMajorSort = ilSetIndex
                ilTemp = cbcSet2.ListIndex
                If cbcSet1.ListIndex = ilTemp - 1 Then    'both major and minor cannot be the same set
                    MsgBox "Select different major and minor sorts"
                    cbcSet1.ListIndex = imPrevMajorIndex
                    Exit Sub
                End If
                If ilSetIndex = 0 Then          'advt
                    rbcSelCInclude(0).Value = True
                ElseIf ilSetIndex = 4 Then     'slsp
                    rbcSelCInclude(1).Value = True
                ElseIf ilSetIndex = 1 Then     'agency
                    rbcSelCInclude(2).Value = True
                ElseIf ilSetIndex = 2 Then     'bus category
                    rbcSelCInclude(3).Value = True
                ElseIf ilSetIndex = 3 Then     'prod protection
                    rbcSelCInclude(4).Value = True
                ElseIf ilSetIndex = 5 Then    'vehicle
                    rbcSelCInclude(5).Value = True
                Else                            'vehicle group
                    rbcSelCInclude(6).Value = True
                End If
                imPrevMajorIndex = ilSetIndex
            ElseIf ilListIndex = CNT_BOBCOMPARE Then
                If ilSetIndex = 0 Then           'none selected, disable the budget selection
                    'turn off any budget selections made
                    'llRg = CLng(lbcSelection(4).ListCount - 1) * &H10000 Or 0
                    'llRet = SendMessageByNum(lbcSelection(4).hwnd, LB_SELITEMRANGE, True, llRg)
                    lbcSelection(4).Enabled = False
                    ckcSelC13(2).Enabled = True         'allow sales source to be used when budgets not used
                Else
                    lbcSelection(4).Enabled = True
                    'turn off any budget selections made
                    llRg = CLng(lbcSelection(4).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(4).HWnd, LB_SELITEMRANGE, False, llRg)
                    ckcSelC13(2).Enabled = False    'do not allow sales source to be used when budgets are selected; it is confusing and overstates the budgets
                                                        'becuase the vehicle could be in multiple sales sources
                    ckcSelC13(2).Value = vbUnchecked
                End If
            End If
            mSetCommands
    End Select
    Exit Sub
End Sub

Private Sub cbcSet2_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilLoop                                                  *
'******************************************************************************************
    Dim ilListIndex As Integer
    Dim ilSetIndex As Integer
    Dim ilTemp As Integer
    Dim ilTemp1 As Integer

    ilSetIndex = cbcSet2.ListIndex
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            rbcSelC7(2).Visible = True              'default to show t-net
            If ilListIndex = CNT_BOB Then
                plcSelC7.Visible = True
                If tgUrf(0).iSlfCode > 0 Then           'slsp/mgr is the user, need to remap the options because
                                                        'owner, vehicle part and vehicle net-net  has been disallowed
                    If ilSetIndex = 0 Then                  'advt
                        rbcSelCInclude(0).Value = True
                    ElseIf ilSetIndex = 1 Then              'agy
                        rbcSelCInclude(5).Value = True
                    ElseIf ilSetIndex = 2 Then              'slsp
                        rbcSelCInclude(1).Value = True
                    ElseIf ilSetIndex = 3 Then              'vehicle
                        rbcSelCInclude(2).Value = True
                    End If
                Else
                    If ilSetIndex = 0 Then                  'advt
                        rbcSelCInclude(0).Value = True
                    ElseIf ilSetIndex = 1 Then              'agy
                        rbcSelCInclude(5).Value = True
                    ElseIf ilSetIndex = 2 Then              'owner
                        rbcSelCInclude(3).Value = True
                    ElseIf ilSetIndex = 3 Then              'slsp
                        rbcSelCInclude(1).Value = True
                    ElseIf ilSetIndex = 4 Then              'vehicle
                        rbcSelCInclude(2).Value = True
                    ElseIf ilSetIndex = 5 Then             'vehicle net-net
                        rbcSelCInclude(6).Value = True
                        rbcSelC7(2).Visible = False
                        rbcSelC7(2).Value = True            'force for net-net ,need to disable acquisition
                        plcSelC7.Visible = False
                    Else                                    ' vehicle/participant
                        rbcSelCInclude(4).Value = True
                    End If
                End If
            ElseIf ilListIndex = CNT_SALESCOMPARE Or ilListIndex = CNT_ADVT_UNITS Or ilListIndex = CNT_AVG_PRICES Then
                'Date:  8/27/2019 added major/minor sorts to ADVERTISER UNITS
                '       11/6/2019 added major/minor sorts to AVG SPOTS PRICE

                ilSetIndex = cbcSet2.ListIndex
                imMinorSort = ilSetIndex
                If ilSetIndex = 0 Then                  'no minor set selected
                    ckcSelC13(0).Value = vbUnchecked    'Include advt totals only used when both sorts are selected and Advt isnt one of them
                    ckcSelC13(0).Enabled = False
                    
                    'Date: 11/9/2019 enable "Use Sales Source as Major Sort" for AVG_PRICES
                    If ilListIndex = CNT_AVG_PRICES Then ckcSelC13(0).Enabled = True
                    
                    lbcSelection(1).Height = 3270
                    lbcSelection(2).Height = 3270
                    lbcSelection(3).Height = 3270
                    lbcSelection(4).Height = 3270
                    lbcSelection(12).Height = 3270      '3-18-16
                    lbcSelection(5).Height = 3270
                    lbcSelection(6).Height = 3270
                    lbcSelection(7).Height = 3270
                    lbcSelection(4).Width = lbcSelection(1).Width / 2 - 30      'vehicle group
                    lbcSelection(12).Width = lbcSelection(1).Width / 2 - 30     '3-18-16 single selection vg
                    lbcSelection(8).Height = 3270
                    lbcSelection(8).Width = lbcSelection(4).Width           'items within group
                    lbcSelection(8).Width = lbcSelection(12).Width
                    ckcAllAAS.Visible = False
                    CkcAllveh.Visible = False
                    CkcAllveh.Value = vbUnchecked
                    If imPrevMinorIndex = 0 Then
                        Exit Sub
                    'turn off only the list box that was selected for the minor sort set
                    'dont want to turn off the major sort selection
                    ElseIf imPrevMinorIndex = 1 Then          'advt
                        lbcSelection(5).Visible = False
                    ElseIf imPrevMinorIndex = 2 Then         'agy
                        lbcSelection(1).Visible = False
                    ElseIf imPrevMinorIndex = 3 Then         'bus cat
                        lbcSelection(3).Visible = False
                    ElseIf imPrevMinorIndex = 4 Then         'prod protection
                        lbcSelection(7).Visible = False
                    ElseIf imPrevMinorIndex = 5 Then         'slsp
                        lbcSelection(2).Visible = False
                    ElseIf imPrevMinorIndex = 6 Then         'vehicle
                        lbcSelection(6).Visible = False
                    ElseIf imPrevMinorIndex = 7 Then        'vehicle group
                        lbcSelection(4).Visible = False
                        lbcSelection(12).Visible = False        '3-18-16
                        lbcSelection(8).Visible = False
                    End If
                    imPrevMinorIndex = 0
                    imPrevMinorIndex = cbcSet2.ListIndex
                    ckcAllAAS.Value = vbUnchecked
                    CkcAllveh.Value = Unchecked
                    Exit Sub
                End If

                'changing the minor sort selection, only turn off the previous list box
                'dont want to turn off the major sort selection
                'verify the previous cbcset2 index selection
                If imPrevMinorIndex = 1 Then          'advt
                    lbcSelection(5).Visible = False
                ElseIf imPrevMinorIndex = 2 Then         'agy
                    lbcSelection(1).Visible = False
                ElseIf imPrevMinorIndex = 3 Then         'bus cat
                    lbcSelection(3).Visible = False
                ElseIf imPrevMinorIndex = 4 Then         'prod protection
                    lbcSelection(7).Visible = False
                ElseIf imPrevMinorIndex = 5 Then         'slsp
                    lbcSelection(2).Visible = False
                ElseIf imPrevMinorIndex = 6 Then         'vehicle
                    lbcSelection(6).Visible = False
                ElseIf imPrevMinorIndex = 7 Then        'vehicle group
                    lbcSelection(4).Visible = False
                    lbcSelection(12).Visible = False     '3-18-16
                    lbcSelection(8).Visible = False
                    CkcAllveh.Visible = False
                    CkcAllveh.Value = vbUnchecked
                End If
                ilTemp = cbcSet2.ListIndex
                If cbcSet1.ListIndex = ilTemp - 1 Then     'error
                    MsgBox "Select different major and minor sorts"
                    cbcSet2.ListIndex = 0                   'default to none
                    Exit Sub
                End If
                lbcSelection(1).Height = 1500
                lbcSelection(2).Height = 1500
                lbcSelection(3).Height = 1500
                lbcSelection(4).Height = 1500
                lbcSelection(12).Height = 1500
                lbcSelection(5).Height = 1500
                lbcSelection(6).Height = 1500
                lbcSelection(7).Height = 1500
                lbcSelection(8).Height = 1500
                ckcAllAAS.Value = vbUnchecked
                If ilSetIndex = 1 Then          'advt
                    lbcSelection(5).Move 120, 2100
                    lbcSelection(5).Visible = True
                    ckcAllAAS.Move 120, 1800
                    ckcAllAAS.Caption = "All Advertisers"
                    ckcAllAAS.Visible = True
                    imPrevMinorIndex = ilSetIndex
               ElseIf ilSetIndex = 2 Then     'agency
                    lbcSelection(1).Move 120, 2100
                    lbcSelection(1).Visible = True
                    ckcAllAAS.Move 120, 1800
                    ckcAllAAS.Caption = "All Agencies"
                    ckcAllAAS.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 3 Then     'bus category
                    lbcSelection(3).Move 120, 2100
                    lbcSelection(3).Visible = True
                    ckcAllAAS.Move 120, 1800
                    ckcAllAAS.Caption = "All Bus Categories"
                    ckcAllAAS.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 4 Then     'prod protection
                    lbcSelection(7).Move 120, 2100
                    lbcSelection(7).Visible = True
                    ckcAllAAS.Move 120, 1800
                    ckcAllAAS.Caption = "All Prod Protection"
                    ckcAllAAS.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 5 Then     'slsp
                    lbcSelection(2).Move 120, 2100
                    lbcSelection(2).Visible = True
                    ckcAllAAS.Move 120, 1800
                    ckcAllAAS.Caption = "All Salespeople"
                    ckcAllAAS.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 6 Then    'vehicle
                    lbcSelection(6).Move 120, 2100
                    lbcSelection(6).Visible = True
                    ckcAllAAS.Move 120, 1800
                    ckcAllAAS.Caption = "All Vehicles"
                    ckcAllAAS.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 7 Then    'vehicle group

                    'lbcSelection(4).Move 15, 2100
                    'lbcSelection(4).Visible = True
                    lbcSelection(12).Move 120, 2100          '3-18-16 change to single selection box
                    lbcSelection(12).Visible = True
                    lbcSelection(8).Move lbcSelection(4).Width + 240, 2100
                    lbcSelection(4).Width = lbcSelection(8).Width
                    lbcSelection(12).Width = lbcSelection(8).Width      '3-18-16
                    lbcSelection(8).Visible = True
                    CkcAllveh.Caption = "All Items"
                    CkcAllveh.Visible = True
                    CkcAllveh.Value = vbUnchecked
                    CkcAllveh.Move lbcSelection(8).Left, 1800
                    ckcAllAAS.Move 120, 1800
                    ckcAllAAS.Visible = False
                
                    imPrevMinorIndex = ilSetIndex
                End If
                'Show advt totals only needed if both sorts did not select advt.
                'i.e. may need advt totals if major vehicle group, minor slsp and want to
                'see subtotals for the advt
                If cbcSet1.ListIndex <> 0 And cbcSet2.ListIndex <> 1 Then
                    ckcSelC13(0).Enabled = True
                Else
                    'Date: 11/9/2019 enable "Use Sales Source as Major Sort" for AVG_PRICES
                    If ilListIndex <> CNT_AVG_PRICES Then
                        ckcSelC13(0).Value = vbUnchecked
                        ckcSelC13(0).Enabled = False
                    End If
                End If
            ElseIf ilListIndex = CNT_BOBCOMPARE Then
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(3).Visible = False
                lbcSelection(5).Visible = False
                lbcSelection(6).Visible = False
                lbcSelection(7).Visible = False
                lbcSelection(8).Visible = False

                lbcSelection(4).Move 2000, 2100, 1800, 1500
                laclbcName(0).Move 2000, 1800
                lbcSelection(6).Move 15, 2100, 1800, 1500
                lbcSelection(6).Visible = True
                ckcAllAAS.Visible = True
                ckcAll.Value = vbUnchecked
                If ilSetIndex = 0 Then          'advt
                    lbcSelection(5).Visible = True
                    ckcAll.Caption = "All Advertisers"
                    ckcAll.Visible = True
                    imPrevMinorIndex = ilSetIndex
               ElseIf ilSetIndex = 1 Then     'agency
                    lbcSelection(1).Visible = True
                    ckcAll.Caption = "All Agencies"
                    ckcAll.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 2 Then     'bus category
                    lbcSelection(3).Visible = True
                    ckcAll.Caption = "All Bus Categories"
                    ckcAll.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 3 Then     'prod protection
                    lbcSelection(7).Visible = True
                    ckcAll.Caption = "All Prod Protection"
                    ckcAll.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 4 Then     'slsp
                    lbcSelection(2).Visible = True
                    ckcAll.Caption = "All Salespeople"
                    ckcAll.Visible = True
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 5 Then    'vehicle
                    lbcSelection(6).Move 15, 280, 4000, 1500  'vehicle list box
                    lbcSelection(6).Visible = True
                    ckcAll.Caption = "All Vehicles"
                    ckcAll.Visible = True
                    ckcAllAAS.Visible = False
                    lbcSelection(4).Move 15, 2100, 4000, 2100
                    laclbcName(0).Move 15, lbcSelection(1).Top + lbcSelection(1).Height + 60
                    imPrevMinorIndex = ilSetIndex
                ElseIf ilSetIndex = 6 Then    'vehicle group
                    lbcSelection(4).Visible = True
                    imPrevMinorIndex = ilSetIndex
                End If
            End If
    End Select
    mSetCommands
End Sub

Private Sub cbcSort1_Click()
    imSort1Index = cbcSort1.ListIndex
    If (imSort1Index = imSort2Index) And (imSort1Index <> 0) Then
        'error, cannot have the same sort parameter defined for 2 sort fields
        cbcSort1.ListIndex = imSort1PrevIndex
        MsgBox "Same sort selection as Sort #2; select another"
    Else
        imSort1PrevIndex = imSort1Index
    End If
    mSetSortOptions imSort1Index
    mSetCommands
End Sub

Private Sub cbcSort2_Click()
    imSort2Index = cbcSort2.ListIndex
    If (imSort2Index = imSort1Index + 1) And (imSort2Index <> 0) Then
        'error, cannot have the same sort parameter defined for 2 sort fields
         MsgBox "Same sort selection as Sort #1; select another"
         cbcSort2.ListIndex = 0
    Else
        imSort2PrevIndex = imSort2Index
    End If
    mSetSortOptions imSort2Index
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
    Dim ilSelect As Integer             'index to lbcselection
    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If igRptCallType = SLSPCOMMSJOB Then
            If ilIndex = COMM_SALESCOMM Or ilIndex = COMM_PROJECTION Then
                If lbcSelection(2).ListCount > 0 Then       'select all slsp
                    llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
            End If
        ElseIf igRptCallType = CONTRACTSJOB Then
            ilListIndex = lbcRptType.ListIndex
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            Select Case ilListIndex
                Case CNT_BR, CNT_INSERTION                                 'proposals/contracts
                    llRg = CLng(lbcSelection(10).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(10).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case 1                               'paperwork summary
                    llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case 2                               'spots by advt
                    If rbcSelCSelect(0).Value Then
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
                    Else
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                Case 3, 14 'Spots by times; Missed Spots
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case 17 'Quarterly
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case CNT_BOB_BYCNT, CNT_BOB_BYSPOT, CNT_BOB_BYSPOT_REPRINT     'formerly Projections
                    If rbcSelCInclude(0).Value Then 'Advt/Contract
                        If Value Then
                            lbcSelection(0).Visible = False
                            lbcSelection(5).Visible = False
                        Else
                            lbcSelection(0).Visible = True
                            lbcSelection(5).Visible = True
                            llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                            llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        End If
                        'llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        'llRet = SendMessagebyNum(lbcSelection(2).Hwnd, LB_SELITEMRANGE, ilValue, llRg)
                        'llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                        'llRet = SendMessagebyNum(lbcSelection(3).Hwnd, LB_SELITEMRANGE, ilValue, llRg)
                    Else                            'slsp or vehicle
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                Case 5  'Recap
                Case 6, 7  'Placement; Discrepancy

                Case 8  'MG
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case 9  'Sales Spot Tracking
                Case 10, 12 'Commercial Change, Affiliate Spot Tracking
                Case 11 'History
                    llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    If Value Then
                        'lbcSelection(0).Visible = False
                        'lbcSelection(5).Visible = False
                    Else
                        lbcSelection(0).Visible = True
                        lbcSelection(5).Visible = True
                    End If
                Case 13 'Spot Sales
                Case 17    'Quarterly Avails
                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case CNT_ADVT_UNITS    ' Adv Units, Avg Rate
                    'Date: 9/2/2019 added same major/minor sort options for ADVT UNITS
                    If rbcSelCInclude(0).Value Then                  'advt
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(1).Value Then   'slsp
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(2).Value Then             'Agency option
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(3).Value Then             'bus cat
                        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(4).Value Then             'prod prot
                        llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(5).Value Then     'vehicles
                        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(6).Value Then         'vg
                        llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                Case CNT_AVGRATE   ' Adv Units, Avg Rate
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case CNT_AVG_PRICES 'Avg Spot Price
                    'Date: 11/1/2019 added Major/Minor sort, using CSI calender for date entry
'                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
'                    llRet = SendMessageByNum(lbcSelection(6).hWnd, LB_SELITEMRANGE, ilValue, llRg)
'                    llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
'                    llRet = SendMessageByNum(lbcSelection(2).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                    
                    If rbcSelCInclude(0).Value Then                  'advt
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(1).Value Then   'slsp
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(2).Value Then             'Agency option
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(3).Value Then             'bus cat
                        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(4).Value Then             'prod prot
                        llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(5).Value Then     'vehicles
                        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(6).Value Then         'vg
                        llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                    
                Case CNT_SALES_CPPCPM       'sales analysis by cpp & cpm
                    llRg = CLng(lbcSelection(11).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(11).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case CNT_TIEOUT
                    If rbcSelCInclude(1).Value Then             'vehicle
                        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    Else                                        'office
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                Case CNT_BOB, CNT_BOBRECAP                                    'Billed & Booked, & B & B Recap
                    If rbcSelCInclude(2).Value Or rbcSelCInclude(4).Value Or rbcSelCInclude(6).Value Then        '8-4-00 vehicle option, vehicle/part, vehicle net-net
                        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(1).Value Or rbcSelCInclude(3).Value Then  'slsp or owners
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(5).Value Then  'agencies 4-12-02
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    Else                                        'advt
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                Case CNT_SALESCOMPARE
                    If rbcSelCInclude(0).Value Then                  'advt
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(1).Value Then   'slsp
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(2).Value Then             'Agency option
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(3).Value Then             'bus cat
                        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(4).Value Then             'prod prot
                        llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(5).Value Then     'vehicles
                        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf rbcSelCInclude(6).Value Then         'vg
                        llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
                Case CNT_CUMEACTIVITY
                    If rbcSelCInclude(0).Value Then             'adv option
                        ilSelect = 5
                    ElseIf rbcSelCInclude(1).Value Then         'agy
                        ilSelect = 1
                    ElseIf rbcSelCInclude(2).Value Then         'demo
                        ilSelect = 11
                    Else
                        ilSelect = 6
                    End If
                    llRg = CLng(lbcSelection(ilSelect).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(ilSelect).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case CNT_MAKEPLAN, CNT_VEHCPPCPM
                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    llRg = CLng(lbcSelection(4).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(4).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case CNT_SALESACTIVITY_SS, CNT_SALESPLACEMENT, CNT_VEH_UNITCOUNT, CNT_LOCKED, CNT_GAMESUMMARY           '4-5-06
                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case CNT_PAPERWORKTAX
                    llRg = CLng(lbcSelection(11).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(11).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Case CNT_BOBCOMPARE             '9-13-07 Billed and Booked Comparisons
                    If cbcSet2.ListIndex = 0 Then                 'advt
                        llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf cbcSet2.ListIndex = 4 Then   'slsp
                        llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf cbcSet2.ListIndex = 1 Then             'Agency option
                        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf cbcSet2.ListIndex = 2 Then             'bus cat
                        llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf cbcSet2.ListIndex = 3 Then             'prod prot
                        llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    ElseIf cbcSet2.ListIndex = 5 Then     'vehicles
                        llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                        llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    End If
            End Select
        End If
    'Else
        imAllClicked = False
    End If
    mSetCommands
    mEnableSeparateFiles
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
            'If ilIndex = 17 Then
            '    ilValue = Value
            '    If imSetAllAAS Then
            '        imAllClickedAAS = True
            '        llRg = CLng(lbcSelection(12).ListCount - 1) * &H10000 Or 0
            '        llRet = SendMessagebyNum(lbcSelection(12).Hwnd, LB_SELITEMRANGE, ilValue, llRg)
            '    End If
            'End If
        'Else
            If (igRptType = 0) And (ilIndex > 1) Then
                ilIndex = ilIndex + 1
            End If
            If ilIndex = CNT_BR Or ilIndex = CNT_INSERTION Then
                ckcAll.Visible = False
                imSetAll = False
                ckcAll.Value = vbUnchecked  '9-12-02 False
                imSetAll = True
                If tgUrf(0).iSlfCode = 0 Then
                    ckcSelC8(0).Enabled = True                      'allow to show mods as differences if requesting all
                End If
                'If rbcSelCSelect(0).Value Then          'advt option
                '    lbcSelection(0).Visible = False     'cnt list box
                '    lbcSelection(5).Visible = False     'advt list box
                'ElseIf rbcSelCSelect(1).Value Then      'agy option
                '    lbcSelection(10).Visible = False    'cnt list box
                '    lbcSelection(8).Visible = False     'agy list box
                'Else                                    'slsp option
                '    lbcSelection(10).Visible = False    'cnt list box
                '    lbcSelection(9).Visible = False     'slsp list box
                'End If
                llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                llRg = CLng(lbcSelection(9).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(9).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                lbcSelection(10).Clear
                lbcSelection(0).Clear
'            ElseIf ilIndex = CNT_SPTSBYADVT Then
'                llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
'                llRet = SendMessageByNum(lbcSelection(6).hwnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = CNT_VEHCPPCPM Then
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = CNT_BOB Then
                If rbcSelCInclude(1).Value And ckcSelC10(1).Value = vbChecked Then        '8-4-00  slsp option with vehicle sub-totals
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
                If rbcSelCInclude(4).Value Then      '8-4-00 vehicle/participant
                    llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If

            ElseIf ilIndex = CNT_QTRLY_AVAILS Then
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = CNT_SALESACTIVITY_SS Or ilIndex = CNT_SALESPLACEMENT Then
                llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = CNT_SALESCOMPARE Then
                If cbcSet2.ListIndex = 1 Then                 'advt
                    llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 2 Then   'agency
                    llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 3 Then             'bus cat
                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 4 Then             'prod prot
                    llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 5 Then             'slsp
                    llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 6 Then     'vehicles
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 7 Then           'vg with items
                    llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
            ElseIf ilIndex = CNT_BOBCOMPARE Then
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = CNT_AVG_PRICES Then      '12-26-08
'                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
'                    llRet = SendMessageByNum(lbcSelection(3).hWnd, LB_SELITEMRANGE, ilValue, llRg)
                
                'Date: 11/1/2019 added Major/Minor sorts, used CSI calendar for date entry
                If cbcSet2.ListIndex = 1 Then                 'advt
                    llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 2 Then   'agency
                    llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 3 Then             'bus cat
                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 4 Then             'prod prot
                    llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 5 Then             'slsp
                    llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 6 Then     'vehicles
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 7 Then           'vg with items
                    llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
            ElseIf ilIndex = CNT_AVGRATE Then                       '12-9-16
                    llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
               ' End If
            ElseIf ilIndex = CNT_ADVT_UNITS Then                       '6-21-18
'                    llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
'                    llRet = SendMessageByNum(lbcSelection(5).hWnd, LB_SELITEMRANGE, ilValue, llRg)
            
                If cbcSet2.ListIndex = 1 Then                 'advt
                    llRg = CLng(lbcSelection(5).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(5).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 2 Then   'agency
                    llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 3 Then             'bus cat
                    llRg = CLng(lbcSelection(3).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(3).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 4 Then             'prod prot
                    llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 5 Then             'slsp
                    llRg = CLng(lbcSelection(2).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(2).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 6 Then     'vehicles
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                ElseIf cbcSet2.ListIndex = 7 Then           'vg with items
                    llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                End If
            End If
        ElseIf igRptCallType = SLSPCOMMSJOB Then
            If ilIndex = COMM_SALESCOMM Then
                llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            End If
        End If
    Else                                                'turned All AAS off
        If igRptCallType = CONTRACTSJOB Then
            If (igRptType = 0) And (ilIndex > 1) Then
                ilIndex = ilIndex + 1
            End If
            If ilIndex = 17 Or ilIndex = CNT_AVGRATE Then       '12-9-16
                ilValue = Value
                If imSetAllAAS Then
                    imAllClickedAAS = True
                    llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Else
                    imAllClickedAAS = False
                End If
            End If
            If ilIndex = CNT_BR Or ilIndex = CNT_INSERTION Then
                'show the AAS boxes again
                mSetupPopAAS                               'setup list box of valid contracts
            End If
        ElseIf igRptCallType = SLSPCOMMSJOB Then
            If ilIndex = COMM_SALESCOMM Then
                If imSetAllAAS Then
                    imAllClickedAAS = True
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                Else
                    imAllClickedAAS = False
                End If
            End If
        End If
    End If
    imAllClickedAAS = False
    mSetCommands
    mEnableSeparateFiles
End Sub

Private Sub CkcAllVeh_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If CkcAllveh.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilIndex As Integer
    ilIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAllVeh Then
        imAllClickedVeh = True
        If igRptCallType = CONTRACTSJOB Then
            If ilIndex = CNT_INSERTION Then
                If lbcSelection(6).ListCount > 0 Then       'select all veh
                    llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                    llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
                    mEnableSeparateFiles
                End If
            ElseIf ilIndex = CNT_SALESACTIVITY_SS Or ilIndex = CNT_SALESPLACEMENT Then      '7-25-02
                llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ilIndex = CNT_BOB Then
                llRg = CLng(lbcSelection(7).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(7).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            ElseIf ((ilIndex = CNT_SALESCOMPARE) Or (ilIndex = CNT_ADVT_UNITS) Or (ilIndex = CNT_AVG_PRICES)) Then
                'Date:  9/9/2019 added check for ADV UNITS
                '       11/9/2019 added check for AVG_PRICES
                llRg = CLng(lbcSelection(8).ListCount - 1) * &H10000 Or 0
                llRet = SendMessageByNum(lbcSelection(8).HWnd, LB_SELITEMRANGE, ilValue, llRg)
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
            If ilListIndex = CNT_BOB Then
                'If Index = 1 And rbcSelCInclude(1).Value Then      '12-3-00 slsp option, sub-totals by vehicle
                If rbcSelCInclude(1).Value Then      '12-3-00 slsp option, sub-totals by vehicle

                    If Index = 1 And Value Then               'turn on vehicle selection for slsp option
                        'Show the vehicle list box for selection
                        ckcAllAAS.Move lbcSelection(2).Left, lbcSelection(2).Top + 1500
                        ckcAllAAS.Caption = "All Vehicles"
                        ckcAllAAS.Visible = True
                        lbcSelection(2).Height = 1500
                        lbcSelection(6).Move lbcSelection(2).Left, ckcAllAAS.Top + ckcAllAAS.Height, lbcSelection(6).Width, lbcSelection(2).Height + 300
                        lbcSelection(6).Visible = True
                        'rbcSelC7(2).Visible = True      'Triple net option
                        rbcSelC7(2).Caption = "T-Net"

                        'if vehicle selection is turned on, need 3 list boxes:  slsp & office which take up top half of section
                        '3rd box is vehicle, taking up bottom half of section
                        lbcSelection(2).Height = 1500
                        lbcSelection(7).Height = 1500
                    ElseIf Index = 2 Then               'office subtotals
                        'do nothing
                    Else                    'turn off vehicle selection
                        ckcAllAAS.Visible = False
                        lbcSelection(2).Height = 3270       'slsp
                        lbcSelection(6).Move lbcSelection(2).Left, lbcSelection(2).Top, 4380, 3370
                        lbcSelection(6).Visible = False

                        lbcSelection(7).Height = 3270       'sales office

                    End If
                ElseIf rbcSelCInclude(2).Value = True Then        'vehicle option, was subtotals by slsp selected?
                    rbcSelC7(2).Visible = True                    'allow t-net subtotals
                    If Index = 1 Then                           '8-6-10 allow slsp subtotal splits with vehicle sort
                        'no need to populate as all slsp and sales offices will be included when igSwapBOBOption is set
                        'mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office, wont be shown to user
                        'mSalesOfficePop lbcSelection(7)     'sales office list, this wont be shown to user
                        ckcSelC10(0).Enabled = True
                        If Value Then                            'slsp subtotals turned on
                            ckcSelC10(0).Value = vbChecked                    'assume user wants splits with slsp
                        Else
                            ckcSelC10(0).Value = vbUnchecked
                        
                        End If
                    ElseIf Index = 0 Then
                        If Value Then
                            ckcSelC10(1).Value = vbChecked
                        End If
                    End If
                Else                                                'not slsp or vehicle option, allow t-net for other options
                    rbcSelC7(2).Visible = True
                End If
            ElseIf ilListIndex = CNT_SALESCOMPARE Then
                If Index = 0 Then                           'top down control?
                    If Value = True Then                    'top down requested, disallow minor subtotals
                        cbcSet2.ListIndex = 0
                        cbcSet2.Enabled = False
                        edcSet2.Enabled = False
                        smPaintCaption4 = "Totals by"
                        plcSelC4_Paint

                        plcSelC4.Move plcSelC2.Left, cbcSet1.Top + cbcSet1.Height + 30, 4380
                        rbcSelC4(0).Caption = "Advertiser"
                        rbcSelC4(0).Left = 900
                        rbcSelC4(0).Width = 1200
                        rbcSelC4(0).Visible = True
                        rbcSelC4(0).Value = True
                        If rbcSelC4(0).Value Then             'default to advt
                            rbcSelC4_click 0
                        Else
                            rbcSelC4(0).Value = True
                        End If
                        rbcSelC4(1).Caption = "Summary"
                        rbcSelC4(1).Left = 2160
                        rbcSelC4(1).Width = 1200
                        rbcSelC4(1).Visible = True
                        rbcSelC4(2).Visible = False
                        plcSelC4.Visible = True

                        'include advt totals; show polticals are separate group are disabled for top down
                        'plcSelC13.Visible = False

                        plcSelC13.Move 120, plcSelC4.Top + plcSelC4.Height, 4380, 240
                        ckcSelC13(2).Move 0, 0
                        ckcSelC13(3).Move 3240, 0, 1200
                        ckcSelC13(0).Visible = False
                        ckcSelC13(1).Visible = False
                        ckcSelC13(3).Visible = True
                    Else                        'not top down, re-enable minor total option
                        cbcSet2.Enabled = True
                        edcSet2.Enabled = True
                        plcSelC4.Visible = False
                        plcSelC13.Move 120, cbcSet1.Top + cbcSet1.Height, 4380, 480
                        ckcSelC13(0).Move 0, 0, 2040
                        ckcSelC13(1).Move 2160, 0, 2360
                        ckcSelC13(2).Move 0, 240, 4380
                        ckcSelC13(3).Move 3240, 240, 1200
                        ckcSelC13(0).Visible = True
                        ckcSelC13(1).Visible = True
                        ckcSelC13(3).Visible = True
                        plcSelC13.Visible = True


                    End If
                End If
            End If
    End Select
    mSetCommands
End Sub

Private Sub ckcSelC12_Click(Index As Integer)
    If Index = 2 And ckcSelC12(Index).Value = vbChecked Then
        ckcSelC12(1).Value = vbChecked
    ElseIf Index = 1 And ckcSelC12(1).Value = vbUnchecked Then
        ckcSelC12(2).Value = vbUnchecked
    End If
End Sub

Private Sub ckcSelC13_Click(Index As Integer)
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex

    If igRptCallType = CONTRACTSJOB Then
        Select Case ilListIndex
            Case CNT_BOB
                If ckcSelC13(Index).Value = vbChecked Then
                    rbcSelC7(2).Visible = False
                    rbcSelC7(0).Value = True
                Else
                    rbcSelC7(2).Visible = True
                End If
            
            Case CNT_SALESACTIVITY_SS
                '09/28/2020 - TTP # 9952 - IF include NTR, Add option to split NTR (or by default: leave NTR grouped together)
                If ckcSelC13(1).Value = vbChecked Or ckcSelC13(2).Value = vbChecked Then  'NTR' or 'Hard Cost'
                    'Include NTR checked, Show Split/Group NTR option
                    'plcSelC14.Visible = True  'Disable this TFN -TTP # 9952
                Else
                    'Include NTR NOT checked, Hide Split/Group NTR option
                    plcSelC14.Visible = False
                    rbcSelC14(0).Value = True
                    rbcSelC14(1).Value = False
                End If
                mSetCommands
            
            Case CNT_DAILY_SALESACTIVITY, CNT_SALESACTIVITY
                '2020-10-30 - TTP # 9955 - add AirTime, NTR, Hardcost Include options to report (Daily Sales Activity by Contract Report, Weekly Sales Activity by Qtr)
                If ckcSelC13(0).Value = vbChecked And (ckcSelC13(1).Value = vbChecked Or ckcSelC13(2).Value = vbChecked) Then
                    'Separate AirTime, NTR, HC option (When Airtime and [NTR or HC] is checked)
                    plcSelC10.Visible = True
                    ckcSelC10(0).Visible = True
                Else
                    plcSelC10.Visible = False
                    ckcSelC10(0).Value = vbUnchecked
                End If
                mSetCommands
        End Select
    End If
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
            If ilListIndex = CNT_BR Then
                If ckcSelC3(0).Value = vbChecked Then
                    rbcOutput(1).Value = True           'force to output
                    rbcOutput(0).Enabled = False       'disallow output method to be changed
                    rbcOutput(1).Enabled = False
                    rbcOutput(2).Enabled = False
                Else
                    rbcOutput(0).Value = True
                    rbcOutput(0).Enabled = True
                    rbcOutput(1).Enabled = True
                    rbcOutput(2).Enabled = True
                End If
            ElseIf ilListIndex = CNT_PAPERWORK Then             '11-7-16 if Rev selected, cannot see holds/orders with them (elminiate showing duplicate)
                If Index = 6 And Value = True Then                               'reject selected?
                    ckcSelC3(0).Value = False
                    ckcSelC3(1).Value = False
                ElseIf (Index = 0 Or Index = 1) And Value = True Then
                    ckcSelC3(6).Value = False
                End If
            End If
        Case SLSPCOMMSJOB
            If ilListIndex = COMM_SALESCOMM Then
                If Value = True Then            'bonus comm version (new & increased sales)
                    edcSelCTo.Text = 1          'default to only 1 month
                    edcSelCTo.Enabled = False
                    'plcSelC1.Enabled = False
                    rbcSelCSelect(0).Enabled = False
                    rbcSelCSelect(1).Enabled = False
                    'plcSelC7.Enabled = False
                    rbcSelC7(1).Value = True        'default to summary
                    'plcSelC8.Enabled = False
                    ckcSelC8(0).Enabled = False
                    'plcSelC11.Enabled = False
                    rbcSelC11(0).Enabled = False
                    rbcSelC11(1).Enabled = False
                    rbcSelC11(2).Enabled = False
                Else
                    edcSelCTo.Enabled = True
                    rbcSelCSelect(0).Enabled = True
                    rbcSelCSelect(1).Enabled = True
                    rbcSelC7(0).Value = True
                    ckcSelC8(0).Enabled = True
                    rbcSelC11(0).Enabled = True
                    rbcSelC11(1).Enabled = True
                    rbcSelC11(2).Enabled = True
                End If
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
            If ilListIndex = CNT_BOB_BYCNT Or ilListIndex = CNT_BOB_BYSPOT Or ilListIndex = CNT_BOB_BYSPOT_REPRINT Then
                lbcSelection_Click 5
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
            If ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION Then
                If Index = 1 And Value Then             'incl research clicked (if on, force rates to be included too)
                    'ckcSelC6(0).Value = True
                    'ckcSelC6(0).Enabled = False
                    plcSelC9.Visible = False
                Else
                    'ckcSelC6(0).Enabled = True

                    'Temporary patch until ABC has authorized it
                    plcSelC9.Visible = False       'True
                End If
            End If
    End Select
    mSetCommands
End Sub

Private Sub ckcSelC8_Click(Index As Integer)
    '09/29/2020 - TTP # 9952 - IF include NTR, Add option to split NTR (or by default: leave NTR grouped together)
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
        If ilListIndex = CNT_CUMEACTIVITY Then
            If ckcSelC8(1).Value = vbChecked Or ckcSelC8(2).Value = vbChecked Then  'NTR' or 'Hard Cost'
                'Include NTR checked, Show Split/Group NTR option
                'plcSelC14.Visible = True 'Disable this TFN - TTP # 9952
            Else
                'Include NTR NOT checked, Hide Split/Group NTR option
                plcSelC14.Visible = False
                rbcSelC14(0).Value = True
                rbcSelC14(1).Value = False
            End If
        End If
End Sub

Private Sub ckcSeparateFile_Click()
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If (igRptType = 0) And (ilListIndex > 1) Then
        ilListIndex = ilListIndex + 1
    End If
    edcFileName.Enabled = True
    cmcBrowse.Enabled = True
    
    If ckcSeparateFile.Value = 1 Then
        If ilListIndex = CNT_INSERTION Then
            'Separate file mode enabled
            edcFileName.Text = "IO"
            edcFileName.Enabled = False
            cmcBrowse.Enabled = False
        End If
    End If
End Sub

Private Sub cmcBrowse_Click()
    'Dan M 8/17/10
    gAdjustCDCFilter imFTSelectedIndex, cdcSetup
    cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    cdcSetup.fileName = edcFileName.Text
    cdcSetup.InitDir = Left$(sgRptSavePath, Len(sgRptSavePath) - 1)
    cdcSetup.Action = 2    'DLG_FILE_SAVE
    edcFileName.Text = cdcSetup.fileName
    mSetCommands
    gChDrDir        '3-25-03
    'ChDrive Left$(sgCurDir, 2)  'Set the default drive
    'ChDir sgCurDir              'set the default pathrDir, 2)
    ChDir sgCurDir
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llcomparestartdates           slLastYear                                              *
'******************************************************************************************
    
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilListIndex As Integer
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim illoop As Integer
    ReDim llStdStartDates(0 To 13) As Long  'Index zero ignored
    Dim llLastBilled As Long
    Dim llBillCycleLastBilled As Long               '1-13-21 for billing method B & B, last cal date billed
    Dim ilLastBilledInx As Integer
    Dim ilBillCycleLastBilledInx As Integer         '1-13-21 for billing method B & b, inx into date array for past/future
    Dim slEarliest As String
    Dim slLatest As String
    Dim llEarliestDate As Long
    Dim llLatestDate As Long
    Dim slStr As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim llCompareDate As Long
    Dim llCompareDate2 As Long
    Dim ilTemp As Integer
    Dim ilSaveMonth As Integer
    Dim slMonthHdr As String * 36
    Dim llPacingDate As Long        '11-21-05
    Dim ilDoIt As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilFilterBRPass As Integer
    Dim slPDFFileName As String
    Dim slStartOfStdTY As String  '2-19-16
    Dim slStartOfStdLY As String  '2-19-16
    Dim slSavePacingDateEntered As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim blOwnerOnly As Boolean      '3-31-16  Process vehicle by vehicle group participant, do not split the particpants but process the owner as 100%.
                                    'required so that future changes of ownership are picked up
    Dim blBillCycle As Boolean      '1-13-21 For Billed and Booked, use billing cycle for results
    Dim blGetCalLastBilled As Boolean   '1-14-21  when getting b & b dates, get the calendar last billed too
    
    'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
    Dim ilNumberOfExports As Integer
    Dim ilExportSplit As Integer
    
    Dim ilValue As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim ilCkcAllVehState As Integer
    ilCkcAllVehState = CkcAllveh.Value
    
    'TTP 10402 - Advertiser Units Ordered Report gives error and shutdown if you enter more than 13 weeks
    ilListIndex = RptSelCt!lbcRptType.ListIndex
    If (igRptCallType = CONTRACTSJOB) And (ilListIndex = CNT_ADVT_UNITS) Then              'advert units sold
        slStr = RptSelCt!edcSelCFrom1.Text       '14-week Qtr max
        ilRet = gVerifyInt(slStr, 1, 14)
        If ilRet = -1 Then                       'bad conversion or illegal #
            MsgBox ("Please enter # Weeks: (Max 14)")
            RptSelCt!edcSelCFrom1.SetFocus       'invalid # weeks
            Exit Sub
        End If
    End If
    
    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    bmHasRecords = False
    
    'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
    If ckcSeparateFile.Value = 1 Then
        'Separate Export for each contract
        mReportSeparateOutputVehicle
        ilNumberOfExports = UBound(tmVehicleList)
    Else
        ilNumberOfExports = 1
    End If
        
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
    cmcGen.Enabled = False
    
    'TTP 10271: STARTOF INDIVIDUAL CONTRACTS LOOP
    For ilExportSplit = 0 To ilNumberOfExports - 1
        If ckcSeparateFile.Value = 1 Then
            'edcTopHowMany.Text = tmContractList(ilExportSplit).lCntrNo
            'Unselect all vehicles
            ilValue = False
            llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            
            'Select 1st Vehicle to export
            lbcSelection(6).Selected(tmVehicleList(ilExportSplit)) = True
            cmcGen.Enabled = False
            bmHasRecords = False
        End If
        
        'LOGSJOB: igRptType = 0 or 2 => Log format; 1 or 3 => Delivery
        'If (igRptCallType = LOGSJOB) And ((igRptType = 0) Or (igRptType = 2)) And ((ilListIndex = 1) Or (ilListIndex = 3)) Then
        If (igRptCallType = CONTRACTSJOB) Then
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
    
            igUsingCrystal = True
    
        Else
            igUsingCrystal = True
        End If
        
        tmBillCycle.blBillCycle = False                     '1-13-21 assume not pulling the B & B by billing method .  This is the only report that uses the contracts bill cycle in header
        If (igRptCallType = CONTRACTSJOB) Then
            If ilListIndex = CNT_BR Then        'And Not rbcSelCInclude(2).Value) Then        'proposals/wide contract
                ilStartJobNo = 1
                ilNoJobs = 5            '1-7-21 chged to 5 with added form for cpm podcast
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear       'can only get the current date & time once for all versions of the contract reports
                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime      '9-14-09 get milliseconds to help prevent multiple users generating at same time
    
            ElseIf (ilListIndex = CNT_INSERTION) Then
                ilNoJobs = 1
                ilStartJobNo = 1
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear       'can only get the current date & time once for all versions of the contract reports
                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime      '9-14-09 get milliseconds to help prevent multiple users generating at same time
            End If
        Else
            ilNoJobs = 1
            ilStartJobNo = 1
        End If
            'dan 10-31-08 combine multiple report jobs into one report.  This means I have to find what the last job is going to be before going into loop.
        'find last print job if ilnojobs > 1
        If ilNoJobs > 1 Then
            Set ogReport = New CReportHelper
            'special case: last print job may not simply = ilnojobs: find if user wants to see report
            If ilListIndex = CNT_BR And igRptCallType = CONTRACTSJOB Then
                For ilJobs = ilNoJobs To ilStartJobNo - 1 Step -1
                    If ilJobs = 5 And Not rbcSelC4(0).Value Then            'summary job and user request both detail/summary (rbcSelc4(0) = both)
                        ogReport.iLastPrintJob = 5          '1-7-21 podcast adds a new form
                        Exit For
                    ElseIf ilJobs = 4 And Not (rbcSelC4(0).Value) And ckcSelC6(1).Value = vbChecked Then       ' (rbcSelc4(0) = both detail /summary; ckcselec6(1) = research)
                       ogReport.iLastPrintJob = 4    '1-7-21 job # adjusted with addition of cpm podcast form; summary w/research
                        Exit For
                    ElseIf ilJobs = 2 And Not (rbcSelC4(0).Value) And (igBR_NTRDefined) Then       'ntr
                        'any NTR?
                        ogReport.iLastPrintJob = 2
                        Exit For
                    ElseIf ilJobs = 3 And Not (rbcSelC4(0).Value) And (igBR_CPMDefined) Then        '1-8-21 cpm podcast defined
                        ogReport.iLastPrintJob = 3
                        Exit For
                    End If      ' default =  1
                Next ilJobs
            End If
        End If
        'dan end
        For ilJobs = ilStartJobNo To ilNoJobs Step 1
            'TTP 11022 - Billed and Booked report: fails to generate
            'If ogReport Is Nothing Then
            '    Exit For
            'End If
            igJobRptNo = ilJobs
            ilDoIt = True
            If ilListIndex = CNT_BR And igRptCallType = CONTRACTSJOB Then
                'ilJobs :  1 = detail, 2 = NTR, 3 = Research Summary with or without rates, 4 = Billing summary (12 month $)
                'ilJobs 1-7-21 Added CPM Podcast as job #3,others adjusted :  1 = detail, 2 = NTR, 3=CPM Podcast, 4 = Research Summary with or without rates, 5 = Billing summary (12 month $)
                If ilJobs = 1 Then          'detail
                    'has detail been requested
                    If (rbcSelC4(1).Value = False) Then           'not summary only  rbcselc4(0)-detail, (1) = summary, 2= both
                        ilDoIt = True
                    Else
                        ilDoIt = False
                    End If
                ElseIf ilJobs = 2 Then      'ntr
                    'any NTR?
                    If (rbcSelC4(0).Value = False) And (igBR_NTRDefined) Then            'user requested summary or both, and NTR defined
                        ilDoIt = True
                    Else
                        ilDoIt = False
                    End If
                    
                ElseIf ilJobs = 3 Then          '1-8-21 cpm
                'any CPM?
                    If (rbcSelC4(0).Value = False) And (igBR_CPMDefined) Then            'user requested summary or both, and CPM defined
                        ilDoIt = True
                    Else
                        ilDoIt = False
                    End If
                    
                ElseIf ilJobs = 4 Then        'summary w/research
                    If rbcSelC4(0).Value = False And ckcSelC6(1).Value = vbChecked Then   'user requested summary or both w/ research
                           ilDoIt = True
                       Else
                           ilDoIt = False
                       End If
                Else                        ' or billing summary
                    If rbcSelC4(0).Value = False Then             'user requested summary or both, and NTR defined
                           ilDoIt = True
                       Else
                           ilDoIt = False
                       End If
                End If
            End If
            Screen.MousePointer = vbHourglass
            If ilDoIt Then
                If Not gGenReportCt() Then
                    igGenRpt = False
                    frcOutput.Enabled = igOutput
                    frcCopies.Enabled = igCopies
                    'frcWhen.Enabled = igWhen
                    frcFile.Enabled = igFile
                    frcOption.Enabled = igOption
                    'frcRptType.Enabled = igReportType
                    Exit Sub
                End If
                
                ilRet = gCmcGenCt(ilListIndex, imGenShiftKey, smLogUserCode)
                '-1 is a Crystal failure of gSetSelection or gSEtFormula
                If ilRet = -1 Then
                    igGenRpt = False
                    frcOutput.Enabled = igOutput
                    frcCopies.Enabled = igCopies
                    frcFile.Enabled = igFile
                    frcOption.Enabled = igOption
                    pbcClickFocus.SetFocus
                    tmcDone.Enabled = True
                    Exit Sub
                ElseIf ilRet = 0 Then   '0 = invalid input data, stay in
                    igGenRpt = False
                    frcOutput.Enabled = igOutput
                    frcCopies.Enabled = igCopies
                    frcFile.Enabled = igFile
                    frcOption.Enabled = igOption
                    Exit Sub
                ElseIf ilRet = 2 Then           'successful return from bridge reports
                    igGenRpt = False
                    frcOutput.Enabled = igOutput
                    frcCopies.Enabled = igCopies
                    frcFile.Enabled = igFile
                    frcOption.Enabled = igOption
                    pbcClickFocus.SetFocus
                    tmcDone.Enabled = True
                    Exit Sub
                End If
            End If                      'ilDoIt
    
           '1 falls thru - successful crystal report
            If igRptCallType = SLSPCOMMSJOB Then
                    Screen.MousePointer = vbHourglass
                If ilListIndex = COMM_SALESCOMM Then    'Or ilListIndex = COMM_PROJECTION Then
                    ilRet = gBuildSlsCommCt
                ElseIf ilListIndex = COMM_PROJECTION Then
                    If gOpenBOBFilesCt() = BTRV_ERR_NONE Then
                        RptSelCt!rbcSelC9(0).Value = False
                        RptSelCt!rbcSelC9(1).Value = True       'forc to project by std bdct month dates
                        '11-21-05 add pacing date to call (n/a to this report)
                        ilTemp = (igMonthOrQtr - 1) * 3 + 1
                        gSetupBOBDates 2, llStdStartDates(), llLastBilled, ilLastBilledInx, 0, ilTemp  'build array of corp start & end dates
                        gGetAllParticipantSplits llStdStartDates(1)
                        If llStdStartDates(1) > llLastBilled Then                       'projection only
                            ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 12, 0, tmBillCycle)  '1-15-21 send the cal months dates for bill cycle method
                        Else
                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 12, 0, tmBillCycle
                            If llLastBilled + 1 < llStdStartDates(13) Then                 'past only or past & projection
                                ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 12, 0, tmBillCycle)            '1-15-21 send the cal months dates for bill cycle method
                            End If
                        End If
    
                        gCloseBOBFilesCt
                        Erase llStdStartDates
                    End If
                End If
                cbcSel.Clear
                Screen.MousePointer = vbDefault
            'If contract spot projection or quarterly avails- create records
            ElseIf (igRptCallType = CONTRACTSJOB) Then
                '11-10-03 remove Portrait contract (no client uses)
                'If ilListIndex = CNT_BR And RptSelCt!rbcSelCInclude(2).Value Then           'Portrait contract converted to Crystal 10-20-00
                If (ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION) Then   '11-10-03 And Not rbcSelCInclude(2).Value) Then       'proposals/wide contract
                    If igJobRptNo = 1 Then                        'only gen the data once
                        Screen.MousePointer = vbHourglass
                        gBRGen
                        Screen.MousePointer = vbDefault
                    Else
                        'Screen.MousePointer = vbHourGlass        'go process part 2 of BR (summary)
                    End If
                End If
                If ilListIndex = CNT_HISTORY Then           'Contract History converted to Crystal 10-20-00
                    Screen.MousePointer = vbHourglass
                    If RptSelCt!rbcOutput(1).Value Then 'print
                        gCntrRptCt True, False      'true = history (vs portrait cntr), false = print & update printables
                    Else
                        gCntrRptCt True, True     'true = history (vs portrait contract, true = display or save to fle (no updating printables flag)
                    End If
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_COMLCHG Then           'Commercial changes converted to Crystal 10-25-00
                    Screen.MousePointer = vbHourglass
                        gCmmlChgRptCt
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_AFFILTRAK Or ilListIndex = CNT_SPOTTRAK Then    'AFffiliate Spot Tracking converted to Crystal 10-24-00, Sales Spot Tracking converted 8-4-09
                    Screen.MousePointer = vbHourglass
                        gTrakAffRptCt
                    Screen.MousePointer = vbDefault
                'removed spots by advt code--see rptselcb
                ElseIf ilListIndex = CNT_BOB_BYSPOT Then    'Spot Projection
                    Screen.MousePointer = vbHourglass
                    gCRSpotProjGenCt
                    cbcSel.Clear
                    Screen.MousePointer = vbDefault
                ElseIf (ilListIndex = CNT_QTRLY_AVAILS) Then      'Quarterly Avails
                    Screen.MousePointer = vbHourglass
                    gCRQAvailsGenCt
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_SALES_CPPCPM Then          'sales analysis cpp & cpm
                    Screen.MousePointer = vbHourglass
                    gSalesCppCpmGen
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_TIEOUT Then                'Tie out
                    Screen.MousePointer = vbHourglass
                    gTieOutGenCt
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_BOB Or ilListIndex = CNT_BOBRECAP Or ilListIndex = CNT_BOBCOMPARE Then     '4-14-05 Billed & Booked
                    ilTemp = igMonthOrQtr               '7-3-08 convert to start month vs start qtr
                    Screen.MousePointer = vbHourglass
                    If gOpenBOBFilesCt() = 0 Then
                        llPacingDate = 0
                        slStr = RptSelCt!edcText.Text           'pacing date
                        If Trim$(slStr) <> "" Then
                            llPacingDate = gDateValue(slStr)
                        End If
                        
                        blGetCalLastBilled = False          '1-14-21 when getting month start dates , do not need the cal last billed date. this is for the bill cycle only on b & b
                        blOwnerOnly = False                 '3-13-16 indicates splits need to occur, but force owner to 100%.  This would pickup different owners across the year if it were changed
                        If ilListIndex = CNT_BOB And RptSelCt!rbcSelCInclude(2).Value And RptSelCt!cbcSet1.ListIndex = 1 Then   'must be billed & booked, by vehicle (rbcselcinclude(2)), and Participant VG (cbcset1)
                            blOwnerOnly = True
                        End If
                        If RptSelCt!rbcSelC9(0).Value Then      'corporate option
                            '11-21-05 add pacing date to call
                            gSetupBOBDates 1, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, ilTemp  'build array of corp start & end dates
                            gGetAllParticipantSplits llStdStartDates(1), blOwnerOnly
                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle    '7-7-14  use # periods vs always 12 periods
                            ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, igPeriods, 0, tmBillCycle) '7-7-14  use # periods vs always 12 periods;'1-15-21 send the cal months dates for bill cycle method
                       ElseIf RptSelCt!rbcSelC9(1).Value Then       'standard
                            gSetupBOBDates 2, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, ilTemp  'build array of std start & end dates
                            '7-12-01 Use airing lines when invoicing by Line so past & future will still balance to contract
                            gGetAllParticipantSplits llStdStartDates(1), blOwnerOnly
                            If llPacingDate = 0 Then                'not pacing
                                If (gUsingBarters()) And (RptSelCt!ckcSelC13(0).Value = vbChecked) And ((Asc(tgSaf(0).sFeatures2) And PAYMENTONCOLLECTION) = PAYMENTONCOLLECTION) Then     '7-25-16 barter and use acq cost instead of spot cost & Acq payment on Collections
                                    ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, igPeriods, 0, tmBillCycle) '7-7-14  use # periods vs always 12 periods;'1-15-21 send the cal months dates for bill cycle method
                                Else
                                    'bypass past if projection only if dates all in future or Calc virtual pkg by line & using airng lines, and not doing a t-net report (because merchand/promotions required from rvf/phf)
                                    If ((llStdStartDates(1) > llLastBilled) Or (tgSpf.iPkageGenMeth = 1 And RptSelCt!rbcSelCSelect(1).Value)) And Not rbcSelC7(2).Value Then    '6-19-08 test for tnet and always go thru rvf/phf
                                        ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, igPeriods, 0, tmBillCycle) '7-7-14  use # periods vs always 12 periods;'1-15-21 send the cal months dates for bill cycle method
                                    Else
                                        'gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 12, 0
                                        gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle   '7-7-14  use # periods vs always 12 periods
                                        If llLastBilled + 1 < llStdStartDates(igPeriods + 1) Then              'past only or past & projection
                                            ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, igPeriods, 0, tmBillCycle) '7-7-14  use # periods vs always 12 periods;'1-15-21 send the cal months dates for bill cycle method
                                        End If
                                    End If
                                End If
                            Else                'pacing std
                                If (gUsingBarters()) And (RptSelCt!ckcSelC13(0).Value = vbChecked) And ((Asc(tgSaf(0).sFeatures2) And PAYMENTONCOLLECTION) = PAYMENTONCOLLECTION) Then     '7-25-16 barter and use acq cost instead of spot cost & Acq payment on Collections
                                    ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, igPeriods, 0, tmBillCycle) '7-7-14  use # periods vs always 12 periods;'1-15-21 send the cal months dates for bill cycle method
                                Else
                                    gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle   '7-7-14  use # periods vs always 12 periods
                                    'TTP 10362 - 12/15/21 - JW; restored, was lost during SOS version 10 - 9/21/21
                                    ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, igPeriods, 0, tmBillCycle)     '7-7-14  use # periods vs always 12 periods;'1-15-21 send the cal months dates for bill cycle method
                                End If
                            End If
                        ElseIf RptSelCt!rbcSelC9(2).Value Then   'calendar month (by line )
                            gSetupBOBDates 3, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, ilTemp  'build array of calendar start & end dates
                            gGetAllParticipantSplits llStdStartDates(1), blOwnerOnly
                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle   '7-7-14  use # periods vs always 12 periods
                            ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, igPeriods, 0, tmBillCycle)         '7-7-14  use # periods vs always 12 periods;'1-15-21 send the cal months dates for bill cycle method
    
                        ElseIf RptSelCt!rbcSelC9(3).Value Then                                    'calendar month by spot
                            gSetupBOBDates 3, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, ilTemp  'build array of calendar start & end dates
                            gGetAllParticipantSplits llStdStartDates(1), blOwnerOnly
                            'get the adjustments and acquisitions adjustments if applicable
                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle      'pick up adjustments and any acquisition adjustments if applicable
                            ilRet = gBOBCalBySpots(llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle) 'gather the spot data, 1-20-21 pass the array for calendar bill cycle dates (wont be used for this option)
                        Else                                                                 '1-13-21 Results by billing cycle rbcselc9(4)
                            'gather past by the way the contract has been invoiced (receivables), gather future by spots (as aired)
                            gSetupBOBDates 2, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, ilTemp  'build array of std start & end dates
                            blGetCalLastBilled = True
                            gSetupBOBDates 3, tmBillCycle.lBillCycleStartDates(), tmBillCycle.lBillCycleLastBilled, tmBillCycle.iBillCycleLastBilledInx, llPacingDate, ilTemp, blGetCalLastBilled  'build array of calendar start & end dates
                            gGetAllParticipantSplits llStdStartDates(1), blOwnerOnly
                            tmBillCycle.blBillCycle = True               'set flag to indicate that the receivables should be based on how the contract is billed
                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle  '7-7-14  use # periods vs always 12 periods
                            ilRet = gBOBCalBySpots(llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle) 'gather the spot data
                        End If
    
                        If ilListIndex = CNT_BOBCOMPARE Then            'do the budgets if applicable
                            gBudgetsForBOBCompare llStdStartDates(), ilTemp         'send array of start/end dates and starting month to process
                        End If
    
                        gCloseBOBFilesCt
                        Erase llStdStartDates
                        Screen.MousePointer = vbDefault
                        Dim llPop As Long
                        Dim llTime As Long
                        ReDim ilNowTime(0 To 1) As Integer
                        slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
                        gPackTime slStr, ilNowTime(0), ilNowTime(1)
                        gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, llTime
                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, llPop   'start time of run
                        llPop = llPop - llTime              'time in seconds in runtime
                    End If
                ElseIf ilListIndex = CNT_AVGRATE Or ilListIndex = CNT_AVG_PRICES Or ilListIndex = CNT_ADVT_UNITS Then   '12-23-08
                    Screen.MousePointer = vbHourglass
                    gCrAvgRateCt (rbcOutput(4).Value) 'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_SALESACTIVITY Or ilListIndex = CNT_DAILY_SALESACTIVITY Then   '6-5-01 add Daily Sales Act
                    Screen.MousePointer = vbHourglass
                    gCrSalesActCt
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_SALESCOMPARE Then
                    Screen.MousePointer = vbHourglass
                    llPacingDate = 0                        '2-18-16 implement pacing
                    slStr = RptSelCt!edcText.Text           'pacing date
                    If Trim$(slStr) <> "" Then
                        llPacingDate = gDateValue(slStr)
                    End If
    
                    '3-20-18 implement calendar months using all contract projections
                    'Determine contracts to process based on their entered and modified dates
                     If RptSelCt!rbcSelC9(2).Value Then      'calendar
                        gSetupBOBDates 3, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, igMonthOrQtr  'build array of corp start & end dates
                    Else                    ' rbcselc9(1) = standard
                        gSetupBOBDates 2, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, igMonthOrQtr  'build array of corp start & end dates
                    End If
                    llEarliestDate = llStdStartDates(1)
                    llLatestDate = llStdStartDates(igPeriods + 1) - 1
                        
                    slEarliest = Format(llEarliestDate, "ddddd")
                    slLatest = Format(llLatestDate, "ddddd")
    
                    illoop = igPeriods
                    If RptSelCt!rbcSelC9(1).Value Then          'std- get last years processing dates
                        'determine # of months for last year, see if asking for all of last year
                        'obtain previous years dates
                        slStr = gObtainEndStd(slEarliest)
                        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                        'code was defaulted Thru Specified Month and hidden
                        If RptSelCt!rbcSelC11(1).Value = True Then      '3-23-16 include all last year (vs thru specified month)
                            illoop = 12
                            slStr = "1/15/" & Trim$(str$(Val(slYear) - 1))
                        Else
                            slStr = slMonth & "/" & "15/" & Trim$(str$(igYear) - 1)
                        End If
                        
                        slEarliest = gObtainStartStd(slStr)
                        slStr = slEarliest
                        Do While illoop <> 0
                            slLatest = gObtainEndStd(slStr)
                            slStr = gObtainStartStd(slLatest)
                            llCompareDate2 = gDateValue(slLatest)
                            llCompareDate2 = llCompareDate2 + 1
                            slStr = Format$(llCompareDate2, "m/d/yy")
                            illoop = illoop - 1
                        Loop
                    Else                            '3-20-18 cal- get last years processing dates
                        'obtain previous years dates
                        slStr = gObtainEndCal(slEarliest)
                        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                        'code was defaulted Thru Specified Month and hidden
                        If RptSelCt!rbcSelC11(1).Value = True Then      '3-23-16 include all last year (vs thru specified month)
                            illoop = 12
                            slStr = "1/15/" & Trim$(str$(Val(slYear) - 1))
                        Else
                            slStr = slMonth & "/" & "15/" & Trim$(str$(igYear) - 1)
                        End If
                        
                        slEarliest = gObtainStartCal(slStr)
                        slStr = slEarliest
                        Do While illoop <> 0
                            slLatest = gObtainEndCal(slStr)
                            slStr = gObtainStartCal(slLatest)
                            llCompareDate2 = gDateValue(slLatest)
                            llCompareDate2 = llCompareDate2 + 1
                            slStr = Format$(llCompareDate2, "m/d/yy")
                            illoop = illoop - 1
                        Loop
                    End If
                    
                    llCompareDate = gDateValue(slEarliest)
                    llCompareDate2 = llCompareDate2 - 1
                    'Determine  month requested, and retrieve all History and Receivables
                    'records that fall within the beginning of the cal year and end of calendar month requested
    
                    If gOpenBOBFilesCt() = 0 Then
                        gGetAllParticipantSplits llCompareDate
                        For illoop = 0 To UBound(llStdStartDates)
                            llStdStartDates(illoop) = 0
                        Next illoop
                        'this year
                        llStdStartDates(1) = llEarliestDate
                        llStdStartDates(2) = llLatestDate + 1
                        'Process Base comparison dates
                        '7-13-01 if billing by line and requesting airing lines, ignore the receivables because
                        'the total wont match the contract
                        If (tgSpf.iPkageGenMeth = 1 And RptSelCt!rbcSelCSelect(1).Value) Then
                            ilLastBilledInx = 1
                            ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 1, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                        Else
                            If llPacingDate = 0 Then                    '2-18-16 not pacing
                                If RptSelCt!rbcSelC9(2).Value Then          '3-20-18 use calendar year (vs std,rbcselc9(2)) - retrieve adjustments and always project with contracts
                                    gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 1, tmBillCycle       'calendar reporting
                                    ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 1, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                Else                                        'std bdcst
                                    'base dates
                                    If llLastBilled < llEarliestDate Then        'everything is in the future
                                        If rbcSelC7(2).Value = True Then        'if t-net, need to adjust merchandising out
                                            ilLastBilledInx = 1
                                            'get any merchandising for future (after last billing until end of requested current year)
            
                                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 2, 1, tmBillCycle
                                        End If
                                        llStdStartDates(1) = llEarliestDate     're-establish the earliest/latest dates in case needed to go into receivables
                                        llStdStartDates(2) = llLatestDate + 1
                                        ilLastBilledInx = 1
                                        ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 1, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                    ElseIf llLastBilled >= llLatestDate Then       'everything in past
                                        ilLastBilledInx = 1
                                        'gCRBob llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 1
                                        '3/22/99
                                        gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 1, tmBillCycle
                                    Else                                        'past & future
                                        llStdStartDates(1) = llEarliestDate
                                        llStdStartDates(2) = llLastBilled + 1
                                        llStdStartDates(3) = llLatestDate + 1
                                        ilLastBilledInx = 1
                                        'gCRBob llStdStartDates(), llLastBilled, ilLastBilledInx, 2, 1
                                        '3/22/99
                                        gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 2, 1, tmBillCycle
                                        ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 2, 1, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                    End If
                                End If
                            Else                                        '2-18-16 pacing
                               '3-22-16 if pacing,  get the adjustments only prior to last billed
                                ilLastBilledInx = 1
                                'get any merchandising for future (after last billing until end of requested current year)
                                gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 2, 1, tmBillCycle
                                llStdStartDates(1) = llEarliestDate     're-establish the earliest/latest dates in case needed to go into receivables
                                llStdStartDates(2) = llLatestDate + 1
                                ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 1, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                            End If
    
                        End If
                        
                        'Process Comparison dates, last year
                        For illoop = 0 To UBound(llStdStartDates)
                            llStdStartDates(illoop) = 0
                        Next illoop
    
                        llStdStartDates(1) = llCompareDate
                        llStdStartDates(2) = llCompareDate2 + 1
                        '7-13-01 if billing by line and requesting airing lines, ignore the receivables because
                        'the total wont match the contract
                        If (tgSpf.iPkageGenMeth = 1 And RptSelCt!rbcSelCSelect(1).Value) Then
                            ilLastBilledInx = 1
                            ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 2, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                        Else
                            If llPacingDate = 0 Then
                                If RptSelCt!rbcSelC9(2).Value Then          '3-20-18 use calendar year (vs std, rbcselc9(2)) - retrieve adjustments and always project with contracts
                                    gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 2, tmBillCycle       'calendar reporting
                                    ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 2, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                Else
                                    If llLastBilled < llCompareDate Then        'everything is in the future
                                        If rbcSelC7(2).Value = True Then        'if t-net, need to adjust merchandising out
                                            ilLastBilledInx = 1
            
                                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 2, tmBillCycle
                                        End If
                                        ilLastBilledInx = 1
                                        ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 2, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                    ElseIf llLastBilled > llCompareDate2 Then       'everything in past
                                        ilLastBilledInx = 1
                                        'gCRBob llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 2
                                        gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 2, tmBillCycle
                                    Else                                        'past & future
                                        llStdStartDates(1) = llCompareDate
                                        llStdStartDates(2) = llLastBilled + 1
                                        llStdStartDates(3) = llCompareDate2 + 1
                                        ilLastBilledInx = 1
                                        'gCRBob llStdStartDates(), llLastBilled, ilLastBilledInx, 2, 2
                                        '3/22/99
                                        gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 2, 2, tmBillCycle
                                        ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 2, 2, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                    End If
                                End If
                            Else                                        '2-18-16 pacing last year
                                If RptSelCt!rbcSelC9(2).Value Then      '3-20-18 implement calendar
                                    'need to calculate the date of last years effective pacing date; based on the # days difference from start of std year
                                    slStartOfStdTY = gObtainYearStartDate(1, RptSelCt!edcText.Text)      'start of current std year for the pacing date; get the difference of pacing date entered against start date of the std year
                                    illoop = llPacingDate - (gDateValue(slStartOfStdTY))                '# days difference for This year
                                    'get the year of the pacing date and backup to previous year
                                    gObtainMonthYear 1, slStartOfStdTY, ilMonth, ilYear
                                    slStartOfStdLY = gObtainYearStartDate(1, "01/15/" & Trim(str(ilYear - 1)))     'start std date of last year
                                    'new effective date
                                    llPacingDate = gDateValue(slStartOfStdLY) + illoop
                                    'Save the current entered effective pacing date, it will be overlapped with the calculated last years effective date.
                                    'once done processing, restore it to the users orig date entered
                                    slSavePacingDateEntered = Trim$(RptSelCt!edcText.Text)
                                    RptSelCt!edcText.Text = Format$(llPacingDate, "m/d/yy")
                                Else                    'std last year pacing date
                                    'need to calculate the date of last years effective pacing date; based on the # days difference from start of cal year
                                    slStartOfStdTY = gObtainYearStartDate(0, RptSelCt!edcText.Text)      'start of current std year for the pacing date; get the difference of pacing date entered against start date of the std year
                                    illoop = llPacingDate - (gDateValue(slStartOfStdTY))                '# days difference for This year
                                    'get the year of the pacing date and backup to previous year
                                    gObtainMonthYear 0, slStartOfStdTY, ilMonth, ilYear
                                    slStartOfStdLY = gObtainYearStartDate(0, "01/15/" & Trim(str(ilYear - 1)))     'start std date of last year
                                    'new effective date
                                    llPacingDate = gDateValue(slStartOfStdLY) + illoop
                                    'Save the current entered effective pacing date, it will be overlapped with the calculated last years effective date.
                                    'once done processing, restore it to the users orig date entered
                                    slSavePacingDateEntered = Trim$(RptSelCt!edcText.Text)
                                    RptSelCt!edcText.Text = Format$(llPacingDate, "m/d/yy")
                                
                                End If
    
                                If rbcSelC7(2).Value = True Then        'if t-net, need to adjust merchandising out
                                    ilLastBilledInx = 1
                                    gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 2, tmBillCycle
                                End If
                                ilLastBilledInx = 1
                                ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 2, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                
                                'finished pacing for last year, now do actuals for last year same requested dates
                                RptSelCt!edcText.Text = ""                  'no pacing for actuals last year
                                If (tgSpf.iPkageGenMeth = 1 And RptSelCt!rbcSelCSelect(1).Value) Then
                                    ilLastBilledInx = 1
                                    ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 3, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                Else
                                    If llLastBilled < llCompareDate Then        'everything is in the future
                                        If rbcSelC7(2).Value = True Then        'if t-net, need to adjust merchandising out
                                            ilLastBilledInx = 1
            
                                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 3, tmBillCycle
                                        End If
                                        ilLastBilledInx = 1
                                        ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 1, 3, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                    ElseIf llLastBilled > llCompareDate2 Then       'everything in past
                                        ilLastBilledInx = 1
                                        'gCRBob llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 2
                                        gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 1, 3, tmBillCycle
                                    Else                                        'past & future
                                        llStdStartDates(1) = llCompareDate
                                        llStdStartDates(2) = llLastBilled + 1
                                        llStdStartDates(3) = llCompareDate2 + 1
                                        ilLastBilledInx = 1
                                        'gCRBob llStdStartDates(), llLastBilled, ilLastBilledInx, 2, 2
                                        '3/22/99
                                        gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, 2, 3, tmBillCycle
                                        ilRet = mBobBuildProjCt(llStdStartDates(), llLastBilled, 2, 3, tmBillCycle)     '1-15-21 send the cal months dates for bill cycle method
                                    End If
                                    
                                    'restore the orig pacing date entered, was overlayed with last years pacing date to use
                                    RptSelCt!edcText.Text = slSavePacingDateEntered
                                    
                                End If
                            End If
                        End If
                        gCloseBOBFilesCt
                        Erase llStdStartDates
                        Screen.MousePointer = vbDefault
                    End If
                
                ElseIf ilListIndex = CNT_CUMEACTIVITY Then
                    Screen.MousePointer = vbHourglass
                    '11-21-05 add pacing date to call (n/a to this report)
                    ilTemp = (igMonthOrQtr - 1) * 3 + 1
                    gSetupBOBDates 2, llStdStartDates(), llLastBilled, ilLastBilledInx, 0, ilTemp 'build array of start & end dates
                    gCrCumeActCt llStdStartDates()
                    Screen.MousePointer = vbDefault
                
                ElseIf ilListIndex = CNT_SALESACTIVITY_SS Then  '8-01-02
                    slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
                    Screen.MousePointer = vbHourglass
                    slStr = RptSelCt!edcSelCTo1.Text             'month in text form (jan..dec)
                    gGetMonthNoFromString slStr, ilSaveMonth         'getmonth #
    
                    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                        ilSaveMonth = Val(slStr)
                    End If
    
                    If rbcSelCInclude(0).Value Then         'corp
                        illoop = 1
                    ElseIf rbcSelCInclude(1).Value Then     'std
                        illoop = 2
                    Else
                        illoop = 3                          'calendar
                    End If
                    gSetupBOBDates illoop, llStdStartDates(), llLastBilled, ilLastBilledInx, 0, ilSaveMonth 'build array of start & end dates
    
                    mSalesFormula       'send formula to crystal for sorting options
                    gCrCumeActCt llStdStartDates()
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_MAKEPLAN Then          'avg prices to make plan
                    Screen.MousePointer = vbHourglass
                    gCrMakePlan
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_VEHCPPCPM Then         'currentcpp or cpm by vehicle
                    Screen.MousePointer = vbHourglass
                    gCrVehCPPCPM
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_SALESANALYSIS Then
                    Screen.MousePointer = vbHourglass
                    'Determine contracts to process based on their entered and modified dates
                    slStr = RptSelCt!CSI_CalFrom.Text           'Date: 12/12/2019 added CSI calendar control for date entrry --> edcSelCFrom.Text               'year
                    ilTemp = Val(slStr)
                    slStr = RptSelCt!edcSelCFrom1.Text             'month in text form (jan..dec)
                    gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
                    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                        ilSaveMonth = Val(slStr)
                    End If
                    ilSaveMonth = (ilSaveMonth - 1) * 3 + 1         'obtain starting month from the starting quarter
                    ilRet = mObtainStartEndDates(ilTemp, ilSaveMonth, 3, llEarliestDate, llLatestDate)
    
                    'obtain previous years dates
                    ilTemp = ilTemp - 1                     'previous year
                    'month (ilSaveMonth) is the same
                    ilRet = mObtainStartEndDates(ilTemp, ilSaveMonth, 3, llCompareDate, llCompareDate2)
    
                    gCrSalesAna llEarliestDate, llLatestDate, llCompareDate, llCompareDate2
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_PAPERWORK Then     '7-18-01
                    Screen.MousePointer = vbHourglass
                    gCrPaperWork
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_SALESPLACEMENT Then        '8-08-02
                    Screen.MousePointer = vbHourglass
                    slStr = RptSelCt!edcSelCTo1.Text             'month in text form (jan..dec)
                    gGetMonthNoFromString slStr, ilSaveMonth         'getmonth #
    
                    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                        ilSaveMonth = Val(slStr)
                    End If
    
                    If rbcSelCInclude(0).Value Then         'corp
                        illoop = 1
                    ElseIf rbcSelCInclude(1).Value Then     'std
                        illoop = 2
                    Else
                        illoop = 3                          'calendar
                    End If
                    gSetupBOBDates illoop, llStdStartDates(), llLastBilled, ilLastBilledInx, 0, ilSaveMonth 'build array of start & end dates
    
                    If gOpenBOBFilesCt() = 0 Then
                        'Determine calendar month requested, and retrieve all History and Receivables
                        'records that fall within the beginning of the cal year and end of calendar month requested
                        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
                        llLastBilled = gDateValue(slStr)            'convert last month billed to long
                        ilLastBilledInx = 12
                        For illoop = 1 To 12 Step 1
                            If llLastBilled > llStdStartDates(illoop) And llLastBilled < llStdStartDates(illoop + 1) Then
                                ilLastBilledInx = illoop
                                Exit For
                            End If
                         Next illoop
    
                        mSalesFormula       'send formula to crystal for sorting options
    
                        gGetAllParticipantSplits llStdStartDates(1)
                        '7-12-01 Use airing lines when invoicing by Line so past & future will still balance to contract
                        If llStdStartDates(1) > llLastBilled Or (tgSpf.iPkageGenMeth = 1 And RptSelCt!rbcSelCSelect(1).Value) Then    'projection only if dates all in future or Calc virtual pkg by line & using airng lines
                            'ilRet = mBuildSlsPlacement(llStdStartDates(), llLastBilled, 12, 0)
                            ilRet = mBuildSlsPlacement(llStdStartDates(), llLastBilled, igPeriods, 0)           '7-7-14  use # periods entered vs always 12
                        Else
                            gCRBOB_PastCT llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, 0, tmBillCycle       '7-7-14  use # periods entered vs always 12
                            If llLastBilled + 1 < llStdStartDates(igPeriods + 1) Then              'past only or past & projection
                                ilRet = mBuildSlsPlacement(llStdStartDates(), llLastBilled, igPeriods, 0)       '7-7-14  use # periods entered vs always 12
                            End If
                        End If
    
                        gCloseBOBFilesCt
                        Erase llStdStartDates
                    End If
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_VEH_UNITCOUNT Then
                    Screen.MousePointer = vbHourglass
                    gCrVehUnitCount
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_LOCKED Then            '4-5-06
                    Screen.MousePointer = vbHourglass
                    gCrLockedAvails
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_PAPERWORKTAX Then          '4-9-07
                    Screen.MousePointer = vbHourglass
                    gCrPaperWkTax
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_MG Then
                    Screen.MousePointer = vbHourglass
                    gCreateMGRpt               ''*CCCCC
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_CONTRACTVERIFY Then        '4--13
                    Screen.MousePointer = vbHourglass
                    gCreateContractVerify
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_INSERTION_ACTIVITY Then    '10-6-15
                    Screen.MousePointer = vbHourglass
                    gCreateInsertionActivity
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_XML_ACTIVITY Then    '3-25-16
                    Screen.MousePointer = vbHourglass
                    gCreateXMLActivity
                    Screen.MousePointer = vbDefault
                End If
            End If

            'tests to prevent a blank page from printing
            ilFilterBRPass = False
            If (ilListIndex = CNT_BR And igRptCallType = CONTRACTSJOB) Then
                If ilDoIt Then                                      '2-7-17 outside of ilJobs loop; determine if more reports of contracts should be processed
                                                                    'if not, ignore
                    If ilJobs = 1 Then      'pass 1 : always detail
                        If igBR_SchLinesExist = True Then       'print the detail if at least 1 sche lines exists in one or more contracts
                            ilFilterBRPass = True
                        End If
                    ElseIf ilJobs = 2 Then                               'pass 2: always NTR
                        If igBR_NTRDefined = True Then                  'could be multiple contracts, any NTR defined at all?
                            ilFilterBRPass = True
                        End If
                    ElseIf ilJobs = 3 Then                              '1-7-21 cpm podcast
                        If igBR_CPMDefined = True Then
                            ilFilterBRPass = True
                        End If
                    Else
                        'Research summary or bill summary
                        'if combining air time and NTR and at least one NTR defined, print it
                        'if at least one sch line exists, print it
                        If igBRSumZer Then          'billing summary, ok to show this if no schedule lines exist
                            'TTP 10884 - Proposals/Contracts report: "rates" only option or "rates and hidden" option doesn't include monthly/quarterly summary page for digital line contract
                            'if combining ntr and air times and the ntrs exists,or at least one air time exists, show the billing summary
                            'If (ckcSelC10(1).Value = vbChecked And igBR_NTRDefined = True) Or (igBR_SchLinesExist = True) Or (ckcSelC6(1).Value = vbChecked And igBR_CPMDefined = True) Then
                            If (ckcSelC10(1).Value = vbChecked And igBR_NTRDefined = True) Or (igBR_SchLinesExist = True) Or (ckcSelC10(1).Value = vbChecked And igBR_CPMDefined = True) Then
                                ilFilterBRPass = True
                            End If
                        Else           'research summary with rates; only do this version if at least one sch lines exists
                            'If (ckcSelC6(0).Value = vbChecked And ckcSelC6(1).Value = vbChecked And igBR_SchLinesExist = True) Then
                            '2-22-10 Need to get the research version, regardless of rates/no rates
                            If (ckcSelC6(1).Value = vbChecked And igBR_SchLinesExist = True) Or (ckcSelC6(1).Value = vbChecked And igBR_CPMDefined = True) Then
                                ilFilterBRPass = True
                            End If
                        End If
                    End If
                    'if passed all the contract parameters with NTR and airtime, should this pass be printed
                    If (ilFilterBRPass = False) Or (ilDoIt = False) Then
                        ilFilterBRPass = False
                    End If

                    If (Not ilFilterBRPass) And ilDoIt = True Then
                        ogReport.RemoveReport
                    End If
                    'dan 10-31-08 for multiple report jobs: don't call form until last job.
                    If Not ogReport Is Nothing Then
                    'dan 6-09-09 took away multi reports 7/09 returned multi reports
                        If ilJobs = 0 Or (ilJobs >= ogReport.iLastPrintJob) Then
                            If rbcOutput(0).Value Then
                                DoEvents            '9-13-02 fix for timing problem with Avails report & Spot Business Booked (random problem)
                                igDestination = 0
                                Screen.MousePointer = vbDefault
                               'If ogReport.ReportCount > 0 Then
                                    Report.Show vbModal
                               ' End If
                            ElseIf rbcOutput(1).Value Then
                                ilCopies = Val(edcCopies.Text)
                                ilRet = gOutputToPrinter(ilCopies)
                            Else
                                slFileName = edcFileName.Text
                                ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
                            End If
                       End If
                    End If
                End If
            'End If
            Else
                If ilListIndex = CNT_INSERTION And rbcOutput(3).Value = True Then
                    'check if email option
                    mCreateInsertionEmailPdfs
                Else
                    'dan 10-31-08 for multiple report jobs: don't call form until last job.
                    If Not ogReport Is Nothing Then
                    'dan 6-09-09 took away multi reports 7/09 returned multi reports
                        If ilJobs = 0 Or (ilJobs >= ogReport.iLastPrintJob) Then
                            If rbcOutput(0).Value Then
                                DoEvents            '9-13-02 fix for timing problem with Avails report & Spot Business Booked (random problem)
                                igDestination = 0
                               'If ogReport.ReportCount > 0 Then
                                    Report.Show vbModal
                               ' End If
                            ElseIf rbcOutput(1).Value Then
                                ilCopies = Val(edcCopies.Text)
                                ilRet = gOutputToPrinter(ilCopies)
                            ElseIf rbcOutput(2).Value Then
                                'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
                                If ckcSeparateFile.Value = 1 Then
                                    'Separate Export for each contract
                                    edcFileName.Text = mGetSeparateFilename("IO", lbcSelection(6).List(tmVehicleList(ilExportSplit)))
                                    slFileName = edcFileName.Text
                                    cmcGen.Enabled = False
                                    If bmHasRecords = True Then
                                        ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
                                    End If
                                Else
                                    slFileName = edcFileName.Text
                                    ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
                                End If
                            ElseIf rbcOutput(3).Value Then 'email
                                'setup filename for email
                                mCreateOrderEmailPDF
                            ElseIf rbcOutput(4).Value Then 'TTP 10119 - Average 30 Rate Report - add option to export to CSV
                                '
                            End If
                       End If
                    End If 'ogreport exists
                End If
            End If
        Next ilJobs
        Set ogReport = Nothing
        imGenShiftKey = 0
        If (igRptCallType = SLSPCOMMSJOB) Then
           If (ilListIndex = COMM_SALESCOMM) Then   'Or ilListIndex = COMM_PROJECTION) Then
                Screen.MousePointer = vbHourglass
                gCRGrfClear
                Screen.MousePointer = vbDefault
            ElseIf (ilListIndex = COMM_PROJECTION) Then
                Screen.MousePointer = vbHourglass
                gCRGrfClear
                Screen.MousePointer = vbDefault
            End If
        ElseIf (igRptCallType = CONTRACTSJOB) Then
            If (ilListIndex = CNT_QTRLY_AVAILS) Then     'Quarterly Avails
                Screen.MousePointer = vbHourglass
                gCRQAvailsClearCt
                Screen.MousePointer = vbDefault
            End If
            If (ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION) Then   '12-10-03 And Not (rbcSelCInclude(2).Value) Then     'produce wide proposal, contract
                'for NTR generation only, may or maynot be present
                Screen.MousePointer = vbHourglass
                'clear printables flag if necessary
                're-open contract file
                If RptSelCt!rbcOutput(1).Value And RptSelCt!ckcSelC3(0).Value = vbChecked Then                'print output and printables only
                    hmCHF = CBtrvTable(TWOHANDLES) 'CBtrvObj()
                    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilRet = btrClose(hmCHF)
                        btrDestroy hmCHF
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    imCHFRecLen = Len(tgChfCT)
                    For illoop = 0 To UBound(lgPrintedCnts) Step 1
                        Do
                            If lgPrintedCnts(illoop) <> 0 Then
                                tmChfSrchKey.lCode = lgPrintedCnts(illoop)
                                imCHFRecLen = Len(tgChfCT)
                                ilRet = btrGetEqual(hmCHF, tgChfCT, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                If ilRet = BTRV_ERR_NONE Then
                                    tgChfCT.sPrint = "P"
                                    ilRet = btrUpdate(hmCHF, tgChfCT, imCHFRecLen)
                                End If
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                    Next illoop
                    ilRet = btrClose(hmCHF)
                End If
                Erase lgPrintedCnts
                
                '2-16-13 #3  clear the temporary file with the user ID along with gendate and time
                gCBFClearWithUserID
                gClearTxr                                                   'clear txr for the split network station list
                Screen.MousePointer = vbDefault
            ElseIf (ilListIndex = CNT_BOB) Or (ilListIndex = CNT_SALESACTIVITY) Or (ilListIndex = CNT_SALESCOMPARE) Or (ilListIndex = CNT_CUMEACTIVITY) Or (ilListIndex = CNT_SALESACTIVITY_SS) Or (ilListIndex = CNT_SALESPLACEMENT) Or (ilListIndex = CNT_BOBRECAP) Or (ilListIndex = CNT_BOBCOMPARE) Then
                Screen.MousePointer = vbHourglass
                gCRGrfClear
                Screen.MousePointer = vbDefault
    '         ElseIf (ilListIndex = CNT_SALESANALYSIS) Or (ilListIndex = CNT_AVGRATE) Or (ilListIndex = CNT_AVG_PRICES) Or (ilListIndex = CNT_ADVT_UNITS) Or (ilListIndex = CNT_TIEOUT) Or (ilListIndex = CNT_SPTSBYADVT) Then
    '10-29-10 remove spotsbyadvt report, its in rptcrcb
             ElseIf (ilListIndex = CNT_SALESANALYSIS) Or (ilListIndex = CNT_AVGRATE) Or (ilListIndex = CNT_AVG_PRICES) Or (ilListIndex = CNT_ADVT_UNITS) Or (ilListIndex = CNT_TIEOUT) Then
                Screen.MousePointer = vbHourglass
                gCRGrfClear
                Screen.MousePointer = vbDefault
            ElseIf ilListIndex = CNT_VEH_UNITCOUNT Or ilListIndex = CNT_LOCKED Then     '4-5-06
                Screen.MousePointer = vbHourglass
                gCRGrfClear
                Screen.MousePointer = vbDefault
            ElseIf (ilListIndex = CNT_AFFILTRAK) Or (ilListIndex = CNT_COMLCHG) Or (ilListIndex = CNT_MG) Or (ilListIndex) = CNT_SPOTTRAK Then
                Screen.MousePointer = vbHourglass
                gCRGrfClear
                Screen.MousePointer = vbDefault
            ElseIf (ilListIndex = CNT_DAILY_SALESACTIVITY) Or (ilListIndex = CNT_PAPERWORK) Or (ilListIndex = CNT_PAPERWORKTAX) Or (ilListIndex = CNT_CONTRACTVERIFY) Then
                Screen.MousePointer = vbHourglass
                gCRGrfClear
                Screen.MousePointer = vbDefault
            ElseIf (ilListIndex = CNT_MAKEPLAN) Or (ilListIndex = CNT_VEHCPPCPM) Then
                Screen.MousePointer = vbHourglass
                gCrAnrClear
                Screen.MousePointer = vbDefault
            ElseIf (ilListIndex = CNT_SALES_CPPCPM) Or (ilListIndex = CNT_HISTORY) Or (ilListIndex = CNT_BR And rbcSelCInclude(2).Value) Then
                Screen.MousePointer = vbHourglass
                gCrCbfClear
                Screen.MousePointer = vbDefault
            ElseIf (ilListIndex = CNT_BOB_BYSPOT) Then
                Screen.MousePointer = vbHourglass
                gJsrClearCt
                Screen.MousePointer = vbDefault
            ElseIf ilListIndex = CNT_INSERTION_ACTIVITY Then    '10-6-15
                Screen.MousePointer = vbHourglass
                gCRGrfClear
                gClearTxr
                Screen.MousePointer = vbDefault
           End If
        End If
    
    Next ilExportSplit
    'TTP 10271: ENDOF INDIVIDUAL CONTRACTS LOOP
    
    If ckcSeparateFile.Value = 1 Then
        'ReSelect user selected vehicles
        ilValue = False
        'Clear auto gen filename
        edcFileName.Text = "IO"
        If ilCkcAllVehState = 1 Then
            CkcAllveh.Value = 0
            CkcAllveh.Value = 1
        Else
            CkcAllveh.Value = ilCkcAllVehState
            llRg = CLng(lbcSelection(6).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(6).HWnd, LB_SELITEMRANGE, ilValue, llRg)
            For ilExportSplit = 0 To UBound(tmVehicleList) - 1
                lbcSelection(6).Selected(tmVehicleList(ilExportSplit)) = True
            Next ilExportSplit
        End If
        'Message complete
        MsgBox "Separate files export complete.", vbOKOnly, "Export Complete"
    End If
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    'frcWhen.Enabled = igWhen
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    'frcRptType.Enabled = igReportType
    'Select Case igRptCallType
    '    Case CONTRACTSJOB
    '        mSetCommands
    'End Select
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    cmcGen.Enabled = True
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
    'cdcSetup.Flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub

Private Sub CSI_CalFrom_Change()
    Dim ilListIndex As Integer
    Dim ilLen As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
    Case CONTRACTSJOB
        If (igRptType = 0) And (ilListIndex > 1) Then
            ilListIndex = ilListIndex + 1
        End If
        'Date: 12/12/2019 added CSI calendar control for date entry
        If ilListIndex = CNT_VEHCPPCPM Then
            ilLen = Len(CSI_CalFrom.Text)
            If ilLen >= 4 Then
                slDate = CSI_CalFrom.Text
                llDate = gDateValue(slDate)

                'populate Rate Cards and bring in Rcf, Rif, and Rdf
                ilRet = gPopRateCardBox(RptSelCt, llDate, RptSelCt!lbcSelection(4), tgRateCardCode(), smRateCardTag, -1)
                lbcSelection(4).Visible = True
            End If
        End If
    End Select
    mSetCommands
End Sub

Private Sub CSI_CalTo_Change()
    mSetCommands
End Sub

Private Sub CSI_From1_Change()
    Dim ilListIndex As Integer
    Dim ilLen As Integer
    Dim slDate As String
    Dim llDate As Long
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            If ilListIndex = CNT_SALESANALYSIS Then
                ilLen = Len(edcSelCTo)
                If ilLen = 4 Then
                    lbcSelection(4).Clear
                    slDate = "1/15/" & Trim$(edcSelCTo)
                    slDate = gObtainStartStd(slDate)
                    llDate = gDateValue(slDate)
                    mBudgetPop
                    lbcSelection(4).Move 15, ckcAll.Top + ckcAll.Height + 30, 4380, 3270
                    lbcSelection(4).Visible = True
                    laclbcName(0).Visible = True
                    laclbcName(0).Caption = "Budget Names"
                    laclbcName(0).Move ckcAll.Left, ckcAll.Top + 30, 1725
                    laclbcName(1).Visible = False
                End If
            End If
    End Select
    mSetCommands
End Sub

Private Sub CSI_To1_Change()
    mSetCommands
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

Private Sub edcResponse_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcResponse_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub

Private Sub edcSelCFrom_Change()
    Dim ilListIndex As Integer
    Dim ilLen As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If

            '10-15-04 change the gathering of rate cards based on the entered date & gather all
            'rate cards from that date on
            If ilListIndex = CNT_QTRLY_AVAILS Then
                ilLen = Len(edcSelCFrom)
                If ilLen >= 4 Then
                    slDate = edcSelCFrom           'retrieve jan thru dec year
                    slDate = gObtainStartStd(slDate)
                    llDate = gDateValue(slDate)
                    'populate Rate Cards and bring in Rcf, Rif, and Rdf
                    ilRet = gPopRateCardBox(RptSelCt, llDate, lbcSelection(12), tgRateCardCode(), sgRateCardCodeTag, -1)
                End If
            ElseIf ilListIndex = CNT_MAKEPLAN Then
                ilLen = Len(edcSelCFrom)
                If ilLen = 4 Then
                    mBudgetPop
                    'if using Corporate calendar, need to show all rate cards applicable, and they must pick 2 years
                        If rbcSelCSelect(0).Value Then          'corp, need to adjust the date for the rate cards to bring in
                            'if asking for corp year 1997, need to bring in 1996 and 1997 since rates cards are input as jan-dec
                            ilLen = Val(edcSelCFrom) - 1
                            ilRet = gGetCorpCalIndex(ilLen)
                            'If ilRet < 1 Then
                            If ilRet < 0 Then
                                slDate = "1/15/" & Trim$(edcSelCFrom)           'retrieve jan thru dec year
                            Else                                                'no error
                                slDate = Trim$(str$(tgMCof(ilRet).iStartMnthNo)) & "/15/" & Trim$(str$(Val((edcSelCFrom) - 1)))      'retrieve corp year
                            End If
                            llDate = gDateValue(slDate)
                            'populate Rate Cards and bring in Rcf, Rif, and Rdf
                            ilRet = gPopRateCardBox(RptSelCt, llDate, RptSelCt!lbcSelection(11), tgRateCardCode(), smRateCardTag, -1)
                            lbcSelection(11).Visible = True
                        Else
                            slDate = "1/15/" & Trim$(edcSelCFrom)           'retrieve jan thru dec year
                            slDate = gObtainStartStd(slDate)
                            llDate = gDateValue(slDate)

                            'populate Rate Cards and bring in Rcf, Rif, and Rdf
                            ilRet = gPopRateCardBox(RptSelCt, llDate, RptSelCt!lbcSelection(12), tgRateCardCode(), smRateCardTag, -1)
                            lbcSelection(12).Visible = True
                    End If
                End If
            Else
                If ilListIndex = CNT_VEHCPPCPM Then
                    ilLen = Len(CSI_CalFrom.Text)           'Date: 12/12/2019 added CSI calendar control for date entry --> edcSelCFrom)
                    If ilLen >= 4 Then
                        slDate = CSI_CalFrom.Text           'Date: 12/12/2019 added CSI calendar control for date entry --> edcSelCFrom.Text          'retrieve jan thru dec year
                        llDate = gDateValue(slDate)

                        'populate Rate Cards and bring in Rcf, Rif, and Rdf
                        ilRet = gPopRateCardBox(RptSelCt, llDate, RptSelCt!lbcSelection(4), tgRateCardCode(), smRateCardTag, -1)
                        lbcSelection(4).Visible = True
                    End If
                End If

            End If
    End Select
    mSetCommands
End Sub

Private Sub edcSelCFrom_GotFocus()
    gCtrlGotFocus edcSelCFrom
End Sub

Private Sub edcSelCFrom_KeyPress(KeyAscii As Integer)
    Dim ilListIndex As Integer
    If igRptCallType = COPYJOB Then
        ilListIndex = lbcRptType.ListIndex
        If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then
            ilListIndex = ilListIndex + 1
        End If
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
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            If ilListIndex = CNT_QTRLY_AVAILS Then    'Quarterly Avail
                'Filter characters (allow only BackSpace, numbers 0 thru 9
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
    End Select
End Sub

Private Sub edcSelCTo_Change()
    Dim ilListIndex As Integer
    Dim ilLen As Integer
    Dim slDate As String
    Dim llDate As Long
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            If ilListIndex = CNT_SALESANALYSIS Then
                ilLen = Len(edcSelCTo)
                If ilLen = 4 Then
                    lbcSelection(4).Clear
                    slDate = "1/15/" & Trim$(edcSelCTo)
                    slDate = gObtainStartStd(slDate)
                    llDate = gDateValue(slDate)
                    mBudgetPop
                    lbcSelection(4).Move 15, ckcAll.Top + ckcAll.Height + 30, 4380, 3270
                    lbcSelection(4).Visible = True
                    laclbcName(0).Visible = True
                    laclbcName(0).Caption = "Budget Names"
                    laclbcName(0).Move ckcAll.Left, ckcAll.Top + 30, 1725
                    laclbcName(1).Visible = False
                End If
            End If
    End Select
    mSetCommands
End Sub

Private Sub edcSelCTo_GotFocus()
    gCtrlGotFocus edcSelCTo
End Sub

Private Sub edcSelCTo_KeyPress(KeyAscii As Integer)
    Dim ilListIndex As Integer
    If igRptCallType = COPYJOB Then
        ilListIndex = lbcRptType.ListIndex
        If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then
            ilListIndex = ilListIndex + 1
        End If
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

Private Sub edcText_Change()
    mSetCommands
End Sub

Private Sub edcText_GotFocus()
    gCtrlGotFocus edcText
End Sub

Private Sub edcTopHowMany_Change()
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = CONTRACTSJOB Then    '11-27-00
        If ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION Then
            mSetCommands
        ElseIf ilListIndex = CNT_HISTORY Then       '1-31-01
            mSetCommands
        End If
    End If
    'Separate files option applies to 1 Contract Selected & Multiple Vehicles selected Only
    ckcSeparateFile.Enabled = False
    If igRptCallType = CONTRACTSJOB And ilListIndex = CNT_INSERTION Then
        mEnableSeparateFiles
    End If
End Sub

Private Sub edcTopHowMany_GotFocus()
    gCtrlGotFocus ActiveControl
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
    RptSelCt.Refresh
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lacExport.ForeColor = vbBlack
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
    'RptSelCt.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmEMailContentCode
    Erase tgCSVNameCode
    'Erase tgSellNameCode
    Erase tgRptSelSalespersonCodeCt
    Erase tgRptSelAgencyCodeCt
    Erase tgRptSelAdvertiserCodeCt
    'Erase tgRptSelNameCode
    Erase tgRptSelBudgetCodeCT
    Erase tgMultiCntrCodeCT
    Erase tgManyCntCodeCT
    Erase tgRptSelDemoCodeCT
    Erase tgSOCodeCT
    Erase tgMnfCodeCT
    Erase tgMNFCodeRpt
    Erase tgVehicleSets1        '10-11-07
    Erase tgVehicleSets2        '10-11-07
    Erase lgPrintedCnts
    Erase tgClfCT
    Erase tgCffCT
    Erase imCodes
    Erase tgEmail_PDFs, tmEmail_Info
'    Erase tgAcqComm
'    Erase tgAcqCommInx
    PECloseEngine
    gChDrDir           '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path

    Set RptSelCt = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lacExport_Click()
    If lacExport.Caption <> "" Then
        'Show exported file in explorer
        Shell "explorer.exe /select, " & Mid(lacExport.Caption, 19), vbNormalFocus
    End If
End Sub

Private Sub lacExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And Y > 0 Then
        lacExport.ForeColor = vbBlue
    End If
End Sub

Private Sub lbcRptType_Click()
    Dim slStr As String
    Dim ilListIndex As Integer
    Dim slMonth As String
    Dim slYear As String
    Dim slDay As String
    ReDim ilAASCodes(0 To 1) As Integer
    Dim slAirOrder As String                'from site pref, bill as ordered (update order),
                                            'update as ordered (update aired), bill as aired
    Dim ilTop As Integer
    
    rbcSelCInclude(2).Visible = False
    Select Case igRptCallType
        Case SLSPCOMMSJOB
            lacSelCTo.Visible = False
            edcSelCTo.Visible = False
            lacSelCFrom1.Visible = False
            edcSelCFrom1.Visible = False
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            plcSelC3.Visible = False
            plcSelC4.Visible = False
            plcSelC5.Visible = False
            plcSelC6.Visible = False
            plcSelC7.Visible = False
            plcSelC8.Visible = False
            ilListIndex = lbcRptType.ListIndex
            Select Case lbcRptType.ListIndex
                Case COMM_SALESCOMM, COMM_PROJECTION                 'sales commission or projection
                    pbcSelC.Visible = True
                    ckcAll.Left = 120
                    ckcAll.Caption = "All Salespeople"
                    lbcSelection(2).Move 120, ckcAll.Height + 30, 4260, 3270     'slsp list box
                    lbcSelection(2).Move 120, lbcSelection(0).Top, 4260, lbcSelection(0).Height
                    lbcSelection(2).Visible = True
                    If ilListIndex = COMM_SALESCOMM Then
                        mSellConvVVPkgPop 6, False                    'lbcselection(6), vehicles
                        lbcSelection(2).Height = 1500
                        ckcAllAAS.Move 120, lbcSelection(2).Top + lbcSelection(2).Height + 30
                        ckcAllAAS.Caption = "All Vehicles"
                        ckcAllAAS.Visible = True
                        lbcSelection(6).Move 120, ckcAllAAS.Top + ckcAllAAS.Height + 30, 4260, 1500
                        lbcSelection(6).Visible = True
                        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr
                        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                        slMonth = gMonthName(slStr)
                        edcSelCFrom.MaxLength = 3           'Jan, ....dec
                        edcSelCFrom.Text = Trim$(slMonth)
                        edcSelCFrom1.MaxLength = 4          'year 1996..2000...
                        edcSelCFrom1.Text = Trim$(slYear)
                        lacSelCFrom.Caption = "Select Month"
                        lacSelCFrom.Visible = True
                        lacSelCFrom.Left = 120
                        lacSelCFrom.Visible = True
                        edcSelCFrom.Visible = True

                        lacSelCFrom1.Caption = "Year"
                        lacSelCFrom1.Visible = True
                        lacSelCFrom1.Left = 1920
                        edcSelCFrom1.Visible = True
                        edcSelCFrom.Move 1320, edcSelCFrom.Top, 480
                        edcSelCFrom1.Move 2400, edcSelCFrom.Top, 600

                        '11-30-04 allow report to be requested for more than 1 month
                        edcSelCTo.Move 4080, edcSelCFrom.Top, 360
                        edcSelCTo.Text = "1"
                        edcSelCTo.MaxLength = 2
                        lacSelCTo.Caption = "# Periods"
                        lacSelCTo.Move 3180, edcSelCFrom.Top + 60, 960
                        lacSelCTo.Visible = True
                        edcSelCTo.Visible = True
                        
                        '8-19-09 Allow std to be run if corporate calendar exists; otherwise only std allowed
                        If tgSpf.sRUseCorpCal = "Y" Then
                            plcSelC4.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30, 4000
                            smPaintCaption4 = "YTD Calendar"
                            plcSelC4_Paint
                            rbcSelC4(0).Caption = "Corporate"
                            rbcSelC4(0).Value = True
                            rbcSelC4(1).Caption = "Standard"
                            rbcSelC4(0).Visible = True
                            rbcSelC4(1).Visible = True
                            rbcSelC4(0).Move 1200, 0, 1200
                            rbcSelC4(1).Move 2400, 0, 1200
                            plcSelC4.Visible = True
                            ilTop = plcSelC4.Top + plcSelC4.Height
                        Else
                            rbcSelC4(1).Value = True        'default to standard
                            'dont show the question since corporate is disallowed if no calendar defined
                            ilTop = edcSelCTo.Top + edcSelCTo.Height
                        End If

                        '9-2-08 new commission report option
                        plcSelC3.Move 120, ilTop, 4000
                        smPaintCaption3 = ""
                        plcSelC3_Paint
                        ckcSelC3(0).Caption = "Add Bonus Comm for New/Increased Sales"
                        ckcSelC3(0).Move 0, 0, 4000
                        plcSelC3.Visible = True
                        ckcSelC3(0).Visible = True

                        '4-20-00
                        lacSelCTo1.Visible = True
                        lacSelCTo1.Move 120, plcSelC3.Top + plcSelC3.Height + 60, 840
                        lacSelCTo1.Caption = "Contract #"
                        edcSelCTo1.MaxLength = 9
                        edcSelCTo1.Move 1320, plcSelC3.Top + plcSelC3.Height + 30, 960
                        edcSelCTo1.Visible = True

                        plcSelC1.Visible = True
                        smPaintCaption1 = "Sort by- "
                        plcSelC1_Paint
                        plcSelC1.Move 120, edcSelCTo1.Top + edcSelCTo1.Height + 30
                        rbcSelCSelect(0).Caption = "Slsp"
                        rbcSelCSelect(1).Caption = "Vehicle, slsp"
                        rbcSelCSelect(0).Visible = True
                        rbcSelCSelect(1).Visible = True
                        If rbcSelCSelect(0).Value Then             'default to SLSP
                            rbcSelCSelect_click 0
                        Else
                            rbcSelCSelect(0).Value = True
                        End If
                        rbcSelCSelect(0).Move 840, 0, 720
                        rbcSelCSelect(1).Move 1560, 0, 1680

                        plcSelC7.Visible = True
                        smPaintCaption7 = "By"
                        plcSelC7_Paint
                        'plcSelC7.Move 120, edcSelCTo1.Top + edcSelCTo.Height + 30
                        plcSelC7.Move 120, plcSelC1.Top + plcSelC1.Height
                        rbcSelC7(0).Value = True
                        rbcSelC7(0).Visible = True
                        rbcSelC7(0).Caption = "Detail"
                        rbcSelC7(0).Left = 360
                        rbcSelC7(0).Width = 800
                        rbcSelC7(1).Caption = "Summary"
                        rbcSelC7(1).Visible = True
                        rbcSelC7(1).Left = 1200
                        rbcSelC7(1).Width = 1200
                        rbcSelC7(2).Visible = False

                        'ifusing bonus commissions for new and increased sales, default to call new report
                        If (Asc(tgSpf.sUsingFeatures7) And BONUSCOMM) = BONUSCOMM Then
                            ckcSelC3(0).Value = vbChecked
                        Else
                            ckcSelC3(0).Value = vbUnchecked
                        End If

                        '2-13-02 subtotals by contract
                        plcSelC8.Move 120, plcSelC7.Top + plcSelC7.Height
                        ckcSelC8(0).Caption = "Include Sub-totals by Contract"
                        ckcSelC8(0).Move 0, 0, 3600
                        ckcSelC8(0).Value = vbUnchecked
                        ckcSelC8(0).Visible = True
                        plcSelC8.Visible = True

                        '5-17-04 Air Time, NTR or Both
                        rbcSelC9(0).Caption = "Air Time"
                        rbcSelC9(0).Move 480, 0, 1080
                        rbcSelC9(2).Value = True
                        rbcSelC9(1).Caption = "NTR"
                        rbcSelC9(1).Move 1560, 0, 700
                        rbcSelC9(2).Caption = "Both"
                        rbcSelC9(2).Move 2260, 0, 720
                        smPaintCaption9 = "For"
                        plcSelC9.Move 120, plcSelC8.Top + plcSelC8.Height + 30, 2980
                        rbcSelC9(0).Visible = True
                        rbcSelC9(1).Visible = True
                        rbcSelC9(2).Visible = True
                        plcSelC9.Visible = True

                        '4-12-07 option to include hard cost
                        plcSelC10.Move 3220, plcSelC9.Top - 30, 1200
                        ckcSelC10(0).Move 0, 0, 1200
                        ckcSelC10(0).Caption = "Hard Cost"
                        ckcSelC10(0).Visible = True
                        plcSelC10.Visible = True
                        '5-11-04 Add option to sort by % within slsp, or vehicle group within slsp
                        'Currently sort is advt/cnt within slsp
                        smPaintCaption11 = "Sort within Slsp by-"
                        plcSelC11.Move 120, plcSelC9.Top + plcSelC9.Height
                        plcSelC11.Height = 480
                        rbcSelC11(0).Caption = "Advertiser"
                        rbcSelC11(1).Caption = "Percent"
                        rbcSelC11(2).Caption = "Owner"

                        'rbcSelC11(0).Move 1790, 0, 1200
                        'rbcSelC11(1).Move 3120, 0, 960
                        'rbcSelC11(2).Move 1790, 240, 1560
                        rbcSelC11(0).Move 360, 240, 1200
                        rbcSelC11(1).Move 1560, 240, 960
                        rbcSelC11(2).Move 2520, 240, 1560
                        rbcSelC11(0).Value = True
                        rbcSelC11(0).Visible = True
                        rbcSelC11(1).Visible = True
                        rbcSelC11(2).Visible = True
                        plcSelC11.Visible = True
                        plcSelC11_Paint
                       ' gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True
                        'edcSet1.Text = "Vehicle Group"
                        'cbcSet1.ListIndex = 0
                        'edcSet1.Move 120, plcSelC11.Top + 30 + plcSelC11.Height
                        'cbcSet1.Move 360 + edcSet1.Width, edcSet1.Top - 45
                        'edcSet1.Visible = True
                        'cbcSet1.Visible = True

                        '7-16-08  Incl acquisition costs
                        plcSelC12.Move plcSelC11.Left, plcSelC11.Top + plcSelC11.Height
                        ckcSelC12(0).Move 0, 0, 2910
                        ckcSelC12(0).Caption = "Include  Acquisition Costs"
                        plcSelC12.Visible = True
                        ckcSelC12(0).Visible = True
                        ckcSelC12(0).Value = vbChecked
                        
                        '4-14-15
                        plcSelC13.Move plcSelC12.Left, plcSelC12.Top + plcSelC12.Height + 60
                        ckcSelC13(0).Move 840, 0, 840
                        ckcSelC13(0).Caption = "Polit"
                        ckcSelC13(0).Visible = True
                        ckcSelC13(0).Value = vbChecked
                        ckcSelC13(1).Move 1680, 0, 1200
                        ckcSelC13(1).Caption = "Non-Polit"
                        ckcSelC13(1).Visible = True
                        ckcSelC13(1).Value = vbChecked
                        plcSelC13.Visible = True
                        smPaintCaption13 = "Include"
                    Else                                         'projection
                        slAirOrder = tgSpf.sInvAirOrder
                        'This report is using the same pre-pass as the Salesperson Billed & Booked
                        'force some answers since all not asked
                        'lacSelCFrom.Caption = "Year"
                        'lacSelCFrom.Visible = True
                        'lacSelCFrom.left = 120
                        'lacSelCFrom.Visible = True
                        'edcSelCFrom.Visible = True
                        'edcSelCFrom.MaxLength = 4           'i.e.-1990....2015
                        'edcSelCFrom.left = 600
                        'edcSelCFrom.Width = 600
                        lacSelCFrom.Caption = "Start Quarter"
                        lacSelCFrom.Visible = True
                        lacSelCFrom.Left = 120
                        edcSelCFrom.MaxLength = 1
                        edcSelCFrom.Width = 480
                        edcSelCFrom.Left = 1500
                        edcSelCFrom.Visible = True
                        lacSelCFrom1.Caption = "Year"
                        lacSelCFrom1.Visible = True
                        lacSelCFrom1.Left = 2100
                        edcSelCFrom1.MaxLength = 4
                        edcSelCFrom1.Move 2580, edcSelCFrom.Top, 600
                        edcSelCFrom1.Visible = True

                        '4-20-00
                        lacSelCTo.Visible = True
                        lacSelCTo.Move 120, lacSelCFrom1.Top + lacSelCFrom1.Height + 120, 840
                        lacSelCTo.Caption = "Contract #"
                        edcSelCTo.MaxLength = 9         '1-30-06
                        edcSelCTo.Move 1320, edcSelCFrom.Top + edcSelCFrom.Height + 30, 960
                        edcSelCTo.Visible = True


                        If rbcSelCInclude(1).Value Then             'default to SLSP
                            rbcSelCInclude_Click 1
                        Else
                            rbcSelCInclude(1).Value = True
                        End If

                        'remove detail version of billed & Booked commissions 3/22/99
                        'plcSelC4.Caption = "Totals by"
                        'plcSelC4.Move 120, lacSelCFrom.Top + lacSelCFrom.Height + 60
                        'rbcSelC4(0).Caption = "Contract"
                        'rbcSelC4(0).Left = 900
                        'rbcSelC4(0).Width = 960
                        'rbcSelC4(0).Visible = True
                        'rbcSelC4(0).Value = True
                        'If rbcSelC4(1).Value Then             'default to advt
                        '    rbcSelC4_click 1, True
                        'Else
                        '    rbcSelC4(1).Value = True
                        'End If
                        'rbcSelC4(1).Caption = "Advertiser"
                        'rbcSelC4(1).Left = 1920
                        'rbcSelC4(1).Width = 1200
                        'rbcSelC4(1).Visible = True
                        'rbcSelC4(2).Caption = "Summary"
                        'rbcSelC4(2).Left = 3120
                        'rbcSelC4(2).Width = 1200
                        'rbcSelC4(2).Visible = True
                        'plcSelC4.Visible = True

                        'default to net, since slsp commissions are based on net
                        rbcSelC7(1).Value = True

                        '4-20-00 plcSelC3.Move 120, lacSelCFrom.Top + lacSelCFrom.Height + 60
                        plcSelC3.Move 120, lacSelCTo.Top + lacSelCTo.Height + 60
                        mAskContractTypes

                        plcSelC1.Move 120, plcSelC6.Top + plcSelC6.Height
                        mAskPkgOrHide ilListIndex

                        'Use same control (different index) for unrelated question (due to lack of controls)
                        ckcSelC8(2).Visible = True
                        ckcSelC8(2).Caption = "Skip to new page each new salesperson"
                        ckcSelC8(2).Move 0, 420, 3720
                        ckcSelC8(2).Visible = True

                        ckcSelC8(0).Value = vbUnchecked 'False
                        ckcSelC8(1).Value = vbChecked   'True
                        If slAirOrder = "S" Then            'bill as ordered, update as ordered; don't ask adjustment qustions
                                                            'always ignore missed & count mgs
                            ckcSelC8(0).Visible = False
                            ckcSelC8(1).Visible = False
                        Else                                'as aired
                            ckcSelC8(0).Visible = True  '9-12-02 vbChecked 'True
                            ckcSelC8(1).Visible = True  '9-12-02 vbChecked 'True
                        End If
                        pbcSelC.Visible = True
                    End If
                End Select
        Case CONTRACTSJOB
            mCntSelectivity0            '3-18-03 split mcntselectivity1 into another module (too large)
            mCntSelectivity1
            mCntSelectivity2
            mCntSelectivity4            '3-28-19
    End Select
    mSetCommands
End Sub

'
'           5-7-98 If selective advertiser and only one contract exists,
'           force "All contracts" selected on
'
Private Sub lbcSelection_Click(Index As Integer)
    Dim ilListIndex As Integer
    Dim illoop As Integer
    Dim ilUpper As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilHOState As Integer
    ReDim ilAASCodes(0 To 1) As Integer
    Dim slCntrStatus As String
    Dim ilHowManyDefined As Integer

    If Not imAllClicked Then
        ilListIndex = lbcRptType.ListIndex
        If igRptCallType = CONTRACTSJOB Then
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            
            Select Case ilListIndex
                Case CNT_BR, CNT_INSERTION, CNT_HISTORY
                    
                    If Index > 6 And Index < 10 Then        'selected an advt, agy or slsp
                        ilUpper = 0
                        Screen.MousePointer = vbHourglass
                        For illoop = 0 To lbcSelection(Index).ListCount - 1
                            If lbcSelection(Index).Selected(illoop) Then
                                If rbcSelCSelect(0).Value Then          'advt
                                    slNameCode = tgRptSelAdvertiserCodeCt(illoop).sKey 'lbcAdvertiserCode.List(ilLoop)
                                ElseIf rbcSelCSelect(1).Value Then      'agy
                                    'slNameCode = tgRptSelAgencyCodeCt(ilLoop).sKey 'lbcAgencyCode.List(ilLoop) '5-16-02
                                    slNameCode = tgAgency(illoop).sKey 'lbcAgencyCode.List(ilLoop)  '5-16-2
                                Else                                    'slsp
                                    'slNameCode = tgRptSelSalespersonCodeCt(ilLoop).sKey    'lbcSalespersonCode.List(ilLoop)
                                    slNameCode = tgSalesperson(illoop).sKey    '5-16-02
                                End If
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                ilAASCodes(ilUpper) = Val(slCode)
                                ilUpper = ilUpper + 1
                                ReDim Preserve ilAASCodes(0 To ilUpper) As Integer
                            End If
                        Next illoop
                        If rbcSelCSelect(0).Value Then
                            'gPopCntrBoxViaArray 0, ilAASCodes(), tmChfAdvtExt(), 0, lbcSelection(10), lbcManyCntCode
                            gPopCntrBoxViaArray 0, ilAASCodes(), tmChfAdvtExt(), 0, lbcSelection(10), tgManyCntCodeCT(), sgManyCntCodeTagCT
                        ElseIf rbcSelCSelect(1).Value Or rbcSelCSelect(2).Value Then
                            'gPopCntrBoxViaArray 1, ilAASCodes(), tmChfAdvtExt(), 0, lbcSelection(10), tgManyCntCodeCT(), sgManyCntCodeTagCT
                            '5-16-02
                            Screen.MousePointer = vbHourglass
                            'If rbcSelCInclude(0).Value Then        'proposals
                            '    slCntrStatus = "WDCI"              'all types, working, dead, complete, incomplete
                            '    ilHOState = 0
                            'Else                                   'contracts/orders (wide & narrow)
                            '    slCntrStatus = "HO"              'default to holds and orders
                            '    ilHOState = 3                   'in addition to sch holds & orders, show latest GN if applicable,
                            '                                    'plus the revised orders turned proposals (WCI)
                            '    ckcSelC8(0).Enabled = False       'disallow mods to shows as differences for selective choice
                            'End If
                            '10-29-03 no longer separate proposals vs orders--show all intermixed
                            slCntrStatus = ""           'default to show all
                            ilHOState = 3
                            If ilListIndex = CNT_HISTORY Then
                                slCntrStatus = "HO"
                                ilHOState = 1                   'only scheduled holds and orders
                            End If
                            ckcAll.Enabled = True
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                            ckcAll.Visible = True

                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS

                            'mCntrPop slCntrStatus, ilHOState
                            If rbcSelCSelect(1).Value Then
                                mAASCntrPop 1, 8, slCntrStatus, ilHOState
                            Else
                                mAASCntrPop 2, 9, slCntrStatus, ilHOState
                            End If

                            ilHowManyDefined = lbcSelection(0).ListCount
                            'ilHowMany = lbcSelection(0).SelectCount
                            If ilHowManyDefined = 1 Then
                                ckcAll.Value = vbChecked    '11-15-01 True
                            End If
                            Screen.MousePointer = vbDefault
                            If imTerminate Then
                                cmcCancel_Click
                                Exit Sub
                            End If
                        'Else
                        '    gPopCntrBoxViaArray 2, ilAASCodes(), tmChfAdvtExt(), 0, lbcSelection(10), tgManyCntCodeCT(), sgManyCntCodeTagCT
                        End If
                            Screen.MousePointer = vbDefault
                            ckcAll.Enabled = True
                            ckcAll.Value = vbUnchecked  '11-15-01 False
                            ckcAll.Visible = True
                    ElseIf Index = 5 Then                      'BR, advt option
                        Screen.MousePointer = vbHourglass
                        If rbcSelCInclude(0).Value = True Then        'proposals
                            slCntrStatus = "WDCI"              'all types, working, dead, complete, incomplete
                            ilHOState = 0
                        ElseIf rbcSelCInclude(1).Value = True Then      'contracts
                            slCntrStatus = "HO"              'default to holds and orders
                            ilHOState = 3                   'in addition to sch holds & orders, show latest GN if applicable,
                                                            'plus the revised orders turned proposals (WCI)
                            ckcSelC8(0).Enabled = False       'disallow mods to shows as differences for selective choice
                        Else
                            slCntrStatus = ""              'combined proposals vs orders
                            ilHOState = 3                   'in addition to sch holds & orders, show latest GN if applicable,
                                                            'plus the revised orders turned proposals (WCI)
                        End If

                        If ilListIndex = CNT_HISTORY Then
                            'slCntrStatus = "HO"
                            'ilHOState = 1                   'only scheduled holds and orders
                            slCntrStatus = ""       '5-9-17
                            ilHOState = 3
                        End If
                        ckcAll.Enabled = True
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                        ckcAll.Visible = True

                        gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        mCntrPop slCntrStatus, ilHOState

                        ilHowManyDefined = lbcSelection(0).ListCount
                        'ilHowMany = lbcSelection(0).SelectCount
                        If ilHowManyDefined = 1 Then
                            ckcAll.Value = vbChecked    '11-15-01 True
                        End If
                        Screen.MousePointer = vbDefault
                        If imTerminate Then
                            cmcCancel_Click
                            Exit Sub
                        End If
                    ElseIf Index = 6 Then                       'vehicle selection on Insrtion Orders
                        imSetAllVeh = False
                        CkcAllveh.Value = vbUnchecked   '11-15-01 False
                        imSetAllVeh = True
                    Else                                        'index 0
                        imSetAll = False
                        ckcAll.Value = vbUnchecked  '11-15-01 False
                        ckcAll.Visible = True
                        imSetAll = True
                    End If
                    'If rbcSelCSelect(0).Value Then      'advt selection
                    '   If index = 5 Then
                    '       mCntrPop igRptType
                    '   End If
                    'Else
                    '    imSetAll = False
                    '    ckcAll.Value = False
                    '    imSetAll = True
                    'End If
                Case 1                                    'paperwork
                    'If rbcSelCSelect(0).Value Then
                    '    If Index = 5 Then
                    '        'mCntrPop slCntrStatus
                    '    End If
                    'Else
                    gUncheckAll RptSelCt!ckcAll, imSetAll

                    'End If
                Case 2                                     'summary/spots by advt
                    If rbcSelCSelect(0).Value Then
                        If Index = 5 Then
                            slCntrStatus = "HO"
                            mCntrPop slCntrStatus, 1        'get only orders (w/o revisions)
                            If imTerminate Then
                                cmcCancel_Click
                                Exit Sub
                            End If
                        Else
                            If Index = 6 Then
                                gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                            End If
                        End If
                    Else
                        If Index = 6 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        Else
                            gUncheckAll RptSelCt!ckcAll, imSetAll

                        End If
                    End If
'                Case 3, 14, CNT_ADVT_UNITS, CNT_AVGRATE  'Spots by times; Missed Spots, advt units ordered, average rate
'                Case 3, 14                    '12-9-16 remove avg rate, remove advt units ordered
                    gUncheckAll RptSelCt!ckcAll, imSetAll
                Case CNT_AVGRATE                                '12-9-16
                    If Index = 6 Then                           'agency box
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    Else
                        gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                    End If
                Case 4, 15, 16  'Projection
                    If rbcSelCInclude(0).Value Then 'Advt/Contract
                        If Index = 5 Then
                            slCntrStatus = ""
                            If ckcSelC5(0).Value = vbChecked Then   '11-15-01
                                slCntrStatus = "H"
                            End If
                            If ckcSelC5(1).Value = vbChecked Then
                                slCntrStatus = slCntrStatus & "O"
                            End If
                            If ilListIndex = CNT_BOB_BYCNT Then             'business booked by contract
                                ilHOState = 2                   'show G & N if latest
                            Else
                                ilHOState = 1                   'dealing with spots, only get the scheduled H & O
                            End If
                            mCntrPop slCntrStatus, ilHOState    'fill the contract list box
                            If imTerminate Then
                                cmcCancel_Click
                                Exit Sub
                            End If
                        End If
                    Else    'Salesperson or Vehicle
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    End If
                Case 5  'Recap
                Case 6, 7  'Placement; Discrepancy

                Case 8  'MG
                    gUncheckAll RptSelCt!ckcAll, imSetAll
                Case 9  'Sales Spot Tracking
                Case 10, 12 'Commercial Change, Affiliate Spot Tracking
                'Case 11 'History
                '    If index = 5 Then
                 '       mCntrPop igRptType
                 '   End If
                Case 13 'Spot Sales
                Case 17 'Quarterly Avails
                    If Index = 6 Or Index = 2 Then
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    End If

                Case CNT_AVG_PRICES
                    'Date: 11/1/2019 added Major/Minor sorts, using CSI calendar for date entry
'                    If Index = 3 Then                   'sales sources
'                        gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
'                    Else                                'vehicle or salesperson
'                        gUncheckAll RptSelCt!ckcAll, imSetAll
'                    End If
                        
                    If UBound(tgVehicle) <> lbcSelection(6).ListCount Then
                        ReDim tgVehicle(0)
                        sgVehicleTag = ""
                        mSellConvVehPop (6)
                    End If
                    If Index = 1 Then           'agency list box
                        If cbcSet1.ListIndex = 1 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 2 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 2 Then       'slsp list box
                        If cbcSet1.ListIndex = 4 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 5 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 3 Then       'bus cat
                        If cbcSet1.ListIndex = 2 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 3 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    'ElseIf Index = 4 Then       'veh group
                    ElseIf Index = 12 Then          '3-18-16 vehicle group single selection
                        If cbcSet1.ListIndex = 6 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                            'ilLoop = lbcSelection(4).ListIndex
                            illoop = lbcSelection(12).ListIndex     '3-18-16
                            ilHowManyDefined = gFindVehGroupInx(illoop, tgVehicleSets1())
                            If ilHowManyDefined = 0 Then
                                lbcSelection(8).Clear
                            Else
                                'smVehGp5CodeTag = ""
                                '3-22-16
                                sgMnfCodeTagCB = ""
                                illoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCB(), sgMnfCodeTagCB, "H" & Trim$(str$(ilHowManyDefined)))
                                'ilLoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(ilHowManyDefined)))
                            End If
                        ElseIf cbcSet2.ListIndex = 7 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                            'ilLoop = lbcSelection(4).ListIndex
                            illoop = lbcSelection(12).ListIndex
                            ilHowManyDefined = gFindVehGroupInx(illoop, tgVehicleSets1())
                            
                            If ilHowManyDefined = 0 Then
                                lbcSelection(8).Clear
                            Else
                                'smVehGp5CodeTag = ""
                                '3-22-16
                                sgMnfCodeTagCB = ""
                                illoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCB(), sgMnfCodeTagCB, "H" & Trim$(str$(ilHowManyDefined)))
                                'ilLoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(ilHowManyDefined)))
                            End If
                        End If

                    ElseIf Index = 5 Then       'advt
                        If cbcSet1.ListIndex = 0 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 1 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 6 Then       'vehicle
                        If cbcSet1.ListIndex = 5 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 6 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 7 Then               'prod prot
                        If cbcSet1.ListIndex = 3 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 4 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 8 Then           'items for vehicle group
                        If cbcSet1.ListIndex = 6 Then
                            gUncheckAll RptSelCt!CkcAllveh, imSetAllVeh
                        ElseIf cbcSet2.ListIndex = 7 Then
                            gUncheckAll RptSelCt!CkcAllveh, imSetAllVeh
                        End If
                    End If
                        
'                Case CNT_ADVT_UNITS                         '6-21-18
'                        gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
'                        gUncheckAll RptSelCt!ckcAll, imSetAll
                Case CNT_SALES_CPPCPM
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                Case CNT_TIEOUT
                    If (Index = 2 Or Index = 6) Then
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    End If
                Case CNT_BOB, CNT_CUMEACTIVITY, CNT_BOBRECAP
                    If ilListIndex = CNT_BOB And rbcSelCInclude(1).Value Then       'Billed & Booked slsp option
                        'If ckcSelC10(1).Value = vbChecked Then      'slsp option with subtotals by vehicle
                            If Index = 6 Then
                                gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                            ElseIf Index = 2 Then
                                gUncheckAll RptSelCt!ckcAll, imSetAll
                            ElseIf Index = 7 Then
                                gUncheckAll RptSelCt!CkcAllveh, imSetAllVeh
                            End If
                        'End If

                    ElseIf ilListIndex = CNT_BOB Then
                        If rbcSelCInclude(2).Value Or rbcSelCInclude(4).Value Or rbcSelCInclude(6).Value Then       'vehicle, vehicle gross/net, vehicle/participant option
                            If Index = 7 Then           'clicked on the vehicle group item list box?
                                gUncheckAll RptSelCt!CkcAllveh, imSetAllVeh
                            Else
                                gUncheckAll RptSelCt!ckcAll, imSetAll
                            End If
                        End If

                        If rbcSelCInclude(4).Value Then   '8-4-00 vehicle with participant splits
                            If Index = 2 Then
                                gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                            Else
                                gUncheckAll RptSelCt!ckcAll, imSetAll
                            End If
                        End If
                        
                        If rbcSelCInclude(3).Value = True Or rbcSelCInclude(0).Value = True Or rbcSelCInclude(5).Value = True Then        '2-16-16 3 = owner, 0 = advt, 5 = agy
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        End If

                    Else
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    End If
                Case CNT_ADVT_UNITS 'determine to turn off ALL box for major set or minor set
                    'Date: 10/9/2019    inititally tgVehicle() was populated from lbcSelection(3) (*see mInitReport) causing the number of records
                    'discrepancy between lbcSelection(6) and tgVehicle() array
                    If UBound(tgVehicle) <> lbcSelection(6).ListCount Then
                        ReDim tgVehicle(0)
                        sgVehicleTag = ""
                        mSellConvVehPop (6)
                    End If
                    If Index = 1 Then           'agency list box
                        If cbcSet1.ListIndex = 1 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 2 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 2 Then       'slsp list box
                        If cbcSet1.ListIndex = 4 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 5 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 3 Then       'bus cat
                        If cbcSet1.ListIndex = 2 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 3 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    'ElseIf Index = 4 Then       'veh group
                    ElseIf Index = 12 Then          'agency
                        If cbcSet1.ListIndex = 6 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                            'ilLoop = lbcSelection(4).ListIndex
                            illoop = lbcSelection(12).ListIndex     '3-18-16
                            ilHowManyDefined = gFindVehGroupInx(illoop, tgVehicleSets1())
                            If ilHowManyDefined = 0 Then
                                lbcSelection(8).Clear
                            Else
                                'smVehGp5CodeTag = ""
                                '3-22-16
                                sgMnfCodeTagCB = ""
                                illoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCB(), sgMnfCodeTagCB, "H" & Trim$(str$(ilHowManyDefined)))
                                'ilLoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(ilHowManyDefined)))
                            End If
                        ElseIf cbcSet2.ListIndex = 7 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                            'ilLoop = lbcSelection(4).ListIndex
                            illoop = lbcSelection(12).ListIndex
                            ilHowManyDefined = gFindVehGroupInx(illoop, tgVehicleSets1())
                            
                            If ilHowManyDefined = 0 Then
                                lbcSelection(8).Clear
                            Else
                                'smVehGp5CodeTag = ""
                                '3-22-16
                                sgMnfCodeTagCB = ""
                                illoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCB(), sgMnfCodeTagCB, "H" & Trim$(str$(ilHowManyDefined)))
                                'ilLoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(ilHowManyDefined)))
                            End If
                        End If

                    ElseIf Index = 5 Then       'advt
                        If cbcSet1.ListIndex = 0 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 1 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 6 Then       'vehicle
                        If cbcSet1.ListIndex = 5 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 6 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 7 Then               'prod prot
                        If cbcSet1.ListIndex = 3 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 4 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 8 Then           'items for vehicle group
                        If cbcSet1.ListIndex = 6 Then
                            gUncheckAll RptSelCt!CkcAllveh, imSetAllVeh
                        ElseIf cbcSet2.ListIndex = 7 Then
                            gUncheckAll RptSelCt!CkcAllveh, imSetAllVeh
                        End If
                    End If
                Case CNT_SALESCOMPARE   'determine to turn off ALL box for major set or minor set
                    If Index = 1 Then           'agency list box
                        If cbcSet1.ListIndex = 1 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 2 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 2 Then       'slsp list box
                        If cbcSet1.ListIndex = 4 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 5 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 3 Then       'bus cat
                        If cbcSet1.ListIndex = 2 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 3 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    'ElseIf Index = 4 Then       'veh group
                    ElseIf Index = 12 Then          '3-18-16 vehicle group single selection
                        If cbcSet1.ListIndex = 6 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                            'ilLoop = lbcSelection(4).ListIndex
                            illoop = lbcSelection(12).ListIndex     '3-18-16
                            ilHowManyDefined = gFindVehGroupInx(illoop, tgVehicleSets1())
                            If ilHowManyDefined = 0 Then
                                lbcSelection(8).Clear
                            Else
                                'smVehGp5CodeTag = ""
                                '3-22-16
                                sgMnfCodeTagCB = ""
                                illoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCB(), sgMnfCodeTagCB, "H" & Trim$(str$(ilHowManyDefined)))
                                'ilLoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(ilHowManyDefined)))
                            End If
                        ElseIf cbcSet2.ListIndex = 7 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                            'ilLoop = lbcSelection(4).ListIndex
                            illoop = lbcSelection(12).ListIndex
                            ilHowManyDefined = gFindVehGroupInx(illoop, tgVehicleSets1())
                            
                            If ilHowManyDefined = 0 Then
                                lbcSelection(8).Clear
                            Else
                                'smVehGp5CodeTag = ""
                                '3-22-16
                                sgMnfCodeTagCB = ""
                                illoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCB(), sgMnfCodeTagCB, "H" & Trim$(str$(ilHowManyDefined)))
                                'ilLoop = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(8), tgMnfCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(ilHowManyDefined)))
                            End If
                        End If

                    ElseIf Index = 5 Then       'advt
                        If cbcSet1.ListIndex = 0 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 1 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 6 Then       'vehicle
                        If cbcSet1.ListIndex = 5 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 6 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 7 Then               'prod prot
                        If cbcSet1.ListIndex = 3 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        ElseIf cbcSet2.ListIndex = 4 Then
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    ElseIf Index = 8 Then           'items for vehicle group
                        If cbcSet1.ListIndex = 6 Then
                            gUncheckAll RptSelCt!CkcAllveh, imSetAllVeh
                        ElseIf cbcSet2.ListIndex = 7 Then
                            gUncheckAll RptSelCt!CkcAllveh, imSetAllVeh
                        End If
                    End If
                Case CNT_MAKEPLAN
                    If (Index = 3) Then
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    End If
                Case CNT_VEHCPPCPM
                    If (Index = 3) Then         'vehicles
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    End If
                    If Index = 2 Then                       'demos
                        imSetAll = False
                        ckcAllAAS.Value = vbUnchecked   '11-15-01 False
                        imSetAll = True
                    End If
                Case CNT_SALESACTIVITY_SS, CNT_SALESPLACEMENT           '7-25-02
                    If (Index = 3) Then             'sales source
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    ElseIf (Index = 2) Then         'sales offices
                        gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                    Else                            'markets or vehicles
                        imSetAllVeh = False
                        CkcAllveh.Value = vbUnchecked   '9-12-02 False
                        imSetAllVeh = True
                    End If
                Case CNT_VEH_UNITCOUNT, CNT_LOCKED, CNT_GAMESUMMARY, CNT_PAPERWORKTAX                 '4-5-06
                    gUncheckAll RptSelCt!ckcAll, imSetAll

                Case CNT_BOBCOMPARE           'determine to turn off ALL box for major set or minor set
                    If cbcSet2.ListIndex = 5 Then           'vehicle selectivity
                        gUncheckAll RptSelCt!ckcAll, imSetAll
                    Else                        'not veh selectivity, but vehicles also shown for selectivity
                        If Index <> 6 Then
                            gUncheckAll RptSelCt!ckcAll, imSetAll
                        Else
                            gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                        End If
                    End If
            End Select

        ElseIf igRptCallType = SLSPCOMMSJOB Then
            If ilListIndex = COMM_PROJECTION Or ilListIndex = COMM_SALESCOMM Then
                If Index = 2 Then               'slsp
                    gUncheckAll RptSelCt!ckcAll, imSetAll
                ElseIf Index = 6 Then          'vehicles
                    gUncheckAll RptSelCt!ckcAllAAS, imSetAllAAS
                End If
            End If
        End If
    Else
        'imSetAll = False
        'ckcAll.Value = False
        'imSetAll = True
    End If

    'If Not imAllClickedAAS Then
    '    illistindex = lbcRptType.ListIndex
    '    ckcAllAAS.Enabled = True
    '    ckcAllAAS.Visible = True
    '    ckcAllAAS.Value = False
    '    lbcSelection(1).Visible = True
    'Else
    '    imSetAllAAS = False
    '    ckcAllAAS.Value = False
    '    imSetAllAAS = True
    'End If
    If ilListIndex = CNT_INSERTION Then
        If Index = 0 Or Index = 5 Or Index = 6 Then
            mEnableSeparateFiles
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
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(RptSelCt, lbcSelection, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(RptSelCt, lbcSelection, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptSelCt
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
    'ilRet = gPopAgyBox(RptSelCt, lbcSelection, Traffic!lbcAgency)
    ilRet = gPopAgyBox(RptSelCt, lbcSelection, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gPopAgyBox)", RptSelCt
        On Error GoTo 0
    End If
    Exit Sub

mAgencyPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************************************************************
'
'                   Sub mAskContractTypes - Ask Holds/Orders contract statuses
'                                         - ask all contract types
'                                         - ask Trades, missed, nc & extra
'                   These currently apply to quarterly avails and Billed/Booked reports
'                   Set the location of plcSelC3 before calling this routine.  Then
'                   all subsequent questions will be in correct locations.
'
'                   Created DH: 9/6/96
' ******************************************************************************************************
'
Private Sub mAskContractTypes()
    Dim ilListIndex As Integer
    ckcSeparateFile.Visible = False
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = CONTRACTSJOB Then
        If (igRptType = 0) And (ilListIndex > 1) Then
            ilListIndex = ilListIndex + 1
        End If
    End If
        'End Select
        'If (igRptType = 0) And (ilListIndex > 1) Then
        '    ilListIndex = ilListIndex + 1
        'End If
        'plcSelC3.Move lacSelCTo.Left, plcSelC2.Top + 230
    'plcSelC3.Caption = "Include"
    
    'Date: 8/26/2019 added drop down list for major/minor sorts; moved controls to accomodate change
    If ilListIndex = CNT_ADVT_UNITS Then
        plcSelC3.Move 120, plcSelC3.Top + 300
    Else
        plcSelC3.Move 120, plcSelC3.Top
    End If
    
    smPaintCaption3 = "Include"
    plcSelC3_Paint
    ckcSelC3(0).Caption = "Holds"
    'ckcSelC3(0).Move 660, -30, 840
    ckcSelC3(0).Move 660, -15, 840
    ckcSelC3(0).Value = vbChecked   'True
    If ckcSelC3(0).Value = vbChecked Then
        ckcSelC3_click 0
    Else
        ckcSelC3(0).Value = vbChecked   'True
    End If
    ckcSelC3(0).Visible = True
    ckcSelC3(1).Value = vbChecked   'True
    ckcSelC3(1).Caption = "Orders"
    'ckcSelC3(1).Move 1500, -30, 900
    ckcSelC3(1).Move 1500, -15, 900
    If ckcSelC3(1).Value = vbChecked Then
        ckcSelC3_click 1
    Else
        ckcSelC3(1).Value = vbChecked   'True
    End If
    ckcSelC3(1).Visible = True
    
    'TTP 10955 - Billed and Booked Cal Spots: slowness reported, possibly due to including digital lines
    If ilListIndex = CNT_BOB Then  'Billed & Booked
        ckcSelC3(10).Value = vbChecked   'True
        ckcSelC3(10).Caption = "Digital"
        ckcSelC3(10).Move 2460, -15, 900
        ckcSelC3(10).Visible = True
    End If
    
    'v81 testing results 3-28-22 - Issue 3: The "Export" radio button has appeared on the Advertiser Units Ordered report
    rbcOutput(4).Visible = False
    ckcSeparateFile.Visible = False
    
    '7-18-01 show selectivity for paperwork summary
    If ilListIndex = CNT_PAPERWORK And igRptCallType = CONTRACTSJOB Then
        ckcSelC3(6).Caption = "Rev"                     '11-17-16  Need to differentiate between Rev working and working proposals
        ckcSelC3(6).Move 2550, -30, 600
        ckcSelC3(6).Visible = True
        ckcSelC3(2).Caption = "Reject"
        ckcSelC3(2).Move 3270, -30, 900
        ckcSelC3(2).Visible = True
        ckcSelC3(3).Caption = "Working"
        ckcSelC3(3).Move 660, 195, 990
        ckcSelC3(3).Visible = True
        ckcSelC3(4).Caption = "Unapproved"      '4-29-09 chged from incomlete
        ckcSelC3(4).Move 1650, 195, 1400
        ckcSelC3(4).Visible = True
        ckcSelC3(5).Caption = "Complete"
        ckcSelC3(5).Move 3120, 195, 1110
        ckcSelC3(5).Visible = True
    End If

    plcSelC3.Visible = True
    'Contract Type selection
    
    plcSelC5.Move 120, plcSelC3.Top + plcSelC3.Height, 4260
    
    plcSelC5.Height = 440
    'plcSelC5.Caption = ""
    smPaintCaption5 = ""
    plcSelC5_Paint
    ckcSelC5(0).Move 660, -30, 1080
    ckcSelC5(0).Caption = "Standard"
    If ckcSelC5(0).Value = vbChecked Then
        ckcSelC5_click 0
    Else
        ckcSelC5(0).Value = vbChecked   'True
    End If
    ckcSelC5(0).Visible = True
    ckcSelC5(1).Move 1800, -30, 1200
    ckcSelC5(1).Caption = "Reserved"
    If ckcSelC5(1).Value = vbChecked Then
        ckcSelC5_click 1
    Else
        ckcSelC5(1).Value = vbChecked   'True
    End If
    ckcSelC5(1).Visible = True
    If tgUrf(0).iSlfCode > 0 Then           'its a slsp thats is asking for this report,
                                            'don't allow them to exclude reserves
        ckcSelC5(1).Enabled = False
    Else
        ckcSelC5(1).Enabled = True
    End If
    ckcSelC5(2).Move 3000, -30, 1080
    ckcSelC5(2).Caption = "Remnant"
    If ckcSelC5(2).Value = vbChecked Then
        ckcSelC5_click 2
    Else
        ckcSelC5(2).Value = vbChecked   'True
    End If
    ckcSelC5(2).Visible = True
    ckcSelC5(3).Move 660, 195, 600
    ckcSelC5(3).Caption = "DR"
    If ckcSelC5(3).Value = vbChecked Then
        ckcSelC5_click 3
    Else
        ckcSelC5(3).Value = vbChecked   'True
    End If
    ckcSelC5(3).Visible = True
    ckcSelC5(4).Move 1260, 195, 1320
    ckcSelC5(4).Caption = "Per Inquiry"
    If ckcSelC5(4).Value = vbChecked Then
        ckcSelC5_click 4
    Else
        ckcSelC5(4).Value = vbChecked   'True
    End If
    ckcSelC5(4).Visible = True
    ckcSelC5(5).Move 2580, 195, 720
    ckcSelC5(5).Caption = "PSA"
    ckcSelC5(5).Value = vbUnchecked 'False
    ckcSelC5(5).Visible = True  '9-12-02 vbChecked 'True
    ckcSelC5(6).Move 3300, 195, 900
    ckcSelC5(6).Caption = "Promo"
    ckcSelC5(6).Value = vbUnchecked 'False
    ckcSelC5(6).Visible = True
    plcSelC5.Visible = True

    If ilListIndex <> CNT_PAPERWORK Then
        plcSelC6.Visible = True
        
        plcSelC6.Move plcSelC5.Left, plcSelC5.Top + plcSelC5.Height
        'plcSelC6.Caption = ""
        smPaintCaption6 = ""
        plcSelC6_Paint
        ckcSelC6(0).Move 660, -30, 840
        ckcSelC6(0).Caption = "Trade"
        If ckcSelC6(0).Value = vbChecked Then
            ckcSelC6_click 0
        Else
            ckcSelC6(0).Value = vbChecked   'True
        End If
        ckcSelC6(0).Visible = True
        ckcSelC6(1).Caption = "Missed"
        ckcSelC6(1).Visible = True
        ckcSelC6(1).Move 1500, -30, 960
        If ckcSelC6(1).Value = vbChecked Then
            ckcSelC6_click 1
        Else
            ckcSelC6(1).Value = vbChecked   'True
        End If
        ckcSelC6(2).Caption = "N/C"
        ckcSelC6(2).Visible = True
        ckcSelC6(2).Move 2460, -30, 600
        If ckcSelC6(2).Value = vbChecked Then
            ckcSelC6_click 2
        Else
            ckcSelC6(2).Value = vbChecked   'True
        End If
        ckcSelC6(3).Caption = "Fill"
        ckcSelC6(3).Visible = True
        ckcSelC6(3).Move 3060, -30, 840
        If ckcSelC6(3).Value = vbChecked Then
            ckcSelC6_click 3
        Else
            ckcSelC6(3).Value = vbChecked   'True
        End If
        If igRptCallType = CONTRACTSJOB Then
            If ilListIndex = CNT_BOB Or ilListIndex = CNT_BOBRECAP Or ilListIndex = CNT_SALESCOMPARE Or ilListIndex = CNT_BOB_BYSPOT Or ilListIndex = CNT_BOBCOMPARE Or ilListIndex = CNT_AVGRATE Or ilListIndex = CNT_AVG_PRICES Or ilListIndex = CNT_ADVT_UNITS Then                  '11-25-02 implement air time vs ntr selectivity
                ckcSelC6(1).Caption = "Air Time"
                ckcSelC6(1).Width = 1200
                If ckcSelC6(1).Value = vbChecked Then
                    ckcSelC6_click 1
                Else
                    ckcSelC6(1).Value = vbChecked   'True
                End If
                ckcSelC6(2).Caption = "NTR"
                If ckcSelC6(2).Value = vbUnchecked Then
                    ckcSelC6_click 2
                Else
                    ckcSelC6(2).Value = vbUnchecked   'True
                End If
                ckcSelC6(2).Move 2580, -30, 720

                ckcSelC6(3).Caption = "HardCost"
                If ckcSelC6(3).Value = vbUnchecked Then
                    ckcSelC6_click 3
                Else
                    ckcSelC6(3).Value = vbUnchecked   'True
                End If
                ckcSelC6(3).Move 3300, -30, 1080

                ckcSelC6(1).Visible = True             'change to Air Time selectivity
                ckcSelC6(2).Visible = True              'change to NTR selectivity
                ckcSelC6(3).Visible = True              'make hard cost option visible
                plcSelC6.Height = 240
                '10-2-06
                plcSelC12.Move 780, plcSelC6.Top + plcSelC6.Height, 1900
                ckcSelC12(0).Caption = "Polit"
                ckcSelC12(0).Value = vbChecked
                ckcSelC12(0).Move 0, -15, 700
                ckcSelC12(1).Value = vbChecked
                ckcSelC12(1).Move 840, -15, 1200
                ckcSelC12(1).Caption = "Non-Polit"
                ckcSelC12(0).Visible = True
                ckcSelC12(1).Visible = True
                plcSelC12.Visible = True
            End If
            If ilListIndex = CNT_AVGRATE Or ilListIndex = CNT_AVG_PRICES Or ilListIndex = CNT_ADVT_UNITS Then        'missed & extra do not apply, see only N/C
                ckcSelC6(1).Visible = False             'default Air time on
                ckcSelC6(3).Visible = False             'no Hard cost option
                'ckcSelC5(5).Visible = True             'dont show psas
                'ckcSelC5(6).Visible = True             'dont show promos
                'move NC option to be next to Trade option, NC replace NTR option
                ckcSelC6(2).Move 1500, -30
                ckcSelC6(2).Caption = "N/C"
                plcSelC10.Move 2380, plcSelC5.Top + plcSelC5.Height, 2500
                ckcSelC10(0).Caption = "AirTime"
                ckcSelC10(0).Move 0, -30, 960
                ckcSelC10(0).Visible = True
                ckcSelC10(0).Value = vbChecked
                ckcSelC10(1).Caption = "Rep"
                ckcSelC10(1).Move 1080, -30, 720
                ckcSelC10(1).Visible = True
                ckcSelC10(1).Value = vbChecked
                plcSelC10.Visible = True
            End If
            
            'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
            'v81 testing results 3-28-22 Issue 1: "separate files per vehicle" checkbox from the Insertion Orders report is appearing
            rbcOutput(4).Visible = False
            ckcSeparateFile.Visible = False
            If ilListIndex = CNT_INSERTION Then
                'TTP 10164 - Move Export where the Email option sits, since email is not visible inthis report
                rbcOutput(4).Top = rbcOutput(3).Top
                ckcSeparateFile.Visible = True
            End If
        End If
        
        If igRptCallType = SLSPCOMMSJOB Then
            If ilListIndex = COMM_PROJECTION Then       'if Billed & Booked report, adjust some of the controls
                ckcSelC3(1).Visible = False             'always include Orders, hide it
                ckcSelC6(1).Visible = False             'missed, nc and extra not applicable for this report, Hide controls
                ckcSelC6(2).Visible = False
                ckcSelC6(3).Visible = False
                plcSelC6.Height = 240
            End If
        End If
    End If
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
    'plcSelC1.Caption = "Month"
    smPaintCaption1 = "Month"
    plcSelC1_Paint
    rbcSelCSelect(0).Caption = "Corporate"
    rbcSelCSelect(0).Left = 660
    rbcSelCSelect(0).Width = 1140
    rbcSelCSelect(1).Caption = "Standard"
    rbcSelCSelect(1).Left = 1840
    rbcSelCSelect(1).Width = 1140
    rbcSelCSelect(0).Visible = True
    rbcSelCSelect(1).Visible = True
    rbcSelCSelect(2).Visible = False
    rbcSelCSelect(3).Visible = False
    rbcSelCSelect(1).Value = True
    If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
        rbcSelCSelect(0).Enabled = False
        rbcSelCSelect(0).Value = False
        rbcSelCSelect(1).Value = True
    Else
        rbcSelCSelect(0).Value = True
    End If
    plcSelC1.Visible = True
End Sub

'
'                   mAskEffDate - Ask Effective Date, Start Year
'                                 and Quarter
'
'                   6/7/97
'
'
Private Sub mAskEffDate()
    lacSelCFrom.Left = 120
    edcSelCFrom.Move 1350, edcSelCFrom.Top, 945
    edcSelCFrom.MaxLength = 10  '8 5/27/99 changed for short form date m/d/yyyy
    lacSelCFrom.Caption = "Effective Date"
    lacSelCFrom.Top = 75
    lacSelCFrom.Visible = True
    edcSelCFrom.Visible = True
    lacSelCTo.Caption = "Year"
    lacSelCTo.Visible = True
    lacSelCTo.Left = 120
    lacSelCTo.Top = edcSelCFrom.Top + edcSelCFrom.Height + 75
    lacSelCTo1.Left = 1580
    lacSelCTo1.Caption = "Quarter"
    lacSelCTo1.Width = 810
    lacSelCTo1.Top = edcSelCFrom.Top + edcSelCFrom.Height + 75
    lacSelCTo1.Visible = True
    edcSelCTo.Move 600, edcSelCFrom.Top + edcSelCFrom.Height + 30, 600
    edcSelCTo1.Move 2340, edcSelCFrom.Top + edcSelCFrom.Height + 30, 300
    edcSelCTo.MaxLength = 4
    edcSelCTo1.MaxLength = 1
    edcSelCTo.Visible = True
    edcSelCTo1.Visible = True
    
    'Date: 1/8/2020 added CSI calendar control for date entry
    CSI_CalFrom.Visible = True: edcSelCFrom.Visible = False
    CSI_CalFrom.Move 1350, edcSelCFrom1.Top, 1080
    CSI_CalFrom.ZOrder 0
End Sub

'
'                 mAskGrossOrNet - Ask By- Gross   or    Net
'                 plcSelC7
'
Private Sub mAskGrossOrNet()
    plcSelC7.Visible = True
    'plcSelC7.Caption = "By"
    smPaintCaption7 = "By"
    plcSelC7_Paint
    rbcSelC7(0).Move 420, 0, 840    'gross button,
    rbcSelC7(0).Caption = "Gross"
    rbcSelC7(1).Move 1380, 0, 660   'net button
    rbcSelC7(1).Caption = "Net"
    rbcSelC7(1).Value = True
    rbcSelC7(2).Caption = "Net-Net"
    rbcSelC7(2).Move 2040, 0, 960
    rbcSelC7(2).Visible = False
End Sub

'
'                   mAskPkgOrHide - Ask
'                        Use-  Package Lines or Hidden Lines
'                        For standard lines - subtract unresolved missed
'                        Count MGs where they air
'                   Use plcSelC1 (rbcSelCSelect)
'                       plcSelC8 (ckcSelC8) for check boxes
Private Sub mAskPkgOrHide(ilListIndex As Integer)
    If ilListIndex = CNT_BOB And igRptCallType = CONTRACTSJOB Then
        smPaintCaption1 = "Use"
        plcSelC12_Paint
        rbcSelCSelect(0).Caption = "Pkg"
        rbcSelCSelect(0).Left = 360
        rbcSelCSelect(0).Width = 720
        If rbcSelCSelect(1).Value Then             'default to hidden
            rbcSelCSelect_click 1
        Else
            rbcSelCSelect(1).Value = True
        End If
        rbcSelCSelect(1).Caption = "Air"
        rbcSelCSelect(1).Left = 1080
        rbcSelCSelect(1).Width = 600
        rbcSelCSelect(2).Visible = False
    Else
        smPaintCaption1 = "Use"
        plcSelC1_Paint
        rbcSelCSelect(0).Caption = "Package Lines"
        rbcSelCSelect(0).Left = 660
        rbcSelCSelect(0).Width = 1540
        If rbcSelCSelect(1).Value Then             'default to hidden
            rbcSelCSelect_click 1
        Else
            rbcSelCSelect(1).Value = True
        End If
        rbcSelCSelect(1).Caption = "Airing Lines"
        rbcSelCSelect(1).Left = 2220
        rbcSelCSelect(1).Width = 1440
        rbcSelCSelect(2).Visible = False
    End If
    plcSelC1.Visible = True 'False
    plcSelC8.Move 120, plcSelC1.Top + plcSelC1.Height, 4400         '2-25-21   following captions were truncated

    plcSelC8.Visible = True
    smPaintCaption8 = ""
    plcSelC8_Paint
    plcSelC8.Height = 690
    ckcSelC8(0).Left = 0   'plcSelC8.Left
    ckcSelC8(0).Caption = "For standard lines- subtract unresolved missed"
    ckcSelC8(0).Width = 4400
    ckcSelC8(0).Visible = True
    ckcSelC8(1).Caption = "Count MGs where they air"
    ckcSelC8(1).Move 0, 195, 4400
    ckcSelC8(1).Visible = True
End Sub

'
'               mAskShowComments
'
'           DH 7-18-01 For paperwork Summary, Ask to show comments
'           on separate lines:  Internal, Other, Change Reason, Cancellation
'
Private Sub mAskShowComments()
    'plcSelC6.Caption = " "
    smPaintCaption6 = ""
    plcSelC6_Paint
    ckcSelC6(0).Caption = "Internal"
    ckcSelC6(0).Value = vbUnchecked 'False
    'ckcSelC6(0).Move 1100, -30, 960
    ckcSelC6(0).Move 0, -30, 960
    ckcSelC6(0).Visible = True
    ckcSelC6(1).Caption = "Other"
    ckcSelC6(1).Value = vbUnchecked 'False
    'ckcSelC6(1).Move 2160, -30, 720
    ckcSelC6(1).Move 960, -30, 800
    ckcSelC6(1).Visible = True
    ckcSelC6(2).Caption = "Change Rsn"
    ckcSelC6(2).Value = vbUnchecked 'False
    'ckcSelC6(2).Move 1100, 195, 1440
    ckcSelC6(2).Move 1740, -30, 1440
    ckcSelC6(2).Visible = True
    ckcSelC6(3).Caption = "Cancel"
    ckcSelC6(3).Value = vbUnchecked 'False
    'ckcSelC6(3).Move 2400, 195, 1440
    ckcSelC6(3).Move 3060, -30, 1440
    ckcSelC6(3).Visible = True
    plcSelC6.Move 120, plcSelC7.Top + plcSelC7.Height, 4000, 480
    plcSelC6.Visible = True
End Sub

'
'                           mAskSumDetailBoth
'                           <input> ilLeft - start position of control
'                                   ilListIndex - report option
'                           rbcSelC4
'
'
Private Sub mAskSumDetailBoth(ilLeft As Integer, ilListIndex As Integer)
    If ilListIndex = CNT_BR Then
        rbcSelC4(0).Caption = "Schedule lines"
        rbcSelC4(0).Left = ilLeft       '840 or 600
        rbcSelC4(0).Width = 1530
    Else
        rbcSelC4(0).Caption = "Detail"
        rbcSelC4(0).Left = ilLeft       '840 or 600
        rbcSelC4(0).Width = 795
    End If
    rbcSelC4(0).Visible = True
    rbcSelC4(1).Caption = "Summary"
    rbcSelC4(1).Left = rbcSelC4(0).Left + rbcSelC4(0).Width     '1635 or 1395
    rbcSelC4(1).Width = 1110
    rbcSelC4(1).Visible = True
    rbcSelC4(2).Caption = "Both"
    rbcSelC4(2).Left = rbcSelC4(1).Left + rbcSelC4(1).Width   '2745 or 2505
    rbcSelC4(2).Width = 675
    rbcSelC4(2).Visible = True
    rbcSelC4(2).Value = True
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
    'ilRet = gPopVehBudgetBox(RptSelCt, 0, 1, lbcSelection(4), lbcBudgetCode)
    ilRet = gPopVehBudgetBox(RptSelCt, 2, 0, 1, lbcSelection(4), tgRptSelBudgetCodeCT(), sgRptSelBudgetCodeTagCT)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mBudgetPopErr
        gCPErrorMsg ilRet, "mBudgetPopErr (gPopVehBudgetBox)", RptSelCt
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
'*      Procedure Name:mCntrPop                        *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                    :7/10/96 -Use new contract status*
'                                                      *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*
'*      19-28-03 gPopCntrForAASBox chged to show Product
'       name instead of ADvt name in cntract list box
'       (new flag: ilshow = 7 for product name)
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
    For illoop = 0 To lbcSelection(5).ListCount - 1 Step 1
        If lbcSelection(5).Selected(illoop) Then
            sgMultiCntrCodeTagCT = ""             'init the date stamp so the box will be populated, 3-10-20 wrong tag was initialized
            ReDim tgMultiCntrCodeCT(0 To 0) As SORTCODE
            lbcMultiCntr.Clear
            slNameCode = tgAdvertiser(illoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
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
            ilShow = 7                  '10-28-03 show product name instead of advt
            ilCurrent = 1
            ilAdfCode = Val(slCode)
            'load up list box with contracts with matching adv
            'ilRet = gPopCntrForAASBox(RptSelCt, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, lbcMultiCntrCode)
            ilRet = gPopCntrForAASBox(RptSelCt, 0, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, tgMultiCntrCodeCT(), sgMultiCntrCodeTagCT)
            If ilRet <> CP_MSG_NOPOPREQ Then
                On Error GoTo mCntrPopErr
                gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", RptSelCt
                On Error GoTo 0
            End If
            For ilIndex = 0 To UBound(tgMultiCntrCodeCT) - 1 Step 1 'lbcMultiCntrCode.ListCount - 1 Step 1
                slName = Trim$(tgMultiCntrCodeCT(ilIndex).sKey)  'lbcMultiCntrCode.List(ilIndex)
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

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imDoubleClickName = False
    imChgMode = False

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

    RptSelCt.Caption = smSelectedRptName & " Report"
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
    smPaintCaption1 = "Select"
    plcSelC1_Paint
    rbcSelCSelect(0).Move 600, 0
    rbcSelCSelect(0).Caption = "Advt"
    rbcSelCSelect(1).Move 1290, 0
    rbcSelCSelect(1).Caption = "Agency"
    rbcSelCSelect(2).Move 2020, 0
    rbcSelCSelect(2).Caption = "Salesperson"
    plcSelC2.Move 120, 885
    'plcSelC2.Caption = "Include"
    smPaintCaption2 = "Include"
    plcSelC2_Paint
    rbcSelCInclude(0).Move 705, 0
    rbcSelCInclude(0).Caption = "All"
    rbcSelCInclude(1).Move 1245, 0
    rbcSelCInclude(2).Move 2655, 0
    plcSelC3.Move 120, 675
    'plcSelC3.Caption = "Zone"
    smPaintCaption3 = "Zone"
    plcSelC3_Paint
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
'    pbcSelC.Move 90, 255, 4515, 4230       'move to init controls
    gCenterStdAlone RptSelCt
    
    mFillSortOptions cbcSort1, False
    mFillSortOptions cbcSort2, True

    If ckcSelC8(0).Enabled = False Then ckcSelC8(0).Value = False
    If ckcSelC8(1).Enabled = False Then ckcSelC8(1).Value = False
End Sub

'
'           mInitControls - set controls to proper positions, sizes
'                   hidden, shown, etc.
'
'           Created :  11/28/98 D Hosaka
'
Private Sub mInitControls()
'   1-12-21 change all left positions for list boxes to the right by a little (from 15 to 120)
    lbcSelection(0).Move 120, ckcAll.Height + 30, 4380, 3270
    lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(2).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(3).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(5).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(5).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width / 2 - 30, lbcSelection(5).Height  '1110
    lbcSelection(6).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height
    lbcSelection(0).Move lbcSelection(5).Left + lbcSelection(5).Width + 60, lbcSelection(0).Top, lbcSelection(0).Width / 2 - 30, lbcSelection(0).Height '840
    lbcSelection(7).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width, lbcSelection(5).Height       'advt
    lbcSelection(8).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width, lbcSelection(5).Height       'agy
    lbcSelection(9).Move lbcSelection(5).Left, lbcSelection(5).Top, lbcSelection(5).Width, lbcSelection(5).Height       'slsp
    lbcSelection(10).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height      'cnt
    lbcSelection(11).Move lbcSelection(0).Left, lbcSelection(0).Top, lbcSelection(0).Width, lbcSelection(0).Height      'demo
    lbcSelection(12).Move lbcSelection(6).Left, lbcSelection(6).Top, lbcSelection(6).Width, lbcSelection(6).Height      '

    lbcSelection(0).Visible = False
    lbcSelection(1).Visible = False
    lbcSelection(2).Visible = False
    lbcSelection(3).Visible = False
    lbcSelection(4).Visible = False
    lbcSelection(5).Visible = False
    lbcSelection(6).Visible = False
    lbcSelection(7).Visible = False
    lbcSelection(8).Visible = False
    lbcSelection(9).Visible = False
    lbcSelection(10).Visible = False
    lbcSelection(11).Visible = False
    lbcSelection(12).Visible = False        'single selection

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
    ckcSelC6(0).Value = vbUnchecked 'False
    ckcSelC6(1).Visible = False
    ckcSelC6(1).Value = vbUnchecked 'False
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
    cbcSel.Visible = False
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
    'pbcSelC.Height = 3195
    edcSelCTo.Text = ""
    ckcAll.Move lbcSelection(1).Left            'readjust 'Check All' location to be above left most list box
    ckcAllAAS.Move ckcAll.Left, ckcAll.Top
    ckcAll.Enabled = True
    laclbcName(0).Visible = False
    laclbcName(1).Visible = False
    
    cbcEMailContent.Top = frcEMail.Top + lacContent.Top + lacContent.Height / 2 - cbcEMailContent.Height / 2
    cbcEMailContent.Left = frcEMail.Left + lacContent.Left + lacContent.Width
    pbcSelC.Height = 4500          '1-12-21 move this to initcontrols to make all selection screens higher
    frcOption.Height = 5000         '1-12-12 (was 4600), 8/10/23 TTP 10745 (was 4740)
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
'*                                                     *
'*******************************************************
Private Sub mInitReport()
    Dim ilIndex As Integer
    Dim ilRet As Integer
    ilRet = gPopExportTypes(cbcFileType)       '10-20-01
    pbcSelC.Visible = False
    sgPhoneImage = mkcPhone.Text
    lbcRptType.Clear
    sgMultiCntrCodeTagCT = ""           '10-16-02 init to reread the contract list
    
    Select Case igRptCallType
        Case SLSPCOMMSJOB
            Screen.MousePointer = vbHourglass
            'rptselct.Caption = "Commission Report Selection"
            'frcOption.Caption = "Commission Selection"
            lbcRptType.AddItem "Sales Commissions on Billing", 0
            lbcRptType.AddItem "Billed and Booked Commissions", 1
            'lbcRptType.ListIndex = COMM_SALESCOMM
            mSPersonPop lbcSelection(2)
            If imTerminate Then
                Exit Sub
            End If
            frcOption.Enabled = True
            pbcOption.Enabled = True
            frcOption.Visible = True
            pbcOption.Visible = True
        Case CONTRACTSJOB
            'If igRptType = 3 Then   'Spot week dump
            '    'Set first parameter to 1 for preview
            '    gSpotWeekDumpRpt 0, "SpotWkDp.Lst", Val(smLogUserCode), imVefCode, smVehName, lmNoRecCreated
            '    imTerminate = True
            '    Exit Sub
            'End If
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
            'rptselct.Caption = "Contract Report Selection"
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
            'ilRet = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(11), lbcDemoCode, "D")
            lbcRptType.AddItem "Proposals/Contracts", 0         '0=proposal
            lbcRptType.AddItem "Paperwork Summary", 1           '1=paperwork summary (contract summaries)

            'If tgUrf(0).islfCode = 0 Then           'its a slsp thats is asking for this report,
                                                    'don't allow them to exclude reserves
                ilIndex = 2
                If igRptType = 0 Then   'Proposal
                    'rbcRptType(2).Visible = False
                Else    'Contract
                    'rbcRptType(2).Caption = "Spots by Advt"
                    lbcRptType.AddItem "Spots by Advertiser", ilIndex   '2=spots by advt
                    ilIndex = ilIndex + 1
                End If
                lbcRptType.AddItem "Spots by Date & Time", ilIndex      '3=spots by date & time
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Business Booked by Contract", ilIndex  '4=projection (named changed to Business Booked)
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Contract Recap", ilIndex            '5=contr recap
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot Placements", ilIndex           '6=Spot placements
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot Discrepancies", ilIndex        '7=spot discrepancies
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "MG's", ilIndex                      '8=makegood
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Spot Tracking", ilIndex       '9=sales spot traking
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Commercial Changes", ilIndex        '10=coml changes
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Contract History", ilIndex          '11 Contract history
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Affiliate Spot Tracking", ilIndex   '12 affil spot traking
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Spot Sales", ilIndex                '13=spot sales
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Missed Spots", ilIndex              '14=missed spots
                ilIndex = ilIndex + 1

                'lbcRptType.AddItem "Business Booked by Spot", ilIndex    '15=spot projection (name changed to Business Booked)
                '1-31-00 name chg from business booked by spot to spot business booked
                lbcRptType.AddItem "Spot Business Booked", ilIndex    '15=spot projection (name changed to Business Booked)
                ilIndex = ilIndex + 1
                'spot reprints - used
                lbcRptType.AddItem "Business Booked by Spot Reprint", ilIndex   '16= Business booked reprint
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Avails", ilIndex                    '17=quarterly summary & detail avails
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Average Spot Prices", ilIndex       '18=avg spot prices
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Advertiser Units Ordered", ilIndex  '19=advt units ordered
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Analysis by CPP & CPM", ilIndex '20=sales analysis by cpp & cpm
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Average 30" & """" & " Unit Rate", ilIndex   'Date: 1/23/2020 changed report name --> 21=Average Rate
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Tie-Out", ilIndex                  '22=Detail Tie Out
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Billed and Booked", ilIndex        '23=Billed & booked by advt, Slsp, owner, vehicle
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Weekly Sales Activity by Quarter", ilIndex   '24=Sales Activity
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Comparison", ilIndex         'Sales Comparison by Advt, Slsp, Agy, comp code, Bus code
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Weekly Sales Activity by Month", ilIndex       'Cumulative Activity Report (pacing)
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Average Prices to Make Plan", ilIndex       'Avg Prices needed to make plan
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "CPP/CPM by Vehicle", ilIndex        'Curent cpp/cpm by vehicle
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Analysis Summary", ilIndex        'Sales Analysis Summary
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Insertion Orders", ilIndex
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Daily Sales Activity by Contract", ilIndex  '6-5-01
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Daily Sales Activity by Month", ilIndex  '7-25-01
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Sales Placement", ilIndex  '7-25-01
                ilIndex = ilIndex + 1                           '7-15-03
                lbcRptType.AddItem "Vehicle Unit Count", ilIndex
                ilIndex = ilIndex + 1                           '7-15-03
                lbcRptType.AddItem "Billed and Booked Recap", ilIndex    '4-14-05
                ilIndex = ilIndex + 1                           '7-15-03
                lbcRptType.AddItem "Locked Avails", ilIndex    '4-5-06
                ilIndex = ilIndex + 1                           '7-14-06
                lbcRptType.AddItem "Event Summary", ilIndex
                ilIndex = ilIndex + 1                           '4-09-07
                lbcRptType.AddItem "Paperwork Tax Summary", ilIndex
                ilIndex = ilIndex + 1                           '9-13-07
                lbcRptType.AddItem "Billed and Booked Comparisons", ilIndex
                ilIndex = ilIndex + 1                           '4-8-13
                lbcRptType.AddItem "Contract Verification", ilIndex
                ilIndex = ilIndex + 1
                lbcRptType.AddItem "Insertion Order Activity Log", ilIndex  'CNT_INSERTION_ACTIVITY
                ilIndex = ilIndex + 1                           '4-1-16
                lbcRptType.AddItem "Proposal XML Activity Log", ilIndex
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
    sgEMailContentType = ""
    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
        'rptselct.Caption = smSelectedRptName & " Report"
        'frcOption.Caption = smSelectedRptName & " Selection"
        'slStr = Trim$(smSelectedRptName)
        'ilLoop = InStr(slStr, "&")
        'If ilLoop > 0 Then
        '    slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
        'End If
        'frcOption.Caption = slStr & " Selection"
        If smSelectedRptName = "Insertion Orders" Then
            sgEMailContentType = "I"
        End If
        If sgEMailContentType <> "" Then
            mEMailContentPop
        End If
    End If
    mSetCommands
    
    'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)

    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
End Sub

'
'           mInvSortPop - Populate invoice sort descriptions
'               from "MNF type V"
'
'           Created:  11/28/98 D Hosaka
'
'
'
Private Sub mInvSortPop()
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = cbcSet1.ListIndex
    If ilIndex > 1 Then
        slName = cbcSet1.List(ilIndex)
    End If
    ilRet = gPopMnfPlusFieldsBox(RptSelCt, cbcSet1, tgMnfCodeCT(), sgMNFCodeTagCT, "V")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mInvSortPop
        gCPErrorMsg ilRet, "mInvSortPop (gPopMnfPlusFieldsBox)", RptSelCt
        On Error GoTo 0
        cbcSet1.AddItem "[None]", 0 'Force as first item on list
        'If ilIndex > 1 Then
        '    gFindMatch slName, 2, cbcSet1
        'Else
        '    cbcSet1.ListIndex = ilIndex
        'End If
    End If
    Exit Sub
mInvSortPop:
    On Error GoTo 0
    imTerminate = True
End Sub

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
    ilRet = gIMoveListBox(RptSelCt, lbcLocal, tlSortCode(), slSortCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMnfPopErr
        gCPErrorMsg ilRet, "mMnfPop (gImoveListBox)", RptSelCt
        On Error GoTo 0
    End If
    Exit Sub
mMnfPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'******************************************************************
'*                                                                *
'*      Procedure Name:mCntSelectivity1                           *
'*                                                                *
'*             Created:5/17/93       By:D. LeVine                 *
'*            Modified:              By:                          *
'*                                                                *
'*            Comments:                                           *
'*              3-18-03 some report options moved to              *
'*              mcntselectivity0 (module too large                *
'*              Contract BR, Paperwork Summary ,                  *
'*              Spots by Advt & Insertion Order moved             *
'*                                                                *
'*      10-15-04 change the gathering of rate cards based         *
'*      on the entered date & gather all rate cards from that date on
'******************************************************************
Private Sub mCntSelectivity1()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDate                                                                                *
'******************************************************************************************
    Dim slStr As String
    Dim ilListIndex As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    ReDim ilAASCodes(0 To 1) As Integer
    Dim slAirOrder As String                'from spf, bill as ordered, aired
    Dim ilSort As Integer                   'for book names, 0 = sort by name only, 1 = sort by date then name
    Dim ilShow As Integer                   'for book names, 0 = show book names, 1 = show names & dates
    Dim illoop As Integer
    Dim ilShowNone As Integer
    Dim ilTop As Integer

    ilListIndex = lbcRptType.ListIndex
    If (igRptType = 0) And (ilListIndex > 1) Then
        ilListIndex = ilListIndex + 1
    End If
    
    Select Case ilListIndex
        Case CNT_BOB_BYCNT                          'Business Booked by Cnt
            plcSelC3.Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = False
            ckcAll.Visible = False
            lacSelCFrom.Caption = "Proj. From Date"
            lacSelCFrom.Visible = True
            edcSelCFrom.Visible = True
            lacSelCTo.Visible = False
            edcSelCTo.Visible = False
            
            'Date: 12/16/2019 added CSI calendar control for date entry
            CSI_CalFrom.Visible = True
            CSI_CalFrom.Left = lacSelCFrom.Left + lacSelCFrom.Width + 10
            CSI_CalFrom.Width = 1170
            CSI_CalFrom.ZOrder 0
            edcSelCFrom.Visible = False
            
            'plcSelC1.Caption = "Select"
            smPaintCaption1 = "Select"
            plcSelC1_Paint
            rbcSelCSelect(0).Caption = "Week"
            rbcSelCSelect(0).Left = 600
            rbcSelCSelect(0).Width = 800
            If rbcSelCSelect(0).Value Then
                rbcSelCSelect_click 0
            Else
                rbcSelCSelect(0).Value = True
            End If
            rbcSelCSelect(1).Caption = "Std Month"
            rbcSelCSelect(1).Left = 1410
            rbcSelCSelect(1).Width = 1200
            rbcSelCSelect(2).Width = 1300
            rbcSelCSelect(2).Caption = "Corp Month"
            rbcSelCSelect(2).Left = 2580
            rbcSelCSelect(2).Visible = True
            rbcSelCSelect(1).Enabled = True
            rbcSelCSelect(2).Enabled = True
            If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
                rbcSelCSelect(2).Enabled = False
            Else
                rbcSelCSelect(2).Value = True
            End If
            plcSelC1.Visible = True
            'plcSelC2.Caption = "By"
            smPaintCaption2 = "By"
            plcSelC2_Paint
            rbcSelCInclude(0).Caption = "Advt"
            rbcSelCInclude(0).Width = 680
            rbcSelCInclude(0).Left = 600
            rbcSelCInclude(1).Caption = "Salesperson"
            rbcSelCInclude(1).Left = 1290
            rbcSelCInclude(1).Width = 1560
            rbcSelCInclude(2).Caption = "Vehicle"
            rbcSelCInclude(2).Left = 2700
            rbcSelCInclude(2).Width = 1560
            rbcSelCInclude(2).Visible = True
            If rbcSelCInclude(0).Value Then
                rbcSelCInclude_Click 0
            Else
                rbcSelCInclude(0).Value = True
            End If
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = True
            lbcSelection(0).Visible = True
            ckcAll.Caption = "All Advertisers"
            ckcAll.Visible = True
            'mSellConvVVPkgPop 6, False  'include pkg veh with VV
            mSellConvNoNTRPop 6
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            'plcSelC4.Move 120, 360, 4275
            'plcSelC4.Caption = "Option"
            smPaintCaption4 = "Option"
            plcSelC4_Paint
            plcSelC4.Move 120, plcSelC1.Top - (plcSelC2.Top - plcSelC1.Top)
            rbcSelC4(0).Move 600, 0, 825    'gross button,
            rbcSelC4(0).Caption = "Gross"
            rbcSelC4(1).Move 1545, 0, 660   'net button
            rbcSelC4(1).Caption = "Net"
            rbcSelC4(2).Move 2295, 0, 1005  'net net button
            rbcSelC4(2).Caption = "Net-Net"
            rbcSelC4(1).Value = True        'default to net
            rbcSelC4(2).Visible = False        '8-4-15 disallow option to show Net-net, No prepass and cannot get to correct participant % in pif in crystal
            plcSelC4.Visible = True
            plcSelC5.Move 120, plcSelC2.Top + plcSelC2.Height
            'for contract and spot projections, ask include holds/orders
            'plcSelC5.Move 120, plcSelC2.Top + plcSelC2.Height
            'plcSelC5.Caption = "Include"
            ckcSelC5(0).Caption = "Holds"
            ckcSelC5(0).Move 720, -30, 825
            ckcSelC5(0).Visible = True
            ckcSelC5(0).Value = vbChecked   'True
            ckcSelC5(1).Caption = "Orders"
            ckcSelC5(1).Move 1545, -30, 900
            ckcSelC5(1).Visible = True
            ckcSelC5(1).Value = vbChecked   'True
            plcSelC5.Visible = True                            'hold/order boxes
            plcSelC2.Visible = True
            pbcSelC.Visible = True
            pbcOption.Visible = True
            ckcAll.Visible = True
    
        Case CNT_BOB_BYSPOT                                     'Business Booked  by Spots
            ilRet = gPopAgyCollectBox(RptSelCt, "A", lbcSelection(1), lbcNoSort)    'get payees (agencies and advertisers)
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = False
            lacSelCFrom.Caption = "Start Date"
            lacSelCFrom.Visible = True
            edcSelCFrom.Left = 1080
            edcSelCFrom.Width = 1170
            edcSelCFrom.Visible = True
            lacSelCTo.Caption = "# Periods"
            lacSelCTo.Move 2490, lacSelCFrom.Top, 1020
            edcSelCTo.Move 3390, edcSelCFrom.Top, 360
            lacSelCTo.Visible = True
            edcSelCTo.Visible = True
            plcSelC1.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30
            plcSelC1.Height = 420
            
            'Date: 1/7/2020 added CSI calendar control for date entry
            CSI_CalFrom.Visible = True: CSI_CalTo.Visible = False
            CSI_CalFrom.Move 1080, lacSelCFrom.Top, 1080
            'lacSelCTo.Move CSI_CalFrom.Left + CSI_CalFrom.Width + 10, lacSelCFrom.Top, 1020
            'CSI_CalTo.Move lacSelCTo.Left + lacSelCTo.Width - 200, lacSelCTo.Top, 1080
            edcSelCFrom.Visible = False
            CSI_CalTo.ZOrder 0: CSI_CalFrom.ZOrder 0
            
            'plcSelC1.Caption = "Select"
            smPaintCaption1 = "Select"
            plcSelC1_Paint
            rbcSelCSelect(0).Caption = "Week"
            rbcSelCSelect(0).Left = 600
            rbcSelCSelect(0).Width = 800
            If rbcSelCSelect(0).Value Then
                rbcSelCSelect_click 0
            Else
                rbcSelCSelect(0).Value = True
            End If
            rbcSelCSelect(1).Caption = "Std Month"
            rbcSelCSelect(1).Left = 1410
            rbcSelCSelect(1).Width = 1140
            rbcSelCSelect(2).Caption = "Corp Month"
            rbcSelCSelect(2).Left = 2580
            rbcSelCSelect(2).Width = 1320        '9-22-00
            rbcSelCSelect(2).Visible = True
            rbcSelCSelect(3).Caption = "Cal Month"
            rbcSelCSelect(3).Move 600, 195, 2850
            rbcSelCSelect(3).Visible = True
            rbcSelCSelect(1).Enabled = True
            rbcSelCSelect(2).Enabled = True
            If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
                rbcSelCSelect(2).Enabled = False
            Else
                rbcSelCSelect(2).Value = True
            End If
            plcSelC1.Visible = True
            plcSelC2.Move 120, plcSelC1.Top + plcSelC1.Height
            'plcSelC2.Caption = "Sort"
            smPaintCaption2 = "Sort"
            plcSelC2_Paint
            rbcSelCInclude(0).Caption = "Advt"
            rbcSelCInclude(0).Width = 680
            rbcSelCInclude(0).Left = 480
            rbcSelCInclude(1).Caption = "Slsp"
            rbcSelCInclude(1).Left = 2310   '1230
            rbcSelCInclude(1).Width = 840   '1560
            rbcSelCInclude(2).Caption = "Vehicle"
            rbcSelCInclude(2).Left = 3030   '2700
            rbcSelCInclude(2).Width = 960   '1560
            rbcSelCInclude(2).Visible = True
    
            '4-08-08 Add Agency option
            rbcSelCInclude(3).Caption = "Agency"
            rbcSelCInclude(3).Left = 1290
            rbcSelCInclude(3).Width = 1020
            rbcSelCInclude(3).Visible = True
    
            If rbcSelCInclude(0).Value Then
                rbcSelCInclude_Click 0
            Else
                rbcSelCInclude(0).Value = True
            End If
            '1-29-00 add vehicle group sorting
            gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True
            edcSet1.Text = "Vehicle Group"
            cbcSet1.ListIndex = 0
            edcSet1.Move 120, plcSelC2.Top + 30 + plcSelC2.Height
            cbcSet1.Move 360 + edcSet1.Width, edcSet1.Top - 45
            edcSet1.Visible = True
            cbcSet1.Visible = True
    
            plcSelC4.Top = edcSet1.Top + edcSet1.Height
            'plcSelC4.Top = plcSelC2.Top + plcSelC2.Height  '1-29-00
            plcSelC4.Left = 120
            plcSelC4.Visible = True    'True                         'as aired or ordered
            'plcSelC4.Caption = "By"
            smPaintCaption4 = "By"
            plcSelC4_Paint
            rbcSelC4(0).Visible = True
            rbcSelC4(0).Caption = "Gross"
            rbcSelC4(0).Left = 480
            rbcSelC4(0).Width = 960
            rbcSelC4(1).Visible = True
            rbcSelC4(1).Caption = "Net"
            rbcSelC4(1).Left = 1380
            rbcSelC4(1).Width = 600
            rbcSelC4(1).Value = True
            rbcSelC4(2).Visible = False
    
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = True
            lbcSelection(0).Visible = True
            ckcAll.Caption = "All Advertisers"
            ckcAll.Visible = True
    
            'plcSelC8.Move 120, plcSelC4.Top + plcSelC4.Height
            plcSelC8.Move 2360, plcSelC4.Top
            smPaintCaption8 = ""
            plcSelC8_Paint
            ckcSelC8(0).Left = 0
            ckcSelC8(0).Width = 1515
            ckcSelC8(0).Caption = "Summary Only"
            ckcSelC8(0).Value = vbUnchecked 'False
            ckcSelC8(0).Visible = True
            ckcSelC8(1).Visible = False
            ckcSelC8(2).Visible = False
            plcSelC8.Visible = True
    
    
            '4-09-08 add option to include adjustments (AN)
            plcSelC10.Move 120, plcSelC8.Top + plcSelC8.Height
            smPaintCaption8 = ""
            plcSelC10_Paint
            ckcSelC10(0).Left = 0
            ckcSelC10(0).Width = 2400
            ckcSelC10(0).Caption = "Include Adjustments"
            ckcSelC10(0).Value = vbUnchecked 'False
            ckcSelC10(0).Visible = True
            ckcSelC10(1).Visible = False
            ckcSelC10(2).Visible = False
            plcSelC10.Visible = True
            'plcSelC5.Move 120, plcSelC2.top + plcSelC2.Height
           'for contract and spot projections, ask include holds/orders
            plcSelC3.Move 120, plcSelC10.Top + plcSelC10.Height + 30
            mAskContractTypes
            'plcSelC3.Move 120, plcSelC6.Top + plcSelC6.Height + 30
            'The height of plcSeC3 encompasses 6 lines of selectivity for all contract types:
            'holds, std, dr, trade, polit lines; plus missed line
            plcSelC3.Height = 1400  '1160  theheight must remain here, after mAscontractTypes; the placement will be incorrect
            'ckcSelC3(2).Move 660, 890, 1080
            ckcSelC3(2).Move 660, 1130, 1080
            ckcSelC3(2).Caption = "Missed"
            ckcSelC3(2).Value = vbUnchecked 'False
            ckcSelC3(2).Visible = True
            ckcSelC3(3).Move 1620, 1130, 1180
            ckcSelC3(3).Caption = "Cancelled"
            ckcSelC3(3).Value = vbUnchecked 'False
            ckcSelC3(3).Visible = True
            ckcSelC3(4).Move 2820, 1130, 1220
            ckcSelC3(4).Caption = "Hidden"
            ckcSelC3(4).Value = vbUnchecked 'False
            ckcSelC3(4).Visible = True
            ckcSelC3(5).Visible = False
            ckcSelC3(6).Visible = False
            'ckcSelC6(1).Visible = False
            'ckcSelC6(2).Visible = False
            'ckcSelC6(3).Visible = False
    
            '11-28-09 selective contract
            lacSelCFrom1.Move 120, plcSelC3.Top + plcSelC3.Height + 30, 1200
            lacSelCFrom1.Caption = "Contract #"
            edcText.Move lacSelCFrom1.Width, lacSelCFrom1.Top - 30, 960
            edcText.MaxLength = 10
            lacSelCFrom1.Visible = True
            edcText.Visible = True
    
            plcSelC3.Visible = True
            pbcSelC.Visible = True
            pbcOption.Visible = True
            ckcAll.Visible = True
            plcSelC2.Visible = True
        Case CNT_BOB_BYSPOT_REPRINT             'Business Booked by Spots reprint
            'Code to build screen selectivity removed due to lack ofmemory.  Reprint of Spot Business Booked no longer used
        Case CNT_RECAP 'Recap
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            plcSelC3.Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = False
            ckcAll.Visible = False
            lacSelCFrom.Caption = "From Contr #"
            lacSelCFrom.Visible = True
            edcSelCFrom.Visible = True
            lacSelCTo.Caption = "To Contr #"
            lacSelCTo.Visible = True
            edcSelCTo.Visible = True
            pbcSelC.Visible = True
            pbcOption.Visible = True
        'Removed placement & Discrepancy code--see rptselcb
        Case CNT_MG  'MG's
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            plcSelC3.Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = False
            lbcSelection(6).Visible = True
            ckcAll.Caption = "All Vehicles"
            ckcAll.Visible = True
    
            lacSelCFrom.Left = 120
            lacSelCFrom1.Left = 2325
            edcSelCFrom.Move 1290, edcSelCFrom.Top, 945
            edcSelCFrom1.Move 2700, edcSelCFrom.Top, 945
            lacSelCTo.Left = 120
            edcSelCTo.Move 1290, edcSelCTo.Top, 945
            lacSelCTo1.Left = 2325
            edcSelCTo1.Move 2700, edcSelCTo1.Top, 945
            edcSelCTo.MaxLength = 10    '8   5/27/99 changed for short form date m/d/yyyy *** Missed To
            edcSelCTo1.MaxLength = 10   '8   5/27/99 changed for short form date m/d/yyyy *** MG To
            edcSelCFrom.MaxLength = 10  '8   5/27/99 changed for short form date m/d/yyyy *** Missed From
            edcSelCFrom1.MaxLength = 10 '8   5/27/99 changed for short form date m/d/yyyy *** MG From
            lacSelCFrom.Caption = "Missed: From"
            lacSelCFrom1.Caption = "To"
            lacSelCFrom.Visible = True
            lacSelCFrom1.Visible = True
            lacSelCTo.Caption = "MG: From"
            lacSelCTo1.Caption = "To"
            lacSelCTo.Visible = True
            lacSelCTo1.Visible = True
            edcSelCFrom.Visible = True
            edcSelCFrom1.Visible = True
            edcSelCTo.Visible = True
            edcSelCTo1.Visible = True
            
            'Date: 1/6/2020 added CSI calendar control for date entry
            CSI_CalFrom.Visible = True: CSI_CalTo.Visible = True: CSI_From1.Visible = True: CSI_To1.Visible = True
            CSI_CalFrom.Move lacSelCFrom.Left + lacSelCFrom.Width - 200, lacSelCFrom.Top, 1080
            lacSelCFrom1.Move CSI_CalFrom.Left + CSI_CalFrom.Width + 10, CSI_CalFrom.Top
            CSI_CalTo.Move lacSelCFrom1.Left + lacSelCFrom1.Width + 10, lacSelCFrom1.Top, 1080
            CSI_From1.Move CSI_CalFrom.Left, lacSelCTo.Top, 1080
            lacSelCTo1.Move CSI_From1.Left + CSI_From1.Width + 10, CSI_From1.Top
            CSI_To1.Move lacSelCTo1.Left + lacSelCTo1.Width + 10, lacSelCTo1.Top, 1080
            CSI_CalTo.Left = CSI_To1.Left
            
            CSI_From1.ZOrder 0: CSI_To1.ZOrder 0
            CSI_CalTo.ZOrder 0: CSI_CalFrom.ZOrder 0
            edcSelCFrom.Visible = False: edcSelCTo.Visible = False: edcSelCFrom1.Visible = False: edcSelCTo1.Visible = False
            'pbcSelC.Width = pbcSelC.Width + 450
            'pbcSelC.ZOrder 0
            
            plcSelC3.Left = 120
            plcSelC3.Top = CSI_To1.Top + CSI_To1.Height + 100    'Date: 1/7/2020 added CSI calendar controls for date entry --> edcSelCTo1.Top + edcSelCTo1.Height
            'plcSelC3.Caption = "Include"
            smPaintCaption3 = "Include"
            plcSelC3_Paint
            plcSelC3.Height = 240
            ckcSelC3(0).Left = 675
            ckcSelC3(0).Width = 1280
            ckcSelC3(0).Caption = "Makegoods"
            If ckcSelC3(0).Value = vbChecked Then
                ckcSelC3_click 0
            Else
                ckcSelC3(0).Value = vbChecked   'True
            End If
            ckcSelC3(0).Visible = True
            ckcSelC3(1).Left = 2010
            ckcSelC3(1).Width = 1080
            ckcSelC3(1).Caption = "Outside"
            ckcSelC3(1).Value = vbUnchecked 'False
            ckcSelC3(1).Visible = True
            ckcSelC3(2).Visible = False
            ckcSelC3(3).Visible = False
            ckcSelC3(4).Visible = False
            ckcSelC3(5).Visible = False
            ckcSelC3(6).Visible = False
            plcSelC3.Visible = True
            'plcSelC1.Caption = "Select vehicles containing"
            'plcSelC1.Caption = "Vehicle selection for"
            smPaintCaption1 = "Vehicle selection for"
            plcSelC1_Paint
            plcSelC1.Move plcSelC3.Left, plcSelC3.Top + plcSelC3.Height, 4000, 435
            rbcSelCSelect(0).Visible = True
            rbcSelCSelect(0).Caption = "Missed"
            rbcSelCSelect(0).Left = 1800        '2220
            rbcSelCSelect(0).Width = 920
            If rbcSelCSelect(1).Value Then
                rbcSelCSelect_click 1
            Else
                rbcSelCSelect(1).Value = True
            End If
    
            rbcSelCSelect(1).Caption = "MG/Outside"
            rbcSelCSelect(1).Left = 2760    '3120
            rbcSelCSelect(1).Width = 1680
            rbcSelCSelect(1).Visible = True
            rbcSelCSelect(2).Caption = "Either"
            rbcSelCSelect(2).Move 1800, 195, 800
            rbcSelCSelect(2).Visible = True
            rbcSelCSelect(3).Caption = "Both"
            rbcSelCSelect(3).Move 2760, 195, 680
            rbcSelCSelect(3).Visible = True
            plcSelC1.Visible = True
            lacText.Text = "Contract #"
            lacText.Move 120, plcSelC1.Top + plcSelC1.Height + 30, 1200
            edcText.Move lacText.Left + lacText.Width, lacText.Top - 30
            lacText.Visible = True
            edcText.Visible = True
            pbcSelC.Visible = True
            pbcOption.Visible = True
            
        Case CNT_SPOTTRAK  'Sales Spot Tracking
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            plcSelC3.Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = False
            ckcAll.Visible = False
            lacSelCFrom.Left = 120
            lacSelCFrom1.Left = 2415    '2385
            edcSelCFrom.Move 1340, edcSelCFrom.Top, 945
            edcSelCFrom1.Move 2700, edcSelCFrom.Top, 945
            lacSelCTo.Left = 120
            edcSelCTo.Move 1340, edcSelCTo.Top, 945
            lacSelCTo1.Left = 2415      '2385
            edcSelCTo1.Move 2700, edcSelCTo1.Top, 945
            edcSelCTo.MaxLength = 10    '8 5/27/99 changed for short form date m/d/yyyy
            edcSelCTo1.MaxLength = 10   '8
            edcSelCFrom.MaxLength = 10  '8
            edcSelCFrom1.MaxLength = 10 '8
            lacSelCFrom.Caption = "Created: From"
            lacSelCFrom1.Caption = "To"
            lacSelCFrom.Visible = True
            lacSelCFrom1.Visible = True
            lacSelCTo.Caption = "Aired: From"
            lacSelCTo1.Caption = "To"
            lacSelCTo.Visible = True
            lacSelCTo1.Visible = True
            edcSelCFrom.Visible = True
            edcSelCFrom1.Visible = True
            edcSelCTo.Visible = True
            edcSelCTo1.Visible = True
            pbcSelC.Visible = True
            
            'Date: 12/10/2019 added CSI calendar control for date entries
            edcSelCFrom.Visible = False: edcSelCFrom1.Visible = False: edcSelCTo.Visible = False: edcSelCTo1.Visible = False
            CSI_CalFrom.Move 1340, edcSelCFrom.Top, 1080
            CSI_CalTo.Move 2700, edcSelCFrom.Top, 1080
            CSI_From1.Move 1340, edcSelCTo.Top, 1080
            CSI_To1.Move 2700, edcSelCTo1.Top, 1080
            CSI_CalFrom.Visible = True
            CSI_From1.Visible = True
            CSI_CalTo.Visible = True
            CSI_To1.Visible = True
            
            CSI_From1.ZOrder 0: CSI_To1.ZOrder 0
            CSI_CalFrom.ZOrder 0: CSI_CalTo.ZOrder 0:
            
            pbcOption.Visible = True
            plcSelC3.Left = 120
            plcSelC3.Top = plcSelC2.Top
            'plcSelC3.Caption = "Include"
            smPaintCaption3 = "Include"
            plcSelC3_Paint
            plcSelC3.Height = 240
            ckcSelC3(0).Left = 675
            ckcSelC3(0).Width = 1020
            ckcSelC3(0).Caption = "New"
            If ckcSelC3(0).Value = vbChecked Then
                ckcSelC3_click 0
            Else
                ckcSelC3(0).Value = vbChecked   'True
            End If
            ckcSelC3(0).Visible = True
            ckcSelC3(1).Left = 1385
            ckcSelC3(1).Width = 1080
            ckcSelC3(1).Caption = "Printed"
            ckcSelC3(1).Value = vbUnchecked 'False
            ckcSelC3(1).Visible = True
            ckcSelC3(2).Left = 2335
            ckcSelC3(2).Width = 1020
            ckcSelC3(2).Caption = "Deleted"
            ckcSelC3(2).Value = vbUnchecked 'False
            ckcSelC3(2).Visible = True
            ckcSelC3(3).Visible = False
            ckcSelC3(4).Visible = False
            ckcSelC3(5).Visible = False
            ckcSelC3(6).Visible = False
            plcSelC3.Visible = True
            pbcOption.Visible = True
            
        Case CNT_COMLCHG, CNT_AFFILTRAK  'Commercial change, Affiliate Spot Tracking
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            plcSelC3.Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = False
            ckcAll.Visible = False
            lacSelCFrom.Left = 120
            lacSelCFrom1.Left = 2415    '2385
            edcSelCFrom.Move 1340, edcSelCFrom.Top, 945
            edcSelCFrom1.Move 2700, edcSelCFrom.Top, 945
            lacSelCTo.Left = 120
            edcSelCTo.Move 1340, edcSelCTo.Top, 945
            lacSelCTo1.Left = 2415
            edcSelCTo1.Move 2700, edcSelCTo1.Top, 945
            edcSelCTo.MaxLength = 10    '8  5/27/99 changed for short form date m/d/yyyy
            edcSelCTo1.MaxLength = 10   '8
            edcSelCFrom.MaxLength = 10  '8
            edcSelCFrom1.MaxLength = 10 '8
            lacSelCFrom.Caption = "Created: From"
            lacSelCFrom1.Caption = "To"
            lacSelCFrom.Visible = True
            lacSelCFrom1.Visible = True
            lacSelCTo.Caption = "Aired: From"
            lacSelCTo1.Caption = "To"
            lacSelCTo.Visible = True
            lacSelCTo1.Visible = True
            edcSelCFrom.Visible = True
            edcSelCFrom1.Visible = True
            edcSelCTo.Visible = True
            edcSelCTo1.Visible = True
            pbcSelC.Visible = True
            
            'Date: 12/6/2019 added CSI calendar control for date entries
            edcSelCFrom.Visible = False: edcSelCFrom1.Visible = False: edcSelCTo.Visible = False: edcSelCTo1.Visible = False
            'CSI_CalFrom.Width = 1050: CSI_From1.Width = 1050
            CSI_CalFrom.Move 1340, edcSelCFrom.Top, 1080
            CSI_CalTo.Move 2700, edcSelCFrom.Top, 1080
            CSI_From1.Move 1340, edcSelCTo.Top, 1080
            CSI_To1.Move 2700, edcSelCTo1.Top, 1080
            CSI_CalFrom.Visible = True
            CSI_From1.Visible = True
            CSI_CalTo.Visible = True
            CSI_To1.Visible = True
            
            CSI_From1.ZOrder 0: CSI_To1.ZOrder 0
            CSI_CalFrom.ZOrder 0: CSI_CalTo.ZOrder 0
            
            'plcSelC1.Visible = True
            plcSelC3.Left = 120
            plcSelC3.Top = plcSelC2.Top + 60
            'plcSelC3.Caption = "Include"
            smPaintCaption3 = "Include"
            plcSelC3_Paint
            plcSelC3.Height = 240
            ckcSelC3(0).Left = 675
            ckcSelC3(0).Width = 1020
            ckcSelC3(0).Caption = "New"
            If ckcSelC3(0).Value = vbChecked Then
                ckcSelC3_click 0
            Else
                ckcSelC3(0).Value = vbChecked   'True
            End If
            ckcSelC3(0).Visible = True
            ckcSelC3(1).Left = 1385
            ckcSelC3(1).Width = 1080
            ckcSelC3(1).Caption = "Printed"
            ckcSelC3(1).Value = vbUnchecked 'False
            ckcSelC3(1).Visible = True
            ckcSelC3(2).Left = 2335
            ckcSelC3(2).Width = 1020
            ckcSelC3(2).Caption = "Deleted"
            ckcSelC3(2).Value = vbUnchecked 'False
            ckcSelC3(2).Visible = True
            ckcSelC3(3).Visible = False
            ckcSelC3(4).Visible = False
            ckcSelC3(5).Visible = False
            ckcSelC3(6).Visible = False
            plcSelC3.Visible = False    'True 10-25-00 only gather "New", Deleted & Printed dont exist
            pbcOption.Visible = True
        Case CNT_HISTORY  'History
            ckcAll.Enabled = False
            ckcAll.Move lbcSelection(10).Left
            lbcSelection(7).Visible = True      'advt list box for valid users
    
            lbcSelection(10).Visible = True     'selected cntrs for valid users
            plcSelC2.Visible = False
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            ckcAll.Caption = "All Contracts"
            ckcAll.Visible = True
            lacSelCFrom.Caption = "Start Date"
            lacSelCFrom.Visible = True
            edcSelCFrom.Visible = True
            lacSelCTo.Caption = "End Date"
            lacSelCTo.Visible = True
            edcSelCTo.Visible = True
            If rbcSelCSelect(0).Value Then
                rbcSelCSelect_click 0             'default to advt cntrs for common subroutine (mSetupPopAAS)
            Else
                rbcSelCSelect(0).Value = True
            End If
    
            lacTopDown.Move 120, edcSelCTo.Top + edcSelCTo.Height + 90
            lacTopDown.Caption = "Contract #"
            lacTopDown.Visible = True
            edcTopHowMany.Move edcSelCTo.Left, edcSelCTo.Top + edcSelCTo.Height + 60, 945
            edcTopHowMany.MaxLength = 9
            edcTopHowMany = ""
            edcTopHowMany.Visible = True
    
            'mSetupPopAASOption (0)              'populate advertiser box
            plcSelC3.Visible = False
            pbcSelC.Visible = True
            pbcOption.Visible = True
        'removed Spot Sales code --see rptselcb
    
            'Date: 12/17/2019 added CSI calendar control for date entry
            CSI_CalFrom.Visible = True
            CSI_CalFrom.Move lacSelCFrom.Left + lacSelCFrom.Width + 10, lacSelCFrom.Top, 1170
            CSI_CalTo.Visible = True
            CSI_CalTo.Move lacSelCTo.Left + lacSelCTo.Width + 10, lacSelCTo.Top, 1170
            CSI_CalTo.ZOrder 0: CSI_CalFrom.ZOrder 0
            edcSelCFrom.Visible = False: edcSelCTo.Visible = False
    
        Case CNT_QTRLY_AVAILS      'Quarterly Avails by minutes or percent
            mSellConvVVActPop 6, False          'ignore dormant vehicles for the avails
            plcSelC1.Visible = False
            lacSelCTo.Visible = False
            edcSelCTo.Visible = False
            lbcSelection(6).Height = 1500
            lbcSelection(6).Visible = True
            lbcSelection(12).Move 120, 2090  '2-1-05
            lbcSelection(12).Height = 1500
            lbcSelection(12).Width = (lbcSelection(12).Width / 2) - 120
            lbcSelection(12).Visible = True
            laclbcName(0).Visible = True
            laclbcName(0).Caption = "Rate Cards"
            '2-1-05  move the top location of the label
            laclbcName(0).Move lbcSelection(12).Left, lbcSelection(12).Top - laclbcName(0).Height - 60, 2205
    
            lbcSelection(2).Move lbcSelection(12).Left + lbcSelection(12).Width + 240, lbcSelection(12).Top, lbcSelection(12).Width, lbcSelection(12).Height
            ckcAllAAS.Move lbcSelection(2).Left, lbcSelection(2).Top - 285
            ckcAllAAS.Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(3).Visible = False
            lbcSelection(5).Visible = False
    
            ckcAll.Caption = "All Vehicles"
            ckcAll.Visible = True
            plcSelC4.Visible = True                     'qtrly summry, detail or daily
            'plcSelC4.Caption = "By"
            smPaintCaption4 = "By"
            plcSelC4_Paint
            rbcSelC4(0).Caption = "Qtrly Summary"
            rbcSelC4(0).Move 360, 0, 1560
            rbcSelC4(1).Caption = "Qtrly Detail"
            rbcSelC4(1).Move 1920, 0, 1260
            rbcSelC4(0).Visible = True
            rbcSelC4(1).Visible = True
            rbcSelC4(0).Value = True                    'default to qtrly summary
            rbcSelC4(2).Visible = True
            rbcSelC4(2).Caption = "Daily"
            rbcSelC4(2).Move 3180, 0, 680
            rbcSelC4(2).Enabled = False
            plcSelC4.Move 120, 60                       'which avails report
            lacSelCFrom.Move plcSelC4.Left, plcSelC4.Top + plcSelC4.Height + 30
            edcSelCFrom.Move 990, lacSelCFrom.Top - 30
            edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
            edcSelCFrom1.MaxLength = 2
            lacSelCFrom.Caption = "Start Date"
            lacSelCFrom1.Caption = "# Weeks"    '"# Quarters"
            lacSelCFrom1.Move edcSelCFrom.Left + edcSelCFrom.Width + 120, lacSelCFrom.Top, 900
            edcSelCFrom1.Move lacSelCFrom1.Left + lacSelCFrom1.Width, edcSelCFrom.Top, 420
            edcSelCFrom1.Text = "13"            '9-15-09 default to 13 weeks (1 qtr)
            lacSelCFrom.Visible = True
            lacSelCFrom1.Visible = True
            edcSelCFrom.Visible = True
            edcSelCFrom1.Visible = True
            plcSelC1.Top = edcSelCFrom.Top + edcSelCFrom.Height + 30
            plcSelC1.Left = 120
            'plcSelC1.Caption = "Show"
            smPaintCaption1 = "Show"
            plcSelC1_Paint
            plcSelC1.Width = 5100
            plcSelC1.Visible = True
            rbcSelCSelect(0).Caption = "Avails"
            rbcSelCSelect(0).Move 510, 0, 800
            If rbcSelCSelect(0).Value Then
                rbcSelCSelect_click 0
            Else
                rbcSelCSelect(0).Value = True           'default
            End If
            rbcSelCSelect(1).Caption = "Sold"
            rbcSelCSelect(1).Move 1290, 0, 690
            rbcSelCSelect(1).Visible = True
            rbcSelCSelect(2).Caption = "Inventory"
            rbcSelCSelect(2).Move 1980, 0, 1500
            rbcSelCSelect(2).Visible = True
            rbcSelCSelect(3).Visible = True
            rbcSelCSelect(3).Caption = "% Sold"
            rbcSelCSelect(3).Move 3090, 0, 1020
            plcSelC2.Top = plcSelC1.Top + 230   'plcSelC1.Height
            'plcSelC2.Caption = ""
            smPaintCaption2 = ""
            plcSelC2_Paint
            plcSelC2.Visible = True
            rbcSelCInclude(0).Caption = "Dayparts"
            'rbcSelCInclude(0).Width = 1040
            rbcSelCInclude(0).Move 0, 0, 1040
            rbcSelCInclude(1).Caption = "Days in Dayparts"
            rbcSelCInclude(1).Move 1020, 0, 1670
           ' rbcSelCInclude(1).Width = 1670
            rbcSelCInclude(2).Caption = "Dayparts in Days"
            rbcSelCInclude(2).Move 2695, 0, 1750
            'rbcSelCInclude(2).Width = 1750
            rbcSelCInclude(2).Visible = True
            If rbcSelCInclude(0).Value Then
                rbcSelCInclude_Click 0
            Else
                rbcSelCInclude(0).Value = True
            End If
            plcSelC2.Width = 4290
            pbcOption.Visible = True
    
        Case CNT_ADVT_UNITS            'advertiser units sold
            rbcSelCInclude(0).Value = False     'insure that the correct list box is tested in mSetCommands
            edcSelCFrom.MaxLength = 10  '8    5/27/99 changed for short form date m/d/yyyy
            lbcSelection(3).Visible = False
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = False
            lbcSelection(2).Visible = False
            lbcSelection(5).Visible = True          '6-21-18 False
            lbcSelection(6).Visible = False         'Date: 9/24/2019 it was blocking lbcSelection(5); set to False -->True
            lbcSelection(5).Height = 1800           'advt
            lbcSelection(6).Height = 1800           'vehicle
            ckcAllAAS.Move lbcSelection(6).Left, lbcSelection(6).Top + lbcSelection(6).Height + 60
            ckcAllAAS.Caption = "All Advertisers"
            ckcAllAAS.Visible = True
            lbcSelection(5).Move lbcSelection(6).Left, ckcAllAAS.Top + ckcAllAAS.Height + 60, lbcSelection(6).Width
            ckcAll.Caption = "All Vehicles"
            ckcAll.Visible = True
            pbcOption.Visible = True
            pbcSelC.Visible = True
            lacSelCFrom.Width = 900
            lacSelCFrom.Caption = "Start Date"
            lacSelCFrom.Visible = True
            
            'Date: 9/22/2019 added CSI calendar
            CSI_CalFrom.Visible = True
            CSI_CalFrom.Left = 1050
            CSI_CalFrom.Width = 1170
            CSI_CalFrom.ZOrder 0
            edcSelCFrom.Visible = False
            
    '                edcSelCFrom.Text = ""
    '                edcSelCFrom.Left = 1050
    '                edcSelCFrom.Visible = True
    '                edcSelCFrom.Width = 1170
            
            '6-21-18 option to select 1-13 weeks, previously always 13
            lacSelCFrom1.Caption = "# Weeks"
            lacSelCFrom1.Move CSI_CalFrom.Left + CSI_CalFrom.Width + 360, lacSelCFrom.Top, 960
            lacSelCFrom1.Visible = True
            edcSelCFrom1.Move lacSelCFrom1.Left + lacSelCFrom.Width, CSI_CalFrom.Top, 360
            edcSelCFrom1.Text = 13
            edcSelCFrom1.Visible = True
            
            'Date: 8/22/2019 added Major/Minor sorts drop down list
            edcSet1.Text = "Sorts-Major"
            edcSet1.Move 120, lacSelCFrom.Top + lacSelCFrom.Height + 90, 1080
            cbcSet1.Move edcSet1.Left + edcSet1.Width, lacSelCFrom.Top + lacSelCFrom.Height + 60, 1260
            cbcSet1.Visible = True
            edcSet1.Visible = True
    
            edcSet2.Text = "Minor"
            edcSet2.Move cbcSet1.Left + cbcSet1.Width + 120, edcSet1.Top, 600
            cbcSet2.Move edcSet2.Left + edcSet2.Width, cbcSet1.Top, 1260
            cbcSet2.Visible = True
            edcSet2.Visible = True
            
            For illoop = 1 To 2
                If illoop = 1 Then
                    ilShowNone = False
                    mFillSalesCompare ilListIndex, cbcSet1, ilShowNone
                Else
                    ilShowNone = True
                    mFillSalesCompare ilListIndex, cbcSet2, ilShowNone
                End If
            Next illoop
            
            plcSelC1.Visible = False
            plcSelC2.Visible = False
            edcSelCTo.Visible = False
            edcSelCTo1.Visible = False
            lacSelCTo.Visible = False
            lacSelCTo1.Visible = False
        'Case CNT_SALES_CPPCPM, CNT_AVGRATE          'sales by CPP & CPM, average rate
        Case CNT_AVGRATE, CNT_AVG_PRICES             'avg rate & avg spot price
            'This code was taken from Sales Analysis screen input
            If ilListIndex = CNT_AVGRATE Then
                mAgyAdvtPop lbcSelection(1)             '12-9-16
    
                '9-28-11 option to gather by week or month
                plcSelC1.Move 120, 0
                smPaintCaption1 = "Show"
                plcSelC1_Paint
                rbcSelCSelect(0).Caption = "Week"
                rbcSelCSelect(0).Left = 720
                rbcSelCSelect(0).Width = 840
                rbcSelCSelect(1).Caption = "Month"
                rbcSelCSelect(1).Left = 1740
                rbcSelCSelect(1).Width = 1470
                rbcSelCSelect(1).Visible = True
                If rbcSelCSelect(0).Value Then
                    rbcSelCSelect_click 0
                Else
                    rbcSelCSelect(0).Value = True
                End If
                rbcSelCSelect(2).Visible = False
                plcSelC1.Visible = True
                
                edcSelCFrom.Text = ""
                edcSelCTo.Text = ""
                edcSelCTo1.Text = ""
                ckcAll.Visible = False
                lacSelCTo.Caption = "Year"
                lacSelCTo.Visible = True
                lacSelCTo.Move 120, plcSelC1.Top + plcSelC1.Height + 60 '75
                lacSelCTo1.Move 1580, plcSelC1.Top + plcSelC1.Height + 60, 810  '75, 810
                lacSelCTo1.Caption = "Quarter"
                lacSelCTo1.Visible = True
                edcSelCTo.Move 600, plcSelC1.Top + plcSelC1.Height + 30, 600    'edcSelCFrom.Top + 30, 600
                edcSelCTo1.Move 2340, plcSelC1.Top + plcSelC1.Height + 30, 300 'edcSelCFrom.Top + 30, 300
                edcSelCTo.MaxLength = 4
                edcSelCTo1.MaxLength = 1
                edcSelCTo.Visible = True
                edcSelCTo1.Visible = True
                plcSelC2.Move 120, edcSelCTo.Top + edcSelCTo.Height
                plcSelC2.Height = 440
                'plcSelC2.Caption = "Month"
                smPaintCaption2 = "Month"
                plcSelC2_Paint
                plcSelC2.Visible = True
                plcSelC2.Visible = False
                'May need to run by Corporate some day (retain, but dont show)
                rbcSelCInclude(0).Caption = "Corporate"
                rbcSelCInclude(0).Move 660, 0, 1140
                rbcSelCInclude(0).Visible = True
                rbcSelCInclude(1).Caption = "Standard"
                rbcSelCInclude(1).Move 1840, 0, 1140
                rbcSelCInclude(1).Visible = True
                If rbcSelCInclude(1).Value Then             'default to std
                    rbcSelCInclude_Click 1
                Else
                    rbcSelCInclude(1).Value = True
                End If
                'If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
                '    rbcSelCInclude(0).Enabled = False
                'Else
                '    rbcSelCInclude(0).Value = True
                'End If
                rbcSelCInclude(2).Visible = False
                plcSelC4.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
                'plcSelC4.Caption = "Totals by"
                smPaintCaption4 = "Totals by"
                plcSelC4_Paint
                rbcSelC4(0).Caption = "Detail"
                rbcSelC4(0).Move 840, 0, 800
                rbcSelC4(1).Caption = "Summary"
                rbcSelC4(1).Move 1680, 0, 1140
    
                plcSelC4.Visible = True
                rbcSelC4(0).Visible = True
                rbcSelC4(1).Visible = True
                rbcSelC4(2).Visible = False
                If rbcSelC4(1).Value Then             'default to summary
                    rbcSelC4_click 1
                Else
                    rbcSelC4(1).Value = True
                End If
                plcSelC3.Move 120, plcSelC4.Top + plcSelC4.Height
            ElseIf ilListIndex = CNT_AVG_PRICES Then
                'Date: 11/5/2019 commented out
                'mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                                                    'populate when needed
                'Date: 11/5/2019 commented out
                'ilRet = gPopMnfPlusFieldsBox(RptSelCt, RptSelCt!lbcSelection(3), tgMnfCodeCT(), sgMNFCodeTagCT, "S")
    
                lbcSelection(6).Height = 1500
                lbcSelection(6).Visible = True
                lbcSelection(2).Height = 1500
    
                lbcSelection(3).Move 15, 2090
                lbcSelection(3).Height = 1500
    
                ckcAllAAS.Move lbcSelection(6).Left, lbcSelection(6).Top + lbcSelection(6).Height + 60
                ckcAllAAS.Value = vbChecked         'default to all sales sources selected
                lbcSelection(3).Move lbcSelection(6).Left, ckcAllAAS.Top + ckcAllAAS.Height + 30
    
                ckcAllAAS.Visible = True
                lbcSelection(3).Visible = False      'sales source
                lbcSelection(0).Visible = False
                lbcSelection(1).Visible = True      'advertiser
                lbcSelection(2).Visible = False
                lbcSelection(5).Visible = True
                lbcSelection(6).Visible = True      'vehicles
                'ckcAll.Caption = "All Vehicles"
    
                ckcAll.Visible = True
                lacSelCFrom.Left = 120
                edcSelCFrom.Move 990, edcSelCFrom.Top, 945
                edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
                lacSelCFrom.Caption = "Start Date"
                lacSelCFrom.Visible = True
                edcSelCFrom.Visible = True
                edcSelCTo.Visible = False
                'pbcSelC.Height = 1215
                plcSelC2.Top = edcSelCFrom.Top + edcSelCFrom.Height
                plcSelC2.Left = 120
                'plcSelC2.Caption = "By"
                smPaintCaption2 = "By"
                plcSelC2_Paint
                'Date: 11/5/2019 commented out
                'rbcSelCInclude(0).Caption = "Salesperson"
                rbcSelCInclude(0).Left = 510
                rbcSelCInclude(0).Width = 1470
                'Date: 11/5/2019 commented out
                'rbcSelCInclude(1).Caption = "Vehicle"
                rbcSelCInclude(1).Left = 1980
                rbcSelCInclude(1).Width = 1080
                rbcSelCInclude(1).Visible = True
                'Date: 11/5/2019 commented out
    '                    If rbcSelCInclude(1).Value Then
    '                        rbcSelCInclude_Click 1
    '                    Else
    '                        rbcSelCInclude(1).Value = True
    '                    End If
                rbcSelCInclude(2).Visible = False
    
    
                'plcSelC1.Top = edcSelCFrom.Top + edcSelCFrom.Height
                plcSelC1.Move 120, plcSelC2.Top + plcSelC2.Height
                'plcSelC1.Left = 120
                'plcSelC1.Caption = "Show"
                'smPaintCaption1 = "Show"
                'plcSelC1_Paint
                rbcSelCSelect(0).Caption = "Weekly"
                rbcSelCSelect(0).Left = 510
                rbcSelCSelect(0).Width = 1080
                rbcSelCSelect(1).Caption = "Monthly"
                rbcSelCSelect(1).Left = 1620
                rbcSelCSelect(1).Width = 1470
                rbcSelCSelect(1).Visible = True
                If rbcSelCSelect(0).Value Then
                    rbcSelCSelect_click 0
                Else
                    rbcSelCSelect(0).Value = True
                End If
                rbcSelCSelect(2).Visible = False
                plcSelC1.Visible = True
                plcSelC2.Visible = True
                plcSelC3.Move 120, plcSelC1.Top + plcSelC1.Height
    
                'Date: 10/31/2019 added major/minor sort and used CSI calendar for date entry
                rbcSelCInclude(0).Value = False     'insure that the correct list box is tested in mSetCommands
                edcSelCFrom.MaxLength = 10  '8    5/27/99 changed for short form date m/d/yyyy
                lbcSelection(3).Visible = False
                lbcSelection(0).Visible = False
                lbcSelection(1).Visible = False
                lbcSelection(2).Visible = False
                lbcSelection(5).Visible = True          '6-21-18 False
                lbcSelection(6).Visible = False         'Date: 9/24/2019 it was blocking lbcSelection(5); set to False -->True
                lbcSelection(5).Height = 1800           'advt
                lbcSelection(6).Height = 1800           'vehicle
                ckcAllAAS.Move lbcSelection(6).Left, lbcSelection(6).Top + lbcSelection(6).Height + 60
                ckcAllAAS.Caption = "All Advertisers"
                ckcAllAAS.Visible = True
                lbcSelection(5).Move lbcSelection(6).Left, ckcAllAAS.Top + ckcAllAAS.Height + 60, lbcSelection(6).Width
                ckcAll.Caption = "All Vehicles"
                ckcAll.Visible = True
                pbcOption.Visible = True
                pbcSelC.Visible = True
                lacSelCFrom.Width = 900
                lacSelCFrom.Caption = "Start Date"
                lacSelCFrom.Visible = True
                
                'Date: 9/22/2019 added CSI calendar
                CSI_CalFrom.Visible = True
                CSI_CalFrom.Left = 1050
                CSI_CalFrom.Width = 1170
                CSI_CalFrom.ZOrder 0
                edcSelCFrom.Visible = False
                
                'Date: 11/6/2019 commented out; no need to display no of weeks
                'lacSelCFrom1.Caption = "# Weeks"
                'lacSelCFrom1.Move CSI_CalFrom.Left + CSI_CalFrom.Width + 360, lacSelCFrom.Top, 960
                lacSelCFrom1.Visible = False
                'edcSelCFrom1.Move lacSelCFrom1.Left + lacSelCFrom.Width, CSI_CalFrom.Top, 360
                'edcSelCFrom1.Text = 13
                edcSelCFrom1.Visible = False
                
                
                'Date: 8/22/2019 added Major/Minor sorts drop down list
                edcSet1.Text = "Sorts-Major"
                edcSet1.Move 120, lacSelCFrom.Top + lacSelCFrom.Height + 90, 1080
                cbcSet1.Move edcSet1.Left + edcSet1.Width, lacSelCFrom.Top + lacSelCFrom.Height + 60, 1260
                cbcSet1.Visible = True
                edcSet1.Visible = True
    
                edcSet2.Text = "Minor"
                edcSet2.Move cbcSet1.Left + cbcSet1.Width + 120, edcSet1.Top, 600
                cbcSet2.Move edcSet2.Left + edcSet2.Width, cbcSet1.Top, 1260
                cbcSet2.Visible = False
                edcSet2.Visible = False
                
                For illoop = 1 To 2
                    If illoop = 1 Then
                        ilShowNone = False
                        mFillSalesCompare ilListIndex, cbcSet1, ilShowNone
                    Else
                        ilShowNone = True
                        mFillSalesCompare ilListIndex, cbcSet2, ilShowNone
                    End If
                Next illoop
                
                plcSelC1.Move edcSet1.Left, cbcSet1.Top + cbcSet1.Height
                plcSelC3.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height
                smPaintCaption1 = "Show"
                plcSelC1_Paint
                
                smPaintCaption2 = "Include"
                plcSelC2_Paint
                
                cbcSet1.ListIndex = 5
                cbcSet2.ListIndex = 0
                
                plcSelC2.Visible = False
                edcSelCTo.Visible = False
                edcSelCTo1.Visible = False
                lacSelCTo.Visible = False
                lacSelCTo1.Visible = False
                
            End If
    
            mAskContractTypes
    
            'Date: 11/5/2019 commented out
            'gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True      'avg prices not using vehicle groups, but need to populate list box as prepass needs it in common code
            If ilListIndex = CNT_AVGRATE Then
                'date: 11/5/2019 moved to accomodate changes for Avg Spots (Major/Minor sorts added)
                gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True      'avg prices not using vehicle groups, but need to populate list box as prepass needs it in common code
                
                '6-13-02
                edcSet1.Text = "Vehicle Group Sort"
                cbcSet1.ListIndex = 0
                edcSet1.Move 120, plcSelC12.Top + plcSelC12.Height + 30, 1800
                cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 45 '45
                edcSet1.Visible = True
                cbcSet1.Visible = True
    
                'gather data by daypart or daypart with overrides
                plcSelC7.Move 120, edcSet1.Top + edcSet1.Height + 30
                plcSelC7.Height = 480
                smPaintCaption7 = "Use"
                rbcSelC7(0).Caption = "Daypart Name"
                rbcSelC7(1).Caption = "Daypart w/Overrides"
                rbcSelC7(2).Caption = "Agency"              '12-9-16 option to show dp or agency subtotals
                rbcSelC7(0).Move 480, 0, 1440
                rbcSelC7(0).Value = True
                rbcSelC7(1).Move 2040, 0, 2280
                rbcSelC7(2).Move 480, 240, 960
                rbcSelC7(2).Visible = True
               
                rbcSelC7(0).Visible = True
                rbcSelC7(1).Visible = True
                plcSelC7.Visible = True
                lacText.Move 120, plcSelC7.Top + plcSelC7.Height + 30, 1080
    
            ElseIf ilListIndex = CNT_AVG_PRICES Then
                plcSelC13.Move 120, plcSelC12.Top + plcSelC12.Height + 30, 3600
                ckcSelC13(0).Caption = "Use Sales Source as major sort"
                ckcSelC13(0).Move 0, 0, 3600
                ckcSelC13(0).Visible = True
                ckcSelC13(1).Visible = False
                ckcSelC13(2).Visible = False
                smPaintCaption13 = ""
                plcSelC13.Visible = True
                lacText.Move 120, plcSelC13.Top + plcSelC13.Height + 30, 1080
            End If
            lacText.Text = "Contract #"
            'lacText.Move 120, plcSelC7.Top + plcSelC7.Height + 30, 1080
            lacText.Visible = True
            edcText.Move 1200, lacText.Top - 30
            edcText.Visible = True
            mAskGrossNetTNet edcText.Top + edcText.Height + 30      '10-29-10 gross, net t-net option
    
        Case Else
            '2/3/21 - Move remaining cases to mCntSelectivity1b, due to procedure too large
            mCntSelectivity1b ilListIndex
    
    End Select
    
    '        If ilListIndex = 99 Then    'CNT_AVG_PRICESOLD Then                        'avg spot rates
    '            plcSelC3.Left = lacSelCTo.Left
    '            plcSelC3.Top = plcSelC1.Top + 230           'plcSelC2.Height
    If ilListIndex = CNT_ADVT_UNITS Then                    'advt units sold
        plcSelC3.Move lacSelCTo.Left, lacSelCTo.Top
        mAdvtPop lbcSelection(5)
        mAskContractTypes
        'Date: 8/27/2019 commented out; need to be populated for major sorting (Adv, Agncy, Bus Cat, etc.)
        'gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True      'avg prices not using vehicle groups, but need to populate list box as prepass needs it in common code
    
        plcSelC3.Visible = True
    '            If ilListIndex = CNT_ADVT_UNITS Then
            plcSelC4.Move 120, plcSelC12.Top + plcSelC12.Height
            'plcSelC4.Caption = "Show"
            smPaintCaption4 = "Show"
            plcSelC4_Paint
            rbcSelC4(0).Left = 660
            rbcSelC4(0).Caption = "Spot counts"
            rbcSelC4(0).Width = 1320
            rbcSelC4(0).Visible = True
            rbcSelC4(1).Left = 2040
            rbcSelC4(1).Width = 1440
            rbcSelC4(1).Caption = "Unit counts"
            rbcSelC4(1).Visible = True
            If rbcSelC4(0).Value Then
                rbcSelC4_click 0
            Else
                rbcSelC4(0).Value = True
            End If
            rbcSelC4(2).Visible = False
            plcSelC4.Visible = True        'True
            plcSelC8.Move plcSelC4.Left, plcSelC4.Top + 195
            'plcSelC8.Caption = ""
            smPaintCaption8 = ""
            plcSelC8_Paint
            ckcSelC8(0).Caption = "Show spot rates"
            ckcSelC8(0).Move 0, 0, 1800
            ckcSelC8(0).Value = vbChecked   'True
            ckcSelC8(0).Visible = True
            ckcSelC8(1).Caption = "Skip to new page each group"
            ckcSelC8(1).Move 0, 240, 3000
            ckcSelC8(1).Visible = True
            plcSelC8.Height = 480
            plcSelC8.Visible = True
    
            lacText.Move 120, plcSelC8.Top + plcSelC8.Height + 30, 1080
            lacText.Text = "Contract #"
            lacText.Visible = True
            edcText.Move 1200, lacText.Top - 30
            edcText.Visible = True
            ilTop = edcText.Top
            mAskGrossNetTNet ilTop + edcText.Height
    
    ElseIf ilListIndex = CNT_AVGRATE Then       'Average Rate
        lbcSelection(3).Visible = False
        lbcSelection(0).Visible = False
        lbcSelection(1).Visible = False
        lbcSelection(2).Visible = False
        lbcSelection(5).Visible = False
        lbcSelection(6).Visible = True
        ckcAll.Caption = "All Vehicles"
        ckcAll.Visible = True
    ElseIf ilListIndex = CNT_SALESACTIVITY Then
         mAskEffDate
        
        '10-4-13 add net option
        plcSelC7.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
        mAskGrossOrNet
        rbcSelC7(0).Value = True            'default to Gross, they way its always been
        rbcSelC7(2).Visible = False
    
        ckcAll.Visible = False
    ElseIf ilListIndex = CNT_SALESCOMPARE Then
    'cbcSet1:           0 = advt, 1 = agy, 2 = bus cat, 3 = prod prot, 4 = slsp, 5 = vehicle, 6 = veh group
    'cbcset2:  0 = none,1 = advt, 2 = agy, 3 = bus cat, 4 = prod prot, 5 = slsp, 6 = vehicle, 7 = veh group
    'converted into radio buttons which was previous selectivity:
    'rbcselection(0) = advt, (1) = agy, (2) = slsp, (3) = bus cat, (4) = prod prot, (5) vehicle, (6) veh group
    'lbcSelection(1) = list of agencies
    'lbcselection(2) = slsp
    'lbcselection(3) = bus categories, lbcselection(7) = prod prot
    'lbcselection(5) = advt, lbcselection(6) = vehicles
    'lbcselection(7) = single select vehicle group
        lbcSelection(7).Width = 4380
        slAirOrder = tgSpf.sInvAirOrder         'bill as ordred, aired
        mAdvtPop lbcSelection(5)
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
    
        mAskYrMonthPeriods ilListIndex   'ask year, Month, # Periods
    
        smPaintCaption9 = "Month"
        plcSelC9_Paint
        rbcSelC9(0).Caption = "Corp"           'hidden, not applicable
        rbcSelC9(0).Left = 3000
        rbcSelC9(0).Width = 600
        rbcSelC9(1).Caption = "Standard"
        rbcSelC9(1).Left = 660
        rbcSelC9(1).Width = 1140
        rbcSelC9(2).Caption = "Calendar"
        rbcSelC9(2).Left = 1840
        rbcSelC9(2).Width = 1140
        rbcSelC9(0).Visible = False
        rbcSelC9(1).Visible = True
        rbcSelC9(2).Visible = True
        rbcSelC9(3).Visible = False
        rbcSelC9(1).Value = True
        plcSelC9.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30, 3000, 240 '9-27-18 adjust spacing between questions
        plcSelC9.Visible = True
        
        plcSelC10.Visible = True
    '            plcSelC10.Move 0, edcSelCTo.Top + edcSelCTo.Height + 30, 1200
        plcSelC10.Move 0, plcSelC9.Top + plcSelC9.Height + 30, 1200
        smPaintCaption10 = ""
        plcSelC10_Paint
        ckcSelC10(0).Visible = True
        ckcSelC10(0).Caption = "Top Down"
        ckcSelC10(0).Move 120
        ckcSelC10(1).Visible = False  ' not used
        ckcSelC10(2).Value = vbChecked  'False    ' not used
        lacTopDown.Move 1440, plcSelC10.Top
        lacTopDown.Caption = "How Many"
        lacTopDown.Visible = True
        edcTopHowMany.Move 2400, plcSelC10.Top - 15
        edcTopHowMany.Text = ""
        edcTopHowMany.Width = 480
        edcTopHowMany.Height = 300
        edcTopHowMany.MaxLength = 4
        edcTopHowMany.Visible = True
    
        lacSelCTo1.Caption = "Cnt#"
        lacSelCTo1.Move 3000, lacTopDown.Top, 600
        edcSelCTo1.Move 3480, edcTopHowMany.Top, 820
        lacSelCTo1.Visible = True
        edcSelCTo1.Visible = True
        edcSelCTo1.MaxLength = 9                '1-30-06
    
         For illoop = 1 To 2
            If illoop = 1 Then
                ilShowNone = False
                mFillSalesCompare ilListIndex, cbcSet1, ilShowNone
            Else
                ilShowNone = True
                mFillSalesCompare ilListIndex, cbcSet2, ilShowNone
            End If
        Next illoop
        If rbcSelCInclude(0).Value Then             'advt defaulted, show the advt list box
            rbcSelCInclude_Click 0
        Else
            rbcSelCInclude(0).Value = True
        End If
        edcSet1.Text = "Sorts-Major"
        edcSet1.Move 120, plcSelC10.Top + plcSelC10.Height + 90, 1080
        cbcSet1.Move edcSet1.Left + edcSet1.Width, plcSelC10.Top + plcSelC10.Height + 60, 1260
        cbcSet1.Visible = True
        edcSet1.Visible = True
    
        edcSet2.Text = "Minor"
        edcSet2.Move cbcSet1.Left + cbcSet1.Width + 120, edcSet1.Top, 600
        cbcSet2.Move edcSet2.Left + edcSet2.Width, cbcSet1.Top, 1260
        cbcSet2.Visible = True
        edcSet2.Visible = True
    
        plcSelC13.Move 120, cbcSet1.Top + cbcSet1.Height, 4380, 480
    
        ckcSelC13(0).Move 0, 0, 2040
        ckcSelC13(0).Caption = "Advertiser totals"
        'ckcSelC13(1).Move 0, 240, 4380
        ckcSelC13(1).Move 2160, 0, 2360
        ckcSelC13(1).Caption = "Separate politicals "
        ckcSelC13(2).Move 0, 260, 4380
        ckcSelC13(2).Caption = "Use Sales Source as major sort"
        ckcSelC13(2).Value = vbChecked
        ckcSelC13(3).Caption = "New Page"
        ckcSelC13(3).Move 3240, 260, 1200
        ckcSelC13(0).Visible = True
        ckcSelC13(1).Visible = True
        ckcSelC13(2).Visible = True
        ckcSelC13(3).Visible = True
        plcSelC13.Visible = True
    
        plcSelC7.Move 120, plcSelC13.Top + plcSelC13.Height + 30
        mAskGrossOrNet
        rbcSelC7(2).Caption = "T-Net"
        rbcSelC7(2).Visible = True
        plcSelC3.Move plcSelC7.Left, plcSelC7.Top + plcSelC7.Height
        mAskContractTypes
    
        plcSelC1.Move 120, plcSelC12.Top + plcSelC12.Height
    
        mAskPkgOrHide ilListIndex
    
        ckcSelC8(2).Visible = False
        ckcSelC8(0).Value = vbUnchecked 'False
        ckcSelC8(1).Value = vbChecked   'True
        If slAirOrder = "S" Then            'bill as ordered, update as ordered; don't ask adjustment qustions
                                            'always ignore missed & count mgs
            ckcSelC8(0).Visible = False
            ckcSelC8(1).Visible = False
            ckcSelC8(0).Value = vbUnchecked 'False       'ignore missed
            ckcSelC8(1).Value = vbUnchecked 'False       'ignore mgs
        Else                                'as aired
            ckcSelC8(0).Visible = True
            ckcSelC8(1).Visible = True
        End If
        plcSelC8.Height = 480
    
        smPaintCaption11 = "Include"
        plcSelC11_Paint
        plcSelC11.Move 120, plcSelC8.Top + plcSelC8.Height, 4530
        rbcSelC11(0).Caption = "Thru specified month"
        rbcSelC11(1).Caption = "All last year"
        rbcSelC11(0).Move 720, 0, 2280
    '            plcSelC11.Visible = True               '4-6-16 hide thru specified month or all last year & default to thru specified month
    '            rbcSelC11(0).Visible = True
        rbcSelC11(1).Move 3000, 0, 1800
    '            rbcSelC11(1).Visible = True
        rbcSelC11(0).Value = True
        
    '            lacText.Move 120, plcSelC11.Top + plcSelC11.Height + 30         '2-19-16 effec pacing date
    '            edcText.Move 1800, plcSelC11.Top + plcSelC11.Height
        lacText.Move 120, plcSelC8.Top + plcSelC8.Height + 30
        edcText.Move 1800, plcSelC8.Top + plcSelC8.Height
    
        lacText.Visible = True
        edcText.Visible = True
        '9/25/20 - TTP 9952 - show Adj on CNT_SALESCOMPARE
        ckcInclRevAdj.Move 680 + ckcInclZero.Width, 3900
        ckcInclRevAdj.Visible = True
        ckcInclRevAdj.Value = vbChecked
    
    ElseIf ilListIndex = CNT_CUMEACTIVITY Then
        ilRet = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(11), tgRptSelDemoCodeCT(), sgRptSelDemoCodeTagCT, "D")
        mSellConvVVPkgPop 6, False
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mAskEffDate                         'Ask EffectiveDate, StartQtr and Year
        lacSelCFrom1.Move 2535, 75, 1200
        lacSelCFrom1.Caption = "Contract #"
        edcSelCFrom1.Move 3490, edcSelCFrom1.Top, 945
        lacSelCFrom1.Visible = True
        edcSelCFrom1.Visible = True
        edcSelCFrom1.MaxLength = 10 '8  5/27/99 changed for short form date m/d/yyyy
        plcSelC2.Move 120, edcSelCTo.Top + edcSelCTo.Height
        plcSelC2.Height = 440
        'plcSelC2.Caption = "Select"
        smPaintCaption2 = "Select"
        plcSelC2_Paint
        plcSelC2.Visible = True
        rbcSelCInclude(0).Caption = "Advertiser"
        rbcSelCInclude(0).Move 720, 0, 1200
        rbcSelCInclude(0).Visible = True
        If rbcSelCInclude(0).Value Then             'default to advt
            rbcSelCInclude_Click 0
        Else
            rbcSelCInclude(0).Value = True
        End If
        rbcSelCInclude(1).Caption = "Agency"
        rbcSelCInclude(1).Move 1920, 0, 920
        rbcSelCInclude(1).Visible = True
        rbcSelCInclude(2).Caption = "Demo"
        rbcSelCInclude(2).Move 720, 195, 1200
        rbcSelCInclude(2).Visible = True
        rbcSelCInclude(3).Caption = "Vehicle"
        rbcSelCInclude(3).Move 1920, 195, 960
        rbcSelCInclude(3).Visible = True
        plcSelC7.Move 120, plcSelC2.Top + plcSelC2.Height
        mAskGrossOrNet
        plcSelC4.Move plcSelC7.Left, plcSelC7.Top + plcSelC7.Height
        'plcSelC4.Caption = "Totals by"
        smPaintCaption4 = "Totals by"
        plcSelC4_Paint
        mAskSumDetailBoth 840, ilListIndex              'send left position of first button
        If rbcSelC4(0).Value Then             'default to detail
            rbcSelC4_click 0
        Else
            rbcSelC4(0).Value = True
        End If
        rbcSelC4(2).Visible = False
        plcSelC4.Visible = True
    
        plcSelC8.Move 120, plcSelC4.Top + plcSelC4.Height, 4000
        ckcSelC8(0).Caption = "Air Time"
        ckcSelC8(1).Caption = "NTR"
        ckcSelC8(2).Caption = "Hard Cost"
        smPaintCaption8 = "Include"
        plcSelC8_Paint
        ckcSelC8(0).Value = vbChecked
        ckcSelC8(1).Value = vbUnchecked
        ckcSelC8(2).Value = vbUnchecked
        ckcSelC8(0).Move 840, 0, 1080
        ckcSelC8(1).Move 1920, 0, 720
        ckcSelC8(2).Move 2640, 0, 1200
        ckcSelC8(0).Visible = True
        ckcSelC8(1).Visible = True
        ckcSelC8(2).Visible = True
        plcSelC8.Visible = True
    
        plcSelC12.Move 120, plcSelC8.Top + plcSelC8.Height + 60, 4000
        ckcSelC12(0).Caption = "New Activity Only"
        ckcSelC12(0).Move 0, -30, 3000
        ckcSelC12(0).Visible = True
        ckcSelC12(0).Value = vbUnchecked
        plcSelC12.Visible = True
    
        '1-17-06 option to see vehicle subtotals.  Only applies to advt option
        plcSelC13.Move 120, plcSelC12.Top + plcSelC12.Height, 4000
        ckcSelC13(0).Caption = "Include Vehicle Subtotals"
        ckcSelC13(0).Move 0, -30, 3000
        ckcSelC13(0).Visible = True
        ckcSelC13(0).Value = vbChecked
        plcSelC13.Visible = True
    ElseIf ilListIndex = CNT_MAKEPLAN Then
        lbcSelection(12).Clear
        lbcSelection(4).Clear
        lbcSelection(11).Clear
        lacSelCFrom.Caption = "Year"
        lacSelCFrom.Visible = True
        lacSelCFrom.Left = 120
        lacSelCFrom1.Left = 1500
        lacSelCFrom1.Caption = "Quarter"
        lacSelCFrom1.Visible = True
        lacSelCFrom1.Width = 810
        lacSelCFrom1.Top = 75
        lacSelCFrom.Top = 75
        edcSelCFrom.Move 720, edcSelCFrom.Top, 600
        edcSelCFrom1.Move 2280, edcSelCFrom.Top, 300
        edcSelCFrom.MaxLength = 4
        edcSelCFrom1.MaxLength = 1
        edcSelCFrom.Visible = True
        edcSelCFrom1.Visible = True
        plcSelC1.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30
        mAskCorpOrStd
        plcSelC4.Move 120, plcSelC1.Top + plcSelC1.Height
        'plcSelC4.Caption = "Show by"
        smPaintCaption4 = "Show by"
        plcSelC4_Paint
        rbcSelC4(0).Caption = "Week"
        rbcSelC4(0).Move 840, 0, 800
        rbcSelC4(1).Caption = "Quarter"
        rbcSelC4(1).Move 1680, 0, 1080
        plcSelC4.Visible = True
        rbcSelC4(0).Visible = True
        rbcSelC4(1).Visible = True
        rbcSelC4(2).Visible = False
        If rbcSelC4(1).Value Then             'default to quarter
            rbcSelC4_click 1
        Else
            rbcSelC4(1).Value = True
        End If
        plcSelC7.Move 120, plcSelC4.Top + plcSelC4.Height
        'plcSelC7.Caption = "Totals by"
        smPaintCaption7 = "Totals by"
        plcSelC7_Paint
        rbcSelC7(0).Caption = "Detail"
        rbcSelC7(0).Move 840, 0, 800
        rbcSelC7(1).Caption = "Summary"
        rbcSelC7(1).Move 1680, 0, 1160
        plcSelC7.Visible = True
        rbcSelC7(0).Visible = True
        rbcSelC7(1).Visible = True
        rbcSelC7(2).Visible = False
        If rbcSelC7(1).Value Then             'default to summary
            rbcSelC7_click 1
        Else
            rbcSelC7(1).Value = True
        End If
        
        '8-4-10 add vehicle group sorting
        gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True
        edcSet1.Text = "Vehicle Group"
        cbcSet1.ListIndex = 0
        edcSet1.Move 120, plcSelC7.Top + 30 + plcSelC7.Height
        cbcSet1.Move 360 + edcSet1.Width, edcSet1.Top - 45
        edcSet1.Visible = True
        cbcSet1.Visible = True
    
        lbcSelection(3).Move 120, ckcAll.Top + ckcAll.Height + 30, 4380, 1500
        lbcSelection(3).Visible = True
        ckcAll.Caption = "All Vehicles"
        ckcAll.Visible = True
        lbcSelection(4).Visible = True          'budgets
        'lbcSelection(12).Visible = True          'ratecard
        lbcSelection(4).Move lbcSelection(3).Left, lbcSelection(3).Top + lbcSelection(3).Height + 300, lbcSelection(3).Width / 2, lbcSelection(3).Height
        lbcSelection(12).Move lbcSelection(3).Left + lbcSelection(3).Width / 2 + 60, lbcSelection(3).Top + lbcSelection(3).Height + 300, lbcSelection(3).Width / 2, lbcSelection(3).Height
        lbcSelection(11).Move lbcSelection(3).Left + lbcSelection(3).Width / 2 + 60, lbcSelection(3).Top + lbcSelection(3).Height + 300, lbcSelection(3).Width / 2, lbcSelection(3).Height
        laclbcName(0).Visible = True
        laclbcName(0).Caption = "Budget Names"
        laclbcName(1).Visible = True
        laclbcName(1).Caption = "Rate Card"
        laclbcName(0).Move lbcSelection(3).Left, lbcSelection(4).Top - laclbcName(0).Height - 30, 1605
        laclbcName(1).Move lbcSelection(3).Left + lbcSelection(3).Width / 2 + 60, lbcSelection(4).Top - laclbcName(1).Height - 30, 1710
    End If
    frcOption.Visible = True
   ' End Select
End Sub

'******************************************************************
'*                                                                *
'*      Procedure Name:mCntSelectivity1b                          *
'*                                                                *
'*             Created:2/3/21        By:J. White                  *
'*            Modified:              By:                          *
'*                                                                *
'*            Comments:                                           *
'*              Moved part of mCntSelectivity1 case logic here    *
'*              mCntSelectivity1b module was too large            *
'*                                                                *
'******************************************************************
Private Sub mCntSelectivity1b(ilListIndex As Integer)
    Dim ilIndex As Integer
    Dim slStr As String
    Dim slAirOrder As String
    
    Select Case ilListIndex
        Case CNT_TIEOUT                             'Tie-out by Vehicle or office
            mSellConvVVPkgPop 6, False
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            lacSelCFrom.Caption = "Start Date"
            lacSelCFrom.Visible = True
            lacSelCFrom.Left = 120
            edcSelCFrom.MaxLength = 10  ' 8  5/27/99 changed for short form date m/d/yyyy
            edcSelCFrom.Width = 990
            edcSelCFrom.Left = 1020
            edcSelCFrom.Visible = True
            lacSelCFrom1.Caption = "Year"
            lacSelCFrom1.Visible = True
            lacSelCFrom1.Left = 2190
            edcSelCFrom1.MaxLength = 4
            edcSelCFrom1.Width = 600
            edcSelCFrom1.Left = 2730
            edcSelCFrom1.Visible = True
            
            'Date: 1/8/2020 added CSI calendar control for date entry
            CSI_CalFrom.Visible = True: edcSelCFrom.Visible = False
            CSI_CalFrom.Move 1020, edcSelCFrom1.Top, 1080
            CSI_CalFrom.ZOrder 0
            
            plcSelC2.Top = edcSelCFrom.Top + edcSelCFrom.Height
            plcSelC2.Left = 120
            'plcSelC2.Caption = "By"
            smPaintCaption2 = "By"
            plcSelC2_Paint
            plcSelC2.Visible = True
            rbcSelCInclude(0).Caption = "Office"
            rbcSelCInclude(0).Left = 360
            rbcSelCInclude(0).Width = 800
            rbcSelCInclude(1).Caption = "Vehicle"
            rbcSelCInclude(1).Left = 1200
            rbcSelCInclude(1).Width = 960
            'rbcSelCInclude(0).Value = True              'default to office
            rbcSelCInclude(0).Visible = True
            rbcSelCInclude(1).Visible = True
            plcSelC1.Move 120, plcSelC2.Top + plcSelC2.Height
            mAskCorpOrStd
            If rbcSelCInclude(0).Value Then             'default to office
                rbcSelCInclude_Click 0
            Else
                rbcSelCInclude(0).Value = True
            End If
            plcSelC1.Visible = True
            plcSelC2.Visible = True
            pbcSelC.Visible = True
            'Dan M added contract selectivity and ntr/hard cost option 7-16-08
            ckcSelC3(0).Caption = "NTR"
            ckcSelC3(0).Width = 750
            ckcSelC3(0).Visible = True
            ckcSelC3(1).Caption = "Hard Cost"
            smPaintCaption3 = "Include"
            plcSelC3_Paint
            ckcSelC3(0).Left = 750
            ckcSelC3(1).Left = ckcSelC3(0).Left + ckcSelC3(0).Width + 30
            ckcSelC3(1).Width = 2000
            ckcSelC3(1).Visible = True

            plcSelC3.Move 120, plcSelC1.Top + plcSelC2.Height
            plcSelC3.Visible = True
            lacTopDown.Caption = "Contract #"
            lacTopDown.Move 120, plcSelC3.Top + plcSelC3.Height + 50, 1000
            edcText.Move lacTopDown.Left + lacTopDown.Width, plcSelC3.Top + plcSelC3.Height
            lacTopDown.Visible = True
            edcText.Visible = True


            mSalesOfficePop lbcSelection(2)
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            'mSellConvVVPkgPop 6, False                   '5-26-06 this pop has already been done
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mBudgetPop                                      'lbcselection(4), one budget only
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            'setup budget comparisons list box
            For ilIndex = 0 To lbcSelection(4).ListCount - 1 Step 1
                slStr = lbcSelection(4).List(ilIndex)
                lbcSelection(12).AddItem slStr               'lbcselection(2) = comparison budgets
            Next ilIndex
            lbcSelection(4).Visible = True                  'show budget name list box (base budget)
            lbcSelection(12).Visible = True                 'split budgets
            laclbcName(0).Visible = True
            laclbcName(0).Caption = "Direct Budget Names"
            laclbcName(1).Visible = True
            laclbcName(1).Caption = "Split Budget Names"
            lbcSelection(4).Move lbcSelection(1).Left, lbcSelection(1).Top + (lbcSelection(1).Height / 2) + 180, lbcSelection(1).Width / 2 - 60, 1500
            laclbcName(0).Move lbcSelection(4).Left, lbcSelection(4).Top - laclbcName(0).Height - 30, 2205
            lbcSelection(12).Move lbcSelection(4).Left + lbcSelection(4).Width + 60, lbcSelection(1).Top + (lbcSelection(1).Height / 2) + 180, lbcSelection(1).Width / 2, 1500
            laclbcName(1).Move lbcSelection(4).Left + lbcSelection(4).Width + 60, lbcSelection(4).Top - laclbcName(0).Height - 30, 2205
            pbcOption.Visible = True
            pbcOption.Enabled = True
            ' dan M

        Case CNT_BOB                                            'Billed & Booked
            'lbcSelection(1) = agy, lbcSelection(2) = slsp, lbcSelection(5) = advt
            'lbcSelection(6) = vehicle, lbcselection(7) = sales office
            'lbcSelection(2) = participants if vehicle/partipants option
            'lbcSelection(2) = owners (for owner option)
            cbcSet2.Clear
            If tgUrf(0).iSlfCode > 0 Then           'slsp or mgr/planner, disallow the participant info
                cbcSet2.AddItem "Advertiser"
                cbcSet2.AddItem "Agency"
                'cbcSet2.AddItem "Owner"            'disallowed if slsp user
                cbcSet2.AddItem "Salesperson"
                cbcSet2.AddItem "Vehicle"
                'cbcSet2.AddItem "Vehicle Gross/Net"    'disallowed if slsp user
                'cbcSet2.AddItem "Vehicle/Participant"  'disallowed if slsp user
            Else
                cbcSet2.AddItem "Advertiser"
                cbcSet2.AddItem "Agency"
                cbcSet2.AddItem "Owner"
                cbcSet2.AddItem "Salesperson"
                cbcSet2.AddItem "Vehicle"
                cbcSet2.AddItem "Vehicle Gross/Net"
                cbcSet2.AddItem "Vehicle/Participant"
            End If

            slAirOrder = tgSpf.sInvAirOrder         'bill as ordred, aired
            mAdvtPop lbcSelection(5)
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If

            'mAskBobInput                    'ask qtr, year & # Periods
            mAskYrMonthPeriods ilListIndex              '7-1-08

            smPaintCaption2 = "Select"
            plcSelC2_Paint
            rbcSelCInclude(0).Caption = "Advt"
            rbcSelCInclude(0).Left = 660
            rbcSelCInclude(0).Width = 690
            rbcSelCInclude(1).Caption = "Salesperson"
            rbcSelCInclude(1).Left = 1380
            rbcSelCInclude(1).Width = 1440
            rbcSelCInclude(1).Visible = True
            rbcSelCInclude(2).Caption = "Vehicle"
            rbcSelCInclude(2).Left = 2820
            rbcSelCInclude(2).Width = 960
            rbcSelCInclude(1).Enabled = True
            rbcSelCInclude(2).Enabled = True
            rbcSelCInclude(2).Visible = True
            rbcSelCInclude(3).Caption = "Owner"
            rbcSelCInclude(3).Left = 360  '660
            rbcSelCInclude(3).Width = 900
            rbcSelCInclude(3).Top = 195
            rbcSelCInclude(3).Visible = True

            rbcSelCInclude(4).Caption = "Vehicle/Participant"  '8-4-00
            rbcSelCInclude(4).Move 1260, 195, 2400   '1540, 195, 2400
            rbcSelCInclude(4).Visible = True
            rbcSelCInclude(0).Value = True                       'default to advt selection
            rbcSelCInclude(5).Caption = "Agency"
            rbcSelCInclude(5).Move 3180, 195, 960
            rbcSelCInclude(5).Visible = True

            plcSelC2.Move 120, lacSelCFrom.Top + lacSelCFrom.Height + 60
            plcSelC2.Height = 435
            'plcSelC2.Visible = True            'turn off the radio button sort options, change to drop down list box
                                                'and remap it in cbcSet2

            lacTopDown.Caption = "Sort"
            lacTopDown.Move 120, lacSelCFrom.Top + lacSelCFrom.Height + 120
            lacTopDown.Visible = True
            cbcSet2.Move 600, lacSelCFrom.Top + lacSelCFrom.Height + 90, 1760
            cbcSet2.ListIndex = 0           'default to advertiser
            cbcSet2.Visible = True

            smPaintCaption4 = "Totals by"
            plcSelC4_Paint
            plcSelC4.Move plcSelC2.Left, cbcSet2.Top + cbcSet2.Height + 30

            rbcSelC4(0).Caption = "Contract"
            rbcSelC4(0).Left = 900
            rbcSelC4(0).Width = 1080
            rbcSelC4(0).Visible = True
            rbcSelC4(0).Value = True
            If rbcSelCInclude(0).Value Then             'default to advt
                rbcSelCInclude_Click 0
            Else
                rbcSelCInclude(0).Value = True
            End If

            rbcSelC4(1).Caption = "Advertiser"
            rbcSelC4(1).Left = 1920
            rbcSelC4(1).Width = 1200
            rbcSelC4(1).Visible = True
            rbcSelC4(2).Caption = "Summary"
            rbcSelC4(2).Left = 3120
            rbcSelC4(2).Width = 1200
            rbcSelC4(2).Visible = True
            plcSelC4.Visible = True

            plcSelC9.Move 120, plcSelC4.Top + plcSelC4.Height, 2600
            mAskBOBCorpOrStd
            rbcSelC9(3).Visible = True
            rbcSelC9(4).Visible = True          '1-12-21 turn on bill method option
            plcSelC9.Width = 4400
            plcSelC9.Height = 375
            plcSelC7.Move 120, plcSelC9.Top + plcSelC9.Height + 120, 2240           '1-12-21 chg the top, move down 60 twips

            plcSelC7.Visible = True
            smPaintCaption7 = ""
            plcSelC7_Paint
            rbcSelC7(0).Move 0, 0, 820   'gross button,
            rbcSelC7(0).Caption = "Gross"
            rbcSelC7(1).Move 820, 0, 600  'net button
            rbcSelC7(1).Caption = "Net"
            rbcSelC7(1).Value = True
            rbcSelC7(2).Move 1440, 0, 1020
            rbcSelC7(2).Caption = "T-Net"
            rbcSelC7(2).Visible = True

            '3-2-02
            lacSelCTo.Visible = True
            lacSelCTo.Move 2920, plcSelC9.Top + plcSelC9.Height + 60, 840

            lacSelCTo.Caption = "Contr #"
            edcSelCTo.MaxLength = 9             '1-30-06
            edcSelCTo.Move 3580, plcSelC9.Top + plcSelC9.Height - 0, 840

            edcSelCTo.Visible = True

            plcSelC3.Move 120, plcSelC7.Top + plcSelC7.Height
            mAskContractTypes

            plcSelC1.Move 2880, plcSelC6.Top + plcSelC6.Height
            mAskPkgOrHide ilListIndex

            'Use same control (different index) for unrelated question (due to lack of controls)
            ckcSelC8(2).Visible = True
            ckcSelC8(2).Caption = "Skip to new page each new group"
            ckcSelC8(2).Move 0, 405, 3120
            ckcSelC8(2).Visible = True

            ckcSelC8(0).Value = vbUnchecked 'False
            ckcSelC8(1).Value = vbChecked   'True

            'smPaintCaption10 = "Slsp"
            smPaintCaption10 = ""           '11-7-16
            plcSelC10_Paint
            plcSelC10.Move 120, plcSelC8.Top + plcSelC8.Height - 30
            'ckcSelC10(0).Move 480, -30, 1920
            ckcSelC10(0).Move 0, -30, 1320          '11-7-16
            plcSelC10.Enabled = False
            plcSelC10.Visible = True
            ckcSelC10(0).Enabled = False
            ckcSelC10(0).Visible = True '^^^
            ckcSelC10(1).Caption = "Veh Sub-Tots"     '11-7-16
            'ckcSelC10(1).Move 2400, -30, 2280
            ckcSelC10(1).Move 1320, -30, 1560
            ckcSelC10(1).Visible = True
            If ckcSelC10(1).Value = vbChecked Then
                ckcSelC10(1).Value = vbUnchecked    ', False
            Else
                ckcSelC10_click 1
            End If
            ckcSelC10(1).Enabled = False
            '11-7-16 allow office subtotals to be hidden
            ckcSelC10(2).Caption = "Ofc Sub-Tots"
            ckcSelC10(2).Move 3000, -30, 1560
            ckcSelC10(2).Visible = True
            If ckcSelC10(2).Value = vbChecked Then
                ckcSelC10(2).Value = vbUnchecked    ', False
            Else
                ckcSelC10_click 2
            End If
            ckcSelC10(2).Enabled = False
            
            If slAirOrder = "S" Then            'bill as ordered, update as ordered; don't ask adjustment qustions
                                                'always ignore missed & count mgs
                ckcSelC8(0).Visible = False
                ckcSelC8(1).Visible = False
                ckcSelC8(0).Value = vbUnchecked 'False       'ignore missed
                ckcSelC8(1).Value = vbUnchecked 'False       'ignore mgs
            Else                                'as aired
                ckcSelC8(0).Visible = True
                ckcSelC8(1).Visible = True
            End If

            '11-21-05 effect pacing date
            lacText.Move 120, plcSelC10.Top + plcSelC10.Height + 30, 1680
            edcText.Move 1800, plcSelC10.Top + plcSelC10.Height
            edcText.Visible = True
            lacText.Visible = True

            'Show Acquisition $ only (dont use spot rates or rvf $
            plcSelC13.Move edcText.Left + 1020, lacText.Top, 1920
            smPaintCaption13 = ""
            ckcSelC13(0).Width = 1680
            ckcSelC13(0).Caption = "Acq. Cost Only"
            ckcSelC13(0).Visible = True
            plcSelC13.Visible = True
            
            '4-26-13 Include $0 contracts
            ckcInclZero.Move 120, plcSelC13.Top + plcSelC13.Height + 30
            ckcInclZero.Visible = True
            
            '7-29-16 all calendar types now allow Inclusion of revenue adjustments to be an option, default it to include
            ckcInclRevAdj.Move 120 + ckcInclZero.Width, ckcInclZero.Top
            ckcInclRevAdj.Visible = True
            ckcInclRevAdj.Value = vbChecked

            'setup the default for corp std after setting up cntr types & pkg/hide questions;
            'Trades will be defaulted off if running standard;
            'Count mgs where they air set off (no adjustments)

            If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
                rbcSelC9(0).Enabled = False
                rbcSelC9(0).Value = False
                rbcSelC9(1).Value = True
            Else
                If rbcSelC9(0).Value Then
                    rbcSelC9_click 0
                Else
                    rbcSelC9(0).Value = True
                End If
            End If


            If tgUrf(0).iSlfCode > 0 Then
                'rbcSelCInclude(1).Value = True
                'If rbcSelCInclude(1).Value Then
                '    rbcSelCInclude_Click 1            'default to slsp
                'End If
                 mSPersonPop lbcSelection(2)
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
                rbcSelCInclude(4).Enabled = False       'vehicle/participant
                rbcSelCInclude(3).Enabled = False       'owner
            End If
                            
        Case CNT_BOBRECAP               '4-14-05
            slAirOrder = tgSpf.sInvAirOrder         'bill as ordred, aired
            mAdvtPop lbcSelection(5)
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            'mAskBobYrQtrPeriods                'ask quarter, year, and # periods
            mAskYrMonthPeriods ilListIndex
            If rbcSelCInclude(2).Value Then             'default to advt
                rbcSelCInclude_Click 2
            Else
                rbcSelCInclude(2).Value = True
            End If
    
            '3-2-02 Selective contract #
            lacSelCTo.Visible = True
            lacSelCTo.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 60, 840
            lacSelCTo.Caption = "Contr #"
            edcSelCTo.MaxLength = 9         '1-30-06
            edcSelCTo.Move 840, edcSelCFrom.Top + edcSelCFrom.Height + 30, 840
            edcSelCTo.Visible = True
    
            smPaintCaption4 = "Totals by"
            'plcSelC4_Paint
            plcSelC4.Move 120, edcSelCTo.Top + edcSelCTo.Height + 60
            rbcSelC4(0).Caption = "Detail"
            rbcSelC4(0).Left = 900
            rbcSelC4(0).Width = 960
            rbcSelC4(0).Visible = True
    
            'rbcSelC4(1).Caption = "Advertiser"
            'rbcSelC4(1).Left = 1920
            'rbcSelC4(1).Width = 1200
            rbcSelC4(1).Visible = False
            rbcSelC4(2).Caption = "Summary"
            rbcSelC4(2).Left = 1920
            rbcSelC4(2).Width = 1200
            rbcSelC4(2).Value = True
            If rbcSelC4(2).Value Then             'default to advt
                rbcSelC4_click 2
            Else
                rbcSelC4(2).Value = True
            End If
            rbcSelC4(2).Visible = True
            plcSelC4.Visible = True
    
            plcSelC9.Move 120, plcSelC4.Top + plcSelC4.Height, 2600
            mAskBOBCorpOrStd
    
            rbcSelC9(2).Visible = False     'disallow report by calendar
            plcSelC7.Move 120, plcSelC9.Top + plcSelC9.Height, 2560
            plcSelC7.Visible = True
            smPaintCaption7 = ""
            plcSelC7_Paint
            rbcSelC7(0).Move 0, 0, 825    'gross button,
            rbcSelC7(0).Caption = "Gross"
            rbcSelC7(1).Move 820, 0, 660   'net button
            rbcSelC7(1).Caption = "Net"
            rbcSelC7(1).Value = True
            rbcSelC7(2).Move 1420, 0, 1020
            rbcSelC7(2).Caption = "T-Net"       '"Net-Net"
            rbcSelC7(2).Visible = True
            rbcSelC7(0).Value = True        'default to gross
    
            'dont show the types, default to include everything except psa/prmos
            plcSelC3.Move 120, plcSelC7.Top + plcSelC7.Height
            'mAskContractTypes
    
            'dont show which type of lines to use, default to airing lines
            plcSelC1.Move 120, plcSelC6.Top + plcSelC6.Height
            'mAskPkgOrHide ilListIndex
            '
            rbcSelC11(0).Caption = "Vehicle"
            rbcSelC11(1).Caption = "Sales Origin"
            rbcSelC11(0).Visible = True
            rbcSelC11(1).Visible = True
            rbcSelC11(0).Move 720, 0, 960
            rbcSelC11(0).Value = True
            rbcSelC11(1).Move 1680, 0, 1680
            plcSelC11.Move 120, plcSelC7.Top + plcSelC7.Height
            smPaintCaption11 = "Sort by-"
            plcSelC11.Visible = True
            plcSelC10.Move 120, plcSelC11.Top + plcSelC11.Height  'sort by vehicle vs sales origin
    
            ckcSelC10(0).Caption = "Skip to new page each vehicle"
            ckcSelC10(0).Move 0, 0, 3600
            ckcSelC10(0).Visible = True
    
            plcSelC10.Visible = True
    
            ckcSelC8(2).Value = vbUnchecked         'no skipping to new page
            ckcSelC8(0).Value = vbUnchecked         'leave as defaults
            ckcSelC8(1).Value = vbChecked           'leave as defaults
    
            If slAirOrder = "S" Then            'bill as ordered, update as ordered; don't ask adjustment qustions
                                                    'always ignore missed & count mgs
                ckcSelC8(0).Visible = False
                ckcSelC8(1).Visible = False
                ckcSelC8(0).Value = vbUnchecked 'False       'ignore missed
                ckcSelC8(1).Value = vbUnchecked 'False       'ignore mgs
            Else                                'as aired
                ckcSelC8(0).Visible = True
                ckcSelC8(1).Visible = True
            End If
    
            If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
                rbcSelC9(0).Enabled = False
                rbcSelC9(0).Value = False
                rbcSelC9(1).Value = True
            Else
                If rbcSelC9(0).Value Then
                    rbcSelC9_click 0
                Else
                    rbcSelC9(0).Value = True
                End If
            End If
    
            plcSelC3.Move 120, plcSelC10.Top + plcSelC10.Height + 30
            mAskContractTypes
    
            lbcSelection(6).Visible = True
            ckcAll.Caption = "All Vehicles"
            ckcAll.Visible = True
    End Select
    If ckcSelC8(0).Enabled = False Then ckcSelC8(0).Value = False
    If ckcSelC8(1).Enabled = False Then ckcSelC8(1).Value = False
End Sub

'
'                   mObtainStartEndDates - obtain Standard Start and
'                   end date of given quarter
'                   <input>  ilYear = year to process
'                            ilmonth = starting month #
'                            ilNoMonths = total months to calc end date
'                   <return> llStartDate - STd start date of period
'                            llEnd Date - Std end date of period
'                            ilRet = true = ok
Private Function mObtainStartEndDates(ilYear As Integer, ilMonth As Integer, ilNoMonths As Integer, llStdStart As Long, llStdEnd As Long) As Integer
    Dim slTemp As String
    Dim slDate As String
    mObtainStartEndDates = False
    slDate = Trim$(str$(ilMonth)) & "/15/" & Trim$(str$(ilYear))
    slDate = gObtainStartStd(slDate)
    Do While ilNoMonths <> 0
        slTemp = gObtainEndStd(slDate)
        slDate = gObtainStartStd(slTemp)
        llStdEnd = gDateValue(slTemp)
        llStdEnd = llStdEnd + 1
        slDate = Format$(llStdEnd, "m/d/yy")
        ilNoMonths = ilNoMonths - 1
    Loop
    llStdStart = gDateValue(slDate)
    llStdEnd = llStdEnd - 1
    mObtainStartEndDates = True
End Function

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
    'gInitStdAlone RptSelCt, slStr, ilTestSystem
    ''ilRet = gParseItem(slCommand, 3, "\", slStr)
    ''igRptCallType = Val(slStr)
    If igRptCallType <> GENERICBUTTON Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)
    Else
        igRptType = 0
    End If
    'If igStdAloneMode Then
    '    ' "Proposals/Contracts"            '0=proposal
    '    ' "Paperwork Summary"              '1=paperwork summary (contract summaries)
    '    ' "Spots by Advertiser"             3=spots by date & time
    '    ' "Business Booked by Contract", ilIndex  '4=projection (named changed to Business Booked)
    '    ' "Contract Recap", ilIndex            '5=contr recap
    '    ' "Spot Placements", ilIndex           '6=Spot placements
    '    ' "Spot Discrepancies", ilIndex        '7=spot discrepancies
    '    ' "MG's", ilIndex                      '8=makegood
    '    ' "Sales Spot Tracking", ilIndex       '9=sales spot traking
    '    ' "Commercial Changes", ilIndex        '10=coml changes
    '    ' "Contract History", ilIndex          '11 Contract history
    '    ' "Affiliate Spot Tracking", ilIndex   '12 affil spot traking
    '    ' "Spot Sales", ilIndex                '13=spot sales
    '    ' "Missed Spots", ilIndex              '14=missed spots
    '    ' "Business Booked by Spot", ilIndex    '15=spot projection (name changed to Business Booked)
    '    ' "Business Booked by Spot Reprint", ilIndex   '16= Business booked reprint
    '    '                                       'reprint unused, screen code removed due to lack of memory 2-28-01
    '    ' "Spot Business Booked"    'changed from Business Booked by Spot 1-31-00
    '    ' "Avails", ilIndex                    '17=quarterly summary & detail avails
    '    ' "Average Spot Prices", ilIndex       '18=avg spot prices
    '    ' "Advertiser Units Ordered", ilIndex  '19=advt units ordered
    '    ' "Sales Analysis by CPP & CPM", ilIndex '20=sales analysis by cpp & cpm
    '    ' "Average Rate", ilIndex            '21=Average Rate
    '    ' "Tie-Out", ilIndex                  '22=Detail Tie Out
    '    ' "Billed and Booked", ilIndex        '23=Billed & booked by advt, Slsp, owner, vehicle
    '    ' "Weekly Sales Activity by Quarter", ilIndex   '24=Sales Activity
    '    ' "Sales Comparison", ilIndex         'Sales Comparison by Advt, Slsp, Agy, comp code, Bus code
    '    ' "Weekly Sales Activity by Month", ilIndex       'Cumulative Activity Report (pacing)
    '    ' "Average Prices to Make Plan", ilIndex       'Avg Prices needed to make plan
    '    ' "CPP/CPM by Vehicle", ilIndex        'Curent cpp/cpm by vehicle
    '    ' "Sales Analysis Summary", ilIndex        'Sales Analysis Summary
    '    ' "Insertion Orders"
    '    ' "Daily Sales Activity by Contract"        '6-5-01
    '
    '    smSelectedRptName = "Business Booked by Contract"  '"Spot Business Booked" '"Billed and Booked"
    '    igRptCallType = CONTRACTSJOB 'LOGSJOB 'CONTRACTSJOB 'COPYJOB 'COLLECTIONSJOB'CONTRACTSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
    '    igRptType = 1   '3 'Log '1   'Contract    '0   'Summary '3 Program  '1  links
    '    slCommand = "x\x\x\x\2\2/6/95\7\12M\12M\1\26" '"" '"CONT0802.ASC\11/20/94\10:11:0 AM" '"x\x\x\x\2"
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)
        End If
    'End If
    If (igRptCallType = CONTRACTSJOB) And (igRptType = 3) Then
        igStdAloneMode = True 'Switch from/to stand alone mode-No DDE
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
'*      Procedure Name:mSalesOfficePop                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSalesOfficePop(lbcSelection As Control)
    Dim ilRet As Integer
    'ilRet = gPopOfficeSourceBox(RptSelCt, lbcSelection, lbcSOCode)
    ilRet = gPopOfficeSourceBox(RptSelCt, lbcSelection, tgSOCodeCT(), sgSOCodeTagCT)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSalesOfficePopErr
        gCPErrorMsg ilRet, "mSalesOfficePop (gPopOfficeSourceBox: Vehicle)", RptSelCt
        On Error GoTo 0
    End If
    Exit Sub
mSalesOfficePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    'TTP 10350 - Advertiser Units Ordered Report Rep Vehicles not shown in vehicle list on the report screen
    ilRet = gPopUserVehicleBox(RptSelCt, VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVehPopErr
        gCPErrorMsg ilRet, "mSellConvVehPop (gPopUserVehicleBox: Vehicle)", RptSelCt
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
Private Sub mSellConvVirtVehPop(ilIndex As Integer, ilUselbcVehicle As Integer, Optional blInsertionsOnly As Boolean = False)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        If blInsertionsOnly Then
            ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_W_CLUSTER + VEHREP_WO_CLUSTER + VEHNTR + ACTIVEVEH + VEHONINSERTION, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
        Else
            ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_W_CLUSTER + VEHREP_WO_CLUSTER + VEHNTR + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
        End If
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelCt
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
'*      Procedure Name:mSellConvVVActPop               *
'*                                                     *
'*             Created:5/5/98        By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box for Active only selling veh*
'*                      virt veh, conve veh            *
'*                                                     *
'*******************************************************
Private Sub mSellConvVVActPop(ilIndex As Integer, ilUselbcVehicle As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)    'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVVActPopErr
        gCPErrorMsg ilRet, "mSellConvVVActPopErr (gPopUserVehicleBox: Vehicle)", RptSelCt
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVVActPopErr:
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
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), Traffic!lbcVehicle)
        ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
        'ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + ACTIVEVEH, lbcSelection(ilIndex), lbcCSVNameCode)    'lbcCSVNameCode)
        ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHPACKAGE + VEHREP_W_CLUSTER + VEHREP_WO_CLUSTER + VEHNTR + ACTIVEVEH + DORMANTVEH, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag) 'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVVPkgPopErr
        gCPErrorMsg ilRet, "mSellConvVVPkgPop (gPopUserVehicleBox: Vehicle)", RptSelCt
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
    Dim ilEnable2 As Integer
    Dim illoop As Integer
    Dim ilListIndex As Integer
    Dim ilIndex As Integer
    Dim Vehicle As Integer
    Dim RateCard As Integer
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = SLSPCOMMSJOB Then       'Commissions
        If ilListIndex = COMM_SALESCOMM Then    'Or ilListIndex = COMM_PROJECTION Then
        'If edcSelCFrom.Text <> "" And (edcSelCFrom1.Text <> "") Then        'dates have been entered
            If (edcSelCFrom.Text <> "") Then                     'check if 1st field for date entered
                ilEnable = False
                If (ckcAll.Value = vbChecked Or lbcSelection(2).SelCount > 0) And (ckcAllAAS.Value = vbChecked Or lbcSelection(6).SelCount > 0) Then                    '9-12-02 check for the correct selection list box here
                    ilEnable = True
'                Else
'                    For ilLoop = 0 To lbcSelection(2).ListCount - 1 Step 1
'                        If lbcSelection(2).Selected(ilLoop) Then
'                            ilEnable = True
'                            Exit For
'                        End If
'                    Next ilLoop
                End If
                If ilListIndex = COMM_SALESCOMM Then
                    If (edcSelCFrom1.Text = "") Then           'Sales commission also has month input field to check
                        ilEnable = False
                    End If
                End If
            Else
                ilEnable = False                    'no year entered
            End If
        ElseIf ilListIndex = COMM_PROJECTION Then
            If edcSelCFrom.Text = "" Or edcSelCFrom1.Text = "" Then         'start qtr and year must be entered
                ilEnable = False
            Else
                If ckcAll.Value = vbChecked Then                   '9-12-02 check for the correct selection list box here
                    ilEnable = True
                Else
                    For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                        If lbcSelection(2).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
            End If
        End If
    ElseIf igRptCallType = CONTRACTSJOB Then
        If (igRptType = 0) And (ilListIndex > 1) Then
            ilListIndex = ilListIndex + 1
        End If
'        If rbcRptType(0).Value Or rbcRptType(1).Value Or rbcRptType(2).Value Then
        If (ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION) Then
            If ckcAllAAS.Value = vbChecked Then '9-12-02
                ilEnable = True
            Else
                'If rbcSelCSelect(0).Value Then                 '5-16-02 advt
                    If edcTopHowMany.Text = "" Then             '11-27-00 if selective cnt # entered, no need to indicate an advt
                        For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                            If lbcSelection(0).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    Else
                        ilEnable = True
                    End If
                'Else                                        'agy or slsp
                '    If edcTopHowMany.Text = "" Then             '11-27-00 if selective cnt # entered, no need to indicate an agy or slsp
                '        For ilLoop = 0 To lbcSelection(10).ListCount - 1 Step 1
                '            If lbcSelection(10).Selected(ilLoop) Then
                '                ilEnable = True
                '                Exit For
                '            End If
                '        Next ilLoop
                '    Else
                '        ilEnable = True
                 '   End If
                'End If
            End If
            
            '7/11/15: Only E-Mail content is mandatory.  Code placed at end of this routine
            'If ilListIndex = CNT_INSERTION Then
            '    If rbcOutput(3).Value = True Then       'if email option, the content and response date are mandatory
            '        If cbcEMailContent.ListIndex < 1 Or edcResponse.Text = "" Then
            '            ilEnable = True
            '        End If
            '    End If
            '
            'End If
        ElseIf (ilListIndex = CNT_PAPERWORK) Then
            If ckcAll.Value = vbChecked Then    '9-12-02
                ilEnable = True
            Else
                If ilListIndex = CNT_PAPERWORK Then                             'paperwork summary report
                    If rbcSelCSelect(0).Value Then                  'advt
                    For illoop = 0 To lbcSelection(5).ListCount - 1 Step 1
                        If lbcSelection(5).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                    End If
                Else
                    If rbcSelCSelect(0).Value Then                  'advt, get selective cnts
                        'Can't use SelCount as property does not exist for ListBoxbox
                        If ckcAll.Value = vbChecked Then    '9-12-02
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
                    If ilEnable And Not ckcAllAAS.Value = vbChecked Then    '9-12-02
                        ilEnable = False
                        'at least one vehicle must be selected
                        For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                            If lbcSelection(6).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    End If
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
            If ilListIndex = CNT_PAPERWORK Then
                'at least one status must be requested
                '3-16-10 wrong controls tested (ckcselc5 instead of ckcselc3)
                If ((Not ckcSelC3(0).Value = vbChecked) And (Not ckcSelC3(1).Value = vbChecked) And (Not ckcSelC3(2).Value = vbChecked) And (Not ckcSelC3(3).Value = vbChecked) And (Not ckcSelC3(4).Value = vbChecked) And (Not ckcSelC3(5).Value = vbChecked) And (Not ckcSelC3(6).Value = vbChecked)) Then       '11-7-16 check Rev (6) selection
                    ilEnable = False
                End If
                'at least one type must be requested
                '3-16-10 wrong controls tested (ckcselc3 instead of ckcselc5)
                If ((Not ckcSelC5(0).Value = vbChecked) And (Not ckcSelC5(1).Value = vbChecked) And (Not ckcSelC5(2).Value = vbChecked) And (Not ckcSelC5(3).Value = vbChecked) And (Not ckcSelC5(4).Value = vbChecked) And (Not ckcSelC5(5).Value = vbChecked) And (Not ckcSelC5(6).Value = vbChecked) And (Not ckcSelC5(7).Value = vbChecked)) Then
                    ilEnable = False
                End If
            End If

        ElseIf (ilListIndex = CNT_BOB_BYCNT) Then
            'Date: 12/16/2019 added CSI calendar control for date entry
            If (CSI_CalFrom.Text <> "") And (rbcSelCSelect(0).Value Or rbcSelCSelect(1).Value Or rbcSelCSelect(2).Value Or rbcSelCSelect(3).Value) Then
                ilEnable = False
                If rbcSelCInclude(0).Value Then
                    'Can't use SelCount as property does not exist for ListBoxbox
                    If ckcAll.Value = vbChecked Then    '9-12-02
                        ilEnable = True
                    Else
                        For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                            If lbcSelection(0).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    End If
                ElseIf rbcSelCInclude(1).Value Then
                    For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                        If lbcSelection(2).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                Else
                    For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                        If lbcSelection(6).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
            Else
                ilEnable = False
            End If
        ElseIf (ilListIndex = CNT_BOB_BYSPOT) Then 'Projection
            'Date: 1/7/2020 added CSI calendar controls for date entry
            'If (edcSelCFrom.Text <> "") And (rbcSelCSelect(0).Value Or rbcSelCSelect(1).Value Or rbcSelCSelect(2).Value Or rbcSelCSelect(3).Value) Then
            If (CSI_CalFrom.Text <> "") And (rbcSelCSelect(0).Value Or rbcSelCSelect(1).Value Or rbcSelCSelect(2).Value Or rbcSelCSelect(3).Value) Then
                ilEnable = False
                If rbcSelCInclude(0).Value Then
                    'Can't use SelCount as property does not exist for ListBoxbox
                    If ckcAll.Value = vbChecked Then    '9-12-02
                        ilEnable = True
                    Else
                        For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                            If lbcSelection(0).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    End If
                ElseIf rbcSelCInclude(1).Value Then
                    For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                        If lbcSelection(2).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                Else
'                    If ilListIndex = CNT_BOB_BYCNT Then
'                        For ilLoop = 0 To lbcSelection(6).ListCount - 1 Step 1
'                            If lbcSelection(6).Selected(ilLoop) Then
'                                ilEnable = True
'                                Exit For
'                            End If
'                        Next ilLoop
'                    Else                'Spot business Booked
                        If rbcSelCInclude(2).Value Then
                            ilIndex = 3             'vehicle
                        Else
                            ilIndex = 1             'agency
                        End If
                        For illoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                            If lbcSelection(ilIndex).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop

'                    End If
                End If
            Else
                ilEnable = False
            End If
        ElseIf (ilListIndex = CNT_BOB_BYSPOT_REPRINT) Then 'Projection Reprint
            'no longer functional
        ElseIf ilListIndex = CNT_RECAP Then 'Recap
            ilEnable = True

        ElseIf ilListIndex = CNT_MG Then 'MG's
            If ckcAll.Value = vbChecked Then    '9-12-02
                ilEnable = True
            Else
                For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                    If lbcSelection(6).Selected(illoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next illoop
            End If
            If Not RptSelCt!ckcSelC3(0).Value = vbChecked And Not RptSelCt!ckcSelC3(1).Value = vbChecked Then
                ilEnable = False
            End If
            'ilEnable = True
        ElseIf ilListIndex = CNT_SPOTTRAK Then  'Sales Spot Tracking
            ilEnable = True
        ElseIf ilListIndex = CNT_COMLCHG Then 'Commercial changes
            ilEnable = True
        ElseIf ilListIndex = CNT_HISTORY Then 'History
            If ckcAll.Value = vbChecked Or edcTopHowMany.Text <> "" Then
                ilEnable = True
            Else
                For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1
                    If lbcSelection(0).Selected(illoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next illoop
            End If

        ElseIf ilListIndex = CNT_AFFILTRAK Then  'Affiliate Spot Tracking
            ilEnable = True

        ElseIf (ilListIndex = CNT_QTRLY_AVAILS) Then
            If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") And (rbcSelCSelect(0).Value Or rbcSelCSelect(1).Value Or rbcSelCSelect(2).Value Or rbcSelCSelect(3).Value) And (rbcSelCInclude(0).Value Or rbcSelCInclude(1).Value Or rbcSelCInclude(2).Value) Then
                ilEnable = False
                For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                    If lbcSelection(6).Selected(illoop) Then
                        Vehicle = True
                        Exit For
                    End If
                Next illoop
                'If ckcAllAAS.Value = True Then
                '    ilenable = True
                'Else
                For illoop = 0 To lbcSelection(12).ListCount - 1 Step 1
                    If lbcSelection(12).Selected(illoop) Then
                        igRCSelectedIndex = illoop
                        RateCard = True
                        Exit For
                    End If
                Next illoop
                'End If
                If Vehicle And RateCard = True Then
                    ilEnable = True
                End If
            End If
            'Else
            '    ilenable = False
            'End If
'        ElseIf (ilListIndex = CNT_AVG_PRICES) Or (ilListIndex = CNT_ADVT_UNITS) Then    'avg spot price, Advertiser Units ordered
         ElseIf (ilListIndex = CNT_AVG_PRICES) Then    '6-21-18 remove adv units
            
            'Date: 9/22/2019 using CSI calendar for date entry; added Major/Minow sorts
            If CSI_CalFrom.Text <> "" Then
                ilEnable = False
                If ckcAll.Value = vbChecked Then '9-12-02
                    ilEnable = True
                Else
                    If rbcSelCInclude(0).Value Then         'advt
                        ilIndex = 5
                    ElseIf rbcSelCInclude(1).Value Then     'slsp
                        ilIndex = 2
                    ElseIf rbcSelCInclude(2).Value Then     'agency
                        ilIndex = 1
                    ElseIf rbcSelCInclude(3).Value Then     'bus cat
                        ilIndex = 3
                    ElseIf rbcSelCInclude(4).Value Then     'prod prot
                        ilIndex = 7
                    '4-25-06 vehicle option added
                    ElseIf rbcSelCInclude(5).Value Then     'vehicle
                        ilIndex = 6
                    Else                                    'vehicle grp
                        'ilIndex = 4
                        ilIndex = 12                        '3-18-16 single selection vg
                    End If
                    For illoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                        If lbcSelection(ilIndex).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                    If ilIndex = 12 Then                'vg group, must have selected the items as well
                        If lbcSelection(8).SelCount > 0 Then
                            ilEnable = True
                        Else
                            ilEnable = False
                        End If
                    End If
                End If
            Else
                ilEnable = False
            End If

        ElseIf (ilListIndex = CNT_ADVT_UNITS) Then     ' Advertiser Units ordered
            If CSI_CalFrom.Text <> "" Then  'Date: 9/22/2019 using CSI calendar for date entry  --> (edcSelCFrom.Text <> "") Then
                ilEnable = False
                If ckcAll.Value = vbChecked Then '9-12-02
                    ilEnable = True
                Else
                    If rbcSelCInclude(0).Value Then         'advt
                        ilIndex = 5
                    ElseIf rbcSelCInclude(1).Value Then     'slsp
                        ilIndex = 2
                    ElseIf rbcSelCInclude(2).Value Then     'agency
                        ilIndex = 1
                    ElseIf rbcSelCInclude(3).Value Then     'bus cat
                        ilIndex = 3
                    ElseIf rbcSelCInclude(4).Value Then     'prod prot
                        ilIndex = 7
                    '4-25-06 vehicle option added
                    ElseIf rbcSelCInclude(5).Value Then     'vehicle
                        ilIndex = 6
                    Else                                    'vehicle grp
                        'ilIndex = 4
                        ilIndex = 12                        '3-18-16 single selection vg
                    End If
                    For illoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                        If lbcSelection(ilIndex).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                    If ilIndex = 12 Then                'vg group, must have selected the items as well
                        If lbcSelection(8).SelCount > 0 Then
                            ilEnable = True
                        Else
                            ilEnable = False
                        End If
                    End If
                End If
                If ilEnable = True Then                         'test the subsort for at least 1 seleced
                    ilEnable = False
                    If ckcAllAAS.Value = vbChecked Or cbcSet2.ListIndex = 0 Then
                        ilEnable = True
                    Else
                        If cbcSet2.ListIndex = 1 Then         'advt
                            ilIndex = 5
                        ElseIf cbcSet2.ListIndex = 2 Then     'agy
                            ilIndex = 1
                        ElseIf cbcSet2.ListIndex = 3 Then     'bus cat
                            ilIndex = 3
                        ElseIf cbcSet2.ListIndex = 4 Then     'prod prot
                            ilIndex = 7
                        ElseIf cbcSet2.ListIndex = 5 Then             'slsp
                            ilIndex = 2
                        ElseIf cbcSet2.ListIndex = 6 Then             'vehicle
                            ilIndex = 6
                        ElseIf cbcSet2.ListIndex = 7 Then               'vehicle grp
                            'ilIndex = 4
                            ilIndex = 12                            '3-18-16 single selection vg
                        End If
                        For illoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                            If lbcSelection(ilIndex).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    End If
                End If
            Else
                ilEnable = False
            End If
        ElseIf ilListIndex = CNT_SALES_CPPCPM Then  'Sales Analysis by CPP/CPM
            'Date: 12/10/2019 added CSI calendar control for date entry
            'If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
            If (CSI_CalFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
                If ckcAll.Value = vbChecked Then    '9-12-02
                    ilEnable = True
                Else
                    For illoop = 0 To lbcSelection(11).ListCount - 1 Step 1
                        If lbcSelection(11).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
            End If
        ElseIf ilListIndex = CNT_AVGRATE Then      'average rate
            If ckcAll.Value = vbChecked Then    '9-12-02
                ilEnable = True
            Else
                For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                    If lbcSelection(6).Selected(illoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next illoop
            End If
            If (edcSelCTo.Text = "" Or edcSelCTo1.Text = "") Then
                ilEnable = False
            End If
        ElseIf ilListIndex = CNT_TIEOUT Then                 'tie out
            If CSI_CalFrom.Text = "" Then       'Date: 1/8/2020 added CSI calendar control for date entry --> edcSelCFrom.Text = "" Then
                ilEnable = False
            Else
                If ckcAll.Value = vbChecked Then                   ' 9-12-02 check for the correct selection list box here
                    ilEnable = True
                Else
                    If rbcSelCInclude(0).Value Then      'office
                        For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                            If lbcSelection(2).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    Else                         ' vehicle option
                        For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                            If lbcSelection(6).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    End If
                End If
                If (ilEnable) Then                      'vehicle or office selected, check on budget names
                    ilEnable = False                    'reset to test budget name
                    For illoop = 0 To lbcSelection(4).ListCount - 1 Step 1
                        If lbcSelection(4).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop

                    'If (ilEnable) And rptselct!rbcSelCInclude(0).Value Then   'continue need splits answered if by office option
                    '    ilEnable = False
                    '    If rbcSelCInclude(0).Value Then         'office option, need split plan
                    '        For ilLoop = 0 To lbcSelection(12).ListCount - 1 Step 1
                    '            If lbcSelection(12).Selected(ilLoop) Then
                    '                ilEnable = True
                    '                Exit For
                    '            End If
                    '        Next ilLoop
                    '    End If
                    'End If
                End If
            End If
        ElseIf ilListIndex = CNT_BOB Then
            '3-15-05 at least 1 contract type (std, resv, remnant, etc) must be set, dates must be entered, and must indicate air time and/or ntr
            If (ckcSelC5(0).Value = vbUnchecked And ckcSelC5(1).Value = vbUnchecked And ckcSelC5(2).Value = vbUnchecked And ckcSelC5(3).Value = vbUnchecked And ckcSelC5(4).Value = vbUnchecked And ckcSelC5(5).Value = vbUnchecked And ckcSelC5(6).Value = vbUnchecked) Or (edcSelCFrom.Text = "" Or edcSelCFrom1.Text = "") Or (ckcSelC6(1).Value = vbUnchecked And ckcSelC6(2).Value = vbUnchecked And ckcSelC6(3).Value = vbUnchecked) Then
                ilEnable = False
            'End If
            'If edcSelCFrom.Text = "" Or edcSelCFrom1.Text = "" Then         'start qtr and year must be entered
                ilEnable = False
            Else
                If ckcAll.Value = vbChecked Then                   ' 9-12-02 check for the correct selection list box here
                    ilEnable = True

                    If rbcSelCInclude(1).Value And ckcSelC10(1).Value = vbChecked Then    'slsp option with vehicle sub-totals
                        ilEnable = False
                        If ckcAllAAS.Value = vbChecked Then '9-12-02
                            ilEnable = True
                        Else
                            ilEnable = False
                            For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                                If lbcSelection(6).Selected(illoop) Then
                                    ilEnable = True
                                    Exit For
                                End If
                            Next illoop
                        End If
                    ElseIf rbcSelCInclude(4).Value Then     '8-4-00 vehicle with participant splits
                        ilEnable = False
                        If ckcAllAAS.Value = vbChecked Then '9-12-02
                            ilEnable = True
                        Else
                            ilEnable = False
                            For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                                If lbcSelection(2).Selected(illoop) Then
                                    ilEnable = True
                                    Exit For
                                End If
                            Next illoop
                        End If
                    End If
                Else
                    If rbcSelCInclude(0).Value Then      'advt
                        For illoop = 0 To lbcSelection(5).ListCount - 1 Step 1
                            If lbcSelection(5).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    ElseIf rbcSelCInclude(2).Value Or rbcSelCInclude(6).Value Then                        'vehicle or vehicle net-net
                        For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                            If lbcSelection(6).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    ElseIf rbcSelCInclude(5).Value Then                       'agency
                        For illoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                            If lbcSelection(1).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    ElseIf rbcSelCInclude(4).Value Then         '11-21-06  vehicle/participant
                        For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                            If lbcSelection(6).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                        If ckcAllAAS.Value = vbChecked And ilEnable Then    '9-12-02
                            ilEnable2 = True
                        Else
                            ilEnable2 = False
                            For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                                If lbcSelection(2).Selected(illoop) Then
                                    ilEnable2 = True
                                    Exit For
                                End If
                            Next illoop
                        End If
                        If Not ilEnable2 Then
                            ilEnable = False
                        End If
                    Else                         'slsp or owner
                        For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1
                            If lbcSelection(2).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                        If rbcSelCInclude(1).Value And ckcSelC10(1).Value = vbChecked Then    'slsp option with vehicle sub-totals
                            If ckcAllAAS.Value = vbChecked And ilEnable Then    '9-12-02
                                ilEnable2 = True
                            Else
                                ilEnable2 = False
                                If lbcSelection(6).SelCount > 0 Then
                                'For ilLoop = 0 To lbcSelection(6).ListCount - 1 Step 1
                                '    If lbcSelection(6).Selected(ilLoop) Then
                                        ilEnable2 = True
                                 '       Exit For
                                 '   End If
                                'Next ilLoop
                                End If
                            End If
                            If Not ilEnable2 Then
                                ilEnable = False
                            End If
                        End If
                    End If
                End If
                 'vehicle, vehicle gross/net or vehicle/participant options, everything tested so far has passed, and a vehicle group  has been selected
                If (rbcSelCInclude(2).Value Or rbcSelCInclude(4).Value Or rbcSelCInclude(6).Value) And ilEnable = True And cbcSet1.ListIndex > 0 Then
                    If lbcSelection(7).SelCount <= 0 Then
                        ilEnable = False
                    End If
                End If
            End If
        ElseIf ilListIndex = CNT_BOBRECAP Or ilListIndex = CNT_PAPERWORKTAX Then               '4-14-05
            ilEnable = True
            'Date: 12/20/2019 added CSI calendar control for date entries
            If ilListIndex = CNT_PAPERWORKTAX Then
                If (CSI_CalFrom.Text = "" Or CSI_CalTo.Text = "") Then
                    ilEnable = False
                End If
            Else
                If (edcSelCFrom.Text = "" Or edcSelCFrom1.Text = "") Then
                    ilEnable = False
                End If
            End If
            If ilListIndex = CNT_BOBRECAP Then
                If lbcSelection(6).SelCount <= 0 Then
                    ilEnable = False
                End If
            Else                'paperwork tax summary must have at least 1 vehicle selected
                If lbcSelection(11).SelCount <= 0 Then
                    ilEnable = False
                End If
            End If

         ElseIf ilListIndex = CNT_SALESACTIVITY Then
            ilEnable = False
            'Date: 1/8/2020 added CSI calendar control for date entries
            'If edcSelCFrom.Text <> "" And edcSelCTo1.Text <> "" And edcSelCTo.Text <> "" Then
            If CSI_CalFrom.Text <> "" And edcSelCTo1.Text <> "" And edcSelCTo.Text <> "" Then
                ilEnable = True
            End If
            If ckcSelC13(0).Value = vbUnchecked And ckcSelC13(1).Value = vbUnchecked And ckcSelC13(2).Value = vbUnchecked Then
                ilEnable = False
            End If
         ElseIf ilListIndex = CNT_SALESCOMPARE Then
            If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") And (edcSelCTo.Text <> "") Then
                'ilEnable = False
                If ckcAll.Value = vbChecked Then '9-12-02
                    ilEnable = True
                Else
                    If rbcSelCInclude(0).Value Then         'advt
                        ilIndex = 5
                    ElseIf rbcSelCInclude(1).Value Then     'slsp
                        ilIndex = 2
                    ElseIf rbcSelCInclude(2).Value Then     'agency
                        ilIndex = 1
                    ElseIf rbcSelCInclude(3).Value Then     'bus cat
                        ilIndex = 3
                    ElseIf rbcSelCInclude(4).Value Then             'prod prot
                        ilIndex = 7
                    '4-25-06 vehicle option added
                    ElseIf rbcSelCInclude(5).Value Then             'vehicle
                        ilIndex = 6
                    Else                                        'vehicle grp
                        'ilIndex = 4
                        ilIndex = 12                           '3-18-16 single selection vg
                    End If
                    For illoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                        If lbcSelection(ilIndex).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                    If ilIndex = 12 Then                'vg group, must have selected the items as well
                        If lbcSelection(8).SelCount > 0 Then
                            ilEnable = True
                        Else
                            ilEnable = False
                        End If
                    End If
                End If
                If ilEnable = True Then                         'test the subsort for at least 1 seleced
                    ilEnable = False
                    If ckcAllAAS.Value = vbChecked Or cbcSet2.ListIndex = 0 Then
                        ilEnable = True
                    Else
                        If cbcSet2.ListIndex = 1 Then         'advt
                            ilIndex = 5
                        ElseIf cbcSet2.ListIndex = 2 Then     'agy
                            ilIndex = 1
                        ElseIf cbcSet2.ListIndex = 3 Then     'bus cat
                            ilIndex = 3
                        ElseIf cbcSet2.ListIndex = 4 Then     'prod prot
                            ilIndex = 7
                        ElseIf cbcSet2.ListIndex = 5 Then             'slsp
                            ilIndex = 2
                        ElseIf cbcSet2.ListIndex = 6 Then             'vehicle
                            ilIndex = 6
                        ElseIf cbcSet2.ListIndex = 7 Then               'vehicle grp
                            'ilIndex = 4
                            ilIndex = 12                            '3-18-16 single selection vg
                        End If
                        For illoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                            If lbcSelection(ilIndex).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    End If
                End If
            Else
                ilEnable = False
            End If
        ElseIf ilListIndex = CNT_CUMEACTIVITY Then
            'Date: 1/8/2020 added CSI calendar control for date entries
            'If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
            If (CSI_CalFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
                If ckcAll.Value = vbChecked Then    '9-12-02
                    ilEnable = True
                Else
                    If rbcSelCInclude(0).Value Then         'advt
                        ilIndex = 5
                    ElseIf rbcSelCInclude(1).Value Then     'agy
                        ilIndex = 1
                    ElseIf rbcSelCInclude(2).Value Then             'demo
                        ilIndex = 11
                    Else                                    'vehicle
                        ilIndex = 6
                    End If
                    For illoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                        If lbcSelection(ilIndex).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
            Else
                ilEnable = False
            End If
        ElseIf ilListIndex = CNT_MAKEPLAN Then
            If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") And (edcSelCTo.Text <> "") Then
                If ckcAll.Value = vbChecked Then    '9-12-02
                    ilEnable = True
                Else
                    For illoop = 0 To lbcSelection(3).ListCount - 1 Step 1      'vehicle entry must be selected
                        If lbcSelection(3).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
                If ilEnable Then                            'rate card entry must be selected
                    ilEnable = False
                    If RptSelCt!rbcSelCSelect(0).Value Then     'corp, more than 1 r/c is necessary
                        ilIndex = 11
                    Else
                        ilIndex = 12
                    End If
                    If lbcSelection(ilIndex).SelCount > 0 Then
                        ilEnable = True
                    End If
                End If
                If ilEnable Then                            'budget entry must be selected
                    ilEnable = False
                    For illoop = 0 To lbcSelection(4).ListCount - 1 Step 1
                        If lbcSelection(4).Selected(illoop) Then
                            igBSelectedIndex = illoop
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
            End If
        ElseIf ilListIndex = CNT_VEHCPPCPM Then
            If (CSI_CalFrom.Text <> "") Then    'Date: 12/12/2019 added CSI calendar control for date entry --> (edcSelCFrom.Text <> "")
                If ckcAll.Value = vbChecked Then    '9-12-02
                    ilEnable = True
                Else
                    For illoop = 0 To lbcSelection(3).ListCount - 1 Step 1      'vehicle entry must be selected
                        If lbcSelection(3).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
                If ilEnable Then
                    ilEnable = False
                    If ckcAllAAS.Value = vbChecked Then '9-12-02
                        ilEnable = True
                    Else
                        For illoop = 0 To lbcSelection(2).ListCount - 1 Step 1      'vehicle entry must be selected
                            If lbcSelection(2).Selected(illoop) Then
                                ilEnable = True
                                Exit For
                            End If
                        Next illoop
                    End If
                End If
                If ilEnable Then
                    ilEnable = False
                    For illoop = 0 To lbcSelection(4).ListCount - 1 Step 1      'vehicle entry must be selected
                        If lbcSelection(4).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
            End If
        ElseIf ilListIndex = CNT_SALESANALYSIS Then
            'Date: 12/11/2019 added CSI calendar control for date entry
            'If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
            If (CSI_CalFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
                'atleast one budget must be selected
                For illoop = 0 To lbcSelection(4).ListCount - 1 Step 1      'budget entry must be selected
                    If lbcSelection(4).Selected(illoop) Then
                        ilEnable = True
                        igBSelectedIndex = illoop                'index of budget selected
                        Exit For
                    End If
                Next illoop
            End If
        ElseIf ilListIndex = CNT_DAILY_SALESACTIVITY Then
            ilEnable = False
            'Date: 11/21/2019 using CSI calendar control for date entries -->
            'If edcSelCTo.Text <> "" And edcSelCTo1.Text <> "" Then
            If CSI_CalFrom.Text <> "" And CSI_CalTo.Text <> "" Then
                ilEnable = True
            End If
            If ckcSelC13(0).Value = vbUnchecked And ckcSelC13(1).Value = vbUnchecked And ckcSelC13(2).Value = vbUnchecked Then
                ilEnable = False
            End If
        ElseIf ilListIndex = CNT_SALESACTIVITY_SS Then      '7-25-02
            ilEnable = False
            'Date: 11/26/2019   added CSI calendar controls for date entries
            'If edcSelCTo.Text <> "" And edcSelCTo1.Text <> "" And edcSelCFrom.Text <> "" And edcSelCFrom1.Text <> "" And edcText.Text <> "" Then
            If edcSelCTo.Text <> "" And edcSelCTo1.Text <> "" And CSI_CalFrom.Text <> "" And CSI_CalTo.Text <> "" And edcText.Text <> "" Then
                'only if everything is set so far, test the list boxes
                If (lbcSelection(6).SelCount = 0 And lbcSelection(6).Visible = True) Or lbcSelection(2).SelCount = 0 Or lbcSelection(3).SelCount = 0 Then
                    ilEnable = False
                Else
                    ilEnable = True
                End If

            End If
        ElseIf ilListIndex = CNT_SALESPLACEMENT Then       '8-14-02
            ilEnable = False
            If edcSelCTo.Text <> "" And edcSelCTo1.Text <> "" And edcText.Text <> "" Then
                'only if everything is set so far, test the list boxes
                If (lbcSelection(6).SelCount = 0 And lbcSelection(6).Visible = True) Or lbcSelection(2).SelCount = 0 Or lbcSelection(3).SelCount = 0 Then
                    ilEnable = False
                Else
                    ilEnable = True
                End If

            End If
        ElseIf ilListIndex = CNT_VEH_UNITCOUNT Or ilListIndex = CNT_LOCKED Or ilListIndex = CNT_GAMESUMMARY Then
            ilEnable = False
            'Date: 12/4/2019 using CSI calendar control for date entries -->
'            If (ilListIndex = CNT_GAMESUMMARY) Or (ilListIndex = CNT_LOCKED) Then
'                If CSI_CalFrom.Text <> "" And CSI_CalTo.Text <> "" Then
'                    'only if everything is set so far, test the list boxes
'                    If lbcSelection(3).SelCount = 0 Then
'                        ilEnable = False
'                    Else
'                        ilEnable = True
'                    End If
'                End If
'            Else
                'Date: 1/8/2020 added CSI calendar control for date entries
                If CSI_CalFrom.Text <> "" And CSI_CalTo.Text <> "" Then         'edcSelCFrom.Text <> "" And edcSelCFrom1.Text <> "" Then
                    'only if everything is set so far, test the list boxes
                    If lbcSelection(3).SelCount = 0 Then
                        ilEnable = False
                    Else
                        ilEnable = True
                    End If
                End If
'            End If
        ElseIf ilListIndex = CNT_BOBCOMPARE Then            '9-13-07
            If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") And (edcSelCTo.Text <> "") Then
                If ckcAll.Value = vbChecked Then '9-12-02
                    ilEnable = True
                Else
                    If cbcSet2.ListIndex = 0 Then         'advt
                        ilIndex = 5
                    ElseIf cbcSet2.ListIndex = 4 Then     'slsp
                        ilIndex = 2
                    ElseIf cbcSet2.ListIndex = 1 Then     'agency
                        ilIndex = 1
                    ElseIf cbcSet2.ListIndex = 2 Then     'bus cat
                        ilIndex = 3
                    ElseIf cbcSet2.ListIndex = 3 Then             'prod prot
                        ilIndex = 7
                    ElseIf cbcSet2.ListIndex = 5 Then             'vehicle
                        ilIndex = 6
                    End If
                    For illoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1
                        If lbcSelection(ilIndex).Selected(illoop) Then
                            ilEnable = True
                            Exit For
                        End If
                    Next illoop
                End If
                '2-23-27 if a vehicle group is selected, then budgets allowed; but only 1 can be selected
                If lbcSelection(4).SelCount > 1 And cbcSet1.ListIndex > 0 Then
                    MsgBox "Only 1 budget selection allowed", vbOKOnly, "Billed and Booked Comparisons"
                    ilEnable = False
                End If

            Else
                ilEnable = False
            End If
        ElseIf ilListIndex = CNT_CONTRACTVERIFY Then         '4-8-13
            If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
                ilEnable = True
            End If
        ElseIf ilListIndex = CNT_INSERTION_ACTIVITY Then        '10-6-15
            'Date: 12/18/2019 added CSI calendar control for date entry
            'If edcSelCFrom.Text = "" Or edcSelCFrom1.Text = "" Or edcSelCTo.Text = "" Or edcSelCTo1.Text = "" Then
            If CSI_CalFrom.Text = "" Or CSI_From1.Text = "" Or CSI_CalTo.Text = "" Or CSI_To1.Text = "" Then
                ilEnable = False
            Else
                ilEnable = True
            End If
        ElseIf ilListIndex = CNT_XML_ACTIVITY Then        '4-1-16
            'Date: 12/24/2019 added CSI calendar control for date entries
            'If edcSelCFrom.Text = "" Or edcSelCFrom1.Text = "" Or edcSelCTo.Text = "" Or edcSelCTo1.Text = "" Or (ckcSelC12(0).Value = vbUnchecked And ckcSelC12(1).Value = vbUnchecked) Then
            If (CSI_CalFrom.Text = "") Or (CSI_CalTo.Text = "") Or (CSI_From1.Text = "") Or (CSI_To1.Text = "") Or (ckcSelC12(0).Value = vbUnchecked And ckcSelC12(1).Value = vbUnchecked) Then
                ilEnable = False
            Else
                ilEnable = True
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
        ElseIf rbcOutput(2).Value Then   'Save As
            If (imFTSelectedIndex >= 0) And (edcFileName.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        ElseIf rbcOutput(3).Value Then                        'email
            imFTSelectedIndex = 0
            'ilEnable = True
            If (cbcEMailContent.ListIndex <= 0) Then
                ilEnable = False
            End If
        End If
    End If
    cmcGen.Enabled = ilEnable
    
    If ckcSelC8(0).Enabled = False Then ckcSelC8(0).Value = vbUnchecked
    If ckcSelC8(1).Enabled = False Then ckcSelC8(1).Value = vbUnchecked
End Sub

' *********************************************************************************************
'
'                           mSetupPopAAS - setup the parameters for the call to retrieve
'                                          contracts for BR report by advt, agency or slsp.
'
'***********************************************************************************************
Private Sub mSetupPopAAS()
    Dim slCntrType As String
    Dim slCntrStatus As String
    Dim ilHOState As Integer
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    lbcSelection(1).Visible = False
    lbcSelection(2).Visible = False
    lbcSelection(0).Visible = False
    lbcSelection(5).Visible = False
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
    'lbcSelection(10).Visible = True        '5-16-02
    imSetAll = False
    ckcAll.Value = vbUnchecked  '9-12-02 False
    ckcAll.Visible = True
    imSetAll = True
    ckcAll.Enabled = False
    ckcAllAAS.Value = False
    lbcSelection(10).Clear
    lbcSelection(0).Clear
    If ilListIndex = CNT_HISTORY Then
        lbcSelection(8).Visible = False
        lbcSelection(9).Visible = False
        'lbcSelection(7).Visible = True
        lbcSelection(7).Visible = False
        lbcSelection(5).Visible = True
        lbcSelection(0).Visible = True
        lbcSelection(10).Visible = False
        ckcAllAAS.Caption = "All Advertisers"
        'slCntrStatus = "HO"
        slCntrStatus = ""                   '5-9-17
        'mCntrPop slCntrStatus, 1            'ilHOState = 1 (show latest holds & orders), exclude  G & N
        mCntrPop slCntrStatus, 3            '5-9-17
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
    Else                                    'contract BR or insertion orders
        '10-29-03 no longer separate proposals vs orders, show them intermixed
        '11-6-03 client doesnt want them mixed, determine proposals, contracts or both

        If rbcSelCInclude(0).Value Then        'proposals
            slCntrStatus = "WDCI"              'all types, working, dead, complete, incomplete
            ilHOState = 0                       'show all versions for all proposals
        ElseIf rbcSelCInclude(1).Value = True Then       'contracts/orders (wide & narrow)
                slCntrStatus = "HO"                'default to holds and orders
                ilHOState = 3                       'show latest version of orders (including revised orders turned into proposals)
        Else
            slCntrStatus = ""                   'show all types
            ilHOState = 3
        End If
        'If rbcSelCInclude(0).Value Or rbcSelCInclude(1).Value Or rbcSelCInclude(2).Value Then
            If rbcSelCSelect(1).Value Then              'agy option
                'populate agency box
                lbcSelection(7).Visible = False
                lbcSelection(9).Visible = False
                lbcSelection(8).Visible = True
                lbcSelection(10).Visible = False '5-16-02
                lbcSelection(0).Visible = True  '5-16-02
                ckcAllAAS.Caption = "All Agencies"
                mAASCntrPop 1, 8, slCntrStatus, ilHOState   '5-16-02

            ElseIf rbcSelCSelect(2).Value Then          'slsp option
                'populate salesperson box
                lbcSelection(7).Visible = False
                lbcSelection(8).Visible = False
                lbcSelection(9).Visible = True
                lbcSelection(10).Visible = False '5-16-02
                lbcSelection(0).Visible = True
                ckcAllAAS.Caption = "All Salespeople"
                mAASCntrPop 2, 9, slCntrStatus, ilHOState   '5-16-02
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
            Else                                        'advt option
                'populate advertiser box
                lbcSelection(8).Visible = False
                lbcSelection(9).Visible = False
                'lbcSelection(7).Visible = True
                lbcSelection(7).Visible = False
                lbcSelection(5).Visible = True
                lbcSelection(0).Visible = True
                lbcSelection(10).Visible = False
                ckcAllAAS.Caption = "All Advertisers"
                mCntrPop slCntrStatus, ilHOState
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
            End If
        'End If
    End If
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
    '11-21-06 testing for all matching group # if slsp entered
    ilRet = gPopSalespersonBox(RptSelCt, 5, True, True, lbcSelection, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSPersonPopErr
        gCPErrorMsg ilRet, "mSPersonPop (gPopSalespersonBox)", RptSelCt
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
    Unload RptSelCt
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcECTab_GotFocus()
    If imDoubleClickName Then
        Exit Sub
    End If
    If GetFocus() <> pbcECTab.HWnd Then
        Exit Sub
    End If
    If imEMailContentSelectedIndex <= 0 Then
        If mEMailContentBranch() Then
            'cbcEMailContent.SetFocus
            Exit Sub
        End If
    End If
    edcResponse.SetFocus
End Sub

Private Sub plcSelC13_Paint()
    plcSelC13.Cls
    plcSelC13.CurrentX = 0
    plcSelC13.CurrentY = 0
    plcSelC13.Print smPaintCaption13
End Sub

Private Sub plcSelC14_Paint()
    plcSelC14.Cls
    plcSelC14.CurrentX = 0
    plcSelC14.CurrentY = 0
    plcSelC14.Print smPaintCaption14
End Sub

Private Sub rbcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    frcCopies.Visible = False   'Print Box
    frcCopies.Enabled = False
    frcFile.Visible = False     'Save to File Box
    frcFile.Enabled = False
    frcEMail.Visible = False    'Email Box
    frcEMail.Enabled = False
    cbcEMailContent.Visible = False
    lacExport.Visible = False
    cbcSet1.Enabled = True
    rbcSelC4(0).Enabled = True
    rbcSelC4(1).Enabled = True
    
    'JW 9/29/21 - Fix TTP 10271 per Jason Email: Tue 9/28/21 10:22 AM (issue #1)
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If Index <> 2 Then ckcSeparateFile.Value = 0
    
    'End of coded added
    If rbcOutput(Index).Value Then
        Select Case Index
            Case 0  'Display
                
            Case 1  'Print
                frcCopies.Enabled = True
                frcCopies.Visible = True
            Case 2  'File
                frcFile.Enabled = True
                frcFile.Visible = True
                'v81 testing results 3-28-22 Issue 1: "separate files per vehicle" checkbox from the Insertion Orders report is appearing
                If igRptCallType = CONTRACTSJOB And ilListIndex = CNT_INSERTION Then
                    ckcSeparateFile.Visible = True
                End If
            Case 3  'Export to pdf & email
                cbcEMailContent.Visible = True
                frcEMail.Visible = True
                frcEMail.Enabled = True
            Case 4  'Export to csv file
                'TTP 10119 - Average 30 Rate Report - add option to export to CSV
                lacExport.Visible = True
                lacExport.Caption = ""
                cbcSet1.Enabled = False
                rbcSelC4(0).Enabled = False
                rbcSelC4(1).Enabled = False
                
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
            If ilListIndex = CNT_BOBRECAP Then
                If Index = 0 Then         'selected sort by vehicle, ok to skip to new page per vehicle
                    ckcSelC10(0).Enabled = True
                Else
                    ckcSelC10(0).Enabled = False
                End If
            End If
    End Select
End Sub

Private Sub rbcSelC4_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC4(Index).Value
    'End of coded added
    Dim ilListIndex As Integer
    Dim illoop As Integer
    Dim ilFound As Integer
    Dim ilLoop2 As Integer
    Dim ilOnly10 As Integer
    ilListIndex = lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
         'If ilListIndex = CNT_BR Then
         '  If index = 1 Then                           'summary only, hidden doesnt apply
         '       If Value Then
         '           ckcSelC6(2).Enabled = False
         '           ckcSelC6(2).Value = False
         '       Else
         '           ckcSelC6(2).Enabled = True
         '       End If
         '   Else
         '       ckcSelC6(2).Enabled = True
         '   End If
         If ilListIndex = CNT_BR Then
            If Index = 0 And Value Then
                plcSelC9.Visible = False
            Else
                If Not ckcSelC6(1).Value = vbChecked Then

                    '****** temporarily patched out until ready to release
                    'plcSelC9.Visible = True
                End If
            End If
        ElseIf ilListIndex = CNT_BOB Then       '12-3-00 billed & booked, if any of the vehicle options (allow vehicle sub-totals to be supressed)
             If (rbcSelCInclude(2).Value Or rbcSelCInclude(4).Value) And Index = 2 And Value = True Then   '12-3-00 check for summary only
                '8-6-10
                If rbcSelC7(2).Value = True Then            'tnet with vehicle option, disallow slsp splits
                    ckcSelC10(1).Enabled = False
                    ckcSelC10(1).Value = vbUnchecked  'True
                    ckcSelC10(0).Value = vbUnchecked
                    plcSelC10.Enabled = False
                Else                                       '4-12-16 summary selected, no need to set default to show slsp subtotals
'                        ckcSelC10(1).Enabled = True
'                        ckcSelC10(1).Value = vbChecked  'True
'                        plcSelC10.Enabled = True
                    ckcSelC10(1).Caption = "Sub-totals by Slsp"     '11-7-16
                    ckcSelC10(1).Move 2400, -30, 2280
                    ckcSelC10(1).Visible = True
                    ckcSelC10(1).Enabled = True             '11-7-16
                   
                    ckcSelC10(2).Enabled = False                '11-7-16
                End If
             Else
                If Not (rbcSelCInclude(1).Value Or rbcSelCInclude(2).Value) Then     '1-4-2001
                    ckcSelC10(1).Enabled = False
                    ckcSelC10(1).Value = vbUnchecked    'False
                    ckcSelC10(2).Value = vbChecked                  '11-7-16 default office subtotals to show
                End If
             End If
        ElseIf ilListIndex = CNT_QTRLY_AVAILS Then                       'avails
            gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True
             ilOnly10 = True
             For illoop = LBound(tgMVef) To UBound(tgMVef) Step 1
                If (tgMVef(illoop).sState <> "D") And (tgMVef(illoop).sType = "S" Or tgMVef(illoop).sType = "C" Or tgMVef(illoop).sType = "V") Then
                    ilFound = gVpfFindIndex(tgMVef(illoop).iCode)
                    If ilFound < 0 Then
                        ilOnly10 = False
                    Else
                        For ilLoop2 = 0 To 9
                            If tgVpf(ilFound).iSLen(ilLoop2) <> 10 And tgVpf(ilFound).iSLen(ilLoop2) <> 0 Then
                                ilOnly10 = False
                                illoop = UBound(tgMVef)
                                Exit For
                            End If
                        Next ilLoop2
                    End If
                End If
             Next illoop
             If Value Then
                 If Index = 0 Then                       'qtrly summary avails
                    plcSelC10.Move lacSelCTo.Left, plcSelC2.Top + 230    '5-16-05 chg from plcselc3 to plcselc10
                    'mAskContractTypes
                    mAskCntrAndSpotTypesForAvails       '5-16-05
                    'ckcSelC5(1).Move 1800, -30, 1160
                    'ckcSelC5(1).Caption = "Reserved"
                    'If ckcSelC5(1).Value = vbChecked Then
                    '    ckcSelC5_click 1
                    'Else
                    '    ckcSelC5(1).Value = vbChecked   'True
                    'End If
                    'ckcSelC5(1).Visible = True
                    'If tgUrf(0).iSlfCode > 0 Then           'its a slsp, don't allow to exclude reserves
                    '    ckcSelC5(1).Enabled = False
                    'Else
                    '    ckcSelC5(1).Enabled = True
                    'End If

                    rbcSelC11(0).Caption = "Units"
                    rbcSelC11(1).Caption = "30/60"
                    plcSelC11.Move 120, plcSelC3.Top + plcSelC3.Height          '5-16-05 chg from plcselc6 to plcselc3 loc. of show 10,30 or 30/60
                    smPaintCaption11 = "Counts by"


                    plcSelC7.Visible = False
                    plcSelC2.Visible = True             'reshow daypart, days in dayp, dayp in days option
                    rbcSelCInclude(0).Value = True      'force to daypart option
                    plcSelC1.Visible = True             'reshow by avails, sellout, invent question
                    rbcSelCSelect(0).Value = True       'force to avails option
                    'gPopVehicleGroups RptSelCt, cbcSet1

                    cbcSet1.ListIndex = 0
                    edcSet1.Move 120, plcSelC11.Top + plcSelC11.Height + 60
                    cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 60
                    edcSet1.Visible = True
                    cbcSet1.Visible = True
                Else                                    'qtrly detail, disallow reservations selection
                    plcSelC8.Visible = False            'turn off orphan option
                    'If tgUrf(0).iSlfCode = 0 Then     'guide or counterpoint password
                        plcSelC10.Top = edcSelCFrom.Top + edcSelCFrom.Height + 30   '5-16-05 chg from plcselc3 to plcsel8
                        'mAskContractTypes
                        mAskCntrAndSpotTypesForAvails       '5-16-05
                        ckcSelC5(1).Enabled = False
                        ckcSelC5(1).Value = vbChecked   'True
                        'plcSelC7.Caption = "Reserved"
                        smPaintCaption7 = "Reserved"
                        plcSelC7_Paint
                        plcSelC7.Move 120, plcSelC3.Top + plcSelC3.Height     '5-16-05
                        plcSelC7.Height = 435
                        rbcSelC7(0).Caption = "Hide"
                        rbcSelC7(0).Move 990, 0, 780
                        rbcSelC7(0).Visible = True
                        rbcSelC7(0).Value = True            'default to hide reservations
                        rbcSelC7(1).Caption = "Show separately"
                        rbcSelC7(1).Move 1770, 0, 1920
                        rbcSelC7(1).Visible = True
                        rbcSelC7(2).Caption = "Exclude"
                        rbcSelC7(2).Move 990, 195, 1080
                        rbcSelC7(2).Visible = True
                        If tgUrf(0).iSlfCode > 0 Then       'disallow slsp from seeing reserves, force to bury them within sold
                            plcSelC7.Visible = False
                        Else
                            plcSelC7.Visible = True
                        End If
                    'Else
                    '    rbcSelC7(0).Value = True        'force to include in sold (Hide)
                    'End If
                    rbcSelC11(0).Caption = "Units"
                    plcSelC11.Move 120, plcSelC7.Top + plcSelC7.Height          'loc. of show 10,30 or 30/60
                    cbcSet1.ListIndex = 0
                    edcSet1.Move 120, plcSelC11.Top + plcSelC11.Height + 60
                    cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 60
                    edcSet1.Visible = True
                    cbcSet1.Visible = True

                    plcSelC2.Visible = False            'hide daypart, days in dayp, dayp in days option
                    rbcSelCInclude(0).Value = True      'force to daypart option
                    plcSelC1.Visible = False            'hide by avails, sellout, invent question
                    rbcSelCSelect(1).Value = True       'force to sellout option

                End If


                plcSelC11.Visible = True                     'Include 10, 30 or 30/60
                rbcSelC11(0).Move 960, 0, 720
                rbcSelC11(1).Move 1800, 0, 720
                'rbcSelC11(1).Caption = "30/60"
                rbcSelC11(0).Visible = True
                rbcSelC11(1).Visible = True
                If ilOnly10 Then                'if only 10s exist, default to units; otherwise take the default from site
                    rbcSelC11(0).Value = True
                Else
                    If tgSpf.sUnitOr3060 = "U" Then         'units (vs 30/60)
                        rbcSelC11(0).Value = True
                    Else
                        rbcSelC11(1).Value = True
                    End If
                End If

                'If Index = 0 Then                   'qtrly avails only, not qtrly detail
                    'For Standard quarter or use the input date entered and wrap around 13 weeks
                    'For standard quarter- the start of the qtr is determined by the date entered.
                    'any date after the start of the qtr until the date entered is blanked (ABC way)
                    plcSelC9.Move plcSelC11.Left, edcSet1.Top + edcSet1.Height
                    'plcSelC9.Caption = "Use"
                    smPaintCaption9 = "Use"
                    plcSelC9_Paint
                    rbcSelC9(0).Caption = "Standard Quarter"
                    rbcSelC9(0).Visible = True
                    rbcSelC9(0).Move 360, 0, 1920
                    rbcSelC9(1).Caption = "Start Date"
                    rbcSelC9(1).Visible = True
                    rbcSelC9(1).Move 2100, 0, 1200
                    plcSelC9.Visible = True
                    If rbcSelC9(0).Value Then
                        rbcSelC9_click 0
                    Else
                        rbcSelC9(0).Value = True
                    End If
                    'plcSelC8.Caption = ""
                    smPaintCaption8 = ""
                    plcSelC8_Paint
                    ckcSelC8(0).Caption = "Show Other Missed Spots Separately"
                    plcSelC8.Move plcSelC9.Left, plcSelC9.Top + plcSelC9.Height
                    ckcSelC8(0).Visible = True
                    ckcSelC8(0).Value = vbChecked   'True
                    ckcSelC8(0).Move 0, 0, 3420
                    plcSelC8.Visible = True


                    ckcSelC12(0).Value = vbChecked
                    ckcSelC12(1).Value = vbChecked
                    If tgSpf.sSystemType = "R" Then         'radio vs network/syndicator
                        'Feed option is placed next to "orders" & "Holds" as contract type spots
                        'plcSelC12.Move 2640, plcSelC3.Top, 1440
                        plcSelC12.Move 2640, plcSelC10.Top, 1440
                        ckcSelC12(0).Move 0, -30, 1440        'local
                        'ckcSelC12(1).Move 2400, 0, 1440      'feed
                        ckcSelC12(0).Visible = True
                        ckcSelC12(0).Caption = "Feed spots"
                        ckcSelC12(1).Visible = False

                        plcSelC12.Visible = True
                        smPaintCaption12 = ""   '"Include"
                        plcSelC12_Paint
                    End If
                'Else
                '    plcSelC9.Visible = False
                'End If
                'temporarily comment out 30 from this option (from 10, 30 or 30/60 to 10 or 30/60 only)
                'rbcSelC11(2).Visible = True
                'rbcSelC11(2).Move 1900, 0, 720
            End If                                      'if value
        ElseIf ilListIndex = CNT_MAKEPLAN Then
            If Index = 0 Then
                lacSelCTo.Visible = False
                lacSelCTo.Visible = False
                edcSelCTo.Visible = False
                edcSelCTo.Text = 1                  'for # quarters to 1
            Else
                lacSelCTo.Caption = "# Quarters"
                lacSelCTo.Visible = True
                lacSelCTo.Move 2760, 75
                edcSelCTo.Move 3780, edcSelCFrom.Top, 300
                edcSelCTo.MaxLength = 1
                edcSelCTo.Visible = True
            End If
        ElseIf ilListIndex = CNT_CUMEACTIVITY Then      '1-17-06
            If Index = 1 Then                           'summary, no vehicle subtotals
                plcSelC13.Visible = False
            Else
                If rbcSelCInclude(0).Value = True Then
                    plcSelC13.Visible = True
                End If
            End If
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
                If ilListIndex = CNT_BOB_BYSPOT Or ilListIndex = CNT_SPOTSALES Then
                    If Index = 0 Then                                'as ordered
                        ckcSelC3(0).Enabled = True
                        ckcSelC3(1).Enabled = True
                        ckcSelC3(2).Enabled = True
                        ckcSelC3(0).Value = vbChecked   'True                'always assume missed is included
                        ckcSelC3(1).Value = vbUnchecked ' False               'always assume cancel is excluded
                        ckcSelC3(2).Value = vbUnchecked 'False               'always asume hidden is excluded
                    Else
                        ckcSelC3(0).Enabled = True
                        ckcSelC3(1).Enabled = True
                        ckcSelC3(2).Enabled = True
                        ckcSelC3(0).Value = vbUnchecked 'False                'always assume missed is excluded
                        ckcSelC3(1).Value = vbUnchecked 'False               'always assume cancel is excluded
                        ckcSelC3(2).Value = vbUnchecked 'False               'always asume hidden is excluded
                    End If
                ElseIf ilListIndex = CNT_QTRLY_AVAILS Then                       'avails report
                    If Index = 2 Then                          'must have selected exclude Reservation
                        RptSelCt!ckcSelC5(1).Value = vbUnchecked    'False        'force to exclude reserves
                    Else                                        'hide or show reserves
                        RptSelCt!ckcSelC5(1).Value = vbChecked  ' True         'force to include reserves
                    End If
                ElseIf ilListIndex = CNT_BOB Then
                    If rbcSelCInclude(2).Value = True Then          'vehicle option
                        If Index = 2 And Value = True Then          'selected t-net
                            ckcSelC10(1).Visible = False            'disallow subtotals by slsp for tnet vehicle option
                            ckcSelC10(1).Value = vbUnchecked
                            ckcSelC10(0).Value = vbUnchecked        '8-6-10 disable slsp splits with t-net
                            ckcSelC10(0).Enabled = False
                        Else                                        'gross or net
                            ckcSelC10(1).Visible = True
                            ckcSelC10(1).Enabled = True             '11-7-16
                            ckcSelC10(0).Enabled = True

                        End If
                    End If
                ElseIf ilListIndex = CNT_PAPERWORK Then
                    If Index = 2 Then                   'show acq only, force to show line option
                        rbcSelCInclude(1).Value = True
                        rbcSelCInclude(0).Enabled = False
                    Else
                        rbcSelCInclude(0).Enabled = True
                    End If
                ElseIf ilListIndex = CNT_AVGRATE Then               '12-9-16        'agency option, show agency list box
                    If Index = 2 Then
                        lbcSelection(6).Height = 1700
                        lbcSelection(1).Height = 1700
    
                        lbcSelection(1).Move 15, 2090
                        lbcSelection(1).Height = 1700
                        lbcSelection(1).Visible = True
                        ckcAllAAS.Move lbcSelection(6).Left, lbcSelection(6).Top + lbcSelection(6).Height + 60
                        ckcAllAAS.Value = vbChecked         'default to all sales sources selected
                        lbcSelection(1).Move lbcSelection(6).Left, ckcAllAAS.Top + ckcAllAAS.Height + 30
                        ckcAllAAS.Value = vbUnchecked
                        ckcAllAAS.Caption = "All Agencies"
                        ckcAllAAS.Visible = True
                    Else
                        lbcSelection(6).Height = 3500
                        lbcSelection(1).Visible = False
                        ckcAllAAS.Visible = False
                    End If
                End If
            Case SLSPCOMMSJOB
                If ilListIndex = COMM_SALESCOMM Then
                    If Index = 1 Then               'summary version
                        'plcSelC8.Visible = False    'contract subtotals
                        ckcSelC8(0).Enabled = False
                        plcSelC11.Visible = True   'sort by
                    Else                            'detail
                        'plcSelC8.Visible = True
                        If ckcSelC3(0).Value = vbChecked Then      'add bonus comm for new/increased
                            ckcSelC8(0).Enabled = False
                        Else
                            ckcSelC8(0).Enabled = True
                        End If
                        plcSelC11.Visible = True
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
             If ilListIndex = CNT_BOB Then                   'Bill & Booked
                 ckcSelC8(0).Enabled = True
                 ckcSelC8(1).Enabled = True
                 If Value Then
                     If Index = 0 Then                      'Corp (vs Std)
                        ckcSelC6(0).Value = vbUnchecked     'False           'default trades off
                        ckcSelC8(1).Value = vbUnchecked 'False           'default show mgs where they air off
                        ckcSelC8(0).Visible = False
                        ckcSelC8(1).Visible = False
                        ckcSelC8(0).Value = vbUnchecked 'False       'ignore missed
                        ckcSelC8(1).Value = vbUnchecked 'False       'ignore mgs
                        edcText.Enabled = True
                        lacText.Enabled = True
                        ckcInclRevAdj.Value = vbChecked
                     ElseIf Index = 3 Then                          '2-6-15 cal spots always shows spots where scheduled
                        ckcSelC8(0).Visible = True
                        'ckcSelC8(1).Visible = False                '7-29-16 reinstate option
                        ckcInclRevAdj.Value = vbChecked             '7-29-16 default to incl Rev Adjustments
                        edcText.Enabled = False                     '8-23-17 disable ability to pace on cal spots
                        lacText.Enabled = False
                        edcText.Text = ""                           '8-24-17 ensure the pacing date is blank, a previous report could have been pacing
                        'TTP 10257: Billed and Booked: Cal spots - request - restore "Count MGs where they air" option for "as ordered" bill method
                        ''12-4-17 show selection to show mg where they air if Aired Billing; otherwise for As Ordered always include the spot as ordered
                        'ckcSelC8(1).Visible = False
                        If tgSpf.sInvAirOrder = "A" Then
                            ckcSelC8(0).Visible = True
                            ckcSelC8(0).Enabled = True
                            ckcSelC8(1).Visible = True
                            ckcSelC8(1).Enabled = True
                        Else
                            If tgSpf.sInvAirOrder = "S" Then
                                'disable MG and Missged options when Bill Method = Bill as Order (option A) sInvAirOrder="S"
                                ckcSelC8(0).Visible = True
                                ckcSelC8(0).Enabled = False
                                ckcSelC8(0).Value = False
                                ckcSelC8(1).Visible = True
                                ckcSelC8(1).Enabled = False
                                ckcSelC8(1).Value = False
                            Else
                                ckcSelC8(0).Visible = True
                                ckcSelC8(0).Enabled = True
                                ckcSelC8(1).Visible = True
                                ckcSelC8(1).Enabled = True
                            End If
                        End If
                        ckcSelC8(1).Visible = True
                     Else                       'std or cal (spots) or Bill Method
                        ckcSelC6(0).Value = vbChecked   'True
                        ckcSelC8(1).Value = vbChecked   'True            'set default to show mgs whre they air
                        ckcSelC8(0).Visible = True
                        ckcSelC8(1).Visible = True
                        ckcSelC8(0).Value = vbUnchecked 'False       'ignore missed
                        ckcSelC8(1).Value = vbChecked   'True       'ignore mgs
                        ckcInclRevAdj.Value = vbChecked
                        If Index = 3 Then                           '1-09-08if cal by spots, disallow the pacing feature
                            edcText.Enabled = False
                            lacText.Enabled = False
                            edcText.Text = ""
                        Else
                            edcText.Enabled = True          'pacing date
                            lacText.Enabled = True
                            '12-8-17 disable sub unresolved missed and show mg where air if billing is as ordered, update ordered for std option
                            If tgSpf.sInvAirOrder = "S" Then             'bill as ordered, update as ordered; don't ask adjustment qustions
                                'disable MG and Missged options when Bill Method = Bill as Order (option A) sInvAirOrder="S"
                                ckcSelC8(0).Visible = True
                                ckcSelC8(1).Visible = True
                                ckcSelC8(0).Enabled = False
                                ckcSelC8(1).Enabled = False
                                ckcSelC8(0).Value = False
                                ckcSelC8(1).Value = False
                            Else                                'as aired
                                ckcSelC8(0).Visible = True  '9-12-02 vbChecked 'True
                                ckcSelC8(1).Visible = True  '9-12-02 vbChecked 'True
                                ckcSelC8(0).Enabled = True
                                ckcSelC8(1).Enabled = True
                            End If
                        End If
                        'TTP 10634 - found out Cal Contracts mode doesnt adjust, hide the treat MG as Aired
                        If Index = 2 Then ckcSelC8(1).Visible = False
                     End If
                End If
            ElseIf ilListIndex = CNT_PAPERWORK Then
                If rbcSelC9(3).Value Then              'by vehicle, contract or line options not available
                    plcSelC2.Visible = False                'disallow summary if by vehicle
                    rbcSelCInclude(1).Value = True          'default to line report
                    cbcSet1.Visible = False
                    edcSet1.Visible = False
                    ckcSelC5(7).Value = vbUnchecked         'disallow NTR with any vehicle/line option
                    ckcSelC5(7).Enabled = False
                Else                                '2-28-05 not vehicle sort, allow report by contract summary or line
                    rbcSelCInclude(0).Value = True      'default to show by contract
                    plcSelC2.Visible = True
                    cbcSet1.Visible = True
                    edcSet1.Visible = True
                    ckcSelC5(7).Value = vbChecked         'turn NTR option back on and default to show on report
                    ckcSelC5(7).Enabled = True
                    If rbcSelCSelect(3).Value = True Then
                        plcSelC2.Visible = False                'disallow summary if by vehicle
                        rbcSelCInclude(1).Value = True          'default to line report
                        cbcSet1.Visible = False
                        edcSet1.Visible = False
                        ckcSelC5(7).Value = vbUnchecked         'disallow NTR with any vehicle/line option
                        ckcSelC5(7).Enabled = False
                    End If
                    If rbcSelC7(2).Value Then                   'show acq cost only
                        rbcSelCInclude(1).Value = True          'force to show by lines to see acq cost
                    End If
                End If
            ElseIf ilListIndex = CNT_SALESACTIVITY_SS Or ilListIndex = CNT_SALESPLACEMENT Then      '2-18-03 allow option to include slsp subtotals
                If Index = 0 Or Index = 1 Then          'by market, source, office or source, office, mkt
    
                    plcSelC12.Height = 440
                    ckcSelC12(1).Move 0, 210, 2640
                    ckcSelC12(1).Caption = "Include Slsp Sub-totals"
                    ckcSelC12(1).Visible = True
                    If ilListIndex = CNT_SALESACTIVITY_SS Then
                        '4-1-11 option to split the slsp
                        ckcSelC12(2).Move 2520, 210, 1920
                        ckcSelC12(2).Caption = "Show Slsp Splits"
                        ckcSelC12(2).Visible = True
                    End If
                Else
                    plcSelC12.Height = 240
                    ckcSelC12(1).Visible = False
                    ckcSelC12(1).Value = vbUnchecked        'force to exclude slsp subtotals
                    ckcSelC12(2).Value = vbUnchecked        '4-1-11 no slsp, no splits
                    ckcSelC12(2).Visible = False
    
                End If
                If ilListIndex = CNT_SALESACTIVITY_SS Then
                    '09/28/2020 - TTP # 9952 - IF include NTR, Add option to split NTR (or by default: leave NTR grouped together)
                    plcSelC13.Move 0, plcSelC12.Top + plcSelC12.Height, 4000
                    ckcSelC13(0).Caption = "Air Time"
                    ckcSelC13(1).Caption = "NTR"
                    ckcSelC13(2).Caption = "Hard Cost"
                    smPaintCaption13 = "Include"
                    plcSelC13_Paint
                    ckcSelC13(0).Value = vbChecked
                    ckcSelC13(1).Value = vbUnchecked
                    ckcSelC13(2).Value = vbUnchecked
                    ckcSelC13(0).Move 840, 0, 1080
                    ckcSelC13(1).Move 1920, 0, 720
                    ckcSelC13(2).Move 2640, 0, 1200
                    ckcSelC13(0).Visible = True
                    ckcSelC13(1).Visible = True
                    ckcSelC13(2).Visible = True
                    plcSelC13.Visible = True
                End If
                
             '6-15-11 remove, use control for airtime/rep opton in Advertiser Units ordered, AVg Rate & Avg Spot Price reports
             'data all retrieved from contract & lines
    '            ElseIf ilListIndex = CNT_ADVT_UNITS Then            '9-23-09
    '                'if TNet, ask to Include Merchandising
    '                If Index = 2 Then
    '                    plcSelC10.Move 120, plcSelC9.Top + plcSelC9.Height
    '                    'plcSelC10.Visible = True
    '                    plcSelC10.Visible = False   'currently commented out to view, feature not implemented
    '                    ckcSelC10(0).Caption = "Include Merchandising/Promotions"
    '                    ckcSelC10(0).Move 0, 0, 3840
    '                    ckcSelC10(0).Visible = True
    '                Else
    '                    plcSelC10.Visible = False
    '                    ckcSelC10(0).Value = vbUnchecked
    '                End If
            End If
        Case SLSPCOMMSJOB
            If ilListIndex = COMM_SALESCOMM Then
                If Index = 0 Then
                    If Value = True Then
                        ckcSelC10(0).Value = vbUnchecked
                        ckcSelC10(0).Enabled = False
                    End If
                Else
                    ckcSelC10(0).Enabled = True
                End If
            End If
    End Select
End Sub

Private Sub rbcSelCInclude_Click(Index As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelCInclude(Index).Value
    'End of coded added
    Dim ilListIndex As Integer
    Dim ilRet As Integer
    Dim ilSetIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    If Value Then
        Select Case igRptCallType
            Case CONTRACTSJOB
                If (igRptType = 0) And (ilListIndex > 1) Then
                    ilListIndex = ilListIndex + 1
                End If
                If ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION Then                 'contract/proposals
                    mSetupPopAAS                        'setup to fill list box with proper contracts
                    If rbcSelCInclude(0).Value Then                 'proposals
                        If ilListIndex <> CNT_INSERTION Then
                            plcSelC3.Enabled = False
                            plcSelC8.Enabled = False
                            ckcSelC3(0).Enabled = False
                            ckcSelC8(0).Enabled = False
                            plcSelC4.Visible = True                     'summary,detail
                            plcSelC7.Visible = True                     'use primary or all demos

                            plcSelC6.Visible = True                     'Include prices
                            plcSelC5.Visible = True                     'diff only
                            ckcSelC5(0).Enabled = False                  'dont allow differenes on proposals
                            ckcSelC5(0).Value = vbUnchecked 'False
                            ckcSelC6(1).Value = vbChecked   'True                   'set default to Research on when running proposals
                            ckcSelC6(0).Value = vbChecked   'True                    'default to  rates with Research version
                            ckcSelC3(0).Value = vbUnchecked 'False
                            ckcSelC8(0).Value = vbUnchecked 'False

                            'Temporary patch until ABC pays
                            'plcSelC9.Visible = True                 'show corp or std option
                        End If
                    ElseIf rbcSelCInclude(1).Value Then             'wide (orders)
                        If ilListIndex = CNT_INSERTION Then
                            plcSelC3.Visible = False
                        Else
                            'wide or narrow (not proposals)
                            ckcSelC3(0).Caption = "Printables Only"
                            plcSelC3.Move plcSelC4.Left, plcSelC5.Top + plcSelC5.Height
                            plcSelC3.Visible = True
                            ckcSelC3(0).Enabled = True
                            ckcSelC3(0).Value = vbUnchecked   'False        5-2-03 was set incorrectly

                            plcSelC4.Visible = True                     'summary, detail
                            plcSelC7.Visible = True                    '5-15-02 chged from false, for orders, force to use primary demo
                            rbcSelC7(0).Value = True
                            plcSelC6.Visible = True                     'include prices
                            plcSelC5.Visible = True                     'diff only
                            ckcSelC5(0).Enabled = True                   'allow differences on orders

                            plcSelC8.Enabled = True                     'for printables only, show mods as differences
                            plcSelC8.Visible = True
                            ckcSelC8(0).Enabled = True
                            ckcSelC3(0).Enabled = True
                            plcSelC3.Enabled = True

                            'Temporary patch until ABC pays
                            'plcSelC9.Visible = True                 'show corp or std option
                        End If
                    Else                                            '11-6-03 this option is no longer Narrow portrait (it has been removed).
                                                                    'this optionis now combined proposals & orders
                        'wide or narrow (not proposals)
                        ckcSelC3(0).Caption = "Printables Only"
                        ckcSelC3(0).Enabled = False                 '11-6-03    narrow version no longer exists
                        ckcSelC3(0).Value = vbUnchecked 'False
                        plcSelC3.Move plcSelC4.Left, plcSelC5.Top + plcSelC5.Height
                        plcSelC3.Visible = True
                        'plcSelC4.Visible = False                    'summary, detail
                        'plcSelC7.Visible = False
                        'plcSelC6.Visible = False                    ' turn off include prices
                        'plcSelC5.Visible = False                    'turn off dif only
                        plcSelC8.Enabled = False
                        ckcSelC8(0).Enabled = False
                        plcSelC9.Visible = False                    'turn off corp or std
                    End If

                End If
                If (ilListIndex = CNT_BOB_BYCNT) Or (ilListIndex = CNT_BOB_BYSPOT) Or (ilListIndex = CNT_BOB_BYSPOT_REPRINT) Then 'Projection

                    Select Case Index
                        Case 0  'Advertiser/Contract #
                            lbcSelection(2).Visible = False
                            lbcSelection(3).Visible = False
                            lbcSelection(6).Visible = False
                            lbcSelection(1).Visible = False
                            If ckcAll.Value = vbChecked Then    '9-12-02
                                lbcSelection(5).Visible = False
                                lbcSelection(0).Visible = False
                            Else
                                lbcSelection(5).Visible = True
                                lbcSelection(0).Visible = True
                            End If
                            ckcAll.Caption = "All Advertisers"
                            'edcSet1.Visible = False
                            'cbcSet1.Visible = False

                        Case 1  'Salesperson
                            lbcSelection(0).Visible = False
                            lbcSelection(3).Visible = False
                            lbcSelection(5).Visible = False
                            lbcSelection(6).Visible = False
                            lbcSelection(2).Visible = True
                            ckcAll.Caption = "All Salespeople"
                            mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                            If imTerminate Then
                                cmcCancel_Click
                                Exit Sub
                            End If
                            'edcSet1.Visible = False
                            'cbcSet1.Visible = False

                        Case 2  'Vehicle
                            lbcSelection(0).Visible = False
                            lbcSelection(2).Visible = False
                            lbcSelection(5).Visible = False
                            If ilListIndex = CNT_BOB_BYCNT Then
                                lbcSelection(3).Visible = False
                                lbcSelection(6).Visible = True
                            Else
                                lbcSelection(6).Visible = False
                                lbcSelection(3).Visible = True
                            End If
                            ckcAll.Caption = "All Vehicles"
                        Case 3  'agency
                            lbcSelection(2).Visible = False
                            lbcSelection(5).Visible = False
                            lbcSelection(6).Visible = False
                            lbcSelection(3).Visible = False
                            lbcSelection(1).Visible = True
                            ckcAll.Caption = "All Agencies"
                    End Select
                    'Date: 11/1/2019 added Major/Minor sorts; used CSI calendar for date entry
'                ElseIf ilListIndex = CNT_AVG_PRICES Then        'Weekly Average Spot Prices
'                    If Index = 1 Then                           'vehicle option
'                        lbcSelection(6).Visible = True
'                        lbcSelection(2).Visible = False
'                        ckcAll.Caption = "All Vehicles"
'                    Else                                        'slsp option
'                        lbcSelection(2).Visible = True
'                        lbcSelection(6).Visible = False
'                        ckcAll.Caption = "All Salespeople"
'                    End If
                ElseIf ilListIndex = CNT_PAPERWORK Then     '6-14-02
                    If Index = 1 And Value = True Then      'if show by line, dont allow comm or NTR to be selected
                        ckcSelC13(0).Value = vbUnchecked
                        ckcSelC13(0).Enabled = False
                        ckcSelC5(7).Value = vbUnchecked     'turn off NTR
                        ckcSelC5(7).Enabled = False         'disable NTR feature
                    Else                                    'selected show by contract
                        ckcSelC13(0).Enabled = True         'enable commission feature
                        ckcSelC5(7).Value = vbChecked       'turn back on NTR
                        ckcSelC5(7).Enabled = True         'enable NTR feature

                    End If
                ElseIf ilListIndex = CNT_TIEOUT Then        'Detail Tie Out
                    If rbcSelCInclude(0).Value Then      'office
                        lbcSelection(2).Visible = True
                        lbcSelection(2).Move 120, ckcAll.Top + ckcAll.Height + 30, 4380, 1500  'office list box
                        lbcSelection(6).Visible = False
                        ckcAll.Caption = "All Offices"
                        lbcSelection(12).Visible = True
                        laclbcName(1).Visible = True
                        ckcAll.Value = vbUnchecked  '9-12-02False
                    Else                                'vehicle
                        lbcSelection(6).Visible = True
                        lbcSelection(6).Move 120, ckcAll.Top + ckcAll.Height + 30, 4380, 1500  'vehicle list box
                        lbcSelection(2).Visible = False
                        ckcAll.Caption = "All Vehicles"
                        lbcSelection(12).Visible = False
                        laclbcName(1).Visible = False
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                    End If
                    ckcAll.Visible = True
                ElseIf ilListIndex = CNT_BOB Then       'Billed & Booked
                    'reentrant when vehicle/participant requested; then a different option;  reposition the list boxes for vehicle and owners
                    lbcSelection(6).Top = lbcSelection(5).Top
                    lbcSelection(6).Width = lbcSelection(5).Width
                    lbcSelection(6).Height = lbcSelection(5).Height
                    lbcSelection(2).Top = lbcSelection(5).Top
                    lbcSelection(2).Height = lbcSelection(5).Height
                    lbcSelection(2).Width = lbcSelection(5).Width
                    ckcAll.Visible = True
                    ckcAllAAS.Visible = False           '1-4-2001
                    plcSelC10.Enabled = False
                    ckcSelC10(0).Enabled = False
                    ckcSelC10(1).Enabled = False        '12-03-00   sub-tots by vehicle
                    ckcSelC10(1).Value = vbUnchecked    'False          '12-3-00
                    ckcSelC10(2).Enabled = False                '11-7-16 subtotals by ofc
                    ckcSelC10(2).Value = vbUnchecked
                    CkcAllveh.Visible = False       'turn off sales office selection incase slsp option selected,
                    lbcSelection(7).Visible = False
                    lbcSelection(2).Visible = False
                    lbcSelection(5).Visible = False
                    lbcSelection(1).Visible = False
                    sgSOCodeTagCT = ""                  'multi used list box, need to make sure the sales office is repopulated when switching using
                                    'vehicle groups and sales offices

                    If Index = 0 Then                   'advt
                        lbcSelection(5).Move lbcSelection(2).Left, lbcSelection(2).Top, 4380, 3270
                        lbcSelection(2).Visible = False
                        lbcSelection(6).Visible = False
                        lbcSelection(5).Visible = True
                        ckcAll.Caption = "All Advertisers"
                        'plcSelC7.Visible = True            'show gross net option
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        ckcSelC10(0).Value = vbUnchecked    'False
                        edcSet1.Visible = False
                        cbcSet1.Visible = False
                    ElseIf Index = 1 Then                   'slsp
                        mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                        If imTerminate Then
                            cmcCancel_Click
                            Exit Sub
                        End If

                        lbcSelection(2).Visible = True      'slsp list
                        lbcSelection(6).Visible = False
                        lbcSelection(5).Visible = False
                        ckcAll.Caption = "All Salespeople"
                        'plcSelC7.Visible = True            'show gross net option
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        lbcSelection(2).Width = 2100
                        lbcSelection(2).Height = 1500
                        CkcAllveh.Caption = "All Sales Offices"        'ckcAllAAS is used for Vehicle list box

                        CkcAllveh.Visible = True
                        lbcSelection(7).Move lbcSelection(2).Left + lbcSelection(2).Width + 90, lbcSelection(2).Top, lbcSelection(2).Width, 1500
                        CkcAllveh.Move lbcSelection(7).Left, ckcAll.Top
                        mSalesOfficePop lbcSelection(7)     'sales office list
                        If imTerminate Then
                            cmcCancel_Click
                            Exit Sub
                        End If
                        lbcSelection(7).Visible = True

                        plcSelC10.Enabled = True
                        ckcSelC10(0).Caption = "Sls Splits"    '8-8-06 changed from Use Primary Only to Show Splits
                        ckcSelC10(0).Value = vbChecked          'default to use splits for slsp option
                        ckcSelC10(0).Enabled = True

                        ckcSelC10(1).Caption = "Veh Sub-tots"
                        ckcSelC10(1).Move 1320, -30, 1560
                        ckcSelC10(1).Visible = True
                        If ckcSelC10(1).Value = vbChecked Then
                            ckcSelC10(1).Value = vbUnchecked    ', False
                        Else
                            ckcSelC10_click 1
                        End If

                        ckcSelC10(1).Enabled = True
                        
                        '11-7-16 allow office subtotals to be hidden
                        ckcSelC10(2).Caption = "Ofc Sub-Tots"
                        ckcSelC10(2).Move 3000, -30, 1560
                        ckcSelC10(2).Visible = True
                        ckcSelC10(2).Value = vbChecked
                        ckcSelC10(2).Enabled = True
                        
                        edcSet1.Visible = False
                        cbcSet1.Visible = False
                    ElseIf Index = 2 Or Index = 4 Or Index = 6 Then               '11-17-06 vehicle, vehicle/part, vehicle net-net
                        ckcSelC10(0).Value = vbUnchecked    'False
                        ckcSelC10(1).Value = vbUnchecked    'False
                        ckcSelC10(2).Value = vbUnchecked                            '11-7-16
                        'plcSelC7.Visible = True            'show  gross net option
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        If rbcSelC4(2).Value Then            '12-3-00 summary only ? allow vehicle sub-totals to be suppressed
                            ckcSelC10(0).Value = vbUnchecked    'False
                            ckcSelC10(1).Value = vbChecked  'True
                            ckcSelC10(1).Enabled = True
                            plcSelC10.Enabled = True
                        End If
                        '3-2-02
                        gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True
                        edcSet1.Text = "Grp"
                        cbcSet1.ListIndex = 0


                        edcSet1.Move 2500, lacTopDown.Top, 360
                        cbcSet1.Move 2920, cbcSet2.Top
                        'edcSet1.Move 120, plcSelC2.Top + plcSelC2.Height + 30, 1080
                        'cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 45  '45
                        edcSet1.Visible = True
                        cbcSet1.Visible = True
                        lbcSelection(1).Visible = False


                        If Index = 2 Or Index = 6 Then           'vehicle or vehicle net-net option
                            'If rbcSelC7(2).Value = False Then
                            If rbcSelC7(2).Value = False Then        't-net, disallow subtotals by slsp
                                '10-18-06 if by vehicle, new option to subtotal by slsp with vehicle
                                ckcSelC10(1).Caption = "Sub-totals by Slsp"
                                ckcSelC10(1).Move 2400, -30, 2280
                                ckcSelC10(1).Visible = True
                                ckcSelC10(1).Enabled = True             '11-7-16
                                If ckcSelC10(1).Value = vbChecked Then
                                    ckcSelC10(1).Value = vbUnchecked    ', False
                                Else
                                    ckcSelC10_click 1
                                End If
                                plcSelC10.Enabled = True
                                ckcSelC10(1).Enabled = True
                                
                                ckcSelC10(2).Enabled = False                '11-7-16
                            Else        't-net with vehicle option; disallow subtotals by slsp
                                ckcSelC10(1).Value = vbUnchecked
                                ckcSelC10(1).Visible = False
                                ckcSelC10(0).Value = vbUnchecked
                                ckcSelC10(0).Enabled = False
                                ckcSelC10(2).Enabled = False            '11-7-16
                            End If
                              lbcSelection(6).Move lbcSelection(2).Left, lbcSelection(2).Top, 4380, 3270
                              lbcSelection(2).Visible = False
                              lbcSelection(6).Visible = True
                              lbcSelection(5).Visible = False
                              ckcAll.Caption = "All Vehicles"
                              ckcAllAAS.Visible = False
                    
                            'End If

                        Else                        'vehicle with participant splits
                            lbcSelection(6).Move lbcSelection(2).Left, lbcSelection(2).Top, 4380, 1500  '1740
                            lbcSelection(2).Visible = False
                            lbcSelection(5).Visible = False
                            ckcAll.Caption = "All Vehicles"
                            'mMnfPop "H", RptSelCt!lbcSelection(2), tgSalesperson(), sgSalespersonTag    'Traffic!lbcSalesperson
                            ilRet = gPopMnfPlusFieldsBox(RptSelCt, RptSelCt!lbcSelection(2), tgSalesperson(), sgSalespersonTag, "H1")
                            If imTerminate Then
                                cmcCancel_Click
                                Exit Sub
                            End If
                            lbcSelection(2).Move lbcSelection(6).Left, lbcSelection(6).Top + lbcSelection(6).Height + 375, 4380, 1500
                            ckcAllAAS.Caption = "All Participants"
                            ckcAllAAS.Left = ckcAll.Left
                            ckcAllAAS.Top = lbcSelection(2).Top - ckcAllAAS.Height
                            ckcAllAAS.Value = vbUnchecked   '9-12-02 False
                            ckcAllAAS.Visible = True
                            lbcSelection(5).Visible = False
                            lbcSelection(6).Visible = True
                            lbcSelection(2).Visible = True      'owner list
                            If ilSetIndex > 0 Then              'vehicle group selected, show the items for it
                                lbcSelection(2).Width = 2100    'width of participant box
                                lbcSelection(7).Width = 2100    'width of vehicle group items box
                                lbcSelection(7).Left = lbcSelection(2).Left + lbcSelection(2).Width + 90
                            End If

                        End If
                    ElseIf Index = 5 Then         '04-12-02 agency option

                        If imTerminate Then
                            cmcCancel_Click
                            Exit Sub
                        End If
                        lbcSelection(1).Visible = True
                        lbcSelection(6).Visible = False
                        lbcSelection(5).Visible = False
                        lbcSelection(2).Visible = False
                        ckcAll.Caption = "All Agencies"
                        'plcSelC7.Visible = True            'show gross net option
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        plcSelC10.Enabled = False
                        ckcSelC10(0).Enabled = False
                        ckcSelC10(1).Enabled = False
                        ckcSelC10(2).Enabled = False        '11-7-16
                        edcSet1.Visible = False
                        cbcSet1.Visible = False
                    Else                                'ownership
                        'mMnfPop "H", RptSelCt!lbcSelection(2), tgSalesperson(), sgSalespersonTag    'Traffic!lbcSalesperson
                        ilRet = gPopMnfPlusFieldsBox(RptSelCt, RptSelCt!lbcSelection(2), tgSalesperson(), sgSalespersonTag, "H1")
                        If imTerminate Then
                            cmcCancel_Click
                            Exit Sub
                        End If
                        lbcSelection(2).Visible = True
                        lbcSelection(6).Visible = False
                        lbcSelection(5).Visible = False
                        ckcAll.Caption = "All Owners"
                        ckcAllAAS.Visible = False           '12-3-00
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        ckcSelC10(0).Value = vbUnchecked  '2-8-16
                        ckcSelC10(1).Value = vbUnchecked  '2-8-16
                        ckcSelC10(2).Value = vbUnchecked                '11-7-16
                        edcSet1.Visible = False
                        cbcSet1.Visible = False

                    End If
                ElseIf ilListIndex = CNT_SALESCOMPARE Or ilListIndex = CNT_ADVT_UNITS Or ilListIndex = CNT_AVG_PRICES Then
                    'Date: 8/27/2019 added major/minor sorts to ADVERTISER UNITS
                    ckcAll.Visible = True
                    mMnfPop "B", RptSelCt!lbcSelection(3), tgMNFCodeRpt(), sgMNFCodeTagRpt    'Traffic!lbcSalesperson
                    If imTerminate Then
                        cmcCancel_Click
                        Exit Sub
                    End If
                    mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                    If imTerminate Then
                        cmcCancel_Click
                        Exit Sub
                    End If

                    mMnfPop "C", RptSelCt!lbcSelection(7), tgMnfCodeCT(), sgMNFCodeTagRpt    'Traffic!lbcSalesperson
                    If imTerminate Then
                        cmcCancel_Click
                        Exit Sub
                    End If

                    'vehicle groups:  participants, formats, markets, etc.  Show each group that at has at least 1 defined
                    'gPopVehicleGroups lbcSelection(4), tgVehicleSets1(), True
                    gPopVehicleGroups lbcSelection(12), tgVehicleSets1(), True          '3-18-16 change to single selection box

                    If cbcSet2.ListIndex = 0 Then
                        'no minor sort set, use the entire screen height for the major sort selection
                        lbcSelection(2).Visible = False                 'slsp, cat or prod list box
                        lbcSelection(6).Visible = False                 'vehicle list box
                        lbcSelection(5).Visible = False                 'advt list box
                        lbcSelection(1).Visible = False                 'agency list box
                        lbcSelection(3).Visible = False                 'bus cat
                        lbcSelection(4).Visible = False                 'Vehicle group
                        lbcSelection(12).Visible = False                '3-18-16 vehicle group , single selection
                        lbcSelection(7).Visible = False                 'Prod Protection
                        lbcSelection(8).Visible = False                 'vehicle group items
                        lbcSelection(1).Height = 3270
                        lbcSelection(2).Height = 3270
                        lbcSelection(3).Height = 3270
                        lbcSelection(4).Height = 3270
                        lbcSelection(12).Height = 3270
                        lbcSelection(5).Height = 3270
                        lbcSelection(6).Height = 3270
                        lbcSelection(7).Height = 3270
                        lbcSelection(8).Height = 3270
                    End If

                    'changing the major sort selection, only turn off the previous list box
                    If imPrevMajorIndex = 0 Then          'advt
                        lbcSelection(5).Visible = False
                    ElseIf imPrevMajorIndex = 1 Then         'agy
                        lbcSelection(1).Visible = False
                    ElseIf imPrevMajorIndex = 2 Then         'bus cat
                        lbcSelection(3).Visible = False
                    ElseIf imPrevMajorIndex = 3 Then         'prod protection
                        lbcSelection(7).Visible = False
                    ElseIf imPrevMajorIndex = 4 Then         'slsp
                        lbcSelection(2).Visible = False
                    ElseIf imPrevMajorIndex = 5 Then         'vehicle
                        lbcSelection(6).Visible = False
                    ElseIf imPrevMajorIndex = 6 Then        'vehicle group
                        lbcSelection(4).Visible = False
                        lbcSelection(12).Visible = False        '3-18-16
                        lbcSelection(8).Visible = False     'vg items
                        CkcAllveh.Visible = False
                        CkcAllveh.Value = vbUnchecked
                    End If

                    If Index = 0 Then                   'advt
                        lbcSelection(5).Move lbcSelection(6).Left, 280, lbcSelection(6).Width, lbcSelection(6).Height
                        lbcSelection(5).Visible = True
                        ckcAll.Caption = "All Advertisers"
                        'plcSelC7.Visible = True            'show gross net option
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                    ElseIf Index = 1 Then                   'slsp
                        lbcSelection(2).Move lbcSelection(6).Left, 280, lbcSelection(6).Width, lbcSelection(6).Height

                        lbcSelection(2).Visible = True
                        ckcAll.Caption = "All Salespeople"
                        'plcSelC7.Visible = True            'show gross net option
                        ckcAll.Value = vbUnchecked  '9-12-02  False
                    ElseIf Index = 2 Then                   'agency
                        lbcSelection(1).Move lbcSelection(6).Left, 280, lbcSelection(6).Width, lbcSelection(6).Height
                        lbcSelection(1).Visible = True
                        ckcAll.Caption = "All Agencies"
                        'plcSelC7.Visible = True            'show  gross net option
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                    ElseIf Index = 3 Then                    'Bus Category
                        'mMnfPop "B", RptSelCt!lbcSelection(2), tgSalesperson(), sgSalespersonTag    'Traffic!lbcSalesperson
                        lbcSelection(3).Move lbcSelection(6).Left, 280, lbcSelection(6).Width, lbcSelection(6).Height
                        lbcSelection(3).Visible = True
                        ckcAll.Caption = "All Categories"
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                    ElseIf Index = 4 Then                                'product Protection categories
                        'mMnfPop "C", RptSelCt!lbcSelection(2), tgSalesperson(), sgSalespersonTag    'Traffic!lbcSalesperson
                        lbcSelection(7).Move lbcSelection(6).Left, 280, lbcSelection(6).Width, lbcSelection(6).Height
                        lbcSelection(7).Visible = True
                        ckcAll.Caption = "All Prod Protection"
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                    ElseIf Index = 5 Then           '4-25-06 vehicle option
                        lbcSelection(6).Move lbcSelection(6).Left, 280, lbcSelection(6).Width, lbcSelection(6).Height
                        lbcSelection(6).Visible = True
                        ckcAll.Caption = "All Vehicles"
                        ckcAll.Value = vbUnchecked
                    Else                            'vehicle group
                        ckcAll.Visible = False      'do not show ALL option, can only select 1 group
'                        lbcSelection(4).Move lbcSelection(6).Left, 280, lbcSelection(6).Width / 2 - 30, lbcSelection(6).Height
'                        lbcSelection(4).Visible = True
                        lbcSelection(12).Move lbcSelection(6).Left, 280, lbcSelection(6).Width / 2 - 120, lbcSelection(6).Height     '1-13-21
                        lbcSelection(12).Visible = True

                        'lbcSelection(8).Move lbcSelection(4).Width + 60, lbcSelection(4).Top
                        lbcSelection(8).Move lbcSelection(12).Width + 240, lbcSelection(12).Top, lbcSelection(12).Width             '3-18-16  items within vehicle group

                        lbcSelection(8).Visible = True
                        CkcAllveh.Caption = "All Items"
                        CkcAllveh.Move lbcSelection(8).Left, 0
                        CkcAllveh.Visible = True
                        CkcAllveh.Value = vbUnchecked
                        laclbcName(0).Caption = "Vehicle Group"
                        laclbcName(0).Move 120, 0, 1800
                        laclbcName(0).Visible = True

                    End If
                ElseIf ilListIndex = CNT_CUMEACTIVITY Then
                    ckcAll.Visible = True
                    lbcSelection(1).Visible = False                 'agy
                    lbcSelection(6).Visible = False                 'vehicle list box
                    lbcSelection(5).Visible = False                 'advt list box
                    lbcSelection(11).Visible = False                 'demo list box
                    plcSelC13.Visible = False
                    If Index = 0 Then                   'advt
                        lbcSelection(5).Move lbcSelection(6).Left, lbcSelection(6).Top, lbcSelection(6).Width, lbcSelection(6).Height
                        lbcSelection(5).Visible = True
                        ckcAll.Caption = "All Advertisers"
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        plcSelC13.Visible = True            '1-17-05 show include vehicle subtotal question
                        If RptSelCt!rbcSelC4(1).Value = True Then   'summary version, no vehicle subtotals
                            plcSelC13.Visible = False
                        End If
                    ElseIf Index = 1 Then                   'agy
                        lbcSelection(1).Move lbcSelection(6).Left, lbcSelection(6).Top, lbcSelection(6).Width, lbcSelection(6).Height
                        lbcSelection(1).Visible = True
                        ckcAll.Caption = "All Agencies"
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        plcSelC13.Visible = False            '1-17-05 hide include vehicle subtotal question
                    ElseIf Index = 2 Then                   'demo
                        lbcSelection(11).Move lbcSelection(6).Left, lbcSelection(6).Top, lbcSelection(6).Width, lbcSelection(6).Height
                        lbcSelection(11).Visible = True
                        ckcAll.Caption = "All Demos"
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        plcSelC13.Visible = False            '1-17-05 hide include vehicle subtotal question
                    ElseIf Index = 3 Then                    'vehicles
                        lbcSelection(6).Visible = True
                        ckcAll.Caption = "All Vehicles"
                        ckcAll.Value = vbUnchecked  '9-12-02 False
                        plcSelC13.Visible = False            '1-17-05 hide include vehicle subtotal question
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
                If ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION Then                        'proposals/contract
                    Screen.MousePointer = vbHourglass
                    mSetupPopAAS                               'setup list box of valid contracts
                    If rbcSelCSelect(0).Value Then
                        lbcSelection(0).Visible = True
                        lbcSelection(5).Visible = True
                        ckcAllAAS.Visible = True
                        ckcAllAAS.Caption = "All Advertisers"
                    ElseIf rbcSelCSelect(1).Value Then
                        lbcSelection(8).Visible = True      'agy list box for valid users
                        lbcSelection(10).Visible = False     '5-16-02 selected cntrs for valid users
                        lbcSelection(0).Visible = True       '5-16-02
                        ckcAllAAS.Visible = True
                        ckcAllAAS.Caption = "All Agencies"
                    ElseIf rbcSelCSelect(2).Value Then
                        lbcSelection(9).Visible = True      'slsp list box for valid users
                        lbcSelection(10).Visible = False    '5-16-02 True     'selected cnts for valid users
                        lbcSelection(0).Visible = True       '5-16-02

                        ckcAllAAS.Visible = True
                        ckcAllAAS.Caption = "All Salespeople"
                    End If
                    Screen.MousePointer = vbDefault
                ElseIf ilListIndex = CNT_PAPERWORK Then
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
                    If ilListIndex = CNT_PAPERWORK Then                         'paperwork summary
                        rbcSelCInclude(0).Value = True                  'default to contract vs line option
                        'Select by-
                        If rbcSelCSelect(3).Value Then              'by vehicle, contract or line options not available
                            plcSelC2.Visible = False                'disallow summary if by vehicle
                            rbcSelCInclude(1).Value = True          'default to line report
                            cbcSet1.Visible = False
                            edcSet1.Visible = False
                            ckcSelC5(7).Value = vbUnchecked         'disallow NTR with any vehicle/line option
                            ckcSelC5(7).Enabled = False
                        Else
                            plcSelC2.Visible = True
                            cbcSet1.Visible = True
                            edcSet1.Visible = True
                            ckcSelC5(7).Value = vbChecked         'turn NTR option back on and default it to obtain
                            ckcSelC5(7).Enabled = True
                            If rbcSelC9(3).Value = True Then            '10-8-13 if either select or sort is by vehicle, it has to default to detail
                                plcSelC2.Visible = False                'disallow summary if by vehicle
                                rbcSelCInclude(1).Value = True          'default to line report
                                cbcSet1.Visible = False
                                edcSet1.Visible = False
                                ckcSelC5(7).Value = vbUnchecked         'disallow NTR with any vehicle/line option
                                ckcSelC5(7).Enabled = False
                            End If
                            If rbcSelC7(2).Value Then               'show act cost only
                                rbcSelCInclude(1).Value = True      'default to show lines
                            End If
                        End If
                        Select Case Index
                            Case 0  'Advertiser/Contract #
                                lbcSelection(0).Visible = False
                                lbcSelection(6).Visible = False
                            Case 1  'Agency
                                lbcSelection(6).Visible = False
                            Case 2  'Salesperson
                                lbcSelection(6).Visible = False
                            Case 3  'vehicles
                                lbcSelection(2).Visible = False
                                rbcSelCInclude(1).Value = True          'force contract/line option to line
                        End Select

                    End If
                ElseIf ilListIndex = CNT_HISTORY Then
                    mSetupPopAAS
                ElseIf ilListIndex = CNT_BOB Then
                    If Index = 0 Then                       'all vehicles including pkg
                        mSellConvVVPkgPop 6, False                    'lbcselection(6), vehicles
                    Else                                    'show all vehicles excl hidden
                        mSellConvVirtVehPop 6, False
                    End If
                ElseIf ilListIndex = CNT_MAKEPLAN Then
                    edcSelCFrom_Change
                ElseIf ilListIndex = CNT_QTRLY_AVAILS Then
                    If Index = 3 Then   '% sellout avails, turn off by units or 30/60
                         'plcSelC11.Visible = False
                        '2-3-05 for sellout %, can choose between showing one percent value for all spots, or
                        '       a 30 vs 60 percent sellout
                        '10-20-11 rbcselc11 also used to indicate whether requesting "Units" vs 30/60
                        'If percent sellout, this was overriding that with the following question.
                        'But found not to be implemented.  Comment it out and allow user to pull based on units and 30/60
                        'In rptvfyct, a different .rpt was called to show 2 columns of 30% sellout vs 60% sellout.
                        'It was never called as the following code was forced to "combined"
'                        rbcSelC11(0).Move 945, 0, 1185
'                        rbcSelC11(1).Move 2250, 0, 2250
'                        rbcSelC11(0).Caption = "Combined"
'                        rbcSelC11(1).Caption = "30/60"
'                        rbcSelC11(0).Value = True
'                        plcSelC11.Move 120, plcSelC6.Top + plcSelC6.Height          'loc. of show 10,30 or 30/60
'                        smPaintCaption11 = "Sellout by"
'                        'plcSelC11.Visible = False
'                        '2-7-05 force show other missed spots separtely to NO
                        ckcSelC8(0).Value = vbUnchecked
                        ckcSelC8(0).Enabled = False
'                        plcSelC11.Visible = False       '3-17-05 disable 30/60 % sellout feature
                    Else
                        If rbcSelC4(0).Value = True Then        '3-8-05 qtrly avails (vs qtrly detail
                            'plcSelC11.Visible = True
                            rbcSelC11(0).Move 960, 0, 720
                            rbcSelC11(1).Move 1800, 0, 720
                            rbcSelC11(0).Caption = "Units"
                            rbcSelC11(1).Caption = "30/60"
                            plcSelC11.Move 120, plcSelC3.Top + plcSelC3.Height          '5-16-05 chg from plcselc6 to plcselc3; loc. of show 10,30 or 30/60
                            smPaintCaption11 = "Counts by"
                            plcSelC11.Visible = True
                            ckcSelC8(0).Value = vbChecked
                            ckcSelC8(0).Enabled = True
                        End If
                    End If
                ElseIf ilListIndex = CNT_AVGRATE Then       '9-28-11
                    If Index = 0 Then               'week
                        lacSelCTo1.Caption = "Quarter"
                        edcSelCTo1.MaxLength = 1
                        lacSelCFrom1.Visible = False
                        edcSelCFrom1.Visible = False
                    Else
                        lacSelCTo1.Caption = "Month"
                        edcSelCTo1.MaxLength = 2
                        lacSelCFrom1.Caption = "# Months"
                        edcSelCFrom1.MaxLength = 2
                        lacSelCFrom1.Move edcSelCTo1.Left + edcSelCTo1.Width + 240, lacSelCTo1.Top, lacSelCTo1.Width
                        edcSelCFrom1.Move lacSelCFrom1.Left + lacSelCFrom1.Width + 120, edcSelCTo1.Top
                        lacSelCFrom1.Visible = True
                        edcSelCFrom1.Visible = True
                    End If
                End If
            Case SLSPCOMMSJOB
                If ilListIndex = COMM_SALESCOMM Then
                    If Index = 0 Then
                        smPaintCaption11 = "Sort within Slsp by-"
                    Else
                        smPaintCaption11 = "Sort within Vehicle & Slsp by-"
                    End If
                    plcSelC11_Paint
                End If
        End Select
        mSetCommands
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub

Private Sub plcSelC12_Paint()
    plcSelC12.Cls
    plcSelC12.CurrentX = 0
    plcSelC12.CurrentY = 0
    plcSelC12.Print smPaintCaption12
End Sub

Private Sub plcSelC11_Paint()
    plcSelC11.Cls
    plcSelC11.CurrentX = 0
    plcSelC11.CurrentY = 0
    plcSelC11.Print smPaintCaption11
End Sub

Private Sub plcSelC10_Paint()
    plcSelC10.Cls
    plcSelC10.CurrentX = 0
    plcSelC10.CurrentY = 0
    plcSelC10.Print smPaintCaption10
End Sub

Private Sub plcSelC9_Paint()
    plcSelC9.Cls
    plcSelC9.CurrentX = 0
    plcSelC9.CurrentY = 0
    plcSelC9.Print smPaintCaption9
End Sub

Private Sub plcSelC8_Paint()
    plcSelC8.Cls
    plcSelC8.CurrentX = 0
    plcSelC8.CurrentY = 0
    plcSelC8.Print smPaintCaption8
End Sub

Private Sub plcSelC7_Paint()
    plcSelC7.Cls
    plcSelC7.CurrentX = 0
    plcSelC7.CurrentY = 0
    plcSelC7.Print smPaintCaption7
End Sub

Private Sub plcSelC6_Paint()
    plcSelC6.Cls
    plcSelC6.CurrentX = 0
    plcSelC6.CurrentY = 0
    plcSelC6.Print smPaintCaption6
End Sub

Private Sub plcSelC5_Paint()
    plcSelC5.Cls
    plcSelC5.CurrentX = 0
    plcSelC5.CurrentY = 0
    plcSelC5.Print smPaintCaption5
End Sub

Private Sub plcSelC3_Paint()
    plcSelC3.Cls
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    plcSelC3.Print smPaintCaption3
End Sub

Private Sub plcSelC2_Paint()
    plcSelC2.Cls
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print smPaintCaption2
End Sub

Private Sub plcSelC1_Paint()
    plcSelC1.Cls
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    plcSelC1.Print smPaintCaption1
End Sub

Private Sub plcSelC4_Paint()
    plcSelC4.Cls
    plcSelC4.CurrentX = 0
    plcSelC4.CurrentY = 0
    plcSelC4.Print smPaintCaption4
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCntSelectivity2                *
'*                                                     *
'*            Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:                                *
'*                                                     *
'*******************************************************
Private Sub mCntSelectivity2()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slAirOrder                                                                            *
'******************************************************************************************

    Dim ilListIndex As Integer
    Dim ilRet As Integer
    ReDim ilAASCodes(0 To 1) As Integer
    Dim illoop As Integer
    Dim ilFound As Integer
    Dim ilTop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefIndex As Integer
    Dim slStr As String
    Dim ilShowNone As Integer
    Dim llStartDate As Long
    Dim ilDay As Integer
    Dim llTodayDate As Long
    Dim llActiveDate As Long

    ilListIndex = lbcRptType.ListIndex
    If (igRptType = 0) And (ilListIndex > 1) Then
        ilListIndex = ilListIndex + 1
    End If
    '2020-10-30 - TTP # 9955 - add AirTime, NTR, Hardcost Include options to report (Daily Sales Activity by Contract Report, Weekly Sales Activity by Qtr)
    If ilListIndex = CNT_DAILY_SALESACTIVITY Or ilListIndex = CNT_SALESACTIVITY Then
        plcSelC13.Move 120, plcSelC2.Top + 60, 4000
        ckcSelC13(0).Caption = "Air Time"
        ckcSelC13(1).Caption = "NTR"
        ckcSelC13(2).Caption = "Hard Cost"
        plcSelC13.Visible = True
        smPaintCaption13 = "Include"
        plcSelC13_Paint
        ckcSelC13(0).Value = vbChecked
        ckcSelC13(1).Value = vbChecked
        ckcSelC13(2).Value = vbChecked
        ckcSelC13(0).Move 840, 0, 1080
        ckcSelC13(1).Move 1920, 0, 720
        ckcSelC13(2).Move 2640, 0, 1200
        ckcSelC13(0).Visible = True
        ckcSelC13(1).Visible = True
        ckcSelC13(2).Visible = True
        
        'IF include NTR, Add option to split  (or by default: leave grouped together)
        smPaintCaption10 = ""
        plcSelC10_Paint
        plcSelC10.Move 0, plcSelC13.Top + plcSelC13.Height + 30, 4000
        ckcSelC10(0).Move 960, 0, 4000
        ckcSelC10(0).Caption = "Separate Air Time and NTR/HC"
        ckcSelC10(0).Value = vbUnchecked
        ckcSelC10(0).Visible = True
        ckcSelC10(1).Visible = False
        If ckcSelC13(1).Value = vbChecked Or ckcSelC13(2).Value = vbChecked Then
            plcSelC10.Visible = True
        End If
    End If
    
    If ilListIndex = CNT_CUMEACTIVITY Then
        '09/29/2020 - TTP # 9952 - IF include NTR, Add option to split NTR (or by default: leave NTR grouped together)
        plcSelC14.Move 120, plcSelC13.Top + plcSelC13.Height, 4000
        rbcSelC14(0).Caption = "Grouped"
        rbcSelC14(1).Caption = "Split"
        smPaintCaption14 = "Show NTR"
        plcSelC14_Paint
        rbcSelC14(0).Value = True
        'rbcSelC14(0).Move 840, 0, 1080
        'rbcSelC14(1).Move 1920, 0, 720
        rbcSelC14(0).Visible = True
        rbcSelC14(1).Visible = True
        plcSelC14.Visible = False
    End If
    If ilListIndex = CNT_DAILY_SALESACTIVITY Then     '6-5-01
        lacSelCTo.Caption = "Activity Dates- Start"
        lacSelCTo.Visible = True
        lacSelCTo.Width = 2640
        lacSelCTo.Left = 120
        lacSelCTo.Top = 105
        lacSelCTo1.Left = 3000
        lacSelCTo1.Caption = "End"
        lacSelCTo1.Width = 360
        lacSelCTo1.Top = 105
        lacSelCTo1.Visible = True
        
        'Date: 11/22/2019 added CSI calenddar control for date entries
'        edcSelCTo.MaxLength = 10
'        edcSelCTo1.MaxLength = 10
'        edcSelCTo.Move 1800, 60, 1080
'        edcSelCTo1.Move 3380, 60, 1080
'        edcSelCTo.Visible = True
'        edcSelCTo1.Visible = True
        
        edcSelCTo.Visible = False
        edcSelCTo1.Visible = False
        
        CSI_CalFrom.Visible = True
        CSI_CalFrom.Left = 1050
        CSI_CalFrom.Width = 1170
        CSI_CalFrom.Move 1800, 60, 1080
        CSI_CalFrom.ZOrder 0
        
        CSI_CalTo.Visible = True
        CSI_CalTo.Left = 1050
        CSI_CalTo.Width = 1170
        CSI_CalTo.Move 3380, 60, 1080
        CSI_CalTo.ZOrder 0
        
        pbcSelC.Width = pbcSelC.Width + 800
        
        'sort by advt or sales office
        'plcSelC4.Caption = "Sort By"
        smPaintCaption4 = "Sort by"
        plcSelC4_Paint
        plcSelC4.Move 120, CSI_CalTo.Top + CSI_CalTo.Height + 40 'Date: 11/22/2019 added CSI calendar control for date entries --> edcSelCTo.Top + edcSelCTo.Height + 60, 3600
        rbcSelC4(0).Move 720, 0, 1200    'Advt button,
        rbcSelC4(0).Caption = "Advertiser"
        rbcSelC4(0).Visible = True
        rbcSelC4(0).Value = True
        rbcSelC4(1).Move 1980, 0, 1560   'Sales Office button
        rbcSelC4(1).Caption = "Sales Office"
        rbcSelC4(1).Visible = True
        rbcSelC7(2).Visible = False
        plcSelC4.Visible = True
        '10-4-13 add net option
        plcSelC7.Move 120, plcSelC4.Top + plcSelC4.Height + 30
        mAskGrossOrNet
        'most default to Net, default this report to gross since it deals with projections
        rbcSelC7(0).Value = True
        rbcSelC7(2).Visible = False     'net-net not applicable


     ElseIf ilListIndex = CNT_SALESANALYSIS Then
        mSPersonPop lbcSelection(2)        'Populate slsp to obtain the closest rollover date date against the user entered date
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        edcSelCFrom.Text = ""
        edcSelCTo.Text = ""
        edcSelCTo1.Text = ""
        ckcAll.Visible = False
        
        'Date: 12/11/2109 added CSI calendar control for date entry
'        CSI_CalFrom.Move 1340, edcSelCFrom.Top, 1080
'        CSI_CalFrom.Visible = True
'        CSI_CalFrom.ZOrder 0
        
        mAskEffDate
        
        edcSelCFrom.Visible = False
        
        plcSelC2.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
        'plcSelC2.Height = 440
        'plcSelC2.Caption = "Month"
        smPaintCaption2 = "Month"
        plcSelC2_Paint
        plcSelC2.Visible = True
        rbcSelCInclude(0).Caption = "Corporate"
        rbcSelCInclude(0).Move 660, 0, 1140
        rbcSelCInclude(0).Visible = True
        rbcSelCInclude(1).Caption = "Standard"
        rbcSelCInclude(1).Move 1840, 0, 1140
        rbcSelCInclude(1).Visible = True
        If rbcSelCInclude(1).Value Then             'default to std
            rbcSelCInclude_Click 1
        Else
            rbcSelCInclude(1).Value = True
        End If
        If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
            rbcSelCInclude(0).Enabled = False
        Else
            rbcSelCInclude(0).Value = True
        End If
        rbcSelCInclude(2).Visible = False

        gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), False
        gPopVehicleGroups RptSelCt!cbcSet2, tgVehicleSets2(), True
        cbcSet1.ListIndex = 0
        cbcSet2.ListIndex = 0
        plcSelC3.Move 120, plcSelC1.Top + plcSelC2.Height
        'for ntr option
        edcSet1.Move 120, plcSelC3.Top + plcSelC3.Height + 90
        'edcSet1.Move 120, plcSelC2.Top + plcSelC2.Height + 90
        cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 30
        edcSet2.Move 120, cbcSet1.Top + cbcSet1.Height + 90
        cbcSet2.Move cbcSet1.Left, edcSet2.Top - 30
        edcSet1.Visible = True
        edcSet2.Visible = True
        cbcSet1.Visible = True
        cbcSet2.Visible = True
        'Dan M added contract selectivity and ntr/hard cost option 7-16-08
        ckcSelC3(0).Caption = "NTR"
        ckcSelC3(0).Width = 750
        ckcSelC3(0).Visible = True
        ckcSelC3(1).Caption = "Hard Cost"
        smPaintCaption3 = "Include"
        plcSelC3_Paint
        ckcSelC3(0).Left = 750
        ckcSelC3(1).Left = ckcSelC3(0).Left + ckcSelC3(0).Width + 30
        ckcSelC3(1).Width = 2000
        ckcSelC3(1).Visible = True
        plcSelC3.Visible = True
        lacTopDown.Caption = "Contract"
        lacTopDown.Move 120, edcSet2.Top + edcSet2.Height + 100, 900
        edcText.Move lacTopDown.Left + lacTopDown.Width, edcSet2.Top + edcSet2.Height + 50
        lacTopDown.Visible = True
        edcText.Visible = True

    ElseIf ilListIndex = CNT_SALESACTIVITY_SS Or ilListIndex = CNT_SALESPLACEMENT Then       '7-25-02
        sgSalespersonTag = ""           'ensure the Sales Source list box is populated and doesn't have incorrect items
        ilTop = 60
        If ilListIndex = CNT_SALESACTIVITY_SS Then      'Sales Placements gets the latest mods of all contracts, no need to ask activity dates
            lacSelCFrom.Caption = "Dates- Start"
            lacSelCFrom.Visible = True
            lacSelCFrom.Move 15, 105, 1680
            edcSelCFrom.MaxLength = 10
            edcSelCFrom.Move 1080, 60, 1080

            lacSelCFrom1.Caption = "End"

            lacSelCFrom1.Move 2300, lacSelCFrom.Top, 360
            lacSelCFrom1.Visible = True

            edcSelCFrom1.MaxLength = 10
            edcSelCFrom1.Move 3000, edcSelCFrom.Top, 1080
            edcSelCFrom.Visible = True
            edcSelCFrom1.Visible = True

            rbcSelCSelect(0).Caption = "Original Entry Date"
            rbcSelCSelect(1).Caption = "Latest Mod Date"
            'plcSelC1.Caption = "Use"
            smPaintCaption1 = "Use"
            
            plcSelC1.Move 15, edcSelCFrom.Top + edcSelCFrom.Height + 30, 4290            'rbcSelCSelect(0).Value = False
            rbcSelCSelect(1).Value = True
            rbcSelCSelect(0).Visible = True
            rbcSelCSelect(1).Visible = True
            rbcSelCSelect(2).Visible = False
            rbcSelCSelect(0).Move 360, 0, 1920
            rbcSelCSelect(1).Move 2280, 0, 2180
            plcSelC1.Visible = True
            ilTop = plcSelC1.Top + plcSelC1.Height + 30
        End If
        lacSelCTo.Caption = "Year"
        lacSelCTo.Visible = True
        lacSelCTo.Move 15, ilTop, 480
        edcSelCTo.Move 570, ilTop - 30, 720
        edcSelCTo.MaxLength = 4

        lacSelCTo1.Move 1620, ilTop, 600
        lacSelCTo1.Caption = "Month"
        lacSelCTo1.Visible = True
        edcSelCTo1.MaxLength = 3
        edcSelCTo1.Move 2220, lacSelCTo.Top - 30, 600
        edcSelCTo.Visible = True
        edcSelCTo1.Visible = True

        lacPeriods.Move 3060, ilTop, 1380
        lacPeriods.Caption = "# Periods"
        edcText.Move 3900, edcSelCTo1.Top, 360
        edcText.MaxLength = 2
        lacPeriods.Visible = True
        edcText.Visible = True
        
        'Date: 11/26/2019   added CSI calendar controls for date entries
'        edcSelCFrom.Visible = False
'        edcSelCFrom1.Visible = False
'
'        CSI_CalFrom.Visible = True
'        CSI_CalFrom.Move edcSelCFrom.Left, 60, 1080
'        CSI_CalFrom.ZOrder 0
'
'        CSI_CalTo.Visible = True
'        CSI_CalTo.Move lacSelCFrom1.Left + lacSelCFrom1.Width + 10, 60, 1080
'        CSI_CalTo.ZOrder 0
        
        '3-5-10 add additional months to gather by :  corp & cal
        If ilListIndex = CNT_SALESACTIVITY_SS Then          'allow different types of month for Sales Activity,
            'Date: 11/26/2019   added CSI calendar controls for date entries
            edcSelCFrom.Visible = False
            edcSelCFrom1.Visible = False
            
            CSI_CalFrom.Visible = True
            CSI_CalFrom.Move edcSelCFrom.Left, 60, 1080
            CSI_CalFrom.ZOrder 0
            
            CSI_CalTo.Visible = True
            CSI_CalTo.Move lacSelCFrom1.Left + lacSelCFrom1.Width + 10, 60, 1080
            CSI_CalTo.ZOrder 0
                                                            'currently disallow for Sales Placeme
            mAskTypeOfMonth smPaintCaption2, rbcSelCInclude(0), rbcSelCInclude(1), rbcSelCInclude(2)
            plcSelC2.Visible = True
            plcSelC2.Move 0, edcSelCTo.Top + edcSelCTo.Height + 30
            plcSelC7.Move 0, plcSelC2.Top + plcSelC2.Height + 15
        Else
            plcSelC2.Visible = False
            rbcSelCInclude(1).Value = True              'default to standard bdcst month and hide the feature
            plcSelC7.Move 0, edcSelCTo.Top + edcSelCTo.Height + 30
        End If

        mAskGrossOrNet
        If ilListIndex = CNT_SALESPLACEMENT Then        'allow net-net option
            rbcSelC7(2).Caption = "T-Net"
            rbcSelC7(2).Visible = True
        End If
  
        gPopVehicleGroups RptSelCt!cbcSet1, tgVehicleSets1(), True
        edcSet1.Text = "Vehicle Group"
        edcSet1.Move 0, plcSelC7.Top + 30 + plcSelC7.Height
        cbcSet1.Move 360 + edcSet1.Width, edcSet1.Top - 45

        ilFound = False
        'first time thru, default to market if client has markets defined, else default to NONE
        'For ilLoop = 1 To UBound(tgVehicleSets1)
        For illoop = LBound(tgVehicleSets1) To UBound(tgVehicleSets1)
            If Trim$(tgVehicleSets1(illoop).sChar) = "Market" Then
                ilFound = True
                Exit For
            End If
        Next illoop

        If ilFound Then
            cbcSet1.ListIndex = illoop
        Else
            cbcSet1.ListIndex = 0
        End If

        'show the sort options
        plcSelC9.Move 0, edcSet1.Top + edcSet1.Height + 30, 3960, 645
        'plcSelC9.Caption = "Sort by "
        smPaintCaption9 = "Sort By"
        rbcSelC9(0).Move 600, 0, 3960
        rbcSelC9(1).Move 600, 195, 3960
        rbcSelC9(2).Move 600, 390, 3960
        rbcSelC9(1).Visible = True
        rbcSelC9(0).Visible = True
        rbcSelC9(2).Visible = True
        rbcSelC9(0).Value = True
        plcSelC9.Visible = True
        edcSet1.Visible = True
        cbcSet1.Visible = True

        plcSelC11.Move 0, plcSelC9.Top + plcSelC9.Height
        'plcSelC11.Caption = ""
        rbcSelC11(0).Caption = "Detail"
        rbcSelC11(1).Caption = "Summary"
        rbcSelC11(0).Move 0, 0, 840
        rbcSelC11(1).Move 840, 0, 1200
        rbcSelC11(0).Visible = True
        rbcSelC11(1).Visible = True
        rbcSelC11(0).Value = True
        plcSelC11.Visible = True

        'lacTopDown.Move 120, plcSelC7.Top + plcSelC7.Height + 30, 1200
        lacTopDown.Move 0, plcSelC11.Top + plcSelC11.Height + 30, 1200
        lacTopDown.Caption = "Contract #"
        edcTopHowMany.Move 1080, lacTopDown.Top - 30, 1080
        edcTopHowMany.MaxLength = 9
        edcTopHowMany.Text = ""
        lacTopDown.Visible = True
        edcTopHowMany.Visible = True

        plcSelC12.Move 0, edcTopHowMany.Top + edcTopHowMany.Height
        'plcSelC12.Caption = ""
        smPaintCaption12 = ""
        ckcSelC12(0).Move 0, 0, 3600
        ckcSelC12(0).Value = vbUnchecked    '9-12-02 False
        ckcSelC12(0).Visible = True
        ckcSelC12(0).Caption = "Skip to new page each new group"
        plcSelC12.Visible = True

        If ilListIndex = CNT_SALESACTIVITY_SS Then
            plcSelC13.Move 0, plcSelC12.Top + plcSelC12.Height, 4000
            ckcSelC13(0).Caption = "Air Time"
            ckcSelC13(1).Caption = "NTR"
            ckcSelC13(2).Caption = "Hard Cost"
            smPaintCaption13 = "Include"
            plcSelC13_Paint
            ckcSelC13(0).Value = vbChecked
            ckcSelC13(1).Value = vbUnchecked
            ckcSelC13(2).Value = vbUnchecked
            ckcSelC13(0).Move 840, 0, 1080
            ckcSelC13(1).Move 1920, 0, 720
            ckcSelC13(2).Move 2640, 0, 1200
            ckcSelC13(0).Visible = True
            ckcSelC13(1).Visible = True
            ckcSelC13(2).Visible = True
            plcSelC13.Visible = True
            
            '09/28/2020 - TTP # 9952 - IF include NTR, Add option to split NTR (or by default: leave NTR grouped together)
            plcSelC14.Move 0, plcSelC13.Top + plcSelC13.Height, 4000
            rbcSelC14(0).Caption = "Grouped"
            rbcSelC14(1).Caption = "Split"
            smPaintCaption14 = "Show NTR"
            plcSelC14_Paint
            rbcSelC14(0).Value = True
            'rbcSelC14(0).Move 840, 0, 1080
            'rbcSelC14(1).Move 1920, 0, 720
            rbcSelC14(0).Visible = True
            rbcSelC14(1).Visible = True
            plcSelC14.Visible = False
        End If

        
        'Sales Source selections:  Sales Source & Sales Office
        ckcAll.Caption = "All Sales Sources"
        ckcAll.Move 15, 0
        ckcAll.Visible = True
        ckcAllAAS.Caption = "All Sales Offices"
        ckcAllAAS.Move lbcSelection(3).Width / 2, 0
        ckcAllAAS.Visible = True
        ilRet = gPopMnfPlusFieldsBox(RptSelCt, RptSelCt!lbcSelection(3), tgSalesperson(), sgSalespersonTag, "S")
        lbcSelection(3).Move 15, ckcAll.Top + ckcAll.Height + 30, lbcSelection(3).Width / 2 - 120, 1500

        mSalesOfficePop lbcSelection(2)
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        lbcSelection(2).Move ckcAllAAS.Left, lbcSelection(3).Top, lbcSelection(3).Width + 120, 1500
        lbcSelection(2).Visible = True
        lbcSelection(3).Visible = True
        'Distribution  :  Markets if exists, else vehicles
        'mMktOrVefPop '
        CkcAllveh.Move 15, lbcSelection(2).Top + lbcSelection(2).Height + 30
        lbcSelection(6).Move 15, CkcAllveh.Top + CkcAllveh.Height + 30
        lbcSelection(6).Height = lbcSelection(6).Height / 2
        'lbcSelection(6).Visible = True
    ElseIf ilListIndex = CNT_VEH_UNITCOUNT Then       '7-15-03
        mStartEndDates
        '2-1-06 option to get export type
        plcSelC3.Move lacSelCFrom.Left, edcSelCFrom.Top + edcSelCFrom.Height + 30
        smPaintCaption3 = "Include"
        ckcSelC3(0).Move 840, 0, 960
        ckcSelC3(0).Caption = "Manual"
        ckcSelC3(0).Value = vbChecked
        ckcSelC3(1).Move 1920, 0, 720
        ckcSelC3(1).Caption = "Web"
        ckcSelC3(1).Value = vbChecked
        ckcSelC3(2).Move 2720, 0, 1200
        ckcSelC3(2).Caption = "Marketron"
        ckcSelC3(2).Value = vbChecked
        ckcSelC3(0).Visible = True
        ckcSelC3(1).Visible = True
        ckcSelC3(2).Visible = True
        plcSelC3.Visible = True
        mSellConvAirVehPop                          'use lbcSelection(3) for selling, air, conv vehicle
        lbcSelection(3).Visible = True
        ckcAll.Caption = "All Vehicles"
        ckcAll.Move 120, 0
        ckcAll.Visible = True
    ElseIf ilListIndex = CNT_LOCKED Or ilListIndex = CNT_GAMESUMMARY Then            '4-5-06 locked avails report, 7-14-06 Game summary
        mStartEndDates
        plcSelC4.Move 0, edcSelCFrom.Top + edcSelCFrom.Height + 30
        smPaintCaption4 = "Sort By"
        rbcSelC4(0).Move 720, 0, 1080
        rbcSelC4(0).Caption = "Vehicle"
        rbcSelC4(1).Move 1800, 0, 720
        rbcSelC4(1).Caption = "Date"
        rbcSelC4(0).Value = True
        rbcSelC4(0).Visible = True
        rbcSelC4(1).Visible = True
        rbcSelC4(2).Visible = False
        'option to exclude cancelled games added 5/13/08 (located in 'else' below)

        If ilListIndex = CNT_LOCKED Then
            mSellConvVehPop 3           'get conv and selling
        Else                            'sports vehicles
            mSportsVehPop 3
            ckcSelC10(0).Caption = "Include Cancelled"
            ckcSelC10(0).Move 0, 0, 2500
            ckcSelC10(0).Value = 1
            plcSelC10.Move 0, plcSelC4.Top + plcSelC4.Height + 30, 2500
            ckcSelC10(0).Visible = True
            plcSelC10.Visible = True
            plcSelC4.Visible = True
        End If
        lbcSelection(3).Visible = True
        ckcAll.Caption = "All Vehicles"
        ckcAll.Move 15, 0
        ckcAll.Visible = True
    ElseIf ilListIndex = CNT_PAPERWORKTAX Then          '4-9-07
        lbcSelection(11).Clear
        ReDim tgSellNameCode(0 To 0) As SORTCODE
        'loop thru the vehicles already populated in box and repopulate in another box
        For illoop = 0 To lbcSelection(6).ListCount - 1 Step 1
            slNameCode = tgCSVNameCode(illoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefIndex = gBinarySearchVef(Val(slCode))
            If ilVefIndex >= 0 Then
                If tgMVef(ilVefIndex).iTrfCode > 0 Then         'its taxable vehicle
                    slStr = lbcSelection(6).List(illoop)
                    lbcSelection(11).AddItem slStr
                    tgSellNameCode(UBound(tgSellNameCode)).sKey = tgCSVNameCode(illoop).sKey
                    ReDim Preserve tgSellNameCode(0 To UBound(tgSellNameCode) + 1) As SORTCODE
                End If
            End If
        Next illoop
        
        mStartEndDates
        smPaintCaption3 = ""
        plcSelC3.Move lacSelCFrom.Left, edcSelCFrom.Top + edcSelCFrom.Height + 90
        ckcSelC3(0).Caption = "Skip Page Each Vehicle"
        ckcSelC3(0).Move lacSelCFrom.Left, 0, 3120
        ckcSelC3(0).Visible = True
        plcSelC3.Visible = True
        lacTopDown.Move plcSelC3.Left, plcSelC3.Top + plcSelC3.Height + 90, 1200
        lacTopDown.Caption = "Contract #"
        edcTopHowMany.Move 1080, lacTopDown.Top - 30, 1080
        edcTopHowMany.MaxLength = 9
        edcTopHowMany.Text = ""
        lacTopDown.Visible = True
        edcTopHowMany.Visible = True
        lbcSelection(11).Move lbcSelection(6).Left, lbcSelection(6).Top, lbcSelection(6).Width
        lbcSelection(11).Visible = True
        
        'Date: 12/20/2019 make sure order of display is set to top
        'lbcSelection(11).ZOrder 0
        
        ckcAll.Visible = True
        ckcAll.Caption = "All Vehicles"
    ElseIf ilListIndex = CNT_BOBCOMPARE Then
        lbcSelection(1).Move 15, 280, 4000, 1500  'agy list box
        lbcSelection(2).Move 15, 280, 4000, 1500  'slsp list box
        lbcSelection(3).Move 15, 280, 4000, 1500  'bus cat list box
        lbcSelection(7).Move 15, 280, 4000, 1500  'prod prot list box
        lbcSelection(5).Move 15, 280, 4000, 1500 'advt list box
        lbcSelection(6).Move 15, 280, 4000, 1500  'vehicle list box
        lbcSelection(8).Move 15, 280, 4000, 1500  'vehicle group items list box
        lbcSelection(4).Move 2000, 2100, 2000, 1500  'budgetlist box, single selection

        mBudgetPop                                      'lbcselection(4), one budget only
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mAdvtPop lbcSelection(5)
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If

        mMnfPop "B", RptSelCt!lbcSelection(3), tgMNFCodeRpt(), sgMNFCodeTagRpt    'Traffic!lbcSalesperson
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If

        mMnfPop "C", RptSelCt!lbcSelection(7), tgMnfCodeCT(), sgMNFCodeTagRpt    'Traffic!lbcSalesperson
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If

        mAskYrMonthPeriods ilListIndex    'ask year, Month, # Periods
        ckcAll.Move 15, 0
        ckcAll.Visible = True

         'For ilLoop = 1 To 2
         '   If ilLoop = 1 Then
        ilShowNone = False
        gPopVehicleGroups cbcSet1, tgVehicleSets1(), True
        lbcSelection(5).Visible = True
        '    Else
                ilShowNone = False
                mFillSalesCompare ilListIndex, cbcSet2, ilShowNone
        '    End If
        'Next ilLoop
        If rbcSelCInclude(0).Value Then             'advt defaulted, show the advt list box
            rbcSelCInclude_Click 0
        Else
            rbcSelCInclude(0).Value = True
        End If
        edcSet1.Text = "Sort- Group"
        edcSet1.Move 120, edcSelCFrom1.Top + edcSelCFrom1.Height + 90, 1080
        cbcSet1.Move edcSet1.Left + edcSet1.Width, edcSelCFrom.Top + edcSelCFrom1.Height + 60, 1260
        cbcSet1.Visible = True
        edcSet1.Visible = True

        edcSet2.Text = "Minor"
        edcSet2.Move cbcSet1.Left + cbcSet1.Width + 120, edcSet1.Top, 600
        cbcSet2.Move edcSet2.Left + edcSet2.Width, cbcSet1.Top, 1260
        cbcSet2.Visible = True
        edcSet2.Visible = True
        plcSelC13.Move 120, cbcSet1.Top + cbcSet1.Height, 4380, 480
        'ckcSelC13(0).Move 0, 0, 2040
        'ckcSelC13(0).Caption = "Advertiser totals"
        '* ** * DO NOT * * *  use ckcSelC13(0) for this report; used in sales compare
        ckcSelC13(1).Move 0, 0, 4000
        ckcSelC13(1).Caption = "Show politicals as separate group"
        ckcSelC13(0).Visible = False
        ckcSelC13(1).Visible = True
        ckcSelC13(2).Move 0, 240, 4000
        ckcSelC13(2).Caption = "Use Sales Source as major sort"
        ckcSelC13(2).Value = vbUnchecked              'default to use sales source as the major sort
        ckcSelC13(2).Visible = True
        plcSelC13.Visible = True

        plcSelC9.Move 120, plcSelC13.Top + plcSelC13.Height + 30
        mAskBOBCorpOrStd
        If tgSpf.sRUseCorpCal = "Y" Then        'default to corp cal if defined
            rbcSelC9(0).Value = True
        End If
        plcSelC7.Move 120, plcSelC9.Top + plcSelC9.Height + 30
        mAskGrossOrNet
        rbcSelC7(2).Caption = "T-Net"
        rbcSelC7(2).Visible = True
        plcSelC7.Width = 3000

        lacSelCTo1.Caption = "Cnt#"
        lacSelCTo1.Move 3120, plcSelC7.Top, 600
        edcSelCTo1.Move 3600, plcSelC7.Top - 30, 820
        lacSelCTo1.Visible = True
        edcSelCTo1.Visible = True
        edcSelCTo1.MaxLength = 9                '1-30-06


        plcSelC3.Move plcSelC7.Left, plcSelC7.Top + plcSelC7.Height + 60
        mAskContractTypes

        plcSelC1.Move 120, plcSelC12.Top + plcSelC12.Height

        mAskPkgOrHide ilListIndex

        ckcSelC8(2).Visible = False
        ckcSelC8(0).Value = vbUnchecked 'False
        ckcSelC8(1).Value = vbChecked   'True
        If tgSpf.sInvAirOrder = "S" Then        'bill as ordred, aired
            ckcSelC8(0).Visible = False
            ckcSelC8(1).Visible = False
            ckcSelC8(0).Value = vbUnchecked 'False       'ignore missed
            ckcSelC8(1).Value = vbUnchecked 'False       'ignore mgs
        Else                                'as aired
            ckcSelC8(0).Visible = True
            ckcSelC8(1).Visible = True
        End If
        plcSelC8.Height = 480

        lbcSelection(4).Visible = True
        lbcSelection(4).Enabled = False     'disable until vehiclegroup selected as option
        laclbcName(0).Caption = "Budgets"
        laclbcName(0).Move 2000, lbcSelection(1).Top + lbcSelection(1).Height + 60, 1800
        laclbcName(0).Visible = True
        lbcSelection(6).Move 15, 2100, 1800, 1500
        lbcSelection(6).Visible = True
        ckcAllAAS.Move 15, laclbcName(0).Top
        ckcAllAAS.Caption = "All Vehicles"
        ckcAllAAS.Visible = True
    ElseIf ilListIndex = CNT_CONTRACTVERIFY Then           '4-8-13
        lacSelCFrom.Left = 120
        lacSelCFrom1.Move 2325, 75
        edcSelCFrom.Move 1290, edcSelCFrom.Top, 945
        edcSelCFrom1.Move 2700, edcSelCFrom.Top, 945
        edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
        edcSelCFrom1.MaxLength = 10 '8  5/27/99 changed for short form date m/d/yyyy
        lacSelCFrom.Caption = "Active: From"
        lacSelCFrom1.Caption = "To"
        lacSelCFrom.Visible = True
        lacSelCFrom1.Visible = True
        edcSelCFrom.Visible = True
        edcSelCFrom1.Visible = True
        plcSelC3.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30, 4380, 495
        ckcSelC3(0).Caption = "Not Verified"
        ckcSelC3(1).Caption = "Sent to Agy"
        ckcSelC3(2).Caption = "Verified"
        ckcSelC3(0).Move 720, 0, 1680
        ckcSelC3(1).Move 2400, 0, 1680
        ckcSelC3(2).Move 720, 255, 1299
        ckcSelC3(0).Visible = True
        ckcSelC3(1).Visible = True
        ckcSelC3(2).Visible = True
        ckcSelC3(0).Value = vbChecked
        ckcSelC3(1).Value = vbChecked
        ckcSelC3(2).Value = vbChecked
        plcSelC3.Visible = True
        smPaintCaption3 = "State"
        plcSelC3_Paint
    ElseIf ilListIndex = CNT_INSERTION_ACTIVITY Then        '10-6-15
        edcSelCTo.MaxLength = 10
        edcSelCTo1.MaxLength = 10
        edcSelCFrom.MaxLength = 10
        edcSelCFrom1.MaxLength = 10
        
        'sent start/end dates
        slStr = Format$(gNow(), "m/d/yy")
        llStartDate = gDateValue(slStr)
        llActiveDate = gDateValue(slStr)
        llTodayDate = gDateValue(slStr)
        'backup to Monday
        ilDay = gWeekDayLong(llStartDate)
        'if today is already a monday, backup to last week
        If ilDay = 0 Then
            ilDay = -1
        End If
        Do While ilDay <> 0
            llStartDate = llStartDate - 1
            ilDay = gWeekDayLong(llStartDate)
        Loop
        
        ilDay = gWeekDayLong(llActiveDate)
        Do While ilDay <> 0
            llActiveDate = llActiveDate - 1
            ilDay = gWeekDayLong(llActiveDate)
        Loop
        
        edcSelCFrom.Text = Format(llStartDate, "m/d/yy")     'sent startdate
        edcSelCFrom1.Text = Format(llTodayDate, "m/d/yy")    'sent end date
        edcSelCTo.Text = Format(llActiveDate, "m/d/yy")       'active start date
        edcSelCTo1.Text = Format(llActiveDate + 6, "m/d/yy")  'active end date
        
        'Date: 12/18/2019 added CSI calendar control for date entries
        CSI_CalFrom.Text = Format(llStartDate, "m/d/yy")    'sent startdate
        CSI_CalTo.Text = Format(llTodayDate, "m/d/yy")    'sent end date
        CSI_From1.Text = Format(llActiveDate, "m/d/yy")       'active start date
        CSI_To1.Text = Format(llActiveDate + 6, "m/d/yy")  'active end date
        
        lacSelCFrom.Move 120, 60, 1530
        edcSelCFrom.Move 1650, edcSelCFrom.Top, 945
        
        lacSelCFrom1.Move 2800  ', 75
        edcSelCFrom1.Move 3060, edcSelCFrom.Top, 945
        
        'active start/end dates
        lacSelCTo.Move 120, edcSelCFrom1.Top + edcSelCFrom1.Height + 120, 1530
        edcSelCTo.Move 1650, lacSelCTo.Top - 60, 945
        
        lacSelCTo1.Move 2800, lacSelCTo.Top
        edcSelCTo1.Move 3060, edcSelCTo.Top, 945
        
        lacSelCFrom.Caption = "Sent Date- From"
        lacSelCFrom1.Caption = "To"
        lacSelCFrom.Visible = True
        lacSelCFrom1.Visible = True
        lacSelCTo.Caption = "Active Date- From"
        lacSelCTo1.Caption = "To"
        lacSelCTo.Visible = True
        lacSelCTo1.Visible = True
        edcSelCFrom.Visible = True
        edcSelCFrom1.Visible = True
        edcSelCTo.Visible = True
        edcSelCTo1.Visible = True
                
        'Date: 12/18/2019 added CSI calendar control for date entry
        CSI_CalFrom.Visible = True: CSI_CalTo.Visible = True: CSI_From1.Visible = True: CSI_To1.Visible = True
        CSI_CalFrom.Move lacSelCFrom.Left + lacSelCFrom.Width + 10, lacSelCFrom.Top, 1080
        CSI_CalTo.Move lacSelCFrom1.Left + lacSelCFrom1.Width + 10, lacSelCFrom1.Top, 1080
        CSI_From1.Move lacSelCTo.Left + lacSelCTo.Width + 10, lacSelCTo.Top, 1080
        CSI_To1.Move lacSelCTo1.Left + lacSelCTo1.Width + 10, lacSelCTo1.Top, 1080
        CSI_CalTo.Left = CSI_To1.Left
        
        CSI_From1.ZOrder 0: CSI_To1.ZOrder 0
        CSI_CalTo.ZOrder 0: CSI_CalFrom.ZOrder 0
        edcSelCFrom.Visible = False: edcSelCTo.Visible = False: edcSelCFrom1.Visible = False: edcSelCTo1.Visible = False
        pbcSelC.Width = pbcSelC.Width + 450
        pbcSelC.ZOrder 0
        
        'show latest version or all versions
        plcSelC11.Move 120, edcSelCTo.Top + edcSelCTo.Height + 120, 2400
        rbcSelC11(0).Value = True
        rbcSelC11(0).Caption = "Latest"
        rbcSelC11(0).Move 840, 0, 840
        rbcSelC11(1).Caption = "All"
        rbcSelC11(1).Move 1800, 0, 600
        rbcSelC11(0).Visible = True
        rbcSelC11(1).Visible = True
        smPaintCaption11 = "Version"
        plcSelC11_Paint
        plcSelC11.Visible = True
        
        cbcSet1.Move 1080, plcSelC11.Top + plcSelC11.Height + 60, 2000
        edcSet1.Text = "Sort by"
        edcSet1.Move 120, cbcSet1.Top + 60, 840
        cbcSet1.AddItem "Advertiser"
        cbcSet1.AddItem "Agency"
        cbcSet1.AddItem "Agency Estimate #"
        cbcSet1.AddItem "Contract #"
        cbcSet1.AddItem "Sender"
        cbcSet1.AddItem "Sent Date"
        cbcSet1.AddItem "Status"
        cbcSet1.ListIndex = 0               'default to advertiser sort
        edcSet1.Visible = True
        cbcSet1.Visible = True
        
        lacText.Text = "Contract #"
        lacText.Move 120, edcSet1.Top + edcSet1.Height + 120, 960
        lacText.Visible = True
        edcText.Move 1200, edcSet1.Top + edcSet1.Height + 60, 1080
        edcText.MaxLength = 9
        edcText.Text = ""
        edcText.Visible = True
        mSetCommands
        ElseIf ilListIndex = CNT_XML_ACTIVITY Then        '3-26-16
        edcSelCTo.MaxLength = 10
        edcSelCTo1.MaxLength = 10
        edcSelCFrom.MaxLength = 10
        edcSelCFrom1.MaxLength = 10
        
        'sent start/end dates
        slStr = Format$(gNow(), "m/d/yy")
        llStartDate = gDateValue(slStr)
        llActiveDate = gDateValue(slStr)
        llTodayDate = gDateValue(slStr)
        'backup to Monday
        ilDay = gWeekDayLong(llStartDate)
        'if today is already a monday, backup to last week
        If ilDay = 0 Then
            ilDay = -1
        End If
        Do While ilDay <> 0
            llStartDate = llStartDate - 1
            ilDay = gWeekDayLong(llStartDate)
        Loop
        
        ilDay = gWeekDayLong(llActiveDate)
        Do While ilDay <> 0
            llActiveDate = llActiveDate - 1
            ilDay = gWeekDayLong(llActiveDate)
        Loop
        
        edcSelCFrom.Text = Format(llStartDate, "m/d/yy")     'sent startdate
        edcSelCFrom1.Text = Format(llTodayDate, "m/d/yy")    'sent end date
        edcSelCTo.Text = Format(llActiveDate, "m/d/yy")       'active start date
        edcSelCTo1.Text = Format(llActiveDate + 6, "m/d/yy")  'active end date
        
        lacSelCFrom.Caption = "Sent Date- From"
        lacSelCFrom1.Caption = "To"
        lacSelCFrom.Visible = True
        lacSelCFrom1.Visible = True
        lacSelCTo.Caption = "Active Date- From"
        lacSelCTo1.Caption = "To"
        lacSelCTo.Visible = True
        lacSelCTo1.Visible = True
        'edcSelCFrom.Visible = True
        'edcSelCFrom1.Visible = True
        'edcSelCTo.Visible = True
        'edcSelCTo1.Visible = True
        
        lacSelCFrom.Move 120, 60, 1530
        'edcSelCFrom.Move lacSelCFrom.Left + lacSelCFrom.Width + 10, 60, 945     ' 1650, edcSelCFrom.Top, 945
        lacSelCFrom1.Move 2810  'edcSelCFrom.Left + edcSelCFrom.Width + 10, edcSelCFrom.Top, 75     '2715, 75
        'edcSelCFrom1.Move lacSelCFrom1.Left + lacSelCFrom1.Width + 10, 945  ' 3060, edcSelCFrom.Top, 945
        'active start/end dates
        lacSelCTo.Move 120, lacSelCFrom.Top + lacSelCFrom.Height + 120, 1530
        'edcSelCTo.Move 1650, lacSelCTo.Top - 60, 945
        lacSelCTo1.Move 2810
        'edcSelCTo1.Move 3060, edcSelCTo.Top, 945
        
        'Date: 12/24/2019 added CSI calendar control for date entries
        edcSelCFrom.Visible = False: edcSelCFrom1.Visible = False: edcSelCTo.Visible = False: edcSelCTo1.Visible = False
        
        pbcSelC.Width = pbcSelC.Width + 800

        CSI_CalFrom.Text = Format(llStartDate, "m/d/yy")     'sent startdate
        CSI_CalTo.Text = Format(llTodayDate, "m/d/yy")    'sent end date
        CSI_From1.Text = Format(llActiveDate, "m/d/yy")       'active start date
        CSI_To1.Text = Format(llActiveDate + 6, "m/d/yy")  'active end date

        CSI_CalFrom.Move lacSelCFrom.Left + lacSelCFrom.Width + 10, lacSelCFrom.Top, 1080
        CSI_From1.Move lacSelCTo.Left + lacSelCTo.Width + 10, lacSelCTo.Top, 1080
        CSI_To1.Move lacSelCTo1.Left + lacSelCTo1.Width + 10, lacSelCTo1.Top, 1080
        CSI_CalTo.Move CSI_To1.Left, lacSelCFrom1.Top, 1080
        
        'lacSelCFrom1.Move 2810
       
        CSI_CalFrom.Visible = True
        CSI_From1.Visible = True
        CSI_CalTo.Visible = True
        CSI_To1.Visible = True
        
        CSI_From1.ZOrder 0: CSI_To1.ZOrder 0
        CSI_CalFrom.ZOrder 0: CSI_CalTo.ZOrder 0
       
        'show latest version or all versions
        plcSelC11.Move 120, edcSelCTo.Top + edcSelCTo.Height + 120, 2400
        rbcSelC11(0).Value = True
        rbcSelC11(0).Caption = "Latest"
        rbcSelC11(0).Move 840, 0, 840
        rbcSelC11(1).Caption = "All"
        rbcSelC11(1).Move 1800, 0, 600
        rbcSelC11(0).Visible = True
        rbcSelC11(1).Visible = True
        smPaintCaption11 = "Version"
        plcSelC11_Paint
        plcSelC11.Visible = True
        
        ckcSelC12(0).Caption = "Sent"
        ckcSelC12(0).Move 720, 0, 720
        ckcSelC12(1).Move 1560, 0, 1320
        ckcSelC12(1).Caption = "Not Sent"
        ckcSelC12(0).Value = vbChecked
        ckcSelC12(1).Value = vbChecked
        ckcSelC12(0).Visible = True
        ckcSelC12(1).Visible = True
        smPaintCaption12 = "Status"
        plcSelC12.Move 120, plcSelC11.Top + plcSelC11.Height + 60, 3000
        plcSelC12.Visible = True
        
        cbcSet1.Move 1080, plcSelC11.Top + plcSelC11.Height + 60, 2000
        edcSet1.Text = "Sort by"
        edcSet1.Move 120, cbcSet1.Top + 60, 840
        cbcSet1.AddItem "Advertiser"
        cbcSet1.AddItem "Agency"
        cbcSet1.AddItem "Agency Estimate #"
        cbcSet1.AddItem "Contract #"
        cbcSet1.AddItem "Sender"
        cbcSet1.AddItem "Sent Date"
        cbcSet1.AddItem "Status"
        cbcSet1.ListIndex = 0               'default to advertiser sort
        'hide all sort features for now; default to advertier
        edcSet1.Visible = False 'True
        cbcSet1.Visible = False   'rue
        
        lacText.Text = "Contract #"
        'lacText.Move 120, edcSet1.Top + edcSet1.Height + 120, 960
        lacText.Move 120, plcSelC12.Top + plcSelC12.Height + 120, 960
        lacText.Visible = True
        'edcText.Move 1200, edcSet1.Top + edcSet1.Height + 60, 1080
        edcText.Move 1200, plcSelC12.Top + plcSelC12.Height + 60, 1080
        edcText.MaxLength = 9
        edcText.Text = ""
        edcText.Visible = True
        mSetCommands
    End If
    frcOption.Visible = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mAASCntrPop                        *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                    :7/10/96 -Use new contract status*
'                                                      *
'*            Comments: Populate the selection combo   *
'*                      box for advt, agy or slsp      *
'*        5-16-02                                      *
'*******************************************************
Public Sub mAASCntrPop(ilAAS As Integer, ilLbcIndex As Integer, slCntrStatus As String, ilHOState As Integer)
'
'   mCntrPop
'   Where:
'       ilAAS - 0 = advt, 1 = agy, 2 = slsp
'       ilLbcIndex - index into List box containing advt, agy or slsp
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
    For illoop = 0 To lbcSelection(ilLbcIndex).ListCount - 1 Step 1
        If lbcSelection(ilLbcIndex).Selected(illoop) Then
            sgMultiCntrCodeTag = ""             'init the date stamp so the box will be populated
            ReDim tgMultiCntrCodeCT(0 To 0) As SORTCODE
            lbcMultiCntr.Clear
            If ilAAS = 0 Then
                slNameCode = tgAdvertiser(illoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
            ElseIf ilAAS = 1 Then
                slNameCode = tgAgency(illoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
            Else
                slNameCode = tgSalesperson(illoop).sKey  'Traffic!lbcAdvertiser.List(ilLoop)
            End If
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
            ilRet = gPopCntrForAASBox(RptSelCt, ilAAS, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilHOState, ilShow, lbcMultiCntr, tgMultiCntrCodeCT(), sgMultiCntrCodeTagCT)
            If ilRet <> CP_MSG_NOPOPREQ Then
                On Error GoTo mCntrPopErr
                gCPErrorMsg ilRet, "mCntrPop (gPopCntrBox)", RptSelCt
                On Error GoTo 0
            End If
            For ilIndex = 0 To UBound(tgMultiCntrCodeCT) - 1 Step 1 'lbcMultiCntrCode.ListCount - 1 Step 1
                slName = Trim$(tgMultiCntrCodeCT(ilIndex).sKey)  'lbcMultiCntrCode.List(ilIndex)
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
    imTerminate = True
    Exit Sub
End Sub

'
'
'       Send formula to Daily Sales Activity by Month and Sales Placement
'       (slsactss.rpt)
'
'
Public Sub mSalesFormula()
    Dim illoop As Integer
    Dim ilMajorSet As Integer
    Dim slVGSelected As String
    Dim slStr As String

    illoop = RptSelCt!cbcSet1.ListIndex
    ilMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
    slVGSelected = ""                       'this for the vehicle group headers if one selected, a vehicle group
                                            'could still be selected even if its not a primary vehicle group sort
    'assume sort by vehicle group (rbcselc9(0).value = true)
    If ilMajorSet = 1 Then
        slStr = "P"
    ElseIf ilMajorSet = 2 Then
        slStr = "S"
    ElseIf ilMajorSet = 3 Then
        slStr = "M"
    ElseIf ilMajorSet = 4 Then
        slStr = "F"
    ElseIf ilMajorSet = 5 Then
        slStr = "R"
    Else
        slStr = "N"
    End If
    slVGSelected = Trim$(slStr)
    If RptSelCt!rbcSelC9(1).Value Then
        slStr = "O"         'sort by office
    ElseIf RptSelCt!rbcSelC9(2).Value Then
        slStr = "A"         'sort by advertiser
    End If
    If Not gSetFormula("Sortby", "'" & slStr & "'") Then
        Exit Sub
    End If
    If Not gSetFormula("VGSelected", "'" & slVGSelected & "'") Then
        Exit Sub
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCntSelectivity0                *
'*                                                     *
'*            Created:3-18-03       By:D. Hosaka       *
'*                                                     *
'*            Comments:                                *
'*              Duplicated from mCntSelectivity1;      *
'*              mCntSelectivity1 too large             *
'*                                                     *
'*******************************************************
Public Sub mCntSelectivity0()
    Dim ilListIndex As Integer
    Dim ilRet As Integer
    ReDim ilAASCodes(0 To 1) As Integer

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    rbcSelCInclude(2).Visible = False
    If tgSpf.sRUseCorpCal = "Y" Then       'if Using corp cal,retain in memory
        ilRet = gObtainCorpCal()
    End If

    ilListIndex = lbcRptType.ListIndex
    If (igRptType = 0) And (ilListIndex > 1) Then
        ilListIndex = ilListIndex + 1
    End If
    mInitControls           'set controls to proper positions, widths, hidden, shown, etc.
    If ilListIndex = CNT_INSERTION Then     'populate list box with Show On Insertion Only is Y (or not an "N")
        mSellConvVirtVehPop 6, False, True
    Else
        mSellConvVirtVehPop 6, False
    End If
    
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If

    Select Case ilListIndex
'        Case CNT_BR, CNT_PAPERWORK, CNT_SPTSBYADVT, CNT_INSERTION     'Contract, Contract summary, spots by advt
        Case CNT_BR, CNT_PAPERWORK, CNT_INSERTION      'Contract, Contract summary, spots by advt
            mSPersonPop lbcSelection(2)        'this list box is used for Salesperson or sales office
                                                'populate when needed
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mSPersonPop lbcSelection(9)        'this list box is used for Salesperson or sales office
                                                'populate when needed
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mAgencyPop lbcSelection(8)  '5-16-02 for contracts/proposals
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            ckcAll.Visible = True
            ckcAll.Enabled = True
            lacSelCFrom.Left = 120
            lacSelCFrom1.Move 2685, 75
            edcSelCFrom.Move 1530, edcSelCFrom.Top, 945
            edcSelCFrom1.Move 3060, edcSelCFrom.Top, 945
            lacSelCTo.Left = 120
            edcSelCTo.Move 1530, edcSelCTo.Top, 945
            lacSelCTo1.Left = 2685
            'edcSelCTo1.Move 2700, edcSelCTo1.Top, 945
            edcSelCTo1.Move 3060, edcSelCTo.Top, 945
            edcSelCTo.MaxLength = 10    '8  5/27/99 changed for short form date m/d/yyyy
            edcSelCTo1.MaxLength = 10   '8  5/27/99 changed for short form date m/d/yyyy
            edcSelCFrom.MaxLength = 10  '8  5/27/99 changed for short form date m/d/yyyy
            edcSelCFrom1.MaxLength = 10 '8  5/27/99 changed for short form date m/d/yyyy
'            If ilListIndex = CNT_SPTSBYADVT Then
'                lacSelCFrom.Caption = "Spots: From"
'            Else
                lacSelCFrom.Caption = "Active: From"
'            End If
            lacSelCFrom1.Caption = "To"
            lacSelCFrom.Visible = True
            lacSelCFrom1.Visible = True
            lacSelCTo.Caption = "Entered: From"
            lacSelCTo1.Caption = "To"
            lacSelCTo.Visible = True
            lacSelCTo1.Visible = True
            edcSelCFrom.Visible = True
            edcSelCFrom1.Visible = True
            edcSelCTo.Visible = True
            edcSelCTo1.Visible = True
            plcSelC3.Visible = False
            
            'Date: 12/19/2019 added CSI calendar control for date entries
            edcSelCFrom.Visible = False: edcSelCFrom1.Visible = False: edcSelCTo.Visible = False: edcSelCTo1.Visible = False
            CSI_CalFrom.Move 1340, edcSelCFrom.Top, 1080
            CSI_CalTo.Move 2700, edcSelCFrom.Top, 1080
            CSI_From1.Move 1340, edcSelCTo.Top, 1080
            CSI_To1.Move 2700, edcSelCTo1.Top, 1080
            CSI_CalFrom.Visible = True
            CSI_From1.Visible = True
            CSI_CalTo.Visible = True
            CSI_To1.Visible = True
            
            lacSelCFrom1.Move CSI_CalFrom.Left + CSI_CalFrom.Width + 10
            lacSelCTo1.Move CSI_From1.Left + CSI_From1.Width + 10
            
            CSI_From1.ZOrder 0: CSI_To1.ZOrder 0
            CSI_CalFrom.ZOrder 0: CSI_CalTo.ZOrder 0
            
            'plcSelC1.Caption = "Select"
            smPaintCaption1 = "Select"
            plcSelC1_Paint
            rbcSelCSelect(0).Caption = "Advt"
            rbcSelCSelect(0).Move 600, 0, 675

            rbcSelCSelect(1).Caption = "Agency"
            rbcSelCSelect(1).Move 1260, 0, 960

            rbcSelCSelect(1).Visible = True
            rbcSelCSelect(2).Caption = "Slsp"   '7-18-01
            rbcSelCSelect(2).Move 2220, 0, 675
             '7-18-01
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
            If ilListIndex = CNT_BR Then                     'proposals/contrs only
                'ckcAll.Move lbcSelection(10).Left
                'If rbcSelCSelect(0).Value Then
                '    rbcSelCSelect_click 0, True             'default to advt cntrs
                'Else
                '    rbcSelCSelect(0).Value = True
                'End If

            Else                                        'contract summary or spots by advt
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
            End If
            ' 10-29-10 remove spotsby advt option,  its in rptcrcb

            If (ilListIndex = CNT_BR) Or (ilListIndex = CNT_PAPERWORK) Or (ilListIndex = CNT_INSERTION) Then      'BR or paperwork                   'Contracts  or summary
                plcSelC3.Visible = True
                'plcSelC3.Caption = ""
                smPaintCaption3 = ""
                plcSelC3_Paint
                ckcSelC3(0).Visible = True
                ckcSelC3(0).Value = vbUnchecked 'False
                ckcSelC3(1).Visible = False
                ckcSelC3(2).Visible = False
                ckcSelC3(3).Visible = False
                ckcSelC3(4).Visible = False
                ckcSelC3(5).Visible = False
            End If
            If ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION Then                     'retain in this order, questions are displayed
                                                        'by previous locations
                ckcAll.Move lbcSelection(10).Left
                If rbcSelCSelect(0).Value Then
                    rbcSelCSelect_click 0             'default to advt cntrs
                Else
                    rbcSelCSelect(0).Value = True
                End If
                

                If ilListIndex = CNT_BR Then
                    rbcOutput(3).Visible = False             'disallow contract PDF & send by email for now

                    plcSelC2.Visible = True
                    '10-29-03 Show proposals/orders intermixed:  proposals show Vrsion #, orders show REvision #
                    'remove ability to request Portrait contract
                    '11-6-03 client doesnt want them combined, ask question for proposals, orders or both
                    smPaintCaption2 = "Select*"
                    plcSelC2_Paint
                    plcSelC2.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height + 30
                    rbcSelCInclude(0).Caption = "Proposals"
                    rbcSelCInclude(0).Left = 600
                    rbcSelCInclude(0).Width = 1200
                    rbcSelCInclude(1).Caption = "Contracts"
                    rbcSelCInclude(1).Left = 1800
                    rbcSelCInclude(1).Width = 1200
                    rbcSelCInclude(2).Caption = "Both"
                    rbcSelCInclude(2).Left = 3000
                    rbcSelCInclude(2).Width = 720
                    laclbcName(0).Caption = "* V = Proposal version #, R = Order revision #"
                    laclbcName(0).FontName = "Arial"
                    laclbcName(0).FontSize = 7
                    laclbcName(0).Move lbcSelection(5).Left, lbcSelection(5).Top + lbcSelection(5).Height + 15, 4000
                    laclbcName(0).Visible = True
                    rbcSelCInclude(2).Visible = True
                    rbcSelCInclude(2).Value = True
                    If (tgSpf.sGUsePropSys <> "Y") Then     'proposals disabled, dont allow that option
                        rbcSelCInclude(0).Enabled = False
                        rbcSelCInclude(1).Value = True
                    End If

                ElseIf ilListIndex = CNT_INSERTION Then
                    'D.S. 07-13-15  Added test for EDS
                    If (Asc(tgSaf(0).sFeatures2) And EMAILDISTRIBUTION) = EMAILDISTRIBUTION And ((tgUrf(0).lEMailCefCode > 0 And tgUrf(0).iCode > 2) Or tgUrf(0).iCode <= 2) Then
                        rbcOutput(3).Visible = True             'allow PDF & send by email
                    End If
                    
                    'v81 testing results 3-28-22 Issue 1: "separate files per vehicle" checkbox from the Insertion Orders report is appearing
                    rbcOutput(4).Top = rbcOutput(3).Top
                    ckcSeparateFile.Visible = True
                
                    rbcSelCInclude(2).Value = True      '1-23-04  force to show both prop and contracts
                    plcSelC2.Visible = False                    'dont show option for proposals, landscape or narrow.  dfaulted to landscape
                    lbcSelection(0).Height = lbcSelection(0).Height / 2
                    lbcSelection(5).Height = lbcSelection(0).Height
                    lbcSelection(8).Height = lbcSelection(0).Height
                    lbcSelection(9).Height = lbcSelection(0).Height
                    lbcSelection(10).Height = lbcSelection(0).Height
                    lbcSelection(6).Move lbcSelection(1).Left, lbcSelection(0).Top + lbcSelection(0).Height + 240 + 60, lbcSelection(6).Width, lbcSelection(0).Height '- CkcAllveh.Height
                    lbcSelection(6).Visible = True
                    CkcAllveh.Move lbcSelection(5).Left, lbcSelection(6).Top - CkcAllveh.Height
                    CkcAllveh.Visible = True

                End If
                ckcSelC3(0).Move 0, -30, 1800
                ckcAll.Caption = "All Contracts"
                ckcAll.Visible = True
                ckcAll.Enabled = False
                plcSelC4.Visible = True

                smPaintCaption4 = "Show"
                plcSelC4_Paint
                plcSelC4.Move plcSelC2.Left, plcSelC2.Top + plcSelC2.Height + 30
                'plcSelC4.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height + 30
                If ilListIndex = CNT_INSERTION Then
                    plcSelC4.Visible = False
                    If rbcSelC4(0).Value Then
                        rbcSelC4_click 0
                    Else
                        rbcSelC4(0).Value = True
                    End If
                Else
                    mAskSumDetailBoth 600, ilListIndex          'send left position of first button
                End If
                If ilListIndex = CNT_INSERTION Then     'for insertions, don't show option for which Demo to use
                    plcSelC7.Visible = False
                Else
                    plcSelC7.Visible = True
                End If
                'plcSelC7.Caption = "Use"
                smPaintCaption7 = "Use"
                plcSelC7_Paint
                plcSelC7.Move plcSelC4.Left, plcSelC4.Top + plcSelC4.Height
                rbcSelC7(0).Caption = "Primary Demo"
                rbcSelC7(0).Left = 600
                rbcSelC7(0).Width = 1520
                rbcSelC7(0).Visible = True
                rbcSelC7(0).Value = True
                rbcSelC7(1).Caption = "All Demo Categories"
                rbcSelC7(1).Left = 2100
                rbcSelC7(1).Width = 2400
                rbcSelC7(1).Visible = True
                rbcSelC7(2).Visible = False

                ilRet = gSocEcoPop(RptSelCt, cbcSet1)


                plcSelC6.Height = plcSelC7.Height
                If ilListIndex = CNT_INSERTION Then
                    'if error when populating socio-economic stuff, dont abort
                    cbcSet1.Move 1295, plcSelC1.Top + plcSelC1.Height
                    edcSet1.Text = "Qualitative"
                    edcSet1.Move 120, cbcSet1.Top + 60, 1080
                    'plcSelC6.Move plcSelC1.Left, plcSelC1.Top + plcSelC1.Height
                    plcSelC6.Move plcSelC1.Left, cbcSet1.Top + cbcSet1.Height + 30
                Else
                    'if error when populating socio-economic stuff, dont abort
                    cbcSet1.Move 1295, plcSelC7.Top + plcSelC7.Height
                    edcSet1.Text = "Qualitative"
                    edcSet1.Move 120, cbcSet1.Top + 60, 1080
                    plcSelC6.Move plcSelC7.Left, cbcSet1.Top + cbcSet1.Height + 50       'Date: 9/13/2018 made sure following controls (Include, Rate, Research) line up correctly   FYM
                End If

                edcSet1.Visible = True
                cbcSet1.Visible = True
                plcSelC6.Visible = True
                smPaintCaption6 = "Include"
                plcSelC6_Paint
                ckcSelC6(0).Caption = "Rates"
                ckcSelC6(0).Left = 720
                ckcSelC6(0).Value = vbUnchecked 'False
                ckcSelC6(0).Visible = True
                ckcSelC6(0).Value = vbChecked   'True                'default to include all prices
                ckcSelC6(0).Width = 800
                ckcSelC6(1).Left = 1560
                ckcSelC6(1).Width = 1080
                ckcSelC6(1).Caption = "Research"
                'ckcSelC6(1).Value = vbChecked   'False
                ckcSelC6(1).Visible = True

'                If ((Asc(tgSpf.sOptionFields) And &H80) = &H80) And ilListIndex <> CNT_INSERTION Then  '8-29-06 chg to test Using Research feature rather than Proposal
                If ((Asc(tgSpf.sOptionFields) And OFRESEARCH) = OFRESEARCH) Then  '8-29-06 chg to test Using Research feature rather than Proposal
                '5-15-12 always leave Research defaulted off for Insertion orders; stations dont need to see research info
                'If tgSpf.sGUsePropSys = "Y" Then
                    ckcSelC6(1).Enabled = True
                    If ilListIndex = CNT_INSERTION Then
                        ckcSelC6(1).Value = vbUnchecked     'insertions orders do not need to see research, default that option off
                    Else
                        ckcSelC6(1).Value = vbChecked
                    End If
                Else
                    ckcSelC6(1).Value = vbUnchecked
                    ckcSelC6(1).Enabled = False
                End If
                ckcSelC6(2).Left = 2700
                ckcSelC6(2).Width = 920
                ckcSelC6(2).Caption = "Hidden"
                ckcSelC6(2).Visible = True
                If ilListIndex = CNT_INSERTION Then
                    'ckcSelC5(0).Visible = False        '6-30-04 allow differences on insertion order
                    'ckcSelC6(2).Visible = False         'disallow proofs on insertion, all lines from hidden (no packages)

                Else
                    ckcSelC5(0).Visible = True
                'End If
                'plcSelC5.Caption = ""
                    smPaintCaption5 = ""
                    plcSelC5_Paint
                    plcSelC5.Move plcSelC5.Left, plcSelC6.Top + plcSelC6.Height
                    plcSelC5.Visible = True
                    ckcSelC5(0).Caption = "Differences Only"
                    ckcSelC5(0).Left = 0
                    ckcSelC5(0).Value = vbUnchecked 'False
                    ckcSelC5(0).Width = 1800
                    ckcSelC5(1).Visible = False
                    ckcSelC5(2).Visible = False
                End If
                If Not rbcSelCInclude(0).Value And ilListIndex <> CNT_INSERTION Then            'printables only apply to holds/orders (not proposals)
                    ckcSelC3(0).Caption = "Printables Only"
                    ckcSelC3(0).Value = vbUnchecked 'False
                    'plcSelC3.Move plcSelC5.Left, plcSelC5.Top + plcSelC5.Height
                    plcSelC3.Move plcSelC4.Left, plcSelC5.Top + plcSelC5.Height
                    plcSelC8.Move plcSelC3.Left, plcSelC3.Top + plcSelC3.Height
                    'plcSelC8.Caption = "For printables"
                    smPaintCaption8 = "For printables"
                    plcSelC8_Paint
                    ckcSelC8(0).Visible = True
                    ckcSelC8(0).Move 1200, -30, 3000
                    ckcSelC8(0).Caption = "Show mods as differences"
                    ckcSelC8(0).Value = vbUnchecked 'False
                Else
                    plcSelC3.Visible = False                    'don't show printables only for proposals
                    plcSelC8.Visible = False
                End If
                If tgUrf(0).iSlfCode > 0 Then           'its a slsp, don't allow to exclude reserves
                    'disable ability for slsp to ask for printables
                    plcSelC3.Enabled = False
                    ckcSelC3(0).Enabled = False
                    plcSelC8.Enabled = False
                    ckcSelC8(0).Enabled = False
                End If
                'Corp or Standard for Proposal or Orders version (not Portrait version)
                plcSelC9.Move 120, plcSelC8.Top + plcSelC8.Height
                'plcSelC9.Caption = "Summary Month Totals"
                smPaintCaption9 = "Summary Month Totals"
                plcSelC9_Paint
                rbcSelC9(0).Caption = "Corporate"
                rbcSelC9(0).Left = 2040     '660
                rbcSelC9(0).Width = 2520    '1140
                rbcSelC9(1).Caption = "Standard"
                rbcSelC9(1).Left = 3220     '1840
                rbcSelC9(1).Width = 2520     '1140
                rbcSelC9(0).Visible = True
                rbcSelC9(1).Visible = True
                rbcSelC9(2).Visible = False
                rbcSelC9(1).Value = True
                If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
                    rbcSelC9(0).Enabled = False
                    rbcSelC9(0).Value = False
                    rbcSelC9(1).Value = True
                Else
                    rbcSelC9(0).Value = True
                End If
                plcSelC9.Visible = True
                '*******temporary patch until ready to release  (See also rbcSelCInclude, ckcSelC6)
                If rbcSelC9(1).Value Then
                    rbcSelC9_click 1
                Else
                    rbcSelC9(1).Value = True
                End If
                plcSelC9.Visible = False

                '8-28-00 Show Commission splits?
                If ilListIndex = CNT_BR Then
                    plcSelC10.Move 120, plcSelC8.Top + plcSelC8.Height - 30, 4000, 510
                    'plcSelC10.Caption = ""
                    smPaintCaption10 = ""
                    plcSelC10_Paint
                    ckcSelC10(0).Move 0, 0, 4000
                    ckcSelC10(0).Caption = "Show Slsp Commission Splits on Summary"
                    ckcSelC10(0).Value = vbChecked 'True
                    ckcSelC10(0).Visible = True
                    ckcSelC10(1).Move 0, 240, 4000          '2-2-10 merge ntr and air time
                    ckcSelC10(1).Caption = "Combine Air Time and NTR/CPM Totals"
                    ckcSelC10(1).Value = vbChecked
                    ckcSelC10(1).Visible = True
                    plcSelC10.Visible = True
                    
                    plcSelC13.Move 120, plcSelC10.Top + plcSelC10.Height, 4000
                    ckcSelC13(0).Move 0, 0, 2040
                    ckcSelC13(0).Caption = "Net Amt on Proposal"
                    ckcSelC13(0).Value = vbChecked
                    ckcSelC13(0).Visible = True
                    
                    '8-25-15 option to show product protection codes
                    ckcSelC13(1).Caption = "Show Product Prot"
                    ckcSelC13(1).Move 2160, 0, 2400
                    ckcSelC13(1).Visible = True
                    plcSelC13.Visible = True
                    '4-18-18 move this after podcast question
'                    lacTopDown.Caption = "Contract #"
'                    lacTopDown.Move 120, plcSelC13.Top + plcSelC13.Height + 30
'                    lacTopDown.Visible = True
'                    edcTopHowMany.Move 1080, plcSelC13.Top + plcSelC13.Height, 945
'                    edcTopHowMany.MaxLength = 9
'                    edcTopHowMany = ""
'                    edcTopHowMany.Visible = True
                    
                    '4-18-18 option to show Aud Percentages for Podcast
                    ckcInclRevAdj.Move 120, plcSelC13.Top + plcSelC13.Height + 30, 2280
                    ckcInclRevAdj.Caption = "Incl Aud % for Podcast"
                    ckcInclRevAdj.Visible = True
                    ckcInclRevAdj.Value = vbUnchecked         '4-18-18 take the default from site, allow user to change
                    If ((Asc(tgSaf(0).sFeatures5) And PODCASTAUDPCT) = PODCASTAUDPCT) Then
                         ckcInclRevAdj.Value = vbChecked
                    End If
                                        
                    'TTP 10382 - Contract report: Option To not show Act1 codes on PDF
                    ckcShowACT1.Move 120, ckcInclRevAdj.Top + ckcInclRevAdj.Height + 30, 4320
                    ckcShowACT1.Visible = True
                    If ((Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES) Then
                        ckcShowACT1.Value = vbChecked
                        ckcShowACT1.Enabled = True
                    Else
                        ckcShowACT1.Value = vbUnchecked
                        ckcShowACT1.Enabled = False
                    End If
                    
                    'TTP 10745 - NTR: add option to only show vehicle, billing date, and description on the contract report, and vehicle and description only on invoice reprint
                    ckcSuppressNTRDetails.Move 120, ckcShowACT1.Top + ckcShowACT1.Height + 30
                    ckcSuppressNTRDetails.Visible = True
                    
                    lacTopDown.Caption = "Contract #"
                    lacTopDown.Move 2520, plcSelC13.Top + plcSelC13.Height + 30
                    lacTopDown.Visible = True

                    edcTopHowMany.Move 3520, plcSelC13.Top + plcSelC13.Height, 945
                    edcTopHowMany.MaxLength = 9
                    edcTopHowMany = ""
                    edcTopHowMany.Visible = True

                End If
                If ilListIndex = CNT_INSERTION Then
                    '11-27-00 selective contr # for Insertion Orders
                    edcTopHowMany.Move 1290, plcSelC6.Top + plcSelC6.Height, 945
                    edcTopHowMany.MaxLength = 9
                    edcTopHowMany = ""
                    lacTopDown.Caption = "Contract #"
                    lacTopDown.Move 120, edcTopHowMany.Top + 30
                    edcTopHowMany.Visible = True
                    lacTopDown.Visible = True
                    '1-16-01 Ask to Show Net-Net values
                    ckcSelC10(0).Value = vbUnchecked    'False
                    ckcSelC10(0).Caption = "Show Net-Net"
                    ckcSelC10(1).Visible = False
                    ckcSelC10(2).Visible = False
                    ckcSelC10(0).Visible = True
                    ckcSelC10(0).Left = 120
                    ckcSelC10(0).Width = 1420
                    plcSelC10.Move 0, edcTopHowMany.Top + edcTopHowMany.Height
                    ckcSelC6(2).Visible = False         'disallow proofs on insertion, all lines from hidden (no packages)
'                    plcSelC5.Top = plcSelC10.Top + plcSelC10.Height
'                    plcSelC5.Visible = True             '6-30-03 allow differences on insertion orders
'                    ckcSelC5(0).Visible = True
'                    plcSelC5.Visible = True
                    smPaintCaption10 = ""
                    plcSelC10_Paint
                    plcSelC10.Visible = True
                    '10-24-08 disable differences only, add option to include NTRs
                    plcSelC12.Move 120, plcSelC10.Top + plcSelC10.Height, 1320
                    ckcSelC12(0).Visible = True
                    ckcSelC12(0).Caption = "Include NTR"
                    ckcSelC12(0).Move 0, 0, 1560
                    plcSelC12.Visible = True
                    smPaintCaption12 = ""
                    
                    '5-23-13 option to show the Product Production code
                    '8-25-15 chg to use index 1 vs 0, so that the Contracts/proposals can use same field
                    plcSelC13.Move 120, plcSelC12.Top + plcSelC12.Height
                    ckcSelC13(1).Move 0, 0, 3000
                    ckcSelC13(1).Caption = "Show Product Protection Code"
                    ckcSelC13(1).Visible = True
                    plcSelC13.Visible = True
                    
                    '5-13-15 Differences only option
                    smPaintCaption5 = ""
                    plcSelC5_Paint
                    plcSelC5.Move 120, plcSelC13.Top + plcSelC13.Height + 60
                    plcSelC5.Visible = True
                    ckcSelC5(0).Caption = "Include Differences"
                    ckcSelC5(0).Left = 0
                    ckcSelC5(0).Value = vbUnchecked 'False
                    ckcSelC5(0).Width = 3000
                    ckcSelC5(0).Visible = True
                    ckcSelC5(0).Value = vbChecked
                    ckcSelC5(1).Visible = False
                    ckcSelC5(2).Visible = False
                    
                    'if Insertion Order is using Site for the name and address (vs payee or vehicle N & A), ask to show the Payee Name
                    If tgSpf.sInsertAddr = "S" Then
                        ckcInclZero.Move 120, plcSelC5.Top + plcSelC5.Height, 4000
                        ckcInclZero.Caption = "Replace Agency Phone # w/Agency Name"
                        ckcInclZero.Visible = True
                    Else
                        ckcInclZero.Value = vbUnchecked         'not using Site, default so the Agency phone # is not replaced with Agy name
                    End If
                  
                End If
            ElseIf ilListIndex = CNT_PAPERWORK Then                         'paperwork
                'The entire question selectivity has been modified to include new questions
                mSellConvVVPkgPop 6, False
                If imTerminate Then
                    cmcCancel_Click
                    Exit Sub
                End If
                rbcSelCSelect(3).Caption = "Vehicle"
                rbcSelCSelect(3).Left = 2895        '7-18-01
                rbcSelCSelect(3).Width = 980
                rbcSelCSelect(3).Top = 0            '7-18-01
                rbcSelCSelect(3).Visible = True
                'plcSelC1.Height = 435              '7-18-01
                rbcSelCSelect(0).Value = True                       'default to advt selection

                'plcSelC9.Caption = "Sort"
                smPaintCaption9 = "Sort"
                plcSelC9_Paint
                rbcSelC9(0).Caption = "Advt"
                rbcSelC9(0).Move 600, 0, 675
                rbcSelC9(0).Visible = False
                rbcSelC9(1).Caption = "Agency"
                rbcSelC9(1).Move 1290, 0, 930
                rbcSelC9(1).Visible = False
                rbcSelC9(2).Caption = "Slsp"   '7-18-01
                rbcSelC9(2).Move 2220, 0, 675
                rbcSelC9(1).Enabled = True
                rbcSelC9(2).Enabled = True
                rbcSelC9(2).Visible = False
                rbcSelC9(3).Caption = "Vehicle"
                rbcSelC9(3).Move 2895, 0, 980
                rbcSelC9(3).Visible = False
                rbcSelC9(0).Value = True                       'default to advt selection
                plcSelC9.Move 120, plcSelC1.Top + plcSelC1.Height, 4400
                plcSelC9.Visible = True
                
                lacSort1.Move 0, 25
                cbcSort1.Move lacSort1.Width + 10, 0
                lacSort2.Move cbcSort1.Left + cbcSort1.Width + 50, 25
                cbcSort2.Move lacSort2.Left + lacSort2.Width + 10, 0
                cbcSort1.Visible = True: cbcSort2.Visible = True
                lacSort1.Visible = True: lacSort2.Visible = True

                plcSelC5.Visible = True                            'hold, order, dead, working,  etc check box
                plcSelC3.Move 120, plcSelC9.Top + plcSelC9.Height  '7-18-01
                plcSelC3.Height = 435           '7-18-01 (was selc3) show 2 lines of questions for this Include
                'plcSelC3.Caption = "Include"
                smPaintCaption3 = "Include"
                plcSelC3_Paint
                mAskContractTypes
                plcSelC3.Visible = True                            'Type check box: trades, DP, PI, psa, etc.

                plcSelC5.Height = plcSelC5.Height + 195
                plcSelC5.Move plcSelC5.Left, 1700

                ckcSelC5(7).Caption = "NTR"
                ckcSelC5(7).Move 660, 390, 840
                ckcSelC5(7).Value = vbChecked           'default to include NTRs
                ckcSelC5(7).Visible = True

                rbcSelC4(0).Caption = "Cash"
                rbcSelC4(0).Visible = True
                rbcSelC4(0).Move 720, 0, 760
                rbcSelC4(1).Caption = "Trade"
                rbcSelC4(1).Visible = True
                rbcSelC4(1).Move 1440, 0, 820
                rbcSelC4(2).Caption = "Both"
                rbcSelC4(2).Width = 100
                rbcSelC4(2).Visible = True
                rbcSelC4(2).Move 2265, 0, 1600
                rbcSelC4(2).Value = True
                plcSelC4.Visible = True
                'plcSelC4.Caption = "Include"
                smPaintCaption4 = "Include"
                plcSelC4_Paint
                'plcSelC4.Move 120, plcSelC5.Top + plcSelC5.Height + 30   '4-2-07
                plcSelC4.Move 120, 2310   '9/10/18

                ckcSelC12(0).Caption = "Skip Page"
                plcSelC12.Move plcSelC4.Left, plcSelC4.Top + plcSelC4.Height, 1440
                ckcSelC12(0).Visible = True
                ckcSelC12(0).Value = vbUnchecked    'False
                ckcSelC12(0).Move 0, -30, 1440
                plcSelC12.Visible = True
                
                rbcSelCInclude(0).Caption = "Contract"
                rbcSelCInclude(0).Visible = True
                rbcSelCInclude(0).Move 780, 0, 1020
                rbcSelCInclude(1).Caption = "Line"
                rbcSelCInclude(1).Visible = True
                rbcSelCInclude(1).Move 1800, 0, 680
                rbcSelCInclude(0).Value = True                  'default to contr option (vs line)
                plcSelC2.Visible = True
                'plcSelC2.Caption = "Show by"
                smPaintCaption2 = "Show by"
                plcSelC2_Paint
                plcSelC2.Move 1800, plcSelC4.Top + plcSelC4.Height
                'adjust for the advertiser box to disallow contract selectivity
                lbcSelection(0).Visible = False             'don't allow selective contracts for the summary report
                lbcSelection(5).Width = lbcSelection(1).Width   'make advt box as wide as agency box

                plcSelC8.Move 120, plcSelC2.Top + plcSelC2.Height, 4000
                ckcSelC8(0).Caption = "Discrepancies Only"
                ckcSelC8(0).Move 0, -30, 1880
                ckcSelC8(0).Visible = True


                ckcSelC8(1).Caption = "Credit Checks Only"
                ckcSelC8(1).Visible = True
                ckcSelC8(1).Move 2040, -30, 1920
                plcSelC8.Visible = True

                ckcSelC10(0).Caption = "Rates"
                plcSelC10.Move plcSelC8.Left, plcSelC8.Top + plcSelC8.Height, 720
                ckcSelC10(0).Visible = True
                ckcSelC10(0).Value = vbChecked  'True
                ckcSelC10(0).Move 0, -30, 1560
                plcSelC10.Visible = True
                '7-18-01 Gross or Net
                plcSelC7.Move 1020, plcSelC10.Top, 2400 '1560
                'plcSelC7.Caption = ""
                smPaintCaption7 = ""
                plcSelC7_Paint
                rbcSelC7(0).Move 0, 0, 900    'gross button,
                rbcSelC7(0).Caption = "Gross"
                rbcSelC7(1).Move 960, 0, 720   'net button
                rbcSelC7(1).Caption = "Net"
                rbcSelC7(1).Value = True
                rbcSelC7(2).Caption = "Acq"
                rbcSelC7(2).Move 1680, 0, 720
                rbcSelC7(0).Visible = True
                rbcSelC7(1).Visible = True
                rbcSelC7(2).Visible = True
                If rbcSelC7(0).Value Then
                    rbcSelC7_click 0
                Else
                    rbcSelC7(0).Value = True
                End If
                plcSelC7.Visible = True

'               8-14-15 move to last question on screen
'                plcSelC13.Move 2760, plcSelC7.Top, 1560
'                ckcSelC13(0).Caption = "Show Comm %"
'                ckcSelC13(0).Move 0, -30, 1560
'                ckcSelC13(0).Visible = True
'                plcSelC13.Visible = True

                '7-18-01 Comments
                mAskShowComments
                '12-7-04 Ask to include Cancel Before contracts
                ckcSelC6(4).Visible = True
                ckcSelC6(4).Caption = "CBS"
                ckcSelC6(4).Value = vbChecked           'default to include
                ckcSelC6(4).Move 0, ckcSelC6(0).Top + ckcSelC6(0).Height, 720

                mInvSortPop
                cbcSet1.ListIndex = 0
                edcSet1.Text = "Invoice Sort"
                edcSet1.Move 1320, plcSelC6.Top + 270 '+ plcSelC6.Height + 60  '7-18-01 was plcselc10
                cbcSet1.Move 1320 + edcSet1.Width, edcSet1.Top - 30, 1860
                edcSet1.Visible = True
                cbcSet1.Visible = True
                                
                plcSelC13.Move 120, plcSelC6.Top + plcSelC6.Height + 30, 1560
                ckcSelC13(0).Caption = "Show Comm %"
                ckcSelC13(0).Move 0, -30, 1560
                ckcSelC13(0).Visible = True
                lacShow.Move 0, ckcSelC13(0).Top + ckcSelC13(0).Height + 10
                rbcShow(0).Move lacShow.Width + 30, ckcSelC13(0).Top + ckcSelC13(0).Height + 10
                rbcShow(1).Move lacShow.Width + rbcShow(0).Width + 30, ckcSelC13(0).Top + ckcSelC13(0).Height + 10
                'TTP 10721 - Paperwork Summary report: add radio button next to "Show GRP/Prop Date" to show the External Contract number
                rbcShow(2).Move lacShow.Width + rbcShow(0).Width + rbcShow(1).Width + 30, ckcSelC13(0).Top + ckcSelC13(0).Height + 10
                rbcShow(0).Visible = True
                rbcShow(1).Visible = True: rbcShow(1).Width = 1695: lacShow.Visible = True
                'TTP 10721 - Paperwork Summary report: add radio button next to "Show GRP/Prop Date" to show the External Contract number
                rbcShow(2).Visible = True: rbcShow(2).Width = 2500
                plcSelC13.Height = 450  'rbcShow(0).Height + ckcSelC13(0).Height * 2
                'TTP 10721 - Paperwork Summary report: add radio button next to "Show GRP/Prop Date" to show the External Contract number
                plcSelC13.Width = lacShow.Width + rbcShow(0).Width + rbcShow(1).Width + rbcShow(2).Width + 80
                plcSelC13.Visible = True
'                pbcSelC.Height = 4250          'move this to make screen height bigger
'                frcOption.Height = 4600
            End If
            pbcSelC.Visible = True
            pbcOption.Visible = True
    End Select
    frcOption.Visible = True
End Sub

Public Sub mSellConvAirVehPop()
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(RptSelCt, VEHCONV_W_FEED + VEHCONV_WO_FEED + VEHAIRING + VEHSELLING + ACTIVEVEH, lbcSelection(3), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvAirVehPopErr
        gCPErrorMsg ilRet, "mSellConvAirVehPop (gPopUserVehicleBox: Vehicle)", RptSelCt
        On Error GoTo 0
    End If
    Exit Sub
mSellConvAirVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'
'           mAskBobInput - ask Start Qtr, Year and # Periods
'           for Billed & Booked Versions
Public Sub mAskBobYrQtrPeriods() 'VBC NR
    lacSelCFrom.Caption = "Start Quarter" 'VBC NR
    lacSelCFrom.Visible = True 'VBC NR
    lacSelCFrom.Left = 120 'VBC NR
    edcSelCFrom.MaxLength = 1 'VBC NR
    edcSelCFrom.Width = 240 'VBC NR
    edcSelCFrom.Left = 1260 'VBC NR
    edcSelCFrom.Visible = True 'VBC NR
    lacSelCFrom1.Caption = "Year" 'VBC NR
    lacSelCFrom1.Visible = True 'VBC NR
    lacSelCFrom1.Left = 1650 'VBC NR
    edcSelCFrom1.MaxLength = 4 'VBC NR
    edcSelCFrom1.Move 2100, edcSelCFrom.Top, 600 'VBC NR
    edcSelCFrom1.Visible = True 'VBC NR

    '3-2-02 Move contract # and ask # periods

    lacSelCTo1.Move 2880, lacSelCFrom.Top, 1200 'VBC NR
    lacSelCTo1.Caption = "# Periods" 'VBC NR
    edcSelCTo1.MaxLength = 2 'VBC NR
    edcSelCTo1.Move 3780, edcSelCFrom.Top, 360 'VBC NR
    edcSelCTo1.Text = 12 'VBC NR
    edcSelCTo1.Visible = True 'VBC NR
    lacSelCTo1.Visible = True 'VBC NR
End Sub 'VBC NR

'
'           mAskBobCorpOrStd - ask Corporate or Std reporting
'           for Billed & Booked versions
'
Public Sub mAskBOBCorpOrStd()
    'plcSelC9.Move 120, plcSelC4.Top + plcSelC4.Height, 2600
    smPaintCaption9 = "Month"
    plcSelC9_Paint
    rbcSelC9(0).Caption = "Corp"
    rbcSelC9(0).Left = 660
    rbcSelC9(0).Width = 720
    rbcSelC9(1).Caption = "Std"
    rbcSelC9(1).Left = 1440
    rbcSelC9(1).Width = 600

    rbcSelC9(2).Caption = "Cal"
    rbcSelC9(2).Move 2080, 0, 600

    rbcSelC9(3).Caption = "Cal (Spots)"
    rbcSelC9(3).Move 2740, 0, 1200
    rbcSelC9(3).Visible = False     'turn off for all places; B & B will enable it
    
    rbcSelC9(4).Move 660, 210, 1440, 120
    rbcSelC9(4).Visible = False     '1-12-21 B & b will turn this on
    rbcSelC9(0).Visible = True
    rbcSelC9(1).Visible = True
    rbcSelC9(2).Visible = True
    rbcSelC9(1).Value = True
    plcSelC9.Visible = True
        
    Exit Sub
End Sub

Private Sub mAskCntrAndSpotTypesForAvails()
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = CONTRACTSJOB Then
        If (igRptType = 0) And (ilListIndex > 1) Then
            ilListIndex = ilListIndex + 1
        End If
    End If

    smPaintCaption10 = "Include"
    plcSelC10_Paint
    ckcSelC10(0).Caption = "Holds"
    ckcSelC10(0).Move 660, -30, 840
    ckcSelC10(0).Value = vbChecked   'True
    If ckcSelC10(0).Value = vbChecked Then
        ckcSelC10_click 0
    Else
        ckcSelC10(0).Value = vbChecked   'True
    End If
    ckcSelC10(0).Visible = True
    ckcSelC10(1).Value = vbChecked   'True
    ckcSelC10(1).Caption = "Orders"
    ckcSelC10(1).Move 1500, -30, 900
    If ckcSelC10(1).Value = vbChecked Then
        ckcSelC10_click 1
    Else
        ckcSelC10(1).Value = vbChecked   'True
    End If
    ckcSelC10(1).Visible = True

    plcSelC10.Visible = True
    'Contract Type selection
    plcSelC5.Move 120, plcSelC10.Top + plcSelC10.Height, 4260    '5-16-05 chf from plcselc3 to plcselc10
    plcSelC5.Height = 440
    'plcSelC5.Caption = ""
    smPaintCaption5 = ""
    plcSelC5_Paint
    ckcSelC5(0).Move 240, -30, 1080     '660
    ckcSelC5(0).Caption = "Standard"
    If ckcSelC5(0).Value = vbChecked Then
        ckcSelC5_click 0
    Else
        ckcSelC5(0).Value = vbChecked   'True
    End If
    ckcSelC5(0).Visible = True
    ckcSelC5(1).Move 1440, -30, 1200        '1800
    ckcSelC5(1).Caption = "Reserved"
    If ckcSelC5(1).Value = vbChecked Then
        ckcSelC5_click 1
    Else
        ckcSelC5(1).Value = vbChecked   'True
    End If
    ckcSelC5(1).Visible = True
    If tgUrf(0).iSlfCode > 0 Then           'its a slsp thats is asking for this report,
                                            'don't allow them to exclude reserves
        ckcSelC5(1).Enabled = False
    Else
        ckcSelC5(1).Enabled = True
    End If
    ckcSelC5(2).Move 2760, -30, 1080        '3000
    ckcSelC5(2).Caption = "Remnant"
    If ckcSelC5(2).Value = vbChecked Then
        ckcSelC5_click 2
    Else
        ckcSelC5(2).Value = vbChecked   'True
    End If
    ckcSelC5(2).Visible = True
    ckcSelC5(3).Move 240, 195, 600      '660
    ckcSelC5(3).Caption = "DR"
    If ckcSelC5(3).Value = vbChecked Then
        ckcSelC5_click 3
    Else
        ckcSelC5(3).Value = vbChecked   'True
    End If
    ckcSelC5(3).Visible = True
    ckcSelC5(4).Move 840, 195, 1320    '1260
    ckcSelC5(4).Caption = "Per Inquiry"
    If ckcSelC5(4).Value = vbChecked Then
        ckcSelC5_click 4
    Else
        ckcSelC5(4).Value = vbChecked   'True
    End If
    ckcSelC5(4).Visible = True
    ckcSelC5(5).Move 2160, 195, 720 '2580
    ckcSelC5(5).Caption = "PSA"
    ckcSelC5(5).Value = vbUnchecked 'False
    ckcSelC5(5).Visible = True  '9-12-02 vbChecked 'True
    ckcSelC5(6).Move 2880, 195, 900     '3300
    ckcSelC5(6).Caption = "Promo"
    ckcSelC5(6).Value = vbUnchecked 'False
    ckcSelC5(6).Visible = True
    plcSelC5.Visible = True

    plcSelC6.Visible = True
    plcSelC6.Move plcSelC5.Left, plcSelC5.Top + plcSelC5.Height
    'plcSelC6.Caption = ""
    smPaintCaption6 = ""
    plcSelC6_Paint
    ckcSelC6(0).Move 240, -30, 840  '660
    ckcSelC6(0).Caption = "Trade"
    If ckcSelC6(0).Value = vbChecked Then
        ckcSelC6_click 0
    Else
        ckcSelC6(0).Value = vbChecked   'True
    End If
    ckcSelC6(0).Visible = True
    ckcSelC6(1).Caption = "Missed"
    ckcSelC6(1).Visible = True
    ckcSelC6(1).Move 1080, -30, 960     '1500
    If ckcSelC6(1).Value = vbChecked Then
        ckcSelC6_click 1
    Else
        ckcSelC6(1).Value = vbChecked   'True
    End If

    ckcSelC6(4).Visible = True
    ckcSelC6(4).Caption = "Locked Avails"
    ckcSelC6(4).Value = vbChecked
    ckcSelC6(4).Move 2220, -30, 1560    '2460

    '5-16-05 chg from plcselc5 to plcselc3 for spot types
    plcSelC3.Visible = True
    smPaintCaption3 = ""
    plcSelC3_Paint

    plcSelC3.Move 0, plcSelC6.Top + plcSelC6.Height, 4300, 440
    ckcSelC3(0).Caption = "Charge"
    ckcSelC3(0).Move 360, -30, 960
    ckcSelC3(0).Visible = True
    ckcSelC3(0).Value = vbChecked   'True
    ckcSelC3(1).Caption = "0.00"
    ckcSelC3(1).Move 1320, -30, 720
    ckcSelC3(1).Visible = True
    ckcSelC3(1).Value = vbChecked   'True
    ckcSelC3(2).Caption = "ADU"
    ckcSelC3(2).Move 2040, -30, 720
    ckcSelC3(2).Visible = True
    ckcSelC3(2).Value = vbChecked   'True
    ckcSelC3(3).Value = vbChecked   'True
    ckcSelC3(3).Caption = "Bonus"
    ckcSelC3(3).Move 2760, -30, 840
    ckcSelC3(3).Visible = True
    ckcSelC3(4).Value = vbChecked   'True
    ckcSelC3(4).Caption = "+Fill"
    'ckcSelC3(4).Move 360, 195, 720
    ckcSelC3(4).Move 3720, -30, 720
    ckcSelC3(4).Visible = True
    ckcSelC3(4).Value = vbChecked   'True
    ckcSelC3(5).Value = vbChecked   'True
    ckcSelC3(5).Caption = "-Fill"
    'ckcSelC3(5).Move 1080, 195, 600
    ckcSelC3(5).Move 360, 195, 600
    ckcSelC3(5).Visible = True
    ckcSelC3(6).Value = vbChecked   'True
    ckcSelC3(6).Caption = "N/C"
    'ckcSelC3(6).Move 1800, 195, 1320
    ckcSelC3(6).Move 1080, 195, 600
    ckcSelC3(6).Visible = True
    ckcSelC3(7).Value = vbChecked   'True
    ckcSelC3(7).Caption = "MG"
    ckcSelC3(7).Move 1800, 195, 600
    ckcSelC3(7).Visible = True             '10-29-10 make option,
    ckcSelC3(8).Value = vbChecked   'True
    ckcSelC3(8).Caption = "Recap"
    'ckcSelC3(8).Move 360, 420, 1440
    ckcSelC3(8).Move 2520, 195, 1440
    ckcSelC3(8).Visible = True
    ckcSelC3(8).Value = vbChecked   'True
    ckcSelC3(9).Value = vbChecked   'True
    ckcSelC3(9).Caption = "Spinoff"
    'ckcSelC3(9).Move 1920, 420, 960
    ckcSelC3(9).Move 3440, 195, 960
    ckcSelC3(9).Visible = True

    plcSelC3.Visible = True
End Sub

Public Sub mStartEndDates()
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If (igRptType = 0) And (ilListIndex > 1) Then
        ilListIndex = ilListIndex + 1
    End If

    lacSelCFrom.Caption = "Dates- Start"
    lacSelCFrom.Visible = True
    lacSelCFrom.Move 15, 105, 1440
    edcSelCFrom.MaxLength = 10
    edcSelCFrom.Move 1080, 60, 1080

    lacSelCFrom1.Caption = "End"

    lacSelCFrom1.Move 2350, lacSelCFrom.Top, 360
    lacSelCFrom1.Visible = True

    edcSelCFrom1.MaxLength = 10
    edcSelCFrom1.Move 2880, edcSelCFrom.Top, 1080
    edcSelCFrom.Visible = True
    edcSelCFrom1.Visible = True
    
    'Date: 12/4/2019   added CSI calendar controls for date entries
    If (ilListIndex = CNT_GAMESUMMARY) Or (ilListIndex = CNT_LOCKED) Or (ilListIndex = CNT_PAPERWORKTAX) Or (ilListIndex = CNT_VEH_UNITCOUNT) Then
        edcSelCFrom.Visible = False
        edcSelCFrom1.Visible = False
        CSI_CalFrom.Visible = True
        CSI_CalFrom.Move edcSelCFrom.Left, 60, 1080
        CSI_CalFrom.ZOrder 0
        CSI_CalTo.Visible = True
        CSI_CalTo.Move lacSelCFrom1.Left + lacSelCFrom1.Width + 10, 60, 1080
        CSI_CalTo.ZOrder 0
    End If
End Sub

'
'           populate only sports vehicles
'
Public Sub mSportsVehPop(ilIndex)
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(RptSelCt, VEHSPORT + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSportsVehPopErr
        gCPErrorMsg ilRet, "mSportsVehPop (gPopUserVehicleBox: Vehicle)", RptSelCt
        On Error GoTo 0
    End If
    Exit Sub
mSportsVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'
'    Fill the "Select" options for the Sales Comparison report
'
Public Sub mFillSalesCompare(ilListIndex As Integer, cbcCombo As Control, ilShowNone As Integer)
    cbcCombo.Clear
    If ilShowNone Then
        cbcCombo.AddItem "None"
    End If
    cbcCombo.AddItem "Advertiser"
    cbcCombo.AddItem "Agency"
    cbcCombo.AddItem "Bus Category"
    cbcCombo.AddItem "Prod Protection"
    cbcCombo.AddItem "Salesperson"
    cbcCombo.AddItem "Vehicle"
    If ilListIndex <> CNT_BOBCOMPARE Then
        cbcCombo.AddItem "Vehicle Group"
    End If
    cbcCombo.ListIndex = 0
End Sub

'       Ask Year, Month, # Periods
'       Yar [    ]   Month  [   ]    # Periods [  ]
'
Public Sub mAskYrMonthPeriods(ilListIndex As Integer)
    edcSelCFrom1.Text = ""
    lacSelCFrom.Caption = "Year"
    lacSelCFrom.Width = 600
    edcSelCFrom.Move 600, 30, 600
    lacSelCFrom.Visible = True
    edcSelCFrom.Visible = True
    lacSelCFrom1.Caption = "Start Month"
    lacSelCFrom1.Move 1440, 60, 1320
    edcSelCFrom1.Move 2460, 30, 490
    lacSelCFrom1.Visible = True
    edcSelCFrom1.Visible = True

    'If (ilListIndex = CNT_BOB Or ilListIndex = CNT_BOBCOMPARE Or ilListIndex = CNT_BOBRECAP) And igRptCallType = CONTRACTSJOB Then
    If (ilListIndex = CNT_BOB Or ilListIndex = CNT_BOBRECAP) And igRptCallType = CONTRACTSJOB Then
        lacSelCTo1.Caption = "# Months"
        lacSelCTo1.Move 3150, 60, 1080
        lacSelCTo1.Visible = True
        edcSelCTo1.Visible = True
        edcSelCTo1.Move 4020, 30, 360
        edcSelCTo1.Text = "12"
    Else            'B & B comparisons, Sales Comparisons
        lacSelCTo.Caption = "# Months"
        lacSelCTo.Move 3150, 60, 1080
        lacSelCTo.Visible = True
        edcSelCTo.Visible = True
        edcSelCTo.Move 4020, 30, 360
        edcSelCTo.Text = "12"
    End If
End Sub

'               Loop thru Email_Pdf array and create a seprate pdf for each unique contract/vehicle
'               mCreateInsertionEmailPdfs
'               <input> tgEmail_Pdf global array
'
Private Sub mCreateInsertionEmailPdfs()
    Dim slDate As String
    Dim slStr As String
    Dim slTime As String
    Dim ilLoopOnDiff As Integer
    Dim slPDFFileName As String
    Dim ilRet As Integer
    Dim blRet As Boolean

    ReDim tmEmail_Info(LBound(tgEmail_PDFs) To UBound(tgEmail_PDFs)) As EMAILINFO
    mFormatGenDateTime slDate, slTime
    
    For ilLoopOnDiff = LBound(tgEmail_PDFs) To UBound(tgEmail_PDFs) - 1
        'new selection based on unique vehicle & contract
        slStr = Trim$(sgSelection) & " and (({CBF_Contract_BR.cbfchfCode} = " & Trim$(str(tgEmail_PDFs(ilLoopOnDiff).lChfCode)) & ") and ({CBF_Contract_BR.cbfvefCode} = " & Trim$(str(tgEmail_PDFs(ilLoopOnDiff).iVefCode)) & " or {CBF_Contract_BR.cbfvefCode} = 0" & "))"
        If Not gSetSelection(slStr) Then
            Return
        End If

        'set blKeepReportOpen to TRUE,
        'other Error message Object Variable or with block not set occurs
        slStr = Trim$(tgEmail_PDFs(ilLoopOnDiff).sVefName) & "_" & Trim$(str(tgEmail_PDFs(ilLoopOnDiff).lCntrNo)) & "_" & Trim$(tgEmail_PDFs(ilLoopOnDiff).sAdvtName) & "_R" & Trim$(str(tgEmail_PDFs(ilLoopOnDiff).iCntRevNo)) & "_" & Trim$(slDate) & "_" & Trim$(slTime)
        'filename:  Vehicle, Cnt #, AdvtName,Revision #, CurrentDate Genned,Current Time Genned
        slPDFFileName = gStripCntrlChars(slStr)
        
        ilRet = gExportCRW(slPDFFileName, 0, True)
        ogReport.DiscardSavedData = True
        
        'prepare array of generated pdfs to send
        tmEmail_Info(ilLoopOnDiff).iVefCode = tgEmail_PDFs(ilLoopOnDiff).iVefCode
        tmEmail_Info(ilLoopOnDiff).lChfCode = tgEmail_PDFs(ilLoopOnDiff).lChfCode
        tmEmail_Info(ilLoopOnDiff).lCntrNo = tgEmail_PDFs(ilLoopOnDiff).lCntrNo
        tmEmail_Info(ilLoopOnDiff).lEmfCode = tgEmail_PDFs(ilLoopOnDiff).lEmfCode
        tmEmail_Info(ilLoopOnDiff).sPDFFileName = slPDFFileName                                 'pdf file dynamically generated
        tmEmail_Info(ilLoopOnDiff).sResponseDate = tgEmail_PDFs(ilLoopOnDiff).sResponseDate     'user input
        tmEmail_Info(ilLoopOnDiff).sRouteTo = "S"               'route to station
        tmEmail_Info(ilLoopOnDiff).sType = "I"                  'I = Insertion order
        tmEmail_Info(ilLoopOnDiff).sAdvtName = tgEmail_PDFs(ilLoopOnDiff).sAdvtName
        tmEmail_Info(ilLoopOnDiff).sProduct = tgEmail_PDFs(ilLoopOnDiff).sProduct
        tmEmail_Info(ilLoopOnDiff).sAgyEstNo = tgEmail_PDFs(ilLoopOnDiff).sAgyEstNo
        tmEmail_Info(ilLoopOnDiff).sStartDate = tgEmail_PDFs(ilLoopOnDiff).sStartDate
        tmEmail_Info(ilLoopOnDiff).sEndDate = tgEmail_PDFs(ilLoopOnDiff).sEndDate
    Next ilLoopOnDiff
    PEClosePrintJob     'export was keeping report open (gexportcrw with true parameter); need to close , all done
    
    'D.S. 07-13-15
    blRet = gSubmitApprovalRequest(tmEmail_Info())
    Exit Sub
End Sub

'               convert generation date and time to string for Email PDF filename
'               Remove slash and replace with dashes in date, remove colon in time
'           mFormatGenDateTime()
'           <input>  global date variable:  igNowDate(0 to 1)
'                    global time variable:  lgNowTime
'           <output>  slDate
'                     slTime
Public Sub mFormatGenDateTime(slDate As String, slTime As String)
    Dim slTemp As String
    Dim illoop As Integer
    Dim slStr As String
    Dim llDate As Long

    gUnpackDateLong igNowDate(0), igNowDate(1), llDate
     slStr = Format$(llDate, "m/d/yy")               'Now date as string
    'replace slash with dash in date
     slDate = ""
     For illoop = 1 To Len(slStr) Step 1
         slTemp = Mid$(slStr, illoop, 1)
         If slTemp <> "/" Then
             slDate = Trim$(slDate) & Trim$(slTemp)
         Else
             slDate = Trim$(slDate) & "-"
         End If
     Next illoop
     
     slStr = gFormatTimeLong(lgNowTime, "A", "1")
     'remove colons from time
     slTime = ""
     For illoop = 1 To Len(slStr) Step 1
         slTemp = Mid$(slStr, illoop, 1)
         If slTemp <> ":" Then
             slTime = Trim$(slTime) & Trim$(slTemp)
         End If
     Next illoop
     Do While Len(slTime) < 8
         slTime = "0" & slTime
     Loop
     
     
     Exit Sub
End Sub

'             Format the filename required for Proposal/contract email pdf
'             one contract per pdf which can contain a detail , and multiple severals summaries
'                mCreateEmailOrderFileName
'               <input> None
Public Sub mCreateOrderEmailPDF()
    Dim ilLoopOnDiff As Integer
    Dim slPDFFileName As String
    Dim slDate As String
    Dim slTime As String
    Dim ilRet As Integer
    
    For ilLoopOnDiff = LBound(tgEmail_PDFs) To UBound(tgEmail_PDFs) - 1
        ogReport.AddToSelection = " and ({chf_contract_header.chfcode} = " & tgEmail_PDFs(ilLoopOnDiff).lChfCode & ")"
        mFormatGenDateTime slDate, slTime
        'set blKeepReportOpen to TRUE,
        'other Error message Object Variable or with block not set occurs
        slPDFFileName = "Cnt" & Trim$(str(tgEmail_PDFs(ilLoopOnDiff).lCntrNo)) & " " & Trim$(tgEmail_PDFs(ilLoopOnDiff).sAdvtName) & " R" & Trim$(str(tgEmail_PDFs(ilLoopOnDiff).iCntRevNo)) & " " & Trim$(slDate) & "_" & Trim$(slTime)
        'filename: Cnt #, AdvtName,Revision #, CurrentDate Genned,Current Time Genned
        ilRet = gExportCRW(slPDFFileName, 0, True)
        ogReport.DiscardSavedData = True
    Next ilLoopOnDiff
    PEClosePrintJob     'export was keeping report open (gexportcrw with true parameter); need to close , all done
    Exit Sub
End Sub

Private Sub mEMailContentPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim illoop As Integer
    Dim slSvName As String
    Dim slName As String
    Dim slNameCode As String
    Dim slCode As String
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer

    slSvName = Trim$(cbcEMailContent.Text)
    ilfilter(0) = CHARFILTER
    slFilter(0) = sgEMailContentType
    ilOffSet(0) = gFieldOffset("Emf", "EmfType") '2
    ilRet = gLPopListBox(RptSelCt, tmEMailContentCode(), smEMailContentCodeTag, "Emf.btr", gFieldOffset("Emf", "EmfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        imChgMode = True
        cbcEMailContent.Clear
        cbcEMailContent.AddItem ("[New]")
        cbcEMailContent.SetItemData = -1  ' Indicates a new title
        
        For illoop = 0 To UBound(tmEMailContentCode) - 1 Step 1
            slNameCode = tmEMailContentCode(illoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                cbcEMailContent.AddItem (Trim$(slName))
                cbcEMailContent.SetItemData = slCode  ' Indicates a new title
            End If
        Next illoop
        cbcEMailContent.SelText (Trim(slSvName))
        imChgMode = False
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEMailContentBranch             *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to event  *
'*                      type and process               *
'*                      communication back from event  *
'*                      type                           *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mEMailContentBranch() As Integer
'
'   ilRet = mEMailContentBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    
    If Not imDoubleClickName And (cbcEMailContent.ListIndex <> 0) Then
        imDoubleClickName = False
        mEMailContentBranch = False
        Exit Function
    End If
    If cbcEMailContent.Text = "[New]" Then
        sgEMailContentName = ""
    Else
        sgEMailContentName = Trim$(cbcEMailContent.Text)
    End If
    EMailContent.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgEMailContentName)
    igEMailContentCallSource = Val(sgEMailContentName)
    ilParse = gParseItem(slStr, 2, "\", sgEMailContentName)
    imDoubleClickName = False
    mEMailContentBranch = True
    smEMailContentCodeTag = ""
    mEMailContentPop
    If imTerminate Then
        mEMailContentBranch = False
        Exit Function
    End If
    If igEMailContentCallSource = CALLDONE Then  'Done
        igEMailContentCallSource = CALLNONE
        cbcEMailContent.SelText (Trim(sgEMailContentName))
        sgEMailContentName = ""
        
        If cbcEMailContent.ListIndex > 0 Then
            imChgMode = True
            mEMailContentBranch = False
            imEMailContentSelectedIndex = cbcEMailContent.ListIndex
            imChgMode = False
        Else
            imChgMode = True
            cbcEMailContent.ListIndex = 0
            imChgMode = False
            cbcEMailContent.SetFocus
            Exit Function
        End If
    End If
    If igEMailContentCallSource = CALLCANCELLED Then  'Cancelled
        'mEMailContentBranch = False
        igEMailContentCallSource = CALLNONE
        sgEMailContentName = ""
        cbcEMailContent.SetFocus
        Exit Function
    End If
    If igEMailContentCallSource = CALLTERMINATED Then
        'mEMailContentBranch = False
        igEMailContentCallSource = CALLNONE
        sgEMailContentName = ""
        cbcEMailContent.SetFocus
        Exit Function
    End If
    mSetCommands
    Exit Function
mEMailContentBranchErr:
    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

Private Sub mAgyAdvtPop(lbcSelection As Control)            '12-9-16
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    ilRet = gPopAgyCollectBox(RptSel, "A", lbcSelection, lbcAgyAdvtCode)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgyAdvtPopErr
        gCPErrorMsg ilRet, "mAgyAdvtPop (gPopAgyCollectBox)", RptSel
        On Error GoTo 0
    End If
    Exit Sub
mAgyAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    'Unload RptSel
    'Set RptSel = Nothing   'Remove data segment
    Exit Sub
End Sub

'           Contract jobs selectivity- break up into different modules due to size
'
Public Sub mCntSelectivity4()
    Dim ilRet As Integer
    Dim ilListIndex As Integer
    Dim ilSort As Integer
    Dim ilShow As Integer

     ilListIndex = lbcRptType.ListIndex
     If (igRptType = 0) And (ilListIndex > 1) Then
         ilListIndex = ilListIndex + 1
     End If

    If ilListIndex = CNT_SALES_CPPCPM Then
         ilRet = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(11), tgRptSelDemoCodeCT(), sgRptSelDemoCodeTagCT, "D")
         lbcSelection(11).Width = 4380
         lbcSelection(11).Left = 15
         lbcSelection(11).Visible = True         'see demo list box
         ckcAll.Caption = "All Demos"
         ckcAll.Visible = True
         ckcAll.Enabled = True
         edcSelCFrom.Text = ""
         edcSelCTo.Text = ""
         edcSelCTo1.Text = ""
         mAskEffDate
         
         'Date: 12/12/2019 added CSI calendar control for date entry
         CSI_CalFrom.Move 1340, edcSelCFrom.Top, 1080
         CSI_CalFrom.Visible = True: CSI_CalFrom.ZOrder 0
         edcSelCFrom.Visible = False
         
         plcSelC2.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
         plcSelC2.Height = 240   '6-4-04 chged from 440 to 240
         'plcSelC2.Caption = "Month"
         smPaintCaption2 = "Month"
         plcSelC2_Paint
         plcSelC2.Visible = True
         rbcSelCInclude(0).Caption = "Corporate"
         rbcSelCInclude(0).Move 660, 0, 1140
         rbcSelCInclude(0).Visible = True
         rbcSelCInclude(1).Caption = "Standard"
         rbcSelCInclude(1).Move 1840, 0, 1140
         rbcSelCInclude(1).Visible = True
         If rbcSelCInclude(1).Value Then             'default to std
             rbcSelCInclude_Click 1
         Else
             rbcSelCInclude(1).Value = True
         End If
         If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
             rbcSelCInclude(0).Enabled = False
         Else
             rbcSelCInclude(0).Value = True
         End If
         rbcSelCInclude(2).Visible = False
         
         plcSelC4.Move plcSelC2.Left, plcSelC2.Top + plcSelC2.Height + 30, 3000
         rbcSelC4(0).Move 480, 0, 840
         rbcSelC4(1).Move 1440, 0, 600
         rbcSelC4(0).Caption = "Gross"
         rbcSelC4(1).Caption = "Net"
         rbcSelC4(0).Visible = True
         rbcSelC4(1).Visible = True
         rbcSelC4(2).Visible = False
         rbcSelC4(0).Value = True
         plcSelC4.Visible = True
         smPaintCaption4 = "By"
         
         '6-4-04 selective contract
         lacTopDown.Caption = "Contract #"
         lacTopDown.Move 120, plcSelC4.Top + plcSelC4.Height + 30
         lacTopDown.Visible = True
         edcTopHowMany.Move 1290, plcSelC4.Top + plcSelC4.Height, 945
         edcTopHowMany.MaxLength = 9
         edcTopHowMany = ""
         edcTopHowMany.Visible = True
     ElseIf ilListIndex = CNT_VEHCPPCPM Then
         plcSelC4.Move 120, 0
         'plcSelC4.Caption = "Show "
         smPaintCaption4 = "Show "
         plcSelC4_Paint
         rbcSelC4(0).Caption = "CPP"
         rbcSelC4(0).Move 720, 0, 630
         rbcSelC4(1).Caption = "CPM"
         rbcSelC4(1).Move 1440, 0, 700
         plcSelC4.Visible = True
         rbcSelC4(0).Visible = True
         rbcSelC4(1).Visible = True
         rbcSelC4(2).Visible = False
         If rbcSelC4(0).Value Then             'default to CPP
             rbcSelC4_click 0
         Else
             rbcSelC4(0).Value = True
         End If
         lacSelCFrom.Move plcSelC4.Left, plcSelC4.Top + plcSelC4.Height + 30
         edcSelCFrom.Move 1350, lacSelCFrom.Top - 30, 1020
         edcSelCFrom.MaxLength = 10  '8    5/27/99 changed for short form date m/d/yyyy
         lacSelCFrom.Caption = "Effective Date"
         lacSelCFrom.Visible = True
         edcSelCFrom.Visible = True

         'Date: 12/12/2019 added CSI calendar control for date entry
         CSI_CalFrom.Move 1340, edcSelCFrom.Top, 1080
         CSI_CalFrom.Visible = True
         CSI_CalFrom.ZOrder 0
         edcSelCFrom.Visible = False
         
         '1-30-18 implement gross net option
         plcSelC11.Move plcSelC4.Left, edcSelCFrom.Top + edcSelCFrom.Height + 60, 3000
         rbcSelC11(0).Move 480, 0, 840
         rbcSelC11(1).Move 1440, 0, 600
         rbcSelC11(0).Caption = "Gross"
         rbcSelC11(1).Caption = "Net"
         rbcSelC11(0).Visible = True
         rbcSelC11(1).Visible = True
         rbcSelC11(2).Visible = False
         rbcSelC11(0).Value = True
         plcSelC11.Visible = True
         smPaintCaption11 = "By"
         ilRet = gPopMnfPlusFieldsBox(RptSelCt, lbcSelection(2), tgRptSelDemoCodeCT(), sgRptSelDemoCodeTagCT, "D")
         ilSort = 1  'sort books names by date, then name
         ilShow = 1  'show books names with date
         'ilRet = gPopBookNameBox(RptSelCt, 0, ilSort, ilShow, lbcSelection(4), tgBookName(), sgBookNameTag)
         lbcSelection(3).Move 120, ckcAll.Top + ckcAll.Height + 30, 4380, 1500
         lbcSelection(3).Visible = True
         ckcAll.Caption = "All Vehicles"
         ckcAll.Visible = True
         lbcSelection(4).Visible = True          'book names
         lbcSelection(2).Visible = True          'Demo Names
         lbcSelection(2).Move lbcSelection(3).Left, lbcSelection(3).Top + lbcSelection(3).Height + 300, lbcSelection(3).Width / 2, lbcSelection(3).Height
         lbcSelection(4).Move lbcSelection(3).Left + lbcSelection(3).Width / 2 + 60, lbcSelection(3).Top + lbcSelection(3).Height + 300, lbcSelection(3).Width / 2, lbcSelection(3).Height
         laclbcName(0).Visible = False
         laclbcName(1).Visible = True
         laclbcName(1).Caption = "Rate Card"
         ckcAllAAS.Move ckcAll.Left, lbcSelection(3).Height + ckcAll.Height + 90
         ckcAllAAS.Caption = "All Demos"
         ckcAllAAS.Visible = True
         laclbcName(0).Move lbcSelection(3).Left, lbcSelection(2).Top - laclbcName(0).Height - 30, 1605
         laclbcName(1).Move lbcSelection(3).Left + lbcSelection(3).Width / 2 + 60, lbcSelection(2).Top - laclbcName(1).Height - 30, 1710

     End If
     Exit Sub
End Sub

'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
'*******************************************************
'*                                                     *
'*      Procedure Name:mReportSeparateOutputVehicle    *
'*                                                     *
'*             Created:09/20/21      By:J.White        *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Gets a list of ListIndexes     *
'*                      for the Vehicle select box     *
'*                      where the vehicle is present   *
'*                      on the selected (or entered)   *
'*                      Contract Number                *
'*                                                     *
'*******************************************************
Sub mReportSeparateOutputVehicle()
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slName As String
    Dim slCode As String
    Dim ilTemp As Integer
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    Dim llCntrNo As Long
    
    ReDim tmVehicleList(0 To 0) As Integer
    If RptSelCt!lbcSelection(6).SelCount <= 0 Then
        RptSelCt!CkcAllveh.Value = vbChecked
    End If
    
    llCntrNo = 0
    If Val(edcTopHowMany.Text) > 0 Then
        'Include Contract #
        llCntrNo = Val(edcTopHowMany.Text)
    Else
        If lbcSelection(0).ListIndex > -1 Then
            'Include Contract #
            If Val(lbcSelection(0).List(lbcSelection(0).ListIndex)) > 0 Then
                llCntrNo = Val(lbcSelection(0).List(lbcSelection(0).ListIndex))
            End If
        End If
    End If
    
    For illoop = 0 To lbcSelection(6).ListCount - 1
        If lbcSelection(6).Selected(illoop) = True Then
            slNameCode = tgCSVNameCode(illoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'check if this vehicle is used on the selected contract
            slSQLQuery = ""
            slSQLQuery = slSQLQuery & " SELECT  chfCode,  chfCntrNo as CntrNo,  CLF.clfVefCode as VefCode"
            slSQLQuery = slSQLQuery & " FROM  ""CHF_Contract_Header"" chf"
            slSQLQuery = slSQLQuery & " JOIN ""CLF_Contract_Line"" clf ON clf.clfchfCode = chf.chfcode"
            slSQLQuery = slSQLQuery & " WHERE chfCntrNo = " & llCntrNo
            slSQLQuery = slSQLQuery & " AND clf.clfVefCode = " & slCode
            slSQLQuery = slSQLQuery & " And chfDelete <> 'Y' "
            'Fix TTP 10271 per Jason Email: Tue 9/28/21 10:22 AM (Issue #5)
            If ckcSelC12(0).Value = 1 Then 'Include NTR
                slSQLQuery = slSQLQuery & " UNION "
                slSQLQuery = slSQLQuery & " SELECT  chfCode,  chfCntrNo as CntrNo,  sbfBillVefCode  as VefCode"
                slSQLQuery = slSQLQuery & " FROM  ""CHF_Contract_Header"" chf"
                slSQLQuery = slSQLQuery & " JOIN  ""SBF_Special_Billing"" sbf ON sbf.sbfchfCode = chf.chfcode"
                slSQLQuery = slSQLQuery & " Where chfCntrNo = " & llCntrNo
                slSQLQuery = slSQLQuery & " AND sbf.sbfBillVefCode = " & slCode
                slSQLQuery = slSQLQuery & " And chfDelete <> 'Y' "
            End If
            Set rst_Temp = gSQLSelectCall(slSQLQuery)
            If Not rst_Temp.EOF Then
                tmVehicleList(UBound(tmVehicleList)) = illoop
                ReDim Preserve tmVehicleList(0 To UBound(tmVehicleList) + 1)
            End If
        End If
    Next illoop
End Sub

'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetSeparateFilename            *
'*                                                     *
'*             Created:09/20/21      By:J.White        *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Gets the Filename for a Select *
'*                      Vehicle and Contract           *
'*                                                     *
'*******************************************************
Function mGetSeparateFilename(slPrefix As String, slVefName As String) As String
    Dim ilContractNo As Long
    Dim ilRet As Integer
    Dim slRepeat As String
    slRepeat = "A"
    'TTP 10373 - Insertion orders Advertiser in Filename...
    Dim illoop As Integer
    Dim ilCount As Integer
    'Fix TTP 10271 per Jason Email: Tue 9/28/21 10:22 AM (Issue #2,#3,#4)
    Do
        mGetSeparateFilename = slPrefix
        'TTP 10373 - Insertion orders Advertiser in Filename...
        If rbcSelCSelect(0).Value = True Then 'by Advertiser
            'get the "selected" adv name
            For illoop = 0 To RptSelCt.lbcSelection(5).ListCount - 1
                If RptSelCt.lbcSelection(5).Selected(illoop) = True Then
                    ilCount = ilCount + 1
                End If
            Next illoop
            'Only add Adv name if ONE adv is selected
            If ilCount = 1 Then
                mGetSeparateFilename = mGetSeparateFilename + "-" & gFileNameFilterNotPath(RptSelCt.lbcSelection(5).Text)
            End If
        End If
        'Vehicle Name
        If Trim(slVefName) <> "" Then
            mGetSeparateFilename = mGetSeparateFilename & "-" & slVefName
        End If
        'Contract #
        If Val(edcTopHowMany.Text) > 0 Then
            'Include Contract #
            mGetSeparateFilename = mGetSeparateFilename & "-CntrNo" & Val(edcTopHowMany.Text)
        Else
            If lbcSelection(0).ListIndex > -1 Then
                'Include Contract #
                If Val(lbcSelection(0).List(lbcSelection(0).ListIndex)) > 0 Then
                    mGetSeparateFilename = mGetSeparateFilename & "-CntrNo" & Val(lbcSelection(0).List(lbcSelection(0).ListIndex))
                End If
            End If
        End If
        
        'Run Date
        ilRet = 0
        mGetSeparateFilename = mGetSeparateFilename & "-"
        mGetSeparateFilename = mGetSeparateFilename & Format(gNow, "mmddyy")
        'Unique filename
        mGetSeparateFilename = mGetSeparateFilename & slRepeat
        'mGetSeparateFilename = mGetSeparateFilename & " " & gFilterNameMatchingKeyPressCheck(Trim$(smClientName))
        'file extension
        mGetSeparateFilename = mGetSeparateFilename & "."
        Select Case cbcFileType.ListIndex
            Case 7 'Rich Text File(RTF)
                mGetSeparateFilename = mGetSeparateFilename & "rtf"
            Case 6 'Comma Separated Values(CSV)
                mGetSeparateFilename = mGetSeparateFilename & "csv"
            Case 5 'Text(TXT)
                mGetSeparateFilename = mGetSeparateFilename & "txt"
            Case 4 'Word(DOC)
                mGetSeparateFilename = mGetSeparateFilename & "doc"
            Case 3 'Excel(XLS)-No headers
                mGetSeparateFilename = mGetSeparateFilename & "xls"
            Case 2 'Excel(XLS)-Column headers
                mGetSeparateFilename = mGetSeparateFilename & "xls"
            Case 1 'Excel(XLS)-All headers
                mGetSeparateFilename = mGetSeparateFilename & "xls"
            Case 0 'Adobe Acrobat(PDF)
                mGetSeparateFilename = mGetSeparateFilename & "pdf"
        End Select
        'Check if exists, make new unique filename character
        ilRet = gFileExist(sgExportPath & CleanFilename(mGetSeparateFilename))
        If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
            slRepeat = Chr(Asc(slRepeat) + 1)
        End If
    Loop While ilRet = 0
    mGetSeparateFilename = CleanFilename(mGetSeparateFilename)
End Function

'TTP 10271: Insertion Orders report: add option to create a separate PDF file for each vehicle
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableSeparateFiles            *
'*                                                     *
'*             Created:09/20/21      By:J.White        *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if CkcSeparateFile   *
'*                      can be enabled                 *
'*                                                     *
'*******************************************************
Sub mEnableSeparateFiles()
    If Val(edcTopHowMany.Text) > 0 Then
        'Single Contract entered in
        If gGetListSelectedCount(lbcSelection(6)) > 1 Then
            'Mutiple Vehicles Selected
            ckcSeparateFile.Enabled = True
        Else
            ckcSeparateFile.Enabled = False
            If igGenRpt = False Then ckcSeparateFile.Value = 0
        End If
    Else
        If gGetListSelectedCount(lbcSelection(0)) = 1 Then
            'A Single Contract Selected
            If gGetListSelectedCount(lbcSelection(6)) > 1 Then
                'Multiple Vehicles Selected
                ckcSeparateFile.Enabled = True
            Else
                ckcSeparateFile.Enabled = False
                If igGenRpt = False Then ckcSeparateFile.Value = 0
            End If
        Else
            ckcSeparateFile.Enabled = False
            If igGenRpt = False Then ckcSeparateFile.Value = 0
        End If
    End If
End Sub

'Fix TTP 10271 per Jason Email: Tue 9/28/21 10:22 AM (Issue #2,#3)
Function CleanFilename(slFileName As String) As String
    slFileName = Replace(slFileName, "&", "-")
    slFileName = Replace(slFileName, "*", "-")
    slFileName = Replace(slFileName, ":", "-")
    slFileName = Replace(slFileName, "?", "-")
    slFileName = Replace(slFileName, "%", "-")
    slFileName = Replace(slFileName, "=", "-")
    slFileName = Replace(slFileName, "<", "-")
    slFileName = Replace(slFileName, ">", "-")
    slFileName = Replace(slFileName, ";", "-")
    slFileName = Replace(slFileName, "@", "-")
    slFileName = Replace(slFileName, "[", "_")
    slFileName = Replace(slFileName, "]", " ")
    slFileName = Replace(slFileName, "{", "_")
    slFileName = Replace(slFileName, "}", " ")
    slFileName = Replace(slFileName, "(", "_")
    slFileName = Replace(slFileName, ")", " ")
    slFileName = Replace(slFileName, "^", "-")
    slFileName = Replace(slFileName, ",", "-")
    slFileName = Replace(slFileName, " ", "_")
    CleanFilename = slFileName
End Function
