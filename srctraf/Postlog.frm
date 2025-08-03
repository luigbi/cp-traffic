VERSION 5.00
Begin VB.Form PostLog 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6000
   ClientLeft      =   240
   ClientTop       =   1845
   ClientWidth     =   9450
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   9450
   Begin VB.CommandButton cmcAdServer 
      Caption         =   "Ad Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   85
      Top             =   5580
      Width           =   1470
   End
   Begin VB.CommandButton cmcManual 
      Appearance      =   0  'Flat
      Caption         =   "&Manual Posting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6795
      TabIndex        =   84
      Top             =   5580
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "&Invoice Import"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5250
      TabIndex        =   49
      Top             =   5580
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox pbcKey 
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
      Height          =   4710
      Left            =   810
      Picture         =   "Postlog.frx":0000
      ScaleHeight     =   4680
      ScaleWidth      =   7500
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   5850
      Visible         =   0   'False
      Width           =   7530
   End
   Begin VB.ListBox lbcGameNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "Postlog.frx":72462
      Left            =   6555
      List            =   "Postlog.frx":72464
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.PictureBox plcCalendar 
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
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   5565
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   30
         Picture         =   "Postlog.frx":72466
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   62
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1620
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   63
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcComplete 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   5805
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   555
      Width           =   5805
      Begin VB.CheckBox ckcDayComplete 
         Caption         =   "Su"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   5280
         TabIndex        =   79
         Top             =   0
         Width           =   675
      End
      Begin VB.CheckBox ckcDayComplete 
         Caption         =   "Sa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   78
         Top             =   0
         Width           =   675
      End
      Begin VB.CheckBox ckcDayComplete 
         Caption         =   "Fr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4080
         TabIndex        =   77
         Top             =   0
         Width           =   675
      End
      Begin VB.CheckBox ckcDayComplete 
         Caption         =   "Th"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   76
         Top             =   0
         Width           =   675
      End
      Begin VB.CheckBox ckcDayComplete 
         Caption         =   "We"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   75
         Top             =   0
         Width           =   675
      End
      Begin VB.CheckBox ckcDayComplete 
         Caption         =   "Tu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   74
         Top             =   0
         Width           =   675
      End
      Begin VB.CheckBox ckcDayComplete 
         Caption         =   "Mo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   73
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      Picture         =   "Postlog.frx":75280
      ScaleHeight     =   225
      ScaleWidth      =   75
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2115
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6645
      Top             =   5730
   End
   Begin VB.PictureBox plcInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
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
      Height          =   315
      Left            =   615
      ScaleHeight     =   285
      ScaleWidth      =   8055
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   8085
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Contract #: xxxxxx  Check #  Transaction: Date xx/xx/xx  Type xx  Action  xx"
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
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   330
         Width           =   7845
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Invoice #: xxxxxx  Vehicle Name:"
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
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   65
         Top             =   45
         Width           =   7890
      End
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   780
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1965
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "Postlog.frx":7558A
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Postlog.frx":76248
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
      End
   End
   Begin VB.TextBox edcDTDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   4530
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1065
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcDTDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5775
      Picture         =   "Postlog.frx":76552
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcCopyNm 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "Postlog.frx":7664C
      Left            =   855
      List            =   "Postlog.frx":7664E
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2925
      Visible         =   0   'False
      Width           =   4425
   End
   Begin VB.PictureBox plcAvailTimes 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3225
      ScaleHeight     =   150
      ScaleWidth      =   705
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   945
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox pbcPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6960
      ScaleHeight     =   210
      ScaleWidth      =   840
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.ListBox lbcAvailTimes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      ItemData        =   "Postlog.frx":76650
      Left            =   3510
      List            =   "Postlog.frx":76652
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2070
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmcTZDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5475
      Picture         =   "Postlog.frx":76654
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcTZDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   4230
      MaxLength       =   20
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcDT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   720
      Left            =   4425
      Picture         =   "Postlog.frx":7674E
      ScaleHeight     =   720
      ScaleWidth      =   1230
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   750
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox plcDT 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   4260
      ScaleHeight     =   1065
      ScaleWidth      =   1470
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1470
      Begin VB.PictureBox pbcDTTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   105
         Left            =   30
         ScaleHeight     =   105
         ScaleWidth      =   75
         TabIndex        =   21
         Top             =   840
         Width           =   75
      End
      Begin VB.PictureBox pbcDTSTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   90
         Left            =   45
         ScaleHeight     =   90
         ScaleWidth      =   15
         TabIndex        =   16
         Top             =   225
         Width           =   15
      End
   End
   Begin VB.Timer tmcClick 
      Interval        =   2000
      Left            =   8370
      Top             =   5475
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
      Left            =   7320
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5715
      Visible         =   0   'False
      Width           =   525
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
      Left            =   7275
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceHelpMsg 
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
      Left            =   7650
      TabIndex        =   53
      Top             =   5610
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2940
      TabIndex        =   47
      Top             =   5580
      Width           =   945
   End
   Begin VB.PictureBox pbcTZZone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2790
      ScaleHeight     =   210
      ScaleWidth      =   765
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1905
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox pbcTZCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   1785
      Left            =   2775
      Picture         =   "Postlog.frx":7AE70
      ScaleHeight     =   1785
      ScaleWidth      =   5250
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox plcTZCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      ForeColor       =   &H00000000&
      Height          =   2130
      Left            =   2640
      ScaleHeight     =   2130
      ScaleWidth      =   5460
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   5460
      Begin VB.PictureBox pbcTZSTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   90
         Left            =   60
         ScaleHeight     =   90
         ScaleWidth      =   45
         TabIndex        =   23
         Top             =   225
         Width           =   45
      End
      Begin VB.PictureBox pbcTZTab 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Height          =   45
         Left            =   120
         ScaleHeight     =   45
         ScaleWidth      =   30
         TabIndex        =   28
         Top             =   1770
         Width           =   30
      End
   End
   Begin VB.PictureBox pbcMissed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   1395
      Left            =   1725
      Picture         =   "Postlog.frx":85402
      ScaleHeight     =   1395
      ScaleWidth      =   7395
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4020
      Width           =   7395
      Begin VB.Label lacMdFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   44
         Top             =   300
         Visible         =   0   'False
         Width           =   7395
      End
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8130
      Top             =   5520
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
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   15
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5685
      Width           =   15
   End
   Begin VB.PictureBox pbcTab 
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
      Height          =   15
      Left            =   15
      ScaleHeight     =   15
      ScaleWidth      =   30
      TabIndex        =   31
      Top             =   3610
      Width           =   30
   End
   Begin VB.PictureBox pbcSTab 
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
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   90
      TabIndex        =   5
      Top             =   255
      Width           =   90
   End
   Begin VB.CommandButton cmcDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4665
      Picture         =   "Postlog.frx":A752C
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   3645
      MaxLength       =   20
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcPosting 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   2730
      Left            =   180
      Picture         =   "Postlog.frx":A7626
      ScaleHeight     =   2730
      ScaleWidth      =   8940
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   855
      Width           =   8940
      Begin VB.Label lacPtFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   30
         Top             =   405
         Visible         =   0   'False
         Width           =   8910
      End
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4095
      TabIndex        =   48
      Top             =   5580
      Width           =   945
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleWidth      =   780
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   780
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1755
      TabIndex        =   46
      Top             =   5580
      Width           =   945
   End
   Begin VB.PictureBox plcSpots 
      ForeColor       =   &H00000000&
      Height          =   1785
      Left            =   1680
      ScaleHeight     =   1725
      ScaleWidth      =   7665
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3675
      Width           =   7725
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Bill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   5115
         TabIndex        =   40
         Top             =   60
         Value           =   1  'Checked
         Width           =   585
      End
      Begin VB.TextBox edcMdDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Height          =   210
         Index           =   1
         Left            =   1890
         MaxLength       =   20
         TabIndex        =   37
         Top             =   60
         Width           =   825
      End
      Begin VB.CommandButton cmcMdDate 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2730
         Picture         =   "Postlog.frx":C2148
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   60
         Width           =   195
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Hidden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   7455
         TabIndex        =   43
         Top             =   60
         Width           =   945
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   6555
         TabIndex        =   42
         Top             =   60
         Width           =   915
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Orders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   5700
         TabIndex        =   41
         Top             =   60
         Width           =   930
      End
      Begin VB.ComboBox cbcVehicle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3060
         TabIndex        =   39
         Top             =   30
         Width           =   2025
      End
      Begin VB.VScrollBar vbcMissed 
         Height          =   1395
         LargeChange     =   5
         Left            =   7410
         TabIndex        =   45
         Top             =   315
         Width           =   240
      End
      Begin VB.TextBox edcMdDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Height          =   210
         Index           =   0
         Left            =   525
         MaxLength       =   20
         TabIndex        =   35
         Top             =   60
         Width           =   825
      End
      Begin VB.CommandButton cmcMdDate 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1365
         Picture         =   "Postlog.frx":C2242
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lacMissed 
         Appearance      =   0  'Flat
         Caption         =   "From                   To    "
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
         Height          =   225
         Left            =   75
         TabIndex        =   34
         Top             =   75
         Width           =   3630
      End
   End
   Begin VB.PictureBox plcPosting 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2850
      Left            =   105
      ScaleHeight     =   2790
      ScaleWidth      =   9210
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   780
      Width           =   9270
      Begin VB.VScrollBar vbcPosting 
         Height          =   2715
         LargeChange     =   11
         Left            =   8970
         TabIndex        =   32
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox plcSelect 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   840
      ScaleHeight     =   360
      ScaleWidth      =   8475
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   8535
      Begin VB.PictureBox pbcWM 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   6855
         ScaleHeight     =   210
         ScaleWidth      =   300
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   45
         Width           =   300
      End
      Begin VB.ComboBox cbcAvailName 
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
         Left            =   5235
         TabIndex        =   3
         Top             =   30
         Width           =   1560
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Vehicles"
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
         Index           =   0
         Left            =   15
         TabIndex        =   68
         Top             =   75
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Packages"
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
         Index           =   1
         Left            =   1050
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   75
         Width           =   1185
      End
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Height          =   210
         Left            =   7155
         TabIndex        =   57
         Top             =   45
         Width           =   1110
      End
      Begin VB.CommandButton cmcDate 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   8280
         Picture         =   "Postlog.frx":C233C
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   45
         Width           =   195
      End
      Begin VB.ComboBox cbcVeh 
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
         Left            =   2235
         TabIndex        =   2
         Top             =   30
         Width           =   2925
      End
   End
   Begin VB.PictureBox plcSort 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   150
      ScaleHeight     =   195
      ScaleWidth      =   3525
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   555
      Width           =   3525
      Begin VB.OptionButton rbcSort 
         Caption         =   "Advertiser"
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
         Height          =   225
         Index           =   1
         Left            =   1875
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Date/Time"
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
         Height          =   225
         Index           =   0
         Left            =   645
         TabIndex        =   70
         Top             =   0
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.PictureBox plcReason 
      Height          =   1770
      Left            =   45
      ScaleHeight     =   1710
      ScaleWidth      =   1515
      TabIndex        =   81
      Top             =   3705
      Width           =   1575
      Begin VB.PictureBox pbcMissedType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   75
         ScaleHeight     =   210
         ScaleWidth      =   1380
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   75
         Width           =   1380
      End
      Begin VB.ListBox lbcMissed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         ItemData        =   "Postlog.frx":C2436
         Left            =   0
         List            =   "Postlog.frx":C2438
         Sorted          =   -1  'True
         TabIndex        =   82
         Top             =   420
         Width           =   1500
      End
   End
   Begin VB.Image imcHidden 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   930
      Picture         =   "Postlog.frx":C243A
      Top             =   5475
      Width           =   465
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   75
      Picture         =   "Postlog.frx":C28CC
      Top             =   255
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8895
      Picture         =   "Postlog.frx":C2BD6
      Top             =   5505
      Width           =   480
   End
End
Attribute VB_Name = "PostLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Postlog.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  lmFeedSpotColor                                                                       *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PostLog.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Post Log input screen code
Option Explicit
Option Compare Text
'Constants
Const vbQuestion = 32
' Drag Status
Const ENTER = 0
Const LEAVE = 1
Const imTimerInterval = 600
Const SCROLLUP = 1
Const SCROLLDN = 2
Dim imButtonIndex As Integer
Dim imFirstActivate As Integer
Dim imWM As Integer '0=Weekly; 1=Monthly view
Dim imScrollDirection As Integer
Dim imCopyNmListIndex As Integer
Dim imSelectDelay As Integer    'True=cbcSelect change mode
Dim imStartMode As Integer
Dim imProcClickMode As Integer  'True = processing selection- disallow any other change until completed
Dim imAvailSelectedIndex As Integer
Dim imTZCopyAllowed As Integer ' True=Veh can have time zone copy, FALSE= no copy allowed
Dim imTZTabDirection As Integer
Dim imTZDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imLastRowSelected As Integer    'Used with multiple selection of rows
'Dim imWeek As Integer   '0=Current Week; 1= Current Month; 2=Current plus Past months
Dim imVehicle As Integer    '0=Current selected Vehicle; 1= All Vehicles
'Dim imSvWeek As Integer   '0=Current Week; 1= Current Month; 2=Current plus Past months
Dim imSvVehicle As Integer    '0=Current selected Vehicle; 1= All Vehicles
Dim imSvVehSelectedIndex As Integer
Dim imSvDateSelectedIndex As Integer
Dim lmMdStartdate As Long
Dim lmMdEndDate As Long
Dim imDoNotUpdate As Integer ' True=do not update, False=ok to update
Dim tmCopyNmCode() As SORTCODE
Dim smCopyNmCodeTag As String
Dim tmMissedCode() As SORTCODE
Dim smMissedCodeTag As String
Dim smCopy As String
'Affidavit Field Areas.  These hold the actual, full detail for each row
Dim tmCtrls(0 To 13)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer  'Current affidavit Box
'  smShow will hold all of the data values from the files
'Dim smShow() As String  'Values shown in affidavit Box (1 to 12 , 1 to #ofRows)
'  smSave will hold only those values the user can change
'Dim smSave() As String  'Values retained for affidavit Box (1 to 7, 1 to #ofRows)
                        '1=Air Date; 2=Air Time; 3=Media Code/Inventory Name & ISCI & Product; 4=ISCI;
                        ' 5=Price, 6=Advt Name, 7=SDF Record #, 8= Product Name (from contract)
                        '9=Sch Time; 10=Date
'Dim imSave() As Integer '1=Price (0=Charge; 1=N/C; -1=Can't Alter); 2=Price (for chg test)
'Dim imPostSpotInfo() As Integer    '1:True=ISCI Required; False=ISCI not required
                                '2:True=ISCI missing; False=ISCI defined
                                '3:True = billed; False=Not billed
                                '4:True=SimulCast; False=Not SimulCast
Dim imRowNo As Integer    'Current spot row number
Dim imSdfChg As Integer  'True=spot detail changed; False=No spot detail changed
Dim imSdfAnyChg(0 To 6) As Integer  'True=any spot detail changed; False=No spot detail changed
'Date/Time
Dim tmDTCtrls(0 To 3) As FIELDAREA
Dim imLBDTCtrls As Integer
Dim imDTBoxNo As Integer
Dim smDTSave(0 To 3) As String  'Index zero ignored
Dim smSpotType As String
'Time zone copy
Dim tmTZCtrls(0 To 2)  As FIELDAREA
Dim imLBTZCtrls As Integer
Dim imTZBoxNo As Integer  'Current affidavit Box
Dim imTZRowNo As Integer    'Current row number
Dim smTZSave(0 To 2, 0 To 8) As String  'Index 1 = Copy ;Index zero ignored
Dim smTZShow(0 To 2, 0 To 8) As String  'Index zero ignored
Dim imTZSave As Integer     'Index 1: 0=All or 1=First zone
Dim smZones(0 To 8) As String   'Index zero ignored
Dim imNoZones As Integer
'Missed field area
Dim tmMdCtrls(0 To 8) As FIELDAREA
Dim imLBMdCtrls As Integer
Dim imMdRowNo As Integer    'Current missed row number
Dim imSaveIndex As Integer
Dim imSdfIndex As Integer
'Dim smMdSave() As String    '1=Missed reason
'Dim smMdShow() As String    'Values shown in affidavit Box (1 to 8 , 1 to #ofRows)
'Dim smMdSchStatus() As String
                            'the value indicates type of spot (M or C or H)
'Dim lmMdRecPos() As Long    'Record positions
'Spot detail record information
Dim hmSdf As Integer        'Spot detail file handle
Dim hmSvSdf As Integer
Dim hmPsf As Integer
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey As SDFKEY1 'SDF key record image
Dim tmSdfSrchKey3 As LONGKEY0
Dim imSdfRecLen As Integer     'SDF record length
'Dim lmTBStartTime(1 To 49) As Long  'Allowed times if time buy (-1 indicates end of times) and library times if library buy
'Dim lmTBEndTime(1 To 49) As Long
Dim lmTBStartTime(0 To 48) As Long  'Allowed times if time buy (-1 indicates end of times) and library times if library buy
Dim lmTBEndTime(0 To 48) As Long
'Advertiser
Dim hmAdf As Integer        'Advertiser file handle
Dim tmAdf As ADF            'ADF record image
Dim tmAdfSrchKey As INTKEY0 'ADF key record image
Dim imAdfRecLen As Integer     'ADF record length
'Agency
Dim hmAgf As Integer        'Agency file handle
Dim tmAgf As AGF            'AGF record image
Dim tmAgfSrchKey As INTKEY0 'AGF key record image
Dim imAgfRecLen As Integer     'AGF record length
'Contract header
Dim hmCHF As Integer        'Contract header file handle
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer     'CHF record length
'Contract line
Dim hmClf As Integer        'Contract line file handle
Dim tmClf As CLF            'CLF record image
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer     'CLF record length
'Contract Games
Dim hmCgf As Integer
Dim tmCgf As CGF
Dim imCgfRecLen As Integer
Dim tmCgfSrchKey1 As CGFKEY1    'CntrNo; CntRevNo; PropVer
Dim tmCgfCff() As CFF
'Copy inventory
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length
'Missed reason: Multi-name file
Dim hmMnf As Integer        'Multi-name file handle
Dim tmMnf As MNF            'MNF record image
Dim imMnfRecLen As Integer     'MNF record length
'Feed
Dim hmFsf As Integer
Dim tmFsf As FSF            'FSF record image
Dim tmFSFSrchKey As LONGKEY0 'FSF key record image
Dim imFsfRecLen As Integer     'FSF record length
'Feed Name
Dim hmFnf As Integer

'Product
Dim hmPrf As Integer

Dim hmSxf As Integer

Dim hmGhf As Integer

Dim hmGsf As Integer


Dim tmGsfInfo() As GSFINFO
Dim tmTeam() As MNF
Dim smTeamTag As String

Dim imSelectedGameNo As Integer
Dim imGameNoComboBoxIndex As Integer
Dim imGameNoChgMode As Integer

'Log Calendar File
Dim hmLcf As Integer        'Log Calendar file handle
Dim tmLcf As LCF            'LCF record image
Dim tmLcfSrchKey As LCFKEY0 'LCF key record image
Dim tmLcfSrchKey1 As LCFKEY1 'LCF key record image
Dim tmLcfSrchKey2 As LCFKEY2 'LCF key record image
Dim imLcfRecLen As Integer     'LCF record length
Dim lmLcfRecPos(0 To 6) As Long     'LCF record position
'Spot summary file
Dim hmSsf As Integer        'Spot summary file handle
'Dim tmSsf As SSF            'SSF record image
Dim lmSsfDate(0 To 6) As Long    'Dates of the days stored into tmSsf
Dim lmSsfRecPos(0 To 6) As Long  'Record positions
Dim tmSsf(0 To 6) As SSF         'Spot summary for one week (0 index for monday;
                                    '1 for tuesday;...; 6 for sunday)
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim imSsfRecLen As Integer     'SSF record length
'Dim lmSsfRecPos As Long     'SSF record position
'Dim lmSsfMemDate As Long
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
'  Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
Dim tmOrigVef As VEF
'  Vehicle File
Dim hmVsf As Integer        'Vehicle file handle
Dim tmVsf As VSF            'VEF record image
Dim imVsfRecLen As Integer     'VEF record length
'Vehicle Links
Dim hmVLF As Integer            'Vehicle links file handle
Dim tmVlf0() As VLF             'Mon-Fri vehicle links
Dim tmVlf6() As VLF             'Sat vehicle links
Dim tmVlf7() As VLF             'Sun vehicle links
'Dim tmVlf() As VLF
'Copy Rotation
Dim hmCrf As Integer
' Copy Combo Inventory File
Dim hmCcf As Integer        'Copy Combo Inventory file handle
Dim tmCcf As CCF            'CCF record image
Dim imCcfRecLen As Integer     'CCF record length
'  Copy Product/Agency File
Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length
' Time Zone Copy FIle
Dim hmTzf As Integer        'Time Zone Copy file handle
Dim tmTzf As TZF            'TZF record image
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer     'TZF record length
' Media Codes File
Dim hmMcf As Integer        'Media Codes file handle
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer     'MCF record length
' Rate Card Programs/Times File
Dim hmRdf As Integer        'Rate Card Programs/Times file handle
Dim tmLnRdf As RDF            'RDF record image
Dim tmRdfSrchKey As INTKEY0 'RDF key record image
Dim imRdfRecLen As Integer     'RDF record length
' Contract Flight File
Dim hmCff As Integer        'Contract Flight file handle
Dim tmCff As CFF            'CFF record image
Dim tmCffSrchKey As CFFKEY0 'CFF key record image
Dim imCffRecLen As Integer     'CFF record length
Dim tmFCff() As CFF
'Smf
Dim hmSmf As Integer        'Spot makegood file handle
Dim tmSmf As SMF            'SMF record image
Dim tmSmfSrchKey2 As LONGKEY0
Dim imSmfRecLen As Integer     'SMF record length


Dim hmIihf As Integer
Dim tmIihf As IIHF        'CFF record image
Dim tmIihfSrchKey0 As LONGKEY0    'CFF key record image
Dim tmIihfSrchKey1 As IIHFKEY1    'CFF key record image
Dim tmIihfSrchKey2 As IIHFKEY2    'CFF key record image
Dim tmIihfSrchKey3 As IIHFKEY3    'CFF key record image
Dim imIihfRecLen As Integer        'CFF record length
Dim smPostLogSource As String

'Record Locks
Dim lmLock1RecCode As Long
Dim lmLock2RecCode As Long
Dim hmRlf As Integer
'Dim tmRec As LPOPREC
Dim tmChfAdvtExt() As CHFADVTEXT
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imDateBox As Integer   '0=Post Log Date; 1=Missed start date; 2=missed end date
Dim lmAvailDate As Long     'Date for Avail Time
Dim imAvailAnfCode As Integer   'Avail Anf Code
Dim imAvailGameNo As Integer
Dim imFirstTime As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imDateChgMode As Integer
Dim imListChgMode As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imVefCode As Integer
Dim imVpfIndex As Integer   'Vehicle option index
Dim imUrfIndex As Integer
Dim imDateSelectedIndex As Integer  'Index of selected record
Dim smSelectedDate As String
Dim imSelectedDay As Integer
Dim imVehSelectedIndex As Integer  'Index of selected record
Dim imAdvtSelectedIndex As Integer  'Index of selected record
Dim imComboBoxIndex As Integer
Dim imDateComboBoxIndex As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imSettingValue As Integer   'True=Don't enable any box woth change
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer 'True = Processing arrow key- retain current list box visibly
                                'False= Make list box invisible
Dim imMissedType As Integer   '0=Missed; 1=Cancel; 2=Hidden
Dim imMissedListIndex As Integer    'Missed reason index- for drop
Dim imDropRowNo As Integer      'Row that missed spot was drop onto
Dim imIndex As Integer          'General index
Dim imUpdateAllowed As Integer
Dim imSvUpdateAllowed As Integer
Dim imButton As Integer 'Value 1= Left button; 2=Right button; 4=Middle button
Dim imButtonRow As Integer
Dim imIgnoreRightMove As Integer
Dim imIgnoreChg As Integer
Dim lmBonusDate As Long
Dim imInTab As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imDragSource As Integer '0=Post log; 1= Missed; 2= Bonus
Dim smDragCntrType As String    'X=Fill; B=Billboard
Dim lmVehLLD As Long        'Vehicle last log date
Dim imDefaultDateIndex As Integer ' Default Date Index
Dim imGetAll As Integer ' True = all records read, False = first screen full read
Dim imBypassFocus As Integer
Dim imBkQH As Integer   'Rank
Dim imPriceLevel As Integer
Dim lmSepLength As Long 'Separation length for advertiser
Dim lmStartDateLen As Long  'Start date that separartion is valid for
Dim lmEndDateLen As Long    'End date that separation is valid for
Dim smSvDate As String
Dim smSvAirTime As String
'6/16/11
Dim smSvAvailTime As String
''Required to be compatible with general schedule routines
''The array are not used by spots except for compatiblity
'Dim imHour(1 To 24) As Integer     'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
'Dim imDay(1 To 7) As Integer       'Day count (Index 1 for monday; Index 2 for tuesday;..)
'Dim imQH(1 To 4) As Integer        'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
''Actual for the day or week be processed- this will be a subset from
''imC---- or imP----
'Dim imAHour(1 To 24) As Integer    'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
'Dim imADay(1 To 7) As Integer      'Day count (Index 1 for monday; Index 2 for tuesday;..)
'Dim imAQH(1 To 4) As Integer       'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
'Dim imSkip(1 To 24, 1 To 4, 0 To 6) As Integer  '-1=Skip all test;0=All test;
'                                    'Bit 0=Skip insert;
'                                    'Bit 1=Skip move;
'                                    'Bit 2=Skip competitive pack;
'                                    'Bit 3=Skip Preempt
Dim imHour(0 To 23) As Integer     'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
Dim imDay(0 To 6) As Integer       'Day count (Index 1 for monday; Index 2 for tuesday;..)
Dim imQH(0 To 3) As Integer        'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
'Actual for the day or week be processed- this will be a subset from
'imC---- or imP----
Dim imAHour(0 To 23) As Integer    'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
Dim imADay(0 To 6) As Integer      'Day count (Index 1 for monday; Index 2 for tuesday;..)
Dim imAQH(0 To 3) As Integer       'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
Dim imSkip(0 To 23, 0 To 3, 0 To 6) As Integer  '-1=Skip all test;0=All test;
                                    'Bit 0=Skip insert;
                                    'Bit 1=Skip move;
                                    'Bit 2=Skip competitive pack;
                                    'Bit 3=Skip Preempt

Dim tmUserVehicle() As SORTCODE
Dim smUserVehicleTag As String
Dim tmAvailCode() As SORTCODE
Dim smAvailCodeTag As String

Private imSvCkcInclude(0 To 3) As Integer
Private bmInPackage As Boolean
Private imSvMissedType As Integer

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Const LBONE = 1

Const DATEINDEX = 1         'Date control/field
Const TIMEINDEX = 2         'Time control/field
Const LENINDEX = 3          'Length control/field
Const ADVTINDEX = 4         'Advertiser/Product control/field
Const TZONEINDEX = 5        'Time zone control/field
Const COPYINDEX = 6         'Copy control/field
Const ISCIINDEX = 7         'ISCI control/field
Const CNTRINDEX = 8         'Contract control/field
Const LINEINDEX = 9         'Line control/field
Const TYPEINDEX = 10         'Type control/field
Const PRICEINDEX = 11       'Price control/field
Const MGOODINDEX = 12       'Make Good control/field
Const AUDINDEX = 13         'Audit trail control/field
'Values set to number larger then AUDINDEX
Const SCHTOMISSED = 20      'Used to indicate that a posted spot should be set to missed
Const MISSEDTOSCH = 21      'Used to indicate that a missed spot should be set to posted
Const MISSEDREASON = 22     'Used to indicate that the missed reason should be changed
Const SCHTOCANCEL = 23      'Used to indicate that a posted spot should be set to cancelled
Const MISSEDTOCANCEL = 24      'Used to indicate that a missed spot should be set to cancelled
Const SCHTOHIDE = 25      'Used to indicate that a posted spot should be set to hide
Const MISSEDTOHIDE = 26      'Used to indicate that a missed spot should be set to hide
Const BONUSTOSCH = 27       'Used to indicate that a bonus spot should be created
Const MDADVTINDEX = 1
Const MDCNTRINDEX = 2
Const MDVEHINDEX = 3
Const MDLENINDEX = 4
Const MDWKMISSINDEX = 5
Const MDENDDATEINDEX = 6
Const MDDPINDEX = 7
Const MDNOSPOTSINDEX = 8
Const MDPRODINDEX = 9
' smSave indicies
' Time Zone indicies
Const ZONEINDEX = 1
Const TZCOPYINDEX = 2
Const DTDATEINDEX = 1
Private imDTAIRTIMEINDEX
'6/16/11
Private imDTAVAILTIMEINDEX

Private Sub cbcAvailName_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcAvailName.Text <> "" Then
            gManLookAhead cbcAvailName, imBSMode, imComboBoxIndex
        End If
        If cbcAvailName.ListIndex >= 0 Then
            If imDateSelectedIndex >= 0 Then
                pbcPosting.Cls
                mPaintPostTitle
                ReDim tgShow(0 To 1) As SHOWINFO
                tmcClick.Enabled = False
                'imSelectDelay = True
                tmcClick.Interval = 2000    '2 seconds
                tmcClick.Enabled = True
            End If
        End If
        imAvailSelectedIndex = cbcAvailName.ListIndex
        imChgMode = False
    End If
End Sub
Private Sub cbcAvailName_Click()
    imAvailSelectedIndex = cbcAvailName.ListIndex
    cbcAvailName_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcAvailName_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub cbcAvailName_DropDown()
    tmcClick.Enabled = False
End Sub
Private Sub cbcAvailName_GotFocus()
    If imTerminate Then
        Exit Sub
    End If
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        'Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    If imAvailSelectedIndex = -1 Then
        cbcAvailName.ListIndex = 0
        imAvailSelectedIndex = 0
    End If
    imComboBoxIndex = imAvailSelectedIndex
    gCtrlGotFocus cbcAvailName
    'tmcClick.Enabled = False
    Exit Sub
End Sub
Private Sub cbcAvailName_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcAvailName_KeyPress(KeyAscii As Integer)
    tmcClick.Enabled = False
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcAvailName.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcVeh_Change()
    Dim slStr As String
    If imStartMode Then
        imStartMode = False
        mCbcVehChange
        Exit Sub
    End If
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        cbcAvailName.Enabled = False
        edcDate.Enabled = False
        cmcDate.Enabled = False
        imDateSelectedIndex = -1
        slStr = Trim$(cbcVeh.Text)
        If slStr <> "" Then
            gManLookAhead cbcVeh, imBSMode, imComboBoxIndex
            If cbcVeh.ListIndex >= 0 Then
                pbcPosting.Cls
                mPaintPostTitle
                pbcMissed.Cls
                ReDim tgShow(0 To 1) As SHOWINFO
                ReDim tgSave(0 To 1) As SAVEINFO
                ReDim tgMdSdfRec(0 To 1) As MDSDFREC
                ReDim tgMdSaveInfo(0 To 1) As MDSAVEINFO
                ReDim tgMdShowInfo(0 To 1) As MDSHOWINFO
                tmcClick.Enabled = False
                imSelectDelay = True
                tmcClick.Interval = 2000    '2 seconds
                tmcClick.Enabled = True
                Exit Sub
            End If
        End If
        cbcAvailName.Enabled = True
        edcDate.Enabled = True
        cmcDate.Enabled = True
    End If
    Exit Sub
End Sub
Private Sub cbcVeh_Click()
    imComboBoxIndex = cbcVeh.ListIndex
    cbcVeh_Change
End Sub
Private Sub cbcVeh_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub cbcVeh_DropDown()
    tmcClick.Enabled = False
    imSelectDelay = False
End Sub
Private Sub cbcVeh_GotFocus()
    If imFirstTime Then
        imFirstTime = False
    End If
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    If cbcVeh.Text = "" Then
    ' get the default vehicle from this global var
        gFindMatch sgUserDefVehicleName, 0, cbcVeh
        If gLastFound(cbcVeh) >= 0 Then
            cbcVeh.ListIndex = gLastFound(cbcVeh)
        Else
            If cbcVeh.ListCount > 0 Then
                cbcVeh.ListIndex = 0
            End If
        End If
    Else
        cbcVeh.ListIndex = imVehSelectedIndex
    End If
    gCtrlGotFocus cbcVeh
    imComboBoxIndex = cbcVeh.ListIndex
    imVehSelectedIndex = imComboBoxIndex
    If cbcVeh.ListCount = 1 Then
        ''cbcDate.SetFocus
        If edcDate.Enabled Then
            edcDate.SetFocus
        End If
    End If
    tmcClick.Enabled = False
End Sub
Private Sub cbcVeh_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcVeh_KeyPress(KeyAscii As Integer)
    tmcClick.Enabled = False
    imSelectDelay = False
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcVeh.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcVeh_LostFocus()
    'If imSelectDelay Then
    '    tmcClick.Enabled = False
    '    imSelectDelay = False
    '    mCbcVehChange
    'End If
End Sub
Private Sub cbcVehicle_Change()
    Dim ilRet As Integer
    If imChgMode = False Then
        imChgMode = True
        Screen.MousePointer = vbHourglass
        If cbcVehicle.Text <> "" Then
            gManLookAhead cbcVehicle, imBSMode, imComboBoxIndex
        End If
        If cbcVehicle.ListIndex >= 0 Then
            imVehicle = cbcVehicle.ListIndex
            ilRet = mReadSdfRec(True)
        End If
        Screen.MousePointer = vbDefault    'Default
        imChgMode = False
    End If
End Sub
Private Sub cbcVehicle_Click()
    cbcVehicle_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcVehicle_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub cbcVehicle_GotFocus()
    If imTerminate Then
        Exit Sub
    End If
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    If imVehicle = -1 Then
        cbcVehicle.ListIndex = 0
        imVehicle = 0
    End If
    imComboBoxIndex = imVehicle
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cbcVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcVehicle_KeyPress(KeyAscii As Integer)
    tmcClick.Enabled = False
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcVehicle.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub ckcDayComplete_Click(Index As Integer)
Dim ilDayComplete As Integer
Dim ilSaveAnyChg As Integer

    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcDayComplete(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    If imDoNotUpdate = True Then
        Exit Sub
    End If

    'save original day flag to determine if anything changed, need to force
    'day complete flag off
    ilSaveAnyChg = imSdfAnyChg(Index)
    'turning day is complete flag off?
    If Not Value Then
        imSdfAnyChg(Index) = True
    End If

    ilDayComplete = mDetermineDayUpdate(Index)      'see if this is a valid day to set complete
    If ilDayComplete >= 0 Then
        'imDateSelectedIndex = ilDayComplete
        mUpdateAffPost Index               'send the day index to mark complete
    End If

    imSdfAnyChg(Index) = ilSaveAnyChg
    Screen.MousePointer = vbHourglass  'Wait
    pbcPosting_Paint
    Screen.MousePointer = vbDefault    'Default
    imDateChgMode = False
    'cbcDate.SetFocus
    edcDate.SetFocus
End Sub
Private Sub ckcDayComplete_DragDrop(Index As Integer, Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub ckcDayComplete_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub ckcDayComplete_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub ckcInclude_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcInclude(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilRet As Integer
    'imSvWeek = -1   'Force read
    lmMdStartdate = -1
    ilRet = mReadSdfRec(True)
End Sub
Private Sub ckcInclude_DragDrop(Index As Integer, Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub ckcInclude_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    If imDTBoxNo = DTDATEINDEX Then
        edcDTDropDown.SelStart = 0
        edcDTDropDown.SelLength = Len(edcDTDropDown.Text)
        edcDTDropDown.SetFocus
    Else
        If imDateBox = 0 Then
            edcDate.SelStart = 0
            edcDate.SelLength = Len(edcDate.Text)
            edcDate.SetFocus
        Else
            edcMdDate(imDateBox - 1).SelStart = 0
            edcMdDate(imDateBox - 1).SelLength = Len(edcMdDate(imDateBox - 1).Text)
            edcMdDate(imDateBox - 1).SetFocus
        End If
    End If
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    If imDTBoxNo = DTDATEINDEX Then
        edcDTDropDown.SelStart = 0
        edcDTDropDown.SelLength = Len(edcDTDropDown.Text)
        edcDTDropDown.SetFocus
    Else
        If imDateBox = 0 Then
            edcDate.SelStart = 0
            edcDate.SelLength = Len(edcDate.Text)
            edcDate.SetFocus
        Else
            edcMdDate(imDateBox - 1).SelStart = 0
            edcMdDate(imDateBox - 1).SelLength = Len(edcMdDate(imDateBox - 1).Text)
            edcMdDate(imDateBox - 1).SetFocus
        End If
    End If
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub cmcCancel_GotFocus()
    lbcCopyNm.Visible = False
    edcTZDropDown.Visible = False
    cmcTZDropDown.Visible = False
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    edcDTDropDown.Visible = False
    edcDTDropDown.Visible = False
    plcTme.Visible = False
    plcDT.Visible = False
    pbcDT.Visible = False
    imDTBoxNo = -1
    'mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    pbcArrow.Visible = False
    lacPtFrame.Visible = False
    lbcAvailTimes.Visible = False
    plcAvailTimes.Visible = False
    plcTZCopy.Visible = False
    pbcTZCopy.Visible = False
    lbcCopyNm.Visible = False
    edcDropDown.Visible = False
    cmcDropDown.Visible = False
    pbcPrice.Visible = False
    'If Not mSetShow(imBoxNo) Then
    '    Exit Sub
    'End If
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDate_Click()
    If tmVef.sType <> "G" Then
        plcCalendar.Visible = Not plcCalendar.Visible
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
        edcDate.SetFocus
    Else
        lbcGameNo.Visible = Not lbcGameNo.Visible
    End If
End Sub
Private Sub cmcDate_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub cmcDate_GotFocus()
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    gCtrlGotFocus ActiveControl
    If imDateBox <> 0 Then
        plcCalendar.Move plcSelect.Left + plcSelect.Width - fgBevelX - plcCalendar.Width, plcSelect.Top + edcDate.height + fgBevelY
        imDateBox = 0
    End If
End Sub
Private Sub cmcDone_Click()
    Dim ilLoop As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ' if day not marked complete, change it to incomplete in the file
    For ilLoop = 0 To 6 Step 1
        If imSdfAnyChg(ilLoop) = True Then
            mUpdateAffPost ilLoop
            Exit For
        End If
    Next ilLoop
    mTerminate
End Sub
Private Sub cmcDone_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case TIMEINDEX
            '5/20/11
            ''plcTme.Visible = Not plcTme.Visible
            'lbcAvailTimes.Visible = Not lbcAvailTimes.Visible
            plcTme.Visible = Not plcTme.Visible
        Case COPYINDEX
            lbcCopyNm.Visible = Not lbcCopyNm.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDTDropDown_Click()
    Select Case imDTBoxNo
        Case DTDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        '6/16/11
        Case imDTAIRTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case imDTAVAILTIMEINDEX
            'If (rbcType(0).Value) And (smSpotType <> "O") And (smSpotType <> "C") Then
                lbcAvailTimes.Visible = Not lbcAvailTimes.Visible
            'Else
            '    plcTme.Visible = Not plcTme.Visible
            'End If
    End Select
    edcDTDropDown.SelStart = 0
    edcDTDropDown.SelLength = Len(edcDTDropDown.Text)
    If imDTBoxNo <> imDTAVAILTIMEINDEX Then
        edcDTDropDown.SetFocus
    End If
End Sub
Private Sub cmcDTDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcImport_Click()
    igBrowserType = 8  'PDF
    sgBrowseMaskFile = ""
    sgBrowserTitle = "Import for Post Log Invoices"
    Browser.Show vbModal
    If igBrowserReturn = 1 Then
        ImportStationSpots.Show vbModal
    End If
End Sub

Private Sub cmcManual_Click()
    'PostManualVeh.Show vbModal
    'If igManualPostVefCode > 0 Then
    '    PostManualCntr.Show vbModal
    'End If
    gShellToPostManual PostLog, "FROM/POSTLOG"

End Sub

Private Sub cmcMdDate_Click(Index As Integer)
    plcCalendar.Visible = Not plcCalendar.Visible
    edcMdDate(Index).SelStart = 0
    edcMdDate(Index).SelLength = Len(edcMdDate(Index).Text)
    edcMdDate(Index).SetFocus
End Sub
Private Sub cmcMdDate_GotFocus(Index As Integer)
    tmcClick.Enabled = False
    If imTerminate Then
        Exit Sub
    End If
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    gCtrlGotFocus ActiveControl
    If imDateBox <> Index + 1 Then
        plcCalendar.Visible = False
        plcCalendar.Move plcSpots.Left + edcMdDate(Index).Left - fgBevelX, plcSpots.Top + edcMdDate(Index).height + fgBevelY
        imDateBox = Index + 1
    End If
End Sub

Private Sub cmcAdServer_Click()
 AdServerFilter.Show vbModal
End Sub

Private Sub cmcReport_Click()
    Dim slStr As String        'General string
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = POSTLOGSJOB
    igRptType = 0
    'Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "PostLog^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        Else
            slStr = "PostLog^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "PostLog^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "PostLog^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'PostLog.Enabled = False
    ''Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'PostLog.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    sgCommandStr = slStr
    RptList.Show vbModal
    'Screen.MousePointer = vbDefault
End Sub
Private Sub cmcReport_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub cmcReport_GotFocus()
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        'Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub cmcReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcTZDropDown_Click()
    Select Case imTZBoxNo
        Case TZCOPYINDEX
            lbcCopyNm.Visible = Not lbcCopyNm.Visible
    End Select
    edcTZDropDown.SelStart = 0
    edcTZDropDown.SelLength = Len(edcTZDropDown.Text)
    edcTZDropDown.SetFocus
End Sub
Private Sub edcDate_Change()
   Dim slStr As String
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim llWkSDate As Long
    Dim llWkEDate As Long
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilDayIndex As Integer
    Dim ilTempDayIndex As Integer
    Dim llTempDate As Long
    Dim slTempStr As String
    Dim llTempLoop As Long
    Dim ilFirstValidDayInx As Integer
    Dim ilSave As Integer
    Dim ilGsf As Integer

    If tmVef.sType <> "G" Then
        slStr = edcDate.Text
        If Not gValidDate(slStr) Then
            pbcPosting.Cls
            mPaintPostTitle
            lacDate.Visible = False
            Exit Sub
        End If
        If imDateChgMode = False Then
            ilDayIndex = gWeekDayStr(slStr)
            llDate = gDateValue(slStr)

            'loop on the days of the week and disable the days prior to the start day selected
            'i.e Sat selected, disable M-Fr
            For ilLoop = 0 To 6
                imDateSelectedIndex = -1            'set flag to ignore going to update of LCF complete flag
                ckcDayComplete(ilLoop).Value = vbUnchecked
                If ilLoop < ilDayIndex Then
                    ckcDayComplete(ilLoop).Enabled = False
                Else
                    ckcDayComplete(ilLoop).Enabled = True
                End If
            Next ilLoop

            ilTempDayIndex = ilDayIndex
            llTempDate = llDate
            'set the day as complete flag on or off for the selected week
            Do While ilTempDayIndex <> 0
                llTempDate = llTempDate - 1
                slTempStr = Format$(llTempDate, "m/d/yy")
                ilTempDayIndex = gWeekDayStr(slTempStr)
            Loop
            'llTempDate = start of the week selected
            'show how the days are set upon entering week
            ilFirstValidDayInx = -1
            For llTempLoop = llTempDate To llTempDate + 6
                ilFound = False
                ilTempDayIndex = gWeekDayStr(Format$(llTempLoop, "m/d/yy"))   'reestablish day of week to set
                For ilLoop = 0 To UBound(tgDates) - 1
                    If llTempLoop = tgDates(ilLoop).lDate Then
                        ilFound = True
                        If ilFirstValidDayInx < 0 Then
                            ilFirstValidDayInx = ilLoop     'first valid day of the week
                        End If
                        ckcDayComplete(ilTempDayIndex).Enabled = True
                        If tgDates(ilLoop).iStatus = 2 Then
                            'show associated day as complete
                            ckcDayComplete(ilTempDayIndex).Value = vbChecked
                            Exit For
                        Else
                            'show associated day as not complete
                            ckcDayComplete(ilTempDayIndex).Value = vbUnchecked
                            Exit For
                        End If
                    End If
                Next ilLoop
                If Not ilFound Then             'no valid date, disable the day of week
                    ckcDayComplete(ilTempDayIndex).Enabled = False
                End If
            Next llTempLoop

            ilFound = False
            imDateChgMode = True
            imDateSelectedIndex = -1

            For ilLoop = 0 To UBound(tgDates) - 1 Step 1
                If tgDates(ilLoop).lDate = llDate Then
                    imDateSelectedIndex = ilLoop
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If (Not ilFound) And (ilFirstValidDayInx >= 0) Then
                imDateSelectedIndex = ilFirstValidDayInx
                ilFound = True
            End If

            If Not ilFound Then
                imDateChgMode = False
                lacDate.Visible = False
                pbcPosting.Cls
                mPaintPostTitle
            '    ReDim smShow(1 To 13, 1 To 1) As String
                ReDim tgShow(0 To 1) As SHOWINFO
            '   smSave will hold only data that can be changed by the user.
            '          This data is complete (not trimmed to fit the controls).
                'ReDim smSave(1 To 10, 1 To 1) As String
                ReDim tgSave(0 To 1) As SAVEINFO
                'ReDim imSave(1 To 2, 1 To 1) As Integer
                'ReDim imPostSpotInfo(1 To 4, 1 To 1) As Integer
                'ReDim tmVlf(1 To 1) As VLF
                imSvDateSelectedIndex = -1
                pbcMissed.Cls
                ReDim tgMdSdfRec(0 To 1) As MDSDFREC
                ReDim tgMdSaveInfo(0 To 1) As MDSAVEINFO
                ReDim tgMdShowInfo(0 To 1) As MDSHOWINFO

                '8-25-04  show all spots in the week from the start date selected even if a valid date isnt selected
                imGetAll = False
                'save the original contents of imDaySelectedIndex
                ilSave = imDateSelectedIndex
                imDateSelectedIndex = ilFirstValidDayInx
                ilRet = mReadSdfRec(True)   '(False) changed 5/12/99 via Jim since date is selected by mouse
                '9/8/06-  Remove setting imDateSelectedIndex back to -1, use first valid date
                'imDateSelectedIndex = ilSave
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
            If imWM <> 1 Then
                llWkSDate = llDate
                Do While gWeekDayLong(llWkSDate) <> 0
                    llWkSDate = llWkSDate - 1
                Loop
                ReDim tgWkDates(0 To 0) As DATES
                llWkEDate = llWkSDate + 6
            Else
                llWkSDate = gDateValue(gObtainStartStd(Format$(llDate, "m/d/yy")))
                ReDim tgWkDates(0 To 0) As DATES
                llWkEDate = gDateValue(gObtainEndStd(Format$(llDate, "m/d/yy")))
            End If
            If rbcType(0).Value Then
                edcMdDate(0).Text = Format$(llWkSDate, "m/d/yy")
                edcMdDate(1).Text = Format$(llWkEDate, "m/d/yy")
            Else
                edcMdDate(0).Text = ""
                edcMdDate(1).Text = ""
            End If
            For ilLoop = 0 To UBound(tgDates) - 1 Step 1
                If (tgDates(ilLoop).lDate >= llWkSDate) And (tgDates(ilLoop).lDate <= llWkEDate) Then
                    tgWkDates(UBound(tgWkDates)).sDate = tgDates(ilLoop).sDate
                    tgWkDates(UBound(tgWkDates)).lDate = tgDates(ilLoop).lDate
                    tgWkDates(UBound(tgWkDates)).iStatus = tgDates(ilLoop).iStatus
                    ReDim Preserve tgWkDates(0 To UBound(tgWkDates) + 1) As DATES
                End If
            Next ilLoop

            pbcPosting.Cls
            mPaintPostTitle
            For ilLoop = 0 To 6 Step 1
                imSdfAnyChg(ilLoop) = False
            Next ilLoop
            'If (tmVef.sType = "S") Then
            '    If (tgVpf(imVpfIndex).sBillSA = "Y") Then
            '        gObtainVlf "S", hmVlf, imVefCode, llDate, tmVlf()
            '    End If
            'End If
            imSvDateSelectedIndex = -1
            imGetAll = False
            ilRet = mReadSdfRec(True)   '(False) changed 5/12/99 via Jim since date is selected by mouse
            Screen.MousePointer = vbHourglass
            ' set flag, iftheelse reset flag
            imDoNotUpdate = True ' avoid having cbcDate update the file

            'If rbcType(0).Value Then
            '    If tmLcf.sAffPost = "C" Then
            '        ckcDayComplete(ilDayIndex).Value = vbChecked    'True
            '    Else
            '        ckcDayComplete(ilDayIndex).Value = vbUnchecked  'False
            '    End If
            'End If

            'mAvailTimePop
            mFinalInvoiceRunning llWkSDate
            imDoNotUpdate = False ' now allow updating of the file
            pbcPosting_Paint
            'pbcMissed.Cls
            'ilRet = mReadMdSdfRec()
            plcCalendar.Visible = False
            Screen.MousePointer = vbDefault    'Default
            imDateChgMode = False

        End If
    Else
        imLbcArrowSetting = True
        gMatchLookAhead edcDate, lbcGameNo, imBSMode, imGameNoComboBoxIndex
        imSelectedGameNo = lbcGameNo.ListIndex
        ilGsf = lbcGameNo.ItemData(imSelectedGameNo)
        ReDim tgWkDates(0 To 0) As DATES
        llWkSDate = tmGsfInfo(ilGsf).lGameDate
        llWkEDate = llWkSDate
        edcMdDate(0).Text = Format$(llWkSDate, "m/d/yy")
        edcMdDate(1).Text = Format$(llWkEDate, "m/d/yy")
        tgWkDates(UBound(tgWkDates)).sDate = Format$(llWkSDate, "m/d/yy")
        tgWkDates(UBound(tgWkDates)).lDate = llWkSDate
        tgWkDates(UBound(tgWkDates)).iStatus = 0
        ReDim Preserve tgWkDates(0 To UBound(tgWkDates) + 1) As DATES
        imDateSelectedIndex = imSelectedGameNo
        'pbcPosting.Cls
        'mPaintPostTitle
        For ilLoop = 0 To 6 Step 1
            imSdfAnyChg(ilLoop) = False
        Next ilLoop
        ckcDayComplete(0).Caption = "Event " & Trim$(lbcGameNo.List(imSelectedGameNo))
        If tgDates(imDateSelectedIndex).iStatus = 2 Then
            ckcDayComplete(0).Value = vbChecked
        Else
            ckcDayComplete(0).Value = vbUnchecked
        End If
        '3/26/11: Moved after DayComplete because pbcPosting_Paint is called with ckcDayComplete
        pbcPosting.Cls
        mPaintPostTitle
        If ((tgSpf.sBActDayCompl <> "N") And (imWM = 0)) Or (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) = USINGLIVELOG) And ((tgVpf(imVpfIndex).sGenLog = "L") Or ((tgVpf(imVpfIndex).sGenLog = "A") And (tgDates(imDateSelectedIndex).sLiveLogMerge <> "M")))) Then
            plcComplete.Visible = True
            plcComplete.Enabled = True
        Else
            plcComplete.Visible = False
            plcComplete.Enabled = False
        End If
        'If (tmVef.sType = "S") Then
        '    If (tgVpf(imVpfIndex).sBillSA = "Y") Then
        '        gObtainVlf "S", hmVlf, imVefCode, llDate, tmVlf()
        '    End If
        'End If
        imSvDateSelectedIndex = -1
        imGetAll = False
        ilRet = mReadSdfRec(True)   '(False) changed 5/12/99 via Jim since date is selected by mouse
        Screen.MousePointer = vbHourglass
        ' set flag, iftheelse reset flag
        imDoNotUpdate = True ' avoid having cbcDate update the file

        'If rbcType(0).Value Then
        '    If tmLcf.sAffPost = "C" Then
        '        ckcDayComplete(ilDayIndex).Value = vbChecked    'True
        '    Else
        '        ckcDayComplete(ilDayIndex).Value = vbUnchecked  'False
        '    End If
        'End If

        'mAvailTimePop
        mFinalInvoiceRunning llWkSDate
        imDoNotUpdate = False ' now allow updating of the file
        pbcPosting_Paint
        'pbcMissed.Cls
        'ilRet = mReadMdSdfRec()
        lbcGameNo.Visible = False
        Screen.MousePointer = vbDefault    'Default
        imDateChgMode = False
        imLbcArrowSetting = False
    End If
End Sub
Private Sub edcDate_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub edcDate_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDayIndex                                                                            *
'******************************************************************************************

    Dim ilLoop As Integer
    If imFirstTime Then
        imFirstTime = False
    End If
    If tmVef.sType = "G" Then
        imGameNoComboBoxIndex = lbcGameNo.ListIndex
    End If
    If imDateBox <> 0 Then
        plcCalendar.Move plcSelect.Left + plcSelect.Width - fgBevelX - plcCalendar.Width, plcSelect.Top + edcDate.height + fgBevelY
        imDateBox = 0
    End If
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    'If cbcDate.Text = "" And imDefaultDateIndex >= 0 Then
    '    cbcDate.ListIndex = imDefaultDateIndex
    'End If
    If Not imBypassFocus Then
        gCtrlGotFocus edcDate
    End If
    imBypassFocus = False
    imDateComboBoxIndex = imDateSelectedIndex
    'imDateSelectedIndex = imDateComboBoxIndex
    ' if any change has occurred, mark this record as incomplete
    For ilLoop = 0 To 6 Step 1
        If imSdfAnyChg(ilLoop) = True Then
            If tmVef.sType <> "G" Then
                'ilDayIndex = gWeekDayStr(edcDate)
                'imDateSelectedIndex = ilDayIndex         '8-25-04
                mUpdateAffPost ilLoop
            Else
                mUpdateAffPost 0
                Exit For
            End If
        End If
    Next ilLoop
End Sub
Private Sub edcDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If tmVef.sType <> "G" Then
            If (Shift And vbAltMask) > 0 Then
                plcCalendar.Visible = Not plcCalendar.Visible
            Else
                slDate = edcDate.Text
                If gValidDate(slDate) Then
                    If KeyCode = KEYUP Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcDate.Text = slDate
                End If
            End If
        Else
            gProcessArrowKey Shift, KeyCode, lbcGameNo, imLbcArrowSetting
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If tmVef.sType <> "G" Then
            If (Shift And vbAltMask) > 0 Then
            Else
                slDate = edcDate.Text
                If gValidDate(slDate) Then
                    If KeyCode = KEYLEFT Then 'Up arrow
                        slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                    Else
                        slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                    End If
                    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    edcDate.Text = slDate
                End If
            End If
            edcDate.SelStart = 0
            edcDate.SelLength = Len(edcDate.Text)
        End If
    End If
End Sub
Private Sub edcDropDown_Change()
    Select Case imBoxNo
        Case TIMEINDEX
            'slStr = edcDropDown.Text
            'gFindMatch slStr, 0, lbcAvailTimes
            'imChgMode = True ' Turn on the switch
            'If gLastFound(lbcAvailTimes) >= 0 Then
            '    lbcAvailTimes.ListIndex = gLastFound(lbcAvailTimes)
            'Else ' No data found so re-display the last good data
            '    lbcAvailTimes.ListIndex = -1
            'End If
            'imChgMode = False
        Case COPYINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcCopyNm, imBSMode, imComboBoxIndex
    End Select
End Sub
Private Sub edcDropDown_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case imBoxNo
        Case TIMEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KEYDOWN) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case TIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    'plcTme.Visible = Not plcTme.Visible
                    lbcAvailTimes.Visible = Not lbcAvailTimes.Visible
                End If
            Case COPYINDEX
                gProcessArrowKey Shift, KeyCode, lbcCopyNm, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDTDropDown_Change()
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim ilFound As Integer
    Select Case imDTBoxNo
        Case DTDATEINDEX
            slStr = edcDTDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            If (rbcType(0).Value) Then
                If imDateChgMode = False Then
                    imDateChgMode = True
                    llDate = gDateValue(slStr)
                    ilFound = False
                    For ilLoop = 0 To UBound(tgWkDates) - 1 Step 1
                        If tgWkDates(ilLoop).lDate = llDate Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        lacDate.Visible = False
                        imDateChgMode = False
                        Exit Sub
                    End If
                    lacDate.Visible = True
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint   'mBoxCalDate called within paint
                    imDateChgMode = False
                End If
            Else
                lacDate.Visible = True
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint   'mBoxCalDate called within paint
            End If
        '6/16/11
        Case imDTAIRTIMEINDEX
        Case imDTAVAILTIMEINDEX
            'If (rbcType(0).Value) And (smSpotType <> "O") And (smSpotType <> "C") Then
                slStr = edcDTDropDown.Text
                gFindMatch slStr, 0, lbcAvailTimes
                imChgMode = True ' Turn on the switch
                If gLastFound(lbcAvailTimes) >= 0 Then
                    lbcAvailTimes.ListIndex = gLastFound(lbcAvailTimes)
                Else ' No data found so re-display the last good data
                    lbcAvailTimes.ListIndex = -1
                End If
                imChgMode = False
            'End If
    End Select
End Sub
Private Sub edcDTDropDown_GotFocus()
    Select Case imDTBoxNo
        Case DTDATEINDEX
        Case imDTAIRTIMEINDEX
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDTDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDTDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDTDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case imDTBoxNo
        Case DTDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case imDTAIRTIMEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
    End Select
End Sub
Private Sub edcDTDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imDTBoxNo
            Case DTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDTDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDTDropDown.Text = slDate
                    End If
                End If
            Case imDTAIRTIMEINDEX
                If rbcType(0).Value Then
                    'If (Shift And vbAltMask) > 0 Then
                    '    lbcAvailTimes.Visible = Not lbcAvailTimes.Visible
                    'End If
                    If (Shift And vbAltMask) > 0 Then
                        plcTme.Visible = Not plcTme.Visible
                    End If
                Else
                    If (Shift And vbAltMask) > 0 Then
                        plcTme.Visible = Not plcTme.Visible
                    End If
                End If
            Case imDTAVAILTIMEINDEX
                If rbcType(0).Value Then
                    If (Shift And vbAltMask) > 0 Then
                        lbcAvailTimes.Visible = Not lbcAvailTimes.Visible
                    End If
                'Else
                '    If (Shift And vbAltMask) > 0 Then
                '        plcTme.Visible = Not plcTme.Visible
                '    End If
                End If
        End Select
        edcDTDropDown.SelStart = 0
        edcDTDropDown.SelLength = Len(edcDTDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imDTBoxNo
            Case DTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDTDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDTDropDown.Text = slDate
                    End If
                End If
                edcDTDropDown.SelStart = 0
                edcDTDropDown.SelLength = Len(edcDTDropDown.Text)
            Case imDTAIRTIMEINDEX
        End Select
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcMdDate_Change(Index As Integer)
    Dim slStr As String
    Dim llDate As Long
    tmcClick.Enabled = False
    slStr = edcMdDate(Index).Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    If imDateChgMode = False Then
        imDateChgMode = True
        llDate = gDateValue(slStr)
        lacDate.Visible = True
        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
        pbcCalendar_Paint   'mBoxCalDate called within paint
        tmcClick.Interval = 2000    '2 seconds
        tmcClick.Enabled = True
        imDateChgMode = False
    End If
End Sub
Private Sub edcMdDate_GotFocus(Index As Integer)
    'tmcClick.Enabled = False
    If imTerminate Then
        Exit Sub
    End If
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    If imDateBox <> Index + 1 Then
        plcCalendar.Visible = False
        plcCalendar.Move plcSpots.Left + edcMdDate(Index).Left - fgBevelX, plcSpots.Top + edcMdDate(Index).height + fgBevelY
        imDateBox = Index + 1
    End If
    imBypassFocus = False
End Sub
Private Sub edcMdDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcMdDate_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcMdDate(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    imBSMode = False
End Sub
Private Sub edcMdDate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcMdDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcMdDate(Index).Text = slDate
            End If
        End If
        edcMdDate(Index).SelStart = 0
        edcMdDate(Index).SelLength = Len(edcMdDate(Index).Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcMdDate(Index).Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcMdDate(Index).Text = slDate
            End If
        End If
        edcMdDate(Index).SelStart = 0
        edcMdDate(Index).SelLength = Len(edcMdDate(Index).Text)
    End If
End Sub
Private Sub edcTZDropDown_Change()
    Select Case imTZBoxNo
        Case TZCOPYINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcTZDropDown, lbcCopyNm, imBSMode, imComboBoxIndex
            imLbcArrowSetting = False
    End Select
End Sub
Private Sub edcTZDropDown_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcTZDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTZDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTZDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcTZDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KEYDOWN) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imTZBoxNo
            Case COPYINDEX
                gProcessArrowKey Shift, KeyCode, lbcCopyNm, imLbcArrowSetting
        End Select
        edcTZDropDown.SelStart = 0
        edcTZDropDown.SelLength = Len(edcTZDropDown.Text)
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True  'To get Alt J and Alt L keys
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(POSTLOGSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    imSvUpdateAllowed = imUpdateAllowed
    DoEvents    'Process events so pending keys are not sent to this
                'form when keypreview turn on
    'If Not imTerminate Then
    '    PostLog.KeyPreview = True  'To get Alt J and Alt L keys
    'End If
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    mPaintMissedTitle
    PostLog.Refresh
    
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
        cmcAdServer.Visible = True
    Else
        cmcAdServer.Visible = False
    End If
    
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Deactivate()
    'PostLog.KeyPreview = False
    Me.KeyPreview = False
End Sub
Private Sub Form_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub Form_DragOver(Source As control, X As Single, Y As Single, State As Integer)
     If ((State = 2) And (X >= lbcMissed.Left) And (X <= (lbcMissed.Left + lbcMissed.Width)) And (Y < lbcMissed.Top)) Then
        Exit Sub
     End If
     If ((State = 1) And (X >= lbcMissed.Left) And (X <= (lbcMissed.Left + lbcMissed.Width)) And (Y > (lbcMissed.Top + lbcMissed.height))) Then
        Exit Sub
     End If
     '  Must not be in Hot Spot so turn off timer
     tmcDrag.Enabled = False
     tmcDrag.Interval = 1000
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        lbcGameNo.Visible = False
        If (imBoxNo > 0) Then
            plcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imDTBoxNo > 0 Then
            mDTEnableBox imDTBoxNo
        ElseIf imTZBoxNo > 0 Then
            mTZEnableBox imTZBoxNo
        Else
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            plcSelect.Enabled = True
        End If
    End If
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100) / Me.height
        Me.height = (lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    ilRet = btrClose(hmSxf)
    btrDestroy hmSxf
    ilRet = btrClose(hmRlf)
    btrDestroy hmRlf
    btrExtClear hmFsf   'Clear any previous extend operation
    ilRet = btrClose(hmFsf)
    btrDestroy hmFsf
    btrExtClear hmFnf   'Clear any previous extend operation
    ilRet = btrClose(hmFnf)
    btrDestroy hmFnf
    btrExtClear hmPrf   'Clear any previous extend operation
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    btrExtClear hmLcf   'Clear any previous extend operation
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    btrExtClear hmGhf   'Clear any previous extend operation
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    btrExtClear hmGsf   'Clear any previous extend operation
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    btrExtClear hmCgf   'Clear any previous extend operation
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    btrExtClear hmClf   'Clear any previous extend operation
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    btrExtClear hmCHF   'Clear any previous extend operation
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    btrExtClear hmSmf   'Clear any previous extend operation
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    btrExtClear hmSdf   'Clear any previous extend operation
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmAgf   'Clear any previous extend operation
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    btrExtClear hmSsf   'Clear any previous extend operation
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    btrExtClear hmVsf   'Clear any previous extend operation
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    btrExtClear hmVLF   'Clear any previous extend operation
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    btrExtClear hmCcf   'Clear any previous extend operation
    ilRet = btrClose(hmCcf)
    btrDestroy hmCcf
    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    btrExtClear hmTzf   'Clear any previous extend operation
    ilRet = btrClose(hmTzf)
    btrDestroy hmTzf
    btrExtClear hmMcf   'Clear any previous extend operation
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    btrExtClear hmRdf   'Clear any previous extend operation
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    btrExtClear hmCff   'Clear any previous extend operation
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    btrExtClear hmIihf   'Clear any previous extend operation
    ilRet = btrClose(hmIihf)
    btrDestroy hmIihf
    ilRet = btrClose(hmPsf)
    btrDestroy hmPsf

    Erase tmCgfCff
    Erase tmTeam

    'Erase smShow
    Erase tgShow
    Erase tgSave
    'Erase smSave
    'Erase imSave
    Erase tgMdSdfRec
    Erase tgMdSaveInfo
    Erase tgMdShowInfo
    'Erase smMdShow
    'Erase lmMdRecPos
    Erase tmChfAdvtExt
    Erase tmCopyNmCode
    Erase tmMissedCode
    Erase tgDates
    Erase tgClfPostLog
    Erase tgCffPostLog

    Erase tmUserVehicle
    smUserVehicleTag = ""
    
    igJobShowing(POSTLOGSJOB) = False
    
    Set PostLog = Nothing

End Sub
Private Sub imcHidden_Click()
    Dim ilLoop As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    Dim ilRepaint As Integer

    imcTrash.Visible = False
    imcHidden.Visible = False
    ilRepaint = False
    'For ilLoop = UBound(tgShow) - 1 To LBound(tgShow) Step -1
    For ilLoop = UBound(tgShow) - 1 To LBONE Step -1
        If tgShow(ilLoop).iChk Then
            imRowNo = ilLoop
            llRecPos = tgSave(tgShow(ilLoop).iSaveInfoIndex).lSdfRecPos   'Val(smSave(SAVRECPOSINDEX, imRowNo))
            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
            tmChfSrchKey.lCode = tmSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            smDragCntrType = tmChf.sType
            If tmSdf.sSpotType = "X" Then
                smDragCntrType = "X"
            ElseIf tmSdf.sSpotType = "O" Then
                smDragCntrType = "B"
            ElseIf tmSdf.sSpotType = "C" Then
                smDragCntrType = "B"
            End If
            If (rbcType(1).Value) Or ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "R") Or (tmChf.sType = "Q") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (tmSdf.sSpotType = "X") Or (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                tgShow(ilLoop).iChk = False
                ilRepaint = True
            Else
                imBoxNo = SCHTOHIDE
                imSdfChg = True
                ilRet = mSaveRec()
                imBoxNo = -1
                imRowNo = -1
            End If
        End If
    Next ilLoop
    If ilRepaint Then
        Beep
        pbcPosting_Paint
    End If
    If (pbcClickFocus.Enabled) And (pbcClickFocus.Visible) Then
        pbcClickFocus.SetFocus
    End If
End Sub
Private Sub imcHidden_DragDrop(Source As control, X As Single, Y As Single)
    Dim ilRet As Integer
    Dim ilCount As Integer
    Dim ilSvCount As Integer
    Dim ilIndex As Integer
    imcTrash.Visible = False
    imcHidden.Visible = False
    If imDragSource = 0 Then    'Post log spot
        If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
        Else
            imBoxNo = SCHTOHIDE
            imSdfChg = True
            ilRet = mSaveRec()
            imBoxNo = -1
            imRowNo = -1
            pbcClickFocus.SetFocus
            'If imRowNo < UBound(smSave, 2) Then
            '    imBoxNo = 1
            '    mEnableBox imBoxNo
            'Else
            '    imBoxNo = -1
            '    imRowNo = -1
            '    pbcArrow.Visible = False
            '    lacPtFrame.Visible = False
            'End If
        End If
        Exit Sub
    ElseIf imDragSource = 1 Then    'Missed
        If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Then
        Else
'            imBoxNo = MISSEDTOHIDE
'            imSdfChg = True
'            ilRet = mSaveRec()
'            imBoxNo = -1
'            imRowNo = -1
'            pbcClickFocus.SetFocus
            'Count number of missed spots
            ilCount = 0
            ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
            Do While ilIndex >= 0
                If (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
                    ilCount = ilCount + 1
                End If
                ilIndex = tgMdSdfRec(ilIndex).iNextIndex
            Loop
            If ilCount > 1 Then
                sgGenMsg = "How many missed spots should be set to " & "Hidden"
                sgCMCTitle(0) = "Change"
                sgCMCTitle(1) = "Cancel"
                sgCMCTitle(2) = ""
                sgCMCTitle(3) = ""
                igDefCMC = 0
                igEditBox = 1
                sgEditValue = Trim$(str$(ilCount))
                GenMsg.Show vbModal
                If igAnsCMC = 0 Then
                    If Val(sgEditValue) <= ilCount Then
                        ilCount = Val(sgEditValue)
                    End If
                Else
                    ilCount = 0
                End If
            End If
            If ilCount > 0 Then
                ilSvCount = ilCount
                ilCount = 0
                ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                Do While ilIndex >= 0
                    If (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
                        imSdfIndex = ilIndex
                        imBoxNo = MISSEDTOHIDE
                        imSdfChg = True
                        ilRet = mSaveRec()
                        ilCount = ilCount + 1
                        If ilCount >= ilSvCount Then
                            Exit Do
                        End If
                    End If
                    ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                Loop
            End If
        End If
        Exit Sub
    End If
End Sub
Private Sub imcHidden_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If imDragSource = 0 Then    'Post log spot
        If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
        Else
            If State = vbEnter Then    'Enter drag over
                lacPtFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
                imcHidden.Picture = IconTraf!imcHideDn.Picture
            ElseIf State = vbLeave Then
                lacPtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                imcHidden.Picture = IconTraf!imcHideUp.Picture
            End If
        End If
        Exit Sub
    ElseIf imDragSource = 1 Then    'Missed
        If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Then
        Else
            If State = vbEnter Then    'Enter drag over
                lacPtFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
                imcHidden.Picture = IconTraf!imcHideDn.Picture
            ElseIf State = vbLeave Then
                lacPtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                imcHidden.Picture = IconTraf!imcHideUp.Picture
            End If
        End If
        Exit Sub
    End If
End Sub

Private Sub imcKey_Click()
    pbcKey.Visible = Not pbcKey.Visible
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    pbcKey.ZOrder vbBringToFront
    'pbcKey.Visible = True
End Sub
Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = False
End Sub
Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    imcTrash.Visible = False
    imcHidden.Visible = False
    'For ilLoop = UBound(tgShow) - 1 To LBound(tgShow) Step -1
    For ilLoop = UBound(tgShow) - 1 To LBONE Step -1
        If tgShow(ilLoop).iChk Then
            imRowNo = ilLoop
            llRecPos = tgSave(tgShow(ilLoop).iSaveInfoIndex).lSdfRecPos   'Val(smSave(SAVRECPOSINDEX, imRowNo))
            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
            tmChfSrchKey.lCode = tmSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            smDragCntrType = tmChf.sType
            If tmSdf.sSpotType = "X" Then
                smDragCntrType = "X"
            ElseIf tmSdf.sSpotType = "O" Then
                smDragCntrType = "B"
            ElseIf tmSdf.sSpotType = "C" Then
                smDragCntrType = "B"
            End If
            If (rbcType(1).Value) Or ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "Q") Or (tmChf.sType = "R") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (tmSdf.sSpotType = "X") Or (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                imBoxNo = SCHTOMISSED
                pbcArrow.Visible = False
                lacPtFrame.Visible = False
                imSdfChg = True
                ilRet = mSaveRec()
                imBoxNo = -1
                imRowNo = -1
            'Else
            '    imBoxNo = SCHTOCANCEL
            '    imSdfChg = True
            '    ilRet = mSaveRec()
            '    imBoxNo = -1
            '    imRowNo = -1
            End If
        End If
    Next ilLoop
    If (pbcClickFocus.Enabled) And (pbcClickFocus.Visible) Then
        pbcClickFocus.SetFocus
    End If
End Sub
Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)
    Dim ilRet As Integer
    Dim ilCount As Integer
    Dim ilSvCount As Integer
    Dim ilIndex As Integer
    imcTrash.Visible = False
    imcHidden.Visible = False
    If imDragSource = 2 Then
        Exit Sub
    End If
    If imDragSource = 0 Then    'Post log spot
        If (rbcType(1).Value) Or ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
            imBoxNo = SCHTOMISSED
            pbcArrow.Visible = False
            lacPtFrame.Visible = False
            imSdfChg = True
            ilRet = mSaveRec()
            imBoxNo = -1
            imRowNo = -1
            pbcClickFocus.SetFocus
            'llRecPos = Val(smSave(SAVRECPOSINDEX, imRowNo))
            'ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
            ''lmSsfMemDate = 0
            'lmSsfDate(imSelectedDay) = 0
            'ilRet = gChgSchSpot("D", hmSdf, tmSdf, hmSmf, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay))
            'ilUpperBound = UBound(smSave, 2) - 1
            'For ilLoop = imRowNo To ilUpperBound - 1 Step 1
            '    For ilIndex = LBound(smShow, 1) To UBound(smShow, 1) Step 1
            '        smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop + 1)
            '    Next ilIndex
            '    For ilIndex = LBound(smSave, 1) To UBound(smSave, 1) Step 1
            '        smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop + 1)
            '    Next ilIndex
            '    For ilIndex = LBound(imSave, 1) To UBound(imSave, 1) Step 1
            '        imSave(ilIndex, ilLoop) = imSave(ilIndex, ilLoop + 1)
            '    Next ilIndex
            '    For ilIndex = LBound(imPostSpotInfo, 1) To UBound(imPostSpotInfo, 1) Step 1
            '        imPostSpotInfo(ilIndex, ilLoop) = imPostSpotInfo(ilIndex, ilLoop + 1)
            '    Next ilIndex
            'Next ilLoop
            'ReDim Preserve smShow(1 To 12, 1 To ilUpperBound) As String
            'ReDim Preserve smSave(1 To 10, 1 To ilUpperBound) As String
            'ReDim Preserve imSave(1 To 2, 1 To ilUpperBound) As Integer
            'ReDim Preserve imPostSpotInfo(1 To 3, 1 To ilUpperBound) As Integer
            'imBoxNo = -1
            'imRowNo = -1
            'pbcArrow.Visible = False
            'lacPtFrame.Visible = False
            'pbcPosting.Cls
            'pbcPosting_Paint
            'pbcClickFocus.SetFocus
        Else
        '    imBoxNo = SCHTOCANCEL
        '    imSdfChg = True
        '    ilRet = mSaveRec()
        '    imBoxNo = -1
        '    imRowNo = -1
        '    pbcClickFocus.SetFocus
        '    'If imRowNo < UBound(smSave, 2) Then
        '    '    imBoxNo = 1
        '    '    mEnableBox imBoxNo
        '    'Else
        '    '    imBoxNo = -1
        '    '    imRowNo = -1
        '    '    pbcArrow.Visible = False
        '    '    lacPtFrame.Visible = False
        '    'End If
        End If
        Exit Sub
    ElseIf imDragSource = 1 Then    'Missed
        If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Then
        Else
        '    ilCount = 0
        '    ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
        '    Do While ilIndex >= 0
        '        If (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
        '            ilCount = ilCount + 1
        '        End If
        '        ilIndex = tgMdSdfRec(ilIndex).iNextIndex
        '    Loop
        '    If ilCount > 1 Then
        '        sgGenMsg = "How many missed spots should be set to " & "Canceled"
        '        sgCMCTitle(0) = "Change"
        '        sgCMCTitle(1) = "Cancel"
        '        sgCMCTitle(2) = ""
        '        sgCMCTitle(3) = ""
        '        igDefCMC = 0
        '        igEditBox = 1
        '        sgEditValue = Trim$(str$(ilCount))
        '        GenMsg.Show vbModal
        '        If igAnsCMC = 0 Then
        '            If Val(sgEditValue) <= ilCount Then
        '                ilCount = Val(sgEditValue)
        '            End If
        '        Else
        '            ilCount = 0
        '        End If
        '    End If
        '    If ilCount > 0 Then
        '        ilSvCount = ilCount
        '        ilCount = 0
        '        ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
        '        Do While ilIndex >= 0
        '            If (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
        '                imSdfIndex = ilIndex
        '                imBoxNo = MISSEDTOCANCEL
        '                imSdfChg = True
        '                ilRet = mSaveRec()
        '                ilCount = ilCount + 1
        '                If ilCount >= ilSvCount Then
        '                    Exit Do
        '                End If
        '            End If
        '            ilIndex = tgMdSdfRec(ilIndex).iNextIndex
        '        Loop
        '    End If
        End If
        Exit Sub
    End If
End Sub
Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If imDragSource = 2 Then
        Exit Sub
    End If
    If imDragSource = 0 Then    'Post log spot
        If (rbcType(1).Value) Or ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
            If State = vbEnter Then    'Enter drag over
                lacPtFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
                imcTrash.Picture = IconTraf!imcFire.Picture
            ElseIf State = vbLeave Then
                lacPtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                imcTrash.Picture = IconTraf!imcFireOut.Picture
            End If
        Else
        '    If State = vbEnter Then    'Enter drag over
        '        lacPtFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        '        imcTrash.Picture = IconTraf!imcBoxOpened.Picture
        '    ElseIf State = vbLeave Then
                lacPtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                imcTrash.Picture = IconTraf!imcBoxClosed.Picture
        '    End If
        End If
        Exit Sub
    ElseIf imDragSource = 1 Then    'Missed
        If (rbcType(1).Value) Or ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
        Else
        '    If State = vbEnter Then    'Enter drag over
        '        lacPtFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        '        imcTrash.Picture = IconTraf!imcBoxOpened.Picture
        '    ElseIf State = vbLeave Then
                lacPtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                imcTrash.Picture = IconTraf!imcBoxClosed.Picture
        '    End If
        End If
        Exit Sub
    End If
End Sub
Private Sub lbcAvailTimes_Click()
    If imChgMode Then
        Exit Sub
    End If
    If lbcAvailTimes.ListIndex >= 0 Then
        'plcAvailTimes.Caption = lbcAvailTimes.List(lbcAvailTimes.ListIndex)
        edcDTDropDown.Text = lbcAvailTimes.List(lbcAvailTimes.ListIndex)
    Else
        edcDTDropDown.Text = ""
    End If
    'edcDTDropDown.SetFocus
End Sub

Private Sub lbcAvailTimes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilAnf As Integer
    
    If (X < 0) Or (X > lbcAvailTimes.Width) Then
        lbcAvailTimes.ToolTipText = ""
        imButtonIndex = -1
        Exit Sub
    End If
    If imButtonIndex <> (Y \ fgListHtArial825) + lbcAvailTimes.TopIndex Then
        imButtonIndex = Y \ fgListHtArial825 + lbcAvailTimes.TopIndex
        If (imButtonIndex >= 0) And (imButtonIndex <= lbcAvailTimes.ListCount - 1) Then
            ilAnf = lbcAvailTimes.ItemData(imButtonIndex)
            If ilAnf <> -1 Then
                lbcAvailTimes.ToolTipText = Trim$(tgAvailAnf(ilAnf).sName)
            Else
                lbcAvailTimes.ToolTipText = ""
            End If
        Else
            imButtonIndex = -1
            lbcAvailTimes.ToolTipText = ""
        End If
    End If
End Sub

Private Sub lbcCopyNm_Click()
    If imTZCopyAllowed Then
        gProcessLbcClick lbcCopyNm, edcTZDropDown, imChgMode, imLbcArrowSetting
    Else
        gProcessLbcClick lbcCopyNm, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcCopyNm_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcGameNo_Click()
    gProcessLbcClick lbcGameNo, edcDate, imGameNoChgMode, imLbcArrowSetting
End Sub

Private Sub lbcMissed_Click()
    Dim ilLoop As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    Dim ilRepaint As Integer
    Dim ilMissedListIndex As Integer
    If imListChgMode Then
        Exit Sub
    End If
    imcTrash.Visible = False
    imcHidden.Visible = False
    ilRepaint = False
    ilMissedListIndex = lbcMissed.ListIndex
    'For ilLoop = UBound(tgShow) - 1 To LBound(tgShow) Step -1
    For ilLoop = UBound(tgShow) - 1 To LBONE Step -1
        If tgShow(ilLoop).iChk Then
            imRowNo = ilLoop
            llRecPos = tgSave(tgShow(ilLoop).iSaveInfoIndex).lSdfRecPos   'Val(smSave(SAVRECPOSINDEX, imRowNo))
            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
            tmChfSrchKey.lCode = tmSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            smDragCntrType = tmChf.sType
            If tmSdf.sSpotType = "X" Then
                smDragCntrType = "X"
            ElseIf tmSdf.sSpotType = "O" Then
                smDragCntrType = "B"
            ElseIf tmSdf.sSpotType = "C" Then
                smDragCntrType = "B"
            End If
            '2/21/13: Allow Package spots to be cancelled
            'If (rbcType(1).Value) Or ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "Q") Or (tmChf.sType = "R") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (tmSdf.sSpotType = "X") Or (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
            If ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "Q") Or (tmChf.sType = "R") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (tmSdf.sSpotType = "X") Or (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                tgShow(ilLoop).iChk = False
                ilRepaint = True
            Else
                imMissedListIndex = ilMissedListIndex
                If imMissedType = 1 Then
                    imBoxNo = SCHTOCANCEL
                ElseIf imMissedType = 2 Then
                    imBoxNo = SCHTOHIDE
                Else
                    imBoxNo = SCHTOMISSED
                End If
                imSdfChg = True
                ilRet = mSaveRec()
                imBoxNo = -1
                imRowNo = -1
            End If
        End If
    Next ilLoop
    If ilRepaint Then
        Beep
        pbcPosting_Paint
        lbcMissed.ListIndex = -1
    End If
    'If (pbcClickFocus.Enabled) And (pbcClickFocus.Visible) Then
    '    pbcClickFocus.SetFocus
    'End If
    imListChgMode = False
End Sub
Private Sub lbcMissed_DragDrop(Source As control, X As Single, Y As Single)
    Dim ilRet As Integer
    Dim ilCount As Integer
    Dim ilIndex As Integer
    Dim ilSvCount As Integer
    Dim slNameCode As String
    Dim slName As String

    imcTrash.Visible = False
    imcHidden.Visible = False
    '2/21/13: Allow Package spots to be cancelled
    'If (rbcType(1).Value) Or ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
    If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
        Exit Sub
    End If
    If imDragSource = 0 Then    'Schedule spot
        'Make spot Missed (imRowNo contains the row number)
        imMissedListIndex = Y \ fgListHtArial825 + lbcMissed.TopIndex
        If imMissedListIndex <= UBound(tmMissedCode) - 1 Then 'lbcMissedCode.ListCount - 1 Then
            If imMissedType = 1 Then
                imBoxNo = SCHTOCANCEL
            ElseIf imMissedType = 2 Then
                imBoxNo = SCHTOHIDE
            Else
                imBoxNo = SCHTOMISSED
            End If
            imSdfChg = True
            ilRet = mSaveRec()
            imBoxNo = -1
            imRowNo = -1
            pbcClickFocus.SetFocus
            'If imRowNo < UBound(smSave, 2) Then
            '    imBoxNo = 1
            '    mEnableBox imBoxNo
            'Else
            '    imBoxNo = -1
            '    imRowNo = -1
            '    pbcArrow.Visible = False
            '    lacPtFrame.Visible = False
            'End If
        End If
'       lacPtFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    ElseIf imDragSource = 1 Then    'Missed spot
        imMissedListIndex = Y \ fgListHtArial825 + lbcMissed.TopIndex
        If imMissedListIndex <= UBound(tmMissedCode) - 1 Then 'lbcMissedCode.ListCount - 1 Then
            'Count number of missed spots
            ilCount = 0
            If tgMdShowInfo(imMdRowNo).iType = 1 Then
                ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                Do While ilIndex >= 0
                    If (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
                        ilCount = ilCount + 1
                    End If
                    ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                Loop
                If ilCount > 1 Then
                    slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    If imMissedType = 1 Then
                        sgGenMsg = "How many Missed spots should be set to Cancel with reason " & slName
                    ElseIf imMissedType = 2 Then
                        sgGenMsg = "How many Missed spots should be set to Hidden"
                    Else
                        sgGenMsg = "How many Missed spots should be set to " & slName
                    End If
                    sgCMCTitle(0) = "Change"
                    sgCMCTitle(1) = "Cancel"
                    sgCMCTitle(2) = ""
                    sgCMCTitle(3) = ""
                    igDefCMC = 0
                    igEditBox = 1
                    sgEditValue = Trim$(str$(ilCount))
                    GenMsg.Show vbModal
                    If igAnsCMC = 0 Then
                        If Val(sgEditValue) <= ilCount Then
                            ilCount = Val(sgEditValue)
                        End If
                    Else
                        ilCount = 0
                    End If
                End If
            ElseIf tgMdShowInfo(imMdRowNo).iType = 3 Then
                ilCount = 0
                ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                Do While ilIndex >= 0
                    If (tgMdSdfRec(ilIndex).sSchStatus = "C") Then
                        ilCount = ilCount + 1
                    End If
                    ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                Loop
                If ilCount > 1 Then
                    slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    If imMissedType = 1 Then
                        sgGenMsg = "How many Canceled spots should be set to " & slName
                    ElseIf imMissedType = 2 Then
                        sgGenMsg = "How many Canceled spots should be set to Hidden"
                    Else
                        sgGenMsg = "How many Canceled spots should be set to Missed with reason " & slName
                    End If
                    sgCMCTitle(0) = "Change"
                    sgCMCTitle(1) = "Cancel"
                    sgCMCTitle(2) = ""
                    sgCMCTitle(3) = ""
                    igDefCMC = 0
                    igEditBox = 1
                    sgEditValue = Trim$(str$(ilCount))
                    GenMsg.Show vbModal
                    If igAnsCMC = 0 Then
                        If Val(sgEditValue) <= ilCount Then
                            ilCount = Val(sgEditValue)
                        End If
                    Else
                        ilCount = 0
                    End If
                End If
            ElseIf tgMdShowInfo(imMdRowNo).iType = 2 Then
                ilCount = 0
                If imMissedType <> 2 Then
                    ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                    Do While ilIndex >= 0
                        If (tgMdSdfRec(ilIndex).sSchStatus = "H") Then
                            ilCount = ilCount + 1
                        End If
                        ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                    Loop
                End If
                If ilCount > 1 Then
                    slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    If imMissedType = 1 Then
                        sgGenMsg = "How many Hidden spots should be set to Cancel with reason " & slName
                    Else
                        sgGenMsg = "How many Hidden spots should be set to Missed with reason " & slName
                    End If
                    sgCMCTitle(0) = "Change"
                    sgCMCTitle(1) = "Cancel"
                    sgCMCTitle(2) = ""
                    sgCMCTitle(3) = ""
                    igDefCMC = 0
                    igEditBox = 1
                    sgEditValue = Trim$(str$(ilCount))
                    GenMsg.Show vbModal
                    If igAnsCMC = 0 Then
                        If Val(sgEditValue) <= ilCount Then
                            ilCount = Val(sgEditValue)
                        End If
                    Else
                        ilCount = 0
                    End If
                End If
            End If
            If ilCount > 0 Then
                ilSvCount = ilCount
                ilCount = 0
                If tgMdShowInfo(imMdRowNo).iType = 1 Then
                    ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                    Do While ilIndex >= 0
                        If (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
                            imSdfIndex = ilIndex
                            If imMissedType = 1 Then
                                imBoxNo = MISSEDTOCANCEL
                            ElseIf imMissedType = 2 Then
                                imBoxNo = MISSEDTOHIDE
                            Else
                                imBoxNo = MISSEDREASON
                            End If
                            imSdfChg = True
                            ilRet = mSaveRec()
                            ilCount = ilCount + 1
                            If ilCount >= ilSvCount Then
                                Exit Do
                            End If
                        End If
                        ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                    Loop
                ElseIf tgMdShowInfo(imMdRowNo).iType = 3 Then
                    ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                    Do While ilIndex >= 0
                        If (tgMdSdfRec(ilIndex).sSchStatus = "C") Then
                            imSdfIndex = ilIndex
                            If imMissedType = 1 Then
                                imBoxNo = MISSEDTOCANCEL
                            ElseIf imMissedType = 2 Then
                                imBoxNo = MISSEDTOHIDE
                            Else
                                imBoxNo = MISSEDREASON
                            End If
                            imSdfChg = True
                            ilRet = mSaveRec()
                            ilCount = ilCount + 1
                            If ilCount >= ilSvCount Then
                                Exit Do
                            End If
                        End If
                        ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                    Loop
                ElseIf tgMdShowInfo(imMdRowNo).iType = 2 Then
                    ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                    Do While ilIndex >= 0
                        If (tgMdSdfRec(ilIndex).sSchStatus = "H") Then
                            imSdfIndex = ilIndex
                            imBoxNo = MISSEDREASON
                            If imMissedType = 1 Then
                                imBoxNo = MISSEDTOCANCEL
                            Else
                                imBoxNo = MISSEDREASON
                            End If
                            imSdfChg = True
                            ilRet = mSaveRec()
                            ilCount = ilCount + 1
                            If ilCount >= ilSvCount Then
                                Exit Do
                            End If
                        End If
                        ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                    Loop
                End If
            End If
        End If
'        lacMdFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    End If
End Sub
Private Sub lbcMissed_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    Dim ilListIndex As Integer
    If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
        'imcTrash.Visible = False
        Exit Sub
    End If
    If imDragSource = 2 Then
        Exit Sub
    End If
    imListChgMode = True
    If imDragSource = 0 Then
        If State = vbEnter Then    'Enter drag over
            lacPtFrame.DragIcon = IconTraf!imcIconMove.DragIcon
            lbcMissed.ListIndex = -1
        ElseIf State = vbLeave Then
            lacPtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
            lbcMissed.ListIndex = -1
        ElseIf State = vbOver Then
            ilListIndex = Y \ fgListHtArial825 + lbcMissed.TopIndex
            If ilListIndex <= lbcMissed.ListCount - 1 Then
                lbcMissed.ListIndex = ilListIndex
            Else
                lbcMissed.ListIndex = -1
            End If
        End If
     End If
    If imDragSource = 1 Then
        If State = vbEnter Then    'Enter drag over
            lacMdFrame.DragIcon = IconTraf!imcIconChg.DragIcon
            lbcMissed.ListIndex = -1
        ElseIf State = vbLeave Then
            lacMdFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
            lbcMissed.ListIndex = -1
        ElseIf State = vbOver Then
            ilListIndex = Y \ fgListHtArial825 + lbcMissed.TopIndex
            If ilListIndex <= lbcMissed.ListCount - 1 Then
                lbcMissed.ListIndex = ilListIndex
            Else
                lbcMissed.ListIndex = -1
            End If
        End If
    End If
    imListChgMode = False
    ' Turn off timer if entering
    If State = ENTER Then
        tmcDrag.Enabled = False ' Turn off timer
        tmcDrag.Interval = 1000 ' reset to one second
    End If
    ' Turn on timer if leaving
    If State = LEAVE Then
        tmcDrag.Interval = imTimerInterval
        tmcDrag.Enabled = True
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailPop                       *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Avail Pop the selection Avail  *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mAvailPop()
'
'   mAvailPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilLp As Integer
    Dim slStr As String
    ilIndex = cbcAvailName.ListIndex
    If ilIndex > 0 Then
        slName = cbcAvailName.List(ilIndex)
    End If
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    'ilRet = gIMoveListBox(PEvent, lbcEvtAvail, lbcEvtAvailCode, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(PostLog, cbcAvailName, tmAvailCode(), smAvailCodeTag, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        'Remove "Post Log" avail name
        For ilLoop = 0 To cbcAvailName.ListCount - 1 Step 1
            slStr = Trim$(cbcAvailName.List(ilLoop))
            If StrComp(slStr, "Post Log", 1) = 0 Then
                cbcAvailName.RemoveItem ilLoop
                For ilLp = ilLoop To UBound(tmAvailCode) - 1 Step 1
                    tmAvailCode(ilLp) = tmAvailCode(ilLp + 1)
                Next ilLp
                ReDim Preserve tmAvailCode(LBound(tmAvailCode) To UBound(tmAvailCode) - 1) As SORTCODE
                Exit For
            End If
        Next ilLoop
        On Error GoTo mAvailPopErr
        gCPErrorMsg ilRet, "mAvailPop (gIMoveListBox)", PostLog
        On Error GoTo 0
        cbcAvailName.AddItem "[All]", 0  'Force as first item on list
        'cbcAvail.Height = gListBoxHeight(cbcAvail.ListCount, 10)
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, cbcAvailName
            If gLastFound(cbcAvailName) > 0 Then
                cbcAvailName.ListIndex = gLastFound(cbcAvailName)
            Else
                cbcAvailName.ListIndex = -1
            End If
        Else
            cbcAvailName.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mAvailPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailRoom                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if room exist for    *
'*                      spot within avail              *
'*                                                     *
'*******************************************************
Private Function mAvailRoom(ilAvailIndex) As Integer
'
'   ilRet = mAvailRoom(ilAvailIndex)
'   where:
'       ilAvailIndex(I)- location of avail within Ssf (use mFindAvail)
'       ilRet(O)- True=Avail has room; False=insufficient room within avail
'
'       tmSdf(I)- spot records
'
'       Code later: ask if avail should be overbooked
'                   If so, create a version zero (0) of the library with the new
'                   units/seconds
'
    Dim ilAvailUnits As Integer
    Dim ilAvailSec As Integer
    Dim ilUnitsSold As Integer
    Dim ilSecSold As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSpotIndex As Integer
    Dim ilNewUnit As Integer
    Dim ilNewSec As Integer
    Dim ilRet As Integer
   LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex)
    ilAvailUnits = tmAvail.iAvInfo And &H1F
    ilAvailSec = tmAvail.iLen
    '10/27/11: Disallow more then 31 spots in any avail
    If tmAvail.iNoSpotsThis >= 31 Then
        ilRet = MsgBox("Move not allowed because Avail contains the maximum number of spots (31).", vbOKOnly + vbExclamation, "Save")
        mAvailRoom = False
        Exit Function
    End If
    For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
       LSet tmSpot = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilSpotIndex)
        If tmSpot.lSdfCode = tmSdf.lCode Then
            mAvailRoom = True
            Exit Function
        End If
        If (tmSpot.iRecType And &HF) >= 10 Then
            ilSpotLen = tmSpot.iPosLen And &HFFF
            If (tgVpf(imVpfIndex).sSSellOut = "T") Then
                ilSpotUnits = ilSpotLen \ 30
                If ilSpotUnits <= 0 Then
                    ilSpotUnits = 1
                End If
                ilSpotLen = 0
            Else
                ilSpotUnits = 1
                'If (tgVpf(imVpfIndex).sSSellOut = "U") Then
                '    ilSpotLen = 0
                'End If
            End If
            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                ilUnitsSold = ilUnitsSold + ilSpotUnits
                ilSecSold = ilSecSold + ilSpotLen
            End If
        End If
    Next ilSpotIndex
    ilSpotLen = tmSdf.iLen
    If (tgVpf(imVpfIndex).sSSellOut = "T") Then
        ilSpotUnits = ilSpotLen \ 30
        If ilSpotUnits <= 0 Then
            ilSpotUnits = 1
        End If
        ilSpotLen = 0
    Else
        ilSpotUnits = 1
        'If (tgVpf(imVpfIndex).sSSellOut = "U") Then
        '    ilSpotLen = 0
        'End If
    End If
    ilNewUnit = 0
    ilNewSec = 0
    If (tgVpf(imVpfIndex).sSSellOut = "M") Then
        If (ilSpotLen + ilSecSold <> ilAvailSec) Or (ilSpotUnits + ilUnitsSold <> ilAvailUnits) Then
            ilNewSec = ilSpotLen + ilSecSold
            ilNewUnit = ilSpotUnits + ilUnitsSold
        Else
            mAvailRoom = True
            Exit Function
        End If
    Else
        If (ilSpotLen + ilSecSold > ilAvailSec) Or (ilSpotUnits + ilUnitsSold > ilAvailUnits) Then
            ilNewSec = ilSpotLen + ilSecSold
            ilNewUnit = ilSpotUnits + ilUnitsSold
        Else
            mAvailRoom = True
            Exit Function
        End If
    End If
    If (tgVpf(imVpfIndex).sSOverBook <> "Y") Then
        ilRet = MsgBox("Move not allowed because Avail would be Overbooked.", vbOKOnly + vbExclamation, "Save")
        mAvailRoom = False
        Exit Function
    End If
    Do
        imSsfRecLen = Len(tmSsf(imSelectedDay))
        ilRet = gSSFGetDirect(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, lmSsfRecPos(imSelectedDay), INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf(imSelectedDay))
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        '5/20/11
        If (tmAvail.iOrigUnit = 0) And (tmAvail.iOrigLen = 0) Then
            tmAvail.iOrigUnit = tmAvail.iAvInfo And &H1F
            tmAvail.iOrigLen = tmAvail.iLen
        End If
        tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilNewUnit
        tmAvail.iLen = ilNewSec
        tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
        imSsfRecLen = igSSFBaseLen + tmSsf(imSelectedDay).iCount * Len(tmProg)
        ilRet = gSSFUpdate(hmSsf, tmSsf(imSelectedDay), imSsfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        mAvailRoom = False
        Exit Function
    End If
    mAvailRoom = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailTimePop                   *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box with avail   *
'*                      times                          *
'*                                                     *
'*******************************************************
Private Sub mAvailTimePop(slAirDate As String, ilGameNo As Integer)
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slTime As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilAnfCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilType As Integer
    Dim ilAnf As Integer
    
    ilType = ilGameNo
    If rbcType(1).Value Then
        Exit Sub
    End If
    If imAvailSelectedIndex <= 0 Then
        ilAnfCode = -1
    Else
        slNameCode = tmAvailCode(imAvailSelectedIndex - 1).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilAnfCode = Val(slCode)
    End If
    If (gDateValue(slAirDate) = lmAvailDate) And (ilAnfCode = imAvailAnfCode) And (imAvailGameNo = ilGameNo) Then
        Exit Sub
    End If
    lbcAvailTimes.Clear
    lmAvailDate = gDateValue(slAirDate)
    imAvailAnfCode = ilAnfCode
    imAvailGameNo = ilGameNo
    gPackDate slAirDate, ilDate0, ilDate1
    imSelectedDay = gWeekDayStr(slAirDate)
    imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
    tmSsfSrchKey.iType = ilType 'slType-On Air
    tmSsfSrchKey.iVefCode = imVefCode
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf(imSelectedDay).iType = ilType) And (tmSsf(imSelectedDay).iVefCode = imVefCode) And (tmSsf(imSelectedDay).iDate(0) = ilDate0) And (tmSsf(imSelectedDay).iDate(1) = ilDate1)
        For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
           LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                ilAnfCode = tmAvail.ianfCode
                If (ilAnfCode = imAvailAnfCode) Or (imAvailAnfCode = -1) Then
                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                    lbcAvailTimes.AddItem slTime
                    lbcAvailTimes.ItemData(lbcAvailTimes.NewIndex) = -1
                    If (imAvailAnfCode = -1) Then
                        For ilAnf = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
                            If tgAvailAnf(ilAnf).iCode = ilAnfCode Then
                                lbcAvailTimes.ItemData(lbcAvailTimes.NewIndex) = ilAnf
                                Exit For
                            End If
                        Next ilAnf
                    End If
                End If
            End If
        Next ilLoop
        imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
        ilRet = gSSFGetNext(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    If imDTBoxNo = DTDATEINDEX Then
        slStr = edcDTDropDown.Text
    Else
        If imDateBox = 0 Then
            slStr = edcDate.Text
        ElseIf (imDateBox = 1) Or (imDateBox = 2) Then
            slStr = edcMdDate(imDateBox - 1).Text
        End If
    End If
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCbcVehChange                   *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process vehicle change         *
'*                                                     *
'*******************************************************
Private Sub mCbcVehChange()
    Dim slDate As String
    Dim ilLoopCount As Integer
    Dim slStr As String
    '  imChgMode is used to avoid entering this routine multiple times
    '            if a vehicle selection change occurs during the
    '            processing of a "change"
    If imChgMode = False Then
        imChgMode = True
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcVeh.ListIndex >= 0 Then
                    cbcVeh.Text = cbcVeh.List(cbcVeh.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            ' If there are characters in the combobox, look ahead
            '    to see if you can find a match
            If cbcVeh.Text <> "" Then
                gManLookAhead cbcVeh, imBSMode, imComboBoxIndex
            End If
            'imVehSelectedIndex is used to hold the index
            '   because VB has a bug
            imVehSelectedIndex = cbcVeh.ListIndex
            ' this function uses imVehSelectedIndex to find the vehicles
            '      option table vehicle index and returns imVpfIndex
            mGetVehIndex
            If imTerminate Then
                cbcAvailName.Enabled = True
                edcDate.Enabled = True
                cmcDate.Enabled = True
                Screen.MousePointer = vbDefault
                imChgMode = False
                Exit Sub
            End If
            ' Take the 2 integers and fetch a date into slDate
            ' Here were getting the LastLog Date
            ' tgVpf is a global array containing Vehicle Option Data
            '  iLLD(0) is mm & dd, iLLD(1) is yy
            pbcPosting.Cls
            mPaintPostTitle
            If tmVef.sType <> "V" Then
                '8/31/06: ttp 1854
                'If tgVpf(imVpfIndex).sGenLog = "N" Then
                'If (tgVpf(imVpfIndex).sGenLog = "N") Or (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) = USINGLIVELOG) And (tgVpf(imVpfIndex).sGenLog = "L")) Then
                If (tgVpf(imVpfIndex).sGenLog = "N") Or (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) = USINGLIVELOG) And ((tgVpf(imVpfIndex).sGenLog = "L") Or (tgVpf(imVpfIndex).sGenLog = "A"))) Then
                    slDate = Format$(gNow(), "m/d/yy")
                    lmVehLLD = gDateValue(slDate)
                Else
                    gUnpackDate tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), slDate
                    lmVehLLD = gDateValue(slDate)
                End If
            Else
                slDate = Format$(gNow(), "m/d/yy")
                lmVehLLD = gDateValue(slDate)
            End If
            mDatePop
            ReDim tgShow(0 To 1) As SHOWINFO
            ReDim tgSave(0 To 1) As SAVEINFO
            imSettingValue = True
            vbcPosting.Min = LBONE  'LBound(tgShow)
            imSettingValue = True
            vbcPosting.Max = LBONE  'LBound(tgShow)
            imSettingValue = True
            vbcPosting.Value = vbcPosting.Min
            imSettingValue = False
            'ReDim smMdShow(1 To 8, 1 To 1) As String
            'ReDim smMdSave(1 To 1, 1 To 1) As String
            'ReDim smMdSchStatus(1 To 1, 1 To 1) As String
            'ReDim lmMdRecPos(1 To 1) As Long
            imSvVehSelectedIndex = -1
            ReDim tgMdSdfRec(0 To 1) As MDSDFREC
            ReDim tgMdSaveInfo(0 To 1) As MDSAVEINFO
            ReDim tgMdShowInfo(0 To 1) As MDSHOWINFO
            vbcMissed.Min = LBONE   'LBound(tgMdShowInfo)
            vbcMissed.Max = LBONE   'LBound(tgMdShowInfo)
            vbcMissed.Value = vbcMissed.Min
            pbcMissed.Cls
            slStr = cbcVeh.List(imVehSelectedIndex)
            cbcVehicle.List(0) = slStr
    '        Read SDF and build smSave, imSave, smShow and call paint
        Loop While imVehSelectedIndex <> cbcVeh.ListIndex
        cbcAvailName.Enabled = True
        edcDate.Enabled = True
        cmcDate.Enabled = True
        pbcMissed_Paint
        Screen.MousePointer = vbDefault    'Default
        imChgMode = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyPop                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mCopyPop()
    Dim ilRet As Integer
    Dim ilAdvtCode As Integer
    Dim ilISCIProd As Integer
    Dim llRecPos As Long
    If tgSpf.sUseCartNo <> "N" Then
        ilISCIProd = 2
    Else
        ilISCIProd = 5
    End If
    ' Read SDF to get advertiser
    llRecPos = tgSave(tgShow(imRowNo).iSaveInfoIndex).lSdfRecPos   'Val(smSave(SAVRECPOSINDEX, imRowNo))
    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
    ilAdvtCode = tmSdf.iAdfCode
    'ilRet = gPopCopyForAdvtBox(PostLog, ilAdvtCode, ilISCIProd, 1, lbcCopyNm, lbcCopyNmCode)
    If (smSpotType = "O") Or (smSpotType = "C") Then
        ilRet = gPopCopyForAdvtBox(PostLog, ilAdvtCode, ilISCIProd, 1 + &H200, lbcCopyNm, tmCopyNmCode(), smCopyNmCodeTag)
    Else
        ilRet = gPopCopyForAdvtBox(PostLog, ilAdvtCode, ilISCIProd, 1, lbcCopyNm, tmCopyNmCode(), smCopyNmCodeTag)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mCopyPopErr
        gCPErrorMsg ilRet, "mCopyPop (gPopCopyForAdvtBox: PostLog)", PostLog
        On Error GoTo 0
        lbcCopyNm.AddItem "[None]", 0   ' Add this at the top of the list
    End If
    Exit Sub
mCopyPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateMdShow                   *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create Missed Show values      *
'*                                                     *
'*******************************************************
Private Sub mCreateMdShow(ilResetToValue As Integer)
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slVehName As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilVef As Integer
    Dim slMultiTimes As String
    Dim slStrTime As String
    Dim slStart As String
    Dim slEnd As String
    Dim slTime As String
    Dim ilDay As Integer
    Dim ilRdf As Integer
    Dim ilValue As Integer
    ilValue = vbcMissed.Value
    ReDim tgMdShowInfo(0 To 1) As MDSHOWINFO
    ilUpper = UBound(tgMdShowInfo)
    'For ilLoop = LBound(tgMdSaveInfo) To UBound(tgMdSaveInfo) - 1 Step 1
    For ilLoop = LBONE To UBound(tgMdSaveInfo) - 1 Step 1
        'Advertiser
        If tgMdSaveInfo(ilLoop).iAdfCode <> tmAdf.iCode Then
            tmAdfSrchKey.iCode = tgMdSaveInfo(ilLoop).iAdfCode
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                'slStr = Trim$(tmAdf.sName)
                If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                    slStr = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "/Direct"
                Else
                    slStr = Trim$(tmAdf.sName)
                End If
            Else
                slStr = ""
                tmAdf.sName = ""
            End If
        Else
            'slStr = Trim$(tmAdf.sName)
            If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                slStr = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "/Direct"
            Else
                slStr = Trim$(tmAdf.sName)
            End If
        End If
        gSetShow pbcMissed, slStr, tmMdCtrls(MDADVTINDEX)
        tgMdShowInfo(ilUpper).sShow(MDADVTINDEX) = tmMdCtrls(MDADVTINDEX).sShow
        'Contract # and Product
        tgMdShowInfo(ilUpper).sShow(MDPRODINDEX) = ""
        If tgMdSaveInfo(ilLoop).lChfCode > 0 Then
            slStr = Trim$(str$(tgMdSaveInfo(ilLoop).lCntrNo))
            tmChfSrchKey.lCode = tgMdSaveInfo(ilLoop).lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                gSetShow pbcMissed, Trim$(tmChf.sProduct), tmMdCtrls(MDVEHINDEX)
                tgMdShowInfo(ilUpper).sShow(MDPRODINDEX) = tmMdCtrls(MDVEHINDEX).sShow
            End If
        Else
            ilRet = mReadChfClfRdfRec(tgMdSaveInfo(ilLoop).lChfCode, 0, tgMdSaveInfo(ilLoop).lFsfCode)
            slStr = "Feed"
        End If
        gSetShow pbcMissed, slStr, tmMdCtrls(MDCNTRINDEX)
        tgMdShowInfo(ilUpper).sShow(MDCNTRINDEX) = tmMdCtrls(MDCNTRINDEX).sShow
        'Vehicle
        slVehName = "                    "
        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMdSaveInfo(ilLoop).iVefCode = tgMVef(ilVef).iCode Then
            ilVef = gBinarySearchVef(tgMdSaveInfo(ilLoop).iVefCode)
            If ilVef <> -1 Then
                slStr = Trim$(tgMVef(ilVef).sName)
                slVehName = tgMVef(ilVef).sName
                gSetShow pbcMissed, slStr, tmMdCtrls(MDVEHINDEX)
                tgMdShowInfo(ilUpper).sShow(MDVEHINDEX) = tmMdCtrls(MDVEHINDEX).sShow
        '        Exit For
            End If
        'Next ilVef
        'Length
        slStr = Trim$(str$(tgMdSaveInfo(ilLoop).iLen))
        gSetShow pbcMissed, slStr, tmMdCtrls(MDLENINDEX)
        tgMdShowInfo(ilUpper).sShow(MDLENINDEX) = tmMdCtrls(MDLENINDEX).sShow
        'Week Missed
        If tgMdSaveInfo(ilLoop).lWkMissed > 0 Then
            slStr = Format$(tgMdSaveInfo(ilLoop).lWkMissed, "m/d/yy")
        Else
            slStr = " "
        End If
        gSetShow pbcMissed, slStr, tmMdCtrls(MDWKMISSINDEX)
        tgMdShowInfo(ilUpper).sShow(MDWKMISSINDEX) = tmMdCtrls(MDWKMISSINDEX).sShow
        'End Date
        If tgMdSaveInfo(ilLoop).lEndDate > 0 Then
            slStr = Format$(tgMdSaveInfo(ilLoop).lEndDate, "m/d/yy")
        Else
            slStr = " "
        End If
        gSetShow pbcMissed, slStr, tmMdCtrls(MDENDDATEINDEX)
        tgMdShowInfo(ilUpper).sShow(MDENDDATEINDEX) = tmMdCtrls(MDENDDATEINDEX).sShow
        'Daypart name or Override infomation
        If tgMdSaveInfo(ilLoop).lChfCode > 0 Then
            tmRdfSrchKey.iCode = tgMdSaveInfo(ilLoop).iRdfCode  ' Daypart File Code
            ilRet = btrGetEqual(hmRdf, tmLnRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmLnRdf.sName = ""
            End If
            If tgMdSaveInfo(ilLoop).iRdfCode <> tmLnRdf.iCode Then
                tmRdfSrchKey.iCode = tgMdSaveInfo(ilLoop).iRdfCode
                ilRet = btrGetEqual(hmRdf, tmLnRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slStr = Trim$(tmLnRdf.sName)
                Else
                    slStr = ""
                End If
            Else
                slStr = Trim$(tmLnRdf.sName)
            End If
        Else
            slStr = Trim$(tmLnRdf.sName)
        End If
        If (tgMdSaveInfo(ilLoop).iStartTime(0) <> 1) Or (tgMdSaveInfo(ilLoop).iStartTime(1) <> 0) Then
            gUnpackTime tgMdSaveInfo(ilLoop).iStartTime(0), tgMdSaveInfo(ilLoop).iStartTime(1), "A", "1", slStartTime
            gUnpackTime tgMdSaveInfo(ilLoop).iEndTime(0), tgMdSaveInfo(ilLoop).iEndTime(1), "A", "1", slEndTime
            slStr = slStartTime & "-" & slEndTime
        Else
            'Add times
            slStrTime = ""
            slMultiTimes = ""
            slStartTime = ""
            For ilRdf = LBound(tmLnRdf.iStartTime, 2) To UBound(tmLnRdf.iStartTime, 2) Step 1 'Row
                If (tmLnRdf.iStartTime(0, ilRdf) <> 1) Or (tmLnRdf.iStartTime(1, ilRdf) <> 0) Then
                    gUnpackTime tmLnRdf.iStartTime(0, ilRdf), tmLnRdf.iStartTime(1, ilRdf), "A", "1", slStart
                    gUnpackTime tmLnRdf.iEndTime(0, ilRdf), tmLnRdf.iEndTime(1, ilRdf), "A", "1", slEnd
                    If slStart <> "" Then
                        slStrTime = slStart & "-" & slEnd
                        slStartTime = slStart
                        If ilRdf < UBound(tmLnRdf.iStartTime, 2) Then
                            If (tmLnRdf.iStartTime(0, ilRdf + 1) <> 1) Or (tmLnRdf.iStartTime(1, ilRdf + 1) <> 0) Then
                                slMultiTimes = "+"
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next ilRdf
            slStr = slStr & " " & slStrTime & slMultiTimes
        End If
        slStr = slStr & " "
        If tgMdSaveInfo(ilLoop).lWkMissed > 0 Then
            For ilDay = 0 To 6 Step 1
                If tgMdSaveInfo(ilLoop).iDay(ilDay) Then
                    slStr = slStr & "Y"
                Else
                    slStr = slStr & "N"
                End If
            Next ilDay
        End If
        gSetShow pbcMissed, slStr, tmMdCtrls(MDDPINDEX)
        tgMdShowInfo(ilUpper).sShow(MDDPINDEX) = tmMdCtrls(MDDPINDEX).sShow
        slTime = Trim$(str$(gTimeToLong(slStartTime, False)))
        Do While Len(slTime) < 5
            slTime = "0" & slTime
        Loop
        If tgMdSaveInfo(ilLoop).lWkMissed > 0 Then
            slStr = Trim$(str$(tgMdSaveInfo(ilLoop).lWkMissed))
        Else
            slStr = "999999"
        End If
        Do While Len(slStr) < 5
            slStr = "0" & slStr
        Loop
        tgMdShowInfo(ilUpper).sBill = tgMdSaveInfo(ilLoop).sBill
        If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
            slStr = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "/Direct"
        Else
            slStr = Trim$(tmAdf.sName)
        End If
        tgMdShowInfo(ilUpper).sKey = slStr & slVehName & slTime & slStr
        If tgMdSaveInfo(ilLoop).lWkMissed > 0 Then
            If tgMdSaveInfo(ilLoop).iMissedCount > 0 Then
                slStr = Trim$(str$(tgMdSaveInfo(ilLoop).iMissedCount))
                gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
                tgMdShowInfo(ilUpper).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
                tgMdShowInfo(ilUpper).iType = 1
                tgMdShowInfo(ilUpper).iMdSaveInfoIndex = ilLoop
                ilUpper = ilUpper + 1
                ReDim Preserve tgMdShowInfo(0 To ilUpper) As MDSHOWINFO
            End If
            If tgMdSaveInfo(ilLoop).iHiddenCount > 0 Then
                If tgMdSaveInfo(ilLoop).iMissedCount > 0 Then
                    tgMdShowInfo(ilUpper) = tgMdShowInfo(ilUpper - 1)
                End If
                slStr = Trim$(str$(tgMdSaveInfo(ilLoop).iHiddenCount))
                gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
                tgMdShowInfo(ilUpper).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
                tgMdShowInfo(ilUpper).iType = 2
                tgMdShowInfo(ilUpper).iMdSaveInfoIndex = ilLoop
                ilUpper = ilUpper + 1
                ReDim Preserve tgMdShowInfo(0 To ilUpper) As MDSHOWINFO
            End If
            If tgMdSaveInfo(ilLoop).iCancelCount > 0 Then
                If (tgMdSaveInfo(ilLoop).iMissedCount > 0) Or (tgMdSaveInfo(ilLoop).iHiddenCount > 0) Then
                    tgMdShowInfo(ilUpper) = tgMdShowInfo(ilUpper - 1)
                End If
                slStr = Trim$(str$(tgMdSaveInfo(ilLoop).iCancelCount))
                gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
                tgMdShowInfo(ilUpper).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
                tgMdShowInfo(ilUpper).iType = 3
                tgMdShowInfo(ilUpper).iMdSaveInfoIndex = ilLoop
                ilUpper = ilUpper + 1
                ReDim Preserve tgMdShowInfo(0 To ilUpper) As MDSHOWINFO
            End If
        Else
            slStr = ""
            gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
            tgMdShowInfo(ilUpper).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
            tgMdShowInfo(ilUpper).iType = 0
            tgMdShowInfo(ilUpper).iMdSaveInfoIndex = ilLoop
            ilUpper = ilUpper + 1
            ReDim Preserve tgMdShowInfo(0 To ilUpper) As MDSHOWINFO
        End If
    Next ilLoop
    'Check for contracts without missed spots
    If UBound(tgMdShowInfo) - 1 > 1 Then
        'ArraySortTyp fnAV(tgMdShowInfo(), 1), UBound(tgMdShowInfo) - 1, 0, LenB(tgMdShowInfo(1)), 0, LenB(tgMdShowInfo(1).sKey), 0
        For ilLoop = LBound(tgMdShowInfo) To UBound(tgMdShowInfo) - 1 Step 1
            tgMdShowInfo(ilLoop) = tgMdShowInfo(ilLoop + 1)
        Next ilLoop
        ReDim Preserve tgMdShowInfo(0 To UBound(tgMdShowInfo) - 1) As MDSHOWINFO
        ArraySortTyp fnAV(tgMdShowInfo(), 0), UBound(tgMdShowInfo), 0, LenB(tgMdShowInfo(0)), 0, LenB(tgMdShowInfo(0).sKey), 0
        ReDim Preserve tgMdShowInfo(0 To UBound(tgMdShowInfo) + 1) As MDSHOWINFO
        For ilLoop = UBound(tgMdShowInfo) - 1 To LBound(tgMdShowInfo) Step -1
            tgMdShowInfo(ilLoop + 1) = tgMdShowInfo(ilLoop)
        Next ilLoop
       
    End If
    vbcMissed.Min = LBONE   'LBound(tgMdShowInfo)
    If UBound(tgMdShowInfo) <= vbcMissed.LargeChange Then
    ' If this is used, there are probably 0 or 1 records
        vbcMissed.Max = LBONE   'LBound(tgMdShowInfo)
    Else
    ' Saves, what amounts to, the count of records just retrieved
        vbcMissed.Max = UBound(tgMdShowInfo) - vbcMissed.LargeChange
    End If
    vbcMissed.Value = vbcMissed.Min
    If ilResetToValue Then
        vbcMissed.Value = ilValue
    End If
    pbcMissed_Paint
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateSaveSpotImage            *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set save values for a spot     *
'*                                                     *
'*******************************************************
Private Function mCreateSaveSpotImage(ilUpperBound As Integer, slAvDate As String, slAvTime As String, ilAnfCode As Integer, ilCount As Integer) As Integer
    Dim ilRet As Integer    'Return status
    Dim llRecPos As Long
    Dim ilCifFound As Integer
    Dim slHoldNames As String
    Dim slActPrice As String
    Dim slStr As String
    Dim llTime As Long
    Dim llTstTime As Long
    Dim ilSimulAirCode As Integer
    Dim ilLoop As Integer
    Dim ilDay As Integer
    ' Get the physical 4-byte position, in the file, of this record
    ilRet = btrGetPosition(hmSdf, llRecPos)
    tgSave(ilUpperBound).lSdfRecPos = llRecPos
    tgSave(ilUpperBound).iCount = ilCount
    tgSave(ilUpperBound).iType = 1
    If tmSdf.lChfCode > 0 Then
        tmChfSrchKey.lCode = tmSdf.lChfCode  ' Contract Hdr File Code
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Else
        tmFSFSrchKey.lCode = tmSdf.lFsfCode
        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        gMoveFeedToCntr tmFsf, tmLnRdf, tmChf, tmClf, tmFCff(), hmFnf, hmPrf
        tmCff = tmFCff(1)
    End If
    'Date
    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), tgSave(ilUpperBound).sAirDate   'smSave(SAVDATEINDEX, ilUpperBound)
    tgSave(ilUpperBound).sSchDate = slAvDate
    tgSave(ilUpperBound).sXMid = tmSdf.sXCrossMidnight
    'Time
    gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", tgSave(ilUpperBound).sAirTime 'smSave(SAVTIMEINDEX, ilUpperBound)
    tgSave(ilUpperBound).sSchTime = slAvTime
    ' Length
    tgSave(ilUpperBound).iLen = tmSdf.iLen
    If tmSdf.sBill = "Y" Then
        'imPostSpotInfo(3, ilUpperBound) = True   'Billed
        tgSave(ilUpperBound).iBilled = True
    Else
        'imPostSpotInfo(3, ilUpperBound) = False  'Not billed
        tgSave(ilUpperBound).iBilled = False
    End If
    tmAdfSrchKey.iCode = tmSdf.iAdfCode
    ' Read Advertiser file to get the Advertiser Name
    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mCreateSaveSpotImageErr
    gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:Advertiser)", PostLog
    On Error GoTo 0
    If tmAdf.sShowISCI = "Y" Or tmAdf.sShowISCI = "W" Then          '3-10-15  show isci (W) on invoice without leader hard-coded prefix of WW_
        tgSave(ilUpperBound).iISCIReq = True
    Else
        tgSave(ilUpperBound).iISCIReq = False
    End If
    tgSave(ilUpperBound).iISCI = True
    tgSave(ilUpperBound).iSimulCast = False
    If rbcType(0).Value Then
        If (tmVef.sType = "S") Then
            If (tgVpf(imVpfIndex).sBillSA = "Y") Then
                ilSimulAirCode = -1
                llTime = gTimeToCurrency(slAvTime, False)
                ilDay = gWeekDayStr(slAvDate)
                If ilDay < 5 Then
                    For ilLoop = LBound(tmVlf0) To UBound(tmVlf0) - 1 Step 1
                        gUnpackTimeLong tmVlf0(ilLoop).iSellTime(0), tmVlf0(ilLoop).iSellTime(1), False, llTstTime
                        If llTstTime = llTime Then
                            If ilSimulAirCode <= 0 Then
                                ilSimulAirCode = tmVlf0(ilLoop).iAirCode
                            Else
                                If ilSimulAirCode <> tmVlf0(ilLoop).iAirCode Then
                                    tgSave(ilUpperBound).iSimulCast = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilLoop
                ElseIf ilDay = 5 Then
                    For ilLoop = LBound(tmVlf6) To UBound(tmVlf6) - 1 Step 1
                        gUnpackTimeLong tmVlf6(ilLoop).iSellTime(0), tmVlf6(ilLoop).iSellTime(1), False, llTstTime
                        If llTstTime = llTime Then
                            If ilSimulAirCode <= 0 Then
                                ilSimulAirCode = tmVlf6(ilLoop).iAirCode
                            Else
                                If ilSimulAirCode <> tmVlf6(ilLoop).iAirCode Then
                                    tgSave(ilUpperBound).iSimulCast = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilLoop
                Else
                    For ilLoop = LBound(tmVlf7) To UBound(tmVlf7) - 1 Step 1
                        gUnpackTimeLong tmVlf7(ilLoop).iSellTime(0), tmVlf7(ilLoop).iSellTime(1), False, llTstTime
                        If llTstTime = llTime Then
                            If ilSimulAirCode <= 0 Then
                                ilSimulAirCode = tmVlf7(ilLoop).iAirCode
                            Else
                                If ilSimulAirCode <> tmVlf7(ilLoop).iAirCode Then
                                    tgSave(ilUpperBound).iSimulCast = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilLoop
                End If
            End If
        End If
    End If
    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
        slStr = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "/Direct"
    Else
        slStr = Trim$(tmAdf.sName)
    End If
    tgSave(ilUpperBound).sAdvtName = slStr  'Trim$(tmAdf.sName)
    ' Read CHF (Contract File) to get ProductName from sProduct
    If tmSdf.lChfCode > 0 Then
        tmChfSrchKey.lCode = tmSdf.lChfCode  ' Contract Hdr File Code
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Else
        tmFSFSrchKey.lCode = tmSdf.lFsfCode
        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        gMoveFeedToCntr tmFsf, tmLnRdf, tmChf, tmClf, tmFCff(), hmFnf, hmPrf
        tmCff = tmFCff(1)
    End If
    On Error GoTo mCreateSaveSpotImageErr
    gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:Contract)", PostLog
    On Error GoTo 0
    'If Not imPostSpotInfo(1, ilUpperBound) Then
    If Not tgSave(ilUpperBound).iISCIReq Then
        If tmChf.iAgfCode > 0 Then
            tmAgfSrchKey.iCode = tmChf.iAgfCode
            ' Read Advertiser file to get the Advertiser Name
            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mCreateSaveSpotImageErr
            gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:Agency)", PostLog
            On Error GoTo 0
            If tmAgf.sShowISCI = "Y" Or tmAgf.sShowISCI = "W" Then      '3-10-15  show isci (W) on invoice without leader hard-coded prefix of WW_
                'imPostSpotInfo(1, ilUpperBound) = True
                tgSave(ilUpperBound).iISCIReq = True
            Else
                'imPostSpotInfo(1, ilUpperBound) = False  'Not required is
                tgSave(ilUpperBound).iISCIReq = False
            End If
        End If
    End If
    tgSave(ilUpperBound).sProd = Trim$(tmChf.sProduct)
    ' Get PRODUCT, COPY, ISCI Code and TZ value
    tgSave(ilUpperBound).sCopy = ""
    tgSave(ilUpperBound).sISCI = ""
    tgSave(ilUpperBound).sCopyProduct = ""
    tgSave(ilUpperBound).sTZone = ""
'    ilCifFound = False
'    If tmSdf.sPtType = "1" Then  '  Single Copy
'        ' Read CIF using lCopyCode from SDF
'        tmCifSrchKey.lCode = tmSdf.lCopyCode
'        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        On Error GoTo mCreateSaveSpotImageErr
'        gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:CIF, Single)", PostLog
'        On Error GoTo 0
'        ilCifFound = True
'    ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
'    ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
'        tgSave(ilUpperBound).sTZone = "4"
'        ' Read TZF using lCopyCode from SDF
'        tmTzfSrchKey.lCode = tmSdf.lCopyCode
'        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        On Error GoTo mCreateSaveSpotImageErr
'        gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:TZF)", PostLog
'        On Error GoTo 0
'        ' Look for the first positive lZone value
'        For imIndex = 1 To 6 Step 1
'            If (tmTzf.lCifZone(imIndex) > 0) And (StrComp(tmTzf.sZone(imIndex), "Oth", 1) = 0) Then ' Process just the first positive Zone
'                ' Read CIF using lCopyCode from SDF
'                tmCifSrchKey.lCode = tmTzf.lCifZone(imIndex)
'                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                On Error GoTo mCreateSaveSpotImageErr
'                gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:CIF, Time zone)", PostLog
'                On Error GoTo 0
'                ilCifFound = True
'                Exit For
'            End If
'        Next imIndex
'        If Not ilCifFound Then
'            For imIndex = 1 To 6 Step 1
'                If tmTzf.lCifZone(imIndex) > 0 Then ' Process just the first positive Zone
'                    ' Read CIF using lCopyCode from SDF
'                    tmCifSrchKey.lCode = tmTzf.lCifZone(imIndex)
'                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                    On Error GoTo mCreateSaveSpotImageErr
'                    gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:CIF, Time zone)", PostLog
'                    On Error GoTo 0
'                    ilCifFound = True
'                    Exit For
'                End If
'            Next imIndex
'        End If
'    End If
'    If ilCifFound Then
'        ' Read CPF using lCpfCode from CIF
'        If tmCif.lCpfCode > 0 Then
'            tmCpfSrchKey.lCode = tmCif.lCpfCode
'            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            On Error GoTo mCreateSaveSpotImageErr
'            gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:CPF)", PostLog
'            On Error GoTo 0
'            'smSave(SAVISCIINDEX, ilUpperBound) = Trim$(tmCpf.sISCI)  ' ISCI Code
'            tgSave(ilUpperBound).sISCI = Trim$(tmCpf.sISCI)
'        Else
'            tmCpf.sISCI = ""
'            'smSave(SAVISCIINDEX, ilUpperBound) = Trim$(tmCpf.sISCI)  ' ISCI Code
'            tgSave(ilUpperBound).sISCI = Trim$(tmCpf.sISCI)
'            tmCpf.sName = ""
'        End If
'        ' Concatinate Copy from Media Code, Inv. Name & Cut#
'        ' First read MCF
'        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
'            tmMcfSrchKey.iCode = tmCif.iMcfCode
'            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'            On Error GoTo mCreateSaveSpotImageErr
'            gBtrvErrorMsg ilRet, "mCreateSaveSpotImage (btrGetEqual:MCF)", PostLog
'            On Error GoTo 0
'            ' Media Code is tmMcf.sName
'            slHoldNames = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
'            If (Len(Trim$(tmCif.sCut)) <> 0) Then
'                slHoldNames = slHoldNames & "-" & tmCif.sCut
'            End If
'            tgSave(ilUpperBound).sCopy = slHoldNames
'        Else
'            tgSave(ilUpperBound).sCopy = ""
'        End If
'        If Trim$(tmCpf.sName) <> "" Then
'            tgSave(ilUpperBound).sCopyProduct = Trim$(tmCpf.sName)
'        End If
'        If Trim$(tmCpf.sISCI) <> "" Then
'            'imPostSpotInfo(2, ilUpperBound) = False
'            tgSave(ilUpperBound).iISCI = False
'        Else
'            'imPostSpotInfo(2, ilUpperBound) = True
'            tgSave(ilUpperBound).iISCI = True
'        End If
'    End If
    If Not mGetCopy(tmSdf, ilUpperBound) Then
        mCreateSaveSpotImage = False
        Exit Function
    End If
    ' Contract #
    tgSave(ilUpperBound).lCntrNo = tmChf.lCntrNo
    ' Line Number
    tgSave(ilUpperBound).iLineNo = tmSdf.iLineNo
    ' T (for Type)
    If tmSdf.sSpotType = "A" Then ' A will be replaced by blank
        tmSdf.sSpotType = " "
    End If
    tgSave(ilUpperBound).sSpotType = tmSdf.sSpotType
    If (tmChf.sType = "M") Or (tmChf.sType = "S") Or (tmChf.sType = "T") Then
        'If tmSdf.sSpotType <> "X" Then
            tgSave(ilUpperBound).sSpotType = tmChf.sType
        'End If
    End If
    tgSave(ilUpperBound).iPrice = -1
    If tmSdf.lChfCode > 0 Then
        If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Or (tmSdf.sSpotType = "E") Then
            If (tmSdf.sSpotType = "O") Then
                slActPrice = "Open BB"
            ElseIf (tmSdf.sSpotType = "C") Then
                slActPrice = "Close BB"
            Else
                slActPrice = "Any BB"
            End If
        ElseIf tmSdf.sSpotType = "X" Then
            'If tmSdf.sPriceType <> "N" Then
            If tmSdf.sPriceType = "+" Then
                slActPrice = "+ Fill"   '"> Fill"
                tgSave(ilUpperBound).iPrice = 1
            ElseIf tmSdf.sPriceType = "-" Then
                slActPrice = "- Fill"   '"< Fill"
                tgSave(ilUpperBound).iPrice = 2
            Else
                tgSave(ilUpperBound).iPrice = 0
                If tmAdf.sBonusOnInv <> "N" Then
                    slActPrice = "+ Fill"
                Else
                    slActPrice = "- Fill"
                End If
            End If
        Else
            tmClfSrchKey.lChfCode = tmSdf.lChfCode
            tmClfSrchKey.iLine = tmSdf.iLineNo
            tmClfSrchKey.iCntRevNo = 32000 '0 ' Plug with very hi number
            tmClfSrchKey.iPropVer = 32000 '0 ' Plug with very hi number
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slActPrice)
    '            If InStr(slActPrice, ".") > 0 Then
    '                If (tmSdf.sPriceType = "N") Or (tmSdf.sPriceType = "P") Then
    '                    'imSave(1, ilUpperBound) = 1
    '                    tgSave(ilUpperBound).iPrice = 1
    '                Else
    '                    'imSave(1, ilUpperBound) = 0
    '                    tgSave(ilUpperBound).iPrice = 0
    '                End If
    '            End If
                'Select Case tmClf.sPriceType
                '    Case "T"    'True
                '        If (tmSdf.sPriceType = "N") Or (tmSdf.sPriceType = "P") Then
                '            imSave(1, ilUpperBound) = 1
                '        Else
                '            imSave(1, ilUpperBound) = 0
                '        End If
                '        gPDNToStr tmClf.sActPrice, 2, slActPrice
                '    Case "N"    'No Charge
                '        imSave(1, ilUpperBound) = -1
                '        slActPrice = "N/C"
                '    Case "M"    'MG Line
                '        imSave(1, ilUpperBound) = -1
                '        slActPrice = "MG"
                '    Case "B"    'Bonus
                '        imSave(1, ilUpperBound) = -1
                '        slActPrice = "Bonus"
                '    Case "S"    'Spinoff
                '        imSave(1, ilUpperBound) = -1
                '        slActPrice = "Spinoff"
                '    Case "P"    'Package
                '        imSave(1, ilUpperBound) = -1
                 '   Case "R"    'Recapturable
                '        imSave(1, ilUpperBound) = -1
                '        slActPrice = "Recapturable"
                '    Case "A"    'ADU
                '        imSave(1, ilUpperBound) = -1
                '        slActPrice = "ADU"
                'End Select
            Else
                slActPrice = " "
                'imSave(1, ilUpperBound) = -1
                tgSave(ilUpperBound).iPrice = -1
            End If
        End If
    Else
        slActPrice = "Feed"
    End If
    tgSave(ilUpperBound).iSvPrice = tgSave(ilUpperBound).iPrice
    tgSave(ilUpperBound).sPrice = slActPrice
    ' MakeGood
    If tmSdf.sSchStatus = "G" Then
        slStr = "G" '"M" the M is too large - it didn't display
    ElseIf tmSdf.sSchStatus = "O" Then
        slStr = "O"
    Else
        slStr = " "
    End If
    tgSave(ilUpperBound).sSchStatus = slStr
    ' Change Code
    If (tmSdf.sAffChg = " ") Or (tmSdf.sAffChg = "N") Then
        slStr = ""
    Else
        slStr = "Y"
    End If
    tgSave(ilUpperBound).sAffChg = slStr
    tgSave(ilUpperBound).ianfCode = ilAnfCode
    tgSave(ilUpperBound).iShowInfoIndex = 0
    mCreateSaveSpotImage = True
    Exit Function
mCreateSaveSpotImageErr:
    On Error GoTo 0
    mCreateSaveSpotImage = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateShowImage                *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set save values for a spot     *
'*                                                     *
'*******************************************************
Private Sub mCreateShowImage(ilIndex As Integer, ilUpperBound As Integer)
    Dim slActPrice As String
    Dim flWidth As Single
    Dim slStr As String
    Dim slProduct As String

    If tgSave(ilIndex).iType = 0 Then
        Exit Sub
    End If
    tgSave(ilIndex).iShowInfoIndex = ilUpperBound
    tgShow(ilUpperBound).iSaveInfoIndex = ilIndex
    'Date
    gSetShow pbcPosting, Trim$(tgSave(ilIndex).sAirDate), tmCtrls(DATEINDEX)
    tgShow(ilUpperBound).sShow(DATEINDEX) = tmCtrls(DATEINDEX).sShow
    'Time
    gSetShow pbcPosting, Trim$(tgSave(ilIndex).sAirTime), tmCtrls(TIMEINDEX)
    tgShow(ilUpperBound).sShow(TIMEINDEX) = tmCtrls(TIMEINDEX).sShow
    ' Length
    tmCtrls(LENINDEX).sShow = Trim$(str$(tgSave(ilIndex).iLen))
    tgShow(ilUpperBound).sShow(LENINDEX) = tmCtrls(LENINDEX).sShow
    'Zone
    tgShow(ilUpperBound).sShow(TZONEINDEX) = tgSave(ilIndex).sTZone
    'Copy
    tgShow(ilUpperBound).sShow(COPYINDEX) = ""
    tgShow(ilUpperBound).sShow(ISCIINDEX) = ""
    If tgSpf.sUseCartNo <> "N" Then
        gSetShow pbcPosting, Trim$(tgSave(ilIndex).sCopy), tmCtrls(COPYINDEX)
        tgShow(ilUpperBound).sShow(COPYINDEX) = tmCtrls(COPYINDEX).sShow
    End If
    gSetShow pbcPosting, Trim$(tgSave(ilIndex).sISCI), tmCtrls(ISCIINDEX)
    tgShow(ilUpperBound).sShow(ISCIINDEX) = tmCtrls(ISCIINDEX).sShow
    'Advertiser Name Plus Contract Product
    tgShow(ilUpperBound).sShow(ADVTINDEX) = ""
    tgShow(ilUpperBound).sShow(TZONEINDEX) = " "
    slProduct = Trim$(tgSave(ilIndex).sCopyProduct)
    If Len(slProduct) <= 0 Then
        slProduct = Trim$(tgSave(ilIndex).sProd)
    End If
    If Len(slProduct) = 0 Then  ' No Product Name
        gSetShow pbcPosting, Trim$(tgSave(ilIndex).sAdvtName), tmCtrls(ADVTINDEX)
        tgShow(ilUpperBound).sShow(ADVTINDEX) = tmCtrls(ADVTINDEX).sShow ' So can be shown
    Else
        ' Get Advertiser field width, divide by 2 (so can append Product)
        flWidth = tmCtrls(ADVTINDEX).fBoxW ' First save copy of real width
        tmCtrls(ADVTINDEX).fBoxW = tmCtrls(ADVTINDEX).fBoxW / 2
        ' Put no more than 1/2 advertiser name into control array
        'gSetShow pbcPosting, smSave(SAVADVTNAMEINDEX, ilIndex), tmCtrls(ADVTINDEX)
        gSetShow pbcPosting, Trim$(tgSave(ilIndex).sAdvtName), tmCtrls(ADVTINDEX)
        ' concantinate  "\ProductName" to it
        slStr = tmCtrls(ADVTINDEX).sShow & "/" & slProduct
        ' Trim the concatinated names to fit
        tmCtrls(ADVTINDEX).fBoxW = flWidth  ' Restore width
        gSetShow pbcPosting, slStr, tmCtrls(ADVTINDEX)
        ' Save concatinated result in the ShowArray
        tgShow(ilUpperBound).sShow(ADVTINDEX) = tmCtrls(ADVTINDEX).sShow
    End If
    ' Contract #
    gSetShow pbcPosting, Trim$(str$(tgSave(ilIndex).lCntrNo)), tmCtrls(CNTRINDEX)
    tgShow(ilUpperBound).sShow(CNTRINDEX) = tmCtrls(CNTRINDEX).sShow
    ' Line Number
    gSetShow pbcPosting, Trim$(str$(tgSave(ilIndex).iLineNo)), tmCtrls(LINEINDEX)
    tgShow(ilUpperBound).sShow(LINEINDEX) = tmCtrls(LINEINDEX).sShow
    ' T (for Type)
    gSetShow pbcPosting, tgSave(ilIndex).sSpotType, tmCtrls(TYPEINDEX)
    tgShow(ilUpperBound).sShow(TYPEINDEX) = tmCtrls(TYPEINDEX).sShow
    'Price
    slActPrice = Trim$(tgSave(ilIndex).sPrice)
'    If tgSave(ilIndex).iPrice = 0 Then 'mSave(1, ilIndex) = 0 Then
'        gFormatStr slActPrice, FMTLEAVEBLANK + FMTCOMMA, 2, slActPrice
'    ElseIf tgSave(ilIndex).iPrice = 1 Then 'imSave(1, ilIndex) = 1 Then
'        slActPrice = "N/C"
'    End If
    If InStr(slActPrice, ".") > 0 Then
        gFormatStr slActPrice, FMTLEAVEBLANK + FMTCOMMA, 2, slActPrice
    End If
    gSetShow pbcPosting, slActPrice, tmCtrls(PRICEINDEX) ' Shorten it
    tgShow(ilUpperBound).sShow(PRICEINDEX) = tmCtrls(PRICEINDEX).sShow
    ' MakeGood
    slStr = Trim$(tgSave(ilIndex).sSchStatus)
    gSetShow pbcPosting, slStr, tmCtrls(MGOODINDEX) ' Shorten it
    tgShow(ilUpperBound).sShow(MGOODINDEX) = tmCtrls(MGOODINDEX).sShow
    ' Change Code
    slStr = Trim$(tgSave(ilIndex).sAffChg)
    gSetShow pbcPosting, slStr, tmCtrls(AUDINDEX)
    tgShow(ilUpperBound).sShow(AUDINDEX) = tmCtrls(AUDINDEX).sShow
    tgShow(ilUpperBound).iChk = False
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDatePop                        *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the date combo box    *
'*                      First date for vehicle in SDF  *
'*                      to last log date               *
'*                                                     *
'*******************************************************
Private Sub mDatePop()

    Dim slDate As String
    Dim llSchDate As Long
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim llEndPostDate As Long
    Dim llNowDate As Long
    Dim ilUpper As Integer
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilGsf As Integer

    imDefaultDateIndex = -1
    'cbcDate.Clear
    ReDim tgDates(0 To 0) As DATES
    ReDim tgWkDates(0 To 0) As DATES
    ilUpper = 0
    ' Exit if last log date is zero. LLVehLLD is obtained from cbcVeh
    'If lmVehLLD = 0 Then ' lmVehLLD is LastLog Date and has already been obtained
    '    Exit Sub
    'End If
    llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    If (lmVehLLD <= llNowDate) And (lmVehLLD > 0) Then
        llEndPostDate = lmVehLLD
    Else
        'If tgSpf.sSMove = "Y" Then   'Move spots between todays Plus 1 And Last Log
        '    llEndPostDate = lmVehLLD    'llNowDate
        'Else
            llEndPostDate = llNowDate   'lmVehLLD
        'End If
    End If
    '11/19/11: Disallow posting on today
    If llEndPostDate = llNowDate Then
        '1/25/13:  Only disallow posting on todays date if live log (TTP 5904)
        If tmVef.sType <> "V" Then
            If (tgVpf(imVpfIndex).sGenLog = "N") Or (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) = USINGLIVELOG) And ((tgVpf(imVpfIndex).sGenLog = "L") Or (tgVpf(imVpfIndex).sGenLog = "A"))) Then
                llEndPostDate = llEndPostDate - 1
            End If
        End If
    End If
    If tmVef.sType <> "G" Then
        pbcWM.Visible = True
        For ilDay = 1 To 6 Step 1
            ckcDayComplete(ilDay).Visible = True
        Next ilDay
        ckcDayComplete(0).Caption = "Mo"
        ckcDayComplete(0).Width = ckcDayComplete(1).Width
        If rbcType(0).Value Then
        '   Obtain the earliest date posted. Comes from the LCF file
            tmLcfSrchKey.iType = 0
            tmLcfSrchKey.sStatus = "C"  ' Current
            tmLcfSrchKey.iVefCode = imVefCode
            tmLcfSrchKey.iLogDate(0) = 257
            tmLcfSrchKey.iLogDate(1) = 1900
            tmLcfSrchKey.iSeqNo = 0
            ' Read the first matching record into tmLcf record structure. If a record exists, put
            '      it into slDate in the form of mm/dd/yy then
            '      it is converted into a number and put into llSchSDate
            ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.sStatus = "C") And (tmLcf.iType = 0) Then
                gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
                llSchDate = gDateValue(slDate)
            Else
                edcDate_Change
                Exit Sub
            End If
            For llLoop = llSchDate To llEndPostDate Step 1
                tmLcfSrchKey.iType = 0
                tmLcfSrchKey.sStatus = "C"  ' Current
                tmLcfSrchKey.iVefCode = imVefCode
                slDate = Format$(llLoop, "m/d/yy")
                gPackDate slDate, ilDate0, ilDate1
                tmLcfSrchKey.iLogDate(0) = ilDate0
                tmLcfSrchKey.iLogDate(1) = ilDate1
                tmLcfSrchKey.iSeqNo = 1
                ' Get a record from LCF
                ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) Then
                    'slName = gFormatDate(slDate)
                    'slName = gAddDayToDate(slName)
                    'If tmLcf.sAffPost = "C" Then
                    '   slName = slName & ":Completed"
                    'ElseIf tmLcf.sAffPost = "I" Then
                    '   slName = slName & ":Incomplete"
                    '   If imDefaultDateIndex = -1 Then
                    '      imDefaultDateIndex = cbcDate.ListCount
                    '   End If
                    'Else
                    '   If imDefaultDateIndex = -1 Then
                    '      imDefaultDateIndex = cbcDate.ListCount
                    '   End If
                    'End If
                    'cbcDate.AddItem slName
                    tgDates(ilUpper).sDate = slDate
                    tgDates(ilUpper).lDate = gDateValue(slDate)
                    tgDates(ilUpper).iGameNo = 0
                    If tmLcf.sAffPost = "C" Then
                       tgDates(ilUpper).iStatus = 2
                    ElseIf tmLcf.sAffPost = "I" Then
                       tgDates(ilUpper).iStatus = 1
                       If imDefaultDateIndex = -1 Then
                          imDefaultDateIndex = ilUpper
                       End If
                    Else
                        tgDates(ilUpper).iStatus = 0
                       If imDefaultDateIndex = -1 Then
                          imDefaultDateIndex = ilUpper
                       End If
                    End If
                    ilUpper = ilUpper + 1
                    ReDim Preserve tgDates(0 To ilUpper) As DATES
                End If
            Next llLoop
        Else
            tmSdfSrchKey.iVefCode = imVefCode
            tmSdfSrchKey.iDate(0) = 0
            tmSdfSrchKey.iDate(1) = 0
            tmSdfSrchKey.iTime(0) = 0
            tmSdfSrchKey.iTime(1) = 0
            tmSdfSrchKey.sSchStatus = "S"
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) And (imVefCode = tmSdf.iVefCode) Then
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                llSchDate = gDateValue(slDate)
            Else
                edcDate_Change
                Exit Sub
            End If
            Do While llSchDate <= llEndPostDate
                tmSdfSrchKey.iVefCode = imVefCode
                slDate = Format$(llSchDate, "m/d/yy")
                gPackDate slDate, tmSdfSrchKey.iDate(0), tmSdfSrchKey.iDate(1)
                ilDate0 = tmSdfSrchKey.iDate(0)
                ilDate1 = tmSdfSrchKey.iDate(1)
                tmSdfSrchKey.iTime(0) = 0
                tmSdfSrchKey.iTime(1) = 0
                tmSdfSrchKey.sSchStatus = "S"
                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) And (imVefCode = tmSdf.iVefCode) And (tmSdf.iDate(0) = ilDate0) And (tmSdf.iDate(1) = ilDate1) Then
                    tgDates(ilUpper).sDate = slDate
                    tgDates(ilUpper).lDate = gDateValue(slDate)
                    tgDates(ilUpper).iStatus = 0
                    tgDates(ilUpper).iGameNo = 0
                    If imDefaultDateIndex = -1 Then
                        imDefaultDateIndex = ilUpper
                    End If
                    ilUpper = ilUpper + 1
                    ReDim Preserve tgDates(0 To ilUpper) As DATES
                    llSchDate = llSchDate + 1
                ElseIf (ilRet = BTRV_ERR_NONE) And (imVefCode = tmSdf.iVefCode) Then
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSchDate
                Else
                    Exit Do
                End If
            Loop
        End If
        'If (imDefaultDateIndex = -1) Or (tgSpf.sBActDayCompl = "N") Then
        'If (imDefaultDateIndex = -1) Or ((tgSpf.sBActDayCompl = "N") And (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) <> USINGLIVELOG) Or (tgVpf(imVpfIndex).sGenLog <> "L"))) Then
        If (imDefaultDateIndex = -1) Or ((tgSpf.sBActDayCompl = "N") And (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) <> USINGLIVELOG) Or ((tgVpf(imVpfIndex).sGenLog <> "L") And (tgVpf(imVpfIndex).sGenLog <> "A")))) Then
            imDefaultDateIndex = ilUpper - 1    'cbcDate.ListCount - 1
        End If
        'cbcDate.ListIndex = imDefaultDateIndex
        imDateComboBoxIndex = imDefaultDateIndex
        If imDefaultDateIndex < LBound(tgDates) Then
            slStr = Format$(gNow(), "m/d/yy")  'Jim- show current moonth even if no dates exist 12/29/01
        Else
            slStr = tgDates(imDefaultDateIndex).sDate 'Format$(gNow(), "m/d/yy")
        End If
        'If edcDate.Text = slStr Then
        '    edcDate_Change
        'Else
        '    edcDate.Text = slStr
        'End If
        'lmNowDate = gDateValue(slStr)
        edcDate.Text = ""
        slStr = gObtainPrevMonday(slStr)
        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
        pbcCalendar_Paint   'mBoxCalDate called within paint
        lacDate.Visible = False
    Else
        lbcGameNo.Clear
        pbcWM.Visible = False
        For ilDay = 1 To 6 Step 1
            ckcDayComplete(ilDay).Visible = False
        Next ilDay
        ckcDayComplete(0).Caption = ""
        ckcDayComplete(0).Width = 5 * ckcDayComplete(1).Width
        ilRet = gGetGameDates(hmLcf, hmGhf, hmGsf, tmVef.iCode, tmTeam(), tmGsfInfo())
        For ilGsf = LBound(tmGsfInfo) To UBound(tmGsfInfo) - 1 Step 1
            If tmGsfInfo(ilGsf).lGameDate <= llEndPostDate Then
                '2/15/18: Verify log date exist
                tmLcfSrchKey.iType = tmGsfInfo(ilGsf).iGameNo
                tmLcfSrchKey.sStatus = "C"  ' Current
                tmLcfSrchKey.iVefCode = imVefCode
                gPackDateLong tmGsfInfo(ilGsf).lGameDate, ilDate0, ilDate1
                tmLcfSrchKey.iLogDate(0) = ilDate0
                tmLcfSrchKey.iLogDate(1) = ilDate1
                tmLcfSrchKey.iSeqNo = 1
                ' Get a record from LCF
                ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) Then
                    slStr = Trim$(str$(tmGsfInfo(ilGsf).iGameNo))
                    slStr = slStr & " " & Format$(tmGsfInfo(ilGsf).lGameDate, "m/d/yy")
                    slStr = slStr & " " & Trim$(Left(tmGsfInfo(ilGsf).sVisitName, 4)) & "@" & Trim$(Left(tmGsfInfo(ilGsf).sHomeName, 4))
                    lbcGameNo.AddItem slStr
                    lbcGameNo.ItemData(lbcGameNo.NewIndex) = ilGsf
                    ilUpper = UBound(tgDates)
                    tgDates(ilUpper).sDate = Format$(tmGsfInfo(ilGsf).lGameDate, "m/d/yy")
                    tgDates(ilUpper).lDate = tmGsfInfo(ilGsf).lGameDate
                    tgDates(ilUpper).iGameNo = tmGsfInfo(ilGsf).iGameNo
                    tgDates(ilUpper).sLiveLogMerge = tmGsfInfo(ilGsf).sLiveLogMerge
                    If tmGsfInfo(ilGsf).sAffPost = "C" Then
                       tgDates(ilUpper).iStatus = 2
                    ElseIf tmGsfInfo(ilGsf).sAffPost = "I" Then
                       tgDates(ilUpper).iStatus = 1
                       If imDefaultDateIndex = -1 Then
                          imDefaultDateIndex = ilUpper
                       End If
                    Else
                        tgDates(ilUpper).iStatus = 0
                        If imDefaultDateIndex = -1 Then
                          imDefaultDateIndex = ilUpper
                        End If
                    End If
                    ilUpper = ilUpper + 1
                    ReDim Preserve tgDates(0 To ilUpper) As DATES
                End If
            End If
        Next ilGsf
        'If (imDefaultDateIndex = -1) Or ((tgSpf.sBActDayCompl = "N") And (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) <> USINGLIVELOG) Or (tgVpf(imVpfIndex).sGenLog <> "L"))) Then
        If (imDefaultDateIndex = -1) Or ((tgSpf.sBActDayCompl = "N") And (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) <> USINGLIVELOG) Or ((tgVpf(imVpfIndex).sGenLog <> "L") And (tgVpf(imVpfIndex).sGenLog <> "A")))) Then
            imDefaultDateIndex = ilUpper - 1    'cbcDate.ListCount - 1
        End If
        lbcGameNo.height = gListBoxHeight(lbcGameNo.ListCount, 13)
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDTEnableBox                    *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mDTEnableBox(ilBoxNo As Integer)
'
'   mDTEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    Dim ilGameNo As Integer

    If (ilBoxNo < imLBDTCtrls) Or (ilBoxNo > UBound(tmDTCtrls)) Then
        Exit Sub ' Bogus box number so get out
    End If
    ilGameNo = tgDates(imDateSelectedIndex).iGameNo
    plcDT.Visible = True
    pbcDT.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case DTDATEINDEX
            edcDTDropDown.Width = tmDTCtrls(DTDATEINDEX).fBoxW - cmcDTDropDown.Width
            edcDTDropDown.MaxLength = 10
            gMoveFormCtrl pbcDT, edcDTDropDown, tmDTCtrls(DTDATEINDEX).fBoxX, tmDTCtrls(DTDATEINDEX).fBoxY
            cmcDTDropDown.Move edcDTDropDown.Left + edcDTDropDown.Width, edcDTDropDown.Top
            'If edcDTDropDown.Top + edcDTDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                imDateBox = -1
                plcCalendar.Move edcDTDropDown.Left, edcDTDropDown.Top + edcDTDropDown.height
            'Else
            '    plcCalendar.Move edcDTDropDown.Left, edcDTDropDown.Top - plcCalendar.Height
            'End If
            slStr = smDTSave(DTDATEINDEX)
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDTDropDown.Text = slStr
            edcDTDropDown.Enabled = True
            edcDTDropDown.SelStart = 0
            edcDTDropDown.SelLength = Len(edcDTDropDown.Text)
            edcDTDropDown.Visible = True
            cmcDTDropDown.Visible = True
            edcDTDropDown.SetFocus

        'Case imDTAIRTIMEINDEX 'Avail Time or Package Air Time
        Case imDTAVAILTIMEINDEX 'Avail Time
            '6/16/11
            'If (rbcType(0).Value) And (smSpotType <> "O") And (smSpotType <> "C") Then
                mAvailTimePop smDTSave(DTDATEINDEX), ilGameNo
            'End If
            edcDTDropDown.Width = tmDTCtrls(imDTAVAILTIMEINDEX).fBoxW '- cmcDTDropDown.Width
            edcDTDropDown.MaxLength = 10
            gMoveFormCtrl pbcDT, edcDTDropDown, tmDTCtrls(imDTAVAILTIMEINDEX).fBoxX, tmDTCtrls(imDTAVAILTIMEINDEX).fBoxY
            cmcDTDropDown.Move edcDTDropDown.Left + edcDTDropDown.Width, edcDTDropDown.Top
            '6/16/11
            'If (rbcType(0).Value) And (smSpotType <> "O") And (smSpotType <> "C") Then
                lbcAvailTimes.height = gListBoxHeight(lbcAvailTimes.ListCount, 5)
                lbcAvailTimes.Move plcAvailTimes.Left, plcAvailTimes.Top + plcAvailTimes.height
                If edcDTDropDown.Top + edcDTDropDown.height + lbcAvailTimes.height < cmcDone.Top Then
                    lbcAvailTimes.Move edcDTDropDown.Left, edcDTDropDown.Top + edcDTDropDown.height
                Else
                    lbcAvailTimes.Move edcDTDropDown.Left, edcDTDropDown.Top - lbcAvailTimes.height
                End If
                gFindMatch Trim$(smDTSave(imDTAVAILTIMEINDEX)), 0, lbcAvailTimes
                imChgMode = True ' Turn on the switch
                If gLastFound(lbcAvailTimes) >= 0 Then
                    lbcAvailTimes.ListIndex = gLastFound(lbcAvailTimes)
                Else ' No data found so re-display the last good data
                    lbcAvailTimes.ListIndex = -1
                End If
                imChgMode = False
                'edcDTDropDown.Text = Trim$(smDTSave(imDTAVAILTIMEINDEX))
            'Else
            '    'If edcDTDropDown.Top + edcDTDropDown.Height + plcTme.Height < cmcDone.Top Then
            '        plcTme.Move edcDTDropDown.Left, edcDTDropDown.Top + edcDTDropDown.Height
            '    'Else
            '    '    plcTme.Move edcDTDropDown.Left, edcDTDropDown.Top - plcTme.Height
            '    'End If
            'End If
            edcDTDropDown.Text = Trim$(smDTSave(imDTAVAILTIMEINDEX))
            edcDTDropDown.Enabled = False
            edcDTDropDown.Visible = True  'Set visibility
            'cmcDTDropDown.Visible = True
            lbcAvailTimes.Visible = True
            'edcDTDropDown.SetFocus
            lbcAvailTimes.SetFocus
        Case imDTAIRTIMEINDEX 'Air Time
            edcDTDropDown.Width = tmDTCtrls(imDTAIRTIMEINDEX).fBoxW - cmcDTDropDown.Width
            edcDTDropDown.MaxLength = 10
            gMoveFormCtrl pbcDT, edcDTDropDown, tmDTCtrls(imDTAIRTIMEINDEX).fBoxX, tmDTCtrls(imDTAIRTIMEINDEX).fBoxY
            cmcDTDropDown.Move edcDTDropDown.Left + edcDTDropDown.Width, edcDTDropDown.Top
            plcTme.Move edcDTDropDown.Left, edcDTDropDown.Top + edcDTDropDown.height
            edcDTDropDown.Text = smDTSave(imDTAIRTIMEINDEX)
            edcDTDropDown.Enabled = True
            edcDTDropDown.Visible = True  'Set visibility
            cmcDTDropDown.Visible = True
            edcDTDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDTSetShow                      *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mDTSetShow(ilBoxNo As Integer)
'
'   mDTSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim ilFound As Integer
    If (ilBoxNo < imLBDTCtrls) Or (ilBoxNo > UBound(tmDTCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case DTDATEINDEX
            plcCalendar.Visible = False
            cmcDTDropDown.Visible = False
            edcDTDropDown.Visible = False  'Set visibility
            slStr = edcDTDropDown.Text
            If gValidDate(slStr) Then
                If rbcType(0).Value Then
                    llDate = gDateValue(slStr)
                    ilFound = False
                    For ilLoop = 0 To UBound(tgWkDates) - 1 Step 1
                        If tgWkDates(ilLoop).lDate = llDate Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        Beep
                        edcDTDropDown.Text = smDTSave(DTDATEINDEX)
                    Else
                        smDTSave(DTDATEINDEX) = slStr
                        slStr = gFormatDate(slStr)
                        gSetShow pbcDT, slStr, tmDTCtrls(ilBoxNo)
                    End If
                Else
                    smDTSave(DTDATEINDEX) = slStr
                    slStr = gFormatDate(slStr)
                    gSetShow pbcDT, slStr, tmDTCtrls(ilBoxNo)
                End If
            Else
                Beep
                edcDTDropDown.Text = smDTSave(DTDATEINDEX)
            End If
        Case imDTAVAILTIMEINDEX
            cmcDTDropDown.Visible = False
            lbcAvailTimes.Visible = False
            edcDTDropDown.Visible = False  'Set visibility
            slStr = edcDTDropDown.Text
            If slStr <> "" Then
                If gValidTime(slStr) Then
                    gFindMatch slStr, 0, lbcAvailTimes
                    If gLastFound(lbcAvailTimes) >= 0 Then
                        smDTSave(imDTAVAILTIMEINDEX) = slStr
                        slStr = gFormatTime(slStr, "A", "1")
                    Else
                        Beep
                        edcDTDropDown.Text = smDTSave(imDTAVAILTIMEINDEX)
                        slStr = smDTSave(imDTAVAILTIMEINDEX)
                    End If
                Else
                    Beep
                    edcDTDropDown.Text = smDTSave(imDTAVAILTIMEINDEX)
                    slStr = smDTSave(imDTAVAILTIMEINDEX)
                End If
            Else
                Beep
                edcDTDropDown.Text = smDTSave(imDTAVAILTIMEINDEX)
                slStr = smDTSave(imDTAVAILTIMEINDEX)
            End If
            gSetShow pbcDT, slStr, tmDTCtrls(ilBoxNo)
        Case imDTAIRTIMEINDEX
            cmcDTDropDown.Visible = False
            plcTme.Visible = False
            edcDTDropDown.Visible = False  'Set visibility
            slStr = edcDTDropDown.Text
            If slStr <> "" Then
                If gValidTime(slStr) Then
                    smDTSave(imDTAIRTIMEINDEX) = slStr
                    slStr = gFormatTime(slStr, "A", "1")
                Else
                    Beep
                    edcDTDropDown.Text = smDTSave(imDTAIRTIMEINDEX)
                    slStr = smDTSave(imDTAIRTIMEINDEX)
                End If
            End If
            gSetShow pbcDT, slStr, tmDTCtrls(ilBoxNo)
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilRowNo As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub ' Bogus box number so get out
    End If
    ' Get out if not a visible row
    If (imRowNo < vbcPosting.Value) Or (imRowNo > vbcPosting.Value + vbcPosting.LargeChange) Then
        pbcArrow.Visible = False
        lacPtFrame.Visible = False
        Exit Sub
    End If
    ilRowNo = tgShow(imRowNo).iSaveInfoIndex
    smSpotType = tgSave(ilRowNo).sSpotType
    ' OK so far, move the frame/arrow to the selected row
    lacPtFrame.Move 0, tmCtrls(TIMEINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15) - 30
    lacPtFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcPosting.Top + tmCtrls(TIMEINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case DATEINDEX, TIMEINDEX 'Vehicle
            'If rbcType(0).Value Then
            If (rbcType(0).Value) And (smSpotType <> "O") And (smSpotType <> "C") Then

                '6/16/11
                pbcDT.height = 1080
                plcDT.height = 1425
                imDTAIRTIMEINDEX = 3
                imDTAVAILTIMEINDEX = 2
                imDTBoxNo = -1
                smSvDate = Trim$(tgSave(ilRowNo).sAirDate)
                smSvAirTime = Trim$(tgSave(ilRowNo).sAirTime)
                '6/16/11
                smSvAvailTime = Trim$(tgSave(ilRowNo).sSchTime)
                smDTSave(DTDATEINDEX) = Trim$(tgSave(ilRowNo).sAirDate)    'smSave(SAVDATEINDEX, imRowNo)
                smDTSave(imDTAIRTIMEINDEX) = Trim$(tgSave(ilRowNo).sAirTime) 'smSave(SAVTIMEINDEX, imRowNo)
                '6/16/11
                smDTSave(imDTAVAILTIMEINDEX) = Trim$(tgSave(ilRowNo).sSchTime) 'smSave(SAVTIMEINDEX, imRowNo)
                gSetShow pbcDT, smDTSave(DTDATEINDEX), tmDTCtrls(DTDATEINDEX)
                gSetShow pbcDT, smDTSave(imDTAIRTIMEINDEX), tmDTCtrls(imDTAIRTIMEINDEX)
                '6/16/11
                gSetShow pbcDT, smDTSave(imDTAVAILTIMEINDEX), tmDTCtrls(imDTAVAILTIMEINDEX)
                'gMoveTableCtrl pbcPosting, plcDT, tmCtrls(TIMEINDEX).fBoxX + tmCtrls(TIMEINDEX).fBoxW, tmCtrls(TIMEINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15)
                gMoveTableCtrl pbcPosting, plcDT, tmCtrls(DATEINDEX).fBoxX, tmCtrls(TIMEINDEX).fBoxY + (imRowNo - vbcPosting.Value + 1) * (fgBoxGridH + 15) - plcDT.height
                pbcDT.Move plcDT.Left + 90, plcDT.Top + 270
                plcDT.Visible = True ' The main (Green) form for Time Zone information display
                pbcDT.Visible = True ' The Time Zone information form
                pbcDT_Paint ' Put data from smTZSave onto form
                pbcDTSTab.SetFocus
            Else

                '6/16/11
                pbcDT.height = 720
                plcDT.height = 1065
                imDTAIRTIMEINDEX = 2
                imDTAVAILTIMEINDEX = -1
                imDTBoxNo = -1
                smSvDate = Trim$(tgSave(ilRowNo).sAirDate)
                smSvAirTime = Trim$(tgSave(ilRowNo).sAirTime)
                smDTSave(DTDATEINDEX) = Trim$(tgSave(ilRowNo).sAirDate)    'smSave(SAVDATEINDEX, imRowNo)
                smDTSave(imDTAIRTIMEINDEX) = Trim$(tgSave(ilRowNo).sAirTime) 'smSave(SAVTIMEINDEX, imRowNo)
                gSetShow pbcDT, smDTSave(DTDATEINDEX), tmDTCtrls(DTDATEINDEX)
                gSetShow pbcDT, smDTSave(imDTAIRTIMEINDEX), tmDTCtrls(imDTAIRTIMEINDEX)
                'gMoveTableCtrl pbcPosting, plcDT, tmCtrls(TIMEINDEX).fBoxX + tmCtrls(TIMEINDEX).fBoxW, tmCtrls(TIMEINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15)
                gMoveTableCtrl pbcPosting, plcDT, tmCtrls(DATEINDEX).fBoxX, tmCtrls(TIMEINDEX).fBoxY + (imRowNo - vbcPosting.Value + 1) * (fgBoxGridH + 15) - plcDT.height
                pbcDT.Move plcDT.Left + 90, plcDT.Top + 270
                plcDT.Visible = True ' The main (Green) form for Time Zone information display
                pbcDT.Visible = True ' The Time Zone information form
                pbcDT_Paint ' Put data from smTZSave onto form
                pbcDTSTab.SetFocus
            End If
        Case COPYINDEX 'Copy or ISCI were selected by user
            If imTZCopyAllowed Then
               ' Reads data into smTZSave, smTZShow, lbcZone & lbcTZCopy
                If fmAdjFactorW > 1 Then
                    If tmCtrls(COPYINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15) + plcPosting.Top - plcTZCopy.height < plcPosting.Top Then
                        plcTZCopy.Move tmCtrls(COPYINDEX).fBoxX, plcPosting.Top '1185
                        pbcTZCopy.Move plcTZCopy.Left + 120, plcTZCopy.Top + 255    '1440
                    Else
                        plcTZCopy.Move tmCtrls(COPYINDEX).fBoxX, tmCtrls(COPYINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15) - plcTZCopy.height + plcPosting.Top
                        pbcTZCopy.Move plcTZCopy.Left + 120, plcTZCopy.Top + 255
                    End If
                End If
               mTZReadCopyData
               mCopyPop
               pbcTZCopy.Move (((plcTZCopy.Width - pbcTZCopy.Width) / 2) + plcTZCopy.Left)
               plcTZCopy.Visible = True ' The main (Green) form for Time Zone information display
               pbcTZCopy.Visible = True ' The Time Zone information form
               pbcTZCopy_Paint ' Put data from smTZSave onto form
               imTZBoxNo = -1
               pbcTZSTab.SetFocus ' Transfer control to the Time Zone Form
            Else ' Copy not allowed
              mCopyPop ' populate lbcCopyNm and lbcCopyNmCode for the selected advertiser
              ' Size the listbox to fit this row
              lbcCopyNm.height = gListBoxHeight(lbcCopyNm.ListCount, 6)
              ' Size the editbox to fit this row
              If tgSpf.sUseCartNo <> "N" Then
                  edcDropDown.Width = tmCtrls(COPYINDEX).fBoxW + tmCtrls(ISCIINDEX).fBoxW - cmcDropDown.Width
                  edcDropDown.MaxLength = 62
                  ' Move the editbox (and the cmc control) into position
                  gMoveTableCtrl pbcPosting, edcDropDown, tmCtrls(COPYINDEX).fBoxX, tmCtrls(COPYINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15)
                  lbcCopyNm.Width = tmCtrls(MGOODINDEX).fBoxX - tmCtrls(COPYINDEX).fBoxX
              Else
                  edcDropDown.Width = tmCtrls(ISCIINDEX).fBoxW '- cmcDropDown.Width
                  edcDropDown.MaxLength = 41
                  ' Move the editbox (and the cmc control) into position
                  gMoveTableCtrl pbcPosting, edcDropDown, tmCtrls(ISCIINDEX).fBoxX, tmCtrls(ISCIINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15)
                  lbcCopyNm.Width = tmCtrls(MGOODINDEX).fBoxX - tmCtrls(ISCIINDEX).fBoxX
              End If
              cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
              ' Find this COPY data in lbcCopyNm
              'gFindPartialMatch smSave(SAVALLCOPYINDEX, imRowNo), 0, Len(smSave(SAVALLCOPYINDEX, imRowNo)), lbcCopyNm
              'gFindMatch smSave(SAVALLCOPYINDEX, imRowNo), 0, lbcCopyNm
              smCopy = ""
              If tgSpf.sUseCartNo <> "N" Then
                  smCopy = Trim$(tgSave(ilRowNo).sCopy) & " " & Trim$(tgSave(ilRowNo).sISCI)
              Else
                  smCopy = Trim$(tgSave(ilRowNo).sISCI)
              End If
              If Trim$(tgSave(ilRowNo).sCopyProduct) <> "" Then
                  smCopy = smCopy & " " & Trim$(tgSave(ilRowNo).sCopyProduct)
              End If
              gFindMatch smCopy, 0, lbcCopyNm
              imChgMode = True ' Turn on the switch
              If gLastFound(lbcCopyNm) >= 0 Then
                  ' An entry was found so select it and put
                  ' its data in the dropdown textbxo
                  lbcCopyNm.ListIndex = gLastFound(lbcCopyNm)
              Else ' No data found so re-display the last good data
                  If smCopy = "" Then
                     lbcCopyNm.ListIndex = 0 ' no copy found
                  Else
                  lbcCopyNm.ListIndex = -1
                  End If
              End If
              imComboBoxIndex = lbcCopyNm.ListIndex
              imCopyNmListIndex = imComboBoxIndex
              If lbcCopyNm.ListIndex >= 0 Then
                 edcDropDown.Text = lbcCopyNm.List(lbcCopyNm.ListIndex)
              Else
                 edcDropDown.Text = ""
              End If
              imChgMode = False
              If edcDropDown.Top + edcDropDown.height + lbcCopyNm.height < cmcDone.Top Then
                  lbcCopyNm.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
              Else
                  lbcCopyNm.Move edcDropDown.Left, edcDropDown.Top - lbcCopyNm.height
              End If
              edcDropDown.SelStart = 0
              edcDropDown.SelLength = Len(edcDropDown.Text)
              edcDropDown.Visible = True
              cmcDropDown.Visible = True
              edcDropDown.SetFocus
            End If
        Case PRICEINDEX
            pbcPrice.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveTableCtrl pbcPosting, pbcPrice, tmCtrls(PRICEINDEX).fBoxX, tmCtrls(PRICEINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            pbcPrice_Paint
            pbcPrice.Visible = True
            pbcPrice.SetFocus
    End Select
    ' Save these in case user clicks another box, thus bypassing pbcTab
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindAvail                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get avail within Ssf           *
'*                                                     *
'*******************************************************
Private Function mFindAvail(slSchDate As String, slFindTime As String, ilGameNo As Integer, ilFindAdjAvail As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mFindAvail(slSchDate, slSchTime, ilAvailIndex)
'   Where:
'       slSchDate(I)- Scheduled Date
'       slSchTime(I)- Time that avail is to be found at
'       ilFindAdjAvail(I)- Find closest avail to specified time
'       llSsfRecPos(O)- Ssf record position
'       ilAvailIndex(O)- Index into Ssf where avail is located
'       ilRet(O)- True=Avail found; False=Avail not found
'       lmSsfRecPos(O)- Ssf record position
'
    Dim ilRet As Integer
    Dim llSchDate As Long
    Dim llTime As Long
    Dim llTstTime As Long
    Dim llFndAdjTime As Long
    Dim ilLoop As Integer
    If rbcType(1).Value Then
        mFindAvail = True
        Exit Function
    End If
    llTime = CLng(gTimeToCurrency(slFindTime, False))
    llSchDate = gDateValue(slSchDate)
    imSelectedDay = gWeekDayStr(slSchDate)
    lmSsfDate(imSelectedDay) = 0
    ilRet = gObtainSsfForDateOrGame(imVefCode, llSchDate, slFindTime, ilGameNo, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay))
    If Not ilRet Then
        mFindAvail = False
        Exit Function
    End If
    llFndAdjTime = -1
    For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
       LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTstTime
            If llTime = llTstTime Then 'Replace
                ilAvailIndex = ilLoop
                mFindAvail = True
                Exit Function
            ElseIf (llTstTime < llTime) And (ilFindAdjAvail) Then
                ilAvailIndex = ilLoop
                llFndAdjTime = llTstTime
            ElseIf (llTime < llTstTime) And (ilFindAdjAvail) Then
                If llFndAdjTime = -1 Then
                    ilAvailIndex = ilLoop
                    mFindAvail = True
                    Exit Function
                Else
                    If (llTime - llFndAdjTime) < (llTstTime - llTime) Then
                        mFindAvail = True
                        Exit Function
                    Else
                        ilAvailIndex = ilLoop
                        mFindAvail = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ilLoop
    If (llFndAdjTime <> -1) And (ilFindAdjAvail) Then
        mFindAvail = True
        Exit Function
    End If
    mFindAvail = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindSpotOrigTime               *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Find spot origianl time        *
'*                                                     *
'*******************************************************
Private Function mFindSpotOrigTime(slSchDate As String, slTime As String, ilGameNo As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mFindAvail(slSchDate, slTime)
'   Where:
'       slSchDate(I)- Scheduled date
'       slTime(O)- Time of avail containing spot
'       5/20/11
'       ilAvailIndex(0)- Avail index which contained the spot
'       llSsfRecPos(O)- Ssf record position
'       ilRet(O)- True=Avail found; False=Avail not found
'       lmSsfRecPos(O)- Ssf record position
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilType As Integer


    If rbcType(1).Value Then
        mFindSpotOrigTime = True
        Exit Function
    End If
    ilType = ilGameNo
    imSelectedDay = gWeekDayStr(slSchDate)
    imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
    tmSsfSrchKey.iType = ilType 'slType-On Air
    tmSsfSrchKey.iVefCode = imVefCode
    gPackDate slSchDate, ilDate0, ilDate1
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf(imSelectedDay).iType = ilType) And (tmSsf(imSelectedDay).iVefCode = imVefCode) And (tmSsf(imSelectedDay).iDate(0) = ilDate0) And (tmSsf(imSelectedDay).iDate(1) = ilDate1)
        For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
           LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                '5/20/11
                ilAvailIndex = ilLoop
                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
            ElseIf (tmAvail.iRecType And &HF) >= 10 Then
               LSet tmSpot = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
                If tmSpot.lSdfCode = tmSdf.lCode Then
                    mFindSpotOrigTime = True
                    ilRet = gSSFGetPosition(hmSsf, lmSsfRecPos(imSelectedDay))
                    Exit Function
                End If
            End If
        Next ilLoop
        imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
        ilRet = gSSFGetNext(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mFindSpotOrigTime = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGenSaveImage                   *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set save values for a spot     *
'*                                                     *
'*******************************************************
Private Sub mGenShowImage()
    Dim ilRet As Integer    'Return status
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilAnfCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slCount As String

    ReDim tgShow(0 To 1) As SHOWINFO
    'For ilIndex = LBound(tgSave) To UBound(tgSave) - 1 Step 1
    For ilIndex = LBONE To UBound(tgSave) - 1 Step 1
        gPackDate Trim$(tgSave(ilIndex).sAirDate), ilDate0, ilDate1
        gUnpackDateForSort ilDate0, ilDate1, slDate
        If tgSave(ilIndex).sXMid = "Y" Then
            slDate = gAddStr(slDate, "1")
        End If
        slTime = Trim$(str$(gTimeToLong(Trim$(tgSave(ilIndex).sAirTime), False)))
        Do While Len(slTime) < 6
            slTime = "0" & slTime
        Loop
        slCount = Trim$(str$(tgSave(ilIndex).iCount))
        Do While Len(slCount) < 5
            slCount = "0" & slCount
        Loop
        If rbcSort(0).Value Then
            tgSave(ilIndex).sKey = slDate & slTime & slCount
        Else
            tgSave(ilIndex).sKey = Trim$(tgSave(ilIndex).sAdvtName) & slDate & slTime & slCount
        End If
    Next ilIndex
    If UBound(tgSave) - 1 > 1 Then
        'ArraySortTyp fnAV(tgSave(), 1), UBound(tgSave) - 1, 0, LenB(tgSave(1)), 0, LenB(tgSave(1).sKey), 0
        For ilIndex = LBound(tgSave) To UBound(tgSave) - 1 Step 1
            tgSave(ilIndex) = tgSave(ilIndex + 1)
        Next ilIndex
        ReDim Preserve tgSave(0 To UBound(tgSave) - 1) As SAVEINFO
        ArraySortTyp fnAV(tgSave(), 0), UBound(tgSave), 0, LenB(tgSave(0)), 0, LenB(tgSave(0).sKey), 0
        ReDim Preserve tgSave(0 To UBound(tgSave) + 1) As SAVEINFO
        For ilIndex = UBound(tgSave) - 1 To LBound(tgSave) Step -1
            tgSave(ilIndex + 1) = tgSave(ilIndex)
        Next ilIndex
    End If
    If imAvailSelectedIndex <= 0 Then
        ilAnfCode = -1
    Else
        slNameCode = tmAvailCode(imAvailSelectedIndex - 1).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilAnfCode = Val(slCode)
    End If
    'For ilIndex = LBound(tgSave) To UBound(tgSave) - 1 Step 1
    For ilIndex = LBONE To UBound(tgSave) - 1 Step 1
        ilFound = False
        If (ilAnfCode < 0) Or (rbcType(1).Value) Then
            ilFound = True
        Else
            If ilAnfCode = tgSave(ilIndex).ianfCode Then
                ilFound = True
            End If
        End If
        If ilFound Then
            ilUpperBound = UBound(tgShow)
            If tgSave(ilIndex).iType = 1 Then
                mCreateShowImage ilIndex, ilUpperBound
                ReDim Preserve tgShow(0 To ilUpperBound + 1) As SHOWINFO
            End If
        End If
    Next ilIndex
    imSettingValue = True
    vbcPosting.Min = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
    imSettingValue = True
    'If UBound(smSave, 2) <= vbcPosting.LargeChange + 1 Then
    If UBound(tgShow) <= vbcPosting.LargeChange + 1 Then
    ' If this is used, there are probably 0 or 1 records
        vbcPosting.Max = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
    Else
    ' Saves, what amounts to, the count of records just retrieved
        'vbcPosting.Max = UBound(smSave, 2) - vbcPosting.LargeChange
        vbcPosting.Max = UBound(tgShow) - vbcPosting.LargeChange
    End If
    imSettingValue = True
    vbcPosting.Value = vbcPosting.Min
    imSettingValue = False
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetVehIndex                    *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get vehicle index and option   *
'*                      index                          *
'*                                                     *
'*******************************************************
Private Sub mGetVehIndex()
'
'   mGetVehIndex
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    Dim ilVff As Integer
    
    slNameCode = tmUserVehicle(imVehSelectedIndex).sKey    'Traffic!lbcUserVehicle.List(imVehSelectedIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mGetVehIndexErr
    gCPErrorMsg ilRet, "mGetVehIndex (gParseItem field 2: Vehicle)", PostLog
    On Error GoTo 0
    imVefCode = Val(slCode)
    tmVefSrchKey.iCode = imVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    imVpfIndex = gBinarySearchVpfPlus(imVefCode)    'gVpfFind(PostLog, imVefCode)
    If ((tgSpf.sBActDayCompl <> "N") And (imWM = 0)) Or (((Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) = USINGLIVELOG) And ((tgVpf(imVpfIndex).sGenLog = "L") Or (tgVpf(imVpfIndex).sGenLog = "A"))) Then
        plcComplete.Visible = True
        plcComplete.Enabled = True
    Else
        plcComplete.Visible = False
        plcComplete.Enabled = False
    End If
    smPostLogSource = ""
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If tgVff(ilVff).iVefCode = imVefCode Then
            smPostLogSource = tgVff(ilVff).sPostLogSource
            Exit For
        End If
    Next ilVff
    ' See if Vehicle can have time zone at all
    imTZCopyAllowed = False
    For ilLoop = LBound(tgVpf(imVpfIndex).sGZone) To UBound(tgVpf(imVpfIndex).sGZone) Step 1
       If Trim$(tgVpf(imVpfIndex).sGZone(ilLoop)) <> "" Then
          imTZCopyAllowed = True
          Exit For
       End If
    Next ilLoop
    For ilLoop = 1 To 8 Step 1
        smZones(ilLoop) = ""
    Next ilLoop
    imNoZones = 0
    If imTZCopyAllowed Then
        smZones(1) = "[All]"
        smZones(2) = "[Other]"
        imNoZones = 2
        For ilLoop = LBound(tgVpf(imVpfIndex).sGZone) To UBound(tgVpf(imVpfIndex).sGZone) Step 1
            If Trim$(tgVpf(imVpfIndex).sGZone(ilLoop)) <> "" Then
                imNoZones = imNoZones + 1
                smZones(imNoZones) = Trim$(tgVpf(imVpfIndex).sGZone(ilLoop))
            End If
        Next ilLoop
        For ilLoop = 1 To imNoZones Step 1
            smTZSave(1, ilLoop) = Trim$(smZones(ilLoop))  ' Save Time Zone
            gSetShow pbcTZCopy, smTZSave(1, ilLoop), tmTZCtrls(ZONEINDEX)
            smTZShow(1, ilLoop) = tmTZCtrls(ZONEINDEX).sShow
        Next ilLoop
    End If
    ' if timezone copy then load up the list box with the possible timezones
'    If imTZCopyAllowed Then
'       lbcZone.Clear
       ' [All] means SDF sPtType is 1
'       lbcZone.AddItem "[All]"
'       For ilLoop = LBound(tgVpf(imVpfIndex).sGZone) To UBound(tgVpf(imVpfIndex).sGZone) Step 1
'          If Trim$(tgVpf(imVpfIndex).sGZone(ilLoop)) <> "" Then
'             lbcZone.AddItem Trim$(tgVpf(imVpfIndex).sGZone(ilLoop))
'          End If
'       Next ilLoop
'    End If
    'Get user index into TgUrf for vehicle
    imUrfIndex = 0
    For ilLoop = 0 To UBound(tgUrf) Step 1
        If imVefCode = tgUrf(ilLoop).iVefCode Then
            imUrfIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    ReDim tmVlf0(0 To 0) As VLF
    ReDim tmVlf6(0 To 0) As VLF
    ReDim tmVlf7(0 To 0) As VLF
    Exit Sub
mGetVehIndexErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRet As Integer    'Return Status
    Dim slStr As String
    imLBCtrls = 1
    imLBDTCtrls = 1
    imLBTZCtrls = 1
    imLBMdCtrls = 1
    imLBCDCtrls = 1
    igJobShowing(POSTLOGSJOB) = True
    imButtonIndex = -1
    imFirstActivate = True
    imFirstTime = True
    imTerminate = False         'terminate if true
    imcKey.Picture = IconTraf!imcKey.Picture
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.height = 165
    'PostLog.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterForm PostLog
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imWM = 0    'Weekly view
    pbcWM_Paint
    imIgnoreChg = False
    lmBonusDate = 0
    imDefaultDateIndex = -1
    imTerminate = False
    imSelectDelay = False
    imStartMode = True
    imVehSelectedIndex = -1
    imDateSelectedIndex = -1
    imAdvtSelectedIndex = -1
    imAvailSelectedIndex = -1
    lmAvailDate = 0
    imAvailAnfCode = 0
    imIgnoreRightMove = False
    imUrfIndex = 0
    imInTab = False
    'imWeek = 0
    imVehicle = 0
    imCalType = 0
    imProcClickMode = False
    imSvVehSelectedIndex = -1
    imSvDateSelectedIndex = -1
    smPostLogSource = ""
    bmInPackage = False
    imSdfRecLen = Len(tmSdf)  'Get and save SDF record length
    imAdfRecLen = Len(tmAdf)  'Get and save ADF record length
    imAgfRecLen = Len(tmAgf)  'Get and save AGF record length
    imCHFRecLen = Len(tmChf)  'Get and save CHF record length
    imClfRecLen = Len(tmClf)  'Get and save CLF record length
    imCifRecLen = Len(tmCif)  'Get and save CIF record length
    imMnfRecLen = Len(tmMnf)  'Get and save MNF record length
    imLcfRecLen = Len(tmLcf)  'Get and save LCF record length
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
    imCcfRecLen = Len(tmCcf)  'Get and save CCF record length
    imCpfRecLen = Len(tmCpf)  'Get and save CPF record length
    imTzfRecLen = Len(tmTzf)  'Get and save TzF record length
    imMcfRecLen = Len(tmMcf)  'Get and save MCF record length
    imRdfRecLen = Len(tmLnRdf)  'Get and save RDF record length
    imCffRecLen = Len(tmCff)  'Get and save CFF record length
    imCgfRecLen = Len(tmCgf)  'Get and save CFF record length
    imSmfRecLen = Len(tmSmf)
    imVsfRecLen = Len(tmVsf)
    imIihfRecLen = Len(tmIihf)
    imBoxNo = -1 'Initialize current Box to N/A
    imTZBoxNo = -1 'Initialize current Box to N/A
    imDTBoxNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imLbcMouseDown = False
    imRowNo = 1
    imTZRowNo = -1
    imMdRowNo = 0
    imChgMode = False
    imBSMode = False
    imSdfChg = False
    imListChgMode = False
    For ilLoop = 0 To 6 Step 1
        imSdfAnyChg(ilLoop) = False
    Next ilLoop
    imSettingValue = False
    imLbcArrowSetting = False
    pbcDT.height = 1080
    plcDT.height = 1425
    imDTAIRTIMEINDEX = 3
    imDTAVAILTIMEINDEX = 2
    ' Establish the initial dimensions of each array
    'ReDim smShow(1 To 13, 1 To 1) As String
    ReDim tgShow(0 To 1) As SHOWINFO        'Index zero ignored
    'ReDim smSave(1 To 10, 1 To 1) As String
    ReDim tgSave(0 To 1) As SAVEINFO
    'ReDim imSave(1 To 2, 1 To 1) As Integer
    'ReDim imPostSpotInfo(1 To 4, 1 To 1) As Integer
    'ReDim smMdShow(1 To 8, 1 To 1) As String
    'ReDim smMdSave(1 To 1, 1 To 1) As String
    'ReDim smMdSchStatus(1 To 1, 1 To 1) As String
    'ReDim llMdRecPos(1 To 1) As Long
    ReDim tgMdSdfRec(0 To 1) As MDSDFREC
    ReDim tgMdSaveInfo(0 To 1) As MDSAVEINFO
    ReDim tgMdShowInfo(0 To 1) As MDSHOWINFO
    ReDim tmVlf0(0 To 0) As VLF
    ReDim tmVlf6(0 To 0) As VLF
    ReDim tmVlf7(0 To 0) As VLF
    imSettingValue = True
    vbcPosting.Min = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
    vbcPosting.Max = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
    vbcMissed.Min = LBONE   'LBound(tgMdShowInfo)
    vbcMissed.Max = LBONE   'LBound(tgMdShowInfo)
    imSettingValue = True
    vbcPosting.Value = LBONE    'LBound(tgShow)   'LBound(smSave, 2)
    ' Spot Detail File
    hmSdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", PostLog
    On Error GoTo 0
    hmSvSdf = hmSdf
    hmPsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPsf, "", sgDBPath & "Psf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Psf.Btr)", PostLog
    On Error GoTo 0
    ' Advertisers Detail File
    hmAdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", PostLog
    On Error GoTo 0
    ' Agency Detail File
    hmAgf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", PostLog
    On Error GoTo 0
    ' Contract Header File
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", PostLog
    On Error GoTo 0
    ' Contract Schedule Line
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", PostLog
    On Error GoTo 0
    ' Contract Games
    hmCgf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cgf.Btr)", PostLog
    On Error GoTo 0
    hmGhf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", PostLog
    On Error GoTo 0
    hmGsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", PostLog
    On Error GoTo 0
    ' Spot makegood File
    hmSmf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", PostLog
    On Error GoTo 0
    ' Copy Inventory File
    hmCif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", PostLog
    On Error GoTo 0
    ' Multi-Name File
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", PostLog
    On Error GoTo 0
    ' Log Calendar File
    hmLcf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", PostLog
    On Error GoTo 0
    'Copy Rotation
    hmCrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", PostLog
    On Error GoTo 0
    ' Spot summary File
    hmSsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", PostLog
    On Error GoTo 0
    ' Vehicle File
    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", PostLog
    On Error GoTo 0
    ' Vehicle File
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", PostLog
    On Error GoTo 0
    ' Vehicle Link File
    hmVLF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vlf.Btr)", PostLog
    On Error GoTo 0
    ' Copy Combo Inventory File
    hmCcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCcf, "", sgDBPath & "Ccf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ccf.Btr)", PostLog
    On Error GoTo 0
    ' Copy Product/Agency File
    hmCpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", PostLog
    On Error GoTo 0
    ' Time Zone Copy File
    hmTzf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Tzf.Btr)", PostLog
    On Error GoTo 0
    ' Media Codes File
    hmMcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", PostLog
    On Error GoTo 0

    ' Rate card Program/Time File
    hmRdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", PostLog
    On Error GoTo 0
    ' Contract Flight File
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", PostLog
    On Error GoTo 0
    'Record Locks
    hmRlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rlf.Btr)", PostLog
    On Error GoTo 0

    'Product
    hmPrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf.Btr)", PostLog
    On Error GoTo 0
    'Feed
    hmFsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Fsf.Btr)", PostLog
    On Error GoTo 0
    imFsfRecLen = Len(tmFsf)  'Get and save CHF record length
    'Feed Name
    hmFnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmFnf, "", sgDBPath & "Fnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Fnf.Btr)", PostLog
    On Error GoTo 0

    hmIihf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmIihf, "", sgDBPath & "Iihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Iihf.Btr)", PostLog
    On Error GoTo 0

    hmSxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sxf.Btr)", PostLog
    On Error GoTo 0

    Screen.MousePointer = vbHourglass
    imcTrash.Picture = IconTraf!imcFireOut.Picture
    imcTrash.Visible = False
    imcHidden.Picture = IconTraf!imcHideUp.Picture
    imcHidden.Visible = False
    gObtainPostLogAvailCode
    'Populate facilty and event type list boxes
    cbcVeh.Clear 'Force population
    mVehPop
    If imTerminate Then ' this is set by mVehPop if error occurs
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    lbcMissed.Clear 'Force list box to be populated
    mMissedPop
    If imTerminate Then
        Exit Sub
    End If
    cbcAvailName.Clear
    mAvailPop
    If imTerminate Then
        Exit Sub
    End If

    mTeamPop
    'cbcWeek.AddItem "Current Week", 0
    'cbcWeek.AddItem "Current Month", 1
    'cbcWeek.AddItem "Current + Past", 2
    'imChgMode = True
    'cbcWeek.ListIndex = 0
    'imChgMode = False
    slStr = cbcVeh.List(0)
    cbcVehicle.AddItem slStr, 0
    cbcVehicle.AddItem "All Vehicles", 1
    imChgMode = True
    cbcVehicle.ListIndex = 0
    imChgMode = False
    'PostLog.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterForm PostLog
    'Traffic!plcHelp.Caption = ""

    If tgSpf.sBActDayCompl <> "Y" Then      'use day is complete feature?
        plcComplete.Visible = False
        plcComplete.Enabled = False
    End If
    If tgSpf.sSystemType = "R" Then         'radio station, dont allow monthly posting
        pbcWM.Enabled = False
    End If

    mInitBox
    gCenterForm PostLog
    tmcStart.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long
    Dim llMdMax As Long
    Dim ilVff As Integer
    Dim blShowImport As Boolean

    flTextHeight = pbcPosting.TextHeight("1") - 35
    If (tgSpf.sCPkOrdered = "Y") Or ((tgSpf.sCPkEqual = "Y") And ((tgSpf.sInvAirOrder = "O") Or (tgSpf.sInvAirOrder = "S"))) Then 'And (tgSpf.sCPkAired <> "Y") And (tgSpf.sCPkEqual <> "Y")Then
        plcSelect.Move 840, 105
    Else
        rbcType(1).Enabled = False
        plcSelect.Move 840, 105
    End If
    imDateBox = 0
    plcCalendar.Move plcSelect.Left + plcSelect.Width - fgBevelX - plcCalendar.Width, plcSelect.Top + edcDate.height + fgBevelY
    lbcGameNo.Move plcSelect.Left + plcSelect.Width - fgBevelX - lbcGameNo.Width, plcSelect.Top + edcDate.height + fgBevelY
    'Position panel and Posting picture areas with panel
    plcPosting.Move 105, 795, pbcPosting.Width + vbcPosting.Width + fgPanelAdj, pbcPosting.height + fgPanelAdj
    pbcPosting.Move plcPosting.Left + fgBevelX, plcPosting.Top + fgBevelY
    pbcArrow.Move plcPosting.Left - pbcArrow.Width - 15
    vbcPosting.Move plcPosting.Width - vbcPosting.Width - 2 * fgBevelX, fgBevelY - 30, vbcPosting.Width, pbcPosting.height
    plcInfo.Move plcPosting.Left + (plcPosting.Left + plcPosting.Width - plcInfo.Width) / 2, plcPosting.Top + plcPosting.height
    pbcKey.Move plcPosting.Left, plcSelect.Top + plcSelect.height + 45
    'Date
    gSetCtrl tmCtrls(DATEINDEX), 30, 375, 660, fgBoxGridH
    'Time
    gSetCtrl tmCtrls(TIMEINDEX), 705, tmCtrls(DATEINDEX).fBoxY, 840, fgBoxGridH
    'Length
    gSetCtrl tmCtrls(LENINDEX), 1560, tmCtrls(DATEINDEX).fBoxY, 345, fgBoxGridH
    'Advertiser
    gSetCtrl tmCtrls(ADVTINDEX), 1920, tmCtrls(DATEINDEX).fBoxY, 2205, fgBoxGridH
    'Time Zone
    gSetCtrl tmCtrls(TZONEINDEX), 4140, tmCtrls(DATEINDEX).fBoxY, 160, fgBoxGridH
    'Copy
    gSetCtrl tmCtrls(COPYINDEX), 4320, tmCtrls(DATEINDEX).fBoxY, 870, fgBoxGridH
    'ISCI
    gSetCtrl tmCtrls(ISCIINDEX), 5205, tmCtrls(DATEINDEX).fBoxY, 1155, fgBoxGridH
    'Contract #
    gSetCtrl tmCtrls(CNTRINDEX), 6375, tmCtrls(DATEINDEX).fBoxY, 780, fgBoxGridH
    'Line #
    gSetCtrl tmCtrls(LINEINDEX), 7170, tmCtrls(DATEINDEX).fBoxY, 360, fgBoxGridH
    'Type
    gSetCtrl tmCtrls(TYPEINDEX), 7550, tmCtrls(DATEINDEX).fBoxY, 160, fgBoxGridH
    'Price
    gSetCtrl tmCtrls(PRICEINDEX), 7725, tmCtrls(DATEINDEX).fBoxY, 840, fgBoxGridH
    'Make Good
    gSetCtrl tmCtrls(MGOODINDEX), 8580, tmCtrls(DATEINDEX).fBoxY, 160, fgBoxGridH
    'Audit
    gSetCtrl tmCtrls(AUDINDEX), 8755, tmCtrls(DATEINDEX).fBoxY, 160, fgBoxGridH
    tmCtrls(AUDINDEX).iReq = False
    'Position panel and Missed picture areas with panel
    lbcMissed.Move 30, 345
    plcSpots.Move 75, 3675, lbcMissed.Width + pbcMissed.Width + vbcMissed.Width + 2 * fgBevelX, 1785
    pbcMissed.Move plcSpots.Left + fgBevelX + lbcMissed.Width + 30, plcSpots.Top + plcSpots.height - pbcMissed.height - fgBevelY
    vbcMissed.Move pbcMissed.Left + pbcMissed.Width - plcSpots.Left - 30, pbcMissed.Top - plcSpots.Top - 30
    lacMdFrame.Width = pbcMissed.Width
    'Advertiser
    gSetCtrl tmMdCtrls(1), 30, 225, 1485, fgBoxGridH
    'Contract #
    gSetCtrl tmMdCtrls(2), 1530, tmMdCtrls(1).fBoxY, 780, fgBoxGridH
    'Vehicle
    gSetCtrl tmMdCtrls(3), 2325, tmMdCtrls(1).fBoxY, 1440, fgBoxGridH
    'Length
    gSetCtrl tmMdCtrls(4), 3780, tmMdCtrls(1).fBoxY, 300, fgBoxGridH
    'Week Date
    gSetCtrl tmMdCtrls(5), 4095, tmMdCtrls(1).fBoxY, 720, fgBoxGridH
    'Contract end date
    gSetCtrl tmMdCtrls(6), 4830, tmMdCtrls(1).fBoxY, 720, fgBoxGridH
    'Daypart
    gSetCtrl tmMdCtrls(7), 5565, tmMdCtrls(1).fBoxY, 1380, fgBoxGridH
    'Spots
    gSetCtrl tmMdCtrls(8), 6960, tmMdCtrls(1).fBoxY, 420, fgBoxGridH
    'Position panel and picture areas for copy
    plcTZCopy.Move 2460, 1185
    pbcTZCopy.Move 2580, 1440
    ' Time Zone Fields
    gSetCtrl tmTZCtrls(ZONEINDEX), 30, 225, 765, fgBoxGridH
    gSetCtrl tmTZCtrls(TZCOPYINDEX), 810, tmTZCtrls(ZONEINDEX).fBoxY, 4425, fgBoxGridH

    'Date/Time
    gSetCtrl tmDTCtrls(1), 30, 30, 1185, fgBoxStH
    gSetCtrl tmDTCtrls(2), 30, tmDTCtrls(1).fBoxY + fgStDeltaY, 1185, fgBoxStH
    '6/16/11
    gSetCtrl tmDTCtrls(3), 30, tmDTCtrls(2).fBoxY + fgStDeltaY, 1185, fgBoxStH
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop

    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxW)
        Do While (tmCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmCtrls(ilLoop).fBoxW = tmCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmCtrls(ilLoop).fBoxX)
            Do While (tmCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmCtrls(ilLoop).fBoxX = tmCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 < tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmCtrls(ilLoop - 1).fBoxX + tmCtrls(ilLoop - 1).fBoxW + 15 > tmCtrls(ilLoop).fBoxX Then
                        tmCtrls(ilLoop - 1).fBoxW = tmCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmCtrls(ilLoop).fBoxX + tmCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop

    pbcPosting.Picture = LoadPicture("")
    pbcPosting.Width = llMax
    plcPosting.Width = llMax + vbcPosting.Width + 2 * fgBevelX + 15
    lacPtFrame.Width = llMax - 15
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    
    blShowImport = False
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If tgVff(ilVff).sPostLogSource = "S" Then
            cmcImport.Visible = True
            cmcManual.Visible = True
            blShowImport = True
            Exit For
        End If
    Next ilVff
    
    If blShowImport Then
        'cmcDone.Left = (PostLog.Width - 5 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
        cmcDone.Left = (PostLog.Width - (cmcDone.Width + cmcCancel.Width + cmcReport.Width + cmcImport.Width + cmcAdServer.Width + (cmcImport.Width - cmcDone.Width) + cmcManual.Width + (cmcManual.Width - cmcDone.Width) + 4 * ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
        cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
        cmcReport.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
        cmcImport.Left = cmcReport.Left + cmcReport.Width + ilSpaceBetweenButtons
        cmcManual.Left = cmcImport.Left + cmcImport.Width + ilSpaceBetweenButtons
        'LB 02/10/21
        cmcAdServer.Left = cmcManual.Left + cmcManual.Width + ilSpaceBetweenButtons
    Else
        cmcDone.Left = (PostLog.Width - 4 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
        cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
        cmcReport.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
        'LB 02/10/21
        cmcAdServer.Left = cmcReport.Left + cmcReport.Width + ilSpaceBetweenButtons
    End If
    cmcDone.Top = PostLog.height - (3 * cmcDone.height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcReport.Top = cmcDone.Top
    cmcImport.Top = cmcDone.Top
    cmcManual.Top = cmcDone.Top
     'LB 02/10/21
    cmcAdServer.Top = cmcDone.Top
    plcSpots.Top = cmcDone.Top - plcSpots.height - 120
    imcTrash.Top = plcSpots.Top + plcSpots.height '+ 30
    imcTrash.Left = PostLog.Width - (3 * imcTrash.Width) / 2
    imcHidden.Top = imcTrash.Top
    llAdjTop = plcSpots.Top - plcSort.Top - plcSort.height - 120 - tmCtrls(1).fBoxY
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcPosting.Top + llAdjTop + 2 * fgBevelY + 240 < plcSpots.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcPosting.height = llAdjTop + 2 * fgBevelY
    pbcPosting.Left = plcPosting.Left + fgBevelX
    pbcPosting.Top = plcPosting.Top + fgBevelY
    pbcPosting.height = plcPosting.height - 2 * fgBevelY
    vbcPosting.Left = plcPosting.Width - vbcPosting.Width - fgBevelX - 30
    vbcPosting.Top = fgBevelY
    vbcPosting.height = pbcPosting.height

    llMdMax = 0
    For ilLoop = imLBMdCtrls To UBound(tmMdCtrls) Step 1
        tmMdCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmMdCtrls(ilLoop).fBoxW)
        Do While (tmMdCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmMdCtrls(ilLoop).fBoxW = tmMdCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmMdCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmMdCtrls(ilLoop).fBoxX)
            Do While (tmMdCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmMdCtrls(ilLoop).fBoxX = tmMdCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmMdCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmMdCtrls(ilLoop - 1).fBoxX + tmMdCtrls(ilLoop - 1).fBoxW + 15 < tmMdCtrls(ilLoop).fBoxX Then
                        tmMdCtrls(ilLoop - 1).fBoxW = tmMdCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmMdCtrls(ilLoop - 1).fBoxX + tmMdCtrls(ilLoop - 1).fBoxW + 15 > tmMdCtrls(ilLoop).fBoxX Then
                        tmMdCtrls(ilLoop - 1).fBoxW = tmMdCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmMdCtrls(ilLoop).fBoxX + tmMdCtrls(ilLoop).fBoxW + 15 > llMdMax Then
            llMdMax = tmMdCtrls(ilLoop).fBoxX + tmMdCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    pbcMissed.Width = llMdMax
    plcSpots.Width = pbcMissed.Width + vbcMissed.Width + 2 * fgBevelX
    plcSpots.Left = plcPosting.Left + plcPosting.Width - plcSpots.Width
    pbcMissed.Left = plcSpots.Left + fgBevelX
    'lbcMissed.Width = plcSpots.Width - pbcMissed.Width - vbcMissed.Width - 2 * fgBevelX - 30
    plcReason.Left = plcPosting.Left
    plcReason.Top = plcSpots.Top
    plcReason.Width = plcSpots.Left - 2 * plcReason.Left
    lbcMissed.Width = plcReason.Width - 2 * fgBevelX - 30
    pbcMissed.Top = plcSpots.Top + plcSpots.height - pbcMissed.height - fgBevelY
    pbcMissed.Left = plcSpots.Left + fgBevelX
    vbcMissed.Left = plcSpots.Width - vbcMissed.Width - fgBevelX
    vbcMissed.Top = pbcMissed.Top - plcSpots.Top
    vbcMissed.height = pbcMissed.height
    pbcMissed.Picture = LoadPicture("")
    lacMdFrame.Width = pbcMissed.Width

    If fmAdjFactorW >= 1.2 Then
        'plcSelect.Width = CLng(1.2 * plcSelect.Width)
        'Do While (plcSelect.Width Mod 15) <> 0
        '    plcSelect.Width = plcSelect.Width + 1
        'Loop
        cbcVeh.Width = (3 * cbcVeh.Width) / 2
        cbcAvailName.Left = cbcVeh.Left + cbcVeh.Width + 90
        cbcAvailName.Width = (3 * cbcAvailName.Width) / 2
        pbcWM.Left = cbcAvailName.Left + cbcAvailName.Width + 90
        edcDate.Left = pbcWM.Left + pbcWM.Width + 90
        cmcDate.Left = edcDate.Left + edcDate.Width
        plcSelect.Width = cmcDate.Left + cmcDate.Width + 2 * fgBevelX
    End If
    If fmAdjFactorW > 1 Then
        plcTZCopy.Move tmCtrls(COPYINDEX).fBoxX, 1185
        pbcTZCopy.Move plcTZCopy.Left + 120, 1440
    End If
    If fmAdjFactorW > 1 Then
        cbcVehicle.Width = fmAdjFactorW * cbcVehicle.Width
        ckcInclude(3).Left = cbcVehicle.Left + cbcVehicle.Width + 60
        ckcInclude(0).Left = ckcInclude(3).Left + ckcInclude(3).Width
        ckcInclude(1).Left = ckcInclude(0).Left + ckcInclude(0).Width
        ckcInclude(2).Left = ckcInclude(1).Left + ckcInclude(1).Width
    End If
    plcCalendar.Move plcSelect.Left + plcSelect.Width - fgBevelX - plcCalendar.Width, plcSelect.Top + edcDate.height + fgBevelY
    lbcGameNo.Move plcSelect.Left + plcSelect.Width - fgBevelX - lbcGameNo.Width, plcSelect.Top + edcDate.height + fgBevelY
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeUnschSpot                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create a Sdf records            *
'*                                                     *
'*                     Similar to code within          *
'*                     CntSchd.Bas                     *
'*                                                     *
'*******************************************************
Private Function mMakeUnschSpot(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long, ilGameNo As Integer, slMissedDate As String, ilVefCode As Integer, llSdfRecPos As Long) As Integer
'
'   ilRet = mMakeUnschSpot(llChfCode, ilLineNo, lMisseddate, slPriceType, llSdfRecPos)
'   Where:
'       llChfCode(I)- Chf Code
'       ilLineNo(I)- Line number
'       slMissedDate(I)- Date to create spot for
'       ilExtraSpot(I)- True=Extra Bonus Spot
'       llSdfRecPos(O)- Sdf Record position
'
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilRet As Integer

    ilRet = mReadChfClfRdfRec(llChfCode, ilLineNo, llFsfCode)
    If Not ilRet Then
        mMakeUnschSpot = False
        Exit Function
    End If
    tmSdf.lCode = 0
    tmSdf.iVefCode = tmClf.iVefCode
    tmSdf.lChfCode = tmChf.lCode    'Contract code
    tmSdf.iLineNo = tmClf.iLine    'Line number
    tmSdf.lFsfCode = llFsfCode
    tmSdf.iAdfCode = tmChf.iAdfCode 'Advertiser code number
    gPackDate slMissedDate, tmSdf.iDate(0), tmSdf.iDate(1)
    llDate = gDateValue(slMissedDate)
    ilDay = gWeekDayLong(llDate)
    If (tmLnRdf.iLtfCode(0) <> 0) Or (tmLnRdf.iLtfCode(1) <> 0) Or (tmLnRdf.iLtfCode(2) <> 0) Then
        tmSdf.iTime(0) = 0
        tmSdf.iTime(1) = 0
    Else    'Time buy- check if override times defined (if so, use them as bump times)
        If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
            tmSdf.iTime(0) = 0
            tmSdf.iTime(1) = 0
        Else
            tmSdf.iTime(0) = tmClf.iStartTime(0)
            tmSdf.iTime(1) = tmClf.iStartTime(1)
        End If
    End If
    tmSdf.sSchStatus = "M"    'S=Scheduled, M=Missed,
                                'G=Makegood, A=on alternate log but not MG, B=on alternate Log and MG,
                                'C=Cancelled
    tmSdf.iMnfMissed = igDefaultMnfMissed   'Missed reason
    tmSdf.sTracer = " "   'M=Mouse move, N=On demand & mouse moved, C=Created in post log,
                            'N=N/A, D=on Demand & created in post log
    tmSdf.sAffChg = " "   'T=Time change, C=Copy change, B=Time and copy changed, blank=no change
    If (llFsfCode <= 0) Or (tmFsf.lCifCode <= 0) Then
        tmSdf.sPtType = "0"
        tmSdf.lCopyCode = 0        'Copy inventory code
        tmSdf.iRotNo = 0
    Else
        tmSdf.sPtType = "1"
        tmSdf.lCopyCode = tmFsf.lCifCode
    End If
    tmSdf.iRotNo = 0
    tmSdf.iLen = tmClf.iLen         'Spot length
    tmSdf.sPriceType = "B"  'tmClf.sPriceType 'T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    tmSdf.sSpotType = "X"   '"T"=Remnant; Q=per Inquire; S=PSA; M=Promo; X=Extra Spot
    tmSdf.sBill = "N"
    tmSdf.iGameNo = ilGameNo
    tmSdf.iUrfCode = tgUrf(0).iCode      'Last user who modified spot
    tmSdf.sXCrossMidnight = "N"
    tmSdf.sWasMG = "N"
    tmSdf.sFromWorkArea = "N"
    tmSdf.sUnused = ""
    ilRet = btrInsert(hmSdf, tmSdf, imSdfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        mMakeUnschSpot = False
        Exit Function
    End If
    ilRet = btrGetPosition(hmSdf, llSdfRecPos)
    If ilRet <> BTRV_ERR_NONE Then
        mMakeUnschSpot = False
    Else
        mMakeUnschSpot = True
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMdCntrPop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Bonus list box if     *
'*                      required                       *
'*                                                     *
'*******************************************************
Private Sub mMdCntrPop(llMdStartdate As Long, llMdEndDate As Long)
    Dim ilLoop As Integer
    Dim slSTime As String
    Dim slETime As String
    Dim llSTime As Long
    Dim llETime As Long
    Dim ilRet As Integer
    Dim llSDate As Long
    Dim llEDate As Long
    Dim slSDate As String
    Dim slEDate As String
    Dim ilDay As Integer
    Dim ilUpper As Integer
    Dim slCntrType As String
    Dim slCntrStatus As String
    Dim ilHOType As Integer
    Dim ilClf As Integer
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim llLnSDate As Long
    Dim llLnEDate As Long
    'ReDim ilAllowedDays(0 To 6) As Integer
    If ckcInclude(0).Value = vbUnchecked Then
        Exit Sub
    End If
    slSTime = "12M"
    slETime = "12M"
    llSTime = CLng(gTimeToCurrency(slSTime, False))
    llETime = CLng(gTimeToCurrency(slETime, True))
    llSDate = llMdStartdate
    llEDate = llMdEndDate
    slSDate = Format$(llSDate, "m/d/yy")
    slEDate = Format$(llEDate, "m/d/yy")
    slCntrType = "C"
    slCntrStatus = "HO"
    ilHOType = 1
    sgCntrForDateStamp = ""
    ilRet = gObtainCntrForDate(PostLog, slSDate, slEDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
    If (ilRet <> CP_MSG_NOPOPREQ) And (ilRet <> CP_MSG_NONE) Then
        Exit Sub
    End If
    For ilLoop = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        ilRet = gObtainChfClf(hmCHF, hmClf, tmChfAdvtExt(ilLoop).lCode, False, tmChf, tgClfPostLog())
        If ilRet And ((tmChf.sType = "C") Or (tmChf.sType = "V") Or ((tmChf.sType = "T") And (tgSpf.sSchdRemnant = "Y")) Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo = "Y"))) Then
            For ilClf = LBound(tgClfPostLog) To UBound(tgClfPostLog) - 1 Step 1
                tmClf = tgClfPostLog(ilClf).ClfRec
                If (tmClf.sType = "S") Or (tmClf.sType = "H") Then
                    If imVehicle = 0 Then
                        If tmClf.iVefCode = imVefCode Then
                            ilFound = True
                        Else
                            ilFound = False
                        End If
                    Else
                        ilFound = True
                    End If
                    If ilFound Then
                        gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLnSDate
                        gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llLnEDate
                        If (llLnEDate < llMdStartdate) Or (llLnSDate > llMdEndDate) Then
                            ilFound = False
                        End If
                    End If
                    If ilFound Then
                        ilFound = False
                        'For ilTest = LBound(tgMdSaveInfo) To UBound(tgMdSaveInfo) - 1 Step 1
                        For ilTest = LBONE To UBound(tgMdSaveInfo) - 1 Step 1
                            If (tgMdSaveInfo(ilTest).lChfCode = tmChf.lCode) And (tgMdSaveInfo(ilTest).iVefCode = tmClf.iVefCode) And (tgMdSaveInfo(ilTest).iLen = tmClf.iLen) Then
                                If (tgMdSaveInfo(ilTest).iRdfCode = tmClf.iRdfCode) And (tgMdSaveInfo(ilTest).iStartTime(0) = tmClf.iStartTime(0)) And (tgMdSaveInfo(ilTest).iStartTime(1) = tmClf.iStartTime(1)) And (tgMdSaveInfo(ilTest).iEndTime(0) = tmClf.iEndTime(0)) And (tgMdSaveInfo(ilTest).iEndTime(1) = tmClf.iEndTime(1)) Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilTest
                    Else
                        ilFound = True
                    End If
                    If Not ilFound Then
                        ilUpper = UBound(tgMdSaveInfo)
                        tgMdSaveInfo(ilUpper).lChfCode = tmChf.lCode
                        tgMdSaveInfo(ilUpper).iAdfCode = tmChf.iAdfCode
                        tgMdSaveInfo(ilUpper).lCntrNo = tmChf.lCntrNo
                        gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), tgMdSaveInfo(ilUpper).lEndDate
                        tgMdSaveInfo(ilUpper).iVefCode = tmClf.iVefCode
                        tgMdSaveInfo(ilUpper).iLineNo = tmClf.iLine
                        tgMdSaveInfo(ilUpper).iLen = tmClf.iLen
                        tgMdSaveInfo(ilUpper).lWkMissed = 0
                        tgMdSaveInfo(ilUpper).iRdfCode = tmClf.iRdfCode
                        tgMdSaveInfo(ilUpper).iStartTime(0) = tmClf.iStartTime(0)
                        tgMdSaveInfo(ilUpper).iStartTime(1) = tmClf.iStartTime(1)
                        tgMdSaveInfo(ilUpper).iEndTime(0) = tmClf.iEndTime(0)
                        tgMdSaveInfo(ilUpper).iEndTime(1) = tmClf.iEndTime(1)
                        For ilDay = 0 To 6 Step 1
                            tgMdSaveInfo(ilUpper).iDay(ilDay) = 0
                        Next ilDay
                        tgMdSaveInfo(ilUpper).iCancelCount = 0
                        tgMdSaveInfo(ilUpper).iHiddenCount = 0
                        tgMdSaveInfo(ilUpper).iMissedCount = 0
                        tgMdSaveInfo(ilUpper).iFirstIndex = -1
                        ReDim Preserve tgMdSaveInfo(0 To ilUpper + 1) As MDSAVEINFO
                    End If
                End If
            Next ilClf
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMissedPop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Missed list box if    *
'*                      required                       *
'*                                                     *
'*******************************************************
Private Sub mMissedPop()
'
'   mMissedPop
'   Where:
'
    ReDim ilfilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slDefaultName As String
    Dim slStr As String
    
    If imMissedType <> 2 Then
        ilIndex = lbcMissed.ListIndex
        If ilIndex >= 0 Then
            slDefaultName = lbcMissed.List(ilIndex)
        Else
            slDefaultName = "Missed"
            slStr = ""
            'ReDim tgMRMnf(1 To 1) As MNF
            ReDim tgMRMnf(0 To 0) As MNF
            ilRet = gObtainMnfForType("M", slStr, tgMRMnf())
            For ilLoop = LBound(tgMRMnf) To UBound(tgMRMnf) - 1 Step 1
                If Trim$(tgMRMnf(ilLoop).sUnitType) = "Y" Then
                    slDefaultName = Trim$(tgMRMnf(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
        End If
        'Repopulate if required- if sales source changed by another user while in this screen
        
        ilfilter(0) = CHARFILTER
        slFilter(0) = "M"
        ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
        ilfilter(1) = CHARFILTERNOT
        slFilter(1) = "A"
        ilOffSet(1) = gFieldOffset("Mnf", "MnfCodeStn")
        'ilRet = gIMoveListBox(PostLog, lbcMissed, lbcMissedCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(PostLog, lbcMissed, tmMissedCode(), smMissedCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mMissedPopErr
            gCPErrorMsg ilRet, "mMissedPop (gIMoveListBox)", PostLog
            On Error GoTo 0
            imChgMode = True
            gFindMatch slDefaultName, 0, lbcMissed
            If gLastFound(lbcMissed) >= 0 Then
                lbcMissed.ListIndex = gLastFound(lbcMissed)
            Else
                lbcMissed.ListIndex = -1
            End If
            imChgMode = False
        End If
    Else
        smMissedCodeTag = ""
        lbcMissed.Clear
        lbcMissed.AddItem "[Hide]"
    End If
    Exit Sub
mMissedPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveTest                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if spot can be moved *
'*                      into specified location        *
'*                                                     *
'*******************************************************
Private Function mMoveTest(slSchDate As String, slMoveDate As String, slMoveTime As String) As String
'
'   slRet = mMoveTest(slSchDate, slMoveDate, slMoveTime)
'       Where:
'           tmSdf (I)- Contains spot
'
'           slRet(O)-   "" = Abort move
'                       "S"=Move
'                       "G"=Move as MG
'                       "O"=Move and set as moved outside contract limits
'
    Dim ilMoveDay As Integer
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilMGMove As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim slMsg As String
    Dim ilWeekMoveOk As Integer
    Dim ilDayMoveOk As Integer
    Dim ilTimeMoveOk As Integer
    Dim llChfCode As Long
    Dim ilLineNo As Integer
    Dim ilAdfCode As Integer
    Dim ilVehComp As Integer
    Dim slDate As String
    Dim slType As String
    Dim slWkDate As String
    Dim llTime As Long
    Dim llMoDate As Long
    Dim llSuDate As Long
    Dim ilGameNo As Integer
    Dim tlSmf As SMF
    ReDim tlCff(0 To 1) As CFF
    Dim llEarliestAllowedDate As Long
    Dim slMissedDate As String

    Screen.MousePointer = vbHourglass
    ilWeekMoveOk = True
    ilDayMoveOk = True
    ilTimeMoveOk = True
    llTime = CLng(gTimeToCurrency(slMoveTime, False))
    ilGameNo = tmSdf.iGameNo
    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
        llMoDate = 0
        If gFindSmf(tmSdf, hmSmf, tlSmf) Then
            ilGameNo = tlSmf.iGameNo
            gUnpackDate tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), slWkDate
            slMissedDate = slWkDate
            llMoDate = gDateValue(gObtainPrevMonday(slWkDate))
            llSuDate = gDateValue(gObtainNextSunday(slWkDate))
        End If
        'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
        'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
        'tmSmfSrchKey.iMissedDate(0) = 0 'sch date =tlSdf.iDate(0)
        'tmSmfSrchKey.iMissedDate(1) = 0 'sch date =tlSdf.iDate(1)
        'imSmfRecLen = Len(tlSmf)
        'ilRet = btrGetGreaterOrEqual(hmSmf, tlSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
        'Do While (ilRet = BTRV_ERR_NONE) And (tlSmf.lChfCode = tmSdf.lChfCode) And (tlSmf.iLineNo = tmSdf.iLineNo)
        '    If (tlSmf.sSchStatus = tmSdf.sSchStatus) And (tlSmf.iActualDate(0) = tmSdf.iDate(0)) And (tlSmf.iActualDate(1) = tmSdf.iDate(1)) And (tlSmf.iActualTime(0) = tmSdf.iTime(0)) And (tlSmf.iActualTime(1) = tmSdf.iTime(1)) Then
        '        gUnpackDate tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), slWkDate
        '        llMoDate = gDateValue(gObtainPrevMonday(slWkDate))
        '        llSuDate = gDateValue(gObtainNextSunday(slWkDate))
        '        Exit Do
        '    End If
        '    ilRet = btrGetNext(hmSmf, tlSmf, imSmfRecLen, BTRV_LOCK_NONE)
        'Loop
        If llMoDate = 0 Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Read error(Smf)- Move ignored", vbOKOnly + vbQuestion, "Error")
            mMoveTest = ""  'Abort
            Exit Function
        End If
    Else
        'gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slWkDate
        slWkDate = slSchDate
        llMoDate = gDateValue(gObtainPrevMonday(slWkDate))
        llSuDate = gDateValue(gObtainNextSunday(slWkDate))
    End If
    slType = tmSdf.sSpotType
    ilMoveDay = gWeekDayStr(slMoveDate)
    llDate = gDateValue(slMoveDate)
    ilRet = mReadChfClfRdfCffRec(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.lFsfCode, ilGameNo, slMoveDate)
    If Not ilRet Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Read error- Move ignored", vbOKOnly + vbQuestion, "Error")
        mMoveTest = ""  'Abort
        Exit Function
    End If
    llDate = gDateValue(slMoveDate) 'smSelectedDate)
    tlCff(0) = tmCff
    llEarliestAllowedDate = 0
    gGetLineSchParameters hmSsf, tmSsf(), lmSsfDate(), lmSsfRecPos(), llDate, imVefCode, tmChf.iAdfCode, ilGameNo, tlCff(), tmClf, tmLnRdf, lmSepLength, lmStartDateLen, lmEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), llEarliestAllowedDate, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, True, imPriceLevel, False
    'If cff not found, then spot is outside date definition
    ilMGMove = 0
    If tmCff.sDelete = "Y" Then
        ilWeekMoveOk = False
    Else
        'Test if within same week
        If (llDate < llMoDate) Or (llDate > llSuDate) Then
            ilWeekMoveOk = False
        End If
        'Test days
        If (tmCff.iSpotsWk > 0) Or (tmCff.iXSpotsWk > 0) Then 'Weekly
            If (tmCff.iDay(ilMoveDay) <= 0) And (tmCff.sXDay(ilMoveDay) <> "Y") Then
                ilDayMoveOk = False
            End If
        Else
            If tmCff.iDay(ilMoveDay) <= 0 Then
                ilDayMoveOk = False
            End If
            If ((tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) And (slMissedDate <> "") Then
                slDate = slMissedDate
            Else
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
            End If
            If gDateValue(slDate) <> gDateValue(slMoveDate) Then
                ilDayMoveOk = False
            End If
        End If
    End If
    'Check Times
    ilFound = False
    For ilLoop = LBound(lmTBStartTime) To UBound(lmTBEndTime) Step 1
        If lmTBStartTime(ilLoop) <> -1 Then
            If (llTime >= lmTBStartTime(ilLoop)) And (llTime < lmTBEndTime(ilLoop)) Then
                ilFound = True
                Exit For
            End If
        End If
    Next ilLoop
    If Not ilFound Then
        ilTimeMoveOk = False
    End If
    slMsg = ""
    If (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
        If tmClf.iVefCode <> imVefCode Then
            slMsg = "Move violates Vehicle"
        End If
        If Not ilWeekMoveOk Then
            If slMsg = "" Then
                slMsg = "Move violates Weeks"
            Else
                slMsg = slMsg & ", weeks"
            End If
        End If
        If Not ilTimeMoveOk Then
            If slMsg = "" Then
                slMsg = "Move violates Times"
            Else
                slMsg = slMsg & ", times"
            End If
        End If
        If Not ilDayMoveOk Then
            If slMsg = "" Then
                slMsg = "Move violates Days"
            Else
                slMsg = slMsg & ", days"
            End If
        End If
    Else
        If (tmSdf.sSchStatus = "G") Then
            If slType = "X" Then
                ilMGMove = vbYes
            End If
            If tmClf.iVefCode <> imVefCode Then
                ilMGMove = vbYes
            End If
            If Not ilWeekMoveOk Then
                ilMGMove = vbYes
            End If
            If Not ilTimeMoveOk Then
                ilMGMove = vbYes
            End If
            If Not ilDayMoveOk Then
                ilMGMove = vbYes
            End If
        Else
            If slType = "X" Then
                ilMGMove = vbNo
            End If
            If tmClf.iVefCode <> imVefCode Then
                ilMGMove = vbNo
            End If
            If Not ilWeekMoveOk Then
                ilMGMove = vbNo
            End If
            If Not ilTimeMoveOk Then
                ilMGMove = vbNo
            End If
            If Not ilDayMoveOk Then
                ilMGMove = vbNo
            End If
        End If
    End If
    If slMsg <> "" Then
        If ((slType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((slType = "M") And (tgSpf.sSchdPromo = "Y")) Or ((slType = "T") And (tgSpf.sSchdRemnant = "Y")) Then
            slType = "A"
        End If
        If (slType <> "S") And (slType <> "M") And (slType <> "T") And (slType <> "Q") Then
            If slType <> "X" Then
                If (tgSpf.sPLMove <> "M") And (tgSpf.sPLMove <> "O") Then  'ask
                    Screen.MousePointer = vbDefault
                    'ilMGMove = MsgBox(slMsg & ": Move as MG", vbYesNoCancel + vbQuestion, "Limits")
                    'If ilMGMove = vbCancel Then
                    '    mMoveTest = ""  'Abort
                    '    Exit Function
                    'End If
                    sgGenMsg = Trim$(slMsg) & ": Move as MG, Outside or Cancel"
                    sgCMCTitle(0) = "MG"
                    sgCMCTitle(1) = "Outside"
                    sgCMCTitle(2) = "Cancel"
                    sgCMCTitle(3) = ""
                    igDefCMC = 0
                    igEditBox = 0
                    GenMsg.Show vbModal
                    If igAnsCMC = 0 Then
                        ilMGMove = vbYes
                    ElseIf igAnsCMC = 1 Then
                        ilMGMove = vbNo
                    Else
                        ilMGMove = vbCancel
                    End If
                    If ilMGMove = vbCancel Then
                        mMoveTest = ""  'Abort
                        Exit Function
                    End If
                Else
                    'Screen.MousePointer = vbDefault
                    ''ilMGMove = MsgBox(slMsg & ": Move as MG", vbOkCancel + vbQuestion + vbDefaultButton2, "Limits")
                    ''If ilMGMove = vbCancel Then
                    ''    mMoveTest = ""  'Abort
                    ''    Exit Function
                    ''End If
                    'sgGenMsg = Trim$(slMsg) & ": Move as MG or Cancel"
                    'sgCMCTitle(0) = "MG"
                    'sgCMCTitle(1) = "Cancel"
                    'sgCMCTitle(2) = ""
                    'igDefCMC = 0
                    'igEditBox = 0
                    'GenMsg.Show vbModal
                    'If igAnsCMC = 0 Then
                    '    ilMGMove = vbYes
                    'Else
                    '    ilMGMove = vbCancel
                    'End If
                    'If ilMGMove = vbCancel Then
                    '    mMoveTest = ""  'Abort
                    '    Exit Function
                    'End If
                    If tgSpf.sPLMove = "M" Then
                        ilMGMove = vbYes       'MG
                    Else
                        ilMGMove = vbNo        'Outside
                    End If
                End If
            Else
                ilMGMove = vbNo
            End If
        Else
            ilMGMove = vbNo
        End If
    End If
    If (ilMGMove = vbYes) Or (ilMGMove = vbOK) Then
        mMoveTest = "G"
    ElseIf ilMGMove = vbNo Then
        mMoveTest = "O"
    Else
        mMoveTest = "S"
    End If
    Screen.MousePointer = vbDefault
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainAvailTime                *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain time spot air           *
'*                                                     *
'*******************************************************
Private Function mObtainAvailTime(slSchDate As String, slDefaultTime As String, slAirDate As String, slAirTime As String) As String
    Dim ilRet As Integer       'Call return value
    Dim slNameCode As String
    Dim slCode As String
    igAvailVefCode = imVefCode
    igAvailGameNo = 0
    If tmVef.sType <> "G" Then
        If Trim$(slSchDate) <> "" Then
            sgAvailDate = slSchDate 'smSelectedDate
        Else
            sgAvailDate = tgDates(imDateSelectedIndex).sDate 'cbcDate.List(imDateSelectedIndex)
        End If
    Else
        sgAvailDate = tgDates(imDateSelectedIndex).sDate
        igAvailGameNo = tgDates(imDateSelectedIndex).iGameNo
    End If
    sgAvailTime = slDefaultTime
    If imAvailSelectedIndex <= 0 Then
        igAvailAnfCode = -1
    Else
        slNameCode = tmAvailCode(imAvailSelectedIndex - 1).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        igAvailAnfCode = Val(slCode)
    End If
    If rbcType(1).Value Then
        igPLogSource = 1
    Else
        igPLogSource = 0
    End If
    PLogTime.Show vbModal
    slAirDate = sgAvailDate
    slAirTime = sgPLogAirTime
    mObtainAvailTime = sgAvailTime
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadCffRec                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadCffRec(slLnStartDate As String, slLnEndDate As String, slNoSpots As String, llPrice As Long, ilDays() As Integer) As Integer
'
'   iRet = mReadCffRec(ilClfIndex)
'   Where:
'       ilClfIndex (I) - CLF index (starting at 0)
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llSpotDate As Long
    Dim ilNoSpots As Integer
    slLnStartDate = ""
    slLnEndDate = ""
    slNoSpots = ""
    llStartDate = 0
    llEndDate = 0
    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slStartDate    'Week Start date
    llSpotDate = gDateValue(slStartDate)
    tmCffSrchKey.lChfCode = tmChf.lCode
    tmCffSrchKey.iClfLine = tmClf.iLine
    tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
    tmCffSrchKey.iPropVer = tmClf.iPropVer
    tmCffSrchKey.iStartDate(0) = 0
    tmCffSrchKey.iStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCff.lChfCode = tmChf.lCode) And (tmCff.iClfLine = tmClf.iLine)
        If (tmCff.iCntRevNo = tmClf.iCntRevNo) And (tmCff.iPropVer = tmClf.iPropVer) And (tmCff.sDelete <> "Y") Then
            gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStartDate    'Week Start date
            gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slEndDate    'Week Start date
            If llStartDate = 0 Then
                llStartDate = gDateValue(slStartDate)
                llEndDate = gDateValue(slEndDate)
            Else
                If gDateValue(slStartDate) < llStartDate Then
                    llStartDate = gDateValue(slStartDate)
                End If
                If gDateValue(slEndDate) > llEndDate Then
                    llEndDate = gDateValue(slEndDate)
                End If
            End If
            If (llSpotDate >= gDateValue(slStartDate)) And (llSpotDate <= gDateValue(slEndDate)) Then
                ilNoSpots = 0
                If (tmCff.iSpotsWk <> 0) Or (tmCff.iXSpotsWk <> 0) Then 'Weekly
                    ilNoSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                    For ilLoop = 0 To 6 Step 1
                        If tmCff.iDay(ilLoop) > 0 Then
                            ilDays(ilLoop) = True
                        Else
                            ilDays(ilLoop) = False
                        End If
                    Next ilLoop
                Else    'Daily
                    For ilLoop = 0 To 6 Step 1
                        ilNoSpots = ilNoSpots + tmCff.iDay(ilLoop)
                        If tmCff.iDay(ilLoop) > 0 Then
                            ilDays(ilLoop) = True
                        Else
                            ilDays(ilLoop) = False
                        End If
                    Next ilLoop
                End If
                slNoSpots = Trim$(str$(ilNoSpots))
                llPrice = tmCff.lActPrice
            End If
        End If
        ilRet = btrGetNext(hmCff, tmCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If llStartDate > 0 Then
        slLnStartDate = Format$(llStartDate, "m/d/yy")
        slLnEndDate = Format$(llEndDate, "m/d/yy")
    End If
    mReadCffRec = True
    Exit Function

    On Error GoTo 0
    mReadCffRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfClfRdfCffRec            *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Private Function mReadChfClfRdfCffRec(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long, ilGameNo As Integer, slSpotDate As String) As Integer
'
'   iRet = mReadChfClfRdpfCffRec(llChfCode, ilLineNo, slMissedDate, SlStartDate, slEndDate, slNoSpots)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       slMissedDate(I)- Missed date or date to find bracketing week
'       tmCff(O)- contains valid flight week (if sDelete = "Y", then week is invalid)
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llSpotDate As Long
    Dim ilDay As Integer
    Dim ilIndex As Integer
    Dim ilLastDay As Integer
    Dim ilFirstDay As Integer
    Dim tlCff As CFF
    ilDay = gWeekDayStr(slSpotDate)
    ilLastDay = -1
    ilFirstDay = -1
    tmCff.sDelete = "Y"  'Set as flag that illegal week
    If mReadChfClfRdfRec(llChfCode, ilLineNo, llFsfCode) Then
        llStartDate = 0
        llEndDate = 0
        llSpotDate = gDateValue(slSpotDate)
        If llChfCode > 0 Then
            If ilGameNo = 0 Then
                tmCffSrchKey.lChfCode = llChfCode
                tmCffSrchKey.iClfLine = ilLineNo
                tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
                tmCffSrchKey.iPropVer = tmClf.iPropVer
                tmCffSrchKey.iStartDate(0) = 0
                tmCffSrchKey.iStartDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hmCff, tlCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tlCff.lChfCode = llChfCode) And (tlCff.iClfLine = ilLineNo)
                    If (tlCff.iCntRevNo = tmClf.iCntRevNo) And (tlCff.iPropVer = tmClf.iPropVer) And (tlCff.sDelete <> "Y") Then
                        gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStartDate    'Week Start date
                        gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slEndDate    'Week Start date
                        If llStartDate = 0 Then
                            llStartDate = gDateValue(slStartDate)
                            llEndDate = gDateValue(slEndDate)
                        Else
                            If gDateValue(slStartDate) < llStartDate Then
                                llStartDate = gDateValue(slStartDate)
                            End If
                            If gDateValue(slEndDate) > llEndDate Then
                                llEndDate = gDateValue(slEndDate)
                            End If
                        End If
                        If (llSpotDate >= gDateValue(slStartDate)) And (llSpotDate <= gDateValue(slEndDate)) Then
                            tmCff = tlCff
                            If (tmCff.iSpotsWk <> 0) Or (tmCff.iXSpotsWk <> 0) Then 'Weekly
                                For ilIndex = 0 To 6 Step 1
                                    If tmCff.iDay(ilIndex) > 0 Then
                                        If ilFirstDay = -1 Then
                                            ilFirstDay = ilIndex
                                        End If
                                        ilLastDay = ilIndex
                                    End If
                                Next ilIndex
                                If (ilDay < ilFirstDay) Or (ilDay > ilLastDay) Then
                                    ilLastDay = -1
                                    ilFirstDay = -1
                                    For ilIndex = 0 To 6 Step 1
                                        If tmCff.sXDay(ilIndex) = "Y" Then
                                            If ilFirstDay = -1 Then
                                                ilFirstDay = ilIndex
                                            End If
                                            ilLastDay = ilIndex
                                        End If
                                    Next ilIndex
                                End If
                            Else    'Daily
                                ilFirstDay = ilDay
                                ilLastDay = ilDay
                            End If
                            Exit Do
                        End If
                    End If
                    ilRet = btrGetNext(hmCff, tlCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            Else
                tmCgfSrchKey1.lClfCode = tmClf.lCode
                ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lCode = tmCgf.lClfCode)
                    If tmCgf.iGameNo = ilGameNo Then
                        gCgfToCff tmClf, tmCgf, tmCgfCff()
                        tmCff = tmCgfCff(0) 'tmCgfCff(1)
                        Exit Do
                    End If
                    ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                Erase tmCgfCff
            End If
        End If
        'Determine times
        For ilLoop = LBound(lmTBStartTime) To UBound(lmTBStartTime) Step 1
            lmTBStartTime(ilLoop) = -1
            lmTBEndTime(ilLoop) = -1
        Next ilLoop
        mReadChfClfRdfCffRec = True
        'set of lmTBStartTime and lmTBEndTime are now set in mMoveTest which calls the function

        'If (tmCff.sDelete <> "Y") And (ilFirstDay <> -1) Then
        '    ilTBIndex = 1
        '    If (tmLnRdf.iLtfCode(0) <> 0) Or (tmLnRdf.iLtfCode(1) <> 0) Or (tmLnRdf.iLtfCode(2) <> 0) Then
        '        'Read Ssf for date- test for library- code removed- as Ssf not read into memory
        '        'this can be added if required
        '        'See gGetLineSchParameters for code
        '        'For now set time as 12m-12m
        '        lmTBStartTime(ilTBIndex) = 0
        '        lmTBEndTime(ilTBIndex) = 86400  '24*3600
        '    Else    'Time buy- check if override times defined (if so, use them as bump times)
        '        If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
        '            For ilLoop = LBound(tmLnRdf.iStartTime, 2) To UBound(tmLnRdf.iStartTime, 2) Step 1
        '                If (tmLnRdf.iStartTime(0, ilLoop) <> 1) Or (tmLnRdf.iStartTime(1, ilLoop) <> 0) Then
        '                    If (tmCff.iSpotsWk = 0) And (tmCff.iXSpotsWk = 0) Then 'Daily- Test if valid day
        '                        If tmLnRdf.sWkDays(ilLoop, ilDay + 1) = "Y" Then
        '                            gUnpackTime tmLnRdf.iStartTime(0, ilLoop), tmLnRdf.iStartTime(1, ilLoop), "A", "1", slLnStart
        '                            gUnpackTime tmLnRdf.iEndTime(0, ilLoop), tmLnRdf.iEndTime(1, ilLoop), "A", "1", slLnEnd
        '                            lmTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
        '                            lmTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
        '                            ilTBIndex = ilTBIndex + 1
        '                        End If
        '                    Else    'Add time for each valid day
        '                        For ilIndex = ilFirstDay To ilLastDay Step 1
        '                            If (tmCff.iDay(ilIndex) = 1) Or (tmCff.sXDay(ilIndex) = "Y") Then
        '                                If tmLnRdf.sWkDays(ilLoop, ilIndex + 1) = "Y" Then
        '                                    gUnpackTime tmLnRdf.iStartTime(0, ilLoop), tmLnRdf.iStartTime(1, ilLoop), "A", "1", slLnStart
        '                                    gUnpackTime tmLnRdf.iEndTime(0, ilLoop), tmLnRdf.iEndTime(1, ilLoop), "A", "1", slLnEnd
        '                                    lmTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
        '                                    lmTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
        '                                    ilTBIndex = ilTBIndex + 1
        '                                End If
        '                            End If
        '                        Next ilIndex
        '                    End If
         '               End If
        '            Next ilLoop
        '        Else
        '            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slLnStart
        '            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slLnEnd
        '            lmTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
        '            lmTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
        '        End If
        '    End If
        '    mReadChfClfRdfCffRec = True
        'Else
        '    mReadChfClfRdfCffRec = True
        'End If
    Else
        mReadChfClfRdfCffRec = False
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfClfRdfRec               *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Private Function mReadChfClfRdfRec(llChfCode As Long, ilLineNo As Integer, llFsfCode As Long) As Integer
'
'   iRet = mReadChfClfRdfRec(llChfCode, ilLineNo)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    If llChfCode > 0 Then
        'If llChfCode <> tmChf.lCode Then
            tmChfSrchKey.lCode = llChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                mReadChfClfRdfRec = False
                Exit Function
            End If
        'End If
        'If (tmClf.lChfCode <> llChfCode) Or (tmClf.iLine <> ilLineNo) Then
            tmClfSrchKey.lChfCode = llChfCode
            tmClfSrchKey.iLine = ilLineNo
            tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
            tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))    'And (tmClf.sSchStatus = "A")
                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        'Else
        '    ilRet = BTRV_ERR_NONE
        'End If
        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) Then
            If tmLnRdf.iCode <> tmClf.iRdfCode Then
                tmRdfSrchKey.iCode = tmClf.iRdfCode  ' Rate card program/time File Code
                ilRet = btrGetEqual(hmRdf, tmLnRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    mReadChfClfRdfRec = False
                    Exit Function
                End If
            End If
            mReadChfClfRdfRec = True
        Else
            mReadChfClfRdfRec = False
        End If
    Else
        tmFSFSrchKey.lCode = llFsfCode
        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        gMoveFeedToCntr tmFsf, tmLnRdf, tmChf, tmClf, tmFCff(), hmFnf, hmPrf
        tmCff = tmFCff(1)
        mReadChfClfRdfRec = True
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadMdSdfRec                   *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read missed spot records       *
'*                                                     *
'*******************************************************
Private Function mReadMdSdfRec(ilGetAll As Integer) As Integer
'
'   iRet = mReadMdSdfRec()
'   Where:
'       iRet (O)- True if records read,
'                 False if not read
'
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slMdStartDate As String
    Dim slMdEndDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slNoSpots As String
    Dim ilLoop As Integer
    Dim ilPass As Integer
    Dim ilPassStart As Integer
    Dim ilPassEnd As Integer
    Dim slSchStatus As String
    Dim ilVefCode As Integer
    Dim ilDay As Integer
    Dim ilFound As Integer
    Dim llMdStartdate As Long
    Dim llMdEndDate As Long
    Dim ilVef As Integer
    Dim llDate As Long
    Dim llWkDate As Long
    Dim ilInclude As Integer
    Dim llPrice As Long
    ReDim ilDays(0 To 6) As Integer
'   Resize the arrays and empty them so that they can be refilled
'          from the file
'   smShow will hold all data for each line.  The data is trimmed
'          to fit the display controls
    'ReDim smMdShow(1 To 8, 1 To 1) As String
    'ReDim smMdSave(1 To 1, 1 To 1) As String
    'ReDim smMdSchStatus(1 To 1, 1 To 1) As String
    'ReDim lmMdRecPos(1 To 1) As Long
    slMdStartDate = edcMdDate(0).Text
    If Not gValidDate(slMdStartDate) Then
        mReadMdSdfRec = True
        Exit Function
    End If
    slMdEndDate = edcMdDate(1).Text
    If Not gValidDate(slMdEndDate) Then
        mReadMdSdfRec = True
        Exit Function
    End If

    'If (imWeek = imSvWeek) And (imVehicle = imSvVehicle) And (imSvVehSelectedIndex = imVehSelectedIndex) And (imSvDateSelectedIndex = imDateSelectedIndex) Then
    If (imVehicle = imSvVehicle) And (imSvVehSelectedIndex = imVehSelectedIndex) And (gDateValue(slMdStartDate) = lmMdStartdate) And (gDateValue(slMdEndDate) = lmMdEndDate) Then
        mReadMdSdfRec = True
        Exit Function
    End If
    pbcMissed.Cls
    ReDim tgMdSdfRec(0 To 1) As MDSDFREC
    ReDim tgMdSaveInfo(0 To 1) As MDSAVEINFO
    ReDim tgMdShowInfo(0 To 1) As MDSHOWINFO
    If (imDateSelectedIndex < 0) Or (imVehSelectedIndex < 0) Or (Not ilGetAll) Then
        vbcMissed.Min = LBONE   'LBound(tgMdShowInfo)
        vbcMissed.Max = LBONE   'LBound(tgMdShowInfo)
        mReadMdSdfRec = True
        Exit Function
    End If
    'imSvWeek = imWeek
    imSvVehicle = imVehicle
    imSvVehSelectedIndex = imVehSelectedIndex
    imSvDateSelectedIndex = imDateSelectedIndex
    lmMdStartdate = gDateValue(slMdStartDate)
    lmMdEndDate = gDateValue(slMdEndDate)
    'slDate = tgDates(imDateSelectedIndex).sDate 'cbcDate.List(imDateSelectedIndex)
    'If imWeek = 1 Then
    '    slMdStartDate = gObtainStartStd(slDate)
    '    slMdEndDate = gObtainNextSunday(slDate)
    'ElseIf imWeek = 2 Then
    '    slMdStartDate = "1/1/94"
    '    slMdEndDate = gObtainNextSunday(slDate)
    'Else
    '    slMdStartDate = gObtainPrevMonday(slDate)
    '    slMdEndDate = gObtainNextSunday(slDate)
    'End If
    llMdStartdate = gDateValue(slMdStartDate)
    llMdEndDate = gDateValue(slMdEndDate)
    For ilVef = 0 To cbcVeh.ListCount - 1 Step 1
        slNameCode = tmUserVehicle(ilVef).sKey    'Traffic!lbcUserVehicle.List(imVehSelectedIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        If (ilVefCode = imVefCode) Or (imVehicle = 1) Then
            ilPassStart = 0
            ilPassEnd = 2
            For ilPass = ilPassStart To ilPassEnd Step 1
                Select Case ilPass
                    Case 0
                        slSchStatus = "M"
                        ilInclude = True
                    Case 1
                        slSchStatus = "C"
                        If ckcInclude(1).Value = vbChecked Then
                            ilInclude = True
                        Else
                            ilInclude = False
                        End If
                    Case 2
                        slSchStatus = "H"
                        If ckcInclude(2).Value = vbChecked Then
                            ilInclude = True
                        Else
                            ilInclude = False
                        End If
               End Select
               If ilInclude Then
                    'ilAdfCode = 0
                    'ilVefCode = imVefCode
                    'ilSortOrder = -1
                    'ReDim tmSdfMdExt(1 To 1) As SDFMDEXT
                    'ilRet = gObtainMissedSpot(slSchStatus, ilVefCode, ilAdfCode, slMdStartDate, slMdEndDate, ilSortOrder, lbcSortCtrl, tlSdfMdExt())
                    tmSdfSrchKey.iVefCode = ilVefCode
                    gPackDate slMdStartDate, tmSdfSrchKey.iDate(0), tmSdfSrchKey.iDate(1)
                    ilDate0 = tmSdfSrchKey.iDate(0)
                    ilDate1 = tmSdfSrchKey.iDate(1)
                    tmSdfSrchKey.iTime(0) = 0
                    tmSdfSrchKey.iTime(1) = 0
                    tmSdfSrchKey.sSchStatus = slSchStatus
                    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode)
                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                        If llDate > llMdEndDate Then
                            Exit Do
                        End If
                        If (tmSdf.sSchStatus = slSchStatus) And ((tmSdf.sBill <> "Y") Or (ckcInclude(3).Value = vbChecked)) Then
                            If (tmSdf.sSpotType = "A") Or (tmSdf.sSpotType = "B") Or (tmSdf.sSpotType = "D") Or (tmSdf.sSpotType = "M") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Y") Then
                                tgMdSdfRec(UBound(tgMdSdfRec)).lSdfCode = tmSdf.lCode
                                ilRet = btrGetPosition(hmSdf, tgMdSdfRec(UBound(tgMdSdfRec)).lSdfRecPos)
                                tgMdSdfRec(UBound(tgMdSdfRec)).lMissedDate = llDate
                                tgMdSdfRec(UBound(tgMdSdfRec)).sSchStatus = tmSdf.sSchStatus
                                If tmSdf.lChfCode > 0 Then
                                    tmChfSrchKey.lCode = tmSdf.lChfCode  ' Contract Hdr File Code
                                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If (ilRet = BTRV_ERR_NONE) Then
                                        tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                        tmClfSrchKey.iLine = tmSdf.iLineNo
                                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Loop
                                    End If
                                    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                                        'tmRdfScrhKey.iCode = tmClf.iRdfCode  ' Rate card program/time File Code
                                        'ilRet = btrGetEqual(hmRdf, tmLnRdf, imRdfRecLen, tmRdfScrhKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        'On Error GoTo mReadMdSdfRecErr
                                        'gBtrvErrorMsg ilRet, "mReadMdSdfRec (btrGetEqual:Rate card Program/Time)", PostLog
                                        'On Error GoTo 0
                                        ilFound = False
                                        llWkDate = llDate
                                        Do While gWeekDayLong(llWkDate) <> 0
                                            llWkDate = llWkDate - 1
                                        Loop
                                        'For ilLoop = LBound(tgMdSaveInfo) To UBound(tgMdSaveInfo) - 1 Step 1
                                        For ilLoop = LBONE To UBound(tgMdSaveInfo) - 1 Step 1
                                            If (tgMdSaveInfo(ilLoop).lChfCode = tmSdf.lChfCode) And (tgMdSaveInfo(ilLoop).iVefCode = tmSdf.iVefCode) And (tgMdSaveInfo(ilLoop).iLen = tmSdf.iLen) Then
                                                If (tgMdSaveInfo(ilLoop).iRdfCode = tmClf.iRdfCode) And (tgMdSaveInfo(ilLoop).iStartTime(0) = tmClf.iStartTime(0)) And (tgMdSaveInfo(ilLoop).iStartTime(1) = tmClf.iStartTime(1)) And (tgMdSaveInfo(ilLoop).iEndTime(0) = tmClf.iEndTime(0)) And (tgMdSaveInfo(ilLoop).iEndTime(1) = tmClf.iEndTime(1)) Then
                                                    If (tgMdSaveInfo(ilLoop).lWkMissed = llWkDate) And (tgMdSaveInfo(ilLoop).sBill = tmSdf.sBill) Then
                                                        ilRet = mReadCffRec(slStartDate, slEndDate, slNoSpots, llPrice, ilDays())
                                                        If ilRet Then
                                                            ilFound = True
                                                            For ilDay = 0 To 6 Step 1
                                                                If tgMdSaveInfo(ilLoop).iDay(ilDay) <> ilDays(ilDay) Then
                                                                    ilFound = False
                                                                    Exit For
                                                                End If
                                                            Next ilDay
                                                            If ilFound Then
                                                                If slSchStatus = "C" Then
                                                                    tgMdSaveInfo(ilLoop).iCancelCount = tgMdSaveInfo(ilLoop).iCancelCount + 1
                                                                ElseIf slSchStatus = "H" Then
                                                                    tgMdSaveInfo(ilLoop).iHiddenCount = tgMdSaveInfo(ilLoop).iHiddenCount + 1
                                                                Else
                                                                    tgMdSaveInfo(ilLoop).iMissedCount = tgMdSaveInfo(ilLoop).iMissedCount + 1
                                                                End If
                                                                tgMdSdfRec(UBound(tgMdSdfRec)).lPrice = llPrice
                                                                tgMdSdfRec(UBound(tgMdSdfRec)).iNextIndex = tgMdSaveInfo(ilLoop).iFirstIndex
                                                                tgMdSaveInfo(ilLoop).iFirstIndex = UBound(tgMdSdfRec)
                                                                ReDim Preserve tgMdSdfRec(0 To UBound(tgMdSdfRec) + 1) As MDSDFREC
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next ilLoop
                                    End If
                                Else
                                    ilRet = mReadChfClfRdfRec(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.lFsfCode)
                                    ilFound = False
                                End If
                                If Not ilFound Then
                                    ilUpper = UBound(tgMdSaveInfo)
                                    ilRet = mReadCffRec(slStartDate, slEndDate, slNoSpots, llPrice, ilDays())
                                    tgMdSaveInfo(ilUpper).lChfCode = tmSdf.lChfCode
                                    tgMdSaveInfo(ilUpper).lFsfCode = tmSdf.lFsfCode
                                    tgMdSaveInfo(ilUpper).iAdfCode = tmSdf.iAdfCode
                                    tgMdSaveInfo(ilUpper).lCntrNo = tmChf.lCntrNo
                                    gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), tgMdSaveInfo(ilUpper).lEndDate
                                    tgMdSaveInfo(ilUpper).iVefCode = tmSdf.iVefCode
                                    tgMdSaveInfo(ilUpper).iLineNo = tmClf.iLine
                                    tgMdSaveInfo(ilUpper).iLen = tmSdf.iLen
                                    tgMdSaveInfo(ilUpper).lWkMissed = llWkDate
                                    tgMdSaveInfo(ilUpper).iRdfCode = tmClf.iRdfCode
                                    tgMdSaveInfo(ilUpper).iStartTime(0) = tmClf.iStartTime(0)
                                    tgMdSaveInfo(ilUpper).iStartTime(1) = tmClf.iStartTime(1)
                                    tgMdSaveInfo(ilUpper).iEndTime(0) = tmClf.iEndTime(0)
                                    tgMdSaveInfo(ilUpper).iEndTime(1) = tmClf.iEndTime(1)
                                    For ilDay = 0 To 6 Step 1
                                        tgMdSaveInfo(ilUpper).iDay(ilDay) = ilDays(ilDay)
                                    Next ilDay
                                    If slSchStatus = "C" Then
                                        tgMdSaveInfo(ilUpper).iCancelCount = 1
                                    ElseIf slSchStatus = "H" Then
                                        tgMdSaveInfo(ilUpper).iHiddenCount = 1
                                    Else
                                        tgMdSaveInfo(ilUpper).iMissedCount = 1
                                    End If
                                    tgMdSaveInfo(ilUpper).sBill = tmSdf.sBill
                                    tgMdSaveInfo(ilUpper).iFirstIndex = UBound(tgMdSdfRec)
                                    ReDim Preserve tgMdSaveInfo(0 To ilUpper + 1) As MDSAVEINFO
                                    tgMdSdfRec(UBound(tgMdSdfRec)).lPrice = llPrice
                                    tgMdSdfRec(UBound(tgMdSdfRec)).iNextIndex = -1
                                    ReDim Preserve tgMdSdfRec(0 To UBound(tgMdSdfRec) + 1) As MDSDFREC
                                End If
                            End If
                        End If
                        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_KEY_NOT_FOUND) Then
                        On Error GoTo mReadMdSdfRecErr
                        gBtrvErrorMsg ilRet, "mReadMdSdfRec (btrGetEqual)", PostLog
                        On Error GoTo 0
                    End If
               End If
            Next ilPass
        End If
    Next ilVef
    'If imWeek = 2 Then
    '    'If Current + Past selected, only use Current Month for contracts
    '    'Showing old contracts is of no value
    '    slDate = tgDates(imDateSelectedIndex).sDate 'cbcDate.List(imDateSelectedIndex)
    '    slMdStartDate = gObtainStartStd(slDate)
    '    slMdEndDate = gObtainNextSunday(slDate)
    '    llMdStartdate = gDateValue(slMdStartDate)
    '    llMdEndDate = gDateValue(slMdEndDate)
    'End If
    mMdCntrPop llMdStartdate, llMdEndDate
    'Create Show Records
    mCreateMdShow False
    mReadMdSdfRec = True
    Exit Function
mReadMdSdfRecErr:
    On Error GoTo 0
    mReadMdSdfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadSdfRec                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read spot records              *
'*                                                     *
'*******************************************************
Private Function mReadSdfRec(ilGetAll As Integer) As Integer
'
'   iRet = mReadSbfRec(ilGetAll)
'   Where:
'       ilGetAll(I) - True means Read All Records
'                     False means Only get 1 screens worth of rows
'       iRet (O)- True if records read,
'                 False if not read
'
    Dim slDate As String
    Dim slStr As String
    Dim ilRet As Integer    'Return status
    Dim ilUpperBound As Integer
    Dim ilNumOfrecToRead As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilSpot As Integer
    Dim ilEvt As Integer
    Dim slAvDate As String
    Dim slAvTime As String
    Dim slTime As String
    Dim llStartDate As Long
    Dim llSunDate As Long
    Dim llDate As Long
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilAnfCode As Integer
    Dim ilDay As Integer
    Dim ilUnits As Integer
    Dim ilSec As Integer
    Dim slUnits As String
    Dim slSpotLen As String
    Dim ilType As Integer
    Dim ilGameNo As Integer

    If ilGetAll And imGetAll Then
        Screen.MousePointer = vbHourglass
        ilRet = mReadMdSdfRec(ilGetAll)
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    imLastRowSelected = -1
    imGetAll = ilGetAll
'   Resize the arrays and empty them so that they can be refilled
'          from the file
'   smShow will hold all data for each line.  The data is trimmed
'          to fit the display controls
    'ReDim smShow(1 To 13, 1 To 1) As String
    ReDim tgShow(0 To 1) As SHOWINFO
'   smSave will hold only data that can be changed by the user.
'          This data is complete (not trimmed to fit the controls).
    'ReDim smSave(1 To 10, 1 To 1) As String
    ReDim tgSave(0 To 1) As SAVEINFO
    'ReDim imSave(1 To 2, 1 To 1) As Integer
    'ReDim imPostSpotInfo(1 To 4, 1 To 1) As Integer
    ilUpperBound = UBound(tgSave)   'UBound(smShow, 2)  ' Initialize to Start at 1
    If imDateSelectedIndex < 0 Or imVehSelectedIndex < 0 Then
        ilRet = mReadMdSdfRec(ilGetAll)
        mReadSdfRec = True
        Exit Function
    End If
    '  Build the file search key
    If ilGetAll Then
       ilNumOfrecToRead = 32000
    Else
       ilNumOfrecToRead = 2 * vbcPosting.LargeChange + 1
    End If
    Screen.MousePointer = vbHourglass  'Wait
    If tmVef.sType <> "G" Then
        slDate = tgDates(imDateSelectedIndex).sDate 'cbcDate.List(imDateSelectedIndex)
        'ilRet = gParseItem(slDate, 2, " ", slDate) 'Remove day
        'ilRet = gParseItem(slDate, 1, ":", slDate) 'Remove :xxxxxxx
        llStartDate = gDateValue(slDate)
        If imWM <> 1 Then
            slDate = gObtainNextSunday(slDate)
        Else
            slDate = gObtainEndStd(slDate)
        End If
        llSunDate = gDateValue(slDate)
        ilType = 0
    Else
        slDate = tgDates(imDateSelectedIndex).sDate
        llStartDate = gDateValue(slDate)
        llSunDate = llStartDate
        ilType = tgDates(imDateSelectedIndex).iGameNo
        'Moved to edcDate_Change
        'ckcDayComplete(0).Caption = "Game " & Trim$(Str$(ilType))
    End If
    ilGameNo = ilType
    For ilLoop = 0 To 6 Step 1
        lmLcfRecPos(ilLoop) = 0
    Next ilLoop
    For llDate = llStartDate To llSunDate Step 1
        ilFound = False
        For ilLoop = 0 To UBound(tgDates) - 1 Step 1
            If tgDates(ilLoop).lDate = llDate Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If ilFound Then
            If (tmVef.sType = "S") Then
                If (tgVpf(imVpfIndex).sBillSA = "Y") Then
                    ilDay = gWeekDayLong(llDate)
                    If ilDay < 5 Then
                        gObtainVlf "S", hmVLF, imVefCode, llDate, tmVlf0()
                    ElseIf ilDay = 5 Then
                        gObtainVlf "S", hmVLF, imVefCode, llDate, tmVlf6()
                    Else
                        gObtainVlf "S", hmVLF, imVefCode, llDate, tmVlf7()
                    End If
                End If
            End If
            slDate = Format$(llDate, "m/d/yy")
            smSelectedDate = slDate
            If tmVef.sType <> "G" Then
                imSelectedDay = gWeekDayStr(slDate)
            Else
                imSelectedDay = 0
            End If
            'If InStr(slDate, ":C") > 0 Then
            '    ilRet = gParseItem(slDate, 1, ":C", slDate)
            'Else
            '    ilRet = gParseItem(slDate, 1, ":I", slDate)
            'End If
            If rbcType(0).Value Then
                tmLcfSrchKey.iType = ilType
                tmLcfSrchKey.sStatus = "C"  ' Current
                tmLcfSrchKey.iVefCode = imVefCode
                gPackDate slDate, ilDate0, ilDate1
                tmLcfSrchKey.iLogDate(0) = ilDate0
                tmLcfSrchKey.iLogDate(1) = ilDate1
                tmLcfSrchKey.iSeqNo = 1
                ' Read the first matching record into tmLcf record structure. If a record exists, put
                '      it into slDate in the form of mm/dd/yy then
                '      it is converted into a number and put into llSchSDate
                ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                On Error GoTo mReadSdfRecErr
                gBtrvErrorMsg ilRet, "mReadSdfRec (btrGetEqual:LCF)", PostLog
                On Error GoTo 0
                ilRet = btrGetPosition(hmLcf, lmLcfRecPos(imSelectedDay))

                gPackDate slDate, ilDate0, ilDate1
                imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
                tmSsfSrchKey.iType = ilType 'slType-On Air
                tmSsfSrchKey.iVefCode = imVefCode
                tmSsfSrchKey.iDate(0) = ilDate0
                tmSsfSrchKey.iDate(1) = ilDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf(imSelectedDay).iType = ilType) And (tmSsf(imSelectedDay).iVefCode = imVefCode) And (tmSsf(imSelectedDay).iDate(0) = ilDate0) And (tmSsf(imSelectedDay).iDate(1) = ilDate1)
                    ilRet = gSSFGetPosition(hmSsf, lmSsfRecPos(imSelectedDay))
                    gUnpackDate tmSsf(imSelectedDay).iDate(0), tmSsf(imSelectedDay).iDate(1), slAvDate
                    ilEvt = 1
                    Do While ilEvt <= tmSsf(imSelectedDay).iCount
                       LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilEvt)
                        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                            If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                                ilUnits = tmAvail.iAvInfo And &H1F
                                slUnits = Trim$(str$(ilUnits)) & ".0"   'For units as thirty
                                ilSec = tmAvail.iLen
                            Else
                                ilUnits = tmAvail.iAvInfo And &H1F
                                ilSec = 0
                            End If
                            gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slAvTime
                            ilAnfCode = tmAvail.ianfCode
                            For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                ilEvt = ilEvt + 1
                               LSet tmSpot = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilEvt)
                                'Get Sdf
                                tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                '   Data is comming from the Spot Detail File (Sdf)
                                If (ilRet = BTRV_ERR_NONE) Then
                                    ''Don't show Feed; PSA or Promo spots
                                    ''If (tmSdf.sSpotType <> "N") And (tmSdf.sSpotType <> "S") And (tmSdf.sSpotType <> "R") Or (tmSdf.sSpotType = "M") Or (tmSdf.sSpotType = "Y") Then
                                    '4/4/07: Show remnant regardless if scheduled or not as remnants are always invoiced, ttp 2639
                                    'If (tmSdf.sSpotType = "A") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant = "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo = "Y")) Or (tmSdf.sSpotType = "Y") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "R") Or (tmSdf.sSpotType = "X") Then
                                    If (tmSdf.sSpotType = "A") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA = "Y")) Or (tmSdf.sSpotType = "T") Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo = "Y")) Or (tmSdf.sSpotType = "Y") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "R") Or (tmSdf.sSpotType = "X") Then
                                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                            ilUnits = ilUnits - 1
                                            ilSec = ilSec - (tmSpot.iPosLen And &HFFF)
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                            ilUnits = ilUnits - 1
                                            ilSec = ilSec - (tmSpot.iPosLen And &HFFF)
                                        ElseIf tgVpf(imVpfIndex).sSSellOut = "T" Then
                                            slSpotLen = Trim$(str$(tmSpot.iPosLen And &HFFF))
                                            slStr = gDivStr(slSpotLen, "30.0")
                                            slUnits = gSubStr(slUnits, slSpotLen)
                                        End If
                                        ilRet = mCreateSaveSpotImage(ilUpperBound, slAvDate, slAvTime, ilAnfCode, ilSpot * 10)
                                        If Not ilRet Then
                                            mReadSdfRec = False
                                            Screen.MousePointer = vbDefault
                                            Exit Function
                                        End If
                                        ' Get another row of space in the array just in case
                                        '     more records are available
                                        ilUpperBound = ilUpperBound + 1
                                        ReDim Preserve tgSave(0 To ilUpperBound) As SAVEINFO
                                        'Create avail record
                                        'If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Or (tgVpf(imVpfIndex).sSSellOut = "M") Then
                                        '    ilRet = mCreateSaveAvailImage(ilUpperBound, slAvDate, slAvTime, ilAnfCode, ilSec, ilUnits)
                                        '    ilUpperBound = ilUpperBound + 1
                                        '    ReDim Preserve tgSave(0 To ilUpperBound) As SAVEINFO
                                        'ElseIf tgVpf(imVpfIndex).sSSellOut = "T" Then
                                        '    ilRet = mCreateSaveAvailImage(ilUpperBound, slAvDate, slAvTime, ilAnfCode, 0, 0)
                                        '    ilUpperBound = ilUpperBound + 1
                                        '    ReDim Preserve tgSave(0 To ilUpperBound) As SAVEINFO
                                        'End If
                                        If ilUpperBound > ilNumOfrecToRead Then
                                            ilRet = 0
                                            Exit Do
                                        End If
                                    End If
                                End If
                            Next ilSpot
                        End If
                        ilEvt = ilEvt + 1
                    Loop
                    If ilUpperBound > ilNumOfrecToRead Then
                        ilRet = 0
                        Exit Do
                    End If
                    imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_KEY_NOT_FOUND) Then
                    On Error GoTo mReadSdfRecErr
                    gBtrvErrorMsg ilRet, "mReadSdfRec (btrGetEqual)", PostLog
                    On Error GoTo 0
                End If
                If tgSpf.sUsingBBs = "Y" Then
                    ReDim tlBBSdf(0 To 0) As SDF
                    ilRet = gGetBBSpots(hmSdf, imVefCode, ilGameNo, slDate, tlBBSdf())
                    For ilEvt = 0 To UBound(tlBBSdf) - 1 Step 1
                        tmSdfSrchKey3.lCode = tlBBSdf(ilEvt).lCode
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                        '   Data is comming from the Spot Detail File (Sdf)
                        If (ilRet = BTRV_ERR_NONE) Then
                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slAvDate
                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                            ilRet = mCreateSaveSpotImage(ilUpperBound, slAvDate, slAvTime, 0, 10)
                            If Not ilRet Then
                                mReadSdfRec = False
                                Screen.MousePointer = vbDefault
                                Exit Function
                            End If
                            ' Get another row of space in the array just in case
                            '     more records are available
                            ilUpperBound = ilUpperBound + 1
                            ReDim Preserve tgSave(0 To ilUpperBound) As SAVEINFO
                            If ilUpperBound > ilNumOfrecToRead Then
                                ilRet = 0
                                Exit For
                            End If
                        End If
                    Next ilEvt
                End If
            Else
                tmSdfSrchKey.iVefCode = imVefCode
                gPackDate smSelectedDate, tmSdfSrchKey.iDate(0), tmSdfSrchKey.iDate(1)
                ilDate0 = tmSdfSrchKey.iDate(0)
                ilDate1 = tmSdfSrchKey.iDate(1)
                tmSdfSrchKey.iTime(0) = 0
                tmSdfSrchKey.iTime(1) = 0
                tmSdfSrchKey.sSchStatus = "S"
                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (imVefCode = tmSdf.iVefCode) And (tmSdf.iDate(0) = ilDate0) And (tmSdf.iDate(1) = ilDate1)
                    'If (tmSdf.sSpotType <> "N") And (tmSdf.sSpotType <> "S") And (tmSdf.sSpotType <> "R") Or (tmSdf.sSpotType = "M") Or (tmSdf.sSpotType = "Y") Then

                    '4/4/07: Show remnant regardless if scheduled or not as remnants are always invoiced, ttp 2639
                    'If (tmSdf.sSpotType = "A") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant = "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo = "Y")) Or (tmSdf.sSpotType = "Y") Or (tmSdf.sSpotType = "R") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "X") Then
                    '2/21/13: Allow Package spots to be cancelled
                    If tmSdf.sSchStatus = "S" Then
                        If (tmSdf.sSpotType = "A") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA = "Y")) Or (tmSdf.sSpotType = "T") Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo = "Y")) Or (tmSdf.sSpotType = "Y") Or (tmSdf.sSpotType = "R") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "X") Then
                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slAvDate
                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                            ilRet = mCreateSaveSpotImage(ilUpperBound, slAvDate, slTime, 0, 10)
                            If Not ilRet Then
                                mReadSdfRec = False
                                Screen.MousePointer = vbDefault
                                Exit Function
                            End If
                            ' Get another row of space in the array just in case
                            '     more records are available
                            ilUpperBound = ilUpperBound + 1
                            'ReDim Preserve smShow(1 To 13, 1 To ilUpperBound) As String
                            'ReDim Preserve smSave(1 To 10, 1 To ilUpperBound) As String
                            ReDim Preserve tgSave(0 To ilUpperBound) As SAVEINFO
                            'ReDim Preserve imSave(1 To 2, 1 To ilUpperBound) As Integer
                            'ReDim Preserve imPostSpotInfo(1 To 4, 1 To ilUpperBound) As Integer
                            If ilUpperBound > ilNumOfrecToRead Then
                                ilRet = 0
                                Exit Do
                            End If
                        End If
                    End If
                    ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
        End If
    Next llDate
    mGenShowImage
    'vbcPosting.Min = LBound(tgSave) 'LBound(smSave, 2)
    ''If UBound(smSave, 2) <= vbcPosting.LargeChange + 1 Then
    'If UBound(tgSave) <= vbcPosting.LargeChange + 1 Then
    '' If this is used, there are probably 0 or 1 records
    '    vbcPosting.Max = LBound(tgSave) 'LBound(smSave, 2)
    'Else
    '' Saves, what amounts to, the count of records just retrieved
    '    'vbcPosting.Max = UBound(smSave, 2) - vbcPosting.LargeChange
    '    vbcPosting.Max = UBound(tgSave) - vbcPosting.LargeChange
    'End If
    If ilNumOfrecToRead <> 32000 Then
        imSettingValue = True
        vbcPosting.Value = vbcPosting.Min
    End If
    ilRet = mReadMdSdfRec(ilGetAll)
    mReadSdfRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mReadSdfRecErr:
    On Error GoTo 0
    mReadSdfRec = False
    Screen.MousePointer = vbDefault
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveAvail                    *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get avail within Ssf           *
'*                                                     *
'*******************************************************
Private Function mRemoveAvail(slSchDate As String, slTime As String) As Integer
    Dim ilRet As Integer
    Dim llTime As Long
    Dim llATime As Long
    Dim ilLoop As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilEvt As Integer
    Dim ilType As Integer

    If rbcType(1).Value Then
        mRemoveAvail = True
        Exit Function
    End If
    llTime = gTimeToCurrency(slTime, False)
    imSelectedDay = gWeekDayStr(slSchDate)
    ilType = tgDates(imDateSelectedIndex).iGameNo
    imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
    tmSsfSrchKey.iType = ilType 'slType-On Air
    tmSsfSrchKey.iVefCode = imVefCode
    gPackDate slSchDate, ilDate0, ilDate1
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf(imSelectedDay).iType = ilType) And (tmSsf(imSelectedDay).iVefCode = imVefCode) And (tmSsf(imSelectedDay).iDate(0) = ilDate0) And (tmSsf(imSelectedDay).iDate(1) = ilDate1)
        For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
           LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llATime
                If llTime = llATime Then
                    If (tmAvail.ianfCode = igPLAnfCode) And (tmAvail.iNoSpotsThis = 0) Then
                        'Remove avail
                        ilRet = gSSFGetPosition(hmSsf, lmSsfRecPos(imSelectedDay))
                        Do
                            imSsfRecLen = Len(tmSsf(imSelectedDay))
                            ilRet = gSSFGetDirect(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, lmSsfRecPos(imSelectedDay), INDEXKEY0, BTRV_LOCK_NONE)
                            If ilRet <> BTRV_ERR_NONE Then
                                mRemoveAvail = False
                                Exit Function
                            End If
                            ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf(imSelectedDay))
                            If ilRet <> BTRV_ERR_NONE Then
                                mRemoveAvail = False
                                Exit Function
                            End If
                            'Move events donw and added avail
                            For ilEvt = ilLoop To tmSsf(imSelectedDay).iCount - 1 Step 1
                                tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilEvt) = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilEvt + 1)
                            Next ilEvt
                            tmSsf(imSelectedDay).iCount = tmSsf(imSelectedDay).iCount - 1
                            imSsfRecLen = igSSFBaseLen + tmSsf(imSelectedDay).iCount * Len(tmProg)
                            ilRet = gSSFUpdate(hmSsf, tmSsf(imSelectedDay), imSsfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            mRemoveAvail = False
                            Exit Function
                        End If
                    End If
                    mRemoveAvail = True
                    Exit Function
                End If
            End If
        Next ilLoop
        imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
        ilRet = gSSFGetNext(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mRemoveAvail = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Update Sdf based on type of    *
'*                      made.  Each field change is    *
'*                      processed when field is left,  *
'*                      not with an update button.     *
'*                      The SSF is also updated.       *
'*                                                     *
'*          12-14-04 Fix subscript out of range
'*******************************************************
Private Function mSaveRec() As Integer
'   Note: per Inquiry, PSA, Promo,.. processed within imcTrash
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim slRet As String
    Dim llSdfRecPos As Long
    Dim llTzfRecPos As Long
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilFirst As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    Dim slSchDate As String
    Dim slSchTime As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slAvTime As String
    Dim slAvDate As String
    Dim ilAvailIndex As Integer
    Dim ilBkQH As Integer
    Dim ilFound As Integer
    Dim ilAnyZoneCopy As Integer
    Dim slOrigDate As String
    Dim slOrigTime As String
    Dim ilUpdateSdf As Integer
    Dim ilRetC As Integer
    Dim llDate As Long
    Dim ilAnfCode As Integer
    Dim tlTzf As TZF
    Dim tlSdf As SDF
    Dim ilRowNo As Integer
    Dim ilCount As Integer
    Dim slDate As String
    Dim slSchStatus As String
    Dim ilBoxNo As Integer
    Dim ilFindAdjAvail As Integer  'Find near avail to air time/date
    Dim ilGameNo As Integer
    Dim ilOrigGameNo As Integer
    Dim slXMid As String

    If Not imUpdateAllowed Then
        mSaveRec = False
        Exit Function
    End If
    ilBoxNo = imBoxNo
    ilFirst = True
    ilUpdateSdf = True
    ilGameNo = tgDates(imDateSelectedIndex).iGameNo
    ilOrigGameNo = ilGameNo
    If imSdfChg Then ' mSetShow sets this flag if data changed
        'If (ilBoxNo <> MISSEDREASON) And (ilBoxNo <> COPYINDEX) Then
        '    imSdfAnyChg = True
        'End If
        slAirDate = ""
        If (ilBoxNo = TIMEINDEX) Or (ilBoxNo = DATEINDEX) Then 'if time changed- set Change status
            ilRowNo = tgShow(imRowNo).iSaveInfoIndex        '12-14-04
            slSchDate = tgSave(ilRowNo).sSchDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
            slAirDate = tgSave(ilRowNo).sAirDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
            If gDateValue(slSchDate) = gDateValue(slAirDate) Then
                slAirDate = ""
            End If
        ElseIf ilBoxNo = SCHTOMISSED Then
            ilRowNo = tgShow(imRowNo).iSaveInfoIndex        '12-14-04
            slSchDate = tgSave(ilRowNo).sSchDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
        ElseIf ilBoxNo = SCHTOCANCEL Then
            ilRowNo = tgShow(imRowNo).iSaveInfoIndex        '12-14-04
            slSchDate = tgSave(ilRowNo).sSchDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
        ElseIf ilBoxNo = SCHTOHIDE Then
            ilRowNo = tgShow(imRowNo).iSaveInfoIndex        '12-14-04
            slSchDate = tgSave(ilRowNo).sSchDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
        ElseIf (ilBoxNo = MISSEDTOSCH) Or (ilBoxNo = BONUSTOSCH) Then
            If imDropRowNo < UBound(tgShow) Then    'UBound(smSave, 2) Then
                slSchDate = tgSave(tgShow(imDropRowNo).iSaveInfoIndex).sAirDate 'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
            Else
                slSchDate = tgDates(imDateSelectedIndex).sDate
            End If
        End If
        If ilGameNo <= 0 Then
            If slAirDate <> "" Then
                If Not mBlockDay(65536 * imVefCode + gDateValue(slSchDate), 65536 * imVefCode + gDateValue(slAirDate)) Then
                    Screen.MousePointer = vbDefault
                    mSaveRec = False
                    imSdfChg = False
                    Exit Function
                End If
            Else
                If Not mBlockDay(65536 * imVefCode + gDateValue(slSchDate), 0) Then
                    Screen.MousePointer = vbDefault
                    mSaveRec = False
                    imSdfChg = False
                    Exit Function
                End If
            End If
        Else
            If Not mBlockDay(65536 * imVefCode + ilGameNo, 0) Then
                Screen.MousePointer = vbDefault
                mSaveRec = False
                imSdfChg = False
                Exit Function
            End If
        End If
        Do
            ilRet = btrBeginTrans(hmSdf, 1000)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                mUnblockDay
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                mSaveRec = False
                imSdfChg = False
                Exit Function
            End If
            If (ilBoxNo = MISSEDTOSCH) Or (ilBoxNo = MISSEDREASON) Or (ilBoxNo = MISSEDTOCANCEL) Or (ilBoxNo = MISSEDTOHIDE) Then
                llSdfRecPos = tgMdSdfRec(imSdfIndex).lSdfRecPos    'lmMdRecPos(imMdRowNo)
            ElseIf (ilBoxNo = BONUSTOSCH) Then
                'tgPLCntSpot(imMdRowNo)
                slDate = tgDates(imDateSelectedIndex).sDate 'cbcDate.List(imDateSelectedIndex)
                ilRet = mMakeUnschSpot(tgMdSaveInfo(imSaveIndex).lChfCode, tgMdSaveInfo(imSaveIndex).iLineNo, tgMdSaveInfo(imSaveIndex).lFsfCode, ilGameNo, slDate, imVefCode, llSdfRecPos)
            Else
                ilRowNo = tgShow(imRowNo).iSaveInfoIndex
                llSdfRecPos = tgSave(ilRowNo).lSdfRecPos    'Val(smSave(SAVRECPOSINDEX, imRowNo))
            End If
            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmSdf)
                mUnblockDay
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                mSaveRec = False
                Exit Function
            End If
            slSchStatus = tmSdf.sSchStatus
            If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                tmSmfSrchKey2.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilOrigGameNo = tmSmf.iGameNo
                End If
            End If
            'tmRec = tmSdf
            'ilRet = gGetByKeyForUpdate("SDF", hmSdf, tmRec)
            'tmSdf = tmRec
            'If ilRet <> BTRV_ERR_NONE Then
            '    ilRet = btrAbortTrans(hmSdf)
            '    Screen.MousePointer = vbDefault
            '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
            '    mSaveRec = False
            '    Exit Function
            'End If
            If (ilBoxNo = TIMEINDEX) Or (ilBoxNo = DATEINDEX) Then 'if time changed- set Change status
                slOrigTime = smSvAvailTime
                slOrigDate = smSvDate
                slSchDate = tgSave(ilRowNo).sSchDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
                slSchTime = tgSave(ilRowNo).sSchTime    'Original Spot Airing time (scheduled time), smSave(SAVTIMEINDEX, imRowNo)
                slAirDate = tgSave(ilRowNo).sAirDate    'New Spot Airing time, smSelectedDate  'cbcDate.List(imDateSelectedIndex)
                slAirTime = tgSave(ilRowNo).sAirTime    'smSave(SAVTIMEINDEX, imRowNo)
                slXMid = tmSdf.sXCrossMidnight
                '6/13/08: Moving a spot after 9pm to before 6a, make it XMid
                If gDateValue(slSchDate) = gDateValue(slAirDate) Then
                    If gTimeToLong(slAirTime, False) >= gTimeToLong("6:00:00AM", False) Then
                        slXMid = "N"
                    Else
                        If gTimeToLong(slSchTime, False) >= gTimeToLong("12:00:00PM", False) Then
                            slXMid = "Y"
                        End If
                    End If
                Else
                    slXMid = "N"
                End If
                If tmVef.sType <> "G" Then
                    imSdfAnyChg(gWeekDayStr(slSchDate)) = True
                Else
                    imSdfAnyChg(0) = True
                End If
                ''If (ilFirst) And (rbcType(0).Value) Then
                'If (ilFirst) And (lbcAvailTimes.ListIndex >= 0) And (rbcType(0).Value) Then
                If (ilFirst) And (rbcType(0).Value) And (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then

                    '6/16/11
                    If (gDateValue(slSchDate) <> gDateValue(slOrigDate)) Or (gTimeToCurrency(slSchTime, False) <> gTimeToCurrency(slOrigTime, False)) Then
                        '6/16/11
                        'If lbcAvailTimes.ListIndex >= 0 Then
                            ilFindAdjAvail = False
                        'Else
                        '    ilFindAdjAvail = True
                        'End If
                        '6/16/11: Set in mFindAvail and used in mAvailRoom
                        'imSelectedDay = gWeekDayStr(slSchDate)
                        ''If Not mFindSpotOrigTime(slSchDate, slOrigTime) Then
                        ''    ilRet = btrAbortTrans(hmSdf)
                        ''    Screen.MousePointer = vbDefault
                        ''    mSaveRec = False
                        ''    imSdfChg = False
                        ''    Exit Function
                        ''End If
                        If Not mFindAvail(slSchDate, slSchTime, ilGameNo, ilFindAdjAvail, ilAvailIndex) Then
                            'If Not mAddAvail(slSchTime, ilAvailIndex) Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            'End If
                        End If
                        If Not mAvailRoom(ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        slRet = mMoveTest(slOrigDate, slAirDate, slAirTime)
                        If slRet = "" Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        'Unschedule, then schedule (gChgSchSpot removes Smf if exist)
                        '6/16/11: Find avail that spot was within
                        If Not mFindSpotOrigTime(slOrigDate, slOrigTime, ilGameNo, ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        imSelectedDay = gWeekDayStr(slOrigDate)
                        gPackDate slOrigDate, tmSdf.iDate(0), tmSdf.iDate(1)
                        gPackTime slOrigTime, tmSdf.iTime(0), tmSdf.iTime(1)
                        ilRet = gChgSchSpot("TM", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                        If Not ilRet Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            'Screen.MousePointer = vbDefault
                            'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                            ilRet = igBtrError
                            If ilRet >= 30000 Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later.  Error in " & sgErrLoc & ", Error #" & str$(ilRet), vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        '6/16/11
                        'If Not mRemoveAvail(slSchDate, slOrigTime) Then
                        If Not mAvailReset(ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        If Not mFindAvail(slSchDate, slSchTime, ilGameNo, ilFindAdjAvail, ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        If imBkQH <= 1000 Then  'Above 1000 is DR; Remnant; PI; Trade; PSA; Promo
                            If (slRet = "G") Or (slRet = "O") Then
                                ilBkQH = 0
                            Else
                                ilBkQH = imBkQH
                            End If
                        Else    'D.R.; Remnant; PI, Trade; Promo; PSA
                            ilBkQH = imBkQH
                        End If
                        'Schedule spot, Smf created if required
                        tlSdf = tmSdf
                        ilRet = gBookSpot(slRet, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf(imSelectedDay), lmSsfRecPos(imSelectedDay), ilAvailIndex, -1, tmChf, tmClf, tmLnRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, imPriceLevel, False, hmSxf, hmGsf)
                        If Not ilRet Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        'Reset copy
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        'tmRec = tmSdf
                        'ilRet = gGetByKeyForUpdate("SDF", hmSdf, tmRec)
                        'tmSdf = tmRec
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    ilRet = btrAbortTrans(hmSdf)
                        '    Screen.MousePointer = vbDefault
                        '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                        '    mSaveRec = False
                        '    imSdfChg = False
                        '    Exit Function
                        'End If
                        tmSdf.iRotNo = tlSdf.iRotNo
                        tmSdf.sPtType = tlSdf.sPtType
                        tmSdf.lCopyCode = tlSdf.lCopyCode
                        tgSave(ilRowNo).sSchDate = slSchDate   'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
                        tgSave(ilRowNo).sSchTime = slSchTime   'smSave(SAVTIMEINDEX, imRowNo)
                        tgSave(ilRowNo).sAirDate = slAirDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
                        tgSave(ilRowNo).sAirTime = slAirTime    'smSave(SAVTIMEINDEX, imRowNo)
                        '6/16/11
                        'If ilFindAdjAvail Then
                            gPackDate slAirDate, tmSdf.iDate(0), tmSdf.iDate(1)
                            gPackTime slAirTime, tmSdf.iTime(0), tmSdf.iTime(1)
                        'End If
                        '6/14/10: handle case where spot moved back from 1a to 11p
                        'If slXMid = "Y" Then
                        '    tmSdf.sXCrossMidnight = "Y"
                        'End If
                        tmSdf.sXCrossMidnight = slXMid
                        tgSave(ilRowNo).sXMid = slXMid
                    Else
                        '1/10/11: Check if time outside contract times.  If so, make spot MG or Outside
                        If (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
                            slRet = mMoveTest(slOrigDate, slAirDate, slAirTime)
                            If (slRet = "G") Or (slRet = "O") Then
                                tmSmf.lChfCode = 0
                                ilRet = gMakeSmf(hmSmf, tmSmf, slRet, tmSdf, tmSdf.iVefCode, slOrigDate, slOrigTime, tmSdf.iGameNo, slAirDate, slAirTime)
                                If ilRet Then
                                    tmSdf.lSmfCode = tmSmf.lCode
                                    tmSdf.sSchStatus = slRet
                                Else
                                    ilRet = btrAbortTrans(hmSdf)
                                    mUnblockDay
                                    Screen.MousePointer = vbDefault
                                    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                                    mSaveRec = False
                                    imSdfChg = False
                                    Exit Function
                                End If
                            ElseIf slRet = "" Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            End If
                        Else
                            '1/10/11: Later, add test if MG/Outside should be removed
                        End If
                        gPackTime slAirTime, tmSdf.iTime(0), tmSdf.iTime(1)
                        tmSdf.sXCrossMidnight = slXMid
                        tgSave(ilRowNo).sXMid = slXMid
                    End If
                Else
                    'If rbcType(1).Value Then
                        gPackDate tgSave(ilRowNo).sAirDate, tmSdf.iDate(0), tmSdf.iDate(1)
                    'End If
                    gPackTime slAirTime, tmSdf.iTime(0), tmSdf.iTime(1)
                End If
                'Update Smf time if required
                'ilRet = gFindSmf(tmSdf, hmSmf, tmSmf)
                'If ilRet Then
                '    gPackTime smSave(SAVTIMEINDEX, imRowNo), tmSmf.iActualTime(0), tmSmf.iActualTime(1)
                '    ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen) '  Update File
                'End If
                'gPackTime smSave(SAVTIMEINDEX, imRowNo), tmSdf.iTime(0), tmSdf.iTime(1)
                'Select Case tmSdf.sAffChg
                '    Case " "    'No changed
                '        tmSdf.sAffChg = "T" 'time change
                '        gSetShow pbcPosting, tmSdf.sAffChg, tmCtrls(AUDINDEX)
                '        smShow(AUDINDEX, imRowNo) = tmCtrls(AUDINDEX).sShow
                '    Case "C"    'Copy Changed
                '        tmSdf.sAffChg = "B"
                '        gSetShow pbcPosting, tmSdf.sAffChg, tmCtrls(AUDINDEX)
                '        smShow(AUDINDEX, imRowNo) = tmCtrls(AUDINDEX).sShow
                'End Select
                If tmSdf.sSchStatus = "G" Then
                    slStr = "G" '"M" the M is too large - it didn't display
                ElseIf tmSdf.sSchStatus = "O" Then
                    slStr = "O"
                Else
                    slStr = " "
                End If
                tgSave(ilRowNo).sSchStatus = slStr
                gSetShow pbcPosting, slStr, tmCtrls(MGOODINDEX) ' Shorten it
                tgShow(imRowNo).sShow(MGOODINDEX) = tmCtrls(MGOODINDEX).sShow
                tgSave(ilRowNo).sAffChg = "Y"
                tmSdf.sAffChg = "Y"
                gSetShow pbcPosting, tmSdf.sAffChg, tmCtrls(AUDINDEX)
                tgShow(imRowNo).sShow(AUDINDEX) = tmCtrls(AUDINDEX).sShow
            ElseIf ilBoxNo = COPYINDEX Then
                If imTZCopyAllowed Then
                    imSdfChg = False
                    ilAnyZoneCopy = False
                    If (StrComp(smTZSave(2, 1), "[None]", 1) = 0) Then
                        For ilLoop = 2 To 8 Step 1
                            If (StrComp(smTZSave(2, ilLoop), "[None]", 1) <> 0) And (Trim$(smTZSave(2, ilLoop)) <> "") Then
                                ilAnyZoneCopy = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilAnyZoneCopy Then  'All
                        If tmSdf.sPtType = "3" Then 'Remove time zone copy
                            ilRet = btrGetPosition(hmTzf, llTzfRecPos)
                            If ilRet <> BTRV_ERR_NONE Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            End If
                            Do
                                'tmRec = tmTzf
                                'ilRet = gGetByKeyForUpdate("TZF", hmTzf, tmRec)
                                'tmTzf = tmRec
                                'If ilRet <> BTRV_ERR_NONE Then
                                '    ilRet = btrAbortTrans(hmSdf)
                                '    Screen.MousePointer = vbDefault
                                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                                '    mSaveRec = False
                                '    imSdfChg = False
                                '    Exit Function
                                'End If
                                ilRet = btrDelete(hmTzf)
                                If ilRet = BTRV_ERR_CONFLICT Then
                                    ilCRet = btrGetDirect(hmTzf, tmTzf, imTzfRecLen, llTzfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            End If
                            imSdfChg = True
                        End If
                        If smTZSave(2, 1) = "[None]" Then
                            If (tmSdf.sPtType <> "0") Or (tmSdf.lCopyCode <> 0) Then
                                tmSdf.sPtType = "0"
                                tmSdf.lCopyCode = 0
                                tmSdf.iRotNo = 0
                                imSdfChg = True
                            End If
                        Else
                            'gFindPartialMatch smTZSave(2, 1), 1, Len(smTZSave(2, 1)), lbcCopyNm
                            gFindMatch smTZSave(2, 1), 1, lbcCopyNm
                            If gLastFound(lbcCopyNm) > 0 Then
                                slNameCode = tmCopyNmCode(gLastFound(lbcCopyNm) - 1).sKey    'lbcCopyNmCode.List(gLastFound(lbcCopyNm) - 1)
                                'Set copy Cif
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If ilRet = CP_MSG_NONE Then
                                    If (tmSdf.sPtType <> "1") Or (tmSdf.lCopyCode <> Val(slCode)) Then
                                        tmSdf.lCopyCode = Val(slCode)
                                        tmSdf.sPtType = "1"
                                        imSdfChg = True
                                    End If
                                Else
                                    If (tmSdf.sPtType <> "0") Or (tmSdf.lCopyCode <> 0) Then
                                        tmSdf.lCopyCode = 0
                                        tmSdf.sPtType = "0"
                                        tmSdf.iRotNo = 0
                                        imSdfChg = True
                                    End If
                                End If
                            Else
                                If (tmSdf.sPtType <> "0") Or (tmSdf.lCopyCode <> 0) Then
                                    tmSdf.sPtType = "0"
                                    tmSdf.lCopyCode = 0
                                    tmSdf.iRotNo = 0
                                    imSdfChg = True
                                End If
                            End If
                        End If
                    Else    'Set up time zone copy
                        tlTzf = tmTzf
                        'For ilLoop = 1 To 6 Step 1
                        For ilLoop = 0 To 5 Step 1
                            tmTzf.sZone(ilLoop) = ""
                            tmTzf.lCifZone(ilLoop) = 0
                            tmTzf.iRotNo(ilLoop) = 0
                        Next ilLoop
                        ilIndex = 0
                        For ilLoop = 3 To 8 Step 1
                            If smTZSave(1, ilLoop) <> "" Then
                                'gFindPartialMatch smTZSave(2, ilLoop), 1, Len(smTZSave(2, ilLoop)), lbcCopyNm
                                gFindMatch smTZSave(2, ilLoop), 1, lbcCopyNm
                                If gLastFound(lbcCopyNm) > 0 Then
                                    slNameCode = tmCopyNmCode(gLastFound(lbcCopyNm) - 1).sKey    'lbcCopyNmCode.List(gLastFound(lbcCopyNm) - 1)
                                    'Set copy Cif
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    If ilRet = CP_MSG_NONE Then
                                        ilIndex = ilIndex + 1
                                        tmTzf.lCifZone(ilIndex - 1) = Val(slCode)
                                        tmTzf.sZone(ilIndex - 1) = smZones(ilLoop)
                                    End If
                                End If
                            End If
                        Next ilLoop
                        If ilIndex < 6 Then
                            If smTZSave(1, 2) <> "" Then
                                'gFindPartialMatch smTZSave(2, 2), 1, Len(smTZSave(2, 2)), lbcCopyNm
                                gFindMatch smTZSave(2, 2), 1, lbcCopyNm
                                If gLastFound(lbcCopyNm) > 0 Then
                                    slNameCode = tmCopyNmCode(gLastFound(lbcCopyNm) - 1).sKey    'lbcCopyNmCode.List(gLastFound(lbcCopyNm) - 1)
                                    'Set copy Cif
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    If ilRet = CP_MSG_NONE Then
                                        ilIndex = ilIndex + 1
                                        tmTzf.lCifZone(ilIndex - 1) = Val(slCode)
                                        tmTzf.sZone(ilIndex - 1) = "Oth"  'smZones(2)
                                    End If
                                End If
                            End If
                        End If
                        'Compare tmTzf and tlTzf
                        If Not imSdfChg Then
                            ilFound = False
                            For ilLoop = 1 To 6 Step 1
                                If tmTzf.lCifZone(ilLoop - 1) > 0 Then
                                    For ilIndex = 1 To 6 Step 1
                                        If (tmTzf.lCifZone(ilLoop - 1) = tlTzf.lCifZone(ilIndex - 1)) And (StrComp(tmTzf.sZone(ilLoop - 1), tlTzf.sZone(ilIndex - 1), 1) = 0) Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next ilIndex
                                    If Not ilFound Then
                                        imSdfChg = True
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                        End If
                        If Not imSdfChg Then
                            ilFound = False
                            For ilIndex = 1 To 6 Step 1
                                If tlTzf.lCifZone(ilIndex - 1) > 0 Then
                                    For ilLoop = 1 To 6 Step 1
                                        If (tmTzf.lCifZone(ilLoop - 1) = tlTzf.lCifZone(ilIndex - 1)) And (StrComp(tmTzf.sZone(ilLoop - 1), tlTzf.sZone(ilIndex - 1), 1) = 0) Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If Not ilFound Then
                                        imSdfChg = True
                                        Exit For
                                    End If
                                End If
                            Next ilIndex
                        End If
                        If tmSdf.sPtType = "3" Then
                            ilRet = btrGetPosition(hmTzf, llTzfRecPos)
                            If ilRet <> BTRV_ERR_NONE Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            End If
                            Do
                                'tmRec = tmTzf
                                'ilRet = gGetByKeyForUpdate("TZF", hmTzf, tmRec)
                                'tlTzf = tmRec
                                'If ilRet <> BTRV_ERR_NONE Then
                                '    ilRet = btrAbortTrans(hmSdf)
                                '    Screen.MousePointer = vbDefault
                                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                                '    mSaveRec = False
                                '    imSdfChg = False
                                '    Exit Function
                                'End If
                                ilRet = btrUpdate(hmTzf, tmTzf, imTzfRecLen) '  Update File
                                If ilRet = BTRV_ERR_CONFLICT Then
                                    ilCRet = btrGetDirect(hmTzf, tmTzf, imTzfRecLen, llTzfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            End If
                        Else
                            imSdfChg = True
                            tmTzf.lCode = 0
                            ilRet = btrInsert(hmTzf, tmTzf, imTzfRecLen, INDEXKEY0)
                            If ilRet <> BTRV_ERR_NONE Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            End If
                        End If
                        tmSdf.sPtType = "3"
                        tmSdf.lCopyCode = tmTzf.lCode
                    End If
                Else
                    If lbcCopyNm.ListIndex <= 0 Then
                        tmSdf.sPtType = "0"
                        tmSdf.lCopyCode = 0
                        tmSdf.iRotNo = 0
                    Else
                        slNameCode = tmCopyNmCode(imCopyNmListIndex - 1).sKey  'lbcCopyNmCode.List(imCopyNmListIndex - 1)
                        'Set copy Cif
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If ilRet = CP_MSG_NONE Then
                            tmSdf.lCopyCode = Val(slCode)
                            tmSdf.sPtType = "1"
                        Else
                            tmSdf.lCopyCode = 0
                            tmSdf.sPtType = "0"
                            tmSdf.iRotNo = 0
                        End If
                    End If
                End If
                If imSdfChg Then
                    slSchDate = tgSave(ilRowNo).sSchDate
                    If tmVef.sType <> "G" Then
                        imSdfAnyChg(gWeekDayStr(slSchDate)) = True  'imSdfAnyChg = True
                    Else
                        imSdfAnyChg(0) = True
                    End If
                    'Select Case tmSdf.sAffChg
                    '    Case " "    'No changed
                    '        tmSdf.sAffChg = "C" 'time change
                    '        gSetShow pbcPosting, tmSdf.sAffChg, tmCtrls(AUDINDEX)
                    '        smShow(AUDINDEX, imRowNo) = tmCtrls(AUDINDEX).sShow
                    '    Case "T"    'Time Changed
                    '        tmSdf.sAffChg = "B"
                    '        gSetShow pbcPosting, tmSdf.sAffChg, tmCtrls(AUDINDEX)
                    '        smShow(AUDINDEX, imRowNo) = tmCtrls(AUDINDEX).sShow
                    'End Select
                    tgSave(ilRowNo).sAffChg = "Y"
                    tmSdf.sAffChg = "Y"
                    gSetShow pbcPosting, tmSdf.sAffChg, tmCtrls(AUDINDEX)
                    tgShow(imRowNo).sShow(AUDINDEX) = tmCtrls(AUDINDEX).sShow
                    ilRet = mGetCopy(tmSdf, ilRowNo)
                    If (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G") Then
                        If (tgSpf.sInvAirOrder = "O") Or (tgSpf.sInvAirOrder = "S") Then
                            'Update SMF
                            tmSmf.sPtType = tmSdf.sPtType
                            tmSmf.lCopyCode = tmSdf.lCopyCode
                            tmSmf.iRotNo = tmSdf.iRotNo
                            ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen)
                        End If
                    End If
                End If
            ElseIf ilBoxNo = PRICEINDEX Then
'                Select Case tgSave(ilRowNo).iPrice 'imSave(1, imRowNo)
'                    Case 1
'                        tmSdf.sPriceType = "P"
'                    Case Else
'                        tmSdf.sPriceType = "L"
'                End Select
                Select Case tgSave(ilRowNo).iPrice
                    Case 1
                        tmSdf.sPriceType = "+"
                    Case 2
                        tmSdf.sPriceType = "-"
                    Case Else
                        If tmAdf.sBonusOnInv <> "N" Then
                            tmSdf.sPriceType = "B"
                        Else
                            tmSdf.sPriceType = "N"
                        End If
                End Select
                slSchDate = tgSave(ilRowNo).sSchDate
                If tmVef.sType <> "G" Then
                    imSdfAnyChg(gWeekDayStr(slSchDate)) = True  'imSdfAnyChg = True
                Else
                    imSdfAnyChg(0) = True
                End If
            ElseIf ilBoxNo = SCHTOMISSED Then
                If ilFirst Then
                    slSchDate = tgSave(ilRowNo).sSchDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
                    slSchTime = tgSave(ilRowNo).sSchTime    'smSave(SAVTIMEINDEX, imRowNo)
                    imSelectedDay = gWeekDayStr(slSchDate)
                    If tmVef.sType <> "G" Then
                        imSdfAnyChg(imSelectedDay) = True
                    Else
                        imSdfAnyChg(0) = True
                    End If
                    'Set Date and time so that gChgSchSpot can locate the spot in SSF
                    If (rbcType(0).Value) And (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                        If Not mFindSpotOrigTime(slSchDate, slOrigTime, ilGameNo, ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        gPackDate slSchDate, tmSdf.iDate(0), tmSdf.iDate(1)
                        gPackTime slOrigTime, tmSdf.iTime(0), tmSdf.iTime(1)
                    End If
                    If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Then
                        ilRet = gChgSchSpot("D", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                    ElseIf (rbcType(1).Value) Or (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                        ilRet = btrDelete(hmSdf)
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = True
                            ilUpdateSdf = False
                        Else
                            igBtrError = ilRet
                            sgErrLoc = "Post Log"
                            ilRet = False
                        End If
                    Else
                        '12/14/15: Set missed reason
                        igMnfMissed = igDefaultMnfMissed
                        If imMissedListIndex >= 0 Then
                            slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                            'Set missed reason
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If ilRet = CP_MSG_NONE Then
                                igMnfMissed = Val(slCode)
                            End If
                        End If
                        ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                    End If
                    If Not ilRet Then
                        ilRet = btrAbortTrans(hmSdf)
                        mUnblockDay
                        'Screen.MousePointer = vbDefault
                        'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                        ilRet = igBtrError
                        If ilRet >= 30000 Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later.  Error in " & sgErrLoc & ", Error #" & str$(ilRet), vbOKOnly + vbExclamation, "Save")
                        mSaveRec = False
                        imSdfChg = False
                        Exit Function
                    End If
                    If (rbcType(0).Value) And (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                        '5/20/11
                        'If Not mRemoveAvail(slSchDate, slOrigTime) Then
                        If Not mAvailReset(ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                    End If
                End If
                'If (smDragCntrType <> "T") And (smDragCntrType <> "Q") And (smDragCntrType <> "S") And (smDragCntrType <> "M") And (smDragCntrType <> "X") Then
                If (rbcType(1).Value) Or ((smDragCntrType = "T") And (tgSpf.sSchdRemnant = "Y")) Or (smDragCntrType = "C") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo = "Y")) Or (smDragCntrType = "V") Then
                    tmSdf.sTracer = "2"
                    tmSdf.sAffChg = ""  'remove any post log changes
                    slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                    'Set missed reason
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If ilRet = CP_MSG_NONE Then
                        tmSdf.iMnfMissed = Val(slCode)
                    Else
                        tmSdf.iMnfMissed = igDefaultMnfMissed
                    End If
                Else
                    ilUpdateSdf = False
                End If
            ElseIf ilBoxNo = SCHTOCANCEL Then
                If ilFirst Then
                    slSchDate = tgSave(ilRowNo).sSchDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
                    slSchTime = tgSave(ilRowNo).sSchTime    'smSave(SAVTIMEINDEX, imRowNo)
                    imSelectedDay = gWeekDayStr(slSchDate)
                    If tmVef.sType <> "G" Then
                        imSdfAnyChg(imSelectedDay) = True
                    Else
                        imSdfAnyChg(0) = True
                    End If
                    'Set Date and time so that gChgSchSpot can locate the spot in SSF
                    If Not mFindSpotOrigTime(slSchDate, slOrigTime, ilGameNo, ilAvailIndex) Then
                        ilRet = btrAbortTrans(hmSdf)
                        mUnblockDay
                        Screen.MousePointer = vbDefault
                        mSaveRec = False
                        imSdfChg = False
                        Exit Function
                    End If
                    gPackDate slSchDate, tmSdf.iDate(0), tmSdf.iDate(1)
                    gPackTime slOrigTime, tmSdf.iTime(0), tmSdf.iTime(1)
                    '2/21/13: Allow Package spots to be cancelled
                    '12/14/15: Set missed reason
                    igMnfMissed = igDefaultMnfMissed
                    If imMissedListIndex >= 0 Then
                        slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                        'Set missed reason
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If ilRet = CP_MSG_NONE Then
                            igMnfMissed = Val(slCode)
                        End If
                    End If
                    If rbcType(1).Value Then
                        tmSdf.sSchStatus = "C"
                        tmSdf.iMnfMissed = igMnfMissed
                        ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = True
                            ilUpdateSdf = False
                        Else
                            igBtrError = ilRet
                            sgErrLoc = "Post Log"
                            ilRet = False
                        End If
                    Else
                        ilRet = gChgSchSpot("C", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                    End If
                    If Not ilRet Then
                        ilRet = btrAbortTrans(hmSdf)
                        mUnblockDay
                        ilRet = igBtrError
                        If ilRet >= 30000 Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later.  Error in " & sgErrLoc & ", Error #" & str$(ilRet), vbOKOnly + vbExclamation, "Save")
                        mSaveRec = False
                        imSdfChg = False
                        Exit Function
                    End If
                    '5/20/11
                    If (rbcType(0).Value) And (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                        If Not mAvailReset(ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                    End If
                End If
                'If (smDragCntrType <> "T") And (smDragCntrType <> "Q") And (smDragCntrType <> "S") And (smDragCntrType <> "M") And (smDragCntrType <> "X") Then
                If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant = "Y")) Or (smDragCntrType = "C") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo = "Y")) Or (smDragCntrType = "V") Then
                    tmSdf.sTracer = "2"
                    tmSdf.sAffChg = ""  'remove any post log changes
                    ''tmSdf.sPtType = "0" 'No copy
                    'tmSdf.iMnfMissed = 0
                    slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                    'Set missed reason
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If ilRet = CP_MSG_NONE Then
                        tmSdf.iMnfMissed = Val(slCode)
                    Else
                        tmSdf.iMnfMissed = igDefaultMnfMissed
                    End If
                Else
                    ilUpdateSdf = False
                End If
            ElseIf ilBoxNo = SCHTOHIDE Then
                If ilFirst Then
                    slSchDate = tgSave(ilRowNo).sSchDate    'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
                    slSchTime = tgSave(ilRowNo).sSchTime    'smSave(SAVTIMEINDEX, imRowNo)
                    imSelectedDay = gWeekDayStr(slSchDate)
                    If tmVef.sType <> "G" Then
                        imSdfAnyChg(imSelectedDay) = True
                    Else
                        imSdfAnyChg(0) = True
                    End If
                    'Set Date and time so that gChgSchSpot can locate the spot in SSF
                    If Not mFindSpotOrigTime(slSchDate, slOrigTime, ilGameNo, ilAvailIndex) Then
                        ilRet = btrAbortTrans(hmSdf)
                        mUnblockDay
                        Screen.MousePointer = vbDefault
                        mSaveRec = False
                        imSdfChg = False
                        Exit Function
                    End If
                    gPackDate slSchDate, tmSdf.iDate(0), tmSdf.iDate(1)
                    gPackTime slOrigTime, tmSdf.iTime(0), tmSdf.iTime(1)
                    ilRet = gChgSchSpot("H", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hmSxf, hmGsf, hmGhf)
                    If Not ilRet Then
                        ilRet = btrAbortTrans(hmSdf)
                        mUnblockDay
                        'Screen.MousePointer = vbDefault
                        'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                        ilRet = igBtrError
                        If ilRet >= 30000 Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Save Not Completed, Try Later.  Error in " & sgErrLoc & ", Error #" & str$(ilRet), vbOKOnly + vbExclamation, "Save")
                        mSaveRec = False
                        imSdfChg = False
                        Exit Function
                    End If
                    '5/20/11
                    If (rbcType(0).Value) And (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                        If Not mAvailReset(ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                    End If
                End If
                'If (smDragCntrType <> "T") And (smDragCntrType <> "Q") And (smDragCntrType <> "S") And (smDragCntrType <> "M") And (smDragCntrType <> "X") Then
                If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant = "Y")) Or (smDragCntrType = "C") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo = "Y")) Or (smDragCntrType = "V") Then
                    tmSdf.sTracer = "2"
                    tmSdf.sAffChg = ""  'remove any post log changes
                    'tmSdf.sPtType = "0" 'No copy
                    tmSdf.iMnfMissed = 0
                Else
                    ilUpdateSdf = False
                End If
            ElseIf (ilBoxNo = MISSEDTOSCH) Or (ilBoxNo = BONUSTOSCH) Then
                If imDropRowNo < UBound(tgShow) Then    'UBound(smSave, 2) Then
                    slSchDate = tgSave(tgShow(imDropRowNo).iSaveInfoIndex).sAirDate 'smSelectedDate  'cbcDate.List(imDateSelectedIndex)
                Else
                    slSchDate = tgDates(imDateSelectedIndex).sDate
                End If
                imSelectedDay = gWeekDayStr(slSchDate)
                'ilRet = gParseItem(slSchDate, 2, " ", slSchDate) 'Remove day
                'ilRet = gParseItem(slSchDate, 1, ":", slSchDate) 'Remove :xxxxxxx
                If ilFirst Then
                    'Add code to determine if MG
                    If imDropRowNo < UBound(tgShow) Then    'UBound(smSave, 2) Then
                        '5/20/11
                        'slAvDate = tgSave(tgShow(imDropRowNo).iSaveInfoIndex).sAirDate 'smSave(SAVDATEINDEX, imDropRowNo)
                        slAvDate = tgSave(tgShow(imDropRowNo).iSaveInfoIndex).sSchDate 'smSave(SAVDATEINDEX, imDropRowNo)
                        ilAnfCode = tgSave(tgShow(imDropRowNo).iSaveInfoIndex).ianfCode
                        ilCount = tgSave(tgShow(imDropRowNo).iSaveInfoIndex).iCount + 1
                    Else
                        slAvDate = ""
                        ilAnfCode = 0
                        ilCount = 10
                    End If
                    If imDropRowNo < UBound(tgShow) Then    'UBound(smSave, 2) Then
                        '5/20/11
                        'slAvTime = tgSave(tgShow(imDropRowNo).iSaveInfoIndex).sAirTime 'smSave(SAVTIMEINDEX, imDropRowNo)
                        slAvTime = tgSave(tgShow(imDropRowNo).iSaveInfoIndex).sSchTime 'smSave(SAVTIMEINDEX, imDropRowNo)
                    Else
                        'slSchTime = "12AM"
                        slAvTime = ""
                    End If
                    '5/20/11: added slAirTime
                    slSchTime = mObtainAvailTime(slAvDate, slAvTime, slSchDate, slAirTime)
                    slAvDate = slSchDate
                    If slSchTime = "" Then
                        ilRet = btrAbortTrans(hmSdf)
                        mUnblockDay
                        Screen.MousePointer = vbDefault
                        mSaveRec = False
                        imSdfChg = False
                        Exit Function
                    End If
                    'If slAvTime = "" Then
                        slAvTime = slSchTime
                    'End If
                    '5/20/11
                    If rbcType(0).Value Then
                        imSelectedDay = gWeekDayStr(slSchDate)
                        If Not mFindAvail(slAvDate, slAvTime, ilGameNo, False, ilAvailIndex) Then
                            'If Not mAddAvail(slSchTime, ilAvailIndex) Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            'End If
                        End If
                        If Not mAvailRoom(ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        If (ilBoxNo = MISSEDTOSCH) Then
                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                            slRet = mMoveTest(slDate, slAvDate, slAvTime)
                            If slRet = "" Then
                                ilRet = btrAbortTrans(hmSdf)
                                mUnblockDay
                                Screen.MousePointer = vbDefault
                                mSaveRec = False
                                imSdfChg = False
                                Exit Function
                            End If
                        Else
                            imBkQH = 0
                            slRet = "O"
                            imPriceLevel = 0
                        End If
                        If tmVef.sType <> "G" Then
                            imSdfAnyChg(gWeekDayStr(slAvDate)) = True
                        Else
                            imSdfAnyChg(0) = True
                        End If
                        tmSmf.lChfCode = 0
                        If imBkQH <= 1000 Then  'Above 1000 is DR; Remnant; PI; Trade; PSA; Promo
                            If (slRet = "G") Or (slRet = "O") Then
                                ilBkQH = 0
                            Else
                                ilBkQH = imBkQH
                            End If
                        Else    'D.R.; Remnant; PI, Trade; Promo; PSA
                            ilBkQH = imBkQH
                        End If
                        'Re-Get Ssf as it is required by gBookSpot (mMoveTest {gGetLineSchParameters} could have changed Ssf for
                        'imSelectedDay if multirecords are involved)
                        If Not mFindAvail(slAvDate, slAvTime, ilGameNo, False, ilAvailIndex) Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        ilRet = gBookSpot(slRet, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf(imSelectedDay), lmSsfRecPos(imSelectedDay), ilAvailIndex, -1, tmChf, tmClf, tmLnRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, imPriceLevel, False, hmSxf, hmGsf)
                        If Not ilRet Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                    Else
                        tmSdf.sSchStatus = "S"
                        gPackDate slSchDate, tmSdf.iDate(0), tmSdf.iDate(1)
                        gPackTime slAirTime, tmSdf.iTime(0), tmSdf.iTime(1)
                        tmSdf.iUrfCode = tgUrf(0).iCode
                        ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmSdf)
                            mUnblockDay
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                            mSaveRec = False
                            imSdfChg = False
                            Exit Function
                        End If
                    End If
                    'tmRec = tmSdf
                    'ilRet = gGetByKeyForUpdate("SDF", hmSdf, tmRec)
                    'tmSdf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hmSdf)
                    '    Screen.MousePointer = vbDefault
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                    '    mSaveRec = False
                    '    imSdfChg = False
                    '    Exit Function
                    'End If
                End If
                'gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slMissedDate
                'gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slMissedTime
                'If slRet <> "S" Then
                '    tmSmf.lChfCode = 0
                '    gMakeSmf hmSmf, tmSmf, slRet, tmSdf.lChfCode, tmSdf.iLineNo, slMissedDate, slMissedTime, slSchDate, slSchTime
                'End If
                tmSdf.sTracer = "2"
                'tmSdf.sSchStatus = slRet
                tmSdf.sAffChg = "Y" '"A" 'Added
                tmSdf.iMnfMissed = 0    'Missed reason and rotation #
                '5/20/11
                'gPackTime slSchTime, tmSdf.iTime(0), tmSdf.iTime(1)
                gPackTime slAirTime, tmSdf.iTime(0), tmSdf.iTime(1)
                'gPackTime slSchTime, tmSdf.iTime(0), tmSdf.iTime(1)
                If imDropRowNo < UBound(tgShow) Then    'UBound(smSave, 2) Then
                '    gPackTime smSave(SAVTIMEINDEX, imDropRowNo), tmSdf.iTime(0), tmSdf.iTime(1)
                Else
                '    gPackTime "12AM", tmSdf.iTime(0), tmSdf.iTime(1)
                    imDropRowNo = UBound(tgShow) - 1    'UBound(smSave, 2) - 1 'Adds after row selected
                End If
                'tmSdf.sPtType = "0" 'No copy
                'slStr = smSelectedDate 'cbcDate.List(imDateSelectedIndex)
                'ilRet = gParseItem(slSchDate, 2, " ", slStr) 'Remove day
                'ilRet = gParseItem(slSchDate, 1, ":", slStr) 'Remove :xxxxxxx
                'gPackDate slStr, tmSdf.iDate(0), tmSdf.iDate(1)
            ElseIf ilBoxNo = MISSEDREASON Then
                tmSdf.sSchStatus = "M"
                slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                'Set missed reason
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    tmSdf.iMnfMissed = Val(slCode)
                Else
                    tmSdf.iMnfMissed = igDefaultMnfMissed
                End If
                'smMdSchStatus(1, imMdRowNo) = "M"
            ElseIf ilBoxNo = MISSEDTOCANCEL Then
                tmSdf.sSchStatus = "C"
                tmSdf.sTracer = "2"
                tmSdf.sAffChg = ""  'remove any post log changes
                ''tmSdf.sPtType = "0" 'No copy
                'tmSdf.iMnfMissed = 0
                ''smMdSchStatus(1, imMdRowNo) = "C"
                'Set missed reason
                slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    tmSdf.iMnfMissed = Val(slCode)
                Else
                    tmSdf.iMnfMissed = igDefaultMnfMissed
                End If
            ElseIf ilBoxNo = MISSEDTOHIDE Then
                tmSdf.sSchStatus = "H"
                tmSdf.sTracer = "2"
                tmSdf.sAffChg = ""  'remove any post log changes
                'tmSdf.sPtType = "0" 'No copy
                tmSdf.iMnfMissed = 0
                'smMdSchStatus(1, imMdRowNo) = "H"
            End If
            'ilFirst = False
            If ilUpdateSdf Then
                'Remove billed flag if moved to missed/cancelled or hidden
                If tgSpf.sInvAirOrder = "2" Then
                    If (ilBoxNo = SCHTOMISSED) Or (ilBoxNo = SCHTOCANCEL) Or (ilBoxNo = SCHTOHIDE) Then
                        tmSdf.sBill = "N"
                    End If
                    'Set bill flag if moved prior to last invoice date
                    If (ilBoxNo = MISSEDTOSCH) Or (ilBoxNo = BONUSTOSCH) Then
                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
                        tmChfSrchKey.lCode = tmSdf.lChfCode
                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            If tmChf.sBillCycle = "C" Then
                                gUnpackDateLong tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), llDate
                            ElseIf tmChf.sBillCycle = "W" Then
                                gUnpackDateLong tgSaf(0).iBLastWeeklyDate(0), tgSaf(0).iBLastWeeklyDate(1), llDate
                            Else
                                gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llDate
                            End If
                            If gDateValue(slSchDate) <= llDate Then
                                tmSdf.sBill = "Y"
                            End If
                        End If
                    End If
                End If
                tmSdf.iUrfCode = tgUrf(0).iCode
                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen) '  Update File
                If (ilBoxNo = MISSEDTOSCH) Then
                    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slSchDate
                    ilRetC = mAddIihf(tmSdf.iVefCode, tmSdf.lChfCode, slSchDate)
                End If
            Else
                ilRet = BTRV_ERR_NONE
            End If
            If ilRet = BTRV_ERR_CONFLICT Then
                ilRetC = btrAbortTrans(hmSdf)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmSdf)
            mUnblockDay
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
            mSaveRec = False
            imSdfChg = False
            Exit Function
        End If
        ilRet = btrEndTrans(hmSdf)
        mUnblockDay
        DoEvents
        If (ilBoxNo = SCHTOMISSED) Or (ilBoxNo = SCHTOCANCEL) Or (ilBoxNo = SCHTOHIDE) Then
            ilUpperBound = UBound(tgSave) - 1 'UBound(smSave, 2) - 1
            For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
                tgSave(ilLoop) = tgSave(ilLoop + 1)
                'tgSave(ilLoop).iShowInfoIndex = tgSave(ilLoop).iShowInfoIndex - 1
            Next ilLoop
            ReDim Preserve tgSave(0 To ilUpperBound) As SAVEINFO
            'For ilLoop = LBound(tgSave) To UBound(tgSave) - 1 Step 1
            For ilLoop = LBONE To UBound(tgSave) - 1 Step 1
                If tgSave(ilLoop).iShowInfoIndex >= imRowNo Then
                    tgSave(ilLoop).iShowInfoIndex = tgSave(ilLoop).iShowInfoIndex - 1
                End If
            Next ilLoop
            ilUpperBound = UBound(tgShow) - 1 'UBound(smSave, 2) - 1
            For ilLoop = imRowNo To ilUpperBound - 1 Step 1
                tgShow(ilLoop) = tgShow(ilLoop + 1)
                'tgShow(ilLoop).iSaveInfoIndex = tgShow(ilLoop).iSaveInfoIndex - 1
            Next ilLoop
            ReDim Preserve tgShow(0 To ilUpperBound) As SHOWINFO
            'For ilLoop = LBound(tgShow) To UBound(tgShow) - 1 Step 1
            For ilLoop = LBONE To UBound(tgShow) - 1 Step 1
                If tgShow(ilLoop).iSaveInfoIndex >= ilRowNo Then
                    tgShow(ilLoop).iSaveInfoIndex = tgShow(ilLoop).iSaveInfoIndex - 1
                End If
            Next ilLoop
            imListChgMode = True
            lbcMissed.ListIndex = -1
            imListChgMode = False
            pbcPosting.Cls
            mPaintPostTitle
            'Determine if missed should be reread (spot is for same advt as that which is shown)
'           If imAdvtSelectedIndex >= 0 Then
            If (ilUpdateSdf) Then
                lacMdFrame.Visible = False
                'If imAdvtSelectedIndex = 0 Then
                '    If (ilBoxNo = SCHTOMISSED) Then
                '        pbcMissed.Cls
                '        'ilRet = mReadMdSdfRec()
                '        pbcMissed_Paint
                '    End If
                'Else
                '    For ilLoop = 0 To UBound(tgAdvertiser) - 1 Step 1'Traffic!lbcAdvertiser.ListCount - 1 Step 1
                '        slNameCode = tgAdvertiser(ilLoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
                '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                '        ilAdvtCode = Val(slCode)
                '        If tmSdf.iAdfCode = ilAdvtCode Then
                '            If cbcAdvt.ListIndex <> ilLoop + 1 Then
                '                   cbcAdvt.ListIndex = ilLoop + 1
                '                   Exit For
                '            Else
                '                pbcMissed.Cls
                '                ilRet = mReadMdSdfRec()
                '                pbcMissed_Paint
                '                Exit For
                '            End If
                '        End If
                '    Next ilLoop
                'End If
                mUpdateMdShow
            End If
            '2/21/13: Allow Package spots to be cancelled
            If rbcType(1).Value Then
                lmMdStartdate = -1
                ilRet = mReadSdfRec(True)
            End If
            imSettingValue = True
            If UBound(tgShow) <= vbcPosting.LargeChange + 1 Then
                vbcPosting.Max = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
            Else
                vbcPosting.Max = UBound(tgShow) - vbcPosting.LargeChange
            End If
        End If
        If (ilBoxNo = MISSEDTOSCH) Then
            ilUpperBound = UBound(tgSave) + 1 'UBound(smSave, 2) + 1
            'ReDim Preserve smSave(1 To 10, 1 To ilUpperBound) As String
            ReDim Preserve tgSave(0 To ilUpperBound) As SAVEINFO
            ilRet = mCreateSaveSpotImage(ilUpperBound - 1, slAvDate, slAvTime, ilAnfCode, ilCount)

            ilUpperBound = UBound(tgShow) + 1 'UBound(smSave, 2) + 1
            'ReDim Preserve smShow(1 To 13, 1 To ilUpperBound) As String
            ReDim Preserve tgShow(0 To ilUpperBound) As SHOWINFO
            For ilLoop = ilUpperBound To imDropRowNo + 2 Step -1
                tgShow(ilLoop) = tgShow(ilLoop - 1)
                If ilLoop <> ilUpperBound Then
                    tgSave(tgShow(ilLoop).iSaveInfoIndex).iShowInfoIndex = ilLoop
                End If
            Next ilLoop
            mCreateShowImage UBound(tgSave) - 1, imDropRowNo + 1
            '2/21/13: Allow Package spots to be cancelled
            If rbcType(1).Value Then
                lmMdStartdate = -1
                ilRet = mReadSdfRec(True)
            Else
                pbcMissed.Cls
                DoEvents
                tgMdSdfRec(imSdfIndex).lSdfCode = 0
                tgMdSdfRec(imSdfIndex).lMissedDate = 999999
                tgMdSdfRec(imSdfIndex).sSchStatus = ""
                If tgMdShowInfo(imMdRowNo).iType = 0 Then   'Contract
                ElseIf tgMdShowInfo(imMdRowNo).iType = 2 Then   'Hidden
                    tgMdSaveInfo(imSaveIndex).iHiddenCount = tgMdSaveInfo(imSaveIndex).iHiddenCount - 1
                    If tgMdSaveInfo(imSaveIndex).iHiddenCount > 0 Then
                        slStr = Trim$(str$(tgMdSaveInfo(imSaveIndex).iHiddenCount))
                    Else
                        slStr = ""
                    End If
                ElseIf tgMdShowInfo(imMdRowNo).iType = 3 Then
                    tgMdSaveInfo(imSaveIndex).iCancelCount = tgMdSaveInfo(imSaveIndex).iCancelCount - 1
                    If tgMdSaveInfo(imSaveIndex).iCancelCount > 0 Then
                        slStr = Trim$(str$(tgMdSaveInfo(imSaveIndex).iCancelCount))
                    Else
                        slStr = ""
                    End If
                Else
                    tgMdSaveInfo(imSaveIndex).iMissedCount = tgMdSaveInfo(imSaveIndex).iMissedCount - 1
                    If tgMdSaveInfo(imSaveIndex).iMissedCount > 0 Then
                        slStr = Trim$(str$(tgMdSaveInfo(imSaveIndex).iMissedCount))
                    Else
                        slStr = ""
                    End If
                End If
                gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
                tgMdShowInfo(imMdRowNo).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
                If (tgMdSaveInfo(imSaveIndex).iMissedCount <= 0) And (tgMdSaveInfo(imSaveIndex).iHiddenCount <= 0) And (tgMdSaveInfo(imSaveIndex).iCancelCount <= 0) Then
                    tgMdSaveInfo(imSaveIndex).lWkMissed = 0
                    slStr = " "
                    gSetShow pbcMissed, slStr, tmMdCtrls(MDWKMISSINDEX)
                    tgMdShowInfo(imMdRowNo).sShow(MDWKMISSINDEX) = tmMdCtrls(MDWKMISSINDEX).sShow
                    tgMdShowInfo(imMdRowNo).iType = 0
                End If
            End If
            'ilUpperBound = UBound(smMdSave, 2) - 1
            'For ilLoop = imMdRowNo To ilUpperBound - 1 Step 1
            '    For ilIndex = LBound(smMdShow, 1) To UBound(smMdShow, 1) Step 1
            '        smMdShow(ilIndex, ilLoop) = smMdShow(ilIndex, ilLoop + 1)
            '    Next ilIndex
            '    For ilIndex = LBound(smMdSave, 1) To UBound(smMdSave, 1) Step 1
            '        smMdSave(ilIndex, ilLoop) = smMdSave(ilIndex, ilLoop + 1)
            '    Next ilIndex
            '    For ilIndex = LBound(smMdSchStatus, 1) To UBound(smMdSchStatus, 1) Step 1
            '        smMdSchStatus(ilIndex, ilLoop) = smMdSchStatus(ilIndex, ilLoop + 1)
            '    Next ilIndex
            '    lmMdRecPos(ilLoop) = lmMdRecPos(ilLoop + 1)
            'Next ilLoop
            'ReDim Preserve smMdShow(1 To 8, 1 To ilUpperBound) As String
            'ReDim Preserve smMdSave(1 To 1, 1 To ilUpperBound) As String
            'ReDim Preserve smMdSchStatus(1 To 1, 1 To ilUpperBound) As String
            'ReDim Preserve lmMdRecPos(1 To ilUpperBound) As Long
            'If UBound(smMdSave, 2) <= vbcMissed.LargeChange Then
            '' If this is used, there are probably 0 or 1 records
            '    vbcMissed.Max = LBound(smMdSave, 2)
            'Else
            '' Saves, what amounts to, the count of records just retrieved
            '    vbcMissed.Max = UBound(smMdSave, 2) - vbcMissed.LargeChange
            'End If
            DoEvents
            pbcMissed_Paint
            pbcPosting.Cls
            mPaintPostTitle
        End If
        If (ilBoxNo = BONUSTOSCH) Then
            ilUpperBound = UBound(tgSave) + 1 'UBound(smSave, 2) + 1
            'ReDim Preserve smSave(1 To 10, 1 To ilUpperBound) As String
            ReDim Preserve tgSave(0 To ilUpperBound) As SAVEINFO
            ilRet = mCreateSaveSpotImage(ilUpperBound - 1, slAvDate, slAvTime, ilAnfCode, ilCount)

            ilUpperBound = UBound(tgShow) + 1 'UBound(smSave, 2) + 1
            'ReDim Preserve smShow(1 To 13, 1 To ilUpperBound) As String
            ReDim Preserve tgShow(0 To ilUpperBound) As SHOWINFO
            For ilLoop = ilUpperBound To imDropRowNo + 2 Step -1
                tgShow(ilLoop) = tgShow(ilLoop - 1)
                If ilLoop <> ilUpperBound Then
                    tgSave(tgShow(ilLoop).iSaveInfoIndex).iShowInfoIndex = ilLoop
                End If
            Next ilLoop
            mCreateShowImage UBound(tgSave) - 1, imDropRowNo + 1
            pbcPosting.Cls
            mPaintPostTitle
        End If
        'If ilBoxNo = MISSEDREASON Then
        '    lbcMissed.ListIndex = -1
        '    slNameCode = tmMissedCode(imMissedListIndex).sKey  'lbcMissedCode.List(imMissedListIndex)
        '    'Set missed reason
        '    ilRet = gParseItem(slNameCode, 1, "\", smMdSave(1, imMdRowNo))
        '    gSetShow pbcMissed, smMdSave(1, imMdRowNo), tmMdCtrls(MDREASONINDEX)
        '    smMdShow(MDREASONINDEX, imMdRowNo) = tmMdCtrls(MDREASONINDEX).sShow
        '    pbcMissed.Cls
        '    pbcMissed_Paint
        '    imSdfChg = False
        '    mSaveRec = True
        '    Exit Function
        'End If
        'If (ilBoxNo = MISSEDTOCANCEL) Or (ilBoxNo = MISSEDTOHIDE) Or ((slSchStatus = "H") And (ilBoxNo = MISSEDREASON)) Or ((slSchStatus = "C") And (ilBoxNo = MISSEDREASON)) Then
        '4/26/10:  Counts don't need to be updated if just changing the reason
        'If (ilBoxNo = MISSEDTOCANCEL) Or (ilBoxNo = MISSEDTOHIDE) Or (ilBoxNo = MISSEDREASON) Then
        If (ilBoxNo = MISSEDTOCANCEL) Or (ilBoxNo = MISSEDTOHIDE) Then
            '2/21/13: Allow Package spots to be cancelled
            If rbcType(1).Value Then
                lmMdStartdate = -1
                ilRet = mReadSdfRec(True)
            Else
                If (slSchStatus = "H") And (tmSdf.sSchStatus <> "H") Then
                    tgMdSaveInfo(imSaveIndex).iHiddenCount = tgMdSaveInfo(imSaveIndex).iHiddenCount - 1
                    If tgMdSaveInfo(imSaveIndex).iHiddenCount > 0 Then
                        slStr = Trim$(str$(tgMdSaveInfo(imSaveIndex).iHiddenCount))
                    Else
                        slStr = ""
                    End If
                ElseIf (slSchStatus = "C") And (tmSdf.sSchStatus <> "C") Then
                    tgMdSaveInfo(imSaveIndex).iCancelCount = tgMdSaveInfo(imSaveIndex).iCancelCount - 1
                    If tgMdSaveInfo(imSaveIndex).iCancelCount > 0 Then
                        slStr = Trim$(str$(tgMdSaveInfo(imSaveIndex).iCancelCount))
                    Else
                        slStr = ""
                    End If
                ElseIf (slSchStatus = "M") And (tmSdf.sSchStatus <> "M") Then
                    tgMdSaveInfo(imSaveIndex).iMissedCount = tgMdSaveInfo(imSaveIndex).iMissedCount - 1
                    If tgMdSaveInfo(imSaveIndex).iMissedCount > 0 Then
                        slStr = Trim$(str$(tgMdSaveInfo(imSaveIndex).iMissedCount))
                    Else
                        slStr = ""
                    End If
                End If
                gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
                tgMdShowInfo(imMdRowNo).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
                If (tgMdSaveInfo(imSaveIndex).iMissedCount <= 0) And (tgMdSaveInfo(imSaveIndex).iHiddenCount <= 0) And (tgMdSaveInfo(imSaveIndex).iCancelCount <= 0) Then
                    tgMdSaveInfo(imSaveIndex).lWkMissed = 0
                    slStr = " "
                    gSetShow pbcMissed, slStr, tmMdCtrls(MDWKMISSINDEX)
                    tgMdShowInfo(imMdRowNo).sShow(MDWKMISSINDEX) = tmMdCtrls(MDWKMISSINDEX).sShow
                    tgMdShowInfo(imMdRowNo).iType = 0
                End If
                mUpdateMdShow
            End If
            DoEvents
            imSdfChg = False
            mSaveRec = True
            Exit Function
        ElseIf (ilBoxNo = MISSEDREASON) Then
            DoEvents
            imSdfChg = False
            mSaveRec = True
            Exit Function
        End If
        imSdfChg = False
        DoEvents
        pbcPosting_Paint
    End If
    mSaveRec = True
    imSdfChg = False
    Exit Function

    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Function mSetShow(ilInBoxNo As Integer) As Integer
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    '6/16/11
    Dim slStr1 As String
    Dim slStr2 As String
    Dim slCopyISCI As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slCopy As String
    Dim slISCI As String
    Dim slProduct As String
    Dim slCpfProduct As String
    Dim ilRet As Integer
    Dim ilTimeChanged As Integer
    Dim flWidth As Single
    Dim ilIndex As Integer
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim ilMRowNo As Integer
    pbcArrow.Visible = False
    lacPtFrame.Visible = False
    ilBoxNo = ilInBoxNo
    ilMRowNo = imRowNo
    On Error GoTo mSetShowErr:
    'If (imRowNo < LBound(tgShow)) Or (imRowNo > UBound(tgShow)) Then
    If (imRowNo < LBONE) Or (imRowNo > UBound(tgShow)) Then
        mSetShow = True
        Exit Function
    End If
    On Error GoTo 0
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        mSetShow = True
        Exit Function
    End If
    ilRowNo = tgShow(ilMRowNo).iSaveInfoIndex
    Select Case ilBoxNo 'Branch on box type (control)
        Case DATEINDEX, TIMEINDEX 'Vehicle
            If (rbcType(0).Value) And (smSpotType <> "O") And (smSpotType <> "C") Then
                '6/16/11
                ''slStr = smDTSave(imDTAIRTIMEINDEX)
                ''5/20/11: Time only
                'slStr = edcDropDown.Text
                'If Not gValidTime(slStr) Then
                '    Beep
                '    edcDTDropDown.SetFocus
                '    mSetShow = False
                '    Exit Function
                'End If
                ''plcCalendar.Visible = False
                ''cmcDTDropDown.Visible = False
                ''edcDTDropDown.Visible = False  'Set visibility
                ''plcDT.Visible = False
                ''pbcDT.Visible = False
                '6/16/11
                plcCalendar.Visible = False
                cmcDTDropDown.Visible = False
                edcDTDropDown.Visible = False  'Set visibility
                plcDT.Visible = False
                pbcDT.Visible = False
                plcTme.Visible = False
                '6/16/11
                'edcDropDown.Visible = False
                'cmcDropDown.Visible = False
                ilTimeChanged = False
                '6/16/11
                'If slStr <> "" Then
                '    slStr = gFormatTime(slStr, "A", "1")
                If smDTSave(imDTAIRTIMEINDEX) <> "" Then
                    slStr1 = gFormatTime(smDTSave(imDTAIRTIMEINDEX), "A", "1")
                    slStr2 = gFormatTime(smDTSave(imDTAVAILTIMEINDEX), "A", "1")
                    '''If gTimeToCurrency(smSave(SAVTIMEINDEX, ilMRowNo), False) <> gTimeToCurrency(slStr, False) Then
                    ''If (gDateValue(tgSave(ilRowNo).sAirDate) <> gDateValue(smDTSave(DTDATEINDEX))) Or (gTimeToCurrency(Trim$(tgSave(ilRowNo).sAirTime), False) <> gTimeToCurrency(slStr, False)) Then
                    '6/16/11
                    'If (gTimeToCurrency(Trim$(tgSave(ilRowNo).sAirTime), False) <> gTimeToCurrency(slStr, False)) Then
                    If (gTimeToCurrency(Trim$(tgSave(ilRowNo).sAirTime), False) <> gTimeToCurrency(slStr1, False)) Or (gTimeToCurrency(Trim$(tgSave(ilRowNo).sSchTime), False) <> gTimeToCurrency(slStr2, False)) Or (gDateValue(tgSave(ilRowNo).sAirDate) <> gDateValue(smDTSave(DTDATEINDEX))) Then
                        imSdfChg = True   ' User CHANGED the time
                        ilTimeChanged = True
                        ''imSdfAnyChg = True
                        'If lbcAvailTimes.ListIndex >= 0 Then
                        '    'smSave(SAVTIMEINDEX, ilMRowNo) = lbcAvailTimes.List(lbcAvailTimes.ListIndex)    'edcDropDown.Text
                        '    tgSave(ilRowNo).sAirTime = lbcAvailTimes.List(lbcAvailTimes.ListIndex)    'edcDropDown.Text
                        'Else
                        '    'smSave(SAVTIMEINDEX, ilMRowNo) = slStr    'edcDropDown.Text
                            '6/16/11
                            'tgSave(ilRowNo).sAirTime = slStr    'edcDropDown.Text
                            tgSave(ilRowNo).sAirTime = slStr1    'edcDropDown.Text
                        'End If
                        '6/16/11
                        tgSave(ilRowNo).sSchTime = slStr2
                        tgSave(ilRowNo).sSchDate = smDTSave(DTDATEINDEX)
                        tgSave(ilRowNo).sAirDate = smDTSave(DTDATEINDEX)
                        If Not mSaveRec() Then ' Disk record updated
                            tgSave(ilRowNo).sAirDate = smSvDate
                            tgSave(ilRowNo).sAirTime = smSvAirTime
                            '6/16/11
                            tgSave(ilRowNo).sSchDate = smSvDate
                            tgSave(ilRowNo).sSchTime = smSvAvailTime
                            smDTSave(DTDATEINDEX) = Trim$(tgSave(ilRowNo).sAirDate)    'smSave(SAVDATEINDEX, imRowNo)
                            smDTSave(imDTAIRTIMEINDEX) = Trim$(tgSave(ilRowNo).sAirTime) 'smSave(SAVTIMEINDEX, imRowNo)
                            mSetShow = False
                            Exit Function
                        End If
                        '6/16/11
                        ''5/20/11
                        ''slStr = smDTSave(imDTAIRTIMEINDEX)
                        'slStr = edcDropDown.Text
                        'slStr = gFormatTime(slStr, "A", "1")
                        'gSetShow pbcPosting, slStr, tmCtrls(ilBoxNo)
                        gSetShow pbcPosting, slStr1, tmCtrls(ilBoxNo)
                        tgShow(ilMRowNo).sShow(TIMEINDEX) = tmCtrls(ilBoxNo).sShow
                        slStr = smDTSave(DTDATEINDEX)
                        gSetShow pbcPosting, slStr, tmCtrls(ilBoxNo)
                        tgShow(ilMRowNo).sShow(DATEINDEX) = tmCtrls(ilBoxNo).sShow
                    End If
                End If
            Else
                '6/16/11
                ''5/20/11
                ''plcDT.Visible = False
                ''pbcDT.Visible = False
                plcDT.Visible = False
                pbcDT.Visible = False
                ilTimeChanged = False
                '6/16/11
                ''slStr = smDTSave(imDTAIRTIMEINDEX)
                'slStr = edcDropDown.Text
                'If Not gValidTime(slStr) Then
                '    Beep
                '    'edcDTDropDown.SetFocus
                '    edcDropDown.SetFocus
                '    mSetShow = False
                '    Exit Function
                'End If
                'slStr = gFormatTime(slStr, "A", "1")
                slStr = gFormatTime(smDTSave(imDTAIRTIMEINDEX), "A", "1")
                '6/16/11
                ''5/20/11
                '''If (gTimeToCurrency(smSave(SAVTIMEINDEX, ilMRowNo), False) <> gTimeToCurrency(slStr, False)) Or (gDateValue(smSave(SAVDATEINDEX, ilMRowNo)) <> gDateValue(smDTSave(DTDATEINDEX))) Then
                ''If (gDateValue(tgSave(ilRowNo).sAirDate) <> gDateValue(smDTSave(DTDATEINDEX))) Or (gTimeToCurrency(Trim$(tgSave(ilRowNo).sAirTime), False) <> gTimeToCurrency(slStr, False)) Then
                'If (gTimeToCurrency(Trim$(tgSave(ilRowNo).sAirTime), False) <> gTimeToCurrency(slStr, False)) Then
                If (gDateValue(tgSave(ilRowNo).sAirDate) <> gDateValue(smDTSave(DTDATEINDEX))) Or (gTimeToCurrency(Trim$(tgSave(ilRowNo).sAirTime), False) <> gTimeToCurrency(slStr, False)) Then
                    imSdfChg = True   ' User CHANGED the time
                    ilTimeChanged = True
                    'imSdfAnyChg = True
                    'smSave(SAVTIMEINDEX, ilMRowNo) = slStr
                    tgSave(ilRowNo).sAirTime = slStr
                    'smSave(SAVDATEINDEX, ilMRowNo) = smDTSave(DTDATEINDEX)
                    tgSave(ilRowNo).sAirDate = smDTSave(DTDATEINDEX)
                    If Not mSaveRec() Then ' Disk record updated
                        mSetShow = False
                        Exit Function
                    End If
                    gSetShow pbcPosting, slStr, tmCtrls(ilBoxNo)
                    tgShow(ilMRowNo).sShow(TIMEINDEX) = tmCtrls(ilBoxNo).sShow
                    '6/16/11
                    slStr = smDTSave(DTDATEINDEX)
                    gSetShow pbcPosting, slStr, tmCtrls(ilBoxNo)
                    tgShow(ilMRowNo).sShow(DATEINDEX) = tmCtrls(ilBoxNo).sShow
                End If
            End If
        'Copy & Isci - These are treated as one field on the picturebox
        Case COPYINDEX
            slISCI = ""
            slProduct = ""
            slCpfProduct = ""
            If imTZCopyAllowed Then
                plcTZCopy.Visible = False
                pbcTZCopy.Visible = False
                If (StrComp(smTZSave(2, 1), "[None]", 1) <> 0) And (Trim$(smTZSave(2, 1)) <> "") Then
                    slCopyISCI = smTZSave(2, 1) 'Take "All" rows for copy
                    tgShow(ilMRowNo).sShow(TZONEINDEX) = " "
                Else
                    If (StrComp(smTZSave(2, 2), "[None]", 1) <> 0) And (Trim$(smTZSave(2, 2)) <> "") Then
                        slCopyISCI = smTZSave(2, 2) 'Take "Other" rows for copy
                        tgShow(ilMRowNo).sShow(TZONEINDEX) = "4"
                    Else
                        slCopyISCI = "" 'Take "All" rows for copy
                        tgShow(ilMRowNo).sShow(TZONEINDEX) = " "
                        For ilIndex = 3 To imNoZones Step 1
                            If Trim$(smTZSave(2, ilIndex)) <> "" Then
                                slCopyISCI = smTZSave(2, ilIndex) 'Take rows with copy
                                tgShow(ilMRowNo).sShow(TZONEINDEX) = "4"
                                Exit For
                            End If
                        Next ilIndex
                    End If
                End If
                If tgSpf.sUseCartNo <> "N" Then
                    ilRet = gParseItem(slCopyISCI, 1, " ", slCopy)
                Else
                    slCopy = ""
                End If
                'ilRet = gParseItem(slCopyISCI, 2, " ", slISCI)
                'ilRet = gParseItem(slCopyISCI, 3, " ", slProduct)
                'gFindPartialMatch slCopyISCI, 1, Len(slCopyISCI), lbcCopyNm
                slProduct = ""
                gFindMatch slCopyISCI, 1, lbcCopyNm
                If gLastFound(lbcCopyNm) > 0 Then
                    slNameCode = tmCopyNmCode(gLastFound(lbcCopyNm) - 1).sKey    'lbcCopyNmCode.List(gLastFound(lbcCopyNm) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If ilRet = CP_MSG_NONE Then
                        tmCifSrchKey.lCode = Val(slCode)
                        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            If (tgSpf.sUseCartNo = "N") Or (tmCif.iMcfCode = 0) Then
                                slCopy = ""
                            End If
                            If tmCif.lcpfCode > 0 Then
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    slISCI = Trim$(tmCpf.sISCI)
                                    slProduct = Trim$(tmCpf.sName)
                                    slCpfProduct = slProduct
                                End If
                            End If
                        End If
                    End If
                End If
                If Trim$(slISCI) <> "" Then
                    'imPostSpotInfo(2, ilMRowNo) = False
                    tgSave(ilRowNo).iISCI = False
                Else
                    'imPostSpotInfo(2, ilMRowNo) = True
                    tgSave(ilRowNo).iISCI = True
                End If
                '     put in the test for slproduct = blank
                If Len(Trim$(slProduct)) <= 0 Then
                    slProduct = Trim$(tgSave(ilRowNo).sProd)   'smSave(SAVPRODNAMEINDEX, ilMRowNo)
                End If
                gSetShow pbcPosting, slCopy, tmCtrls(ilBoxNo)
                tgShow(ilMRowNo).sShow(COPYINDEX) = tmCtrls(ilBoxNo).sShow
                gSetShow pbcPosting, slISCI, tmCtrls(ISCIINDEX)
                tgShow(ilMRowNo).sShow(ISCIINDEX) = tmCtrls(ISCIINDEX).sShow
                imSdfChg = True 'Assume time zone copy changed- test within saverec
            Else
                lbcCopyNm.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False

                slCopyISCI = edcDropDown.Text
                'ilPos = InStr(slCopyISCI, "(Purged)")
                'If ilPos > 0 Then
                '    slCopyISCI = Left$(slCopyISCI, ilPos - 1)
                'Else
                '    ilPos = InStr(slCopyISCI, "(Reused)")
                '    If ilPos > 0 Then
                '        slCopyISCI = Left$(slCopyISCI, ilPos - 1)
                '    End If
                'End If
                If tgSpf.sUseCartNo <> "N" Then
                    ilRet = gParseItem(slCopyISCI, 1, " ", slCopy)
                Else
                    slCopy = ""
                End If
                'Since ISCI can have blanks and product might not exist, read CIF and CPF
                'to obtain ISCI and Product
                'ilRet = gParseItem(slCopyISCI, 2, " ", slISCI)
                'ilRet = gParseItem(slCopyISCI, 3, " ", slProduct)
                'gFindPartialMatch slCopyISCI, 1, Len(slCopyISCI), lbcCopyNm
                slProduct = ""
                gFindMatch slCopyISCI, 1, lbcCopyNm
                If gLastFound(lbcCopyNm) > 0 Then
                    slNameCode = tmCopyNmCode(gLastFound(lbcCopyNm) - 1).sKey    'lbcCopyNmCode.List(gLastFound(lbcCopyNm) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If ilRet = CP_MSG_NONE Then
                        tmCifSrchKey.lCode = Val(slCode)
                        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            If (tgSpf.sUseCartNo = "N") Or (tmCif.iMcfCode = 0) Then
                                slCopy = ""
                            End If
                            If tmCif.lcpfCode > 0 Then
                                tmCpfSrchKey.lCode = tmCif.lcpfCode
                                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    slISCI = Trim$(tmCpf.sISCI)
                                    slProduct = Trim$(tmCpf.sName)
                                    slCpfProduct = slProduct
                                End If
                            End If
                        End If
                    End If
                End If
                '     put in the test for slproduct = blank
                If Len(Trim$(slProduct)) <= 0 Then
                    slProduct = Trim$(tgSave(ilRowNo).sProd)  'smSave(SAVPRODNAMEINDEX, ilMRowNo)
                End If
                gSetShow pbcPosting, slCopy, tmCtrls(ilBoxNo)
                tgShow(ilMRowNo).sShow(COPYINDEX) = tmCtrls(ilBoxNo).sShow
                gSetShow pbcPosting, slISCI, tmCtrls(ISCIINDEX)
                tgShow(ilMRowNo).sShow(ISCIINDEX) = tmCtrls(ISCIINDEX).sShow
                If Trim$(slISCI) <> "" Then
                    'imPostSpotInfo(2, ilMRowNo) = False
                    tgSave(ilRowNo).iISCI = False
                Else
                    'imPostSpotInfo(2, ilMRowNo) = True
                    tgSave(ilRowNo).iISCI = True
                End If
                ' Test to see if copy has changed
                'If Trim$(smSave(SAVALLCOPYINDEX, ilMRowNo)) <> Trim$(edcDropDown.Text) Then
                If Trim$(smCopy) <> Trim$(edcDropDown.Text) Then
                    imSdfChg = True
                '    imSdfAnyChg = True
                End If
                'smSave(SAVALLCOPYINDEX, ilMRowNo) = edcDropDown.Text
                'tgSave(ilRowNo).sMedia = edcDropDown.Text
                tgSave(ilRowNo).sCopy = slCopy
                tgSave(ilRowNo).sISCI = slISCI
                tgSave(ilRowNo).sCopyProduct = slCpfProduct
                'ilPos = InStr(smSave(SAVALLCOPYINDEX, ilMRowNo), "(Purged)")
                'If ilPos > 0 Then
                '    smSave(SAVALLCOPYINDEX, ilMRowNo) = Left$(smSave(SAVALLCOPYINDEX, ilMRowNo), ilPos - 1)
                'Else
                '    ilPos = InStr(smSave(SAVALLCOPYINDEX, ilMRowNo), "(Reused)")
                '    If ilPos > 0 Then
                '        smSave(SAVALLCOPYINDEX, ilMRowNo) = Left$(smSave(SAVALLCOPYINDEX, ilMRowNo), ilPos - 1)
                '    End If
                'End If
                imCopyNmListIndex = lbcCopyNm.ListIndex
            End If
            If Len(slProduct) <= 0 Then  ' No Product Name
                ' Use saved Advertiser Name
'                smSave(SAVPRODNAMEINDEX, ilMRowNo) = smSave(SAVADVTNAMEINDEX, ilMRowNo)
                gSetShow pbcPosting, Trim$(tgSave(ilRowNo).sAdvtName), tmCtrls(ADVTINDEX)
                tgShow(ilMRowNo).sShow(ADVTINDEX) = tmCtrls(ADVTINDEX).sShow ' So can be shown
            Else
                ' Get Advertiser field width, divide by 2 (so can append Product)
                flWidth = tmCtrls(ADVTINDEX).fBoxW ' First save copy of real width
                tmCtrls(ADVTINDEX).fBoxW = tmCtrls(ADVTINDEX).fBoxW / 2
                ' Put no more than 1/2 advertiser name into control array
                gSetShow pbcPosting, Trim$(tgSave(ilRowNo).sAdvtName), tmCtrls(ADVTINDEX)
                ' concantinate  "\ProductName" to it
                slStr = tmCtrls(ADVTINDEX).sShow & "/" & slProduct
                ' Trim the concatinated names to fit
                tmCtrls(ADVTINDEX).fBoxW = flWidth  ' Restore width
                gSetShow pbcPosting, slStr, tmCtrls(ADVTINDEX)
                ' Save concatinated result in the ShowArray
                tgShow(ilMRowNo).sShow(ADVTINDEX) = tmCtrls(ADVTINDEX).sShow
            End If
'            pbcPosting_Paint ' cause a paint event
            If Not mSaveRec() Then ' Disk record updated
                mSetShow = False
                Exit Function
            End If
        Case PRICEINDEX 'Price index
            pbcPrice.Visible = False
'            If tgSave(ilRowNo).iPrice = 0 Then 'imSave(1, ilMRowNo) = 0 Then
'                slStr = tgSave(ilRowNo).sPrice  'smSave(SAVPRICEINDEX, ilMRowNo)
'                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
'            Else
'                slStr = "N/C"
'            End If
            Select Case tgSave(ilRowNo).iPrice
                Case 1
                    slStr = "+ Fill"
                Case 2
                    slStr = "- Fill"
                Case Else
                    If tmAdf.sBonusOnInv <> "N" Then
                        slStr = "+ Fill"
                    Else
                        slStr = "- Fill"
                    End If
            End Select


            gSetShow pbcPosting, slStr, tmCtrls(ilBoxNo)
            tgShow(ilMRowNo).sShow(PRICEINDEX) = tmCtrls(ilBoxNo).sShow
            'If imSave(1, ilMRowNo) <> imSave(2, ilMRowNo) Then
            If tgSave(ilRowNo).iPrice <> tgSave(ilRowNo).iSvPrice Then
                imSdfChg = True   ' User CHANGED the date
            '    imSdfAnyChg = True
            End If
            'imSave(2, ilMRowNo) = imSave(1, ilMRowNo)
            tgSave(ilRowNo).iSvPrice = tgSave(ilRowNo).iPrice
            If Not mSaveRec() Then
                mSetShow = False
                Exit Function
            End If
    End Select
    mSetShow = True
    Exit Function
mSetShowErr:
    mSetShow = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mShowInfo                       *
'*                                                     *
'*             Created:5/13/94       By:D. Hannifan    *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show Sdf information           *
'*                                                     *
'*******************************************************
Private Sub mShowInfo()
    Dim ilRet As Integer
    Dim llSdfRecPos As Long
    Dim slSchTime As String
    Dim slSchDate As String
    Dim slMissedDate As String
    Dim slMissedTime As String
    Dim slDate As String
    Dim slDays As String
    Dim slStr As String
    Dim ilRowNo As Integer

    ilRowNo = tgShow(imButtonRow).iSaveInfoIndex
    llSdfRecPos = tgSave(ilRowNo).lSdfRecPos    'Val(smSave(SAVRECPOSINDEX, imButtonRow))
    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        plcInfo.Visible = False
        Exit Sub
    End If
    lacInfo(1).Visible = False
    slSchDate = Trim$(tgSave(ilRowNo).sSchDate)
    slSchTime = Trim$(tgSave(ilRowNo).sSchTime)
    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
        tmSmfSrchKey2.lCode = tmSdf.lCode
        ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slMissedDate
            gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slMissedTime
            tmVefSrchKey.iCode = tmSmf.iOrigSchVef
            ilRet = btrGetEqual(hmVef, tmOrigVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If tmSdf.sSpotType <> "X" Then
                    If slSchTime = "" Then
                        lacInfo(0).Caption = Trim$(tmOrigVef.sName) & ", Missed " & slMissedDate & " @" & slMissedTime
                    Else
                        lacInfo(0).Caption = "Scheduled " & slSchDate & " @" & slSchTime & " " & Trim$(tmOrigVef.sName) & ", Missed " & slMissedDate & " @" & slMissedTime
                    End If
                Else
                    lacInfo(0).Caption = "Fill Spot " & Trim$(tmOrigVef.sName)
                End If
                plcInfo.Visible = True
            Else
                plcInfo.Visible = False
            End If
        Else
            plcInfo.Visible = False
        End If
    Else
        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
        ilRet = mReadChfClfRdfCffRec(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.lFsfCode, tmSdf.iGameNo, slDate)
        If Not ilRet Then
            plcInfo.Visible = False
            Exit Sub
        End If
        If tmCff.sDelete = "Y" Then 'Flight not found
            If slSchTime = "" Then
                plcInfo.Visible = False
                Exit Sub
            End If
            lacInfo(0).Caption = "Scheduled  " & slSchDate & " @" & slSchTime
            plcInfo.Visible = True
        Else
            slDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slStr)
            If slSchTime = "" Then
                lacInfo(0).Caption = slDays
            Else
                lacInfo(0).Caption = "Scheduled " & slSchDate & " @" & slSchTime & " Allowed Days " & slDays
            End If
            plcInfo.Visible = True
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'

    imTerminate = False
    Screen.MousePointer = vbDefault
    'Unload IconTraf
    igManUnload = YES
    Unload PostLog
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTZEnableBox                    *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mTZEnableBox(ilBoxNo As Integer)
'
'   mTZEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBTZCtrls) Or (ilBoxNo > UBound(tmTZCtrls)) Then
        Exit Sub ' Bogus box number so get out
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case ZONEINDEX
            pbcTZZone.Width = tmTZCtrls(ilBoxNo).fBoxW
            gMoveTableCtrl pbcTZCopy, pbcTZZone, tmTZCtrls(ZONEINDEX).fBoxX, tmTZCtrls(ZONEINDEX).fBoxY + (imTZRowNo - 1) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            If imTZSave < 0 Then
                imTZSave = 0
            End If
            pbcTZZone_Paint
            pbcTZZone.Visible = True
            pbcTZZone.SetFocus
        Case TZCOPYINDEX 'Copy or ISCI were selected by user
            mCopyPop ' populate lbcCopyNm and lbcCopyNmCode for the selected advertiser
            ' Size the listbox to fit this row
            lbcCopyNm.height = gListBoxHeight(lbcCopyNm.ListCount, 6)
            lbcCopyNm.Width = tmTZCtrls(TZCOPYINDEX).fBoxW
            ' Size the editbox to fit this row
            edcTZDropDown.Width = tmTZCtrls(TZCOPYINDEX).fBoxW - cmcTZDropDown.Width
            edcTZDropDown.MaxLength = 62
            ' Move the editbox (and the cmc control) into position
            gMoveTableCtrl pbcTZCopy, edcTZDropDown, tmTZCtrls(TZCOPYINDEX).fBoxX, tmTZCtrls(TZCOPYINDEX).fBoxY + (imTZRowNo - 1) * (fgBoxGridH + 15)
            cmcTZDropDown.Move edcTZDropDown.Left + edcTZDropDown.Width, edcTZDropDown.Top
            ' Find this COPY data in lbcCopyNm (might contain (Purged) or (Reused) at end within list box
            'gFindPartialMatch smTZSave(TZCOPYINDEX, imTZRowNo), 0, Len(smTZSave(TZCOPYINDEX, imTZRowNo)), lbcCopyNm
            gFindMatch smTZSave(TZCOPYINDEX, imTZRowNo), 0, lbcCopyNm
            imChgMode = True ' Turn on the switch
            If gLastFound(lbcCopyNm) >= 0 Then
                ' An entry was found so select it and put
                ' its data in the dropdown textbxo
                lbcCopyNm.ListIndex = gLastFound(lbcCopyNm)
            Else ' No data found so re-display the last good data
                If smTZSave(TZCOPYINDEX, imTZRowNo) = "" Then
                    lbcCopyNm.ListIndex = 0 ' no copy found
                Else
                lbcCopyNm.ListIndex = -1
                End If
            End If
            imComboBoxIndex = lbcCopyNm.ListIndex
            imCopyNmListIndex = imComboBoxIndex
            If lbcCopyNm.ListIndex >= 0 Then
                edcTZDropDown.Text = lbcCopyNm.List(lbcCopyNm.ListIndex)
            Else
                edcTZDropDown.Text = ""
            End If
            imChgMode = False
'            If edcTZDropDown.Top + edcTZDropDown.Height + lbcCopyNm.Height < cmcDone.Top Then
                lbcCopyNm.Move edcTZDropDown.Left, edcTZDropDown.Top + edcTZDropDown.height
'            Else
'                lbcCopyNm.Move edcDropDown.Left, edcDropDown.Top - lbcCopyNm.Height
'            End If
            edcTZDropDown.SelStart = 0
            edcTZDropDown.SelLength = Len(edcTZDropDown.Text)
            edcTZDropDown.Visible = True
            cmcTZDropDown.Visible = True
            edcTZDropDown.SetFocus
    End Select
    ' Save these in case user clicks another box, thus bypassing pbcTab
End Sub
Private Sub mTZReadCopyData()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    Dim slProductName As String
    Dim slCopy As String
    Dim slISCI As String
    Dim ilRowNo As Integer
    ' Re-initialize the data storage arrays to save Zone and Copy/ISCI/Product
'    ReDim smTZShow(1 To 2, 1 To 6) As String
'    ReDim smTZSave(1 To 2, 1 To 6) As String
    ' Read SDF to get the sPtType code
    ilRowNo = tgShow(imRowNo).iSaveInfoIndex
    llRecPos = tgSave(ilRowNo).lSdfRecPos   'Val(smSave(SAVRECPOSINDEX, imRowNo))
    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
    On Error GoTo mTZReadCopyDataErr
    gBtrvErrorMsg ilRet, "mTZReadCopyData (btrGetEqual:SdF)", PostLog
    On Error GoTo 0
    ' Read CHF (Contract File) to get ProductName from sProduct
'    tmChfSrchKey.lCode = tmSdf.lChfCode  ' Contract Hdr File Code
'    ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'    On Error GoTo mTZReadCopyDataErr
'    gBtrvErrorMsg ilRet, "mTZReadCopyDataErr (btrGetEqual:Contract)", PostLog
'    On Error GoTo 0
'    slProductName = Trim$(tmChf.sProduct) ' save the Product name

'    slNameCode = Traffic!lbcUserVehicle.List(imVehSelectedIndex)
'    ilRet = gParseItem(slNameCode, 2, "\", slCode)
'    On Error GoTo mTZReadCopyDataErr
'    gCPErrorMsg ilRet, "mGetVehIndex (gParseItem field 2: Vehicle)", PostLog
'    On Error GoTo 0
'    imVefCode = Val(slCode)
'    imVpfIndex = gVpfFind(PostLog, imVefCode)
    ' Put all of the Time Zones from Vehicle Option into the Save/Show Arrays
    ilUpperBound = 1
    imTZSave = -1
    smTZSave(2, 1) = "[None]"
    smTZShow(2, 1) = "[None]"
    For ilLoop = 2 To 8 Step 1
        smTZSave(2, ilLoop) = ""
        smTZShow(2, ilLoop) = ""
    Next ilLoop
    If tmSdf.sPtType = "0" Then ' None Assigned
        imTZSave = 0
        Exit Sub
    End If
    ilUpperBound = 1
    If tmSdf.sPtType = "1" Then ' All Time Zones have same copy
        imTZSave = 0
        'smTZShow(1, ilUpperBound) = "[All]"
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mTZReadCopyDataErr
        gBtrvErrorMsg ilRet, "mTZReadCopyData (btrGetEqual:CIF, Single)", PostLog
        On Error GoTo 0

        ' Read CPF using lCpfCode from CIF to get COPY data
        If tmCif.lcpfCode > 0 Then
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mTZReadCopyDataErr
            gBtrvErrorMsg ilRet, "mTZReadCopyDataErr (btrGetEqual:CPF)", PostLog
            On Error GoTo 0
            slISCI = Trim$(tmCpf.sISCI)  ' ISCI Code
            ' See if sName is Blank
'            If (Len(Trim$(tmCpf.sName)) <> 0) Then
                slProductName = Trim$(tmCpf.sName)
'            End If
        Else
            slISCI = ""
            slProductName = ""
        End If
        ' Concatinate Copy from Media Code, Inv. Name
        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
            tmMcfSrchKey.iCode = tmCif.iMcfCode
            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mTZReadCopyDataErr
            gBtrvErrorMsg ilRet, "mTZReadCopyData (btrGetEqual:MCF)", PostLog
            On Error GoTo 0
            ' Media Code is tmMcf.sName
            slCopy = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
            ' Save COPY/ISCI/PRODUCT
            If slProductName <> "" Then
                If Len(slISCI) > 0 Then
                    smTZSave(2, 1) = slCopy & " " & slISCI & " " & slProductName
                    smTZShow(2, 1) = slCopy & " " & slISCI & " " & slProductName
                Else
                    smTZSave(2, 1) = slCopy & " " & slProductName
                    smTZShow(2, 1) = slCopy & " " & slProductName
                End If
            Else
                If Len(slISCI) > 0 Then
                    smTZSave(2, 1) = slCopy & " " & slISCI ' & " " & slProductName
                    smTZShow(2, 1) = slCopy & " " & slISCI ' & " " & slProductName
                Else
                    smTZSave(2, 1) = slCopy '& " " & slISCI ' & " " & slProductName
                    smTZShow(2, 1) = slCopy '& " " & slISCI ' & " " & slProductName
                End If
            End If
        Else
            If slProductName <> "" Then
                If Len(slISCI) > 0 Then
                    smTZSave(2, 1) = slISCI & " " & slProductName
                    smTZShow(2, 1) = slISCI & " " & slProductName
                Else
                    smTZSave(2, 1) = slProductName
                    smTZShow(2, 1) = slProductName
                End If
            Else
                If Len(slISCI) > 0 Then
                    smTZSave(2, 1) = slISCI ' & " " & slProductName
                    smTZShow(2, 1) = slISCI ' & " " & slProductName
                Else
                    smTZSave(2, 1) = "" '& " " & slISCI ' & " " & slProductName
                    smTZShow(2, 1) = "" '& " " & slISCI ' & " " & slProductName
                End If
            End If
        End If
        Exit Sub
    End If
    If tmSdf.sPtType = "3" Then ' Each Time Zone might have unique Copy
'        ReDim Preserve smTZShow(2, 6) As String    ' Allocate space for 6 Time Zones
'        ReDim Preserve smTZSave(2, 6) As String
        ' Read TZF using lCopyCode from SDF: TZF points to as many as 6 CIF records
        tmTzfSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mTZReadCopyDataErr
        gBtrvErrorMsg ilRet, "mTZReadCopyData (btrGetEqual:TzF, Single)", PostLog
        On Error GoTo 0
        ' Loop through the 6 Time Zone fields and put the data in smTZSave and smTZShow
        For ilLoop = 1 To 6 Step 1
        ' Look for the first positive lZone value
            'imTZSave = 1
            'ilUpperBound = ilLoop
            'smTZShow(1, ilLoop) = smTZSave(1, ilLoop)   'Time zones
            If tmTzf.lCifZone(ilLoop - 1) > 0 Then
                ' Read CIF using lCopyCode from SDF to get Inventory Name and CpfCode
                tmCifSrchKey.lCode = tmTzf.lCifZone(ilLoop - 1)
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                On Error GoTo mTZReadCopyDataErr
                gBtrvErrorMsg ilRet, "mTZReadCopyDataa (btrGetEqual:CIF, Time zone)", PostLog
                On Error GoTo 0
                ' Read CPF to get Product Name and ISCI
                If tmCif.lcpfCode > 0 Then
                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo mTZReadCopyDataErr
                    gBtrvErrorMsg ilRet, "mTZReadCopyDataErr (btrGetEqual:CPF)", PostLog
                    On Error GoTo 0
                Else
                    tmCpf.sName = ""
                    tmCpf.sISCI = ""
                End If
                ' Read MCF to get Media Code, Inv. Name
                If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                    tmMcfSrchKey.iCode = tmCif.iMcfCode
                    ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo mTZReadCopyDataErr
                    gBtrvErrorMsg ilRet, "mTZReadCopyData (btrGetEqual:MCF)", PostLog
                    On Error GoTo 0

                    ' Media Code is tmMcf.sName
                    slCopy = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
                    slISCI = Trim$(tmCpf.sISCI)  ' ISCI Code
                    ' See if sName is Blank
    '                If (Len(Trim$(tmCpf.sName)) <> 0) Then
                        slProductName = Trim$(tmCpf.sName)
    '                End If
                    ' Save COPY/ISCI/PRODUCT
                    'Find Matching index
                    If StrComp(tmTzf.sZone(ilLoop - 1), "Oth", 1) = 0 Then
                        If slProductName <> "" Then
                            If Len(slISCI) > 0 Then
                                smTZSave(2, 2) = slCopy & " " & slISCI & " " & slProductName
                                smTZShow(2, 2) = slCopy & " " & slISCI & " " & slProductName
                            Else
                                smTZSave(2, 2) = slCopy & " " & slProductName
                                smTZShow(2, 2) = slCopy & " " & slProductName
                            End If
                        Else
                            If Len(slISCI) > 0 Then
                                smTZSave(2, 2) = slCopy & " " & slISCI ' & " " & slProductName
                                smTZShow(2, 2) = slCopy & " " & slISCI ' & " " & slProductName
                            Else
                                smTZSave(2, 2) = slCopy '& " " & slISCI ' & " " & slProductName
                                smTZShow(2, 2) = slCopy '& " " & slISCI ' & " " & slProductName
                            End If
                        End If
                    Else
                        For ilIndex = 3 To imNoZones Step 1
                            If StrComp(tmTzf.sZone(ilLoop - 1), smZones(ilIndex), 1) = 0 Then
                                If slProductName <> "" Then
                                    If Len(slISCI) > 0 Then
                                        smTZSave(2, ilIndex) = slCopy & " " & slISCI & " " & slProductName
                                        smTZShow(2, ilIndex) = slCopy & " " & slISCI & " " & slProductName
                                    Else
                                        smTZSave(2, ilIndex) = slCopy & " " & slProductName
                                        smTZShow(2, ilIndex) = slCopy & " " & slProductName
                                    End If
                                Else
                                    If Len(slISCI) > 0 Then
                                        smTZSave(2, ilIndex) = slCopy & " " & slISCI ' & " " & slProductName
                                        smTZShow(2, ilIndex) = slCopy & " " & slISCI ' & " " & slProductName
                                    Else
                                        smTZSave(2, ilIndex) = slCopy '& " " & slISCI ' & " " & slProductName
                                        smTZShow(2, ilIndex) = slCopy '& " " & slISCI ' & " " & slProductName
                                    End If
                                End If
                            End If
                        Next ilIndex
                    End If
                Else
                    slISCI = Trim$(tmCpf.sISCI)  ' ISCI Code
                    slProductName = Trim$(tmCpf.sName)
                    If StrComp(tmTzf.sZone(ilLoop - 1), "Oth", 1) = 0 Then
                        If slProductName <> "" Then
                            If Len(slISCI) > 0 Then
                                smTZSave(2, 2) = slISCI & " " & slProductName
                                smTZShow(2, 2) = slISCI & " " & slProductName
                            Else
                                smTZSave(2, 2) = slProductName
                                smTZShow(2, 2) = slProductName
                            End If
                        Else
                            If Len(slISCI) > 0 Then
                                smTZSave(2, 2) = slISCI ' & " " & slProductName
                                smTZShow(2, 2) = slISCI ' & " " & slProductName
                            Else
                                smTZSave(2, 2) = "" '& " " & slISCI ' & " " & slProductName
                                smTZShow(2, 2) = "" '& " " & slISCI ' & " " & slProductName
                            End If
                        End If
                    Else
                        For ilIndex = 3 To imNoZones Step 1
                            If StrComp(tmTzf.sZone(ilLoop - 1), smZones(ilIndex), 1) = 0 Then
                                If slProductName <> "" Then
                                    If Len(slISCI) > 0 Then
                                        smTZSave(2, ilIndex) = slISCI & " " & slProductName
                                        smTZShow(2, ilIndex) = slISCI & " " & slProductName
                                    Else
                                        smTZSave(2, ilIndex) = slProductName
                                        smTZShow(2, ilIndex) = slProductName
                                    End If
                                Else
                                    If Len(slISCI) > 0 Then
                                        smTZSave(2, ilIndex) = slISCI ' & " " & slProductName
                                        smTZShow(2, ilIndex) = slISCI ' & " " & slProductName
                                    Else
                                        smTZSave(2, ilIndex) = "" '& " " & slISCI ' & " " & slProductName
                                        smTZShow(2, ilIndex) = "" '& " " & slISCI ' & " " & slProductName
                                    End If
                                End If
                            End If
                        Next ilIndex
                    End If
                End If
'                lbcTZCopy.AddItem smTZShow(2, ilUpperBound) ' Put COPY ISCI PRODUCT into listbox
           Else ' No CIF code so set smTZSave/smTZShow to blanks
                'smTZSave(2, ilUpperBound) = ""
                'If smTZSave(1, ilUpperBound) <> "" Then
                '    smTZShow(2, ilUpperBound) = "[None]"
                'Else
                '    smTZShow(2, ilUpperBound) = ""
                'End If
           End If
        Next ilLoop
    End If
mTZReadCopyDataErr:
    On Error GoTo 0
    'mTZReadCopyData = False
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTZSetShow                      *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mTZSetShow(ilBoxNo As Integer)
'
'   mTZSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBTZCtrls) Or (ilBoxNo > UBound(tmTZCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case ZONEINDEX
            pbcTZZone.Visible = False
            'If imTZSave = 0 Then
            '    For ilLoop = 1 To 6 Step 1
            '        If ilLoop = 1 Then
            '            slStr = "[All]"
            '        Else
            '            slStr = ""
            '        End If
            '        gSetShow pbcTZCopy, slStr, tmTZCtrls(ilBoxNo)
            '        smTZShow(1, ilLoop) = tmTZCtrls(ilBoxNo).sShow
            '        slStr = ""
            '        gSetShow pbcTZCopy, slStr, tmTZCtrls(TZCOPYINDEX)
            '        smTZShow(2, ilLoop) = tmTZCtrls(TZCOPYINDEX).sShow
            '    Next ilLoop
            'Else
            '    For ilLoop = 1 To 6 Step 1
            '        slStr = smTZSave(1, ilLoop)
            '        gSetShow pbcTZCopy, slStr, tmTZCtrls(ilBoxNo)
            '        smTZShow(1, ilLoop) = tmTZCtrls(ilBoxNo).sShow
            '        slStr = smTZSave(2, ilLoop)
            '        If (slStr = "") And (smTZSave(1, ilLoop) <> "") Then
            '            slStr = "[None]"
            '        End If
            '        gSetShow pbcTZCopy, slStr, tmTZCtrls(TZCOPYINDEX)
            '        smTZShow(2, ilLoop) = tmTZCtrls(TZCOPYINDEX).sShow
            '    Next ilLoop
            'End If
        Case TZCOPYINDEX
            lbcCopyNm.Visible = False
            edcTZDropDown.Visible = False
            cmcTZDropDown.Visible = False
            slStr = edcTZDropDown.Text
            'ilPos = InStr(slStr, "(Purged)")
            'If ilPos > 0 Then
            '    slStr = Left$(slStr, ilPos - 1)
            'Else
            '    ilPos = InStr(slStr, "(Reused)")
            '    If ilPos > 0 Then
            '        slStr = Left$(slStr, ilPos - 1)
            '    End If
            'End If
            smTZSave(2, imTZRowNo) = slStr
            gSetShow pbcTZCopy, slStr, tmTZCtrls(ilBoxNo)
            smTZShow(2, imTZRowNo) = tmTZCtrls(ilBoxNo).sShow
    End Select
End Sub
'********************************************************************
'*
'*      Procedure Name:mUpdateAffPost
'*          <input> ilDayIndex - day of week (0-6) to
'*                  clear or set
'*
'*             Created:7/19/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Update day is complete flag
'*
'*          8-24-04 Set/clear day is complete flag based on day of
'                   week index
'********************************************************************
Private Sub mUpdateAffPost(ilDaySelected As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDay                                                                                 *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slDate As String
    Dim ilDayIndex As Integer
    Dim llLcfDate As Long
    
    If imDateSelectedIndex < 0 Then
        Exit Sub
    End If
    If rbcType(1).Value Then
        Exit Sub
    End If
    If imWM = 1 Then
        Exit Sub
    End If
    'slDate = smSelectedDate 'cbcDate.List(imDateSelectedIndex)
    'ilRet = gParseItem(slDate, 2, " ", slDate) 'Remove day
    'ilRet = gParseItem(slDate, 1, ":", slDate) 'Remove :xxxxxxx
    'If InStr(slDate, ":C") > 0 Then
    '    ilRet = gParseItem(slDate, 1, ":C", slDate)
    'Else
    '    ilRet = gParseItem(slDate, 1, ":I", slDate)
    'End If
    If tmVef.sType <> "G" Then
        '9/7/06-Use ilDaySelected as it is 0-6
        slDate = tgDates(imDateSelectedIndex).sDate
        ilDayIndex = gWeekDayStr(slDate)
        If ckcDayComplete(ilDaySelected).Value = vbChecked Then
            '9/7/06-Use ilDaySelected as it is 0-6
            'slDate = tgDates(imDateSelectedIndex).sDate
            ''ilDay = gWeekDayStr(slDate)
            Do
                '9/7/06-Use ilDaySelected as it is 0-6
                'ilRet = btrGetDirect(hmLcf, tmLcf, imLcfRecLen, lmLcfRecPos(ilDayIndex), INDEXKEY0, BTRV_LOCK_NONE)
                ilRet = btrGetDirect(hmLcf, tmLcf, imLcfRecLen, lmLcfRecPos(ilDaySelected), INDEXKEY0, BTRV_LOCK_NONE)
               'tmRec = tmLcf
                'ilRet = gGetByKeyForUpdate("LCF", hmLcf, tmRec)
                'tmLcf = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = MsgBox("Complete Not Set, Try Later", vbOkOnly + vbExclamation, "Post Log")
                '    Exit Sub
                'End If
                If ilRet <> BTRV_ERR_NONE Then      'no day exists
                    Exit Sub
                End If
                If tmLcf.sAffPost = "C" Then
                    Exit Sub
                End If
                tmLcf.sAffPost = "C"  ' Complete
                ' Update the file
                ilRet = btrUpdate(hmLcf, tmLcf, imLcfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If (ilDayIndex <= ilDaySelected) Then
                If imDateSelectedIndex + ilDaySelected - ilDayIndex < UBound(tgDates) Then
                    tgDates(imDateSelectedIndex + ilDaySelected - ilDayIndex).iStatus = 2
                End If
            End If
        Else
            'For ilDay = 0 To 6 Step 1
            '    If imSdfAnyChg(ilDay) And (ilDay >= ilDayIndex) Then
            '        Do
            '            ilRet = btrGetDirect(hmLcf, tmLcf, imLcfRecLen, lmLcfRecPos(ilDay), INDEXKEY0, BTRV_LOCK_NONE)
            '            'tmRec = tmLcf
            '            'ilRet = gGetByKeyForUpdate("LCF", hmLcf, tmRec)
            '            'tmLcf = tmRec
            '            'If ilRet <> BTRV_ERR_NONE Then
            '            '    ilRet = MsgBox("Complete Not Set, Try Later", vbOkOnly + vbExclamation, "Post Log")
            '            '    Exit Sub
            '            'End If
            '
            '            If ilRet <> BTRV_ERR_NONE Then
            '                Exit Sub
            '            End If
            '
            '            If tmLcf.sAffPost = "I" Then
            '                Exit Do
            '            End If
            '            tmLcf.sAffPost = "I"  ' Incomplete
            '            ' Update the file
            '            ilRet = btrUpdate(hmLcf, tmLcf, imLcfRecLen)
            '        Loop While ilRet = BTRV_ERR_CONFLICT
            '        tgDates(imDateSelectedIndex + ilDay - ilDayIndex).iStatus = 1
            '    End If
            'Next ilDay
            Do
                '9/7/06-Use ilDaySelected as it is 0-6
                'ilRet = btrGetDirect(hmLcf, tmLcf, imLcfRecLen, lmLcfRecPos(ilDayIndex), INDEXKEY0, BTRV_LOCK_NONE)
                ilRet = btrGetDirect(hmLcf, tmLcf, imLcfRecLen, lmLcfRecPos(ilDaySelected), INDEXKEY0, BTRV_LOCK_NONE)
               'tmRec = tmLcf
                'ilRet = gGetByKeyForUpdate("LCF", hmLcf, tmRec)
                'tmLcf = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = MsgBox("Complete Not Set, Try Later", vbOkOnly + vbExclamation, "Post Log")
                '    Exit Sub
                'End If
                If ilRet <> BTRV_ERR_NONE Then      'no day exists
                    Exit Sub
                End If
                If tmLcf.sAffPost = "I" Then
                    Exit Sub
                End If
                tmLcf.sAffPost = "I"  ' InComplete
                ' Update the file
                ilRet = btrUpdate(hmLcf, tmLcf, imLcfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If (ilDayIndex <= ilDaySelected) Then
                If imDateSelectedIndex + ilDaySelected - ilDayIndex < UBound(tgDates) Then
                    tgDates(imDateSelectedIndex + ilDaySelected - ilDayIndex).iStatus = 1
                End If
            End If
        End If
    Else
        Do
            'tmLcfSrchKey1.iVefCode = imVefCode
            'tmLcfSrchKey1.iType = tgDates(imDateSelectedIndex).iGameNo
            'ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            'If ilRet = BTRV_ERR_NONE Then
            slDate = tgDates(imDateSelectedIndex).sDate
            tmLcfSrchKey2.iVefCode = imVefCode
            gPackDate slDate, tmLcfSrchKey2.iLogDate(0), tmLcfSrchKey2.iLogDate(1)
            ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode)
                gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llLcfDate
                If llLcfDate <> gDateValue(slDate) Then
                    Exit Do
                End If
                If tmLcf.iType = tgDates(imDateSelectedIndex).iGameNo Then
                    If ckcDayComplete(0).Value = vbChecked Then
                        If tmLcf.sAffPost = "C" Then
                            Exit Sub
                        End If
                        tmLcf.sAffPost = "C"  ' Complete
                        ' Update the file
                        ilRet = btrUpdate(hmLcf, tmLcf, imLcfRecLen)
                        tgDates(imDateSelectedIndex).iStatus = 2
                    Else
                        If imSdfAnyChg(0) Then
                            If tmLcf.sAffPost = "I" Then
                                Exit Sub
                            End If
                            tmLcf.sAffPost = "I"  ' Incomplete
                            ' Update the file
                            ilRet = btrUpdate(hmLcf, tmLcf, imLcfRecLen)
                            tgDates(imDateSelectedIndex).iStatus = 1
                        End If
                    End If
                    Exit Do
                Else
                    ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
            Loop
            'Else
            '    Exit Do
            'End If
        Loop While ilRet = BTRV_ERR_CONFLICT

    End If
    'imDateChgMode = True
    'slDate = gAddDayToDate(slDate)
    'If ckcDayComplete.Value Then
    '    'cbcDate.List(imDateSelectedIndex) = slDate & ":Completed"
    '    tgDates(imDateSelectedIndex).iStatus = 2
    'Else
    '    'cbcDate.List(imDateSelectedIndex) = slDate & ":Incomplete"
    '    tgDates(imDateSelectedIndex).iStatus = 1
    'End If
    'imDateChgMode = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mUpdateMdShow                   *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Update tgMdShowInfo            *
'*                                                     *
'*******************************************************
Private Sub mUpdateMdShow()
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slNoSpots As String
    Dim llWkDate As Long
    Dim ilDay As Integer
    Dim ilUpper As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim slStr As String
    Dim llPrice As Long
    ReDim ilDays(0 To 6) As Integer

    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
    ilFound = False
    llWkDate = gDateValue(slDate)
    llDate = llWkDate
    Do While gWeekDayLong(llWkDate) <> 0
        llWkDate = llWkDate - 1
    Loop
    ilRet = mReadChfClfRdfRec(tmSdf.lChfCode, tmSdf.iLineNo, tmSdf.lFsfCode)
    If ilRet Then
        'For ilLoop = LBound(tgMdSaveInfo) To UBound(tgMdSaveInfo) - 1 Step 1
        For ilLoop = LBONE To UBound(tgMdSaveInfo) - 1 Step 1
            If (tgMdSaveInfo(ilLoop).lChfCode = tmSdf.lChfCode) And (tgMdSaveInfo(ilLoop).iVefCode = tmSdf.iVefCode) And (tgMdSaveInfo(ilLoop).iLen = tmSdf.iLen) Then
                If (tgMdSaveInfo(ilLoop).iRdfCode = tmClf.iRdfCode) And (tgMdSaveInfo(ilLoop).iStartTime(0) = tmClf.iStartTime(0)) And (tgMdSaveInfo(ilLoop).iStartTime(1) = tmClf.iStartTime(1)) And (tgMdSaveInfo(ilLoop).iEndTime(0) = tmClf.iEndTime(0)) And (tgMdSaveInfo(ilLoop).iEndTime(1) = tmClf.iEndTime(1)) Then
                    If (tgMdSaveInfo(ilLoop).lWkMissed = llWkDate) And (tgMdSaveInfo(ilLoop).sBill = tmSdf.sBill) Then
                        ilRet = mReadCffRec(slStartDate, slEndDate, slNoSpots, llPrice, ilDays())
                        If ilRet Then
                            ilFound = True
                            For ilDay = 0 To 6 Step 1
                                If tgMdSaveInfo(ilLoop).iDay(ilDay) <> ilDays(ilDay) Then
                                    ilFound = False
                                    Exit For
                                End If
                            Next ilDay
                            If ilFound Then
                                tgMdSdfRec(UBound(tgMdSdfRec)).lSdfCode = tmSdf.lCode
                                tgMdSdfRec(UBound(tgMdSdfRec)).lPrice = llPrice
                                ilRet = btrGetPosition(hmSdf, tgMdSdfRec(UBound(tgMdSdfRec)).lSdfRecPos)
                                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), tgMdSdfRec(UBound(tgMdSdfRec)).lMissedDate
                                tgMdSdfRec(UBound(tgMdSdfRec)).sSchStatus = tmSdf.sSchStatus
                                If tmSdf.sSchStatus = "C" Then
                                    tgMdSaveInfo(ilLoop).iCancelCount = tgMdSaveInfo(ilLoop).iCancelCount + 1
                                ElseIf tmSdf.sSchStatus = "H" Then
                                    tgMdSaveInfo(ilLoop).iHiddenCount = tgMdSaveInfo(ilLoop).iHiddenCount + 1
                                ElseIf tmSdf.sSchStatus = "M" Then
                                    tgMdSaveInfo(ilLoop).iMissedCount = tgMdSaveInfo(ilLoop).iMissedCount + 1
                                End If
                                tgMdSdfRec(UBound(tgMdSdfRec)).iNextIndex = tgMdSaveInfo(ilLoop).iFirstIndex
                                tgMdSaveInfo(ilLoop).iFirstIndex = UBound(tgMdSdfRec)
                                ReDim Preserve tgMdSdfRec(0 To UBound(tgMdSdfRec) + 1) As MDSDFREC
                                ilUpper = ilLoop
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next ilLoop
        If Not ilFound Then
            ilUpper = UBound(tgMdSaveInfo)
            'For ilLoop = LBound(tgMdSaveInfo) To UBound(tgMdSaveInfo) - 1 Step 1
            For ilLoop = LBONE To UBound(tgMdSaveInfo) - 1 Step 1
                If (tgMdSaveInfo(ilLoop).lChfCode = tmSdf.lChfCode) And (tgMdSaveInfo(ilLoop).iVefCode = tmSdf.iVefCode) And (tgMdSaveInfo(ilLoop).iLen = tmSdf.iLen) Then
                    If (tgMdSaveInfo(ilLoop).iRdfCode = tmClf.iRdfCode) And (tgMdSaveInfo(ilLoop).iStartTime(0) = tmClf.iStartTime(0)) And (tgMdSaveInfo(ilLoop).iStartTime(1) = tmClf.iStartTime(1)) And (tgMdSaveInfo(ilLoop).iEndTime(0) = tmClf.iEndTime(0)) And (tgMdSaveInfo(ilLoop).iEndTime(1) = tmClf.iEndTime(1)) Then
                        If (tgMdSaveInfo(ilLoop).iLineNo = tmSdf.iLineNo) And (tgMdSaveInfo(ilLoop).lWkMissed = 0) Then
                            'ilRet = mReadCffRec(slStartDate, slEndDate, slNoSpots, llPrice, ilDays())
                            ilFound = True
                            ilUpper = ilLoop
                            llWkDate = 0
                            Exit For
                        End If
                    End If
                End If
            Next ilLoop
            If llWkDate = 0 Then
                For ilDay = 0 To 6 Step 1
                    ilDays(ilDay) = 0
                Next ilDay
            Else
                ilRet = mReadCffRec(slStartDate, slEndDate, slNoSpots, llPrice, ilDays())
            End If
            tgMdSaveInfo(ilUpper).lChfCode = tmSdf.lChfCode
            tgMdSaveInfo(ilUpper).iAdfCode = tmSdf.iAdfCode
            tgMdSaveInfo(ilUpper).lFsfCode = 0
            tgMdSaveInfo(ilUpper).lCntrNo = tmChf.lCntrNo
            gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), tgMdSaveInfo(ilUpper).lEndDate
            tgMdSaveInfo(ilUpper).iVefCode = tmSdf.iVefCode
            tgMdSaveInfo(ilUpper).iLineNo = tmClf.iLine
            tgMdSaveInfo(ilUpper).iLen = tmSdf.iLen
            tgMdSaveInfo(ilUpper).lWkMissed = llWkDate
            tgMdSaveInfo(ilUpper).iRdfCode = tmClf.iRdfCode
            tgMdSaveInfo(ilUpper).iStartTime(0) = tmClf.iStartTime(0)
            tgMdSaveInfo(ilUpper).iStartTime(1) = tmClf.iStartTime(1)
            tgMdSaveInfo(ilUpper).iEndTime(0) = tmClf.iEndTime(0)
            tgMdSaveInfo(ilUpper).iEndTime(1) = tmClf.iEndTime(1)
            For ilDay = 0 To 6 Step 1
                tgMdSaveInfo(ilUpper).iDay(ilDay) = ilDays(ilDay)
            Next ilDay
            If tmSdf.sSchStatus = "C" Then
                tgMdSaveInfo(ilUpper).iCancelCount = 1
            ElseIf tmSdf.sSchStatus = "H" Then
                tgMdSaveInfo(ilUpper).iHiddenCount = 1
            ElseIf tmSdf.sSchStatus = "M" Then
                tgMdSaveInfo(ilUpper).iMissedCount = 1
            End If
            tgMdSaveInfo(ilUpper).sBill = tmSdf.sBill
            tgMdSaveInfo(ilUpper).iFirstIndex = UBound(tgMdSdfRec)
            If Not ilFound Then
                ReDim Preserve tgMdSaveInfo(0 To ilUpper + 1) As MDSAVEINFO
            End If
            tgMdSdfRec(UBound(tgMdSdfRec)).lPrice = llPrice
            tgMdSdfRec(UBound(tgMdSdfRec)).iNextIndex = -1
            tgMdSdfRec(UBound(tgMdSdfRec)).lSdfCode = tmSdf.lCode
            ilRet = btrGetPosition(hmSdf, tgMdSdfRec(UBound(tgMdSdfRec)).lSdfRecPos)
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), tgMdSdfRec(UBound(tgMdSdfRec)).lMissedDate
            tgMdSdfRec(UBound(tgMdSdfRec)).sSchStatus = tmSdf.sSchStatus
            ReDim Preserve tgMdSdfRec(0 To UBound(tgMdSdfRec) + 1) As MDSDFREC
        End If
        ilFound = False
        'For ilLoop = LBound(tgMdShowInfo) To UBound(tgMdShowInfo) - 1 Step 1
        For ilLoop = LBONE To UBound(tgMdShowInfo) - 1 Step 1
            If (tgMdShowInfo(ilLoop).iMdSaveInfoIndex = ilUpper) Then
                If ((tmSdf.sSchStatus = "M") And ((tgMdShowInfo(ilLoop).iType = 1) Or (tgMdShowInfo(ilLoop).iType = 0))) Or ((tmSdf.sSchStatus = "H") And ((tgMdShowInfo(ilLoop).iType = 2) Or (tgMdShowInfo(ilLoop).iType = 0))) Or ((tmSdf.sSchStatus = "C") And ((tgMdShowInfo(ilLoop).iType = 3) Or (tgMdShowInfo(ilLoop).iType = 0))) Then
                    ilFound = True
                    If (tmSdf.sSchStatus = "M") Then
                        tgMdShowInfo(ilLoop).iType = 1
                        If tgMdSaveInfo(ilUpper).iMissedCount > 0 Then
                            slStr = Trim$(str$(tgMdSaveInfo(ilUpper).iMissedCount))
                        Else
                            slStr = ""
                        End If
                        gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
                        tgMdShowInfo(ilLoop).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
                    ElseIf tmSdf.sSchStatus = "H" Then
                        tgMdShowInfo(ilLoop).iType = 2
                        If tgMdSaveInfo(ilUpper).iHiddenCount > 0 Then
                            slStr = Trim$(str$(tgMdSaveInfo(ilUpper).iHiddenCount))
                        Else
                            slStr = ""
                        End If
                        gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
                        tgMdShowInfo(ilLoop).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
                    ElseIf tmSdf.sSchStatus = "C" Then
                        tgMdShowInfo(ilLoop).iType = 3
                        If tgMdSaveInfo(ilUpper).iCancelCount > 0 Then
                            slStr = Trim$(str$(tgMdSaveInfo(ilUpper).iCancelCount))
                        Else
                            slStr = ""
                        End If
                        gSetShow pbcMissed, slStr, tmMdCtrls(MDNOSPOTSINDEX)
                        tgMdShowInfo(ilLoop).sShow(MDNOSPOTSINDEX) = tmMdCtrls(MDNOSPOTSINDEX).sShow
                    End If
                    'Week Missed
                    If tgMdSaveInfo(ilUpper).lWkMissed > 0 Then
                        slStr = Format$(tgMdSaveInfo(ilUpper).lWkMissed, "m/d/yy")
                    Else
                        slStr = " "
                    End If
                    gSetShow pbcMissed, slStr, tmMdCtrls(MDWKMISSINDEX)
                    tgMdShowInfo(ilLoop).sShow(MDWKMISSINDEX) = tmMdCtrls(MDWKMISSINDEX).sShow
                    'pbcMissed.Cls
                    'DoEvents
                    'pbcMissed_Paint
                    'Force paint because of multi-spots at one time
                    ilFound = False
                    Exit For
                End If
            End If
        Next ilLoop
        If Not ilFound Then
            pbcMissed.Cls
            DoEvents
            mCreateMdShow True
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim llFilter As Long
    'ilRet = gPopUserVehicleBox(PostLog, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcVeh, Traffic!lbcUserVehicle)
    If rbcType(0).Value Then
        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + VEHEXCLUDEPODNOPRGM
    Else
        'Need to include conventional vehicle as dynamic packages can be made referencing conventional
        llFilter = VEHPACKAGE + VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH + VEHEXCLUDEPODNOPRGM
    End If
    ilRet = gPopUserVehicleBox(PostLog, llFilter, cbcVeh, tmUserVehicle(), smUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", PostLog
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcArrow_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub pbcArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                If imDTBoxNo = DTDATEINDEX Then
                    edcDTDropDown.Text = Format$(llDate, "m/d/yy")
                    edcDTDropDown.SelStart = 0
                    edcDTDropDown.SelLength = Len(edcDTDropDown.Text)
                    imBypassFocus = True
                    DoEvents
                    edcDTDropDown.SetFocus
                    Exit Sub
                Else
                    If imDateBox = 0 Then
                        edcDate.Text = Format$(llDate, "m/d/yy")
                        edcDate.SelStart = 0
                        edcDate.SelLength = Len(edcDate.Text)
                        imBypassFocus = True
                        DoEvents
                        edcDate.SetFocus
                        Exit Sub
                    Else
                        edcMdDate(imDateBox - 1).Text = Format$(llDate, "m/d/yy")
                        edcMdDate(imDateBox - 1).SelStart = 0
                        edcMdDate(imDateBox - 1).SelLength = Len(edcDate.Text)
                        imBypassFocus = True
                        edcMdDate(imDateBox - 1).SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imDTBoxNo = DTDATEINDEX Then
        edcDTDropDown.SetFocus
    Else
        If imDateBox = 0 Then
            edcDate.SetFocus
        ElseIf (imDateBox = 1) Or (imDateBox = 2) Then
            edcMdDate(imDateBox - 1).SetFocus
        End If
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    If (imDTBoxNo = DTDATEINDEX) And (rbcType(1).Value) Then
        slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
        lacCalName.Caption = gMonthYearFormat(slStr)
        gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
        mBoxCalDate
    ElseIf (imDTBoxNo = DTDATEINDEX) Then
        slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
        lacCalName.Caption = gMonthYearFormat(slStr)
        gPLPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate, tgWkDates()  'tgDates()
        mBoxCalDate
    Else
        slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
        lacCalName.Caption = gMonthYearFormat(slStr)
        gPLPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate, tgDates()
        mBoxCalDate
    End If
End Sub
Private Sub pbcClickFocus_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcDT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBDTCtrls To UBound(tmDTCtrls) Step 1
        If (X >= tmDTCtrls(ilBox).fBoxX) And (X <= (tmDTCtrls(ilBox).fBoxX + tmDTCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmDTCtrls(ilBox).fBoxY)) And (Y <= (tmDTCtrls(ilBox).fBoxY + tmDTCtrls(ilBox).fBoxH)) Then
                mDTSetShow imDTBoxNo
                imDTBoxNo = ilBox
                mDTEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcDT_Paint()
    Dim ilBox As Integer
    '6/16/11
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    
    '6/16/11
    llColor = pbcDT.ForeColor
    If (rbcType(0).Value) And (smSpotType <> "O") And (smSpotType <> "C") Then
        pbcDT.CurrentX = tmDTCtrls(imDTAVAILTIMEINDEX).fBoxX + 15
        pbcDT.CurrentY = tmDTCtrls(imDTAVAILTIMEINDEX).fBoxY
    Else
        pbcDT.CurrentX = tmDTCtrls(imDTAIRTIMEINDEX).fBoxX + 15
        pbcDT.CurrentY = tmDTCtrls(imDTAIRTIMEINDEX).fBoxY
    End If
    slFontName = pbcDT.FontName
    flFontSize = pbcDT.FontSize
    pbcDT.ForeColor = BLUE
    pbcDT.FontBold = False
    pbcDT.FontSize = 7
    pbcDT.FontName = "Arial"
    pbcDT.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    If (rbcType(0).Value) And (smSpotType <> "O") And (smSpotType <> "C") Then
        pbcDT.Print "Avail Time"
    Else
        pbcDT.Print "Air Time"
    End If
    pbcDT.FontSize = flFontSize
    pbcDT.FontName = slFontName
    pbcDT.FontSize = flFontSize
    pbcDT.ForeColor = llColor
    pbcDT.FontBold = True
    
    For ilBox = imLBDTCtrls To UBound(tmDTCtrls) Step 1
        pbcDT.CurrentX = tmDTCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDT.CurrentY = tmDTCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcDT.Print tmDTCtrls(ilBox).sShow ' Print one box of data at a time
    Next ilBox
End Sub
Private Sub pbcDTSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim llDate As Long
    If GetFocus() <> pbcDTSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    ilBox = imDTBoxNo
    Select Case ilBox
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If imBoxNo = DATEINDEX Then
                imDTBoxNo = DTDATEINDEX
            Else
                imDTBoxNo = imDTAIRTIMEINDEX
            End If
            mDTEnableBox imDTBoxNo
            Exit Sub
        Case DTDATEINDEX
            slStr = edcDTDropDown.Text
            If Not gValidDate(slStr) Then
                Beep
                edcDTDropDown.SetFocus
                Exit Sub
            End If
            If rbcType(0).Value Then
                llDate = gDateValue(slStr)
                ilFound = False
                For ilLoop = 0 To UBound(tgWkDates) - 1 Step 1
                    If tgWkDates(ilLoop).lDate = llDate Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    edcDTDropDown.SetFocus
                    Exit Sub
                End If
            End If
            mDTSetShow imDTBoxNo
            imDTBoxNo = -1
            pbcSTab.SetFocus
            Exit Sub
        Case imDTAIRTIMEINDEX
            slStr = Trim$(edcDTDropDown.Text)
            If (Not gValidTime(slStr)) Or (slStr = "") Then
                Beep
                edcDTDropDown.SetFocus
                Exit Sub
            End If
            ilBox = ilBox - 1
        Case imDTAVAILTIMEINDEX
            slStr = Trim$(edcDTDropDown.Text)
            If (Not gValidTime(slStr)) Or (slStr = "") Then
                Beep
                'edcDTDropDown.SetFocus
                Exit Sub
            End If
            ilBox = ilBox - 1
        Case Else
            ilBox = ilBox - 1
    End Select
    mDTSetShow imDTBoxNo
    imDTBoxNo = ilBox
    mDTEnableBox ilBox
End Sub
Private Sub pbcDTTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim ilFound As Integer
    If GetFocus() <> pbcDTTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    ilBox = imDTBoxNo
    Select Case ilBox
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            imDTBoxNo = imDTAIRTIMEINDEX
            mDTEnableBox imDTBoxNo
            Exit Sub
        Case DTDATEINDEX
            slStr = edcDTDropDown.Text
            If Not gValidDate(slStr) Then
                Beep
                edcDTDropDown.SetFocus
                Exit Sub
            End If
            If rbcType(0).Value Then
                llDate = gDateValue(slStr)
                ilFound = False
                For ilLoop = 0 To UBound(tgWkDates) - 1 Step 1
                    If tgWkDates(ilLoop).lDate = llDate Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    edcDTDropDown.SetFocus
                    Exit Sub
                End If
            End If
            ilBox = ilBox + 1
        Case imDTAIRTIMEINDEX
            slStr = Trim$(edcDTDropDown.Text)
            If (Not gValidTime(slStr)) Or (slStr = "") Then
                Beep
                edcDTDropDown.SetFocus
                Exit Sub
            End If
            mDTSetShow imDTBoxNo
            imDTBoxNo = -1
            pbcTab.SetFocus
            Exit Sub
        Case imDTAVAILTIMEINDEX
            slStr = Trim$(edcDTDropDown.Text)
            If (Not gValidTime(slStr)) Or (slStr = "") Then
                Beep
                'edcDTDropDown.SetFocus
                Exit Sub
            End If
            ilBox = ilBox + 1
        Case Else
            ilBox = ilBox + 1
    End Select
    mDTSetShow imDTBoxNo
    imDTBoxNo = ilBox
    mDTEnableBox ilBox
End Sub
Private Sub pbcMissed_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub pbcMissed_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    ' Turn off timer
    tmcDrag.Enabled = False
End Sub
Private Sub pbcMissed_GotFocus()
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        'Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    'ilRet = mReadSdfRec(True) ' Get Remaining data
End Sub
Private Sub pbcMissed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRet As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ilRet = mReadSdfRec(True) ' Get Remaining data
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragSource = 1
    tmcDrag.Enabled = True  'Start timer to see if drag or clickEnd Sub
End Sub
Private Sub pbcMissed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If Not imUpdateAllowed Then
        Exit Sub
    End If
End Sub
Private Sub pbcMissed_Paint()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slFont                                                                                *
'******************************************************************************************

' paint the Missed Form from smShow
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    mPaintMissedTitle
    ilStartRow = vbcMissed.Value   'Top location
    ilEndRow = vbcMissed.Value + vbcMissed.LargeChange
    If ilEndRow >= UBound(tgMdShowInfo) Then
        ilEndRow = UBound(tgMdShowInfo) - 1
    End If
    If ilEndRow = 0 Then  ' There are no rows to display
       Exit Sub
    End If
    llColor = pbcMissed.ForeColor
    slFontName = pbcPosting.FontName
    flFontSize = pbcPosting.FontSize
    'pbcMissed.ForeColor = BLUE
    'pbcMissed.FontBold = False
    'pbcMissed.FontSize = 7
    'pbcMissed.FontName = "Arial"
    'pbcMissed.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    'pbcMissed.CurrentX = tmMdCtrls(MDVEHINDEX).fBoxX + fgBoxInsetX
    'pbcMissed.CurrentY = 15 '- 30 '+ fgBoxInsetY
    'If cbcVehicle.ListIndex > 0 Then
    '    pbcMissed.Print "Vehicle"
    'Else
    '    pbcMissed.Print "Product"
    'End If
    'pbcMissed.FontSize = flFontSize
    'pbcMissed.FontName = slFontName
    'pbcMissed.FontSize = flFontSize
    'pbcMissed.ForeColor = llColor
    'pbcMissed.FontBold = True
   '
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBMdCtrls To UBound(tmMdCtrls) Step 1
            pbcMissed.CurrentX = tmMdCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcMissed.CurrentY = tmMdCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            If Trim$(tgMdShowInfo(ilRow).sShow(MDCNTRINDEX)) = "Feed" Then
                pbcMissed.ForeColor = DARKGRAY
            End If
            If tgMdShowInfo(ilRow).iType = 3 Then 'cancel
                pbcMissed.ForeColor = RED
            ElseIf tgMdShowInfo(ilRow).iType = 2 Then 'hidden
                pbcMissed.ForeColor = CYAN
            End If
            If (ilBox = imLBMdCtrls) And (tgMdShowInfo(ilRow).sBill = "Y") Then
                pbcMissed.ForeColor = DARKGREEN
            End If
            If (ilBox = MDCNTRINDEX) And (Trim$(tgMdShowInfo(ilRow).sShow(MDCNTRINDEX)) = "Feed") Then
                pbcMissed.Print ""
            ElseIf (ilBox = MDVEHINDEX) Then
                If cbcVehicle.ListIndex > 0 Then
                    pbcMissed.Print tgMdShowInfo(ilRow).sShow(ilBox)
                Else
                    pbcMissed.Print tgMdShowInfo(ilRow).sShow(MDPRODINDEX)
                End If
            Else
                pbcMissed.Print tgMdShowInfo(ilRow).sShow(ilBox)
            End If
            pbcMissed.ForeColor = llColor
        Next ilBox
    Next ilRow
End Sub

Private Sub pbcMissedType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcMissedType_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("M")) Or (KeyAscii = Asc("m")) Then
        imMissedType = 0
        pbcMissedType_Paint
        Exit Sub
    End If
    If (KeyAscii = Asc("C")) Or (KeyAscii = Asc("c")) Then
        imMissedType = 1
        pbcMissedType_Paint
        Exit Sub
    End If
    If (KeyAscii = Asc("H")) Or (KeyAscii = Asc("h")) Then
        imMissedType = 2
        pbcMissedType_Paint
        Exit Sub
    End If
    If KeyAscii = Asc(" ") Then
        If imMissedType = 0 Then 'imSave(1, imRowNo) = 0 Then  'True price
            imMissedType = 1
            pbcMissedType_Paint
        ElseIf imMissedType = 1 Then 'imSave(1, imRowNo) = 1 Then  'N/C
            imMissedType = 2
            pbcMissedType_Paint
        ElseIf imMissedType = 2 Then 'imSave(1, imRowNo) = 1 Then  'N/C
            imMissedType = 0
            pbcMissedType_Paint
        End If
    End If
    mMissedPop
End Sub

Private Sub pbcMissedType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imMissedType = 0 Then 'imSave(1, imRowNo) = 0 Then  'True price
        imMissedType = 1
        pbcMissedType_Paint
    ElseIf imMissedType = 1 Then 'imSave(1, imRowNo) = 1 Then  'N/C
        imMissedType = 2
        pbcMissedType_Paint
    ElseIf imMissedType = 2 Then 'imSave(1, imRowNo) = 1 Then  'N/C
        imMissedType = 0
        pbcMissedType_Paint
    End If
    mMissedPop
End Sub

Private Sub pbcMissedType_Paint()
    pbcMissedType.Cls
    pbcMissedType.CurrentX = fgBoxInsetX
    pbcMissedType.CurrentY = 0 'fgBoxInsetY
    If imMissedType = 0 Then
        pbcMissedType.Print "Missed"    'smSave(SAVPRICEINDEX, imRowNo)
    ElseIf imMissedType = 1 Then
        pbcMissedType.Print "Cancel"
    ElseIf imMissedType = 2 Then
        pbcMissedType.Print "Hidden"
    Else
        pbcMissedType.Print ""
    End If
End Sub

Private Sub pbcPosting_DragDrop(Source As control, X As Single, Y As Single)
    Dim ilRet As Integer
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer

    imcTrash.Visible = False
    imcHidden.Visible = False
    If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "R") Or (smDragCntrType = "Q") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
        Exit Sub
    End If
    If imDragSource = 0 Then
'        lacPtFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    ElseIf imDragSource = 1 Then
        'Insert spot
        lacMdFrame.Visible = False
        ilCompRow = vbcPosting.LargeChange + 1
        'If UBound(smSave, 2) > ilCompRow Then
        If UBound(tgShow) > ilCompRow Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(tgShow) - 1   'UBound(smSave, 2) - 1
        End If
        For ilRow = 1 To ilMaxRow Step 1
            If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(TIMEINDEX).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(TIMEINDEX).fBoxY + tmCtrls(TIMEINDEX).fBoxH)) Then
'                lacPtFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                imDropRowNo = ilRow + vbcPosting.Value - 1
                imBoxNo = MISSEDTOSCH
                imSdfChg = True
                ilRet = mSaveRec()
                If Not ilRet Then
                    imBoxNo = -1
                    imRowNo = -1
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
                imSettingValue = True
                'If UBound(smSave, 2) <= vbcPosting.LargeChange + 1 Then
                If UBound(tgShow) <= vbcPosting.LargeChange + 1 Then
                ' If this is used, there are probably 0 or 1 records
                    vbcPosting.Max = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
                Else
                ' Saves, what amounts to, the count of records just retrieved
                    'vbcPosting.Max = UBound(smSave, 2) - vbcPosting.LargeChange
                    vbcPosting.Max = UBound(tgShow) - vbcPosting.LargeChange
                End If
                ''If imDropRowNo + 1 < UBound(smSave, 2) Then
                'If imDropRowNo + 1 < UBound(tgShow) Then
                '    imRowNo = imDropRowNo + 1
                '    imBoxNo = 1
                '    mEnableBox imBoxNo
                'Else
                    imBoxNo = -1
                    imRowNo = -1
                    pbcClickFocus.SetFocus
                'End If
                Exit Sub
            End If
        Next ilRow
        imDropRowNo = UBound(tgShow)    'UBound(smSave, 2)
        imBoxNo = MISSEDTOSCH
        imSdfChg = True
        ilRet = mSaveRec()
        If Not ilRet Then
            imBoxNo = -1
            imRowNo = -1
            pbcClickFocus.SetFocus
            Exit Sub
        End If
        If UBound(tgShow) <= vbcPosting.LargeChange + 1 Then
        ' If this is used, there are probably 0 or 1 records
            vbcPosting.Max = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
        Else
        ' Saves, what amounts to, the count of records just retrieved
            'vbcPosting.Max = UBound(smSave, 2) - vbcPosting.LargeChange
            vbcPosting.Max = UBound(tgShow) - vbcPosting.LargeChange
        End If
        'If imDropRowNo + 1 < UBound(tgShow) Then    'UBound(smSave, 2) Then
        '    imRowNo = imDropRowNo + 1
        '    imBoxNo = 1
        '    mEnableBox imBoxNo
        'Else
            imBoxNo = -1
            imRowNo = -1
            pbcClickFocus.SetFocus
        'End If
        Exit Sub
'        lacMdFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    ElseIf imDragSource = 2 Then
        'Insert spot
        lacMdFrame.Visible = False
        ilCompRow = vbcPosting.LargeChange + 1
        'If UBound(smSave, 2) > ilCompRow Then
        If UBound(tgShow) > ilCompRow Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(tgShow) - 1   'UBound(smSave, 2) - 1
        End If
        For ilRow = 1 To ilMaxRow Step 1
            If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(TIMEINDEX).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(TIMEINDEX).fBoxY + tmCtrls(TIMEINDEX).fBoxH)) Then
'                lacPtFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                imDropRowNo = ilRow + vbcPosting.Value - 1
                imBoxNo = BONUSTOSCH
                imSdfChg = True
                ilRet = mSaveRec()
                If Not ilRet Then
                    imBoxNo = -1
                    imRowNo = -1
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
                imSettingValue = True
                'If UBound(smSave, 2) <= vbcPosting.LargeChange + 1 Then
                If UBound(tgShow) <= vbcPosting.LargeChange + 1 Then
                ' If this is used, there are probably 0 or 1 records
                    vbcPosting.Max = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
                Else
                ' Saves, what amounts to, the count of records just retrieved
                    'vbcPosting.Max = UBound(smSave, 2) - vbcPosting.LargeChange
                    vbcPosting.Max = UBound(tgShow) - vbcPosting.LargeChange
                End If
                If imDropRowNo + 1 < UBound(tgShow) Then    'UBound(smSave, 2) Then
                    imRowNo = imDropRowNo + 1
                    imBoxNo = COPYINDEX '1
                    mEnableBox imBoxNo
                Else
                    imBoxNo = -1
                    imRowNo = -1
                    pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
        Next ilRow
        imDropRowNo = UBound(tgShow)    'UBound(smSave, 2)
        imBoxNo = BONUSTOSCH
        imSdfChg = True
        ilRet = mSaveRec()
        If Not ilRet Then
            imBoxNo = -1
            imRowNo = -1
            pbcClickFocus.SetFocus
            Exit Sub
        End If
        imSettingValue = True
        'If UBound(smSave, 2) <= vbcPosting.LargeChange + 1 Then
        If UBound(tgShow) <= vbcPosting.LargeChange + 1 Then
        ' If this is used, there are probably 0 or 1 records
            vbcPosting.Max = LBONE  'LBound(tgShow) 'LBound(smSave, 2)
        Else
        ' Saves, what amounts to, the count of records just retrieved
            'vbcPosting.Max = UBound(smSave, 2) - vbcPosting.LargeChange
            vbcPosting.Max = UBound(tgShow) - vbcPosting.LargeChange
        End If
        If imDropRowNo + 1 < UBound(tgShow) Then    'UBound(smSave, 2) Then
            imRowNo = imDropRowNo + 1
            imBoxNo = COPYINDEX '1
            mEnableBox imBoxNo
        Else
            imBoxNo = -1
            imRowNo = -1
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
End Sub
Private Sub pbcPosting_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If ((smDragCntrType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (smDragCntrType = "Q") Or (smDragCntrType = "R") Or ((smDragCntrType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((smDragCntrType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (smDragCntrType = "X") Or (smDragCntrType = "B") Then
        Exit Sub
    End If
    If (imDragSource = 1) Or (imDragSource = 2) Then
        If State = vbEnter Then    'Enter drag over
            lacMdFrame.DragIcon = IconTraf!imcIconMove.DragIcon
        ElseIf State = vbLeave Then
            lacMdFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
     End If
End Sub
Private Sub pbcPosting_GotFocus()
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
End Sub

Private Sub pbcPosting_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRet As Integer
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    imButton = Button
    If Button = 2 Then  'Right Mouse
        ilCompRow = vbcPosting.LargeChange + 1
        'If UBound(smSave, 2) - 1 > ilCompRow Then
        If UBound(tgShow) - 1 > ilCompRow Then
            'If UBound(smSave, 2) = vbcPosting.Value + ilCompRow - 1 Then
            If UBound(tgShow) = vbcPosting.Value + ilCompRow - 1 Then
                ilMaxRow = ilCompRow - 1
            Else
                ilMaxRow = ilCompRow
            End If
        Else
            ilMaxRow = UBound(tgShow) - 1   'UBound(smSave, 2) - 1
        End If
        ' Look through all rows
        For ilRow = 1 To ilMaxRow Step 1
            If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY + tmCtrls(1).fBoxH)) Then
                imButtonRow = ilRow + vbcPosting.Value - 1
                mShowInfo
                Exit For
            End If
        Next ilRow
        Exit Sub
    End If
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ilRet = mReadSdfRec(True) ' Get Remaining data
    'If rbcType(1).Value Then
    '    Exit Sub
    'End If
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    imDragSource = 0
    tmcDrag.Enabled = True  'Start timer to see if drag or clickEnd Sub
End Sub
Private Sub pbcPosting_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    If imIgnoreRightMove Then
        Exit Sub
    End If
    imButton = Button
    If Button <> 2 Then  'Right Mouse
        Exit Sub
    End If
    imButton = Button
    imIgnoreRightMove = True
    ilCompRow = vbcPosting.LargeChange + 1
    'If UBound(smSave, 2) - 1 > ilCompRow Then
    If UBound(tgShow) - 1 > ilCompRow Then
        'If UBound(smSave, 2) = vbcPosting.Value + ilCompRow - 1 Then
        If UBound(tgShow) = vbcPosting.Value + ilCompRow - 1 Then
            ilMaxRow = ilCompRow - 1
        Else
            ilMaxRow = ilCompRow
        End If
    Else
        ilMaxRow = UBound(tgShow) - 1   'UBound(smSave, 2) - 1
    End If
    ' Look through all rows
    For ilRow = 1 To ilMaxRow Step 1
        If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(1).fBoxY + tmCtrls(1).fBoxH)) Then
            imButtonRow = ilRow + vbcPosting.Value - 1
            mShowInfo
            Exit For
        End If
    Next ilRow
    imIgnoreRightMove = False
    Exit Sub
End Sub
Private Sub pbcPosting_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilLoop As Integer
    Dim ilSelected As Integer
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If Button = 2 Then
        plcInfo.Visible = False
        Exit Sub
    End If
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ilCompRow = vbcPosting.LargeChange + 1
    'If UBound(smSave, 2) - 1 > ilCompRow Then
    If UBound(tgShow) - 1 > ilCompRow Then
        'If UBound(smSave, 2) = vbcPosting.Value + ilCompRow - 1 Then
        If UBound(tgShow) = vbcPosting.Value + ilCompRow - 1 Then
            ilMaxRow = ilCompRow - 1
        Else
            ilMaxRow = ilCompRow
        End If
    Else
        ilMaxRow = UBound(tgShow) - 1   'UBound(smSave, 2) - 1
    End If
    ' Look through all rows
    For ilRow = 1 To ilMaxRow Step 1
    ' Look through all columns
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            ' See if this is the column
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                ' See of this is the row
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ' See if can format the data for this row/column
                    mDTSetShow imDTBoxNo
                    imDTBoxNo = -1
                    mTZSetShow imTZBoxNo
                    imTZBoxNo = -1
                    imTZRowNo = -1
                    If Not mSetShow(imBoxNo) Then
                        Beep
                        mEnableBox imBoxNo
                        Exit Sub
                    End If
                    ' These two are treated as one large box
                    If ilBox = ISCIINDEX Then
                        ilBox = COPYINDEX
                    End If
                    ' These three are the only ones that can accept user input
                    ilRowNo = ilRow + vbcPosting.Value - 1
                    'If (ilBox <> DATEINDEX) And (ilBox <> TIMEINDEX) And (ilBox <> COPYINDEX) And (ilBox <> PRICEINDEX) Then
                    'Note: Price field also tested later
'                    If (ilBox <> DATEINDEX) And (ilBox <> TIMEINDEX) And (ilBox <> COPYINDEX) Then
                    ''5/20/11
                    ''If (ilBox <> DATEINDEX) And (ilBox <> TIMEINDEX) And (ilBox <> COPYINDEX) And (ilBox <> PRICEINDEX) Then
                    '6/16/11
                    'If (ilBox <> TIMEINDEX) And (ilBox <> COPYINDEX) And (ilBox <> PRICEINDEX) Then
                    If (ilBox <> DATEINDEX) And (ilBox <> TIMEINDEX) And (ilBox <> COPYINDEX) And (ilBox <> PRICEINDEX) Then
                        If tgSave(tgShow(ilRowNo).iSaveInfoIndex).iBilled And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
                            Beep
                            Exit Sub
                        End If
                        'Assume row is to be checked
                        If (Button And vbLeftButton) > 0 Then
                            If (Shift = 0) Then
                                'For ilLoop = LBound(tgShow) To UBound(tgShow) - 1 Step 1
                                For ilLoop = LBONE To UBound(tgShow) - 1 Step 1
                                    tgShow(ilLoop).iChk = False
                                Next ilLoop
                                tgShow(ilRowNo).iChk = Not tgShow(ilRowNo).iChk
                                imLastRowSelected = ilRowNo
                            ElseIf (Shift And vbCtrlMask) > 0 Then
                                tgShow(ilRowNo).iChk = Not tgShow(ilRowNo).iChk
                                imLastRowSelected = ilRowNo
                            ElseIf (Shift And vbShiftMask) > 0 Then
                                'For ilLoop = LBound(tgShow) To UBound(tgShow) - 1 Step 1
                                For ilLoop = LBONE To UBound(tgShow) - 1 Step 1
                                    tgShow(ilLoop).iChk = False
                                Next ilLoop
                                If imLastRowSelected <> -1 Then
                                    If imLastRowSelected < ilRowNo Then
                                        For ilLoop = imLastRowSelected To ilRowNo Step 1
                                            tgShow(ilLoop).iChk = True
                                        Next ilLoop
                                    Else
                                        For ilLoop = ilRowNo To imLastRowSelected Step 1
                                            tgShow(ilLoop).iChk = True
                                        Next ilLoop
                                    End If
                                Else
                                    tgShow(ilRowNo).iChk = True
                                End If
                                imLastRowSelected = ilRowNo
                            End If
                            pbcPosting_Paint
                            ilSelected = False
                            'For ilLoop = LBound(tgShow) To UBound(tgShow) - 1 Step 1
                            For ilLoop = LBONE To UBound(tgShow) - 1 Step 1
                                If tgShow(ilLoop).iChk Then
                                    ilSelected = True
                                    imcTrash.Picture = IconTraf!imcBoxClosed.Picture
                                    'imcTrash.Visible = True
                                    'imcTrash.Enabled = True
                                    'If tgUrf(imUrfIndex).sHideSpots = "I" Then
                                    '    imcHidden.Picture = IconTraf!imcHideUp.Picture
                                    '    imcHidden.Visible = True
                                    '    imcHidden.Enabled = True
                                    'End If
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilSelected Then
                                imcTrash.Visible = False
                                imcHidden.Visible = False
                            End If
                        Else
                            Beep
                        End If
                        Exit Sub
                    End If
                    'If imPostSpotInfo(3, ilRowNo) And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
                    If (tgSave(tgShow(ilRowNo).iSaveInfoIndex).iBilled) And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
                        Beep
                        Exit Sub
                    End If
                    'Can only change price if true price field
                    'If (ilBox = PRICEINDEX) And ((imSave(1, ilRowNo) = -1) Or (tgUrf(imUrfIndex).sPrice <> "I")) Then
                    'Disallow any price change instead as stated above
                    If (ilBox = PRICEINDEX) Then
                        If (InStr(1, tgSave(ilRowNo).sPrice, "Fill", vbTextCompare) <= 0) Or (tgUrf(imUrfIndex).sPrice <> "I") Then
                            Beep
                            Exit Sub
                        End If
                    End If
                    'For ilLoop = LBound(tgShow) To UBound(tgShow) - 1 Step 1
                    For ilLoop = LBONE To UBound(tgShow) - 1 Step 1
                        tgShow(ilLoop).iChk = False
                    Next ilLoop
                    imcTrash.Visible = False
                    imcHidden.Visible = False
                    imLastRowSelected = -1
                    imRowNo = ilRowNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcPosting_Paint()
' paint the Post Log Form from smShow
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilSRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slFont As String
    Dim llColor As Long
    Dim llSvDateColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim slStr As String

    mPaintPostTitle
    ilStartRow = vbcPosting.Value   'Top location
    ilEndRow = vbcPosting.Value + vbcPosting.LargeChange
    If ilEndRow >= UBound(tgShow) Then  'UBound(smShow, 2) Then
        ilEndRow = UBound(tgShow) - 1 'UBound(smShow, 2) - 1
    End If
    llColor = pbcPosting.ForeColor
    pbcPosting.CurrentX = tmCtrls(COPYINDEX).fBoxX + fgBoxInsetX
    pbcPosting.CurrentY = 15 '- 30 '+ fgBoxInsetY
    slFontName = pbcPosting.FontName
    flFontSize = pbcPosting.FontSize
'    pbcPosting.ForeColor = BLUE
'    pbcPosting.FontBold = False
'    pbcPosting.FontSize = 7
'    pbcPosting.FontName = "Arial"
'    pbcPosting.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
'    If tgSpf.sUseCartNo <> "N" Then
'        pbcPosting.Print "Copy"
'    Else
'        pbcPosting.Print " "
'    End If
'    pbcPosting.FontSize = flFontSize
'    pbcPosting.FontName = slFontName
'    pbcPosting.FontSize = flFontSize
'    pbcPosting.ForeColor = llColor
'    pbcPosting.FontBold = True
    For ilRow = ilStartRow To ilEndRow Step 1
        ilSRow = tgShow(ilRow).iSaveInfoIndex
        'If imPostSpotInfo(1, ilRow) And imPostSpotInfo(2, ilRow) Then
        If StrComp(Trim$(tgSave(ilSRow).sPrice), "Feed", vbTextCompare) = 0 Then
            pbcPosting.ForeColor = DARKGRAY
        Else
            If tgSave(ilSRow).iISCIReq And tgSave(ilSRow).iISCI Then
                pbcPosting.ForeColor = RED
            End If
            'If imPostSpotInfo(3, ilRow) Then    'If billed- override any other color
            If tgSave(ilSRow).iBilled Then    'If billed- override any other color
                pbcPosting.ForeColor = DARKGREEN
            End If
        End If
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If ilBox = TIMEINDEX Then
                'If imPostSpotInfo(4, ilRow) Then    'If billed- override any other color
                If tgSave(ilSRow).iSimulCast Then    'If billed- override any other color
                    pbcPosting.ForeColor = MAGENTA
                End If
            End If
            If tgShow(ilRow).iChk Then
                gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, GRAY
            Else
                ''5/20/11
                ''If (ilBox = ADVTINDEX) Or (ilBox = LENINDEX) Or (ilBox = TZONEINDEX) Or (ilBox >= CNTRINDEX) Then
                '6/16/11
                'If (ilBox = DATEINDEX) Or (ilBox = ADVTINDEX) Or (ilBox = LENINDEX) Or (ilBox = TZONEINDEX) Or (ilBox >= CNTRINDEX) Then
                If (ilBox = ADVTINDEX) Or (ilBox = LENINDEX) Or (ilBox = TZONEINDEX) Or (ilBox >= CNTRINDEX) Then
                    gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                Else
                    gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
                End If
            End If
            If Not tgShow(ilRow).iChk Then
                ' Paint the box background with white or yellow
                If imBoxNo = COPYINDEX Then
                    If (ilBox = COPYINDEX) Or (ilBox = ISCIINDEX) Then
                        gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
                    ''5/20/11
                    ''ElseIf (ilBox = ADVTINDEX) Or (ilBox = TZONEINDEX) Then
                    '6/16/11
                    'ElseIf (ilBox = DATEINDEX) Or (ilBox = ADVTINDEX) Or (ilBox = TZONEINDEX) Then
                    ElseIf (ilBox = ADVTINDEX) Or (ilBox = TZONEINDEX) Then
                        gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                    End If
                End If
                If ilBox = AUDINDEX Then
                    gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                End If
                ''5/20/11
                '6/16/11
                'If ilBox = DATEINDEX Then
                '    gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                'End If
                '''If (ilBox = TIMEINDEX) Or (ilBox = PRICEINDEX) Then
                ''5/20/11
                ''If (ilBox = DATEINDEX) Or (ilBox = TIMEINDEX) Then
                '6/16/11
                'If (ilBox = TIMEINDEX) Then
                If (ilBox = DATEINDEX) Or (ilBox = TIMEINDEX) Then
                    gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
                'ElseIf (ilBox = LENINDEX) Or (ilBox = CNTRINDEX) Or (ilBox = LINEINDEX) Or (ilBox = TYPEINDEX) Or (ilBox = MGOODINDEX) Then
                ElseIf (ilBox = LENINDEX) Or (ilBox = CNTRINDEX) Or (ilBox = LINEINDEX) Or (ilBox = TYPEINDEX) Or (ilBox = MGOODINDEX) Or (ilBox = PRICEINDEX) Then
                    gPaintArea pbcPosting, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                End If
            End If
            If ilBox <> PRICEINDEX Then
                pbcPosting.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            Else
                'pbcPosting.CurrentX = gRightJustifyShowStr(pbcPosting, smShow(ilBox, ilRow), tmCtrls(ilBox))
                pbcPosting.CurrentX = gRightJustifyShowStr(pbcPosting, tgShow(ilRow).sShow(ilBox), tmCtrls(ilBox))
            End If
            If (ilBox = TZONEINDEX) Then
                pbcPosting.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                slFont = pbcPosting.FontName
                pbcPosting.FontName = "Monotype Sorts"
                pbcPosting.FontBold = False
                pbcPosting.Print tgShow(ilRow).sShow(ilBox)
                pbcPosting.FontName = slFont
                pbcPosting.FontBold = True
            ElseIf ilBox = PRICEINDEX Then
                If (tgUrf(imUrfIndex).sPrice <> "H") Then
                    pbcPosting.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30
                    pbcPosting.Print tgShow(ilRow).sShow(ilBox) ' Print one box of data at a time
                End If
            ElseIf ((ilBox = CNTRINDEX) Or (ilBox = LINEINDEX)) And (StrComp(Trim$(tgSave(ilSRow).sPrice), "Feed", vbTextCompare) = 0) Then
                'Leave blank
            ElseIf ilBox = DATEINDEX Then
                pbcPosting.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30
                If tgSave(ilSRow).sXMid <> "Y" Then
                    pbcPosting.Print tgShow(ilRow).sShow(ilBox) ' Print one box of data at a time
                Else
                    llSvDateColor = pbcPosting.ForeColor
                    pbcPosting.ForeColor = MAGENTA
                    slStr = Trim$(tgSave(ilSRow).sAirDate)
                    slStr = gIncOneDay(slStr)
                    gSetShow pbcPosting, slStr, tmCtrls(DATEINDEX)
                    pbcPosting.Print tmCtrls(DATEINDEX).sShow
                    pbcPosting.ForeColor = llSvDateColor
                End If
            Else
                pbcPosting.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30
                pbcPosting.Print tgShow(ilRow).sShow(ilBox) ' Print one box of data at a time
            End If
            If ilBox = TIMEINDEX Then
                pbcPosting.ForeColor = llColor
                        'If imPostSpotInfo(1, ilRow) And imPostSpotInfo(2, ilRow) Then
                If StrComp(Trim$(tgSave(ilSRow).sPrice), "Feed", vbTextCompare) = 0 Then
                    pbcPosting.ForeColor = DARKGRAY
                Else
                    If tgSave(ilSRow).iISCIReq And tgSave(ilSRow).iISCI Then
                        pbcPosting.ForeColor = RED
                    End If
                    'If imPostSpotInfo(3, ilRow) Then    'If billed- override any other color
                    If tgSave(ilSRow).iBilled Then    'If billed- override any other color
                        pbcPosting.ForeColor = DARKGREEN
                    End If
                End If
            End If
        Next ilBox
        pbcPosting.ForeColor = llColor
    Next ilRow
    '6/16/11
    'For ilRow = ilEndRow + 1 To vbcPosting.Value + vbcPosting.LargeChange Step 1
    '    gPaintArea pbcPosting, tmCtrls(DATEINDEX).fBoxX, tmCtrls(DATEINDEX).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(DATEINDEX).fBoxW - 15, tmCtrls(DATEINDEX).fBoxH - 15, LIGHTYELLOW
    'Next ilRow
End Sub
Private Sub pbcPrice_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub pbcPrice_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcPrice_KeyPress(KeyAscii As Integer)
    Dim ilRowNo As Integer
    If KeyAscii = Asc(" ") Then
        ilRowNo = tgShow(imRowNo).iSaveInfoIndex
'        If tgSave(ilRowNo).iPrice = 0 Then 'imSave(1, imRowNo) = 0 Then  'True price
'            'imSave(1, imRowNo) = 1
'            tgSave(ilRowNo).iPrice = 1
'            pbcPrice_Paint
'        ElseIf tgSave(ilRowNo).iPrice = 1 Then 'imSave(1, imRowNo) = 1 Then  'N/C
'            'imSave(1, imRowNo) = 0
'            tgSave(ilRowNo).iPrice = 0
'            pbcPrice_Paint
'        End If
        If tgSave(ilRowNo).iPrice = 0 Then '0=Inv-Advt
            tgSave(ilRowNo).iPrice = 1
            pbcPrice_Paint
        ElseIf tgSave(ilRowNo).iPrice = 1 Then '1=Inv-Yes
            tgSave(ilRowNo).iPrice = 2
            pbcPrice_Paint
        ElseIf tgSave(ilRowNo).iPrice = 2 Then '2=Inv-No
            tgSave(ilRowNo).iPrice = 0
            pbcPrice_Paint
        End If
    End If
End Sub
Private Sub pbcPrice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    ilRowNo = tgShow(imRowNo).iSaveInfoIndex
'    If tgSave(ilRowNo).iPrice = 0 Then 'imSave(1, imRowNo) = 0 Then
'        'imSave(1, imRowNo) = 1
'        tgSave(ilRowNo).iPrice = 1
'    Else
'        'imSave(1, imRowNo) = 0
'        tgSave(ilRowNo).iPrice = 0
'    End If
    If tgSave(ilRowNo).iPrice = 0 Then '0=Inv-Advt
        tgSave(ilRowNo).iPrice = 1
    ElseIf tgSave(ilRowNo).iPrice = 1 Then '1=Inv-Yes
        tgSave(ilRowNo).iPrice = 2
    Else
        tgSave(ilRowNo).iPrice = 0
    End If
    pbcPrice_Paint
End Sub
Private Sub pbcPrice_Paint()
    Dim ilRowNo As Integer
    pbcPrice.Cls
    pbcPrice.CurrentX = fgBoxInsetX
    pbcPrice.CurrentY = 0 'fgBoxInsetY
    ilRowNo = tgShow(imRowNo).iSaveInfoIndex
'    If tgSave(ilRowNo).iPrice = 0 Then  'imSave(1, imRowNo) = 0 Then
'        pbcPrice.Print Trim$(tgSave(ilRowNo).sPrice)    'smSave(SAVPRICEINDEX, imRowNo)
'    Else
'        pbcPrice.Print "N/C"
'    End If
    If tgSave(ilRowNo).iPrice = 1 Then
        pbcPrice.Print "Inv-Yes"
    ElseIf tgSave(ilRowNo).iPrice = 2 Then
        pbcPrice.Print "Inv-No"
    Else
        pbcPrice.Print "Inv-Advt"
    End If
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilFound As Integer
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    If imDateSelectedIndex < 0 Then
        If edcDate.Enabled Then
            edcDate.SetFocus
        Else
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    ilBox = imBoxNo
    ilRow = imRowNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                'If (UBound(smSave, 2) <= 1) Then
                If (UBound(tgShow) <= 1) Then
                    'cbcDate.SetFocus
                    edcDate.SetFocus
                    Exit Sub
                End If
                imTabDirection = 0  'Set-Left to right
                imRowNo = 1
                imBoxNo = 1
                imSettingValue = True
                vbcPosting.Value = vbcPosting.Min
                'If imPostSpotInfo(3, imRowNo) And (tgUrf(imUrfIndex).sHideSpots <> "I") Then  'Billed
                If tgSave(tgShow(imRowNo).iSaveInfoIndex).iBilled And (tgUrf(imUrfIndex).sHideSpots <> "I") Then  'Billed
                    imBoxNo = 0
                    pbcTab.SetFocus
                    Exit Sub
                Else
                    mEnableBox imBoxNo
                    Exit Sub
                End If
            Case 0  'form pbcTab when last row is billed
                ilBox = 1
            Case DATEINDEX, TIMEINDEX 'time (first control within header)
            ' hide control, save data, test for data change (set imSdfChg if changed)
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    If Not mSetShow(imBoxNo) Then
                        Beep
                        ''5/20/11
                        ''mDTEnableBox ilBox
                        'mEnableBox ilBox
                        '6/16/11
                        'mDTEnableBox ilBox
                        '1/11/12
                        cmcCancel.SetFocus
                        Exit Sub
                    End If
                End If
                ' From first row, return to date
                If imRowNo = 1 Then
                   imBoxNo = -1
                   imRowNo = -1
                   'cbcDate.SetFocus
                   edcDate.SetFocus
                   Exit Sub
                End If
                ' Back up one row
                imRowNo = imRowNo - 1
                Do While vbcPosting.Value > imRowNo
                    imSettingValue = True
                    vbcPosting.Value = vbcPosting.Value - 1
                Loop
                ilFound = False
                ilBox = PRICEINDEX
            Case COPYINDEX
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    If Not mSetShow(imBoxNo) Then
                        Beep
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
                ilFound = True  'next control
                ilBox = TIMEINDEX ' Point to the next control in line
            Case PRICEINDEX
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    If Not mSetShow(imBoxNo) Then
                        Beep
                        mEnableBox ilBox
                        Exit Sub
                    End If
                End If
                ilFound = False
                ilBox = COPYINDEX
        End Select
        'If imPostSpotInfo(3, imRowNo) And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
        If tgSave(tgShow(imRowNo).iSaveInfoIndex).iBilled And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
            ilFound = False
        End If
    Loop While Not ilFound
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilRet As Integer
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilFound As Integer
    Dim ilBottomLine As Integer
    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    If imInTab Then
        Exit Sub
    End If
    imInTab = True
    If imDateSelectedIndex < 0 Then
        If edcDate.Enabled Then
            edcDate.SetFocus
        Else
            pbcClickFocus.SetFocus
        End If
        imInTab = False
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    ilBox = imBoxNo
    ilRow = imRowNo
    Do
        ilFound = True
        Select Case ilBox

            Case -1 'Tab from control prior to form area
                imTabDirection = -1 ' Right to Left
                ilRet = mReadSdfRec(True) ' Get Remaining data
                imRowNo = UBound(tgShow) - 1    'UBound(smSave, 2) - 1' Find out Total Rows in array
                imBoxNo = 1
                ilBox = 1
                If imRowNo <= 0 Then
                    'cbcDate.SetFocus
                    edcDate.SetFocus
                    imInTab = False
                    Exit Sub
                End If
                imSettingValue = True
                If imRowNo <= vbcPosting.LargeChange + 1 Then
                    vbcPosting.Value = vbcPosting.Min
                Else
                    vbcPosting.Value = imRowNo - vbcPosting.LargeChange
                End If
                'If imPostSpotInfo(3, imRowNo) And (tgUrf(imUrfIndex).sChgBilled <> "I") Then   'Billed
                If tgSave(tgShow(imRowNo).iSaveInfoIndex).iBilled And (tgUrf(imUrfIndex).sChgBilled <> "I") Then   'Billed
                    imBoxNo = 0
                    pbcSTab.SetFocus
                    imInTab = False
                    Exit Sub
                Else
                    mEnableBox imBoxNo
                    imInTab = False
                    Exit Sub
                End If
            Case 0  'Tab from pbcSTab when first row is billed
                ilBox = 1
            Case DATEINDEX, TIMEINDEX ' Came here from TIME
                ' hide control, save data, test for data change (set imSdfChg if changed)
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    'If lbcAvailTimes.ListIndex < 0 Then
                    '    Beep
                    '    lbcAvailTimes.SetFocus
                    '    Exit Sub
                    'End If
                    If Not mSetShow(imBoxNo) Then
                        Beep
                        ''5/20/11
                        ''mDTEnableBox ilBox
                        'mEnableBox ilBox
                        '6/16/11
                        'mDTEnableBox ilBox
                        imInTab = False
                        '1/11/12
                        cmcCancel.SetFocus
                        Exit Sub
                    End If
                End If
                ilFound = False
                ilBox = COPYINDEX
            Case COPYINDEX  ' Came here from COPY/ISCI
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    If Not mSetShow(imBoxNo) Then
                        Beep
                        mEnableBox ilBox
                        imInTab = False
                        Exit Sub
                    End If
                End If
                ilFound = False  ' Skips the next control
                ilBox = PRICEINDEX ' Point to the next control in line

            Case PRICEINDEX  ' Came here from PRICE
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    If Not mSetShow(imBoxNo) Then
                        Beep
                        mEnableBox ilBox
                        imInTab = False
                        Exit Sub
                    End If
                End If
                imRowNo = imRowNo + 1 ' Tab to next row
              ' Cause a scroll event which then causes a pbcPosting_Paint event
              ilBottomLine = imRowNo - vbcPosting.Value ' difference relative to the top
              'If ilBottomLine <= vbcPosting.LargeChange Or ((UBound(smSave, 2) - imRowNo) = 0) Then
              If ilBottomLine <= vbcPosting.LargeChange Or ((UBound(tgShow) - imRowNo) = 0) Then
                 ilBottomLine = 0
              End If
              If ilBottomLine Then
                  imSettingValue = True
                  vbcPosting.Value = vbcPosting.Value + 1 ' Move thumb down
              End If
             ' if on last screen of data, tab to DATE
             If imRowNo >= UBound(tgShow) Then  'UBound(smSave, 2) Then
                pbcPosting_Paint    'required because of a timing problem
                imBoxNo = -1
                imRowNo = -1
                'cbcDate.SetFocus
                edcDate.SetFocus
                imInTab = False
                Exit Sub
             End If

            ilBox = TIMEINDEX ' Point to the next control in line
            ilFound = True ' Not a tab-stop field
        End Select
        'If imPostSpotInfo(3, imRowNo) And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
        'Fix for fast tabbing on time with changes 9/17/03
        'If (imRowNo < LBound(tgShow)) Or (imRowNo > UBound(tgShow)) Then
        If (imRowNo < LBONE) Or (imRowNo > UBound(tgShow)) Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
        'If (tgShow(imRowNo).iSaveInfoIndex < LBound(tgSave)) Or (tgShow(imRowNo).iSaveInfoIndex > UBound(tgSave)) Then
        If (tgShow(imRowNo).iSaveInfoIndex < LBONE) Or (tgShow(imRowNo).iSaveInfoIndex > UBound(tgSave)) Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
        If tgSave(tgShow(imRowNo).iSaveInfoIndex).iBilled And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
            ilFound = False
        End If
    Loop While Not ilFound
    imBoxNo = ilBox
    mEnableBox ilBox
    imInTab = False
    Exit Sub

    imInTab = False
    On Error GoTo 0
End Sub
Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    '5/20/11
                    'If imDTBoxNo = imDTAIRTIMEINDEX Then
                    '    imBypassFocus = True    'Don't change select text
                    '    edcDTDropDown.SetFocus
                    '    'SendKeys slKey
                    '    gSendKeys edcDTDropDown, slKey
                    'Else
                    '6/16/11
                    If imDTBoxNo = imDTAIRTIMEINDEX Then
                        imBypassFocus = True    'Don't change select text
                        edcDTDropDown.SetFocus
                        'SendKeys slKey
                        gSendKeys edcDTDropDown, slKey
                    Else
                        Select Case imBoxNo
                            Case TIMEINDEX
                                imBypassFocus = True    'Don't change select text
                                edcDropDown.SetFocus
                                'SendKeys slKey
                                gSendKeys edcDropDown, slKey
                        End Select
                    '6/16/11
                    End If
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTZCopy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If tmcDrag.Enabled Then
        imTZDragType = -1
        tmcDrag.Enabled = False
    End If
    ' Look through all rows
    For ilRow = 1 To imNoZones Step 1
    ' Look through all columns
        For ilBox = 2 To 2 Step 1   'Disallow zone box
            ' See if this is the column
            If (X >= tmTZCtrls(ilBox).fBoxX) And (X <= (tmTZCtrls(ilBox).fBoxX + tmTZCtrls(ilBox).fBoxW)) Then
                ' See of this is the row
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmTZCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmTZCtrls(ilBox).fBoxY + tmTZCtrls(ilBox).fBoxH)) Then
                    ' See if can format the data for this row/column
                    'If (ilBox = 1) And (ilRow > 1) Then
                    '    Beep
                    '    Exit Sub
                    'End If
                    'If (ilRow > 1) And (imTZSave = 0) Then
                    '    Beep
                    '    Exit Sub
                    'End If
                    mTZSetShow imTZBoxNo
                    ilRowNo = ilRow
                    imTZRowNo = ilRowNo
                    imTZBoxNo = ilBox
                    mTZEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcTZCopy_Paint()
' paint the Time Zone Form from smTZShow
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    llColor = pbcTZCopy.ForeColor
    pbcTZCopy.CurrentX = tmTZCtrls(TZCOPYINDEX).fBoxX + fgBoxInsetX
    pbcTZCopy.CurrentY = 15 '- 30 '+ fgBoxInsetY
    slFontName = pbcTZCopy.FontName
    flFontSize = pbcTZCopy.FontSize
    pbcTZCopy.ForeColor = BLUE
    pbcTZCopy.FontBold = False
    pbcTZCopy.FontSize = 7
    pbcTZCopy.FontName = "Arial"
    pbcTZCopy.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    If tgSpf.sUseCartNo <> "N" Then
        pbcTZCopy.Print "Copy/ISCI/Product"
    Else
        pbcTZCopy.Print "ISCI/Product"
    End If
    pbcTZCopy.FontSize = flFontSize
    pbcTZCopy.FontName = slFontName
    pbcTZCopy.FontSize = flFontSize
    pbcTZCopy.ForeColor = llColor
    pbcTZCopy.FontBold = True
    ilStartRow = 1   'Top location
    ilEndRow = 6
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBTZCtrls To UBound(tmTZCtrls) Step 1
            pbcTZCopy.CurrentX = tmTZCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcTZCopy.CurrentY = tmTZCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 15
            pbcTZCopy.Print smTZShow(ilBox, ilRow) ' Print one box of data at a time
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcTZSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    Dim ilRowNo As Integer
    'If imRowNo = -1 Then
    '    mTZSetShow imTZBoxNo
    '    imTZBoxNo = -1
    '    imTZRowNo = -1
    '    pbcClickFocus.SetFocus
    '    Exit Sub
    'End If
    imTZTabDirection = -1 'Set- Right to left
    ilBox = imTZBoxNo
    ilRowNo = imTZRowNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTZTabDirection = 0  'Set-Left to right
                imTZRowNo = 1
                imTZBoxNo = TZCOPYINDEX 'ZONEINDEX
                mTZEnableBox imTZBoxNo
                Exit Sub
            Case ZONEINDEX 'time (first control within header)
            ' hide control, save data, test for data change (set imSdfChg if changed)
                ' From first row, return to edcDropDown
                If imTZRowNo = 1 Then
                   mTZSetShow imTZBoxNo
                   imTZBoxNo = -1
                   imTZRowNo = -1
                   pbcSTab.SetFocus
                   Exit Sub
                End If
                ' Back up one row
                ilRowNo = ilRowNo - 1
                'If (ilRowNo <> 1) And (imTZSave = 0) Then
                '    ilFound = False
                'End If
                ilBox = TZCOPYINDEX
            Case TZCOPYINDEX
                ilBox = ZONEINDEX ' Point to the next control in line
                'If ilRowNo <> 1 Then
                    ilFound = False
                'End If
        End Select
    Loop While Not ilFound
    mTZSetShow imTZBoxNo
    imTZRowNo = ilRowNo
    imTZBoxNo = ilBox
    mTZEnableBox ilBox
End Sub
Private Sub pbcTZTab_GotFocus()
    Dim ilBox As Integer
    Dim ilFound As Integer
    Dim ilRowNo As Integer
    Dim ilLoop As Integer
    imTabDirection = 0  'Set-Left to right
    ilBox = imTZBoxNo
    ilRowNo = imTZRowNo
    Do
        ilFound = True
        Select Case ilBox

            Case -1 'Tab from control prior to form area
                imTZTabDirection = -1 ' Right to Left
                ilRowNo = 1
                For ilLoop = imNoZones To 1 Step -1
                    If smTZSave(1, ilLoop) <> "" Then
                        ilRowNo = ilLoop
                        Exit For
                    End If
                Next ilLoop
                ilBox = TZCOPYINDEX
            Case ZONEINDEX ' Came here from ZONE
                ' hide control, save data, test for data change (set imSdfChg if changed)
                'If (smTZSave(1, ilRowNo) = "") Or ((ilRowNo <> 1) And (imTZSave = 0)) Then
                If (smTZSave(1, ilRowNo) = "") Then
                    ilFound = False
                End If
                ilBox = TZCOPYINDEX
            Case TZCOPYINDEX  ' Came here from COPY/ISCI/PRODUCT
                If ilBox = imTZBoxNo Then
                    mTZSetShow imTZBoxNo
                End If
                ilRowNo = ilRowNo + 1 ' Tab to next row
                If (ilRowNo > imNoZones) Or ((ilRowNo = 2) And (StrComp(smTZSave(2, 1), "[None]", 1) <> 0)) Then
                    mTZSetShow imTZBoxNo
                    imTZBoxNo = -1
                    imTZRowNo = -1
                    pbcTab.SetFocus
                    Exit Sub
                End If
                'If ilRowNo <> 1 Then
                    ilFound = False
                'End If
                ilBox = ZONEINDEX
        End Select
    Loop While Not ilFound
    mTZSetShow imTZBoxNo
    imTZRowNo = ilRowNo
    imTZBoxNo = ilBox
    mTZEnableBox ilBox
End Sub
Private Sub pbcTZZone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcTZZone_KeyPress(KeyAscii As Integer)
    Dim ilLoop As Integer
    Dim slChar As String
    For ilLoop = 1 To imNoZones Step 1
        If ilLoop <= 3 Then
            slChar = Mid$(smZones(ilLoop), 2, 1)
        Else
            slChar = Mid$(smZones(ilLoop), 1, 1)
        End If
        If (KeyAscii = Asc(UCase(slChar))) Or (KeyAscii = Asc(LCase(slChar))) Then
            imTZSave = ilLoop - 1
            pbcTZZone_Paint
            Exit Sub
        End If
    Next ilLoop
    If KeyAscii = Asc(" ") Then
        imTZSave = imTZSave + 1
        If imTZSave > imNoZones Then
            imTZSave = 0
        End If
        pbcTZZone_Paint
    End If
End Sub
Private Sub pbcTZZone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imTZSave = imTZSave + 1
    If imTZSave > imNoZones Then
        imTZSave = 0
    End If
    pbcTZZone_Paint
End Sub
Private Sub pbcTZZone_Paint()
    pbcTZZone.Cls
    pbcTZZone.CurrentX = fgBoxInsetX
    pbcTZZone.CurrentY = 0 'fgBoxInsetY
    pbcTZZone.Print smZones(imTZSave - 1)
End Sub
Private Sub pbcWM_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub pbcWM_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcWM_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("W")) Or (KeyAscii = Asc("w")) Then
        imWM = 0
        pbcWM_Paint
        edcDate_Change
        Exit Sub
    End If
    If (KeyAscii = Asc("M")) Or (KeyAscii = Asc("m")) Then
        imWM = 1
        pbcWM_Paint
        edcDate_Change
        Exit Sub
    End If
    If KeyAscii = Asc(" ") Then
        If imWM = 0 Then 'imSave(1, imRowNo) = 0 Then  'True price
            imWM = 1
            pbcWM_Paint
        ElseIf imWM = 1 Then 'imSave(1, imRowNo) = 1 Then  'N/C
            imWM = 0
            pbcWM_Paint
        End If
        edcDate_Change
    End If
End Sub
Private Sub pbcWM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imWM = 0 Then
        imWM = 1
    Else
        imWM = 0
    End If
    pbcWM_Paint
    edcDate_Change
End Sub
Private Sub pbcWM_Paint()
    pbcWM.Cls
    pbcWM.CurrentX = fgBoxInsetX
    pbcWM.CurrentY = 0 'fgBoxInsetY
    If imWM = 0 Then
        pbcWM.Print "Wk"    'smSave(SAVPRICEINDEX, imRowNo)
    Else
        pbcWM.Print "Mo"
    End If
End Sub

Private Sub plcComplete_Paint()
    plcComplete.CurrentX = 0
    plcComplete.CurrentY = 0
    plcComplete.Print "Day is Complete"
End Sub


Private Sub plcPosting_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcPosting_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub plcPosting_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub plcSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcSpots_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSpots_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub plcSpots_DragOver(Source As control, X As Single, Y As Single, State As Integer)
' fgListHTArial825 is the size of a standard row in listboxes
     Dim ilLeftEdge As Integer
     Dim ilRightEdge As Integer
     Dim ilLowerBotEdge As Integer
     Dim ilUpperBotEdge As Integer
     Dim ilMissedBottom As Integer
     If imDragSource = 2 Then
        Exit Sub
     End If
     ilLeftEdge = lbcMissed.Left
     ilRightEdge = lbcMissed.Left + lbcMissed.Width
     ilUpperBotEdge = lbcMissed.Top - 1
     ilLowerBotEdge = lbcMissed.Top + lbcMissed.height + fgListHtArial825
     ilMissedBottom = lbcMissed.Top + lbcMissed.height
     ' On entering this panel turn off scrolling if not in Hot Box
     If (X < ilLeftEdge) Or (X > ilRightEdge) And State = ENTER Then
        tmcDrag.Enabled = False
        tmcDrag.Interval = 1000
     End If
     If ((X < ilLeftEdge) Or (X > ilRightEdge) Or (Y > ilLowerBotEdge)) And State = ENTER Then
        tmcDrag.Enabled = False
        tmcDrag.Interval = 1000
     End If
     ' Scroll if inside Hot Box
     If (X >= ilLeftEdge) And (X <= ilRightEdge) And (Y <= ilUpperBotEdge) And State = ENTER Then
        imScrollDirection = SCROLLUP
        tmcDrag.Enabled = True
        tmcDrag.Interval = imTimerInterval
     End If
     If (X >= ilLeftEdge) And (X <= ilRightEdge) And (Y > ilMissedBottom) And (Y <= ilLowerBotEdge) And State = ENTER Then
        imScrollDirection = SCROLLDN
        tmcDrag.Enabled = True
        tmcDrag.Interval = imTimerInterval
     End If
End Sub
Private Sub plcSpots_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        imDragSource = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub rbcSort_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSort(Index).Value
    'End of coded added
    If Value Then
        Screen.MousePointer = vbHourglass
        pbcPosting.Cls
        mPaintPostTitle
        mGenShowImage
        pbcPosting_Paint
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcSort_GotFocus(Index As Integer)
    If imFirstTime Then
        imFirstTime = False
    End If
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub rbcType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim ilLoop As Integer
    Dim Value As Integer
    Value = rbcType(Index).Value
    'End of coded added
    If Value Then
        pbcPosting.Cls
        mPaintPostTitle
        imChgMode = True
        imVehSelectedIndex = -1
        If Index = 0 Then
            ''ckcDayComplete.Visible = True
            '9/7/06- Show once vehicle selected
            'plcComplete.Visible = True
            plcComplete.Visible = False
            plcSpots.Visible = True
            pbcMissed.Visible = True
            hmSdf = hmSvSdf
            '6/16/11
            pbcDT.height = 1080
            plcDT.height = 1425
            imDTAIRTIMEINDEX = 3
            imDTAVAILTIMEINDEX = 2
        Else
            'ckcDayComplete.Visible = False
            plcComplete.Visible = False
            '2/21/13: Allow package spots to be cancelled
            'plcSpots.Visible = False
            'pbcMissed.Visible = False
            plcSpots.Visible = True
            pbcMissed.Visible = True
            hmSdf = hmPsf
            '6/16/11
            pbcDT.height = 720
            plcDT.height = 1065
            imDTAIRTIMEINDEX = 2
            imDTAVAILTIMEINDEX = -1
        End If
        mVehPop
        imChgMode = False
        gFindMatch sgUserDefVehicleName, 0, cbcVeh
        If gLastFound(cbcVeh) >= 0 Then
            cbcVeh.ListIndex = gLastFound(cbcVeh)
        Else
            If cbcVeh.ListCount > 0 Then
                If cbcVeh.ListIndex = 0 Then
                    cbcVeh_Change
                Else
                    cbcVeh.ListIndex = 0
                End If
            End If
        End If
        If Index = 0 Then
            If bmInPackage Then
                imMissedType = imSvMissedType
                pbcMissedType_Paint
                pbcMissedType.Enabled = True
                mMissedPop
                bmInPackage = False
                For ilLoop = 0 To 3 Step 1
                    ckcInclude(ilLoop) = imSvCkcInclude(ilLoop)
                    ckcInclude(ilLoop).Enabled = True
                Next ilLoop
            End If
        Else
            bmInPackage = True
            imSvMissedType = imMissedType
            imMissedType = 1
            pbcMissedType_Paint
            pbcMissedType.Enabled = False
            mMissedPop
            For ilLoop = 0 To 3 Step 1
                imSvCkcInclude(ilLoop) = ckcInclude(ilLoop)
            Next ilLoop
            For ilLoop = 0 To 3 Step 1
                If (ilLoop = 0) Or (ilLoop = 2) Then
                    ckcInclude(ilLoop) = vbUnchecked
                    ckcInclude(ilLoop).Enabled = False
                Else
                    ckcInclude(ilLoop) = vbChecked
                    ckcInclude(ilLoop).Enabled = True
                End If
            Next ilLoop
        End If
        pbcMissed_Paint
    End If
End Sub
Private Sub rbcType_GotFocus(Index As Integer)
    If imFirstTime Then
        imFirstTime = False
    End If
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
End Sub
Private Sub tmcClick_Timer()
    Dim ilRet As Integer
    tmcClick.Enabled = False
    If imSelectDelay Then
        imProcClickMode = True
        imSelectDelay = False
        mCbcVehChange
        imProcClickMode = False
    Else
        Screen.MousePointer = vbHourglass
        plcCalendar.Visible = False
        lbcGameNo.Visible = False
        imProcClickMode = True
        pbcPosting.Cls
        mPaintPostTitle
        ilRet = mReadSdfRec(True) ' Get Remaining data
        mGenShowImage
        pbcPosting_Paint
        imProcClickMode = False
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Dim ilRet As Integer
    Dim llRecPos As Long
    Dim ilIndex As Integer
    Dim llDate As Long
    Dim llPrice As Long
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            imcTrash.Visible = False
            imcHidden.Visible = False
            'For ilRow = LBound(tgShow) To UBound(tgShow) - 1 Step 1
            For ilRow = LBONE To UBound(tgShow) - 1 Step 1
                tgShow(ilRow).iChk = False
            Next ilRow
            imLastRowSelected = -1
            If imDragSource = 0 Then
                ilCompRow = vbcPosting.LargeChange + 1
                'If UBound(smSave, 2) > ilCompRow Then
                If UBound(tgShow) > ilCompRow Then
                    ilMaxRow = ilCompRow
                Else
                    ilMaxRow = UBound(tgShow) - 1   'UBound(smSave, 2) - 1
                End If
                For ilRow = 1 To ilMaxRow Step 1
                    If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(TIMEINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(TIMEINDEX).fBoxY + tmCtrls(TIMEINDEX).fBoxH)) Then
                        mDTSetShow imDTBoxNo
                        imDTBoxNo = -1
                        mTZSetShow imTZBoxNo
                        imTZBoxNo = -1
                        imTZRowNo = -1
                        If Not mSetShow(imBoxNo) Then
                            Exit Sub
                        End If
                        'If imPostSpotInfo(3, ilRow + vbcPosting.Value - 1) And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
                        If tgSave(ilRow + vbcPosting.Value - 1).iBilled And (tgUrf(imUrfIndex).sChgBilled <> "I") Then  'Billed
                            Beep
                            Exit Sub
                        End If
                        imBoxNo = -1
                        imRowNo = -1
                        imRowNo = ilRow + vbcPosting.Value - 1
                        llRecPos = tgSave(tgShow(imRowNo).iSaveInfoIndex).lSdfRecPos   'Val(smSave(SAVRECPOSINDEX, imRowNo))
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                        tmChfSrchKey.lCode = tmSdf.lChfCode
                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        '2/21/13: Allow Package spots to be cancelled
                        If (rbcType(1).Value) Or ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "Q") Or (tmChf.sType = "R") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (tmSdf.sSpotType = "X") Or (tmSdf.sSpotType = "C") Or (tmSdf.sSpotType = "O") Then
                        'If ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "Q") Or (tmChf.sType = "R") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (tmSdf.sSpotType = "X") Or (tmSdf.sSpotType = "C") Or (tmSdf.sSpotType = "O") Then
                            imcTrash.Picture = IconTraf!imcFireOut.Picture
                            imcTrash.Visible = True
                            imcTrash.Enabled = True
                            smDragCntrType = tmChf.sType
                            If tmSdf.sSpotType = "X" Then
                                smDragCntrType = "X"
                            ElseIf tmSdf.sSpotType = "O" Then
                                smDragCntrType = "B"
                            ElseIf tmSdf.sSpotType = "C" Then
                                smDragCntrType = "B"
                            End If
                        Else
                            imcTrash.Picture = IconTraf!imcBoxClosed.Picture
                            'imcTrash.Visible = True
                            'imcTrash.Enabled = True
                            'If tgUrf(imUrfIndex).sHideSpots = "I" Then
                            '    imcHidden.Picture = IconTraf!imcHideUp.Picture
                            '    imcHidden.Visible = True
                            '    imcHidden.Enabled = True
                            'End If
                            smDragCntrType = tmChf.sType
                        End If
                        lacPtFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                        lacPtFrame.Move 0, tmCtrls(TIMEINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacPtFrame.Visible = True
                        pbcArrow.Move pbcArrow.Left, plcPosting.Top + tmCtrls(TIMEINDEX).fBoxY + (imRowNo - vbcPosting.Value) * (fgBoxGridH + 15) + 45
                        pbcArrow.Visible = True
                        lacPtFrame.Drag vbBeginDrag
                        lacPtFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                        Exit Sub
                    End If
                Next ilRow
            ElseIf imDragSource = 1 Then
                '2/21/13: Allow Package spots to be cancelled
                'If rbcType(1).Value Then
                '    imDragSource = -1
                '    Exit Sub
                'End If
                ilCompRow = vbcMissed.LargeChange + 1
                If UBound(tgMdShowInfo) > ilCompRow Then
                    ilMaxRow = ilCompRow
                Else
                    ilMaxRow = UBound(tgMdShowInfo) - 1
                End If
                For ilRow = 1 To ilMaxRow Step 1
                    If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmMdCtrls(MDADVTINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmMdCtrls(MDADVTINDEX).fBoxY + tmMdCtrls(MDADVTINDEX).fBoxH)) Then
                        mDTSetShow imDTBoxNo
                        imDTBoxNo = -1
                        mTZSetShow imTZBoxNo
                        imTZBoxNo = -1
                        imTZRowNo = -1
                        If Not mSetShow(imBoxNo) Then
                            Exit Sub
                        End If
                        imBoxNo = -1
                        imRowNo = -1
                        imMdRowNo = ilRow + vbcMissed.Value - 1
                        imSaveIndex = tgMdShowInfo(imMdRowNo).iMdSaveInfoIndex
                        If tgMdSaveInfo(imSaveIndex).lChfCode > 0 Then
                            tmChfSrchKey.lCode = tgMdSaveInfo(imSaveIndex).lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                imDragSource = -1
                                Exit Sub
                            End If
                        Else
                            ilRet = mReadChfClfRdfRec(tgMdSaveInfo(imSaveIndex).lChfCode, 0, tgMdSaveInfo(imSaveIndex).lFsfCode)
                        End If
                        If ((tmChf.sType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmChf.sType = "Q") Or (tmChf.sType = "R") Or ((tmChf.sType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmChf.sType = "M") And (tgSpf.sSchdPromo <> "Y")) Then
                            imDragSource = -1
                            Exit Sub
                        End If
                        smDragCntrType = tmChf.sType
                        If tgMdShowInfo(imMdRowNo).iType = 0 Then   'Contract
                            imDragSource = 2
                        ElseIf tgMdShowInfo(imMdRowNo).iType = 2 Then   'Hidden
                            llDate = 0
                            llPrice = -1
                            imSdfIndex = -1
                            ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                            Do While ilIndex >= 0
                                'If (tgMdSdfRec(ilIndex).lMissedDate > llDate) And (tgMdSdfRec(ilIndex).sSchStatus = "H") Then
                                '    llDate = tgMdSdfRec(ilIndex).lMissedDate
                                '    imSdfIndex = ilIndex
                                'End If
                                If (tgMdSdfRec(ilIndex).lPrice > llPrice) And (tgMdSdfRec(ilIndex).sSchStatus = "H") Then
                                    llPrice = tgMdSdfRec(ilIndex).lPrice
                                    imSdfIndex = ilIndex
                                End If
                                ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                            Loop
                            If imSdfIndex = -1 Then
                                ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                                Do While ilIndex >= 0
                                    If (tgMdSdfRec(ilIndex).sSchStatus = "H") Then
                                        imSdfIndex = ilIndex
                                        Exit Do
                                    End If
                                    ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                                Loop
                            End If
                            If imSdfIndex = -1 Then
                                imDragSource = 2
                            Else
                                imDragSource = 1
                            End If
                        ElseIf tgMdShowInfo(imMdRowNo).iType = 3 Then
                            llDate = 0
                            llPrice = -1
                            imSdfIndex = -1
                            ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                            Do While ilIndex >= 0
                                'If (tgMdSdfRec(ilIndex).lMissedDate > llDate) And (tgMdSdfRec(ilIndex).sSchStatus = "C") Then
                                '    llDate = tgMdSdfRec(ilIndex).lMissedDate
                                '    imSdfIndex = ilIndex
                                'End If
                                If (tgMdSdfRec(ilIndex).lPrice > llPrice) And (tgMdSdfRec(ilIndex).sSchStatus = "C") Then
                                    llPrice = tgMdSdfRec(ilIndex).lPrice
                                    imSdfIndex = ilIndex
                                End If
                                ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                            Loop
                            If imSdfIndex = -1 Then
                                ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                                Do While ilIndex >= 0
                                    If (tgMdSdfRec(ilIndex).sSchStatus = "C") Then
                                        imSdfIndex = ilIndex
                                        Exit Do
                                    End If
                                    ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                                Loop
                            End If
                            If imSdfIndex = -1 Then
                                imDragSource = 2
                            Else
                                imDragSource = 1
                            End If
                        Else
                            llDate = 0
                            llPrice = -1
                            imSdfIndex = -1
                            ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                            Do While ilIndex >= 0
                                'If (tgMdSdfRec(ilIndex).lMissedDate > llDate) And (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
                                '    llDate = tgMdSdfRec(ilIndex).lMissedDate
                                '    imSdfIndex = ilIndex
                                'End If
                                If (tgMdSdfRec(ilIndex).lPrice > llPrice) And (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
                                    llPrice = tgMdSdfRec(ilIndex).lPrice
                                    imSdfIndex = ilIndex
                                End If
                                ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                            Loop
                            If imSdfIndex = -1 Then
                                ilIndex = tgMdSaveInfo(imSaveIndex).iFirstIndex
                                Do While ilIndex >= 0
                                    If (tgMdSdfRec(ilIndex).sSchStatus = "M") Then
                                        imSdfIndex = ilIndex
                                        Exit Do
                                    End If
                                    ilIndex = tgMdSdfRec(ilIndex).iNextIndex
                                Loop
                            End If
                            If imSdfIndex = -1 Then
                                imDragSource = 2
                            Else
                                imDragSource = 1
                            End If
                        End If
                        If imDragSource = 1 Then
                            'If smMdSchStatus(1, imMdRowNo) <> "C" Then
                            If tgMdShowInfo(imMdRowNo).iType <> 3 Then
                                imcTrash.Picture = IconTraf!imcBoxClosed.Picture
                                'imcTrash.Visible = True
                                'imcTrash.Enabled = True
                            End If
                            ''If smMdSchStatus(1, imMdRowNo) <> "H" Then
                            'If tgMdShowInfo(imMdRowNo).iType <> 2 Then
                            '    If tgUrf(imUrfIndex).sHideSpots = "I" Then
                            '        imcHidden.Picture = IconTraf!imcHideUp.Picture
                            '        imcHidden.Visible = True
                            '        imcHidden.Enabled = True
                            '    End If
                            'End If
                        Else
                        End If
                        lacMdFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                        lacMdFrame.Move 0, tmMdCtrls(MDADVTINDEX).fBoxY + (imMdRowNo - vbcMissed.Value) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacMdFrame.Visible = True
                        lacMdFrame.Drag vbBeginDrag
                        lacMdFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                        Exit Sub
                    End If
                Next ilRow
            End If
    End Select
    ' These values are set in plcSpot
    If imScrollDirection = SCROLLUP Then  'scroll up
        If lbcMissed.TopIndex > 0 Then
            lbcMissed.TopIndex = lbcMissed.TopIndex - 1
        End If
    End If
    If imScrollDirection = SCROLLDN Then  'Scroll down
        If lbcMissed.TopIndex < lbcMissed.ListCount Then
            lbcMissed.TopIndex = lbcMissed.TopIndex + 1
        End If
    End If
End Sub
Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    pbcWM_Paint
End Sub
Private Sub vbcMissed_Change()
    If imSettingValue Then
        pbcMissed.Cls
        pbcMissed_Paint
        imSettingValue = False
    Else
        mTZSetShow imTZBoxNo
        imTZBoxNo = -1
        imTZRowNo = -1
        If Not mSetShow(imBoxNo) Then
            'Exit Sub
        End If
        imBoxNo = -1
        imRowNo = -1
        pbcMissed.Cls
        pbcMissed_Paint
        'mEnableBox imBoxNo
    End If
End Sub
Private Sub vbcMissed_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub vbcMissed_GotFocus()
    Dim ilRet As Integer
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        'Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    ilRet = mReadSdfRec(True) ' Get Remaining data
End Sub
Private Sub vbcPosting_Change()
    Dim ilRet As Integer
    If imSettingValue Then
        pbcPosting.Cls
        pbcPosting_Paint
        imSettingValue = False
    Else
        'getting change event prior to focus event
        mDTSetShow imDTBoxNo
        imDTBoxNo = -1
        mTZSetShow imTZBoxNo
        imTZBoxNo = -1
        imTZRowNo = -1
        If Not mSetShow(imBoxNo) Then
            'Exit Sub
        End If
        imBoxNo = -1
        imRowNo = -1
        ilRet = mReadSdfRec(True) ' Get Remaining data
        pbcPosting.Cls
        pbcPosting_Paint
        mEnableBox imBoxNo
    End If
End Sub
Private Sub vbcPosting_DragDrop(Source As control, X As Single, Y As Single)
    imcTrash.Visible = False
    imcHidden.Visible = False
End Sub
Private Sub vbcPosting_GotFocus()
    Dim ilRet As Integer
    plcCalendar.Visible = False
    lbcGameNo.Visible = False
    mDTSetShow imDTBoxNo
    imDTBoxNo = -1
    mTZSetShow imTZBoxNo
    imTZBoxNo = -1
    imTZRowNo = -1
    If Not mSetShow(imBoxNo) Then
        'Exit Sub
    End If
    imBoxNo = -1
    imRowNo = -1
    ilRet = mReadSdfRec(True) ' Get Remaining data
End Sub
Private Sub plcDT_Paint()
    plcDT.CurrentX = 0
    plcDT.CurrentY = 0
    plcDT.Print "Aired Date/Time"
End Sub
Private Sub plcTZCopy_Paint()
    plcTZCopy.CurrentX = 0
    plcTZCopy.CurrentY = 0
    plcTZCopy.Print "Copy"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Post Log"
End Sub
Private Sub plcSort_Paint()
    plcSort.CurrentX = 0
    plcSort.CurrentY = 0
    plcSort.Print "Sort by"
End Sub
'
'
'               mDetermineDayUpdate - determine the actual date to set
'               as complete based on the date user selected from the
'               calendar and the index of the day of week to set complete
'               <input> index to day of week to set as complete
'               <return> -1 if day not found in tgdates array (not valid
'                       day to set as complete; otherwise day index
'
Public Function mDetermineDayUpdate(ilDayCompleteInx As Integer) As Integer
Dim slStr As String
Dim ilDayIndex As Integer
Dim llDate As Long
Dim ilLoop As Integer

    If tmVef.sType <> "G" Then
        slStr = edcDate.Text                    'date selected in calendar
        ilDayIndex = gWeekDayStr(slStr)
        llDate = gDateValue(slStr)

        'calculate start of week from date selected
        Do While ilDayIndex <> 0
            llDate = llDate - 1
            slStr = Format$(llDate, "m/d/yy")
            ilDayIndex = gWeekDayStr(slStr)
        Loop
        'llDate = start of the week selected
        llDate = llDate + ilDayCompleteInx           'calculate actual date to set as complete
        ilDayIndex = -1
        For ilLoop = 0 To UBound(tgDates) - 1
            If llDate = tgDates(ilLoop).lDate Then
                ilDayIndex = ilLoop  'reestablish day of week to set
            End If
        Next ilLoop
        mDetermineDayUpdate = ilDayIndex
    Else
        mDetermineDayUpdate = imDateSelectedIndex
    End If
End Function

Private Function mBlockDay(llLock1 As Long, llLock2 As Long) As Integer
    Dim slUserName As String
    Dim ilRet As Integer

    'MAI/Sirius: Added 10/4 Added Record Lock
    mUnblockDay
    lmLock1RecCode = gCreateLockRec(hmRlf, "S", "P", llLock1, False, slUserName)
    If lmLock1RecCode = 0 Then
        ilRet = MsgBox("Unable to perform requested task as " & slUserName & " is working on day", vbOKOnly + vbInformation, "Block")
        mBlockDay = False
        Exit Function
    End If
    If llLock2 > 0 Then
        lmLock2RecCode = gCreateLockRec(hmRlf, "S", "P", llLock2, False, slUserName)
        If lmLock2RecCode = 0 Then
            ilRet = MsgBox("Unable to perform requested task as " & slUserName & " is working on day", vbOKOnly + vbInformation, "Block")
            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLock1RecCode)
            mBlockDay = False
            Exit Function
        End If
    Else
        lmLock2RecCode = -1
    End If
    mBlockDay = True
End Function

Private Sub mUnblockDay()
    Dim ilRet As Integer
    If lmLock1RecCode > 0 Then
        ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLock1RecCode)
    End If
    If lmLock2RecCode > 0 Then
        ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLock2RecCode)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTeamPop                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Tema list box         *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mTeamPop()
'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gObtainMnfForType("Z", smTeamTag, tmTeam())
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintLnTitle                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Header Titles            *
'*                                                     *
'*******************************************************
Private Sub mPaintPostTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcPosting.ForeColor
    slFontName = pbcPosting.FontName
    flFontSize = pbcPosting.FontSize
    ilFillStyle = pbcPosting.FillStyle
    llFillColor = pbcPosting.FillColor
    pbcPosting.ForeColor = BLUE
    pbcPosting.FontBold = False
    pbcPosting.FontSize = 7
    pbcPosting.FontName = "Arial"
    pbcPosting.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmCtrls(DATEINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcPosting.Line (tmCtrls(DATEINDEX).fBoxX - 15, 15)-Step(tmCtrls(DATEINDEX).fBoxW + 15, tmCtrls(DATEINDEX).fBoxY - 30), BLUE, B
    ''5/20/11
    '6/16/11
    'pbcPosting.Line (tmCtrls(DATEINDEX).fBoxX, 30)-Step(tmCtrls(DATEINDEX).fBoxW - 15, tmCtrls(DATEINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(DATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPosting.Print "Date"
    pbcPosting.Line (tmCtrls(TIMEINDEX).fBoxX - 15, 15)-Step(tmCtrls(TIMEINDEX).fBoxW + 15, tmCtrls(TIMEINDEX).fBoxY - 30), BLUE, B
    pbcPosting.CurrentX = tmCtrls(TIMEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPosting.Print "Time"
    pbcPosting.Line (tmCtrls(LENINDEX).fBoxX - 15, 15)-Step(tmCtrls(LENINDEX).fBoxW + 15, tmCtrls(LENINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(LENINDEX).fBoxX, 30)-Step(tmCtrls(LENINDEX).fBoxW - 15, tmCtrls(LENINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(LENINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPosting.Print "Len"
    pbcPosting.Line (tmCtrls(ADVTINDEX).fBoxX - 15, 15)-Step(tmCtrls(ADVTINDEX).fBoxW + 15, tmCtrls(ADVTINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(ADVTINDEX).fBoxX, 30)-Step(tmCtrls(ADVTINDEX).fBoxW - 15, tmCtrls(ADVTINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(ADVTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15
    pbcPosting.Print "Advertiser/Product"
    pbcPosting.Line (tmCtrls(TZONEINDEX).fBoxX - 15, 15)-Step(tmCtrls(TZONEINDEX).fBoxW + 15, tmCtrls(TZONEINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(TZONEINDEX).fBoxX, 30)-Step(tmCtrls(TZONEINDEX).fBoxW - 15, tmCtrls(TZONEINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(TZONEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15
    pbcPosting.Print "T"
    pbcPosting.CurrentX = tmCtrls(TZONEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = ilHalfY + 15
    pbcPosting.Print "Z"
    pbcPosting.Line (tmCtrls(COPYINDEX).fBoxX - 15, 15)-Step(tmCtrls(COPYINDEX).fBoxW + 15, tmCtrls(COPYINDEX).fBoxY - 30), BLUE, B
    If tgSpf.sUseCartNo <> "N" Then
        pbcPosting.CurrentX = tmCtrls(COPYINDEX).fBoxX + 15  'fgBoxInsetX
        pbcPosting.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcPosting.Print "Copy"
    End If
    pbcPosting.Line (tmCtrls(ISCIINDEX).fBoxX - 15, 15)-Step(tmCtrls(ISCIINDEX).fBoxW + 15, tmCtrls(ISCIINDEX).fBoxY - 30), BLUE, B
    pbcPosting.CurrentX = tmCtrls(ISCIINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPosting.Print "ISCI Code"
    pbcPosting.Line (tmCtrls(CNTRINDEX).fBoxX - 15, 15)-Step(tmCtrls(CNTRINDEX).fBoxW + 15, tmCtrls(CNTRINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(CNTRINDEX).fBoxX, 30)-Step(tmCtrls(CNTRINDEX).fBoxW - 15, tmCtrls(CNTRINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(CNTRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15
    pbcPosting.Print "Contract #"
    pbcPosting.Line (tmCtrls(LINEINDEX).fBoxX - 15, 15)-Step(tmCtrls(LINEINDEX).fBoxW + 15, tmCtrls(LINEINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(LINEINDEX).fBoxX, 30)-Step(tmCtrls(LINEINDEX).fBoxW - 15, tmCtrls(LINEINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(LINEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15
    pbcPosting.Print "Line"
    pbcPosting.Line (tmCtrls(TYPEINDEX).fBoxX - 15, 15)-Step(tmCtrls(TYPEINDEX).fBoxW + 15, tmCtrls(TYPEINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(TYPEINDEX).fBoxX, 30)-Step(tmCtrls(TYPEINDEX).fBoxW - 15, tmCtrls(TYPEINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(TYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15
    pbcPosting.Print "T"
    pbcPosting.Line (tmCtrls(PRICEINDEX).fBoxX - 15, 15)-Step(tmCtrls(PRICEINDEX).fBoxW + 15, tmCtrls(PRICEINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(PRICEINDEX).fBoxX, 30)-Step(tmCtrls(PRICEINDEX).fBoxW - 15, tmCtrls(PRICEINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(PRICEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15
    pbcPosting.Print "Spot Price"
    pbcPosting.Line (tmCtrls(MGOODINDEX).fBoxX - 15, 15)-Step(tmCtrls(MGOODINDEX).fBoxW + 15, tmCtrls(MGOODINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(MGOODINDEX).fBoxX, 30)-Step(tmCtrls(MGOODINDEX).fBoxW - 15, tmCtrls(MGOODINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(MGOODINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15
    pbcPosting.Print "M"
    pbcPosting.CurrentX = tmCtrls(MGOODINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = ilHalfY + 15
    pbcPosting.Print "G"
    pbcPosting.Line (tmCtrls(AUDINDEX).fBoxX - 15, 15)-Step(tmCtrls(AUDINDEX).fBoxW + 15, tmCtrls(AUDINDEX).fBoxY - 30), BLUE, B
    pbcPosting.Line (tmCtrls(AUDINDEX).fBoxX, 30)-Step(tmCtrls(AUDINDEX).fBoxW - 15, tmCtrls(AUDINDEX).fBoxY - 45), LIGHTYELLOW, BF
    pbcPosting.CurrentX = tmCtrls(AUDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = 15
    pbcPosting.Print "C"
    pbcPosting.CurrentX = tmCtrls(AUDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPosting.CurrentY = ilHalfY + 15
    pbcPosting.Print "h"

    ilLineCount = 0
    llTop = tmCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            pbcPosting.Line (tmCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmCtrls(ilLoop).fBoxW + 15, tmCtrls(ilLoop).fBoxH + 15), BLUE, B
            '6/16/11
            'If (ilLoop = DATEINDEX) Or (ilLoop = LENINDEX) Or (ilLoop = ADVTINDEX) Or (ilLoop = TZONEINDEX) Or (ilLoop >= CNTRINDEX) Then
            If (ilLoop = LENINDEX) Or (ilLoop = ADVTINDEX) Or (ilLoop = TZONEINDEX) Or (ilLoop >= CNTRINDEX) Then
                pbcPosting.Line (tmCtrls(ilLoop).fBoxX, llTop + 15)-Step(tmCtrls(ilLoop).fBoxW - 15, tmCtrls(ilLoop).fBoxH - 30), LIGHTYELLOW, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmCtrls(1).fBoxH + 15
    Loop While llTop + tmCtrls(1).fBoxH < pbcPosting.height
    vbcPosting.LargeChange = ilLineCount - 1
    pbcPosting.FontSize = flFontSize
    pbcPosting.FontName = slFontName
    pbcPosting.FontSize = flFontSize
    pbcPosting.ForeColor = llColor
    pbcPosting.FontBold = True
End Sub

Private Sub mPaintMissedTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer

    If imFirstActivate Then
        Exit Sub
    End If
    llColor = pbcMissed.ForeColor
    slFontName = pbcMissed.FontName
    flFontSize = pbcMissed.FontSize
    ilFillStyle = pbcMissed.FillStyle
    llFillColor = pbcMissed.FillColor
    pbcMissed.ForeColor = BLUE
    pbcMissed.FontBold = False
    pbcMissed.FontSize = 7
    pbcMissed.FontName = "Arial"
    pbcMissed.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    pbcMissed.Line (tmMdCtrls(MDADVTINDEX).fBoxX - 15, 15)-Step(tmMdCtrls(MDADVTINDEX).fBoxW + 15, tmMdCtrls(MDADVTINDEX).fBoxH + 15), BLUE, B
    pbcMissed.Line (tmMdCtrls(MDADVTINDEX).fBoxX, 30)-Step(tmMdCtrls(MDADVTINDEX).fBoxW - 15, tmMdCtrls(MDADVTINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcMissed.CurrentX = tmMdCtrls(MDADVTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMissed.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcMissed.Print "Advertiser"
    pbcMissed.Line (tmMdCtrls(MDCNTRINDEX).fBoxX - 15, 15)-Step(tmMdCtrls(MDCNTRINDEX).fBoxW + 15, tmMdCtrls(MDCNTRINDEX).fBoxH + 15), BLUE, B
    pbcMissed.Line (tmMdCtrls(MDCNTRINDEX).fBoxX, 30)-Step(tmMdCtrls(MDCNTRINDEX).fBoxW - 15, tmMdCtrls(MDCNTRINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcMissed.CurrentX = tmMdCtrls(MDCNTRINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMissed.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcMissed.Print "Contract #"
    pbcMissed.Line (tmMdCtrls(MDVEHINDEX).fBoxX - 15, 15)-Step(tmMdCtrls(MDVEHINDEX).fBoxW + 15, tmMdCtrls(MDVEHINDEX).fBoxH + 15), BLUE, B
    pbcMissed.Line (tmMdCtrls(MDVEHINDEX).fBoxX, 30)-Step(tmMdCtrls(MDVEHINDEX).fBoxW - 15, tmMdCtrls(MDVEHINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcMissed.CurrentX = tmMdCtrls(MDVEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMissed.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    If cbcVehicle.ListIndex > 0 Then
        pbcMissed.Print "Vehicle"
    Else
        pbcMissed.Print "Product"
    End If
    pbcMissed.Line (tmMdCtrls(MDLENINDEX).fBoxX - 15, 15)-Step(tmMdCtrls(MDLENINDEX).fBoxW + 15, tmMdCtrls(MDLENINDEX).fBoxH + 15), BLUE, B
    pbcMissed.Line (tmMdCtrls(MDLENINDEX).fBoxX, 30)-Step(tmMdCtrls(MDLENINDEX).fBoxW - 15, tmMdCtrls(MDLENINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcMissed.CurrentX = tmMdCtrls(MDLENINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMissed.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcMissed.Print "Len"
    pbcMissed.Line (tmMdCtrls(MDWKMISSINDEX).fBoxX - 15, 15)-Step(tmMdCtrls(MDWKMISSINDEX).fBoxW + 15, tmMdCtrls(MDWKMISSINDEX).fBoxH + 15), BLUE, B
    pbcMissed.Line (tmMdCtrls(MDWKMISSINDEX).fBoxX, 30)-Step(tmMdCtrls(MDWKMISSINDEX).fBoxW - 15, tmMdCtrls(MDWKMISSINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcMissed.CurrentX = tmMdCtrls(MDWKMISSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMissed.CurrentY = 15
    pbcMissed.Print "Week Date"
    pbcMissed.Line (tmMdCtrls(MDENDDATEINDEX).fBoxX - 15, 15)-Step(tmMdCtrls(MDENDDATEINDEX).fBoxW + 15, tmMdCtrls(MDENDDATEINDEX).fBoxH + 15), BLUE, B
    pbcMissed.Line (tmMdCtrls(MDENDDATEINDEX).fBoxX, 30)-Step(tmMdCtrls(MDENDDATEINDEX).fBoxW - 15, tmMdCtrls(MDENDDATEINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcMissed.CurrentX = tmMdCtrls(MDENDDATEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMissed.CurrentY = 15
    pbcMissed.Print "End Date"
    pbcMissed.Line (tmMdCtrls(MDDPINDEX).fBoxX - 15, 15)-Step(tmMdCtrls(MDDPINDEX).fBoxW + 15, tmMdCtrls(MDDPINDEX).fBoxH + 15), BLUE, B
    pbcMissed.Line (tmMdCtrls(MDDPINDEX).fBoxX, 30)-Step(tmMdCtrls(MDDPINDEX).fBoxW - 15, tmMdCtrls(MDDPINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcMissed.CurrentX = tmMdCtrls(MDDPINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMissed.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcMissed.Print "Daypart"
    pbcMissed.Line (tmMdCtrls(MDNOSPOTSINDEX).fBoxX - 15, 15)-Step(tmMdCtrls(MDNOSPOTSINDEX).fBoxW + 15, tmMdCtrls(MDNOSPOTSINDEX).fBoxH + 15), BLUE, B
    pbcMissed.Line (tmMdCtrls(MDNOSPOTSINDEX).fBoxX, 30)-Step(tmMdCtrls(MDNOSPOTSINDEX).fBoxW - 15, tmMdCtrls(MDNOSPOTSINDEX).fBoxH - 15), LIGHTYELLOW, BF
    pbcMissed.CurrentX = tmMdCtrls(MDNOSPOTSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcMissed.CurrentY = 15
    pbcMissed.Print "Spots"
    ilLineCount = 0
    llTop = tmMdCtrls(1).fBoxY
    Do
        For ilLoop = imLBMdCtrls To MDNOSPOTSINDEX Step 1
            pbcMissed.Line (tmMdCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmMdCtrls(ilLoop).fBoxW + 15, tmMdCtrls(ilLoop).fBoxH + 15), BLUE, B
            pbcMissed.Line (tmMdCtrls(ilLoop).fBoxX, llTop + 15)-Step(tmMdCtrls(ilLoop).fBoxW - 15, tmMdCtrls(ilLoop).fBoxH - 30), LIGHTYELLOW, BF
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmMdCtrls(1).fBoxH + 15
    Loop While llTop + tmMdCtrls(1).fBoxH < pbcMissed.height
    'vbcPosting.LargeChange = ilLineCount - 1
    pbcMissed.FontSize = flFontSize
    pbcMissed.FontName = slFontName
    pbcMissed.FontSize = flFontSize
    pbcMissed.ForeColor = llColor
    pbcMissed.FontBold = True
End Sub

Private Function mGetCopy(tlSdf As SDF, ilSaveIndex As Integer) As Integer
    Dim ilCifFound As Integer
    Dim ilRet As Integer
    Dim slHoldNames As String
    
    tgSave(ilSaveIndex).sCopy = ""
    tgSave(ilSaveIndex).sISCI = ""
    tgSave(ilSaveIndex).sCopyProduct = ""
    tgSave(ilSaveIndex).sTZone = ""
    ilCifFound = False
    If tmSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mGetCopyErr
        gBtrvErrorMsg ilRet, "mGetCopy (btrGetEqual:CIF, Single)", PostLog
        On Error GoTo 0
        ilCifFound = True
    ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
        tgSave(ilSaveIndex).sTZone = "4"
        ' Read TZF using lCopyCode from SDF
        tmTzfSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo mGetCopyErr
        gBtrvErrorMsg ilRet, "mGetCopy (btrGetEqual:TZF)", PostLog
        On Error GoTo 0
        ' Look for the first positive lZone value
        For imIndex = 1 To 6 Step 1
            If (tmTzf.lCifZone(imIndex - 1) > 0) And (StrComp(tmTzf.sZone(imIndex - 1), "Oth", 1) = 0) Then ' Process just the first positive Zone
                ' Read CIF using lCopyCode from SDF
                tmCifSrchKey.lCode = tmTzf.lCifZone(imIndex - 1)
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                On Error GoTo mGetCopyErr
                gBtrvErrorMsg ilRet, "mGetCopy (btrGetEqual:CIF, Time zone)", PostLog
                On Error GoTo 0
                ilCifFound = True
                Exit For
            End If
        Next imIndex
        If Not ilCifFound Then
            For imIndex = 1 To 6 Step 1
                If tmTzf.lCifZone(imIndex - 1) > 0 Then ' Process just the first positive Zone
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(imIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo mGetCopyErr
                    gBtrvErrorMsg ilRet, "mGetCopy (btrGetEqual:CIF, Time zone)", PostLog
                    On Error GoTo 0
                    ilCifFound = True
                    Exit For
                End If
            Next imIndex
        End If
    End If
    If ilCifFound Then
        ' Read CPF using lCpfCode from CIF
        If tmCif.lcpfCode > 0 Then
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mGetCopyErr
            gBtrvErrorMsg ilRet, "mGetCopy (btrGetEqual:CPF)", PostLog
            On Error GoTo 0
            'smSave(SAVISCIINDEX, ilSaveIndex) = Trim$(tmCpf.sISCI)  ' ISCI Code
            tgSave(ilSaveIndex).sISCI = Trim$(tmCpf.sISCI)
        Else
            tmCpf.sISCI = ""
            'smSave(SAVISCIINDEX, ilSaveIndex) = Trim$(tmCpf.sISCI)  ' ISCI Code
            tgSave(ilSaveIndex).sISCI = Trim$(tmCpf.sISCI)
            tmCpf.sName = ""
        End If
        ' Concatinate Copy from Media Code, Inv. Name & Cut#
        ' First read MCF
        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
            tmMcfSrchKey.iCode = tmCif.iMcfCode
            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mGetCopyErr
            gBtrvErrorMsg ilRet, "mGetCopy (btrGetEqual:MCF)", PostLog
            On Error GoTo 0
            ' Media Code is tmMcf.sName
            slHoldNames = Trim$(tmMcf.sName) & Trim$(tmCif.sName)
            If (Len(Trim$(tmCif.sCut)) <> 0) Then
                slHoldNames = slHoldNames & "-" & tmCif.sCut
            End If
            tgSave(ilSaveIndex).sCopy = slHoldNames
        Else
            tgSave(ilSaveIndex).sCopy = ""
        End If
        If Trim$(tmCpf.sName) <> "" Then
            tgSave(ilSaveIndex).sCopyProduct = Trim$(tmCpf.sName)
        End If
        If Trim$(tmCpf.sISCI) <> "" Then
            'imPostSpotInfo(2, ilSaveIndex) = False
            tgSave(ilSaveIndex).iISCI = False
        Else
            'imPostSpotInfo(2, ilSaveIndex) = True
            tgSave(ilSaveIndex).iISCI = True
        End If
    End If
    mGetCopy = True
    Exit Function
mGetCopyErr:
    On Error GoTo 0
    mGetCopy = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailReset                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Reduce units/sec or reset to   *
'*                      original values                *
'*                                                     *
'*******************************************************
Private Function mAvailReset(ilAvailIndex) As Integer
'   ilAvailIndex(I)- location of avail within Ssf (use mFindAvail)
'   ilRet(O)- True=Avail has been reset successfully; False=Reset failed
    Dim ilAvailUnits As Integer
    Dim ilAvailSec As Integer
    Dim ilOrigAvailUnits As Integer
    Dim ilOrigAvailSec As Integer
    Dim ilUnitsSold As Integer
    Dim ilSecSold As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSpotIndex As Integer
    Dim ilNewUnit As Integer
    Dim ilNewSec As Integer
    Dim ilRet As Integer
   LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex)
    ilAvailUnits = tmAvail.iAvInfo And &H1F
    ilAvailSec = tmAvail.iLen
    ilOrigAvailUnits = tmAvail.iOrigUnit And &H1F
    ilOrigAvailSec = tmAvail.iOrigLen
    If (ilOrigAvailUnits = 0) And (ilOrigAvailSec = 0) Then
        mAvailReset = True
        Exit Function
    End If
    For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
        If (tmSpot.iRecType And &HF) >= 10 Then
            ilSpotLen = tmSpot.iPosLen And &HFFF
            If (tgVpf(imVpfIndex).sSSellOut = "T") Then
                ilSpotUnits = ilSpotLen \ 30
                If ilSpotUnits <= 0 Then
                    ilSpotUnits = 1
                End If
                ilSpotLen = 0
            Else
                ilSpotUnits = 1
            End If
            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                ilUnitsSold = ilUnitsSold + ilSpotUnits
                ilSecSold = ilSecSold + ilSpotLen
            End If
        End If
    Next ilSpotIndex
    ilNewUnit = 0
    ilNewSec = 0
    If (tgVpf(imVpfIndex).sSSellOut = "M") Then
        If (ilSecSold < ilAvailSec) Or (ilUnitsSold < ilAvailUnits) Then
            ilNewSec = ilSecSold
            ilNewUnit = ilUnitsSold
        Else
            mAvailReset = True
            Exit Function
        End If
    Else
        If (ilSecSold < ilAvailSec) Or (ilUnitsSold < ilAvailUnits) Then
            ilNewSec = ilSecSold
            ilNewUnit = ilUnitsSold
        Else
            mAvailReset = True
            Exit Function
        End If
    End If
    Do
        imSsfRecLen = Len(tmSsf(imSelectedDay))
        ilRet = gSSFGetDirect(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, lmSsfRecPos(imSelectedDay), INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mAvailReset = False
            Exit Function
        End If
        ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf(imSelectedDay))
        If ilRet <> BTRV_ERR_NONE Then
            mAvailReset = False
            Exit Function
        End If
        If ilNewUnit < ilOrigAvailUnits Then
            tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilOrigAvailUnits
        Else
            tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilNewUnit
        End If
        If ilNewSec < ilOrigAvailSec Then
            tmAvail.iLen = ilOrigAvailSec
        Else
            tmAvail.iLen = ilNewSec
        End If
        tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
        imSsfRecLen = igSSFBaseLen + tmSsf(imSelectedDay).iCount * Len(tmProg)
        ilRet = gSSFUpdate(hmSsf, tmSsf(imSelectedDay), imSsfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        mAvailReset = False
        Exit Function
    End If
    mAvailReset = True
    Exit Function
End Function


Private Sub mFinalInvoiceRunning(llDate As Long)
    Dim ilRet As Integer
    Dim slUserName As String
    Dim llStdStart As Long
    Dim llLockRecCode As Long
    
    imUpdateAllowed = imSvUpdateAllowed
    llStdStart = gDateValue(gObtainStartStd(Format(llDate, "m/d/yy")))
    llLockRecCode = gCreateLockRec(hmRlf, "I", "F", llStdStart, False, slUserName)
    If llLockRecCode = 0 Then
        ilRet = MsgBox("Final Invoices being currently run by " & slUserName & ", Spot Posting Disallowed", vbOKOnly + vbExclamation, "Message")
        imUpdateAllowed = False
    Else
        ilRet = gDeleteLockRec_ByRlfCode(hmRlf, llLockRecCode)
        'Create Lock records by User
        
    End If
End Sub

Private Function mAddIihf(ilVefCode As Integer, llChfCode As Long, slPostDate As String) As Integer
    Dim ilRet As Integer
    Dim ilVff As Integer
    Dim llInvStartdate As Long
    
    If Not cmcImport.Visible Then
        mAddIihf = True
        Exit Function
    End If
    If smPostLogSource <> "S" Then
        mAddIihf = True
        Exit Function
    End If
    llInvStartdate = gDateValue(gObtainStartStd(slPostDate))
    tmIihfSrchKey2.lChfCode = llChfCode
    tmIihfSrchKey2.iVefCode = ilVefCode
    gPackDateLong llInvStartdate, tmIihfSrchKey2.iInvStartDate(0), tmIihfSrchKey2.iInvStartDate(1)
    ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        mAddIihf = True
        Exit Function
    End If
    tmIihf.lCode = 0
    tmIihf.iVefCode = ilVefCode
    tmIihf.lChfCode = llChfCode
    gPackDateLong llInvStartdate, tmIihf.iInvStartDate(0), tmIihf.iInvStartDate(1)
    tmIihf.sFileName = "Post Log"
    tmIihf.sStnEstimateNo = ""
    tmIihf.sStnInvoiceNo = ""
    tmIihf.sStnContractNo = ""
    tmIihf.lAmfCode = 0
    tmIihf.sSourceForm = "P"
    tmIihf.sUnused = ""
    ilRet = btrInsert(hmIihf, tmIihf, imIihfRecLen, INDEXKEY0)
    If ilRet = BTRV_ERR_NONE Then
        mAddIihf = True
    Else
        mAddIihf = False
    End If
End Function
