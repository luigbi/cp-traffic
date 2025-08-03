VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExptGen 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   5610
   ClientTop       =   4275
   ClientWidth     =   9135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   9135
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2400
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   32
      Top             =   1170
      Visible         =   0   'False
      Width           =   1995
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
         Left            =   1635
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
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
         Left            =   45
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
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
         Left            =   45
         Picture         =   "Exptgen.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   34
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   25
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcCalendarTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   4695
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   40
      Top             =   1170
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendarTo 
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
         Left            =   45
         Picture         =   "Exptgen.frx":2E1A
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDateTo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   35
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.CommandButton cmcCalDnTo 
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
         Left            =   45
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalUpTo 
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
         Left            =   1635
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalNameTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   44
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.Timer tmcWideOrbit00 
      Interval        =   100
      Left            =   600
      Top             =   5025
   End
   Begin VB.ListBox lbcMsg 
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
      Height          =   2760
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1290
      Width           =   5235
   End
   Begin VB.CheckBox ckcWideOrbit00 
      Caption         =   "Append ""00"" to file name"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2790
      TabIndex        =   21
      ToolTipText     =   "This option is used only for Wide Orbit automation and will append ""00"" to the end of the file name."
      Top             =   5415
      Width           =   2895
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "RadioMan"
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
      Height          =   210
      Index           =   21
      Left            =   7065
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Station Playlist"
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
      Height          =   210
      Index           =   20
      Left            =   7440
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Zetta"
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
      Height          =   210
      Index           =   19
      Left            =   7320
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Scott V5"
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
      Height          =   210
      Index           =   18
      Left            =   7680
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Linkup"
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
      Height          =   210
      Index           =   17
      Left            =   7710
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Jelli"
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
      Height          =   210
      Index           =   16
      Left            =   7695
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Wide Orbit"
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
      Height          =   210
      Index           =   15
      Left            =   7065
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4665
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Audio Vault Air"
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
      Height          =   210
      Index           =   14
      Left            =   7065
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Audio Vault RPS"
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
      Height          =   210
      Index           =   13
      Left            =   7065
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Rivendell"
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
      Height          =   210
      Index           =   12
      Left            =   7050
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4410
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "RCS-5"
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
      Height          =   210
      Index           =   11
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CheckBox ckcAllItems 
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7125
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   405
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.ListBox lbcGroupItems 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "Exptgen.frx":5C34
      Left            =   7140
      List            =   "Exptgen.frx":5C36
      MultiSelect     =   2  'Extended
      TabIndex        =   53
      Top             =   645
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.TextBox edcSet1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5475
      TabIndex        =   52
      Text            =   "Vehicle Group"
      Top             =   405
      Width           =   1305
   End
   Begin VB.ComboBox cbcSet1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   5475
      TabIndex        =   51
      Top             =   645
      Width           =   1590
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Simian"
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
      Height          =   210
      Index           =   10
      Left            =   5670
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame frcZone 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   75
      TabIndex        =   45
      Top             =   630
      Width           =   5835
      Begin VB.OptionButton rbcZone 
         Caption         =   "Pacific"
         Height          =   240
         Index           =   3
         Left            =   4230
         TabIndex        =   50
         Top             =   30
         Width           =   1035
      End
      Begin VB.OptionButton rbcZone 
         Caption         =   "Mountain"
         Height          =   240
         Index           =   2
         Left            =   2970
         TabIndex        =   49
         Top             =   30
         Width           =   1245
      End
      Begin VB.OptionButton rbcZone 
         Caption         =   "Central"
         Height          =   240
         Index           =   1
         Left            =   1905
         TabIndex        =   48
         Top             =   30
         Width           =   1110
      End
      Begin VB.OptionButton rbcZone 
         Caption         =   "Eastern"
         Height          =   240
         Index           =   0
         Left            =   780
         TabIndex        =   47
         Top             =   30
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.Label lacZone 
         Caption         =   "Zone"
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   45
         Width           =   555
      End
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "WireReady"
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
      Height          =   210
      Index           =   9
      Left            =   6600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Audio  Vault Sat"
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
      Height          =   210
      Index           =   8
      Left            =   4665
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "iMediaTouch"
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
      Height          =   210
      Index           =   7
      Left            =   2400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcEndDate 
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
      Left            =   4680
      Picture         =   "Exptgen.frx":5C38
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   390
      Width           =   195
   End
   Begin VB.TextBox edcEndDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   17
      Top             =   390
      Width           =   930
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4140
      Width           =   1410
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "RCS"
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
      Height          =   210
      Index           =   4
      Left            =   4425
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Prophet NexGen"
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
      Height          =   210
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Dalet"
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
      Height          =   210
      Index           =   0
      Left            =   780
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Drake"
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
      Height          =   210
      Index           =   1
      Left            =   1140
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Scott"
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
      Height          =   210
      Index           =   5
      Left            =   4695
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Prophet Wizard"
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
      Height          =   210
      Index           =   3
      Left            =   1920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.OptionButton rbcAutoType 
      Caption         =   "Prophet MediaStar"
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
      Height          =   210
      Index           =   6
      Left            =   3165
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2370
      Top             =   4950
   End
   Begin VB.CommandButton cmcStartDate 
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
      Left            =   2685
      Picture         =   "Exptgen.frx":5D32
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   390
      Width           =   195
   End
   Begin VB.TextBox edcStartDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1755
      MaxLength       =   10
      TabIndex        =   14
      Top             =   390
      Width           =   930
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   1260
      Top             =   4875
      _ExtentX        =   2646
      _ExtentY        =   1323
      _Version        =   393216
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   270
      Left            =   105
      ScaleHeight     =   270
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   90
      Width           =   2175
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2790
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcVehicle 
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
      Height          =   2760
      ItemData        =   "Exptgen.frx":5E2C
      Left            =   120
      List            =   "Exptgen.frx":5E2E
      MultiSelect     =   2  'Extended
      TabIndex        =   19
      Top             =   1290
      Width           =   3375
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
      Enabled         =   0   'False
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
      Left            =   2790
      TabIndex        =   22
      Top             =   5040
      Width           =   1050
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
      Left            =   5145
      TabIndex        =   23
      Top             =   5040
      Width           =   1050
   End
   Begin VB.Frame frcSplit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      TabIndex        =   55
      Top             =   900
      Width           =   5865
      Begin VB.OptionButton rbcSplit 
         Caption         =   "Primary Only"
         Height          =   240
         Index           =   0
         Left            =   2385
         TabIndex        =   57
         Top             =   45
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton rbcSplit 
         Caption         =   "Secondary Only"
         Height          =   240
         Index           =   1
         Left            =   3900
         TabIndex        =   56
         Top             =   45
         Width           =   1800
      End
      Begin VB.Label lacSplit 
         Caption         =   "For Split Networks- Export"
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   45
         Width           =   2430
      End
   End
   Begin VB.Label lacEndDate 
      Appearance      =   0  'Flat
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3480
      TabIndex        =   16
      Top             =   375
      Width           =   465
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   195
      TabIndex        =   39
      Top             =   4620
      Width           =   8730
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   5025
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   165
      TabIndex        =   37
      Top             =   4380
      Width           =   8730
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Export Date- From"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      TabIndex        =   13
      Top             =   375
      Width           =   1785
   End
   Begin VB.Label lacErrors 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5055
      TabIndex        =   31
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5175
      TabIndex        =   29
      Top             =   1395
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3225
      TabIndex        =   27
      Top             =   1395
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ExptGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExptGen.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export feed (for Dalet, Scott, Drake & Prophet) input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim smScreenCaption As String
Dim hmTo As Integer   'From file hanle
Dim hmMsg As Integer   'From file hanle
Dim lmNowDate As Long   'Todays date
'Required by gMakeSsf
Dim tmSsf As SSF                'SSF record image
Dim hmSsf As Integer
Dim hmCTSsf As Integer
Dim tmCTSsf As SSF               'Ssf for conflict test
'Dim tmSsfOld As SSF
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmAvailTest As AVAILSS
Dim tmOpenAvail As AVAILSS
Dim tmCloseAvail As AVAILSS
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ANF key record image
Dim imAdfRecLen As Integer  'ANF record length

'Arf table for the export location
Dim hmArf As Integer
Dim tmArf As ARF
Dim tmArfSrchKey As INTKEY0 'ARF key record image
Dim imArfRecLen As Integer  'ARF record length

'Agf table for the export location, 6-21-12
Dim hmAgf As Integer
Dim tmAgf As AGF
Dim tmAgfSrchKey As INTKEY0 'AGF key record image
Dim imAgfRecLen As Integer  'AGF record length

'Multi- name
Dim hmMnf As Integer
Dim tmMnf As MNF
Dim tmMnfSrchKey As INTKEY0 'MNF key record image
Dim imMnfRecLen As Integer  'MNF record length
'Copy/Product
Dim hmCpf As Integer
Dim tmCpf As CPF
Dim imCpfRecLen As Integer  'CPF record length
'Avail name
Dim hmAnf As Integer
Dim tmAnf As ANF
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length
'Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
'Contract Line record information
Dim hmClf As Integer        'Contract Line file handle
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
'Short Title Vehicle Table record information
Dim hmVsf As Integer        'Short Title Vehicle Table file handle
Dim imVsfRecLen As Integer  'VSF record length
Dim tmVsf As VSF            'VSF record image
'Media code record information
Dim hmMcf As Integer        'Contract line file handle
Dim imMcfRecLen As Integer  'MCF record length
Dim tmMcf As MCF            'MCF record image
'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image
Dim hmTzf As Integer        'Time zone Copy file handle
Dim imTzfRecLen As Integer  'TZF record length
Dim tmTzf As TZF            'TZF record image
Dim imCopyMissing As Integer
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmSVef As VEF           'Selling Vehicle
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
Dim smVehName As String
'Vehicle Options
Dim tmVpf As VPF                'VPF record image
Dim tmVpfSrchKey As VPFKEY0     'VPF key 0 image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmVpf As Integer            'Vehicle preference file handle
'Vehicle linkage record information
Dim hmVlf As Integer        'Vehicle linkage file handle
Dim tmVlfSrchKey1 As VLFKEY1 'VLF key record image
Dim imVlfRecLen As Integer  'VLF record length
Dim tmVlf As VLF            'VLF record image
'Delivery file (DLF)
Dim hmDlf As Integer        'Delivery link file
Dim imDlfRecLen As Integer  'DLF record length
Dim tmDlfSrchKey As DLFKEY0 'DLF key record image
Dim tmDlf As DLF            'DLF record image
'Air Copy
Dim hmRsf As Integer        'Region Copy file handle
Dim imRsfRecLen As Integer  'RSF record length
Dim tmRsf As RSF            'RSF record image
Dim tmRsfSrchKey1 As LONGKEY0 'RSF key record image

'Audio Vault To X-Digital Indicator ID
Dim hmAxf As Integer
Dim tmAxf As AXF
Dim tmAxfSrchKey1 As INTKEY0 'ANF key record image
Dim imAxfRecLen As Integer  'ANF record length

'Copy Rotation
Dim hmCrf As Integer        'Copy Rotation file handle
'Copy Vehicle
Dim hmCvf As Integer        'Copy Vehicle file handle
'Spot record
Dim tmSdf As SDF
Dim hmSdf As Integer
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey3 As LONGKEY0
Dim tmBBSdfInfo() As BBSDFINFO
'Game Info
Dim hmGsf As Integer
Dim tmGsf As GSF        'GSF record image
Dim tmGsfSrchKey As LONGKEY0
Dim tmGsfSrchKey3 As GSFKEY3    'GSF key record image
Dim tmGsfSrchKey4 As GSFKEY4    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

Dim tmTeam() As MNF
Dim smTeamTag As String

'Event Type
Dim hmEtf As Integer        'event type
Dim tmEtf As ETF
Dim imEtfRecLen As Integer
Dim tmEtfSrchKey0 As INTKEY0

'CEF Type
Dim hmCef As Integer        'comment for other type
Dim tmCef As CEF
Dim imCefRecLen As Integer
Dim tmCefSrchKey0 As LONGKEY0

'LCF Calendar Log
Dim hmLcf As Integer        'Log Calendar handle
Dim tmLcf As LCF
Dim imLcfRecLen As Integer
Dim tmLcfSrchKey0 As LCFKEY0

'9/9/13: Handle Merge (Parent/Child)
Dim hmLvf As Integer

Dim tmLLC() As LLC  'Image
Dim smMcfPrefix As String
Dim slMcfCode As String
Dim smROS As String
Dim smPty As String
Dim smFixed As String
Dim smAdvtCode As String
Dim smCompCode As String
Dim smCifName As String
Dim smCreativeTitle As String
Dim smISCI As String
Dim imMediaCodeLen As Integer
Dim smMediaCodeSuppressSpot As String

Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imEvtType(0 To 14) As Integer
Dim tmSpotTimes() As SPOTTIMES
Dim smCurrentLines() As String
Dim tmExpRecImage() As EXPRECIMAGE
'Dim smNewLines() As String * 140        '4-7-08 chg from 118 to 140 for new prophet nextgen fields
Dim smNewLines() As String * 255        '7/21/09 chg from 140 to 255 for new rcs-5 fields

Dim imSetIndex As Integer           '11-13-08 vehicle group selected

Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)

Dim imSetAllItems As Integer
Dim imAllClickedItems As Integer
Dim imFoundspot As Integer  '1-5-05 flag to indicate if at least one spot found; if not, dont retain export file
Dim imUsingMediaCodeForAV As Integer   '10-11-06 using media codes with inventory

Dim smVehicleExportLoc As String    '2-12-07 location where export should be placed
'Dim imTempExportPath As String      '2-12-07, 6-21-12 should be name "sm" not "im"
Dim smTempExportPath As String      '2-12-07, 6-21-12 should be name "sm" not "im"

'9/9/13: Handle ProgCodeID = Merge
Dim tmProgTimeRange() As PROGTIMERANGE
Dim tmBreakByProg() As BREAKBYPROG

Dim imXMidNight As Integer
Dim lmLastEvtTime As Long
Dim imGameVehicle As Integer

'7/11/14: Lock avails between 12am-3pm
Dim smLockDate As String
Dim smLockStartTime As String
Dim smLockEndTime As String
Dim imLockVefCode As Integer
Dim lmLockStartTime As Long

' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0

Const AUTOTYPE_WIDEORBIT = 16
Const AUTOTYPE_JELLI = 17               '6-21-12
Const AUTOTYPE_ENCOESPN = 18
Const AUTOTYPE_SCOTT = 3                '5-30-13
Const AUTOTYPE_SCOTT_V5 = 19
Const AUTOTYPE_ZETTA = 20               '1-5-16
Const AUTOTYPE_STATIONPL = 21           '5-10-18
Const AUTOTYPE_RADIOMAN = 22            '5-1-20

Private Sub cbcSet1_Click()
    Dim ilLoop As Integer
    Dim ilRet As Integer
            ilLoop = cbcSet1.ListIndex
            imSetIndex = gFindVehGroupInx(ilLoop, tgVehicleSets1())

            If imSetIndex > 0 Then
                smVehGp5CodeTag = ""
                ilRet = gPopMnfPlusFieldsBox(ExptGen, lbcGroupItems, tgSOCodeCT(), smVehGp5CodeTag, "H" & Trim$(str$(imSetIndex)))
                If imSetIndex = 1 Then              'participants vehicle sets
                    lbcGroupItems.Visible = True
                    ckcAllItems.Caption = "All Participants"
                ElseIf imSetIndex = 2 Then          'subtotals vehicle sets
                    lbcGroupItems.Visible = True
                    ckcAllItems.Caption = "All Sub-totals"
                ElseIf imSetIndex = 3 Then          'market vehicle sets
                    lbcGroupItems.Visible = True
                    ckcAllItems.Caption = "All Markets"
                ElseIf imSetIndex = 4 Then          'format vehicle sets
                    lbcGroupItems.Visible = True
                    ckcAllItems.Caption = "All Formats"
                ElseIf imSetIndex = 5 Then          'research vehicle sets
                    lbcGroupItems.Visible = True
                    ckcAllItems.Caption = "All Research"
                ElseIf imSetIndex = 6 Then          'sub-company vehicle sets
                    lbcGroupItems.Visible = True
                    ckcAllItems.Caption = "All Sub-Companies"
                End If
                ckcAllItems.Visible = True
                ckcAllItems.Value = vbUnchecked
            Else
                lbcGroupItems.Visible = False
                ckcAllItems.Value = vbUnchecked   '9-12-02 False
                ckcAllItems.Visible = False
            End If
            mSetCommands
End Sub

Private Sub cbcSet1_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    If lbcVehicle.ListCount <= 0 Then
        Exit Sub
    End If
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub ckcAllItems_Click()
  'Code added because Value removed as parameter
    Dim Value As Integer
    If lbcGroupItems.ListCount <= 0 Then
        Exit Sub
    End If
    Value = False
    If ckcAllItems.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAllItems Then
        imAllClickedItems = True
        llRg = CLng(lbcGroupItems.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcGroupItems.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClickedItems = False
    End If
    mSetCommands
End Sub

Private Sub ckcAllItems_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub

Private Sub cmcCalDnTo_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendarTo_Paint
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
End Sub

Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub

Private Sub cmcCalUpTo_Click()
    plcCalendar.Visible = False
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendarTo_Paint
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
End Sub

Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub cmcEndDate_Click()
    plcCalendarTo.Visible = Not plcCalendarTo.Visible
    edcEndDate.SelStart = 0
    edcEndDate.SelLength = Len(edcEndDate.Text)
    edcEndDate.SetFocus
    mSetCommands
End Sub

Private Sub cmcEndDate_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub

Private Sub cmcExport_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoopOnVef                   ilVef                                                   *
'******************************************************************************************

    Dim slToFile As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStr As String
    Dim ilDays As Integer
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slLetter As String
    Dim slFileName As String
    Dim ilLine As Integer
    Dim slDateTime As String
    Dim slMissingCopyNames As String
    Dim ilAutomationType As Integer
    Dim ilRec As Integer
    Dim ilLoopSet As Integer
    Dim ilFoundGroup As Integer
    Dim llGsfCode As Long
'       ilautomationType                                           ilAutomationType
'         (as input)                                               (converted from inpu screen)
'        spfAutoType &                          rbc control
'        spfAutoType1     Automation            from input
'            &h1            DALET               rbc(0)              1
'            &H2            Prophet NexGen      rbc(2)              2
'            &h4            Scott               rbc(5)              3
'            &h8            Drake               rbc(1)              4
'            &h10           RCS-4               rbc(4)              5
'            &h20           Prophet Wizard      rbc(3)              6    2-5-03
'            &h40           Prophet MediaStar   rbc(6)              7    9-25-03
'            &h80           imediaTouch         rbc(7)              8    6-25-05
'            &h100          Audio Vault sat     rbc(8)              9    8-10-05
'            &h400          Wire Ready          rbc(9)              10
'            &h4000         Simian              rbc(10)             11   8-21-08
'            &h8000         RCS-5               rbc(11)             12
'                           Rivendell           rbc(12)             13
'                           AudioVault RPS      rbc(13)             14  11-15-10
'        spfAutoType3:
'                           AudioVault Air      rbc(14)             15
'                           Wide Orbit          rbc(15)             16   1-6-12     WIDEORBIT = &H2
'                           JELLI               rbc(16)             17   6-21-12    &H4
'                           ENCO-ESPN           rbc(17)             18   10/16/12   ENCOESPN = &H8
'                           SCOTT_V5            rbc(18)             19   8-16-13     Scott_V5 = &H10
'                           ZETTA               rbc(19)             20   1-5-16      Zetta = &H20
'                           Station Playlist    rbc(20)             21   5-11-18     Station Playlist = &H40
'                           RadidoMan           rbc(21)             22   5-1-20      RadioMan = &H80
    Dim slExt As String * 4         '5-3-01  .txt or .log
    Dim ilRecdLen As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim llTestDate As Long
    Dim slDayOfWeek As String * 21
    Dim ilVpfIndex As Integer
    Dim ilGameNo As Integer
    Dim ilGameVpfIndex As Integer
    Dim llGameDate As Long
    Dim llIndex As Long
    Dim ilTeam As Integer
    Dim ilVff As Integer
    Dim ilCRet As Integer
    Dim ilVef As Integer
    Dim tlVef As VEF    'Used to retain vehicle so that it can be reset after combine vehicle
    Dim fs As New FileSystemObject          'folders


    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    
    lacProcessing.Caption = ""
    lacMsg.Caption = ""
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If
    llStartDate = gDateValue(slStr)

    slStr = edcEndDate.Text
    If slStr = "" Then
        'assume same end date as start
        slStr = edcStartDate.Text
    End If
    If Not gValidDate(slStr) Then
        Beep
        edcEndDate.SetFocus
        Exit Sub
    End If
    llEndDate = gDateValue(slStr)
    'ilDays = 1  'Val(edcNoDays.Text)
    'If ilDays = 0 Then
    '    Beep
    '    edcNoDays.SetFocus
    '    Exit Sub
    'End If
    'slEndDate = Format$(gDateValue(slStartDate) + ilDays - 1, "m/d/yy")


    lbcMsg.Clear
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    imUsingMediaCodeForAV = False

    'Determine file name from date: YYMMDDn.NY
    'where n= version letter
    'gObtainYearMonthDayStr slStartDate, True, slFYear, slFMonth, slFDay
    'ilAutomationType = Val(tgSpf.sAutoType) '5-21-01
    If rbcAutoType(0).Value Then    'Dalet
        ilAutomationType = 1
    ElseIf rbcAutoType(2).Value Then    'Prophet nexgen
        ilAutomationType = 2
    ElseIf rbcAutoType(5).Value Then    'Scott
        ilAutomationType = AUTOTYPE_SCOTT
    ElseIf rbcAutoType(1).Value Then    'Drake
        ilAutomationType = 4
    ElseIf rbcAutoType(4).Value Then    'RCS
        ilAutomationType = 5
    ElseIf rbcAutoType(3).Value Then      '2-5-03 prophet wizard
        ilAutomationType = 6
    ElseIf rbcAutoType(6).Value Then      '2-5-03 prophet media star
        ilAutomationType = 7
    ElseIf rbcAutoType(7).Value Then        '6-25-05 iMediatouch
        ilAutomationType = 8
    ElseIf rbcAutoType(8).Value Then        '8-10-05 Aduio Vault Sat
        ilAutomationType = 9
        If (Asc(tgSpf.sUsingFeatures3) And INCMEDIACODEAUDIOVAULT) = INCMEDIACODEAUDIOVAULT Then
            imUsingMediaCodeForAV = True
        Else
            imUsingMediaCodeForAV = False
        End If

    ElseIf rbcAutoType(9).Value Then        '9-11-06 WireReady
        ilAutomationType = 10
    ElseIf rbcAutoType(10).Value Then       '8-21-08 Simian
        ilAutomationType = 11
    ElseIf rbcAutoType(11).Value Then
        ilAutomationType = 12
    ElseIf rbcAutoType(12).Value Then       '2/1/10: Rivendell
        ilAutomationType = 13
    ElseIf rbcAutoType(13).Value Then       '11-15-10 audio vault prs
        ilAutomationType = 14
    ElseIf rbcAutoType(14).Value Then       '2-18-11 Audio Vault AIR (same as WireReady)
        ilAutomationType = 15
    ElseIf rbcAutoType(15).Value Then       '1-6-12 Wide Orbit (same as Scott)
        ilAutomationType = AUTOTYPE_WIDEORBIT
    ElseIf rbcAutoType(16).Value Then       '6-21-12 Jelli
        ilAutomationType = AUTOTYPE_JELLI
    ElseIf rbcAutoType(17).Value Then       '10/16/12 Enco-ESPN
        ilAutomationType = AUTOTYPE_ENCOESPN
    ElseIf rbcAutoType(18).Value Then           '8-16-13 Scott V5
        ilAutomationType = AUTOTYPE_SCOTT_V5
    ElseIf rbcAutoType(19).Value Then           '1-5-16
        ilAutomationType = AUTOTYPE_ZETTA
    ElseIf rbcAutoType(20).Value Then       '5-11-18
        ilAutomationType = AUTOTYPE_STATIONPL
    ElseIf rbcAutoType(21).Value Then           '5-1-20
        ilAutomationType = AUTOTYPE_RADIOMAN
    End If


    slDayOfWeek = "MONTUEWEDTHUFRISATSUN"
    Screen.MousePointer = vbHourglass
    imExporting = True
    '4/18/14: Moved within loop
    'ilRet = mBBSpots()
    slMissingCopyNames = ""
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            smVehName = Trim$(slName)
            '6-1-05 assign copy to air time spots
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            tmVefSrchKey.iCode = ilVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then       'no error, get the vpf for the possible definition of export location
                '4/18/14: Create BB's just prior to gathering spots to min time if running report that removes BB spots
                ilRet = mBBSpots(ilVefCode)

                ilFoundGroup = True
                If (Not ExptGen!ckcAllItems.Value = vbChecked And imSetIndex > 0) Then                 'all vehicle sets selected
                    ilFoundGroup = False
                    If imSetIndex = 1 Then
                        ilRec = tmVef.iOwnerMnfCode
                    ElseIf imSetIndex = 2 Then
                        ilRec = tmVef.iMnfVehGp2
                    ElseIf imSetIndex = 3 Then
                        ilRec = tmVef.iMnfVehGp3Mkt
                    ElseIf imSetIndex = 4 Then
                        ilRec = tmVef.iMnfVehGp4Fmt
                    ElseIf imSetIndex = 5 Then
                        ilRec = tmVef.iMnfVehGp5Rsch
                    ElseIf imSetIndex = 6 Then
                        ilRec = tmVef.iMnfVehGp6Sub
                    End If
                    For ilLoopSet = 0 To ExptGen!lbcGroupItems.ListCount - 1 Step 1
                        If ExptGen!lbcGroupItems.Selected(ilLoopSet) Then
                            slNameCode = tgSOCodeCT(ilLoopSet).sKey
                            ilRet = gParseItem(slNameCode, 1, "\", slName)
                            ilRet = gParseItem(slName, 3, "|", slName)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            'Determine which vehicle set to test
                            If ilRec = Val(slCode) Then
                                ilFoundGroup = True
                                Exit For
                            End If
                        End If
                    Next ilLoopSet
                End If
                If ilFoundGroup Then

                    ilVpfIndex = -1
                    ilRet = gBinarySearchVpf(tmVef.iCode)
                    If ilRet <> -1 Then
                        ilVpfIndex = ilRet
                    End If
                    If ilVpfIndex = -1 Then
                        ''MsgBox smVehName & " preference get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                        gAutomationAlertAndLogHandler smVehName & " preference get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                        'Print #hmMsg, "** Terminated " & smVehName & " Options not found  **"
                        gAutomationAlertAndLogHandler "** Terminated " & smVehName & " Options not found  **"
                        Close #hmMsg
                        Close #hmTo
                        imExporting = False
                        Screen.MousePointer = vbDefault
                        cmcCancel.SetFocus
                        Exit Sub
                    End If

                    'get the ARF that contains the export location
                    'not all automation types are coded to allow for export paths to be defined in the traffic.ini;
                    'Prophet next gen(2), prophet wizard (6), prophet media star(7), wide orbit or jelli has implementation of paths defined in traffic.ini
                    'If (tgVpf(ilVpfIndex).iAutoExptArfCode > 0) And (ilAutomationType = 2 Or ilAutomationType = 6 Or ilAutomationType = 7 Or ilAutomationType = AUTOTYPE_WIDEORBIT Or ilAutomationType = AUTOTYPE_JELLI Or ilAutomationType = AUTOTYPE_ENCOESPN) Or ilAutomationType = AUTOTYPE_SCOTT Or ilAutomationType = AUTOTYPE_SCOTT_V5 Or ilAutomationType = 9 Then
                    '11-1-13 the export location defined with the vehicle can be applied to any export; it should not be associated with export locations defined with .ini; WRong Test previously implemented.
                     If (tgVpf(ilVpfIndex).iAutoExptArfCode > 0) Then
                        tmArfSrchKey.iCode = tgVpf(ilVpfIndex).iAutoExptArfCode
                        ilRet = btrGetEqual(hmArf, tmArf, imArfRecLen, tmArfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                        If (ilRet <> BTRV_ERR_NONE) Then
                            'MsgBox smVehName & " get Export Loc error, default export loc used", vbOkOnly + vbCritical + vbApplicationModal
                           'Print #hmMsg, "** get Export Loc error, default export loc used for " & smVehName & " Export Loc not found  **"
                           gAutomationAlertAndLogHandler "** get Export Loc error, default export loc used for " & smVehName & " Export Loc not found  **"
                            'Close #hmMsg
                            'Close #hmTo
                            'imExporting = False
                            'Screen.MousePointer = vbDefault
                            'cmcCancel.SetFocus
                            'Exit Sub
                            smVehicleExportLoc = ""
                        Else
                            smVehicleExportLoc = Trim$(tmArf.sFTP)
                        End If
                        'smVehicleExportLoc = Trim$(tmArf.sFTP)
                    Else
                        smVehicleExportLoc = ""        'no export path exists, determine which automation type, may need to create a folder
                    End If


                    'assign copy for all airtime spots for this vehicle for entire date span
                    mAirTimeCopy Format$(llStartDate, "m/d/yy"), Format$(llEndDate, "m/d/yy"), "12M", "12M"
                    For llDate = llStartDate To llEndDate         'loop on all days within the selected vehicle
                        slStartDate = Format(llDate, "m/d/yy")
                        'Determine file name from date: YYMMDDn.NY
                        'where n= version letter
                        gObtainYearMonthDayStr slStartDate, True, slFYear, slFMonth, slFDay
                        slEndDate = slStartDate     'only do one day at a time since each day is a separate file
                        'Print #hmMsg, " "
                        gAutomationAlertAndLogHandler " "
                        'Print #hmMsg, "** Generating Data for " & Trim$(slName) & " for " & slStartDate & " **"
                        gAutomationAlertAndLogHandler "** Generating Data for " & Trim$(slName) & " for " & slStartDate & " **"
                        
                        lacProcessing.Caption = "Generating Data for " & Trim$(slName) & " for " & slStartDate
                        
                        '10/16/12: Create separate file for each Game if Enco-ESPN
                        ReDim llLoopGsfCode(0 To 1) As Long
                        llLoopGsfCode(0) = 0
                        llIndex = 0
                        If ilAutomationType = AUTOTYPE_ENCOESPN Then
                            If tmVef.sType = "G" Then
                                If tmVef.iVefCode = 0 Then
                                    tmGsfSrchKey3.iVefCode = tmVef.iCode
                                    tmGsfSrchKey3.iGameNo = 0
                                    ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                                    Do While (ilCRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = tmVef.iCode)
                                        gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llGameDate
                                        If (llGameDate >= llDate) And (llGameDate <= llDate) Then
                                            If (tmGsf.sGameStatus <> "C") And (tmGsf.iAirVefCode = 0) Then
                                                If llIndex = UBound(llLoopGsfCode) Then
                                                    ReDim Preserve llLoopGsfCode(0 To llIndex + 1)
                                                End If
                                                llLoopGsfCode(llIndex) = tmGsf.lCode
                                                llIndex = llIndex + 1
                                            End If
                                        End If
                                        ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                    Loop
                                End If
                            End If
                        End If
                        
                        For llIndex = 0 To UBound(llLoopGsfCode) - 1 Step 1
                            ilGameNo = 0
                            If (ilAutomationType = 5) Or (ilAutomationType = 12) Then       'RCS has file name as MMDDYY, all others YYMMDD followed by station code
                                slFileName = slFMonth & slFDay & right$(slFYear, 2)
                                slExt = ".log"
                                'ilRecdLen = 67
                                If (ilAutomationType = 12) Then
                                    ilRecdLen = 253
                                Else
                                    ilRecdLen = 67
                                End If
                            ElseIf ilAutomationType = 7 Then
                                slFileName = slFMonth & slFDay & right$(slFYear, 2)
                                slExt = ".trf"
                                ilRecdLen = 73      'Time; Cart #; Spot ID; Advt/Prod; Spot Length (starts in column 70)
                            ElseIf ilAutomationType = 8 Then        '6-25-05 iMediaTouch
                                'slFileName = slFMonth & slFDay & right$(slFYear, 2)
                                slFileName = right$(slFYear, 2) & slFMonth & slFDay
                                slExt = ".trf"
                                ilRecdLen = 74
                            '11-17-10 add Audio Vault RPS
                            ElseIf ilAutomationType = 9 Or ilAutomationType = 14 Then        '8-10-05 Audio Vault Sat
                                slExt = ""              '".txt" no extension
                                ilDays = gWeekDayLong(llDate)
                                slFileName = Mid$(slDayOfWeek, (ilDays) * 3 + 1, 3)
                            ElseIf ilAutomationType = 2 Then            '4-7-06 more fields added for prophet nextgen
                                slFileName = right$(slFYear, 2) & slFMonth & slFDay  '05-01-01
                                slExt = ".txt"
                                ilRecdLen = 131
                            ElseIf ilAutomationType = 11 Then             '8-21-08 simian
                                slFileName = right$(slFYear, 2) & slFMonth & slFDay  '05-01-01
                                slExt = ".txt"
                                ilRecdLen = 109
                            ElseIf ilAutomationType = 13 Then             'Rivendell
                                slFileName = slFYear & slFMonth & slFDay
                                slExt = ".cpi"
                                ilRecdLen = 134
                            ElseIf ilAutomationType = AUTOTYPE_WIDEORBIT Then                               '1-10-12 Wide Orbit
                               'slFileName = right$(slFYear, 2) & slFMonth & slFDay & "00"                  '2-28-12 hard code the 00 required for filename
                                slFileName = right$(slFYear, 2) & slFMonth & slFDay                         '8-15-23 jjb TTP 10803 - Added checkbox to allow user to decide to append "00" to file name
                                If ckcWideOrbit00.Value = vbChecked And ckcWideOrbit00.Visible = True Then
                                    slFileName = slFileName & "00"                                          '8-15-23 jjb TTP 10803 - Added checkbox to allow user to decide to append "00" to file name
                                End If
                                slExt = ".log"                                                              '2-27-12, was .skd
                                ilRecdLen = 108
                            ElseIf ilAutomationType = AUTOTYPE_JELLI Then         '2-21-12 Jelli
                                slFileName = "Jelli-" & slFMonth & slFDay & right$(slFYear, 2)     '2-28-12 hard code the 00 required for filename
                                slExt = ".txt"
                                ilRecdLen = 210
                            ElseIf ilAutomationType = AUTOTYPE_ENCOESPN Then         '10/16/12 Enco-ESPN
                                If tmVef.sType = "G" Then
                                    slFileName = slFMonth & slFDay & right$(slFYear, 2)     '2-28-12 hard code the 00 required for filename
                                    tmGsfSrchKey.lCode = llLoopGsfCode(llIndex)
                                    ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                    If ilRet = BTRV_ERR_NONE Then
                                        ilGameNo = tmGsf.iGameNo
                                        If Trim(tmGsf.sXDSProgCodeID) <> "" Then
                                            slFileName = "_" & Trim(tmGsf.sXDSProgCodeID) & "_"
                                        Else
                                            slFileName = "_" & tmGsf.iGameNo & "_"
                                        End If
                                        For ilTeam = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1
                                            If tmTeam(ilTeam).iCode = tmGsf.iVisitMnfCode Then
                                                slFileName = slFileName & Trim$(tmTeam(ilTeam).sName) & "_Vs_"
                                                Exit For
                                            End If
                                        Next ilTeam
                                        For ilTeam = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1
                                            If tmTeam(ilTeam).iCode = tmGsf.iHomeMnfCode Then
                                                slFileName = slFileName & Trim$(tmTeam(ilTeam).sName)
                                                Exit For
                                            End If
                                        Next ilTeam
                                    End If
                                Else
                                    slFileName = "_" & slFMonth & slFDay & right$(slFYear, 2)     '2-28-12 hard code the 00 required for filename
                                    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                                        If tgVff(ilVff).iVefCode = tmVef.iCode Then
                                            slFileName = "_" & Trim(tgVff(ilVff).sXDProgCodeID) & slFileName
                                            Exit For
                                        End If
                                    Next ilVff
                                End If
                                slExt = ".csv"
                                ilRecdLen = 210
                            ElseIf ilAutomationType = AUTOTYPE_SCOTT_V5 Then        '8-16-13
                                slFileName = right$(slFYear, 2) & slFMonth & slFDay
                                slExt = ".txt"
                                ilRecdLen = 108
                            ElseIf ilAutomationType = AUTOTYPE_ZETTA Then        '1-5-16
                                slFileName = right$(slFYear, 2) & slFMonth & slFDay & Trim$(Left$(tmVef.sCodeStn, 5))
                                slExt = ".txt"
                                ilRecdLen = 218
                             ElseIf ilAutomationType = AUTOTYPE_STATIONPL Then             '5-11-18 Station playlist
                                slFileName = right$(slFYear, 2) & slFMonth & slFDay
                                slExt = ".txt"
                                ilRecdLen = 130
                            Else                                    'everything else falls here if not defined.  one is RADIOMAN & WIRE READY
                                slFileName = right$(slFYear, 2) & slFMonth & slFDay  '05-01-01
                                slExt = ".txt"
                                ilRecdLen = 140
                            End If
                            imFoundspot = False         '1-5-05 assume no spots yet
                            imCopyMissing = False
                            'ReDim smNewLines(0 To 0) As String * 140        '4-7-06 chg from 118 to 140 for prophet nextgen fields
                            ReDim smNewLines(0 To 0) As String * 255        '4-7-06 chg from 118 to 140 for prophet nextgen fields
    
    
                            slLetter = Trim$(Left$(tmVef.sCodeStn, 5))  '1-3-05 chg from 2 char to 5 char station code
                            ilRet = 0
                            'On Error GoTo cmcExportErr:
                            '9-19-05 If ilAutomationType = 7 Or ilAutomationType = 8 Then    'prophet wizard or iMediaTouch
                            If ilAutomationType = 7 Then     'prophet wizard
                                 slToFile = slLetter & slFileName & slExt
                            ElseIf ilAutomationType = 8 Then        '9-19-05 iMediaTouch will be yymmddssss.trf
                                slToFile = slFileName & slLetter & slExt
                            ElseIf ilAutomationType = 9 Or ilAutomationType = 14 Then            'audio vault sat
                                slToFile = slLetter & slFileName & slExt          '8-10-05 Station Code + Day of week + .txt
                            ElseIf ilAutomationType = 13 Then            'Rivendell
                                slToFile = slLetter & slFileName & slExt
                            ElseIf ilAutomationType = AUTOTYPE_WIDEORBIT Or ilAutomationType = AUTOTYPE_JELLI Then           '1-10-12 wide orbit & Jelli do not use station vehicle codes
                                slToFile = slFileName & slExt           'no station code reference, wide orbit cannot handle different filename
                            ElseIf ilAutomationType = AUTOTYPE_ENCOESPN Then            'Rivendell
                                slToFile = Trim$(tmVef.sName) & slFileName & slExt
                            ElseIf ilAutomationType = AUTOTYPE_SCOTT_V5 Then
                                slToFile = slFileName & slExt
                            ElseIf ilAutomationType = AUTOTYPE_ZETTA Then       '1-5-16
                                slToFile = slFileName & slExt
                            Else
                                slToFile = slFileName & slLetter & slExt  '05-01-01    yymmddssss.ext; everything else not defined defaults here :  one is RADIOMAN & Wire Ready
                            End If
    
                            On Error GoTo cmcExportErr:
                            '1-10-12 vehicle can have an export path as an override to the generic export path.
                            'there is also a path for prophet or wide orbit (if using either of those exports) which will be used if no path in the vehicle table is defined.
                            'if no path in traffic.ini for prophet or wide orbit, then the gernic export path is used.
                            If smVehicleExportLoc = "" Then
                                'no vehicle path defined, use generic or specific automation defined in traffic.ini
                                If ilAutomationType = AUTOTYPE_WIDEORBIT Then           'wide orbit
                                    smTempExportPath = sgWideOrbitExportPath
                                ElseIf ilAutomationType = AUTOTYPE_JELLI Then
                                    smTempExportPath = sgJelliExportPath            'export path defined in traffic.ini
                                ElseIf ilAutomationType = 2 Or ilAutomationType = 6 Or ilAutomationType = 7 Then        'any prophet
                                    smTempExportPath = sgProphetExportPath
                                ElseIf ilAutomationType = AUTOTYPE_SCOTT Or ilAutomationType = AUTOTYPE_SCOTT_V5 Then                 '5-30-13
                                    smTempExportPath = sgScottExportPath
                                ElseIf ilAutomationType = AUTOTYPE_ZETTA Then               '1-7-16
                                    smTempExportPath = sgZettaExportPath
                                Else
                                    smTempExportPath = sgExportPath
                                End If
                                If right$(smTempExportPath, 1) <> "\" Then
                                    smTempExportPath = smTempExportPath & "\"
                                End If
                                '6-22-12 Create an export folder if none exists, but its defined in traffic.ini
                                If Not fs.FolderExists(Left$(smTempExportPath, Len(smTempExportPath) - 1)) Then
                                    slFileName = Left$(smTempExportPath, Len(smTempExportPath))
                                    fs.CreateFolder (slFileName)
                                    If Not fs.FolderExists(slFileName) Then
                                        ''gMsgBox ("Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder."), vbOkOnly + vbApplicationModal, "gObtainIniValue"
                                        gAutomationAlertAndLogHandler "Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder.", vbOkOnly + vbApplicationModal, "gObtainIniValue"
                                        'Print #hmMsg, "Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder. "
                                        gAutomationAlertAndLogHandler "Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder. "
                                        lacProcessing.Caption = slName & " export aborted"
                                        'Print #hmMsg, "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                        Close #hmMsg
        
                                        imExporting = False
                                        cmcCancel.Caption = "&Done"
                                        cmcCancel.SetFocus
                                        Screen.MousePointer = vbDefault
                                        Exit Sub
                                    End If
                                End If
                            Else
                                If Not fs.FolderExists(Left$(smVehicleExportLoc, Len(smVehicleExportLoc) - 1)) Then
                                    'gMsgBox ("WideOrbitImport = " & sgWideOrbitImportPath & " path within Traffic.Ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
                                    slFileName = Left$(smVehicleExportLoc, Len(smVehicleExportLoc))
                                    fs.CreateFolder (slFileName)
                                    If Not fs.FolderExists(slFileName) Then
                                        ''gMsgBox ("Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder."), vbOkOnly + vbApplicationModal, "gObtainIniValue"
                                        gAutomationAlertAndLogHandler "Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder.", vbOkOnly + vbApplicationModal, "gObtainIniValue"
                                        'Print #hmMsg, "Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder. "
                                        gAutomationAlertAndLogHandler "Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder. "
                                        lacProcessing.Caption = slName & " export aborted"
                                        'Print #hmMsg, "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                        Close #hmMsg
        
                                        imExporting = False
                                        cmcCancel.Caption = "&Done"
                                        cmcCancel.SetFocus
                                        Screen.MousePointer = vbDefault
                                        Exit Sub
                                    End If
                                End If
    
                                smTempExportPath = smVehicleExportLoc
                                'If right$(smTempExportPath, 1) <> "\" Then
                                '    smTempExportPath = smTempExportPath & "\"
                                'End If
                                smTempExportPath = gSetPathEndSlash(smTempExportPath, True)
                            End If
    
                            ''slDateTime = FileDateTime(sgProphetExportPath & slToFile)    '1-6-05 chged from sgExportPath to new sgProphetExportPath
                            'slDateTime = FileDateTime(smTempExportPath & slToFile)    '2-12-07
                            ilRet = gFileExist(smTempExportPath & slToFile)
    
                            If ilRet = 0 Then
                               ' Kill sgProphetExportPath & slToFile  '1-6-05 chg from sgExportpath to new sgProphetExportPath
                                 Kill smTempExportPath & slToFile  '2-12-07
                            End If
                            On Error GoTo 0
                            'If ilAutomationType = 7 Or ilAutomationType = 8 Or ilAutomationType = 9 Then
                            If ilAutomationType = 7 Or ilAutomationType = 9 Or ilAutomationType = 13 Then
                                'slToFile = sgProphetExportPath & slToFile   'slLetter & slFileName & slExt  '".txt" '05-01-01 slToFile
                                slToFile = smTempExportPath & slToFile   '2-12-07 slLetter & slFileName & slExt  '".txt" '05-01-01 slToFile
                            Else
                                'slToFile = sgProphetExportPath & slToFile   'slFileName & slLetter & slExt '".txt" '05-01-01 slToFile
                                slToFile = smTempExportPath & slToFile   '2-12-07 slFileName & slLetter & slExt '".txt" '05-01-01 slToFile
                            End If
                            ReDim smCurrentLines(0 To 0) As String
                            On Error GoTo 0
                            ilRet = 0
                            'On Error GoTo cmcExportErr:
                            'hmTo = FreeFile
                            'Open slToFile For Output As hmTo
                            ilRet = gFileOpen(slToFile, "Output", hmTo)
                            If ilRet <> 0 Then
                                'Print #hmMsg, "** Terminated **"
                                gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                Close #hmMsg
                                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                                'edcTo.SetFocus
                                Exit Sub
                            End If
                            'Print #hmMsg, "** Storing Output into " & slToFile & " **"
                            gAutomationAlertAndLogHandler "* Storing Output into " & slToFile & " **"
                            ReDim tmExpRecImage(0 To 0) As EXPRECIMAGE
                            ReDim tmProgTimeRange(0 To 0) As PROGTIMERANGE
                            If tmVef.sType = "L" Then
                                '9/9/13: Obtain Parent times
                                If ilAutomationType = AUTOTYPE_ENCOESPN Then
                                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                        If tgMVef(ilVef).iVefCode = tmVef.iCode Then
                                            ilRet = gGetProgramTimes(hmLcf, hmLvf, tgMVef(ilVef).iCode, slStartDate, slEndDate, tmProgTimeRange())
                                        End If
                                    Next ilVef
                                End If
                                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    If tgMVef(ilVef).iVefCode = tmVef.iCode Then
                                        tlVef = tmVef
                                        ilVefCode = tgMVef(ilVef).iCode
                                        tmVefSrchKey.iCode = ilVefCode
                                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                        If ilRet = BTRV_ERR_NONE Then       'no error, get the vpf for the possible definition of export location
                                            'If Not mExptGenDay("C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType(), ilAutomationType, 0) Then
                                            '7/27/12: Added Game test
                                            If tmVef.sType <> "G" Then
                                                mAirTimeCopy Format$(llDate, "m/d/yy"), Format$(llDate, "m/d/yy"), "12M", "12M"
                                                If Not mExptGenDay("C", ilVefCode, Format$(llDate, "m/d/yy"), Format$(llDate, "m/d/yy"), "12AM", "12AM", imEvtType(), ilAutomationType, 0) Then
                                                    'Print #hmMsg, "** Terminated **"
                                                    gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                                    Close #hmMsg
                                                    Close #hmTo
                                                    imExporting = False
                                                    'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                    Screen.MousePointer = vbDefault
                                                    cmcCancel.SetFocus
                                                    Exit Sub
                                                End If
                                            Else
                                                tmGsfSrchKey3.iVefCode = tmVef.iCode
                                                tmGsfSrchKey3.iGameNo = 0
                                                ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                                                Do While (ilCRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = tgMVef(ilVef).iCode)
                                                    gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llGameDate
                                                    If (llGameDate >= llDate) And (llGameDate <= llDate) Then
                                                        If tmGsf.sGameStatus <> "C" Then
                                                            mAirTimeCopy Format$(llDate, "m/d/yy"), Format$(llDate, "m/d/yy"), "12M", "12M"
                                                            If Not mExptGenDay("C", ilVefCode, Format$(llDate, "m/d/yy"), Format$(llDate, "m/d/yy"), "12AM", "12AM", imEvtType(), ilAutomationType, tmGsf.iGameNo) Then
                                                                'Print #hmMsg, "** Terminated **"
                                                                gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                                                Close #hmMsg
                                                                Close #hmTo
                                                                imExporting = False
                                                                'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                Screen.MousePointer = vbDefault
                                                                cmcCancel.SetFocus
                                                                Exit Sub
                                                            End If
                                                        End If
                                                    End If
                                                    ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                Loop
                                            End If
                                        End If
                                        '1/21/09:  Restore vehicle
                                        tmVef = tlVef
                                        ilVefCode = tmVef.iCode
                                    End If
                                Next ilVef
                            Else
                                If (ilAutomationType = AUTOTYPE_ENCOESPN) And (tmVef.sType = "G") Then
                                    ilRet = mExptGenDay("C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType(), ilAutomationType, ilGameNo)
                                Else
                                    ilRet = mExptGenDay("C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType(), ilAutomationType, 0)
                                End If
                                'If Not mExptGenDay("C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType(), ilAutomationType, 0) Then
                                If Not ilRet Then
                                    'Print #hmMsg, "** Terminated **"
                                    gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                    Close #hmMsg
                                    Close #hmTo
                                    imExporting = False
                                    'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                    Screen.MousePointer = vbDefault
                                    cmcCancel.SetFocus
                                    Exit Sub
                                End If
                                If tmVef.iCombineVefCode > 0 Then
                                    '1/21/09:  Retain vehicle so that it can be restored after the combine vehicle data is generated
                                    tlVef = tmVef
                                    ilVefCode = tmVef.iCombineVefCode
                                    tmVefSrchKey.iCode = ilVefCode
                                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                    If ilRet = BTRV_ERR_NONE Then       'no error, get the vpf for the possible definition of export location
                                        'If Not mExptGenDay("C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType(), ilAutomationType, 0) Then
                                        mAirTimeCopy Format$(llDate, "m/d/yy"), Format$(llDate, "m/d/yy"), "12M", "12M"
                                        If Not mExptGenDay("C", ilVefCode, Format$(llDate, "m/d/yy"), Format$(llDate, "m/d/yy"), "12AM", "12AM", imEvtType(), ilAutomationType, 0) Then
                                            'Print #hmMsg, "** Terminated **"
                                            gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                            Close #hmMsg
                                            Close #hmTo
                                            imExporting = False
                                            'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                            Screen.MousePointer = vbDefault
                                            cmcCancel.SetFocus
                                            Exit Sub
                                        End If
                                    End If
                                    '1/21/09:  Restore vehicle
                                    tmVef = tlVef
                                    ilVefCode = tmVef.iCode
                                End If
                            End If
                            If (Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS Then
                                tmGsfSrchKey4.iAirVefCode = tmVef.iCode
                                gPackDate slStartDate, tmGsfSrchKey4.iAirDate(0), tmGsfSrchKey4.iAirDate(1)
                                ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
                                Do While (ilCRet = BTRV_ERR_NONE) And (tmGsf.iAirVefCode = tmVef.iCode)
                                    llGsfCode = tmGsf.lCode
                                    gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llTestDate
                                    If (llTestDate = llDate) Then
                                        ilGameNo = tmGsf.iGameNo
                                        tlVef = tmVef
                                        ilVefCode = tmGsf.iVefCode
                                        tmVefSrchKey.iCode = ilVefCode
                                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                        If ilRet = BTRV_ERR_NONE Then       'no error, get the vpf for the possible definition of export location
                                            'If Not mExptGenDay("C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType(), ilAutomationType, ilGameNo) Then
                                            ilGameVpfIndex = -1
                                            ilRet = gBinarySearchVpf(tmVef.iCode)
                                            If ilRet <> -1 Then
                                                ilGameVpfIndex = ilRet
                                                If ((tgVpf(ilGameVpfIndex).sGenLog <> "L") And (tgVpf(ilGameVpfIndex).sGenLog <> "A")) Or ((tgVpf(ilGameVpfIndex).sGenLog = "A") And (tmGsf.sLiveLogMerge = "M")) Then
                                                    mAirTimeCopy Format$(llDate, "m/d/yy"), Format$(llDate, "m/d/yy"), "12M", "12M"
                                                    If Not mExptGenDay("C", ilVefCode, Format$(llDate, "m/d/yy"), Format$(llDate, "m/d/yy"), "12AM", "12AM", imEvtType(), ilAutomationType, ilGameNo) Then
                                                        'Print #hmMsg, "** Terminated **"
                                                        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                                        Close #hmMsg
                                                        Close #hmTo
                                                        imExporting = False
                                                        'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                        Screen.MousePointer = vbDefault
                                                        cmcCancel.SetFocus
                                                        Exit Sub
                                                    End If
                                                End If
                                            End If
                                        End If
                                        '1/21/09:  Restore vehicle
                                        tmVef = tlVef
                                        ilVefCode = tmVef.iCode
                                    Else
                                        Exit Do
                                    End If
                                    'ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                    '11/20/09:  Reset for key 4 because gsf is read in mExptGenDay with key 3
                                    tmGsfSrchKey4.iAirVefCode = tmVef.iCode
                                    gPackDate slStartDate, tmGsfSrchKey4.iAirDate(0), tmGsfSrchKey4.iAirDate(1)
                                    ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
                                    Do While (ilCRet = BTRV_ERR_NONE)
                                        If (tmGsf.lCode = llGsfCode) Then
                                            Exit Do
                                        End If
                                        ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                    Loop
                                    If (ilCRet = BTRV_ERR_NONE) Then
                                        ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                    End If
                                Loop
                            End If
                            If UBound(tmExpRecImage) > 0 Then
                                ArraySortTyp fnAV(tmExpRecImage(), 0), UBound(tmExpRecImage), 0, LenB(tmExpRecImage(0)), 0, LenB(tmExpRecImage(0).sKey), 0
                            End If
                            'ReDim smNewLines(0 To UBound(tmExpRecImage)) As String * 140    '4-7-06 chg from 118 to 140 for prophet nextgen fields
                            ReDim smNewLines(0 To UBound(tmExpRecImage)) As String * 255    '4-7-06 chg from 118 to 140 for prophet nextgen fields
                            For ilLine = 0 To UBound(tmExpRecImage) - 1 Step 1
                                smNewLines(ilLine) = tmExpRecImage(ilLine).sRecord
                            Next ilLine
                            If imCopyMissing Then
                                If slMissingCopyNames = "" Then
                                    slMissingCopyNames = slName
                                Else
                                    slMissingCopyNames = slMissingCopyNames & ", " & slName
                                End If
                            End If
                            If ilAutomationType = 7 Or ilAutomationType = 8 Or ilAutomationType = 9 Then
                                lacProcessing.Caption = "Writing Data to " & slLetter & slFileName & slExt   '".txt" '09-25-01
                            Else
                                lacProcessing.Caption = "Writing Data to " & slFileName & slLetter & slExt  '".txt" '05-01-01
                            End If
                            DoEvents
                            ilRet = 0
                            'On Error GoTo cmcExportErr:
                            'Write out the exported spots to disk
                            mReSetBreakNumbers ilAutomationType
                            For ilLine = LBound(smNewLines) To UBound(smNewLines) - 1 Step 1
                                '3= scott, 19 = scott V5, 16 = wide orbit, 9 = audio vault sat, 14 = audio vault prs, 6-21-12 Jelli added comma delimited
                                '19 = scottV5 comma delimited added 8-16-13
                                If ilAutomationType = AUTOTYPE_SCOTT Or ilAutomationType = AUTOTYPE_SCOTT_V5 Or ilAutomationType = 9 Or ilAutomationType = 14 Or ilAutomationType = AUTOTYPE_WIDEORBIT Or ilAutomationType = AUTOTYPE_JELLI Then
                                                                                            'Audio Vault Sat is vertical bar delimited, not fixed length.
                                    ilRecdLen = Len(Trim(smNewLines(ilLine)))
                                End If
                                Print #hmTo, Left(smNewLines(ilLine), ilRecdLen)
                                If ilRet <> 0 Then
                                    imExporting = False
                                    'Print #hmMsg, "** Terminated **"
                                    gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                                    Close #hmMsg
                                    Close #hmTo
                                    Screen.MousePointer = vbDefault
                                    ''MsgBox "Error writing to " & sgDBPath & "Messages\" & "ExptGen.Txt" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                    gAutomationAlertAndLogHandler "Error writing to " & sgDBPath & "Messages\" & "ExptGen.Txt", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                    cmcCancel.SetFocus
                                    Exit Sub
                                End If
                            Next ilLine
                            'Print #hmMsg, "** Completed " & Trim$(tmVef.sName) & " for " & slStartDate & " **"
                            gAutomationAlertAndLogHandler "** Completed " & Trim$(tmVef.sName) & " for " & slStartDate & " **"
                            Close #hmTo         'close the message file
    
                            '5-30-13 by option, allow empty files without spots to be created
                            If imFoundspot Or tgSaf(0).sGenAutoFileWOSpt = "Y" Then          '1-5-05 if no spots found, dont retain export file and dont show file sent to message
                                lacProcessing.Caption = "Output for " & smScreenCaption & " sent to " & slToFile
                            Else
                                lacProcessing.Caption = "No spots found for " & smVehName & " on " & slStartDate & ", no export created"
                                Kill (slToFile)
                                'Print #hmMsg, "No spots found for " & smVehName & " on " & slStartDate & ", no export created"
                                gAutomationAlertAndLogHandler "No spots found for " & smVehName & " on " & slStartDate & ", no export created"
                            End If
                            'Print #hmMsg, "** Completed " & smScreenCaption & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                            gAutomationAlertAndLogHandler "** Completed " & smScreenCaption & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                        
                        Next llIndex 'next llLoopGsfCode
                        'Else
                        'End If
                    Next llDate         'next date
                    'Set Log dates for vehicles that Logs are not generated
                    If (tgVpf(ilVpfIndex).sGenLog = "N") Then
                        ilRet = mSetLogDate()
                        If tmVef.iCombineVefCode > 0 Then
                            '1/21/09:  Retain vehicle so that it can be restored after the combine vehicle data is generated
                            tlVef = tmVef
                            ilVefCode = tmVef.iCombineVefCode
                            tmVefSrchKey.iCode = ilVefCode
                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                            If ilRet = BTRV_ERR_NONE Then       'no error, get the vpf for the possible definition of export location
                                ilRet = mSetLogDate()
                            End If
                            '1/21/09:  Restore vehicle
                            tmVef = tlVef
                            ilVefCode = tmVef.iCode
                        End If
                        If (Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS Then
                            tmGsfSrchKey4.iAirVefCode = tmVef.iCode
                            gPackDate slStartDate, tmGsfSrchKey4.iAirDate(0), tmGsfSrchKey4.iAirDate(1)
                            ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
                            Do While (ilCRet = BTRV_ERR_NONE) And (tmGsf.iAirVefCode = tmVef.iCode)
                                gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llTestDate
                                If (llTestDate >= llStartDate) And (llTestDate <= llEndDate) Then
                                    tlVef = tmVef
                                    ilVefCode = tmGsf.iVefCode
                                    tmVefSrchKey.iCode = ilVefCode
                                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                    If ilRet = BTRV_ERR_NONE Then       'no error, get the vpf for the possible definition of export location
                                        ilGameVpfIndex = -1
                                        ilRet = gBinarySearchVpf(tmVef.iCode)
                                        If ilRet <> -1 Then
                                            ilGameVpfIndex = ilRet
                                            If ((tgVpf(ilGameVpfIndex).sGenLog <> "L") And (tgVpf(ilGameVpfIndex).sGenLog <> "A")) Or ((tgVpf(ilGameVpfIndex).sGenLog = "A") And (tmGsf.sLiveLogMerge = "M")) Then
                                                ilRet = mSetLogDate()
                                            End If
                                        End If
                                    End If
                                    '1/21/09:  Restore vehicle
                                    tmVef = tlVef
                                    ilVefCode = tmVef.iCode
                                End If
                                Exit Do
                            Loop
                        End If
                    End If
                End If              'ilfoundgroup
            Else                    'error, vehicle not found
                'Print #hmMsg, " "
                gAutomationAlertAndLogHandler " "
                'Print #hmMsg, "Name: " & slName & " not found"
                gAutomationAlertAndLogHandler "Name: " & slName & " not found"
                lacProcessing.Caption = slName & " not found: vehicle aborted"
            End If                  'ilret <> BTRV_err_none
            ''Set Log dates for vehicles that Logs are not generated
            'ilRet = mSetLogDate(-1)
        End If                  'lbcVehicle.Selected(ilLoop)
    Next ilLoop                 'next vehicle

    'Print #hmMsg, "** Completed " & smScreenCaption & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Close #hmMsg
    On Error GoTo 0

    'lacProcessing.Caption = "Output for " & smScreenCaption & " sent to " & slToFile
    lacMsg.Caption = "Messages sent to " & sgDBPath & "Messages\" & "ExptGen.Txt"
    Screen.MousePointer = vbDefault
    'If slMissingCopyNames <> "" Then
    '    MsgBox "Copy missing on " & slMissingCopyNames, vbOkOnly + vbExclamation + vbApplicationModal, "Copy Missing"
    'End If
    imExporting = False
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    Exit Sub

cmcExportErr:
    ilRet = err.Number
    Resume Next
    
    
ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
    
End Sub
Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub


Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
    mSetCommands
End Sub
Private Sub cmcStartDate_GotFocus()
    plcCalendarTo.Visible = False
    gCtrlGotFocus ActiveControl
    mSetCommands
End Sub

Private Sub edcEndDate_Change()
    Dim slStr As String
    plcCalendar.Visible = False

    slStr = edcEndDate.Text
    If Not gValidDate(slStr) Then
        lacDateTo.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendarTo_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcEndDate_Click()
    plcCalendar.Visible = False
    mSetCommands
End Sub

Private Sub edcEndDate_GotFocus()
    plcCalendar.Visible = False

    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDate
    mSetCommands
End Sub

Private Sub edcEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
End Sub

Private Sub edcEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcEndDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub

Private Sub edcEndDate_KeyUp(KeyCode As Integer, Shift As Integer)
Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendarTo.Visible = Not plcCalendarTo.Visible
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcEndDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcEndDate.Text = slDate
            End If
        End If
        edcEndDate.SelStart = 0
        edcEndDate.SelLength = Len(edcEndDate.Text)
    End If
    mSetCommands
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcSet1_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub edcStartDate_Change()
    Dim slStr As String
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcStartDate_Click()
    mSetCommands
End Sub
Private Sub edcStartDate_GotFocus()
    plcCalendarTo.Visible = False
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcStartDate
    mSetCommands
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
    mSetCommands
End Sub
Private Sub edcStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    mSetCommands
End Sub
Private Sub edcStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
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
    'Me.Visible = False
    'Me.Visible = True
    frcZone.Visible = False
    frcSplit.Visible = False
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
    frcZone.Visible = True
    '3-23-09 if using network split, ask question to include primary or secondary
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        frcSplit.Visible = True
    End If
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcCalendarTo.Visible = False
        gFunctionKeyBranch KeyCode
'        If plcAutoType.Visible = True Then
'            plcAutoType.Visible = False
'            plcAutoType.Visible = True
'        End If
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        'cmcCancel_Click
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width  'move off the screen so screen won't flash
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmExpRecImage
    Erase tmTeam
    Erase tmProgTimeRange
    Erase tmBreakByProg
    Erase tmLLC
    
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmCTSsf)
    btrDestroy hmCTSsf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmTzf)
    btrDestroy hmTzf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    ilRet = btrClose(hmDlf)
    btrDestroy hmDlf
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmAxf)
    btrDestroy hmAxf
    ilRet = btrClose(hmRsf)
    btrDestroy hmRsf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmEtf)
    btrDestroy hmEtf
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmLvf)
    btrDestroy hmLvf
    ilRet = btrClose(hmAgf)         '6-21-12
    btrDestroy hmAgf
    ilRet = btrClose(hmArf)         '6-21-12
    btrDestroy hmArf
    
    Set ExptGen = Nothing   'Remove data segment
    
End Sub



Private Sub frcSplit_Click()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub frcZone_Click()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lacCntr_Click(Index As Integer)
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub lacDate_Click()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub lacDateTo_Click()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub lacEndDate_Click()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
    mSetCommands
End Sub

Private Sub lacStartDate_Click()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
    mSetCommands
End Sub

Private Sub lacZone_Click()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub lbcGroupItems_Click()
    If Not imAllClickedItems Then
        imSetAllItems = False
        ckcAllItems.Value = vbUnchecked  '9-12-02 False
        imSetAllItems = True
    End If
    mSetCommands
End Sub

Private Sub lbcGroupItems_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  '9-12-02 False
        imSetAll = True
    End If
    mSetCommands
End Sub
Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
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
Private Sub mBoxCalDate(EditDate As Control, LabelDate As Control)
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = EditDate.Text   'edcStartDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    LabelDate.Caption = slDay
                    LabelDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    LabelDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            LabelDate.Visible = False
        Else
            LabelDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
'***********************************************************
'*                                                         *
'*      Procedure Name:mExptGenDay                         *
'*                                                         *
'*             Created:5/18/93       By:D. LeVine          *
'*            Modified:              By:                   *
'*                                                         *
'*            Comments:Build export record                 *
'*                     ("EST" zone, Cmml Sch = "Y",        *
'*                      and Avail Name = "N")              *
'*                                                         *

'*******************************************************
Private Function mExptGenDay(sLCP As String, ilVefCode As Integer, slSDate As String, slEDate As String, slStartTime As String, slEndTime As String, ilEvtType() As Integer, ilAutomationType As Integer, ilGameNo As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slPricePty                                                                            *
'******************************************************************************************

'
'   ilRet = mExptGenDay (slCp, ilVefCode, slSDate, slEDate, slStartTime, slEndTime, ilEvtType())
'
'   Where:
'       slCP (I)- "C"=Current only; "P"=Pending only; "B"=Both
'       ilVefCode (I)-Vehicle code number(slFor = L or C) or feed code (slFor = D)
'       slSDate (I)- Start Date that events are to be obtained
'       slEDate (I)- Start Date that events are to be obtained
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'       ilEvtType (I)- Array of which events are to be included (True or False)
'                       Index description
'                         0   Library
'                         1   Program event
'                         2   Contract avail
'                         3   Open BB
'                         4   Floating BB
'                         5   Close BB
'                         6   Cmml promo
'                         7   Feed avail
'                         8   PSA avail
'                         9   Promo avail
'                         10  Page eject
'                         11  Line space 1
'                         12  Line space 2
'                         13  Line space 3
'                         14  Other event types
'       Automation type changed from byte to bit value 5-21-01
'
'       ilautomationType                                           ilAutomationType
'         (as input)                                               (converted from inpu screen)
'        spfAutoType &                          rbc control
'        spfAutoType1     Automation            from input
'            &h1            DALET               rbc(0)              1
'            &H2            Prophet NexGen      rbc(2)              2
'            &h4            Scott               rbc(5)              3
'            &h8            Drake               rbc(1)              4
'            &h10           RCS-4               rbc(4)              5
'            &h20           Prophet Wizard      rbc(3)              6    2-5-03
'            &h40           Prophet MediaStar   rbc(6)              7    9-25-03
'            &h80           imediaTouch         rbc(7)              8    6-25-05
'            &h100          Audio Vault sat     rbc(8)              9    8-10-05
'            &h400          Wire Ready          rbc(9)              10
'            &h4000         Simian              rbc(10)             11   8-21-08
'            &h8000         RCS-5               rbc(11)             12
'                           Rivendell           rbc(12)             13
'            &h1            AudioVault RPS      rbc(13)             14   11-15-10
'       spfAutoType3
'
'             &H2           WideOrbit           rbc(15)             16
'             &H4           Jelli               rbc(16)             17  6-21-12
'             &H8           ENCO-ESPN           rbc(17)             18  10/16/12
'             &H10          Scott - V5          rbc(18)             19  8-16-13
'             &H20          Zetta               rbc(19)             20  1-5-16
'             &H40          StationPlaylist     rbc(20)             21   5-11-18
'             &H80          RadidoMan           rbc(21)             22   5-1-20
'       ilGameNo (I)- Game Number to generate export for
'
    Dim ilType As Integer
    Dim ilRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilSsfDate0 As Integer
    Dim ilSsfDate1 As Integer
    Dim ilEvt As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilSpot As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilTerminated As Integer
    Dim slTime As String
    Dim ilVpfIndex As Integer
    Dim ilVehCode As Integer
    Dim ilDlfDate0 As Integer
    Dim ilDlfDate1 As Integer
    Dim ilDlfFound As Integer
    Dim ilVlfDate0 As Integer
    Dim ilVlfDate1 As Integer
    Dim ilAirHour As Integer
    Dim ilLocalHour As Integer
    Dim ilSIndex As Integer
    Dim slSsfDate As String
    'Spot summary
    Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim tlSsfSrchKey2 As SSFKEY2 'SSF key record image
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim llEvtTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilWithinTime As Integer
    Dim ilEvtFdIndex As Integer
    Dim ilFirstEvtShown As Integer      '11-18-10
    Dim slXMid As String                '11-18-10
    Dim llEvtCefCode As Long            '11-18-10
    Dim slAirDate As String
    Dim llAirDate As Long
    Dim ilTest As Integer
    Dim llSpotTime As Long
    Dim llHour As Long
    Dim llMin As Long
    Dim llSec As Long
    Dim llAvailTime As Long
    Dim llPrevTime As Long
    Dim ilCopy As Integer
    Dim ilDlfExist As Integer
    Dim llCopyMissingSdfCode As Long
    'Spot detail record information
    Dim slStr As String
    Dim slMsg As String
    Dim ilAvailUnits As Integer
    Dim ilAvailLen As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSubFeed As Integer    'True=Output only Dlf records with subfeeds;
                                'False=Output only records without subfeeds
    Dim slZone As String
    Dim slAdvtName As String
    Dim slProdName As String
    Dim slAgyName As String
    Dim slCodeStn As String
    Dim slRecord As String
    Dim ilLineNo As Integer
    Dim ilPageNo As Integer
    Dim ilShowTime As Integer
    ReDim tmSpotTimes(0 To 0) As SPOTTIMES
    'Dim slCifName As String
    'Dim slCreativeTitle As String
    'Dim slISCI As String
    'Dim slMcfPrefix As String           '1-19-12 Wide Orbit will use the Media Code Prefix for the category type in export
    Dim llTempTime As Long
    Dim slHour As String * 2
    Dim slLenInSec As String
    Dim slLenInMinSec As String
    Dim llPrevAvailEndTime As Long          '1-27-03
    Dim llSpotEndTime As Long
    Dim llAvailStartTime As Long        '1-27-03
    Dim ilBBPass As Integer
    Dim llBBTime As Long
    Dim tlBBAvail As AVAILSS
    Dim tlSdf As SDF
    Dim ilFound As Integer
    Dim ilAddSpot As Integer
    Dim ilBB As Integer
    Dim ilFdOpen As Integer
    Dim ilFdClose As Integer
    Dim ilEvtRet As Integer
    Dim ilStartTime0 As Integer
    Dim ilStartTime1 As Integer
    Dim slSortDate As String
    Dim slSortType As String
    Dim llOpenTime As Long
    Dim llCloseTime As Long
    Dim llTestTime As Long
    Dim tlBBSdf As SDF
    Dim slDelimiter As String      'delimeter to separate fields
    Dim ilSameBreak As Integer
    Dim slTemp As String
    Dim ilPos As Integer
    Dim ilVefIndex As Integer
'    Dim slAdvtCode As String     '4-7-06 addl fields for prophet nextgen
'    Dim slCompCode As String
'    Dim slROS As String
'    Dim slFixed As String
'    Dim slPty As String
'    Dim ilMediaCodeLen As Integer       'len of material code string (may need to exclude them for Audio Vault option)
    Dim ilBBLen As Integer
    Dim ilAdjZone As Integer
    Dim llAdjTime As Long
    Dim ilWegenerOLA As Integer
    Dim ilIncludeFullOrSplit As Integer '3-23-09
    Dim ilLLCIndex As Integer
    Dim ilLLCLoop As Integer    'ReDim ilEvtType(0 To 14) As Integer
    Dim ilFoundLLC As Integer
    ReDim llPrevOpenFdBBSpots(0 To 0) As Long
    ReDim llPrevCloseFdBBSpots(0 To 0) As Long
    Dim ilLcfRet As Integer
    Dim ilAtLeastOneSpot As Integer
    Dim blAvailOk As Boolean
    Dim ilAnf As Integer
'    Dim slMediaCodeSuppressSpot As String
    Dim ilIndex As Integer
    Dim slLogDay As String
    Dim slLogMonth As String
    Dim slLogYear As String
    Dim slEventName As String
    Dim llESPNPrevAvailTime As Long
    Dim llESPNPrevDate As Long
    Dim ilESPNHour As Integer
    Dim ilESPNBreak As Integer
    Dim ilESPNPosition As Integer
    Dim slProgCodeID As String
    Dim ilVff As Integer
    Dim ilParentVefCode As Integer
    '7/11/14: Lock avails
    Dim blInProcSpot As Boolean
    Dim ilSub As Integer
    Dim blBypassZeroUnits As Boolean
    
    ilLineNo = 66  'Force header
    ilPageNo = 0
    ilType = 0
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    'these entries would be to export open/close billboards
    ilEvtType(2) = True
    ilEvtType(3) = True
    ilEvtType(5) = True
    If ilAutomationType = 14 Then           'audio vault rps
        ilEvtType(14) = True                'include comments
    End If
    If ilAutomationType = AUTOTYPE_ENCOESPN Then           'ESPN
        ilEvtType(1) = True                'include comments
    End If
    
    'ReDim tmExpRecImage(0 To 0) As EXPRECIMAGE
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    ilSsfRecLen = Len(tmSsf)  'Get and save SSF record length
    On Error GoTo 0
    'tmVef.iCode = 0
    llSDate = gDateValue(slSDate)
    llEDate = gDateValue(slEDate)
    ilESPNHour = -1
    blInProcSpot = False
    For llDate = llSDate To llEDate + 1 Step 1  'Process next date for avails that map back one day
        imXMidNight = False
        lmLastEvtTime = -1
        'ilProphetFlag = False           '10-04-01
        ReDim tmBBSdfInfo(0 To 0) As BBSDFINFO
        ilWithinTime = False
        slDate = Format$(llDate, "m/d/yy")
        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        gObtainYearMonthDayStr slDate, True, slLogYear, slLogMonth, slLogDay

        ilVehCode = ilVefCode
        If ilVehCode <> tmVef.iCode Then
            'tmVefSrchKey.iCode = ilVehCode
            'ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            'If (ilRet <> BTRV_ERR_NONE) Then
            '    Screen.MousePointer = vbDefault
            '    mExptGenDay = False
            '    MsgBox smVehName & " get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            '    Exit Function
            '    Screen.MousePointer = vbHourGlass
            'End If
        End If
            ilVpfIndex = -1
            'For ilLoop = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
            '    If tmVef.iCode = tgVpf(ilLoop).iVefKCode Then
                ilLoop = gBinarySearchVpf(tmVef.iCode)
                If ilLoop <> -1 Then
                    ilVpfIndex = ilLoop
            '        Exit For
                End If
            'Next ilLoop
            If ilVpfIndex = -1 Then
                Screen.MousePointer = vbDefault
                mExptGenDay = False
                ''MsgBox smVehName & " preference get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                gAutomationAlertAndLogHandler smVehName & " preference get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                Screen.MousePointer = vbHourglass
                Exit Function
            End If
        'End If

        'get the defined export location if applicable
        '12/29/08:  If cart # not defined, then use the Reel # if Wegener or OLA
        ilWegenerOLA = False
        'If (tgVpf(ilVpfIndex).sWegenerExport = "Y") Or (tgVpf(ilVpfIndex).sOLAExport = "Y") Then
        If (tgVpf(ilVpfIndex).sOLAExport = "Y") Then
            ilWegenerOLA = True
        End If
        ilDlfFound = False
        'If (tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(1) <> 0)) Then
        If (tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(0) <> 0)) Then
            'Obtain Engineering records for date
            If (ilDay >= 0) And (ilDay <= 4) Then
                slDay = "0"
            ElseIf ilDay = 5 Then
                slDay = "6"
            Else
                slDay = "7"
            End If
            'Obtain the start date of Dlf
            tmDlfSrchKey.iVefCode = ilVehCode
            tmDlfSrchKey.sAirDay = slDay
            tmDlfSrchKey.iStartDate(0) = ilLogDate0
            tmDlfSrchKey.iStartDate(1) = ilLogDate1
            tmDlfSrchKey.iAirTime(0) = 0
            tmDlfSrchKey.iAirTime(1) = 6144 '24*256
            ilRet = btrGetLessOrEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) Then
                ilDlfDate0 = tmDlf.iStartDate(0)
                ilDlfDate1 = tmDlf.iStartDate(1)
                ilDlfFound = True
            Else
                ilDlfDate0 = 0
                ilDlfDate1 = 0
                tmDlf.sZone = ""
            End If
            'Obtain the start date of VLF
            If tmVef.sType = "A" Then
                ilVlfDate0 = 0
                ilVlfDate1 = 0
                tmVlfSrchKey1.iAirCode = ilVehCode
                tmVlfSrchKey1.iAirDay = Val(slDay)
                tmVlfSrchKey1.iEffDate(0) = ilLogDate0
                tmVlfSrchKey1.iEffDate(1) = ilLogDate1
                tmVlfSrchKey1.iAirTime(0) = 0
                tmVlfSrchKey1.iAirTime(1) = 6144    '24*256
                tmVlfSrchKey1.iAirPosNo = 32000
                tmVlfSrchKey1.iAirSeq = 32000
                ilRet = btrGetLessOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode)
                    ilTerminated = False
                    'Check for CBS
                    If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                        If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                            ilTerminated = True
                        End If
                    End If
                    If (tmVlf.sStatus <> "P") And (tmVlf.iAirDay = Val(slDay)) And (Not ilTerminated) Then
                        ilVlfDate0 = tmVlf.iEffDate(0)
                        ilVlfDate1 = tmVlf.iEffDate(1)
                        Exit Do
                    End If
                    ilRet = btrGetPrevious(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                'If (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) Then
                '    ilVlfDate0 = tmVlf.iEffDate(0)
                '    ilVlfDate1 = tmVlf.iEffDate(1)
                'Else
                '    ilVlfDate0 = 0
                '    ilVlfDate1 = 0
                'End If
            End If
        End If

        blBypassZeroUnits = False
        If (tmVef.sType = "A") Then
            ilVff = gBinarySearchVff(tmVef.iCode)
            If ilVff <> -1 Then
                If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                    blBypassZeroUnits = True
                End If
            End If
        End If


        'If Not ilDlfFound Then
        '    Screen.MousePointer = vbDefault
        '    MsgBox smVehName & " delivery missing", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
        '    mExptGenDay = False
        '    Screen.MousePointer = vbHourGlass
        '    Exit Function
        'End If
        'gObtainVlf hlVlf, ilVehCode, llDate, tlVlf0(), tlVlf5(), tlVlf6()


        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilSsfDate0 = ilLogDate0
        ilSsfDate1 = ilLogDate1
        ilVefIndex = gBinarySearchVef(ilVehCode)
        If ilVefIndex = -1 Then
            mExptGenDay = False
            Exit Function
        End If
        If tgMVef(ilVefIndex).sType <> "G" Then
            imGameVehicle = False
            ilType = 0
            tlSsfSrchKey.iType = ilType
            tlSsfSrchKey.iVefCode = ilVehCode
            tlSsfSrchKey.iDate(0) = ilSsfDate0
            tlSsfSrchKey.iDate(1) = ilSsfDate1
            tlSsfSrchKey.iStartTime(0) = 0
            tlSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        Else
            imGameVehicle = True
            tlSsfSrchKey2.iVefCode = ilVehCode
            tlSsfSrchKey2.iDate(0) = ilSsfDate0
            tlSsfSrchKey2.iDate(1) = ilSsfDate1
            ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY) 'Get last current record to obtain date
            If ilGameNo <> 0 Then
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVehCode)
                    If tmSsf.iType = ilGameNo Then
                        Exit Do
                    End If
                    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            ilType = tmSsf.iType
        End If
        If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> ilVehCode) Or (tmSsf.iDate(0) <> ilSsfDate0) Or (tmSsf.iDate(1) <> ilSsfDate1) Then
            'If airing- then use first Ssf prior to date defined
            If tmVef.sType = "A" Then
                ilSsfDate0 = 0
                ilSsfDate1 = 0
                ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                tlSsfSrchKey.iType = ilType
                tlSsfSrchKey.iVefCode = ilVehCode
                tlSsfSrchKey.iDate(0) = ilLogDate0
                tlSsfSrchKey.iDate(1) = ilLogDate1
                tlSsfSrchKey.iStartTime(0) = 0
                tlSsfSrchKey.iStartTime(1) = 6144   '24*256
                ilRet = gSSFGetLessOrEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode)
                    gUnpackDate tmSsf.iDate(0), tmSsf.iDate(1), slSsfDate
                    If (ilDay = gWeekDayStr(slSsfDate)) And (tmSsf.iStartTime(0) = 0) And (tmSsf.iStartTime(1) = 0) Then
                        ilSsfDate0 = tmSsf.iDate(0)
                        ilSsfDate1 = tmSsf.iDate(1)
                        Exit Do
                    End If
                    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetPrevious(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
        End If
        DoEvents
        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode) Then
            gUnpackDate ilSsfDate0, ilSsfDate1, slSsfDate
            ReDim tmLLC(0 To 0) As LLC  'Image
            If (tgSpf.sUsingBBs) = "Y" Or (ilAutomationType = 14) Or (ilAutomationType = AUTOTYPE_ENCOESPN) Then      '11-23 add option to see comments with audio Vault RPS
                ilEvtRet = gBuildEventDay(ilType, "C", ilVehCode, slSsfDate, "12M", "12M", ilEvtType(), tmLLC())
                ilLLCIndex = 0      'keep track of the last library event processed for spots
            End If
            ReDim llPrevOpenFdBBSpots(0 To 0) As Long
            ReDim llPrevCloseFdBBSpots(0 To 0) As Long
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
                'Loop thru Ssf and move records to export
                llPrevAvailEndTime = 0      '1-27-03 initalize first time thru for block notations
                ilEvt = 1
                '9/21/09: Bypass and game set a Live Log
                If tgVpf(ilVpfIndex).sGenLog = "A" Then
                    tmGsfSrchKey3.iVefCode = ilVehCode
                    tmGsfSrchKey3.iGameNo = ilType
                    ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    'If ilRet = BTRV_ERR_NONE Then
                    Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVehCode) And (tmGsf.iGameNo = ilType)
                        If (tmGsf.iAirDate(0) = ilSsfDate0) And (tmGsf.iAirDate(1) = ilSsfDate1) Then
                            If tmGsf.sLiveLogMerge = "L" Then
                                ilEvt = tmSsf.iCount + 1
                            End If
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    Loop
                Else
                    '7/10/13:  Get Game so that Prog Code ID would be available
                    If ilType > 0 Then
                        tmGsfSrchKey3.iVefCode = ilVehCode
                        tmGsfSrchKey3.iGameNo = ilType
                        ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                        Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVehCode) And (tmGsf.iGameNo = ilType)
                            If (tmGsf.iAirDate(0) = ilSsfDate0) And (tmGsf.iAirDate(1) = ilSsfDate1) Then
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        Loop
                    End If
                End If
                
                ilAtLeastOneSpot = False
                If ilEvt <= tmSsf.iCount Then       'ilEvt set to avoid going thru export if live log; do not set complete flag
                    '4-26-11 if an export is created,set the Day is Not complete flag
                    gUpdateLCFCompleteFlag hmLcf, tmLcf, ilSsfDate0, ilSsfDate1, ilType, ilVehCode, "I"
                End If
                slEventName = ""
                Do While ilEvt <= tmSsf.iCount
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If (tmProg.iRecType = 1) Or ((tmProg.iRecType >= 2) And (tmProg.iRecType <= 9)) Then
                        gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llEvtTime
                        If llEvtTime > llEndTime Then
                            ilWithinTime = False
                            Exit Do
                        End If
                        If llEvtTime >= llStartTime Then
                            ilWithinTime = True
                        End If
                    End If
                    '4/26/11: Test if avail spot should be exported
                    blAvailOk = True
                    If (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        ilAnf = gBinarySearchAnf(tmAvail.ianfCode, tgAvailAnf())
                        If ilAnf <> -1 Then
                            If tgAvailAnf(ilAnf).sAutomationExport = "N" Then
                                blAvailOk = False
                            End If
                        End If
                        If blBypassZeroUnits Then
                            If (tmAvail.iAvInfo And &H1F <= 0) Or (tmAvail.iLen <= 0) Then
                                blAvailOk = False
                            End If
                        End If
                    End If
                    ilEvtFdIndex = -1
                    If ilWithinTime And blAvailOk Then
                        If tmProg.iRecType = 1 Then    'Program
                            ilAirHour = tmProg.iStartTime(1) \ 256  'Obtain Hour
                            If (llDate > llEDate) And (ilAirHour >= 6) Then
                                'Done processing
                                mExptGenDay = True
                                Exit Function
                            End If
                            If ilAutomationType = AUTOTYPE_ENCOESPN Then
                                slEventName = ""
                                For ilLoop = LBound(tmLLC) To UBound(tmLLC) - 1 Step 1
                                    'Match start time and length
                                    If (tmLLC(ilLoop).iEtfCode = 1) Then
                                        gPackTime tmLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                        If (ilStartTime0 = tmProg.iStartTime(0)) And (ilStartTime1 = tmProg.iStartTime(1)) Then
                                            slEventName = tmLLC(ilLoop).sName
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If
                        ''ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
                        ElseIf (tmProg.iRecType = 2) Or ((tmProg.iRecType >= 6) And (tmProg.iRecType <= 9)) Then 'Avail
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            tmOpenAvail.iRecType = -1
                            tmCloseAvail.iRecType = -1
                            If (tgSpf.sUsingBBs = "Y") And ((tmProg.iRecType = 2) Or ((tmProg.iRecType >= 6) And (tmProg.iRecType <= 9))) Then
                                For ilLoop = LBound(tmLLC) To UBound(tmLLC) - 1 Step 1
                                    'Match start time and length
                                    If (tmLLC(ilLoop).iEtfCode >= 2) And (tmLLC(ilLoop).iEtfCode <= 9) Then
                                        gPackTime tmLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                        If (ilStartTime0 = tmAvail.iTime(0)) And (ilStartTime1 = tmAvail.iTime(1)) Then
                                            'Scan to find Open & Close avail
                                            For ilFdOpen = ilLoop To LBound(tmLLC) Step -1
                                                If tmLLC(ilFdOpen).iEtfCode = 2 Then
                                                    llTestTime = gTimeToLong(tmLLC(ilFdOpen).sStartTime, False)
                                                End If
                                                If tmLLC(ilFdOpen).iEtfCode = 3 Then
                                                    tmOpenAvail.iRecType = Val(tmLLC(ilFdOpen).sType)
                                                    gPackTime tmLLC(ilFdOpen).sStartTime, tmOpenAvail.iTime(0), tmOpenAvail.iTime(1)
                                                    tmOpenAvail.iLtfCode = tmLLC(ilFdOpen).iLtfCode
                                                    tmOpenAvail.iAvInfo = 0
                                                    tmOpenAvail.iLen = 0
                                                    tmOpenAvail.ianfCode = Val(tmLLC(ilFdOpen).sName)
                                                    tmOpenAvail.iNoSpotsThis = 0
                                                    tmOpenAvail.iOrigUnit = 0
                                                    tmOpenAvail.iOrigLen = 0
                                                    gUnpackTimeLong tmOpenAvail.iTime(0), tmOpenAvail.iTime(1), False, llOpenTime '1-27-03
                                                    If ilAutomationType = 2 Then         'prophet NexGen only
                                                        '10-4-01  If we dont have to worry about open avails, remove this code
                                                        '1-27-03 create "Block" notation only if avails are not back-to-back
                                                        If (llDate <= llEDate + 1) Then    'And ilProphetFlag = True Then
                                                           LSet tmAvail = tmOpenAvail
                                                            '6/1/16: Replace GoSub
                                                            'GoSub lProphetBlockTime
                                                            mProphetBlockTime slDelimiter, ilAutomationType, slRecord, ilSameBreak, ilDlfFound, ilAdjZone, llAdjTime, slAirDate, llAirDate, llAvailTime, llSpotTime, blInProcSpot, ilVpfIndex, slDate, llDate, llSDate, llEDate, slAdvtName, slSortType, slSortDate, ilRet
                                                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                                            If llOpenTime + 15 >= llTestTime Then
                                                                If llTestTime > llPrevAvailEndTime Then
                                                                    llPrevAvailEndTime = llTestTime  '1-27-03 calc end time of current avail for test against next avails start time
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    If (ilAutomationType = 6) Or (ilAutomationType = 7) Then         'prophet Wizard, always create BLOCK events if avails are back-to-back
                                                        If llDate <= llEDate + 1 Then
                                                           LSet tmAvail = tmOpenAvail
                                                            '6/1/16: Replace GoSub
                                                            'GoSub lProphetBlockTime
                                                            mProphetBlockTime slDelimiter, ilAutomationType, slRecord, ilSameBreak, ilDlfFound, ilAdjZone, llAdjTime, slAirDate, llAirDate, llAvailTime, llSpotTime, blInProcSpot, ilVpfIndex, slDate, llDate, llSDate, llEDate, slAdvtName, slSortType, slSortDate, ilRet
                                                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                                        End If
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilFdOpen
                                            For ilFdClose = ilLoop To UBound(tmLLC) - 1 Step 1
                                                If tmLLC(ilFdClose).iEtfCode = 2 Then
                                                    llTestTime = gTimeToLong(tmLLC(ilFdClose).sStartTime, False)
                                                    llTestTime = llTestTime + gLengthToLong(tmLLC(ilFdClose).sLength)
                                                End If
                                                If tmLLC(ilFdClose).iEtfCode = 5 Then
                                                    tmCloseAvail.iRecType = Val(tmLLC(ilFdClose).sType)
                                                    gPackTime tmLLC(ilFdClose).sStartTime, tmCloseAvail.iTime(0), tmCloseAvail.iTime(1)
                                                    tmCloseAvail.iLtfCode = tmLLC(ilFdClose).iLtfCode
                                                    tmCloseAvail.iAvInfo = 0
                                                    tmCloseAvail.iLen = 0
                                                    tmCloseAvail.ianfCode = Val(tmLLC(ilFdClose).sName)
                                                    tmCloseAvail.iNoSpotsThis = 0
                                                    tmCloseAvail.iOrigUnit = 0
                                                    tmCloseAvail.iOrigLen = 0
                                                    gUnpackTimeLong tmCloseAvail.iTime(0), tmCloseAvail.iTime(1), False, llCloseTime '1-27-03
                                                    If ilAutomationType = 2 Then         'prophet NexGen only
                                                        '10-4-01  If we dont have to worry about open avails, remove this code
                                                        '1-27-03 create "Block" notation only if avails are not back-to-back
                                                        If (llDate <= llEDate + 1) And (llCloseTime > llTestTime) Then   'And ilProphetFlag = True Then
                                                           LSet tmAvail = tmCloseAvail
                                                            '6/1/16: Replace GoSub
                                                            'GoSub lProphetBlockTime
                                                            mProphetBlockTime slDelimiter, ilAutomationType, slRecord, ilSameBreak, ilDlfFound, ilAdjZone, llAdjTime, slAirDate, llAirDate, llAvailTime, llSpotTime, blInProcSpot, ilVpfIndex, slDate, llDate, llSDate, llEDate, slAdvtName, slSortType, slSortDate, ilRet
                                                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                                        End If
                                                    End If
                                                    If (ilAutomationType = 6) Or (ilAutomationType = 7) Then         'prophet Wizard, always create BLOCK events if avails are back-to-back
                                                        If llDate <= llEDate + 1 Then
                                                           LSet tmAvail = tmCloseAvail
                                                            '6/1/16: Replace GoSub
                                                            'GoSub lProphetBlockTime
                                                            mProphetBlockTime slDelimiter, ilAutomationType, slRecord, ilSameBreak, ilDlfFound, ilAdjZone, llAdjTime, slAirDate, llAirDate, llAvailTime, llSpotTime, blInProcSpot, ilVpfIndex, slDate, llDate, llSDate, llEDate, slAdvtName, slSortType, slSortDate, ilRet
                                                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                                        End If
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilFdClose
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If
                            ilAirHour = tmAvail.iTime(1) \ 256  'Obtain Hour
                            If (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
                                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAvailStartTime '1-27-03
                                'create the "Break record" with the description field
                                'if Prophet, create a "BLOCK" notation record for each new avail   '05-03-01
                                If ilAutomationType = 2 Then          'prophet NexGen or Audio Vault Sat
                                    '10-4-01  If we dont have to worry about open avails, remove this code
                                    '1-27-03 create "Block" notation only if avails are not back-to-back
                                    If (llDate <= llEDate + 1) And (llAvailStartTime > llPrevAvailEndTime) Then  'test back to back avails based on len of avail
                                        '6/1/16: Replace GoSub
                                        'GoSub lProphetBlockTime
                                        mProphetBlockTime slDelimiter, ilAutomationType, slRecord, ilSameBreak, ilDlfFound, ilAdjZone, llAdjTime, slAirDate, llAirDate, llAvailTime, llSpotTime, blInProcSpot, ilVpfIndex, slDate, llDate, llSDate, llEDate, slAdvtName, slSortType, slSortDate, ilRet
                                    End If
                                    'ilProphetFlag = True
                                ElseIf ilAutomationType = 9 Then
                                    '10-4-01  If we dont have to worry about open avails, remove this code
                                    '1-27-03 create "Block" notation only if avails are not back-to-back
                                    If (llDate <= llEDate + 1) And (llAvailStartTime > llSpotEndTime) Then    'test back to back avails based on spot lengths within the avail to next avail
                                        '6/1/16: Replace GoSub
                                        'GoSub lProphetBlockTime
                                        mProphetBlockTime slDelimiter, ilAutomationType, slRecord, ilSameBreak, ilDlfFound, ilAdjZone, llAdjTime, slAirDate, llAirDate, llAvailTime, llSpotTime, blInProcSpot, ilVpfIndex, slDate, llDate, llSDate, llEDate, slAdvtName, slSortType, slSortDate, ilRet
                                    End If
                                End If
                                If (ilAutomationType = 6) Or (ilAutomationType = 7) Then         'prophet Wizard, always create BLOCK events if avails are back-to-back
                                    If llDate <= llEDate + 1 Then
                                        '6/1/16: Replace GoSub
                                        'GoSub lProphetBlockTime
                                        mProphetBlockTime slDelimiter, ilAutomationType, slRecord, ilSameBreak, ilDlfFound, ilAdjZone, llAdjTime, slAirDate, llAirDate, llAvailTime, llSpotTime, blInProcSpot, ilVpfIndex, slDate, llDate, llSDate, llEDate, slAdvtName, slSortType, slSortDate, ilRet
                                    End If
                                    'ilProphetFlag = True
                                End If
                                llPrevAvailEndTime = llAvailStartTime + tmAvail.iLen  '1-27-03 calc end time of current avail for test against next avails start time
                             End If
                            If (llDate > llEDate) And (ilAirHour >= 6) Then
                                'Done processing
                                mExptGenDay = True
                                Exit Function
                            End If
                            ilSubFeed = False
                            ilDlfExist = False
                            'If (tmVef.sType = "A") Or (ilDlfFound) Then
                            If (ilDlfFound) Then
                                'Obtain engineering entry to see if avail is sent
                                tmDlfSrchKey.iVefCode = ilVehCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                    ilDlfExist = True
                                    ilTerminated = False
                                    If (tmDlf.sFed = "N") Or (tmDlf.sZone <> "CST") Then
                                        ilTerminated = True
                                    Else
                                        If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                            If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                    End If
                                    If Not ilTerminated Then
                                        If tmDlf.iMnfSubFeed > 0 Then
                                            ilSubFeed = True
                                            Exit Do
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            Else
                                ilDlfExist = True   'Set to aviod error message that link is missing
                            End If
                            'If Not ilDlfExist Then
                            If (Not ilDlfExist) And (tgSpf.sSDelNet = "Y") Then
                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                ilRet = MsgBox(smVehName & " missing Delivery link on " & slDate & " at " & slTime, vbOKCancel + vbQuestion + vbApplicationModal, "Find Error")
                                If ilRet = 2 Then   'Cancel
                                    mExptGenDay = False
                                    Exit Function
                                End If
                            Else
                                'Loop on spots, then add conflicting spots
                                If (tmVef.sType = "A") Then
                                    tmVlfSrchKey1.iAirCode = ilVehCode
                                    tmVlfSrchKey1.iAirDay = Val(slDay)
                                    tmVlfSrchKey1.iEffDate(0) = ilVlfDate0
                                    tmVlfSrchKey1.iEffDate(1) = ilVlfDate1
                                    tmVlfSrchKey1.iAirTime(0) = tmAvail.iTime(0)
                                    tmVlfSrchKey1.iAirTime(1) = tmAvail.iTime(1)
                                    tmVlfSrchKey1.iAirPosNo = 0
                                    tmVlfSrchKey1.iAirSeq = 1
                                    ilRet = btrGetGreaterOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) And (tmVlf.iEffDate(0) = ilVlfDate0) And (tmVlf.iEffDate(1) = ilVlfDate1) And (tmVlf.iAirTime(0) = tmAvail.iTime(0)) And (tmVlf.iAirTime(1) = tmAvail.iTime(1))
                                        ilTerminated = False
                                        If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                            If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                        If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                            If (tmCTSsf.iType <> ilType) Or (tmCTSsf.iVefCode <> tmVlf.iSellCode) Or (tmCTSsf.iDate(0) <> ilLogDate0) Or (tmCTSsf.iDate(1) <> ilLogDate1) Then
                                                tmVefSrchKey.iCode = tmVlf.iSellCode
                                                ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                If (ilRet = BTRV_ERR_NONE) Then
                                                    ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                    tlSsfSrchKey.iType = ilType
                                                    tlSsfSrchKey.iVefCode = tmVlf.iSellCode
                                                    tlSsfSrchKey.iDate(0) = ilLogDate0
                                                    tlSsfSrchKey.iDate(1) = ilLogDate1
                                                    tlSsfSrchKey.iStartTime(0) = 0
                                                    tlSsfSrchKey.iStartTime(1) = 0
                                                    ilRet = gSSFGetEqual(hmCTSsf, tmCTSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                End If
                                            End If
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmCTSsf.iType = ilType) And (tmCTSsf.iVefCode = tmVlf.iSellCode) And (tmCTSsf.iDate(0) = ilLogDate0) And (tmCTSsf.iDate(1) = ilLogDate1)
                                                For ilSIndex = 1 To tmCTSsf.iCount Step 1
                                                    tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSIndex)
                                                    If ((tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9)) Then
                                                        If (tmAvailTest.iTime(0) = tmVlf.iSellTime(0)) And (tmAvailTest.iTime(1) = tmVlf.iSellTime(1)) Then
                                                            'Determine if any unsold time
                                                            ilAvailUnits = tmAvailTest.iAvInfo And &H1F
                                                            ilAvailLen = tmAvailTest.iLen
                                                            ilSpotLen = 0
                                                            For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                               LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                                                    ilSpotLen = ilSpotLen + (tmSpot.iPosLen And &HFFF)
                                                                    ilSpotUnits = ilSpotUnits + 1
                                                                End If
                                                            Next ilSpot
                                                            If ilAvailLen > ilSpotLen Then
                                                                If ilDlfFound Then
                                                                    tmDlfSrchKey.iVefCode = ilVehCode
                                                                    tmDlfSrchKey.sAirDay = slDay
                                                                    tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                    tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                    tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                    tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                    ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                        ilTerminated = False
                                                                        If (tmDlf.sFed = "N") Or (StrComp(Trim$(tmDlf.sZone), "CST", 1) <> 0) Then
                                                                            ilTerminated = True
                                                                        Else
                                                                            If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                                If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                    ilTerminated = True
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        If Not ilTerminated Then
                                                                            If ilSubFeed Then
                                                                                If tmDlf.iMnfSubFeed > 0 Then
                                                                                    '6/1/16: Replace GoSub
                                                                                    'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                                    mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour, blInProcSpot, ilVpfIndex, llEDate
                                                                                    If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                                        slMsg = "Unsold: " & Trim$(tmSVef.sName)
                                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                        slMsg = slMsg & " at " & slTime
                                                                                        tmAnfSrchKey.iCode = tmAvailTest.ianfCode
                                                                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                        If (ilRet = BTRV_ERR_NONE) Then
                                                                                            slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                                        End If
                                                                                        'Print #hmMsg, slMsg
                                                                                        gAutomationAlertAndLogHandler slMsg
                                                                                        lbcMsg.AddItem slMsg
                                                                                    End If
                                                                                    DoEvents
                                                                                    If ilRet <> 0 Then
                                                                                        mExptGenDay = False
                                                                                        Exit Function
                                                                                    End If
                                                                                    Exit Do
                                                                                End If
                                                                            Else
                                                                                If tmDlf.iMnfSubFeed = 0 Then
                                                                                    '6/1/16: Replace GoSub
                                                                                    'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                                    mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour, blInProcSpot, ilVpfIndex, llEDate
                                                                                    If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                                        slMsg = "Unsold: " & Trim$(tmSVef.sName)
                                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                        slMsg = slMsg & " at " & slTime
                                                                                        tmAnfSrchKey.iCode = tmAvailTest.ianfCode
                                                                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                        If (ilRet = BTRV_ERR_NONE) Then
                                                                                            slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                                        End If
                                                                                        'Print #hmMsg, slMsg
                                                                                        gAutomationAlertAndLogHandler slMsg
                                                                                        lbcMsg.AddItem slMsg
                                                                                    End If
                                                                                    DoEvents
                                                                                    If ilRet <> 0 Then
                                                                                        mExptGenDay = False
                                                                                        Exit Function
                                                                                    End If
                                                                                    Exit Do
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                    Loop
                                                                Else
                                                                    tmDlf.iAirTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iAirTime(1) = tmAvail.iTime(1)
                                                                    tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                    tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                    tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                    tmDlf.sZone = ""
                                                                    tmDlf.iEtfCode = 0
                                                                    tmDlf.iEnfCode = 0
                                                                    tmDlf.sProgCode = ""
                                                                    tmDlf.iMnfFeed = 0
                                                                    tmDlf.sBus = ""
                                                                    tmDlf.sSchedule = ""
                                                                    tmDlf.iMnfSubFeed = 0
                                                                    '6/1/16: Replace GoSub
                                                                    'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                    mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour, blInProcSpot, ilVpfIndex, llEDate
                                                                    If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                        slMsg = "Unsold: " & Trim$(tmSVef.sName)
                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                        slMsg = slMsg & " at " & slTime
                                                                        tmAnfSrchKey.iCode = tmAvailTest.ianfCode
                                                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                        If (ilRet = BTRV_ERR_NONE) Then
                                                                            slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                        End If
                                                                        'Print #hmMsg, slMsg
                                                                        gAutomationAlertAndLogHandler slMsg
                                                                        lbcMsg.AddItem slMsg
                                                                    Else
                                                                        ilRet = 0
                                                                    End If
                                                                    DoEvents
                                                                    If ilRet <> 0 Then
                                                                        mExptGenDay = False
                                                                        Exit Function
                                                                    End If
                                                                End If
                                                            End If
                                                            '???????
                                                            tlBBAvail = tmAvail
                                                            gUnpackTimeLong tmAvailTest.iTime(0), tmAvailTest.iTime(1), False, llBBTime
                                                            mCheckEmptyAvail ilAutomationType, llDate, llBBTime, tmAvailTest.iNoSpotsThis      'if avail is totally empty, audio vault needs an entry

                                                            For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1       'airing vehicle
                                                               LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                smMediaCodeSuppressSpot = "N"
                                                                '3-23-09 if including primary split network spot and its a primary, include it
                                                                'or if including secondary split network spot and its a secondary spot, include it
                                                                'if not a split network spot, always include it
                                                                ilIncludeFullOrSplit = False
                                                                If ((rbcSplit(0).Value = True) And (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI) Or ((rbcSplit(1).Value = True) And (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC) Or (((tmSpot.iRecType And SSSPLITPRI) <> SSSPLITPRI) And ((tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC)) Then
                                                                    ilIncludeFullOrSplit = True
                                                                End If

                                                                If ilIncludeFullOrSplit Then
                                                                   tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                                   ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                   If ilRet = BTRV_ERR_NONE Then
                                                                       tlBBSdf = tmSdf
                                                                       tmChfSrchKey.lCode = tmSdf.lChfCode
                                                                       ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                       If ilRet = BTRV_ERR_NONE Then
                                                                           If tgSpf.sUsingBBs = "Y" Or ilAutomationType = 2 Then       '4-7-06 using BB or prophet nextgen
                                                                               tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                                                               tmClfSrchKey.iLine = tmSdf.iLineNo
                                                                               tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                                                               tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                                                               imClfRecLen = Len(tmClf)
                                                                               ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                                                               Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))   'And (tmClf.sSchStatus = "A")
                                                                                   ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                               Loop
                                                                           End If
                                                                           tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                                           ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation

                                                                           If ilRet = BTRV_ERR_NONE Then
                                                                               '4-7-06 Gather extra prophet fields in necessary
                                                                               mGetProphetNextGenFields ilAutomationType, smAdvtCode, smCompCode, smPty, smFixed, smROS
                                                                               ilBBPass = 0
                                                                               Do
                                                                                   slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                                                   gGetAirCopy tmVef.sType, tmVef.iCode, ilVpfIndex, tmSdf, hmCrf, hmRsf, hmCvf, slZone
                                                                                   ilCopy = gObtainCopy(slZone, tmSdf, hmMcf, hmTzf, hmCif, hmCpf, ilWegenerOLA, smCifName, smCreativeTitle, imMediaCodeLen, smISCI, smMcfPrefix, smMediaCodeSuppressSpot, slMcfCode)
                                                                                   If Not ilCopy Then
                                                                                       slMsg = "Copy Missing: " & Trim$(tmSVef.sName)
                                                                                       gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                       slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                                       slMsg = slMsg & " for " & Trim$(str$(tmChf.lCntrNo)) & " " & Trim$(tmAdf.sName)
                                                                                       'Print #hmMsg, slMsg
                                                                                       'lbcMsg.AddItem slMsg
                                                                                   End If
                                                                                   
                                                                                   If ilDlfFound Then
                                                                                       tmDlfSrchKey.iVefCode = ilVehCode
                                                                                       tmDlfSrchKey.sAirDay = slDay
                                                                                       tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                                       tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                                       tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                                       tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                                       ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                       Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                                           ilTerminated = False
                                                                                           If (tmDlf.sFed = "N") Or (StrComp(Trim$(tmDlf.sZone), "CST", 1) <> 0) Then
                                                                                               ilTerminated = True
                                                                                           Else
                                                                                               If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                                                   If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                                       ilTerminated = True
                                                                                                   End If
                                                                                               End If
                                                                                           End If
                                                                                           If Not ilTerminated Then
                                                                                               If ilSubFeed Then
                                                                                                   If tmDlf.iMnfSubFeed > 0 Then
                                                                                                        '6/1/16: Replaced GoSub
                                                                                                        'GoSub lProcSpot
                                                                                                        mProcSpot ilAutomationType, ilAdjZone, llAdjTime, ilDlfFound, slAirDate, llAirDate, llAvailTime, llSpotTime, llPrevTime, ilShowTime, llSpotEndTime, blInProcSpot, ilVefCode, ilVehCode, ilVpfIndex, slDate, llDate, llSDate, llEDate, ilAirHour, ilLocalHour, ilCopy, llCopyMissingSdfCode, slSortType, slSortDate, slRecord, slLogYear, slLogMonth, slLogDay, slEventName, llESPNPrevDate, ilESPNHour, ilESPNBreak, ilESPNPosition, llESPNPrevAvailTime, ilParentVefCode, ilAtLeastOneSpot, ilSameBreak, ilRet, slMsg
                                                                                                        DoEvents
                                                                                                        If ilRet <> 0 Then
                                                                                                           mExptGenDay = False
                                                                                                           Exit Function
                                                                                                        End If
                                                                                                   End If
                                                                                               Else
                                                                                                   If tmDlf.iMnfSubFeed = 0 Then
                                                                                                        '6/1/16: Replaced GoSub
                                                                                                        'GoSub lProcSpot
                                                                                                        mProcSpot ilAutomationType, ilAdjZone, llAdjTime, ilDlfFound, slAirDate, llAirDate, llAvailTime, llSpotTime, llPrevTime, ilShowTime, llSpotEndTime, blInProcSpot, ilVefCode, ilVehCode, ilVpfIndex, slDate, llDate, llSDate, llEDate, ilAirHour, ilLocalHour, ilCopy, llCopyMissingSdfCode, slSortType, slSortDate, slRecord, slLogYear, slLogMonth, slLogDay, slEventName, llESPNPrevDate, ilESPNHour, ilESPNBreak, ilESPNPosition, llESPNPrevAvailTime, ilParentVefCode, ilAtLeastOneSpot, ilSameBreak, ilRet, slMsg
                                                                                                       DoEvents
                                                                                                       If ilRet <> 0 Then
                                                                                                           mExptGenDay = False
                                                                                                           Exit Function
                                                                                                       End If
                                                                                                   End If
                                                                                               End If
                                                                                           End If
                                                                                           ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                                       Loop
                                                                                       ilBBPass = 2
                                                                                   Else
                                                                                       tmDlf.iAirTime(0) = tmAvail.iTime(0)
                                                                                       tmDlf.iAirTime(1) = tmAvail.iTime(1)
                                                                                       tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                                       tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                                       tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                                       tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                                       tmDlf.sZone = ""
                                                                                       tmDlf.iEtfCode = 0
                                                                                       tmDlf.iEnfCode = 0
                                                                                       tmDlf.sProgCode = ""
                                                                                       tmDlf.iMnfFeed = 0
                                                                                       tmDlf.sBus = ""
                                                                                       tmDlf.sSchedule = ""
                                                                                       tmDlf.iMnfSubFeed = 0
                                                                                        '6/1/16: Replaced GoSub
                                                                                        'GoSub lProcSpot
                                                                                        mProcSpot ilAutomationType, ilAdjZone, llAdjTime, ilDlfFound, slAirDate, llAirDate, llAvailTime, llSpotTime, llPrevTime, ilShowTime, llSpotEndTime, blInProcSpot, ilVefCode, ilVehCode, ilVpfIndex, slDate, llDate, llSDate, llEDate, ilAirHour, ilLocalHour, ilCopy, llCopyMissingSdfCode, slSortType, slSortDate, slRecord, slLogYear, slLogMonth, slLogDay, slEventName, llESPNPrevDate, ilESPNHour, ilESPNBreak, ilESPNPosition, llESPNPrevAvailTime, ilParentVefCode, ilAtLeastOneSpot, ilSameBreak, ilRet, slMsg
                                                                                       If (ilRet <> 0) Or (imTerminate = True) Then
                                                                                           mExptGenDay = False
                                                                                           Exit Function
                                                                                       End If
                                                                                   End If
                                                                                   tmSdf = tlBBSdf
                                                                                  LSet tmAvail = tlBBAvail
                                                                                   If (tgSpf.sUsingBBs <> "Y") Or (ilBBPass >= 2) Then
                                                                                       Exit Do
                                                                                   End If
                                                                                   ilBBLen = tmClf.iBBOpenLen
                                                                                   If (ilBBLen > 0) And (tmOpenAvail.iRecType <> -1) And (ilBBPass = 0) Then
                                                                                       'Find BB spot
                                                                                       ilFound = gFindBBSpot(hmSdf, "O", tmSdf.iVefCode, tmSdf.lChfCode, tmSdf.iLineNo, llDate, llBBTime, tlSdf, llPrevOpenFdBBSpots())
                                                                                       If Not ilFound Then
                                                                                           ilBBPass = 1
                                                                                       Else
                                                                                          LSet tmAvail = tmOpenAvail
                                                                                           tmSdf = tlSdf
                                                                                           tmSdf.iTime(0) = tmOpenAvail.iTime(0)
                                                                                           tmSdf.iTime(1) = tmOpenAvail.iTime(1)
                                                                                       End If
                                                                                   Else
                                                                                       ilBBPass = 1
                                                                                   End If
                                                                                   If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                                                                       ilBBLen = tmClf.iBBOpenLen
                                                                                   Else
                                                                                       ilBBLen = tmClf.iBBCloseLen
                                                                                   End If
                                                                                   If (ilBBLen > 0) And (tmCloseAvail.iRecType <> -1) And (ilBBPass = 1) Then
                                                                                       'Find BB spot
                                                                                       ilFound = gFindBBSpot(hmSdf, "C", tmSdf.iVefCode, tmSdf.lChfCode, tmSdf.iLineNo, llDate, llBBTime, tlSdf, llPrevCloseFdBBSpots())
                                                                                       If Not ilFound Then
                                                                                           Exit Do
                                                                                       Else
                                                                                          LSet tmAvail = tmCloseAvail
                                                                                           tmSdf = tlSdf
                                                                                           tmSdf.iTime(0) = tmCloseAvail.iTime(0)
                                                                                           tmSdf.iTime(1) = tmCloseAvail.iTime(1)
                                                                                       End If
                                                                                   Else
                                                                                       If ilBBPass = 1 Then
                                                                                           Exit Do
                                                                                       End If
                                                                                   End If
                                                                                   ilBBPass = ilBBPass + 1
                                                                               Loop
                                                                           Else
                                                                               Screen.MousePointer = vbDefault
                                                                               ''MsgBox Trim$(tmSVef.sName) & " advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                               gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                               Screen.MousePointer = vbHourglass
                                                                               mExptGenDay = False
                                                                               Exit Function
                                                                           End If
                                                                       Else
                                                                           Screen.MousePointer = vbDefault
                                                                           ''MsgBox Trim$(tmSVef.sName) & " contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                           gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                           Screen.MousePointer = vbHourglass
                                                                           mExptGenDay = False
                                                                           Exit Function
                                                                       End If
                                                                   Else
                                                                       Screen.MousePointer = vbDefault
                                                                       ''MsgBox Trim$(tmSVef.sName) & " spot get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                       gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " spot get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                       Screen.MousePointer = vbHourglass
                                                                       mExptGenDay = False
                                                                       Exit Function
                                                                   End If
                                                                End If          'ilIncludeFullOrSplit
                                                            Next ilSpot
                                                            Exit Do
                                                        End If
                                                    End If
                                                Next ilSIndex
                                                ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                ilRet = gSSFGetNext(hmCTSsf, tmCTSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Loop
                                        End If
                                        ilRet = btrGetNext(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                Else
                                    ilAvailUnits = tmAvail.iAvInfo And &H1F
                                    ilAvailLen = tmAvail.iLen
                                    ilSpotLen = 0
                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpot + ilEvt)
                                        If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                            ilSpotLen = ilSpotLen + (tmSpot.iPosLen And &HFFF)
                                            ilSpotUnits = ilSpotUnits + 1
                                        End If
                                    Next ilSpot
                                    If ilAvailLen > ilSpotLen Then              'test for time unsold
                                        If ilDlfFound Then
                                            'Obtain engineering entry to see is avail is sent
                                            tmDlfSrchKey.iVefCode = ilVehCode
                                            tmDlfSrchKey.sAirDay = slDay
                                            tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                            tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                            tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                            tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                            ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                ilTerminated = False
                                                If (tmDlf.sFed = "N") Or (StrComp(Trim$(tmDlf.sZone), "CST", 1) <> 0) Then
                                                    ilTerminated = True
                                                Else
                                                    If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                        If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                            ilTerminated = True
                                                        End If
                                                    End If
                                                End If
                                                If Not ilTerminated Then
                                                    If ilSubFeed Then
                                                        If tmDlf.iMnfSubFeed > 0 Then
                                                            '6/1/16: Replace GoSub
                                                            'GoSub lProcAdjDate  'Result stored into slAirDate
                                                            mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour, blInProcSpot, ilVpfIndex, llEDate
                                                            If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                If (ilRet = BTRV_ERR_NONE) Then
                                                                    slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                End If
                                                                'Print #hmMsg, slMsg
                                                                gAutomationAlertAndLogHandler slMsg
                                                                lbcMsg.AddItem slMsg
                                                            End If
                                                            DoEvents
                                                            If ilRet <> 0 Then
                                                                mExptGenDay = False
                                                                Exit Function
                                                            End If
                                                            Exit Do
                                                        End If
                                                    Else
                                                        If tmDlf.iMnfSubFeed = 0 Then
                                                            '6/1/16: Replace GoSub
                                                            'GoSub lProcAdjDate  'Result stored into slAirDate
                                                            mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour, blInProcSpot, ilVpfIndex, llEDate
                                                            If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                If (ilRet = BTRV_ERR_NONE) Then
                                                                    slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                End If
                                                                'Print #hmMsg, slMsg
                                                                gAutomationAlertAndLogHandler slMsg
                                                                lbcMsg.AddItem slMsg
                                                            End If
                                                            DoEvents
                                                            If ilRet <> 0 Then
                                                                mExptGenDay = False
                                                                Exit Function
                                                            End If
                                                            Exit Do
                                                        End If
                                                    End If
                                                End If
                                                ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Loop
                                        Else
                                            tmDlf.iAirTime(0) = tmAvail.iTime(0)
                                            tmDlf.iAirTime(1) = tmAvail.iTime(1)
                                            tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                            tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                            tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                            tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                            tmDlf.sZone = ""
                                            tmDlf.iEtfCode = 0
                                            tmDlf.iEnfCode = 0
                                            tmDlf.sProgCode = ""
                                            tmDlf.iMnfFeed = 0
                                            tmDlf.sBus = ""
                                            tmDlf.sSchedule = ""
                                            tmDlf.iMnfSubFeed = 0
                                            '6/1/16: Replace GoSub
                                            'GoSub lProcAdjDate  'Result stored into slAirDate
                                            mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour, blInProcSpot, ilVpfIndex, llEDate
                                            If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                If (ilRet = BTRV_ERR_NONE) Then
                                                    slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                End If
                                                'Print #hmMsg, slMsg
                                                gAutomationAlertAndLogHandler slMsg
                                                lbcMsg.AddItem slMsg
                                            Else
                                                ilRet = 0
                                            End If
                                            DoEvents
                                            If ilRet <> 0 Then
                                                mExptGenDay = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                    tlBBAvail = tmAvail
                                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llBBTime
                                    llSpotEndTime = llBBTime

                                    mCheckEmptyAvail ilAutomationType, llDate, llBBTime, tmAvail.iNoSpotsThis     'if avail is totally empty, audio vault needs an entry

                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1       'process spots for conventional vehicle
                                        ilEvt = ilEvt + 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                        smMediaCodeSuppressSpot = "N"
                                        '3-23-09 if including primary split network spot and its a primary, include it
                                        'or if including secondary split network spot and its a secondary spot, include it
                                        'if not a split network spot, always include it
                                        ilIncludeFullOrSplit = False
                                        If ((rbcSplit(0).Value = True) And (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI) Or ((rbcSplit(1).Value = True) And (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC) Or (((tmSpot.iRecType And SSSPLITPRI) <> SSSPLITPRI) And ((tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC)) Then
                                            ilIncludeFullOrSplit = True
                                        End If

                                        If ilIncludeFullOrSplit Then
                                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                tlBBSdf = tmSdf
                                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet = BTRV_ERR_NONE Then
                                                    If tgSpf.sUsingBBs = "Y" Or ilAutomationType = 2 Then       '4-7-06 using bb or Prophet nextgen
                                                        tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                                        tmClfSrchKey.iLine = tmSdf.iLineNo
                                                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                                        imClfRecLen = Len(tmClf)
                                                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                                                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Loop
                                                    End If
                                                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        '4-7-06 gather extra prophet nextgen fields if necessary
                                                        mGetProphetNextGenFields ilAutomationType, smAdvtCode, smCompCode, smPty, smFixed, smROS

                                                        ilBBPass = 0
                                                        Do
                                                            slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                            gGetAirCopy tmVef.sType, tmVef.iCode, ilVpfIndex, tmSdf, hmCrf, hmRsf, hmCvf, slZone
                                                            ilCopy = gObtainCopy(slZone, tmSdf, hmMcf, hmTzf, hmCif, hmCpf, ilWegenerOLA, smCifName, smCreativeTitle, imMediaCodeLen, smISCI, smMcfPrefix, smMediaCodeSuppressSpot, slMcfCode)
                                                            If Not ilCopy Then
                                                                slMsg = "Copy Missing: " & Trim$(smVehName)
                                                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                slMsg = slMsg & " for " & Trim$(str$(tmChf.lCntrNo)) & " " & Trim$(tmAdf.sName)
                                                                'Print #hmMsg, slMsg
                                                                'lbcMsg.AddItem slMsg
                                                            End If
                                                            If ilDlfFound Then
                                                                'Obtain engineering entry to see is avail is sent
                                                                tmDlfSrchKey.iVefCode = ilVehCode
                                                                tmDlfSrchKey.sAirDay = slDay
                                                                tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                ilRet = btrGetEqual(hmDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                    ilTerminated = False
                                                                    If (tmDlf.sFed = "N") Or (StrComp(Trim$(tmDlf.sZone), "CST", 1) <> 0) Then
                                                                        ilTerminated = True
                                                                    Else
                                                                        If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                            If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                ilTerminated = True
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    If Not ilTerminated Then
                                                                        If ilSubFeed Then
                                                                            If tmDlf.iMnfSubFeed > 0 Then
                                                                                '6/1/16: Replaced GoSub
                                                                                'GoSub lProcSpot
                                                                                mProcSpot ilAutomationType, ilAdjZone, llAdjTime, ilDlfFound, slAirDate, llAirDate, llAvailTime, llSpotTime, llPrevTime, ilShowTime, llSpotEndTime, blInProcSpot, ilVefCode, ilVehCode, ilVpfIndex, slDate, llDate, llSDate, llEDate, ilAirHour, ilLocalHour, ilCopy, llCopyMissingSdfCode, slSortType, slSortDate, slRecord, slLogYear, slLogMonth, slLogDay, slEventName, llESPNPrevDate, ilESPNHour, ilESPNBreak, ilESPNPosition, llESPNPrevAvailTime, ilParentVefCode, ilAtLeastOneSpot, ilSameBreak, ilRet, slMsg
                                                                                DoEvents
                                                                                If ilRet <> 0 Then
                                                                                    mExptGenDay = False
                                                                                    Exit Function
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            If tmDlf.iMnfSubFeed = 0 Then
                                                                                '6/1/16: Replaced GoSub
                                                                                'GoSub lProcSpot
                                                                                mProcSpot ilAutomationType, ilAdjZone, llAdjTime, ilDlfFound, slAirDate, llAirDate, llAvailTime, llSpotTime, llPrevTime, ilShowTime, llSpotEndTime, blInProcSpot, ilVefCode, ilVehCode, ilVpfIndex, slDate, llDate, llSDate, llEDate, ilAirHour, ilLocalHour, ilCopy, llCopyMissingSdfCode, slSortType, slSortDate, slRecord, slLogYear, slLogMonth, slLogDay, slEventName, llESPNPrevDate, ilESPNHour, ilESPNBreak, ilESPNPosition, llESPNPrevAvailTime, ilParentVefCode, ilAtLeastOneSpot, ilSameBreak, ilRet, slMsg
                                                                                DoEvents
                                                                                If ilRet <> 0 Then
                                                                                    mExptGenDay = False
                                                                                    Exit Function
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    ilRet = btrGetNext(hmDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                Loop
                                                                ilBBPass = 2
                                                            Else
                                                                tmDlf.iAirTime(0) = tmAvail.iTime(0)
                                                                tmDlf.iAirTime(1) = tmAvail.iTime(1)
                                                                tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                tmDlf.sZone = ""
                                                                tmDlf.iEtfCode = 0
                                                                tmDlf.iEnfCode = 0
                                                                tmDlf.sProgCode = ""
                                                                tmDlf.iMnfFeed = 0
                                                                tmDlf.sBus = ""
                                                                tmDlf.sSchedule = ""
                                                                tmDlf.iMnfSubFeed = 0
                                                                '6/1/16: Replaced GoSub
                                                                'GoSub lProcSpot
                                                                mProcSpot ilAutomationType, ilAdjZone, llAdjTime, ilDlfFound, slAirDate, llAirDate, llAvailTime, llSpotTime, llPrevTime, ilShowTime, llSpotEndTime, blInProcSpot, ilVefCode, ilVehCode, ilVpfIndex, slDate, llDate, llSDate, llEDate, ilAirHour, ilLocalHour, ilCopy, llCopyMissingSdfCode, slSortType, slSortDate, slRecord, slLogYear, slLogMonth, slLogDay, slEventName, llESPNPrevDate, ilESPNHour, ilESPNBreak, ilESPNPosition, llESPNPrevAvailTime, ilParentVefCode, ilAtLeastOneSpot, ilSameBreak, ilRet, slMsg
                                                                If (ilRet <> 0) Or (imTerminate = True) Then
                                                                    mExptGenDay = False
                                                                    Exit Function
                                                                End If
                                                            End If
                                                            tmSdf = tlBBSdf
                                                           LSet tmAvail = tlBBAvail
                                                            If (tgSpf.sUsingBBs <> "Y") Or (ilBBPass >= 2) Then
                                                                Exit Do
                                                            End If
                                                            ilBBLen = tmClf.iBBOpenLen
                                                            If (ilBBLen > 0) And (tmOpenAvail.iRecType <> -1) And (ilBBPass = 0) Then
                                                                'Find BB spot
                                                                ilFound = gFindBBSpot(hmSdf, "O", tmSdf.iVefCode, tmSdf.lChfCode, tmSdf.iLineNo, llDate, llBBTime, tlSdf, llPrevOpenFdBBSpots())
                                                                If Not ilFound Then
                                                                    ilBBPass = 1
                                                                Else
                                                                   LSet tmAvail = tmOpenAvail
                                                                    tmSdf = tlSdf
                                                                    tmSdf.iTime(0) = tmOpenAvail.iTime(0)
                                                                    tmSdf.iTime(1) = tmOpenAvail.iTime(1)
                                                                End If
                                                            Else
                                                                ilBBPass = 1
                                                            End If
                                                            If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                                                ilBBLen = tmClf.iBBOpenLen
                                                            Else
                                                                ilBBLen = tmClf.iBBCloseLen
                                                            End If
                                                            If (ilBBLen > 0) And (tmCloseAvail.iRecType <> -1) And (ilBBPass = 1) Then
                                                                'Find BB spot
                                                                ilFound = gFindBBSpot(hmSdf, "C", tmSdf.iVefCode, tmSdf.lChfCode, tmSdf.iLineNo, llDate, llBBTime, tlSdf, llPrevCloseFdBBSpots())
                                                                If Not ilFound Then
                                                                    Exit Do
                                                                Else
                                                                   LSet tmAvail = tmCloseAvail
                                                                    tmSdf = tlSdf
                                                                    tmSdf.iTime(0) = tmCloseAvail.iTime(0)
                                                                    tmSdf.iTime(1) = tmCloseAvail.iTime(1)
                                                                End If
                                                            Else
                                                                If ilBBPass = 1 Then
                                                                    Exit Do
                                                                End If
                                                            End If
                                                            ilBBPass = ilBBPass + 1
                                                        Loop
                                                    Else
                                                        Screen.MousePointer = vbDefault
                                                        ''MsgBox smVehName & " advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                        gAutomationAlertAndLogHandler smVehName & " advertiser get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                        Screen.MousePointer = vbHourglass
                                                        mExptGenDay = False
                                                        Exit Function
                                                    End If
                                                Else
                                                    Screen.MousePointer = vbDefault
                                                    ''MsgBox smVehName & " contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                    gAutomationAlertAndLogHandler smVehName & " contract get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                    Screen.MousePointer = vbHourglass
                                                    mExptGenDay = False
                                                    Exit Function
                                                End If
                                            Else
                                                Screen.MousePointer = vbDefault
                                                ''MsgBox smVehName & " spot get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                gAutomationAlertAndLogHandler smVehName & " spot get error", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                Screen.MousePointer = vbHourglass
                                                mExptGenDay = False
                                                Exit Function
                                            End If
                                        End If              'ilIncludeFullOrSplit
                                    Next ilSpot
                                End If
                            End If
                       End If              'tmProg.iRecType = 1 elseif tmProg.iRecType >= 2 and tmProg.iRecType <= 9
                    End If                  'if within time                    ilEvt = ilEvt + 1
                    ilEvt = ilEvt + 1
                Loop
                
                'if the day/game didnt have any spots, make it as complete
                If Not ilAtLeastOneSpot Then
                    gUpdateLCFCompleteFlag hmLcf, tmLcf, ilSsfDate0, ilSsfDate1, ilType, ilVehCode, "C"
                End If
                
                ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tgMVef(ilVefIndex).sType = "G" Then
                    If ilGameNo <> 0 Then
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVehCode)
                            If tmSsf.iType = ilGameNo Then
                                Exit Do
                            End If
                            ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                            ilRet = gSSFGetNext(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                    ilType = tmSsf.iType
                End If
            Loop        '(ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
            If ilAutomationType = 14 Then
                'process the PRS comment events
                mProcRPSEvent llDate, llStartTime, llEndTime, tmLLC
            End If
        End If          '(ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode)
                   
        'for Audio Vault, create the the reqquired jump record
        mCreateJump ilAutomationType, llDate
        If (Not ilDlfFound) And (llDate = llEDate) And (rbcZone(0).Value) Then
            Exit For
        End If
    Next llDate         'for lldate =llSDate To llEDate + 1
    Erase tmBBSdfInfo
    'Erase tlLLC
    Erase llPrevOpenFdBBSpots
    Erase llPrevCloseFdBBSpots
    mExptGenDay = True
    Exit Function
'lProcSpot:
'    'Record format for Prophet(wizard & nextgen), Dalet (Fixed Length):
'    'Column  Length  Field
'    '  1        8    Start Time HR:MN:SC  Military Hours
'    '  9       20    Cart Number (w/ material code) or ISCI code
'    ' 29       46    Advertiser Name (30) / Product name (15)
'    ' 75       20    Creative Title
'    ' 95        4    Length  MMSS
'    ' 99       10    SDF Code
'    '109        1    c/r
'    '110        1    l/f
'
'   'Record format for StationPlaylist 5-11-18:
'    'Column  Length  Field
'    '  1        8    Start Time HR:MN:SC  Military Hours
'    '  9       20    Cart Number (w/ material code) or ISCI code
'    ' 29       46    Advertiser Name (30) / Product name (15)
'    ' 75       20    Creative Title
'    ' 95        4    Length  MMSS
'    ' 99       10    SDF Code
'    '109       20    ISCI
'    '129        1    c/r
'    '130        1    l/f
'    'following 5 fields for nextgen ONLY)  4-7-06
'    '109        1    blank
'    '110        1    ROS flag (R/N)
'    '111        1    blank
'    '112        6    5-10-06 Pty code was Fixed buy (F/M)
'    '118        1    blank
'    '119        1    5-10-06 fixed buy, was Pry code
'    '120        1    blank
'    '121        5    Advertiser code
'    '126        1    blank
'    '127        5    competitive code (chfcompcode(0))
'    '132        1    c/r
'    '133        1    l/f
'
'
'    'Record format for Prophet, MediaStar:
'    'Column  Length  Field
'    '  6        8    Start Time HR:MN:SC  Military Hours
'    ' 15        5    Cart Number (w/o material code)
'    ' 21       10    Spot ID
'    ' 39       29    Advertiser Name (30)/Product (remaining characters) , plus slash
'    ' 69        1    * asterisk to indicate that this spot is from CSI
'    ' 70        4    Length  SSSS
'    '           1    c/r
'    '           1    l/f
'
' 'Record format for RCS (Fixed length):
'    'Column  Length  Field
'    '  1        1    Commercial Indicator always C
'    '  2        7    Start Time HRMN:SC  Military Hours
'    '  9        4    Cart Number (without material code) or first 4 char of ISCI
'    ' 13       24    Advertiser Abbr/Product
'    ' 37        3    Unused (Priority number)
'    ' 40        4    Length  MNSC (60=> 0100; 90=> 0130; 30=>0030)
'    ' 44        4    Unused (Commercial Type)
'    ' 48        6    First Part of Sdf.lCode (was Unused (Customer ID))
'    ' 54        2    Second Part of Sdf.lCode (was Unused (Internal Code))
'    '                use 8 bytes to make up sdd.lcode
'    ' 56        4    Unused (Product Code)
'    ' 60        8    Unused (Ordered Time)
'    ' 68        1    Carriage Return <cr>
'    ' 69        1    Line Feed <lf>
'    '
'
''Record format for RCS-5 (Fixed length):
'    'Column  Length  Field
'    '  1        1    Commercial Indicator always C
'    '  2        7    Start Time HRMN:SC  Military Hours
'    '  9       15    Cart Number (without material code) or first 5 char of ISCI
'    ' 24       24    Advertiser Abbr/Product
'    ' 48        3    Unused (Priority number)
'    ' 51        4    Length  MNSC (60=> 0100; 90=> 0130; 30=>0030)
'    ' 55        4    Unused (Commercial Type)
'    ' 59        6    First Part of Sdf.lCode (was Unused (Customer ID))
'    ' 65        2    Second Part of Sdf.lCode (was Unused (Internal Code))
'    '                use 8 bytes to make up sdd.lcode
'    ' 67        4    Unused (Product Code)
'    ' 71        8    Unused (Ordered Time)
'    ' 79       24    Unused (Sponsor 1)
'    '103       24    Unused (Sponsor 2)
'    '127       24    Unused (Comment)
'    '151        8    Unused (Start Date mm/dd/yy)
'    '159        5    Unused (Start time hh:mm)
'    '164        8    Unused (Stop Date)
'    '172        5    Unused (Stop Time)
'    '177        8    Unused (Kill Date)
'    '185       15    Unused (Product 1)
'    '200       15    Unused (Product 2)
'    '215        1    Unused (Is Live? Y or N)
'    '216        1    Unused (Is External? Y or N)
'    '217        2    Unused (CBSI Stopset number)
'    '219       35    Unused (Native-Specific)
'    '254        1    Carriage Return <cr>
'    '255        1    Line Feed <lf>
'
''Record Format for Scott
'    ' Field #     Max length   Description
'    '  1              8         Start Time HR:MN:SC Military time
'    '  2              --        N/a
'    '  3              3         Comml notation (CA)
'    '  4              6         DA followed for cart # without material code
'    '  5              66        Advt (max 30) / Prod (max 35) (in quotes)
'    '  6              10        SDF Spot Code (in quotes)
'    '  7              5         Length of spot MM:SS
'
''Record Format for Scott V5 (added 8-16-13)
'    ' Field #     Max length   Description
'    '  1              8         Start Time HR:MN:SC Military time
'    '  2              --        N/a
'    '  3              3         Comml notation (Material Code)
'    '  4              6         cart # without material code
'    '  5              66        Advt (max 30) / Prod (max 35) (in quotes)
'    '  6              10        SDF Spot Code (in quotes)
'    '  7              5         N/A
'    '  9                        N/A
'    '  10                       N/A
'    '  11             6         Must be 6 blanks enclosed in "
'    '  12                       N/A
'    '  13                       N/A
'    '  14                       N/A
'    '  15                       N/A
'    '  16                       N/A
'
''Record Format for Wide Orbit (comma delimited, variable length) 1-10-12
'    ' Field #     Max length   Description
'    '  1              8         Start Time HR:MN:SC Military time
'    '  2              --        N/a
'    '  3              3         Comml notation (CA)
'    '  4              6         DA followed for cart # without material code
'    '  5              45        Advt (max 30) / Prod (max 35) (in quotes)
'    '  6              21        SDF Spot Code followed by ":", then station code all in quotes 1-10-12 station code added
'    '  7              5         Length of spot MM:SS
'
' 'Record format for iMedia Touch (fixed length, blank after each field)
' '   Field #       Max Length    Description
' '    1-8              8          Event Time HH:MM:SS   24hr time
' '    9                1          Blank
' '    10-14            5          Event Duration MM:SS
' '    15               1          Blank
' '    16-21            6          ZM#### ZM (Hard-coded) followed by cart#
' '    22               1          Blank
' '    23-25            3          Event index (Unused)
' '    26               1          Blank
' '    27-29            3          COM  (Hard-coded)
' '    30               1          Blank
' '    31-38            8          Live Copy identifier (unused)
' '    39               1          Blank
' '    40               1          Sync Char (unused)
' '    41               1          Blank
' '    42               1          Item function (unused)
' '    43               1          Blank
' '    44-73            30         20 char advertisr/prod, 10 char internal spot code
' '    74               1          Blank
' '    75               1          c/r
' '    76               1          linefeed
'
' 'Record Format for Audio Vault Sat spot (vertical line delimited, variable length)
'    ' Field #     Max length   Description
'    '  1              8         Start Time HR:MN:SC Military time
'    '  2              |
'    '  3              1         blank or + (for other than 1st spot in break)
'    '  4              |
'    '  4              10        Cart # (with or without media definition based on the site option)
'    '  5              |
'    '  6              40       Advt/Prod (optional)
'    '  7              |
'
'    '11-17-10 Audio Vault RPS.  There are 2 record types:  Comments and spots
'    ' Field #     Max length   Description
'    ' Comment type record
'    '  1              |
'    '  2              8         Start Time HR:MN:SC Military time
'    '  3              |         blank or + (for other than 1st spot in break)
'    '  4              +
'    '  5              |
'    '  6              40       Comment
'
'    '  1              |
'    '  2              8         Start Time HR:MN:SC Military time
'    '  3              |         blank or + (for other than 1st spot in break)
'    '  4              +
'    '  5              |
'    '  6              4       cart (without media code)
'    '  7              ||
'    '  Following 2 fields may need to be removed.  Client needs to let us know if it interferes with the equipment
'    '  8              20      Advertiser name
'    '  9              10      Internal Spot Code
'
'    '9-11-06 WireReady or 2-18-11 AudioVault Air
'    'Record Format    Max Length    Description
'    '1                  8           Scheduled Spot Time HH:MM:SS  military tie
'    '9                  20          Cart # (w/material Code) or ISCI
'    '29                 46          Advt name/prod name (up to 30 char for advt name)
'    '75                 20          Creative Title
'    '95                 4           Spot Length MMSS
'    '99                 10          Spot Internal code
'    '109                1           C/R
'    '110                1           L/F
'    '
' 'Record format for Rivendell (fixed length, blank after each field)
' '   Field #       Max Length    Description
' '    1-8              8          Event Time HH:MM:SS   24hr time
' '    9                1          Blank
' '    10-10            1          Single alpha-numeric chanracter, ignored by Rivendell
' '    11-24            14         cart#, numeric, 6 digits max, right padded with spaces
' '    25               1          Blank
' '    26-59            34         Cart Title, right padded with spaces
' '    60               1          Blank
' '    61-68            8          Spot Length HH:MM:SS
' '    69               1          Blank
' '    70-101           32         Generic element ISCI Code, right padded with spaces
' '    102              1          Blank
' '    103-134          32         Event GUID, Spot internal code
' '    135              1          c/r
' '    136              1          linefeed
'
'
'    '11-17-10 Audio Vault RPS.  There are 2 record types:  Comments and spots
'    ' Field #     Max length   Description
'    ' Comment type record
'    '  1              |
'    '  2              10        Start Time HR:MN:SCAM/PM (06:12:30PM)
'    '  3              |
'    '  4              40 variable length- Comment (new record created for each string separated by ;
'
'    'spot
'    '  1              |
'    '  2              10         Start Time HR:MN:SCAM/PM (06:12:30PM)
'    '  3              |
'    '  4              +
'    '  5              |
'    '  6              4       cart (without media code)
'    '  7              ||
'    '  8              20      Advertiser name
'    '  9              10      Internal Spot Code
'    '3/3/12: Test if Media code set to suppress spots.  This flag is obtained from gObtainCopy call
'
'    '6-21-12 Jelli Comma delimited
'    '   Field Name       Format          Max Length
'    '   ISCI Code        AN              20
'    '   Start Date/Time  YYYY-mm-dd hh:mm:ss    military 19
'    '   Length           nnn             3
'    '   Creative Title   AN              30
'    '   Veh STation Code AN              5
'    '   contract         N               9
'    '   Advt Name        AN              30
'    '   Agency Name      AN              40
'    '   Internal Code    nnn             10
'
'    '10/16/12 ENCO-ESPN Comma delimited
'    '   Field Name       Format          Max Length
'    '   Date             m/dd/yy         8
'    '   Time             hh:mm:ss        8 (military)
'    '   Length           nnn             3
'    '   Event Name       AN              20
'    '   ISCI Code        AN              20
'    '   Advt/Prod        AN              35
'    '   ProgCodeHnBnPn   AN
'    '   Internal Code    nnn             10
'
''Record format for Zetta(Fixed length):  All unused fields with blanks
''this is a subset of RCS-5
'    'Column  Length  Field
'    '  1        1    Commercial Indicator always C
'    '  2        7    Start Time HRMN:SC  Military Hours
'    '  9       15    Cart Number (3 char material code + 4 char cart #)
'    ' 24       24    Creative Title
'    ' 48        3    Unused (Priority number)
'    ' 51        4    Length  MNSC (60=> 0100; 90=> 0130; 30=>0030)
'    ' 55        4    Unused (Commercial Type)
'    ' 59        10    CSI Internal SDF code
'    ' 69        2    Unused (Product Code)
'    ' 71        8    Unused (Ordered Time)
'    ' 79       24    Advertiser Name
'    '103       24    Agency Name (blank if direct)
'    '127       24    Unused (Comment)
'    '151        8    Unused (Start Date mm/dd/yy)
'    '159        5    Unused (Start time hh:mm)
'    '164        8    Unused (Stop Date)
'    '172        5    Unused (Stop Time)
'    '177        8    Unused (Kill Date)
'    '185       15    Product Name
'    '200       15    Unused (Product 2)
'    '215        1    Live Y/N
'    '216        1    Unused (Is External? Y or N) , Always N
'    '217        2    Unused (CBSI Stopset number)
'

End Function

Private Function mGetProgCode(slProgCodeID As String) As String
    'GSF is read in cmcExport_Click
    mGetProgCode = slProgCodeID
    If UCase(slProgCodeID) = "EVENT" Then
        mGetProgCode = Trim$(tmGsf.sXDSProgCodeID)
    End If
    Exit Function
End Function
Private Function mGetMergeProgCode(ilVefCode As Integer, llDate As Long, llTime As Long, slProgCodeID As String, ilParentVefCode As Integer) As String
    '9/9/13: Handle Merge
    Dim ilLoop As Integer
    
    mGetMergeProgCode = slProgCodeID
    ilParentVefCode = ilVefCode
    If UCase(mGetMergeProgCode) = "MERGE" Then
        For ilLoop = LBound(tmProgTimeRange) To UBound(tmProgTimeRange) - 1 Step 1
            If tmProgTimeRange(ilLoop).iVefCode <> ilVefCode Then
                If tmProgTimeRange(ilLoop).lDate = llDate Then
                    If (llTime >= tmProgTimeRange(ilLoop).lStartTime) And (llTime < tmProgTimeRange(ilLoop).lEndTime) Then
                        ilParentVefCode = tmProgTimeRange(ilLoop).iVefCode
                        mGetMergeProgCode = Trim(tmProgTimeRange(ilLoop).sProgID)
                        Exit Function
                    End If
                End If
            End If
        Next ilLoop
    End If
    Exit Function
End Function
'
'
'               Format Copy string into record array
'               Some formats need material code, some do not
'
'           <input> ilAutomation Type : from tgSpf.sAutoType
'               ilCopy - true if copy exists
'               slCopyStr - Cart # (with or without material code, or ISCI)
'               llCopyMissSdfCode - spot code from missing copy
'               llSpotCode - current spot processing
'               slMsg - text for error msg & audit trail
'           <output>
'               Return - copy String
'
'    '1-17-02 special code implemented for Prophet (Sirius) to append a hard-coded "cut-# " to their copy
'   11-18-02 For MAI & Prophet- append an -000 if material code F, all others append -001
'   6-25-05 implemention iMediaTouch (ilAutomationType = 8)
'   10-12-06 Audio Vault has an additional option to include/excl the media code
Private Function mFormatCopy(ilAutomationType As Integer, ilCopy As Integer, slCopyStr As String, llCopyMissingSdfCode As Long, llSpotCode As Long, slMsg As String, ilMediaCodeLen As Integer) As String
'
Dim ilCopyLen As Integer
Dim slStr As String
Dim ilCopyInx As Integer
    ilCopyInx = 1
    '1-6-12 Wide Orbit
    If ilAutomationType = 5 Or ilAutomationType = AUTOTYPE_SCOTT Or ilAutomationType = AUTOTYPE_SCOTT_V5 Or ilAutomationType = 8 Or ilAutomationType = 12 Or ilAutomationType = 13 Or ilAutomationType = 15 Or ilAutomationType = AUTOTYPE_WIDEORBIT Then           'RCS-4, rcs-5 or Scott ,iMediaTouch,rivendell, audio vault air  can only have 4 char copy
    'If ilAutomationType = 5 Or ilAutomationType = 3 Or ilAutomationType = 8 Or ilAutomationType = 12 Or ilAutomationType = 13 Then            'RCS or Scott or iMediaTouch   can only have 4 char copy
        ilCopyLen = 4
        If ilAutomationType = 12 Then
            ilCopyLen = 15
        End If
        If ilAutomationType = 13 Then
            ilCopyLen = 14
        End If
        If ilAutomationType = 15 Then           'audiovault air
            ilCopyLen = 20
        End If
        
        If tgSpf.sUseCartNo <> "N" Then
            'see how many char the material code is
            ilCopyInx = ilMediaCodeLen + 1          '2-27-12
            'ilCopyInx = 2       'ignore the material code
        Else
            ilCopyInx = 1       'using agency tape #s
        End If
    ElseIf ilAutomationType = 7 Then            'Phophet
        ilCopyLen = 5
        'ilCopyInx = 2   'Ignore Material code  2-27-12 media code may be more than 1 char,
        ilCopyInx = ilMediaCodeLen + 1      '2-27-12
    ElseIf ilAutomationType = 9 Then            'audio vault

        ilCopyLen = 5
        If tgSpf.sUseCartNo <> "N" Then
            If imUsingMediaCodeForAV Then
                ilCopyInx = 1       'use the material code
                ilCopyLen = 10      'use max 5 char mater code & 5 char inv #
            Else

                'ilCopyInx = 2        'ignore material code definitions
                ilCopyInx = ilMediaCodeLen + 1      'material code may be more than 1 char , ignore all of them
            End If
        Else
            ilCopyInx = 1       'using agency tape #s
        End If
    '11-17-10 audio vault RPS added
    ElseIf ilAutomationType = 14 Then       'audiovault RPS (ignores material code)
        'ilCopyInx = 2           'ignore material code
        ilCopyInx = ilMediaCodeLen + 1          '2-27-12
        ilCopyLen = 4           'length of copy cart
    ElseIf ilAutomationType = AUTOTYPE_ZETTA Then
         ilCopyLen = 15
         ilCopyInx = 1
    Else
        ilCopyLen = 20
    End If
    '1-17-02 special code implemented for Prophet (Sirius) to append a hard-coded "cut-# " to their copy
    If ilAutomationType = 2 Then            'Prophet, but for Sirius client only
        slStr = Mid(slCopyStr, 1, 1)        '11-18-02 get the media code and see if its an "F"
        If Trim$(slStr) = "F" Then
            slCopyStr = Trim$(slCopyStr) & "-000"
        Else
            slCopyStr = Trim$(slCopyStr) & "-001"
        End If
    End If

    slStr = Mid(slCopyStr, ilCopyInx, ilCopyLen)

    If Not ilCopy Then
        imCopyMissing = True
        If llCopyMissingSdfCode <> tmSdf.lCode Then
            'Print #hmMsg, slMsg
            gAutomationAlertAndLogHandler slMsg
            lbcMsg.AddItem slMsg
            llCopyMissingSdfCode = tmSdf.lCode
        End If
        slStr = ""
    End If
    Do While Len(slStr) < ilCopyLen
        slStr = slStr & " "
    Loop
    mFormatCopy = slStr
End Function
'
'
'           Format SDF spot length as 4 characters seconds or
'           MM:SS  5-4-01
'       <input> ilSpotLen -  spot length in seconds
'       <output> slLenInSec - left filled with 0, always in total seconds
'                slLenInMinSec - MMSS (MM & SS left filled with 0).  60" shows as 0100
'       7-12-01 length in min & sec did not show the minutes properly
'
Private Sub mFormatSpotLen(ilSpotLen As Integer, slLenInSec As String, slLenInMinSec As String)
Dim llMin As Long
Dim llSec As Long
Dim slTemp As String
    'show in MMSS, left fill with zeros
    If ilSpotLen < 60 Then
        slLenInMinSec = Trim$(str$(ilSpotLen))
        Do While Len(slLenInMinSec) < 4
            slLenInMinSec = "0" & slLenInMinSec
        Loop
    Else
        llMin = ilSpotLen
        llSec = llMin Mod 60
        llMin = llMin \ 60
        slLenInMinSec = Trim$(str$(llMin))
        Do While Len(slLenInMinSec) < 2
            slLenInMinSec = "0" & slLenInMinSec
        Loop
        'slLenInMinSec = Trim$(Str$(llSec))
        slTemp = Trim$(str$(llSec))
        Do While Len(slTemp) < 2
            'slLenInMinSec = "0" & slLenInMinSec
            slTemp = "0" & slTemp
        Loop
        slLenInMinSec = slLenInMinSec & Trim$(slTemp)
    End If
        'Spot Length , all seconds
        slLenInSec = Trim$(str$(ilSpotLen))
        Do While Len(slLenInSec) < 4
            slLenInSec = "0" & slLenInSec
        Loop

End Sub '
'
'           Format SDF spot length as 4 characters seconds or
'           MM:SS
'           If 60", make it 60s vs 1m
'       <input> ilSpotLen -  spot length in seconds
'       <output> slLenInSec - left filled with 0, always in total seconds
'                slLenInMinSec - MMSS (MM & SS left filled with 0).  60" shows as 0100
'
Private Sub mFormatSpotLenForMin(ilSpotLen As Integer, slLenInSec As String, slLenInMinSec As String)
Dim llMin As Long
Dim llSec As Long
Dim slTemp As String
    'show in MMSS, left fill with zeros
    If ilSpotLen <= 60 Then
        slLenInMinSec = Trim$(str$(ilSpotLen))
        Do While Len(slLenInMinSec) < 4
            slLenInMinSec = "0" & slLenInMinSec
        Loop
    Else
        llMin = ilSpotLen
        llSec = llMin Mod 60
        llMin = llMin \ 60
        slLenInMinSec = Trim$(str$(llMin))
        Do While Len(slLenInMinSec) < 2
            slLenInMinSec = "0" & slLenInMinSec
        Loop
        'slLenInMinSec = Trim$(Str$(llSec))
        slTemp = Trim$(str$(llSec))
        Do While Len(slTemp) < 2
            'slLenInMinSec = "0" & slLenInMinSec
            slTemp = "0" & slTemp
        Loop
        slLenInMinSec = slLenInMinSec & Trim$(slTemp)
    End If
        'Spot Length , all seconds
        slLenInSec = Trim$(str$(ilSpotLen))
        Do While Len(slLenInSec) < 4
            slLenInSec = "0" & slLenInSec
        Loop

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilAutomationType As Integer
    Dim ilHowMany As Integer
    Dim ilStartPos As Integer
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imAllClicked = False
    imSetAll = True
    imAllClickedItems = False
    imSetAllItems = True
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    'open ARF table for the location to store export file
    hmArf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmArf, "", sgDBPath & "Arf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:ARF)", ExptGen
    On Error GoTo 0
    imArfRecLen = Len(tmArf)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:ADF)", ExptGen
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:MNF)", ExptGen
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:CPF)", ExptGen
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:ANF)", ExptGen
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:CHF)", ExptGen
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:CLF)", ExptGen
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:VSF)", ExptGen
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:MCF)", ExptGen
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:CIF)", ExptGen
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:TZF)", ExptGen
    On Error GoTo 0
    imTzfRecLen = Len(tmTzf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:VSF)", ExptGen
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVpf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:VPF)", ExptGen
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:VLF)", ExptGen
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmDlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:DLF)", ExptGen
    On Error GoTo 0
    imDlfRecLen = Len(tmDlf)
    hmRsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:RSF)", ExptGen
    On Error GoTo 0
    imRsfRecLen = Len(tmRsf)
    hmAxf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAxf, "", sgDBPath & "Axf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:AXF)", ExptGen
    On Error GoTo 0
    imAxfRecLen = Len(tmAxf)
    hmCvf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:CVF)", ExptGen
    On Error GoTo 0
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:CRF)", ExptGen
    On Error GoTo 0
    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:SDF)", ExptGen
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpenSSF)", ExptGen
    On Error GoTo 0
    hmCTSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCTSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:SSF)", ExptGen
    On Error GoTo 0
    hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:GSF)", ExptGen
    On Error GoTo 0
    imGsfRecLen = Len(tmGsf)
         
    hmEtf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmEtf, "", sgDBPath & "Etf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:ETF)", ExptGen
    On Error GoTo 0
    imEtfRecLen = Len(tmEtf)
    
    hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:CEF)", ExptGen
    On Error GoTo 0
    imCefRecLen = Len(tmCef)
    
    '4-26-11 Lcf table to set "Day is Not Complete flag"
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:LCF)", ExptGen
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)
    
    hmLvf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:LVF)", ExptGen
    On Error GoTo 0
    
    If (Asc(tgSpf.sAutoType3) And JELLI) = JELLI Then       '6-21-12
        hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen:AGF)", ExptGen
        On Error GoTo 0
        imAgfRecLen = Len(tmAgf)
        gObtainAgency
    End If


    ilRet = gVffRead()
    
    'Populate arrays to determine if records exist
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        'mTerminate
        Exit Sub
    End If
    
    
    '4/26/11: Add test of avail attribute
    ilRet = gObtainAvail()
    
    smTeamTag = ""
    ilRet = gObtainMnfForType("Z", smTeamTag, tmTeam())
    
    'Select Satellite Music; SMN Mix; Tom Joyner
    'For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '    slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    ilVefCode = Val(slCode)
    '    For ilTest = 0 To UBound(tgVpf) Step 1
    '        If ilVefCode = tgVpf(ilTest).iVefKCode Then
    '            If tgVpf(ilTest).sExpHiPhoenix = "Y" Then
    '                lbcVehicle.Selected(ilLoop) = True
    '            End If
    '            Exit For
    '        End If
    '    Next ilTest
    'Next ilLoop
    For ilLoop = LBound(imEvtType) To UBound(imEvtType) Step 1
        imEvtType(ilLoop) = True
    Next ilLoop
    imEvtType(0) = False 'Don't include library names
    'plcGauge.Move ExptGen.Width / 2 - plcGauge.Width / 2
    'cmcFileConv.Move ExptGen.Width / 2 - cmcFileConv.Width / 2
    'cmcCancel.Move ExptGen.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    'cmcReport.Move ExptGen.Width / 2 - cmcReport.Width / 2 + cmcReport.Width
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    '5-21-01 changed from 1 byte string to binary value
    'ilAutomationType = Asc(tgSpf.sAutoType2) * 256
    ilAutomationType = (Asc(tgSpf.sAutoType2) And &H7F) * 256
    ilAutomationType = ilAutomationType + Asc(tgSpf.sAutoType)
    ilHowMany = 0
    ilStartPos = 780

    'ilAutomationType     '&H1 Dalet (was 1), &H2 Prophet NexGen (was 2), &H4 Scott (was 3), &H8 Drake (was 4), &H10 RCS (was 5), Prophet Wizard &H20 (6)
    If (ilAutomationType And DALET) = DALET Then
        smScreenCaption = "Export Dalet"
        plcScreen_Paint
        rbcAutoType(0).Left = ilStartPos
        rbcAutoType(0).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(0).Width + 60
        ilHowMany = ilHowMany + 1
        rbcAutoType(0).Value = True
    End If
    If (ilAutomationType And DRAKE) = DRAKE Then
        smScreenCaption = "Export Drake"
        plcScreen_Paint
        rbcAutoType(1).Left = ilStartPos
        rbcAutoType(1).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(1).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(1).Value = True
        End If
    End If
    If (ilAutomationType And IMEDIATOUCH) = IMEDIATOUCH Then          '6-25-05
        smScreenCaption = "Export iMediaTouch"
        plcScreen_Paint
        rbcAutoType(7).Left = ilStartPos
        rbcAutoType(7).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(7).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(7).Value = True
        End If
    End If
    If (ilAutomationType And PROPHETNEXGEN) = PROPHETNEXGEN Then
        smScreenCaption = "Export Prophet NexGen"
        plcScreen_Paint
        rbcAutoType(2).Left = ilStartPos
        rbcAutoType(2).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(2).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(2).Value = True
        End If
       End If
    If (ilAutomationType And PROPHETWIZARD) = PROPHETWIZARD Then
        smScreenCaption = "Export Prophet Wizard"
        plcScreen_Paint
        rbcAutoType(3).Left = ilStartPos
        rbcAutoType(3).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(3).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(3).Value = True
        End If
    End If
    If (ilAutomationType And RCS4DIGITCART) = RCS4DIGITCART Then
        smScreenCaption = "Export RCS 4 Digit Cart #"
        plcScreen_Paint
        rbcAutoType(4).Left = ilStartPos
        rbcAutoType(4).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(4).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(4).Value = True
        End If
    End If
    If (Asc(tgSpf.sAutoType2) And RCS5DIGITCART) = RCS5DIGITCART Then
        smScreenCaption = "Export RCS 5 Digit Cart #"
        plcScreen_Paint
        rbcAutoType(11).Left = ilStartPos
        rbcAutoType(11).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(11).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(11).Value = True
        End If
    End If
    If (ilAutomationType And SCOTT) = SCOTT Then          '2-5-03
        smScreenCaption = "Export Scott"
        plcScreen_Paint
        rbcAutoType(5).Left = ilStartPos
        rbcAutoType(5).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(5).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(5).Value = True
        End If
    End If
    If (ilAutomationType And PROPHETMEDIASTAR) = PROPHETMEDIASTAR Then          '9-24-03
        smScreenCaption = "Export Prophet MediaStar"
        plcScreen_Paint
        rbcAutoType(6).Left = ilStartPos
        rbcAutoType(6).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(6).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(6).Value = True
        End If
    End If
    If (ilAutomationType And (256 * AUDIOVAULT)) = (256 * AUDIOVAULT) Then        '8-10-05
        smScreenCaption = "Export Audio Vault Sat"
        plcScreen_Paint
        rbcAutoType(8).Left = ilStartPos
        rbcAutoType(8).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(8).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(8).Value = True
        End If
    End If
    '9-11-06 Wire Ready
    If (ilAutomationType And (256 * WIREREADY)) = (256 * WIREREADY) Then        '8-10-05
        smScreenCaption = "Export WireReady"
        plcScreen_Paint
        rbcAutoType(9).Left = ilStartPos
        rbcAutoType(9).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(9).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(9).Value = True
        End If
    End If

    '8-21-08
    If (ilAutomationType And (256 * SIMIAN)) = (256 * SIMIAN) Then        '8-21-08
        smScreenCaption = "Export Simian"
        plcScreen_Paint
        rbcAutoType(10).Left = ilStartPos
        rbcAutoType(10).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(10).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(10).Value = True
        End If
    End If
    If (Asc(tgSpf.sUsingFeatures8) And RIVENDELLEXPORT) = RIVENDELLEXPORT Then
        smScreenCaption = "Export Rivendell"
        plcScreen_Paint
        rbcAutoType(12).Left = ilStartPos
        rbcAutoType(12).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(12).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(12).Value = True
        End If
    End If

    
    '11-15-10 AudioVault RPS
    If (Asc(tgSpf.sAutoType2) And AUDIOVAULTRPS) = AUDIOVAULTRPS Then
        smScreenCaption = "Export Audio Vault RPS"
        plcScreen_Paint
        rbcAutoType(13).Left = ilStartPos
        'rbcAutoType(13).Width = 1800
        rbcAutoType(13).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(13).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(13).Value = True
        End If
    End If
    
    '2-18-11 AudioVaultAir
    If (Asc(tgSpf.sAutoType3) And AUDIOVAULTAIR) = AUDIOVAULTAIR Then
        smScreenCaption = "Export Audio Vault Air"
        plcScreen_Paint
        rbcAutoType(14).Left = ilStartPos
        'rbcAutoType(14).Width = 1800
        rbcAutoType(14).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(14).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(14).Value = True
        End If
    End If
    
    '1-6-12 Wide Orbit
    If (Asc(tgSpf.sAutoType3) And WIDEORBIT) = WIDEORBIT Then
        smScreenCaption = "Export Wide Orbit"
        plcScreen_Paint
        rbcAutoType(15).Left = ilStartPos
        'rbcAutoType(15).Width = 1800
        rbcAutoType(15).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(14).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(15).Value = True
        End If
    End If
    
    '6-21-12 Jelli
    If (Asc(tgSpf.sAutoType3) And JELLI) = JELLI Then
        smScreenCaption = "Export Jelli"
        plcScreen_Paint
        rbcAutoType(16).Left = ilStartPos
        'rbcAutoType(16).Width = 1800
        rbcAutoType(16).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(16).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(16).Value = True
        End If
    End If
    
    '10/16/12 Enco-ESPN
    If (Asc(tgSpf.sAutoType3) And ENCOESPN) = ENCOESPN Then
        'smScreenCaption = "Export Enco-ESPN"
        smScreenCaption = "Export Linkup"   'TTP 10876 JJB
        plcScreen_Paint
        rbcAutoType(17).Left = ilStartPos
        'rbcAutoType(17).Width = 1800
        rbcAutoType(17).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(17).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(17).Value = True
        End If
    End If
    
    '8-16-13 Scott V5
    If (Asc(tgSpf.sAutoType3) And SCOTT_V5) = SCOTT_V5 Then
        smScreenCaption = "Export Scott_V5"
        plcScreen_Paint
        rbcAutoType(18).Left = ilStartPos
        'rbcAutoType(17).Width = 1800
        rbcAutoType(18).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(18).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(18).Value = True
        End If
    End If

    '1-5-16
    If (Asc(tgSpf.sAutoType3) And ZETTA) = ZETTA Then
        smScreenCaption = "Export Zetta"
        plcScreen_Paint
        rbcAutoType(19).Left = ilStartPos
        'rbcAutoType(17).Width = 1800
        rbcAutoType(19).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(19).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(19).Value = True
        End If
    End If
    '5-11-18 station playlist
    If (Asc(tgSpf.sAutoType3) And STATIONPLAYLIST) = STATIONPLAYLIST Then
        smScreenCaption = "Export Station PlayList"
        plcScreen_Paint
        rbcAutoType(20).Left = ilStartPos
        'rbcAutoType(17).Width = 1800
        rbcAutoType(20).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(20).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(20).Value = True
        End If
    End If
    
    '5-1-20 RadidoMan
    If (Asc(tgSpf.sAutoType3) And RADIOMAN) = RADIOMAN Then        '8-10-05
        smScreenCaption = "Export RadioMan"
        plcScreen_Paint
        rbcAutoType(21).Left = ilStartPos
        rbcAutoType(21).Top = 90
        ilStartPos = ilStartPos + rbcAutoType(9).Width + 60
        ilHowMany = ilHowMany + 1
        If ilHowMany = 1 Then
            rbcAutoType(21).Value = True
        End If
    End If


    If ilHowMany = 0 Then           'At least one export type indicated?
        ''MsgBox "Invalid or no Automation type defined" & Trim$(tgSpf.sAutoType), vbOkOnly + vbCritical + vbApplicationModal, "mInit"
        gAutomationAlertAndLogHandler "Invalid or no Automation type defined" & Trim$(tgSpf.sAutoType), vbOkOnly + vbCritical + vbApplicationModal, "mInit"
        imTerminate = True
        Exit Sub
    End If
    If ilHowMany > 1 Then
'        plcScreen.Visible = False           'turn off the single automation caption
        smScreenCaption = "Export"
        'loop to create which one to show
        'ilAutomationType     '&H1 Dalet (was 1), &H2 Prophet NexGen (was 2), &H4 Scott (was 3), &H8 Drake (was 4), &H10 RCS (was 5), Prophet Wizard &H20 (6)
        If (ilAutomationType And DALET) = DALET Then        'dalet
            rbcAutoType(0).Visible = True
        End If
        If (ilAutomationType And DRAKE) = DRAKE Then        'drake
            rbcAutoType(1).Visible = True
        End If
        If (ilAutomationType And PROPHETNEXGEN) = PROPHETNEXGEN Then        'prophet nexgen
            rbcAutoType(2).Visible = True
        End If
        If (ilAutomationType And PROPHETWIZARD) = PROPHETWIZARD Then        'prophet wizard
            rbcAutoType(3).Visible = True
        End If
        If (ilAutomationType And RCS4DIGITCART) = RCS4DIGITCART Then      'rcs
            rbcAutoType(4).Visible = True
        End If
        If (ilAutomationType And SCOTT) = SCOTT Then      'scott
            rbcAutoType(5).Visible = True
        End If
        If (ilAutomationType And PROPHETMEDIASTAR) = PROPHETMEDIASTAR Then      'prophet mediastar
            rbcAutoType(6).Visible = True
        End If
        '6-25-05 iMediaTouch
        If (ilAutomationType And IMEDIATOUCH) = IMEDIATOUCH Then
            rbcAutoType(7).Visible = True
        End If
         '8-10-05 Audio Vault Sat
        If (ilAutomationType And (256 * AUDIOVAULT)) = (256 * AUDIOVAULT) Then
            rbcAutoType(8).Visible = True
        End If

        '9-11-06
        If (ilAutomationType And (256 * WIREREADY)) = (256 * WIREREADY) Then
            rbcAutoType(9).Visible = True
        End If

        '8-21-08
        If (ilAutomationType And (256 * SIMIAN)) = (256 * SIMIAN) Then
            rbcAutoType(10).Visible = True
        End If
        If (Asc(tgSpf.sAutoType2) And RCS5DIGITCART) = RCS5DIGITCART Then
            rbcAutoType(11).Visible = True
        End If
        If (Asc(tgSpf.sUsingFeatures8) And RIVENDELLEXPORT) = RIVENDELLEXPORT Then
            rbcAutoType(12).Visible = True
        End If
        
        If (Asc(tgSpf.sAutoType2) And AUDIOVAULTRPS) = AUDIOVAULTRPS Then       '11-15-10
            'rbcAutoType(13).Width = 1800
            rbcAutoType(13).Visible = True
        End If

        If (Asc(tgSpf.sAutoType3) And AUDIOVAULTAIR) = AUDIOVAULTAIR Then       '2-18-11
            'rbcAutoType(14).Width = 1800
            rbcAutoType(14).Visible = True
        End If
        '1-6-12 Wide Orbit
        If (Asc(tgSpf.sAutoType3) And WIDEORBIT) = WIDEORBIT Then       '1-6-12
            'rbcAutoType(15).Width = 1800
            rbcAutoType(15).Visible = True
        End If
        
        '6-21-12 Jelli
        If (Asc(tgSpf.sAutoType3) And JELLI) = JELLI Then       '6-21-21
            'rbcAutoType(16).Width = 1800
            rbcAutoType(16).Visible = True
        End If
 
         '10/16/12 Enco-ESPN
        If (Asc(tgSpf.sAutoType3) And ENCOESPN) = ENCOESPN Then
            'rbcAutoType(17).Width = 1800
            rbcAutoType(17).Visible = True
        End If
        
        '8-16-13 Scott_v5
        If (Asc(tgSpf.sAutoType3) And SCOTT_V5) = SCOTT_V5 Then
            'rbcAutoType(17).Width = 1800
            rbcAutoType(18).Visible = True
        End If
        
        '1-5-16 Zetta
        If (Asc(tgSpf.sAutoType3) And ZETTA) = ZETTA Then
            rbcAutoType(19).Visible = True
        End If
        
        '5-11-18 Station Playlist
        If (Asc(tgSpf.sAutoType3) And STATIONPLAYLIST) = STATIONPLAYLIST Then
            rbcAutoType(20).Visible = True
        End If
        
        '5-1-20
        If (Asc(tgSpf.sAutoType3) And RADIOMAN) = RADIOMAN Then
            rbcAutoType(21).Visible = True
        End If

'        plcAutoType.Left = 860
'        plcAutoType.Top = 60
'        plcAutoType.Visible = True
    End If


    gPopVehicleGroups ExptGen!cbcSet1, tgVehicleSets1(), True

    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear

    plcCalendar.Move 870, 600
    plcCalendarTo.Move 2895, 600
    pbcCalendar_Paint   'mBoxCalDate called within paint
    pbcCalendarTo_Paint
    lacDate.Visible = False
    lacDateTo.Visible = False
    gCenterStdAlone ExptGen
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
    Exit Sub
mInitErr:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.Height
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExptGen.Txt"
    slToFile = sgDBPath & "Messages\" & "ExptGen.Txt"
    sgMessageFile = slToFile
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, ""
    'Print #hmMsg, "** Export Dalet Systems: " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & " **"
    'Print #hmMsg, "** " & smScreenCaption & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    gAutomationAlertAndLogHandler "** " & smScreenCaption & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function

Private Sub mSetCommands()
Dim ilEnabled As Integer
Dim ilLoop As Integer
    ilEnabled = False
    'at least one vehicle must be selected
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            ilEnabled = True
            Exit For
        End If
    Next ilLoop

    'Automation type must be selected (either using a single one which is retrieved from the Site Pref,
    'or if multiple ones must be selected on front screen
    If ilEnabled Then
        ilEnabled = False
        If rbcAutoType(0).Value Or rbcAutoType(1).Value Or rbcAutoType(2).Value Or rbcAutoType(3).Value Or rbcAutoType(4).Value Or rbcAutoType(5).Value Or rbcAutoType(6).Value Or rbcAutoType(7).Value Or rbcAutoType(8).Value Or rbcAutoType(9).Value Or rbcAutoType(10).Value Or rbcAutoType(11).Value Or rbcAutoType(12).Value Or rbcAutoType(13).Value Or rbcAutoType(14).Value Or rbcAutoType(15).Value Or rbcAutoType(16).Value Or rbcAutoType(17).Value Or rbcAutoType(18).Value Or rbcAutoType(19).Value Or rbcAutoType(20).Value Or rbcAutoType(21).Value Then
            ilEnabled = True
        End If
        If ilEnabled Then
            ilEnabled = False
            If edcStartDate <> "" Then
                ilEnabled = True
            End If
        End If
    End If

    If ilEnabled Then
        If (imSetIndex > 0 And lbcGroupItems.SelCount <= 0) Then
            ilEnabled = False
        End If
    End If
    cmcExport.Enabled = ilEnabled
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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


    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExptGen
    igManUnload = NO
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
    Dim ilLoop As Integer
    Dim ilVff As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    
    'ilRet = gPopUserVehicleBox(ExptGen, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHBYPASSWEGENER_OLA + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    If rbcAutoType(17).Value Then   'Enco-ESPN
        ilRet = gPopUserVehicleBox(ExptGen, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHLOG + VEHLOGVEHICLE + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    Else
        ilRet = gPopUserVehicleBox(ExptGen, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHSPORTMINUELIVE + VEHLOG + VEHLOGVEHICLE + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    End If
    
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExptGen
        On Error GoTo 0
    End If
    
    If rbcAutoType(17).Value Then   'Enco-ESPN
        For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
            slNameCode = tgUserVehicle(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                If ilVefCode = tgVff(ilVff).iVefCode Then
                    If tgVff(ilVff).sExportEncoESPN = "Y" Then
                        lbcVehicle.Selected(ilLoop) = True
                    End If
                    Exit For
                End If
            Next ilVff
        Next ilLoop
    End If
    
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub lbcVehicle_Scroll()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
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
                edcStartDate.Text = Format$(llDate, "m/d/yy")
                edcStartDate.SelStart = 0
                edcStartDate.SelLength = Len(edcStartDate.Text)
                imBypassFocus = True
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcStartDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate edcStartDate, lacDate
End Sub

Private Sub pbcCalendarTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
                edcEndDate.Text = Format$(llDate, "m/d/yy")
                edcEndDate.SelStart = 0
                edcEndDate.SelLength = Len(edcEndDate.Text)
                imBypassFocus = True
                edcEndDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcEndDate.SetFocus
End Sub
Private Sub pbcCalendarTo_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalNameTo.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendarTo, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate edcEndDate, lacDateTo
End Sub


Private Sub plcScreen_GotFocus()
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub rbcAutoType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcAutoType(Index).Value
    'If Value = vbTrue Then
    If Value = True Then
        mVehPop
        mHighlightVehicles Index
    End If
    'End of coded added
    mSetCommands
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub

Private Sub rbcAutoType_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub rbcSplit_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    plcCalendarTo.Visible = False
End Sub

Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub

Private Function mBBSpots(ilVefCode As Integer) As Integer
    'Dim ilLoop As Integer
    Dim ilRet As Integer
    'Dim slNameCode As String
    'Dim slCode As String
    'Dim ilVefCode As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slStr As String

    mBBSpots = True
    If tgSpf.sUsingBBs <> "Y" Then
        Exit Function
    End If
    slStr = edcStartDate.Text
    llStartDate = gDateValue(slStr)
    slStr = edcEndDate.Text
    llEndDate = gDateValue(slStr)
    'Determine vehicles to create Billboard spots
    'For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '    If lbcVehicle.Selected(ilLoop) Then
    '        slNameCode = tgUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
    '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '        ilVefCode = Val(slCode)
            ilRet = gMakeBBAndAssignCopy(hmSdf, hmVlf, ilVefCode, llStartDate, llEndDate)
            If Not ilRet Then
                mBBSpots = False
            End If
    '    End If
    'Next ilLoop
End Function



Private Function mSetLogDate() As Integer
    Dim ilVpf As Integer
    Dim ilVef As Integer
    Dim slEndDate As String
    Dim slStartDate As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilRet As Integer
    Dim slDate As String

    ReDim ilVefCode(0 To 0) As Integer
    mSetLogDate = True
    slStartDate = edcStartDate.Text
    slEndDate = edcEndDate.Text
    ilVpf = gBinarySearchVpf(tmVef.iCode)
    If ilVpf <> -1 Then
        'If (tgVpf(ilVpf).sGenLog = "N") Then
            If tmVef.sType = "A" Then
                gBuildLinkArray hmVlf, tmVef, slStartDate, ilVefCode()
                ilVefCode(UBound(ilVefCode)) = tmVef.iCode
                ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
            ElseIf tmVef.sType = "L" Then
                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '7/27/12: Include Sports within Log vehicles
                    'If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                    If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                        ilVefCode(UBound(ilVefCode)) = tgMVef(ilVef).iCode
                        ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                    End If
                Next ilVef
                ilVefCode(UBound(ilVefCode)) = tmVef.iCode
                ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
            Else
                ReDim ilVefCode(0 To 1) As Integer
                ilVefCode(0) = tmVef.iCode
            End If
            gGetSyncDateTime slSyncDate, slSyncTime
            For ilVef = 0 To UBound(ilVefCode) - 1 Step 1
                Do
                    tmVpfSrchKey.iVefKCode = ilVefCode(ilVef)
                    ilRet = btrGetEqual(hmVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slDate
                    If gDateValue(slEndDate) > gDateValue(slDate) Then
                        gPackDate slEndDate, tmVpf.iLLD(0), tmVpf.iLLD(1)
                        'gPackDate slSyncDate, tmVpf.iSyncDate(0), tmVpf.iSyncDate(1)
                        'gPackTime slSyncTime, tmVpf.iSyncTime(0), tmVpf.iSyncTime(1)
                    End If
                    gPackDate slEndDate, tmVpf.iLPD(0), tmVpf.iLPD(1)
                    ilRet = btrUpdate(hmVpf, tmVpf, imVpfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilVpf = gBinarySearchVpf(ilVefCode(ilVef))
                If ilVpf <> -1 Then
                    tgVpf(ilVpf) = tmVpf
                End If
            Next ilVef
            '11/26/17
            gFileChgdUpdate "vpf.btr", False
            
        'End If
    End If
End Function
'
'           mAirTimeCopy - assign copy (or reassign superceded copy) to air time spots
'                           Billboards done in separate subroutine
'           <input> slSDate - start date
'                   slEDate - end date
'                   slSTime = 12M
'                   slETime = 12M
'
'
Private Sub mAirTimeCopy(slSDate As String, slEDate As String, slSTime As String, slETime As String)
Dim ilRet As Integer
Dim ilLink As Integer
Dim ilVefIndex As Integer
Dim ilZoneExist As Integer
Dim ilVpfIndex As Integer
Dim ilZone As Integer
Dim llTZDate As Long
Dim slTZDate As String

'   6-2-05 Currently, only assign copy to conventional vehicles.  Airing vehicles do not
'   have spots to update copy pointers.  (may need code similar to blackout code using rsf)
'   6-27-05 open up copy for airing vehicles

    ilVpfIndex = gVpfFind(ExptGen, tmVef.iCode)
    If tgVpf(ilVpfIndex).sGenLog = "Y" Then     'only assign copy if not generating a log (to speed up)
        Exit Sub
    End If
    llTZDate = gDateValue(slSDate) + 1
    slTZDate = Format$(llTZDate, "m/d/yy")       'for timezone copy on airing vehicles

    If tmVef.sType = "A" Then           'airing vehicle, build array of selling if copy entered by selling

        ilZoneExist = False
        For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
            If Trim$(tgVpf(ilVpfIndex).sGZone(ilZone)) <> "" Then
                ilZoneExist = True
                Exit For
            End If
        Next ilZone
        gBuildLinkArray hmVlf, tmVef, slSDate, igSVefCode() 'build array of selling vehicles
        For ilLink = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
            ilVefIndex = igSVefCode(ilLink)
            ilRet = gAssignCopyToSpots(0, ilVefIndex, 1, slSDate, slEDate, slSTime, slETime)
            If ilZoneExist Then
                ilRet = gAssignCopyToSpots(0, ilVefIndex, 1, slTZDate, slTZDate, slSTime, slETime)
            End If
        Next ilLink

    Else
        ilRet = gAssignCopyToSpots(0, tmVef.iCode, 1, slSDate, slEDate, slSTime, slETime)
    End If

End Sub
'
'               mCreateJump - this is required for Audio Vault export only
'               This creates the END record
'           <input> ilAutomation Type - reqd for type 9 only (audio vault)
'                   llDate - date of export
'
Private Sub mCreateJump(ilAutomationType As Integer, llDate As Long)
Dim slDayOfWeek As String * 21
Dim ilDays As Integer
Dim slDayName As String
Dim ilTime(0 To 1) As Integer
Dim llTime As Long
Dim slRecord As String

    slDayOfWeek = "MONTUEWEDTHUFRISATSUN"

    If ilAutomationType = 9 Then           'must be audio vault Sat to create the Jump record
        ilDays = gWeekDayLong(llDate + 1)
        slDayName = Mid$(slDayOfWeek, (ilDays) * 3 + 1, 3)
        slRecord = "|23:59:50|@|" & Trim$(tmVef.sCodeStn) & Trim$(slDayName) & "||J"
        gPackTime "11:59:50PM", ilTime(0), ilTime(1)
        gUnpackTimeLong ilTime(0), ilTime(1), False, llTime
        mSaveImage slRecord, Format$(llDate, "m/d/yy"), llTime, "C", "N"
    End If
    Exit Sub
End Sub
'
'            mSaveImage - save image in array of records so
'           they can be sorted
'           <input>
Private Sub mSaveImage(slRecord As String, slAirDate As String, llSpotTime As Long, slSortType As String, slXMid As String)
Dim slSortDate As String
Dim slStr As String
Dim llXMidDate As Long
Dim slLLCIndex As String
Dim slRPSCommentSub As String

    
    If slXMid = "Y" Then
        llXMidDate = gDateValue(slAirDate)
        slSortDate = Format$(llXMidDate + 1, "m/d/yy")
    Else
        slSortDate = Trim$(str$(gDateValue(slAirDate)))
    End If
    
    slSortDate = Trim$(str$(gDateValue(slAirDate)))
    Do While Len(slSortDate) < 6
        slSortDate = "0" & slSortDate
    Loop
    slStr = Trim$(str$(llSpotTime))
    Do While Len(slStr) < 6
        slStr = "0" & slStr
    Loop
    tmExpRecImage(UBound(tmExpRecImage)).sKey = slSortDate & slStr & slSortType
    tmExpRecImage(UBound(tmExpRecImage)).sRecord = slRecord
    ReDim Preserve tmExpRecImage(0 To UBound(tmExpRecImage) + 1) As EXPRECIMAGE
    Exit Sub
End Sub
'
'               mCheckEmptyAvail - Audio Vault requires entry if the avail is empty
'               <input> ilAutomationType - only Audio Vault Sat (Type = 9)
'                       llDate  - date processing
'                       llTime - time of avail
'                       ilNoSpotsThisAvail - # spots from the avail
Public Sub mCheckEmptyAvail(ilAutomationType As Integer, llInDate As Long, llTime As Long, ilNoSpotsThisAvail As Integer)
Dim slRecord As String
Dim slTime As String
Dim llDate As Long

    llDate = llInDate
    If ilAutomationType = 9 Then
        If ilNoSpotsThisAvail = 0 Then
            slTime = gFormatSpotTime(llTime)
            slRecord = "|" & Trim$(slTime) & "||Empty||" & """" & "*EmptyAvail"
            If imGameVehicle Then
                If lmLastEvtTime = -1 Then
                    lmLastEvtTime = llTime
                Else
                    If (llTime < lmLastEvtTime) And (Not imXMidNight) Then
                        imXMidNight = True
                    End If
                End If
                If imXMidNight Then
                    llDate = llDate + 1
                End If
            End If
            mSaveImage slRecord, Format$(llDate, "m/d/yy"), llTime, "B", "N"
        End If
    End If
    Exit Sub
End Sub
'
'           mGetProphetNextGenFields - 5 extra fields are required for Prophet Nexxtgen
'           4-6-06
'           ilAutomation : 2 = Prophet nextgen
'           Advertiser code (5 char autocode)
'           Competitve code (5 char autocode)
'           ROS flag - "R" = overrride or DP greater/equal than 18 hours, else "N"
'           Priority # - 6 characters (1-4 = line pty, last 2 char = price pty (reversed # from
'           what is stored in spot record (SSF entry).  i.e. if stored pty is 15, the value is 0, if 14, value is 1, etc.)
'           stored price pty is subtracted from 15
Public Sub mGetProphetNextGenFields(ilAutomation As Integer, slAdvtCode As String, slCompCode As String, slPty As String, slFixed As String, slROS As String)
Dim ilFound As Integer
Dim llOvStartTime As Long
Dim llOvEndTime As Long
Dim slStr As String
Dim ilLoop2 As Integer
Dim llTotalSec As Long

            If ilAutomation = 2 Then
                '4-7-06 gather the additional fields required for Prophet nextgen
                'Advt code
                slAdvtCode = Trim$(str$(tmAdf.iCode))
                Do While Len(slAdvtCode) < 5
                    slAdvtCode = "0" & slAdvtCode
                Loop
                'competitive code
                slCompCode = Trim$(str$(tmChf.iMnfComp(0)))
                Do While Len(slCompCode) < 5
                    slCompCode = "0" & slCompCode
                Loop

                slPty = Trim$(str$(tmSpot.iRank And RANKMASK))
                Do While Len(slPty) < 4
                    slPty = "0" & slPty
                Loop

                slStr = Trim$(str$(15 - ((tmSpot.iRank And PRICELEVELMASK) / SHIFT11)))    'reverse the priority, ie 15 becomes 0, 14 becomes 1, etc.
                Do While Len(slStr) < 2
                    slStr = "0" & slStr
                Loop
                slPty = slPty & slStr 'Concatenate the line priority with the spot priority

                'solo avail or 1st position in break
                If (tmClf.iPosition = 1) And ((Asc(tmClf.sOV2DefinedBits) And LN1STPOSITION) = LN1STPOSITION) Or tmClf.sSoloAvail = "Y" Then
                    slFixed = "F"
                Else
                    slFixed = "M"
                End If

                If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                    gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                    gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                    llTotalSec = (llOvEndTime - llOvStartTime)
                Else
                    'no overrides, use the DP start/end times
                    ilFound = gBinarySearchRdf(tmClf.iRdfCode)
                    If ilFound <> -1 Then           'found the dp
                        llTotalSec = 0
                        'accum all the segments of the daypart
                        For ilLoop2 = LBound(tgMRdf(ilFound).iStartTime, 2) To UBound(tgMRdf(ilFound).iStartTime, 2) Step 1
                            If (tgMRdf(ilFound).iStartTime(0, ilLoop2) <> 1) Or (tgMRdf(ilFound).iStartTime(1, ilLoop2) <> 0) Then
                                gUnpackTimeLong tgMRdf(ilFound).iStartTime(0, ilLoop2), tgMRdf(ilFound).iStartTime(1, ilLoop2), False, llOvStartTime
                                gUnpackTimeLong tgMRdf(ilFound).iEndTime(0, ilLoop2), tgMRdf(ilFound).iEndTime(1, ilLoop2), True, llOvEndTime
                                llTotalSec = llTotalSec + (llOvEndTime - llOvStartTime)
                             Else
                                'llTotalSec = 0
                                'Exit For
                            End If
                        Next ilLoop2
                    End If
                End If

                If llTotalSec / 3600 >= 18 Then 'dp times (or override times greater or equal to 18 hours, consider ROS
                    slROS = "R"
                Else
                    slROS = "N"
                End If
            Else
                slAdvtCode = ""
                slCompCode = ""
                slPty = ""
                slFixed = ""
                slROS = ""
            End If
End Sub
'
'           mProcPRSEvent - create a new record for each string of text in the comment
'           separted by a semicolon (;).
'           The event type to process must be an "Other" type whose name is "PRS Comment"
'            The comment field contains a string of text.  each record created from that string
'           is separated by a ;
'           <input> ilautomationtype - automation option (this sub is for audio vault rps only, but
'                   needs to be sent to another subrotuine from here
'                   lldate - date processing
'                   llstarttime -earliest time of day
'                   llendtime - end time of day
'                   tlllc() - calendar of days events
Private Sub mProcRPSEvent(llDate As Long, llStartTime As Long, llEndTime As Long, tlLLC() As LLC)
Dim ilFound As Integer
Dim ilLLCLoop As Integer
Dim llTestTime As Long
Dim ilRet As Integer
Dim slStr As String
Dim slRecord As String
Dim ilStart As Integer
Dim ilLen As Integer
Dim slSemiColon As String * 1
Dim ilPos As Integer
Dim ilSub As Integer

            For ilLLCLoop = 0 To UBound(tlLLC) - 1
                llTestTime = gTimeToLong(tlLLC(ilLLCLoop).sStartTime, False)
                
                If llTestTime >= llStartTime And llTestTime <= llEndTime Then
                    If tlLLC(ilLLCLoop).sType = "Y" Then            'Other event in calendar
                        tmEtfSrchKey0.iCode = tlLLC(ilLLCLoop).iEtfCode
                        ilRet = btrGetEqual(hmEtf, tmEtf, imEtfRecLen, tmEtfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            If tlLLC(ilLLCLoop).lCefCode > 0 Then       'comment exists
                                tmCefSrchKey0.lCode = tlLLC(ilLLCLoop).lCefCode
                                imCefRecLen = Len(tmCef)
                                ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    ilSub = 1
                                    ilFound = True
                                    ilStart = 1
                                    slStr = gStripChr0(tmCef.sComment)
                                    'does text end with a ;.  If not, place one at the end
                                    ilLen = Len(Trim$(slStr))
                                    If ilLen > 0 Then
                                        slSemiColon = Mid$(slStr, ilLen, 1)
                                        If slSemiColon <> ";" Then
                                            slStr = slStr & ";"
                                        End If

                                        Do While ilFound = True
                                            ilPos = InStr(ilStart, slStr, ";")
                                            If ilPos > 0 Then           'found a text to create
                                                'every comment record is |time|, followed by whatever the user enters in the comments section
                                                'of a RPS Comment.  A new record is created for each group of text followed by a ;
                                                'i.e. +|Indicator5||X;+|ID05||;   2 export records are created
                                                '|time|+|Indicator5||X
                                                '|time|+|ID05||
                                                ilLen = ilPos - ilStart
                                                slRecord = "|" & tlLLC(ilLLCLoop).sStartTime & "|"
                                                slRecord = slRecord & Trim$(Mid(slStr, ilStart, ilLen))
                                                ilStart = ilPos + 1   'get past the text and ;
                                                mSaveImageForRPS slRecord, Format$(llDate, "m/d/yy"), llTestTime, "B", tlLLC(ilLLCLoop).sXMid, ilLLCLoop, ilSub
                                                ilSub = ilSub + 1
                                            Else
                                                ilFound = False
                                            End If
                                        Loop
                                    End If     'ilLen > 0
                                End If  'btrv_err_none
                            End If      'tlLLC(ilLLCLoop).lCefCode > 0
                        End If          'btrv_err_none
                    End If              'tlLLC(ilLLCLoop).sType = "Y"
                End If                  'llTestTime >= llStartTime And llTestTime <= llEndTime
            Next ilLLCLoop              'next calendar event, look only for other (comments)
            Exit Sub
End Sub
'
'            mSaveImageForRPS - save image in array of records so
'           they can be sorted for Audio Vault RPS option only
'           This automation export uses "Other" events defined as
'           comment name "RPS Comment".  Each text separated by a ";"
'           is created as a separate record.  In order to keep these
'           records in their correct sequence, use the library index along
'           with a sub number to keep in order.
'           Library index is used for multiple events at the same time:
'           for example a comment before an avail.  The sub number is
'           used when the comment library event has a comment defined with
'           multiple ";" embedded.
'           <input>
'           slRecord - the string containing automation format
'           slairdate - date of show (vehicle)
'           llSpotTime - time of event
'           slSortType - A,  B or C for sorting
'           slXMid - Y/No for cross midnight flag
'           ilLLCIndex - index of event into library array
'           ilSubNo - sub number for Library index for "Other" comment events
'
Private Sub mSaveImageForRPS(slRecord As String, slAirDate As String, llSpotTime As Long, slSortType As String, slXMid As String, ilLLCIndex As Integer, ilSub As Integer)
Dim slSortDate As String
Dim slStr As String
Dim llXMidDate As Long
Dim slLLCIndex As String
Dim slSub As String

    
    If slXMid = "Y" Then                    'cross midnight flag, adjust date if crossing midnight
        llXMidDate = gDateValue(slAirDate)
        slSortDate = Format$(llXMidDate + 1, "m/d/yy")
    Else
        slSortDate = Trim$(str$(gDateValue(slAirDate)))
    End If
    
    Do While Len(slSortDate) < 6                'date
        slSortDate = "0" & slSortDate
    Loop
    slStr = Trim$(str$(llSpotTime))             'time of event
    Do While Len(slStr) < 6
        slStr = "0" & slStr
    Loop
    
    slLLCIndex = Trim$(str$(ilLLCIndex))     'library index (due to comments, multiple events could be at same time)
    Do While Len(slLLCIndex) < 4
        slLLCIndex = "0" & slLLCIndex
    Loop
    
    slSub = Trim$(str$(ilSub))        'multiple events within 1 "Other" event type, defined in the comments field
    Do While Len(slSub) < 3
        slSub = "0" & slSub
    Loop
    
    tmExpRecImage(UBound(tmExpRecImage)).sKey = slSortDate & slStr & slSortType & slLLCIndex & slSub
    
    tmExpRecImage(UBound(tmExpRecImage)).sRecord = slRecord
    ReDim Preserve tmExpRecImage(0 To UBound(tmExpRecImage) + 1) As EXPRECIMAGE
    Exit Sub
End Sub
'
'       Update LCF record with either "I" (not complete) or "C" (complete
'       <input>  Date to search
'                ilType (game # or 0 for regular programming)
'               ilVehCode - vehicle code
'       <return> None
Public Sub mUpdateLCFCompleteFlag(ilSsfDate0 As Integer, ilSsfDate1 As Integer, ilType As Integer, ilVehCode As Integer, slCompleteFlag As String)
Dim ilLcfRet As Integer

        tmLcfSrchKey0.iLogDate(0) = ilSsfDate0
        tmLcfSrchKey0.iLogDate(1) = ilSsfDate1
        tmLcfSrchKey0.iSeqNo = 1
        tmLcfSrchKey0.iType = ilType
        tmLcfSrchKey0.iVefCode = ilVehCode
        tmLcfSrchKey0.sStatus = "C"
        ilLcfRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If ilLcfRet = BTRV_ERR_NONE Then
            tmLcf.sAffPost = slCompleteFlag
            ilLcfRet = btrUpdate(hmLcf, tmLcf, imLcfRecLen)
        End If
        Exit Sub
End Sub
'
'           gRemoveIllegalChar- strip illegal characters, retain any commas
'
Public Function gStripIllegalChr(slInStr As String) As String
    Dim slStr As String
    Dim slChr As String
    Dim ilIndex As Integer
    
    slStr = ""
    If Len(slInStr) > 0 Then
        ilIndex = 1
        Do While ilIndex <= Len(slInStr)
            slChr = Mid(slInStr, ilIndex, 1)
            If ((Asc(slChr) >= Asc(" ")) And (Asc(slChr) <= Asc("~"))) Then
                slStr = slStr & slChr
            End If
            ilIndex = ilIndex + 1
        Loop
    End If
    gStripIllegalChr = Trim$(slStr)
End Function

Private Sub mHighlightVehicles(ilIndex As Integer)
    Dim ilVff As Integer
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilRet As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim ilValue As Integer
    
    If rbcAutoType(0).Value Then    'Dalet
    ElseIf rbcAutoType(2).Value Then    'Prophet nexgen
    ElseIf rbcAutoType(5).Value Then    'Scott
    ElseIf rbcAutoType(1).Value Then    'Drake
    ElseIf rbcAutoType(4).Value Then    'RCS
    ElseIf rbcAutoType(3).Value Then      '2-5-03 prophet wizard
    ElseIf rbcAutoType(6).Value Then      '2-5-03 prophet media star
    ElseIf rbcAutoType(7).Value Then        '6-25-05 iMediatouch
    ElseIf rbcAutoType(8).Value Then        '8-10-05 Aduio Vault Sat
    ElseIf rbcAutoType(9).Value Then        '9-11-06 WireReady
    ElseIf rbcAutoType(10).Value Then       '8-21-08 Simian
    ElseIf rbcAutoType(11).Value Then
    ElseIf rbcAutoType(12).Value Then       '2/1/10: Rivendell
    ElseIf rbcAutoType(13).Value Then       '11-15-10 audio vault prs
    ElseIf rbcAutoType(14).Value Then       '2-18-11 Audio Vault AIR (same as WireReady)
    ElseIf rbcAutoType(15).Value Then       '1-6-12 Wide Orbit (same as Scott)
    ElseIf rbcAutoType(16).Value Then       '6-21-12 Jelli
    ElseIf rbcAutoType(17).Value Then       '10/16/12 Enco-ESPN
        If ckcAll.Value = vbChecked Then
            ckcAll.Value = vbUnchecked
        Else
            imAllClicked = True
            llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
            'llRet = SendMessageByNum(lbcVehicle.hwnd, LB_SELITEMRANGE, vbFalse, llRg)
            ilValue = True
            llRet = SendMessageByNum(lbcVehicle.HWnd, LB_SELITEMRANGE, ilValue, llRg)
            imAllClicked = False
        End If
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tgVff(ilVff).sExportEncoESPN = "Y" Then
                For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
                    slNameCode = tgUserVehicle(ilVef).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilVefCode = Val(slCode)
                    If ilVefCode = tgVff(ilVff).iVefCode Then
                        lbcVehicle.Selected(ilVef) = True
                    End If
                Next ilVef
            End If
        Next ilVff
    ElseIf rbcAutoType(18).Value Then       '8-16-13 scott v5
    ElseIf rbcAutoType(19).Value Then       '1-5-16 Zetta
    ElseIf rbcAutoType(20).Value Then       '5-11-18
    End If
End Sub

Private Sub mReSetBreakNumbers(ilAutomationType As Integer)
    Dim ilLine As Integer
    Dim ilPoint1 As Integer
    Dim ilPoint2 As Integer
    Dim ilPoint3 As Integer
    Dim ilBreakNumber As Integer
    Dim slHour As String
    Dim ilPrevHour As Integer
    Dim slScan As String
    Dim slPosition As String
    Dim ilPosition As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slLen As String
    Dim llDate As Long
    Dim llPrevDate As Long
    Dim llAvailTime As Long
    '11/20/13: TTP 6507
    'Dim llNextAvailTime As Long
    Dim llPrevAvailTime As Long
    
    Dim slHBP As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilParentVefCode As Integer
    Dim blFound As Boolean
    Dim ilIndex As Integer
    Dim ilPrevIndex As Integer
    
    If ilAutomationType = AUTOTYPE_ENCOESPN Then
        ReDim tmBreakByProg(0 To 0) As BREAKBYPROG
        llPrevDate = -1
        ilPrevHour = -1
        For ilLine = LBound(smNewLines) To UBound(smNewLines) - 1 Step 1
            '11/20/13: TTP 6507
            'ilPoint1 = InStr(1, smNewLines(ilLine), "^", vbTextCompare)
            'ilParentVefCode = Val(Mid(smNewLines(ilLine), ilPoint1 + 1))
            ilPoint1 = InStr(1, smNewLines(ilLine), "^", vbTextCompare)
            ilPoint2 = InStr(ilPoint1 + 1, smNewLines(ilLine), "^", vbTextCompare)
            ilParentVefCode = Val(Mid(smNewLines(ilLine), ilPoint1 + 1, ilPoint2 - ilPoint1 - 1))
            llAvailTime = Val(Mid(smNewLines(ilLine), ilPoint2 + 1))
            
            smNewLines(ilLine) = Left(smNewLines(ilLine), ilPoint1 - 1)
            blFound = False
            For ilLoop = 0 To UBound(tmBreakByProg) - 1 Step 1
                If tmBreakByProg(ilLoop).iVefCode = ilParentVefCode Then
                    blFound = True
                    ilIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
            If Not blFound Then
                ilIndex = UBound(tmBreakByProg)
                tmBreakByProg(ilIndex).iVefCode = ilParentVefCode
                tmBreakByProg(ilIndex).iBreakNo = 0
                tmBreakByProg(ilIndex).iPositionNo = 0
                ReDim Preserve tmBreakByProg(0 To ilIndex + 1) As BREAKBYPROG
            End If
            ilRet = gParseItem(smNewLines(ilLine), 1, ",", slDate)
            llDate = gDateValue(slDate)
            If llDate <> llPrevDate Then
                ilPrevHour = -1
            End If
            ilRet = gParseItem(smNewLines(ilLine), 2, ",", slTime)
            '11/20/13: TTP 6507
            'llAvailTime = gTimeToLong(slTime, False)
            slHour = Left$(gFormatSpotTime(llAvailTime), 2)
            If Val(slHour) <> ilPrevHour Then
                For ilLoop = 0 To UBound(tmBreakByProg) - 1 Step 1
                    tmBreakByProg(ilLoop).iBreakNo = 0
                    tmBreakByProg(ilLoop).iPositionNo = 0
                Next ilLoop
                '11/20/13: TTP 6507
                'llNextAvailTime = -1
                llPrevAvailTime = -1
            Else
                If ilIndex <> ilPrevIndex Then
                    '11/20/13: TTP 6507
                    'llNextAvailTime = -1
                    llPrevAvailTime = -1
                End If
            End If
            ilPrevIndex = ilIndex
            ilBreakNumber = tmBreakByProg(ilIndex).iBreakNo
            ilPosition = tmBreakByProg(ilIndex).iPositionNo
            '11/20/13: TTP 6507
            'If llNextAvailTime <> llAvailTime Then
            If llPrevAvailTime <> llAvailTime Then
                ilBreakNumber = ilBreakNumber + 1
                ilPosition = 1
            Else
                ilPosition = ilPosition + 1
            End If
            tmBreakByProg(ilIndex).iBreakNo = ilBreakNumber
            tmBreakByProg(ilIndex).iPositionNo = ilPosition
            slHBP = "H" & slHour & "B" & Trim$(str$(ilBreakNumber)) & "P" & Trim$(str$(ilPosition))
            llPrevDate = gDateValue(slDate)
            ilPrevHour = Val(slHour)
            ilRet = gParseItem(smNewLines(ilLine), 3, ",", slLen)
            '11/20/13: TTP 6507
            'llNextAvailTime = llAvailTime + Val(slLen)
            llPrevAvailTime = llAvailTime   ' + Val(slLen)
            ilPoint1 = InStr(1, smNewLines(ilLine), ",", vbTextCompare)
            ilPoint2 = InStr(1, smNewLines(ilLine), ":", vbTextCompare)
            If (ilPoint1 > 0) And (ilPoint2 > 0) And (ilPoint1 < ilPoint2) Then
                slHour = Mid(smNewLines(ilLine), ilPoint1 + 1, ilPoint2 - ilPoint1 - 1)
                slScan = "H" & slHour & "B"
                ilPoint1 = InStr(1, smNewLines(ilLine), slScan, vbTextCompare)
                ilPoint2 = InStr(ilPoint1 + 4, smNewLines(ilLine), "P", vbTextCompare)
                If (ilPoint1 > 0) And (ilPoint2 > 0) And (ilPoint1 < ilPoint2) Then
                    ilPoint3 = InStr(ilPoint2 + 1, smNewLines(ilLine), ",", vbTextCompare)
                    smNewLines(ilLine) = Left(smNewLines(ilLine), ilPoint1 - 1) & slHBP & Mid(smNewLines(ilLine), ilPoint3)
                End If
            End If
        Next ilLine
    End If
End Sub
Private Sub mLockAvails(blInProcSpot As Boolean, ilVpfIndex As Integer, llDate As Long, llEDate As Long)
    Dim llEndDate As Long
    
    If (tgVpf(ilVpfIndex).sGenLog = "N") And (blInProcSpot) Then
        llEndDate = gDateValue(edcEndDate.Text)
        If llDate = llEndDate + 1 Then
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), smLockDate
            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", smLockStartTime
            smLockEndTime = gFormatTimeLong(gTimeToLong(smLockStartTime, False) + 1, "A", "1")
            If (imLockVefCode <> tmSdf.iVefCode) Or (gTimeToLong(smLockStartTime, False) <> lmLockStartTime) Then
                gSetLockStatus tmSdf.iVefCode, 1, -1, smLockDate, smLockDate, tmSdf.iGameNo, smLockStartTime, smLockEndTime
                imLockVefCode = tmSdf.iVefCode
                lmLockStartTime = gTimeToLong(smLockStartTime, False)
            End If
        End If
    End If
End Sub
Private Function mSplitCopy() As Boolean
    Dim ilRet As Integer
    
    tmRsfSrchKey1.lCode = tmSdf.lCode
    ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode) Then
        tmAxfSrchKey1.iCode = tmSdf.iAdfCode
        ilRet = btrGetEqual(hmAxf, tmAxf, imAxfRecLen, tmAxfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            mSplitCopy = True
        Else
            mSplitCopy = False
        End If
    Else
        mSplitCopy = False
    End If
End Function

Private Sub mProcAdjDate(slAirDate As String, llDate As Long, ilAirHour As Integer, ilLocalHour As Integer, blInProcSpot As Boolean, ilVpfIndex As Integer, llEDate As Long)
    'Test if Air time is AM and Feed Time is PM. If so, adjust date
    slAirDate = Format$(llDate, "m/d/yy")
    ilAirHour = tmDlf.iAirTime(1) \ 256  'Obtain Hour
    ilLocalHour = tmDlf.iLocalTime(1) \ 256  'Obtain Hour
    If (ilAirHour < 6) And (ilLocalHour > 17) Then
        mLockAvails blInProcSpot, ilVpfIndex, llDate, llEDate
        slAirDate = gDecOneDay(slAirDate)
    End If
End Sub

Private Sub mProcSpotTime(llAirDate As Long, llAvailTime As Long, ilVehCode As Integer, slCodeStn As String, llSpotTime As Long, llPrevTime As Long, ilShowTime As Integer)
    Dim ilTest As Integer
    
    For ilTest = LBound(tmSpotTimes) To UBound(tmSpotTimes) - 1 Step 1
        If (tmSpotTimes(ilTest).iVefCode = ilVehCode) And (tmSpotTimes(ilTest).sCodeStn = slCodeStn) And (tmSpotTimes(ilTest).lAirDate = llAirDate) And (tmSpotTimes(ilTest).lAvailTime = llAvailTime) Then
            ilShowTime = False
            llSpotTime = tmSpotTimes(ilTest).lNextSpotTime
            tmSpotTimes(ilTest).lNextSpotTime = tmSpotTimes(ilTest).lNextSpotTime + tmSdf.iLen
            llPrevTime = tmSpotTimes(ilTest).lNextSpotTime
            'Return
            Exit Sub
        End If
    Next ilTest
    ReDim Preserve tmSpotTimes(0 To UBound(tmSpotTimes) + 1) As SPOTTIMES
    tmSpotTimes(UBound(tmSpotTimes) - 1).iVefCode = ilVehCode
    tmSpotTimes(UBound(tmSpotTimes) - 1).sCodeStn = slCodeStn
    tmSpotTimes(UBound(tmSpotTimes) - 1).lAirDate = llAirDate
    tmSpotTimes(UBound(tmSpotTimes) - 1).lAvailTime = llAvailTime
    tmSpotTimes(UBound(tmSpotTimes) - 1).lNextSpotTime = llAvailTime + tmSdf.iLen
    llSpotTime = llAvailTime
    If llPrevTime = llAvailTime Then
        ilShowTime = False
    Else
        ilShowTime = True
    End If
    llPrevTime = tmSpotTimes(UBound(tmSpotTimes) - 1).lNextSpotTime
End Sub

Private Sub mTimeZoneAdj(ilAdjZone As Integer, llAdjTime As Long, ilDlfFound As Integer, llAirDate As Long, llAvailTime As Long, blInProcSpot As Boolean, ilVpfIndex As Integer, llDate As Long, llEDate As Long)
    ilAdjZone = False
    llAdjTime = 0
    If (Trim$(tmDlf.sZone) <> "") And (ilDlfFound) Then
        Select Case Left$(tmDlf.sZone, 1)
            Case "E"
                If rbcZone(0).Value = False Then
                    ilAdjZone = True
                    If rbcZone(1).Value Then
                        llAdjTime = -3600
                    ElseIf rbcZone(2).Value Then
                        llAdjTime = -7200
                    ElseIf rbcZone(3).Value Then
                        llAdjTime = -10800
                    End If
                End If
            Case "C"
                If rbcZone(1).Value = False Then
                    ilAdjZone = True
                    If rbcZone(0).Value Then
                        llAdjTime = 3600
                    ElseIf rbcZone(2).Value Then
                        llAdjTime = -3600
                    ElseIf rbcZone(3).Value Then
                        llAdjTime = -7200
                    End If
                End If
            Case "M"
                If rbcZone(2).Value = False Then
                    ilAdjZone = True
                    If rbcZone(0).Value Then
                        llAdjTime = 7200
                    ElseIf rbcZone(1).Value Then
                        llAdjTime = 3600
                    ElseIf rbcZone(3).Value Then
                        llAdjTime = -3600
                    End If
                End If
            Case "P"
                If rbcZone(3).Value = False Then
                    ilAdjZone = True
                    If rbcZone(0).Value Then
                        llAdjTime = 10800
                    ElseIf rbcZone(1).Value Then
                        llAdjTime = 7200
                    ElseIf rbcZone(2).Value Then
                        llAdjTime = 3600
                    End If
                End If
        End Select
    Else
        If rbcZone(0).Value = False Then
            ilAdjZone = True
            If rbcZone(1).Value Then
                llAdjTime = -3600
            ElseIf rbcZone(2).Value Then
                llAdjTime = -7200
            ElseIf rbcZone(3).Value Then
                llAdjTime = -10800
            End If
        End If
    End If
    If ilAdjZone Then
        llAvailTime = llAvailTime + llAdjTime
        If llAvailTime < 0 Then
            llAvailTime = llAvailTime + 86400
            llAirDate = llAirDate - 1
            mLockAvails blInProcSpot, ilVpfIndex, llDate, llEDate
        ElseIf llAvailTime >= 86400 Then
            llAvailTime = llAvailTime - 86400
            llAirDate = llAirDate + 1
        End If
    End If

End Sub


Private Sub mProphetBlockTime(slDelimiter As String, ilAutomationType As Integer, slRecord As String, ilSameBreak As Integer, ilDlfFound As Integer, ilAdjZone As Integer, llAdjTime As Long, slAirDate As String, llAirDate As Long, llAvailTime As Long, llSpotTime As Long, blInProcSpot As Boolean, ilVpfIndex As Integer, slDate As String, llDate As Long, llSDate As Long, llEDate As Long, slAdvtName As String, slSortType As String, slSortDate As String, ilRet As Integer)
    Dim llTempTime As Long
    Dim slStr As String
    Dim slHour As String
    Dim llHour As Long
    Dim llMin As Long
    Dim llSec As Long
    Dim ilFound As Integer
    Dim ilTest As Integer
    
    slDelimiter = ""
    If ilAutomationType = 7 Then    'Prophet MediaStar (Same as Prohet Wizard except 5 blanks in front of time)
        slRecord = "     "
    ElseIf ilAutomationType = 9 Then    'audio vault sat
        slRecord = "|"
    Else
        slRecord = ""
    End If
    'Time (HH:MM:SS)
    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTempTime
    If ilAutomationType = 9 Then            'audio vaultsat, decrease the align format time back 10"
        If llTempTime - 10 >= 0 Then        'check for start of day so the computation doesnt get negative
            llTempTime = llTempTime - 10
        End If
        slDelimiter = "|@|Align|Automatic time Alignment|:"           'Example: 23:40:50|@|Align|Automatic time Alignment|:
        ilSameBreak = True                   'if spots are all back to back, regardless if in different avails, they are
                                             'continuation spots.  flag them in export with "+"
    End If
    If Not ilDlfFound Then
        llAvailTime = llTempTime
        llAirDate = gDateValue(slDate)
        '6/1/16: Replace GoSub
        'GoSub lTimeZoneAdj
        mTimeZoneAdj ilAdjZone, llAdjTime, ilDlfFound, llAirDate, llAvailTime, blInProcSpot, ilVpfIndex, llDate, llEDate
        llTempTime = llAvailTime
        slAirDate = Format$(llAirDate, "m/d/yy")
        If llAirDate > llEDate Then
            'Return
            Exit Sub
        End If
        If llAirDate < llSDate Then
            'Return
            Exit Sub
        End If
    End If
    llHour = llTempTime \ 3600
    llMin = llTempTime Mod 3600
    llSec = llMin Mod 60
    llMin = llMin \ 60
    slStr = Trim$(str$(llHour))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slHour = Trim$(slStr)
    slRecord = slRecord & slStr & ":"
    slStr = Trim$(str$(llMin))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slRecord = slRecord & slStr
    slStr = Trim$(str$(llSec))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slRecord = slRecord & ":" & slStr & Trim$(slDelimiter)
    'Cart Number filler
    slStr = ""

        If ilAutomationType = 7 Then    'Prophet MediaStar (Same as Prohet Wizard except 5 blanks in front of time)
            Do While Len(slStr) < 25
                slStr = slStr & " "
            Loop

        Else
            Do While Len(slStr) < 20
                slStr = slStr & " "
            Loop
        End If

        slRecord = slRecord & slStr
        'BLOCK description in advt/prod field
        If ilAutomationType = 9 Then                'audio vault sat
            slAdvtName = ""
        Else
            slAdvtName = "Block#"
        End If
        Do While Len(slAdvtName) < 66
            slAdvtName = slAdvtName & " "
        Loop
        slRecord = slRecord & slAdvtName

        'Spot Length filler
        slStr = ""
        Do While Len(slStr) < 4
            slStr = " " & slStr
        Loop
        slRecord = slRecord & slStr
        slStr = ""      'spot code
        Do While Len(slStr) < 10
            slStr = " " & slStr
        Loop
        slRecord = slRecord & slStr
        ilRet = 0
        'smNewLines(UBound(smNewLines)) = slRecord
        'ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 118
        '4/2/08:  Moved up to time zone adjustment
        'slAirDate = slDate
        llSpotTime = llTempTime
        slSortType = "A"
        slSortDate = Trim$(str$(gDateValue(slAirDate)))
        Do While Len(slSortDate) < 6
            slSortDate = "0" & slSortDate
        Loop
        slStr = Trim$(str$(llSpotTime))
        Do While Len(slStr) < 6
            slStr = "0" & slStr
        Loop

    ilFound = False
    For ilTest = 0 To UBound(tmExpRecImage) - 1 Step 1
        If slSortDate & slStr & slSortType = Trim$(tmExpRecImage(ilTest).sKey) Then
            ilFound = True
            Exit For
        End If
    Next ilTest
    If Not ilFound Then
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    End If
End Sub

Private Sub mSaveRecImage(slAirDate As String, llSpotTime As Long, slSortDate As String, slSortType As String, slRecord As String)
    Dim slStr As String
    
    If imGameVehicle Then
        If lmLastEvtTime = -1 Then
            lmLastEvtTime = llSpotTime
        Else
            If (llSpotTime < lmLastEvtTime) And (Not imXMidNight) Then
                imXMidNight = True
            End If
        End If
        If imXMidNight Then
            slSortDate = Trim$(str$(gDateValue(slAirDate) + 1))
        Else
            slSortDate = Trim$(str$(gDateValue(slAirDate)))
        End If
    Else
        slSortDate = Trim$(str$(gDateValue(slAirDate)))
    End If
    Do While Len(slSortDate) < 6
        slSortDate = "0" & slSortDate
    Loop
    slStr = Trim$(str$(llSpotTime))
    Do While Len(slStr) < 6
        slStr = "0" & slStr
    Loop
    tmExpRecImage(UBound(tmExpRecImage)).sKey = slSortDate & slStr & slSortType
    tmExpRecImage(UBound(tmExpRecImage)).sRecord = slRecord
    ReDim Preserve tmExpRecImage(0 To UBound(tmExpRecImage) + 1) As EXPRECIMAGE
End Sub

Private Sub mProcTestForDuplBB(ilAddSpot As Integer)
    Dim ilBB As Integer
    
    ilAddSpot = True
    If tgSpf.sUsingBBs <> "Y" Then
        'Return
        Exit Sub
    End If
    If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
        'Return
        Exit Sub
    End If
    For ilBB = 0 To UBound(tmBBSdfInfo) - 1 Step 1
        If tmBBSdfInfo(ilBB).sType <> "A" Then
            If (tmBBSdfInfo(ilBB).lChfCode = tmSdf.lChfCode) And (tmBBSdfInfo(ilBB).iLen = tmSdf.iLen) And (tmBBSdfInfo(ilBB).sType = tmSdf.sSpotType) Then
                If (tmBBSdfInfo(ilBB).iTime(0) = tmSdf.iTime(0)) And (tmBBSdfInfo(ilBB).iTime(1) = tmSdf.iTime(1)) Then
                    ilAddSpot = False
                    'Return
                    Exit Sub
                End If
            End If
        End If
    Next ilBB
    tmBBSdfInfo(UBound(tmBBSdfInfo)).sType = tmSdf.sSpotType
    tmBBSdfInfo(UBound(tmBBSdfInfo)).lChfCode = tmSdf.lChfCode
    tmBBSdfInfo(UBound(tmBBSdfInfo)).iLen = tmSdf.iLen
    tmBBSdfInfo(UBound(tmBBSdfInfo)).iTime(0) = tmSdf.iTime(0)
    tmBBSdfInfo(UBound(tmBBSdfInfo)).iTime(1) = tmSdf.iTime(1)
    ReDim Preserve tmBBSdfInfo(0 To UBound(tmBBSdfInfo) + 1) As BBSDFINFO
End Sub

Private Sub mProcSpot(ilAutomationType As Integer, ilAdjZone As Integer, llAdjTime As Long, ilDlfFound As Integer, slAirDate As String, llAirDate As Long, llAvailTime As Long, llSpotTime As Long, llPrevTime As Long, ilShowTime As Integer, llSpotEndTime As Long, blInProcSpot As Boolean, ilVefCode As Integer, ilVehCode As Integer, ilVpfIndex As Integer, slDate As String, llDate As Long, llSDate As Long, llEDate As Long, ilAirHour As Integer, ilLocalHour As Integer, ilCopy As Integer, llCopyMissingSdfCode As Long, slSortType As String, slSortDate As String, slRecord As String, slLogYear As String, slLogMonth As String, slLogDay As String, slEventName As String, llESPNPrevDate As Long, ilESPNHour As Integer, ilESPNBreak As Integer, ilESPNPosition As Integer, llESPNPrevAvailTime As Long, ilParentVefCode As Integer, ilAtLeastOneSpot As Integer, ilSameBreak As Integer, ilRet As Integer, slMsg As String)

    'Record format for Prophet(wizard & nextgen), Dalet (Fixed Length):
    'Column  Length  Field
    '  1        8    Start Time HR:MN:SC  Military Hours
    '  9       20    Cart Number (w/ material code) or ISCI code
    ' 29       46    Advertiser Name (30) / Product name (15)
    ' 75       20    Creative Title
    ' 95        4    Length  MMSS
    ' 99       10    SDF Code
    '109        1    c/r
    '110        1    l/f

    'following 5 fields for nextgen ONLY)  4-7-06
    '109        1    blank
    '110        1    ROS flag (R/N)
    '111        1    blank
    '112        6    5-10-06 Pty code was Fixed buy (F/M)
    '118        1    blank
    '119        1    5-10-06 fixed buy, was Pry code
    '120        1    blank
    '121        5    Advertiser code
    '126        1    blank
    '127        5    competitive code (chfcompcode(0))
    '132        1    c/r
    '133        1    l/f


    'Record format for Prophet, MediaStar:
    'Column  Length  Field
    '  6        8    Start Time HR:MN:SC  Military Hours
    ' 15        5    Cart Number (w/o material code)
    ' 21       10    Spot ID
    ' 39       29    Advertiser Name (30)/Product (remaining characters) , plus slash
    ' 69        1    * asterisk to indicate that this spot is from CSI
    ' 70        4    Length  SSSS
    '           1    c/r
    '           1    l/f

 'Record format for RCS (Fixed length):
    'Column  Length  Field
    '  1        1    Commercial Indicator always C
    '  2        7    Start Time HRMN:SC  Military Hours
    '  9        4    Cart Number (without material code) or first 4 char of ISCI
    ' 13       24    Advertiser Abbr/Product
    ' 37        3    Unused (Priority number)
    ' 40        4    Length  MNSC (60=> 0100; 90=> 0130; 30=>0030)
    ' 44        4    Unused (Commercial Type)
    ' 48        6    First Part of Sdf.lCode (was Unused (Customer ID))
    ' 54        2    Second Part of Sdf.lCode (was Unused (Internal Code))
    '                use 8 bytes to make up sdd.lcode
    ' 56        4    Unused (Product Code)
    ' 60        8    Unused (Ordered Time)
    ' 68        1    Carriage Return <cr>
    ' 69        1    Line Feed <lf>
    '
    
'Record format for RCS-5 (Fixed length):
    'Column  Length  Field
    '  1        1    Commercial Indicator always C
    '  2        7    Start Time HRMN:SC  Military Hours
    '  9       15    Cart Number (without material code) or first 5 char of ISCI
    ' 24       24    Advertiser Abbr/Product
    ' 48        3    Unused (Priority number)
    ' 51        4    Length  MNSC (60=> 0100; 90=> 0130; 30=>0030)
    ' 55        4    Unused (Commercial Type)
    ' 59        6    First Part of Sdf.lCode (was Unused (Customer ID))
    ' 65        2    Second Part of Sdf.lCode (was Unused (Internal Code))
    '                use 8 bytes to make up sdd.lcode
    ' 67        4    Unused (Product Code)
    ' 71        8    Unused (Ordered Time)
    ' 79       24    Unused (Sponsor 1)
    '103       24    Unused (Sponsor 2)
    '127       24    Unused (Comment)
    '151        8    Unused (Start Date mm/dd/yy)
    '159        5    Unused (Start time hh:mm)
    '164        8    Unused (Stop Date)
    '172        5    Unused (Stop Time)
    '177        8    Unused (Kill Date)
    '185       15    Unused (Product 1)
    '200       15    Unused (Product 2)
    '215        1    Unused (Is Live? Y or N)
    '216        1    Unused (Is External? Y or N)
    '217        2    Unused (CBSI Stopset number)
    '219       35    Unused (Native-Specific)
    '254        1    Carriage Return <cr>
    '255        1    Line Feed <lf>

'Record Format for Scott
    ' Field #     Max length   Description
    '  1              8         Start Time HR:MN:SC Military time
    '  2              --        N/a
    '  3              3         Comml notation (CA)
    '  4              6         DA followed for cart # without material code
    '  5              66        Advt (max 30) / Prod (max 35) (in quotes)
    '  6              10        SDF Spot Code (in quotes)
    '  7              5         Length of spot MM:SS
    
'Record Format for Scott V5 (added 8-16-13)
    ' Field #     Max length   Description
    '  1              8         Start Time HR:MN:SC Military time
    '  2              --        N/a
    '  3              3         Comml notation (Material Code)
    '  4              6         cart # without material code
    '  5              66        Advt (max 30) / Prod (max 35) (in quotes)
    '  6              10        SDF Spot Code (in quotes)
    '  7              5         N/A
    '  9                        N/A
    '  10                       N/A
    '  11             6         Must be 6 blanks enclosed in "
    '  12                       N/A
    '  13                       N/A
    '  14                       N/A
    '  15                       N/A
    '  16                       N/A
    
'Record Format for Wide Orbit (comma delimited, variable length) 1-10-12
    ' Field #     Max length   Description
    '  1              8         Start Time HR:MN:SC Military time
    '  2              --        N/a
    '  3              3         Comml notation (CA)
    '  4              6         DA followed for cart # without material code
    '  5              45        Advt (max 30) / Prod (max 35) (in quotes)
    '  6              21        SDF Spot Code followed by ":", then station code all in quotes 1-10-12 station code added
    '  7              5         Length of spot MM:SS
    
 'Record format for iMedia Touch (fixed length, blank after each field)
 '   Field #       Max Length    Description
 '    1-8              8          Event Time HH:MM:SS   24hr time
 '    9                1          Blank
 '    10-14            5          Event Duration MM:SS
 '    15               1          Blank
 '    16-21            6          ZM#### ZM (Hard-coded) followed by cart#
 '    22               1          Blank
 '    23-25            3          Event index (Unused)
 '    26               1          Blank
 '    27-29            3          COM  (Hard-coded)
 '    30               1          Blank
 '    31-38            8          Live Copy identifier (unused)
 '    39               1          Blank
 '    40               1          Sync Char (unused)
 '    41               1          Blank
 '    42               1          Item function (unused)
 '    43               1          Blank
 '    44-73            30         20 char advertisr/prod, 10 char internal spot code
 '    74               1          Blank
 '    75               1          c/r
 '    76               1          linefeed

 'Record Format for Audio Vault Sat spot (vertical line delimited, variable length)
    ' Field #     Max length   Description
    '  1              8         Start Time HR:MN:SC Military time
    '  2              |
    '  3              1         blank or + (for other than 1st spot in break)
    '  4              |
    '  4              10        Cart # (with or without media definition based on the site option)
    '  5              |
    '  6              40       Advt/Prod (optional)
    '  7              |

    '11-17-10 Audio Vault RPS.  There are 2 record types:  Comments and spots
    ' Field #     Max length   Description
    ' Comment type record
    '  1              |
    '  2              8         Start Time HR:MN:SC Military time
    '  3              |         blank or + (for other than 1st spot in break)
    '  4              +
    '  5              |
    '  6              40       Comment

    '  1              |
    '  2              8         Start Time HR:MN:SC Military time
    '  3              |         blank or + (for other than 1st spot in break)
    '  4              +
    '  5              |
    '  6              4       cart (without media code)
    '  7              ||
    '  Following 2 fields may need to be removed.  Client needs to let us know if it interferes with the equipment
    '  8              20      Advertiser name
    '  9              10      Internal Spot Code
    
    '9-11-06 WireReady or 2-18-11 AudioVault Air
    'Record Format    Max Length    Description
    '1                  8           Scheduled Spot Time HH:MM:SS  military tie
    '9                  20          Cart # (w/material Code) or ISCI
    '29                 46          Advt name/prod name (up to 30 char for advt name)
    '75                 20          Creative Title
    '95                 4           Spot Length MMSS
    '99                 10          Spot Internal code
    '109                1           C/R
    '110                1           L/F
    '
 'Record format for Rivendell (fixed length, blank after each field)
 '   Field #       Max Length    Description
 '    1-8              8          Event Time HH:MM:SS   24hr time
 '    9                1          Blank
 '    10-10            1          Single alpha-numeric chanracter, ignored by Rivendell
 '    11-24            14         cart#, numeric, 6 digits max, right padded with spaces
 '    25               1          Blank
 '    26-59            34         Cart Title, right padded with spaces
 '    60               1          Blank
 '    61-68            8          Spot Length HH:MM:SS
 '    69               1          Blank
 '    70-101           32         Generic element ISCI Code, right padded with spaces
 '    102              1          Blank
 '    103-134          32         Event GUID, Spot internal code
 '    135              1          c/r
 '    136              1          linefeed


    '11-17-10 Audio Vault RPS.  There are 2 record types:  Comments and spots
    ' Field #     Max length   Description
    ' Comment type record
    '  1              |
    '  2              10        Start Time HR:MN:SCAM/PM (06:12:30PM)
    '  3              |
    '  4              40 variable length- Comment (new record created for each string separated by ;

    'spot
    '  1              |
    '  2              10         Start Time HR:MN:SCAM/PM (06:12:30PM)
    '  3              |
    '  4              +
    '  5              |
    '  6              4       cart (without media code)
    '  7              ||
    '  8              20      Advertiser name
    '  9              10      Internal Spot Code
    '3/3/12: Test if Media code set to suppress spots.  This flag is obtained from gObtainCopy call
    
    '6-21-12 Jelli Comma delimited
    '   Field Name       Format          Max Length
    '   ISCI Code        AN              20
    '   Start Date/Time  YYYY-mm-dd hh:mm:ss    military 19
    '   Length           nnn             3
    '   Creative Title   AN              30
    '   Veh STation Code AN              5
    '   contract         N               9
    '   Advt Name        AN              30
    '   Agency Name      AN              40
    '   Internal Code    nnn             10

    '10/16/12 ENCO-ESPN Comma delimited
    '   Field Name       Format          Max Length
    '   Date             m/dd/yy         8
    '   Time             hh:mm:ss        8 (military)
    '   Length           nnn             3
    '   Event Name       AN              20
    '   ISCI Code        AN              20
    '   Advt/Prod        AN              35
    '   ProgCodeHnBnPn   AN
    '   Internal Code    nnn             10

'Record format for Zetta(Fixed length):  All unused fields with blanks
'this is a subset of RCS-5
    'Column  Length  Field
    '  1        1    Commercial Indicator always C
    '  2        7    Start Time HRMN:SC  Military Hours
    '  9       15    Cart Number (3 char material code + 4 char cart #)
    ' 24       24    Creative Title
    ' 48        3    Unused (Priority number)
    ' 51        4    Length  MNSC (60=> 0100; 90=> 0130; 30=>0030)
    ' 55        4    Unused (Commercial Type)
    ' 59        10    CSI Internal SDF code
    ' 69        2    Unused (Product Code)
    ' 71        8    Unused (Ordered Time)
    ' 79       24    Advertiser Name
    '103       24    Agency Name (blank if direct)
    '127       24    Unused (Comment)
    '151        8    Unused (Start Date mm/dd/yy)
    '159        5    Unused (Start time hh:mm)
    '164        8    Unused (Stop Date)
    '172        5    Unused (Stop Time)
    '177        8    Unused (Kill Date)
    '185       15    Product Name
    '200       15    Unused (Product 2)
    '215        1    Live Y/N
    '216        1    Unused (Is External? Y or N) , Always N
    '217        2    Unused (CBSI Stopset number)
    
'   'Record format for StationPlaylist (similar to simian with ISCI added to end) 5-11-18:
'    'Column  Length  Field
'    '  1        8    Start Time HR:MN:SC  Military Hours
'    '  9       20    Cart Number (w/ material code) or ISCI code
'    ' 29       46    Advertiser Name (30) / Product name (15)
'    ' 75       20    Creative Title
'    ' 95        4    Length  MMSS
'    ' 99       10    SDF Code
'    '109       20    ISCI
'    '129        1    c/r
'    '130        1    l/f
    'Dim slMsg As String
    Dim slStr As String
    Dim slTemp As String
    Dim slLenInSec As String
    Dim slLenInMinSec As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilLLCLoop As Integer
    Dim ilFoundLLC As Integer
    Dim llTestTime As Long
    Dim ilLLCIndex As Integer
    Dim ilSub As Integer
    Dim slProgCodeID As String
    Dim slAdvtName As String
    Dim slAgyName As String
    Dim ilAddSpot As Integer
    Dim ilIndex As Integer
    Dim slHour As String
    Dim ilVff As Integer
    Dim slCodeStn As String
    
    ilRet = 0
    If smMediaCodeSuppressSpot = "Y" Then
        'Return
        Exit Sub
    End If
    If tmDlf.iMnfSubFeed > 0 Then
        tmMnfSrchKey.iCode = tmDlf.iMnfSubFeed
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        slCodeStn = Left$(tmMnf.sCodeStn, 5)
    Else
        slCodeStn = Left$(tmVef.sCodeStn, 5)
    End If
    blInProcSpot = True
    If ilDlfFound Then
        '6/1/16: Replace GoSub
        'GoSub lProcAdjDate  'Result stored into slAirDate
        mProcAdjDate slAirDate, llDate, ilAirHour, ilLocalHour, blInProcSpot, ilVpfIndex, llEDate
        llAirDate = gDateValue(slAirDate)
        gUnpackTimeLong tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), False, llAvailTime
    Else
        slAirDate = Format$(llDate, "m/d/yy")
        llAirDate = gDateValue(slAirDate)
        gUnpackTimeLong tmDlf.iLocalTime(0), tmDlf.iLocalTime(1), False, llAvailTime
        '6/1/16: Replace GoSub
        'GoSub lTimeZoneAdj
        mTimeZoneAdj ilAdjZone, llAdjTime, ilDlfFound, llAirDate, llAvailTime, blInProcSpot, ilVpfIndex, llDate, llEDate
        slAirDate = Format(llAirDate, "m/d/yy")
    End If
    blInProcSpot = False
    If llAirDate > llEDate Then
        'Return
        Exit Sub
    End If
    If llAirDate < llSDate Then
        'Return
        Exit Sub
    End If
    '6/1/16: Replace GoSub
    'GoSub lProcTestForDuplBB
    mProcTestForDuplBB ilAddSpot
    If Not ilAddSpot Then
        'Return
        Exit Sub
    End If
    '10-4-01
    'If ilProphetFlag = True Then
    '    GoSub lProphetBlockTime
    'End If
    'ilProphetFlag = False

    imFoundspot = True      '1-5-05 at least one spot to retain export file
    ilAtLeastOneSpot = True     '4-26-11 at least one spot found per game or vehicle to set day is incomplete flag
    '6/1/16: Replace GoSub
    'GoSub lProcSpotTime 'Result stored into slSpotTime
    mProcSpotTime llAirDate, llAvailTime, ilVehCode, slCodeStn, llSpotTime, llPrevTime, ilShowTime
    slRecord = ""
    If (ilAutomationType = 5) Or (ilAutomationType = 12) Then            'rcs only
        'Record type
        slRecord = "C"

        'Time
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = slRecord & Left(slStr, 2)    'HH
        slRecord = slRecord & Mid(slStr, 4, 2)  'MM
        slRecord = slRecord & Mid(slStr, 6, 3)  ':SS
        'Cart Number only
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        slRecord = slRecord & slStr

        'Advertiser/Product
        slAdvtName = Trim$(tmAdf.sAbbr) & "/" & Trim$(tmChf.sProduct)
        slAdvtName = Left$(slAdvtName, 24)
        Do While Len(slAdvtName) < 24
            slAdvtName = slAdvtName & " "
        Loop
        slRecord = slRecord & slAdvtName

        'Unused (3)
        slRecord = slRecord & "   "

        mFormatSpotLen tmSdf.iLen, slLenInSec, slLenInMinSec    'length in MMSS
        slRecord = slRecord & Left(slLenInMinSec, 2)
        slRecord = slRecord & Mid(slLenInMinSec, 3)
        'Unused (4)
        slRecord = slRecord & "    "

        ''Unused (6)
        'slRecord = slRecord & "      "
        '
        ''Unused (2)
        'slRecord = slRecord & "  "
        
        slStr = Trim$(str$(tmSdf.lCode))
        Do While Len(slStr) < 8
            slStr = " " & slStr
        Loop
        slRecord = slRecord & slStr

        'Unused (4)
        slRecord = slRecord & "    "

        'Unused (8)
        slRecord = slRecord & "        "
        
        If ilAutomationType = 12 Then
            Do While Len(slRecord) <= 253
                slRecord = slRecord & " "
            Loop
        End If

        ilRet = 0
        'smNewLines(UBound(smNewLines)) = slRecord
        'ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 118
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = AUTOTYPE_ZETTA Then               '1-6-16
     'Record type
        slRecord = "C"
        'Time
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = slRecord & Left(slStr, 2)    'HH
        slRecord = slRecord & Mid(slStr, 4, 2)  'MM
        slRecord = slRecord & Mid(slStr, 6, 3)  ':SS
        'Cart Number only
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)      '15 char field
        slRecord = slRecord & slStr

        'Creative Title
        Do While Len(smCreativeTitle) < 24
            smCreativeTitle = smCreativeTitle & " "
        Loop
        slRecord = slRecord & Mid(smCreativeTitle, 1, 24)

        'Unused (3)
        slRecord = slRecord & "   "

        mFormatSpotLen tmSdf.iLen, slLenInSec, slLenInMinSec    'length in MMSS
        slRecord = slRecord & Left(slLenInMinSec, 2)
        slRecord = slRecord & Mid(slLenInMinSec, 3)
        'Unused (4)
        slRecord = slRecord & "    "
        
        slStr = Trim$(str$(tmSdf.lCode))
        Do While Len(slStr) < 10
            slStr = slStr & " "
        Loop
        slRecord = slRecord & slStr

        'Unused (2)
        slRecord = slRecord & "  "

        'Unused (8)
        slRecord = slRecord & "        "
        'Advertiser/Product
        slAdvtName = Trim$(tmAdf.sName)
        slAdvtName = Left$(slAdvtName, 24)
        Do While Len(slAdvtName) < 24
            slAdvtName = slAdvtName & " "
        Loop
        slRecord = slRecord & slAdvtName
    
        slAgyName = ""
        If tmChf.iAgfCode > 0 Then      'agency name
            ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
            If ilIndex >= 0 Then
                slAgyName = Trim$(tgCommAgf(ilIndex).sName)
                slAgyName = gStripIllegalChr(slAgyName)
            End If
        End If
        
        Do While Len(slAgyName) < 24
            slAgyName = slAgyName & " "
        Loop
        slRecord = slRecord & Mid(slAgyName, 1, 24)
        
        slStr = ""
        Do While Len(slStr) < 58
            slStr = slStr & " "
        Loop
        slRecord = slRecord & slStr         'blank fill 58 char
        slStr = Trim$(tmChf.sProduct)
        Do While Len(slStr) < 15
            slStr = slStr & " "
        Loop
        slRecord = slRecord & Mid(slStr, 1, 15)       'product name
        
        slStr = ""
        Do While Len(slStr) < 15            'product name 2 , fill blanks
            slStr = slStr & " "
        Loop
        slRecord = slRecord & slStr
        
        If tmClf.sLiveCopy = "L" Or tmClf.sLiveCopy = "M" Then      'live coml or promo; pos. 215
            slRecord = slRecord & "Y"           'determine live
        Else
            slRecord = slRecord & "N"
        End If
        slRecord = slRecord & "N"           'Unused: external (Y/N)
        slRecord = slRecord & "  "          'unused: stopset
        
        ilRet = 0
        'smNewLines(UBound(smNewLines)) = slRecord
        'ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 118
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = AUTOTYPE_SCOTT Or ilAutomationType = AUTOTYPE_SCOTT_V5 Then             'scott
        'Time
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = Left(slStr, 8)               'use HH:MM:SS
        If ilAutomationType = AUTOTYPE_SCOTT Then
            slRecord = slRecord & ",,CA,DA"
        Else
            slRecord = slRecord & ",," & Trim$(Mid(smMcfPrefix, 1, 3)) & ",DA"
        End If
        'Cart Number only
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        slRecord = slRecord & Trim$(slStr) & ","

        'Advertiser/Product
        If Trim$(tmChf.sProduct) = "" Then      'omit slash if no product
            slAdvtName = Trim$(tmAdf.sName)
        Else
            slAdvtName = Trim$(tmAdf.sName) & "/" & Trim$(tmChf.sProduct)
        End If
        slRecord = slRecord & """" & slAdvtName & """" & ","

        slStr = Trim$(str$(tmSdf.lCode))
        slRecord = slRecord & """" & slStr & """" & ","
        mFormatSpotLen tmSdf.iLen, slLenInSec, slLenInMinSec
        slRecord = slRecord & Left(slLenInMinSec, 2)  'save as length MM:SS
        slRecord = slRecord & ":"
        slRecord = slRecord & Mid(slLenInMinSec, 3, 2)

        If ilAutomationType = AUTOTYPE_SCOTT Then                   'terminate the record with a ,
            slRecord = slRecord & ","
        Else
            slRecord = slRecord & ",,,," & """" & "      " & """" & ",,,,,"          'terminate the remaining null fields
        End If
        ilRet = 0
        'smNewLines(UBound(smNewLines)) = Trim(slRecord)
        'ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 118
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = AUTOTYPE_WIDEORBIT Then                           'wide orbit (1-6-12)
        'Time
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = Left(slStr, 8)               'use HH:MM:SS
        slRecord = slRecord & ",,"      'CA,DA"
        If smMcfPrefix = "" Then
            smMcfPrefix = "CA"
        End If
        slRecord = slRecord & smMcfPrefix & ",DA"
        'Cart Number only
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        slRecord = slRecord & slStr & ","

        'Advertiser/Product
        If Trim$(tmChf.sProduct) = "" Then      'omit slash if no product
            slAdvtName = Trim$(tmAdf.sName)
        Else
            slAdvtName = Trim$(tmAdf.sName) & "/" & Trim$(tmChf.sProduct)
        End If
        slRecord = slRecord & """" & slAdvtName & """" & ","

        slStr = Trim$(str$(tmSdf.lCode)) & ":" & Trim$(tmVef.sCodeStn)         'add the vehicle station code with quotes for reference (wide Orbit 1-10-12)
        slRecord = slRecord & """" & slStr & """" & ","
        mFormatSpotLen tmSdf.iLen, slLenInSec, slLenInMinSec
        slRecord = slRecord & Left(slLenInMinSec, 2)  'save as length MM:SS
        slRecord = slRecord & ":"
        slRecord = slRecord & Mid(slLenInMinSec, 3, 2)

        slRecord = slRecord & ","           'terminate the record with a ,
        ilRet = 0
        'smNewLines(UBound(smNewLines)) = Trim(slRecord)
        'ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 118
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = 7 Then    'Prophet MediaStar (Same as Prohet Wizard except 5 blanks in front of time)
        'Time (HH:MM:ss)
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        'Place time into 6 column
        slRecord = "     " & Left(slStr, 8)     'use HH:MM:SS
        slRecord = slRecord & " "   'Leave column 14 blank, media code removed in mFormatCopy
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        slRecord = slRecord & slStr
        'Place ID into 21 column
        Do While Len(slRecord) < 20
            slRecord = slRecord & " "
        Loop
        'Spot ID
        slStr = Trim$(str$(tmSdf.lCode))
        slRecord = slRecord & slStr
        'Place Advteriser/Product into 39 column
        Do While Len(slRecord) < 38
            slRecord = slRecord & " "
        Loop
        'Advertiser/Product
        If Trim$(tmChf.sProduct) <> "" Then
            slAdvtName = Trim$(tmAdf.sName) & "/" & Trim$(tmChf.sProduct)
            If Len(slAdvtName) > 30 Then
                slAdvtName = Left$(slAdvtName, 30)
            End If
        Else
            slAdvtName = Trim$(tmAdf.sName)
        End If
        'advt & prod is max 30 character
        Do While Len(slAdvtName) < 30
            slAdvtName = slAdvtName & " "
        Loop
        'Replace the 30th character with an asterisk so that when return we will know that this spot came from CSI
        slAdvtName = Left$(slAdvtName, 29) & "*"
        slRecord = slRecord & slAdvtName
        'Place Length into 70 column
        Do While Len(slRecord) < 69
            slRecord = slRecord & " "
        Loop
        slStr = Trim$(str$(tmSdf.iLen))
        Do While Len(slStr) < 4
            slStr = "0" & slStr
        Loop
        slRecord = slRecord & slStr
        ilRet = 0
        'smNewLines(UBound(smNewLines)) = slRecord
        'ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 118
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = 8 Then            'iMediaTouch
        'Time
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = Left(slStr, 8)               'use HH:MM:SS
        slRecord = slRecord & " "               'blank between field
        mFormatSpotLen tmSdf.iLen, slLenInSec, slLenInMinSec
        'duration
        slRecord = slRecord & Left(slLenInMinSec, 2)  'save as length MM:SS
        slRecord = slRecord & ":"
        slRecord = slRecord & Mid(slLenInMinSec, 3, 2)
        slRecord = slRecord & " "           'blank between field
        'copy
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        slRecord = slRecord & "ZM" & slStr & " "
        'event index + blank between field (unused(3) blanks + blank between field)
        slRecord = slRecord & "    "
        '11/5/20 - TTP # 10013 - iMediaTouch Replace COM with Media Code
        If (Asc(tgSaf(0).sFeatures7) And IMEDIA_MEDIACODE) = IMEDIA_MEDIACODE Then
            If Trim(Mid(slMcfCode, 1, 3)) = "" Then
                'event category hard-coded "COM" + blank between field (No Copy)
                slRecord = slRecord & "COM "
            Else
                slRecord = slRecord & Mid(slMcfCode, 1, 3) & Space(4 - Len(Mid(slMcfCode, 1, 3)))
            End If
        Else
            'event category hard-coded "COM" + blank between field
            slRecord = slRecord & "COM "
        End If
        'live copy identifier(8 unused) + blank
        slRecord = slRecord & "         "
        'synchronization char(unused) + blank
        slRecord = slRecord & "  "
        'item function (unused & blank)
        slRecord = slRecord & "  "
        'event description (20 char advt name & 10 char internal spot code)
        slAdvtName = Trim$(tmAdf.sAbbr) & "/" & Left$(Trim$(tmChf.sProduct), 15)
        'advt & prod is max 20 character
        Do While Len(slAdvtName) < 20
            slAdvtName = slAdvtName & " "
        Loop

        slRecord = slRecord & Mid$(slAdvtName, 1, 20)
        slStr = Trim$(str$(tmSdf.lCode))
        Do While Len(slStr) < 10
            slStr = " " & slStr
        Loop
        slRecord = slRecord & slStr & " "       'blank after field
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = 9 Then           '8-10-05 Audio Vault Sat
        'Time (HH:MM:ss)
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = "|" & Left(slStr, 8) & "|"              'use HH:MM:SS
        'first spot in break or not?
        If ilSameBreak Then
            slRecord = slRecord & "|"
            ilSameBreak = False
        Else
            slRecord = slRecord & "+|"
        End If
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        slRecord = slRecord & Trim$(slStr) & "|"

        slAdvtName = Trim$(Mid$(tmAdf.sName, 1, 19)) & "/" '& Trim$(Mid$(tmChf.sProduct, 1, 10))
        'take more of the product if the advt didnt fill 19 char
        ilPos = Len(slAdvtName)
        If ilPos < 20 Then
            ilPos = 20 - ilPos + 10         'get the remainder that doesnt fill 19 char and add 10 more for the product
        Else
            ilPos = 10
        End If
        slAdvtName = slAdvtName & Trim$(Mid$(tmChf.sProduct, 1, ilPos))
        'look for vertical bar
        slStr = ""
        For ilLoop = 1 To 30
            If Mid$(slAdvtName, ilLoop, 1) <> "|" Then
                slStr = slStr & (Mid$(slAdvtName, ilLoop, 1))
            End If
        Next ilLoop
        Do While Len(slStr) < 30           'right fill with blanks to fill 30 char
            slStr = slStr & " "
        Loop
        'last 10 characters of the advt/prod is the sdf code
        slTemp = Trim$(str$(tmSdf.lCode))
        Do While Len(slTemp) < 10
            slTemp = " " & slTemp
        Loop

        slRecord = slRecord & slStr & slTemp & "|" & " "        'combine time, copy, advt/Prod, and spot code
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = 14 Then               'audio vault rps
        slStr = gFormatSpotTimeByType(llSpotTime, 12)
        slRecord = "|" & (Trim$(slStr)) & "|+"
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        slRecord = slRecord & "|" & Trim$(slStr) & "||"
        
        'Add advertiser name and internal spot code for reference
        slAdvtName = Trim$(Mid$(tmAdf.sName, 1, 20))
        Do While Len(slAdvtName) < 20       '20 char max
            slAdvtName = slAdvtName & " "
        Loop
        slRecord = slRecord & slAdvtName
        slStr = Trim$(str$(tmSdf.lCode))
        Do While Len(slStr) < 8
            slStr = " " & slStr
        Loop
        slRecord = slRecord & slStr
 
        slSortType = "B"
        'GoSub SaveRecImage

        ilFoundLLC = False
        For ilLLCLoop = ilLLCIndex To UBound(tmLLC) - 1
            llTestTime = gTimeToLong(tmLLC(ilLLCLoop).sStartTime, False)
            If llSpotTime = llTestTime And Val(tmLLC(ilLLCLoop).sType) >= 2 And Val(tmLLC(ilLLCLoop).sType) <= 9 Then
                ilFoundLLC = True
                Exit For
            End If
        Next ilLLCLoop
        If Not ilFoundLLC Then
            ilLLCIndex = 0
        End If
        'Test if split copy exist
        'Test if split copy exist
        ilSub = 1
        If mSplitCopy() Then
            slStr = "|" & gFormatSpotTimeByType(llSpotTime, 12) & "|+|" & Trim$(tmAxf.sAudioVaultID)
            mSaveImageForRPS slStr, slAirDate, llSpotTime, slSortType, tmLLC(ilLLCLoop).sXMid, ilLLCLoop, ilSub
            ilSub = 2
        End If
        'format the key for sorting
        mSaveImageForRPS slRecord, slAirDate, llSpotTime, slSortType, tmLLC(ilLLCLoop).sXMid, ilLLCLoop, ilSub
 
    ElseIf ilAutomationType = 13 Then           '2/2/10:  Rivendell
        'Time
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = Left(slStr, 8)               'use HH:MM:SS
        slRecord = slRecord & " "               'blank between field
        slRecord = slRecord & "R"               'Rivendell character that is ignored
        'copy
        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        Do While Len(slStr) < 14
            slStr = slStr & " "
        Loop
        slRecord = slRecord & slStr
        slRecord = slRecord & " "               'blank between field
        'Creative Title
        Do While Len(smCreativeTitle) < 34
            smCreativeTitle = smCreativeTitle & " "
        Loop
        slRecord = slRecord & smCreativeTitle
        slRecord = slRecord & " "               'blank between field
        'Event Length
        mFormatSpotLen tmSdf.iLen, slLenInSec, slLenInMinSec
        slRecord = slRecord & "00:" & Left(slLenInMinSec, 2) 'save as length MM:SS
        slRecord = slRecord & ":"
        slRecord = slRecord & Mid(slLenInMinSec, 3, 2)
        slRecord = slRecord & " "               'blank between field
        'ISCI
        Do While Len(smISCI) < 32
            smISCI = smISCI & " "
        Loop
        slRecord = slRecord & smISCI
        slRecord = slRecord & " "               'blank between field
        slTemp = Trim$(str$(tmSdf.lCode))
        slTemp = slTemp & ":" & Trim$(str$(tmVef.iCode))
        slTemp = slTemp & ":" & Left$(Trim$(tmVef.sName), 31 - Len(slTemp))
        Do While Len(slTemp) < 32
            slTemp = slTemp & " "
        Loop
        slRecord = slRecord & slTemp
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = AUTOTYPE_JELLI Then           '6-21-12 Jelli
        slRecord = """" & Trim$(smISCI) & """" & ","            'ISCI
        'Date and Time
        slRecord = slRecord & Trim$(slLogYear) & "-" & slLogMonth & "-" & slLogDay & " "
        slStr = gFormatSpotTime(llSpotTime)
        slRecord = slRecord & Trim$(Left(slStr, 8)) & ","               'use HH:MM:SS
        slStr = Trim$(str$(tmSdf.iLen))         'spot length: nnn
        Do While Len(slStr) < 3
            slStr = "0" & slStr
        Loop
        slRecord = slRecord & slStr & ","
        slRecord = slRecord & """" & Trim$(smCreativeTitle) & """" & ","
        slRecord = slRecord & """" & Trim$(tmVef.sCodeStn) & """" & ","                'vehicle station code (5 char max
        slStr = Trim$(str$(tmChf.lCntrNo))         'contract #
        slRecord = slRecord & Trim$(slStr) & ","
        'Advertiser
        slAdvtName = Trim$(tmAdf.sName)
        slAdvtName = gStripIllegalChr(slAdvtName)           'remove special char, retain commas if any embedded in text
        slRecord = slRecord & """" & Trim$(slAdvtName) & """" & ","
        slAgyName = ""
        If tmChf.iAgfCode > 0 Then      'agency name
            ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
            If ilIndex >= 0 Then
                slAgyName = Trim$(tgCommAgf(ilIndex).sName)
                slAgyName = gStripIllegalChr(slAgyName)
            End If
        End If
        slRecord = slRecord & """" & slAgyName & """" & ","
        slRecord = slRecord & Trim$(str(tmSdf.lCode))               'internal spot code
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    ElseIf ilAutomationType = AUTOTYPE_ENCOESPN Then           '6-21-12 Jelli
        slRecord = Format(Trim$(slDate), "m/d/yy") & ","
        slStr = gFormatSpotTime(llSpotTime)
        slRecord = slRecord & Trim$(Left(slStr, 8)) & ","               'use HH:MM:SS
        slStr = Trim$(str$(tmSdf.iLen))         'spot length: nnn
        slRecord = slRecord & slStr & ","
        slRecord = slRecord & """" & Trim$(slEventName) & """" & ","
        slAdvtName = Left(Trim$(tmAdf.sAbbr) & "," & Trim$(tmChf.sProduct), 15)
        slAdvtName = UCase$(slAdvtName)
        smISCI = UCase$(smISCI)
        slRecord = slRecord & """" & gFileNameFilter(slAdvtName & "(" & Trim$(smISCI) & ")") & """" & ","            'ISCI
        slRecord = slRecord & """" & Trim$(tmAdf.sName) & "/" & Trim$(tmChf.sProduct) & """" & ","
        'slRecord = slRecord & Trim$(str(tmSdf.lCode))               'internal spot code
        '
        slHour = Left$(gFormatSpotTime(llAvailTime), 2)
        If llESPNPrevDate <> gDateValue(slDate) Then
            ilESPNHour = -1
        End If
        If Val(slHour) <> ilESPNHour Then
            ilESPNBreak = 1
            ilESPNPosition = 1
        Else
            If llESPNPrevAvailTime <> llAvailTime Then
                ilESPNBreak = ilESPNBreak + 1
                ilESPNPosition = 1
            Else
                ilESPNPosition = ilESPNPosition + 1
            End If
        End If
        llESPNPrevDate = gDateValue(slDate)
        ilESPNHour = Val(slHour)
        llESPNPrevAvailTime = llAvailTime
        slStr = "H" & slHour & "B" & Trim$(str$(ilESPNBreak)) & "P" & Trim$(str$(ilESPNPosition))
        slProgCodeID = ""
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tgVff(ilVff).iVefCode = ilVehCode Then
                slProgCodeID = Trim$(tgVff(ilVff).sXDProgCodeID)
                Exit For
            End If
        Next ilVff
        'If Game and Program Id is Event, get ID from gsf
        slProgCodeID = mGetProgCode(slProgCodeID)
        '9/9/13: If ProgCode = Merge, get code from parent
        slProgCodeID = mGetMergeProgCode(ilVefCode, llESPNPrevDate, llAvailTime, slProgCodeID, ilParentVefCode)
        slRecord = slRecord & slProgCodeID & slStr & ","
        slRecord = slRecord & Trim$(str(tmSdf.lCode))               'internal spot code
        '9/3/13: Temporarily add parent VefCode
        '11/20/13: Add AvailTime (ttp 6507)
        'slRecord = slRecord & "^" & ilParentVefCode
        slRecord = slRecord & "^" & ilParentVefCode & "^" & llESPNPrevAvailTime

        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord

    Else                    '(1)Dalet, (2)Prophet,  (4) Drake (no specs yet), wireready, audio vault, simian
        'Time (HH:MM:ss)
        slStr = gFormatSpotTime(llSpotTime)     'obtain time time as string as HH:MM:SS and extract whats needed
        slRecord = Left(slStr, 8)               'use HH:MM:SS

        slStr = mFormatCopy(ilAutomationType, ilCopy, smCifName, llCopyMissingSdfCode, tmSdf.lCode, slMsg, imMediaCodeLen)
        slRecord = slRecord & slStr

        'Advertiser/Product: 7/5/01 Prophet required creative title to be in with the advt&prod field.
        'reduce advt/prod from 66 to 46 and last 20 character will be replaced with creative title
        slAdvtName = Trim$(tmAdf.sName) & "/" & Left$(Trim$(tmChf.sProduct), 15)
        'advt & prod is max 46 character
        Do While Len(slAdvtName) < 46
            slAdvtName = slAdvtName & " "
        Loop


        'slAdvtName = Left$(slAdvtName, 66)
        'Do While Len(slAdvtName) < 66
        '    slAdvtName = slAdvtName & " "
        'Loop
        slRecord = slRecord & slAdvtName
        'next 20 is the creative title.
        '4-22-13 creative title could be 30char max; this export takes only 20
        If ilAutomationType = 13 Then          'all but rivendell are 20 character creative title fields
            Do While Len(smCreativeTitle) < 34
                smCreativeTitle = smCreativeTitle & " "
            Loop
        Else
            smCreativeTitle = Mid(smCreativeTitle, 1, 20)
            Do While Len(smCreativeTitle) < 20
                smCreativeTitle = smCreativeTitle & " "
            Loop
        End If
        slRecord = slRecord & smCreativeTitle
        'Spot Length , all seconds
        'slStr = Trim$(Str$(tmSdf.iLen))
        'Do While Len(slStr) < 4
        '    slStr = "0" & slStr
        'Loop
        'slRecord = slRecord & slStr
        
        '9-28-18 Station Playlist 1M length needs to be in seconds
        If ilAutomationType = AUTOTYPE_STATIONPL Then
            mFormatSpotLenForMin tmSdf.iLen, slLenInSec, slLenInMinSec
            slRecord = slRecord & Left(slLenInMinSec, 2)  'save as length MM & SS
            slRecord = slRecord & Mid(slLenInMinSec, 3, 2)
            slStr = Trim$(str$(tmSdf.lCode))
            Do While Len(slStr) < 10
                slStr = " " & slStr
            Loop
            slRecord = slRecord & slStr
            'ISCI
            Do While Len(smISCI) < 20
                smISCI = smISCI & " "
            Loop
            slRecord = slRecord & Trim$(smISCI)
        Else
            mFormatSpotLen tmSdf.iLen, slLenInSec, slLenInMinSec
            slRecord = slRecord & Left(slLenInMinSec, 2)  'save as length MM & SS
            slRecord = slRecord & Mid(slLenInMinSec, 3, 2)
            slStr = Trim$(str$(tmSdf.lCode))
            Do While Len(slStr) < 10
                slStr = " " & slStr
            Loop
            slRecord = slRecord & slStr
        End If
        
        '9-28-18 move to above because length is formatted differently
'        If ilAutomationType = AUTOTYPE_STATIONPL Then          '5-11-18
'            'ISCI
'            Do While Len(smISCI) < 20
'                smISCI = smISCI & " "
'            Loop
'            slRecord = slRecord & Trim$(smISCI)
'        End If
        
        '4-7-06 if Prophet nextgen, add 5 more fields
        '5-10-06 swap Fixed time buy and pty fields
        If ilAutomationType = 2 Then
            slRecord = slRecord & " " & Trim$(smROS) & " " & Trim$(smPty) & " " & Trim$(smFixed) & " " & Trim$(smAdvtCode) & " " & Trim$(smCompCode)
        End If

        ilRet = 0
        'smNewLines(UBound(smNewLines)) = slRecord
        'ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 118
        slSortType = "B"
        '6/1/16: Replaced GoSub
        'GoSub SaveRecImage
        mSaveRecImage slAirDate, llSpotTime, slSortDate, slSortType, slRecord
    End If
    llSpotEndTime = llSpotEndTime + Val(tmSdf.iLen)
End Sub

Private Sub tmcWideOrbit00_Timer()

    ckcWideOrbit00.Visible = IIF(rbcAutoType(15).Value = True, True, False)    '8-15-23 jjb TTP 10803 - Added checkbox to allow user to decide to append "00" to file name
    
End Sub
