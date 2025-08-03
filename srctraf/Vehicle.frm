VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Vehicle 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4140
   ClientLeft      =   855
   ClientTop       =   2235
   ClientWidth     =   8175
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
   Icon            =   "Vehicle.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   8175
   Begin VB.TextBox edcACT1Lineup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   1080
      MaxLength       =   11
      TabIndex        =   47
      Top             =   2955
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.TextBox edcAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   3
      Left            =   3015
      MaxLength       =   25
      TabIndex        =   15
      Top             =   2820
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.TextBox edcRNLinkDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   570
      MaxLength       =   30
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2085
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.ListBox lbcRNLink 
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
      ItemData        =   "Vehicle.frx":08CA
      Left            =   2475
      List            =   "Vehicle.frx":08CC
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1770
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.ListBox lbcMultiGameVehLog 
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
      ItemData        =   "Vehicle.frx":08CE
      Left            =   3510
      List            =   "Vehicle.frx":08D0
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.TextBox edcTaxDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   5625
      TabIndex        =   19
      Top             =   1740
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.ListBox lbcTax 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "Vehicle.frx":08D2
      Left            =   450
      List            =   "Vehicle.frx":08D4
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1425
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.TextBox edcHubDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   135
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.ListBox lbcHub 
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
      ItemData        =   "Vehicle.frx":08D6
      Left            =   3120
      List            =   "Vehicle.frx":08D8
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.TextBox edcMultiVehDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   390
      MaxLength       =   40
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.ListBox lbcMultiConvVehLog 
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
      Left            =   3570
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2175
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.TextBox edcContact 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2760
      MaxLength       =   40
      TabIndex        =   11
      Top             =   630
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.PictureBox pbcStartNew 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   8055
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   5
      Top             =   30
      Width           =   105
   End
   Begin VB.TextBox edcVehGp6DropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   195
      MaxLength       =   20
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.ListBox lbcVehGp6 
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
      Left            =   4320
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2145
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.TextBox edcReallDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   525
      MaxLength       =   40
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.TextBox edcVehGp5DropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   255
      MaxLength       =   20
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2535
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.TextBox edcVehGp4DropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   60
      MaxLength       =   20
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.TextBox edcVehGp3DropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   330
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.ListBox lbcVehGp5 
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
      Left            =   3075
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.ListBox lbcVehGp4 
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
      Left            =   3570
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.ListBox lbcVehGp3 
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
      Left            =   4275
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.TextBox edcVehGp2DropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   360
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2565
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.ListBox lbcVehGp2 
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
      Left            =   3915
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.CheckBox ckcLock 
      Alignment       =   1  'Right Justify
      Caption         =   "Locked"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   270
      Width           =   960
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8415
      Top             =   5070
   End
   Begin VB.ListBox lbcType 
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
      Left            =   300
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1785
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lbcBook 
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
      Left            =   3855
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.ListBox lbcDemo 
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
      Left            =   4185
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox edcSort 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   5475
      MaxLength       =   5
      TabIndex        =   42
      Top             =   2580
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox edcDemoDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   585
      MaxLength       =   20
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox edcBookDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   -90
      MaxLength       =   40
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.ListBox lbcLogVeh 
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
      Left            =   3885
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2175
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.TextBox edcTypeDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   195
      MaxLength       =   20
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2550
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.TextBox edcLogVehDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   165
      MaxLength       =   40
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   2610
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
      Left            =   3210
      Picture         =   "Vehicle.frx":08DA
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcStationCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   1065
      MaxLength       =   5
      TabIndex        =   31
      Top             =   2580
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   3075
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   15
      ScaleHeight     =   105
      ScaleWidth      =   90
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3015
      Width           =   90
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6450
      ScaleHeight     =   210
      ScaleWidth      =   1470
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox edcDialPos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   5700
      MaxLength       =   10
      TabIndex        =   18
      Top             =   2550
      Visible         =   0   'False
      Width           =   1470
   End
   Begin MSMask.MaskEdBox mkcFax 
      Height          =   210
      Left            =   4935
      TabIndex        =   17
      Top             =   1020
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16776960
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(AAA) AAA-AAAA"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mkcPhone 
      Height          =   210
      Left            =   990
      TabIndex        =   16
      Top             =   900
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   0
      BackColor       =   16776960
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
   Begin VB.TextBox edcAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   0
      Left            =   720
      MaxLength       =   25
      TabIndex        =   12
      Top             =   1335
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.TextBox edcAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   1
      Left            =   855
      MaxLength       =   25
      TabIndex        =   13
      Top             =   1215
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.TextBox edcAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   9
      Index           =   2
      Left            =   3210
      MaxLength       =   25
      TabIndex        =   14
      Top             =   2565
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   4500
      TabIndex        =   10
      Top             =   2565
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.PictureBox plcSelect 
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   1005
      ScaleHeight     =   360
      ScaleWidth      =   6960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   7020
      Begin VB.ComboBox cbcSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   2265
         TabIndex        =   4
         Top             =   30
         Width           =   4650
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Packages"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1110
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   75
         Width           =   1170
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Vehicles"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   75
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
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
      Left            =   5265
      TabIndex        =   57
      Top             =   3285
      Width           =   975
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
      Left            =   2625
      TabIndex        =   55
      Top             =   3285
      Width           =   975
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
      Left            =   1305
      TabIndex        =   54
      Top             =   3285
      Width           =   975
   End
   Begin VB.PictureBox pbcVeh 
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
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   165
      Picture         =   "Vehicle.frx":09D4
      ScaleHeight     =   2385
      ScaleWidth      =   7860
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   615
      Width           =   7860
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   765
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   135
      Width           =   15
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   45
      TabIndex        =   52
      Top             =   1695
      Width           =   45
   End
   Begin VB.PictureBox plcVeh 
      ForeColor       =   &H00000000&
      Height          =   2430
      Left            =   135
      ScaleHeight     =   2370
      ScaleWidth      =   7875
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Width           =   7935
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   60
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   2085
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   75
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   2445
      Visible         =   0   'False
      Width           =   525
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
      Left            =   7065
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3390
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmcMerge 
      Appearance      =   0  'Flat
      Caption         =   "&Merge"
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
      Left            =   7020
      TabIndex        =   60
      Top             =   3615
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmcOptions 
      Appearance      =   0  'Flat
      Caption         =   "&Options"
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
      Left            =   3945
      TabIndex        =   59
      Top             =   3675
      Width           =   975
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
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
      Left            =   2610
      TabIndex        =   58
      Top             =   3675
      Width           =   975
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      Left            =   3945
      TabIndex        =   56
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label plcScreen 
      Caption         =   "Vehicle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   900
   End
   Begin VB.Label lacCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   6990
      TabIndex        =   65
      Top             =   435
      Width           =   1035
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   3105
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Vehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Vehicle.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Vehicle.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Vehicle input screen code
'
Option Explicit
Option Compare Text
'Vehicle Field Areas
Dim imVefChanged As Integer 'If so, force read of vef and vpf
Dim tmVehicle() As SORTCODE
Dim smVehicleTag As String
Dim tmCtrls(0 To 26)  As FIELDAREA
Dim imLBCtrls As Integer
Dim tmDemoCode() As SORTCODE
Dim smDemoCodeTag As String
Dim tmLogVehCode() As SORTCODE
Dim smLogVehCodeTag As String
Dim tmHubCode() As SORTCODE
Dim smHubCodeTag As String
Dim tmMultiConvVehLogCode() As SORTCODE
Dim smMultiVehLogCodeTag As String
Dim tmMultiGameVehLogCode() As SORTCODE
Dim smMultiGameVehLogCodeTag As String
Dim tmVehGp2Code() As SORTCODE
Dim smVehGp2CodeTag As String
Dim tmVehGp3Code() As SORTCODE
Dim smVehGp3CodeTag As String
Dim tmVehGp4Code() As SORTCODE
Dim smVehGp4CodeTag As String
Dim tmVehGp5Code() As SORTCODE
Dim smVehGp5CodeTag As String
Dim tmVehGp6Code() As SORTCODE
Dim smVehGp6CodeTag As String
Dim tmBookCode() As SORTCODE
Dim smBookNameTag As String
Dim tmTaxSortCode() As SORTCODE
Dim smTaxSortCodeTag As String
Dim tmRNLinkCode() As SORTCODE
Dim smRNLinkCodeTag As String
Dim imMaxCtrlNo As Integer
Dim imState As Integer  '0=Active; 1=Dormant
Dim imBoxNo As Integer   'Current Vehicle Box
Dim tmVef As VEF        'VEF record image
Dim tmMVef() As VEF     'Used to check vehicle names
Dim tmVefSrchKey As INTKEY0    'VEF key record image
Dim imVefRecLen As Integer        'VEF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imLVChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imMVLChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imTaxChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imHubChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imTypeChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imVehGp2ChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imVehGp3ChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imVehGp4ChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imVehGp5ChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imVehGp6ChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBookChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imDemoChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imRNLinkChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imInNew As Integer
Dim imSvSelectedIndex As Integer
Dim imFromOldSave As Integer
Dim imDoubleClickName As Integer
Dim hmVef As Integer 'Vehicle file handle
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim smPhoneImage As String  'Blank phone image- obtained from mkcPhone.text before input
Dim smFaxImage As String    'Blank fax image
Dim imComboBoxIndex As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imLbcArrowSetting As Integer
Dim imLbcMouseDown As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imFirstTimeSelect As Integer
Dim smVehicleType As String
Dim smLogVeh As String
Dim smMultiVehLog As String
Dim smRNLink As String
Dim smOrigTax As String
Dim smOrigHub As String
Dim smOrigVehGp2 As String          'Subtotal (Veh Gp 2)
Dim smVehGp2 As String
Dim smOrigVehGp3 As String          'Market (Veh Gp 3)
Dim smVehGp3 As String
Dim smOrigVehGp4 As String          'Format (Veh Gp 4)
Dim smVehGp4 As String
Dim smOrigVehGp5 As String          'Research (Veh Gp 5)
Dim smVehGp5 As String
Dim smOrigVehGp6 As String          'Sub-Company (Veh Gp 6)
Dim smVehGp6 As String
Dim smBook As String
Dim smReall As String
Dim smDemo As String
Dim smOrigType As String
Dim imOrigState As Integer  '0=Active; 1=Dormant
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imTaxDefined As Integer
Dim bmPrgLibDefined As Boolean
Dim imACT1CodesDefined As Integer
'SDF
Dim tmSdf As SDF
Dim hmSdf As Integer
Dim imSdfRecLen As Integer
'VSF
Dim tmVsf As VSF
Dim hmVsf As Integer
Dim tmVsfSrchKey As LONGKEY0
Dim imVsfRecLen As Integer
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length

Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF

Dim hmVaf As Integer            'Multiname file handle
Dim imVafRecLen As Integer      'MNF record length
Dim tmVafSrchKey1 As INTKEY0
Dim tmVaf As VAF

Dim hmVbf As Integer            'Multiname file handle
Dim imVBfRecLen As Integer      'MNF record length
Dim tmVbfSrchKey1 As VBFKEY1
Dim tmVbf As VBF

Dim hmVff As Integer            'Multiname file handle
Dim imVffRecLen As Integer      'MNF record length
Dim tmVffSrchKey1 As INTKEY0
Dim tmVff As VFF

'Library calendar
Dim hmLcf As Integer        'Library calendar file handle
Dim tmLcf As LCF
Dim imLcfRecLen As Integer
Dim tmLcfSrchKey2 As LCFKEY2

'NIF network inventory added 7-21-05
Dim tmNif As NIF                'NIF record image
Dim tmNifSrchKey1 As NIFKEY1     'NIF key 1 image
Dim imNifRecLen As Integer      'NIF record length
Dim hmNif As Integer            'NIF file handle

Dim imUpdateAllowed As Integer    'User can update records

'6/8/18
Dim smOrigName As String

Const NAMEINDEX = 1     'Name control/field
Const CONTACTINDEX = 2
Const PHONEINDEX = 3    'Phone/extension control/field
Const FAXINDEX = 4      'Fax control/field
Const ADDRESSINDEX = 5  'Address control/field
Const TYPEINDEX = 9     'Type (conventional, selling, airing, Log, Virtual) control/field
Const LOGVEHINDEX = 10   'Log vehicle
Const MULTIVEHLOGINDEX = 11   'Log vehicle
Const RNLINKINDEX = 12
Const DIALPOSINDEX = 13  'Dial position control/field
Const SCODEINDEX = 14   'Station vehicle code control/field
Const MKTNAMEINDEX = 15  'Market Name (Veh Gp 3) control/field
Const RSCHINDEX = 16     'Research (Veh Gp 5)
Const SUBCOMPINDEX = 17     'Sub-Company (Veh Gp 6)
Const FORMATINDEX = 18   'Format (Veh Gp 4) control/index
Const SUBTOTALINDEX = 19 'Subtotal (Veh Gp 2)
Const BOOKINDEX = 20
Const DEMOINDEX = 21
'Const REALLINDEX = 20
Const ACT1CODESINDEX = 22
Const TAXINDEX = 23
Const SORTINDEX = 24
Const HUBINDEX = 25
Const STATEINDEX = 26   'State (Active; Dormant)
'Const SSOURCE1INDEX = 25
'Const PART1INDEX = 26  'Participant (Veh Gp 1)
'Const INTUPDATE1INDEX = 27
'Const EXTUPDATE1INDEX = 28
'Const PRODPCT1INDEX = 29
'Const SSOURCE2INDEX = 30
'Const PART2INDEX = 31  'Participant (Veh Gp 1)
'Const INTUPDATE2INDEX = 32
'Const EXTUPDATE2INDEX = 33
'Const PRODPCT2INDEX = 34
'Const SSOURCE3INDEX = 35
'Const PART3INDEX = 36  'Participant (Veh Gp 1)
'Const INTUPDATE3INDEX = 37
'Const EXTUPDATE3INDEX = 38
'Const PRODPCT3INDEX = 39
'Const SSOURCE4INDEX = 40
'Const PART4INDEX = 41  'Participant (Veh Gp 1)
'Const INTUPDATE4INDEX = 42
'Const EXTUPDATE4INDEX = 43
'Const PRODPCT4INDEX = 44
'Const SSOURCE5INDEX = 45
'Const PART5INDEX = 46  'Participant (Veh Gp 1)
'Const INTUPDATE5INDEX = 47
'Const EXTUPDATE5INDEX = 48
'Const PRODPCT5INDEX = 49
'Const SSOURCE6INDEX = 50
'Const PART6INDEX = 51  'Participant (Veh Gp 1)
'Const INTUPDATE6INDEX = 52
'Const EXTUPDATE6INDEX = 53
'Const PRODPCT6INDEX = 54
'Const SSOURCE7INDEX = 55
'Const PART7INDEX = 56  'Participant (Veh Gp 1)
'Const INTUPDATE7INDEX = 57
'Const EXTUPDATE7INDEX = 58
'Const PRODPCT7INDEX = 59
'Const SSOURCE8INDEX = 60
'Const PART8INDEX = 61  'Participant (Veh Gp 1)
'Const INTUPDATE8INDEX = 62
'Const EXTUPDATE8INDEX = 63
'Const PRODPCT8INDEX = 64
Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    cbcSelect.Text = gFilterNameMatchingKeyPressCheck(cbcSelect.Text, "/[]")
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cbcSelectErr
        End If
        If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            lacCode.Caption = Str$(tmVef.iCode)
        Else
            lacCode.Caption = ""
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
        lacCode.Caption = ""
    End If
    '10071  this may be set too late.  Put next to imChgMode
    igVefCodeModel = 0
    If cbcSelect.ListIndex <= 0 Then
        igVehMode = 0    'New
    Else
        igVehMode = 1   'Change
    End If
    imFirstTimeSelect = True
    pbcVeh.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
        bmPrgLibDefined = mDetermineIfPrgDefined()
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcName.Text = slStr
        End If
    End If
    For ilLoop = imLBCtrls To imMaxCtrlNo Step 1
        If (ilLoop <> MKTNAMEINDEX) And (ilLoop <> RSCHINDEX) And (ilLoop <> FORMATINDEX) And (ilLoop <> SUBTOTALINDEX) And (ilLoop <> SUBCOMPINDEX) Then
            mSetShow ilLoop  'Set show strings
        Else
            Select Case ilLoop
               Case MKTNAMEINDEX 'Vehicle
                    slStr = smVehGp3
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
               Case RSCHINDEX 'Vehicle
                    slStr = smVehGp5
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
               Case SUBCOMPINDEX 'Sub-Company
                    slStr = smVehGp6
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
               Case FORMATINDEX 'Vehicle
                    slStr = smVehGp4
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
               Case SUBTOTALINDEX 'Vehicle
                    slStr = smVehGp2
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
            End Select
        End If
    Next ilLoop
    pbcVeh_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
    imBypassSetting = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_DropDown()
    'mPopulate
    'If imTerminate Then
    '    Exit Sub
    'End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        If igVehCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgVehName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgVehName  'New Name
            End If
            cbcSelect_Change
            If sgVehName <> "" Then
                mSetCommands
                gFindMatch sgVehName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            pbcSTab.SetFocus
            Exit Sub
        End If
    End If
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount = 1 Then
        igVehMode = 0   'New
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        pbcSTab.SetFocus
        Exit Sub
    End If
    gCtrlGotFocus ActiveControl
'    If cbcSelect.ListCount = 2 Then
'        cbcSelect.ListIndex = 1
'        cbcSelect_Change
'        Exit Sub
'    End If
    If slSvText = "[New]" Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        cbcSelect_Change    'Call change so picture area repainted
    ElseIf (slSvText = "") Then
        gFindMatch sgUserDefVehicleName, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
            cbcSelect_Change
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
End Sub
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub ckcLock_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcLock.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    mSetCommands
End Sub
Private Sub cmcCancel_Click()
    Screen.MousePointer = vbHourglass
    If igVehCallSource <> CALLNONE Then
        If igVehCallSource = CALLSOURCEENAME Then
            igVehCallSource = CALLCANCELLED
        End If
        If igVehCallSource = CALLSOURCEPRG Then
            igVehCallSource = CALLCANCELLED
        End If
        mTerminate    'Placed after setfocus no make sure which window gets focus
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If igVehCallSource <> CALLNONE Then
        sgVehName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgVehName = "[New]"
            If Not imTerminate Then
                mEnableBox imBoxNo
                Exit Sub
            Else
                cmcCancel_Click
                Exit Sub
            End If
        End If
    Else
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    igVehicleDefined = True
    If igVehCallSource <> CALLNONE Then
        If igVehCallSource = CALLSOURCEENAME Then
            If sgVehName = "[New]" Then
                igVehCallSource = CALLCANCELLED
            Else
                igVehCallSource = CALLDONE
            End If
        End If
        If igVehCallSource = CALLSOURCEPRG Then
            If sgVehName = "[New]" Then
                igVehCallSource = CALLCANCELLED
            Else
                igVehCallSource = CALLDONE
            End If
        End If
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    If imBoxNo = -1 Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If Not cmcUpdate.Enabled Then
        'Cycle to first unanswered mandatory
        For ilLoop = imLBCtrls To imMaxCtrlNo Step 1
            If mTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
                Screen.MousePointer = vbDefault  'Wait
                Beep
                imBoxNo = ilLoop
                mEnableBox imBoxNo
                Exit Sub
            End If
            Screen.MousePointer = vbDefault  'Wait
        Next ilLoop
    End If
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case TAXINDEX
            lbcTax.Visible = Not lbcTax.Visible
            edcTaxDropdown.SelStart = 0
            edcTaxDropdown.SelLength = Len(edcTaxDropdown.Text)
            edcTaxDropdown.SetFocus
        Case HUBINDEX
            lbcHub.Visible = Not lbcHub.Visible
            edcHubDropdown.SelStart = 0
            edcHubDropdown.SelLength = Len(edcHubDropdown.Text)
            edcHubDropdown.SetFocus
        Case MKTNAMEINDEX
            lbcVehGp3.Visible = Not lbcVehGp3.Visible
            edcVehGp3DropDown.SelStart = 0
            edcVehGp3DropDown.SelLength = Len(edcVehGp3DropDown.Text)
            edcVehGp3DropDown.SetFocus
        Case RSCHINDEX
            lbcVehGp5.Visible = Not lbcVehGp5.Visible
            edcVehGp5DropDown.SelStart = 0
            edcVehGp5DropDown.SelLength = Len(edcVehGp5DropDown.Text)
            edcVehGp5DropDown.SetFocus
        Case SUBCOMPINDEX
            lbcVehGp6.Visible = Not lbcVehGp6.Visible
            edcVehGp6DropDown.SelStart = 0
            edcVehGp6DropDown.SelLength = Len(edcVehGp6DropDown.Text)
            edcVehGp6DropDown.SetFocus
        Case FORMATINDEX
            lbcVehGp4.Visible = Not lbcVehGp4.Visible
            edcVehGp4DropDown.SelStart = 0
            edcVehGp4DropDown.SelLength = Len(edcVehGp4DropDown.Text)
            edcVehGp4DropDown.SetFocus
        Case SUBTOTALINDEX
            lbcVehGp2.Visible = Not lbcVehGp2.Visible
            edcVehGp2DropDown.SelStart = 0
            edcVehGp2DropDown.SelLength = Len(edcVehGp2DropDown.Text)
            edcVehGp2DropDown.SetFocus
        Case BOOKINDEX
            If lbcBook.Visible = False Then
                Screen.MousePointer = vbHourglass
                mBookPop True
                lbcBook.Height = gListBoxHeight(lbcBook.ListCount, 6)
                Screen.MousePointer = vbDefault
            End If
            lbcBook.Visible = Not lbcBook.Visible
            edcBookDropDown.SelStart = 0
            edcBookDropDown.SelLength = Len(edcBookDropDown.Text)
            edcBookDropDown.SetFocus
        'Case REALLINDEX
        '    lbcBook.Visible = Not lbcBook.Visible
        '    edcReallDropDown.SelStart = 0
        '    edcReallDropDown.SelLength = Len(edcReallDropDown.Text)
        '    edcReallDropDown.SetFocus
        Case DEMOINDEX
            lbcDemo.Visible = Not lbcDemo.Visible
            edcDemoDropDown.SelStart = 0
            edcDemoDropDown.SelLength = Len(edcDemoDropDown.Text)
            edcDemoDropDown.SetFocus
        Case TYPEINDEX
            lbcType.Visible = Not lbcType.Visible
            edcTypeDropDown.SelStart = 0
            edcTypeDropDown.SelLength = Len(edcTypeDropDown.Text)
            edcTypeDropDown.SetFocus
        Case LOGVEHINDEX
            lbcLogVeh.Visible = Not lbcLogVeh.Visible
            edcLogVehDropDown.SelStart = 0
            edcLogVehDropDown.SelLength = Len(edcLogVehDropDown.Text)
            edcLogVehDropDown.SetFocus
        Case MULTIVEHLOGINDEX
            If (smVehicleType = "C") Or (smVehicleType = "A") Then
                lbcMultiConvVehLog.Visible = Not lbcMultiConvVehLog.Visible
            ElseIf smVehicleType = "G" Then
                lbcMultiGameVehLog.Visible = Not lbcMultiGameVehLog.Visible
            End If
            edcMultiVehDropdown.SelStart = 0
            edcMultiVehDropdown.SelLength = Len(edcMultiVehDropdown.Text)
            edcMultiVehDropdown.SetFocus
        Case RNLINKINDEX
            lbcRNLink.Visible = Not lbcRNLink.Visible
            edcRNLinkDropdown.SelStart = 0
            edcRNLinkDropdown.SelLength = Len(edcRNLinkDropdown.Text)
            edcRNLinkDropdown.SetFocus
    End Select
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    Dim ilCode As Integer
    Dim slMsg As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If (imSelectedIndex > 0) And (igVehMode = 1) Then
        Screen.MousePointer = vbHourglass
        ilCode = tmVef.iCode
        'Check that record is not referenced-Code missing
        'Can't have ast without Att so only test Att
        'ilRet = gIICodeRefExist(Vehicle, ilCode, "Aet.Mkd", "AetVefCode")    'chfvefCode
        'If ilRet Then
        '    Screen.MousePointer = vbDefault
        '    slMsg = "Cannot erase - a Affiliate History Spot (aet) references this Name"
        '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Alf.Btr", "AlfVefCode")    'chfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Avail Locks references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        'Can't have ast without Att so only test Att
        'ilRet = gIICodeRefExist(Vehicle, ilCode, "Ast.Mkd", "AstVefCode")    'chfvefCode
        'If ilRet Then
        '    Screen.MousePointer = vbDefault
        '    slMsg = "Cannot erase - an Avail Locks references this Name"
        '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Att.Mkd", "AttVefCode")    'chfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Affiliate Agreement references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Auf.Btr", "AufVefCode")    'chfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a User Alert references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Bof.Btr", "BofVefCode")    'chfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Blackout references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Bvf.Btr", "BvfVefCode")    'chfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Budget by Vehicle references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Cct.Mkd", "CctVefCode")    'chfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Affiliate Comment (cct) references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gLLCodeRefExist(Vehicle, CLng(ilCode), "Chf.Btr", "ChfVefCode")    'chfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Clf.Btr", "ClfVefCode")  'clfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Cptt.Mkd", "CpttVefCode")    'chfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Affiliate Post CP (cptt) references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Crf.Btr", "CrfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Rotation references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Cyf.Btr", "CyfVefCode")  'clfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Copy Feed Date references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        'Can't have dat without Att so only test Att
        'ilRet = gIICodeRefExist(Vehicle, ilCode, "Dat.Mkd", "DatVefCode")    'chfvefCode
        'If ilRet Then
        '    Screen.MousePointer = vbDefault
        '    slMsg = "Cannot erase - a Affiliate Pledge (dat) references this Name"
        '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Dlf.Btr", "DlfVefCode")  'clfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Delivery Links references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Drf.Btr", "DrfVefCode")  'clfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Research references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        'Can't have edf without Att so only test Att
        'ilRet = gIICodeRefExist(Vehicle, ilCode, "Edf.Mkd", "EdfVefCode")    'chfvefCode
        'If ilRet Then
        '    Screen.MousePointer = vbDefault
        '    slMsg = "Cannot erase - a Affiliate Export Detail (edf) references this Name"
        '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Egf.Btr", "EgfVefCode")  'clfvefCode
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Engineering Links references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Enf.Btr", "EnfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Event Names references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Fpf.Btr", "FpfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Feed Pledge references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Fsf.Btr", "FsfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Feed Spot references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Fxf.Btr", "FxfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Feed Cross Reference references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Ghf.Btr", "GhfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Game Header references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Gsf.Btr", "GsfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Game Schedule references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Gsf.Btr", "GsfAirVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Game schedule preempt vehicle references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Ihf.Btr", "IhfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Game Inventory Header references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Isf.Btr", "IsfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Game Inventory Specification references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Lcf.Btr", "LcfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Log Date references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Llf.Btr", "LlfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Live Log references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Lst.Mkd", "LstLnVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Affiliate Log Spot references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Lst.Mkd", "LstLogVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Affiliate Log Spot references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Ltf.Btr", "LtfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Library Title references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Mat.Mkd", "MatVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Market Group (mat) references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "mcf.Btr", "McfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Media references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "msf.Btr", "MsfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Media Inventory Sold references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "mtf.Btr", "MtfSdfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a MG Tracking references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Paf.Btr", "PafVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Program Airing Information references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Pff.Btr", "PffVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Pre-Feed references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Phf.Btr", "PhfAirVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Revenue History references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Phf.Btr", "PhfBillVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Revenue History references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Pjf.Btr", "PjfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Projection references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
'        ilRet = gIICodeRefExist(Vehicle, ilCode, "Lnf.Btr", "LnfVefCode")
'        If ilRet Then
'            slMsg = "Cannot erase - a Log Library references this Name"
'            ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
'            Exit Sub
'        End If
        'ilRet = gIICodeRefExist(Vehicle, ilCode, "Fsf.Btr", "FsfVefCode")
        'If ilRet Then
        '    slMsg = "Cannot erase - a Feed Spot references this Name"
        '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        'ilRet = gIICodeRefExist(Vehicle, ilCode, "Fxf.Btr", "FxfVefCode")
        'If ilRet Then
        '    slMsg = "Cannot erase - a Feed X-Ref references this Name"
        '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Rcf.Btr", "RcfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Rate card references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Rht.Mkd", "RhtVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a RADAR Header references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        'ilRet = gIICodeRefExist(Vehicle, ilCode, "Rpf.Btr", "RpfVefCode")
        'If ilRet Then
        '    slMsg = "Cannot erase - a Rate Card Program/Time references this Name"
        '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '    Exit Sub
        'End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Rif.Btr", "RifVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Rate card Item references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Rsf.Btr", "RsfBVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Regional Assigned Spots references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Rvf.Btr", "RvfAirVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Receivables references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Rvf.Btr", "RvfBillVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Receivables references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Saf.Btr", "SafVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Schedule Attributes references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Sbf.Btr", "SbfAirVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Special Billing references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Sbf.Btr", "SbfBillVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Special Billing references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Scf.Btr", "ScfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Salesperson Commission references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Sdf.Btr", "SdfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Spot Detail references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Ssf.Btr", "SsfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Spot Summary references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Stf.Btr", "StfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Spot Tracking references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Psf.Btr", "PsfVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Package Spot references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Vef.Btr", "VefVefCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Package Spot references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Vlf.Btr", "VlfSellCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Links references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = gIICodeRefExist(Vehicle, ilCode, "Vlf.Btr", "VlfAirCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Links references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        If mCheckPvfVefCode(ilCode) Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Package Vehicle reference this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        
'        ilRet = gIICodeRefExist(Vehicle, ilCode, "Urf.Btr", "UrfVefCode")
        ilRet = gCodeInUser(Vehicle, "V", ilCode)
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a User Options references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
'        ilRet = gIICodeRefExist(Vehicle, ilCode, "Urf.Btr", "UrfDefaultVeh")
'        If ilRet Then
'            slMsg = "Cannot erase - a User Options references this Name"
'            ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
'            Exit Sub
'        End If
        ilRet = gIFSCodeRefExist(Vehicle, "F", ilCode)
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Vehicle Combo references this Name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmVef.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        gGetSyncDateTime slSyncDate, slSyncTime
        slStamp = gFileDateTime(sgDBPath & "Vef.btr")
        If tmVef.sType <> "V" Then
            If Not mDeleteVpf(tmVef.iCode) Then
                cmcCancel_Click
                Exit Sub
            End If
            If Not mDeleteVof(tmVef.iCode, tmVef.sType) Then
                cmcCancel_Click
                Exit Sub
            End If
        Else
            ilRet = mDeleteVsf(tmVef.lVsfCode)
            On Error GoTo cmcEraseErr
            gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", Vehicle
            On Error GoTo 0
            If Not mDeleteVpf(tmVef.iCode) Then
                cmcCancel_Click
                Exit Sub
            End If
            If Not mDeleteVof(tmVef.iCode, tmVef.sType) Then
                cmcCancel_Click
                Exit Sub
            End If
        End If
        ilRet = mDeleteVff(tmVef.iCode)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete: Vff)", Vehicle
        On Error GoTo 0
        ilRet = mDeleteVaf(tmVef.iCode)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete: Vaf)", Vehicle
        On Error GoTo 0
        ilRet = mDeleteVbf(tmVef.iCode)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete: Vbf)", Vehicle
        On Error GoTo 0
        ilRet = mDeleteNif(tmVef.iCode)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete: Nif)", Vehicle
        On Error GoTo 0
        ilRet = btrDelete(hmVef)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", Vehicle
        On Error GoTo 0
'        If tgSpf.sRemoteUsers = "Y" Then
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "VEF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmVef.iRemoteID
'            tmDsf.lAutoCode = tmVef.iAutoCode
'            tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'            tmDsf.lCntrNo = 0
'            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'        End If
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If Traffic!lbcVehicle.Tag <> "" Then
        '    If slStamp = Traffic!lbcVehicle.Tag Then
        '        Traffic!lbcVehicle.Tag = FileDateTime(sgDBPath & "Vef.btr")
        '    End If
        'End If
'        If smVehicleTag <> "" Then
'            If slStamp = smVehicleTag Then
'                smVehicleTag = gFileDateTime(sgDBPath & "Vef.btr")
'            End If
'        End If
'        'Traffic!lbcVehicle.RemoveItem imSelectedIndex - 1
'        gRemoveItemFromSortCode imSelectedIndex - 1, tmVehicle()
'        cbcSelect.RemoveItem imSelectedIndex
'        On Error GoTo cmcEraseErr
'        ilRet = gVpfRead(Vehicle)  'Reset Vehicle preference table
        On Error GoTo 0
        sgVehicleTag = ""
        sgMVefStamp = ""
        ilRet = csiSetStamp("VEF", sgMVefStamp)
        DoEvents
        smVehicleTag = ""
        sgMVefStamp = "~"   'Force Read
        mPopulate
        gObtainVefIgnoreHub tmMVef()
        sgVpfStamp = "~"    'Force read
        ilRet = gVpfRead()
        gAnyClustersDef
        gAnyRepDef
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcVeh.Cls
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcEraseErr:
    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub
Private Sub cmcErase_GotFocus()
    gCtrlGotFocus cmcErase
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub cmcMerge_Click()
    Dim slMsg As String
    Dim ilRet As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    'If tgSpf.sRemoteUsers = "Y" Then
    '    slMsg = "Cannot Merge - Remote User System in Use"
    If tgUrf(0).iRemoteID > 0 Then
        slMsg = "Remote User Cannot Run Merge"
        ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Merge")
        Exit Sub
    End If
    ilRet = MsgBox("Backup of database must be done before merge, has it been done", vbYesNo + vbQuestion, "Merge Vehicle")
    If ilRet = vbNo Then
        Exit Sub
    End If
    ilRet = MsgBox("Are all other users off the traffic system", vbYesNo + vbQuestion, "Merge Vehicle")
    If ilRet = vbNo Then
        Exit Sub
    End If
    igMergeCallSource = VEHICLESLIST
    Merge.Show vbModal
    Screen.MousePointer = vbHourglass
    pbcVeh.Cls
    cbcSelect.Clear
    mPopulate
    gObtainVefIgnoreHub tmMVef()
    cbcSelect.ListIndex = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcMerge_GotFocus()
    gCtrlGotFocus cmcMerge
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub cmcOptions_Click()
    'If Not gWinRoom(igNoExeWinRes(VEHOPTEXE)) Then
    '    Exit Sub
    'End If
    'If tmVef.sType <> "V" Then
    'If (tmVef.sType <> "P") And (tmVef.sType <> "T") Then
    'If (tmVef.sType <> "T") And (tmVef.sType <> "R") And (tmVef.sType <> "U") Then
    'If (tmVef.sType <> "T") And (tmVef.sType <> "N") Then
        igVehOptCallSource = CALLSOURCEVEHICLE
        sgVehNameToVehOpt = cbcSelect.Text
        igVehNewToVehOpt = False
        'Screen.MousePointer = vbHourGlass  'Wait
        VehOpt.Show vbModal
        'Screen.MousePointer = vbDefault    'Default
        igVehOptCallSource = CALLNONE
    'Else
    '    'Screen.MousePointer = vbHourGlass  'Wait
    '    sgVsfName = Trim$(tmVef.sName)
    '    VirtVeh.Show vbModal
    '    'Screen.MousePointer = vbDefault    'Default
    'End If
    'End If
End Sub
Private Sub cmcOptions_GotFocus()
    gCtrlGotFocus cmcOptions
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcReport_Click()
    '10071 vehicle in its own project.  Lose reports
'    Dim slStr As String
'    'Unload IconTraf
'    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
'    '    Exit Sub
'    'End If
'    igRptCallType = VEHICLESLIST
'    igRptType = 0
'    ''Screen.MousePointer = vbHourGlass  'Wait
'    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
'    'edcLinkSrceDoneMsg.Text = ""
'    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
'        If igTestSystem Then
'            slStr = "Vehicle^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
'        Else
'            slStr = "Vehicle^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
'        End If
'    'Else
'    '    If igTestSystem Then
'    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
'    '    Else
'    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
'    '    End If
'    'End If
'    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
'    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
'    'Vehicle.Enabled = False
'    'Do While Not igChildDone
'    '    DoEvents
'    'Loop
'    'slStr = sgDoneMsg
'    'Vehicle.Enabled = True
'    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
'    'For ilLoop = 0 To 10
'    '    DoEvents
'    'Next ilLoop
'    ''Screen.MousePointer = vbDefault    'Default
'    sgCommandStr = slStr
'    RptList.Show vbModal
End Sub
Private Sub cmcReport_GotFocus()
    gCtrlGotFocus cmcReport
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    Dim slStr As String
    If imSelectedIndex > 0 Then
        ilIndex = imSelectedIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cmcUndoErr
        End If
        pbcVeh.Cls
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To imMaxCtrlNo Step 1
            If (ilLoop <> MKTNAMEINDEX) And (ilLoop <> RSCHINDEX) And (ilLoop <> FORMATINDEX) And (ilLoop <> SUBTOTALINDEX) And (ilLoop <> SUBCOMPINDEX) Then
                mSetShow ilLoop  'Set show strings
            Else
                Select Case ilLoop
                   Case MKTNAMEINDEX 'Vehicle
                        slStr = smVehGp3
                        gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
                   Case RSCHINDEX 'Vehicle
                        slStr = smVehGp5
                        gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
                   Case SUBCOMPINDEX 'Sub-Company
                        slStr = smVehGp6
                        gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
                   Case FORMATINDEX 'Vehicle
                        slStr = smVehGp4
                        gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
                   Case SUBTOTALINDEX 'Vehicle
                        slStr = smVehGp2
                        gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
                End Select
            End If
        Next ilLoop
        pbcVeh_Paint
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcVeh.Cls
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    gCtrlGotFocus cmcUndo
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = Trim$(edcName.Text)   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    imBoxNo = -1
    'Must reset display so altered flag is cleared and setcommand will turn select on
    If imSvSelectedIndex <> 0 Then
        gFindMatch slName, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        Else
            cbcSelect.ListIndex = 0
        End If
        imFromOldSave = True
    Else
        cbcSelect.ListIndex = 0
        imFromOldSave = False
    End If
    DoEvents
    'If imSvSelectedIndex = cbcSelect.ListIndex Then
    '    cbcSelect_Change    'Call change so picture area repainted
    'End If
    imFirstTimeSelect = True
    igVehicleDefined = True
    igVefCodeModel = 0
    mSetCommands
    If cbcSelect.Enabled Then
        On Error Resume Next
        cbcSelect.SetFocus
    Else
        pbcClickFocus.SetFocus
    End If
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus cmcUpdate
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub edcACT1Lineup_Change()
    mSetChg imBoxNo
End Sub

Private Sub edcACT1Lineup_GotFocus()
    gCtrlGotFocus edcACT1Lineup
End Sub

Private Sub edcAddr_Change(Index As Integer)
    mSetChg imBoxNo
End Sub
Private Sub edcAddr_GotFocus(Index As Integer)
    gCtrlGotFocus edcAddr(Index)
End Sub
Private Sub edcAddr_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcBookDropDown_Change()
    imLbcArrowSetting = True
    gMatchLookAhead edcBookDropDown, lbcBook, imBSMode, imComboBoxIndex
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcBookDropDown_GotFocus()
    If lbcBook.ListCount = 1 Then
        lbcBook.ListIndex = 0
        'If imTabDirection = -1 Then  'Right To Left
        '    pbcSTab.SetFocus
        'Else
        '    pbcTab.SetFocus
        'End If
        'Exit Sub
    End If
End Sub
Private Sub edcBookDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcBookDropDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcBookDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub edcBookDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcBook, imLbcArrowSetting
        edcBookDropDown.SelStart = 0
        edcBookDropDown.SelLength = Len(edcBookDropDown.Text)
    End If
End Sub

Private Sub edcContact_Change()
    mSetChg CONTACTINDEX 'Use NAMEINDEX instead of imBoxNo to handle calling from another function
End Sub

Private Sub edcContact_GotFocus()
    gCtrlGotFocus edcContact
End Sub

Private Sub edcContact_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcDemoDropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcDemoDropDown, lbcDemo, imBSMode, slStr)
    If ilRet = 1 Then
        lbcDemo.ListIndex = 1
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcDemoDropDown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub
Private Sub edcDemoDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDemoDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDemoDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDemoDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcDemo, imLbcArrowSetting
        edcDemoDropDown.SelStart = 0
        edcDemoDropDown.SelLength = Len(edcDemoDropDown.Text)
    End If
End Sub
Private Sub edcDemoDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub
Private Sub edcDialPos_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcDialPos_GotFocus()
    gCtrlGotFocus edcDialPos
End Sub
Private Sub edcDialPos_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcHubDropdown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    
    imLbcArrowSetting = True
    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
        gMatchLookAhead edcHubDropdown, lbcHub, imBSMode, imComboBoxIndex
    Else
        ilRet = gOptionalLookAhead(edcRNLinkDropdown, lbcRNLink, imBSMode, slStr)
        If ilRet = 1 Then
            lbcRNLink.ListIndex = 1
        End If
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub

Private Sub edcHubDropdown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub

Private Sub edcHubDropdown_GotFocus()
    If lbcHub.ListCount <= 2 Then
        '12/27/18: Select [None]
        'lbcHub.ListIndex = 1    '0
    End If
End Sub

Private Sub edcHubDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcHubDropdown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcHubDropdown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcHubDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcHub, imLbcArrowSetting
        edcHubDropdown.SelStart = 0
        edcHubDropdown.SelLength = Len(edcHubDropdown.Text)
    End If
End Sub

Private Sub edcHubDropdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcLogVehDropDown_Change()
    imLbcArrowSetting = True
    gMatchLookAhead edcLogVehDropDown, lbcLogVeh, imBSMode, imComboBoxIndex
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcLogVehDropDown_GotFocus()
    If lbcLogVeh.ListCount = 1 Then
        lbcLogVeh.ListIndex = 0
        'If imTabDirection = -1 Then  'Right To Left
        '    pbcSTab.SetFocus
        'Else
        '    pbcTab.SetFocus
        'End If
        'Exit Sub
    End If
End Sub
Private Sub edcLogVehDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcLogVehDropDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcLogVehDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub edcLogVehDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcLogVeh, imLbcArrowSetting
        edcLogVehDropDown.SelStart = 0
        edcLogVehDropDown.SelLength = Len(edcLogVehDropDown.Text)
    End If
End Sub

Private Sub edcMultiVehDropdown_Change()
    imLbcArrowSetting = True
    If (smVehicleType = "C") Or (smVehicleType = "A") Then
        gMatchLookAhead edcMultiVehDropdown, lbcMultiConvVehLog, imBSMode, imComboBoxIndex
    ElseIf smVehicleType = "G" Then
        gMatchLookAhead edcMultiVehDropdown, lbcMultiGameVehLog, imBSMode, imComboBoxIndex
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub

Private Sub edcMultiVehDropdown_GotFocus()
    If (smVehicleType = "C") Or (smVehicleType = "A") Then
        If lbcMultiConvVehLog.ListCount = 1 Then
            lbcMultiConvVehLog.ListIndex = 0
            'If imTabDirection = -1 Then  'Right To Left
            '    pbcSTab.SetFocus
            'Else
            '    pbcTab.SetFocus
            'End If
            'Exit Sub
        End If
    ElseIf smVehicleType = "G" Then
        If lbcMultiGameVehLog.ListCount = 1 Then
            lbcMultiGameVehLog.ListIndex = 0
        End If
    End If
End Sub

Private Sub edcMultiVehDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcMultiVehDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcMultiVehDropdown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub edcMultiVehDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        If (smVehicleType = "C") Or (smVehicleType = "A") Then
            gProcessArrowKey Shift, KeyCode, lbcMultiConvVehLog, imLbcArrowSetting
        ElseIf smVehicleType = "G" Then
            gProcessArrowKey Shift, KeyCode, lbcMultiGameVehLog, imLbcArrowSetting
        End If
        edcMultiVehDropdown.SelStart = 0
        edcMultiVehDropdown.SelLength = Len(edcMultiVehDropdown.Text)
    End If
End Sub

Private Sub edcName_Change()
    mSetChg NAMEINDEX 'Use NAMEINDEX instead of imBoxNo to handle calling from another function
End Sub
Private Sub edcName_GotFocus()
    gCtrlGotFocus edcName
End Sub
Private Sub edcName_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAsciiIncludeSlash(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcName_LostFocus()
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKName()
    '9760
    edcName.Text = gRemoveIllegalPastedChar(edcName.Text)
    edcName.Text = gFilterNameMatchingKeyPressCheck(edcName.Text, "/[]")
End Sub

Private Sub edcReallDropDown_Change()
    imLbcArrowSetting = True
    gMatchLookAhead edcReallDropDown, lbcBook, imBSMode, imComboBoxIndex
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcReallDropDown_GotFocus()
    If lbcBook.ListCount = 1 Then
        lbcBook.ListIndex = 0
        'If imTabDirection = -1 Then  'Right To Left
        '    pbcSTab.SetFocus
        'Else
        '    pbcTab.SetFocus
        'End If
        'Exit Sub
    End If
End Sub
Private Sub edcReallDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcReallDropDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcReallDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub edcReallDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcBook, imLbcArrowSetting
        edcReallDropDown.SelStart = 0
        edcReallDropDown.SelLength = Len(edcReallDropDown.Text)
    End If
End Sub

Private Sub edcRNLinkDropdown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcRNLinkDropdown, lbcRNLink, imBSMode, slStr)
    If ilRet = 1 Then
        lbcRNLink.ListIndex = 1
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub

Private Sub edcRNLinkDropdown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub

Private Sub edcRNLinkDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcRNLinkDropdown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcRNLinkDropdown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcRNLinkDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcRNLink, imLbcArrowSetting
        edcRNLinkDropdown.SelStart = 0
        edcRNLinkDropdown.SelLength = Len(edcRNLinkDropdown.Text)
    End If
End Sub

Private Sub edcRNLinkDropdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub

Private Sub edcSort_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcSort_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcSort_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcSort.Text
    slStr = Left$(slStr, edcSort.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSort.SelStart - edcSort.SelLength)
    If gCompNumberStr(slStr, "10000") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcStationCode_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcStationCode_GotFocus()
    gCtrlGotFocus edcStationCode
End Sub
Private Sub edcStationCode_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcTaxDropdown_Change()
    imLbcArrowSetting = True
    gMatchLookAhead edcTaxDropdown, lbcTax, imBSMode, imComboBoxIndex
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub

Private Sub edcTaxDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcTaxDropdown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTaxDropdown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcTaxDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcTax, imLbcArrowSetting
        edcTaxDropdown.SelStart = 0
        edcTaxDropdown.SelLength = Len(edcTaxDropdown.Text)
    End If
End Sub

Private Sub edcTypeDropDown_Change()
    imLbcArrowSetting = True
    gMatchLookAhead edcTypeDropDown, lbcType, imBSMode, imComboBoxIndex
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcTypeDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTypeDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTypeDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcTypeDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcType, imLbcArrowSetting
        edcTypeDropDown.SelStart = 0
        edcTypeDropDown.SelLength = Len(edcTypeDropDown.Text)
    End If
End Sub
Private Sub edcVehGp2DropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcVehGp2DropDown, lbcVehGp2, imBSMode, slStr)
    If ilRet = 1 Then
        lbcVehGp2.ListIndex = 0
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcVehGp2DropDown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub
Private Sub edcVehGp2DropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcVehGp2DropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcVehGp2DropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcVehGp2DropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcVehGp2, imLbcArrowSetting
        edcVehGp2DropDown.SelStart = 0
        edcVehGp2DropDown.SelLength = Len(edcVehGp2DropDown.Text)
    End If
End Sub
Private Sub edcVehGp2DropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub
Private Sub edcVehGp3DropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcVehGp3DropDown, lbcVehGp3, imBSMode, slStr)
    If ilRet = 1 Then
        lbcVehGp3.ListIndex = 0
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcVehGp3DropDown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub
Private Sub edcVehGp3DropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcVehGp3DropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcVehGp3DropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcVehGp3DropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcVehGp3, imLbcArrowSetting
        edcVehGp3DropDown.SelStart = 0
        edcVehGp3DropDown.SelLength = Len(edcVehGp3DropDown.Text)
    End If
End Sub
Private Sub edcVehGp3DropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub
Private Sub edcVehGp4DropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcVehGp4DropDown, lbcVehGp4, imBSMode, slStr)
    If ilRet = 1 Then
        lbcVehGp4.ListIndex = 0
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcVehGp4DropDown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub
Private Sub edcVehGp4DropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcVehGp4DropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcVehGp4DropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcVehGp4DropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcVehGp4, imLbcArrowSetting
        edcVehGp4DropDown.SelStart = 0
        edcVehGp4DropDown.SelLength = Len(edcVehGp4DropDown.Text)
    End If
End Sub
Private Sub edcVehGp4DropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub
Private Sub edcVehGp5DropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcVehGp5DropDown, lbcVehGp5, imBSMode, slStr)
    If ilRet = 1 Then
        lbcVehGp5.ListIndex = 0
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcVehGp5DropDown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub
Private Sub edcVehGp5DropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcVehGp5DropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcVehGp5DropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcVehGp5DropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcVehGp5, imLbcArrowSetting
        edcVehGp5DropDown.SelStart = 0
        edcVehGp5DropDown.SelLength = Len(edcVehGp5DropDown.Text)
    End If
End Sub
Private Sub edcVehGp5DropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub
Private Sub edcVehGp6DropDown_Change()
    Dim ilRet As Integer
    Dim slStr As String
    imLbcArrowSetting = True
    ilRet = gOptionalLookAhead(edcVehGp6DropDown, lbcVehGp6, imBSMode, slStr)
    If ilRet = 1 Then
        lbcVehGp6.ListIndex = 0
    End If
    imLbcArrowSetting = False
    mSetChg imBoxNo
End Sub
Private Sub edcVehGp6DropDown_DblClick()
    imDoubleClickName = True    'Double click event followed by mouse up
End Sub
Private Sub edcVehGp6DropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcVehGp6DropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcVehGp6DropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcVehGp6DropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcVehGp6, imLbcArrowSetting
        edcVehGp6DropDown.SelStart = 0
        edcVehGp6DropDown.SelLength = Len(edcVehGp6DropDown.Text)
    End If
End Sub
Private Sub edcVehGp6DropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
        Exit Sub
        imDoubleClickName = False
    End If
End Sub
Private Sub Form_Activate()
    Me.KeyPreview = True
    If imInNew Then
        '10071
        Vehicle.Refresh
        Exit Sub
    End If
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        '10071-trying to get vehicle to show when screen loses focus
        rbcType(0).Refresh
        rbcType(1).Refresh
        cbcSelect.Refresh
        Exit Sub
    End If
    imFirstActivate = False
    '10071 move to later
    If (igWinStatus(VEHICLESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcVeh.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcVeh.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Vehicle.Refresh
    mSetCommands
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcSelect.Enabled) And (imBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
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
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    If Not igManUnload Then
        mSetShow imBoxNo
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            mEnableBox imBoxNo
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    
    gBuildDormantVef
    
    Erase tmVehicle
    Erase tmDemoCode
    Erase tmLogVehCode
    Erase tmHubCode
    Erase tmMultiConvVehLogCode
    Erase tmMultiGameVehLogCode
    Erase tmVehGp2Code
    Erase tmVehGp3Code
    Erase tmVehGp4Code
    Erase tmVehGp5Code
    Erase tmVehGp6Code
    Erase tmBookCode
    Erase tmMVef
    Erase tmTaxSortCode

    btrExtClear hmNif   'Clear any previous extend operation
    ilRet = btrClose(hmNif)
    btrDestroy hmNif
    btrExtClear hmLcf   'Clear any previous extend operation
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    btrExtClear hmVaf   'Clear any previous extend operation
    ilRet = btrClose(hmVaf)
    btrDestroy hmVaf
    btrExtClear hmVbf   'Clear any previous extend operation
    ilRet = btrClose(hmVbf)
    btrDestroy hmVbf
    btrExtClear hmVff   'Clear any previous extend operation
    ilRet = btrClose(hmVff)
    btrDestroy hmVff
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    btrExtClear hmVsf   'Clear any previous extend operation
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    btrExtClear hmSdf   'Clear any previous extend operation
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    
    Set Vehicle = Nothing   'Remove data segment
    End
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcBook_Click()
    'If imBoxNo = REALLINDEX Then
    '    gProcessLbcClick lbcBook, edcReallDropDown, imBookChgMode, imLbcArrowSetting
    'Else
        gProcessLbcClick lbcBook, edcBookDropDown, imBookChgMode, imLbcArrowSetting
    'End If
End Sub
Private Sub lbcDemo_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcDemo, edcDemoDropDown, imDemoChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcDemo_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcDemo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcDemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcDemo, edcDemoDropDown, imDemoChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcHub_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcHub, edcHubDropdown, imHubChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcHub_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcHub_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcHub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcHub_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcHub, edcHubDropdown, imHubChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcLogVeh_Click()
    gProcessLbcClick lbcLogVeh, edcLogVehDropDown, imLVChgMode, imLbcArrowSetting
End Sub

Private Sub lbcMultiConvVehLog_Click()
    If (smVehicleType = "C") Or (smVehicleType = "A") Then
        gProcessLbcClick lbcMultiConvVehLog, edcMultiVehDropdown, imMVLChgMode, imLbcArrowSetting
    ElseIf smVehicleType = "G" Then
        gProcessLbcClick lbcMultiGameVehLog, edcMultiVehDropdown, imMVLChgMode, imLbcArrowSetting
    End If
End Sub


Private Sub lbcMultiGameVehLog_Click()
    If (smVehicleType = "C") Or (smVehicleType = "A") Then
        gProcessLbcClick lbcMultiConvVehLog, edcMultiVehDropdown, imMVLChgMode, imLbcArrowSetting
    ElseIf smVehicleType = "G" Then
        gProcessLbcClick lbcMultiGameVehLog, edcMultiVehDropdown, imMVLChgMode, imLbcArrowSetting
    End If

End Sub

Private Sub lbcRNLink_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcRNLink, edcRNLinkDropdown, imRNLinkChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcRNLink_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcRNLink_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcRNLink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcRNLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcRNLink, edcRNLinkDropdown, imRNLinkChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcTax_Click()
    gProcessLbcClick lbcTax, edcTaxDropdown, imTaxChgMode, imLbcArrowSetting
End Sub

Private Sub lbcType_Click()
    gProcessLbcClick lbcType, edcTypeDropDown, imTypeChgMode, imLbcArrowSetting
End Sub

Private Sub lbcVehGp2_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcVehGp2, edcVehGp2DropDown, imVehGp2ChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcVehGp2_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcVehGp2_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehGp2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcVehGp2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehGp2, edcVehGp2DropDown, imVehGp2ChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcVehGp3_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcVehGp3, edcVehGp3DropDown, imVehGp3ChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcVehGp3_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcVehGp3_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehGp3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcVehGp3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehGp3, edcVehGp3DropDown, imVehGp3ChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcVehGp4_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcVehGp4, edcVehGp4DropDown, imVehGp4ChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcVehGp4_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcVehGp4_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehGp4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcVehGp4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehGp4, edcVehGp4DropDown, imVehGp4ChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcVehGp5_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcVehGp5, edcVehGp5DropDown, imVehGp5ChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcVehGp5_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcVehGp5_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehGp5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcVehGp5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehGp5, edcVehGp5DropDown, imVehGp5ChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcVehGp6_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcVehGp6, edcVehGp6DropDown, imVehGp6ChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcVehGp6_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcVehGp6_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcVehGp6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcVehGp6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcVehGp6, edcVehGp6DropDown, imVehGp6ChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddVofModel                    *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add Vof                        *
'*                                                     *
'*******************************************************
Private Function mAddVofModel() As Integer
    Dim hlVof As Integer        'site Option file handle
    Dim ilRecLen As Integer     'Vof record length
    Dim tlVof As VOF
    Dim tlSrchKey As VOFKEY0
    Dim hlCef As Integer    'Comment file handle
    Dim tlCef As CEF        'CEF record image
    Dim tlCefSrchKey As LONGKEY0    'CEF key record image
    Dim ilCefRecLen As Integer        'CEF record length
    Dim ilRet As Integer
    Dim slType As String
    Dim ilPass As Integer
    hlVof = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlVof, "", sgDBPath & "Vof.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mAddVofModel = False
        Exit Function
    End If
    ilRecLen = Len(tlVof)  'btrRecordLength(hlVof)  'Get and save record length
    hlCef = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlCef, "", sgDBPath & "Cef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlVof)
        btrDestroy hlVof
        mAddVofModel = False
        Exit Function
    End If
    ilCefRecLen = Len(tlCef)  'btrRecordLength(hlVof)  'Get and save record length
    For ilPass = 0 To 2 Step 1
        If ilPass = 0 Then
            slType = "L"
        ElseIf ilPass = 1 Then
            slType = "C"
        Else
            slType = "O"
        End If
        tlSrchKey.iVefCode = igVefCodeModel
        tlSrchKey.sType = slType
        ilRet = btrGetEqual(hlVof, tlVof, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tlCefSrchKey.lCode = tlVof.lHd1CefCode
            If tlCefSrchKey.lCode <> 0 Then
                tlCef.sComment = ""
                ilCefRecLen = Len(tlCef)    '1009
                ilRet = btrGetEqual(hlCef, tlCef, ilCefRecLen, tlCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    ilCefRecLen = Len(tlCef)    '5 + Len(Trim$(tlCef.sComment)) + 2
                    tlCef.lCode = 0
                    ilRet = btrInsert(hlCef, tlCef, ilCefRecLen, INDEXKEY0)
                    If ilRet = BTRV_ERR_NONE Then
                        tlVof.lHd1CefCode = tlCef.lCode
                    Else
                        tlVof.lHd1CefCode = 0
                    End If
                Else
                    tlVof.lHd1CefCode = 0
                End If
            Else
                tlVof.lHd1CefCode = 0
            End If
            tlCefSrchKey.lCode = tlVof.lFt1CefCode
            If tlCefSrchKey.lCode <> 0 Then
                tlCef.sComment = ""
                ilCefRecLen = Len(tlCef)    '1009
                ilRet = btrGetEqual(hlCef, tlCef, ilCefRecLen, tlCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    ilCefRecLen = Len(tlCef)   '5 + Len(Trim$(tlCef.sComment)) + 2
                    tlCef.lCode = 0
                    ilRet = btrInsert(hlCef, tlCef, ilCefRecLen, INDEXKEY0)
                    If ilRet = BTRV_ERR_NONE Then
                        tlVof.lFt1CefCode = tlCef.lCode
                    Else
                        tlVof.lFt1CefCode = 0
                    End If
                Else
                    tlVof.lFt1CefCode = 0
                End If
            Else
                tlVof.lFt1CefCode = 0
            End If
            tlCefSrchKey.lCode = tlVof.lFt2CefCode
            If tlCefSrchKey.lCode <> 0 Then
                tlCef.sComment = ""
                ilCefRecLen = Len(tlCef)    '1009
                ilRet = btrGetEqual(hlCef, tlCef, ilCefRecLen, tlCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    ilCefRecLen = Len(tlCef)    '5 + Len(Trim$(tlCef.sComment)) + 2
                    tlCef.lCode = 0
                    ilRet = btrInsert(hlCef, tlCef, ilCefRecLen, INDEXKEY0)
                    If ilRet = BTRV_ERR_NONE Then
                        tlVof.lFt2CefCode = tlCef.lCode
                    Else
                        tlVof.lFt2CefCode = 0
                    End If
                Else
                    tlVof.lFt2CefCode = 0
                End If
            Else
                tlVof.lFt2CefCode = 0
            End If
            tlVof.iVefCode = tmVef.iCode
            ilRet = btrInsert(hlVof, tlVof, ilRecLen, INDEXKEY0)
        End If
    Next ilPass
    ilRet = btrClose(hlCef)
    btrDestroy hlCef
    ilRet = btrClose(hlVof)
    btrDestroy hlVof
    mAddVofModel = True
    Exit Function

    imTerminate = True
    On Error GoTo 0
    btrDestroy hlCef
    btrDestroy hlVof
    mAddVofModel = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddVpfModel                    *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add VPF                        *
'*                                                     *
'*******************************************************
Private Function mAddVpfModel(slSyncDate As String, slSyncTime As String)
    Dim hlVpf As Integer        'site Option file handle
    Dim ilRecLen As Integer     'Vpf record length
    Dim tlVpf As VPF
    Dim tlSrchKey As VPFKEY0
    Dim ilRet As Integer
    hlVpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mAddVpfModel = False
        Exit Function
    End If
    ilRecLen = Len(tlVpf)  'btrRecordLength(hlVpf)  'Get and save record length
    tlSrchKey.iVefKCode = igVefCodeModel
    ilRet = btrGetEqual(hlVpf, tlVpf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlVpf)
        btrDestroy hlVpf
        mAddVpfModel = False
        Exit Function
    End If
    tlVpf.iVefKCode = tmVef.iCode
    tlVpf.sStnFdCode = ""
    tlVpf.sStnFdCart = "N"
    tlVpf.sStnFdXRef = "Y"
    tlVpf.iFTPArfCode = 0
    tlVpf.iProducerArfCode = 0
    tlVpf.iProgProvArfCode = 0
    tlVpf.iCommProvArfCode = 0
    tlVpf.iAutoExptArfCode = 0
    tlVpf.iAutoImptArfCode = 0
    tlVpf.iInterfaceID = 0
    'gPackDate slSyncDate, tlVpf.iSyncDate(0), tlVpf.iSyncDate(1)
    'gPackTime slSyncTime, tlVpf.iSyncTime(0), tlVpf.iSyncTime(1)
    ilRet = btrInsert(hlVpf, tlVpf, ilRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlVpf)
        btrDestroy hlVpf
        mAddVpfModel = False
        Exit Function
    End If
    ilRet = btrClose(hlVpf)
    btrDestroy hlVpf
    
    ReDim Preserve tgVpf(0 To UBound(tgVpf) + 1) As VPF
    tgVpf(UBound(tgVpf)) = tlVpf
    If UBound(tgVpf) > 1 Then
        ArraySortTyp fnAV(tgVpf(), 0), UBound(tgVpf) + 1, 0, LenB(tgVpf(0)), 0, -1, 0
    End If
    
    '11/26/17
    gFileChgdUpdate "vpf.btr", False
            
    mAddVpfModel = True
    Exit Function

    imTerminate = True
    On Error GoTo 0
    btrDestroy hlVpf
    mAddVpfModel = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mBookNamePop                    *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Book Name list box    *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mBookPop(ilAllBooks As Integer)
'
'   mBookNamePop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcBook.ListIndex
    If ilIndex > 1 Then
        slName = lbcBook.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen

    If imSelectedIndex = 0 Then
        ''ilRet = gPopBookNameBox(Vehicle, 0, 1, 1, lbcBook, lbcBookCode)
        'Don't show any book names for newe vehicles
        'ilRet = gPopBookNameBox(Vehicle, 0, 0, 0, 1, 1, lbcBook, tmBookCode(), smBookNameTag)
        If smBookNameTag = "New" Then
            Exit Sub
        End If
        lbcBook.Clear
        smBookNameTag = "New"
        ReDim tmBookCode(0 To 0) As SORTCODE
        lbcBook.AddItem "[None]", 0
        Exit Sub
    Else
        ''ilRet = gPopBookNameBox(Vehicle, tmVef.iCode, 1, 1, lbcBook, lbcBookCode)
        '5/2/06: Premiere was unable to set books when vehicle had only one book that could be used
        'If ilAllBooks Then
            ilRet = gPopBookNameBox(Vehicle, 0, 0, tmVef.iCode, 1, 1, lbcBook, tmBookCode(), smBookNameTag)
        'Else
        '    ilRet = gPopBookNameBox(Vehicle, 1, 0, tmVef.iCode, 1, 1, lbcBook, tmBookCode(), smBookNameTag)
        'End If
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mBookPopErr
        gCPErrorMsg ilRet, "mBookPop (gPopBookBox)", Vehicle
        On Error GoTo 0
        lbcBook.AddItem "[None]", 0
        imBookChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcBook
            If gLastFound(lbcBook) > 1 Then
                lbcBook.ListIndex = gLastFound(lbcBook)
            Else
                lbcBook.ListIndex = -1
            End If
        Else
            lbcBook.ListIndex = ilIndex
        End If
        imBookChgMode = False
    End If
    Exit Sub
mBookPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer
    smOrigName = ""
    edcName.Text = ""
    edcContact.Text = ""
    'edcMktName.Text = ""
    For ilLoop = 0 To 2 Step 1
        edcAddr(ilLoop).Text = ""
    Next ilLoop
    mkcPhone.Text = smPhoneImage
    mkcFax.Text = smFaxImage
    'edcFormat.Text = ""
    edcDialPos.Text = ""
    edcACT1Lineup.Text = ""
    imLVChgMode = True
    edcLogVehDropDown.Text = ""
    imLVChgMode = False
    smLogVeh = ""
    imMVLChgMode = True
    edcMultiVehDropdown.Text = ""
    imMVLChgMode = False
    smMultiVehLog = ""
    imBookChgMode = True
    edcBookDropDown.Text = ""
    imBookChgMode = False
    smBook = ""
    imBookChgMode = True
    edcReallDropDown.Text = ""
    imBookChgMode = False
    smReall = ""
    imDemoChgMode = True
    edcDemoDropDown.Text = ""
    imDemoChgMode = False
    smDemo = ""
    imRNLinkChgMode = True
    edcRNLinkDropdown.Text = ""
    imRNLinkChgMode = False
    smRNLink = ""
    imTaxChgMode = True
    edcTaxDropdown.Text = ""
    imTaxChgMode = False
    smOrigTax = ""
    imHubChgMode = True
    edcHubDropdown.Text = ""
    imHubChgMode = False
    smOrigHub = ""
    edcSort.Text = ""
    edcStationCode.Text = ""
    'If tgSpf.sSSellNet = "Y" Then
        smOrigType = "" 'type index
    imTypeChgMode = True
    edcTypeDropDown.Text = ""
    imTypeChgMode = False
    smVehicleType = ""
    'Else
    '    imType = 0
    'End If
    smOrigType = ""
    imState = -1
    imOrigState = -1
    imVehGp2ChgMode = True
    edcVehGp2DropDown.Text = ""
    imVehGp2ChgMode = False
    imVehGp3ChgMode = True
    edcVehGp3DropDown.Text = ""
    imVehGp3ChgMode = False
    imVehGp4ChgMode = True
    edcVehGp4DropDown.Text = ""
    imVehGp4ChgMode = False
    imVehGp5ChgMode = True
    edcVehGp5DropDown.Text = ""
    imVehGp5ChgMode = False
    imVehGp6ChgMode = True
    edcVehGp6DropDown.Text = ""
    imVehGp6ChgMode = False
    smOrigVehGp2 = ""
    smVehGp2 = ""
    smOrigVehGp3 = ""
    smVehGp3 = ""
    smOrigVehGp4 = ""
    smVehGp4 = ""
    smOrigVehGp5 = ""
    smVehGp5 = ""
    smOrigVehGp6 = ""
    smVehGp6 = ""
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    ckcLock.Visible = False
    bmPrgLibDefined = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateVsf                      *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create VSF                     *
'*                                                     *
'*******************************************************
Private Function mCreateVsf() As Long
    Dim ilLoop As Integer
    Dim ilRet As Integer
    tmVsf.lCode = 0
    tmVsf.sType = "V"
    tmVsf.sName = Left$(tmVef.sName, 20)
    tmVsf.lLkVsfCode = 0
    'tmVsf.sMktName = ""
    For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
        tmVsf.iFSCode(ilLoop) = 0
        'slStr = ""
        'gStrToPDN slStr, 4, 4, tmVsf.sFSComm(ilLoop)
        tmVsf.lFSComm(ilLoop) = 0
        tmVsf.iNoSpots(ilLoop) = 0
    Next ilLoop
    tmVsf.sSource = "U"
    tmVsf.iMerge = 0
    ilRet = btrInsert(hmVsf, tmVsf, imVsfRecLen, INDEXKEY0)
    mCreateVsf = tmVsf.lCode
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteVof                      *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Delete Vof                     *
'*                                                     *
'*******************************************************
Function mDeleteVof(ilCode As Integer, slType As String) As Integer
    Dim hlVof As Integer        'site Option file handle
    Dim ilRecLen As Integer     'Vof record length
    Dim tlVof As VOF
    Dim tlSrchKey As VOFKEY0
    Dim ilRet As Integer

    If (slType = "C") Or (slType = "L") Or (slType = "A") Then
        hlVof = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hlVof, "", sgDBPath & "Vof.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mDeleteVof = False
            Exit Function
        End If
        ilRecLen = Len(tlVof)  'btrRecordLength(hlVof)  'Get and save record length
        tlSrchKey.iVefCode = ilCode
        tlSrchKey.sType = "L"
        ilRet = btrGetEqual(hlVof, tlVof, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hlVof)
        End If
        tlSrchKey.iVefCode = ilCode
        tlSrchKey.sType = "C"
        ilRet = btrGetEqual(hlVof, tlVof, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hlVof)
        End If
        tlSrchKey.iVefCode = ilCode
        tlSrchKey.sType = "O"
        ilRet = btrGetEqual(hlVof, tlVof, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hlVof)
        End If
        On Error GoTo 0
        ilRet = btrClose(hlVof)
        btrDestroy hlVof
    End If
    mDeleteVof = True
    Exit Function

   'Ignore error as record might not ahve been created
    'imTerminate = True
    On Error GoTo 0
    btrDestroy hlVof
    mDeleteVof = True   'False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteVpf                      *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Delete VPF                     *
'*                                                     *
'*******************************************************
Private Function mDeleteVpf(ilCode As Integer) As Integer
    Dim hlVpf As Integer        'site Option file handle
    Dim ilRecLen As Integer     'Vpf record length
    Dim tlVpf As VPF
    Dim tlSrchKey As VPFKEY0
    Dim ilRet As Integer
    hlVpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mDeleteVpf = False
        Exit Function
    End If
    ilRecLen = Len(tlVpf)  'btrRecordLength(hlVpf)  'Get and save record length
    tlSrchKey.iVefKCode = ilCode
    ilRet = btrGetEqual(hlVpf, tlVpf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    On Error GoTo mDeleteVpfErr
    gBtrvErrorMsg ilRet, "mDeleteVpf (btrGetEqual)", Vehicle
    On Error GoTo 0
    ilRet = btrDelete(hlVpf)
    On Error GoTo mDeleteVpfErr
    gBtrvErrorMsg ilRet, "mDeleteVpf (btrDelete)", Vehicle
    On Error GoTo 0
    ilRet = btrClose(hlVpf)
    btrDestroy hlVpf
    mDeleteVpf = True
    Exit Function
mDeleteVpfErr:
    imTerminate = True
    On Error GoTo 0
    btrDestroy hlVpf
    mDeleteVpf = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteVsf                      *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Delete VSF                     *
'*                                                     *
'*******************************************************
Private Function mDeleteVsf(llVsfCode As Long) As Integer
    Dim ilRet As Integer
    ilRet = BTRV_ERR_NONE
    If llVsfCode <= 0 Then
        Exit Function
    End If
    Do
        tmVsfSrchKey.lCode = llVsfCode
        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmVsf)
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    mDeleteVsf = ilRet
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteVff                      *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Delete VFF                     *
'*                                                     *
'*******************************************************
Private Function mDeleteVff(ilVefCode As Integer) As Integer
    Dim ilRet As Integer
    ilRet = BTRV_ERR_NONE
    Do
        tmVffSrchKey1.iCode = ilVefCode
        ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmVff)
        Else
            ilRet = BTRV_ERR_NONE
            Exit Function
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    mDeleteVff = ilRet
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteVaf                      *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Delete VAF                     *
'*                                                     *
'*******************************************************
Private Function mDeleteVaf(ilVefCode As Integer) As Integer
    Dim ilRet As Integer
    ilRet = BTRV_ERR_NONE
    Do
        tmVafSrchKey1.iCode = ilVefCode
        ilRet = btrGetEqual(hmVaf, tmVaf, imVafRecLen, tmVafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmVaf)
        Else
            ilRet = BTRV_ERR_NONE
            Exit Function
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    mDeleteVaf = ilRet
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteVbf                      *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Delete VBF                     *
'*                                                     *
'*******************************************************
Private Function mDeleteVbf(ilVefCode As Integer) As Integer
    Dim ilRet As Integer
    tmVbfSrchKey1.iVefCode = ilVefCode
    tmVbfSrchKey1.iStartDate(0) = 0
    tmVbfSrchKey1.iStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmVbf, tmVbf, imVBfRecLen, tmVbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmVbf.iVefCode = ilVefCode)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmVbf)
        Else
            ilRet = BTRV_ERR_NONE
            Exit Do
        End If
        ilRet = btrGetNext(hmVbf, tmVbf, imVBfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mDeleteVbf = BTRV_ERR_NONE
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteNif                      *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Delete NIF                     *
'*                                                     *
'*******************************************************
Private Function mDeleteNif(ilVefCode As Integer) As Integer
    Dim ilRet As Integer
    tmNifSrchKey1.iVefCode = ilVefCode
    tmNifSrchKey1.iYear = 0
    ilRet = btrGetGreaterOrEqual(hmNif, tmNif, imNifRecLen, tmNifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmNif.iVefCode = ilVefCode)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmNif)
        Else
            ilRet = BTRV_ERR_NONE
            Exit Do
        End If
        ilRet = btrGetNext(hmNif, tmNif, imNifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mDeleteNif = BTRV_ERR_NONE
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveGroupName                *
'*                                                     *
'*             Created:12/21/95      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove group name              *
'*                                                     *
'*******************************************************
Private Function mRemoveGroupName(ilVefCode As Integer) As Integer
    Dim ilRet As Integer
    ilRet = BTRV_ERR_NONE
    Do
        tmVffSrchKey1.iCode = ilVefCode
        ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tmVff.sGroupName = ""
            ilRet = btrUpdate(hmVff, tmVff, imVffRecLen)
        Else
            ilRet = BTRV_ERR_NONE
            Exit Function
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    mRemoveGroupName = ilRet
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mDemoBranch                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Demo and process               *
'*                      communication back from        *
'*                      Demo                           *
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
Private Function mDemoBranch() As Integer
'
'   ilRet = mDemoBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcDemoDropDown, lbcDemo, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcDemoDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mDemoBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mDemoBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "D"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcDemoDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mDemoBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcDemo.Clear
        smDemoCodeTag = ""
        sgDemoMnfStamp = ""
        mDemoPop
        If imTerminate Then
            mDemoBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcDemo
        sgMNmName = ""
        If gLastFound(lbcDemo) > 0 Then
            imDemoChgMode = True
            lbcDemo.ListIndex = gLastFound(lbcDemo)
            edcDemoDropDown.Text = lbcDemo.List(lbcDemo.ListIndex)
            imDemoChgMode = False
            mDemoBranch = False
            mSetChg imBoxNo
        Else
            imDemoChgMode = True
            lbcDemo.ListIndex = 1
            edcDemoDropDown.Text = lbcDemo.List(1)
            imDemoChgMode = False
            mSetChg imBoxNo
            edcDemoDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mDemoPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Demo List box if      *
'*                      required                       *
'*                                                     *
'*******************************************************
Private Sub mDemoPop()
'
'   mDemoPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcDemo.ListIndex
    If ilIndex > 1 Then
        slName = lbcDemo.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilfilter(0) = CHARFILTER
    slFilter(0) = "D"
    ilOffset(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(Vehicle, lbcDemo, lbcDemoCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Vehicle, lbcDemo, tmDemoCode(), smDemoCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mDemoPopErr
        gCPErrorMsg ilRet, "mDemoPop (gIMoveListBox)", Vehicle
        On Error GoTo 0
        lbcDemo.AddItem "[None]", 0  'Force as first item on list
        lbcDemo.AddItem "[New]", 0  'Force as first item on list
        imDemoChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcDemo
            If gLastFound(lbcDemo) > 1 Then
                lbcDemo.ListIndex = gLastFound(lbcDemo)
            Else
                lbcDemo.ListIndex = -1
            End If
        Else
            lbcDemo.ListIndex = ilIndex
        End If
        imDemoChgMode = False
    End If
    Exit Sub
mDemoPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilMax                                                   *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxCtrlNo) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Width = tmCtrls(ilBoxNo).fBoxW
            If tgSpf.iVehLen <= 40 Then
                edcName.MaxLength = tgSpf.iVehLen
            Else
                edcName.MaxLength = 20
            End If
            gMoveFormCtrl pbcVeh, edcName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case CONTACTINDEX 'Name
            edcContact.Width = tmCtrls(ilBoxNo).fBoxW
            edcContact.MaxLength = 40
            gMoveFormCtrl pbcVeh, edcContact, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcContact.Visible = True  'Set visibility
            edcContact.SetFocus
        Case MKTNAMEINDEX 'Market Name
            'edcMktName.Width = tmCtrls(ilBoxNo).fBoxW
            'edcMktName.MaxLength = 20
            'gMoveFormCtrl pbcVeh, edcMktName, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            'edcMktName.Visible = True  'Set visibility
            'edcMktName.SetFocus
            mVehGp3Pop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehGp3.Height = gListBoxHeight(lbcVehGp3.ListCount, 6)
            edcVehGp3DropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 40
            lbcVehGp3.Width = 5000
            edcVehGp3DropDown.MaxLength = 50
            gMoveFormCtrl pbcVeh, edcVehGp3DropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcVehGp3DropDown.Left + edcVehGp3DropDown.Width, edcVehGp3DropDown.Top
            imVehGp3ChgMode = True
            slStr = smVehGp3
            gFindMatch slStr, 1, lbcVehGp3
            If gLastFound(lbcVehGp3) >= 1 Then
                lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
                edcVehGp3DropDown.Text = lbcVehGp3.List(lbcVehGp3.ListIndex)
            Else
                If lbcVehGp3.ListCount > 1 Then
                    lbcVehGp3.ListIndex = 1
                    edcVehGp3DropDown.Text = lbcVehGp3.List(1)
                Else
                    lbcVehGp3.ListIndex = 0
                    edcVehGp3DropDown.Text = lbcVehGp3.List(0)
                End If
            End If
            imVehGp3ChgMode = False
            lbcVehGp3.Move edcVehGp3DropDown.Left, edcVehGp3DropDown.Top + edcVehGp3DropDown.Height
            edcVehGp3DropDown.SelStart = 0
            edcVehGp3DropDown.SelLength = Len(edcVehGp3DropDown.Text)
            edcVehGp3DropDown.Visible = True
            cmcDropDown.Visible = True
            edcVehGp3DropDown.SetFocus
        Case RSCHINDEX
            mVehGp5Pop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehGp5.Height = gListBoxHeight(lbcVehGp5.ListCount, 6)
            edcVehGp5DropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 40
            lbcVehGp5.Width = 5000
            edcVehGp5DropDown.MaxLength = 50
            gMoveFormCtrl pbcVeh, edcVehGp5DropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcVehGp5DropDown.Left + edcVehGp5DropDown.Width, edcVehGp5DropDown.Top
            
            imVehGp5ChgMode = True
            slStr = smVehGp5
            gFindMatch slStr, 1, lbcVehGp5
            If gLastFound(lbcVehGp5) >= 1 Then
                lbcVehGp5.ListIndex = gLastFound(lbcVehGp5)
                edcVehGp5DropDown.Text = lbcVehGp5.List(lbcVehGp5.ListIndex)
            Else
                If lbcVehGp5.ListCount > 1 Then
                    lbcVehGp5.ListIndex = 1
                    edcVehGp5DropDown.Text = lbcVehGp5.List(1)
                Else
                    lbcVehGp5.ListIndex = 0
                    edcVehGp5DropDown.Text = lbcVehGp5.List(0)
                End If
            End If
            imVehGp5ChgMode = False
            'lbcVehGp.Move edcVehGpDropDown.Left, edcVehGpDropDown.Top + edcVehGpDropDown.Height
            'lbcVehGp5.Move cmcDropDown.Left + cmcDropDown.Width - lbcVehGp5.Width, edcVehGp5DropDown.Top + edcVehGp5DropDown.Height
            'lbcVehGp5.Move cmcDropDown.Left + cmcDropDown.Width - lbcVehGp5.Width, edcVehGp5DropDown.Top + edcVehGp5DropDown.Height
            
            lbcVehGp5.Move edcVehGp5DropDown.Left, edcVehGp5DropDown.Top + edcVehGp5DropDown.Height
            
            
            edcVehGp5DropDown.SelStart = 0
            edcVehGp5DropDown.SelLength = Len(edcVehGp5DropDown.Text)
            edcVehGp5DropDown.Visible = True
            cmcDropDown.Visible = True
            edcVehGp5DropDown.SetFocus
        Case SUBCOMPINDEX 'Sub-Company
            mVehGp6Pop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehGp6.Height = gListBoxHeight(lbcVehGp6.ListCount, 6)
            edcVehGp6DropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 40
            lbcVehGp6.Width = 5000
            edcVehGp6DropDown.MaxLength = 50
            gMoveFormCtrl pbcVeh, edcVehGp6DropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcVehGp6DropDown.Left + edcVehGp6DropDown.Width, edcVehGp6DropDown.Top
            imVehGp6ChgMode = True
            slStr = smVehGp6
            gFindMatch slStr, 1, lbcVehGp6
            If gLastFound(lbcVehGp6) >= 1 Then
                lbcVehGp6.ListIndex = gLastFound(lbcVehGp6)
                edcVehGp6DropDown.Text = lbcVehGp6.List(lbcVehGp6.ListIndex)
            Else
                If lbcVehGp6.ListCount > 1 Then
                    lbcVehGp6.ListIndex = 1
                    edcVehGp6DropDown.Text = lbcVehGp6.List(1)
                Else
                    lbcVehGp6.ListIndex = 0
                    edcVehGp6DropDown.Text = lbcVehGp6.List(0)
                End If
            End If
            imVehGp6ChgMode = False
            lbcVehGp6.Move edcVehGp6DropDown.Left, edcVehGp6DropDown.Top + edcVehGp6DropDown.Height
            edcVehGp6DropDown.SelStart = 0
            edcVehGp6DropDown.SelLength = Len(edcVehGp6DropDown.Text)
            edcVehGp6DropDown.Visible = True
            cmcDropDown.Visible = True
            edcVehGp6DropDown.SetFocus
        Case ADDRESSINDEX 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).Width = tmCtrls(ADDRESSINDEX).fBoxW
            edcAddr(ilBoxNo - ADDRESSINDEX).MaxLength = 25
            gMoveFormCtrl pbcVeh, edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAddr(ilBoxNo - ADDRESSINDEX).Visible = True  'Set visibility
            edcAddr(ilBoxNo - ADDRESSINDEX).SetFocus
        Case ADDRESSINDEX + 1 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).Width = tmCtrls(ADDRESSINDEX).fBoxW
            edcAddr(ilBoxNo - ADDRESSINDEX).MaxLength = 25
            gMoveFormCtrl pbcVeh, edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAddr(ilBoxNo - ADDRESSINDEX).Visible = True  'Set visibility
            edcAddr(ilBoxNo - ADDRESSINDEX).SetFocus
        Case ADDRESSINDEX + 2 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).Width = tmCtrls(ADDRESSINDEX).fBoxW
            edcAddr(ilBoxNo - ADDRESSINDEX).MaxLength = 25
            gMoveFormCtrl pbcVeh, edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAddr(ilBoxNo - ADDRESSINDEX).Visible = True  'Set visibility
            edcAddr(ilBoxNo - ADDRESSINDEX).SetFocus
        Case ADDRESSINDEX + 3 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).Width = tmCtrls(ADDRESSINDEX).fBoxW
            edcAddr(ilBoxNo - ADDRESSINDEX).MaxLength = 25
            gMoveFormCtrl pbcVeh, edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcAddr(ilBoxNo - ADDRESSINDEX).Visible = True  'Set visibility
            edcAddr(ilBoxNo - ADDRESSINDEX).SetFocus
        Case PHONEINDEX 'Phone and extension
            mkcPhone.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcVeh, mkcPhone, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            mkcPhone.Visible = True  'Set visibility
            mkcPhone.SetFocus
        Case FAXINDEX 'Fax
            mkcFax.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcVeh, mkcFax, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            mkcFax.Visible = True  'Set visibility
            mkcFax.SetFocus
        Case FORMATINDEX 'Format
            'edcFormat.Width = tmCtrls(ilBoxNo).fBoxW
            'edcFormat.MaxLength = 20
            'gMoveFormCtrl pbcVeh, edcFormat, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            'edcFormat.Visible = True  'Set visibility
            'edcFormat.SetFocus
            mVehGp4Pop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehGp4.Height = gListBoxHeight(lbcVehGp4.ListCount, 6)
            edcVehGp4DropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 40
            lbcVehGp4.Width = 5000
            edcVehGp4DropDown.MaxLength = 50
            gMoveFormCtrl pbcVeh, edcVehGp4DropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcVehGp4DropDown.Left + edcVehGp4DropDown.Width, edcVehGp4DropDown.Top
            imVehGp4ChgMode = True
            slStr = smVehGp4
            gFindMatch slStr, 1, lbcVehGp4
            If gLastFound(lbcVehGp4) >= 1 Then
                lbcVehGp4.ListIndex = gLastFound(lbcVehGp4)
                edcVehGp4DropDown.Text = lbcVehGp4.List(lbcVehGp4.ListIndex)
            Else
                If lbcVehGp4.ListCount > 1 Then
                    lbcVehGp4.ListIndex = 1
                    edcVehGp4DropDown.Text = lbcVehGp4.List(1)
                Else
                    lbcVehGp4.ListIndex = 0
                    edcVehGp4DropDown.Text = lbcVehGp4.List(0)
                End If
            End If
            imVehGp4ChgMode = False
            'lbcVehGp.Move edcVehGpDropDown.Left, edcVehGpDropDown.Top + edcVehGpDropDown.Height
            'lbcVehGp4.Move edcVehGp4DropDown.Left, edcVehGp4DropDown.Top + edcVehGp4DropDown.Height
            lbcVehGp4.Move cmcDropDown.Left + cmcDropDown.Width - lbcVehGp4.Width, edcVehGp5DropDown.Top + edcVehGp4DropDown.Height
            edcVehGp4DropDown.SelStart = 0
            edcVehGp4DropDown.SelLength = Len(edcVehGp4DropDown.Text)
            edcVehGp4DropDown.Visible = True
            cmcDropDown.Visible = True
            edcVehGp4DropDown.SetFocus
        Case SUBTOTALINDEX
            mVehGp2Pop
            If imTerminate Then
                Exit Sub
            End If
            lbcVehGp2.Height = gListBoxHeight(lbcVehGp2.ListCount, 6)
            edcVehGp2DropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 40
            lbcVehGp2.Width = 5000
            edcVehGp2DropDown.MaxLength = 50
            gMoveFormCtrl pbcVeh, edcVehGp2DropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcVehGp2DropDown.Left + edcVehGp2DropDown.Width, edcVehGp2DropDown.Top
            imVehGp2ChgMode = True
            slStr = smVehGp2
            gFindMatch slStr, 1, lbcVehGp2
            If gLastFound(lbcVehGp2) >= 1 Then
                lbcVehGp2.ListIndex = gLastFound(lbcVehGp2)
                edcVehGp2DropDown.Text = lbcVehGp2.List(lbcVehGp2.ListIndex)
            Else
                If lbcVehGp2.ListCount > 1 Then
                    lbcVehGp2.ListIndex = 1
                    edcVehGp2DropDown.Text = lbcVehGp2.List(1)
                Else
                    lbcVehGp2.ListIndex = 0
                    edcVehGp2DropDown.Text = lbcVehGp2.List(0)
                End If
            End If
            imVehGp2ChgMode = False
            'lbcVehGp.Move edcVehGpDropDown.Left, edcVehGpDropDown.Top + edcVehGpDropDown.Height
            lbcVehGp2.Move cmcDropDown.Left + cmcDropDown.Width - lbcVehGp2.Width, edcVehGp2DropDown.Top + edcVehGp2DropDown.Height
            edcVehGp2DropDown.SelStart = 0
            edcVehGp2DropDown.SelLength = Len(edcVehGp2DropDown.Text)
            edcVehGp2DropDown.Visible = True
            cmcDropDown.Visible = True
            edcVehGp2DropDown.SetFocus
        Case LOGVEHINDEX 'Vehicle
            mLogVehPop
            If imTerminate Then
                Exit Sub
            End If
            lbcLogVeh.Height = gListBoxHeight(lbcLogVeh.ListCount, 8)
            edcLogVehDropDown.Width = tmCtrls(LOGVEHINDEX).fBoxW - cmcDropDown.Width
            If tgSpf.iVehLen <= 40 Then
                edcLogVehDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcLogVehDropDown.MaxLength = 20
            End If
            gMoveFormCtrl pbcVeh, edcLogVehDropDown, tmCtrls(LOGVEHINDEX).fBoxX, tmCtrls(LOGVEHINDEX).fBoxY
            cmcDropDown.Move edcLogVehDropDown.Left + edcLogVehDropDown.Width, edcLogVehDropDown.Top
            imLVChgMode = True
            slStr = edcLogVehDropDown.Text
            gFindMatch slStr, 0, lbcLogVeh
            If gLastFound(lbcLogVeh) >= 0 Then
                lbcLogVeh.ListIndex = gLastFound(lbcLogVeh)
                edcLogVehDropDown.Text = lbcLogVeh.List(lbcLogVeh.ListIndex)
            Else
                lbcLogVeh.ListIndex = 0
                edcLogVehDropDown.Text = lbcLogVeh.List(0)
            End If
            imComboBoxIndex = lbcLogVeh.ListIndex
            imLVChgMode = False
            lbcLogVeh.Move edcLogVehDropDown.Left, edcLogVehDropDown.Top + edcLogVehDropDown.Height
            edcLogVehDropDown.SelStart = 0
            edcLogVehDropDown.SelLength = Len(edcLogVehDropDown.Text)
            edcLogVehDropDown.Visible = True
            cmcDropDown.Visible = True
            edcLogVehDropDown.SetFocus
        Case MULTIVEHLOGINDEX 'Vehicle
            If (smVehicleType = "C") Or (smVehicleType = "A") Then
                mMultiConvVehLogPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcMultiConvVehLog.Height = gListBoxHeight(lbcMultiConvVehLog.ListCount, 8)
                edcMultiVehDropdown.Width = tmCtrls(MULTIVEHLOGINDEX).fBoxW - cmcDropDown.Width
                If tgSpf.iVehLen <= 40 Then
                    edcMultiVehDropdown.MaxLength = tgSpf.iVehLen
                Else
                    edcMultiVehDropdown.MaxLength = 20
                End If
                gMoveFormCtrl pbcVeh, edcMultiVehDropdown, tmCtrls(MULTIVEHLOGINDEX).fBoxX, tmCtrls(MULTIVEHLOGINDEX).fBoxY
                cmcDropDown.Move edcMultiVehDropdown.Left + edcMultiVehDropdown.Width, edcMultiVehDropdown.Top
                imMVLChgMode = True
                slStr = edcMultiVehDropdown.Text
                gFindMatch slStr, 0, lbcMultiConvVehLog
                If gLastFound(lbcMultiConvVehLog) >= 0 Then
                    lbcMultiConvVehLog.ListIndex = gLastFound(lbcMultiConvVehLog)
                    edcMultiVehDropdown.Text = lbcMultiConvVehLog.List(lbcMultiConvVehLog.ListIndex)
                Else
                    lbcMultiConvVehLog.ListIndex = 0
                    edcMultiVehDropdown.Text = lbcMultiConvVehLog.List(0)
                End If
                imComboBoxIndex = lbcMultiConvVehLog.ListIndex
                imMVLChgMode = False
                lbcMultiConvVehLog.Move edcMultiVehDropdown.Left, edcMultiVehDropdown.Top + edcMultiVehDropdown.Height
                edcMultiVehDropdown.SelStart = 0
                edcMultiVehDropdown.SelLength = Len(edcMultiVehDropdown.Text)
                edcMultiVehDropdown.Visible = True
                cmcDropDown.Visible = True
                edcMultiVehDropdown.SetFocus
            ElseIf smVehicleType = "G" Then
                mMultiGameVehLogPop
                If imTerminate Then
                    Exit Sub
                End If
                lbcMultiGameVehLog.Height = gListBoxHeight(lbcMultiGameVehLog.ListCount, 8)
                edcMultiVehDropdown.Width = tmCtrls(MULTIVEHLOGINDEX).fBoxW - cmcDropDown.Width
                If tgSpf.iVehLen <= 40 Then
                    edcMultiVehDropdown.MaxLength = tgSpf.iVehLen
                Else
                    edcMultiVehDropdown.MaxLength = 20
                End If
                gMoveFormCtrl pbcVeh, edcMultiVehDropdown, tmCtrls(MULTIVEHLOGINDEX).fBoxX, tmCtrls(MULTIVEHLOGINDEX).fBoxY
                cmcDropDown.Move edcMultiVehDropdown.Left + edcMultiVehDropdown.Width, edcMultiVehDropdown.Top
                imMVLChgMode = True
                slStr = edcMultiVehDropdown.Text
                gFindMatch slStr, 0, lbcMultiGameVehLog
                If gLastFound(lbcMultiGameVehLog) >= 0 Then
                    lbcMultiGameVehLog.ListIndex = gLastFound(lbcMultiGameVehLog)
                    edcMultiVehDropdown.Text = lbcMultiGameVehLog.List(lbcMultiGameVehLog.ListIndex)
                Else
                    lbcMultiGameVehLog.ListIndex = 0
                    edcMultiVehDropdown.Text = lbcMultiGameVehLog.List(0)
                End If
                imComboBoxIndex = lbcMultiGameVehLog.ListIndex
                imMVLChgMode = False
                lbcMultiGameVehLog.Move edcMultiVehDropdown.Left, edcMultiVehDropdown.Top + edcMultiVehDropdown.Height
                edcMultiVehDropdown.SelStart = 0
                edcMultiVehDropdown.SelLength = Len(edcMultiVehDropdown.Text)
                edcMultiVehDropdown.Visible = True
                cmcDropDown.Visible = True
                edcMultiVehDropdown.SetFocus
            End If
        Case RNLINKINDEX 'Vehicle
            mRNLinkPop
            If imTerminate Then
                Exit Sub
            End If
            lbcRNLink.Height = gListBoxHeight(lbcRNLink.ListCount, 8)
            edcRNLinkDropdown.Width = tmCtrls(RNLINKINDEX).fBoxW - cmcDropDown.Width
            edcRNLinkDropdown.MaxLength = 30
            gMoveFormCtrl pbcVeh, edcRNLinkDropdown, tmCtrls(RNLINKINDEX).fBoxX, tmCtrls(RNLINKINDEX).fBoxY
            cmcDropDown.Move edcRNLinkDropdown.Left + edcRNLinkDropdown.Width, edcRNLinkDropdown.Top
            imRNLinkChgMode = True
            slStr = edcRNLinkDropdown.Text
            gFindMatch slStr, 0, lbcRNLink
            If gLastFound(lbcRNLink) >= 0 Then
                lbcRNLink.ListIndex = gLastFound(lbcRNLink)
                edcRNLinkDropdown.Text = lbcRNLink.List(lbcRNLink.ListIndex)
            Else
                lbcRNLink.ListIndex = 0
                edcRNLinkDropdown.Text = lbcRNLink.List(0)
            End If
            imComboBoxIndex = lbcRNLink.ListIndex
            imRNLinkChgMode = False
            lbcRNLink.Move edcRNLinkDropdown.Left, edcRNLinkDropdown.Top + edcRNLinkDropdown.Height
            edcRNLinkDropdown.SelStart = 0
            edcRNLinkDropdown.SelLength = Len(edcRNLinkDropdown.Text)
            edcRNLinkDropdown.Visible = True
            cmcDropDown.Visible = True
            edcRNLinkDropdown.SetFocus
        Case DIALPOSINDEX 'Dial Position
            edcDialPos.Width = tmCtrls(ilBoxNo).fBoxW
            edcDialPos.MaxLength = 5
            gMoveFormCtrl pbcVeh, edcDialPos, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcDialPos.Visible = True  'Set visibility
            edcDialPos.SetFocus
        Case ACT1CODESINDEX 'ACT1 Codes
            edcACT1Lineup.Width = tmCtrls(ilBoxNo).fBoxW
            edcACT1Lineup.MaxLength = 11
            gMoveFormCtrl pbcVeh, edcACT1Lineup, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcACT1Lineup.Visible = True  'Set visibility
            edcACT1Lineup.SetFocus
        Case TAXINDEX 'Tax
            lbcTax.Height = gListBoxHeight(lbcTax.ListCount, 8)
            If imACT1CodesDefined Then
                edcTaxDropdown.Width = tmCtrls(ACT1CODESINDEX).fBoxW + tmCtrls(TAXINDEX).fBoxW - cmcDropDown.Width
            Else
                edcTaxDropdown.Width = tmCtrls(TAXINDEX).fBoxW - cmcDropDown.Width
            End If
            edcTaxDropdown.MaxLength = 0
            gMoveFormCtrl pbcVeh, edcTaxDropdown, tmCtrls(TAXINDEX).fBoxX, tmCtrls(TAXINDEX).fBoxY
            cmcDropDown.Move edcTaxDropdown.Left + edcTaxDropdown.Width, edcTaxDropdown.Top
            imTaxChgMode = True
            slStr = edcTaxDropdown.Text
            gFindMatch slStr, 0, lbcTax
            If gLastFound(lbcTax) >= 0 Then
                lbcTax.ListIndex = gLastFound(lbcTax)
                edcTaxDropdown.Text = lbcTax.List(lbcTax.ListIndex)
            Else
                lbcTax.ListIndex = 0
                edcTaxDropdown.Text = lbcTax.List(0)
            End If
            imComboBoxIndex = lbcTax.ListIndex
            imTaxChgMode = False
            ''lbcTax.Move edcTaxDropdown.Left + edcTaxDropdown.Width + cmcDropDown.Width - lbcTax.Width, edcTaxDropdown.Top + edcTaxDropdown.Height
            'lbcTax.Move edcTaxDropdown.Left + edcTaxDropdown.Width + cmcDropDown.Width - lbcTax.Width, edcTaxDropdown.Top - lbcTax.Height
            lbcTax.Move edcTaxDropdown.Left, edcTaxDropdown.Top - lbcTax.Height
            edcTaxDropdown.SelStart = 0
            edcTaxDropdown.SelLength = Len(edcTaxDropdown.Text)
            edcTaxDropdown.Visible = True
            cmcDropDown.Visible = True
            edcTaxDropdown.SetFocus
        Case HUBINDEX 'Hub
            mHubPop
            If imTerminate Then
                Exit Sub
            End If
            lbcHub.Height = gListBoxHeight(lbcHub.ListCount, 8)
            edcHubDropdown.Width = tmCtrls(HUBINDEX).fBoxW - cmcDropDown.Width
            edcHubDropdown.MaxLength = 20
            gMoveFormCtrl pbcVeh, edcHubDropdown, tmCtrls(HUBINDEX).fBoxX, tmCtrls(HUBINDEX).fBoxY
            cmcDropDown.Move edcHubDropdown.Left + edcHubDropdown.Width, edcHubDropdown.Top
            imHubChgMode = True
            slStr = edcHubDropdown.Text
            gFindMatch slStr, 0, lbcHub
            If gLastFound(lbcHub) >= 0 Then
                lbcHub.ListIndex = gLastFound(lbcHub)
                edcHubDropdown.Text = lbcHub.List(lbcHub.ListIndex)
            Else
                '12/27/18: selected [None]
                lbcHub.ListIndex = 1    '0
                edcHubDropdown.Text = lbcHub.List(1)    '0)
            End If
            imComboBoxIndex = lbcHub.ListIndex
            imHubChgMode = False
            lbcHub.Move edcHubDropdown.Left, edcHubDropdown.Top - lbcHub.Height
            edcHubDropdown.SelStart = 0
            edcHubDropdown.SelLength = Len(edcHubDropdown.Text)
            edcHubDropdown.Visible = True
            cmcDropDown.Visible = True
            edcHubDropdown.SetFocus
        Case SORTINDEX 'Station Code
            edcSort.Width = tmCtrls(ilBoxNo).fBoxW
            edcSort.MaxLength = 4
            gMoveFormCtrl pbcVeh, edcSort, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcSort.Visible = True  'Set visibility
            edcSort.SetFocus
        Case SCODEINDEX 'Station Code
            edcStationCode.Width = tmCtrls(ilBoxNo).fBoxW
            edcStationCode.MaxLength = 5
            gMoveFormCtrl pbcVeh, edcStationCode, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcStationCode.Visible = True  'Set visibility
            edcStationCode.SetFocus
        Case BOOKINDEX
            If lbcBook.ListCount <= 2 Then
                mBookPop False
            Else
                mBookPop True
            End If
            If imTerminate Then
                Exit Sub
            End If
            lbcBook.Height = gListBoxHeight(lbcBook.ListCount, 8)
            edcBookDropDown.Width = tmCtrls(BOOKINDEX).fBoxW - cmcDropDown.Width
            edcBookDropDown.MaxLength = 40  'Name plus date
            gMoveFormCtrl pbcVeh, edcBookDropDown, tmCtrls(BOOKINDEX).fBoxX, tmCtrls(BOOKINDEX).fBoxY
            cmcDropDown.Move edcBookDropDown.Left + edcBookDropDown.Width, edcBookDropDown.Top
            imBookChgMode = True
            slStr = edcBookDropDown.Text
            gFindMatch slStr, 0, lbcBook
            If gLastFound(lbcBook) >= 0 Then
                lbcBook.ListIndex = gLastFound(lbcBook)
                edcBookDropDown.Text = lbcBook.List(lbcBook.ListIndex)
            Else
                lbcBook.ListIndex = 0
                edcBookDropDown.Text = lbcBook.List(0)
            End If
            imComboBoxIndex = lbcBook.ListIndex
            imBookChgMode = False
            lbcBook.Move edcBookDropDown.Left, edcBookDropDown.Top - lbcBook.Height
            edcBookDropDown.SelStart = 0
            edcBookDropDown.SelLength = Len(edcBookDropDown.Text)
            edcBookDropDown.Visible = True
            cmcDropDown.Visible = True
            edcBookDropDown.SetFocus
        'Case REALLINDEX
        '    If lbcBook.ListCount <= 2 Then
        '        mBookPop False
        '    Else
        '        mBookPop True
        '    End If
        '    If imTerminate Then
        '        Exit Sub
        '    End If
        '    lbcBook.Height = gListBoxHeight(lbcBook.ListCount, 8)
        '    edcReallDropDown.Width = tmCtrls(REALLINDEX).fBoxW - cmcDropDown.Width
        '    edcReallDropDown.MaxLength = 40  'Name plus date
        '    gMoveFormCtrl pbcVeh, edcReallDropDown, tmCtrls(REALLINDEX).fBoxX, tmCtrls(REALLINDEX).fBoxY
        '    cmcDropDown.Move edcReallDropDown.Left + edcReallDropDown.Width, edcReallDropDown.Top
        '    imBookChgMode = True
        '    slStr = edcReallDropDown.Text
        '    gFindMatch slStr, 0, lbcBook
        '    If gLastFound(lbcBook) >= 0 Then
        '        lbcBook.ListIndex = gLastFound(lbcBook)
        '        edcReallDropDown.Text = lbcBook.List(lbcBook.ListIndex)
        '    Else
        '        lbcBook.ListIndex = 0
        '        edcReallDropDown.Text = lbcBook.List(0)
        '    End If
        '    imComboBoxIndex = lbcBook.ListIndex
        '    imBookChgMode = False
        '    lbcBook.Move edcReallDropDown.Left, edcReallDropDown.Top - lbcBook.Height
        '    edcReallDropDown.SelStart = 0
        '    edcReallDropDown.SelLength = Len(edcReallDropDown.Text)
        '    edcReallDropDown.Visible = True
        '    cmcDropDown.Visible = True
        '    edcReallDropDown.SetFocus
        Case DEMOINDEX
            mDemoPop
            If imTerminate Then
                Exit Sub
            End If
            lbcDemo.Height = gListBoxHeight(lbcDemo.ListCount, 8)
            edcDemoDropDown.Width = tmCtrls(DEMOINDEX).fBoxW - cmcDropDown.Width
            edcDemoDropDown.MaxLength = 20
            gMoveFormCtrl pbcVeh, edcDemoDropDown, tmCtrls(DEMOINDEX).fBoxX, tmCtrls(DEMOINDEX).fBoxY
            cmcDropDown.Move edcDemoDropDown.Left + edcDemoDropDown.Width, edcDemoDropDown.Top
            imDemoChgMode = True
            slStr = edcDemoDropDown.Text
            gFindMatch slStr, 0, lbcDemo
            If gLastFound(lbcDemo) >= 0 Then
                lbcDemo.ListIndex = gLastFound(lbcDemo)
                edcDemoDropDown.Text = lbcDemo.List(lbcDemo.ListIndex)
            Else
                lbcDemo.ListIndex = 1
                edcDemoDropDown.Text = lbcDemo.List(1)
            End If
            imDemoChgMode = False
            lbcDemo.Move edcDemoDropDown.Left, edcDemoDropDown.Top - lbcDemo.Height
            edcDemoDropDown.SelStart = 0
            edcDemoDropDown.SelLength = Len(edcDemoDropDown.Text)
            edcDemoDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDemoDropDown.SetFocus
        Case TYPEINDEX
            lbcType.Height = gListBoxHeight(lbcType.ListCount, 8)
            edcTypeDropDown.Width = tmCtrls(TYPEINDEX).fBoxW - cmcDropDown.Width
            edcTypeDropDown.MaxLength = 20
            gMoveFormCtrl pbcVeh, edcTypeDropDown, tmCtrls(TYPEINDEX).fBoxX, tmCtrls(TYPEINDEX).fBoxY
            cmcDropDown.Move edcTypeDropDown.Left + edcTypeDropDown.Width, edcTypeDropDown.Top
            imTypeChgMode = True
            slStr = edcTypeDropDown.Text
            gFindMatch slStr, 0, lbcType
            If gLastFound(lbcType) >= 0 Then
                lbcType.ListIndex = gLastFound(lbcType)
                edcTypeDropDown.Text = lbcType.List(lbcType.ListIndex)
            Else
                'If tgSpf.sSSellNet = "Y" Then
                '    slStr = "Selling"
                'Else
                    slStr = "Conventional"
                'End If
                gFindMatch slStr, 0, lbcType
                If gLastFound(lbcType) >= 0 Then
                    lbcType.ListIndex = gLastFound(lbcType)
                    edcTypeDropDown.Text = lbcType.List(lbcType.ListIndex)
                Else
                    lbcType.ListIndex = 0
                    edcTypeDropDown.Text = lbcType.List(0)
                End If
            End If
            imComboBoxIndex = lbcType.ListIndex
            imTypeChgMode = False
            lbcType.Move edcTypeDropDown.Left, edcTypeDropDown.Top + edcTypeDropDown.Height
            edcTypeDropDown.SelStart = 0
            edcTypeDropDown.SelLength = Len(edcTypeDropDown.Text)
            edcTypeDropDown.Visible = True
            cmcDropDown.Visible = True
            edcTypeDropDown.SetFocus
        Case STATEINDEX 'Selling or Airing
            If imState < 0 Then
                imState = 0
                tmCtrls(ilBoxNo).iChg = True
                mSetCommands
            End If
            pbcState.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcVeh, pbcState, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcState_Paint
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInitParameters
'   Where:
'
    Dim ilRet As Integer
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imVefChanged = False
    mParseCmmdLine
    '10071
    mPrepVehicle
    'this is after mPrepVehicle!
    If imTerminate Then
        Exit Sub
    End If
    mSetAllowUpdate
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
        If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
            ilRet = gPopTaxRateBox(False, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
            imTaxDefined = True
        ElseIf (Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA Then
            ilRet = gPopTaxRateBox(True, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
            imTaxDefined = True
        ElseIf (Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR Then
            ilRet = gPopTaxRateBox(False, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
            imTaxDefined = True
        Else
            imTaxDefined = False
        End If
    Else
        imTaxDefined = False
    End If
    imACT1CodesDefined = False
    If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then
        imACT1CodesDefined = True
    End If
    mInitBox
    '10071
'    Vehicle.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    Vehicle.Height = cmcOptions.Top + 5 * cmcOptions.Height / 3
    gCenterStdAlone Vehicle
    'Vehicle.Show
    Screen.MousePointer = vbHourglass

    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    imVefRecLen = Len(tmVef)  'Get and save ARF record length
    imBoxNo = -1 'Initialize current Box to N/A
    imInNew = False
    imChgMode = False
    imLVChgMode = False
    imMVLChgMode = False
    imBookChgMode = False
    imDemoChgMode = False
    imFirstTimeSelect = True
    imVehGp2ChgMode = False
    imVehGp3ChgMode = False
    imVehGp4ChgMode = False
    imVehGp5ChgMode = False
    imRNLinkChgMode = False
    imBSMode = False
    imPopReqd = False
    imBypassSetting = False
    imDoubleClickName = False
    imDirProcess = -1
    imSvSelectedIndex = -1
    igVefCodeModel = 0
    imFromOldSave = False
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imLbcMouseDown = False
    bmPrgLibDefined = False
    smPhoneImage = mkcPhone.Text
    smFaxImage = mkcFax.Text
    If ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        sgRNCallType = "R"
    ElseIf ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
        sgRNCallType = "N"
    Else
        sgRNCallType = ""
    End If
    hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Vehicle
    On Error GoTo 0
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Vehicle
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmVsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Vehicle
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Vehicle
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Vehicle
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmVaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVaf, "", sgDBPath & "Vaf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Vehicle
    On Error GoTo 0
    imVafRecLen = Len(tmVaf)
    hmVbf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVbf, "", sgDBPath & "Vbf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Vehicle
    On Error GoTo 0
    imVBfRecLen = Len(tmVbf)
    hmVff = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Vehicle
    On Error GoTo 0
    imVffRecLen = Len(tmVff)
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:LCF)", Vehicle
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)
    hmNif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmNif, "", sgDBPath & "Nif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen:NIF)", Vehicle
    On Error GoTo 0
    imNifRecLen = Len(tmNif)
'    Vehicle.Height = cmcReport.Top + 5 * cmcReport.Height / 3
'    gCenterModalForm Vehicle
'    Traffic!plcHelp.Caption = ""
    mInitBox
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0
        mSetCommands
    End If
    gObtainVefIgnoreHub tmMVef()
    If Not imTerminate Then
        lbcLogVeh.Clear
        'mLogVehPop
    End If
    If Not imTerminate Then
        lbcMultiConvVehLog.Clear
        'mLogVehPop
    End If
    If Not imTerminate Then
        lbcMultiGameVehLog.Clear
        'mLogVehPop
    End If
    If Not imTerminate Then
        lbcVehGp2.Clear
        mVehGp2Pop
    End If
    If Not imTerminate Then
        lbcVehGp3.Clear
        mVehGp3Pop
    End If
    If Not imTerminate Then
        lbcVehGp4.Clear
        mVehGp4Pop
    End If
    If Not imTerminate Then
        lbcVehGp5.Clear
        mVehGp5Pop
    End If
    If Not imTerminate Then
        lbcVehGp6.Clear
        mVehGp6Pop
    End If
    If Not imTerminate Then
        lbcBook.Clear
        mBookPop False
    End If
    If Not imTerminate Then
        lbcDemo.Clear
        mDemoPop
    End If
    If Not imTerminate Then
        lbcHub.Clear
        mHubPop
    End If
    If Not imTerminate Then
        lbcRNLink.Clear
        mRNLinkPop
    End If
    'If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
    '    If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
    '        ilRet = gPopTaxRateBox(False, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
    '        imTaxDefined = True
    '    ElseIf (Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA Then
    '        ilRet = gPopTaxRateBox(True, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
    '        imTaxDefined = True
    '    ElseIf (Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR Then
    '        ilRet = gPopTaxRateBox(False, lbcTax, tmTaxSortCode(), smTaxSortCodeTag)
    '        imTaxDefined = True
    '    Else
    '        imTaxDefined = False
    '    End If
    'Else
    '    imTaxDefined = False
    'End If
    If Not imTaxDefined Then
        ReDim tmTaxSortCode(0 To 0) As SORTCODE
    End If
    lbcTax.AddItem "[None]", 0
    If tgSpf.sSSellNet = "Y" Then
        imMaxCtrlNo = UBound(tmCtrls)
    Else
        imMaxCtrlNo = UBound(tmCtrls)
    End If
    Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    flTextHeight = pbcVeh.TextHeight("1") - 35
    If (tgSpf.sCPkOrdered <> "Y") And (tgSpf.sCPkAired <> "Y") Then
        rbcType(1).Enabled = False
        plcSelect.Move 1005, 30
    Else
        plcSelect.Move 1005, 30
    End If
    'Position panel and picture areas with panel
    plcVeh.Move 135, 600, pbcVeh.Width + fgPanelAdj, pbcVeh.Height + fgPanelAdj
    pbcVeh.Move plcVeh.Left + fgBevelX, plcVeh.Top + fgBevelY
    'Name
    gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
    'Contact
    gSetCtrl tmCtrls(CONTACTINDEX), 2850, 30, 1770, fgBoxStH
    tmCtrls(CONTACTINDEX).iReq = False
    'Phone
    gSetCtrl tmCtrls(PHONEINDEX), 4635, 30, 1935, fgBoxStH
    tmCtrls(PHONEINDEX).iReq = False
    'Fax
    gSetCtrl tmCtrls(FAXINDEX), 6585, 30, 1260, fgBoxStH
    tmCtrls(FAXINDEX).iReq = False
    'Address
    gSetCtrl tmCtrls(ADDRESSINDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 7815, fgBoxStH  'fgBoxAddH
    tmCtrls(ADDRESSINDEX).iReq = False
    gSetCtrl tmCtrls(ADDRESSINDEX + 1), 30, tmCtrls(ADDRESSINDEX).fBoxY + flTextHeight, tmCtrls(ADDRESSINDEX).fBoxW, flTextHeight
    tmCtrls(ADDRESSINDEX + 1).iReq = False
    gSetCtrl tmCtrls(ADDRESSINDEX + 2), 30, tmCtrls(ADDRESSINDEX + 1).fBoxY + flTextHeight, tmCtrls(ADDRESSINDEX).fBoxW, flTextHeight
    tmCtrls(ADDRESSINDEX + 2).iReq = False
    gSetCtrl tmCtrls(ADDRESSINDEX + 3), 30, tmCtrls(ADDRESSINDEX + 2).fBoxY + flTextHeight, tmCtrls(ADDRESSINDEX).fBoxW, flTextHeight
    tmCtrls(ADDRESSINDEX + 3).iReq = False
    'Type: Selling or Airing or conventional or Log or Virtual Or Package
    'gSetCtrl tmCtrls(TYPEINDEX), 30, tmCtrls(ADDRESSINDEX).fBoxY + fgAddDeltaY, 1515, fgBoxStH
    'gSetCtrl tmCtrls(TYPEINDEX), 30, tmCtrls(ADDRESSINDEX).fBoxY + 960, 1515, fgBoxStH
    gSetCtrl tmCtrls(TYPEINDEX), 30, tmCtrls(ADDRESSINDEX).fBoxY + 945, 1515, fgBoxStH
    'Log vehicle code
    gSetCtrl tmCtrls(LOGVEHINDEX), 1560, tmCtrls(TYPEINDEX).fBoxY, 1275, fgBoxStH
    tmCtrls(LOGVEHINDEX).iReq = False
    'Multi vehicle log code
    gSetCtrl tmCtrls(MULTIVEHLOGINDEX), 2850, tmCtrls(TYPEINDEX).fBoxY, 1725, fgBoxStH
    tmCtrls(MULTIVEHLOGINDEX).iReq = False
    'Rep-Net Link
    gSetCtrl tmCtrls(RNLINKINDEX), 4590, tmCtrls(TYPEINDEX).fBoxY, 1470, fgBoxStH
    tmCtrls(RNLINKINDEX).iReq = False
    'Dial Position
    gSetCtrl tmCtrls(DIALPOSINDEX), 6075, tmCtrls(TYPEINDEX).fBoxY, 855, fgBoxStH
    tmCtrls(DIALPOSINDEX).iReq = False
    'Station code
    gSetCtrl tmCtrls(SCODEINDEX), 6945, tmCtrls(TYPEINDEX).fBoxY, 900, fgBoxStH
    tmCtrls(SCODEINDEX).iReq = False
    'Market
    gSetCtrl tmCtrls(MKTNAMEINDEX), 30, tmCtrls(TYPEINDEX).fBoxY + fgStDeltaY, 1515, fgBoxStH
    If tgSpf.sMktBase <> "Y" Then
        tmCtrls(MKTNAMEINDEX).iReq = False
    End If
    'Research
    gSetCtrl tmCtrls(RSCHINDEX), 1560, tmCtrls(MKTNAMEINDEX).fBoxY, 1275, fgBoxStH
    tmCtrls(RSCHINDEX).iReq = False
    'Sub-Company
    If tgSpf.sSubCompany = "Y" Then
        gSetCtrl tmCtrls(SUBCOMPINDEX), 2850, tmCtrls(MKTNAMEINDEX).fBoxY, 1725, fgBoxStH
        tmCtrls(SUBCOMPINDEX).iReq = True
    Else
        'gSetCtrl tmCtrls(SUBCOMPINDEX), 0, tmCtrls(MKTNAMEINDEX).fBoxY, 0, fgBoxStH
        gSetCtrl tmCtrls(SUBCOMPINDEX), 2850, tmCtrls(MKTNAMEINDEX).fBoxY, 1725, fgBoxStH
        tmCtrls(SUBCOMPINDEX).iReq = False
    End If
    'Format
    gSetCtrl tmCtrls(FORMATINDEX), 4590, tmCtrls(MKTNAMEINDEX).fBoxY, 1620, fgBoxStH
    tmCtrls(FORMATINDEX).iReq = False
    'Subtotal
    gSetCtrl tmCtrls(SUBTOTALINDEX), 6225, tmCtrls(MKTNAMEINDEX).fBoxY, 1620, fgBoxStH
    tmCtrls(SUBTOTALINDEX).iReq = False
    'Book Name
    'If tgSpf.sCAudPkg <> "Y" Then
    '    gSetCtrl tmCtrls(BOOKINDEX), 30, tmCtrls(MKTNAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
    '    tmCtrls(BOOKINDEX).iReq = False
    '    gSetCtrl tmCtrls(REALLINDEX), 0, tmCtrls(MKTNAMEINDEX).fBoxY + fgStDeltaY, 0, fgBoxStH
    '    tmCtrls(REALLINDEX).iReq = False
    'Else
    '    gSetCtrl tmCtrls(BOOKINDEX), 30, tmCtrls(MKTNAMEINDEX).fBoxY + fgStDeltaY, 1515, fgBoxStH
    '    tmCtrls(BOOKINDEX).iReq = False
    '    gSetCtrl tmCtrls(REALLINDEX), 1560, tmCtrls(MKTNAMEINDEX).fBoxY + fgStDeltaY, 1275, fgBoxStH
    '    tmCtrls(REALLINDEX).iReq = False
    'End If
    gSetCtrl tmCtrls(BOOKINDEX), 30, tmCtrls(MKTNAMEINDEX).fBoxY + fgStDeltaY, 1515, fgBoxStH
    tmCtrls(BOOKINDEX).iReq = False
    gSetCtrl tmCtrls(DEMOINDEX), 1560, tmCtrls(BOOKINDEX).fBoxY, 1275, fgBoxStH
    tmCtrls(DEMOINDEX).iReq = False
    If imACT1CodesDefined And imTaxDefined Then
        gSetCtrl tmCtrls(ACT1CODESINDEX), 2850, tmCtrls(BOOKINDEX).fBoxY, 1275, fgBoxStH
        tmCtrls(ACT1CODESINDEX).iReq = False
        gSetCtrl tmCtrls(TAXINDEX), 4140, tmCtrls(BOOKINDEX).fBoxY, 435, fgBoxStH
        tmCtrls(TAXINDEX).iReq = True
    ElseIf imACT1CodesDefined Then
        'ACT1 Codes
        gSetCtrl tmCtrls(ACT1CODESINDEX), 2850, tmCtrls(BOOKINDEX).fBoxY, 1725, fgBoxStH
        tmCtrls(ACT1CODESINDEX).iReq = False
        gSetCtrl tmCtrls(TAXINDEX), 0, 0, 0, 0
        tmCtrls(TAXINDEX).iReq = False
    ElseIf imTaxDefined Then
        'Tax
        gSetCtrl tmCtrls(ACT1CODESINDEX), 0, 0, 0, 0
        tmCtrls(ACT1CODESINDEX).iReq = False
        gSetCtrl tmCtrls(TAXINDEX), 2850, tmCtrls(BOOKINDEX).fBoxY, 1725, fgBoxStH
        tmCtrls(TAXINDEX).iReq = True
    Else
        gSetCtrl tmCtrls(ACT1CODESINDEX), 0, 0, 0, 0
        tmCtrls(ACT1CODESINDEX).iReq = False
        gSetCtrl tmCtrls(TAXINDEX), 0, 0, 0, 0
        tmCtrls(TAXINDEX).iReq = False
    End If
    'Sort
    gSetCtrl tmCtrls(SORTINDEX), 4590, tmCtrls(BOOKINDEX).fBoxY, 855, fgBoxStH
    tmCtrls(SORTINDEX).iReq = False
    'Hub
    gSetCtrl tmCtrls(HUBINDEX), 5460, tmCtrls(BOOKINDEX).fBoxY, 750, fgBoxStH
    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
        tmCtrls(HUBINDEX).iReq = True
    Else
        tmCtrls(HUBINDEX).iReq = False
    End If
    'State
    gSetCtrl tmCtrls(STATEINDEX), 6225, tmCtrls(BOOKINDEX).fBoxY, 1620, fgBoxStH

End Sub
Private Sub mkcFax_Change()
    mSetChg imBoxNo
End Sub
Private Sub mkcFax_GotFocus()
    gCtrlGotFocus mkcFax
End Sub
Private Sub mkcPhone_Change()
    mSetChg imBoxNo
End Sub
Private Sub mkcPhone_GotFocus()
    gCtrlGotFocus mkcPhone
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mLogVehPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer

'    If (tgSpf.sSSellNet = "Y") Or (tgSpf.sSDelNet = "Y") Then
'        ilFilter(0) = NOFILTER
'        slFilter(0) = ""
'        ilOffset(0) = 0
'    Else
    If edcTypeDropDown.Text = "Simulcast" Then
        ilfilter(0) = CHARFILTER
        slFilter(0) = "C"
        ilOffset(0) = gFieldOffset("Vef", "VefType")
    Else
        ilfilter(0) = CHARFILTER
        slFilter(0) = "L"
        ilOffset(0) = gFieldOffset("Vef", "VefType")
    End If
'    End If
    ilOffset(0) = gFieldOffset("Vef", "VefType") '167
    'ilRet = gIMoveListBox(Vehicle, lbcLogVeh, lbcLogVehCode, "Vef.btr", gFieldOffset("Vef", "VefName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Vehicle, lbcLogVeh, tmLogVehCode(), smLogVehCodeTag, "Vef.btr", gFieldOffset("Vef", "VefName"), 20, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLogVehPopErr
        gCPErrorMsg ilRet, "mLogVehPop (gIMoveListBox)", Vehicle
        On Error GoTo 0
        lbcLogVeh.AddItem "[None]", 0  'Force as first item on list
    End If
    Exit Sub
mLogVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                      and set defaults               *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                       ilIndexS                      ilIndexV                  *
'*  ilIndexP                      ilIndexIU                     ilIndexEU                 *
'*  ilAllZero                     ilSSourceLastFd                                         *
'******************************************************************************************

'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slStr As String
    Dim slCode As String
    Dim ilRet As Integer

    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        tmVef.sName = edcName.Text
    End If
    If Not ilTestChg Or tmCtrls(CONTACTINDEX).iChg Then
        tmVef.sContact = edcContact.Text
    End If
    'If Not ilTestChg Or tmCtrls(MKTNAMEINDEX).iChg Then
    '    tmVef.sMktName = edcMktName.Text
    'End If
    imVehGp3ChgMode = True
    If Not ilTestChg Or tmCtrls(MKTNAMEINDEX).iChg Then
        tmVef.iMnfVehGp3Mkt = 0
        slStr = smVehGp3
        gFindMatch slStr, 2, lbcVehGp3
        If gLastFound(lbcVehGp3) > 1 Then
            slNameCode = tmVehGp3Code(gLastFound(lbcVehGp3) - 2).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmVef.iMnfVehGp3Mkt = Val(slCode)
        End If
    End If
    imVehGp3ChgMode = False
    imVehGp5ChgMode = True
    If Not ilTestChg Or tmCtrls(RSCHINDEX).iChg Then
        tmVef.iMnfVehGp5Rsch = 0
        slStr = smVehGp5
        gFindMatch slStr, 2, lbcVehGp5
        If gLastFound(lbcVehGp5) > 1 Then
            slNameCode = tmVehGp5Code(gLastFound(lbcVehGp5) - 2).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmVef.iMnfVehGp5Rsch = Val(slCode)
        End If
    End If
    imVehGp5ChgMode = False
    imVehGp6ChgMode = True
    If Not ilTestChg Or tmCtrls(SUBCOMPINDEX).iChg Then
        tmVef.iMnfVehGp6Sub = 0
        slStr = smVehGp6
'        gFindMatch slStr, 1, lbcVehGp6
'        If gLastFound(lbcVehGp6) >= 1 Then
'            slNameCode = tmVehGp6Code(gLastFound(lbcVehGp6) - 1).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
'            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            tmVef.iMnfVehGp6Sub = Val(slCode)
'        End If
        If tgSpf.sSubCompany <> "Y" Then
            gFindMatch slStr, 2, lbcVehGp6
            If gLastFound(lbcVehGp6) >= 2 Then
                slNameCode = tmVehGp6Code(gLastFound(lbcVehGp6) - 2).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmVef.iMnfVehGp6Sub = Val(slCode)
            End If
        Else
            gFindMatch slStr, 1, lbcVehGp6
            If gLastFound(lbcVehGp6) >= 1 Then
                slNameCode = tmVehGp6Code(gLastFound(lbcVehGp6) - 1).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmVef.iMnfVehGp6Sub = Val(slCode)
            End If
        End If
    End If
    imVehGp6ChgMode = False
    For ilLoop = 0 To 3 Step 1
        If Not ilTestChg Or tmCtrls(ADDRESSINDEX + ilLoop).iChg Then
            If ilLoop <= 2 Then
                tmVef.sAddr(ilLoop) = edcAddr(ilLoop).Text
            Else
                tmVff.sAddr4 = edcAddr(ilLoop).Text
            End If
        End If
    Next ilLoop
    If Not ilTestChg Or tmCtrls(PHONEINDEX).iChg Then
        gGetPhoneNo mkcPhone, tmVef.sPhone
    End If
    If Not ilTestChg Or tmCtrls(FAXINDEX).iChg Then
        gGetPhoneNo mkcFax, tmVef.sFax
    End If
    'If Not ilTestChg Or tmCtrls(FORMATINDEX).iChg Then
    '    tmVef.sFormat = edcFormat.Text
    'End If
    imVehGp4ChgMode = True
    If Not ilTestChg Or tmCtrls(FORMATINDEX).iChg Then
        tmVef.iMnfVehGp4Fmt = 0
        slStr = smVehGp4
        gFindMatch slStr, 2, lbcVehGp4
        If gLastFound(lbcVehGp4) > 1 Then
            slNameCode = tmVehGp4Code(gLastFound(lbcVehGp4) - 2).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmVef.iMnfVehGp4Fmt = Val(slCode)
        End If
    End If
    imVehGp4ChgMode = False
    imVehGp2ChgMode = True
    If Not ilTestChg Or tmCtrls(SUBTOTALINDEX).iChg Then
        tmVef.iMnfVehGp2 = 0
        slStr = smVehGp2
        gFindMatch slStr, 2, lbcVehGp2
        If gLastFound(lbcVehGp2) > 1 Then
            slNameCode = tmVehGp2Code(gLastFound(lbcVehGp2) - 2).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmVef.iMnfVehGp2 = Val(slCode)
        End If
    End If
    imVehGp2ChgMode = False
    If Not ilTestChg Or tmCtrls(DIALPOSINDEX).iChg Then
        tmVef.sDialPos = edcDialPos.Text
    End If
    If imTaxDefined Then
        If Not ilTestChg Or tmCtrls(TAXINDEX).iChg Then
            tmVef.iTrfCode = 0
            imTaxChgMode = True
            slStr = edcTaxDropdown.Text
            gFindMatch slStr, 0, lbcTax
            If gLastFound(lbcTax) > 0 Then
                slNameCode = tmTaxSortCode(gLastFound(lbcTax) - 1).sKey    'lbcLogVehCode.List(gLastFound(lbcLogVeh) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmVef.iTrfCode = Val(slCode)
            End If
            imTaxChgMode = False
        End If
    End If
    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
        If Not ilTestChg Or tmCtrls(HUBINDEX).iChg Then
            tmVef.iMnfHubCode = 0
            imHubChgMode = True
            slStr = edcHubDropdown.Text
            gFindMatch slStr, 0, lbcHub
            '12/27/18
            If gLastFound(lbcHub) > 1 Then  '0 Then
                slNameCode = tmHubCode(gLastFound(lbcHub) - 2).sKey '- 1).sKey    'lbcLogVehCode.List(gLastFound(lbcLogVeh) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmVef.iMnfHubCode = Val(slCode)
            End If
            imHubChgMode = False
        End If
    End If
    If imACT1CodesDefined Then
        If Not ilTestChg Or tmCtrls(ACT1CODESINDEX).iChg Then
            tmVff.sACT1LineupCode = edcACT1Lineup.Text
        End If
    End If
    If Not ilTestChg Or tmCtrls(TYPEINDEX).iChg Then
        'tmVef.sUnused3 = ""
        imTypeChgMode = True
        slStr = edcTypeDropDown.Text
        gFindMatch slStr, 0, lbcType
        If gLastFound(lbcType) >= 0 Then
            slStr = lbcType.List(gLastFound(lbcType))
            If StrComp(slStr, "Conventional", 1) = 0 Then
                tmVef.sType = "C"
            ElseIf StrComp(slStr, "Selling", 1) = 0 Then
                tmVef.sType = "S"
            ElseIf StrComp(slStr, "Airing", 1) = 0 Then
                tmVef.sType = "A"
            ElseIf StrComp(slStr, "Log", 1) = 0 Then
                tmVef.sType = "L"
            ElseIf StrComp(slStr, "Virtual", 1) = 0 Then
                tmVef.sType = "V"
            ElseIf StrComp(slStr, "Simulcast", 1) = 0 Then
                tmVef.sType = "T"
            ElseIf StrComp(slStr, "Package", 1) = 0 Then
                tmVef.sType = "P"
            ElseIf StrComp(slStr, "Rep", 1) = 0 Then
                tmVef.sType = "R"
            ElseIf StrComp(slStr, "Sport", 1) = 0 Then
                tmVef.sType = "G"
            ElseIf StrComp(slStr, "NTR", 1) = 0 Then
                tmVef.sType = "N"
            End If
        End If
        imTypeChgMode = False
    End If

    If Not ilTestChg Or tmCtrls(LOGVEHINDEX).iChg Then
        tmVef.iVefCode = 0
        imLVChgMode = True
        slStr = edcLogVehDropDown.Text
        gFindMatch slStr, 0, lbcLogVeh
        If gLastFound(lbcLogVeh) > 0 Then
            slNameCode = tmLogVehCode(gLastFound(lbcLogVeh) - 1).sKey    'lbcLogVehCode.List(gLastFound(lbcLogVeh) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmVef.iVefCode = Val(slCode)
        End If
        imLVChgMode = False
    End If
    If Not ilTestChg Or tmCtrls(MULTIVEHLOGINDEX).iChg Then
        tmVef.iCombineVefCode = 0
        imMVLChgMode = True
        slStr = edcMultiVehDropdown.Text
        If (smVehicleType = "C") Or (smVehicleType = "A") Then
            gFindMatch slStr, 0, lbcMultiConvVehLog
            If gLastFound(lbcMultiConvVehLog) > 0 Then
                slNameCode = tmMultiConvVehLogCode(gLastFound(lbcMultiConvVehLog) - 1).sKey    'lbcLogVehCode.List(gLastFound(lbcLogVeh) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmVef.iCombineVefCode = Val(slCode)
            End If
        ElseIf smVehicleType = "G" Then
            gFindMatch slStr, 0, lbcMultiGameVehLog
            If gLastFound(lbcMultiGameVehLog) > 0 Then
                slNameCode = tmMultiGameVehLogCode(gLastFound(lbcMultiGameVehLog) - 1).sKey    'lbcLogVehCode.List(gLastFound(lbcLogVeh) - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmVef.iCombineVefCode = Val(slCode)
            End If
        End If
        imMVLChgMode = False
    End If

    If Not ilTestChg Or tmCtrls(RNLINKINDEX).iChg Then
        tmVef.iNrfCode = 0
        imRNLinkChgMode = True
        slStr = edcRNLinkDropdown.Text
        gFindMatch slStr, 2, lbcRNLink
        If gLastFound(lbcRNLink) > 1 Then
            slNameCode = tmRNLinkCode(gLastFound(lbcRNLink) - 2).sKey    'lbcLogVehCode.List(gLastFound(lbcLogVeh) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmVef.iNrfCode = Val(slCode)
        End If
        imRNLinkChgMode = False
    End If

    If Not ilTestChg Or tmCtrls(SCODEINDEX).iChg Then
        tmVef.sCodeStn = edcStationCode.Text
    End If
    If Not ilTestChg Or tmCtrls(SORTINDEX).iChg Then
        tmVef.iSort = Val(edcSort.Text)
    End If
    If Not ilTestChg Or tmCtrls(BOOKINDEX).iChg Then
        tmVef.iDnfCode = 0
        imBookChgMode = True
        slStr = edcBookDropDown.Text
        gFindMatch slStr, 0, lbcBook
        If gLastFound(lbcBook) > 0 Then
            slNameCode = tmBookCode(gLastFound(lbcBook) - 1).sKey    'lbcBookCode.List(gLastFound(lbcBook) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmVef.iDnfCode = Val(slCode)
        End If
        imBookChgMode = False
    End If
    'If tgSpf.sCAudPkg = "Y" Then
    '    If Not ilTestChg Or tmCtrls(REALLINDEX).iChg Then
    '        tmVef.iReallDnfCode = 0
    '        imBookChgMode = True
    '        slStr = edcReallDropDown.Text
    '        gFindMatch slStr, 0, lbcBook
    '        If gLastFound(lbcBook) > 0 Then
    '            slNameCode = tmBookCode(gLastFound(lbcBook) - 1).sKey    'lbcBookCode.List(gLastFound(lbcBook) - 1)
    '            ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '            tmVef.iReallDnfCode = Val(slCode)
    '        End If
    '        imBookChgMode = False
    '    End If
    'Else
        tmVef.iReallDnfCode = 0
    'End If
    If Not ilTestChg Or tmCtrls(DEMOINDEX).iChg Then
        tmVef.iMnfDemo = 0
        imDemoChgMode = True
        slStr = edcDemoDropDown.Text
        If (slStr = "[None]") Or (slStr = "") Then
            tmVef.iMnfDemo = 0
        Else
            gFindMatch slStr, 0, lbcDemo
            If gLastFound(lbcDemo) > 0 Then
                slNameCode = tmDemoCode(gLastFound(lbcDemo) - 2).sKey    'lbcDemoCode.List(gLastFound(lbcDemo) - 2)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmVef.iMnfDemo = Val(slCode)
            End If
        End If
        imDemoChgMode = False
    End If
    If Not ilTestChg Or tmCtrls(STATEINDEX).iChg Then
        Select Case imState
            Case 0  'Active
                tmVef.sState = "A"
            Case 1  'Dormant
                tmVef.sState = "D"
        End Select
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    edcName.Text = Trim$(tmVef.sName)
    smOrigName = edcName.Text
    edcContact.Text = Trim$(tmVef.sContact)
    mVehGp3Pop
    smOrigVehGp3 = ""
    smVehGp3 = ""
    If tmVef.iMnfVehGp3Mkt > 0 Then
        For ilVef = 0 To UBound(tmVehGp3Code) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
            slNameCode = tmVehGp3Code(ilVef).sKey   'lbcVehGpCode.List(ilVef)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iMnfVehGp3Mkt Then
                ilRet = gParseItem(slNameCode, 1, "\", smVehGp3)
                smOrigVehGp3 = smVehGp3
                Exit For
            End If
        Next ilVef
    End If
    mVehGp5Pop
    smOrigVehGp5 = ""
    smVehGp5 = ""
    If tmVef.iMnfVehGp5Rsch > 0 Then
        For ilVef = 0 To UBound(tmVehGp5Code) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
            slNameCode = tmVehGp5Code(ilVef).sKey   'lbcVehGpCode.List(ilVef)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iMnfVehGp5Rsch Then
                ilRet = gParseItem(slNameCode, 1, "\", smVehGp5)
                smOrigVehGp5 = smVehGp5
                Exit For
            End If
        Next ilVef
    End If
    mVehGp6Pop
    smOrigVehGp6 = ""
    smVehGp6 = ""
    If tmVef.iMnfVehGp6Sub > 0 Then
        For ilVef = 0 To UBound(tmVehGp6Code) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
            slNameCode = tmVehGp6Code(ilVef).sKey   'lbcVehGpCode.List(ilVef)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iMnfVehGp6Sub Then
                ilRet = gParseItem(slNameCode, 1, "\", smVehGp6)
                smOrigVehGp6 = smVehGp6
                Exit For
            End If
        Next ilVef
    End If
    For ilLoop = 0 To 3 Step 1
        If ilLoop <= 2 Then
            edcAddr(ilLoop).Text = Trim$(tmVef.sAddr(ilLoop))
        Else
            edcAddr(ilLoop).Text = Trim$(tmVff.sAddr4)
        End If
    Next ilLoop
    gSetPhoneNo tmVef.sPhone, mkcPhone
    gSetPhoneNo tmVef.sFax, mkcFax
    edcDialPos.Text = Trim$(tmVef.sDialPos)
    smOrigTax = ""
    edcTaxDropdown.Text = ""
    If imTaxDefined Then
        If tmVef.iTrfCode > 0 Then
            For ilVef = 0 To UBound(tmTaxSortCode) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
                slNameCode = tmTaxSortCode(ilVef).sKey   'lbcVehGpCode.List(ilVef)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tmVef.iTrfCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", smOrigTax)
                    edcTaxDropdown.Text = smOrigTax
                    Exit For
                End If
            Next ilVef
        End If
    End If
    smOrigHub = ""
    edcHubDropdown.Text = ""
    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
        If tmVef.iMnfHubCode > 0 Then
            For ilVef = 0 To UBound(tmHubCode) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
                slNameCode = tmHubCode(ilVef).sKey   'lbcVehGpCode.List(ilVef)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tmVef.iMnfHubCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", smOrigHub)
                    edcHubDropdown.Text = smOrigHub
                    Exit For
                End If
            Next ilVef
        End If
    End If
    edcACT1Lineup.Text = ""
    If imACT1CodesDefined Then
        edcACT1Lineup.Text = Trim$(tmVff.sACT1LineupCode)
    End If
    imTypeChgMode = True
    smVehicleType = tmVef.sType
    Select Case tmVef.sType
        Case "C"
            edcTypeDropDown.Text = "Conventional"
        Case "S"
            edcTypeDropDown.Text = "Selling"
        Case "A"
            edcTypeDropDown.Text = "Airing"
        Case "L"
            edcTypeDropDown.Text = "Log"
        Case "T"
            edcTypeDropDown.Text = "Simulcast"
        Case "V"
            edcTypeDropDown.Text = "Virtual"
        Case "P"
            edcTypeDropDown.Text = "Package"
        Case "R"
            edcTypeDropDown.Text = "Rep"
        Case "G"
            edcTypeDropDown.Text = "Sport"
        Case "N"
            edcTypeDropDown.Text = "NTR"
            'cmcOptions.Enabled = False
    End Select
    smOrigType = edcTypeDropDown.Text
    imTypeChgMode = False
    mLogVehPop
    smLogVeh = ""
    If tmVef.iVefCode > 0 Then
        For ilLoop = 0 To UBound(tmLogVehCode) - 1 Step 1 'lbcLogVehCode.ListCount - 1 Step 1
            slNameCode = tmLogVehCode(ilLoop).sKey 'lbcLogVehCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iVefCode Then
                ilRet = gParseItem(slNameCode, 1, "\", smLogVeh)
                Exit For
            End If
        Next ilLoop
    End If
    imLVChgMode = True
    edcLogVehDropDown.Text = smLogVeh
    imLVChgMode = False

    mMultiConvVehLogPop
    mMultiGameVehLogPop
    smMultiVehLog = ""
    If tmVef.iCombineVefCode > 0 Then
        If (smVehicleType = "C") Or (smVehicleType = "A") Then
            For ilLoop = 0 To UBound(tmMultiConvVehLogCode) - 1 Step 1 'lbcLogVehCode.ListCount - 1 Step 1
                slNameCode = tmMultiConvVehLogCode(ilLoop).sKey 'lbcLogVehCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tmVef.iCombineVefCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", smMultiVehLog)
                    ilRet = gParseItem(smMultiVehLog, 3, "|", smMultiVehLog)
                    Exit For
                End If
            Next ilLoop
        ElseIf smVehicleType = "G" Then
            For ilLoop = 0 To UBound(tmMultiGameVehLogCode) - 1 Step 1 'lbcLogVehCode.ListCount - 1 Step 1
                slNameCode = tmMultiGameVehLogCode(ilLoop).sKey 'lbcLogVehCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tmVef.iCombineVefCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", smMultiVehLog)
                    ilRet = gParseItem(smMultiVehLog, 3, "|", smMultiVehLog)
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    imMVLChgMode = True
    edcMultiVehDropdown.Text = smMultiVehLog
    imMVLChgMode = False

    mRNLinkPop
    smRNLink = ""
    If tmVef.iNrfCode > 0 Then
        For ilLoop = 0 To UBound(tmRNLinkCode) - 1 Step 1 'lbcLogVehCode.ListCount - 1 Step 1
            slNameCode = tmRNLinkCode(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iNrfCode Then
                ilRet = gParseItem(slNameCode, 1, "\", smRNLink)
                Exit For
            End If
        Next ilLoop
    End If
    imRNLinkChgMode = True
    edcRNLinkDropdown.Text = smRNLink
    imRNLinkChgMode = False

    edcStationCode.Text = Trim$(tmVef.sCodeStn)
    'edcFormat.Text = Trim$(tmVef.sFormat)
    mVehGp4Pop
    smOrigVehGp4 = ""
    smVehGp4 = ""
    If tmVef.iMnfVehGp4Fmt > 0 Then
        For ilVef = 0 To UBound(tmVehGp4Code) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
            slNameCode = tmVehGp4Code(ilVef).sKey   'lbcVehGpCode.List(ilVef)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iMnfVehGp4Fmt Then
                ilRet = gParseItem(slNameCode, 1, "\", smVehGp4)
                smOrigVehGp4 = smVehGp4
                Exit For
            End If
        Next ilVef
    End If
    mVehGp2Pop
    smOrigVehGp2 = ""
    smVehGp2 = ""
    If tmVef.iMnfVehGp2 > 0 Then
        For ilVef = 0 To UBound(tmVehGp2Code) - 1 Step 1 'lbcVehGpCode.ListCount - 1 Step 1
            slNameCode = tmVehGp2Code(ilVef).sKey   'lbcVehGpCode.List(ilVef)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iMnfVehGp2 Then
                ilRet = gParseItem(slNameCode, 1, "\", smVehGp2)
                smOrigVehGp2 = smVehGp2
                Exit For
            End If
        Next ilVef
    End If
    If tmVef.iSort > 0 Then
        edcSort.Text = Trim$(Str$(tmVef.iSort))
    Else
        edcSort.Text = ""
    End If
    mBookPop False
    smBook = ""
    If tmVef.iDnfCode > 0 Then
        For ilLoop = 0 To UBound(tmBookCode) - 1 Step 1 'lbcBookCode.ListCount - 1 Step 1
            slNameCode = tmBookCode(ilLoop).sKey   'lbcBookCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iDnfCode Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 2, "|", smBook)
                Exit For
            End If
        Next ilLoop
    End If
    imBookChgMode = True
    edcBookDropDown.Text = smBook
    imBookChgMode = False
    smReall = ""
    If tgSpf.sCAudPkg = "Y" Then
        If tmVef.iReallDnfCode > 0 Then
            For ilLoop = 0 To UBound(tmBookCode) - 1 Step 1 'lbcBookCode.ListCount - 1 Step 1
                slNameCode = tmBookCode(ilLoop).sKey   'lbcBookCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = tmVef.iReallDnfCode Then
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slName, 2, "|", smReall)
                    Exit For
                End If
            Next ilLoop
        End If
        imBookChgMode = True
        edcReallDropDown.Text = smReall
        imBookChgMode = False
    End If
    mDemoPop
    smDemo = ""
    If tmVef.iMnfDemo > 0 Then
        For ilLoop = 0 To UBound(tmDemoCode) - 1 Step 1 'lbcDemoCode.ListCount - 1 Step 1
            slNameCode = tmDemoCode(ilLoop).sKey   'lbcDemoCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If Val(slCode) = tmVef.iMnfDemo Then
                ilRet = gParseItem(slNameCode, 1, "\", smDemo)
                Exit For
            End If
        Next ilLoop
    End If
    imDemoChgMode = True
    edcDemoDropDown.Text = smDemo
    imDemoChgMode = False
    'cmcOptions.Enabled = True
    Select Case tmVef.sState
        Case "A"
            imState = 0
        Case "D"
            imState = 1
        Case Else
            imState = -1
    End Select
    imOrigState = imState
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim slStr As String
    Dim ilVef As Integer
    Dim ilRet As Integer

    If edcName.Text <> "" Then    'Test name
        
        'cbcSelect contains only non-package vehicles when rbcType(0) selected or only package vehicle when rbcType(1) selected
        slStr = Trim$(edcName.Text)
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(edcName.Text) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Vehicle already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    edcName.Text = Trim$(tmVef.sName) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
        If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
            '6/8/18
            'tmMVef contains all non-hub vehicle plus user hub vehicles (urf.iMnfHubCode = vef.iMnfHubCode)
            'If name changed, then repopulate array in case another user has entered the vehicle name
            If StrComp(Trim$(smOrigName), Trim$(edcName.Text), vbTextCompare) <> 0 Then
                gObtainVefIgnoreHub tmMVef()
            End If
            For ilVef = LBound(tmMVef) To UBound(tmMVef) - 1 Step 1
                If tmVef.iCode <> tmMVef(ilVef).iCode Then
                    If ((rbcType(0).Value) And (tmMVef(ilVef).sType <> "P")) Or ((rbcType(1).Value) And (tmMVef(ilVef).sType = "P")) Then
                        If StrComp(Trim$(tmMVef(ilVef).sName), Trim$(edcName.Text), vbTextCompare) = 0 Then
                            Beep
                            MsgBox "Vehicle already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                            edcName.Text = Trim$(tmVef.sName) 'Reset text
                            mSetShow imBoxNo
                            mSetChg imBoxNo
                            imBoxNo = 1
                            mEnableBox imBoxNo
                            mOKName = False
                            Exit Function
                        End If
                    End If
                End If
            Next ilVef
        Else
            '6/8/18
            'tgMVef contains all vehicles (Non-package plus package)
            'If name changed, then repopulate array in case another user has entered the vehicle name
            If StrComp(Trim$(smOrigName), Trim$(edcName.Text), vbTextCompare) <> 0 Then
                sgMVefStamp = "~"
                ilRet = gObtainVef()    'Re-Populate tgMVef only if chnaged since the last time array loaded (fct_File_Chg_Table checked)
            End If
            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If tmVef.iCode <> tgMVef(ilVef).iCode Then
                    If StrComp(Trim$(tgMVef(ilVef).sName), Trim$(edcName.Text), vbTextCompare) = 0 Then
                        Beep
                        If ((rbcType(0).Value) And (tgMVef(ilVef).sType <> "P")) Then
                            MsgBox "Vehicle already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                        ElseIf ((rbcType(1).Value) And (tgMVef(ilVef).sType = "P")) Then
                            MsgBox "Vehicle already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                        ElseIf rbcType(0).Value Then
                            MsgBox "Vehicle name already defined as a Package vehicle name, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                        Else
                            MsgBox "Vehicle name already defined as a non-package vehicle name, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                        End If
                        edcName.Text = Trim$(tmVef.sName) 'Reset text
                        mSetShow imBoxNo
                        mSetChg imBoxNo
                        imBoxNo = 1
                        mEnableBox imBoxNo
                        mOKName = False
                        Exit Function
                    End If
                End If
            Next ilVef
        End If
    End If
    mOKName = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintReall                     *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Reallocate Book Name     *
'*                      if required                    *
'*                                                     *
'*******************************************************
Private Sub mPaintReall()
    'Dim llColor As Long
    'Dim slFontName As String
    'Dim flFontSize As Single
    'If tgSpf.sCAudPkg <> "Y" Then
    '    Exit Sub
    'End If
    'llColor = pbcVeh.ForeColor
    'slFontName = pbcVeh.FontName
    'flFontSize = pbcVeh.FontSize
    'pbcVeh.ForeColor = BLUE
    'pbcVeh.FontBold = False
    'pbcVeh.FontSize = 7
    'pbcVeh.FontName = "Arial"
    'pbcVeh.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    'pbcVeh.Line (1455, tmCtrls(REALLINDEX).fBoxY)-Step(0, 330)
    'pbcVeh.CurrentX = tmCtrls(REALLINDEX).fBoxX + 15  'fgBoxInsetX
    'pbcVeh.CurrentY = tmCtrls(REALLINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    'pbcVeh.Print "Reallocation Book"
    'pbcVeh.FontSize = flFontSize
    'pbcVeh.FontName = slFontName
    'pbcVeh.FontSize = flFontSize
    'pbcVeh.ForeColor = llColor
    'pbcVeh.FontBold = True
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
'10071
Private Sub mParseCmmdLine()

    Dim slCommand As String
    Dim slStr As String
    Dim slStartIn As String
    Dim blIsCsi As Boolean
    Dim blIsFromTraffic As Boolean
    
    sgCommandStr = Command$
    slStartIn = CurDir$
    blIsCsi = False
    igShowVersionNo = 0
    slCommand = sgCommandStr
    lgCurrHRes = GetDeviceCaps(Traffic!pbcList.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(Traffic!pbcList.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(Traffic!pbcList.hdc, BITSPIXEL)
    '10365 we need to connect to database now, so need 'test' or 'prod' now.
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    mTestPervasive
    lgUlfCode = 0
    blIsFromTraffic = False
    'not coming from traffic
    If InStr(1, sgCommandStr, "^", vbTextCompare) <= 0 Then
        'if not allowing stand alone:
        'MsgBox "Contract Projection must be run from Traffic->Proposals"
        'igExitTraffic = True
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If
        '10365 already did this
'        If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
'            igTestSystem = False
'        Else
'            igTestSystem = True
'        End If
        slStr = sgUserName
        sgCallAppName = "Traffic"
    Else
        blIsFromTraffic = True
        gParseItem slCommand, 1, "\", slStr    'Get application name
        gParseItem slStr, 1, "^", sgCallAppName    'Get application name
        '10365 already did this.  No longer testing what gets sent from traffic.
'        gParseItem slStr, 2, "^", slStr   'Get Test or prod
'        If slStr = "Prod" Then
'            igTestSystem = False
'        Else
'            igTestSystem = True
'        End If
        gParseItem slCommand, 2, "\", slStr    'user
        gParseItem slStr, 1, "/", slStr
        sgUrfStamp = "~" 'Clear time stamp incase same name
        sgUserName = Trim$(slStr)
    End If
    If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
        sgSpecialPassword = mDetermineCsiLogin()
    End If
    'sets  sgCPName & sgSUName to 'counterpoint' and 'Guide'. Sets super user because sgSpecialPassword above is set.
    gUrfRead Vehicle, sgUserName, True, tgUrf(), False  'Obtain user records
    If Len(sgSpecialPassword) > 0 Then
        'sets the igWinStatus that we need.  Reruns gInitSuperUser (from gUrfRead) but that doesn't do anything because tlUrf.sName is blank from gUrfRead
        gExpandGuideAsUser tgUrf(0)
        sgUserName = "Guide"
        tgUrf(0).sName = sgUserName
    Else
        If Not blIsFromTraffic Then
            'is this user allowed to access traffic?
            If igWinStatus(VEHICLESLIST) < 1 Then
                gMsgBox "This user does not have access to Vehicle screens", vbExclamation, "Access Denied"
                imTerminate = True
            End If
        End If
    End If
    mGetUlfCode
    sgVehName = ""
    DoEvents
    mInitStdAloneVehicle
    mCheckForDate
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim llFilter As Long
    Dim ilValue As Integer

    imPopReqd = False
    If rbcType(0).Value Then
        igVpfType = 0
        'ilFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHLOG + VEHVIRTUAL + VEHLOGVEHICLE + VEHSIMUL + ACTIVEVEH + DORMANTVEH
        'Don't include virtual vehicles 4/2/02- Jim/Mary
        llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHLOG + VEHLOGVEHICLE + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + VEHSIMUL + VEHSPORT + VEHIMPORTAFFILIATESPOTS + ACTIVEVEH + DORMANTVEH
    Else
        igVpfType = 1
        llFilter = VEHPACKAGE + ACTIVEVEH + DORMANTVEH
    End If
    'ilRet = gPopUserVehicleBox(Vehicle, ilFilter, cbcSelect, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(Vehicle, llFilter, cbcSelect, tmVehicle(), smVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", Vehicle
        On Error GoTo 0
        If cbcSelect.List(0) <> "[New]" Then
            cbcSelect.AddItem "[New]", 0  'Force as first item on list
        End If
        imPopReqd = True
    End If
    If rbcType(0).Value Then
        cmcOptions.Enabled = True
        lbcType.Clear
        lbcType.AddItem "Conventional"
        If tgSpf.sSSellNet = "Y" Then
            lbcType.AddItem "Selling"
            lbcType.AddItem "Airing"
        End If
        lbcType.AddItem "Log"
        lbcType.AddItem "Simulcast"
        'If tgSpf.sMktBase = "Y" Then
        If (Asc(tgSpf.sUsingFeatures) And USINGREP) = USINGREP Then
            lbcType.AddItem "Rep"
        End If
        lbcType.AddItem "NTR"
        ilValue = Asc(tgSpf.sSportInfo)
        If (ilValue And USINGSPORTS) = USINGSPORTS Then
            lbcType.AddItem "Sport"
        End If
        'Don't include virtual vehicles 4/2/02- Jim/Mary
        'lbcType.AddItem "Virtual"
    Else
        cmcOptions.Enabled = False
        lbcType.Clear
        lbcType.AddItem "Package"
    End If
    If Not imFromOldSave Then
        cbcSelect.ListIndex = 0
    End If
    imFromOldSave = False
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer) As Integer
'
'   iRet = ENmRead(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    slNameCode = tmVehicle(ilSelectIndex - 1).sKey 'Traffic!lbcVehicle.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", Vehicle
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmVefSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", Vehicle
    On Error GoTo 0
    tmVffSrchKey1.iCode = tmVef.iCode
    ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        mAddVff
    End If

    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slStr As String
    Dim slStamp As String   'Date/Time stamp for file
    '10071 this is never used
    'Dim ilInitOption As Integer
    Dim ilSelectedIndex As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilRetainGroupName As Integer
    Dim ilVpfIndex As Integer
    
 '10071 never used
'    If imSelectedIndex = 0 Then
'        'If Not gWinRoom(igNoExeWinRes(VEHOPTEXE)) Then
'        '    Exit Function
'        'End If
'        ilInitOption = True
'    Else
'        ilInitOption = False
'    End If
    mSetShow imBoxNo
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        Screen.MousePointer = vbDefault  'Wait
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        Screen.MousePointer = vbDefault  'Wait
        mSaveRec = False
        Exit Function
    End If
    If (imSelectedIndex > 0) And (imOrigState <> imState) And (imState = 1) Then
        slMsg = ""
        slStr = edcTypeDropDown.Text
        gFindMatch slStr, 0, lbcType
        If gLastFound(lbcType) >= 0 Then
            slStr = lbcType.List(gLastFound(lbcType))
            If StrComp(slStr, "Conventional", 1) = 0 Then
                slMsg = "Before changing the Vehicle to Dormant:  End Programming and Remove References in Currently Used Rate Cards"
            ElseIf StrComp(slStr, "Selling", 1) = 0 Then
                slMsg = "Before changing the Vehicle to Dormant:  Remove Links Only if No Logs Generated or Terminate Links if Logs Generated, then End Programming and Remove References in Currently Used Rate Cards"
            ElseIf StrComp(slStr, "Airing", 1) = 0 Then
                slMsg = "Before changing the Vehicle to Dormant:  Remove Links Only if No Logs Generated or Terminate Links if Logs Generated, then End Programming "
            ElseIf StrComp(slStr, "Log", 1) = 0 Then
                slMsg = "Before changing the Vehicle to Dormant:  Remove References in Other Vehicles if No logs Generated"
            ElseIf StrComp(slStr, "Virtual", 1) = 0 Then
                slMsg = ""
            ElseIf StrComp(slStr, "Simulcast", 1) = 0 Then
                slMsg = "Before changing the Vehicle to Dormant:  Remove References in Other Vehicles if No Logs Generated, and Remove References in Currently Used Rate Cards"
            ElseIf StrComp(slStr, "Package", 1) = 0 Then
                slMsg = "Before changing the Vehicle to Dormant:  Remove References in Currently Used Rate Cards"
            ElseIf StrComp(slStr, "Rep", 1) = 0 Then
                slMsg = "Before changing the Vehicle to Dormant:  Remove References in Currently Used Rate Cards"
            ElseIf StrComp(slStr, "Sport", 1) = 0 Then
                slMsg = "Before changing the Vehicle to Dormant:  End Programming and Remove References in Currently Used Rate Cards"
            ElseIf StrComp(slStr, "NTR", 1) = 0 Then
                slMsg = ""
            End If
        End If
        If slMsg <> "" Then
            ilRet = MsgBox(slMsg & ", Task Completed and Continue with Save", vbYesNo + vbQuestion, "Save")
            If ilRet = vbNo Then
                Screen.MousePointer = vbDefault  'Wait
                mSaveRec = False
                Exit Function
            End If
        End If
    End If
    If (imSelectedIndex > 0) Then
        slStr = edcTypeDropDown.Text
        If (StrComp(smOrigType, slStr, 1) <> 0) And (smOrigType <> "") Then
            If tgSpf.sGUseAffSys = "Y" Then
                If (StrComp(smOrigType, "Conventional", 1) = 0) And (StrComp(slStr, "Selling", 1) = 0) Then
                    If (Asc(tgSpf.sUsingFeatures7) And WEGENEREXPORT) = WEGENEREXPORT Then
                        If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                            GoTo mSaveRecErr
                        End If
                        tmVffSrchKey1.iCode = tmVef.iCode
                        ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet = BTRV_ERR_NONE Then
                            If Trim$(tmVff.sGroupName) <> "" Then
                                ilRetainGroupName = MsgBox("Retain Group Name " & Trim$(tmVff.sGroupName) & " with Selling vehicle", vbYesNoCancel + vbExclamation + vbDefaultButton2, "Group Name")
                                If ilRetainGroupName = vbCancel Then
                                    Screen.MousePointer = vbDefault  'Wait
                                    mSaveRec = False
                                    Exit Function
                                End If
                                If ilRetainGroupName = vbNo Then
                                    ilRet = mRemoveGroupName(tmVef.iCode)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass  'Wait
    gGetSyncDateTime slSyncDate, slSyncTime
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Vef.btr")
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                GoTo mSaveRecErr
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            If tmVef.sType = "V" Then
                tmVef.lVsfCode = mCreateVsf()
            Else
                tmVef.lVsfCode = 0
            End If
            tmVef.sExportRAB = "N"
            tmVef.iCode = 0  'Autoincrement
            tmVef.iRemoteID = tgUrf(0).iRemoteUserID
            tmVef.iAutoCode = tmVef.iCode
            ilRet = btrInsert(hmVef, tmVef, imVefRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            If (Left$(smOrigType, 1) = "V") And (tmVef.sType <> "V") Then
                'Remove Virtual vehicle, add preference
                ilRet = mDeleteVsf(tmVef.lVsfCode)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                tmVef.lVsfCode = 0
                ilRet = gVpfFind(Vehicle, tmVef.iCode)
            ElseIf (Left$(smOrigType, 1) <> "V") And (tmVef.sType = "V") Then
                'Add virtual vehicle, remove preference
                tmVef.lVsfCode = mCreateVsf()
                'ilRet = mDeleteVpf(tmVef.iCode)
            ElseIf tmVef.sType = "V" Then   'Update name
                Do
                    tmVsfSrchKey.lCode = tmVef.lVsfCode
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        tmVsf.sName = Left$(tmVef.sName, 20)
                        ilRet = btrUpdate(hmVsf, tmVsf, imVsfRecLen)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
            ElseIf (Left$(smOrigType, 1) = "C") And (tmVef.sType = "S") Then
                ilRet = mDeleteVff(tmVef.iCode)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
            End If
            'tmVef.iSourceID = tgUrf(0).iRemoteUserID
            'gPackDate slSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
            'gPackTime slSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            slMsg = "mSaveRec (btr(Update)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Vehicle
    On Error GoTo 0
    If imSelectedIndex = 0 Then 'New selected
        Do
            'tmVefSrchKey.iCode = tmVef.iCode
            'ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'slMsg = "mSaveRec (btrGetEqual:Vehicle)"
            'On Error GoTo mSaveRecErr
            'gBtrvErrorMsg ilRet, slMsg, Vehicle
            'On Error GoTo 0
            tmVef.iRemoteID = tgUrf(0).iRemoteUserID
            tmVef.iAutoCode = tmVef.iCode
            'tmVef.iSourceID = tgUrf(0).iRemoteUserID
            'gPackDate slSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
            'gPackTime slSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            slMsg = "mSaveRec (btrUpdate:Vehicle)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Vehicle
        On Error GoTo 0
        tgMVef(UBound(tgMVef)) = tmVef
        'ReDim Preserve tgMVef(1 To UBound(tgMVef) + 1) As VEF
        ReDim Preserve tgMVef(0 To UBound(tgMVef) + 1) As VEF
        'If UBound(tgMVef) > 2 Then
        If UBound(tgMVef) > 1 Then
            'ArraySortTyp fnAV(tgMVef(), 1), UBound(tgMVef) - 1, 0, LenB(tgMVef(1)), 0, -1, 0
            ArraySortTyp fnAV(tgMVef(), 0), UBound(tgMVef), 0, LenB(tgMVef(0)), 0, -1, 0
        End If
    Else
        ilRet = gBinarySearchVef(tmVef.iCode)
        If ilRet <> -1 Then
            tgMVef(ilRet) = tmVef
        End If
    End If
    
    '7/25/14: Moved code above VFF as gVpfFind is called some times resulting in a duplicate key error and values not retained
    'If insert- add Vehicle preference file if not virtual
    ilSelectedIndex = imSelectedIndex
    If ilSelectedIndex = 0 Then
        'If tmVef.sType <> "V" Then
        If igVefCodeModel > 0 Then
            ilRet = mAddVpfModel(slSyncDate, slSyncTime)
            On Error GoTo mSaveRecErr
            ilRet = gVpfFind(Vehicle, tmVef.iCode)
            On Error GoTo 0
            ilRet = mAddVofModel()
            ilRet = mAddPifModel()
        Else
            On Error GoTo mSaveRecErr
            ilRet = gVpfFind(Vehicle, tmVef.iCode)
            On Error GoTo 0
        End If
        'End If
    Else
        On Error GoTo mSaveRecErr
        ilRet = gVpfFind(Vehicle, tmVef.iCode)
        On Error GoTo 0
    End If
    tmVffSrchKey1.iCode = tmVef.iCode
    ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        mAddVff
        tmVffSrchKey1.iCode = tmVef.iCode
        ilRet = btrGetEqual(hmVff, tmVff, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    End If
    If ilRet = BTRV_ERR_NONE Then
        tmVff.sAddr4 = edcAddr(3).Text
        If tmCtrls(NAMEINDEX).iChg Then
            ilVpfIndex = gVpfFind(Vehicle, tmVef.iCode)
            If ilVpfIndex > 0 Then
                If tgVpf(ilVpfIndex).iInterfaceID > 0 Then
                    If tmVff.sSentToXDSStatus <> "N" Then
                        tmVff.sSentToXDSStatus = "M"
                    End If
                End If
            End If
        End If
        tmVff.sACT1LineupCode = edcACT1Lineup.Text
        ilRet = btrUpdate(hmVff, tmVff, imVffRecLen)
    End If
    
   
    'If Traffic!lbcVehicle.Tag <> "" Then
    '    If slStamp = Traffic!lbcVehicle.Tag Then
    '        Traffic!lbcVehicle.Tag = FileDateTime(sgDBPath & "Vef.btr")
    '    End If
    'End If
    'If imSelectedIndex <> 0 Then
    '    Traffic!lbcVehicle.RemoveItem imSelectedIndex - 1
    '    cbcSelect.RemoveItem imSelectedIndex
    'End If
    'Traffic!lbcVehicle.Tag = "" 'Force Population
    sgVehNameToVehOpt = RTrim$(tmVef.sName)
    'sgVehicleTag = ""
    'Not required as tgMVef changed above
    'sgMVefStamp = ""
    'ilRet = csiSetStamp("VEF", sgMVefStamp)
    DoEvents
    '11/26/17
    gFileChgdUpdate "vef.btr", True
    smVehicleTag = ""
    mPopulate
    gObtainVefIgnoreHub tmMVef()
'    sgVpfStamp = "~"    'Force read
'    ilRet = gVpfRead()
    gAnyClustersDef
    gAnyRepDef
    'cbcSelect.RemoveItem 0 'Remove [New]
    'slName = RTrim$(tmVef.sName)
    'cbcSelect.AddItem slName
    'sgVehNameToVehOpt = slName
    'slName = tmVef.sName + "\" + LTrim$(Str$(tmVef.iCode))
    'Traffic!lbcVehicle.AddItem slName
    'cbcSelect.AddItem "[New]", 0
    
    If ilSelectedIndex = 0 Then
        'If tmVef.sType <> "V" Then
        'If (tmVef.sType <> "P") And (tmVef.sType <> "T") Then
        'If (tmVef.sType <> "T") And (tmVef.sType <> "R") And (tmVef.sType <> "U") Then
        'Allow rep vehicles to option screen so that spot length can be defined 12/10/02
        If (tmVef.sType <> "T") And (tmVef.sType <> "N") Then
            'slName = RTrim$(tmVef.sName)
            'sgVehNameToVehOpt = slName
            igVehOptCallSource = CALLSOURCEVEHICLE
            igVehNewToVehOpt = True
            'Screen.MousePointer = vbHourGlass  'Wait
            VehOpt.Show vbModal
            '10071
            pbcClickFocus.SetFocus
            'Screen.MousePointer = vbDefault    'Default
            igVehOptCallSource = CALLNONE
        'Else
        '    'Screen.MousePointer = vbHourGlass  'Wait
        '    sgVsfName = Trim$(tmVef.sName)
        '    VirtVeh.Show vbModal
        '    'Screen.MousePointer = vbDefault    'Default
        'End If
        End If
    End If
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                      *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    Dim ilAltered As Integer
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If ckcLock.Visible Then
        If ckcLock.Value = vbUnchecked Then
            ilAltered = True
        End If
    End If
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        Screen.MousePointer = vbDefault  'Wait
        If ilAltered = YES Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & edcName.Text
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcVeh_Paint
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    Exit Function
                End If
                If ilRes = vbNo Then
                    cbcSelect.ListIndex = 0
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
        End If
    End If
    Screen.MousePointer = vbDefault  'Wait
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/11/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            gSetChgFlag tmVef.sName, edcName, tmCtrls(ilBoxNo)
        Case CONTACTINDEX 'Name
            gSetChgFlag tmVef.sContact, edcContact, tmCtrls(ilBoxNo)
        Case MKTNAMEINDEX 'Market Name
            'gSetChgFlag tmVef.sMktName, edcMktName, tmCtrls(ilBoxNo)
            gSetChgFlag smOrigVehGp3, lbcVehGp3, tmCtrls(ilBoxNo)
        Case RSCHINDEX
            gSetChgFlag smOrigVehGp5, lbcVehGp5, tmCtrls(ilBoxNo)
        Case SUBCOMPINDEX
            gSetChgFlag smOrigVehGp6, lbcVehGp6, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX 'Address
            gSetChgFlag tmVef.sAddr(ilBoxNo - ADDRESSINDEX), edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 1 'Address
            gSetChgFlag tmVef.sAddr(ilBoxNo - ADDRESSINDEX), edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 2 'Address
            gSetChgFlag tmVef.sAddr(ilBoxNo - ADDRESSINDEX), edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 3 'Address
            gSetChgFlag tmVff.sAddr4, edcAddr(ilBoxNo - ADDRESSINDEX), tmCtrls(ilBoxNo)
        Case PHONEINDEX 'Phone number plus extension
            gSetChgFlag tmVef.sPhone, mkcPhone, tmCtrls(ilBoxNo)
        Case FAXINDEX 'Fax number
            gSetChgFlag tmVef.sFax, mkcFax, tmCtrls(ilBoxNo)
        Case FORMATINDEX 'Format
            'gSetChgFlag tmVef.sFormat, edcFormat, tmCtrls(ilBoxNo)
            gSetChgFlag smOrigVehGp4, lbcVehGp4, tmCtrls(ilBoxNo)
        Case SUBTOTALINDEX   'Vehicle Group Set 2
            gSetChgFlag smOrigVehGp2, lbcVehGp2, tmCtrls(ilBoxNo)
        Case DIALPOSINDEX 'Dial Position
            gSetChgFlag tmVef.sDialPos, edcDialPos, tmCtrls(ilBoxNo)
        Case ACT1CODESINDEX 'ACT1 Codes
            If imACT1CodesDefined Then
                gSetChgFlag tmVff.sACT1LineupCode, edcACT1Lineup, tmCtrls(ilBoxNo)
            End If
        Case TAXINDEX
            If imTaxDefined Then
                gSetChgFlag smOrigTax, lbcTax, tmCtrls(ilBoxNo)
            End If
        Case HUBINDEX
            If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
                gSetChgFlag smOrigHub, lbcHub, tmCtrls(ilBoxNo)
            End If
        Case LOGVEHINDEX   'Log Vehicle
            gSetChgFlag smLogVeh, lbcLogVeh, tmCtrls(ilBoxNo)
        Case MULTIVEHLOGINDEX   'Log Vehicle
            If (smVehicleType = "C") Or (smVehicleType = "A") Then
                gSetChgFlag smMultiVehLog, lbcMultiConvVehLog, tmCtrls(ilBoxNo)
            ElseIf smVehicleType = "G" Then
                gSetChgFlag smMultiVehLog, lbcMultiGameVehLog, tmCtrls(ilBoxNo)
            End If
        Case RNLINKINDEX   'Log Vehicle
            gSetChgFlag smRNLink, lbcRNLink, tmCtrls(ilBoxNo)
        Case SORTINDEX 'Sort
            slStr = Trim$(Str$(tmVef.iSort))
            gSetChgFlag slStr, edcSort, tmCtrls(ilBoxNo)
        Case SCODEINDEX 'Dial Position
            gSetChgFlag tmVef.sCodeStn, edcStationCode, tmCtrls(ilBoxNo)
        Case BOOKINDEX   'Book
            gSetChgFlag smBook, lbcBook, tmCtrls(ilBoxNo)
        'Case REALLINDEX   'Book
        '    gSetChgFlag smReall, lbcBook, tmCtrls(ilBoxNo)
        Case DEMOINDEX   'Demo
            gSetChgFlag smDemo, lbcDemo, tmCtrls(ilBoxNo)
        Case TYPEINDEX
            gSetChgFlag smOrigType, lbcType, tmCtrls(ilBoxNo)
        Case STATEINDEX
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
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
    Dim ilAltered As Integer
    'If rbcType(0).Value Then
        If imSelectedIndex > 0 Then
            cmcOptions.Enabled = True
        Else
            cmcOptions.Enabled = False
        End If
    'End If
    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If ckcLock.Visible Then
        If ckcLock.Value = vbUnchecked Then
            ilAltered = True
        End If
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    Screen.MousePointer = vbDefault  'Wait
    'Revert button set if any field changed
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
        If imUpdateAllowed Then
            cmcErase.Enabled = True
        Else
            cmcErase.Enabled = False
        End If
    Else
        cmcErase.Enabled = False
    End If
    Screen.MousePointer = vbDefault  'Wait
    If Not ilAltered Then
        cbcSelect.Enabled = True
        rbcType(0).Enabled = True
        rbcType(1).Enabled = True
    Else
        cbcSelect.Enabled = False
        rbcType(0).Enabled = False
        rbcType(1).Enabled = False
    End If
    'If rbcType(0).Value Then
        If (imSelectedIndex > 0) And (Not ilAltered) Then
            cmcOptions.Enabled = True
        Else
            cmcOptions.Enabled = False
        End If
    'End If
'    If cbcSelect.ListCount <= 2 Then
'        cmcCombo.Enabled = False
'    Else
'        cmcCombo.Enabled = True
'    End If
    If (Not ilAltered) And (tgUrf(0).sMerge = "I") And (imUpdateAllowed) Then
        'Disallow Merge until all files added to the merge logic
        cmcMerge.Enabled = False    'True
    Else
        cmcMerge.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxCtrlNo) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.SetFocus
        Case CONTACTINDEX 'Name
            edcContact.SetFocus
        Case MKTNAMEINDEX 'Market Name
            'edcMktName.SetFocus
            edcVehGp3DropDown.SetFocus
        Case RSCHINDEX
            edcVehGp5DropDown.SetFocus
        Case SUBCOMPINDEX 'Sub-Company
            edcVehGp6DropDown.SetFocus
        Case ADDRESSINDEX 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).SetFocus
        Case ADDRESSINDEX + 1 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).SetFocus
        Case ADDRESSINDEX + 2 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).SetFocus
        Case ADDRESSINDEX + 3 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).SetFocus
        Case PHONEINDEX 'Phone and extension
            mkcPhone.SetFocus
        Case FAXINDEX 'Fax
            mkcFax.SetFocus
        Case FORMATINDEX 'Format
            'edcFormat.SetFocus
            edcVehGp4DropDown.SetFocus
        Case SUBTOTALINDEX 'Vehicle Group
            edcVehGp2DropDown.SetFocus
        Case DIALPOSINDEX 'Dial Position
            edcDialPos.SetFocus
        Case ACT1CODESINDEX 'ACT1 Codes
            edcACT1Lineup.SetFocus
        Case TAXINDEX 'Market Name
            edcTaxDropdown.SetFocus
        Case HUBINDEX 'Market Name
            edcHubDropdown.SetFocus
        Case LOGVEHINDEX 'Log Vehicle
            edcLogVehDropDown.SetFocus
        Case MULTIVEHLOGINDEX 'Log Vehicle
            edcMultiVehDropdown.SetFocus
        Case RNLINKINDEX 'Log Vehicle
            edcRNLinkDropdown.SetFocus
        Case SORTINDEX 'Sort
            edcSort.SetFocus
        Case SCODEINDEX 'Station code
            edcStationCode.SetFocus
        Case BOOKINDEX 'Book
            edcBookDropDown.SetFocus
        'Case REALLINDEX 'Book
        '    edcReallDropDown.SetFocus
        Case DEMOINDEX 'Demo
            edcDemoDropDown.SetFocus
        Case TYPEINDEX 'Type
            edcTypeDropDown.SetFocus
        Case STATEINDEX 'State
            pbcState.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imMaxCtrlNo) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case CONTACTINDEX 'Name
            edcContact.Visible = False  'Set visibility
            slStr = edcContact.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case MKTNAMEINDEX 'Market Name
            'edcMktName.Visible = False  'Set visibility
            'slStr = edcMktName.Text
            'gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
            lbcVehGp3.Visible = False
            edcVehGp3DropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcVehGp3DropDown.Text
            smVehGp3 = slStr
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case RSCHINDEX
            lbcVehGp5.Visible = False
            edcVehGp5DropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcVehGp5DropDown.Text
            smVehGp5 = slStr
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case SUBCOMPINDEX 'Sub-Company
            lbcVehGp6.Visible = False
            edcVehGp6DropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcVehGp6DropDown.Text
            smVehGp6 = slStr
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).Visible = False
            slStr = edcAddr(ilBoxNo - ADDRESSINDEX).Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 1 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).Visible = False
            slStr = edcAddr(ilBoxNo - ADDRESSINDEX).Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 2 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).Visible = False
            slStr = edcAddr(ilBoxNo - ADDRESSINDEX).Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case ADDRESSINDEX + 3 'Address
            edcAddr(ilBoxNo - ADDRESSINDEX).Visible = False
            slStr = edcAddr(ilBoxNo - ADDRESSINDEX).Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case PHONEINDEX 'Phone number plus extension
            mkcPhone.Visible = False  'Set visibility
            If mkcPhone.Text = smPhoneImage Then
                slStr = ""
            Else
                slStr = mkcPhone.Text
            End If
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case FAXINDEX 'Fax number
            mkcFax.Visible = False  'Set visibility
            If mkcFax.Text = smFaxImage Then
                slStr = ""
            Else
                slStr = mkcFax.Text
            End If
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case FORMATINDEX 'Format
            'edcFormat.Visible = False  'Set visibility
            'slStr = edcFormat.Text
            'gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
            lbcVehGp4.Visible = False
            edcVehGp4DropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcVehGp4DropDown.Text
            smVehGp4 = slStr
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case SUBTOTALINDEX 'Vehicle
            lbcVehGp2.Visible = False
            edcVehGp2DropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcVehGp2DropDown.Text
            smVehGp2 = slStr
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case DIALPOSINDEX 'Dial Position
            edcDialPos.Visible = False  'Set visibility
            slStr = edcDialPos.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case ACT1CODESINDEX 'ACT1 Codes
            edcACT1Lineup.Visible = False  'Set visibility
            slStr = edcACT1Lineup.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case TAXINDEX 'Market Name
            lbcTax.Visible = False
            edcTaxDropdown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcTaxDropdown.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case HUBINDEX 'Market Name
            lbcHub.Visible = False
            edcHubDropdown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcHubDropdown.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case LOGVEHINDEX 'Vehicle
            lbcLogVeh.Visible = False
            edcLogVehDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcLogVehDropDown.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case MULTIVEHLOGINDEX 'Vehicle
            lbcMultiConvVehLog.Visible = False
            lbcMultiGameVehLog.Visible = False
            edcMultiVehDropdown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcMultiVehDropdown.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case RNLINKINDEX 'Vehicle
            lbcRNLink.Visible = False
            edcRNLinkDropdown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcRNLinkDropdown.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case SORTINDEX 'Sort
            edcSort.Visible = False  'Set visibility
            slStr = edcSort.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case SCODEINDEX 'Dial Position
            edcStationCode.Visible = False  'Set visibility
            slStr = edcStationCode.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case BOOKINDEX 'Vehicle
            lbcBook.Visible = False
            edcBookDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcBookDropDown.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        'Case REALLINDEX 'Vehicle
        '    lbcBook.Visible = False
        '    edcReallDropDown.Visible = False
        '    cmcDropDown.Visible = False
        '    slStr = edcReallDropDown.Text
        '    gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case DEMOINDEX 'Vehicle
            lbcDemo.Visible = False
            edcDemoDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDemoDropDown.Text
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case TYPEINDEX 'Vehicle
            lbcType.Visible = False
            edcTypeDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcTypeDropDown.Text
            If igVpfType = 1 Then
                If StrComp(slStr, "Package", vbTextCompare) = 0 Then
                    If tmVef.lPvfCode <= 0 Then
                        slStr = slStr & "-Dynamic"
                    Else
                        slStr = slStr & "-Standard"
                    End If
                End If
                smVehicleType = "P"
            Else
                slStr = edcTypeDropDown.Text
                gFindMatch slStr, 0, lbcType
                If gLastFound(lbcType) >= 0 Then
                    slStr = lbcType.List(gLastFound(lbcType))
                    If StrComp(slStr, "Conventional", 1) = 0 Then
                        smVehicleType = "C"
                    ElseIf StrComp(slStr, "Selling", 1) = 0 Then
                        smVehicleType = "S"
                    ElseIf StrComp(slStr, "Airing", 1) = 0 Then
                        smVehicleType = "A"
                    ElseIf StrComp(slStr, "Log", 1) = 0 Then
                        smVehicleType = "L"
                    ElseIf StrComp(slStr, "Virtual", 1) = 0 Then
                        smVehicleType = "V"
                    ElseIf StrComp(slStr, "Simulcast", 1) = 0 Then
                        smVehicleType = "T"
                    ElseIf StrComp(slStr, "Package", 1) = 0 Then
                        smVehicleType = "P"
                    ElseIf StrComp(slStr, "Rep", 1) = 0 Then
                        smVehicleType = "R"
                    ElseIf StrComp(slStr, "Sport", 1) = 0 Then
                        smVehicleType = "G"
                    ElseIf StrComp(slStr, "NTR", 1) = 0 Then
                        smVehicleType = "N"
                    End If
                End If
            End If
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
        Case STATEINDEX 'State
            pbcState.Visible = False  'Set visibility
            If imState = 0 Then
                slStr = "Active"
            ElseIf imState = 1 Then
                slStr = "Dormant"
            Else
                slStr = ""
            End If
            gSetShow pbcVeh, slStr, tmCtrls(ilBoxNo)
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mStartNew                       *
'*                                                     *
'*             Created:7/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up a New rate card and     *
'*                      initiate RCTerms               *
'*                                                     *
'*******************************************************
Private Function mStartNew() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slName As String
    'If Not gWinRoom(igNoExeWinRes(RCTERMSEXE)) Then
    '    mStartNew = False
    '    Screen.MousePointer = vbDefault
    '    Exit Function
    'End If
    slName = cbcSelect.Text
    If slName = "[New]" Then
        slName = ""
    End If
    imInNew = True
    If (cbcSelect.ListCount > 1) And (rbcType(0).Value) Then
            VehModel.Show vbModal
            If (igVehReturn = 0) Or (igVefCodeModel = 0) Then    'Cancelled
                igVefCodeModel = 0
                mStartNew = True
                imInNew = False
                Screen.MousePointer = vbDefault
                Exit Function
            End If
    Else
        igVefCodeModel = 0
        mStartNew = True
        imInNew = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    Screen.MousePointer = vbHourglass    '
    'Build program images from newest
    tmVefSrchKey.iCode = igVefCodeModel
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        igVefCodeModel = 0
        mStartNew = False
        imInNew = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    tmVef.iCode = 0
    tmVef.sName = ""
    tmVef.sDialPos = ""
    tmVef.iReallDnfCode = 0
    tmVef.sCodeStn = ""
    tmVef.sState = "A"
    tmVef.iSort = 0 'Later set to next number
    tmVef.iDnfCode = 0
    mMoveRecToCtrl
    smOrigType = ""
    For ilLoop = imLBCtrls To imMaxCtrlNo Step 1
        If (ilLoop <> MKTNAMEINDEX) And (ilLoop <> RSCHINDEX) And (ilLoop <> FORMATINDEX) And (ilLoop <> SUBTOTALINDEX) And (ilLoop <> SUBCOMPINDEX) Then
            mSetShow ilLoop  'Set show strings
        Else
            Select Case ilLoop
               Case MKTNAMEINDEX 'Vehicle
                    slStr = smVehGp3
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
               Case RSCHINDEX 'Vehicle
                    slStr = smVehGp5
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
               Case SUBCOMPINDEX 'Sub-Company
                    slStr = smVehGp6
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
               Case FORMATINDEX 'Vehicle
                    slStr = smVehGp4
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
               Case SUBTOTALINDEX 'Vehicle
                    slStr = smVehGp2
                    gSetShow pbcVeh, slStr, tmCtrls(ilLoop)
            End Select
        End If
    Next ilLoop
    If slName <> "[New]" Then
        edcName.Text = slName
    End If
    pbcVeh_Paint
    bmPrgLibDefined = False
    imChgMode = False
    mStartNew = True
    Screen.MousePointer = vbDefault
    mSetCommands
    imInNew = False
    Exit Function

    On Error GoTo 0
    mStartNew = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    Dim ilRet As Integer

    sgVffStamp = ""
    ilRet = gVffRead()



    sgDoneMsg = Trim$(Str$(igVehCallSource)) & "\" & sgVehName
    Screen.MousePointer = vbDefault
    'Unload IconTraf
    igManUnload = YES
    'Unload Traffic
    Unload Vehicle
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilCode As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slType As String
    Dim slStationCode As String

    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcName, "", "Name must be specified", tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = CONTACTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcContact, "", "Contact name must be specified", tmCtrls(CONTACTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = CONTACTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If rbcType(0).Value Then
        If tgSpf.sSubCompany = "Y" Then
            tmCtrls(SUBCOMPINDEX).iReq = True
        Else
            tmCtrls(SUBCOMPINDEX).iReq = False
        End If
    Else
        tmCtrls(SUBCOMPINDEX).iReq = False
    End If
    If (ilCtrlNo = MKTNAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If gFieldDefinedCtrl(edcMktName, "", "Market Name must be specified", tmCtrls(MKTNAMEINDEX).iReq, ilState) = NO Then
        '    If ilState = (ALLMANDEFINED + SHOWMSG) Then
        '        imBoxNo = MKTNAMEINDEX
        '    End If
        '    mTestFields = NO
        '    Exit Function
        'End If

        If (ilCtrlNo = TESTALLCTRLS) Then
            slStr = smVehGp3
            slType = Trim$(edcTypeDropDown.Text)
            If (tgSpf.sMktBase = "Y") And (StrComp(slType, "NTR", vbTextCompare) <> 0) Then
                If StrComp(slStr, "[None]", vbTextCompare) = 0 Then
                    slStr = ""
                End If
            End If
            If gFieldDefinedStr(slStr, "", "Market must be specified", tmCtrls(MKTNAMEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = MKTNAMEINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        Else
            If gFieldDefinedCtrl(lbcVehGp3, "", "Market must be specified", tmCtrls(MKTNAMEINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = MKTNAMEINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = RSCHINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smVehGp5, "", "Research must be specified", tmCtrls(RSCHINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = RSCHINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        Else
            If gFieldDefinedCtrl(lbcVehGp5, "", "Research must be specified", tmCtrls(RSCHINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = RSCHINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = SUBCOMPINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smVehGp6, "", "Sub-Company must be specified", tmCtrls(SUBCOMPINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = SUBCOMPINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        Else
            If gFieldDefinedCtrl(lbcVehGp6, "", "Sub-Company must be specified", tmCtrls(SUBCOMPINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = SUBCOMPINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = ADDRESSINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcAddr(0), "", "Address must be specified", tmCtrls(ADDRESSINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = ADDRESSINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = PHONEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(mkcPhone, smPhoneImage, "Phone # must be specified", tmCtrls(PHONEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = PHONEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = FAXINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(mkcFax, smFaxImage, "Fax # must be specified", tmCtrls(FAXINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = FAXINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = FORMATINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        'If gFieldDefinedCtrl(edcFormat, "", "Format must be specified", tmCtrls(FORMATINDEX).iReq, ilState) = NO Then
        '    If ilState = (ALLMANDEFINED + SHOWMSG) Then
        '        imBoxNo = FORMATINDEX
        '    End If
        '    mTestFields = NO
        '    Exit Function
        'End If
        If (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smVehGp4, "", "Format must be specified", tmCtrls(FORMATINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = FORMATINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        Else
            If gFieldDefinedCtrl(lbcVehGp4, "", "Format must be specified", tmCtrls(FORMATINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = FORMATINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = SUBTOTALINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smVehGp2, "", "Subtotal must be specified", tmCtrls(SUBTOTALINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = SUBTOTALINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        Else
            If gFieldDefinedCtrl(lbcVehGp2, "", "Subtotal must be specified", tmCtrls(SUBTOTALINDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = SUBTOTALINDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (ilCtrlNo = DIALPOSINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDialPos, "", "Dial Position must be specified", tmCtrls(DIALPOSINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = DIALPOSINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = LOGVEHINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcLogVehDropDown, "", "Log Vehicle must be specified", tmCtrls(LOGVEHINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = LOGVEHINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = MULTIVEHLOGINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcMultiVehDropdown, "", "Multi-Vehicle Log must be specified", tmCtrls(MULTIVEHLOGINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = MULTIVEHLOGINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = RNLINKINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcRNLinkDropdown, "", "Rep-Net Link must be specified", tmCtrls(RNLINKINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = RNLINKINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcSort, "", "Sort must be specified", tmCtrls(SORTINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SORTINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    'If (ilCtrlNo = SORTINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
    '    ilSort = Val(edcSort.Txt)
    '    slNameCode = tmVehicle(imSelectedIndex - 1).sKey 'Traffic!lbcVehicle.List(ilSelectIndex - 1)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    ilCode = Trim$(slCode)
    '    If (ilSort <> 0) And (ilState = (ALLMANDEFINED + SHOWMSG)) Then
    '        For ilLoop = 0 To UBound(tgMVef) - 1 Step 1
    '            If (ilSort = tgMVef(ilLoop).iSort) And (ilCode <> tgMVef(ilLoop)) Then
    '                Screen.MousePointer = vbDefault  'Wait
    '                ilRes = MsgBox("Sort code previously used", vbOkOnly + vbExclamation, "Incomplete")
    '                imBoxNo = SORTINDEX
    '                mTestFields = NO
    '                Exit Function
    '            End If
    '        Next ilLoop
    '    End If
    'End If
    If (ilCtrlNo = SCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcStationCode, "", "Station Vehicle Code must be specified", tmCtrls(SCODEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = SCODEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = SCODEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        slStationCode = Trim$(edcStationCode.Text)
        If imSelectedIndex >= 1 Then
            slNameCode = tmVehicle(imSelectedIndex - 1).sKey 'Traffic!lbcVehicle.List(ilSelectIndex - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilCode = Trim$(slCode)
        Else
            ilCode = -1
        End If
        If (slStationCode <> "") And (ilState = (ALLMANDEFINED + SHOWMSG)) Then
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If (StrComp(Trim$(tgMVef(ilLoop).sCodeStn), slStationCode, vbTextCompare) = 0) And (ilCode <> tgMVef(ilLoop).iCode) Then
                    Screen.MousePointer = vbDefault  'Wait
                    ilRes = MsgBox("Station code previously used by " & Trim$(tgMVef(ilLoop).sName), vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = SCODEINDEX
                    mTestFields = NO
                    Exit Function
                End If
            Next ilLoop
        End If
    End If
    If (ilCtrlNo = BOOKINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcBookDropDown, "", "Rating Book Name must be specified", tmCtrls(BOOKINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = BOOKINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    'If tgSpf.sCAudPkg = "Y" Then
    '    If (ilCtrlNo = REALLINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
    '        If gFieldDefinedCtrl(edcReallDropDown, "", "Reallocation Book Name must be specified", tmCtrls(REALLINDEX).iReq, ilState) = NO Then
    '            If ilState = (ALLMANDEFINED + SHOWMSG) Then
    '                imBoxNo = REALLINDEX
    '            End If
    '            mTestFields = NO
    '            Exit Function
    '        End If
    '    End If
    'End If
    If (ilCtrlNo = DEMOINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDemoDropDown, "", "Prime Demo must be specified", tmCtrls(DEMOINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = DEMOINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = TYPEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcTypeDropDown, "", "Vehicle type must be specified", tmCtrls(TYPEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = TYPEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If ilState = (ALLMANDEFINED + SHOWMSG) Then 'SaveRec call
            slStr = Trim$(edcTypeDropDown.Text)
            If (StrComp(smOrigType, slStr, 1) <> 0) And (smOrigType <> "") Then
                Screen.MousePointer = vbDefault
                If (tgSpf.sGUseAffSys = "Y") And (imSelectedIndex > 0) Then
                    If (StrComp(smOrigType, "Conventional", 1) = 0) Or (StrComp(smOrigType, "Selling", 1) = 0) Then
                        ilRes = MsgBox("Affiliate System in Use, Agreements must be Changed", vbOKOnly + vbExclamation, "Warning")
'                        edcTypeDropDown.Text = smOrigType
'                        imBoxNo = TYPEINDEX
'                        mTestFields = NO
'                        Exit Function
                    End If
                End If
                If bmPrgLibDefined Then
                    slStr = Trim$(edcTypeDropDown.Text)
                    If (smOrigType = "Sport") Or (slStr = "Sport") Then
                        ilRes = MsgBox("Changing to or from Sport Vehicle type not allowed", vbOKOnly + vbExclamation, "Warning")
                        imBoxNo = TYPEINDEX
                        mTestFields = NO
                        Exit Function
                    ElseIf (smOrigType = "Airing") And (slStr = "Rep") Then
                        ilRes = MsgBox("Changing Airing to Rep Vehicle type not allowed", vbOKOnly + vbExclamation, "Warning")
                        imBoxNo = TYPEINDEX
                        mTestFields = NO
                        Exit Function
                    Else
                        ilCode = tmVef.iCode
                        ilRet = gIICodeRefExist(Vehicle, ilCode, "Clf.Btr", "ClfVefCode")  'clfvefCode
                        If ilRet Then
                            If (smOrigType = "Conventional") And ((slStr = "Rep") Or (slStr = "Selling")) Then
                                Screen.MousePointer = vbDefault  'Wait
                                ilRes = MsgBox("Contract Lines reference vehicle, Change Conventional anyway", vbYesNo + vbQuestion, "Invoice")
                                If ilRes = vbNo Then
                                    imBoxNo = TYPEINDEX
                                    mTestFields = NO
                                    Exit Function
                                End If
                            Else
                                Screen.MousePointer = vbDefault  'Wait
                                ilRes = MsgBox("Contract Lines reference vehicle, can't change Vehicle type", vbOKOnly + vbExclamation, "Incomplete")
                                imBoxNo = TYPEINDEX
                                mTestFields = NO
                                Exit Function
                            End If
                        End If
                    End If
                End If
                slStr = Trim$(edcTypeDropDown.Text)
                If slStr = "NTR" Then
                    ilRet = gIICodeRefExist(Vehicle, ilCode, "Rif.Btr", "RifVefCode")  'clfvefCode
                    If ilRet Then
                        Screen.MousePointer = vbDefault  'Wait
                        ilRes = MsgBox("Rate Card reference vehicle, can't change Vehicle type", vbOKOnly + vbExclamation, "Incomplete")
                        imBoxNo = TYPEINDEX
                        mTestFields = NO
                        Exit Function
                    End If
                End If
            End If
            'If (imType = 2) And ((imOrigType <> 2) And (imOrigType <> -1)) Then
            '    Screen.MousePointer = vbHourGlass  'Wait
            '    'Test if any unbilled spots
            '    tmSdfSrchKey1.iVefCode = tmVef.iCode
            '    tmSdfSrchKey1.iDate(0) = 0
            '    tmSdfSrchKey1.iDate(1) = 0
            '    tmSdfSrchKey1.iTime(0) = 0
            '    tmSdfSrchKey1.iTime(1) = 0
            '    tmSdfSrchKey1.sSchStatus = ""
            '    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            '    If (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmVef.iCode) Then
            '        Screen.MousePointer = vbDefault  'Wait
            '        ilRes = MsgBox("Spots exists, can't change Vehicle type to Airing", vbOkOnly + vbExclamation, "Incomplete")
            '        imBoxNo = TYPEINDEX
            '        mTestFields = NO
            '        Exit Function
            '    End If
            '    Screen.MousePointer = vbDefault
            'End If
            'If (imType = 3) And ((imOrigType <> 3) And (imOrigType <> -1)) Then
            '    Screen.MousePointer = vbHourGlass  'Wait
            '    'Test if any unbilled spots
            '    tmSdfSrchKey1.iVefCode = tmVef.iCode
            '    tmSdfSrchKey1.iDate(0) = 0
            '    tmSdfSrchKey1.iDate(1) = 0
            '    tmSdfSrchKey1.iTime(0) = 0
            '    tmSdfSrchKey1.iTime(1) = 0
            '    tmSdfSrchKey1.sSchStatus = ""
            '    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            '    If (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmVef.iCode) Then
            '        Screen.MousePointer = vbDefault  'Wait
            '        ilRes = MsgBox("Spots exists, can't change Vehicle type to Log", vbOkOnly + vbExclamation, "Incomplete")
            '        imBoxNo = TYPEINDEX
            '        mTestFields = NO
            '        Exit Function
            '    End If
            '    Screen.MousePointer = vbDefault
            'End If
        End If
    End If
    If (ilCtrlNo = STATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If imState = 0 Then
            slStr = "Active"
        ElseIf imState = 1 Then
            slStr = "Dormant"
        Else
            slStr = ""
        End If
        If gFieldDefinedStr(slStr, "", "Active Or Dormant must be specified", tmCtrls(STATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STATEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
        If ilState = (ALLMANDEFINED + SHOWMSG) Then 'SaveRec call
            ''Jim said remove as of 6/25/97 he was at ABC
            'Dormant vehicle are now included with all other vehicles on the invoice screen- remove test jim 9/22/98
            'If (imState = 1) And ((imOrigState <> 1) And (imOrigState <> -1)) And (Not igDemo) Then
            '    Screen.MousePointer = vbHourGlass  'Wait
            '    'Test if any unbilled spots
            '    tmSdfSrchKey1.iVefCode = tmVef.iCode
            '    tmSdfSrchKey1.iDate(0) = 0
            '    tmSdfSrchKey1.iDate(1) = 0
            '    tmSdfSrchKey1.iTime(0) = 0
            '    tmSdfSrchKey1.iTime(1) = 0
            '    tmSdfSrchKey1.sSchStatus = ""
            '    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            '    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmVef.iCode)
            '        If tmSdf.sBill = "N" Then
            '            'Ignore missed- billed flag not set
            '            If (tmSdf.sSchStatus <> "C") And (tmSdf.sSchStatus <> "H") Then
            '                If (tmSdf.sSpotType <> "S") And (tmSdf.sSpotType <> "M") Then
            '                    Screen.MousePointer = vbDefault  'Wait
            '                    ilRes = MsgBox("Unbilled spots exist and will not be Billed if changed to Dormant, Ok to Continue", vbYesNo + vbDefaultButton2 + vbQuestion, "Dormant")
            '                    If ilRes = vbNo Then
            '                        imBoxNo = STATEINDEX
            '                        mTestFields = NO
            '                        Exit Function
            '                    Else
            '                        Exit Do
            '                    End If
            '                End If
            '            End If
            '        End If
            '        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            '    Loop
            '    Screen.MousePointer = vbDefault
            'End If
        End If
    End If
    'If market specified, verify that market cluster is not expect for Rep vehicles.
    If (ilCtrlNo = TESTALLCTRLS) And (ilState = (ALLMANDEFINED + SHOWMSG)) Then
        slStr = edcTypeDropDown.Text
        gFindMatch slStr, 0, lbcType
        If gLastFound(lbcType) >= 0 Then
            slStr = lbcType.List(gLastFound(lbcType))
            If StrComp(slStr, "Rep", 1) <> 0 Then
                slStr = smVehGp3
                gFindMatch slStr, 2, lbcVehGp3
                If gLastFound(lbcVehGp3) > 1 Then
                    slNameCode = tmVehGp3Code(gLastFound(lbcVehGp3) - 2).sKey  'lbcVehGpCode.List(gLastFound(lbcVehGp) - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmVef.iMnfVehGp3Mkt = Val(slCode)
                    tmMnfSrchKey.iCode = Val(slCode)
                    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching demo name
                    If ilRet = BTRV_ERR_NONE Then
                        If Trim$(tmMnf.sRPU) = "Y" Then
                            ilRes = MsgBox("Market Cluster set to Yes which is not allowed with this vehicle type", vbOKOnly + vbExclamation, "Incomplete")
                            imBoxNo = MKTNAMEINDEX
                            mTestFields = NO
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    mTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp2Branch                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Vehicle group and process      *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mVehGp2Branch() As Integer
'
'   ilRet = mVehGpBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcVehGp2DropDown, lbcVehGp2, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcVehGp2DropDown.Text = "[None]") Then
        imDoubleClickName = False
        mVehGp2Branch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mVehGp2Branch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "H"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcVehGp2DropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\2"
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\2"
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\2"
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\2"
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mVehGp2Branch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcVehGp2.Clear
        smVehGp2CodeTag = ""
        mVehGp2Pop
        If imTerminate Then
            mVehGp2Branch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 2, lbcVehGp2
        sgMNmName = ""
        If gLastFound(lbcVehGp2) > 0 Then
            imVehGp2ChgMode = True
            lbcVehGp2.ListIndex = gLastFound(lbcVehGp2)
            edcVehGp2DropDown.Text = lbcVehGp2.List(lbcVehGp2.ListIndex)
            imVehGp2ChgMode = False
            mVehGp2Branch = False
            mSetChg imBoxNo
        Else
            imVehGp2ChgMode = True
            lbcVehGp2.ListIndex = 1
            edcVehGp2DropDown.Text = lbcVehGp2.List(1)
            imVehGp2ChgMode = False
            mSetChg imBoxNo
            edcVehGp2DropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGpPop                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mVehGp2Pop()
'
'   mVehGpPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcVehGp2.ListIndex
    If ilIndex > 1 Then
        slName = lbcVehGp2.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gIMoveListBox(Vehicle, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gPopMnfPlusFieldsBox(Vehicle, lbcVehGp2, tmVehGp2Code(), smVehGp2CodeTag, "H2")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehGp2PopErr
        gCPErrorMsg ilRet, "mVehGp2Pop (gPopMnfPlusFieldsBox)", Vehicle
        On Error GoTo 0
        lbcVehGp2.AddItem "[None]", 0  'Force as first item on list
        lbcVehGp2.AddItem "[New]", 0  'Force as first item on list
        imVehGp2ChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcVehGp2
            If gLastFound(lbcVehGp2) >= 2 Then
                lbcVehGp2.ListIndex = gLastFound(lbcVehGp2)
            Else
                lbcVehGp2.ListIndex = -1
            End If
        Else
            lbcVehGp2.ListIndex = ilIndex
        End If
        imVehGp2ChgMode = False
    End If
    Exit Sub
mVehGp2PopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp3Branch                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Vehicle group and process      *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mVehGp3Branch() As Integer
'
'   ilRet = mVehGp3Branch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcVehGp3DropDown, lbcVehGp3, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcVehGp3DropDown.Text = "[None]") Then
        imDoubleClickName = False
        mVehGp3Branch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mVehGp3Branch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "H"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcVehGp3DropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\3"
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\3"
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\3"
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\3"
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mVehGp3Branch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcVehGp3.Clear
        smVehGp3CodeTag = ""
        mVehGp3Pop
        If imTerminate Then
            mVehGp3Branch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 2, lbcVehGp3
        sgMNmName = ""
        If gLastFound(lbcVehGp3) > 0 Then
            imVehGp3ChgMode = True
            lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
            edcVehGp3DropDown.Text = lbcVehGp3.List(lbcVehGp3.ListIndex)
            imVehGp3ChgMode = False
            mVehGp3Branch = False
            mSetChg imBoxNo
        Else
            imVehGp3ChgMode = True
            lbcVehGp3.ListIndex = 1
            edcVehGp3DropDown.Text = lbcVehGp3.List(1)
            imVehGp3ChgMode = False
            mSetChg imBoxNo
            edcVehGp3DropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp3Pop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mVehGp3Pop()
'
'   mVehGpPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcVehGp3.ListIndex
    If ilIndex > 1 Then
        slName = lbcVehGp3.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gIMoveListBox(Vehicle, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gPopMnfPlusFieldsBox(Vehicle, lbcVehGp3, tmVehGp3Code(), smVehGp3CodeTag, "H3")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehGp3PopErr
        gCPErrorMsg ilRet, "mVehGp3Pop (gPopMnfPlusFieldsBox)", Vehicle
        On Error GoTo 0
        lbcVehGp3.AddItem "[None]", 0  'Force as first item on list
        lbcVehGp3.AddItem "[New]", 0  'Force as first item on list
        imVehGp3ChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcVehGp3
            If gLastFound(lbcVehGp3) >= 2 Then
                lbcVehGp3.ListIndex = gLastFound(lbcVehGp3)
            Else
                lbcVehGp3.ListIndex = -1
            End If
        Else
            lbcVehGp3.ListIndex = ilIndex
        End If
        imVehGp3ChgMode = False
    End If
    Exit Sub
mVehGp3PopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp4Branch                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Vehicle group and process      *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mVehGp4Branch() As Integer
'
'   ilRet = mVehGp4Branch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcVehGp4DropDown, lbcVehGp4, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcVehGp4DropDown.Text = "[None]") Then
        imDoubleClickName = False
        mVehGp4Branch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mVehGp4Branch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "H"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcVehGp4DropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\4"
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\4"
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\4"
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\4"
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mVehGp4Branch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcVehGp4.Clear
        smVehGp4CodeTag = ""
        mVehGp4Pop
        If imTerminate Then
            mVehGp4Branch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 2, lbcVehGp4
        sgMNmName = ""
        If gLastFound(lbcVehGp4) > 0 Then
            imVehGp4ChgMode = True
            lbcVehGp4.ListIndex = gLastFound(lbcVehGp4)
            edcVehGp4DropDown.Text = lbcVehGp4.List(lbcVehGp4.ListIndex)
            imVehGp4ChgMode = False
            mVehGp4Branch = False
            mSetChg imBoxNo
        Else
            imVehGp4ChgMode = True
            lbcVehGp4.ListIndex = 1
            edcVehGp4DropDown.Text = lbcVehGp4.List(1)
            imVehGp4ChgMode = False
            mSetChg imBoxNo
            edcVehGp4DropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp4Pop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mVehGp4Pop()
'
'   mVehGpPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcVehGp4.ListIndex
    If ilIndex > 1 Then
        slName = lbcVehGp4.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gIMoveListBox(Vehicle, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gPopMnfPlusFieldsBox(Vehicle, lbcVehGp4, tmVehGp4Code(), smVehGp4CodeTag, "H4")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehGp4PopErr
        gCPErrorMsg ilRet, "mVehGp4Pop (gPopMnfPlusFieldsBox)", Vehicle
        On Error GoTo 0
        lbcVehGp4.AddItem "[None]", 0  'Force as first item on list
        lbcVehGp4.AddItem "[New]", 0  'Force as first item on list
        imVehGp4ChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcVehGp4
            If gLastFound(lbcVehGp4) >= 2 Then
                lbcVehGp4.ListIndex = gLastFound(lbcVehGp4)
            Else
                lbcVehGp4.ListIndex = -1
            End If
        Else
            lbcVehGp4.ListIndex = ilIndex
        End If
        imVehGp4ChgMode = False
    End If
    Exit Sub
mVehGp4PopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp5Branch                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Vehicle group and process      *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mVehGp5Branch() As Integer
'
'   ilRet = mVehGp5Branch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcVehGp5DropDown, lbcVehGp5, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcVehGp5DropDown.Text = "[None]") Then
        imDoubleClickName = False
        mVehGp5Branch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mVehGp5Branch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "H"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcVehGp5DropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\5"
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\5"
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\5"
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\5"
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mVehGp5Branch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcVehGp5.Clear
        smVehGp5CodeTag = ""
        mVehGp5Pop
        If imTerminate Then
            mVehGp5Branch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 2, lbcVehGp5
        sgMNmName = ""
        If gLastFound(lbcVehGp5) > 0 Then
            imVehGp5ChgMode = True
            lbcVehGp5.ListIndex = gLastFound(lbcVehGp5)
            edcVehGp5DropDown.Text = lbcVehGp5.List(lbcVehGp5.ListIndex)
            imVehGp5ChgMode = False
            mVehGp5Branch = False
            mSetChg imBoxNo
        Else
            imVehGp5ChgMode = True
            lbcVehGp5.ListIndex = 1
            edcVehGp5DropDown.Text = lbcVehGp5.List(1)
            imVehGp5ChgMode = False
            mSetChg imBoxNo
            edcVehGp5DropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp5Pop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mVehGp5Pop()
'
'   mVehGpPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcVehGp5.ListIndex
    If ilIndex > 1 Then
        slName = lbcVehGp5.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gIMoveListBox(Vehicle, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gPopMnfPlusFieldsBox(Vehicle, lbcVehGp5, tmVehGp5Code(), smVehGp5CodeTag, "H5")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehGp5PopErr
        gCPErrorMsg ilRet, "mVehGp5Pop (gPopMnfPlusFieldsBox)", Vehicle
        On Error GoTo 0
        lbcVehGp5.AddItem "[None]", 0  'Force as first item on list
        lbcVehGp5.AddItem "[New]", 0  'Force as first item on list
        imVehGp5ChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcVehGp5
            If gLastFound(lbcVehGp5) >= 2 Then
                lbcVehGp5.ListIndex = gLastFound(lbcVehGp5)
            Else
                lbcVehGp5.ListIndex = -1
            End If
        Else
            lbcVehGp5.ListIndex = ilIndex
        End If
        imVehGp5ChgMode = False
    End If
    Exit Sub
mVehGp5PopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp6Branch                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Vehicle group and process      *
'*                      communication back from        *
'*                      invoice sorting                *
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
Private Function mVehGp6Branch() As Integer
'
'   ilRet = mVehGp6Branch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcVehGp6DropDown, lbcVehGp6, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcVehGp6DropDown.Text = "[None]") Then
        imDoubleClickName = False
        mVehGp6Branch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mVehGp6Branch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "H"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcVehGp6DropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\6"
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\6"
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\6"
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName & "\6"
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mVehGp6Branch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcVehGp6.Clear
        smVehGp6CodeTag = ""
        mVehGp6Pop
        If imTerminate Then
            mVehGp6Branch = False
            Exit Function
        End If
        If tgSpf.sSubCompany <> "Y" Then
            gFindMatch sgMNmName, 2, lbcVehGp6
        Else
            gFindMatch sgMNmName, 1, lbcVehGp6
        End If
        sgMNmName = ""
        If gLastFound(lbcVehGp6) > 0 Then
            imVehGp6ChgMode = True
            lbcVehGp6.ListIndex = gLastFound(lbcVehGp6)
            edcVehGp6DropDown.Text = lbcVehGp6.List(lbcVehGp6.ListIndex)
            imVehGp6ChgMode = False
            mVehGp6Branch = False
            mSetChg imBoxNo
        Else
            imVehGp6ChgMode = True
            lbcVehGp6.ListIndex = 1
            edcVehGp6DropDown.Text = lbcVehGp6.List(1)
            imVehGp6ChgMode = False
            mSetChg imBoxNo
            edcVehGp6DropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGp6Pop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mVehGp6Pop()
'
'   mVehGpPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcVehGp6.ListIndex
    If ilIndex >= 1 Then
        slName = lbcVehGp6.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gIMoveListBox(Vehicle, lbcVehGp, tmVehGpCode(), smVehGpCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gPopMnfPlusFieldsBox(Vehicle, lbcVehGp6, tmVehGp6Code(), smVehGp6CodeTag, "H6")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehGp6PopErr
        gCPErrorMsg ilRet, "mVehGp6Pop (gPopMnfPlusFieldsBox)", Vehicle
        On Error GoTo 0
        lbcVehGp6.AddItem "[New]", 0  'Force as first item on list
        If tgSpf.sSubCompany <> "Y" Then
            lbcVehGp6.AddItem "[None]", 1
        End If
        imVehGp6ChgMode = True
        If ilIndex >= 1 Then
            gFindMatch slName, 1, lbcVehGp6
            If gLastFound(lbcVehGp6) >= 1 Then
                lbcVehGp6.ListIndex = gLastFound(lbcVehGp6)
            Else
                lbcVehGp6.ListIndex = -1
            End If
        Else
            lbcVehGp6.ListIndex = ilIndex
        End If
        imVehGp6ChgMode = False
    End If
    Exit Sub
mVehGp6PopErr:
    On Error GoTo 0
    imTerminate = True
End Sub


Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub



Private Sub pbcSTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = SUBTOTALINDEX) Then
        If mVehGp2Branch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = HUBINDEX) Then
        If mHubBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = MKTNAMEINDEX) Then
        If mVehGp3Branch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = FORMATINDEX) Then
        If mVehGp4Branch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = RSCHINDEX) Then
        If mVehGp5Branch() Then
            Exit Sub
        End If
    End If
    'If tgSpf.sSubCompany = "Y" Then
        If (imBoxNo = SUBCOMPINDEX) Then
            If mVehGp6Branch() Then
                Exit Sub
            End If
        End If
    'End If
    If imBoxNo = DEMOINDEX Then
        If mDemoBranch() Then
            Exit Sub
        End If
    End If
    If imBoxNo = RNLINKINDEX Then
        If mRNLinkBranch() Then
            Exit Sub
        End If
    End If
    imTabDirection = -1  'Set-Right to left
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxCtrlNo) Then
        If (imBoxNo <> NAMEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Screen.MousePointer = vbDefault  'Wait
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
            Screen.MousePointer = vbDefault  'Wait
        End If
    End If
    Select Case imBoxNo
        Case -1
            imTabDirection = 0  'Set-Left to right
            If (igVehMode = 0) And (imFirstTimeSelect) Then
                Exit Sub
            End If
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            Else
                mSetChg 1
                ilBox = 2
            End If
        Case NAMEINDEX 'Name (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
        Case DIALPOSINDEX
            If sgRNCallType = "" Then
                If (smVehicleType = "C") Or (smVehicleType = "A") Or (smVehicleType = "G") Then
                    ilBox = MULTIVEHLOGINDEX
                Else
                    ilBox = LOGVEHINDEX
                End If
            Else
                If ((sgRNCallType = "N") And (smVehicleType = "R")) Or ((sgRNCallType = "R") And ((smVehicleType = "C") Or (smVehicleType = "S"))) Then
                    ilBox = RNLINKINDEX
                Else
                    If (smVehicleType = "C") Or (smVehicleType = "A") Or (smVehicleType = "G") Then
                        ilBox = MULTIVEHLOGINDEX
                    Else
                        ilBox = LOGVEHINDEX
                    End If
                End If
            End If
        Case RNLINKINDEX
            If (smVehicleType = "C") Or (smVehicleType = "A") Or (smVehicleType = "G") Then
                ilBox = MULTIVEHLOGINDEX
            Else
                ilBox = LOGVEHINDEX
            End If
        Case FORMATINDEX
            'If tgSpf.sSubCompany = "Y" Then
                ilBox = SUBCOMPINDEX
            'Else
            '    ilBox = RSCHINDEX
            'End If
        Case SCODEINDEX
            ilBox = DIALPOSINDEX
        Case STATEINDEX
            If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
                ilBox = HUBINDEX
            Else
                ilBox = SORTINDEX
            End If
        Case MKTNAMEINDEX
            If tgSpf.sAStnCodes = "N" Then
                ilBox = DIALPOSINDEX
            Else
                ilBox = SCODEINDEX
            End If
        Case DEMOINDEX
            'If tgSpf.sCAudPkg <> "Y" Then
                ilBox = BOOKINDEX
            'Else
            '    ilBox = REALLINDEX
            'End If
        Case SORTINDEX
            If imTaxDefined Then
                ilBox = TAXINDEX
            Else
                If imACT1CodesDefined And ((smVehicleType = "C") Or (smVehicleType = "S") Or (smVehicleType = "R") Or (smVehicleType = "G") Or (smVehicleType = "P")) Then
                    ilBox = ACT1CODESINDEX
                Else
                    ilBox = DEMOINDEX
                End If
            End If
        Case TAXINDEX
            If imACT1CodesDefined And ((smVehicleType = "C") Or (smVehicleType = "S") Or (smVehicleType = "R") Or (smVehicleType = "G") Or (smVehicleType = "P")) Then
                ilBox = ACT1CODESINDEX
            Else
                ilBox = DEMOINDEX
            End If
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcStartNew_GotFocus()
    Dim ilRet As Integer
    If imInNew Then
        Exit Sub
    End If
    If (igVehMode = 0) And (imFirstTimeSelect) Then
        imFirstTimeSelect = False
        ilRet = mStartNew()
        If Not ilRet Then
            imTerminate = True
            mTerminate
            Exit Sub
        End If
    End If
    mSetCommands
    '2/7/09: Added to handle case where focus can't be set
    On Error Resume Next
    pbcSTab.SetFocus
    On Error GoTo 0
End Sub
Private Sub pbcState_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If imState <> 0 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imState = 0
        pbcState_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If imState <> 1 Then
            tmCtrls(imBoxNo).iChg = True
        End If
        imState = 1
        pbcState_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If imState = 0 Then  'Active
            imState = 1
            tmCtrls(imBoxNo).iChg = True
            pbcState_Paint
        ElseIf imState = 1 Then  'Dormant
            tmCtrls(imBoxNo).iChg = True
            imState = 0  'Active
            pbcState_Paint
        End If
    End If
    mSetCommands
End Sub
Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imState = 0 Then  'Active
        tmCtrls(imBoxNo).iChg = True
        imState = 1  'Dormant
    ElseIf imState = 1 Then  'Dormant
        tmCtrls(imBoxNo).iChg = True
        imState = 0  'Active
    End If
    pbcState_Paint
    mSetCommands
End Sub
Private Sub pbcState_Paint()
    pbcState.Cls
    pbcState.CurrentX = fgBoxInsetX
    pbcState.CurrentY = 0 'fgBoxInsetY
    Select Case imState
        Case 0  'Active
            pbcState.Print "Active"
        Case 1  'Dormant
            pbcState.Print "Dormant"
        Case Else
            pbcState.Print "       "
    End Select
End Sub
Private Sub pbcTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    Dim ilBox As Integer
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = SUBTOTALINDEX) Then
        If mVehGp2Branch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = HUBINDEX) Then
        If mHubBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = MKTNAMEINDEX) Then
        If mVehGp3Branch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = FORMATINDEX) Then
        If mVehGp4Branch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = RSCHINDEX) Then
        If mVehGp5Branch() Then
            Exit Sub
        End If
    End If
    'If tgSpf.sSubCompany = "Y" Then
        If (imBoxNo = SUBCOMPINDEX) Then
            If mVehGp6Branch() Then
                Exit Sub
            End If
        End If
    'End If
    If imBoxNo = DEMOINDEX Then
        If mDemoBranch() Then
            Exit Sub
        End If
    End If
    If (imBoxNo = RNLINKINDEX) Then
        If mRNLinkBranch() Then
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imMaxCtrlNo) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Screen.MousePointer = vbDefault  'Wait
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
        Screen.MousePointer = vbDefault  'Wait
    End If
    Select Case imBoxNo
        Case -1
            imTabDirection = -1  'Set-Right to left
            ilBox = STATEINDEX   'imMaxCtrlNo
        Case ADDRESSINDEX  'Address
            If edcAddr(0).Text = "" Then
                edcAddr(1).Text = ""
                edcAddr(2).Text = ""
                edcAddr(3).Text = ""
                gSetShow pbcVeh, "", tmCtrls(ADDRESSINDEX + 1)
                gSetShow pbcVeh, "", tmCtrls(ADDRESSINDEX + 2)
                gSetShow pbcVeh, "", tmCtrls(ADDRESSINDEX + 3)
                pbcVeh.Cls
                pbcVeh_Paint
                ilBox = TYPEINDEX
            Else
                ilBox = imBoxNo + 1
            End If
        Case ADDRESSINDEX + 1 'Address
            If edcAddr(1).Text = "" Then
                edcAddr(2).Text = ""
                edcAddr(3).Text = ""
                gSetShow pbcVeh, "", tmCtrls(ADDRESSINDEX + 2)
                gSetShow pbcVeh, "", tmCtrls(ADDRESSINDEX + 3)
                pbcVeh.Cls
                pbcVeh_Paint
                ilBox = TYPEINDEX
            Else
                ilBox = imBoxNo + 1
            End If
        Case ADDRESSINDEX + 2 'Address
            If edcAddr(2).Text = "" Then
                edcAddr(3).Text = ""
                gSetShow pbcVeh, "", tmCtrls(ADDRESSINDEX + 3)
                pbcVeh.Cls
                pbcVeh_Paint
                ilBox = TYPEINDEX
            Else
                ilBox = imBoxNo + 1
            End If
        Case LOGVEHINDEX
            If (smVehicleType = "C") Or (smVehicleType = "A") Or (smVehicleType = "G") Then
                'ilBox = imBoxNo + 1
                ilBox = MULTIVEHLOGINDEX
            Else
                'ilBox = imBoxNo + 2
                If sgRNCallType = "" Then
                    ilBox = DIALPOSINDEX
                Else
                    If ((sgRNCallType = "N") And (smVehicleType = "R")) Or ((sgRNCallType = "R") And ((smVehicleType = "C") Or (smVehicleType = "S"))) Then
                        ilBox = RNLINKINDEX
                    Else
                        ilBox = DIALPOSINDEX
                    End If
                End If
            End If
        Case MULTIVEHLOGINDEX
            If sgRNCallType = "" Then
                ilBox = DIALPOSINDEX
            Else
                If ((sgRNCallType = "N") And (smVehicleType = "R")) Or ((sgRNCallType = "R") And ((smVehicleType = "C") Or (smVehicleType = "S"))) Then
                    ilBox = RNLINKINDEX
                Else
                    ilBox = DIALPOSINDEX
                End If
            End If
        Case RSCHINDEX
            'If tgSpf.sSubCompany = "Y" Then
                ilBox = SUBCOMPINDEX
            'Else
            '    ilBox = FORMATINDEX
            'End If
        Case DIALPOSINDEX
            If tgSpf.sAStnCodes = "N" Then
                ilBox = MKTNAMEINDEX
            Else
                ilBox = SCODEINDEX
            End If
        Case SORTINDEX
            If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
                ilBox = HUBINDEX
            Else
                ilBox = STATEINDEX
            End If
        Case BOOKINDEX
            'If tgSpf.sCAudPkg <> "Y" Then
                ilBox = DEMOINDEX
            'Else
            '    ilBox = REALLINDEX
            'End If
        Case DEMOINDEX
            If imACT1CodesDefined And ((smVehicleType = "C") Or (smVehicleType = "S") Or (smVehicleType = "R") Or (smVehicleType = "G") Or (smVehicleType = "P")) Then
                ilBox = ACT1CODESINDEX
            Else
                If imTaxDefined Then
                    ilBox = TAXINDEX
                Else
                    ilBox = SORTINDEX
                End If
            End If
        Case ACT1CODESINDEX
            If imTaxDefined Then
                ilBox = TAXINDEX
            Else
                ilBox = SORTINDEX
            End If
        Case imMaxCtrlNo 'last control
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcUpdate.Enabled) And (igVehCallSource = CALLNONE) Then
                cmcUpdate.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcVeh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    Dim ilBox As Integer
    Dim flAdj As Single
    If imBoxNo = NAMEINDEX Then
        If Not mOKName() Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To imMaxCtrlNo Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (ilBox = ADDRESSINDEX + 1) Or (ilBox = ADDRESSINDEX + 2) Or (ilBox = ADDRESSINDEX + 3) Then
                flAdj = fgBoxInsetY
            Else
                flAdj = 0
            End If
            If (Y >= tmCtrls(ilBox).fBoxY + flAdj) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH + flAdj) Then
                If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) <> USINGHUB Then
                    If (ilBox = HUBINDEX) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                End If
                If (ilBox = SCODEINDEX) And (tgSpf.sAStnCodes = "N") Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                'If (ilBox = REALLINDEX) And (tgSpf.sCAudPkg <> "Y") Then
                '    Beep
                '    mSetFocus imBoxNo
                '    Exit Sub
                'End If
                If (ilBox = TAXINDEX) And (Not imTaxDefined) Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                If (ilBox = ACT1CODESINDEX) And (Not imACT1CodesDefined) Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                If (ilBox = ACT1CODESINDEX) And ((smVehicleType <> "C") And (smVehicleType <> "S") And (smVehicleType <> "R") And (smVehicleType <> "G") And (smVehicleType <> "P")) Then
                    Beep
                    mSetFocus imBoxNo
                    Exit Sub
                End If
                If (ilBox = RNLINKINDEX) Then
                    If (sgRNCallType = "") Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    Else
                        If ((sgRNCallType = "N") And (smVehicleType = "R")) Or ((sgRNCallType = "R") And ((smVehicleType = "C") Or (smVehicleType = "S"))) Then
                        Else
                            Beep
                            mSetFocus imBoxNo
                            Exit Sub
                        End If
                    End If
                End If
                If (ilBox = MULTIVEHLOGINDEX) Then
                    If (smVehicleType <> "C") And (smVehicleType <> "A") And (smVehicleType <> "G") Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                End If
                'If (ilBox = SUBCOMPINDEX) And (tgSpf.sSubCompany <> "Y") Then
                '    Beep
                '    mSetFocus imBoxNo
                '    Exit Sub
                'End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcVeh_Paint()
    Dim ilBox As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    mPaintReall
    If imACT1CodesDefined And imTaxDefined Then
        pbcVeh.Line (tmCtrls(ACT1CODESINDEX).fBoxX - 15, tmCtrls(ACT1CODESINDEX).fBoxY - 15)-Step(tmCtrls(ACT1CODESINDEX).fBoxW + 15, tmCtrls(ACT1CODESINDEX).fBoxH + 15), BLUE, B
    End If
    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
        llColor = pbcVeh.ForeColor
        slFontName = pbcVeh.FontName
        flFontSize = pbcVeh.FontSize
        pbcVeh.ForeColor = BLUE
        pbcVeh.FontBold = False
        pbcVeh.FontSize = 7
        pbcVeh.FontName = "Arial"
        pbcVeh.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcVeh.CurrentX = tmCtrls(HUBINDEX).fBoxX + 15  'fgBoxInsetX
        pbcVeh.CurrentY = tmCtrls(HUBINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcVeh.Print "Hub"
        pbcVeh.FontSize = flFontSize
        pbcVeh.FontName = slFontName
        pbcVeh.FontSize = flFontSize
        pbcVeh.ForeColor = llColor
        pbcVeh.FontBold = True
    End If
    If imTaxDefined Then
        llColor = pbcVeh.ForeColor
        slFontName = pbcVeh.FontName
        flFontSize = pbcVeh.FontSize
        pbcVeh.ForeColor = BLUE
        pbcVeh.FontBold = False
        pbcVeh.FontSize = 7
        pbcVeh.FontName = "Arial"
        pbcVeh.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcVeh.CurrentX = tmCtrls(TAXINDEX).fBoxX + 15  'fgBoxInsetX
        pbcVeh.CurrentY = tmCtrls(TAXINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcVeh.Print "Tax"
        pbcVeh.FontSize = flFontSize
        pbcVeh.FontName = slFontName
        pbcVeh.FontSize = flFontSize
        pbcVeh.ForeColor = llColor
        pbcVeh.FontBold = True
    End If
    If imACT1CodesDefined Then
        llColor = pbcVeh.ForeColor
        slFontName = pbcVeh.FontName
        flFontSize = pbcVeh.FontSize
        pbcVeh.ForeColor = BLUE
        pbcVeh.FontBold = False
        pbcVeh.FontSize = 7
        pbcVeh.FontName = "Arial"
        pbcVeh.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcVeh.CurrentX = tmCtrls(ACT1CODESINDEX).fBoxX + 15  'fgBoxInsetX
        pbcVeh.CurrentY = tmCtrls(ACT1CODESINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcVeh.Print "ACT1 Lineup"
        pbcVeh.FontSize = flFontSize
        pbcVeh.FontName = slFontName
        pbcVeh.FontSize = flFontSize
        pbcVeh.ForeColor = llColor
        pbcVeh.FontBold = True
    End If
    For ilBox = imLBCtrls To imMaxCtrlNo Step 1
        If (ilBox = SUBTOTALINDEX) Then
            gPaintArea pbcVeh, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + fgOffset - 15, tmCtrls(ilBox).fBoxW - 15, fgBoxGridH, WHITE
        End If
        pbcVeh.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcVeh.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcVeh.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub

Private Sub rbcType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcType(Index).Value
    'End of coded added
    mPopulate
End Sub
Private Sub rbcType_GotFocus(Index As Integer)
    If imFirstFocus Then
        cbcSelect.SetFocus
    End If
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imBoxNo
        Case HUBINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcHub, edcHubDropdown, imHubChgMode, imLbcArrowSetting
        Case MKTNAMEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcVehGp3, edcVehGp3DropDown, imVehGp3ChgMode, imLbcArrowSetting
        Case RSCHINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcVehGp5, edcVehGp5DropDown, imVehGp5ChgMode, imLbcArrowSetting
        Case SUBCOMPINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcVehGp6, edcVehGp6DropDown, imVehGp6ChgMode, imLbcArrowSetting
        Case FORMATINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcVehGp4, edcVehGp4DropDown, imVehGp4ChgMode, imLbcArrowSetting
        Case SUBTOTALINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcVehGp2, edcVehGp2DropDown, imVehGp2ChgMode, imLbcArrowSetting
        Case DEMOINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcDemo, edcDemoDropDown, imDemoChgMode, imLbcArrowSetting
        Case RNLINKINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcRNLink, edcRNLinkDropdown, imRNLinkChgMode, imLbcArrowSetting
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMultiConvVehLogPop             *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mMultiConvVehLogPop()
    Dim ilRet As Integer
    Dim llFilter As Long
    llFilter = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH + DORMANTVEH
    ilRet = gPopUserVehicleBox(Vehicle, llFilter, lbcMultiConvVehLog, tmMultiConvVehLogCode(), smMultiVehLogCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMultiConvVehLogPopErr
        gCPErrorMsg ilRet, "mMultiConvVehLogPop (gPopUserVehicleBox: Vehicle)", Vehicle
        On Error GoTo 0
        lbcMultiConvVehLog.AddItem "[None]", 0  'Force as first item on list
    End If
    Exit Sub
mMultiConvVehLogPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mVehGpPop                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Vehicle Group Code    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mHubPop()
'
'   mVehGpPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcHub.ListIndex
    If ilIndex > 0 Then
        slName = lbcHub.List(ilIndex)
    End If
    ilfilter(0) = CHARFILTER
    slFilter(0) = "W"
    ilOffset(0) = gFieldOffset("Mnf", "MnfType") '2
    'ilRet = gIMoveListBox(Vehicle, lbcDemo, lbcDemoCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Vehicle, lbcHub, tmHubCode(), smHubCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mHubPopErr
        gCPErrorMsg ilRet, "mHubPop (gIMoveListBox)", Vehicle
        On Error GoTo 0
        lbcHub.AddItem "[None]", 0  'Force as first item on list
        lbcHub.AddItem "[New]", 0  'Force as first item on list
        imHubChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcHub
            If gLastFound(lbcHub) >= 1 Then
                lbcHub.ListIndex = gLastFound(lbcHub)
            Else
                lbcHub.ListIndex = -1
            End If
        Else
            lbcHub.ListIndex = ilIndex
        End If
        imHubChgMode = False
    End If
    Exit Sub
mHubPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mHubBranch                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Hub and process               *
'*                      communication back from        *
'*                      Hub                           *
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
Private Function mHubBranch() As Integer
'
'   ilRet = mHubBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcHubDropdown, lbcHub, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcHubDropdown.Text = "[None]") Then
        imDoubleClickName = False
        mHubBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(INVOICESORTLIST)) Then
    '    imDoubleClickName = False
    '    mHubBranch = True
    '    mEnableBox imBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "W"
    igMNmCallSource = CALLSOURCEVEHICLE
    If edcHubDropdown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Vehicle^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Vehicle^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Vehicle^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Vehicle^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Vehicle.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Vehicle.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mHubBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcHub.Clear
        smHubCodeTag = ""
        mHubPop
        If imTerminate Then
            mHubBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcHub
        sgMNmName = ""
        If gLastFound(lbcHub) > 0 Then
            imHubChgMode = True
            lbcHub.ListIndex = gLastFound(lbcHub)
            edcHubDropdown.Text = lbcHub.List(lbcHub.ListIndex)
            imHubChgMode = False
            mHubBranch = False
            mSetChg imBoxNo
        Else
            imHubChgMode = True
            If lbcHub.ListCount > 1 Then
                lbcHub.ListIndex = 1
                edcHubDropdown.Text = lbcHub.List(1)
            Else
                lbcHub.ListIndex = 0
                edcHubDropdown.Text = lbcHub.List(0)
            End If
            imHubChgMode = False
            mSetChg imBoxNo
            edcHubDropdown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMultiGameVehLogPop             *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mMultiGameVehLogPop()
    Dim ilRet As Integer
    Dim llFilter As Long
    llFilter = VEHSPORT + ACTIVEVEH + DORMANTVEH
    ilRet = gPopUserVehicleBox(Vehicle, llFilter, lbcMultiGameVehLog, tmMultiGameVehLogCode(), smMultiGameVehLogCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mMultiGameVehLogPopErr
        gCPErrorMsg ilRet, "mMultiGameVehLogPop (gPopUserVehicleBox: Vehicle)", Vehicle
        On Error GoTo 0
        lbcMultiGameVehLog.AddItem "[None]", 0  'Force as first item on list
    End If
    Exit Sub
mMultiGameVehLogPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Function mAddPifModel() As Integer
    Dim hlPif As Integer        'site Option file handle
    Dim ilRecLen As Integer     'Vpf record length
    Dim tlPif As PIF
    Dim tlSrchKey1 As PIFKEY1
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llNowDate As Long

    llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    hlPif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hlPif, "", sgDBPath & "Pif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mAddPifModel = False
        Exit Function
    End If
    ilRecLen = Len(tlPif)  'btrRecordLength(hlVpf)  'Get and save record length
    tlSrchKey1.iVefCode = igVefCodeModel
    gPackDate "", tlSrchKey1.iStartDate(0), tlSrchKey1.iStartDate(1)
    tlSrchKey1.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hlPif, tlPif, ilRecLen, tlSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlPif.iVefCode = igVefCodeModel)
        gUnpackDateLong tlPif.iEndDate(0), tlPif.iEndDate(1), llDate
        If llDate >= llNowDate Then
            tlPif.lCode = 0
            tlPif.iVefCode = tmVef.iCode
            ilRet = btrInsert(hlPif, tlPif, ilRecLen, INDEXKEY0)
        End If
        ilRet = btrGetNext(hlPif, tlPif, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hlPif)
    btrDestroy hlPif
    mAddPifModel = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mRNLinkBranch                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to        *
'*                      Demo and process               *
'*                      communication back from        *
'*                      Demo                           *
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
Private Function mRNLinkBranch() As Integer
'
'   ilRet = mRNLinkBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer

    ilRet = gOptionalLookAhead(edcRNLinkDropdown, lbcRNLink, imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (edcRNLinkDropdown.Text = "[None]") Then
        imDoubleClickName = False
        mRNLinkBranch = False
        Exit Function
    End If
    igRNCallSource = CALLSOURCEVEHICLE
    If edcRNLinkDropdown.Text = "[New]" Then
        sgRNName = ""
    Else
        sgRNName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed

    If igTestSystem Then
        slStr = "Vehicle^Test\" & sgUserName & "\" & sgRNCallType & "\" & Trim$(Str$(igRNCallSource)) & "\" & sgRNName
    Else
        slStr = "Vehicle^Prod\" & sgUserName & "\" & sgRNCallType & "\" & Trim$(Str$(igRNCallSource)) & "\" & sgRNName
    End If
    sgCommandStr = slStr
    RepNet.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgRNName)
    igRNCallSource = Val(sgRNName)
    ilParse = gParseItem(slStr, 2, "\", sgRNName)

    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mRNLinkBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igRNCallSource = CALLDONE Then  'Done
        igRNCallSource = CALLNONE
'        gSetMenuState True
        smRNLinkCodeTag = ""
        lbcRNLink.Clear
        mRNLinkPop
        If imTerminate Then
            mRNLinkBranch = False
            Exit Function
        End If
        gFindMatch sgRNName, 1, lbcRNLink
        sgRNName = ""
        If gLastFound(lbcRNLink) > 0 Then
            imRNLinkChgMode = True
            lbcRNLink.ListIndex = gLastFound(lbcRNLink)
            edcRNLinkDropdown.Text = lbcRNLink.List(lbcRNLink.ListIndex)
            imRNLinkChgMode = False
            mRNLinkBranch = False
            mSetChg imBoxNo
        Else
            imRNLinkChgMode = True
            lbcRNLink.ListIndex = 1
            edcRNLinkDropdown.Text = lbcRNLink.List(1)
            imRNLinkChgMode = False
            mSetChg imBoxNo
            edcRNLinkDropdown.SetFocus
            Exit Function
        End If
    End If
    If igRNCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igRNCallSource = CALLNONE
        sgRNName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    If igRNCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igRNCallSource = CALLNONE
        sgRNName = ""
        mEnableBox imBoxNo
        Exit Function
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRNLinkPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Demo List box if      *
'*                      required                       *
'*                                                     *
'*******************************************************
Private Sub mRNLinkPop()
'
'   mDemoPop
'   Where:
'
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcRNLink.ListIndex
    If ilIndex > 1 Then
        slName = lbcRNLink.List(ilIndex)
    End If
    ilfilter(0) = CHARFILTER
    slFilter(0) = sgRNCallType
    ilOffset(0) = gFieldOffset("NRF", "nrfType") '2
    ilRet = gIMoveListBox(Vehicle, lbcRNLink, tmRNLinkCode(), smRNLinkCodeTag, "Nrf.btr", gFieldOffset("Nrf", "NrfName"), 30, ilfilter(), slFilter(), ilOffset())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mRNLinkPopErr
        gCPErrorMsg ilRet, "mRNLinkPop (gIMoveListBox)", Vehicle
        On Error GoTo 0
        lbcRNLink.AddItem "[None]", 0  'Force as first item on list
        lbcRNLink.AddItem "[New]", 0  'Force as first item on list
        imRNLinkChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, lbcRNLink
            If gLastFound(lbcRNLink) > 1 Then
                lbcRNLink.ListIndex = gLastFound(lbcRNLink)
            Else
                lbcRNLink.ListIndex = -1
            End If
        Else
            lbcRNLink.ListIndex = ilIndex
        End If
        imRNLinkChgMode = False
    End If
    Exit Sub
mRNLinkPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Private Function mDetermineIfPrgDefined() As Boolean
    Dim ilRet As Integer
    
    mDetermineIfPrgDefined = False
    If imSelectedIndex = 0 Then
        Exit Function
    End If
    tmLcfSrchKey2.iVefCode = tmVef.iCode
    tmLcfSrchKey2.iLogDate(0) = 0   'ilStartDate(0)
    tmLcfSrchKey2.iLogDate(1) = 0   'ilStartDate(1)
    ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
    If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = tmVef.iCode) Then
        mDetermineIfPrgDefined = True
    Else
        '5/7/11: Test RLF instead of Clf.  Not as good as testing Clf but it is faster
        '        Lines can't be added without rif record created
        '        However, user could remove rif with clf records created referencing the vehicle
        'ilRet = gIICodeRefExist(Vehicle, tmVef.iCode, "Clf.Btr", "ClfVefCode")  'clfvefCode
        'If ilRet Then
        '    mDetermineIfPrgDefined = True
        'Else
        '    ilRet = gIICodeRefExist(Vehicle, tmVef.iCode, "Sbf.Btr", "SbfBillVefCode")
        '    If ilRet Then
        '        mDetermineIfPrgDefined = True
        '    End If
        'End If
        ilRet = gIICodeRefExist(Vehicle, tmVef.iCode, "Rif.Btr", "RifVefCode")  'clfvefCode
        If ilRet Then
            mDetermineIfPrgDefined = True
        End If
    End If

End Function


Private Sub mAddVff()
    'dan note: if it fails, didn't find anything--are we in 'modelling'? Then alternate to intiailizing
    'insert both cases and vefcode

    Dim ilRet As Integer
    '10050
    Dim tlVFFFromModelled As VFF
    'Add Record
    tmVff.iCode = 0
    tmVff.iVefCode = tmVef.iCode
    tmVff.sGroupName = ""
    tmVff.sWegenerExportID = ""
    tmVff.sOLAExportID = ""
    tmVff.iLiveCompliantAdj = 5
    tmVff.iUstCode = 0
    tmVff.iUrfCode = tgUrf(0).iCode
    'tmVff.sXDXMLForm = "S"
    If ((Asc(tgSpf.sUsingFeatures8) And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT) Then
        tmVff.sXDXMLForm = "S"
    Else
        If ((Asc(tgSpf.sUsingFeatures7) And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT) Then
            tmVff.sXDXMLForm = "P"
        Else
            tmVff.sXDXMLForm = ""
        End If
    End If
    tmVff.sXDISCIPrefix = ""
    tmVff.sXDProgCodeID = ""
    tmVff.sXDSaveCF = "Y"
    tmVff.sXDSaveHDD = "N"
    tmVff.sXDSaveNAS = "N"
    'tmVff.sUnused = ""
    tmVff.iCwfCode = 0
    tmVff.sAirWavePrgID = ""
    tmVff.sExportAirWave = ""
    tmVff.sExportNYESPN = ""
    tmVff.sPledgeVsAir = "N"
    tmVff.sFedDelivery(0) = ""
    tmVff.sFedDelivery(1) = ""
    tmVff.sFedDelivery(2) = ""
    tmVff.sFedDelivery(3) = ""
    tmVff.sFedDelivery(4) = ""
    'tmVff.sFedDelivery(5) = ""
    gPackDate "1/1/1990", tmVff.iLastAffExptDate(0), tmVff.iLastAffExptDate(1)
    tmVff.sMoveSportToNon = "N"
    tmVff.sMoveSportToSport = "N"
    tmVff.sMoveNonToSport = "N"
    tmVff.sMergeTraffic = "S"
    tmVff.sMergeAffiliate = "S"
    tmVff.sMergeWeb = "S"
    tmVff.sPledgeByEvent = "N"
    tmVff.lPledgeHdVtfCode = 0
    tmVff.lPledgeFtVtfCode = 0
    tmVff.iPledgeClearance = 0
    tmVff.sExportEncoESPN = "N"
    tmVff.sWebName = ""
    tmVff.lSeasonGhfCode = 0
    tmVff.iMcfCode = 0
    tmVff.sExportAudio = "N"
    tmVff.sExportMP2 = "N"
    tmVff.sExportCnCSpot = "N"
    tmVff.sExportEnco = "N"
    tmVff.sExportCnCNetInv = "N"
    tmVff.sIPumpEventTypeOV = ""
    tmVff.sExportIPump = "N"
    tmVff.sAddr4 = ""
    tmVff.lBBOpenCefCode = 0
    tmVff.lBBCloseCefCode = 0
    tmVff.lBBBothCefCode = 0
    tmVff.sXDSISCIPrefix = ""
    tmVff.sXDSSaveCF = "Y"
    tmVff.sXDSSaveHDD = "N"
    tmVff.sXDSSaveNAS = "N"
    tmVff.sMGsOnWeb = "N"   '"Y"
    tmVff.sReplacementOnWeb = "N"   '"Y"
    tmVff.sExportMatrix = "N"
    tmVff.sSentToXDSStatus = "N"
    tmVff.sStationComp = "N"
    tmVff.sExportSalesForce = "N"
    tmVff.sExportEfficio = "N"
    tmVff.sExportJelli = "N"
    tmVff.sOnXMLInsertion = "N"
    tmVff.sOnInsertions = "N"
    tmVff.sPostLogSource = "N"
    tmVff.sExportTableau = "N"
    tmVff.sStationPassword = ""
    tmVff.sHonorZeroUnits = "N"
    tmVff.sHideCommOnLog = "N"
    tmVff.sHideCommOnWeb = "N"
    tmVff.iConflictWinLen = 0
    tmVff.sACT1LineupCode = ""
    tmVff.sPrgmmaticAllow = "N"
    tmVff.sSalesBrochure = ""
    tmVff.sCartOnWeb = "N"
    tmVff.sDefaultAudioType = "R"
    tmVff.iLogExptArfCode = 0
    'tmVff.sUnused = ""
    tmVff.sASICallLetters = ""
    tmVff.sASIBand = ""
    tmVff.sExportCustom = "" 'TTP 9992
    '10050 modelling?
    If igVefCodeModel > 0 Then
        tmVffSrchKey1.iCode = igVefCodeModel
        ilRet = btrGetEqual(hmVff, tlVFFFromModelled, imVffRecLen, tmVffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tmVff.iAvfCode = tlVFFFromModelled.iAvfCode
            '10981
'            tmVff.lAdVehNameCefCode = tlVFFFromModelled.lAdVehNameCefCode
            tmVff.lAdVehNameCefCode = 0
        End If
    End If
    ilRet = btrInsert(hmVff, tmVff, imVffRecLen, INDEXKEY0)
End Sub

Private Function mCheckPvfVefCode(ilVefCode As Integer) As Boolean
    Dim pvf_rst As ADODB.Recordset
    Dim slSQLQuery As String
    Dim ilLoop As Integer
    Dim slStr As String
    
    On Error Resume Next
    mCheckPvfVefCode = False
    slStr = "("
    For ilLoop = 1 To 25 Step 1
        slStr = slStr & "CASE pvfVefCode" & ilLoop & " WHEN " & ilVefCode & " THEN 1 ELSE 0 END + "
    Next ilLoop
    slStr = Left(slStr, Len(slStr) - 2) & ") > 0"
    slSQLQuery = "SELECT Count(1) as TestCount FROM pvf_Package_Vehicle WHERE " & slStr
    Set pvf_rst = gSQLSelectCall(slSQLQuery)
    If Not pvf_rst.EOF Then
        If pvf_rst!TestCount > 0 Then
            mCheckPvfVefCode = True
        End If
    End If
    pvf_rst.Close
End Function
'10071
Private Sub mTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    gInitGlobalVar
    hlSpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    ilRecLen = Len(tlSpf)
    ilRet = btrGetFirst(hlSpf, tlSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    btrDestroy hlSpf
End Sub
Private Sub mCheckForDate()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    Dim slSetDate As String
    Dim ilRet As Integer

    ilPos = InStr(1, sgCommandStr, "/D:", 1)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gValidDate(slDate) Then
            slDate = gAdjYear(slDate)
            slSetDate = slDate
        End If
    End If
    If Trim$(slSetDate) = "" Then
        If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Then
            slSetDate = "12/15/1999"
            slDate = slSetDate
        End If
    End If
    If Trim$(slSetDate) <> "" Then
        ilRet = gCsiSetName(slDate)
    End If
End Sub
Private Sub mGetUlfCode()
    Dim ilPos As Integer
    Dim ilSpace As Integer

    ilPos = InStr(1, sgCommandStr, "/ULF:", 1)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5)))
        Else
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5, ilSpace - ilPos - 3)))
        End If
    End If
End Sub
Private Sub mInitStdAloneVehicle()
    'set sgusername before calling?
    
    Dim ilRet As Integer
    Dim slVehType As String
    Dim hlUrf As Integer        'User Option file handle
    Dim tlUrf As URF
    Dim ilRecLen As Integer

    If igStdAloneMode Then
        ilRet = csiSetAlloc("NAMES", 0, 2)
    End If
    sgSystemDate = gAdjYear(Format$(gNow(), "m/d/yy"))    'Used to reset date when exiting traffic
    igResetSystemDate = False
   ' gInitGlobalVar   'Initialize global variables
    ReDim tgJobHelp(0 To 0) As HLF
    ReDim tgListHelp(0 To 0) As HLF
    gSpfRead
    sgCPName = gGetCSIName("CPNAME")
    sgSUName = gGetCSIName("SUNAME")
    If (sgCPName = "") Or (sgSUName = "") Then
        hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            ilRecLen = Len(tlUrf)  'btrRecordLength(hlUrf)  'Get and save record length
            ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                gUrfDecrypt tlUrf
                If tlUrf.iCode = 1 Then
                    sgCPName = Trim$(tlUrf.sName)
                End If
                ilRet = btrGetNext(hlUrf, tlUrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUrfDecrypt tlUrf
                    If tlUrf.iCode = 2 Then
                        sgSUName = Trim$(tlUrf.sName)
                    End If
                End If
            End If
        End If
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
    End If
    igUpdateAllowed = True
    'gUrfRead Vehicle, sgUserName, True, tgUrf(), False  'Obtain user records
    'sgUserName = Trim$(tgUrf(0).sName)
    'gInitSuperUser tgUrf(0)
End Sub
Private Sub mPrepVehicle()
    gObtainSAF
    gVpfRead
End Sub
Private Function mDetermineCsiLogin() As String
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slRet As String
    
    slDate = Format$(Now(), "m/d/yy")
    slMonth = Month(slDate)
    slYear = Year(slDate)
    llValue = Val(slMonth) * Val(slYear)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    llValue = ilValue
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slRet = Trim$(Str$(ilValue))
    Do While Len(slRet) < 4
        slRet = "0" & slRet
    Loop
    mDetermineCsiLogin = slRet
End Function
Private Sub mSetAllowUpdate()
    If (igWinStatus(VEHICLESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcVeh.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcVeh.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Vehicle.Refresh
    mSetCommands
End Sub
