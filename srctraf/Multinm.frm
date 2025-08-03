VERSION 5.00
Begin VB.Form MultiNm 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   615
   ClientTop       =   1935
   ClientWidth     =   4845
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
   FontTransparent =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   10
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   4845
   Begin VB.Timer tmcHide 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   510
      Top             =   2550
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   750
      TabIndex        =   1
      Text            =   "cbcSelect"
      Top             =   330
      Width           =   3210
   End
   Begin VB.ListBox lbcDemos 
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
      Left            =   4485
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   285
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcCtrl5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2085
      MaxLength       =   10
      TabIndex        =   28
      Top             =   1110
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   1995
      Picture         =   "Multinm.frx":0000
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1140
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   33
      Top             =   690
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox edcCtrl4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2340
      MaxLength       =   10
      TabIndex        =   27
      Top             =   750
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcCtrl3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   3210
      MaxLength       =   10
      TabIndex        =   26
      Top             =   195
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbcOrigin 
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
      Left            =   3405
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3300
      ScaleHeight     =   210
      ScaleWidth      =   1365
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox edcName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   1935
      MaxLength       =   10
      TabIndex        =   23
      Top             =   150
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.TextBox edcSlspComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2910
      MaxLength       =   10
      TabIndex        =   25
      Top             =   615
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3675
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox edcRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2610
      MaxLength       =   10
      TabIndex        =   24
      Top             =   795
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3285
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2730
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   30
      ScaleHeight     =   75
      ScaleWidth      =   15
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1845
      Width           =   15
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
      HelpContextID   =   1
      Left            =   975
      TabIndex        =   35
      Top             =   1980
      Width           =   945
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
      HelpContextID   =   2
      Left            =   1995
      TabIndex        =   36
      Top             =   1980
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      HelpContextID   =   3
      Left            =   3015
      TabIndex        =   37
      Top             =   1980
      Width           =   945
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
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
      HelpContextID   =   4
      Left            =   1455
      TabIndex        =   38
      Top             =   2340
      Width           =   945
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
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
      HelpContextID   =   5
      Left            =   2475
      TabIndex        =   39
      Top             =   2340
      Width           =   945
   End
   Begin VB.CommandButton cmcMerge 
      Appearance      =   0  'Flat
      Caption         =   "&Merge into"
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
      HelpContextID   =   6
      Left            =   4500
      TabIndex        =   40
      Top             =   480
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   75
      ScaleHeight     =   30
      ScaleWidth      =   30
      TabIndex        =   34
      Top             =   1710
      Width           =   30
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   1695
      ScaleHeight     =   30
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   165
      Width           =   15
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   1065
      Index           =   0
      Left            =   2370
      Picture         =   "Multinm.frx":00FA
      ScaleHeight     =   1065
      ScaleWidth      =   2850
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   2850
      Begin VB.Label lacCover 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2280
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3390
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4530
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   810
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   735
      Index           =   11
      Left            =   135
      Picture         =   "Multinm.frx":A33C
      ScaleHeight     =   735
      ScaleWidth      =   2850
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   17
      Left            =   285
      Picture         =   "Multinm.frx":C88E
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   18
      Left            =   450
      Picture         =   "Multinm.frx":15B0C
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   1080
      Index           =   10
      Left            =   3930
      Picture         =   "Multinm.frx":1C98E
      ScaleHeight     =   1080
      ScaleWidth      =   2850
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   1065
      Index           =   4
      Left            =   750
      Picture         =   "Multinm.frx":20020
      ScaleHeight     =   1065
      ScaleWidth      =   2850
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1815
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   15
      Left            =   3690
      Picture         =   "Multinm.frx":29BB2
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1770
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   16
      Left            =   3600
      Picture         =   "Multinm.frx":2C104
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1695
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   1065
      Index           =   14
      Left            =   4200
      Picture         =   "Multinm.frx":2E656
      ScaleHeight     =   1065
      ScaleWidth      =   2850
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1935
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   13
      Left            =   3525
      Picture         =   "Multinm.frx":31CE8
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   1065
      Index           =   3
      Left            =   3885
      Picture         =   "Multinm.frx":3423A
      ScaleHeight     =   1065
      ScaleWidth      =   2850
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   1065
      Index           =   7
      Left            =   3825
      Picture         =   "Multinm.frx":4195C
      ScaleHeight     =   1065
      ScaleWidth      =   2850
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   9
      Left            =   120
      Picture         =   "Multinm.frx":44FEE
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1665
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   5
      Left            =   300
      Picture         =   "Multinm.frx":47540
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   6
      Left            =   3750
      Picture         =   "Multinm.frx":48822
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1410
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   12
      Left            =   390
      Picture         =   "Multinm.frx":49B04
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1665
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   1
      Left            =   3660
      Picture         =   "Multinm.frx":50986
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1365
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   8
      Left            =   3615
      Picture         =   "Multinm.frx":52ED8
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   375
      Index           =   2
      Left            =   240
      Picture         =   "Multinm.frx":5542A
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Height          =   375
      Index           =   19
      Left            =   945
      Picture         =   "Multinm.frx":55E6C
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   20
      Left            =   15
      Picture         =   "Multinm.frx":5992E
      ScaleHeight     =   720
      ScaleWidth      =   2850
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox pbcMNm 
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
      Index           =   21
      Left            =   1230
      Picture         =   "Multinm.frx":601B8
      ScaleHeight     =   720
      ScaleWidth      =   2835
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   825
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.PictureBox plcMNm 
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   825
      ScaleHeight     =   1140
      ScaleWidth      =   2910
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   2970
   End
   Begin VB.Label plcScreen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2955
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   90
      Top             =   2340
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imcUnused 
      Appearance      =   0  'Flat
      Height          =   15
      Left            =   30
      Top             =   1770
      Width           =   30
   End
End
Attribute VB_Name = "MultiNm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Multinm.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: MultiNm.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the NTR Type, Annoucer, Sports,
'   Sales Source, Sales Team, Revenue sets, Missed reason
'   product protection and network input screen code
Option Explicit
Option Compare Text
Dim Ctrl2 As control    'Second control on form
Dim Ctrl3 As control    'Third control on form
Dim Ctrl4 As control
Dim Ctrl5 As control
Dim Ctrl6 As control
Dim Ctrl7 As control
Dim Ctrl8 As control
Dim smMnfCallType As String 'I=NTR Type; A=Announcer; P=Sports; S=Sales Source; T=Sales Team;
                        'R=Revenue Sets; M=Missed Reason; C=Product Protection; H=Vehicle Group;
                        'B=Business Categories; P=Potential; D=Demos; O=Competitor; Y=Transaction Type
                        'G=Sales Region; E=Genres; N=Network Feed; V=Invoice Sort; X=Program Exclusions
                        'K=Segments; F=Soc Eco Groups; J=Terms; L=Language; Z=Team Name; U = Forcast
                        'W=Hub;1=Subtotal 1; 2=Subtotal 2; 3=Daypart Group; 4=Copy Type; 5=Podcast Categories
                        '5=Position
                        'Q=Unused
Dim imPaintIndex As Integer 'Paint index 0=NTR Type; 1= Announcer; 2= Sales Team;
                        '3=Sales source; 4=Missed reason; 5=Exclusions; 6=Invoice Sort; Sales Region;
                        '7=Revenue Set; 8=Genre; 9=Product Protection; 10=Feed; 11=Transaction;
                        '12=Vehicle Groups; 13=Business Categories; 14=Potential Codes;
                        '15=Custom Demos; 16=Comptitors; 17=Segments
Dim imVehGpSetNo As Integer 'Vehicle Group Set #
Dim imUpdateAllowed As Integer    'User can update records
'MultiNm Box Field Areas
Dim tmCtrls(0 To 8)  As FIELDAREA   'Control fields
Dim imLBCtrls As Integer
Dim imNoCtrls As Integer    'Total number of controls for the form
Dim imBoxNo As Integer  'Current MultiNm Box
Dim tmMnf As MNF       'Mnf record image
Dim tmMnfSrchKey As INTKEY0    'Mnf key record image
Dim imMnfRecLen As Integer        'Mnf record length
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim hmMnf As Integer 'Name and address file handle
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imOriginFirst As Integer
Dim imComboBoxIndex As Integer
Dim imTaxDefined As Integer
Dim imAcqCostDefined As Integer
Dim imSaleType As Integer   '0=NTR; 1=Agency; 2=Direct
Dim imTaxable As Integer    '0=Yes; 1=No
Dim imHardCost As Integer   '0=Yes; 1=No
Dim imEnglish As Integer    '0=Yes; 1=No
Dim imUpdateRvf As Integer  '0=Receivables (store Y); 1=History (store N); 2=Export+History (store E); 3=Export+A/R (Store F); 4=Ask (store A)
Dim imDollars As Integer  '0=Yes; 1= No
Dim imManOpt As Integer    '1=Mandatory; 2=Optional
Dim imBillMGMissed As Integer   '1=Bill MG, not Missed; 2=Bill Missed, not MG; 3=Bill MG & Missed
Dim imMissedFor As Integer   '1=Network Missed; 2=Stations Missed; 3=Network & Station Missed; 4=Station Replacement 'Old values:1=Traffic; 2=Affiliate Web; 3=Both
Dim imDefReason As Integer  '0=Yes; 1=No
Dim imUsThem As Integer '0=Us; 1=Them
Dim imTypeOfFeed As Integer '0=Dish, 1=Antenna; 2=CD; 3=Subfeed
Dim imSubFeedAllowed As Integer '0=Yes; 1=No
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imTestAddStdDemo As Integer
Dim smScreenCaption As String
Dim imMaxCustomDemoNumber As Integer
Dim smAdServerName As String
Dim imAdServerCode As Integer

'9/6/14: Show titles if defined as the screen caption.
Dim smEventTitle1 As String
Dim smEventTitle2 As String

Const NAMEINDEX = 1     'Name control/index
Const CTRL2INDEX = 2    'Control2/field2
Const CTRL3INDEX = 3    'Control3/Field3
Const CTRL4INDEX = 4    'Salesperson commission
Const CTRL5INDEX = 5    'Taxable
Const CTRL6INDEX = 6    'Hard Cost
Const CTRL7INDEX = 7    'Acquisition Cost
Const CTRL8INDEX = 8    'Sale Type

Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim slTranType As String
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    pbcMNm(imPaintIndex).Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcName.Text = slStr
        End If
    End If
    If sgMnfCallType = "Y" Then
        slTranType = UCase$(Trim$(tmMnf.sUnitType))
        If (slTranType = "IN") Or (slTranType = "AN") Or (slTranType = "PI") Or (slTranType = "PO") Or (slTranType = "WB") Or (slTranType = "WV") Then
            pbcMNm(11).Enabled = False
            imUpdateAllowed = False
        Else
            If (igWinStatus(TRANSACTIONSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(11).Enabled = False
                imUpdateAllowed = False
            Else
                pbcMNm(11).Enabled = True
                imUpdateAllowed = True
            End If
        End If
        If imUpdateAllowed Then
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
        Else
            pbcSTab.Enabled = False
            pbcTab.Enabled = False
        End If
    End If
    For ilLoop = imLBCtrls To imNoCtrls Step 1
        mSetShow ilLoop  'Set show strings
    Next ilLoop
    pbcMNm_Paint imPaintIndex
    Screen.MousePointer = vbDefault
    imChgMode = False
'    mSetCommands
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
Private Sub cbcSelect_DblClick()
    'Currently you can't get a double click event on a drop down
    cbcSelect_Click
    imBoxNo = -1
    pbcSTab.SetFocus
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
        If igMNmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
            If sgMNmName = "" Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.Text = sgMNmName   'New name
            End If
            cbcSelect_Change
            If sgMNmName <> "" Then
                mSetCommands
                gFindMatch sgMNmName, 1, cbcSelect
                If gLastFound(cbcSelect) > 0 Then
                    cmcDone.SetFocus
                    Exit Sub
                End If
            End If
            If pbcSTab.Enabled Then
                pbcSTab.SetFocus
            Else
                cmcCancel.SetFocus
            End If
            Exit Sub
        End If
    End If
    slSvText = cbcSelect.Text
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        If pbcSTab.Enabled Then
            pbcSTab.SetFocus
        Else
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
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
    If KeyAscii = KEYBACKSPACE Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcCancel_Click()
    If igMNmCallSource <> CALLNONE Then
        igMNmCallSource = CALLCANCELLED
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
    If igMNmCallSource <> CALLNONE Then
        sgMNmName = edcName.Text 'Save name for returning
        If mSaveRecChg(False) = False Then
            sgMNmName = "[None]"
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
    If igMNmCallSource <> CALLNONE Then
        If smMnfCallType = "Y" Then 'Add transaction type to the description
            sgMNmName = Trim$(Ctrl2.Text) & " " & Trim$(sgMNmName)
        End If
        If sgMNmName = "[New]" Then
            igMNmCallSource = CALLCANCELLED
        Else
            igMNmCallSource = CALLDONE
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
        For ilLoop = imLBCtrls To imNoCtrls Step 1
            If mTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                imBoxNo = ilLoop
                mEnableBox imBoxNo
                Exit Sub
            End If
        Next ilLoop
    End If
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    If (imBoxNo = CTRL3INDEX) And ((smMnfCallType = "S") Or (smMnfCallType = "H")) Then
        lbcOrigin.Visible = Not lbcOrigin.Visible
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
        edcDropDown.SetFocus
    End If
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slStamp As String   'Date/Time stamp for file
    Dim slMsg As String
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        If tgSpf.sRemoteUsers = "Y" Then
            slMsg = "Cannot erase - Remote User System in Use"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        'Check that record is not referenced-Code missing
        Screen.MousePointer = vbHourglass
        Select Case smMnfCallType
            Case "I"    'NTR Type
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Sbf.Btr", "sbfMnfItem")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a NTR Item references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                If StrComp(Trim$(tmMnf.sName), "MultiMedia", vbTextCompare) = 0 Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase MultiMedia"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "A"    'Announcer
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Cif.Btr", "CifMnfAnn")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Copy Inventory references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "S"    'Sales Source
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfSSCode1") 'sofmnfSSCode
                If Not ilRet Then
                    ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfSSCode2") 'sofmnfSSCode
                End If
                If Not ilRet Then
                    ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfSSCode3") 'sofmnfSSCode
                End If
                If Not ilRet Then
                    ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfSSCode4") 'sofmnfSSCode
                End If
                If Not ilRet Then
                    ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfSSCode5") 'sofmnfSSCode
                End If
                If Not ilRet Then
                    ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfSSCode6") 'sofmnfSSCode
                End If
                If Not ilRet Then
                    ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfSSCode7") 'sofmnfSSCode
                End If
                If Not ilRet Then
                    ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfSSCode8") 'sofmnfSSCode
                End If
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Vehicle references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Sof.Btr", "SofMnfSSCode") 'sofmnfSSCode
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Sales Office references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "T"    'Sales Team
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Slf.Btr", "SlfMnfSlsTeam")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Salesperson references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "W"    'Hub
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfHubCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Vehicle references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Urf.Btr", "UrfMnfHubCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a User references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "J"    'Terms
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Adf.Btr", "adfmnfInvTerms")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Agf.Btr", "agfmnfInvTerms")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Agency references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "L"    'Language
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Gsf.Btr", "gsfLangMnfCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Event references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "Z"    'Team Name
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Gsf.Btr", "gsfVisitMnfCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Event Visiting Team references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Gsf.Btr", "gsfHomeMnfCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Event Home Team references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "1"    'Event Subtotal 1
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Gsf.Btr", "gsfSubtotal1MnfCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Event references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "2"    'Language
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Gsf.Btr", "gsfSubtotal2MnfCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Event references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "R"    'Revenue Sets
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfRevBk1")   'chfmnfRevBk1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfRevBk2")   'chfmnfRevBk2
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfRevBk3")   'chfmnfRevBk1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfRevBk4")   'chfmnfRevBk1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfRevBk5")   'chfmnfRevBk1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "M"    'Missed Reason
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Sdf.Btr", "SdfMnfMissed")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Spot Detail references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "C"    'Product Protection
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Adf.Btr", "AdfMnfComp1")   'adfmnfComp1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Adf.Btr", "AdfMnfComp2")   'adfmnfComp2
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfComp1")   'chfmnfComp1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfComp2")   'chfmnfComp2
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Prf.Btr", "PrfMnfComp1")   'adfmnfComp1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser Product Name references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Prf.Btr", "PrfMnfComp2")   'adfmnfComp2
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser Product Name references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Cif.Btr", "CifMnfComp1")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Copy Inventory references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Cif.Btr", "CifMnfComp2")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Copy Inventory references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "X"    'Program Exclusions
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Adf.Btr", "AdfMnfExcl1")   'adfmnfExcl1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Adf.Btr", "AdfMnfExcl2")   'adfmnfExcl2
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfExcl1")   'chfmnfExcl1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfExcl2")   'chfmnfExcl2
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Prf.Btr", "PrfMnfExcl1")   'adfmnfExcl1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser Product Name references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Prf.Btr", "PrfMnfExcl2")   'adfmnfExcl2
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser Product Name references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Lef.Btr", "LefMnfExcl1")    'lefmnfExcl1
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Program Library Events references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Lef.Btr", "LefMnfExcl2")    'lefmnfExcl2
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Program Library Events references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "N"    'Network
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Fxf.Btr", "FxfMnfFeed")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Feed X-Ref references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "G"    'Sales Region
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Sof.Btr", "SofMnfRegion") 'sofmnfRegion
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Sales Office references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "V"    'Invoice sorts
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Adf.Btr", "AdfMnfSort")   'adfmnfSort
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Advertiser references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Agf.Btr", "AgfMnfSort")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Agency references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "E"
            Case "H"    'Vehicle Group
                Select Case Val(tmMnf.sUnitType)
                    Case 1  'Participant
                        ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfGroup")   'Group
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfGroup2")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfGroup3")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfGroup4")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfGroup5")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfGroup6")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfGroup7")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfGroup8")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Rvf.Btr", "RvfMnfGroup")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Phf.Btr", "PhfMnfGroup")   'Group
                        End If
                        If Not ilRet Then
                            ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Pif.Btr", "PifMnfGroup")   'Group
                        End If
                    Case 2  'Subtotal
                        ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfVehGp2")   'Group
                    Case 3  'Market
                        ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfVehGp3Mkt")   'Group
                    Case 4  'Format
                        ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfVehGp4Fmt")   'Group
                    Case 5  'Research
                        ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfVehGp5Rsch")   'Group
                    Case 6  'Sub-Company
                        ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Vef.Btr", "VefMnfVehGp6Sub")   'Group
                End Select
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    If Val(tmMnf.sUnitType) = 1 Then
                        slMsg = "Cannot erase - a Vehicle and/or Receivables and/or Participant references this name"
                    Else
                        slMsg = "Cannot erase - a Vehicle references this name"
                    End If
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "B"    'Business Categories
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfBus")   'Group
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                'ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Ctf.Btr", "CtfMnfBus")   'Group
                'If ilRet Then
                '    Screen.MousePointer = vbDefault
                '    slMsg = "Cannot erase - a Contract Total references this name"
                '    ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
                '    Exit Sub
            Case "K"    'Segmentss
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfSeg")   'Segment
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "P"    'Potential
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfPotnType")   'Group
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "O"    'Competitor
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfCmpy1")   'Competitors
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfCmpy2")   'Competitors
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfCmpy3")   'Competitors
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfCmpy4")   'Competitors
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfCmpy5")   'Competitors
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfCmpy6")   'Competitors
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Chf.Btr", "ChfMnfCmpy7")   'Competitors
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Contract references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "3"    'Daypart Group
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Rdf.Btr", "rdfDPGroupMnfCode")   'adfmnfSort
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Daypart references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "4"    'Copy Type
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Pcf.Btr", "pcfCopyTypeMnfCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - a Copy Type references this name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "5"    'Category
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "thf.Btr", "thfCategoryMnfCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - There is an item that references this category name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
            Case "6"    'Position
                ilRet = gIICodeRefExist(MultiNm, tmMnf.iCode, "Pcf.Btr", "pcfPositionMnfCode")
                If ilRet Then
                    Screen.MousePointer = vbDefault
                    slMsg = "Cannot erase - There is an Ad Server Contract Item that references this Position Name"
                    ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
        End Select
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & tmMnf.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                ilRet = MsgBox("Erase not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                Exit Sub
            End If
        End If
        slStamp = gFileDateTime(sgDBPath & "Mnf.btr")
        ilRet = btrDelete(hmMnf)
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", MultiNm
        On Error GoTo 0
'        If tgSpf.sRemoteUsers = "Y" Then
'            gGetSyncDateTime slSyncDate, slSyncTime
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "MNF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmMnf.iRemoteID
'            tmDsf.lAutoCode = tmMnf.iAutoCode
'            tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'            tmDsf.lCntrNo = 0
'            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'        End If
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        'If lbcNameCode.Tag <> "" Then
        '    If slStamp = lbcNameCode.Tag Then
        '        lbcNameCode.Tag = FileDateTime(sgDBPath & "Mnf.btr")
        '    End If
        'End If
        If sgNameCodeTag <> "" Then
            If slStamp = sgNameCodeTag Then
                sgNameCodeTag = gFileDateTime(sgDBPath & "Mnf.btr")
            End If
        End If
        'lbcNameCode.RemoveItem imSelectedIndex - 1
        gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
        cbcSelect.RemoveItem imSelectedIndex
    End If
    'Remove focus from control and make invisible
    mSetShow imBoxNo
    imBoxNo = -1
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcMNm(imPaintIndex).Cls
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
    If Not imUpdateAllowed Then
        Exit Sub
    End If
End Sub

Private Sub cmcMerge_GotFocus()
    gCtrlGotFocus cmcMerge
    mSetShow imBoxNo
    imBoxNo = -1
End Sub

Private Sub cmcUndo_Click()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    ilIndex = imSelectedIndex
    If ilIndex > 0 Then
        If Not mReadRec(ilIndex, SETFORREADONLY) Then
            GoTo cmcUndoErr
        End If
        mMoveRecToCtrl
        For ilLoop = imLBCtrls To imNoCtrls Step 1
            mSetShow ilLoop  'Set show strings
        Next ilLoop
        pbcMNm(imPaintIndex).Cls
        pbcMNm_Paint imPaintIndex
        mSetCommands
        imBoxNo = -1
        pbcSTab.SetFocus
        Exit Sub
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcMNm(imPaintIndex).Cls
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
    Dim slName As String    'Save name as mSaveRec set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilRet As Integer
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
    Screen.MousePointer = vbHourglass
    imBoxNo = -1
    ''Must reset display so altered flag is cleared and setcommand will turn select on
    'If imSvSelectedIndex <> 0 Then
    '    cbcSelect.Text = slName
    'Else
    '    cbcSelect.ListIndex = 0
    'End If
    'cbcSelect_Change    'Call change so picture area repainted
    ilCode = tmMnf.iCode
    cbcSelect.Clear
    sgNameCodeTag = ""
    If smMnfCallType = "B" Then
        sgBusCatMnfStamp = ""
    End If
    If smMnfCallType = "C" Then
        sgCompMnfStamp = ""
    End If
    If smMnfCallType = "D" Then
        sgDemoMnfStamp = ""
    End If
    If smMnfCallType = "F" Then
        sgSocEcoMnfStamp = ""
    End If
    If smMnfCallType = "P" Then
        sgPotMnfStamp = ""
    End If
    If smMnfCallType = "X" Then
        sgExclMnfStamp = ""
    End If
    mPopulate
    For ilLoop = 0 To UBound(tgNameCode) - 1 Step 1 'lbcDPNameCode.ListCount - 1 Step 1
        slNameCode = tgNameCode(ilLoop).sKey 'lbcDPNameCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Val(slCode) = ilCode Then
            If cbcSelect.ListIndex = ilLoop + 1 Then
                cbcSelect_Change
            Else
                cbcSelect.ListIndex = ilLoop + 1
            End If
            Exit For
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault
    mSetCommands
    If cbcSelect.Enabled Then
        cbcSelect.SetFocus
    Else
        cmcCancel.SetFocus
    End If
End Sub
Private Sub cmcUpdate_GotFocus()
    gCtrlGotFocus cmcUpdate
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub edcComm_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcComm_GotFocus()
    gCtrlGotFocus edcComm
End Sub
Private Sub edcComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer    'Decimal point position
    Dim slStr As String
    Dim ilKey As Integer
    ilPos = InStr(edcComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcComm.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcComm.Text
    slStr = Left$(slStr, edcComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcComm.SelStart - edcComm.SelLength)
    If gCompNumberStr(slStr, "99.9999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcCtrl3_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcCtrl3_GotFocus()
    If smMnfCallType = "C" Then
        If edcCtrl3.Text = "" Then
            edcCtrl3.Text = Left$(edcName.Text, 5)
            mSetChg imBoxNo   'Change event not generated
            mSetCommands
        End If
    End If
    If smMnfCallType = "X" Then
        If edcCtrl3.Text = "" Then
            edcCtrl3.Text = Left$(edcName.Text, 5)
            mSetChg imBoxNo   'Change event not generated
            mSetCommands
        End If
    End If
    gCtrlGotFocus edcCtrl3
End Sub
Private Sub edcCtrl3_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If smMnfCallType = "A" Then 'Announcer- group number
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl3.Text
        slStr = Left$(slStr, edcCtrl3.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl3.SelStart - edcCtrl3.SelLength)
        If gCompNumberStr(slStr, "99") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If smMnfCallType = "V" Then 'Invoice sort number
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl3.Text
        slStr = Left$(slStr, edcCtrl3.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl3.SelStart - edcCtrl3.SelLength)
        If gCompNumberStr(slStr, "999") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If smMnfCallType = "E" Then 'Group number
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl3.Text
        slStr = Left$(slStr, edcCtrl3.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl3.SelStart - edcCtrl3.SelLength)
        If gCompNumberStr(slStr, "99") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If smMnfCallType = "Y" Then 'Transaction type
        If Len(edcCtrl3.Text) = 0 Then
            slStr = UCase$(Chr$(KeyAscii))
            If (slStr <> "P") And (slStr <> "I") And (slStr <> "A") And (slStr <> "W") Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            KeyAscii = Asc(slStr)
        Else
            slStr = UCase$(edcCtrl3.Text & Chr$(KeyAscii))
            If (slStr = "IN") Then   'Or (slStr = "II") Or (slStr = "IL") Or (slStr = "IX") Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            If (slStr = "AN") Then  'Or (slStr = "AI") Or (slStr = "AL") Or (slStr = "AX") Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            If (slStr = "PI") Or (slStr = "PO") Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            If (slStr = "WB") Or (slStr = "WV") Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = Chr$(KeyAscii)
            KeyAscii = Asc(UCase$(slStr))
        End If
    End If
    If smMnfCallType = "H" Then 'Vehicle Group Sort Order #
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl3.Text
        slStr = Left$(slStr, edcCtrl3.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl3.SelStart - edcCtrl3.SelLength)
        If gCompNumberStr(slStr, "999") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If smMnfCallType = "B" Then 'Potential
    End If
    If smMnfCallType = "K" Then 'Segment
    End If
    If smMnfCallType = "P" Then 'Invoice sort number
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl3.Text
        slStr = Left$(slStr, edcCtrl3.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl3.SelStart - edcCtrl3.SelLength)
        If gCompNumberStr(slStr, "100") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    '5/7/10: Assign group number automatically
    'If smMnfCallType = "D" Then 'Vehicle Group Sort Order #
    '    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
    '        Beep
    '        KeyAscii = 0
    '        Exit Sub
    '    End If
    '    slStr = edcCtrl3.Text
    '    slStr = Left$(slStr, edcCtrl3.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl3.SelStart - edcCtrl3.SelLength)
    '    If gCompNumberStr(slStr, "990") > 0 Then
    '        Beep
    '        KeyAscii = 0
    '        Exit Sub
    '    End If
    'End If
    If smMnfCallType = "3" Then 'Daypart Group
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl3.Text
        slStr = Left$(slStr, edcCtrl3.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl3.SelStart - edcCtrl3.SelLength)
        If gCompNumberStr(slStr, "999") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcCtrl4_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcCtrl4_GotFocus()
    gCtrlGotFocus edcCtrl4
End Sub
Private Sub edcCtrl4_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If smMnfCallType = "R" Then 'Revenue Set
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl4.Text
        slStr = Left$(slStr, edcCtrl4.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl4.SelStart - edcCtrl4.SelLength)
        If gCompNumberStr(slStr, "5") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If smMnfCallType = "P" Then 'Potential Code
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl4.Text
        slStr = Left$(slStr, edcCtrl4.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl4.SelStart - edcCtrl4.SelLength)
        If gCompNumberStr(slStr, "100") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
    If smMnfCallType = "H" Then 'Vehicle Group number (set # field)
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl4.Text
        slStr = Left$(slStr, edcCtrl4.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl4.SelStart - edcCtrl4.SelLength)
        If gCompNumberStr(slStr, "99") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcCtrl5_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcCtrl5_GotFocus()
    gCtrlGotFocus edcCtrl5
End Sub
Private Sub edcCtrl5_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If smMnfCallType = "P" Then 'Potential Code
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcCtrl5.Text
        slStr = Left$(slStr, edcCtrl5.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcCtrl5.SelStart - edcCtrl5.SelLength)
        If gCompNumberStr(slStr, "100") > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcDropDown_Change()
    If (imBoxNo = CTRL3INDEX) And ((smMnfCallType = "S") Or (smMnfCallType = "H")) Then
        imLbcArrowSetting = True
        gMatchLookAhead edcDropDown, lbcOrigin, imBSMode, imComboBoxIndex
        imLbcArrowSetting = False
        mSetChg imBoxNo
    End If
End Sub
Private Sub edcDropDown_GotFocus()
    If (imBoxNo = CTRL3INDEX) And ((smMnfCallType = "S") Or (smMnfCallType = "H")) Then
        If lbcOrigin.ListCount = 1 Then
            lbcOrigin.ListIndex = 0
            'If imTabDirection = -1 Then  'Right To Left
            '    pbcSTab.SetFocus
            'Else
            '    pbcTab.SetFocus
            'End If
            'Exit Sub
        End If
    End If
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
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
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (imBoxNo = CTRL3INDEX) And ((smMnfCallType = "S") Or (smMnfCallType = "H")) Then
            gProcessArrowKey Shift, KeyCode, lbcOrigin, imLbcArrowSetting
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
        End If
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcName_Change()
    mSetChg NAMEINDEX   'can't use imBoxNo as not set when edcProd set via cbcSelect- altered flag set so field is saved
End Sub
Private Sub edcName_GotFocus()
    gCtrlGotFocus edcName
End Sub
Private Sub edcName_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If smMnfCallType = "S" Then
        If (KeyAscii = KEYSLASH) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
Private Sub edcName_LostFocus()
    'Dim ilRet As Integer
    ''Test if name changed and if new name is valid
    'ilRet = mOKName(True)
End Sub
Private Sub edcRate_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcRate_GotFocus()
    gCtrlGotFocus edcRate
End Sub
Private Sub edcRate_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    Dim ilKey As Integer
    ilPos = InStr(edcRate.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcRate.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcRate.Text
    slStr = Left$(slStr, edcRate.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcRate.SelStart - edcRate.SelLength)
    If gCompNumberStr(slStr, "9999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcSlspComm_Change()
    mSetChg imBoxNo
End Sub
Private Sub edcSlspComm_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcSlspComm_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer    'Decimal point position
    Dim slStr As String
    Dim ilKey As Integer
    ilPos = InStr(edcSlspComm.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcSlspComm.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcSlspComm.Text
    slStr = Left$(slStr, edcSlspComm.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSlspComm.SelStart - edcSlspComm.SelLength)
    If gCompNumberStr(slStr, "99.9999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Select Case sgMnfCallType
        Case "I"    'NTR Type
            If (igWinStatus(ITEMBILLINGTYPESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(0).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(ITEMBILLINGTYPESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(0).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "V"    'Invoice sort
            If (igWinStatus(INVOICESORTLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(6).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(INVOICESORTLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(6).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "A"    'Announcer
            If (igWinStatus(ANNOUNCERNAMESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(1).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(ANNOUNCERNAMESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(1).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "S"    'Sales Source
            If (igWinStatus(SALESSOURCESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(3).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(SALESSOURCESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(3).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "T"    'Sales Team
            If (igWinStatus(SALESTEAMSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(2).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(SALESTEAMSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(2).Enabled = True
                imUpdateAllowed = True
            End If
        Case "W"    'Hub
            If (igWinStatus(VEHICLESLIST) <= 1) And (igWinStatus(USERLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(2).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(VEHICLESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(2).Enabled = True
                imUpdateAllowed = True
            End If
        Case "J"    'Terms
            If (igWinStatus(AGENCIESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(19).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(AGENCIESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(19).Enabled = True
                imUpdateAllowed = True
            End If
        Case "L"    'Language
            If (igWinStatus(PROGRAMMINGJOB) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(21).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(PROGRAMMINGJOB) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(21).Enabled = True
                imUpdateAllowed = True
            End If
        Case "Z"    'Team
            If (igWinStatus(PROGRAMMINGJOB) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(20).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(PROGRAMMINGJOB) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(20).Enabled = True
                imUpdateAllowed = True
            End If
        Case "1"    'Event Subtotal 1
            If (igWinStatus(PROGRAMMINGJOB) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(2).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(PROGRAMMINGJOB) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(2).Enabled = True
                imUpdateAllowed = True
            End If
        Case "2"    'Event Subtotal 2
            If (igWinStatus(PROGRAMMINGJOB) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(2).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(PROGRAMMINGJOB) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(2).Enabled = True
                imUpdateAllowed = True
            End If
        Case "R"    'Revenue Sets
            If (igWinStatus(REVENUESETSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(7).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(REVENUESETSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(7).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "M"    'Missed Reason
            If (igWinStatus(MISSEDREASONSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(4).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(MISSEDREASONSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(4).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "C"    'Product Protection
            If (igWinStatus(COMPETITIVESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(9).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(COMPETITIVESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(9).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "X"    'Program Exclusions
            If (igWinStatus(EXCLUSIONSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(5).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(EXCLUSIONSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(5).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "N"    'Feed
            If (igWinStatus(FEEDTYPESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(10).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(FEEDTYPESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(10).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "G"    'Sales region
            If (igWinStatus(SALESREGIONSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(6).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(SALESREGIONSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(6).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "E"    'Genres
            If (igWinStatus(GENRESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(8).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(GENRESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(8).Enabled = True
                imUpdateAllowed = True
            End If
'            gShowBranner
        Case "Y"    'Transaction Type
            If (igWinStatus(TRANSACTIONSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(11).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(TRANSACTIONSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(11).Enabled = True
                imUpdateAllowed = True
            End If
        Case "H"    'Vehicle Group
            If (igWinStatus(VEHICLEGROUPSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(12).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(VEHICLEGROUPSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(12).Enabled = True
                imUpdateAllowed = True
            End If
        Case "B"    'Business Catories
            If (igWinStatus(BUSCATEGORIESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(13).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(BUSCATEGORIESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(13).Enabled = True
                imUpdateAllowed = True
            End If
        Case "P"    'Potential
            If (igWinStatus(POTENTIALCODESLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(14).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(POTENTIALCODESLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(14).Enabled = True
                imUpdateAllowed = True
            End If
        Case "D"    'Vehicle Group
            If (igWinStatus(DEMOSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(15).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(DEMOSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(15).Enabled = True
                imUpdateAllowed = True
            End If
        Case "F"    'Vehicle Group
            If (igWinStatus(RESEARCHLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(18).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(RESEARCHLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(18).Enabled = True
                imUpdateAllowed = True
            End If
        Case "O"    'Competitor
            If (igWinStatus(COMPETITORSLIST) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(16).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(COMPETITORSLIST) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(16).Enabled = True
                imUpdateAllowed = True
            End If
        Case "K"    'Segments
            If (tgSpf.sCUseSegments <> "Y") And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(17).Enabled = False
                imUpdateAllowed = False
            Else
                pbcMNm(17).Enabled = True
                imUpdateAllowed = True
            End If
        Case "3"    'Daypart Group
            If (igWinStatus(RATECARDSJOB) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(6).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(RATECARDSJOB) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(6).Enabled = True
                imUpdateAllowed = True
            End If
        Case "4"    'Copy Type
            If ((igWinStatus(PROPOSALSJOB) <= 1) And (igWinStatus(CONTRACTSJOB) <= 1)) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                pbcMNm(2).Enabled = False
                imUpdateAllowed = False
                If (igWinStatus(PROGRAMMINGJOB) = 0) Then
                    tmcHide.Enabled = True
                End If
            Else
                pbcMNm(2).Enabled = True
                imUpdateAllowed = True
            End If
        Case "5"    'Category
            'If (igWinStatus(PROGRAMMINGJOB) <= 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            '    pbcMNm(2).Enabled = False
            '    imUpdateAllowed = False
            '    If (igWinStatus(PROGRAMMINGJOB) = 0) Then
            '        tmcHide.Enabled = True
            '    End If
            'Else
                pbcMNm(2).Enabled = True
                imUpdateAllowed = True
            'End If
        Case "6"    'Position
            pbcMNm(20).Enabled = True
            imUpdateAllowed = True
    End Select
    If imUpdateAllowed Then
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
    Else
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
    End If
    gShowBranner imUpdateAllowed
    mSetCommands
    Me.KeyPreview = True
    MultiNm.Refresh
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
    
    Erase tgNameCode
    
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf

    Set MultiNm = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcOrigin_Click()
    gProcessLbcClick lbcOrigin, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcOrigin_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddStdDemo                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add Standard Demos              *
'*                                                     *
'*******************************************************
Private Function mAddStdDemo() As Integer
'
'   ilRet = mAddStdDemo ()
'   Where:
'       ilRet (O)- True = populated; False = error
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilAddMissingOnly As Integer

    If Not imTestAddStdDemo Then
        mAddStdDemo = True
        Exit Function
    End If
    imTestAddStdDemo = False
    ReDim ilFilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    ilFilter(0) = CHARFILTER
    slFilter(0) = "D"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilFilter(1) = INTEGERFILTER
    slFilter(1) = "0"
    ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
    lbcDemos.Clear
    ilRet = gIMoveListBox(MultiNm, lbcDemos, tgNameCode(), sgNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    sgNameCodeTag = ""
    If lbcDemos.ListCount > 0 Then
        'Test if 20 exist
        For ilLoop = 1 To lbcDemos.ListCount - 1 Step 1
            If InStr(1, lbcDemos.List(ilLoop), "20", vbTextCompare) > 0 Then
                mAddStdDemo = True
                Exit Function
            End If
        Next ilLoop
        'Add in missing demos
        ilAddMissingOnly = True
    Else
        ilAddMissingOnly = False
    End If
    lbcDemos.Clear
    gDemoPop lbcDemos   'Get demo names
    gGetSyncDateTime slSyncDate, slSyncTime
    For ilLoop = 1 To lbcDemos.ListCount - 1 Step 1
        ilFound = False
        If ilAddMissingOnly Then
            For ilIndex = LBound(tgNameCode) To UBound(tgNameCode) - 1 Step 1
                If InStr(1, Trim$(tgNameCode(ilIndex).sKey), Trim$(lbcDemos.List(ilLoop)), vbTextCompare) > 0 Then
                    ilFound = True
                    Exit For
                End If
            Next ilIndex
        End If
        If Not ilFound Then
            tmMnf.iCode = 0
            tmMnf.sType = "D"
            tmMnf.sName = lbcDemos.List(ilLoop)
            tmMnf.sRPU = ""
            tmMnf.sUnitType = ""
            tmMnf.iMerge = 0
            tmMnf.iGroupNo = 0
            tmMnf.sCodeStn = ""
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
            Do
                'tmMnfSrchKey.iCode = tmMnf.iCode
                'ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
                tmMnf.iAutoCode = tmMnf.iCode
                gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
                gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
                ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
    Next ilLoop
    mAddStdDemo = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    edcName.Text = ""
    If TypeOf Ctrl2 Is Image Then
    Else
        If (smMnfCallType <> "R") And (smMnfCallType <> "M") And (smMnfCallType <> "N") And (smMnfCallType <> "L") Then
            Ctrl2.Text = ""
        End If
    End If
    If TypeOf Ctrl3 Is Image Then
    Else
        If (smMnfCallType = "S") Or (smMnfCallType = "H") Then
            Ctrl3.ListIndex = -1
        Else
            If (smMnfCallType <> "N") And (smMnfCallType <> "M") And (smMnfCallType <> "O") Then
                Ctrl3.Text = ""
            End If
        End If
        imOriginFirst = True
    End If
    If TypeOf Ctrl4 Is Image Then
    Else
        If (smMnfCallType <> "S") And (smMnfCallType <> "H") And (smMnfCallType <> "M") Then
            Ctrl4.Text = ""
        End If
    End If
    If (smMnfCallType = "S") Then
        Ctrl5.Text = ""
    End If
    If smMnfCallType = "I" Then
        imTaxable = -1
        imHardCost = -1
        If imAcqCostDefined Then
            Ctrl7.Text = ""
        End If
        imSaleType = -1
    End If
    If smMnfCallType = "L" Then
        imEnglish = -1
    End If
    If smMnfCallType = "S" Then
        imUpdateRvf = -1
    End If
    If smMnfCallType = "H" Then
        imDollars = -1
    End If
    If smMnfCallType = "M" Then
        imBillMGMissed = -1
        imMissedFor = -1
        imDefReason = -1
    End If
    If smMnfCallType = "R" Then
        imManOpt = -1
    End If
    If smMnfCallType = "N" Then
        imTypeOfFeed = -1
        imSubFeedAllowed = -1
    End If
    If smMnfCallType = "O" Then
        imUsThem = -1
    End If
    lbcOrigin.ListIndex = -1
    mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    Dim slStr As String

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imNoCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = True  'Set visibility
            edcName.SetFocus
        Case CTRL2INDEX 'Second control
            '5/7/10: Assign group number automatically
            'If (smMnfCallType <> "R") And (smMnfCallType <> "N") And (smMnfCallType <> "M") And (smMnfCallType <> "L") And (smMnfCallType <> "D") Then
            If (smMnfCallType <> "R") And (smMnfCallType <> "N") And (smMnfCallType <> "M") And (smMnfCallType <> "L") Then
                Ctrl2.Visible = True  'Set visibility
                Ctrl2.SetFocus
            Else    'Revenue sets
                If smMnfCallType = "R" Then
                    If imManOpt < 0 Then
                        imManOpt = 1                           '1
                        tmCtrls(ilBoxNo).iChg = True
                        mSetCommands
                    End If
                ElseIf smMnfCallType = "M" Then
                    If imBillMGMissed < 0 Then
                        imBillMGMissed = 1    'Bill MG, not Missed
                        tmCtrls(ilBoxNo).iChg = True
                        mSetCommands
                    End If
                ElseIf smMnfCallType = "L" Then
                    If imEnglish < 0 Then
                        If cbcSelect.ListCount = 1 Then
                            imEnglish = 0
                        Else
                            imEnglish = 1
                        End If
                        tmCtrls(ilBoxNo).iChg = True
                        mSetCommands
                    End If
                '5/7/10:  Add more custom groups
                '5/7/10: Assign group number automatically
                'ElseIf smMnfCallType = "D" Then
                '    'Set Group number to next possible value
                '    If imSelectedIndex = 0 Then 'New selected
                '        If Ctrl2.Text = "" Then
                '            Ctrl2.Text = Trim$(str$(imMaxCustomDemoNumber + 1))
                '        End If
                '    End If
                '    Ctrl2.Visible = True  'Set visibility
                '    Ctrl2.SetFocus
                Else
                    If imTypeOfFeed < 0 Then
                        imTypeOfFeed = 0    'Dish
                        tmCtrls(ilBoxNo).iChg = True
                        mSetCommands
                    End If
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                If smMnfCallType = "R" Then
                    gMoveFormCtrl pbcMNm(7), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                ElseIf smMnfCallType = "M" Then
                    pbcYN.Width = (3 * tmCtrls(ilBoxNo).fBoxW) / 2 + 180
                    gMoveFormCtrl pbcMNm(4), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                ElseIf smMnfCallType = "L" Then
                    gMoveFormCtrl pbcMNm(21), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                Else
                    gMoveFormCtrl pbcMNm(10), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                End If
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            End If
        Case CTRL3INDEX 'Second control
            If (smMnfCallType = "S") Then
                lbcOrigin.height = gListBoxHeight(lbcOrigin.ListCount, 3)
                edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 8
                gMoveFormCtrl pbcMNm(3), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcOrigin.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                imChgMode = True
                If lbcOrigin.ListIndex < 0 Then
                    lbcOrigin.ListIndex = 0
                End If
                imComboBoxIndex = lbcOrigin.ListIndex
                If lbcOrigin.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcOrigin.List(lbcOrigin.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            ElseIf (smMnfCallType = "H") Then
                lbcOrigin.height = gListBoxHeight(lbcOrigin.ListCount, 3)
                edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
                edcDropDown.MaxLength = 0
                gMoveFormCtrl pbcMNm(12), edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
                lbcOrigin.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.height
                imChgMode = True
                If lbcOrigin.ListIndex < 0 Then
                    lbcOrigin.ListIndex = 0
                End If
                imComboBoxIndex = lbcOrigin.ListIndex
                If lbcOrigin.ListIndex < 0 Then
                    edcDropDown.Text = ""
                Else
                    edcDropDown.Text = lbcOrigin.List(lbcOrigin.ListIndex)
                End If
                imChgMode = False
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                edcDropDown.Visible = True
                cmcDropDown.Visible = True
                edcDropDown.SetFocus
            ElseIf smMnfCallType = "N" Then
                If imSubFeedAllowed < 0 Then
                    imSubFeedAllowed = 1    'No
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(10), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            ElseIf smMnfCallType = "O" Then
                If imUsThem < 0 Then
                    imUsThem = 2    'Them
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(16), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            ElseIf smMnfCallType = "M" Then
                If imDefReason < 0 Then
                    imDefReason = 1    'No
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(4), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            Else
                Ctrl3.Visible = True  'Set visibility
                Ctrl3.SetFocus
            End If
        Case CTRL4INDEX 'Second control
            If smMnfCallType = "S" Then
                If imUpdateRvf < 0 Then
                    imUpdateRvf = 0 'Receiables
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(3), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            ElseIf smMnfCallType = "H" Then
                If imDollars < 0 Then
                    'imDollars = 0 'Yes
                    If lbcOrigin.ListIndex < 0 Then
                        slStr = ""
                    Else
                        slStr = lbcOrigin.List(lbcOrigin.ListIndex)
                    End If
                    If (StrComp(slStr, "Research", 1) = 0) Then
                        imDollars = 0 'Yes
                    Else
                        imDollars = 1   'No
                    End If
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(12), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            ElseIf smMnfCallType = "M" Then
                If imMissedFor < 0 Then
                    imMissedFor = 1    'Traffic
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(4), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            Else
                Ctrl4.Visible = True  'Set visibility
                Ctrl4.SetFocus
            End If
        Case CTRL5INDEX 'Second control
            If smMnfCallType = "I" Then
                If imHardCost < 0 Then
                    imHardCost = 1    'No
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(0), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            ElseIf smMnfCallType = "S" Then
                Ctrl5.Visible = True  'Set visibility
                Ctrl5.SetFocus
            End If
        Case CTRL6INDEX 'Second control
            If smMnfCallType = "I" Then
                If imTaxable < 0 Then
                    imTaxable = 1    'No
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(0), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            End If
        Case CTRL7INDEX 'Second control
            If smMnfCallType = "I" Then
                Ctrl7.Visible = True  'Set visibility
                Ctrl7.SetFocus
            End If
        Case CTRL8INDEX 'Second control
            If smMnfCallType = "I" Then
                If imSaleType < 0 Then
                    imSaleType = 0    'NTR
                    tmCtrls(ilBoxNo).iChg = True
                    mSetCommands
                End If
                pbcYN.Width = tmCtrls(ilBoxNo).fBoxW
                gMoveFormCtrl pbcMNm(0), pbcYN, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            End If
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    imPopReqd = False
    imTestAddStdDemo = True
    imFirstFocus = True
    smMnfCallType = sgMnfCallType
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imMaxCustomDemoNumber = 0
    'gPDNToStr tgSpf.sBTax(0), 2, slStr1
    'gPDNToStr tgSpf.sBTax(1), 2, slStr2
    'If (Val(slStr1) = 0) And (Val(slStr2) = 0) Then
    '12/17/06-Change to tax by agency or vehicle
    'If (tgSpf.iBTax(0) = 0) And (tgSpf.iBTax(1) = 0) Then
    If (Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR Then
        imTaxDefined = True
    Else
        imTaxDefined = False
    End If
    '6/7/15: replaced acquisition from site override with Barter in system options
    If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) <> SPNTRACQUISITION Then
        imAcqCostDefined = False
    Else
        imAcqCostDefined = True
    End If
    Select Case smMnfCallType
        Case "I"    'NTR Type
            imPaintIndex = 0
            'If imTaxDefined Then
            '    imNoCtrls = 6
            'Else
            '    imNoCtrls = 5
            'End If
            'If imAcqCostDefined Then
            '    imNoCtrls = imNoCtrls + 1
            'End If
            imNoCtrls = 8
            smScreenCaption = "NTR Types"
        Case "A"    'Announcer
            imPaintIndex = 1
            imNoCtrls = 3
            smScreenCaption = "Announcer Names"
        Case "S"    'Sales Source
            imPaintIndex = 3
            imNoCtrls = 5
            smScreenCaption = "Sales Sources"
            lbcOrigin.AddItem "Local"
            lbcOrigin.AddItem "Regional"
            lbcOrigin.AddItem "National"
            imOriginFirst = True
        Case "T"    'Sales Team
            imPaintIndex = 2
            imNoCtrls = 1
            smScreenCaption = "Sales Teams"
        Case "W"    'Hub
            imPaintIndex = 2
            imNoCtrls = 1
            smScreenCaption = "Hub"
        Case "J"    'Terms
            imPaintIndex = 19
            imNoCtrls = 1
            smScreenCaption = "Terms"
        Case "L"    'Language
            imPaintIndex = 21
            imNoCtrls = 2
            smScreenCaption = "Language"
        Case "Z"    'Team
            gGetEventTitles igGameSchdVefCode, smEventTitle1, smEventTitle2
            imPaintIndex = 20
            imNoCtrls = 2
            If ((smEventTitle1 = "Visiting Team") And (smEventTitle2 = "Home Team")) Or ((smEventTitle1 = "") And (smEventTitle2 = "")) Then
                smScreenCaption = "Team Name"
            Else
                If (smEventTitle1 <> "") And (smEventTitle2 = "") Then
                    smScreenCaption = smEventTitle1
                ElseIf (smEventTitle1 = "") And (smEventTitle2 <> "") Then
                    smScreenCaption = smEventTitle2
                Else
                    smScreenCaption = smEventTitle1 & "/" & smEventTitle2
                End If
            End If
        Case "1"    'Event Subtotal 1
            imPaintIndex = 2
            imNoCtrls = 1
            smScreenCaption = Trim$(tgSaf(0).sEventSubtotal1)  '"Event Subtotal 1"
        Case "2"    'Event Subtotal 2
            imPaintIndex = 2
            imNoCtrls = 1
            smScreenCaption = Trim$(tgSaf(0).sEventSubtotal1)  '"Event Subtotal 2"
        Case "R"    'Revenue Sets
            imPaintIndex = 7
            If tgSpf.sAStnCodes = "N" Then
                imNoCtrls = 3   '2
            Else
                imNoCtrls = 4   '3
            End If
            smScreenCaption = "Revenue Sets"
        Case "M"    'Missed Reason
            imPaintIndex = 4
            imNoCtrls = 4
            smScreenCaption = "Missed Reasons"
        Case "C"    'Product Protection
            imPaintIndex = 9
            If tgSpf.sAStnCodes = "N" Then
                imNoCtrls = 2
            Else
                imNoCtrls = 3
            End If
            smScreenCaption = "Product Protection"
        Case "X"    'Exclusions
            imPaintIndex = 5
            imNoCtrls = 2
            smScreenCaption = "Program Exclusions"
        Case "N"    'Feed
            imPaintIndex = 10
            imNoCtrls = 4
            smScreenCaption = "Feed Types"
        Case "V"    'Invoice sort
            imPaintIndex = 6
            imNoCtrls = 2
            smScreenCaption = "Invoice Sort"
        Case "G"    'Sales Regions
            imPaintIndex = 6
            imNoCtrls = 2
            smScreenCaption = "Sales Regions"
        Case "E"    'Genre
            imPaintIndex = 8
            imNoCtrls = 2
            smScreenCaption = "Genre Names"
        Case "Y"    'Transaction Type
            imPaintIndex = 11
            imNoCtrls = 2
            smScreenCaption = "Transaction Types"
        Case "H"    'Vehicle Group
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 40
            cbcSelect.Left = 250
            cbcSelect.Width = 4400
            imPaintIndex = 12
            imNoCtrls = 3
            If imVehGpSetNo = 1 Then
                lbcOrigin.AddItem "Participant"
                smScreenCaption = "Vehicle Group: Participant"
            ElseIf imVehGpSetNo = 2 Then
                lbcOrigin.AddItem "Subtotal"
                smScreenCaption = "Vehicle Group: Subtotal"
            ElseIf imVehGpSetNo = 3 Then
                lbcOrigin.AddItem "Market"
                smScreenCaption = "Vehicle Group: Market"
                imNoCtrls = 4
            ElseIf imVehGpSetNo = 4 Then
                lbcOrigin.AddItem "Format"
                smScreenCaption = "Vehicle Group: Format"
            ElseIf imVehGpSetNo = 5 Then
                lbcOrigin.AddItem "Research"
                smScreenCaption = "Vehicle Group: Research"
                imNoCtrls = 4
            ElseIf imVehGpSetNo = 6 Then
                lbcOrigin.AddItem "Sub-Company"
                smScreenCaption = "Vehicle Group: Sub-Company"
            Else
                lbcOrigin.AddItem "Participant"
                lbcOrigin.AddItem "Subtotal"
                lbcOrigin.AddItem "Market"
                lbcOrigin.AddItem "Format"
                lbcOrigin.AddItem "Research"
                'If tgSpf.sSubCompany = "Y" Then
                    lbcOrigin.AddItem "Sub-Company"
                'End If
                smScreenCaption = "Vehicle Groups"
                imNoCtrls = 4
            End If
        Case "B"    'Business Categories
            imPaintIndex = 13
            imNoCtrls = 2
            smScreenCaption = "Business Categories"
        Case "P"    'Potential Codes
            imPaintIndex = 14
            imNoCtrls = 4
            smScreenCaption = "Potential Codes"
        Case "D"    'Custom Demos
            imPaintIndex = 15
            '5/7/10: Assign group number automatically
            'imNoCtrls = 2
            imNoCtrls = 1
            smScreenCaption = "Custom Demos"
        Case "F"    'Soc Eco Group
            imPaintIndex = 18
            imNoCtrls = 2
            smScreenCaption = "Qualitative Groups"
            cmcErase.Enabled = False
        Case "O"    'Competitors
            imPaintIndex = 16
            imNoCtrls = 3
            smScreenCaption = "Competitors"
        Case "K"    'Segments
            imPaintIndex = 17
            imNoCtrls = 2
            smScreenCaption = "Segments"
        Case "3"    'Daypart Group
            imPaintIndex = 6
            imNoCtrls = 2
            smScreenCaption = "Daypart Group"
        Case "4"    'Copy Type
            imPaintIndex = 2
            imNoCtrls = 1
            smScreenCaption = "Copy Type"
        Case "5"    'Podcast Category
            imPaintIndex = 2
            imNoCtrls = 1
            smScreenCaption = "Ad Server Category"
        Case "6"    'Ad Server Position
            imPaintIndex = 20
            imNoCtrls = 1
            smScreenCaption = smAdServerName & " Position"
    End Select

    mInitBox
    MultiNm.height = cmcUndo.Top + 5 * cmcUndo.height / 3
    gCenterStdAlone MultiNm
    'MultiNm.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imMnfRecLen = Len(tmMnf)  'Get and save Mnf record length
    imBoxNo = -1 'Initialize current MNm Box to N/A
    imChgMode = False
    imBSMode = False
    imBypassSetting = False
    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", MultiNm
    On Error GoTo 0
'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", MultiNm
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)  'Get and save Mnf record length
    MultiNm.height = cmcUndo.Top + 5 * cmcUndo.height / 3
    gCenterStdAlone MultiNm
'    gCenterModalForm MultiNm
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not imTerminate Then
        cbcSelect.ListIndex = 0
        mSetCommands
    End If
    'plcScreen.Cls
    'plcScreen_Paint
    plcScreen.Caption = smScreenCaption
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
    Dim ilLoop As Integer    'For loop control parameter
    flTextHeight = pbcMNm(0).TextHeight("1")
    'Position panel and picture areas with panel
    'Set picture visible
    For ilLoop = 0 To 10 Step 1
        pbcMNm(ilLoop).Visible = False
    Next ilLoop
    Select Case smMnfCallType
        Case "I"    'NTR Type
            plcMNm.Move 930, 735, pbcMNm(0).Width + fgPanelAdj, pbcMNm(0).height + fgPanelAdj
            pbcMNm(0).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(0).Visible = True
            Set Ctrl2 = edcRate
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = edcCtrl3
            Ctrl3.ZOrder vbBringToFront   'Place in front
            Set Ctrl4 = edcSlspComm
            Ctrl4.ZOrder vbBringToFront   'Place in front
            Set Ctrl5 = pbcYN
            Ctrl5.ZOrder vbBringToFront
            If imTaxDefined Then
                Set Ctrl6 = pbcYN
                Ctrl6.ZOrder vbBringToFront
                'lacCover.Visible = False
            Else
                Set Ctrl6 = imcUnused
                'lacCover.Visible = True
            End If
            If imAcqCostDefined Then
                Set Ctrl7 = edcCtrl4
                Ctrl7.ZOrder vbBringToFront   'Place in front
            Else
                Set Ctrl7 = imcUnused
            End If
            Set Ctrl8 = pbcYN
            Ctrl8.ZOrder vbBringToFront
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(0), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Rate per unit (x,xxx,xxx.xx)
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 900, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 10
            gMoveFormCtrl pbcMNm(0), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'Unit type
            gSetCtrl tmCtrls(CTRL3INDEX), 945, tmCtrls(CTRL2INDEX).fBoxY, 960, fgBoxStH
            Ctrl3.Width = tmCtrls(CTRL3INDEX).fBoxW
            'Changed to 15 from 6 on 12/9/02 added mnfUnitsPer to DDFs
            Ctrl3.MaxLength = 15    '6
            gMoveFormCtrl pbcMNm(0), Ctrl3, tmCtrls(CTRL3INDEX).fBoxX, tmCtrls(CTRL3INDEX).fBoxY
            'Salesperson commission (xx.xxxx)
            gSetCtrl tmCtrls(CTRL4INDEX), 1920, tmCtrls(CTRL2INDEX).fBoxY, 915, fgBoxStH
            Ctrl4.Width = tmCtrls(CTRL4INDEX).fBoxW
            Ctrl4.MaxLength = 8
            gMoveFormCtrl pbcMNm(0), Ctrl4, tmCtrls(CTRL4INDEX).fBoxX, tmCtrls(CTRL4INDEX).fBoxY
            'Acquisition Cost (xxxxxx.xx)
            gSetCtrl tmCtrls(CTRL7INDEX), 30, tmCtrls(CTRL2INDEX).fBoxY + fgStDeltaY, 900, fgBoxStH
            If imAcqCostDefined Then
                Ctrl7.Width = tmCtrls(CTRL7INDEX).fBoxW
                Ctrl7.MaxLength = 9
                gMoveFormCtrl pbcMNm(0), Ctrl7, tmCtrls(CTRL7INDEX).fBoxX, tmCtrls(CTRL7INDEX).fBoxY
            End If
            'Sales Type
            gSetCtrl tmCtrls(CTRL8INDEX), 945, tmCtrls(CTRL7INDEX).fBoxY, 645, fgBoxStH
            Ctrl8.Width = tmCtrls(CTRL8INDEX).fBoxW
            gMoveFormCtrl pbcMNm(0), Ctrl8, tmCtrls(CTRL8INDEX).fBoxX, tmCtrls(CTRL8INDEX).fBoxY
            'Hard Cost
            gSetCtrl tmCtrls(CTRL5INDEX), 1605, tmCtrls(CTRL7INDEX).fBoxY, 660, fgBoxStH
            Ctrl5.Width = tmCtrls(CTRL5INDEX).fBoxW
            gMoveFormCtrl pbcMNm(0), Ctrl5, tmCtrls(CTRL7INDEX).fBoxX, tmCtrls(CTRL5INDEX).fBoxY
            'gSetCtrl tmCtrls(CTRL5INDEX), 1440, tmCtrls(CTRL4INDEX).fBoxY, 1395, fgBoxStH
            gSetCtrl tmCtrls(CTRL6INDEX), 2280, tmCtrls(CTRL7INDEX).fBoxY, 555, fgBoxStH
            If imTaxDefined Then
                'Taxable
                Ctrl6.Width = tmCtrls(CTRL6INDEX).fBoxW
                gMoveFormCtrl pbcMNm(0), Ctrl6, tmCtrls(CTRL6INDEX).fBoxX, tmCtrls(CTRL6INDEX).fBoxY
            Else
                'lacCover.Move tmCtrls(CTRL6INDEX).fBoxX, tmCtrls(CTRL6INDEX).fBoxY
            End If
        Case "A"    'Announcer
            plcMNm.Move 930, 735 + 180, pbcMNm(1).Width + fgPanelAdj, pbcMNm(1).height + fgPanelAdj
            pbcMNm(1).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(1).Visible = True
            Set Ctrl2 = edcRate
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = edcCtrl3
            Ctrl3.ZOrder vbBringToFront   'Place in front
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(1), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Rate per spot (x,xxx,xxx.xx)
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 10
            gMoveFormCtrl pbcMNm(1), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'Group number
            gSetCtrl tmCtrls(CTRL3INDEX), 1440, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl3.Width = tmCtrls(CTRL3INDEX).fBoxW
            Ctrl3.MaxLength = 5
            gMoveFormCtrl pbcMNm(1), Ctrl3, tmCtrls(CTRL3INDEX).fBoxX, tmCtrls(CTRL3INDEX).fBoxY
            tmCtrls(CTRL3INDEX).iReq = False
        Case "S"    'Sales Source
            plcMNm.Move 930, 735, pbcMNm(3).Width + fgPanelAdj, pbcMNm(3).height + fgPanelAdj
            pbcMNm(3).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(3).Visible = True
            Set Ctrl2 = edcComm
            Ctrl2.ZOrder vbBringToFront   'Move to front
            Set Ctrl3 = lbcOrigin
            Ctrl3.ZOrder vbBringToFront
            Set Ctrl4 = pbcYN
            Set Ctrl5 = edcCtrl5
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20  'Jim request on 12/26/01 for Premiere Traffic was 10  'Only name that is 10
            gMoveFormCtrl pbcMNm(3), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Rep commission (xx.xxxx)
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 8
            gMoveFormCtrl pbcMNm(3), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'Origin
            gSetCtrl tmCtrls(CTRL3INDEX), 1440, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            'Update Receivables
            gSetCtrl tmCtrls(CTRL4INDEX), 30, tmCtrls(CTRL3INDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl4.Width = tmCtrls(CTRL4INDEX).fBoxW
            gMoveFormCtrl pbcMNm(3), Ctrl4, tmCtrls(CTRL4INDEX).fBoxX, tmCtrls(CTRL4INDEX).fBoxY
            gSetCtrl tmCtrls(CTRL5INDEX), 1440, tmCtrls(CTRL4INDEX).fBoxY, 1395, fgBoxStH
            Ctrl5.Width = tmCtrls(CTRL5INDEX).fBoxW
            Ctrl5.MaxLength = 15
            gMoveFormCtrl pbcMNm(3), Ctrl5, tmCtrls(CTRL5INDEX).fBoxX, tmCtrls(CTRL5INDEX).fBoxY
            tmCtrls(CTRL5INDEX).iReq = False
        Case "T"    'Sales Team
            plcMNm.Move 930, 735 + 180, pbcMNm(2).Width + fgPanelAdj, pbcMNm(2).height + fgPanelAdj
            pbcMNm(2).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(2).Visible = True
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(2), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
        Case "W"    'Hub
            plcMNm.Move 930, 735 + 180, pbcMNm(2).Width + fgPanelAdj, pbcMNm(2).height + fgPanelAdj
            pbcMNm(2).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(2).Visible = True
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(2), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
        Case "J"    'Terms
            plcMNm.Move 930, 735 + 180, pbcMNm(19).Width + fgPanelAdj, pbcMNm(19).height + fgPanelAdj
            pbcMNm(19).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(19).Visible = True
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(19), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
        Case "L"    'Language
            plcMNm.Move 930, 735 + 180, pbcMNm(21).Width + fgPanelAdj, pbcMNm(21).height + fgPanelAdj
            pbcMNm(21).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(21).Visible = True
            Set Ctrl2 = pbcYN
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(21), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Bill Mg/Missed
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
        Case "Z"    'Team
            plcMNm.Move 930, 735 + 180, pbcMNm(20).Width + fgPanelAdj, pbcMNm(20).height + fgPanelAdj
            pbcMNm(20).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(20).Visible = True
            Set Ctrl2 = edcCtrl3
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(20), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Abbreviation
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 4
            gMoveFormCtrl pbcMNm(20), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
        Case "1"    'Event Subtotal 1
            plcMNm.Move 930, 735 + 180, pbcMNm(2).Width + fgPanelAdj, pbcMNm(2).height + fgPanelAdj
            pbcMNm(2).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(2).Visible = True
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(2), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
        Case "2"    'Event Subtotal 2
            plcMNm.Move 930, 735 + 180, pbcMNm(2).Width + fgPanelAdj, pbcMNm(2).height + fgPanelAdj
            pbcMNm(2).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(2).Visible = True
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(2), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
        Case "R"    'Revenue Sets
            plcMNm.Move 930, 735, pbcMNm(7).Width + fgPanelAdj, pbcMNm(7).height + fgPanelAdj
            pbcMNm(7).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(7).Visible = True
            Set Ctrl2 = pbcYN
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = edcCtrl4
            Ctrl3.ZOrder vbBringToFront   'Place in front
            Set Ctrl4 = edcCtrl5
            Ctrl4.ZOrder vbBringToFront   'Place in front
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(7), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Mandatory/Optional
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            gMoveFormCtrl pbcMNm(7), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'Set #
            gSetCtrl tmCtrls(CTRL3INDEX), 30, tmCtrls(CTRL2INDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl3.Width = tmCtrls(CTRL3INDEX).fBoxW
            Ctrl3.MaxLength = 1
            gMoveFormCtrl pbcMNm(7), Ctrl3, tmCtrls(CTRL3INDEX).fBoxX, tmCtrls(CTRL3INDEX).fBoxY
            'Station Code
            gSetCtrl tmCtrls(CTRL4INDEX), 1440, tmCtrls(CTRL2INDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl4.Width = tmCtrls(CTRL4INDEX).fBoxW
            Ctrl4.MaxLength = 5
            gMoveFormCtrl pbcMNm(7), Ctrl4, tmCtrls(CTRL4INDEX).fBoxX, tmCtrls(CTRL4INDEX).fBoxY
            tmCtrls(CTRL4INDEX).iReq = False
        Case "M"    'Missed Reason
            'plcMNm.Move 930, 735 + 180, pbcMNm(4).Width + fgPanelAdj, pbcMNm(4).Height + fgPanelAdj
            plcMNm.Move 930, 735, pbcMNm(4).Width + fgPanelAdj, pbcMNm(4).height + fgPanelAdj
            pbcMNm(4).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(4).Visible = True
            Set Ctrl2 = pbcYN
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = pbcYN
            Set Ctrl4 = pbcYN
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(4), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            ''Bill Mg/Missed
            'gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1770, fgBoxStH
            'Bill Mg/Missed
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            'Default Reason
            gSetCtrl tmCtrls(CTRL3INDEX), 1440, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            'For
            gSetCtrl tmCtrls(CTRL4INDEX), 30, tmCtrls(CTRL2INDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
        Case "C"    'Product Protection
            plcMNm.Move 930, 735 + 180, pbcMNm(9).Width + fgPanelAdj, pbcMNm(9).height + fgPanelAdj
            pbcMNm(9).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(9).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = edcCtrl4
            Ctrl3.ZOrder vbBringToFront   'Place in front
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(9), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Abbreviation
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 5
            gMoveFormCtrl pbcMNm(9), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'Station Code
            gSetCtrl tmCtrls(CTRL3INDEX), 1440, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl3.Width = tmCtrls(CTRL3INDEX).fBoxW
            Ctrl3.MaxLength = 5
            gMoveFormCtrl pbcMNm(9), Ctrl3, tmCtrls(CTRL3INDEX).fBoxX, tmCtrls(CTRL3INDEX).fBoxY
            tmCtrls(CTRL3INDEX).iReq = False
        Case "X"    'Program Exclusions
            plcMNm.Move 930, 735 + 180, pbcMNm(5).Width + fgPanelAdj, pbcMNm(5).height + fgPanelAdj
            pbcMNm(5).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(5).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(5), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Abbreviation
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 5
            gMoveFormCtrl pbcMNm(5), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
        Case "N"    'Feed
            plcMNm.Move 930, 735, pbcMNm(10).Width + fgPanelAdj, pbcMNm(10).height + fgPanelAdj
            pbcMNm(10).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(10).Visible = True
            Set Ctrl2 = pbcYN
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = pbcYN
            Set Ctrl4 = edcCtrl4
            Ctrl4.ZOrder vbBringToFront   'Place in front
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(10), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Type of Feed
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            gMoveFormCtrl pbcMNm(10), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'SubFeed Allowed
            gSetCtrl tmCtrls(CTRL3INDEX), 1440, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl3.Width = tmCtrls(CTRL3INDEX).fBoxW
            gMoveFormCtrl pbcMNm(10), Ctrl3, tmCtrls(CTRL3INDEX).fBoxX, tmCtrls(CTRL3INDEX).fBoxY
            'Station Code
            gSetCtrl tmCtrls(CTRL4INDEX), 30, tmCtrls(CTRL3INDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl4.Width = tmCtrls(CTRL4INDEX).fBoxW
            Ctrl4.MaxLength = 5
            gMoveFormCtrl pbcMNm(10), Ctrl4, tmCtrls(CTRL4INDEX).fBoxX, tmCtrls(CTRL4INDEX).fBoxY
            tmCtrls(CTRL4INDEX).iReq = False
       Case "V"    'Invoice sort
            plcMNm.Move 930, 735 + 180, pbcMNm(6).Width + fgPanelAdj, pbcMNm(6).height + fgPanelAdj
            pbcMNm(6).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(6).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(6), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Sort order #
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 2
            gMoveFormCtrl pbcMNm(6), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
        Case "G"    'Sales regions
            plcMNm.Move 930, 735 + 180, pbcMNm(6).Width + fgPanelAdj, pbcMNm(6).height + fgPanelAdj
            pbcMNm(6).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(6).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(6), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Sort order #
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 2
            gMoveFormCtrl pbcMNm(6), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
        Case "E"    'Genres
            plcMNm.Move 930, 735 + 180, pbcMNm(8).Width + fgPanelAdj, pbcMNm(8).height + fgPanelAdj
            pbcMNm(8).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(8).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(8), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Group #
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 2
            gMoveFormCtrl pbcMNm(8), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
        Case "Y"    'Transaction Type
            plcMNm.Move 930, 735 + 180, pbcMNm(11).Width + fgPanelAdj, pbcMNm(11).height + fgPanelAdj
            pbcMNm(11).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(11).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Ctrl2.MaxLength = 2
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(11), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Transaction Type
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            gMoveFormCtrl pbcMNm(11), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
       Case "H"    'Vehicle Group
            plcMNm.Move 930, 735 + 180, pbcMNm(12).Width + fgPanelAdj, pbcMNm(12).height + fgPanelAdj
            pbcMNm(12).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(12).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            'Set Ctrl3 = edcCtrl4
            Set Ctrl3 = lbcOrigin
            Ctrl3.ZOrder vbBringToFront   'Place in front
            If imNoCtrls = 3 Then
                Set Ctrl4 = imcUnused
            Else
                Set Ctrl4 = pbcYN
            End If
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            'TTP 10466 - Vehicle Group Names (mnfName): need to expand character limit to 40
            edcName.MaxLength = 50
            gMoveFormCtrl pbcMNm(12), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Sort order #
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 855, fgBoxStH
            tmCtrls(CTRL2INDEX).iReq = False
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 3
            gMoveFormCtrl pbcMNm(12), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'Set #
            'gSetCtrl tmCtrls(CTRL3INDEX), 1440, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            'Ctrl3.Width = tmCtrls(CTRL3INDEX).fBoxW
            'Ctrl3.MaxLength = 2
            'gMoveFormCtrl pbcMnm(12), Ctrl3, tmCtrls(CTRL3INDEX).fBoxX, tmCtrls(CTRL3INDEX).fBoxY
            gSetCtrl tmCtrls(CTRL3INDEX), 900, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1380, fgBoxStH
            If imNoCtrls = 4 Then
                'Dollars
                gSetCtrl tmCtrls(CTRL4INDEX), 2295, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 540, fgBoxStH
                Ctrl4.Width = tmCtrls(CTRL4INDEX).fBoxW
                gMoveFormCtrl pbcMNm(12), Ctrl4, tmCtrls(CTRL4INDEX).fBoxX, tmCtrls(CTRL4INDEX).fBoxY
            End If
        Case "B"    'Business Categories
            plcMNm.Move 930, 735 + 180, pbcMNm(13).Width + fgPanelAdj, pbcMNm(13).height + fgPanelAdj
            pbcMNm(13).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(13).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Ctrl2.MaxLength = 2
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(13), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Type
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            tmCtrls(CTRL2INDEX).iReq = False
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            gMoveFormCtrl pbcMNm(13), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
        Case "P"    'Potential Code
            plcMNm.Move 930, 735, pbcMNm(14).Width + fgPanelAdj, pbcMNm(14).height + fgPanelAdj
            pbcMNm(14).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(14).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Ctrl2.MaxLength = 3
            Set Ctrl3 = edcCtrl4
            Ctrl3.ZOrder vbBringToFront   'Place in front
            Ctrl3.MaxLength = 3
            Set Ctrl4 = edcCtrl5
            Ctrl4.ZOrder vbBringToFront   'Place in front
            Ctrl4.MaxLength = 3
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(14), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Optimistic %
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            gMoveFormCtrl pbcMNm(14), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'Most Likely
            gSetCtrl tmCtrls(CTRL3INDEX), 1440, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl3.Width = tmCtrls(CTRL3INDEX).fBoxW
            gMoveFormCtrl pbcMNm(14), Ctrl3, tmCtrls(CTRL3INDEX).fBoxX, tmCtrls(CTRL3INDEX).fBoxY
            'Pessimistic %
            gSetCtrl tmCtrls(CTRL4INDEX), 30, tmCtrls(CTRL3INDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl4.Width = tmCtrls(CTRL4INDEX).fBoxW
            gMoveFormCtrl pbcMNm(14), Ctrl4, tmCtrls(CTRL4INDEX).fBoxX, tmCtrls(CTRL4INDEX).fBoxY
            'tmCtrls(CTRL4INDEX).iReq = False
       Case "D"    'Custom Demo
            '5/7/10: Assign group number automatically
            pbcMNm(15).height = pbcMNm(2).height
            plcMNm.Move 930, 735 + 180, pbcMNm(15).Width + fgPanelAdj, pbcMNm(15).height + fgPanelAdj
            pbcMNm(15).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(15).Visible = True
            'Set Ctrl2 = edcCtrl3
            'Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(15), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            ''Sort order #
            'gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            'tmCtrls(CTRL2INDEX).iReq = True
            'Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            'Ctrl2.MaxLength = 3
            'gMoveFormCtrl pbcMNm(15), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
       Case "F"    'Soc Eco Groups
            plcMNm.Move 930, 735 + 180, pbcMNm(15).Width + fgPanelAdj, pbcMNm(18).height + fgPanelAdj
            pbcMNm(18).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(18).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 4
            gMoveFormCtrl pbcMNm(18), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Sort order #
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            tmCtrls(CTRL2INDEX).iReq = True
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 20
            gMoveFormCtrl pbcMNm(18), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
        Case "O"    'Competitors
            plcMNm.Move 930, 735 + 180, pbcMNm(16).Width + fgPanelAdj, pbcMNm(16).height + fgPanelAdj
            pbcMNm(16).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(16).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = pbcYN
            Ctrl3.ZOrder vbBringToFront   'Place in front
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(16), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Abbreviation
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 5
            gMoveFormCtrl pbcMNm(16), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
            'Us/Them
            gSetCtrl tmCtrls(CTRL3INDEX), 1440, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 1395, fgBoxStH
            Ctrl3.Width = tmCtrls(CTRL3INDEX).fBoxW
            gMoveFormCtrl pbcMNm(16), Ctrl3, tmCtrls(CTRL3INDEX).fBoxX, tmCtrls(CTRL3INDEX).fBoxY
        Case "K"    'Segments
            plcMNm.Move 930, 735 + 180, pbcMNm(17).Width + fgPanelAdj, pbcMNm(17).height + fgPanelAdj
            pbcMNm(17).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(17).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Ctrl2.MaxLength = 2
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(17), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Abbreviation
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            tmCtrls(CTRL2INDEX).iReq = True
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            gMoveFormCtrl pbcMNm(17), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
       Case "3"    'Daypart Group
            plcMNm.Move 930, 735 + 180, pbcMNm(6).Width + fgPanelAdj, pbcMNm(6).height + fgPanelAdj
            pbcMNm(6).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(6).Visible = True
            Set Ctrl2 = edcCtrl3
            Ctrl2.ZOrder vbBringToFront   'Place in front
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(6), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
            'Sort order #
            gSetCtrl tmCtrls(CTRL2INDEX), 30, tmCtrls(NAMEINDEX).fBoxY + fgStDeltaY, 2805, fgBoxStH
            Ctrl2.Width = tmCtrls(CTRL2INDEX).fBoxW
            Ctrl2.MaxLength = 2
            gMoveFormCtrl pbcMNm(6), Ctrl2, tmCtrls(CTRL2INDEX).fBoxX, tmCtrls(CTRL2INDEX).fBoxY
        Case "4"    'Copy Type
            plcMNm.Move 930, 735 + 180, pbcMNm(2).Width + fgPanelAdj, pbcMNm(2).height + fgPanelAdj
            pbcMNm(2).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(2).Visible = True
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(2), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
        Case "5"    'Podcast Category
            plcMNm.Move 930, 735 + 180, pbcMNm(2).Width + fgPanelAdj, pbcMNm(2).height + fgPanelAdj
            pbcMNm(2).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(2).Visible = True
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(2), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
        Case "6"    'Position
            pbcMNm(20).height = 360
            plcMNm.Move 930, 735 + 180, pbcMNm(20).Width + fgPanelAdj, pbcMNm(20).height + fgPanelAdj
            pbcMNm(20).Move plcMNm.Left + fgBevelX, plcMNm.Top + fgBevelY
            pbcMNm(20).Visible = True
            Set Ctrl2 = imcUnused
            Set Ctrl3 = imcUnused
            Set Ctrl4 = imcUnused
            Set Ctrl5 = imcUnused
            Set Ctrl6 = imcUnused
            'Name
            gSetCtrl tmCtrls(NAMEINDEX), 30, 30, 2805, fgBoxStH
            edcName.Width = tmCtrls(NAMEINDEX).fBoxW
            edcName.MaxLength = 20
            gMoveFormCtrl pbcMNm(20), edcName, tmCtrls(NAMEINDEX).fBoxX, tmCtrls(NAMEINDEX).fBoxY
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
    Dim slStr As String     'String buffer
    If Not ilTestChg Or tmCtrls(NAMEINDEX).iChg Then
        If (smMnfCallType <> "F") Then
            tmMnf.sName = edcName.Text
        Else
            tmMnf.sUnitType = edcName.Text
        End If
    End If
    If (smMnfCallType = "I") Or (smMnfCallType = "A") Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            slStr = Ctrl2.Text
            gStrToPDN slStr, 2, 5, tmMnf.sRPU
        End If
    ElseIf smMnfCallType = "P" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            tmMnf.sRPU = slStr
            slStr = Ctrl2.Text  'Optimistic %
            gStrToPDN slStr, 2, 5, tmMnf.sRPU
        End If
    Else
        tmMnf.sRPU = ""
    End If
    'Product Protection or revenue sets use station codes
    If (smMnfCallType = "C") Then
        If Not ilTestChg Or tmCtrls(CTRL3INDEX).iChg Then
            slStr = Ctrl3.Text
            tmMnf.sCodeStn = slStr
        End If
    ElseIf (smMnfCallType = "R") Then
        If Not ilTestChg Or tmCtrls(CTRL4INDEX).iChg Then
            slStr = Ctrl4.Text
            tmMnf.sCodeStn = slStr
        End If
    ElseIf (smMnfCallType = "N") Then
        If Not ilTestChg Or tmCtrls(CTRL4INDEX).iChg Then
            slStr = Ctrl4.Text
            tmMnf.sCodeStn = slStr
        End If
    ElseIf (smMnfCallType = "O") Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            slStr = Ctrl2.Text
            tmMnf.sCodeStn = slStr
        End If
    ElseIf smMnfCallType = "M" Then
        If Not ilTestChg Or tmCtrls(CTRL4INDEX).iChg Then
            Select Case imMissedFor
                Case 1
                    tmMnf.sCodeStn = "T"    'Network Missed replacing Traffic
                Case 2
                    tmMnf.sCodeStn = "A"    'Station Missed replacing Affiliate Web
                Case 3
                    tmMnf.sCodeStn = "B"    'Network + Station Missed replacing Both
                Case 4
                    tmMnf.sCodeStn = "R"    'Station Replacement (new)
                Case Else
                    tmMnf.sCodeStn = "T"    'Network Missed replacing Traffic
            End Select
        End If
    Else
        tmMnf.sCodeStn = ""
    End If
    tmMnf.lCost = 0
    If smMnfCallType = "I" Then
        If Not ilTestChg Or tmCtrls(CTRL3INDEX).iChg Then
            slStr = Ctrl3.Text
            'tmMnf.sUnitType = slStr
            tmMnf.sUnitsPer = slStr
        End If
        If Not ilTestChg Or tmCtrls(CTRL4INDEX).iChg Then
            slStr = Ctrl4.Text
            gStrToPDN slStr, 4, 4, tmMnf.sSSComm
        End If
        Select Case imHardCost
            Case 0  'Yes
                tmMnf.sCodeStn = "Y"  'Yes
            Case 1  'No
                tmMnf.sCodeStn = "N"  'No
            Case Else
                tmMnf.sCodeStn = "N"  'No
        End Select
        If imHardCost <> 0 Then
            Select Case imTaxable
                Case 0  'Yes
                    tmMnf.iGroupNo = 1  'Yes
                Case 1  'No
                    tmMnf.iGroupNo = 0  'No
                Case Else
                    tmMnf.iGroupNo = 0  'No
            End Select
        Else
            tmMnf.iGroupNo = 0
        End If
        If imAcqCostDefined Then
            If Not ilTestChg Or tmCtrls(CTRL7INDEX).iChg Then
                slStr = Ctrl7.Text
                tmMnf.lCost = gStrDecToLong(slStr, 2)
            End If
        End If
        Select Case imSaleType
            Case 1
                tmMnf.sUnitType = "A"
            Case 2
                tmMnf.sUnitType = "D"
            Case Else
                tmMnf.sUnitType = "N"
            End Select
    ElseIf (smMnfCallType = "Y") Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            slStr = Ctrl2.Text
            tmMnf.sUnitType = slStr
        End If
    ElseIf (smMnfCallType = "B") Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            slStr = Ctrl2.Text
            tmMnf.sUnitType = slStr
        End If
    ElseIf (smMnfCallType = "K") Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            slStr = Ctrl2.Text
            tmMnf.sUnitType = slStr
        End If
    ElseIf (smMnfCallType = "Z") Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            slStr = Ctrl2.Text
            tmMnf.sUnitType = slStr
        End If
    ElseIf (smMnfCallType = "H") Then
        If Not ilTestChg Or tmCtrls(CTRL3INDEX).iChg Then
            'slStr = Ctrl3.Text
            If imVehGpSetNo <= 0 Then
                slStr = Trim$(str$(lbcOrigin.ListIndex + 1))
            Else
                slStr = Trim$(str$(imVehGpSetNo))
            End If
            tmMnf.sUnitType = slStr
        End If
    ElseIf (smMnfCallType = "P") Then
        If Not ilTestChg Or tmCtrls(CTRL3INDEX).iChg Then
            slStr = Ctrl3.Text
            tmMnf.sUnitType = slStr
        End If
    ElseIf (smMnfCallType = "M") Then
        If Not ilTestChg Or tmCtrls(CTRL3INDEX).iChg Then
            If imDefReason = 0 Then
                tmMnf.sUnitType = "Y"
            Else
                tmMnf.sUnitType = "N"
            End If
        End If
    ElseIf (smMnfCallType = "R") Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If imManOpt = 1 Then
                tmMnf.sUnitType = "M"
            Else
                tmMnf.sUnitType = "O"
            End If
        End If
    ElseIf (smMnfCallType = "J") Then
        tmMnf.sUnitType = "T"   'Indicate which type of J record (T=Terms-Set here; D=Default Terms-Set in Site; A=Client Abbreviation-Set in Site)
    Else
        If (smMnfCallType <> "C") And (smMnfCallType <> "N") And (smMnfCallType <> "X") And (smMnfCallType <> "S") And (smMnfCallType <> "F") And (smMnfCallType <> "J") Then
            tmMnf.sUnitType = ""
        End If
    End If
    If (smMnfCallType = "L") Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If imEnglish = 0 Then
                tmMnf.sUnitType = "Y"
            Else
                tmMnf.sUnitType = "N"
            End If
        End If
    End If
    If smMnfCallType = "A" Then
        If Ctrl3.Text = "" Then
            tmMnf.iGroupNo = 0
        Else
            tmMnf.iGroupNo = Val(Ctrl3.Text)
        End If
    ElseIf smMnfCallType = "V" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If Ctrl2.Text = "" Then
                tmMnf.iGroupNo = 0
            Else
                tmMnf.iGroupNo = Val(Ctrl2.Text)
            End If
        End If
    ElseIf smMnfCallType = "G" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If Ctrl2.Text = "" Then
                tmMnf.iGroupNo = 0
            Else
                tmMnf.iGroupNo = Val(Ctrl2.Text)
            End If
        End If
    ElseIf smMnfCallType = "R" Then
        If Not ilTestChg Or tmCtrls(CTRL3INDEX).iChg Then
            If Ctrl3.Text = "" Then
                tmMnf.iGroupNo = 0
            Else
                tmMnf.iGroupNo = Val(Ctrl3.Text)
            End If
        End If
    ElseIf smMnfCallType = "E" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If Ctrl2.Text = "" Then
                tmMnf.iGroupNo = 0
            Else
                tmMnf.iGroupNo = Val(Ctrl2.Text)
            End If
        End If
    ElseIf smMnfCallType = "M" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            Select Case imBillMGMissed
                Case 1
                    tmMnf.iGroupNo = 1  'Bill MG, not Missed
                Case 2
                    tmMnf.iGroupNo = 2  'Bill Missed, not MG
                Case 3
                    tmMnf.iGroupNo = 3  'Bill MG & Missed
                Case Else
                    tmMnf.iGroupNo = 1  'Bill MG, not Missed
            End Select
        End If
    ElseIf smMnfCallType = "N" Then
        If Not ilTestChg Or tmCtrls(CTRL3INDEX).iChg Then
            Select Case imSubFeedAllowed
                Case 0  'Yes
                    tmMnf.iGroupNo = 1  'Yes
                Case 1  'No
                    tmMnf.iGroupNo = 0  'No
                Case Else
                    tmMnf.iGroupNo = 0  'No
            End Select
        End If
    ElseIf smMnfCallType = "H" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If Ctrl2.Text = "" Then
                tmMnf.iGroupNo = 0
            Else
                tmMnf.iGroupNo = Val(Ctrl2.Text)
            End If
        End If
    ElseIf smMnfCallType = "D" Then
        '5/7/10: Assign group number automatically
        'If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
        '    If Ctrl2.Text = "" Then
        '        tmMnf.iGroupNo = 0
        '    Else
        '        tmMnf.iGroupNo = Val(Ctrl2.Text)
        '    End If
        'End If
    ElseIf (smMnfCallType = "1") Then
        tmMnf.iGroupNo = 1   'Indicate which Subtotal
    ElseIf (smMnfCallType = "2") Then
        tmMnf.iGroupNo = 2   'Indicate which Subtotal
    ElseIf smMnfCallType = "F" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            tmMnf.sName = Ctrl2.Text
        End If
    ElseIf smMnfCallType = "O" Then
        If Not ilTestChg Or tmCtrls(CTRL3INDEX).iChg Then
            If imUsThem = 1 Then
                tmMnf.iGroupNo = 1
            Else
                tmMnf.iGroupNo = 2
            End If
        End If
    ElseIf smMnfCallType = "3" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If Ctrl2.Text = "" Then
                tmMnf.iGroupNo = 0
            Else
                tmMnf.iGroupNo = Val(Ctrl2.Text)
            End If
        End If
    ElseIf smMnfCallType = "6" Then
        tmMnf.iGroupNo = imAdServerCode
    ElseIf (smMnfCallType <> "I") And (smMnfCallType <> "S") Then   'For I group # set above, S set below
        tmMnf.iGroupNo = 0
    End If
    If smMnfCallType = "S" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            slStr = Ctrl2.Text
            gStrToPDN slStr, 4, 4, tmMnf.sSSComm
        End If
        If lbcOrigin.ListIndex > -1 Then
            tmMnf.iGroupNo = lbcOrigin.ListIndex + 1
        Else
            tmMnf.iGroupNo = 1 'default to local
        End If
        If Not ilTestChg Or tmCtrls(CTRL4INDEX).iChg Then
            If imUpdateRvf = 1 Then
                tmMnf.sUnitType = "N"   'History
            ElseIf imUpdateRvf = 2 Then
                tmMnf.sUnitType = "E"   'Export+History
            ElseIf imUpdateRvf = 3 Then
                tmMnf.sUnitType = "F"   'Export+A/R
            ElseIf imUpdateRvf = 4 Then
                tmMnf.sUnitType = "A"   'Ask
            Else
                tmMnf.sUnitType = "Y"   'Receivables
            End If
        End If
        If Not ilTestChg Or tmCtrls(CTRL5INDEX).iChg Then
            tmMnf.sUnitsPer = Ctrl5.Text
        End If
    ElseIf (smMnfCallType = "P") Then
        If Not ilTestChg Or tmCtrls(CTRL4INDEX).iChg Then
            'slStr = Ctrl4.Text
            'tmMnf.sSSComm = slStr
            slStr = Ctrl4.Text
            gStrToPDN slStr, 4, 4, tmMnf.sSSComm
        End If
    ElseIf (smMnfCallType = "H") Then
        If (Trim$(tmMnf.sUnitType) = "5") Or (Trim$(tmMnf.sUnitType) = "3") Then
            If Not ilTestChg Or tmCtrls(CTRL4INDEX).iChg Then
                If (Trim$(tmMnf.sUnitType) = "5") Then
                    If imDollars = 1 Then
                        tmMnf.sRPU = "N"
                    Else
                        tmMnf.sRPU = "Y"
                    End If
                Else
                    If imDollars = 0 Then
                        tmMnf.sRPU = "Y"
                    Else
                        tmMnf.sRPU = "N"
                    End If
                End If
            End If
        Else
            tmMnf.sRPU = ""
        End If
    ElseIf smMnfCallType <> "I" Then
        tmMnf.sSSComm = ""
    End If
    If smMnfCallType = "C" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If Ctrl2.Text = "" Then
                tmMnf.sUnitType = Trim$(Left$(edcName.Text, 5))
            Else
                tmMnf.sUnitType = Ctrl2.Text
            End If
        End If
    End If
    If smMnfCallType = "X" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If Ctrl2.Text = "" Then
                tmMnf.sUnitType = Trim$(Left$(edcName.Text, 5))
            Else
                tmMnf.sUnitType = Ctrl2.Text
            End If
        End If
    End If
    If smMnfCallType = "N" Then
        If Not ilTestChg Or tmCtrls(CTRL2INDEX).iChg Then
            If imTypeOfFeed = 1 Then
                tmMnf.sUnitType = "A"
            ElseIf imTypeOfFeed = 2 Then
                tmMnf.sUnitType = "C"
            ElseIf imTypeOfFeed = 3 Then
                tmMnf.sUnitType = "S"
            Else
                tmMnf.sUnitType = "D"
            End If
        End If
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
'
'   mMoveRecToCtrl
'   Where:
'
    Dim slStr As String
    Dim ilLoop As Integer
    If (smMnfCallType <> "F") Then
        edcName.Text = Trim$(tmMnf.sName)
    Else
        edcName.Text = Trim$(tmMnf.sUnitType)
    End If
    If (smMnfCallType = "I") Or (smMnfCallType = "A") Then
        gPDNToStr tmMnf.sRPU, 2, slStr
        Ctrl2.Text = slStr
    End If
    If smMnfCallType = "I" Then
        'If tmMnf.sUnitType = "" Then
        '    Ctrl3.Text = ""
        'Else
        '    Ctrl3.Text = Trim$(tmMnf.sUnitType)
        'End If
        If tmMnf.sUnitsPer = "" Then
            Ctrl3.Text = ""
        Else
            Ctrl3.Text = Trim$(tmMnf.sUnitsPer)
        End If
        gPDNToStr tmMnf.sSSComm, 4, slStr
        Ctrl4.Text = slStr
        If Trim$(tmMnf.sCodeStn) = "Y" Then
            imHardCost = 0 'Yes
        Else
            imHardCost = 1 'No
        End If
        If tmMnf.iGroupNo = 1 Then
            imTaxable = 0 'Yes
        Else
            imTaxable = 1 'No
        End If
        If imAcqCostDefined Then
            Ctrl7.Text = gLongToStrDec(tmMnf.lCost, 2)
        Else
            'Ctrl7.Text = ""
        End If
        If Trim$(tmMnf.sUnitType) = "A" Then
            imSaleType = 1  'Agency
        ElseIf Trim$(tmMnf.sUnitType) = "D" Then
            imSaleType = 2  'Direct
        Else
            imSaleType = 0  'NTR
        End If
    End If
    If smMnfCallType = "L" Then
        If Trim$(tmMnf.sUnitType) = "N" Then
            imEnglish = 1
        Else
            imEnglish = 0
        End If
    End If
    If smMnfCallType = "A" Then
        Ctrl3.Text = Trim$(str$(tmMnf.iGroupNo))
    End If
    If smMnfCallType = "S" Then
        gPDNToStr tmMnf.sSSComm, 4, slStr
        Ctrl2.Text = slStr
        If tmMnf.iGroupNo > 0 Then
            Ctrl3.ListIndex = tmMnf.iGroupNo - 1
        Else
            Ctrl3.Text = ""
        End If
        If Trim$(tmMnf.sUnitType) = "N" Then
            imUpdateRvf = 1 'History
        ElseIf Trim$(tmMnf.sUnitType) = "E" Then
            imUpdateRvf = 2 'Export+History
        ElseIf Trim$(tmMnf.sUnitType) = "F" Then
            imUpdateRvf = 3 'Export+A/R
        ElseIf Trim$(tmMnf.sUnitType) = "A" Then
            imUpdateRvf = 4 'Ask
        ElseIf Trim$(tmMnf.sUnitType) = "Y" Then
            imUpdateRvf = 0 'Receivables
        Else
            imUpdateRvf = -1
        End If
        Ctrl5.Text = Trim$(tmMnf.sUnitsPer)
    End If
    If smMnfCallType = "C" Then
        Ctrl2.Text = Trim$(tmMnf.sUnitType)
        Ctrl3.Text = Trim$(tmMnf.sCodeStn)
    End If
    If smMnfCallType = "X" Then
        Ctrl2.Text = Trim$(tmMnf.sUnitType)
    End If
    If smMnfCallType = "Z" Then
        Ctrl2.Text = Trim$(tmMnf.sUnitType)
    End If
    If smMnfCallType = "R" Then
        If Trim$(tmMnf.sUnitType) = "M" Then
            imManOpt = 1
        Else
            imManOpt = 2
        End If
        Ctrl3.Text = Trim$(str$(tmMnf.iGroupNo))
        Ctrl4.Text = Trim$(tmMnf.sCodeStn)
    End If
    If smMnfCallType = "V" Then
        Ctrl2.Text = Trim$(str$(tmMnf.iGroupNo))
    End If
    If smMnfCallType = "G" Then
        Ctrl2.Text = Trim$(str$(tmMnf.iGroupNo))
    End If
    If smMnfCallType = "E" Then
        Ctrl2.Text = Trim$(str$(tmMnf.iGroupNo))
    End If
    If smMnfCallType = "M" Then
        If Trim$(tmMnf.sUnitType) = "Y" Then
            imDefReason = 0
        ElseIf Trim$(tmMnf.sUnitType) = "N" Then
            imDefReason = 1
        Else
            imDefReason = -1
        End If
        If tmMnf.iGroupNo = 1 Then
            imBillMGMissed = 1  'Bill MG, not Missed
        ElseIf tmMnf.iGroupNo = 2 Then
            imBillMGMissed = 2  'Bill Missed, not MG
        ElseIf tmMnf.iGroupNo = 3 Then
            imBillMGMissed = 3  'Bill MG and Missed
        Else
            imBillMGMissed = -1
        End If
        If Trim$(tmMnf.sCodeStn) = "T" Then
            imMissedFor = 1
        ElseIf Trim$(tmMnf.sCodeStn) = "A" Then
            imMissedFor = 2
        ElseIf Trim$(tmMnf.sCodeStn) = "B" Then
            imMissedFor = 3
        ElseIf Trim$(tmMnf.sCodeStn) = "R" Then
            imMissedFor = 4
        Else
            imMissedFor = -1
        End If
    End If
    If smMnfCallType = "N" Then
        If Trim$(tmMnf.sUnitType) = "A" Then
            imTypeOfFeed = 1
        ElseIf Trim$(tmMnf.sUnitType) = "C" Then
            imTypeOfFeed = 2
        ElseIf Trim$(tmMnf.sUnitType) = "S" Then
            imTypeOfFeed = 3
        ElseIf Trim$(tmMnf.sUnitType) = "D" Then
            imTypeOfFeed = 0
        Else
            imTypeOfFeed = -1
        End If
        If tmMnf.iGroupNo = 1 Then
            imSubFeedAllowed = 0 'Yes
        Else
            imSubFeedAllowed = 1 'No
        End If
        Ctrl4.Text = Trim$(tmMnf.sCodeStn)
    End If
    If smMnfCallType = "Y" Then
        Ctrl2.Text = Trim$(tmMnf.sUnitType)
    End If
    If smMnfCallType = "B" Then
        Ctrl2.Text = Trim$(tmMnf.sUnitType)
    End If
    If smMnfCallType = "K" Then
        Ctrl2.Text = Trim$(tmMnf.sUnitType)
    End If
    If smMnfCallType = "H" Then
        Ctrl2.Text = Trim$(str$(tmMnf.iGroupNo))
        imDollars = -1
        'Ctrl3.Text = Trim$(tmMnf.sUnitType)
        Select Case Trim$(tmMnf.sUnitType)
            Case "1"
                If imVehGpSetNo <= 0 Then
                    lbcOrigin.ListIndex = 0
                Else
                    lbcOrigin.ListIndex = 0
                End If
            Case "2"
                If imVehGpSetNo <= 0 Then
                    lbcOrigin.ListIndex = 1
                Else
                    lbcOrigin.ListIndex = 0
                End If
            Case "3"
                If imVehGpSetNo <= 0 Then
                    lbcOrigin.ListIndex = 2
                Else
                    lbcOrigin.ListIndex = 0
                End If
                If Trim$(tmMnf.sRPU) = "Y" Then
                    imDollars = 0
                Else
                    imDollars = 1
                End If
            Case "4"
                If imVehGpSetNo <= 0 Then
                    lbcOrigin.ListIndex = 3
                Else
                    lbcOrigin.ListIndex = 0
                End If
            Case "5"
                If imVehGpSetNo <= 0 Then
                    lbcOrigin.ListIndex = 4
                Else
                    lbcOrigin.ListIndex = 0
                End If
                If Trim$(tmMnf.sRPU) = "N" Then
                    imDollars = 1
                Else
                    imDollars = 0
                End If
            Case "6"
                If imVehGpSetNo <= 0 Then
                    lbcOrigin.ListIndex = 5
                Else
                    lbcOrigin.ListIndex = 0
                End If
            Case Else
                lbcOrigin.ListIndex = -1
        End Select
    End If
    If smMnfCallType = "P" Then
        gPDNToStr tmMnf.sRPU, 2, slStr
        Ctrl2.Text = slStr  'Trim$(tmMnf.sRPU)
        Ctrl3.Text = Trim$(tmMnf.sUnitType)
        gPDNToStr tmMnf.sSSComm, 4, slStr
        Ctrl4.Text = slStr  'Trim$(tmMnf.sSSComm)
    End If
    If smMnfCallType = "D" Then
        '5/7/10: Assign group number automatically
        'Ctrl2.Text = Trim$(str$(tmMnf.iGroupNo))
    End If
    If smMnfCallType = "F" Then
        Ctrl2.Text = Trim$(tmMnf.sName)
    End If
    If smMnfCallType = "O" Then
        Ctrl2.Text = Trim$(tmMnf.sCodeStn)
        If tmMnf.iGroupNo = 1 Then
            imUsThem = 1 'Us
        Else
            imUsThem = 2 'Them
        End If
    End If
    If smMnfCallType = "3" Then
        Ctrl2.Text = Trim$(str$(tmMnf.iGroupNo))
    End If
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
Private Function mOKName(ilTestNameOnly As Integer)
    Dim slStr As String
    Dim tlMnf As MNF
    Dim ilGroupNo As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    If edcName.Text <> "" Then    'Test name
        slStr = edcName.Text
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If edcName.Text = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    If smMnfCallType = "M" Then 'Missed Reason
                        MsgBox "Reason already defined, enter a different Reason", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    ElseIf smMnfCallType = "Y" Then 'Transaction Type
                        MsgBox "Description already defined, enter a different Description", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    Else
                        MsgBox "Name already specified, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    End If
                    edcName.Text = Trim$(tmMnf.sName) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
        '5/7/10: Assign group number automatically
        ''If Custom demo- check Uniqueness of GroupNo
        'If (smMnfCallType = "D") And (Not ilTestNameOnly) Then
        '    If imSelectedIndex = 0 Then 'New selected
        '        tmMnf.iCode = 0
        '    End If
        '    If Ctrl2.Text = "" Then
        '        ilGroupNo = 0
        '    Else
        '        ilGroupNo = Val(Ctrl2.Text)
        '    End If
        '    For ilLoop = 0 To UBound(tgNameCode) - 1 Step 1 'lbcNameCode.ListCount - 1 Step 1
        '        slNameCode = tgNameCode(ilLoop).sKey   'lbcNameCode.List(ilLoop)
        '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        '        slCode = Trim$(slCode)
        '        tmMnfSrchKey.iCode = CInt(slCode)
        '        If tmMnfSrchKey.iCode <> tmMnf.iCode Then
        '            ilRet = btrGetEqual(hmMnf, tlMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        '            If ilRet = BTRV_ERR_NONE Then
        '                If ilGroupNo = tlMnf.iGroupNo Then
        '                    MsgBox "Sort Order # already specified, enter a different #", vbOkOnly + vbExclamation + vbApplicationModal, "Error"
        '                    mSetShow imBoxNo
        '                    mSetChg imBoxNo
        '                    imBoxNo = 2
        '                    mEnableBox imBoxNo
        '                    mOKName = False
        '                    Exit Function
        '                End If
        '            End If
        '        End If
        '    Next ilLoop
        'End If
    End If
    mOKName = True
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
    slCommand = sgCommandStr    'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
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
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone MultiNm, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", sgMnfCallType)    'Get call type "L" or "A"
    ilRet = gParseItem(slCommand, 4, "\", slStr)    'Get call source
    igMNmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igMNmCallSource = CALLNONE
    '    sgMnfCallType = "K" '"S" '"D"
    '    imVehGpSetNo = 0    '0=All; 1=Participant; 2=Subtotal; 3=Market; 4=Format; 5=Research
    'End If
    If igMNmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 5, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgMNmName = slStr
        Else
            sgMNmName = ""
        End If
        If (sgMnfCallType = "Y") And (sgMNmName <> "") Then 'Remove the transaction type
            sgMNmName = right$(sgMNmName, Len(sgMNmName) - 3)
        End If
        ilRet = gParseItem(slCommand, 6, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            imVehGpSetNo = Val(slStr)
        Else
            imVehGpSetNo = 0
        End If
        If sgMnfCallType = "6" Then  'Ad Server
            ilRet = gParseItem(sgMNmName, 1, "/", smAdServerName)
            ilRet = gParseItem(sgMNmName, 2, "/", slStr)
            imAdServerCode = Val(slStr)
            sgMNmName = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    Dim slType As String
    Dim slSyncDate As String
    Dim slSyncTime As String

    If (smMnfCallType <> "H") And (smMnfCallType <> "F") Then
        If smMnfCallType = "D" Then
            ilRet = mAddStdDemo()
            ReDim ilFilter(0 To 1) As Integer
            ReDim slFilter(0 To 1) As String
            ReDim ilOffSet(0 To 1) As Integer
            ilFilter(0) = CHARFILTER
            slFilter(0) = smMnfCallType
            ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
            ilFilter(1) = INTEGERFILTERNOT
            slFilter(1) = "0"
            ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
        ElseIf smMnfCallType = "J" Then
            ReDim ilFilter(0 To 1) As Integer
            ReDim slFilter(0 To 1) As String
            ReDim ilOffSet(0 To 1) As Integer
            ilFilter(0) = CHARFILTER
            slFilter(0) = smMnfCallType
            ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
            ilFilter(1) = CHARFILTER
            slFilter(1) = "T"
            ilOffSet(1) = gFieldOffset("Mnf", "MnfUnitType") '2
        ElseIf smMnfCallType = "1" Then
            ilRet = mAddStdDemo()
            ReDim ilFilter(0 To 1) As Integer
            ReDim slFilter(0 To 1) As String
            ReDim ilOffSet(0 To 1) As Integer
            ilFilter(0) = CHARFILTER
            slFilter(0) = "1"   'smMnfCallType
            ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
            ilFilter(1) = INTEGERFILTER
            slFilter(1) = "1"
            ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
        ElseIf smMnfCallType = "2" Then
            ilRet = mAddStdDemo()
            ReDim ilFilter(0 To 1) As Integer
            ReDim slFilter(0 To 1) As String
            ReDim ilOffSet(0 To 1) As Integer
            ilFilter(0) = CHARFILTER
            slFilter(0) = "2"   'smMnfCallType
            ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
            ilFilter(1) = INTEGERFILTER
            slFilter(1) = "2"
            ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
        ElseIf smMnfCallType = "6" Then
            ReDim ilFilter(0 To 1) As Integer
            ReDim slFilter(0 To 1) As String
            ReDim ilOffSet(0 To 1) As Integer
            ilFilter(0) = CHARFILTER
            slFilter(0) = "6"   'smMnfCallType
            ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
            ilFilter(1) = INTEGERFILTER
            slFilter(1) = imAdServerCode
            ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
        Else
            ReDim ilFilter(0) As Integer
            ReDim slFilter(0) As String
            ReDim ilOffSet(0) As Integer
            ilFilter(0) = CHARFILTER
            slFilter(0) = smMnfCallType
            ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
        End If
        imPopReqd = False
        'ilRet = gIMoveListBox(MultiNm, cbcSelect, lbcNameCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(MultiNm, cbcSelect, tgNameCode(), sgNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    ElseIf smMnfCallType = "H" Then
        If imVehGpSetNo <= 0 Then
            slType = "H"
        Else
            slType = "H" & Trim$(str$(imVehGpSetNo))
        End If
        ilRet = gPopMnfPlusFieldsBox(MultiNm, cbcSelect, tgNameCode(), sgNameCodeTag, slType)
    ElseIf smMnfCallType = "F" Then
        'Moved to Research screen
        'ilRet = mAddStdDemo()
        slType = "F"    '"FG"
        ilRet = gPopMnfPlusFieldsBox(MultiNm, cbcSelect, tgNameCode(), sgNameCodeTag, slType)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        If smMnfCallType = "Y" Then
            If cbcSelect.ListCount <= 0 Then
                gGetSyncDateTime slSyncDate, slSyncTime
                For ilLoop = 1 To 9 Step 1
                'Add hard coded transactions
                    tmMnf.iCode = 0
                    tmMnf.sType = smMnfCallType
                    Select Case ilLoop
                        Case 1  'Invoice
                            tmMnf.sName = "Invoice"
                            tmMnf.sUnitType = "IN"
                        Case 2  'Invoice Adjustment
                            tmMnf.sName = "Invoice Adjustment"
                            tmMnf.sUnitType = "AN"
                        Case 3  'Payment to Invoice
                            tmMnf.sName = "Payment to Invoice"
                            tmMnf.sUnitType = "PI"
                        Case 4  'Payment On Account
                            tmMnf.sName = "Payment On Account"
                            tmMnf.sUnitType = "PO"
                        Case 5  'Write-Off Bad Debt
                            tmMnf.sName = "Write-Off Bad Debt"
                            tmMnf.sUnitType = "WB"
                        Case 6  'Write-Off Variance
                            tmMnf.sName = "Write-Off Variance"
                            tmMnf.sUnitType = "WV"
                        Case 7  'Transfer
                            tmMnf.sName = "Transfer"
                            tmMnf.sUnitType = "WT"
                        Case 8  'Transfer
                            tmMnf.sName = "Returned Check"
                            tmMnf.sUnitType = "WU"
                        Case 9  'Transfer
                            tmMnf.sName = "Redeposit Check"
                            tmMnf.sUnitType = "WD"
                    End Select
                    tmMnf.sRPU = ""
                    tmMnf.sSSComm = ""
                    tmMnf.iMerge = 0
                    tmMnf.iGroupNo = 0
                    tmMnf.sCodeStn = ""
                    tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmMnf.iAutoCode = tmMnf.iCode
                    ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
                    Do
                        'tmMnfSrchKey.iCode = tmMnf.iCode
                        'ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
                        tmMnf.iAutoCode = tmMnf.iCode
                        gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
                        gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
                        ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                Next ilLoop
                'lbcNameCode.Tag = ""
                sgNameCodeTag = ""
                cbcSelect.Tag = ""
                'ilRet = gIMoveListBox(MultiNm, cbcSelect, lbcNameCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
                ilRet = gIMoveListBox(MultiNm, cbcSelect, tgNameCode(), sgNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
            End If
        End If
        If smMnfCallType = "D" Then
            imMaxCustomDemoNumber = 0
            For ilLoop = LBound(tgDemoMnf) To UBound(tgDemoMnf) - 1 Step 1
                If tgDemoMnf(ilLoop).iGroupNo > imMaxCustomDemoNumber Then
                    imMaxCustomDemoNumber = tgDemoMnf(ilLoop).iGroupNo
                End If
            Next ilLoop
        End If
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", MultiNm
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
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
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer)
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    slNameCode = tgNameCode(ilSelectIndex - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", MultiNm
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmMnfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "EMnRead (btrGetEqual)", MultiNm
    On Error GoTo 0
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
'*             Created:4/21/93       By:D. LeVine      *
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
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim slStamp As String   'Date/Time stamp for file
    Screen.MousePointer = vbHourglass  'Wait
    mSetShow imBoxNo
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        Screen.MousePointer = vbDefault  'Wait
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName(False) Then
        Screen.MousePointer = vbDefault  'Wait
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    gGetSyncDateTime slSyncDate, slSyncTime
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Mnf.btr")
        'If Len(lbcNameCode.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(lbcNameCode.Tag, Len(lbcNameCode.Tag) - Len(slStamp))
        'End If
        If Len(sgNameCodeTag) > Len(slStamp) Then
            slStamp = slStamp & right$(sgNameCodeTag, Len(sgNameCodeTag) - Len(slStamp))
        End If
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            tmMnf.iCode = 0  'Autoincrement
            tmMnf.sType = smMnfCallType 'L=Lock box; A=Agency DP; S=Sales Office
            If smMnfCallType = "D" Then
                imMaxCustomDemoNumber = imMaxCustomDemoNumber + 1
                tmMnf.iGroupNo = imMaxCustomDemoNumber
            End If
            tmMnf.iMerge = 0   'Merge code number
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
            gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
            ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
            slMsg = "mSaveRec (btrUpdate)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, MultiNm
    On Error GoTo 0
    If imSelectedIndex = 0 Then 'New selected
        Do
            'tmMnfSrchKey.iCode = tmMnf.iCode
            'ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
            gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
            ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    ''If lbcNameCode.Tag <> "" Then
    ''    If slStamp = lbcNameCode.Tag Then
    ''        lbcNameCode.Tag = FileDateTime(sgDBPath & "Mnf.btr")
    ''        If Len(slStamp) > Len(lbcNameCode.Tag) Then
    ''            lbcNameCode.Tag = lbcNameCode.Tag & Right$(slStamp, Len(slStamp) - Len(lbcNameCode.Tag))
    ''        End If
    ''    End If
    ''End If
    'If sgNameCodeTag <> "" Then
    '    If slStamp = sgNameCodeTag Then
    '        sgNameCodeTag = FileDateTime(sgDBPath & "Mnf.btr")
    '        If Len(slStamp) > Len(sgNameCodeTag) Then
    '            sgNameCodeTag = sgNameCodeTag & Right$(slStamp, Len(slStamp) - Len(sgNameCodeTag))
    '        End If
    '    End If
    'End If
    'If imSelectedIndex <> 0 Then
    '    'lbcNameCode.RemoveItem imSelectedIndex - 1
    '    gRemoveItemFromSortCode imSelectedIndex - 1, tgNameCode()
    '    cbcSelect.RemoveItem imSelectedIndex
    'End If
    'cbcSelect.RemoveItem 0 'Remove [New]
    'slName = RTrim$(tmMnf.sName)
    'If (smMnfCallType = "H") And (imVehGpSetNo <= 0) Then
    '    Select Case Trim$(tmMnf.sUnitType)
    '        Case "1"
    '            slName = slName & "/Participant"
    '        Case "2"
    '            slName = slName & "/Subtotal"
    '        Case "3"
    '            slName = slName & "/Market"
    '        Case "4"
    '            slName = slName & "/Format"
    '        Case "5"
    '            slName = slName & "/Research"
    '    End Select
    'End If
    'cbcSelect.AddItem slName
    'If (smMnfCallType = "H") Then
    '    slName = slName & "\" & LTrim$(Str$(tmMnf.iCode))
    'Else
    '    slName = tmMnf.sName + "\" + LTrim$(Str$(tmMnf.iCode))'slName + "\" + LTrim$(Str$(tmMnf.iCode))
    'End If
    ''lbcNameCode.AddItem slName
    'gAddItemToSortCode slName, tgNameCode(), True
    'cbcSelect.AddItem "[New]", 0
    '5/10/10:  This save required so that tgDemoMnf regardless if Done or Save pressed
    If smMnfCallType = "D" Then
        sgDemoMnfStamp = ""
        ilRet = gObtainMnfForType("D", sgDemoMnfStamp, tgDemoMnf())
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
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    If mTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
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
                    pbcMNm_Paint imPaintIndex
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
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/12/93       By:D. LeVine      *
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
    If ilBoxNo < imLBCtrls Or ilBoxNo > imNoCtrls Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            If smMnfCallType <> "F" Then
                gSetChgFlag tmMnf.sName, edcName, tmCtrls(ilBoxNo)
            Else
                gSetChgFlag tmMnf.sUnitType, edcName, tmCtrls(ilBoxNo)
            End If
        Case CTRL2INDEX 'Control 2
            If (smMnfCallType = "I") Or (smMnfCallType = "A") Then
                gPDNToStr tmMnf.sRPU, 2, slStr
                gSetChgFlag slStr, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "S" Then
                gPDNToStr tmMnf.sSSComm, 4, slStr
                gSetChgFlag slStr, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "C" Then
                gSetChgFlag tmMnf.sUnitType, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "X" Then
                gSetChgFlag tmMnf.sUnitType, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "V" Then
                gSetChgFlag Trim$(str$(tmMnf.iGroupNo)), Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "G" Then
                gSetChgFlag Trim$(str$(tmMnf.iGroupNo)), Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "R" Then 'Change set as user enters info
            End If
            If smMnfCallType = "E" Then
                gSetChgFlag Trim$(str$(tmMnf.iGroupNo)), Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "N" Then 'Change set as user enters info
            End If
            If smMnfCallType = "Y" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sUnitType, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "Z" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sUnitType, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "H" Then
                gSetChgFlag Trim$(str$(tmMnf.iGroupNo)), Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "B" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sUnitType, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "K" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sUnitType, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "P" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sRPU, Ctrl2, tmCtrls(ilBoxNo)
            End If
            '5/7/10: Assign group number automatically
            'If smMnfCallType = "D" Then
            '    gSetChgFlag Trim$(str$(tmMnf.iGroupNo)), Ctrl2, tmCtrls(ilBoxNo)
            'End If
            If smMnfCallType = "F" Then
                gSetChgFlag Trim$(tmMnf.sName), Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "O" Then
                gSetChgFlag tmMnf.sCodeStn, Ctrl2, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "3" Then
                gSetChgFlag Trim$(str$(tmMnf.iGroupNo)), Ctrl2, tmCtrls(ilBoxNo)
            End If
        Case CTRL3INDEX 'Control 3
            If smMnfCallType = "I" Then
                'gSetChgFlag tmMnf.sUnitType, Ctrl3, tmCtrls(ilBoxNo)
                gSetChgFlag tmMnf.sUnitsPer, Ctrl3, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "A" Then
                gSetChgFlag Trim$(str$(tmMnf.iGroupNo)), Ctrl3, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "S" Then
                slStr = ""
                If tmMnf.iGroupNo > 0 Then
                    slStr = lbcOrigin.List(tmMnf.iGroupNo - 1)
                End If
                gSetChgFlag slStr, Ctrl3, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "C" Then
                gSetChgFlag tmMnf.sCodeStn, Ctrl3, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "H" Then
                'gSetChgFlag tmMnf.sUnitType, Ctrl3, tmCtrls(ilBoxNo)
                Select Case Trim$(tmMnf.sUnitType)
                    Case "1"
                        slStr = "Participant"
                    Case "2"
                        slStr = "Subtotal"
                    Case "3"
                        slStr = "Market"
                    Case "4"
                        slStr = "Format"
                    Case "5"
                        slStr = "Research"
                    Case "6"
                        slStr = "Sub-Company"
                    Case Else
                        slStr = ""
                End Select
                If StrComp(slStr, Ctrl3.Text, 1) <> 0 Then
                    'If (StrComp(slStr, "Research", 1) <> 0) And (StrComp(slStr, "Market", 1) <> 0) Then
                        imDollars = -1
                        slStr = "         "
                        gSetShow pbcMNm(0), slStr, tmCtrls(CTRL4INDEX)
                    'End If
                End If
                gSetChgFlag slStr, Ctrl3, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "R" Then
                slStr = Trim$(str$(tmMnf.iGroupNo))
                gSetChgFlag slStr, Ctrl3, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "N" Then 'Change set as user enters info
            End If
            If smMnfCallType = "P" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sUnitType, Ctrl3, tmCtrls(ilBoxNo)
            End If
        Case CTRL4INDEX
            If smMnfCallType = "I" Then
                gPDNToStr tmMnf.sSSComm, 4, slStr
                gSetChgFlag slStr, Ctrl4, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "R" Then
                gSetChgFlag tmMnf.sCodeStn, Ctrl4, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "N" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sCodeStn, Ctrl4, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "P" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sSSComm, Ctrl4, tmCtrls(ilBoxNo)
            End If
        Case CTRL5INDEX
            If smMnfCallType = "I" Then 'Change set as user enters info
            End If
            If smMnfCallType = "S" Then 'Change set as user enters info
                gSetChgFlag tmMnf.sUnitsPer, Ctrl5, tmCtrls(ilBoxNo)
            End If
        Case CTRL6INDEX
            If smMnfCallType = "I" Then 'Change set as user enters info
            End If
        Case CTRL7INDEX
            If smMnfCallType = "I" Then 'Change set as user enters info
                slStr = gLongToStrDec(tmMnf.lCost, 2)
                gSetChgFlag slStr, Ctrl7, tmCtrls(ilBoxNo)
            End If
        Case CTRL8INDEX
            If smMnfCallType = "I" Then 'Change set as user enters info
            End If
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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

    If imBypassSetting Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    'Revert button set if any field changed
    If ilAltered Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If smMnfCallType <> "F" Then
        If (imSelectedIndex > 0) Or (mTestFields(TESTALLCTRLS, ALLBLANK + NOMSG) = NO) Then
            If imUpdateAllowed Then
                cmcErase.Enabled = True
            Else
                cmcErase.Enabled = False
            End If
        Else
            cmcErase.Enabled = False
        End If
    Else
        cmcErase.Enabled = False
    End If
    '9/12/16: Removed Merge button as no support code added to Merge.Frm
    'Merge set only if change mode
    'If (smMnfCallType <> "N") And (smMnfCallType <> "V") And (smMnfCallType <> "E") And (smMnfCallType <> "Y") And (smMnfCallType <> "H") And (smMnfCallType <> "B") And (smMnfCallType <> "K") And (smMnfCallType <> "P") And (smMnfCallType <> "D") And (smMnfCallType <> "F") Then
    '    If (imSelectedIndex > 0) And (tgUrf(0).sMerge = "I") And (imUpdateAllowed) Then
    '        cmcMerge.Enabled = True
    '    Else
    '        cmcMerge.Enabled = False
    '    End If
    'Else
    '    cmcMerge.Enabled = False
    'End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:Set Focus                       *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
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
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imNoCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.SetFocus
        Case CTRL2INDEX 'Second control
            If (smMnfCallType <> "R") And (smMnfCallType <> "M") And (smMnfCallType <> "N") Then
                Ctrl2.SetFocus
            Else    'Revenue sets
                pbcYN.SetFocus
            End If
        Case CTRL3INDEX 'Second control
            If smMnfCallType = "S" Then
                edcDropDown.SetFocus
            ElseIf smMnfCallType = "H" Then
                edcDropDown.SetFocus
            ElseIf smMnfCallType = "M" Then
                pbcYN.SetFocus
            ElseIf smMnfCallType = "N" Then
                pbcYN.SetFocus
            Else
                Ctrl3.SetFocus
            End If
        Case CTRL4INDEX 'Second control
            If smMnfCallType = "S" Then
                pbcYN.SetFocus
            ElseIf smMnfCallType = "H" Then
                pbcYN.SetFocus
            ElseIf smMnfCallType = "M" Then
                pbcYN.SetFocus
            Else
                Ctrl4.SetFocus
            End If
        Case CTRL5INDEX 'Second control
            If smMnfCallType = "I" Then
                pbcYN.SetFocus
            End If
            If smMnfCallType = "S" Then
                Ctrl5.SetFocus
            End If
        Case CTRL6INDEX 'Second control
            If smMnfCallType = "I" Then
                pbcYN.SetFocus
            End If
        Case CTRL7INDEX 'Second control
            If smMnfCallType = "I" Then
                Ctrl7.SetFocus
            End If
        Case CTRL8INDEX 'Second control
            If smMnfCallType = "I" Then
                pbcYN.SetFocus
            End If
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > imNoCtrls) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case NAMEINDEX 'Name
            edcName.Visible = False  'Set visibility
            slStr = edcName.Text
            gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
        Case CTRL2INDEX 'Control 2
            Ctrl2.Visible = False  'Set visibility
            If (smMnfCallType = "I") Or (smMnfCallType = "A") Then
                slStr = Ctrl2.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "S" Then
                slStr = Ctrl2.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 4, slStr
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "C" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "X" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "R" Then
                pbcYN.Visible = False  'Set visibility
                If imManOpt = 1 Then
                    slStr = "Mandatory"
                ElseIf imManOpt = 2 Then
                    slStr = "Optional"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "M" Then
                pbcYN.Visible = False  'Set visibility
                If imBillMGMissed = 1 Then
                    slStr = "Bill MG"
                ElseIf imBillMGMissed = 2 Then
                    slStr = "Bill Missed (No MG)"
                ElseIf imBillMGMissed = 3 Then
                    slStr = "Bill MG & Missed"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "V" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "G" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "E" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "N" Then
                pbcYN.Visible = False  'Set visibility
                If imTypeOfFeed = 0 Then
                    slStr = "Dish"
                ElseIf imTypeOfFeed = 1 Then
                    slStr = "Antenna"
                ElseIf imTypeOfFeed = 2 Then
                    slStr = "CD"
                ElseIf imTypeOfFeed = 3 Then
                    slStr = "Subfeed"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "Y" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "H" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "B" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "K" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "P" Then
                slStr = Ctrl2.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 0, slStr
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "D" Then
                '5/7/10: Assign group number automatically
                'slStr = Ctrl2.Text
                'gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
                'If Val(slStr) > imMaxCustomDemoNumber Then
                '    imMaxCustomDemoNumber = Val(slStr)
                'End If
            End If
            If smMnfCallType = "F" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "O" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "Z" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "L" Then
                pbcYN.Visible = False  'Set visibility
                If imEnglish = 0 Then
                    slStr = "Yes"
                ElseIf imEnglish = 1 Then
                    slStr = "No"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(21), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "3" Then
                slStr = Ctrl2.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
        Case CTRL3INDEX 'Control 3
            If smMnfCallType = "I" Then
                Ctrl3.Visible = False  'Set visibility
                slStr = Ctrl3.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "A" Then
                Ctrl3.Visible = False  'Set visibility
                slStr = Ctrl3.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "S" Then
                lbcOrigin.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                If lbcOrigin.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcOrigin.List(lbcOrigin.ListIndex)
                End If
                Ctrl3.Visible = False  'Set visibility
                slStr = Ctrl3.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "C" Then
                Ctrl3.Visible = False  'Set visibility
                slStr = Ctrl3.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "H" Then
                'Ctrl3.Visible = False  'Set visibility
                'slStr = Ctrl3.Text
                'gSetShow pbcMnm(0), slStr, tmCtrls(ilBoxNo)
                lbcOrigin.Visible = False
                edcDropDown.Visible = False
                cmcDropDown.Visible = False
                If lbcOrigin.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcOrigin.List(lbcOrigin.ListIndex)
                End If
                Ctrl3.Visible = False  'Set visibility
                slStr = Ctrl3.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "R" Then
                Ctrl3.Visible = False  'Set visibility
                slStr = Ctrl3.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "M" Then
                pbcYN.Visible = False  'Set visibility
                If imDefReason = 0 Then
                    slStr = "Yes"
                ElseIf imDefReason = 1 Then
                    slStr = "No"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "N" Then
                pbcYN.Visible = False  'Set visibility
                If imSubFeedAllowed = 0 Then
                    slStr = "Yes"
                ElseIf imSubFeedAllowed = 1 Then
                    slStr = "No"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "P" Then
                Ctrl3.Visible = False  'Set visibility
                slStr = Ctrl3.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 0, slStr
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "O" Then
                pbcYN.Visible = False  'Set visibility
                If imUsThem = 1 Then
                    slStr = "Us"
                ElseIf imUsThem = 2 Then
                    slStr = "Them"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
        Case CTRL4INDEX
            Ctrl4.Visible = False
            If smMnfCallType = "I" Then
                slStr = Ctrl4.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 4, slStr
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "R" Then
                slStr = Ctrl4.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "N" Then
                slStr = Ctrl4.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "P" Then
                slStr = Ctrl4.Text
                gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 0, slStr
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "S" Then
                pbcYN.Visible = False  'Set visibility
                If imUpdateRvf = 1 Then
                    slStr = "History"
                ElseIf imUpdateRvf = 2 Then
                    slStr = "Export+History"
                ElseIf imUpdateRvf = 3 Then
                    slStr = "Export+A/R"
                ElseIf imUpdateRvf = 4 Then
                    slStr = "Ask by Vehicle"
                ElseIf imUpdateRvf = 0 Then
                    slStr = "Receivables"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(3), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "H" Then
                pbcYN.Visible = False  'Set visibility
                If imDollars = 1 Then
                    slStr = "No"
                ElseIf imDollars = 0 Then
                    slStr = "Yes"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "M" Then
                pbcYN.Visible = False  'Set visibility
                If imMissedFor = 1 Then
                    slStr = "Network Missed"    '"Traffic"
                ElseIf imMissedFor = 2 Then
                    slStr = "Station Missed"    '"Affiliate Web"
                ElseIf imMissedFor = 3 Then
                    slStr = "Network & Station Missed"  '"Both"
                ElseIf imMissedFor = 4 Then
                    slStr = "Station Replacement"    'New selection
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
        Case CTRL5INDEX
            Ctrl5.Visible = False
            If smMnfCallType = "I" Then
                pbcYN.Visible = False  'Set visibility
                If imHardCost = 0 Then
                    slStr = "Yes"
                    If imAcqCostDefined Then
                        Ctrl7.Text = ""
                        tmCtrls(CTRL7INDEX).sShow = ""
                    End If
                    imTaxable = 1
                    tmCtrls(CTRL6INDEX).sShow = ""
                ElseIf imHardCost = 1 Then
                    slStr = "No"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
            If smMnfCallType = "S" Then
                slStr = Ctrl5.Text
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
        Case CTRL6INDEX
            Ctrl6.Visible = False
            If smMnfCallType = "I" Then
                pbcYN.Visible = False  'Set visibility
                If imTaxDefined Then
                    If imTaxable = 0 Then
                        slStr = "Yes"
                    ElseIf imTaxable = 1 Then
                        slStr = "No"
                    Else
                        slStr = ""
                    End If
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
        Case CTRL7INDEX
            Ctrl7.Visible = False
            If smMnfCallType = "I" Then
                If imAcqCostDefined Then
                    slStr = Ctrl7.Text
                    If gStrDecToLong(slStr, 2) <> 0 Then
                        imHardCost = 1
                        tmCtrls(CTRL5INDEX).sShow = "No"
                    End If
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
        Case CTRL8INDEX
            Ctrl8.Visible = False
            If smMnfCallType = "I" Then
                pbcYN.Visible = False  'Set visibility
                If imSaleType = 0 Then
                    slStr = "NTR"
                ElseIf imSaleType = 1 Then
                    slStr = "Agency"
                ElseIf imSaleType = 2 Then
                    slStr = "Direct"
                Else
                    slStr = ""
                End If
                gSetShow pbcMNm(0), slStr, tmCtrls(ilBoxNo)
            End If
    End Select
End Sub
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
'   FacTerminate
'   Where:
'
    Dim ilRet As Integer

    sgNameCodeTag = ""
    
    If smMnfCallType = "R" Then
        sgRevSetStamp = ""
        ilRet = gObtainMnfForType("R", sgRevSetStamp, tgRevSet())
    End If
    If smMnfCallType = "O" Then
        sgShareBudgetStamp = ""
        ilRet = gObtainMnfForType("O", sgShareBudgetStamp, tgShareBudget())
    End If
    If smMnfCallType = "X" Then
        sgExclMnfStamp = ""
        ilRet = gObtainMnfForType("X", sgExclMnfStamp, tgExclMnf())
    End If
    If smMnfCallType = "B" Then
        sgBusCatMnfStamp = ""
        ilRet = gObtainMnfForType("B", sgBusCatMnfStamp, tgBusCatMnf())
    End If
    If smMnfCallType = "P" Then
        sgPotMnfStamp = ""
        ilRet = gObtainMnfForType("P", sgPotMnfStamp, tgPotMnf())
    End If
    If smMnfCallType = "F" Then
        sgSocEcoMnfStamp = ""
        ilRet = gObtainMnfForType("F", sgSocEcoMnfStamp, tgSocEcoMnf())
    End If


    sgDoneMsg = Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload MultiNm
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test mandatory and blank fields *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
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
    Dim slMess As String    'Message string
    Dim slStr As String
    Dim ilReq As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    If (ilCtrlNo = NAMEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If smMnfCallType = "M" Then 'Missed reason
            slMess = "Reason must be specified"
        ElseIf smMnfCallType = "Y" Then 'Missed reason
            slMess = "Description must be specified"
        Else
            slMess = "Name must be specified"
        End If
        If gFieldDefinedCtrl(edcName, "", slMess, tmCtrls(NAMEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = NAMEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If TypeOf Ctrl2 Is Image Then 'Not used
    Else
        If (ilCtrlNo = CTRL2INDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If (smMnfCallType = "I") Or (smMnfCallType = "R") Then 'Lock Box
                slMess = "Rate must be specified"
            ElseIf smMnfCallType = "S" Then
                slMess = "% Commission must be specified"
            ElseIf smMnfCallType = "C" Then
                slMess = "Abbreviation must be specified"
            ElseIf smMnfCallType = "X" Then
                slMess = "Abbreviation must be specified"
            ElseIf smMnfCallType = "R" Then
                slMess = "Mandatory/Optional must be specified"
            ElseIf smMnfCallType = "M" Then
                slMess = "Bill MG/Missed must be specified"
            ElseIf smMnfCallType = "V" Then
                slMess = "Sort Order # must be specified"
            ElseIf smMnfCallType = "G" Then
                slMess = "Sort Order # must be specified"
            ElseIf smMnfCallType = "E" Then
                slMess = "Group # must be specified"
            ElseIf smMnfCallType = "N" Then
                slMess = "Type of Feed must be specified"
            ElseIf smMnfCallType = "Y" Then
                slMess = "Transaction Type must be specified"
            ElseIf smMnfCallType = "H" Then
                slMess = "Sort Order # must be specified"
            ElseIf smMnfCallType = "B" Then
                slMess = "Type must be specified"
            ElseIf smMnfCallType = "K" Then
                slMess = "Abbreviation must be specified"
            ElseIf smMnfCallType = "P" Then
                slMess = "Optimistic % must be specified"
            '5/7/10: Assign group number automatically
            'ElseIf smMnfCallType = "D" Then
            '    slMess = "Sort Order # must be specified"
            ElseIf smMnfCallType = "F" Then
                slMess = "Description must be specified"
            ElseIf smMnfCallType = "O" Then
                slMess = "Abbreviation must be specified"
            ElseIf smMnfCallType = "L" Then
                slMess = "English must be specified"
            ElseIf smMnfCallType = "Z" Then
                slMess = "Abbreviation must be specified"
            ElseIf smMnfCallType = "3" Then
                slMess = "Sort Order # must be specified"
            End If
            If (smMnfCallType <> "R") And (smMnfCallType <> "M") And (smMnfCallType <> "N") And (smMnfCallType <> "L") Then
                If gFieldDefinedCtrl(Ctrl2, "", slMess, tmCtrls(CTRL2INDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL2INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
                If smMnfCallType = "Y" Then
                    If Len(Trim$(Ctrl2.Text)) <> 2 Then
                        slStr = ""
                        If gFieldDefinedStr(slStr, "", "Transaction Type must be 2 Characters", tmCtrls(CTRL2INDEX).iReq, ilState) = NO Then
                            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                                imBoxNo = CTRL2INDEX
                            End If
                            mTestFields = NO
                            Exit Function
                        End If
                    End If
                End If
                '5/7/10: Assign group number automatically
                'If smMnfCallType = "D" Then
                '    If (Val(Trim$(Ctrl2.Text)) <= 0) Or (Val(Trim$(Ctrl2.Text)) > 990) Then
                '        slStr = ""
                '        If gFieldDefinedStr(slStr, "", "Sort Number must be a unique # from 1 thru 990", tmCtrls(CTRL2INDEX).iReq, ilState) = NO Then
                '            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                '                imBoxNo = CTRL2INDEX
                '            End If
                '            mTestFields = NO
                '            Exit Function
                '        End If
                '    End If
                'End If
            Else
                If smMnfCallType = "R" Then
                    If imManOpt = 1 Then
                        slStr = "Mandatory"
                    ElseIf imManOpt = 2 Then
                        slStr = "Optional"
                    Else
                        slStr = ""
                    End If
                    If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL2INDEX).iReq, ilState) = NO Then
                        If ilState = (ALLMANDEFINED + SHOWMSG) Then
                            imBoxNo = CTRL2INDEX
                        End If
                        mTestFields = NO
                        Exit Function
                    End If
                ElseIf smMnfCallType = "M" Then
                    If imBillMGMissed = 1 Then
                        slStr = "Bill MG"
                    ElseIf imBillMGMissed = 2 Then
                        slStr = "Bill Missed (No MG)"
                    ElseIf imBillMGMissed = 3 Then
                        slStr = "Bill MG & Missed"
                    Else
                        slStr = ""
                    End If
                    If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL2INDEX).iReq, ilState) = NO Then
                        If ilState = (ALLMANDEFINED + SHOWMSG) Then
                            imBoxNo = CTRL2INDEX
                        End If
                        mTestFields = NO
                        Exit Function
                    End If
                ElseIf smMnfCallType = "L" Then
                    If imEnglish = 0 Then
                        slStr = "Yes"
                    ElseIf imEnglish = 1 Then
                        slStr = "No"
                    Else
                        slStr = ""
                    End If
                    If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL2INDEX).iReq, ilState) = NO Then
                        If ilState = (ALLMANDEFINED + SHOWMSG) Then
                            imBoxNo = CTRL2INDEX
                        End If
                        mTestFields = NO
                        Exit Function
                    End If
                Else
                    If imTypeOfFeed = 0 Then
                        slStr = "D"
                    ElseIf imTypeOfFeed = 1 Then
                        slStr = "A"
                    ElseIf imTypeOfFeed = 2 Then
                        slStr = "C"
                    ElseIf imTypeOfFeed = 3 Then
                        slStr = "S"
                    Else
                        slStr = ""
                    End If
                    If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL2INDEX).iReq, ilState) = NO Then
                        If ilState = (ALLMANDEFINED + SHOWMSG) Then
                            imBoxNo = CTRL2INDEX
                        End If
                        mTestFields = NO
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    If TypeOf Ctrl3 Is Image Then 'Not used
    Else
        If (ilCtrlNo = CTRL3INDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            ilReq = tmCtrls(CTRL3INDEX).iReq
            If smMnfCallType = "I" Then
                slMess = "Unit per must be specified"
                If StrComp(Trim$(tmMnf.sName), "MultiMedia", vbTextCompare) = 0 Then
                    ilReq = False
                End If
            ElseIf smMnfCallType = "A" Then
                slMess = "Group Number must be specified"
            ElseIf smMnfCallType = "S" Then
                slMess = "Sales Origin must be specified"
            ElseIf smMnfCallType = "C" Then
                slMess = "Station Product Protection Code must be specified"
            ElseIf smMnfCallType = "H" Then
                slMess = "Group Name must be specified"
            ElseIf smMnfCallType = "R" Then
                slMess = "Set # must be specified"
            ElseIf smMnfCallType = "M" Then
                slMess = "Default Reason must be specified"
            ElseIf smMnfCallType = "N" Then
                slMess = "Subfeed Allowed must be specified"
            ElseIf smMnfCallType = "P" Then
                slMess = "Most Likely % must be specified"
            ElseIf smMnfCallType = "O" Then
                slMess = "Us/Them must be specified"
            End If
            If (smMnfCallType <> "M") And (smMnfCallType <> "N") And (smMnfCallType <> "O") Then
                If gFieldDefinedCtrl(Ctrl3, "", slMess, ilReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL3INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            ElseIf (smMnfCallType = "O") Then
                If imUsThem = 1 Then
                    slStr = "Us"
                ElseIf imUsThem = 2 Then
                    slStr = "Them"
                Else
                    slStr = ""
                End If
                If gFieldDefinedStr(slStr, "", slMess, ilReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL3INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            ElseIf (smMnfCallType = "M") Then
                If imDefReason = 0 Then
                    slStr = "Yes"
                ElseIf imDefReason = 1 Then
                    slStr = "No"
                Else
                    slStr = ""
                End If
                If gFieldDefinedStr(slStr, "", slMess, ilReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL3INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            Else
                If imSubFeedAllowed = 0 Then
                    slStr = "Y"
                ElseIf imSubFeedAllowed = 1 Then
                    slStr = "N"
                Else
                    slStr = ""
                End If
                If gFieldDefinedStr(slStr, "", slMess, ilReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL3INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
        End If
    End If
    If TypeOf Ctrl4 Is Image Then 'Not used
    Else
        If (ilCtrlNo = CTRL4INDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If smMnfCallType = "I" Then
                slMess = "% Commission must be specified"
            End If
            If smMnfCallType = "N" Then
                slMess = "Station Code must be specified"
            End If
            If smMnfCallType = "R" Then
                slMess = "Station Revenue Set Code must be specified"
            End If
            If smMnfCallType = "P" Then
                slMess = "Pessimistic % must be specified"
            End If
            If smMnfCallType = "S" Then
                slMess = "Update Receivables must be specified"
            End If
            If smMnfCallType = "M" Then
                slMess = "Missed For must be specified"
            End If
            If (smMnfCallType = "H") Then
                If lbcOrigin.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcOrigin.List(lbcOrigin.ListIndex)
                End If
                If (StrComp(slStr, "Research", 1) = 0) Then
                    slMess = "Dollars must be specified"
                Else
                    slMess = "Cluster must be specified"
                End If
            End If
            If (smMnfCallType = "I") Or (smMnfCallType = "N") Or (smMnfCallType = "R") Or (smMnfCallType = "P") Then
                If gFieldDefinedCtrl(Ctrl4, "", slMess, tmCtrls(CTRL4INDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL4INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            ElseIf (smMnfCallType = "S") Then
                If imUpdateRvf = 0 Then
                    slStr = "Y"
                ElseIf imUpdateRvf = 1 Then
                    slStr = "N"
                ElseIf imUpdateRvf = 2 Then
                    slStr = "E"
                ElseIf imUpdateRvf = 3 Then
                    slStr = "F"
                ElseIf imUpdateRvf = 4 Then
                    slStr = "A"
                Else
                    slStr = ""
                End If
                If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL4INDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL4INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            ElseIf (smMnfCallType = "H") Then
                If lbcOrigin.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcOrigin.List(lbcOrigin.ListIndex)
                End If
                If (StrComp(slStr, "Research", 1) = 0) Or (StrComp(slStr, "Market", 1) = 0) Then
                    If imDollars = 0 Then
                        slStr = "Yes"
                    ElseIf imDollars = 1 Then
                        slStr = "No"
                    Else
                        slStr = ""
                    End If
                    If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL4INDEX).iReq, ilState) = NO Then
                        If ilState = (ALLMANDEFINED + SHOWMSG) Then
                            imBoxNo = CTRL4INDEX
                        End If
                        mTestFields = NO
                        Exit Function
                    End If
                End If
            ElseIf smMnfCallType = "M" Then
                If imMissedFor = 1 Then
                    slStr = "Network Missed" '"Traffic"
                ElseIf imMissedFor = 2 Then
                    slStr = "Station Missed"    '"Affiliate Web"
                ElseIf imMissedFor = 3 Then
                    slStr = "Network & Station Missed"  '"
                ElseIf imMissedFor = 4 Then
                    slStr = "Station Replacement"    'New
                Else
                    slStr = ""
                End If
                If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL2INDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL4INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
        End If
    End If
    If (smMnfCallType = "I") Then
        If (ilCtrlNo = CTRL5INDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If smMnfCallType = "I" Then
                slMess = "Hard Cost must be specified"
                If imHardCost = 0 Then
                    slStr = "Yes"
                ElseIf imHardCost = 1 Then
                    slStr = "No"
                Else
                    slStr = ""
                End If
            End If
            If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL5INDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = CTRL5INDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (smMnfCallType = "S") Then
        If (ilCtrlNo = CTRL5INDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            slMess = "Logo must be specified"
            If gFieldDefinedCtrl(Ctrl5, "", slMess, tmCtrls(CTRL5INDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = CTRL5INDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (smMnfCallType = "I") And (imTaxDefined) Then
        If (ilCtrlNo = CTRL6INDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If smMnfCallType = "I" Then
                slMess = "Taxable must be specified"
                If imTaxable = 0 Then
                    slStr = "Yes"
                ElseIf imTaxable = 1 Then
                    slStr = "No"
                Else
                    slStr = ""
                End If
            End If
            If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL6INDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = CTRL6INDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (smMnfCallType = "I") And (imAcqCostDefined) Then
        If imHardCost <> 0 Then
            If (ilCtrlNo = CTRL7INDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
                If smMnfCallType = "I" Then
                    slMess = "NTR Acquisition Cost must be specified"
                    slStr = Ctrl7.Text
                End If
                If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL7INDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = CTRL7INDEX
                    End If
                    mTestFields = NO
                    Exit Function
                End If
            End If
        End If
    End If
    If (smMnfCallType = "I") Then
        If (ilCtrlNo = CTRL8INDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
            If smMnfCallType = "I" Then
                slMess = "Sale Type must be specified"
                If imSaleType = 0 Then
                    slStr = "NTR"
                ElseIf imSaleType = 1 Then
                    slStr = "Agency"
                ElseIf imSaleType = 2 Then
                    slStr = "Direct"
                Else
                    slStr = ""
                End If
            End If
            If gFieldDefinedStr(slStr, "", slMess, tmCtrls(CTRL8INDEX).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = CTRL8INDEX
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    End If
    If (smMnfCallType = "M") And (imDefReason = 0) And (ilState = ALLMANDEFINED + SHOWMSG) Then 'Missed reason
        'Test that no other reason has default set
        'Missed Reason
        slStr = ""
        'ReDim tgMRMnf(1 To 1) As MNF
        ReDim tgMRMnf(0 To 0) As MNF
        ilRet = gObtainMnfForType("M", slStr, tgMRMnf())
        For ilLoop = LBound(tgMRMnf) To UBound(tgMRMnf) - 1 Step 1
            If Trim$(tgMRMnf(ilLoop).sUnitType) = "Y" Then
                If imSelectedIndex <> 0 Then
                    slNameCode = tgNameCode(imSelectedIndex - 1).sKey    'lbcNameCode.List(ilSelectIndex - 1)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) <> tgMRMnf(ilLoop).iCode Then
                        MsgBox "Default Reason already defined for " & Trim$(tgMRMnf(ilLoop).sName), vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                        mTestFields = NO
                        Exit Function
                    End If
                Else
                    MsgBox "Default Reason already defined for " & Trim$(tgMRMnf(ilLoop).sName), vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    mTestFields = NO
                    Exit Function
                End If
            End If
        Next ilLoop
    End If
    mTestFields = YES
End Function
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
Private Sub pbcMNm_MouseUp(iIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim slStr As String
    If imBoxNo = NAMEINDEX Then
        If Not mOKName(True) Then
            Exit Sub
        End If
    End If
    For ilBox = imLBCtrls To imNoCtrls Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                If (smMnfCallType = "H") And (ilBox = CTRL4INDEX) Then
                    If lbcOrigin.ListIndex < 0 Then
                        slStr = ""
                    Else
                        slStr = lbcOrigin.List(lbcOrigin.ListIndex)
                    End If
                    If (StrComp(slStr, "Research", 1) <> 0) And (StrComp(slStr, "Market", 1) <> 0) Then
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                End If
                '1/26/11: Disallow Vehicle group name type to be altered in change mode
                If (smMnfCallType = "H") And (ilBox = CTRL3INDEX) Then
                    If imSelectedIndex > 0 Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                End If
                If (smMnfCallType = "I") Then
                    If (ilBox = CTRL6INDEX) And ((Not imTaxDefined) Or (imHardCost = 0)) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If (ilBox = CTRL7INDEX) And (Not imAcqCostDefined) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    If (ilBox = NAMEINDEX) And (StrComp(Trim$(tmMnf.sName), "MultiMedia", vbTextCompare) = 0) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                End If
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcMNm_Paint(ilIndex As Integer)
    Dim ilBox As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    If smMnfCallType = "H" Then
        pbcMNm(ilIndex).Cls
        llColor = pbcMNm(ilIndex).ForeColor
        slFontName = pbcMNm(ilIndex).FontName
        flFontSize = pbcMNm(ilIndex).FontSize
        pbcMNm(ilIndex).ForeColor = BLUE
        pbcMNm(ilIndex).FontBold = False
        pbcMNm(ilIndex).FontSize = 7
        pbcMNm(ilIndex).FontName = "Arial"
        pbcMNm(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcMNm(ilIndex).CurrentX = tmCtrls(CTRL4INDEX).fBoxX + 15  'fgBoxInsetX
        pbcMNm(ilIndex).CurrentY = tmCtrls(CTRL4INDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If lbcOrigin.ListIndex >= 0 Then
            slStr = lbcOrigin.List(lbcOrigin.ListIndex)
        End If
        If StrComp(slStr, "Research", 1) = 0 Then
            slStr = "Dollars"
        ElseIf StrComp(slStr, "Market", 1) = 0 Then
            slStr = "Cluster"
        Else
            slStr = "                 "
        End If
        pbcMNm(ilIndex).Print slStr
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).FontName = slFontName
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).ForeColor = llColor
        pbcMNm(ilIndex).FontBold = True
    End If
    If smMnfCallType = "I" Then
        pbcMNm(ilIndex).Cls
        llColor = pbcMNm(ilIndex).ForeColor
        slFontName = pbcMNm(ilIndex).FontName
        flFontSize = pbcMNm(ilIndex).FontSize
        pbcMNm(ilIndex).ForeColor = BLUE
        pbcMNm(ilIndex).FontBold = False
        pbcMNm(ilIndex).FontSize = 7
        pbcMNm(ilIndex).FontName = "Arial"
        pbcMNm(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcMNm(ilIndex).CurrentX = tmCtrls(CTRL7INDEX).fBoxX + 15  'fgBoxInsetX
        pbcMNm(ilIndex).CurrentY = tmCtrls(CTRL7INDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If imAcqCostDefined Then
            slStr = "Acquisition $"
            pbcMNm(ilIndex).Print slStr
        End If
        pbcMNm(ilIndex).CurrentX = tmCtrls(CTRL6INDEX).fBoxX + 15  'fgBoxInsetX
        pbcMNm(ilIndex).CurrentY = tmCtrls(CTRL6INDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        If imTaxDefined Then
            slStr = "Taxable"
            pbcMNm(ilIndex).Print slStr
        End If
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).FontName = slFontName
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).ForeColor = llColor
        pbcMNm(ilIndex).FontBold = True
    End If
    '9/6/14: Change title
    If smMnfCallType = "Z" Then
        pbcMNm(ilIndex).Cls
        llColor = pbcMNm(ilIndex).ForeColor
        slFontName = pbcMNm(ilIndex).FontName
        flFontSize = pbcMNm(ilIndex).FontSize
        pbcMNm(ilIndex).ForeColor = BLUE
        pbcMNm(ilIndex).FontBold = False
        pbcMNm(ilIndex).FontSize = 7
        pbcMNm(ilIndex).FontName = "Arial"
        pbcMNm(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcMNm(ilIndex).CurrentX = tmCtrls(NAMEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcMNm(ilIndex).CurrentY = tmCtrls(NAMEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcMNm(ilIndex).Print "Name"
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).FontName = slFontName
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).ForeColor = llColor
        pbcMNm(ilIndex).FontBold = True
    End If
    If smMnfCallType = "S" Then
        pbcMNm(ilIndex).Cls
        llColor = pbcMNm(ilIndex).ForeColor
        slFontName = pbcMNm(ilIndex).FontName
        flFontSize = pbcMNm(ilIndex).FontSize
        pbcMNm(ilIndex).ForeColor = BLUE
        pbcMNm(ilIndex).FontBold = False
        pbcMNm(ilIndex).FontSize = 7
        pbcMNm(ilIndex).FontName = "Arial"
        pbcMNm(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcMNm(ilIndex).CurrentX = tmCtrls(CTRL5INDEX).fBoxX + 15  'fgBoxInsetX
        pbcMNm(ilIndex).CurrentY = tmCtrls(CTRL5INDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcMNm(ilIndex).Print "Logo"
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).FontName = slFontName
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).ForeColor = llColor
        pbcMNm(ilIndex).FontBold = True
    End If
    If smMnfCallType = "6" Then
        pbcMNm(ilIndex).Cls
        llColor = pbcMNm(ilIndex).ForeColor
        slFontName = pbcMNm(ilIndex).FontName
        flFontSize = pbcMNm(ilIndex).FontSize
        pbcMNm(ilIndex).ForeColor = BLUE
        pbcMNm(ilIndex).FontBold = False
        pbcMNm(ilIndex).FontSize = 7
        pbcMNm(ilIndex).FontName = "Arial"
        pbcMNm(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        pbcMNm(ilIndex).CurrentX = tmCtrls(NAMEINDEX).fBoxX + 15  'fgBoxInsetX
        pbcMNm(ilIndex).CurrentY = tmCtrls(NAMEINDEX).fBoxY - 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
        pbcMNm(ilIndex).Print "Position"
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).FontName = slFontName
        pbcMNm(ilIndex).FontSize = flFontSize
        pbcMNm(ilIndex).ForeColor = llColor
        pbcMNm(ilIndex).FontBold = True
    End If
    For ilBox = imLBCtrls To imNoCtrls Step 1
        pbcMNm(ilIndex).CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcMNm(ilIndex).CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcMNm(ilIndex).Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName(True) Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imNoCtrls) Then
        If (imBoxNo <> NAMEINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    imTabDirection = -1  'Set-right to left
    Select Case imBoxNo
        Case -1
            imTabDirection = 0  'Set-Left to right
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then '[New]
                ilBox = 1
                mSetCommands
            Else
                mSetChg 1
                ilBox = 2
                If imSelectedIndex < 0 Then
                    'Check if only one field
                    'If (smMnfCallType = "T") Or (smMnfCallType = "W") Or (smMnfCallType = "J") Then   'Or (smMnfCallType = "R") Or (smMnfCallType = "M") Or (smMnfCallType = "N") Then
                    If imNoCtrls = 1 Then
                        mSetShow imBoxNo
                        imBoxNo = -1
                        If cmcUpdate.Enabled Then
                            cmcUpdate.SetFocus
                        Else
                            cmcDone.SetFocus
                        End If
                        Exit Sub
                    End If
                Else
                    'If (smMnfCallType = "T") Or (smMnfCallType = "W") Or (smMnfCallType = "J") Then   'Or (smMnfCallType = "R") Or (smMnfCallType = "M") Then
                    If imNoCtrls = 1 Then
                        ilBox = 1
                    End If
                End If
            End If
        Case 1 'Name
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
        Case Else
            ilBox = imBoxNo - 1
            If (smMnfCallType = "H") And (imBoxNo = CTRL4INDEX) Then
                '1/26/11: Bypass Vehicle Group type in change mode
                If imSelectedIndex > 0 Then
                    ilBox = CTRL2INDEX
                End If
            End If
            If smMnfCallType = "I" Then
                If imBoxNo = CTRL8INDEX Then
                    If (imAcqCostDefined) Then
                        ilBox = CTRL7INDEX
                    Else
                        ilBox = CTRL4INDEX
                    End If
                End If
                If imBoxNo = CTRL5INDEX Then
                    ilBox = CTRL8INDEX
                End If
                If imBoxNo = CTRL7INDEX Then
                    ilBox = CTRL4INDEX
                End If
                If imBoxNo = CTRL2INDEX Then
                    If StrComp(Trim$(tmMnf.sName), "MultiMedia", vbTextCompare) = 0 Then
                        mSetShow imBoxNo
                        imBoxNo = -1
                        If cbcSelect.Enabled Then
                            cbcSelect.SetFocus
                        ElseIf cmcCancel.Enabled Then
                            cmcCancel.SetFocus
                        End If
                        Exit Sub
                    End If
                End If
            End If
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imBoxNo = NAMEINDEX Then
        If Not mOKName(True) Then
            Exit Sub
        End If
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= imNoCtrls) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    Select Case imBoxNo
        Case -1
            imTabDirection = -1  'Set-Right to left
            If (smMnfCallType = "H") Then
                ilBox = CTRL2INDEX
            ElseIf (smMnfCallType = "I") Then
                If (imTaxDefined) And (imHardCost = 1) Then
                    ilBox = CTRL6INDEX
                Else
                    ilBox = CTRL5INDEX
                End If
            Else
                ilBox = imNoCtrls
            End If
        Case imNoCtrls
            If smMnfCallType = "I" Then 'Coming from CTRL8INDEX
                ilBox = CTRL5INDEX
            Else
                mSetShow imBoxNo
                imBoxNo = -1
                If (cmcUpdate.Enabled) And (igMNmCallSource = CALLNONE) Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
            End If
        Case Else
            If (smMnfCallType = "H") And (imBoxNo = CTRL3INDEX) Then
                If lbcOrigin.ListIndex < 0 Then
                    slStr = ""
                Else
                    slStr = lbcOrigin.List(lbcOrigin.ListIndex)
                End If
                If (StrComp(slStr, "Research", 1) <> 0) And (StrComp(slStr, "Market", 1) <> 0) Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    If (cmcUpdate.Enabled) And (igMNmCallSource = CALLNONE) Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
            End If
            ilBox = imBoxNo + 1
            If (smMnfCallType = "H") And (imBoxNo = CTRL2INDEX) Then
                '1/26/11: Bypass Vehicle Group type in change mode
                If imSelectedIndex > 0 Then
                    If lbcOrigin.ListIndex < 0 Then
                        slStr = ""
                    Else
                        slStr = lbcOrigin.List(lbcOrigin.ListIndex)
                    End If
                    If (StrComp(slStr, "Research", 1) <> 0) And (StrComp(slStr, "Market", 1) <> 0) Then
                        mSetShow imBoxNo
                        imBoxNo = -1
                        If (cmcUpdate.Enabled) And (igMNmCallSource = CALLNONE) Then
                            cmcUpdate.SetFocus
                        Else
                            cmcDone.SetFocus
                        End If
                        Exit Sub
                    End If
                    ilBox = CTRL4INDEX
                End If
            End If
            If smMnfCallType = "I" Then
                If imBoxNo = CTRL4INDEX Then
                    If (Not imAcqCostDefined) Then
                        ilBox = CTRL8INDEX
                    Else
                        ilBox = CTRL7INDEX
                    End If
                End If
                If imBoxNo = CTRL8INDEX Then
                    ilBox = CTRL5INDEX
                End If
                If imBoxNo = CTRL5INDEX Then
                    If (Not imTaxDefined) Or (imHardCost = 0) Then
                        mSetShow imBoxNo
                        imBoxNo = -1
                        If (cmcUpdate.Enabled) And (igMNmCallSource = CALLNONE) Then
                            cmcUpdate.SetFocus
                        Else
                            cmcDone.SetFocus
                        End If
                        Exit Sub
                    End If
                End If
                If imBoxNo = CTRL6INDEX Then
                    mSetShow imBoxNo
                    imBoxNo = -1
                    If (cmcUpdate.Enabled) And (igMNmCallSource = CALLNONE) Then
                        cmcUpdate.SetFocus
                    Else
                        cmcDone.SetFocus
                    End If
                    Exit Sub
                End If
            End If
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If (imBoxNo = CTRL5INDEX) And (smMnfCallType = "I") Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imHardCost <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imHardCost = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imHardCost <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imHardCost = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imHardCost = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imHardCost = 1
                pbcYN_Paint
            ElseIf imHardCost = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imHardCost = 0
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imHardCost = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL6INDEX) And (smMnfCallType = "I") Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imTaxable <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTaxable = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imTaxable <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTaxable = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imTaxable = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imTaxable = 1
                pbcYN_Paint
            ElseIf imTaxable = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imTaxable = 0
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imTaxable = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
     If (imBoxNo = CTRL8INDEX) And (smMnfCallType = "I") Then
        If KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imSaleType <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSaleType = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
            If imSaleType <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSaleType = 1
            pbcYN_Paint
        ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
            If imSaleType <> 2 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSaleType = 2
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSaleType = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSaleType = 1
                pbcYN_Paint
            ElseIf imSaleType = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSaleType = 2
                pbcYN_Paint
            ElseIf imSaleType = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imSaleType = 0
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imSaleType = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
   If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "S") Then
        If KeyAscii = Asc("R") Or (KeyAscii = Asc("r")) Then
            If imUpdateRvf <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imUpdateRvf = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("H") Or (KeyAscii = Asc("h")) Then
            If imUpdateRvf <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imUpdateRvf = 1
            pbcYN_Paint
        ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
            If imUpdateRvf <> 2 Then
                tmCtrls(imBoxNo).iChg = True
                imUpdateRvf = 2
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imUpdateRvf = 3
                pbcYN_Paint
            End If
        ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
            If imUpdateRvf <> 4 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imUpdateRvf = 4
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imUpdateRvf = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imUpdateRvf = 1
                pbcYN_Paint
            ElseIf imUpdateRvf = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imUpdateRvf = 2
                pbcYN_Paint
            ElseIf imUpdateRvf = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imUpdateRvf = 3
                pbcYN_Paint
            ElseIf imUpdateRvf = 3 Then
                tmCtrls(imBoxNo).iChg = True
                imUpdateRvf = 4
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imUpdateRvf = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "H") Then
        If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
            If imDollars <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imDollars = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imDollars <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imDollars = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imDollars = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imDollars = 1
                pbcYN_Paint
            ElseIf imDollars = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imDollars = 0
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imDollars = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "R") Then
        If (KeyAscii = Asc("M")) Or (KeyAscii = Asc("m")) Then
            If imManOpt <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imManOpt = 1
            pbcYN_Paint
        ElseIf (KeyAscii = Asc("O")) Or (KeyAscii = Asc("o")) Then
            If imManOpt <> 2 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imManOpt = 2
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imManOpt = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imManOpt = 2
                pbcYN_Paint
            ElseIf imManOpt = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imManOpt = 1
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imManOpt = 1
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "M") Then
        If KeyAscii = Asc(" ") Then
            If imBillMGMissed = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imBillMGMissed = 2
                pbcYN_Paint
            ElseIf imBillMGMissed = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imBillMGMissed = 1  '3
                pbcYN_Paint
            'ElseIf imBillMGMissed = 3 Then
            '    tmCtrls(imBoxNo).iChg = True
            '    imBillMGMissed = 1
            '    pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imBillMGMissed = 1
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "L") Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            If imEnglish <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imEnglish = 0
            pbcYN_Paint
        ElseIf (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
            If imEnglish <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imEnglish = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imEnglish = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imEnglish = 1
                pbcYN_Paint
            ElseIf imEnglish = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imEnglish = 0
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imEnglish = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "M") Then
        If KeyAscii = Asc(" ") Then
            If imDefReason = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imDefReason = 1
                pbcYN_Paint
            ElseIf imDefReason = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imDefReason = 0
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imDefReason = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "M") Then
        If KeyAscii = Asc(" ") Then
            If imMissedFor = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imMissedFor = 2
                pbcYN_Paint
            ElseIf imMissedFor = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imMissedFor = 3
                pbcYN_Paint
            ElseIf imMissedFor = 3 Then
                tmCtrls(imBoxNo).iChg = True
                imMissedFor = 4
                pbcYN_Paint
            ElseIf imMissedFor = 4 Then
                tmCtrls(imBoxNo).iChg = True
                imMissedFor = 1
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imMissedFor = 1
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "N") Then
        If (KeyAscii = Asc("D")) Or (KeyAscii = Asc("d")) Then
            If imTypeOfFeed <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTypeOfFeed = 0
            pbcYN_Paint
        ElseIf (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
            If imTypeOfFeed <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTypeOfFeed = 1
            pbcYN_Paint
        ElseIf (KeyAscii = Asc("C")) Or (KeyAscii = Asc("c")) Then
            If imTypeOfFeed <> 2 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTypeOfFeed = 2
            pbcYN_Paint
        ElseIf (KeyAscii = Asc("S")) Or (KeyAscii = Asc("s")) Then
            If imTypeOfFeed <> 3 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imTypeOfFeed = 3
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imTypeOfFeed = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imTypeOfFeed = 1
                pbcYN_Paint
            ElseIf imTypeOfFeed = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imTypeOfFeed = 2
                pbcYN_Paint
            ElseIf imTypeOfFeed = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imTypeOfFeed = 3
                pbcYN_Paint
            ElseIf imTypeOfFeed = 3 Then
                tmCtrls(imBoxNo).iChg = True
                imTypeOfFeed = 0
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imTypeOfFeed = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "N") Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            If imSubFeedAllowed <> 0 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSubFeedAllowed = 0
            pbcYN_Paint
        ElseIf (KeyAscii = Asc("N")) Or (KeyAscii = Asc("n")) Then
            If imSubFeedAllowed <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imSubFeedAllowed = 1
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSubFeedAllowed = 0 Then
                tmCtrls(imBoxNo).iChg = True
                imSubFeedAllowed = 1
                pbcYN_Paint
            ElseIf imSubFeedAllowed = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imSubFeedAllowed = 0
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imSubFeedAllowed = 0
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "O") Then
        If (KeyAscii = Asc("U")) Or (KeyAscii = Asc("u")) Then
            If imUsThem <> 1 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imUsThem = 1
            pbcYN_Paint
        ElseIf (KeyAscii = Asc("T")) Or (KeyAscii = Asc("t")) Then
            If imUsThem <> 2 Then
                tmCtrls(imBoxNo).iChg = True
            End If
            imUsThem = 2
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imUsThem = 1 Then
                tmCtrls(imBoxNo).iChg = True
                imUsThem = 2
                pbcYN_Paint
            ElseIf imUsThem = 2 Then
                tmCtrls(imBoxNo).iChg = True
                imUsThem = 1
                pbcYN_Paint
            Else
                tmCtrls(imBoxNo).iChg = True
                imUsThem = 1
                pbcYN_Paint
            End If
        End If
        mSetCommands
    End If
End Sub
Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (imBoxNo = CTRL5INDEX) And (smMnfCallType = "I") Then
        If imHardCost = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imHardCost = 1
        ElseIf imHardCost = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imHardCost = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL6INDEX) And (smMnfCallType = "I") Then
        If imTaxable = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imTaxable = 1
        ElseIf imTaxable = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imTaxable = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL8INDEX) And (smMnfCallType = "I") Then
        If imSaleType = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSaleType = 1
        ElseIf imSaleType = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSaleType = 2
        ElseIf imSaleType = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imSaleType = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "S") Then
        If imUpdateRvf = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imUpdateRvf = 1
        ElseIf imUpdateRvf = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imUpdateRvf = 2
        ElseIf imUpdateRvf = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imUpdateRvf = 3
        ElseIf imUpdateRvf = 3 Then
            tmCtrls(imBoxNo).iChg = True
            imUpdateRvf = 4
        ElseIf imUpdateRvf = 4 Then
            tmCtrls(imBoxNo).iChg = True
            imUpdateRvf = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "H") Then
        If imDollars = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imDollars = 1
        ElseIf imDollars = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imDollars = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "R") Then
        If imManOpt = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imManOpt = 2
        ElseIf imManOpt = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imManOpt = 1
        Else
            tmCtrls(imBoxNo).iChg = True
            imManOpt = 2
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "L") Then
        If imEnglish = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imEnglish = 1
        ElseIf imEnglish = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imEnglish = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "M") Then
        If imBillMGMissed = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imBillMGMissed = 2
        ElseIf imBillMGMissed = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imBillMGMissed = 1  '3
        'ElseIf imBillMGMissed = 3 Then
        '    tmCtrls(imBoxNo).iChg = True
        '    imBillMGMissed = 1
        Else
            tmCtrls(imBoxNo).iChg = True
            imBillMGMissed = 1
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "M") Then
        If imDefReason = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imDefReason = 1
        ElseIf imDefReason = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imDefReason = 0
        Else
            tmCtrls(imBoxNo).iChg = True
            imDefReason = 1
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "M") Then
        If imMissedFor = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imMissedFor = 2
        ElseIf imMissedFor = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imMissedFor = 3
        ElseIf imMissedFor = 3 Then
            tmCtrls(imBoxNo).iChg = True
            imMissedFor = 4
        ElseIf imMissedFor = 4 Then
            tmCtrls(imBoxNo).iChg = True
            imMissedFor = 1
        Else
            tmCtrls(imBoxNo).iChg = True
            imMissedFor = 1
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "N") Then
        If imTypeOfFeed = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imTypeOfFeed = 1
        ElseIf imTypeOfFeed = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imTypeOfFeed = 2
        ElseIf imTypeOfFeed = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imTypeOfFeed = 3
        ElseIf imTypeOfFeed = 3 Then
            tmCtrls(imBoxNo).iChg = True
            imTypeOfFeed = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "N") Then
        If imSubFeedAllowed = 0 Then
            tmCtrls(imBoxNo).iChg = True
            imSubFeedAllowed = 1
        ElseIf imSubFeedAllowed = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imSubFeedAllowed = 0
        End If
        pbcYN_Paint
        mSetCommands
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "O") Then
        If imUsThem = 1 Then
            tmCtrls(imBoxNo).iChg = True
            imUsThem = 2
        ElseIf imUsThem = 2 Then
            tmCtrls(imBoxNo).iChg = True
            imUsThem = 1
        End If
        pbcYN_Paint
        mSetCommands
    End If
End Sub
Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If (imBoxNo = CTRL5INDEX) And (smMnfCallType = "I") Then
        If imHardCost = 0 Then
            pbcYN.Print "Yes"
        ElseIf imHardCost = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL6INDEX) And (smMnfCallType = "I") Then
        If imTaxable = 0 Then
            pbcYN.Print "Yes"
        ElseIf imTaxable = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL8INDEX) And (smMnfCallType = "I") Then
        If imSaleType = 0 Then
            pbcYN.Print "NTR"
        ElseIf imSaleType = 1 Then
            pbcYN.Print "Agency"
        ElseIf imSaleType = 2 Then
            pbcYN.Print "Direct"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "S") Then
        If imUpdateRvf = 0 Then
            pbcYN.Print "Receivables"
        ElseIf imUpdateRvf = 1 Then
            pbcYN.Print "History"
        ElseIf imUpdateRvf = 2 Then
            pbcYN.Print "Export+History"
        ElseIf imUpdateRvf = 3 Then
            pbcYN.Print "Export+A/R"
        ElseIf imUpdateRvf = 4 Then
            pbcYN.Print "Ask by Vehicle"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "H") Then
        If imDollars = 0 Then
            pbcYN.Print "Yes"
        ElseIf imDollars = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "R") Then
        If imManOpt = 1 Then
            pbcYN.Print "Mandatory"
        ElseIf imManOpt = 2 Then
            pbcYN.Print "Optional"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "L") Then
        If imEnglish = 0 Then
            pbcYN.Print "Yes"
        ElseIf imEnglish = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "M") Then
        If imBillMGMissed = 1 Then
            pbcYN.Print "Bill MG"
        ElseIf imBillMGMissed = 2 Then
            pbcYN.Print "Bill Missed (No MG)"
        'ElseIf imBillMGMissed = 3 Then
        '    pbcYN.Print "Bill MG & Missed"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "M") Then
        If imDefReason = 0 Then
            pbcYN.Print "Yes"
        ElseIf imDefReason = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL4INDEX) And (smMnfCallType = "M") Then
        If imMissedFor = 1 Then
            pbcYN.Print "Network Missed"    '"Traffic"
        ElseIf imMissedFor = 2 Then
            pbcYN.Print "Station Missed"    '"Affiliate Web"
        ElseIf imMissedFor = 3 Then
            pbcYN.Print "Network & Station Missed"  '"
        ElseIf imMissedFor = 4 Then
            pbcYN.Print "Station Replacement"    'New
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL2INDEX) And (smMnfCallType = "N") Then
        If imTypeOfFeed = 0 Then
            pbcYN.Print "Dish"
        ElseIf imTypeOfFeed = 1 Then
            pbcYN.Print "Antenna"
        ElseIf imTypeOfFeed = 2 Then
            pbcYN.Print "CD"
        ElseIf imTypeOfFeed = 3 Then
            pbcYN.Print "Subfeed"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "N") Then
        If imSubFeedAllowed = 0 Then
            pbcYN.Print "Yes"
        ElseIf imSubFeedAllowed = 1 Then
            pbcYN.Print "No"
        Else
            pbcYN.Print "   "
        End If
    End If
    If (imBoxNo = CTRL3INDEX) And (smMnfCallType = "O") Then
        If imUsThem = 1 Then
            pbcYN.Print "Us"
        ElseIf imUsThem = 2 Then
            pbcYN.Print "Them"
        Else
            pbcYN.Print "   "
        End If
    End If
End Sub
Private Sub plcMNm_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub tmcHide_Timer()
    tmcHide.Enabled = False
    cmcCancel_Click
End Sub
