VERSION 5.00
Begin VB.Form RateCard 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6150
   ClientLeft      =   1095
   ClientTop       =   1815
   ClientWidth     =   9210
   ClipControls    =   0   'False
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
   Icon            =   "Ratecard.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   9210
   Begin VB.CommandButton cmcCPMpkg 
      Appearance      =   0  'Flat
      Caption         =   "&CPM Pkg"
      Height          =   285
      HelpContextID   =   7
      Left            =   5760
      TabIndex        =   37
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmcDupl 
      Appearance      =   0  'Flat
      Caption         =   "Duplicate"
      Height          =   285
      HelpContextID   =   7
      Left            =   6675
      TabIndex        =   38
      Top             =   5640
      Width           =   900
   End
   Begin VB.CommandButton cmcResetStdPrice 
      Appearance      =   0  'Flat
      Caption         =   "Recompute Package Rates from Hidden"
      Height          =   225
      HelpContextID   =   1
      Left            =   75
      TabIndex        =   66
      Top             =   5880
      Width           =   3465
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "Import Rates from a file"
      Height          =   225
      HelpContextID   =   1
      Left            =   90
      TabIndex        =   65
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   1100
      Left            =   9195
      Top             =   3870
   End
   Begin VB.ComboBox cbcSelect 
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
      Left            =   5400
      TabIndex        =   1
      Top             =   60
      Width           =   3795
   End
   Begin VB.CommandButton cmcRealloc 
      Appearance      =   0  'Flat
      Caption         =   "Re&alloc $"
      Height          =   285
      HelpContextID   =   7
      Left            =   8265
      TabIndex        =   40
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmcStdPkg 
      Appearance      =   0  'Flat
      Caption         =   "Std &Pkg"
      Height          =   285
      HelpContextID   =   7
      Left            =   4950
      TabIndex        =   36
      Top             =   5640
      Width           =   855
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Left            =   3060
      ScaleHeight     =   210
      ScaleWidth      =   315
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Height          =   285
      Left            =   2655
      TabIndex        =   33
      Top             =   5640
      Width           =   855
   End
   Begin VB.ListBox lbcBudget2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1950
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2115
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.CommandButton cmcImpact 
      Appearance      =   0  'Flat
      Caption         =   "&Impact"
      Height          =   285
      HelpContextID   =   7
      Left            =   7515
      TabIndex        =   39
      Top             =   5640
      Width           =   855
   End
   Begin VB.ListBox lbcDPName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1830
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   330
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   2550
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
      Height          =   2040
      Left            =   150
      Picture         =   "Ratecard.frx":08CA
      ScaleHeight     =   2010
      ScaleWidth      =   3210
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.ListBox lbcDPNameRow 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4845
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2055
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.ListBox lbcLen 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   5445
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1395
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcGrid 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   7440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1695
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ListBox lbcBudget 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1860
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1695
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.TextBox edcSPDropDown 
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
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmcSPDropDown 
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
      Left            =   2400
      Picture         =   "Ratecard.frx":15A24
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   405
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcSPSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   135
      Left            =   6675
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   195
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox plcRCInfo 
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
      Height          =   705
      Left            =   210
      ScaleHeight     =   675
      ScaleWidth      =   9135
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Label lacRCInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Start Time xx:xx:xxam  Length xx:xx:xx"
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
         Index           =   2
         Left            =   105
         TabIndex        =   48
         Top             =   435
         Width           =   8925
      End
      Begin VB.Label lacRCInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Start Time xx:xx:xxam  Length xx:xx:xx"
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
         Left            =   105
         TabIndex        =   47
         Top             =   225
         Width           =   8925
      End
      Begin VB.Label lacRCInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Library Name xxxxxxxxxxxxxxxxxx  Version xx"
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
         Left            =   105
         TabIndex        =   46
         Top             =   30
         Width           =   8970
      End
   End
   Begin VB.PictureBox pbcView 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Left            =   210
      ScaleHeight     =   210
      ScaleWidth      =   1305
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   855
      Width           =   1305
   End
   Begin VB.PictureBox pbcSP 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1275
      Picture         =   "Ratecard.frx":15B1E
      ScaleHeight     =   390
      ScaleWidth      =   2340
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   165
      Visible         =   0   'False
      Width           =   2340
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
      Left            =   9165
      TabIndex        =   51
      Top             =   4605
      Visible         =   0   'False
      Width           =   255
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
      Left            =   9135
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   9180
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4155
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      HelpContextID   =   7
      Left            =   4185
      TabIndex        =   35
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Enabled         =   0   'False
      Height          =   285
      HelpContextID   =   5
      Left            =   3420
      TabIndex        =   34
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      HelpContextID   =   3
      Left            =   1890
      TabIndex        =   32
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      HelpContextID   =   2
      Left            =   1005
      TabIndex        =   31
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      HelpContextID   =   1
      Left            =   120
      TabIndex        =   30
      Top             =   5640
      Width           =   855
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
      Left            =   1455
      Picture         =   "Ratecard.frx":16B20
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1575
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
      Left            =   4305
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Left            =   630
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5775
      Width           =   105
   End
   Begin VB.PictureBox pbcStartNew 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   135
      Left            =   4320
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   -30
      Width           =   105
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   135
      Left            =   15
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   27
      Top             =   3915
      Width           =   105
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   135
      Left            =   -15
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   15
      Top             =   810
      Width           =   105
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   105
      ScaleHeight     =   180
      ScaleWidth      =   885
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   885
   End
   Begin VB.PictureBox plcSP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1170
      ScaleHeight     =   420
      ScaleWidth      =   2280
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   180
      Left            =   0
      Picture         =   "Ratecard.frx":16C1A
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1425
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.CommandButton cmcTerms 
      Appearance      =   0  'Flat
      Caption         =   "&Terms"
      Height          =   285
      HelpContextID   =   7
      Left            =   4155
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   315
      Width           =   855
   End
   Begin VB.PictureBox pbcDaypart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2970
      Left            =   255
      Picture         =   "Ratecard.frx":16F24
      ScaleHeight     =   2970
      ScaleWidth      =   8745
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   8745
      Begin VB.Label lacDPFrame 
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
         Left            =   -30
         TabIndex        =   21
         Top             =   675
         Visible         =   0   'False
         Width           =   8760
      End
   End
   Begin VB.PictureBox pbcRateCard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   150
      Picture         =   "Ratecard.frx":33366
      ScaleHeight     =   2955
      ScaleWidth      =   8745
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   855
      Width           =   8745
      Begin VB.PictureBox pbcLnWkArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3960
         Picture         =   "Ratecard.frx":4F7A8
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   64
         Top             =   15
         Width           =   270
      End
      Begin VB.PictureBox pbcLnWkArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   7875
         Picture         =   "Ratecard.frx":4FA5A
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   63
         Top             =   105
         Width           =   270
      End
      Begin VB.TextBox edcBdDropDown 
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
         Left            =   1365
         MaxLength       =   20
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmcBdDropDown 
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
         Left            =   3585
         Picture         =   "Ratecard.frx":4FD0C
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lacRCFrame 
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
         Left            =   15
         TabIndex        =   18
         Top             =   645
         Visible         =   0   'False
         Width           =   8760
      End
   End
   Begin VB.PictureBox plcRateCard 
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
      Height          =   3075
      Left            =   105
      ScaleHeight     =   3015
      ScaleWidth      =   9075
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   660
      Width           =   9135
      Begin VB.VScrollBar vbcRateCard 
         Height          =   2940
         LargeChange     =   12
         Left            =   8820
         TabIndex        =   12
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.PictureBox pbcStatic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   165
      Picture         =   "Ratecard.frx":4FE06
      ScaleHeight     =   1410
      ScaleWidth      =   8745
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3975
      Visible         =   0   'False
      Width           =   8745
   End
   Begin VB.PictureBox plcStatic 
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
      Height          =   1530
      Left            =   225
      ScaleHeight     =   1470
      ScaleWidth      =   8820
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   8880
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3915
      Top             =   -285
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3420
      Top             =   -270
   End
   Begin VB.ListBox lbcDPNameCode 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6180
      Sorted          =   -1  'True
      TabIndex        =   43
      Top             =   -45
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox plcShow 
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
      Height          =   225
      Left            =   615
      ScaleHeight     =   225
      ScaleWidth      =   3225
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   3795
      Width           =   3225
      Begin VB.OptionButton rbcShow 
         Caption         =   "Standard"
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
         Left            =   1920
         TabIndex        =   54
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton rbcShow 
         Caption         =   "Corporate"
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
         Left            =   795
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox plcType 
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
      Height          =   225
      Left            =   5070
      ScaleHeight     =   225
      ScaleWidth      =   3960
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3795
      Width           =   3960
      Begin VB.OptionButton rbcType 
         Caption         =   "Flight"
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
         Index           =   3
         Left            =   2580
         TabIndex        =   59
         Top             =   0
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Week"
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
         Index           =   2
         Left            =   1785
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   0
         Width           =   840
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Month"
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
         Left            =   945
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   0
         Width           =   885
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Quarter"
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
         Left            =   0
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   0
         Width           =   990
      End
   End
   Begin VB.PictureBox pbcSPTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   135
      Left            =   6780
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   14
      Top             =   660
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   165
      Picture         =   "Ratecard.frx":5D508
      Top             =   420
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   9015
      Picture         =   "Ratecard.frx":5D812
      Top             =   5505
      Width           =   480
   End
End
Attribute VB_Name = "RateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Ratecard.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RateCard.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Rate Card input screen code
'
'   Dormant vehicles are not showed on the rate card screen,  to show add DORMANTVEH to mVehPop
'   and remove tmVef.sState <> "D" test in mReadRifRec
'
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
'Rate Card Spec Areas
Dim tmSPCtrls(0 To 3)  As FIELDAREA
Dim imLBSPCtrls As Integer
Dim imSPBoxNo As Integer   'Current Rate Card Box
Dim imAdjIndex As Integer
'Rate Card Field Areas
Dim tmRCCtrls(0 To 12)  As FIELDAREA
Dim imLBRCCtrls As Integer
Dim tmWKCtrls(0 To 4)  As FIELDAREA
Dim imLBWKCtrls As Integer
Dim tmNWCtrls(0 To 5)  As FIELDAREA
Dim imLBNWCtrls As Integer
Public imRCBoxNo As Integer   'Current Rate Card Box
Public lmRCRowNo As Long      'Current row number in Program area (start at 0)
Dim bmInStdPrice As Boolean 'Ignore GotFocus
Dim bmInImportPrice As Boolean
'Daypart Field Areas
Dim tmDPCtrls(0 To 13)  As FIELDAREA
Dim imLBDPCtrls As Integer
Dim imDPBoxNo As Integer   'Current Daypart Box
Dim imMaxTDRows As Integer
'Static Field Areas
Dim tmSTCtrls(0 To 7)  As FIELDAREA
Dim imLBSTCtrls As Integer
'Rate Card
Dim hmRcf As Integer    'Rate Card file handle
Dim tmRcf As RCF
Dim tmRcfSrchKey As INTKEY0    'Rcf key record image
Dim imRcfRecLen As Integer        'Rcf record length
'Rate Card Item
Dim hmRif As Integer    'Rate Card item file handle
Dim tmTrashRifRec() As RIFREC  'Multiyear records
Dim tmRifSrchKey1 As LONGKEY0
Dim imRifRecLen As Integer        'Rpf record length
Dim imRifChg As Integer  'True=Vehicle or daypart value changed; False=No changes
'Dim tmDPBudgetInfo() As DPBUDGETINFO
Dim tmMRif() As RIF
Dim lmAutoDelRif() As Long
'Daypart
Dim hmRdf As Integer
Dim tmRdf As RDF
Dim imRdfRecLen As Integer
'Dim sgMRdfStamp As String
'Dim tgMRdf() As RDF
'Library
Dim hmLtf As Integer        'Library title file handle
Dim tmLtf As LTF            'LTF record image
Dim tmLtfSrchKey As INTKEY0 'LTF key record image
Dim imLtfRecLen As Integer  'LTF record length
Dim hmLcf As Integer
'Avail Name
Dim hmAnf As Integer        'Avail Name file handle
Dim tmAnf As ANF            'ANF record image
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length
'Vehicle
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
Dim tmUserVeh() As USERVEH
Dim smSvVehName As String
'Dim imVefCode() As Integer
'Virtual Vechcle
Dim hmVsf As Integer        'Virtual Vehicle file handle
Dim tmVsf As VSF            'VSF record image
Dim tmVsfSrchKey As LONGKEY0 'VSF key record image
Dim imVsfRecLen As Integer  'VSF record length
'Multi-Name: Vehicle Group
Dim hmMnf As Integer        'Multi-Name file handle
Dim tmMnf As MNF            'MNF record image
Dim tmMnfSrchKey As INTKEY0 'MNF key record image
Dim imMnfRecLen As Integer  'MNF record length
'Contract
Dim hmCHF As Integer        'Contract file handle
Dim tmChf As CHF            'CHF record image
Dim imCHFRecLen As Integer  'CHF record length
'Line
Dim hmClf As Integer        'Line file handle
Dim tmClf As CLF            'CLF record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClfModel() As CLFLIST
'Flight
Dim hmCff As Integer        'Flight file handle
Dim tmCff As CFF            'CFF record image
Dim imCffRecLen As Integer  'CFF record length
'Spot
Dim hmSdf As Integer    'file handle
Dim imSdfRecLen As Integer  'Record length
Dim tmSdf As SDF
'MG File
Dim hmSmf As Integer    'file handle
Dim imSmfRecLen As Integer  'Record length
Dim tmSmf As SMF
'Standard Package Vechcle
Dim hmPvf As Integer        'Standard Vehicle file handle
Dim tmPvf() As PVF            'PVF record image
Dim tmTPvf As PVF
Dim tmPvfSrchKey As LONGKEY0 'PVF key record image
Dim imPvfRecLen As Integer  'PVF record length
'Site Options
Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim tmSafSrchKey1 As SAFKEY1    'Vef key record image
Dim imSafRecLen As Integer

''Daypart Weekly Avails
'Dim hmDaf As Integer        'Daypart Weekly Aavils file handle
'Dim tmDaf As DAF            'DAF record image
'Dim tmDafSrchKey As DAFKEY0 'DAF key record image
'Dim imDafRecLen As Integer  'DAF record length
Dim imShowIndex As Integer
Dim imTypeIndex As Integer
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length
'Spot Summary
Dim hmSsf As Integer    'file handle
'Dim tmRec As LPOPREC
'Period (column) Information
Dim imPdYear As Integer
Dim imPdStartWk As Integer 'start week number
Dim imPdStartFltNo As Integer
Dim imRifStartYear As Integer
Dim imRifNoYears As Integer
Dim tmPdGroups(0 To 4) As PDGROUPS      'Index zero ignored
Dim imHotSpot(0 To 4, 0 To 4) As Integer    'Index zero ignored
Dim imInHotSpot As Integer
'Budget information
Dim hmBvf As Integer    'Rate Card file handle
Dim imBvfRecLen As Integer        'Rcf record length
Dim tmBvf As BVF
Dim tmBudgetCode() As SORTCODE
Dim smBudgetCodeTag As String
Dim lmBdStartDate As Long
Dim lmBdEndDate As Long
Dim smBdMnfName As String
Dim imBdMnfCode As Integer
Dim imBdYear As Integer
Dim imBdMnf() As Integer
Dim imBdYr() As Integer
Dim imBSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new and not hub)
Dim imInNew As Integer
Dim imIgnoreSetting As Integer
Dim imGDSelectedIndex As Integer  'Index of Grid Level
Dim imLenSelectedIndex As Integer 'Index of length
Dim imCGDSelectedIndex As Integer  'Index of Current Grid Level
Dim imComboBoxIndex As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imLen As Integer    'Number of lengths defined in the rate card
Dim imFirstTimeSelect As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imSettingValue As Integer   'True=Don't enable any box with change
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim smDefVehicle As String
Dim imBypassFocus As Integer
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imUpdateAllowed As Integer
Dim imIgnoreRightMove As Integer
Dim imButtonIndex As Integer
Dim imView As Integer   '0=Rate; 1=Daypart
Dim imRetBranch As Integer
Dim lmNowDate As Long
Dim imNowYear As Integer
'Dim tgTempRCUserVehicle() As SORTCODE
'Dim tgTempRCUserVehicleTag As String
Dim tmTempRifRec() As RIFREC

Dim tmChfAdvtExt() As CHFADVTEXT
Dim tmRCModelInfo() As RCMODELINFO
Dim tgTempRCUserVehicle() As SORTCODE
Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Dim bmInDupicate As Boolean

Const LBONE = 1

'Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system
Const VEHINDEX = 1          'Vehicle control/field
Const DAYPARTINDEX = 2      'Daypart name control/field
'Const DOLLARINDEX = 3
'Const PCTINVINDEX = 4
Const CPMINDEX = 3
Const ACQUISITIONINDEX = 4
Const BASEINDEX = 5
Const RPTINDEX = 6
Const SORTINDEX = 7
Const DOLLAR1INDEX = 8  '5
Const DOLLAR2INDEX = 9  '6
Const DOLLAR3INDEX = 10  '7
Const DOLLAR4INDEX = 11  '8
Const AVGINDEX = RCAVGRATEINDEX 'JW, this Col is shared between StdPkg and CPMPkg screens -  12 'RCAVGRATEINDEX '9      'Also in StdPkg.Frm
Const TOTALINDEX = 13   '10

Const TIMESINDEX = 4
Const DAYINDEX = 5
Const AVAILINDEX = 12
Const HRSINDEX = 13
Const DPBASEINDEX = 14    'used to know which color to use when painting daypart name
'Const BUDGETINDEX = 1       'Comparison Budget control/field
Const GRIDINDEX = 1 '2    'Grid Level control/field
Const LENGTHINDEX = 2   '3       'Length control/field
Const CURGRIDINDEX = 3  '4      'Current Grid Level control/field
Const STVEHINDEX = 1
Const STTITLEINDEX = 2
Const STDOLLAR1INDEX = 3
Const STDOLLAR2INDEX = 4
Const STDOLLAR3INDEX = 5
Const STDOLLAR4INDEX = 6
Const STAVGINDEX = 7
Const WK1INDEX = 1
Const WK2INDEX = 2
Const WK3INDEX = 3
Const WK4INDEX = 4
Const NW1INDEX = 1
Const NW2INDEX = 2
Const NW3INDEX = 3
Const NW4INDEX = 4
Const NWAVGINDEX = 5

Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilLoopCount As Integer
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcSelect.ListIndex >= 0 Then
                    cbcSelect.Text = cbcSelect.List(cbcSelect.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
            If ilRet = 0 Then
                ilIndex = cbcSelect.ListIndex   'cbcSelect.ListCount - cbcSelect.ListIndex
                If Not mReadRec(ilIndex) Then
                    GoTo cbcSelectErr
                End If
                If Not mReadRifRec(tgRcfI.iCode, True) Then
                    GoTo cbcSelectErr
                End If
                mInitRateCardCtrls
                mInitRif UBound(tmRifRec)
                igRCMode = 1    'Change
            Else
                If ilRet = 1 Then
                    cbcSelect.ListIndex = 0
                End If
                ilRet = 1   'Clear fields as no match name found
                igRCMode = 0    'New
            End If
            pbcRateCard.Cls
            pbcDaypart.Cls
            If ilRet = 0 Then
                imSelectedIndex = cbcSelect.ListIndex
                mMoveRecToCtrl
                'mInitShow
            Else
                If imAdjIndex = 1 Then
                    imSelectedIndex = 0
                Else
                    imSelectedIndex = -1
                End If
                mClearCtrlFields
            End If
            'pbcRateCard_Paint
            If imView = 1 Then
                pbcDaypart_Paint
            Else
                pbcRateCard_Paint
            End If
        Loop While (imSelectedIndex <> cbcSelect.ListIndex) And ((imSelectedIndex <> 0) Or (cbcSelect.ListIndex >= 0))
        imFirstTimeSelect = True
        mSetCommands
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault    'Default
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcSelect_DropDown()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub

Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   ' so list box can be reset
    If bmInStdPrice Then
        'bmInStdPrice = False
        Exit Sub
    End If
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If imFirstTime Then
        imFirstTime = False
    End If
    If cbcSelect.ListCount <= 1 Then
        igRCMode = 0    'New
        imFirstTimeSelect = True
        pbcStartNew.SetFocus
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 1 'Force to newest instead of [New]
'        cbcSelect_Change
'        cbcSelect.ListIndex = 0
'        cbcSelect_Change    'Call change so picture area repainted
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
    mSetCommands
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

Private Sub cmcBdDropDown_Click()
    If tgSpf.sRUseCorpCal <> "Y" Then
        lbcBudget.Visible = Not lbcBudget.Visible
        edcBdDropDown.SelStart = 0
        edcBdDropDown.SelLength = Len(edcBdDropDown.Text)
        edcBdDropDown.SetFocus
    Else
        cmcBdDropDown.Visible = False
        edcBdDropDown.Visible = False
        lbcBudget2.Visible = Not lbcBudget2.Visible
    End If
End Sub

Private Sub cmcBdDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus cmcCancel
End Sub

Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcCPMpkg_Click()
    'PODCAST - 12/22/2020 - Add CPM Pkg to Rate Card Screen:
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim llRowNo As Long
    Dim ilTest As Integer
    Dim slStr As String
    Dim ilVpf As Integer
    
    igDPNameCallSource = CALLSOURCERATECARD
    If lmRCRowNo <= 0 Then
        sgDPName = ""
    Else
        'Test if Package
        sgDPName = ""
        gFindMatch Trim$(smRCSave(VEHINDEX, lmRCRowNo)), 0, lbcVehicle
        ilIndex = gLastFound(lbcVehicle)
        'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
        If ilIndex >= 0 And ilIndex < UBound(tgTempRCUserVehicle) Then
            slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilCode = CInt(slCode)
            ilLoop = gBinarySearchVef(ilCode)
            If ilLoop <> -1 Then
                If (tgMVef(ilLoop).sType = "P") And (tgMVef(ilLoop).lPvfCode > 0) Then
                    ilVpf = gBinarySearchVpf(ilCode)
                    If ilVpf <> -1 Then
                        If tgVpf(ilVpf).sGMedium = "P" Then
                            'make sure this is a POD-CAST CPM
                            sgDPName = Trim$(smRCSave(VEHINDEX, lmRCRowNo))
                        End If
                    End If
                End If
            End If
        End If
    End If
    'PODCAST - 12/22/2020 - Add CPM Pkg screen
    CPMPkg.Show vbModal
    sgRCUserVehicleTag = ""
    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
    lbcVehicle.Clear
    mVehPop lbcVehicle
    For llRowNo = LBONE To UBound(tmRifRec) - 1 Step 1
        If imRCSave(11, llRowNo) = 1 Then
            'For ilTest = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
            '    If tmUserVeh(ilTest).iCode = tmRifRec(llRowNo).tRif.iVefCode Then
                    ilTest = mBinarySearch(tmRifRec(llRowNo).tRif.iVefCode)
                    If ilTest <> -1 Then
                        smRCSave(VEHINDEX, llRowNo) = tmUserVeh(ilTest).sName
                        slStr = smRCSave(VEHINDEX, llRowNo)
                        gSetShow pbcRateCard, slStr, tmRCCtrls(VEHINDEX)
                        smRCShow(VEHINDEX, llRowNo) = tmRCCtrls(VEHINDEX).sShow
                        smDPShow(VEHINDEX, llRowNo) = tmRCCtrls(VEHINDEX).sShow
                    End If
            '        Exit For
            '    End If
            'Next ilTest
        End If
    Next llRowNo
    If imView = 1 Then
        pbcDaypart.Cls
        DoEvents
        pbcDaypart_Paint
    Else
        pbcRateCard.Cls
        DoEvents
        pbcRateCard_Paint
    End If
    mSetCommands
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If (mSaveRecChg(True) = False) And (imUpdateAllowed) Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If imRCBoxNo > 0 Then
            mRCEnableBox imRCBoxNo
        End If
        If imSPBoxNo > 0 Then
            mSPEnableBox imSPBoxNo
        End If
        Exit Sub
    End If
    tmcDelay.Enabled = False
    mTerminate
End Sub

Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    Dim slStr As String
    If (lmRCRowNo = UBound(tmRifRec)) Then
        For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
            slStr = ""
            gSetShow pbcRateCard, slStr, tmRCCtrls(ilLoop)
            smRCShow(ilLoop, lmRCRowNo) = tmRCCtrls(ilLoop).sShow
        Next ilLoop
        pbcRateCard_Paint
        mSetDefInSave   'Set defaults for extra row
    End If
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus cmcDone
End Sub

Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcDropDown_Click()
    Select Case imRCBoxNo
        Case VEHINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
        Case DAYPARTINDEX
            lbcDPNameRow.Visible = Not lbcDPNameRow.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDupl_Click()
    Dim llUpper As Long
    Dim ilWk As Integer
    
    llUpper = UBound(tmRifRec)
    igRCNoDollarColumns = 0
    For ilWk = imLBWKCtrls To UBound(tmWKCtrls) Step 1
        If Trim$(tmWKCtrls(ilWk).sShow) <> "" Then
            igRCNoDollarColumns = igRCNoDollarColumns + 1
        End If
    Next ilWk
    bmInDupicate = True
    RCItemDupl.Show vbModal
    bmInDupicate = False
    If llUpper <> UBound(tmRifRec) Then
        imRifChg = True
    End If
    If imView = 1 Then
        pbcDaypart.Cls
        DoEvents
        pbcDaypart_Paint
    Else
        pbcRateCard.Cls
        DoEvents
        pbcRateCard_Paint
    End If
End Sub

Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim slMsg As String
    Dim llLkYear As Long
    'Dim ilSvLkYear As Integer
    Dim llRif As Long
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim tlRif As RIF
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imAdjIndex = 0 Then
        Exit Sub
    End If
    If (imSelectedIndex > 0) And (igRCMode <> 0) And (tgRcfI.iCode <> 0) Then
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(RateCard, tgRcfI.iCode, "Chf.Btr", "ChfRcfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - a Contract references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        ilRet = MsgBox("OK to remove " & tgRcfI.sName, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        gGetSyncDateTime slSyncDate, slSyncTime
        If Not mReadRec(imSelectedIndex) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        If Not mReadRifRec(tgRcfI.iCode, True) Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        For llRif = LBONE To UBound(tmRifRec) - 1 Step 1
            If tmRifRec(llRif).iStatus = 1 Then
                Do  'Loop until record updated or added
                    'ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmRifRec(ilRif).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    tmRifSrchKey1.lCode = tmRifRec(llRif).tRif.lCode
                    ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    llLkYear = tmRifRec(llRif).lLkYear
                    If ilRet <> BTRV_ERR_NONE Then
                        Screen.MousePointer = vbDefault
                        ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                        Exit Sub
                    End If
                    'tmRec = tlRif
                    'ilRet = gGetByKeyForUpdate("RIF", hmRif, tmRec)
                    'tlRif = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    Screen.MousePointer = vbDefault
                    '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                    '    Exit Sub
                    'End If
                    'ilRet = btrDelete(hmRif)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                    Exit Sub
                End If
                llLkYear = tmRifRec(llRif).lLkYear
                Do While llLkYear > 0
                    If tmLkRifRec(llLkYear).iStatus = 1 Then
                        Do  'Loop until record updated or added
                            'ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmLkRifRec(llLkYear).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            tmRifSrchKey1.lCode = tmLkRifRec(llLkYear).tRif.lCode
                            ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet <> BTRV_ERR_NONE Then
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                                Exit Sub
                            End If
                            'tmRec = tlRif
                            'ilRet = gGetByKeyForUpdate("RIF", hmRif, tmRec)
                            'tlRif = tmRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    Screen.MousePointer = vbDefault
                            '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                            '    Exit Sub
                            'End If
                            ilRet = btrDelete(hmRif)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            Screen.MousePointer = vbDefault
                            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
                            Exit Sub
                        End If
                    End If
                    llLkYear = tmLkRifRec(llLkYear).lLkYear
                Loop
            End If
        Next llRif
        'ilRet = btrGetPosition(hmRcf, llRcfRecPos)
        'tmRec = tgRcfI
        'ilRet = gGetByKeyForUpdate("RCF", hmRcf, tmRec)
        ilRet = btrDelete(hmRcf)
        If ilRet <> BTRV_ERR_NONE Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
'        If tgSpf.sRemoteUsers = "Y" Then
'            tmDsf.lCode = 0
'            tmDsf.sFileName = "RCF"
'            gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'            gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'            tmDsf.iRemoteID = tmRcf.iRemoteID
'            tmDsf.lAutoCode = tmRcf.iAutoCode
'            tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'            tmDsf.lCntrNo = 0
'            ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'            If ilRet <> BTRV_ERR_NONE Then
'                Screen.MousePointer = vbDefault
'                ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
'                Exit Sub
'            End If
'        End If
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    pbcRateCard.Cls
    smRateCardTag = ""
    sgMRcfStamp = ""
    mPopulate
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Resume Next
End Sub

Private Sub cmcImpact_Click()
    'If Not gWinRoom(igNoExeWinRes(RCTERMSEXE)) Then
    '    Exit Sub
    'End If
    RCImpact.Show vbModal
End Sub

Private Sub cmcImpact_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus cmcImpact
End Sub

Private Sub cmcImpact_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcImport_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ImptRCPrice.Show vbModal
End Sub

Private Sub cmcImport_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcRealloc_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    'If Not gWinRoom(igNoExeWinRes(RCTERMSEXE)) Then
    '    Exit Sub
    'End If
    RCReallo.Show vbModal
End Sub

Private Sub cmcRealloc_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus cmcRealloc
End Sub

Private Sub cmcRealloc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcReport_Click()
    Dim slStr As String        'General string
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = RATECARDSJOB
    igRptType = 0
    'Screen.MousePointer = vbHourglass  'Wait
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    ''Traffic!edcLinkSrceHelpMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "RateCard^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
        Else
            slStr = "RateCard^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "RateCard^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    Else
    '        slStr = "RateCard^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'RateCard.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'RateCard.Enabled = True
    ''Traffic!edcLinkSrceHelpMsg.Text = "Ok"
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    sgCommandStr = slStr
    RptList.Show vbModal
    'MousePointer = vbDefault
End Sub

Private Sub cmcReport_GotFocus()
    Dim ilLoop As Integer
    Dim slStr As String
    If (lmRCRowNo = UBound(tmRifRec)) Then
        For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
            slStr = ""
            gSetShow pbcRateCard, slStr, tmRCCtrls(ilLoop)
            smRCShow(ilLoop, lmRCRowNo) = tmRCCtrls(ilLoop).sShow
        Next ilLoop
        pbcRateCard_Paint
        mSetDefInSave   'Set defaults for extra row
    End If
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus cmcReport
End Sub

Private Sub cmcReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcResetStdPrice_Click()
    Dim ilRet As Integer
    Dim ilRow As Integer
    
    ilRet = MsgBox("All Package Rates will be recomputed from Hidden Vehicles, OK to Continue?", vbYesNo + vbQuestion, "Rate Card")
    If ilRet = vbNo Then
        Exit Sub
    End If
    If imView <> 2 Then
        Screen.MousePointer = vbHourglass
        mResetStdPrice
        pbcRateCard.Cls
        pbcRateCard_Paint
        imRifChg = True
        Screen.MousePointer = vbDefault
    Else
        'If ilEndRow > UBound(smBdShow, 2) Then
        '    ilEndRow = UBound(smBdShow, 2) 'include blank row as it might have data
        'End If
        'For ilRow = ilStartRow To ilEndRow Step 1
        'Next llRow
    End If
    mSetCommands
End Sub

Private Sub cmcResetStdPrice_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcSPDropDown_Click()
    Select Case imSPBoxNo
        'Case BUDGETINDEX
        '    lbcBudget.Visible = Not lbcBudget.Visible
        Case GRIDINDEX
            lbcGrid.Visible = Not lbcGrid.Visible
        Case LENGTHINDEX
            lbcLen.Visible = Not lbcLen.Visible
        Case CURGRIDINDEX
            lbcGrid.Visible = Not lbcGrid.Visible
    End Select
    edcSPDropDown.SelStart = 0
    edcSPDropDown.SelLength = Len(edcSPDropDown.Text)
    edcSPDropDown.SetFocus
End Sub

Private Sub cmcSPDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcStdPkg_Click()
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim llRowNo As Long
    Dim ilTest As Integer
    Dim slStr As String
    Dim ilVpf As Integer
    
    igDPNameCallSource = CALLSOURCERATECARD
    If lmRCRowNo <= 0 Then
        sgDPName = ""
    Else
        'Test if Package
        sgDPName = ""
        gFindMatch Trim$(smRCSave(VEHINDEX, lmRCRowNo)), 0, lbcVehicle
        ilIndex = gLastFound(lbcVehicle)
        'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
        If ilIndex >= 0 And ilIndex < UBound(tgTempRCUserVehicle) Then
            slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilCode = CInt(slCode)
            'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If ilCode = tgMVef(ilLoop).iCode Then
                ilLoop = gBinarySearchVef(ilCode)
                If ilLoop <> -1 Then
                    If (tgMVef(ilLoop).sType = "P") And (tgMVef(ilLoop).lPvfCode > 0) Then
                        ilVpf = gBinarySearchVpf(ilCode)
                        If ilVpf <> -1 Then
                            If tgVpf(ilVpf).sGMedium <> "P" Then
                                'make sure this is NOT a POD-CAST CPM
                                sgDPName = Trim$(smRCSave(VEHINDEX, lmRCRowNo))
                            End If
                        End If
                    End If
                End If
            '        Exit For
            '    End If
            'Next ilLoop
        End If
    End If
    StdPkg.Show vbModal
    sgRCUserVehicleTag = ""
    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
    lbcVehicle.Clear
    mVehPop lbcVehicle
    For llRowNo = LBONE To UBound(tmRifRec) - 1 Step 1
        If imRCSave(11, llRowNo) = 1 Then
            'For ilTest = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
            '    If tmUserVeh(ilTest).iCode = tmRifRec(llRowNo).tRif.iVefCode Then
                    ilTest = mBinarySearch(tmRifRec(llRowNo).tRif.iVefCode)
                    If ilTest <> -1 Then
                        smRCSave(VEHINDEX, llRowNo) = tmUserVeh(ilTest).sName
                        slStr = smRCSave(VEHINDEX, llRowNo)
                        gSetShow pbcRateCard, slStr, tmRCCtrls(VEHINDEX)
                        smRCShow(VEHINDEX, llRowNo) = tmRCCtrls(VEHINDEX).sShow
                        smDPShow(VEHINDEX, llRowNo) = tmRCCtrls(VEHINDEX).sShow
                    End If
            '        Exit For
            '    End If
            'Next ilTest
        End If
    Next llRowNo
    If imView = 1 Then
        pbcDaypart.Cls
        DoEvents
        pbcDaypart_Paint
    Else
        pbcRateCard.Cls
        DoEvents
        pbcRateCard_Paint
    End If
    mSetCommands
End Sub

Private Sub cmcTerms_Click()
    Dim ilRCMode As Integer
    'If Not gWinRoom(igNoExeWinRes(RCTERMSEXE)) Then
    '    Exit Sub
    'End If
    ilRCMode = igRCMode
    igRCMode = 1    'Change
    RCTerms.Show vbModal
    igRCMode = ilRCMode
    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
    lbcVehicle.Clear
    mVehPop lbcVehicle
    mSetCommands
End Sub

Private Sub cmcTerms_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus cmcTerms
End Sub

Private Sub cmcTerms_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcUndo_Click()
    Dim ilIndex As Integer
    ilIndex = imSelectedIndex
    If (((imSelectedIndex > 0) And (imAdjIndex = 1)) Or ((imSelectedIndex >= 0) And (imAdjIndex = 0))) And (tgRcfI.iCode <> 0) Then 'Not New selected
        If ilIndex > 0 Then
            If Not mReadRec(ilIndex) Then
                GoTo cmcUndoErr
            End If
            If Not mReadRifRec(tgRcfI.iCode, True) Then
                GoTo cmcUndoErr
                Exit Sub
            End If
            pbcRateCard.Cls
            mInitRateCardCtrls
            mMoveRecToCtrl
            pbcRateCard_Paint
            igRcfChg = False
            imRifChg = False
            mSetCommands
            imRCBoxNo = -1
            imSPBoxNo = -1 'Initialize current Box to N/A
            pbcSTab.SetFocus
            Exit Sub
        End If
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    igRcfChg = False
    imRifChg = False
    imRCBoxNo = -1
    imSPBoxNo = -1 'Initialize current Box to N/A
    imSelectedIndex = 0
    cbcSelect.RemoveItem 1
    pbcRateCard.Cls
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
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus cmcUndo
End Sub

Private Sub cmcUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub cmcUpdate_Click()
    Dim slName As String    'Save name as MNmSave set listindex to 0 which clears values from controls
    Dim imSvSelectedIndex As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slName = cbcSelect.Text   'Save name
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mRCEnableBox imRCBoxNo
        Exit Sub
    End If
    
    smRateCardTag = ""
    mPopulate
'    imRCBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
'    imFirstTimeSelect = True
'    igRcfChg = False
'    imRifChg = False
'    'llRowNo = UBound(tmRifRec)
'    'For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
'    '    slStr = ""
'    '    gSetShow pbcRateCard, slStr, tmRCCtrls(ilLoop)
'    '    smShow(ilLoop, llRowNo) = tmRCCtrls(ilLoop).sShow
'    'Next ilLoop
'    pbcRateCard_Paint
'    mSetDefInSave   'Set defaults for extra row
'    mSetCommands
'    pbcSTab.SetFocus
End Sub

Private Sub cmcUpdate_GotFocus()
    Dim ilLoop As Integer
    Dim slStr As String
    If (lmRCRowNo = UBound(tmRifRec)) Then
        For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
            slStr = ""
            gSetShow pbcRateCard, slStr, tmRCCtrls(ilLoop)
            smRCShow(ilLoop, lmRCRowNo) = tmRCCtrls(ilLoop).sShow
        Next ilLoop
        pbcRateCard_Paint
        mSetDefInSave   'Set defaults for extra row
    End If
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus cmcUpdate
End Sub

Private Sub cmcUpdate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub edcBdDropDown_Change()
    Dim ilRet As Integer
    If tgSpf.sRUseCorpCal = "Y" Then
        Exit Sub
    End If
    If imChgMode = False Then
        imChgMode = True
        imLbcArrowSetting = True
        gMatchLookAhead edcBdDropDown, lbcBudget, imBSMode, imComboBoxIndex
        imLbcArrowSetting = False
        If lbcBudget.ListIndex >= 0 Then
            imBSelectedIndex = lbcBudget.ListIndex
            'pbcRateCard.Cls
            ilRet = mBdBuildCompBudget()
            If Not ilRet Then
                imChgMode = False
                Exit Sub
            End If
            pbcRateCard_Paint
        End If
        imChgMode = False
    End If
End Sub

Private Sub edcBdDropDown_GotFocus()
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus edcBdDropDown
End Sub

Private Sub edcBdDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcBdDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If tgSpf.sRUseCorpCal = "Y" Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcBdDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcBdDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If tgSpf.sRUseCorpCal = "Y" Then
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        gProcessArrowKey Shift, KeyCode, lbcBudget, imLbcArrowSetting
        edcBdDropDown.SelStart = 0
        edcBdDropDown.SelLength = Len(edcBdDropDown.Text)
    End If
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imRCBoxNo
        Case VEHINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
        Case DAYPARTINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcDPNameRow, imBSMode, slStr)
            If ilRet = 1 Then
                If lbcDPNameRow.ListCount > 0 Then
                    lbcDPNameRow.ListIndex = 0
                End If
            End If
        Case DOLLAR1INDEX To DOLLAR4INDEX
            'If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sGUsePropSys <> "Y") Then
            '    edcDropDown.Text = "0"
            'End If
    End Select
    imLbcArrowSetting = False
End Sub

Private Sub edcDropDown_DblClick()
    If imRCBoxNo = DAYPARTINDEX Then
        imDoubleClickName = True    'Double click event is followed by a mouse up event
                                    'Process the double click event in the mouse up event
                                    'to avoid the mouse up event being in next form
    End If
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case imRCBoxNo
        Case VEHINDEX
            If lbcVehicle.ListCount = 1 Then
                lbcVehicle.ListIndex = 0
                'If imTabDirection = -1 Then 'Right to left
                '    pbcSTab.SetFocus
                'Else
                '    pbcTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case DAYPARTINDEX
        Case DOLLAR1INDEX To DOLLAR4INDEX
            If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sGUsePropSys <> "Y") Then
                If imTabDirection = -1 Then 'Right to left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            End If
    End Select
    gCtrlGotFocus edcDropDown
End Sub

Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    Dim slComp As String
    Dim ilPos As Integer

    Select Case imRCBoxNo
        'Case DOLLARINDEX, PCTINVINDEX
        '    ilPos = InStr(edcDropDown.SelText, ".")
        '    If ilPos = 0 Then
        '        ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
        '        If ilPos > 0 Then
        '            If KeyAscii = KEYDECPOINT Then
        '                Beep
        '                KeyAscii = 0
        '                Exit Sub
        '            End If
        '        End If
        '    End If
        '    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
        '    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        '        Beep
        '        KeyAscii = 0
        '        Exit Sub
        '    End If
        '    slStr = edcDropDown.Text
        '    slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & Right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
        '    If imRCBoxNo = DOLLARINDEX Then
        '        slComp = "99.999"
        '    ElseIf imRCBoxNo = PCTINVINDEX Then
        '        slComp = "100.00"
        '    End If
        '    If gCompNumberStr(slStr, slComp) > 0 Then
        '        Beep
        '        KeyAscii = 0
        '        Exit Sub
        '    End If
        Case SORTINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            slComp = "99"
            If gCompNumberStr(slStr, slComp) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case DOLLAR1INDEX To DOLLAR4INDEX
            If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sGUsePropSys <> "Y") Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            slComp = "9999999"
            If gCompNumberStr(slStr, slComp) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case CPMINDEX
            ilPos = InStr(edcDropDown.SelText, ".")
            If ilPos = 0 Then
                ilPos = InStr(edcDropDown, ".")    'Disallow multi-decimal points
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
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "9999999.99") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case Else
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
    End Select
End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imRCBoxNo
            Case VEHINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
            Case DAYPARTINDEX
                gProcessArrowKey Shift, KeyCode, lbcDPNameRow, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imRCBoxNo
            Case DAYPARTINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcSPDropDown_Change()
    Select Case imSPBoxNo
        'Case BUDGETINDEX
        '    imLbcArrowSetting = True
        '    gMatchLookAhead edcSPDropDown, lbcBudget, imBSMode, imComboBoxIndex
        Case GRIDINDEX
            If imChgMode = False Then
                imChgMode = True
                imLbcArrowSetting = True
                gMatchLookAhead edcSPDropDown, lbcGrid, imBSMode, imComboBoxIndex
                imGDSelectedIndex = lbcGrid.ListIndex
                mGetShowPrices -1
                pbcRateCard.Cls
                pbcRateCard_Paint
                mSetCommands
                imChgMode = False
            End If
        Case LENGTHINDEX
            If imChgMode = False Then
                imChgMode = True
                gMatchLookAhead edcSPDropDown, lbcLen, imBSMode, imComboBoxIndex
                imLenSelectedIndex = lbcLen.ListIndex
                mGetShowPrices -1
                pbcRateCard.Cls
                pbcRateCard_Paint
                imChgMode = False
            End If
        Case CURGRIDINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcSPDropDown, lbcGrid, imBSMode, imComboBoxIndex
            If lbcGrid.ListIndex <> tgRcfI.iTodayGrid Then
                igRcfChg = True
            End If
    End Select
    imLbcArrowSetting = False
End Sub

Private Sub edcSPDropDown_GotFocus()
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus edcSPDropDown
End Sub

Private Sub edcSPDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcSPDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSPDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcSPDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbCtrlMask) > 0) And (KeyCode >= KEYLEFT) And (KeyCode <= KeyDown) Then
        imDirProcess = KeyCode 'mDirection 0
        pbcSPTab.SetFocus
        Exit Sub
    End If
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
        Select Case imSPBoxNo
            'Case BUDGETINDEX
            '    gProcessArrowKey Shift, KeyCode, lbcBudget, imLbcArrowSetting
            Case GRIDINDEX
                gProcessArrowKey Shift, KeyCode, lbcGrid, imLbcArrowSetting
            Case LENGTHINDEX
                gProcessArrowKey Shift, KeyCode, lbcLen, imLbcArrowSetting
            Case CURGRIDINDEX
                gProcessArrowKey Shift, KeyCode, lbcGrid, imLbcArrowSetting
        End Select
        edcSPDropDown.SelStart = 0
        edcSPDropDown.SelLength = Len(edcSPDropDown.Text)
    End If
End Sub

Private Sub Form_Activate()
    Me.KeyPreview = True
    If imInNew Then
        Exit Sub
    End If
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True  'To get Alt J and Alt L keys
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcRateCard.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        cmcRealloc.Enabled = False
    Else
        pbcRateCard.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
        'Disallow Reallocate at this time 4/2/02-  Mary
        'The button was made invisible prior to this request.  It still is invisible.
        'cmcRealloc.Enabled = True
        If tgSpf.sCAudPkg <> "Y" Then
            cmcRealloc.Enabled = False
        Else
            cmcRealloc.Enabled = True
        End If
    End If
    If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sGUsePropSys <> "Y") Then
        cmcImpact.Enabled = False
    End If
    gShowBranner imUpdateAllowed
    'DoEvents    'Process events so pending keys are not sent to this
                'form when keypreview turn on
    'If Not imTerminate Then
    '    RateCard.KeyPreview = True   'To get Alt J and Alt L keys
    'End If
    Me.ZOrder 0 'Send to front
    RateCard.Refresh
    pbcView_Paint
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    'RateCard.KeyPreview = False
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcSelect.Enabled) And ((imSPBoxNo > 0) Or (imRCBoxNo > 0)) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imSPBoxNo > 0 Then
            mSPEnableBox imSPBoxNo
        ElseIf imRCBoxNo > 0 Then
            mRCEnableBox imRCBoxNo
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
    lgPercentAdjW = 95
    lgPercentAdjH = 90
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100)
    End If
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    If Not igManUnload Then
        mSPSetShow imSPBoxNo    'Remove focus
        imSPBoxNo = -1
        mRCSetShow imRCBoxNo
        imRCBoxNo = -1
        pbcArrow.Visible = False
        lacRCFrame.Visible = False
        lacDPFrame.Visible = False
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            If imRCBoxNo <> -1 Then
                mRCEnableBox imRCBoxNo
            ElseIf imSPBoxNo <> -1 Then
                mSPEnableBox imSPBoxNo
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    btrExtClear hmPvf   'Clear any previous extend operation
    ilRet = btrClose(hmPvf)
    btrDestroy hmPvf
    btrExtClear hmSsf   'Clear any previous extend operation
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    btrExtClear hmBvf   'Clear any previous extend operation
    ilRet = btrClose(hmBvf)
    btrDestroy hmBvf
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    btrExtClear hmSmf   'Clear any previous extend operation
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    btrExtClear hmSdf   'Clear any previous extend operation
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    btrExtClear hmCff   'Clear any previous extend operation
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    btrExtClear hmClf   'Clear any previous extend operation
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    btrExtClear hmCHF   'Clear any previous extend operation
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    btrExtClear hmVsf   'Clear any previous extend operation
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    btrExtClear hmLtf   'Clear any previous extend operation
    ilRet = btrClose(hmLtf)
    btrDestroy hmLtf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    btrExtClear hmAnf   'Clear any previous extend operation
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    btrExtClear hmRdf   'Clear any previous extend operation
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    btrExtClear hmRif   'Clear any previous extend operation
    ilRet = btrClose(hmRif)
    btrDestroy hmRif
    btrExtClear hmRcf   'Clear any previous extend operation
    ilRet = btrClose(hmRcf)
    btrDestroy hmRcf
    
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
    
    Erase lmAutoDelRif

    Erase tmPvf
    Erase smRCShow
    Erase smRCSave
    Erase lmRCSave
    Erase imRCSave
    Erase smDPShow
    Erase tmRifRec
    Erase tmLkRifRec
    Erase tmTrashRifRec
    Erase tmUserVeh
    'Erase imVefCode
    'Erase tmDPBudgetInfo
    Erase smBdShow

    Erase tmRateCard
    Erase tmBudgetCode
    Erase imBdMnf
    Erase imBdYr
    Erase tmMRif
'    Erase tgMRdf

    Erase tgImpactRec
    Erase tgDollarRec

    Erase tmChfAdvtExt
    Erase tmTempRifRec
    Erase tmRCModelInfo
    Erase tmClfModel
    Erase tgTempRCUserVehicle
    
    igJobShowing(RATECARDSJOB) = False
    
    Set RateCard = Nothing   'Remove data segment
    End
End Sub

Private Sub imcKey_Click()
    pbcKey.Visible = Not pbcKey.Visible
End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = True
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'pbcKey.Visible = False
End Sub

Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim llLoop As Long
    Dim ilIndex As Integer
    Dim llUpperBound As Long
    Dim llRowNo As Long
    Dim slStr As String
    Dim ilVbcValue As Integer
    'Dim tlRif As RIF
    If (lmRCRowNo < 1) Then
        Exit Sub
    End If
    'If tmRifRec(lmRCRowNo).iStatus <> 0 Then
    '    Do
    '        ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmRifRec(lmRCRowNo).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    '        ilRet = btrDelete(hmRif)
    '    Loop While ilRet = BTRV_ERR_CONFLICT
    'End If
    llRowNo = lmRCRowNo
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    lmRCRowNo = 0
    gCtrlGotFocus ActiveControl
    llUpperBound = UBound(smRCSave, 2)
    If llRowNo = llUpperBound Then
        For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
            slStr = ""
            gSetShow pbcRateCard, slStr, tmRCCtrls(ilLoop)
            smRCShow(ilLoop, llRowNo) = tmRCCtrls(ilLoop).sShow
        Next ilLoop
        pbcRateCard_Paint
        mSetDefInSave   'Set defaults for extra row
    Else
        Screen.MousePointer = vbHourglass
        'This code was added because of an error in Contracts
        'The error in contracts was fixed so this code is not needed (3/22/01)
        'The error was related to missing vehicles from the rate card, and another user
        'alters vehicles causing the vehicle list to be refreshed without
        'readding missing vehicles from the rate card.
        'If (imSelectedIndex <> 0) And (tgRcfI.iCode <> 0) And (tmRifRec(llRowNo).iStatus = 1) Then
        '    gFindMatch Trim$(smRCSave(VEHINDEX, llRowNo)), 0, lbcVehicle
        '    ilIndex = gLastFound(lbcVehicle)
        '    If ilIndex >= 0 Then
        '        slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
        '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        '        ilVefCode = Val(slCode)
        '        ilExtLen = Len(llChfCode)  'Extract operation record size
        '        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        '        btrExtClear hmClf   'Clear any previous extend operation
        '        ilRet = btrGetFirst(hmClf, tmClf, imClfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        '        If ilRet <> BTRV_ERR_END_OF_FILE Then
        '            Call btrExtSetBounds(hmClf, llNoRec, -1, "UC") '"EG") 'Set extract limits (all records)
        '            ilOffset = gFieldOffset("Clf", "ClfVefCode")
        '            ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, ilVefCode, 2)
        '            ilOffset = gFieldOffset("Clf", "clfChfCode")
        '            ilRet = btrExtAddField(hmClf, ilOffset, ilExtLen)'Extract the whole record
        '            ilRet = btrExtGetNext(hmClf, llChfCode, ilExtLen, llRecPos)
        '            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        '                ilExtLen = Len(llChfCode)  'Extract operation record size
        '                Do While ilRet = BTRV_ERR_REJECT_COUNT
        '                    ilRet = btrExtGetNext(hmClf, llChfCode, ilExtLen, llRecPos)
        '                Loop
        '                Do While ilRet = BTRV_ERR_NONE
        '                    If tmChfSrchKey.lCode <> llChfCode Then
        '                        tmChfSrchKey.lCode = llChfCode
        '                        ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        '                        If (ilRet = BTRV_ERR_NONE) And (tmChf.iRcfCode = tgRcfI.iCode) Then
        '                            Screen.MousePointer = vbDefault
        '                            slMsg = "Cannot erase - a Contract references vehicle, as alternative set vehicle to Dormant"
        '                            ilRet = MsgBox(slMsg, vbOkOnly + vbExclamation, "Erase")
        '                            Exit Sub
        '                        End If
        '                    End If
        '                    ilRet = btrExtGetNext(hmClf, llChfCode, ilExtLen, llRecPos)
        '                    Do While ilRet = BTRV_ERR_REJECT_COUNT
        '                        ilRet = btrExtGetNext(hmClf, llChfCode, ilExtLen, llRecPos)
        '                    Loop
        '                Loop
        '            End If
        '        End If
        '    End If
        'End If
        For llLoop = llRowNo To llUpperBound - 1 Step 1
            If (tmRifRec(llLoop).iStatus = 1) And (llLoop = llRowNo) Then
                'The link by year records can stay in tmLkRifRec
                tmTrashRifRec(UBound(tmTrashRifRec)) = tmRifRec(llLoop)
                ReDim Preserve tmTrashRifRec(0 To UBound(tmTrashRifRec) + 1) As RIFREC
            End If
            tmRifRec(llLoop) = tmRifRec(llLoop + 1)
            For ilIndex = 1 To UBound(smRCSave, 1) Step 1
                smRCSave(ilIndex, llLoop) = smRCSave(ilIndex, llLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(lmRCSave, 1) Step 1
                lmRCSave(ilIndex, llLoop) = lmRCSave(ilIndex, llLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(imRCSave, 1) Step 1
                imRCSave(ilIndex, llLoop) = imRCSave(ilIndex, llLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smRCShow, 1) Step 1
                smRCShow(ilIndex, llLoop) = smRCShow(ilIndex, llLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smDPShow, 1) Step 1
                smDPShow(ilIndex, llLoop) = smDPShow(ilIndex, llLoop + 1)
            Next ilIndex
        Next llLoop
        llUpperBound = UBound(smRCSave, 2)
        ReDim Preserve smRCSave(0 To SORTINDEX, 0 To llUpperBound - 1) As String * 40
        ReDim Preserve lmRCSave(0 To TOTALINDEX - SORTINDEX, 0 To llUpperBound - 1) As Long
        ReDim Preserve imRCSave(0 To 17, 0 To llUpperBound - 1) As Integer
        ReDim Preserve smRCShow(0 To AVGINDEX, 0 To llUpperBound - 1) As String * 40
        ReDim Preserve smDPShow(0 To DPBASEINDEX, 0 To llUpperBound - 1) As String * 40 'Values shown in program area
        ReDim Preserve tmRifRec(0 To llUpperBound - 1) As RIFREC
        imSettingValue = True
        ilVbcValue = vbcRateCard.Value
        vbcRateCard.Min = LBONE 'LBound(tmRifRec)
        imSettingValue = True
        If UBound(tmRifRec) <= vbcRateCard.LargeChange Then ' + 1 Then
            vbcRateCard.Max = LBONE 'LBound(tmRifRec)
        Else
            vbcRateCard.Max = UBound(tmRifRec) - vbcRateCard.LargeChange
        End If
        imSettingValue = True
        'vbcRateCard.Value = vbcRateCard.Min
        If ilVbcValue <= vbcRateCard.Max Then
            vbcRateCard.Value = ilVbcValue
        Else
            vbcRateCard.Value = vbcRateCard.Max
        End If
        imSettingValue = False
        Screen.MousePointer = vbDefault
    End If
    imRifChg = True
    mSetCommands
    If imView = 0 Then
        lacRCFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    ElseIf imView = 1 Then
        lacDPFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    End If
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcRateCard.Cls
    pbcRateCard_Paint
End Sub

Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
'    lacRCFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash_Click
End Sub

Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        If imView = 0 Then
            lacRCFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        ElseIf imView = 1 Then
            lacDPFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        End If
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        If imView = 0 Then
            lacRCFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        ElseIf imView = 1 Then
            lacDPFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub

Private Sub imcTrash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbcClickFocus.SetFocus
End Sub

Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub lbcBudget_Click()
    gProcessLbcClick lbcBudget, edcBdDropDown, imChgMode, imLbcArrowSetting
    If imBSelectedIndex <> lbcBudget.ListIndex Then
        edcBdDropDown_Change
    End If
End Sub

Private Sub lbcBudget_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcBudget2_Click()
    Dim ilRet As Integer
    'pbcRateCard.Cls
    ilRet = mBdBuildCompBudget()
    If Not ilRet Then
        Exit Sub
    End If
    pbcRateCard_Paint
End Sub

Private Sub lbcDPName_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcDPNameRow_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcDPNameRow, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcDPNameRow_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub lbcDPNameRow_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcDPNameRow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcDPNameRow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcDPNameRow, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcGrid_Click()
    gProcessLbcClick lbcGrid, edcSPDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcLen_Click()
    gProcessLbcClick lbcLen, edcSPDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcLen_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBdBuildCompBudget              *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Comparison dollars         *
'*                                                     *
'*                                                     *
'*******************************************************
Private Function mBdBuildCompBudget() As Integer
    Dim ilRet As Integer
    Dim llRif As Long
    Dim slDate As String
    Dim slEnd As String
    Dim slStart As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slYear As String
    Dim slNameYear As String
    Dim ilType As Integer
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slName As String
    If tgSpf.sRUseCorpCal <> "Y" Then
        If imBSelectedIndex < 0 Then
            ReDim tgImpactRec(0 To 1) As IMPACTREC
            ReDim tgDollarRec(0 To 1, 0 To 1) As DOLLARREC
            mBdBuildCompBudget = False
            Exit Function
        End If
    Else
        If lbcBudget2.SelCount < 1 Then
            ReDim tgImpactRec(0 To 1) As IMPACTREC
            ReDim tgDollarRec(0 To 1, 0 To 1) As DOLLARREC
            mBdBuildCompBudget = False
            Exit Function
        End If
    End If
    ilRet = MsgBox("The Comparison generation will take some time, Proceed", vbYesNo + vbQuestion, "Rate Card")
    If ilRet = vbNo Then
        ReDim tgImpactRec(0 To 1) As IMPACTREC
        ReDim tgDollarRec(0 To 1, 0 To 1) As DOLLARREC
        mBdBuildCompBudget = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    pbcRateCard.Cls
    If tgSpf.sRUseCorpCal <> "Y" Then
        lbcBudget.Visible = False
        slNameCode = tmBudgetCode(imBSelectedIndex).sKey  'lbcBudget.List(ilIndex - 1)
        ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
        ilRet = gParseItem(slNameYear, 2, "/", smBdMnfName)
        ilRet = gParseItem(slNameYear, 1, "/", slYear)
        slYear = gSubStr("9999", slYear)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imBdMnfCode = Val(slCode)
        imBdYear = Val(slYear)
        ReDim imBdMnf(0 To 0) As Integer
        ReDim imBdYr(0 To 0) As Integer
        imBdMnf(0) = imBdMnfCode
        imBdYr(0) = imBdYear
    Else
        lbcBudget2.Visible = False
        edcBdDropDown.Visible = True
        cmcBdDropDown.Visible = True
        If lbcBudget2.SelCount = 1 Then
            ReDim imBdMnf(0 To 0) As Integer
            ReDim imBdYr(0 To 0) As Integer
        Else
            ReDim imBdMnf(0 To 1) As Integer
            ReDim imBdYr(0 To 1) As Integer
        End If
        ilCount = 0
        slName = ""
        For ilLoop = 0 To lbcBudget2.ListCount - 1 Step 1
            If lbcBudget2.Selected(ilLoop) Then
                slNameCode = tmBudgetCode(ilLoop).sKey  'lbcBudget.List(ilIndex - 1)
                ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
                ilRet = gParseItem(slNameYear, 2, "/", smBdMnfName)
                ilRet = gParseItem(slNameYear, 1, "/", slYear)
                slYear = gSubStr("9999", slYear)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilCount = 0 Then
                    slName = lbcBudget2.List(ilLoop)
                    imBdMnf(0) = Val(slCode)
                    imBdYr(0) = Val(slYear)
                    ilCount = 1
                Else
                    slName = slName & "; " & lbcBudget2.List(ilLoop)
                    If imBdYr(0) < Val(slYear) Then
                        imBdMnf(1) = Val(slCode)
                        imBdYr(1) = Val(slYear)
                    Else
                        imBdMnf(1) = imBdMnf(0)
                        imBdYr(1) = imBdYr(0)
                        imBdMnf(0) = Val(slCode)
                        imBdYr(0) = Val(slYear)
                    End If
                    Exit For
                End If
            End If
        Next ilLoop
        edcBdDropDown.Text = slName
        slYear = Trim$(Str$(imBdYr(0)))
    End If
    slYear = Trim$(Str$(tgRcfI.iYear))
    If rbcShow(0).Value Then
        ilType = 4
        slDate = "1/15/" & slYear
        slDate = gObtainStartCorp(slDate, True)
        lmBdStartDate = gDateValue(slDate)
        'slDate = "12/15/" & slYear
        'slDate = gObtainEndCorp(slDate, True)
        'lmBdEndDate = gDateValue(slDate)
        slStart = slDate
        For ilLoop = 1 To 12 Step 1
            slEnd = gObtainEndCorp(slStart, True)
            slStart = gIncOneDay(slEnd)
        Next ilLoop
        lmBdEndDate = gDateValue(slEnd)
    Else
        ilType = 0
        slDate = "1/15/" & slYear
        slDate = gObtainStartStd(slDate)
        lmBdStartDate = gDateValue(slDate)
        slDate = "12/15/" & slYear
        slDate = gObtainEndStd(slDate)
        lmBdEndDate = gDateValue(slDate)
    End If
    'imBdNoWks = (lmBdEndDate - lmBdStartDate) / 7 + 1
    'ReDim tgImpactRec(1 To 1) As IMPACTREC
    'ReDim tgDollarRec(1 To imBdNoWks, 1 To 1) As DOLLARREC
    'If Not mReadBvfRec(hmBvf, imBdMnfCode, imBdYear, tmBvfVeh()) Then
    '    mBdBuildCompBudget = False
    '    Exit Function
    'End If
    ReDim tmMRif(0 To 0) As RIF
    For llRif = LBONE To UBound(tmRifRec) - 1 Step 1
        If (tmRifRec(llRif).iStatus = 0) Or (tmRifRec(llRif).iStatus = 1) Then
            tmMRif(UBound(tmMRif)) = tmRifRec(llRif).tRif
            ReDim Preserve tmMRif(0 To UBound(tmMRif) + 1) As RIF
        End If
    Next llRif
    If tgSpf.sRUseCorpCal <> "Y" Then
        mBdGetBudgetDollars 0, hmCHF, hmClf, hmCff, hmSdf, hmSmf, hmVef, hmVsf, hmSsf, hmBvf, hmLcf, imBdMnf(), imBdYr(), lmBdStartDate, lmBdEndDate, tmMRif(), tgMRdf()
    Else
        mBdGetBudgetDollars 2, hmCHF, hmClf, hmCff, hmSdf, hmSmf, hmVef, hmVsf, hmSsf, hmBvf, hmLcf, imBdMnf(), imBdYr(), lmBdStartDate, lmBdEndDate, tmMRif(), tgMRdf()
    End If
    mBdGetShowPrices True
    Screen.MousePointer = vbDefault
    mBdBuildCompBudget = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetBdShowPrices                *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate show dates           *
'*                                                     *
'*******************************************************
Private Sub mBdGetShowPrices(ilSetVbc As Integer)
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim ilStartWk As Integer
    Dim ilEndWk As Integer
    Dim slStr As String
    Dim ilVefDp As Integer
    Dim llCBudget As Long   'Current budget
    Dim llDollarSold As Long      'Sold
    Dim llAvail As Long
    Dim llBudget1 As Long
    Dim ilWksUndefined As Integer
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim slNameYear As String
    Dim slColor As String
    ReDim smBdShow(0 To AVGINDEX, 0 To 3 * (UBound(tgImpactRec) - 1) + 1) As String * 40
    For ilLoop = LBound(smBdShow, 1) To UBound(smBdShow, 1) Step 1
        For ilIndex = LBound(smBdShow, 2) To UBound(smBdShow, 2) Step 1
            smBdShow(ilLoop, ilIndex) = ""
        Next ilIndex
    Next ilLoop
    For ilVefDp = LBONE To UBound(tgImpactRec) - 1 Step 1
        For ilLoop = LBONE To AVGINDEX Step 1
            smBdShow(ilLoop, 3 * (ilVefDp - 1) + 1) = Trim$(smRCShow(ilLoop, tgImpactRec(ilVefDp).lPtRifRec(0) + 1))
        Next ilLoop
        llColor = pbcRateCard.ForeColor
        slFontName = pbcRateCard.FontName
        flFontSize = pbcRateCard.FontSize
        pbcRateCard.ForeColor = BLUE
        pbcRateCard.FontBold = False
        pbcRateCard.FontSize = 7
        pbcRateCard.FontName = "Arial"
        pbcRateCard.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        slNameYear = "M" & "Price to Make Budget"   'lbcBudget.List(imBSelectedIndex)
        gSetShow pbcRateCard, slNameYear, tmRCCtrls(VEHINDEX)
        smBdShow(VEHINDEX, 3 * (ilVefDp - 1) + 2) = "  " & Mid$(tmRCCtrls(VEHINDEX).sShow, 2)
        slStr = "MDifference"
        gSetShow pbcRateCard, slStr, tmRCCtrls(VEHINDEX)
        smBdShow(VEHINDEX, 3 * (ilVefDp - 1) + 3) = "  " & Mid$(tmRCCtrls(VEHINDEX).sShow, 2)
        pbcRateCard.FontSize = flFontSize
        pbcRateCard.FontName = slFontName
        pbcRateCard.FontSize = flFontSize
        pbcRateCard.ForeColor = llColor
        pbcRateCard.FontBold = True
        For ilGroup = LBONE To UBound(tmPdGroups) Step 1
            If tmPdGroups(ilGroup).sStartDate <> "" Then
                'ilStartWk = (gDateValue(tmPdGroups(ilGroup).sStartDate) - lmBdStartDate) \ 7 + 1
                'ilEndWk = (gDateValue(tmPdGroups(ilGroup).sEndDate) - lmBdStartDate) \ 7 + 1
                ilStartWk = tmPdGroups(ilGroup).iStartWkNo
                ilEndWk = ilStartWk + tmPdGroups(ilGroup).iNoWks - 1
                llCBudget = 0
                llDollarSold = 0
                llAvail = 0
                ilWksUndefined = 0
                For ilWk = ilStartWk To ilEndWk Step 1
                    If ilWk >= LBONE Then    'LBound(tgDollarRec, 1) Then
                        llAvail = llAvail + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).l30Inv - tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).l30Sold
                        If ((imBSelectedIndex >= 0) And (tgSpf.sRUseCorpCal <> "Y")) Or ((lbcBudget2.SelCount >= 1) And (tgSpf.sRUseCorpCal = "Y")) Then
                            llCBudget = llCBudget + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).lBudget   'lDollarSold
                            llDollarSold = llDollarSold + tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).lDollarSold
                        End If
                        If tgDollarRec(ilWk, tgImpactRec(ilVefDp).iPtDollarRec).iAvailDefined = 0 Then
                            ilWksUndefined = ilWksUndefined + 1
                        End If
                    Else
                        ilWksUndefined = ilWksUndefined + 1
                    End If
                Next ilWk
                If ((imBSelectedIndex >= 0) And (tgSpf.sRUseCorpCal <> "Y")) Or ((lbcBudget2.SelCount >= 1) And (tgSpf.sRUseCorpCal = "Y")) Then
                    If ilWksUndefined <> ilEndWk - ilStartWk + 1 Then
                        If llAvail > 0 Then
                            slColor = ""
                            llBudget1 = (llCBudget - llDollarSold) / llAvail
                            slStr = gLongToStrDec(llBudget1, 0)
                            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                            gSetShow pbcRateCard, slStr, tmRCCtrls(SORTINDEX + ilGroup)
                            smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 2) = tmRCCtrls(SORTINDEX + ilGroup).sShow
                            If llBudget1 >= 0 Then
                                If Len(Trim$(smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 1))) <> 0 Then
                                    If llBudget1 <= CLng(Trim$(smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 1))) Then
                                        slColor = "G"
                                    Else
                                        slColor = "R"
                                    End If
                                    llBudget1 = CLng(Trim$(smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 1))) - llBudget1
                                Else
                                    llBudget1 = 0 - llBudget1
                                    slColor = "R"
                                End If
                                slStr = gLongToStrDec(llBudget1, 0)
                                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                            Else
                                If llBudget1 <> 0 Then
                                    slColor = "G"
                                End If
                                slStr = ""
                            End If
                            gSetShow pbcRateCard, slStr, tmRCCtrls(SORTINDEX + ilGroup)
                            smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 3) = tmRCCtrls(SORTINDEX + ilGroup).sShow
                            smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 2) = slColor & Trim$(smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 2))
                        Else
                            gSetShow pbcRateCard, "Sold Out", tmRCCtrls(SORTINDEX + ilGroup)
                            smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 2) = tmRCCtrls(SORTINDEX + ilGroup).sShow
                            smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 3) = ""
                        End If
                    Else
                        gSetShow pbcRateCard, "Undefined", tmRCCtrls(SORTINDEX + ilGroup)
                        smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 2) = tmRCCtrls(SORTINDEX + ilGroup).sShow
                        smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 3) = ""
                    End If
                Else
                    smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 2) = ""
                    smBdShow(SORTINDEX + ilGroup, 3 * (ilVefDp - 1) + 3) = ""
                End If
            End If
        Next ilGroup
    Next ilVefDp
    If ilSetVbc Then
        imSettingValue = True
        vbcRateCard.Min = LBONE 'LBound(smBdShow, 2)
        imSettingValue = True
        If UBound(smBdShow, 2) - 1 <= vbcRateCard.LargeChange + 1 Then ' + 1 Then
            vbcRateCard.Max = LBONE 'LBound(smBdShow, 2)
        Else
            vbcRateCard.Max = UBound(smBdShow, 2) - vbcRateCard.LargeChange
        End If
        imSettingValue = True
        vbcRateCard.Value = vbcRateCard.Min
        imSettingValue = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBudgetPop                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the budget dropdown   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mBudgetPop()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    'ilRet = gPopVehBudgetBox(RateCard, 0, 1, lbcBudget, lbcBudgetCode)
    ilRet = gPopVehBudgetBox(RateCard, 2, 0, 1, lbcBudget, tmBudgetCode(), smBudgetCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mBudgetPopErr
        gCPErrorMsg ilRet, "mBudgetPop (gPopUserVehicleBox: Vehicle)", RateCard
        On Error GoTo 0
        lbcBudget2.Clear
        For ilLoop = 0 To lbcBudget.ListCount - 1 Step 1
            lbcBudget2.AddItem lbcBudget.List(ilLoop)
        Next ilLoop
        lbcBudget.Height = gListBoxHeight(lbcBudget.ListCount, 10)
        lbcBudget2.Height = gListBoxHeight(lbcBudget2.ListCount, 10)
    End If
    Exit Sub
mBudgetPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim slDate As String

    lbcVehicle.ListIndex = -1
    tgRcfI.iVefCode = -32000  'Force this field to be reset in mMoveCtrlToRec
    'tgRcfI.iStartDate(0) = 0
    'tgRcfI.iStartDate(1) = 0
    'tgRcfI.iEndDate(0) = 0
    'tgRcfI.iEndDate(1) = 0
    tgRcfI.iGridsUsed = 0
    tgRcfI.iBaseLen = 0
    tgRcfI.iTodayGrid = 0
    For ilLoop = LBound(tgRcfI.iLen) To UBound(tgRcfI.iLen) Step 1
        tgRcfI.iLen(ilLoop) = 0
    Next ilLoop
    For ilLoop = LBound(tgRcfI.iSpotMin) To UBound(tgRcfI.iSpotMin) Step 1
        tgRcfI.iSpotMin(ilLoop) = 0
        tgRcfI.iSpotMax(ilLoop) = 0
    Next ilLoop
    For ilLoop = LBound(tgRcfI.iWkMin) To UBound(tgRcfI.iWkMin) Step 1
        tgRcfI.iWkMin(ilLoop) = 0
        tgRcfI.iWkMax(ilLoop) = 0
    Next ilLoop
    For ilLoop = LBound(tgRcfI.iHrStartTime, 2) To UBound(tgRcfI.iHrStartTime, 2) Step 1
        tgRcfI.iHrStartTime(0, ilLoop) = 0
        tgRcfI.iHrStartTime(1, ilLoop) = 0
        tgRcfI.iHrEndTime(0, ilLoop) = 0
        tgRcfI.iHrEndTime(1, ilLoop) = 0
    Next ilLoop
    For ilLoop = LBound(tgRcfI.iFltNo) To UBound(tgRcfI.iFltNo) Step 1
        tgRcfI.iFltNo(ilLoop) = 0
    Next ilLoop
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    slDate = gObtainEndStd(slDate)
    gObtainMonthYear 0, slDate, ilMonth, ilYear
    tgRcfI.iYear = ilYear
    mInitRateCardCtrls
'   mMoveCtrlToRec
    For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
        tmRCCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCompMonths                     *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Months for Year        *
'*                                                     *
'*******************************************************
Private Sub mCompMonths(ilYear As Integer, ilStartWk() As Integer, ilNoWks() As Integer)
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    Dim ilLoop As Integer
    If rbcShow(0).Value Then    'Corporate
        slDate = "1/15/" & Trim$(Str$(ilYear))
        slStart = gObtainStartCorp(slDate, True)
        For ilLoop = 1 To 12 Step 1
            slEnd = gObtainEndCorp(slStart, True)
            ilNoWks(ilLoop) = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
            If ilLoop = 1 Then
                ilStartWk(1) = 1      'Set 1 0 below
            Else
                ilStartWk(ilLoop) = ilStartWk(ilLoop - 1) + ilNoWks(ilLoop - 1)
            End If
            slStart = gIncOneDay(slEnd)
        Next ilLoop
    Else                        'Standard
        'Compute start week number for each month
        slDate = "1/15/" & Trim$(Str$(ilYear))
        slStart = gObtainStartStd(slDate)
        For ilLoop = 1 To 12 Step 1
            slEnd = gObtainEndStd(slStart)
            ilNoWks(ilLoop) = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
            If ilLoop = 1 Then
                ilStartWk(1) = 1      'Set 1 0 below
            Else
                ilStartWk(ilLoop) = ilStartWk(ilLoop - 1) + ilNoWks(ilLoop - 1)
            End If
            slStart = gIncOneDay(slEnd)
        Next ilLoop
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mDirection                      *
'*                                                     *
'*             Created:9/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move to box indicated by       *
'*                      user direction                 *
'*                                                     *
'*******************************************************
Private Sub mDirection(ilMoveDir As Integer)
'
'   mDirection ilMove
'   Where:
'       ilMove (I)- 0=Up; 1= down; 2= left; 3= right
'
    mRCSetShow imRCBoxNo
    Select Case ilMoveDir
        Case KEYUP  'Up
            If lmRCRowNo > 1 Then
                lmRCRowNo = lmRCRowNo - 1
                If lmRCRowNo < vbcRateCard.Value + 1 Then
                    imSettingValue = True
                    vbcRateCard.Value = vbcRateCard.Value - 1
                End If
            Else
                lmRCRowNo = UBound(tmRifRec)
                imSettingValue = True
                If lmRCRowNo <= vbcRateCard.LargeChange + 1 Then
                    vbcRateCard.Value = 1
                Else
                    vbcRateCard.Value = lmRCRowNo - vbcRateCard.LargeChange - 1
                End If
            End If
        Case KeyDown  'Down
            If lmRCRowNo < UBound(tmRifRec) Then
                lmRCRowNo = lmRCRowNo + 1
                If lmRCRowNo > vbcRateCard.Value + vbcRateCard.LargeChange Then '+ 1 Then
                    imSettingValue = True
                    vbcRateCard.Value = vbcRateCard.Value + 1
                End If
            Else
                lmRCRowNo = 1
                imSettingValue = True
                vbcRateCard.Value = 1
            End If
        Case KEYLEFT  'Left
            If imRCBoxNo > VEHINDEX Then
                imRCBoxNo = imRCBoxNo - 1
            Else
                imRCBoxNo = DAYPARTINDEX
            End If
        Case KEYRIGHT  'Right
            If imRCBoxNo < DAYPARTINDEX Then
                imRCBoxNo = imRCBoxNo + 1
            Else
                imRCBoxNo = VEHINDEX
            End If
    End Select
    imSettingValue = False
    mRCEnableBox imRCBoxNo
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mDPBranch                       *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to DP     *
'*                      names and process communication*
'*                      back from daypart names        *
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
Private Function mDPBranch()
'
'   ilRet = mDPBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilUpdateAllowed As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    ilRet = gOptionalLookAhead(edcDropDown, lbcDPNameRow, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mDPBranch = False
        Exit Function
    End If
    'If Not gWinRoom(igNoLJWinRes(AVAILNAMESLIST)) Then
    '    mDPBranch = True
    '    mEnableBox imRCBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    igDPNameCallSource = CALLSOURCERATECARD
    If edcDropDown.Text = "[New]" Then
        sgDPName = ""
    Else
        sgDPName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    gFindMatch Trim$(smRCSave(VEHINDEX, lmRCRowNo)), 0, lbcVehicle
    ilIndex = gLastFound(lbcVehicle)
    If ilIndex < 0 Then
        mDPBranch = False
        Exit Function
    End If
    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
    ilRet = 0
    igVefCode = 0
    If ilIndex >= 0 And ilIndex < UBound(tgTempRCUserVehicle) Then
        slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        igVefCode = Val(Trim$(slCode))
    End If
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'edcLinkSrceDoneMsg.Text = ""
    'If igTestSystem Then
    '    slStr = "Daypart^Test\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
    'Else
    '    slStr = "Daypart^Prod\" & sgUserName & "\" & Trim$(Str$(igANmCallSource)) & "\" & sgANmName
    'End If
    'lgShellRet = Shell(sgExePath & "AName.Exe " & slStr, 1)
    imRetBranch = True
    'RateCard.Enabled = False
    Daypart.Show vbModal
    'Daypart.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'ilParse = gParseItem(slStr, 1, "\", sgANmName)
    'igANmCallSource = Val(sgANmName)
    'ilParse = gParseItem(slStr, 2, "\", sgANmName)
    'Daypart.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    imDoubleClickName = False
    mDPBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'RateCard.Enabled = True
    If igDPAltered Then
        Screen.MousePointer = vbHourglass
        mMoveCtrlToRec  'Save current values to record to retain changed fields
        'sgMRdfStamp = ""
        lbcDPName.Clear
        mDPNamePop
        mRemakeDpName
        Screen.MousePointer = vbDefault
    End If
    If igDPNameCallSource = CALLDONE Then  'Done
        igDPNameCallSource = CALLNONE
        mDPNameRowPop
        If imTerminate Then
            imRetBranch = False
            mDPBranch = False
            Exit Function
        End If
        'mDPNameRowPop
        gFindMatch sgDPName, 1, lbcDPNameRow
        sgDPName = ""
        If gLastFound(lbcDPNameRow) > 0 Then
            imChgMode = True
            lbcDPNameRow.ListIndex = gLastFound(lbcDPNameRow)
            edcDropDown.Text = lbcDPNameRow.List(lbcDPNameRow.ListIndex)
            imChgMode = False
            mDPBranch = False
        Else
            imChgMode = True
            lbcDPNameRow.ListIndex = 0
            edcDropDown.Text = lbcDPNameRow.List(lbcDPName.ListIndex)
            imChgMode = False
            edcDropDown.SetFocus
            pbcRateCard_Paint
            imRetBranch = False
            Exit Function
        End If
    End If
    pbcRateCard.Cls
    pbcRateCard_Paint
    If igDPNameCallSource = CALLCANCELLED Then  'Cancelled
        igDPNameCallSource = CALLNONE
        sgDPName = ""
        mRCEnableBox imRCBoxNo
        pbcRateCard_Paint
        imRetBranch = False
        Exit Function
    End If
    If igDPNameCallSource = CALLTERMINATED Then
        igDPNameCallSource = CALLNONE
        sgDPName = ""
        mRCEnableBox imRCBoxNo
        pbcRateCard_Paint
        imRetBranch = False
        Exit Function
    End If
    imRetBranch = False
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mDPNamePop                      *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mDPNamePop()
    'ReDim ilFilter(0) As Integer
    'ReDim slFilter(0) As String
    'ReDim ilOffset(0) As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim ilIndex As Integer
    Dim slMRdfStamp As String
    Dim ilLoop As Integer
    ilIndex = lbcDPName.ListIndex
    If ilIndex > 0 Then
        slName = lbcDPName.List(ilIndex)
    End If
    slMRdfStamp = sgMRdfStamp
    ilRet = gObtainRdf(sgMRdfStamp, tgMRdf())
    If (slMRdfStamp <> sgMRdfStamp) Or (lbcDPName.ListCount <= 0) Then
        lbcDPName.Clear
        lbcDPNameCode.Clear
        For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
            lbcDPNameCode.AddItem tgMRdf(ilLoop).sName & "\" & Trim(Str$(tgMRdf(ilLoop).iCode))
        Next ilLoop
        For ilLoop = 0 To lbcDPNameCode.ListCount - 1 Step 1
            slNameCode = lbcDPNameCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            lbcDPName.AddItem Trim$(slName)
        Next ilLoop
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 2, lbcDPName
            If gLastFound(lbcDPName) > 1 Then
                lbcDPName.ListIndex = gLastFound(lbcDPName)
            Else
                lbcDPName.ListIndex = -1
            End If
        Else
            lbcDPName.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    'ilFilter(0) = NOFILTER
    'slFilter(0) = ""
    'ilOffset(0) = 0
    'ilRet = gIMoveListBox(RateCard, lbcDPName, lbcDPNameCode, "Rdf.Btr", gFieldOffset("Rdf", "RdfName"), 20, ilFilter(), slFilter(), ilOffset())
    'If ilRet <> CP_MSG_NOPOPREQ Then
    '    On Error GoTo mDPNameErr
    '    gCPErrorMsg ilRet, "mDPNamePop (gIMoveListBox: DP Name)", RateCard
    '    On Error GoTo 0
    '    imChgMode = True
    '    If ilIndex > 0 Then
    '        gFindMatch slName, 2, lbcDPName
    '        If gLastFound(lbcDPName) > 1 Then
    '            lbcDPName.ListIndex = gLastFound(lbcDPName)
    '        Else
    '            lbcDPName.ListIndex = -1
    '        End If
    '    Else
    '        lbcDPName.ListIndex = ilIndex
    '    End If
    '    imChgMode = False
    'End If
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mDPNameRowPop                   *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mDPNameRowPop()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim slDPName As String
    Dim slName As String
    Dim ilRdf As Integer
    lbcDPNameRow.Clear
    gFindMatch Trim$(smRCSave(VEHINDEX, lmRCRowNo)), 0, lbcVehicle
    If gLastFound(lbcVehicle) >= 0 Then
        slName = lbcVehicle.List(gLastFound(lbcVehicle))
        'Populate only with names not used or with the one from the current row
        For ilLoop = 0 To lbcDPName.ListCount - 1 Step 1
            ilFound = False
            slDPName = lbcDPName.List(ilLoop)
            For ilIndex = LBONE To UBound(smRCSave, 2) - 1 Step 1
                If StrComp(slName, Trim$(smRCSave(VEHINDEX, ilIndex)), 1) = 0 Then
                    If StrComp(slDPName, Trim$(smRCSave(DAYPARTINDEX, ilIndex)), 1) = 0 Then
                        If ilIndex = lmRCRowNo Then
                            'Add here to get dormant names
                            gFindMatch slDPName, 0, lbcDPNameRow
                            If gLastFound(lbcDPNameRow) < 0 Then
                                lbcDPNameRow.AddItem slDPName
                            End If
                        End If
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilIndex
            If Not ilFound Then
                gFindMatch slDPName, 0, lbcDPNameRow
                If gLastFound(lbcDPNameRow) < 0 Then
                    'Test if dormant
                    ilFound = False
                    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        If StrComp(slDPName, Trim$(tgMRdf(ilRdf).sName), 1) = 0 Then
                            If tgMRdf(ilRdf).sState = "D" Then
                                ilFound = True
                            End If
                            Exit For
                        End If
                    Next ilRdf
                    If Not ilFound Then
                        lbcDPNameRow.AddItem slDPName
                    End If
                End If
            End If
        Next ilLoop
    End If
    lbcDPNameRow.AddItem "[New]", 0
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetAvg                         *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute average week dollars   *
'*                                                     *
'*******************************************************
Private Sub mGetAvg(ilYear As Integer, llRowNo As Long)
    Dim llDollar As Long
    Dim ilNoWks As Integer
    Dim llLkYear As Long
    Dim llRif As Long
    Dim slStr As String
    Dim ilWk As Integer
    Dim llERowNo As Long
    Dim vefMediumType As String
    llERowNo = UBound(tmRifRec) - 1
    If (Trim$(smRCSave(VEHINDEX, UBound(tmRifRec))) <> "") Then
        llERowNo = UBound(tmRifRec)
    End If
    For llRif = LBONE To llERowNo Step 1
        If (llRif = llRowNo) Or (llRowNo = -1) Then
            smRCShow(AVGINDEX, llRif) = ""
            If tmRifRec(llRif).tRif.iYear = ilYear Then
                llDollar = 0
                ilNoWks = 0
                If (rbcShow(1).Value) Then
                    llDollar = llDollar + tmRifRec(llRif).tRif.lRate(0)
                End If
                For ilWk = 1 To 53 Step 1
                    If tmRifRec(llRif).tRif.lRate(ilWk) > 0 Then
                        ilNoWks = ilNoWks + 1
                        llDollar = llDollar + tmRifRec(llRif).tRif.lRate(ilWk)
                    End If
                Next ilWk
                If ilNoWks > 0 Then
                    lmRCSave(AVGINDEX - SORTINDEX, llRif) = llDollar / ilNoWks
                Else
                    lmRCSave(AVGINDEX - SORTINDEX, llRif) = 0
                End If
                slStr = Trim$(Str$(lmRCSave(AVGINDEX - SORTINDEX, llRif)))
                'vefMediumType = mGetVehicleMediumType(tmRifRec(llRif).tRif.iVefCode)
                
                ' LB 02/10/21
                'If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER And imRCSave(17, llRif) = 1 Then
                '    slStr = 0
                'End If
                
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                gSetShow pbcRateCard, slStr, tmRCCtrls(AVGINDEX)
                smRCShow(AVGINDEX, llRif) = tmRCCtrls(AVGINDEX).sShow
            Else
                llLkYear = tmRifRec(llRif).lLkYear
                Do While llLkYear > 0
                    If tmLkRifRec(llLkYear).tRif.iYear = ilYear Then
                        llDollar = 0
                        ilNoWks = 0
                        If (rbcShow(1).Value) Then
                            llDollar = llDollar + tmLkRifRec(llLkYear).tRif.lRate(0)
                        End If
                        For ilWk = 1 To 53 Step 1
                            If tmLkRifRec(llLkYear).tRif.lRate(ilWk) > 0 Then
                                ilNoWks = ilNoWks + 1
                                llDollar = llDollar + tmLkRifRec(llLkYear).tRif.lRate(ilWk)
                            End If
                        Next ilWk
                        If ilNoWks > 0 Then
                            lmRCSave(AVGINDEX - SORTINDEX, llRif) = llDollar / ilNoWks
                        Else
                            lmRCSave(AVGINDEX - SORTINDEX, llRif) = 0
                        End If
                        slStr = Trim$(Str$(lmRCSave(AVGINDEX - SORTINDEX, llRif)))
                         'vefMediumType = mGetVehicleMediumType(tmRifRec(llRif).tRif.iVefCode)
                        ' LB 02/10/21
                        'If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER And imRCSave(17, llRif) = 1 Then
                        '    slStr = 0
                        'End If
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcRateCard, slStr, tmRCCtrls(AVGINDEX)
                        smRCShow(AVGINDEX, llRif) = tmRCCtrls(AVGINDEX).sShow
                        Exit Do
                    Else
                        llLkYear = tmLkRifRec(llLkYear).lLkYear
                    End If
                Loop
            End If
        End If
    Next llRif
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShowDates                   *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate show dates           *
'*                                                     *
'*******************************************************
Private Sub mGetShowDates()
'
'   mGetShowDates
'   Where:
'
    'Dim ilIndex As Integer
    'Dim ilLoop As Integer
    'Dim ilLen As Integer
    'Dim ilSpot As Integer
    'Dim ilWeek As Integer
    'Dim slStr As String
    'Dim slLeft As String
    'Dim slRight As String
    'Dim ilGrid As Integer
    'Dim ilPos As Integer
    'Dim ilStore As Integer
    'Dim slUnpacked As String
    'Dim slRound As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilLp1 As Integer
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    Dim slWkEnd As String
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilYearOk As Integer
    Dim ilWkNo As Integer
    Dim ilWkCount As Integer
    ReDim ilStartWk(0 To 12) As Integer 'Index zero ignored
    ReDim ilNoWks(0 To 12) As Integer
    Dim slFontName As String
    Dim flFontSize As Single
    If UBound(tmRifRec) <= 1 Then
        For ilIndex = 1 To 4 Step 1
            tmPdGroups(ilIndex).iStartWkNo = -1
            tmPdGroups(ilIndex).iNoWks = 0
            tmPdGroups(ilIndex).iTrueNoWks = 0
            tmPdGroups(ilIndex).iFltNo = 0
            tmPdGroups(ilIndex).sStartDate = ""
            tmPdGroups(ilIndex).sEndDate = ""
            gSetShow pbcRateCard, "", tmWKCtrls(ilIndex)
            gSetShow pbcRateCard, "", tmNWCtrls(ilIndex)
        Next ilIndex
        tmNWCtrls(5).sShow = "Average"
        If tgRcfI.iYear <= 0 Then
            Exit Sub
        End If
    End If
    slFontName = pbcRateCard.FontName
    flFontSize = pbcRateCard.FontSize
    pbcRateCard.FontBold = False
    pbcRateCard.FontSize = 7
    pbcRateCard.FontName = "Arial"
    pbcRateCard.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    gSetShow pbcRateCard, "Average", tmNWCtrls(5)
    If Not rbcType(3).Value Then        'Flight
        tmPdGroups(1).iYear = imPdYear
        tmPdGroups(1).iStartWkNo = imPdStartWk
        ilIndex = 1
        Do
            ilFound = False
            If ilIndex > 1 Then
                If tmPdGroups(ilIndex).iYear <> tmPdGroups(ilIndex - 1).iYear Then
                    ilYearOk = False
                Else
                    ilYearOk = True
                End If
            Else
                ilYearOk = False
            End If
            If Not ilYearOk Then
                mCompMonths tmPdGroups(ilIndex).iYear, ilStartWk(), ilNoWks()
            End If
            If rbcType(0).Value Then        'Quarter
                For ilLoop = 1 To 12 Step 1
                    If tmPdGroups(ilIndex).iStartWkNo = ilStartWk(ilLoop) Then
                        tmPdGroups(ilIndex).iNoWks = ilNoWks(ilLoop) + ilNoWks(ilLoop + 1) + ilNoWks(ilLoop + 2)
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
            ElseIf rbcType(1).Value Then    'Month
                For ilLoop = 1 To 12 Step 1
                    If tmPdGroups(ilIndex).iStartWkNo = ilStartWk(ilLoop) Then
                        tmPdGroups(ilIndex).iNoWks = ilNoWks(ilLoop)
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
            ElseIf rbcType(2).Value Then    'Week
                If tmPdGroups(ilIndex).iStartWkNo <= ilStartWk(12) + ilNoWks(12) - 1 Then
                    tmPdGroups(ilIndex).iNoWks = 1
                    ilFound = True
                End If
            End If
            If ilFound Then
                If ilIndex <> 4 Then
                    tmPdGroups(ilIndex + 1).iStartWkNo = tmPdGroups(ilIndex).iStartWkNo + tmPdGroups(ilIndex).iNoWks
                    tmPdGroups(ilIndex + 1).iYear = tmPdGroups(ilIndex).iYear   'imPdYear
                End If
                ilIndex = ilIndex + 1
            Else
                tmPdGroups(ilIndex).iYear = tmPdGroups(ilIndex).iYear + 1
                tmPdGroups(ilIndex).iStartWkNo = 1
                'Test if year exist
                If tmPdGroups(ilIndex).iYear > imRifStartYear + imRifNoYears - 1 Then
                    For ilLoop = ilIndex To 4 Step 1
                        tmPdGroups(ilLoop).iStartWkNo = -1
                        tmPdGroups(ilLoop).iTrueNoWks = 0
                        tmPdGroups(ilLoop).iNoWks = 0
                    Next ilLoop
                    Exit Do
                End If
            End If
        Loop Until ilIndex > 4
    Else
        tmPdGroups(1).iYear = imPdYear
        tmPdGroups(1).iFltNo = imPdStartFltNo
        ilIndex = 1
        Do
            ilFound = False
            For ilLoop = LBound(tgRcfI.iFltNo) + 1 To UBound(tgRcfI.iFltNo) Step 1
                If tmPdGroups(ilIndex).iFltNo = tgRcfI.iFltNo(ilLoop) Then
                    tmPdGroups(ilIndex).iStartWkNo = ilLoop
                    tmPdGroups(ilIndex).iNoWks = 1
                    For ilLp1 = ilLoop + 1 To UBound(tgRcfI.iFltNo) Step 1
                        If tmPdGroups(ilIndex).iFltNo = tgRcfI.iFltNo(ilLp1) Then
                            tmPdGroups(ilIndex).iNoWks = tmPdGroups(ilIndex).iNoWks + 1
                        Else
                            Exit For
                        End If
                    Next ilLp1
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If ilFound Then
                If ilIndex <> 4 Then
                    tmPdGroups(ilIndex + 1).iFltNo = tmPdGroups(ilIndex).iFltNo + 1
                    tmPdGroups(ilIndex + 1).iYear = tmPdGroups(ilIndex).iYear
                End If
                ilIndex = ilIndex + 1
            Else
                tmPdGroups(ilIndex).iYear = tmPdGroups(ilIndex).iYear + 1
                tmPdGroups(ilIndex).iFltNo = 1
                'Test if year exist
                If tmPdGroups(ilIndex).iYear > imRifStartYear + imRifNoYears - 1 Then
                    For ilLoop = ilIndex To 4 Step 1
                        tmPdGroups(ilLoop).iStartWkNo = -1
                        tmPdGroups(ilLoop).iTrueNoWks = 0
                        tmPdGroups(ilLoop).iNoWks = 0
                    Next ilLoop
                    Exit Do
                End If
            End If
        Loop Until ilIndex > 4
    End If
    'Compute Start/End Date if groups
    For ilIndex = 1 To 4 Step 1
        If tmPdGroups(ilIndex).iStartWkNo > 0 Then
            If rbcShow(0).Value Then    'Corporate
                slDate = "1/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                slStart = gObtainStartCorp(slDate, True)
                'slDate = "12/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                'slEnd = gObtainEndCorp(slDate, True)
                slDate = slStart
                For ilLoop = 1 To 12 Step 1
                    slEnd = gObtainEndCorp(slDate, True)
                    slDate = gIncOneDay(slEnd)
                Next ilLoop
            Else
                slDate = "1/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                slStart = gObtainStartStd(slDate)
                slDate = "12/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                slEnd = gObtainEndStd(slDate)
            End If
            ilWkNo = 1
            Do
                If ilWkNo = tmPdGroups(ilIndex).iStartWkNo Then
                    tmPdGroups(ilIndex).sStartDate = slStart
                    slWkEnd = gObtainNextSunday(slStart)
                    ilWkCount = 1
                    Do
                        If ilWkNo = tmPdGroups(ilIndex).iStartWkNo + tmPdGroups(ilIndex).iNoWks - 1 Then
                            tmPdGroups(ilIndex).sEndDate = slWkEnd
                            tmPdGroups(ilIndex).iTrueNoWks = ilWkCount
                            'slDate = tmPdGroups(ilIndex).sStartDate & "-" & tmPdGroups(ilIndex).sEndDate
                            slDate = tmPdGroups(ilIndex).sStartDate
                            slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                            gSetShow pbcRateCard, slDate, tmWKCtrls(ilIndex)
                            slStr = Trim$(Str$(ilWkCount))
                            slStr = "# Weeks " & slStr
                            gSetShow pbcRateCard, slStr, tmNWCtrls(ilIndex)
                            Exit Do
                        Else
                            ilWkNo = ilWkNo + 1
                            ilWkCount = ilWkCount + 1
                            slWkEnd = gIncOneWeek(slWkEnd)
                            If gDateValue(slWkEnd) > gDateValue(slEnd) Then
                                tmPdGroups(ilIndex).sEndDate = slEnd
                                tmPdGroups(ilIndex).iTrueNoWks = ilWkCount - 1
                                'slDate = Left$(tmPdGroups(ilIndex).sStartDate, Len(tmPdGroups(ilIndex).sStartDate) - 3)
                                'slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                                slDate = tmPdGroups(ilIndex).sStartDate
                                slDate = slDate & "-" & Left$(tmPdGroups(ilIndex).sEndDate, Len(tmPdGroups(ilIndex).sEndDate) - 3)
                                gSetShow pbcRateCard, slDate, tmWKCtrls(ilIndex)
                                slStr = Trim$(Str$(ilWkCount - 1))
                                slStr = "# Weeks " & slStr
                                gSetShow pbcRateCard, slStr, tmNWCtrls(ilIndex)
                                Exit Do
                            End If
                        End If
                    Loop
                    Exit Do
                Else
                    ilWkNo = ilWkNo + 1
                    slStart = gIncOneWeek(slStart)
                End If
            Loop
        Else
            tmPdGroups(ilIndex).sStartDate = ""
            tmPdGroups(ilIndex).sEndDate = ""
            gSetShow pbcRateCard, "", tmWKCtrls(ilIndex)
            gSetShow pbcRateCard, "", tmNWCtrls(ilIndex)
        End If
    Next ilIndex
    pbcRateCard.FontSize = flFontSize
    pbcRateCard.FontName = slFontName
    pbcRateCard.FontSize = flFontSize
    pbcRateCard.FontBold = True
    If imView <> 2 Then
        mGetShowPrices -1
    Else
        mGetShowPrices -1
        mBdGetShowPrices False
    End If
    'ilIndex = 0
    'ilLen = 1
    'gPDNToStr tgRcfI.sRound, 2, slRound
    'Do
    '    ilSpot = 1
    '    Do
    '        ilWeek = 1
    '        Do
    '            ilStore = 0
    '            If tgRcfI.sUseLen <> "N" Then
    '                For ilLoop = LBound(tgRcfI.sValue) To UBound(tgRcfI.sValue) Step 1
    '                    If tgRcfI.iLen(ilLoop) = tmRgf(ilIndex, imRowNo - 1).iLen Then
    '                        gPDNToStr tgRcfI.sValue(ilLoop), 2, slUnpacked
    '                        For ilGrid = 1 To tgRcfI.iGridsUsed Step 1
    '                            Select Case tgRcfI.sUseLen
    '                                Case "A"    'Actuals
    '                                Case "D"    '+,- dollars
    '                                    'If ilStore = 0 Then
    '                                    '    smGDSave(ilGrid, ilIndex + 1) = gAddStr(smReCalcSave(ilGrid), slUnpacked)
    '                                    'Else
    '                                    '    smGDSave(ilGrid, ilIndex + 1) = gAddStr(smGDSave(ilGrid, ilIndex + 1), slUnpacked)
    '                                    'End If
    '                                Case "P"    '+,- %
    '                                    slStr = gAddStr("100", slUnpacked)
    '                                    ilPos = InStr(slStr, ".")
    '                                    If ilPos = 0 Then
    '                                        slLeft = slStr
    '                                        slRight = "00"
    '                                    Else
    '                                        slLeft = Left$(slStr, ilPos - 1)
    '                                        slRight = Right$(slStr, Len(slStr) - ilPos)
    '                                    End If
    '                                    slRight = Right$(slLeft, 2) & slRight
    '                                    slLeft = Left$(slLeft, Len(slLeft) - 2)
    '                                    slStr = slLeft & "." & slRight
    '                                    'If ilStore = 0 Then
    '                                    '    smGDSave(ilGrid, ilIndex + 1) = gMulStr(smReCalcSave(ilGrid), slStr)
    '                                    'Else
    '                                    '    smGDSave(ilGrid, ilIndex + 1) = gMulStr(smGDSave(ilGrid, ilIndex + 1), slStr)
    '                                    'End If
    '                            End Select
    '                        Next ilGrid
    '                        ilStore = 1
    '                        Exit For
    '                    End If
    '                Next ilLoop
    '            End If
    '            If tgRcfI.sUseSpot <> "N" Then
    '                gPDNToStr tgRcfI.sSpotVal(ilSpot), 2, slUnpacked
    '                For ilGrid = 1 To tgRcfI.iGridsUsed Step 1
    '                    Select Case tgRcfI.sUseSpot
    '                        Case "A"    'Actuals
    '                        Case "D"    '+,- dollars
    '                            'If ilStore = 0 Then
    '                            '    smGDSave(ilGrid, ilIndex + 1) = gAddStr(smReCalcSave(ilGrid), slUnpacked)
    '                            'Else
    '                            '    smGDSave(ilGrid, ilIndex + 1) = gAddStr(smGDSave(ilGrid, ilIndex + 1), slUnpacked)
    '                            'End If
    '                        Case "P"    '+,- %
    '                            slStr = gAddStr("100", slUnpacked)
    '                            ilPos = InStr(slStr, ".")
    '                            If ilPos = 0 Then
    '                                slLeft = slStr
    '                                slRight = "00"
    '                            Else
    '                                slLeft = Left$(slStr, ilPos - 1)
    '                                slRight = Right$(slStr, Len(slStr) - ilPos)
    '                            End If
    '                            slRight = Right$(slLeft, 2) & slRight
    '                            slLeft = Left$(slLeft, Len(slLeft) - 2)
    '                            slStr = slLeft & "." & slRight
    '                            'If ilStore = 0 Then
    '                            '    smGDSave(ilGrid, ilIndex + 1) = gMulStr(smReCalcSave(ilGrid), slStr)
    '                            'Else
    '                            '    smGDSave(ilGrid, ilIndex + 1) = gMulStr(smGDSave(ilGrid, ilIndex + 1), slStr)
    '                            'End If
    '                        End Select
    '                Next ilGrid
    '                ilStore = 1
    '            End If
    '            If tgRcfI.sUseWeek <> "N" Then
    '                gPDNToStr tgRcfI.sWkVal(ilWeek), 2, slUnpacked
    '                For ilGrid = 1 To tgRcfI.iGridsUsed Step 1
    '                    Select Case tgRcfI.sUseWeek
    '                        Case "A"    'Actuals
    '                        Case "D"    '+,- dollars
    '                            'If ilStore = 0 Then
    '                            '    smGDSave(ilGrid, ilIndex + 1) = gAddStr(smReCalcSave(ilGrid), slUnpacked)
    '                            'Else
    '                            '    smGDSave(ilGrid, ilIndex + 1) = gAddStr(smGDSave(ilGrid, ilIndex + 1), slUnpacked)
    '                            'End If
    '                        Case "P"    '+,- %
    '                            slStr = gAddStr("100", slUnpacked)
    '                            ilPos = InStr(slStr, ".")
    '                            If ilPos = 0 Then
    '                                slLeft = slStr
    '                                slRight = "00"
    '                            Else
    '                                slLeft = Left$(slStr, ilPos - 1)
    '                                slRight = Right$(slStr, Len(slStr) - ilPos)
    '                            End If
    '                            slRight = Right$(slLeft, 2) & slRight
    '                            slLeft = Left$(slLeft, Len(slLeft) - 2)
    '                            slStr = slLeft & "." & slRight
    '                            'If ilStore = 0 Then
    '                            '    smGDSave(ilGrid, ilIndex + 1) = gMulStr(smReCalcSave(ilGrid), slStr)
    '                            'Else
    '                            '    smGDSave(ilGrid, ilIndex + 1) = gMulStr(smGDSave(ilGrid, ilIndex + 1), slStr)
    '                            'End If
    '                        End Select
    '                Next ilGrid
    '                ilStore = 1
    '            End If
    '            ilIndex = ilIndex + 1
    '            ilWeek = ilWeek + 1
    '        Loop While ilWeek <= imWeek
    '        ilSpot = ilSpot + 1
    '    Loop While ilSpot <= imSpot
    '    ilLen = ilLen + 1
    'Loop While ilLen <= imLen
    'For ilIndex = LBound(smGDSave, 2) To UBound(smGDSave, 2) Step 1
    '    For ilLoop = 1 To tgRcfI.iGridsUsed Step 1
    '        'smGDSave(ilLoop, ilIndex) = gAddStr(smGDSave(ilLoop, ilIndex), ".005")
    '        'smGDSave(ilLoop, ilIndex) = gRoundStr(smGDSave(ilLoop, ilIndex), slRound, 2)
    '        'slStr = smGDSave(ilLoop, ilIndex)
    '        'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
    '        'gSetShow pbcGdArea, slStr, tmGdCtrls(2)
    '        'smGDShow(ilLoop + imNoTitles, ilIndex + 1) = tmGdCtrls(2).sShow
    '    Next ilLoop
    'Next ilIndex
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShowPrices                  *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Calculate show dates           *
'*                                                     *
'*******************************************************
Private Sub mGetShowPrices(llRowNo As Long)
    Dim ilLoop As Integer
    Dim llRif As Long
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim llLkYear As Long
    Dim llRate As Long
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    'Dim slNoWks As String
    Dim slStr As String
    Dim llSRowNo As Long
    Dim llERowNo As Long
    Dim llLoop As Long
    Dim vefMediumType As String
    
    If llRowNo < 0 Then
        llSRowNo = LBONE    'LBound(lmRCSave, 2)
        llERowNo = UBound(lmRCSave, 2) - 1
        If (Trim$(smRCSave(VEHINDEX, UBound(tmRifRec))) <> "") Then
            llERowNo = UBound(tmRifRec)
        End If
    Else
        llSRowNo = llRowNo
        llERowNo = llRowNo
    End If
    For llLoop = llSRowNo To llERowNo Step 1
        lmRCSave(DOLLAR1INDEX - DOLLAR1INDEX + 1, llLoop) = 0
        lmRCSave(DOLLAR2INDEX - DOLLAR1INDEX + 1, llLoop) = 0
        lmRCSave(DOLLAR3INDEX - DOLLAR1INDEX + 1, llLoop) = 0
        lmRCSave(DOLLAR4INDEX - DOLLAR1INDEX + 1, llLoop) = 0
        lmRCSave(AVGINDEX - DOLLAR1INDEX + 1, llLoop) = 0    'Average for year
        lmRCSave(TOTALINDEX - DOLLAR1INDEX + 1, llLoop) = 0    'Average for year
        For ilWk = 1 To 8 Step 1
            imRCSave(ilWk, llLoop) = 0
        Next ilWk
        For ilWk = 12 To 15 Step 1
            imRCSave(ilWk, llLoop) = 0
        Next ilWk
    Next llLoop
    'If not using proposal system, then bypass the computation
    If tgSpf.sGUsePropSys <> "Y" Then
        Exit Sub
    End If
    'Sum value, then avaerage
    If llRowNo < 0 Then
        llSRowNo = LBONE    'LBound(tmRifRec)
        llERowNo = UBound(tmRifRec) - 1
        If (Trim$(smRCSave(VEHINDEX, UBound(tmRifRec))) <> "") Then
            llERowNo = UBound(tmRifRec)
        End If
    Else
        llSRowNo = llRowNo
        llERowNo = llRowNo
    End If
    For ilGroup = LBONE To UBound(tmPdGroups) Step 1
        For llRif = llSRowNo To llERowNo Step 1
            If tmRifRec(llRif).tRif.iYear = tmPdGroups(ilGroup).iYear Then
                If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                    ''Add in the first part of the standard week
                    'If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                    '    'gPDNToStr tmRifRec(llRif).tRif.sRate(0), 2, slDollar
                    '    'smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif) = gAddStr(smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif), slDollar)
                    '    lmRCSave(ilGroup, ilRif) = lmRCSave(ilGroup, ilRif) + tmRifRec(ilRif).tRif.lRate(0)
                    '    If (tmRifRec(ilRif).tRif.lRate(0) > 0) And (tmRifRec(ilRif).tRif.lRate(1) = 0) Then
                    '        imRCSave(ilGroup, ilRif) = imRCSave(ilGroup, ilRif) + 1
                    '    End If
                    'End If
                    For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                        'gPDNToStr tmRifRec(ilRif).tRif.sRate(ilWk), 2, slDollar
                        'smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif) = gAddStr(smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif), slDollar)
                        slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                        If rbcShow(0).Value Then
                            mRifGetRate llRif, slStart, tmRifRec(), tmLkRifRec(), llRate
                        Else
                            llRate = tmRifRec(llRif).tRif.lRate(ilWk)
                        End If
                        lmRCSave(ilGroup, llRif) = lmRCSave(ilGroup, llRif) + llRate 'tmRifRec(llRif).tRif.lRate(ilWk)
                        imRCSave(ilGroup + 4, llRif) = imRCSave(ilGroup + 4, llRif) + 1
                        If llRate > 0 Then      'tmRifRec(ilRif).tRif.lRate(ilWk) > 0 Then
                            imRCSave(ilGroup, llRif) = imRCSave(ilGroup, llRif) + 1
                        Else
                            imRCSave(ilGroup + 11, llRif) = 1
                        End If
                    Next ilWk
                    'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 = 52) And (rbcShow(0).Value) Then
                    '    'gPDNToStr tmRifRec(ilRif).tRif.sRate(0), 2, slDollar
                    '    'smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif) = gAddStr(smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif), slDollar)
                    '    lmRCSave(ilGroup, ilRif) = lmRCSave(ilGroup, ilRif) + tmRifRec(ilRif).tRif.lRate(0)
                    'End If
                    If ilGroup = LBONE Then 'LBound(tmPdGroups) Then    'Average Dollar
                        'If (rbcShow(1).Value) Then
                        '    'gPDNToStr tmRifRec(ilRif).tRif.sRate(0), 2, slDollar
                        '    'smRCSave(AVGINDEX, ilRif) = gAddStr(smRCSave(AVGINDEX, ilRif), slDollar)
                        '    lmRCSave(TOTALINDEX - SORTINDEX, ilRif) = lmRCSave(TOTALINDEX - SORTINDEX, ilRif) + tmRifRec(ilRif).tRif.lRate(0)
                        'End If
                        For ilWk = 1 To 53 Step 1
                            'gPDNToStr tmRifRec(ilRif).tRif.sRate(ilWk), 2, slDollar
                            'smRCSave(AVGINDEX, ilRif) = gAddStr(smRCSave(AVGINDEX, ilRif), slDollar)
                            lmRCSave(TOTALINDEX - SORTINDEX, llRif) = lmRCSave(TOTALINDEX - SORTINDEX, llRif) + tmRifRec(llRif).tRif.lRate(ilWk)
                        Next ilWk
                    End If
                End If
            Else
                llLkYear = tmRifRec(llRif).lLkYear
                Do While llLkYear > 0
                    If tmLkRifRec(llLkYear).tRif.iYear = tmPdGroups(ilGroup).iYear Then
                        If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                            ''Add in the first part of the standard week
                            'If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                            '    'gPDNToStr tmLkRifRec(llLkYear).tRif.sRate(0), 2, slDollar
                            '    'smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif) = gAddStr(smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif), slDollar)
                            '    lmRCSave(ilGroup, ilRif) = lmRCSave(ilGroup, ilRif) + tmLkRifRec(llLkYear).tRif.lRate(0)
                            '    If (tmLkRifRec(llLkYear).tRif.lRate(0) > 0) And (tmLkRifRec(llLkYear).tRif.lRate(1) = 0) Then
                            '        imRCSave(ilGroup, ilRif) = imRCSave(ilGroup, ilRif) + 1
                            '    End If
                            'End If
                            For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                                'gPDNToStr tmLkRifRec(llLkYear).tRif.sRate(ilWk), 2, slDollar
                                'smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif) = gAddStr(smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif), slDollar)
                                slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                                If rbcShow(0).Value Then
                                    mRifGetRate llRif, slStart, tmRifRec(), tmLkRifRec(), llRate
                                Else
                                    llRate = tmLkRifRec(llLkYear).tRif.lRate(ilWk)
                                End If
                                lmRCSave(ilGroup, llRif) = lmRCSave(ilGroup, llRif) + llRate 'tmLkRifRec(llLkYear).tRif.lRate(ilWk)
                                imRCSave(ilGroup + 4, llRif) = imRCSave(ilGroup + 4, llRif) + 1
                                If llRate > 0 Then      'tmLkRifRec(llLkYear).tRif.lRate(ilWk) > 0 Then
                                    imRCSave(ilGroup, llRif) = imRCSave(ilGroup, llRif) + 1
                                Else
                                    imRCSave(ilGroup + 11, llRif) = 1
                                End If
                            Next ilWk
                            'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 = 52) And (rbcShow(0).Value) Then
                            '    'gPDNToStr tmLkRifRec(llLkYear).tRif.sRate(0), 2, slDollar
                            '    'smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif) = gAddStr(smRCSave(ilGroup + DOLLAR1INDEX - 1, ilRif), slDollar)
                            '    lmRCSave(ilGroup, ilRif) = lmRCSave(ilGroup, ilRif) + tmLkRifRec(llLkYear).tRif.lRate(0)
                            'End If
                            If ilGroup = LBONE Then 'LBound(tmPdGroups) Then    'Average Dollar
                                'If (rbcShow(1).Value) Then
                                '    'gPDNToStr tmLkRifRec(llLkYear).tRif.sRate(0), 2, slDollar
                                '    'smRCSave(AVGINDEX, ilRif) = gAddStr(smRCSave(AVGINDEX, ilRif), slDollar)
                                '    lmRCSave(TOTALINDEX - SORTINDEX, ilRif) = lmRCSave(TOTALINDEX - SORTINDEX, ilRif) + tmLkRifRec(llLkYear).tRif.lRate(0)
                                'End If
                                For ilWk = 1 To 53 Step 1
                                    'gPDNToStr tmLkRifRec(llLkYear).tRif.sRate(ilWk), 2, slDollar
                                    'smRCSave(AVGINDEX, ilRif) = gAddStr(smRCSave(AVGINDEX, ilRif), slDollar)
                                    lmRCSave(TOTALINDEX - SORTINDEX, llRif) = lmRCSave(TOTALINDEX - SORTINDEX, llRif) + tmLkRifRec(llLkYear).tRif.lRate(ilWk)
                                Next ilWk
                            End If
                        End If
                        Exit Do
                    Else
                        llLkYear = tmLkRifRec(llLkYear).lLkYear
                    End If
                Loop
            End If
        Next llRif
    Next ilGroup
    'Average value
    
    For ilGroup = LBONE To UBound(tmPdGroups) Step 1
        
        For llRif = llSRowNo To llERowNo Step 1
            If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                'lmRCSave(ilGroup, llRif) = lmRCSave(ilGroup, llRif) / tmPdGroups(ilGroup).iTrueNoWks
                If imRCSave(ilGroup, llRif) > 0 Then
                    lmRCSave(ilGroup, llRif) = lmRCSave(ilGroup, llRif) / imRCSave(ilGroup, llRif)
                Else
                    lmRCSave(ilGroup, llRif) = 0
                End If
                slStr = Trim$(Str$(lmRCSave(ilGroup, llRif)))
                'L.Bianchi
                'vefMediumType = mGetVehicleMediumType(tmRifRec(llRif).tRif.iVefCode)
                ' LB 02/10/21
                'If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER And imRCSave(17, llRif) = 1 Then
                '    slStr = 0
                'End If
                
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                gSetShow pbcRateCard, slStr, tmRCCtrls(DOLLAR1INDEX)
                smRCShow(ilGroup + DOLLAR1INDEX - 1, llRif) = tmRCCtrls(DOLLAR1INDEX).sShow
            Else
                smRCShow(ilGroup + DOLLAR1INDEX - 1, llRif) = ""
            End If
        Next llRif
    Next ilGroup
    'Compute number of weeks for year
    If rbcShow(0).Value Then    'Corporate
        slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
        slStart = gObtainStartCorp(slDate, True)
        'slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
        'slEnd = gObtainEndCorp(slDate, True)
        slDate = slStart
        For ilLoop = 1 To 12 Step 1
            slEnd = gObtainEndCorp(slDate, True)
            slDate = gIncOneDay(slEnd)
        Next ilLoop
    Else
        slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
        slStart = gObtainStartStd(slDate)
        slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
        slEnd = gObtainEndStd(slDate)
    End If
    'slNoWks = Trim$(Str$((gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7))
    'ilNoWks = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7

    'For ilRif = LBound(tmRifRec) To UBound(tmRifRec) - 1 Step 1
    '    'ilNoWks = imRCSave(1, ilRif) + imRCSave(2, ilRif) + imRCSave(3, ilRif) + imRCSave(4, ilRif)
    '
    '    If ilNoWks > 0 Then
    '        lmRCSave(AVGINDEX - SORTINDEX, ilRif) = lmRCSave(TOTALINDEX - SORTINDEX, ilRif) / ilNoWks
    '    Else
    '        lmRCSave(AVGINDEX - SORTINDEX, ilRif) = 0
    '    End If
    '    slStr = Trim$(Str$(lmRCSave(AVGINDEX - SORTINDEX, ilRif)))
    '    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
    '    gSetShow pbcRateCard, slStr, tmRCCtrls(AVGINDEX)
    '    smRCShow(AVGINDEX, ilRif) = tmRCCtrls(AVGINDEX).sShow
    'Next ilRif
    mGetAvg tmPdGroups(LBONE).iYear, -1
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetStdPkgPrice                 *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Standard Package price     *
'*                                                     *
'*******************************************************
Private Sub mGetStdPkgPrice(ilAskQuestion As Boolean)
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilLoop As Integer
    Dim ilWk As Integer
    Dim ilPvf As Integer
    Dim llRif As Long
    Dim ilVef As Integer
    Dim llPvfCode As Long
    Dim ilVefCode As Integer
    Dim ilRdfCode As Integer
    Dim llLkYear As Long
    Dim ilFound As Integer
    gFindMatch Trim$(smRCSave(VEHINDEX, lmRCRowNo)), 0, lbcVehicle
    ilIndex = gLastFound(lbcVehicle)
    If ilIndex < 0 Then
        Exit Sub
    End If
    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
    slNameCode = ""
    ilCode = 0
    If ilIndex >= 0 And ilIndex < UBound(tgTempRCUserVehicle) Then
        slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilCode = CInt(slCode)
    End If
    imRCSave(11, lmRCRowNo) = 0
    'For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
    '    If tgMVef(ilLoop).iCode = ilCode Then
        ilLoop = gBinarySearchVef(ilCode)
        If ilLoop <> -1 Then
            
            If ((tgMVef(ilLoop).sType <> "C") And (tgMVef(ilLoop).sType <> "S")) Or ((Asc(tgSpf.sUsingFeatures2) And BARTER) <> BARTER) Then
                smRCSave(ACQUISITIONINDEX, lmRCRowNo) = "Y"
                smRCShow(ACQUISITIONINDEX, lmRCRowNo) = ""
            End If
            
            If (tgMVef(ilLoop).sType = "P") Then
                imRCSave(11, lmRCRowNo) = 1
            End If
            If (tgMVef(ilLoop).sType = "P") And (tgMVef(ilLoop).lPvfCode > 0) Then
                If ilAskQuestion Then
                    ilRet = MsgBox("Compute Package Price", vbYesNo + vbQuestion, "Rate Card")
                    If ilRet = vbNo Then
                        'Exit For
                        Exit Sub
                    End If
                End If
                llPvfCode = tgMVef(ilLoop).lPvfCode
                ReDim tmPvf(0 To 0) As PVF
                Do While llPvfCode > 0
                    tmPvfSrchKey.lCode = llPvfCode
                    ilRet = btrGetEqual(hmPvf, tmPvf(UBound(tmPvf)), imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    llPvfCode = tmPvf(UBound(tmPvf)).lLkPvfCode
                    ReDim Preserve tmPvf(0 To UBound(tmPvf) + 1) As PVF
                Loop
                'Clear price
                For ilWk = LBound(tmRifRec(lmRCRowNo).tRif.lRate) To UBound(tmRifRec(lmRCRowNo).tRif.lRate) Step 1
                    tmRifRec(lmRCRowNo).tRif.lRate(ilWk) = 0
                Next ilWk
                llLkYear = tmRifRec(lmRCRowNo).lLkYear
                Do While llLkYear > 0
                    For ilWk = LBound(tmLkRifRec(llLkYear).tRif.lRate) To UBound(tmLkRifRec(llLkYear).tRif.lRate) Step 1
                        tmLkRifRec(llLkYear).tRif.lRate(ilWk) = 0
                    Next ilWk
                    llLkYear = tmLkRifRec(llLkYear).lLkYear
                Loop
                For llRif = LBONE To UBound(tmRifRec) - 1 Step 1
                    gFindMatch Trim$(smRCSave(VEHINDEX, llRif)), 0, lbcVehicle
                    ilIndex = gLastFound(lbcVehicle)
                    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
                    If ilIndex >= 0 And ilIndex < UBound(tgTempRCUserVehicle) Then
                        slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
                        'ilRet = gParseItem(slNameCode, 1, "\", slVehName)
                        'ilRet = gParseItem(slVehName, 3, "|", slVehName)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilVefCode = Val(slCode)
                        gFindMatch Trim$(smRCSave(DAYPARTINDEX, llRif)), 0, lbcDPName
                        ilIndex = gLastFound(lbcDPName)
                        slNameCode = lbcDPNameCode.List(ilIndex)
                        'ilRet = gParseItem(slNameCode, 1, "\", slDPName)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilRdfCode = Val(slCode)
                        ilFound = False
                        For ilPvf = LBound(tmPvf) To UBound(tmPvf) - 1 Step 1
                            For ilVef = LBound(tmPvf(ilPvf).iVefCode) To UBound(tmPvf(ilPvf).iVefCode) Step 1
                                If (tmPvf(ilPvf).iVefCode(ilVef) = ilVefCode) And (tmPvf(ilPvf).iRdfCode(ilVef) = ilRdfCode) Then
                                    ilFound = True
                                    For ilWk = LBound(tmRifRec(lmRCRowNo).tRif.lRate) To UBound(tmRifRec(lmRCRowNo).tRif.lRate) Step 1
                                        tmRifRec(lmRCRowNo).tRif.lRate(ilWk) = tmRifRec(lmRCRowNo).tRif.lRate(ilWk) + tmPvf(ilPvf).iNoSpot(ilVef) * ((tgMVef(ilLoop).iStdIndex * CSng(tmRifRec(llRif).tRif.lRate(ilWk))) / 100)
                                    Next ilWk
                                    Exit For
                                End If
                            Next ilVef
                            If ilFound Then
                                imRifChg = True
                                Exit For
                            End If
                        Next ilPvf
                    End If
                Next llRif
                mGetShowPrices lmRCRowNo
            End If
        End If
    '        Exit For
    '    End If
    'Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim ilRet As Integer    'Return Status
    Dim slDate As String
    Dim ilMonth As Integer
    Dim ilValue As Integer
    
    Screen.MousePointer = vbHourglass
    
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    
    imLBRCCtrls = 1
    imLBSPCtrls = 1
    imLBWKCtrls = 1
    imLBNWCtrls = 1
    imLBDPCtrls = 1
    imLBSTCtrls = 1
    igJobShowing(RATECARDSJOB) = True
    imFirstActivate = True
    imFirstTime = True
    imTerminate = False
    imPopReqd = False
    imInNew = False
    bmInStdPrice = False
    bmInImportPrice = False
    imcKey.Picture = IconTraf!imcKey.Picture
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    igDPCallSource = 1
    'RateCard.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterForm RateCard
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imRetBranch = False
    imRcfRecLen = Len(tmRcf)  'Get and save ARF record length
    imRCBoxNo = -1 'Initialize current Box to N/A
    lmRCRowNo = 1 'Initialize current Box to N/A
    imDPBoxNo = -1 'Initialize current Box to N/A
    imSPBoxNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imLbcMouseDown = False
    imBypassFocus = False
    imChgMode = False
    imBSMode = False
    igRcfChg = False
    imRifChg = False
    imIgnoreSetting = False
    imButtonIndex = -1
    If tgUrf(0).sRCView = "D" Then
        imView = 1
        pbcRateCard.Visible = False
        pbcDaypart.Visible = True
    Else
        imView = 0
        pbcRateCard.Visible = True
        pbcDaypart.Visible = False
    End If
    imShowIndex = 1 'Std Month
    imTypeIndex = 3 'Flight
    bmInDupicate = False
'    ReDim tgMRdf(1 To 1) As RDF
'    sgMRdfStamp = ""
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    slDate = gObtainEndStd(slDate)
    gObtainMonthYear 0, slDate, ilMonth, imNowYear
    imIgnoreRightMove = False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imFirstTimeSelect = True
    imSettingValue = False
    imLbcArrowSetting = False
    smDefVehicle = Trim$(sgUserDefVehicleName) 'Use rate card vehicle as default if possible
    hmRcf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rcf.Btr)", RateCard
    On Error GoTo 0
    imRcfRecLen = Len(tgRcfI)
    hmRif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rif.btr)", RateCard
    On Error GoTo 0
    ReDim tmRifRec(0 To 1) As RIFREC
    ReDim tmLkRifRec(0 To 1) As RIFREC
    ReDim tmTrashRifRec(0 To 1) As RIFREC
    imRifRecLen = Len(tmRifRec(1).tRif)
'    hmDsf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dsf.Btr)", RateCard
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)
    hmRdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", RateCard
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf)
    imMaxTDRows = UBound(tmRdf.iStartTime, 2)
    hmLtf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ltf.Btr)", RateCard
    On Error GoTo 0
    imLtfRecLen = Len(tmLtf)
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", RateCard
    On Error GoTo 0
    hmAnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Anf.Btr)", RateCard
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)
    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", RateCard
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", RateCard
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", RateCard
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", RateCard
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", RateCard
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", RateCard
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", RateCard
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", RateCard
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmBvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bvf.Btr)", RateCard
    On Error GoTo 0
    imBvfRecLen = Len(tmBvf)
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", RateCard
    On Error GoTo 0
    hmPvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPvf, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pvf.Btr)", RateCard
    On Error GoTo 0
    imPvfRecLen = Len(tmTPvf)
    ilRet = gObtainCorpCal()
    'Get Site Options
    hmSaf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Saf.Btr)", RateCard
    On Error GoTo 0
    imSafRecLen = Len(tmSaf)
    
    tmSafSrchKey1.iVefCode = 0
    ilRet = btrGetEqual(hmSaf, tmSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    
    'Populate facilty and event type list boxes
    lbcVehicle.Clear 'Force population
    mVehPop lbcVehicle
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
'    sgMRdfStamp = ""
    lbcDPName.Clear
    mDPNamePop
    If imTerminate Then
        Exit Sub
    End If
    lbcBudget.Clear 'Force population
    mBudgetPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    'RateCard.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterForm RateCard
    'Traffic!plcHelp.Caption = ""
    mInitBox
    imInHotSpot = False
    imHotSpot(1, 1) = pbcLnWkArrow(0).Left  '3945  'Left
    imHotSpot(1, 2) = 15    'Top
    imHotSpot(1, 3) = imHotSpot(1, 1) + 150 '3945 + 150 'Right
    imHotSpot(1, 4) = 15 + 180  'Bottom
    imHotSpot(2, 1) = pbcLnWkArrow(0).Left + 150 '4095  'Left
    imHotSpot(2, 2) = 15    'Top
    imHotSpot(2, 3) = imHotSpot(2, 1) + 150 '4095 + 150 'Right
    imHotSpot(2, 4) = 15 + 180  'Bottom
    imHotSpot(3, 1) = pbcLnWkArrow(1).Left  '7845  'Left
    imHotSpot(3, 2) = 15    'Top
    imHotSpot(3, 3) = imHotSpot(3, 1) + 150 '7845 + 150 'Right
    imHotSpot(3, 4) = 15 + 180  'Bottom
    imHotSpot(4, 1) = pbcLnWkArrow(1).Left + 150 '7995  'Left
    imHotSpot(4, 2) = 15    'Top
    imHotSpot(4, 3) = imHotSpot(4, 1) + 150 '7995 + 150 'Right
    imHotSpot(4, 4) = 15 + 180  'Bottom
    'mCenterForm RateCard
    gCenterStdAlone RateCard
    
    cbcSelect.Clear  'Force list box to be populated
    mPopulate
    Screen.MousePointer = vbHourglass
    If Not imTerminate Then
        cbcSelect.ListIndex = 0 'This will generate a select_change event
        mSetCommands
    End If
    If tgSpf.sRUseCorpCal <> "Y" Then
        rbcShow(0).Enabled = False
    End If
    ilValue = Asc(tmSaf.sFeatures8)
    If (ilValue And PODADSERVER) = PODADSERVER Then
        cmcCPMpkg.Enabled = True
    Else
        cmcCPMpkg.Enabled = False
    End If
    
    DoEvents
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
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    Dim llRCMax As Long
    Dim llDPMax As Long
    Dim llSTMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long

    flTextHeight = pbcRateCard.TextHeight("1") - 35
    cbcSelect.Move 5400, 60
    'Position panel and picture areas with panel
    plcSP.Move 1170, 120, pbcSP.Width + fgPanelAdj, pbcSP.Height + fgPanelAdj
    pbcSP.Move plcSP.Left + fgBevelX, plcSP.Top + fgBevelY
    plcRateCard.Move 105, 660, pbcRateCard.Width + vbcRateCard.Width + fgPanelAdj, pbcRateCard.Height + fgPanelAdj
    pbcRateCard.Move plcRateCard.Left + fgBevelX, plcRateCard.Top + fgBevelY
    vbcRateCard.Move pbcRateCard.Width + fgBevelX - 15, fgBevelY - 15
    pbcDaypart.Move pbcRateCard.Left, pbcRateCard.Top
    pbcView.Move pbcRateCard.Left + 30, pbcRateCard.Top
    pbcArrow.Move plcRateCard.Left - pbcArrow.Width - 15
    plcShow.Move 615, plcRateCard.Top + plcRateCard.Height + 15
    plcType.Move 5070, plcShow.Top
    plcStatic.Move 105, plcShow.Top + plcShow.Height, pbcStatic.Width + fgPanelAdj, pbcStatic.Height + fgPanelAdj
    pbcStatic.Move plcStatic.Left + fgBevelX, plcStatic.Top + fgBevelY
    plcRCInfo.Move 120, 570
    pbcKey.Move 105, 660
    ''Budget Comparison
    'gSetCtrl tmSPCtrls(BUDGETINDEX), 30, 30, 1965, fgBoxStH
    'Grid Level
    'gSetCtrl tmSPCtrls(GRIDINDEX), 2010, tmSPCtrls(BUDGETINDEX).fBoxY, 765, fgBoxStH
    gSetCtrl tmSPCtrls(GRIDINDEX), 30, 30, 765, fgBoxStH
    'Length
    gSetCtrl tmSPCtrls(LENGTHINDEX), 810, tmSPCtrls(GRIDINDEX).fBoxY, 525, fgBoxStH
    'Current Grid
    gSetCtrl tmSPCtrls(CURGRIDINDEX), 1350, tmSPCtrls(GRIDINDEX).fBoxY, 870, fgBoxStH
    'Vehicle
    'gSetCtrl tmRCCtrls(VEHINDEX), 30, 420, 1800, fgBoxGridH
    gSetCtrl tmRCCtrls(VEHINDEX), 30, 420, 1300, fgBoxGridH
    'Daypart
    gSetCtrl tmRCCtrls(DAYPARTINDEX), 1345, tmRCCtrls(VEHINDEX).fBoxY, 1255, fgBoxGridH
    'CPM
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
        gSetCtrl tmRCCtrls(CPMINDEX), 2300, tmRCCtrls(VEHINDEX).fBoxY, 300, fgBoxGridH
    End If
    'Acquisition Cost
    gSetCtrl tmRCCtrls(ACQUISITIONINDEX), 2610, tmRCCtrls(VEHINDEX).fBoxY, 600, fgBoxGridH
    ''Dollar Index
    'gSetCtrl tmRCCtrls(DOLLARINDEX), 3210, tmRCCtrls(VEHINDEX).fBoxY, 525, fgBoxGridH
    'tmRCCtrls(DOLLARINDEX).iReq = False
    ''Percent Inventory
    'gSetCtrl tmRCCtrls(PCTINVINDEX), 3750, tmRCCtrls(VEHINDEX).fBoxY, 480, fgBoxGridH
    'tmRCCtrls(PCTINVINDEX).iReq = False
    'Base Index
    'gSetCtrl tmRCCtrls(BASEINDEX), 3210, tmRCCtrls(VEHINDEX).fBoxY, 330, fgBoxGridH
    gSetCtrl tmRCCtrls(BASEINDEX), 3225, tmRCCtrls(VEHINDEX).fBoxY, 330, fgBoxGridH
    'Report Index
    'gSetCtrl tmRCCtrls(RPTINDEX), 3555, tmRCCtrls(VEHINDEX).fBoxY, 330, fgBoxGridH
    gSetCtrl tmRCCtrls(RPTINDEX), 3570, tmRCCtrls(VEHINDEX).fBoxY, 315, fgBoxGridH
    'Sort Index
    gSetCtrl tmRCCtrls(SORTINDEX), 3900, tmRCCtrls(VEHINDEX).fBoxY, 330, fgBoxGridH
    tmRCCtrls(SORTINDEX).iReq = False
    'Dollar 1
    gSetCtrl tmRCCtrls(DOLLAR1INDEX), 4245, tmRCCtrls(VEHINDEX).fBoxY, 885, fgBoxGridH
    'Dollar 2
    gSetCtrl tmRCCtrls(DOLLAR2INDEX), 5145, tmRCCtrls(VEHINDEX).fBoxY, 885, fgBoxGridH
    'Dollar 3
    gSetCtrl tmRCCtrls(DOLLAR3INDEX), 6045, tmRCCtrls(VEHINDEX).fBoxY, 885, fgBoxGridH
    'Dollar 4
    gSetCtrl tmRCCtrls(DOLLAR4INDEX), 6945, tmRCCtrls(VEHINDEX).fBoxY, 885, fgBoxGridH
    'Total
    gSetCtrl tmRCCtrls(AVGINDEX), 7845, tmRCCtrls(VEHINDEX).fBoxY, 885, fgBoxGridH

    'Vehicle
    gSetCtrl tmDPCtrls(VEHINDEX), 30, 420, 1600, fgBoxGridH
    'Daypart
    gSetCtrl tmDPCtrls(DAYPARTINDEX), 1645, tmDPCtrls(VEHINDEX).fBoxY, 1555, fgBoxGridH
    'added by L. Bianchi
    'CPM
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
        gSetCtrl tmDPCtrls(CPMINDEX), 2760, tmDPCtrls(VEHINDEX).fBoxY, 385, fgBoxGridH
    End If
    'Times
    gSetCtrl tmDPCtrls(TIMESINDEX), 3210, tmDPCtrls(VEHINDEX).fBoxY, 2640, fgBoxGridH
    'Days of the week
    For ilLoop = 0 To 6 Step 1
        gSetCtrl tmDPCtrls(DAYINDEX + ilLoop), 5865 + 225 * (ilLoop), tmDPCtrls(VEHINDEX).fBoxY, 210, fgBoxGridH
        tmDPCtrls(DAYINDEX + ilLoop).iReq = False
    Next ilLoop
    'Avails
    gSetCtrl tmDPCtrls(AVAILINDEX), 7440, tmDPCtrls(VEHINDEX).fBoxY, 990, fgBoxGridH
    'gSetCtrl tmDPCtrls(AVAILINDEX + 1), 7440, tmDPCtrls(VEHINDEX).fBoxY, 990, fgBoxGridH
    'tmDPCtrls(AVAILINDEX + 1).iReq = False
    'Hour flag
    gSetCtrl tmDPCtrls(HRSINDEX), 8445, tmDPCtrls(VEHINDEX).fBoxY, 285, fgBoxGridH
    tmDPCtrls(HRSINDEX).iReq = False
    'Vehicle
    gSetCtrl tmSTCtrls(STVEHINDEX), 30, 30, 1800, fgBoxGridH
    'Title
    gSetCtrl tmSTCtrls(STTITLEINDEX), 1845, tmSTCtrls(VEHINDEX).fBoxY, 2385, fgBoxGridH
    'Dollar 1
    gSetCtrl tmSTCtrls(STDOLLAR1INDEX), 4245, tmSTCtrls(STVEHINDEX).fBoxY, 885, fgBoxGridH
    'Dollar 2
    gSetCtrl tmSTCtrls(STDOLLAR2INDEX), 5145, tmSTCtrls(STVEHINDEX).fBoxY, 885, fgBoxGridH
    'Dollar 3
    gSetCtrl tmSTCtrls(STDOLLAR3INDEX), 6045, tmSTCtrls(STVEHINDEX).fBoxY, 885, fgBoxGridH
    'Dollar 4
    gSetCtrl tmSTCtrls(STDOLLAR4INDEX), 6945, tmSTCtrls(STVEHINDEX).fBoxY, 885, fgBoxGridH
    'Total
    gSetCtrl tmSTCtrls(STAVGINDEX), 7845, tmSTCtrls(STVEHINDEX).fBoxY, 885, fgBoxGridH

    'Week 1
    gSetCtrl tmWKCtrls(WK1INDEX), 4245, 30, 885, fgBoxGridH
    'Week 2
    gSetCtrl tmWKCtrls(WK2INDEX), 5145, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Week 3
    gSetCtrl tmWKCtrls(WK3INDEX), 6045, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Week 4
    gSetCtrl tmWKCtrls(WK4INDEX), 6945, tmWKCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    '# Week 1
    gSetCtrl tmNWCtrls(NW1INDEX), 4245, 225, 885, fgBoxGridH
    '# Week 2
    gSetCtrl tmNWCtrls(NW2INDEX), 5145, tmNWCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    '# Week 3
    gSetCtrl tmNWCtrls(NW3INDEX), 6045, tmNWCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    '# Week 4
    gSetCtrl tmNWCtrls(NW4INDEX), 6945, tmNWCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH
    'Total
    gSetCtrl tmNWCtrls(NWAVGINDEX), 7845, tmNWCtrls(WK1INDEX).fBoxY, 885, fgBoxGridH



    llRCMax = 0
    For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
         If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER And ilLoop = CPMINDEX Then
                    GoTo Skip_Loop
          End If
        tmRCCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmRCCtrls(ilLoop).fBoxW)
        Do While (tmRCCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmRCCtrls(ilLoop).fBoxW = tmRCCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmRCCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmRCCtrls(ilLoop).fBoxX)
            Do While (tmRCCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmRCCtrls(ilLoop).fBoxX = tmRCCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmRCCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmRCCtrls(ilLoop - 1).fBoxX + tmRCCtrls(ilLoop - 1).fBoxW + 15 < tmRCCtrls(ilLoop).fBoxX Then
                        tmRCCtrls(ilLoop - 1).fBoxW = tmRCCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmRCCtrls(ilLoop - 1).fBoxX + tmRCCtrls(ilLoop - 1).fBoxW + 15 > tmRCCtrls(ilLoop).fBoxX Then
                        tmRCCtrls(ilLoop - 1).fBoxW = tmRCCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmRCCtrls(ilLoop).fBoxX + tmRCCtrls(ilLoop).fBoxW + 15 > llRCMax Then
            llRCMax = tmRCCtrls(ilLoop).fBoxX + tmRCCtrls(ilLoop).fBoxW + 15
        End If
Skip_Loop:
    Next ilLoop
    
    llDPMax = 0
    For ilLoop = imLBDPCtrls To HRSINDEX Step 1
        'added by L. Bianchi
        If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER And ilLoop = CPMINDEX Then
                    GoTo Skip_DPLoop
        End If
        tmDPCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmDPCtrls(ilLoop).fBoxW)
        Do While (tmDPCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmDPCtrls(ilLoop).fBoxW = tmDPCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmDPCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmDPCtrls(ilLoop).fBoxX)
            Do While (tmDPCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDPCtrls(ilLoop).fBoxX = tmDPCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmDPCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmDPCtrls(ilLoop - 1).fBoxX + tmDPCtrls(ilLoop - 1).fBoxW + 15 < tmDPCtrls(ilLoop).fBoxX Then
                        tmDPCtrls(ilLoop - 1).fBoxW = tmDPCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDPCtrls(ilLoop - 1).fBoxX + tmDPCtrls(ilLoop - 1).fBoxW + 15 > tmDPCtrls(ilLoop).fBoxX Then
                        tmDPCtrls(ilLoop - 1).fBoxW = tmDPCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmDPCtrls(ilLoop).fBoxX + tmDPCtrls(ilLoop).fBoxW + 15 > llDPMax Then
            llDPMax = tmDPCtrls(ilLoop).fBoxX + tmDPCtrls(ilLoop).fBoxW + 15
        End If
Skip_DPLoop:
    Next ilLoop

    llSTMax = 0
    For ilLoop = imLBSTCtrls To UBound(tmSTCtrls) Step 1
        tmSTCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmSTCtrls(ilLoop).fBoxW)
        Do While (tmSTCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmSTCtrls(ilLoop).fBoxW = tmSTCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 1 Then
            tmSTCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmSTCtrls(ilLoop).fBoxX)
            Do While (tmSTCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmSTCtrls(ilLoop).fBoxX = tmSTCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmSTCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmSTCtrls(ilLoop - 1).fBoxX + tmSTCtrls(ilLoop - 1).fBoxW + 15 < tmSTCtrls(ilLoop).fBoxX Then
                        tmSTCtrls(ilLoop - 1).fBoxW = tmSTCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmSTCtrls(ilLoop - 1).fBoxX + tmSTCtrls(ilLoop - 1).fBoxW + 15 > tmSTCtrls(ilLoop).fBoxX Then
                        tmSTCtrls(ilLoop - 1).fBoxW = tmSTCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmSTCtrls(ilLoop).fBoxX + tmSTCtrls(ilLoop).fBoxW + 15 > llSTMax Then
            llSTMax = tmSTCtrls(ilLoop).fBoxX + tmSTCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop

    For ilLoop = DOLLAR1INDEX To DOLLAR4INDEX Step 1
        tmWKCtrls(ilLoop - DOLLAR1INDEX + 1).fBoxX = tmRCCtrls(ilLoop).fBoxX
        tmWKCtrls(ilLoop - DOLLAR1INDEX + 1).fBoxW = tmRCCtrls(ilLoop).fBoxW
    Next ilLoop
    For ilLoop = DOLLAR1INDEX To AVGINDEX Step 1
        tmNWCtrls(ilLoop - DOLLAR1INDEX + 1).fBoxX = tmRCCtrls(ilLoop).fBoxX
        tmNWCtrls(ilLoop - DOLLAR1INDEX + 1).fBoxW = tmRCCtrls(ilLoop).fBoxW
    Next ilLoop

    If llRCMax > llDPMax Then
        tmDPCtrls(HRSINDEX).fBoxW = tmDPCtrls(HRSINDEX).fBoxW + llRCMax - llDPMax
    ElseIf llRCMax < llDPMax Then
        tmDPCtrls(HRSINDEX).fBoxW = tmDPCtrls(HRSINDEX).fBoxW + llRCMax - llDPMax
    End If
    pbcRateCard.Width = llRCMax
    plcRateCard.Width = llRCMax + vbcRateCard.Width + 2 * fgBevelX + 15
    lacRCFrame.Width = llRCMax - 15
    lacDPFrame.Width = lacRCFrame.Width
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    
    cmcImport.Left = plcRateCard.Left
    cmcResetStdPrice.Left = cmcImport.Left
    
    cmcDone.Left = (RateCard.Width - 10 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    If cmcResetStdPrice.Left + cmcResetStdPrice.Width + cmcDone.Width / 3 > cmcDone.Left Then
        cmcDone.Left = cmcResetStdPrice.Left + cmcResetStdPrice.Width + cmcDone.Width / 3
    End If
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcErase.Left = cmcUpdate.Left + cmcUpdate.Width + ilSpaceBetweenButtons
    cmcUndo.Left = cmcErase.Left + cmcErase.Width + ilSpaceBetweenButtons
    cmcReport.Left = cmcUndo.Left + cmcUndo.Width + ilSpaceBetweenButtons
    cmcStdPkg.Left = cmcReport.Left + cmcReport.Width + ilSpaceBetweenButtons
    cmcCPMpkg.Left = cmcStdPkg.Left + cmcStdPkg.Width + ilSpaceBetweenButtons
    cmcDupl.Left = cmcCPMpkg.Left + cmcCPMpkg.Width + ilSpaceBetweenButtons
    cmcImpact.Left = cmcDupl.Left + cmcDupl.Width + ilSpaceBetweenButtons
    cmcDone.Top = RateCard.Height - cmcDone.Height - 180
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    cmcErase.Top = cmcDone.Top
    cmcUndo.Top = cmcDone.Top
    cmcReport.Top = cmcDone.Top
    cmcStdPkg.Top = cmcDone.Top
    cmcDupl.Top = cmcDone.Top
    cmcImpact.Top = cmcDone.Top
    imcTrash.Top = RateCard.Height - imcTrash.Height - 120
    imcTrash.Left = RateCard.Width - (3 * imcTrash.Width) / 2
    'PODCAST - 12/22/2020 - Add CPM Pkg to Rate Card Screen:
    cmcCPMpkg.Top = cmcDone.Top
    
    plcShow.Top = cmcDone.Top - plcShow.Height - 120
    plcType.Top = plcShow.Top
    llAdjTop = plcShow.Top - plcRateCard.Top - fgBevelY - 120
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcRateCard.Top + llAdjTop + 2 * fgBevelY + 240 < plcShow.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcRateCard.Height = llAdjTop + 2 * fgBevelY + 30
    pbcRateCard.Left = plcRateCard.Left + fgBevelX
    pbcRateCard.Top = plcRateCard.Top + fgBevelY
    pbcRateCard.Height = plcRateCard.Height - 2 * fgBevelY
    vbcRateCard.Left = plcRateCard.Width - vbcRateCard.Width - fgBevelX - 30
    vbcRateCard.Top = fgBevelY
    vbcRateCard.Height = pbcRateCard.Height
    pbcView.Top = pbcRateCard.Top + 15
    pbcView.Left = pbcRateCard.Left + 15
    pbcDaypart.Width = pbcRateCard.Width
    pbcDaypart.Height = pbcRateCard.Height
    pbcDaypart.Move pbcRateCard.Left, pbcRateCard.Top, pbcDaypart.Width, pbcDaypart.Height
    pbcRateCard.Picture = LoadPicture("")
    pbcDaypart.Picture = LoadPicture("")
    cbcSelect.Left = plcRateCard.Left + plcRateCard.Width - cbcSelect.Width
    cmcTerms.Left = RateCard.Width / 2 - cmcTerms.Width / 2
    If fmAdjFactorW >= 1.2 Then
        ilSpaceBetweenButtons = plcType.Left - plcShow.Left - plcShow.Width
        'plcType.Left = pbcRateCard.Left + tmRCCtrls(DOLLAR1INDEX).fBoxX
        'plcShow.Left = plcType.Left - ilSpaceBetweenButtons - plcShow.Width
        plcShow.Left = cmcDone.Left
        plcType.Left = plcShow.Left + ilSpaceBetweenButtons + plcShow.Width
    End If
    pbcLnWkArrow(0).Left = tmWKCtrls(WK1INDEX).fBoxX - pbcLnWkArrow(0).Width - 30
    pbcLnWkArrow(0).Top = 15
    pbcLnWkArrow(1).Left = tmWKCtrls(WK4INDEX).fBoxX + tmWKCtrls(WK4INDEX).fBoxW + 60
    pbcLnWkArrow(1).Top = 15
    pbcSPTab.Left = -100
    pbcSPSTab.Left = -100
    pbcTab.Left = -100
    pbcSTab.Left = -100
    pbcClickFocus.Left = -100
    
    cmcImport.Top = plcRateCard.Top + plcRateCard.Height + 60
    cmcResetStdPrice.Top = cmcImport.Top + cmcImport.Height + 60
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitShowPrice                  *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize price to be shown   *
'*                      on rate card                   *
'*                                                     *
'*******************************************************
Private Sub mInitLastShow()
    Dim ilLoop As Integer
    Dim ilLast As Integer
    ilLast = UBound(smRCShow, 2)
    For ilLoop = LBound(smRCSave, 1) To UBound(smRCSave, 1) Step 1
        smRCSave(ilLoop, ilLast) = ""
    Next ilLoop
    For ilLoop = LBound(smRCShow, 1) To UBound(smRCShow, 1) Step 1
        smRCShow(ilLoop, ilLast) = ""
    Next ilLoop
    ilLast = UBound(smDPShow, 2)
    For ilLoop = LBound(smDPShow, 1) To UBound(smDPShow, 1) Step 1
        smDPShow(ilLoop, ilLast) = ""
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitRateCardCtrls              *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Sub mInitRateCardCtrls()
    Dim ilLoop As Integer
    Dim llIndex As Long
    Dim llUpperBound As Long
    Dim slStr As String
    Dim ilLower As Integer
    'gPDNToStr tgRcfI.sRound, 2, smRound
    slStr = gLongToStrDec(tgRcfI.lRound, 2)
    imPdYear = imNowYear
    imPdStartWk = 1
    imPdStartFltNo = 1
    imRifStartYear = 0
    imRifNoYears = 0
    If UBound(tmRifRec) > 1 Then
        ilLower = LBONE 'LBound(tmRifRec)
        imRifStartYear = tmRifRec(ilLower).tRif.iYear
        imRifNoYears = 1
'Only allow one year
'        ilIndex = tmRifRec(ilLower).lLkYear
'        Do While ilIndex > 0
'            imRifNoYears = imRifNoYears + 1
'            ilIndex = tmLkRifRec(ilIndex).lLkYear
'        Loop
        'Adjust Period to be viewed
        If imRifStartYear > imPdYear Then
            imPdYear = imRifStartYear
        ElseIf imRifStartYear + imRifNoYears - 1 < imPdYear Then
            imPdYear = imRifStartYear + imRifNoYears - 1
        End If
    Else
        imPdYear = tgRcfI.iYear
        imRifStartYear = imPdYear
        imRifNoYears = 1
    End If
    lbcGrid.Clear
    If tgRcfI.iGridsUsed <= 1 Then
        lbcGrid.AddItem "I"
    Else
        lbcGrid.Enabled = True
        For ilLoop = 1 To tgRcfI.iGridsUsed Step 1
            lbcGrid.AddItem gIntToRoman(ilLoop)
        Next ilLoop
    End If
    imLen = 0
    lbcLen.Clear
    If tgRcfI.sUseLen <> "N" Then
        For ilLoop = LBound(tgRcfI.iLen) To UBound(tgRcfI.iLen) Step 1
            If tgRcfI.iLen(ilLoop) <> 0 Then
                imLen = imLen + 1
                lbcLen.AddItem Trim$(Str$(tgRcfI.iLen(ilLoop)))
            End If
        Next ilLoop
    End If
    llUpperBound = UBound(tmRifRec)
    ReDim smRCShow(0 To AVGINDEX, 0 To llUpperBound) As String * 40 'Values shown in program area
    ReDim smRCSave(0 To SORTINDEX, 0 To llUpperBound) As String * 40 'Values saved (program name) in program area
    ReDim lmRCSave(0 To TOTALINDEX - SORTINDEX, 0 To llUpperBound) As Long 'Values saved (program name) in program area
   'updated by l.bianchi
    ReDim imRCSave(0 To 17, 0 To llUpperBound) As Integer 'Values saved (program name) in program area
    
    ReDim smDPShow(0 To DPBASEINDEX, 0 To llUpperBound) As String * 40 'Values shown in program area
    For ilLoop = imLBWKCtrls To UBound(tmWKCtrls) Step 1
        tmWKCtrls(ilLoop).sShow = ""
    Next ilLoop
    For ilLoop = imLBNWCtrls To UBound(tmNWCtrls) Step 1
        tmNWCtrls(ilLoop).sShow = ""
    Next ilLoop
    For ilLoop = LBound(smRCShow, 1) To UBound(smRCShow, 1) Step 1
        For llIndex = LBound(smRCShow, 2) To UBound(smRCShow, 2) Step 1
            smRCShow(ilLoop, llIndex) = ""
        Next llIndex
    Next ilLoop
    For ilLoop = LBound(smRCSave, 1) To UBound(smRCSave, 1) Step 1
        For llIndex = LBound(smRCSave, 2) To UBound(smRCSave, 2) Step 1
            smRCSave(ilLoop, llIndex) = ""
        Next llIndex
    Next ilLoop
    For ilLoop = LBound(smDPShow, 1) To UBound(smDPShow, 1) Step 1
        For llIndex = LBound(smDPShow, 2) To UBound(smDPShow, 2) Step 1
            smDPShow(ilLoop, llIndex) = ""
        Next llIndex
    Next ilLoop
    imSettingValue = True
    vbcRateCard.Min = LBONE 'LBound(tmRifRec)
    imSettingValue = True
    If UBound(tmRifRec) <= vbcRateCard.LargeChange Then ' + 1 Then
        vbcRateCard.Max = LBONE 'LBound(tmRifRec)
    Else
        vbcRateCard.Max = UBound(tmRifRec) - vbcRateCard.LargeChange
    End If
    imSettingValue = True
    vbcRateCard.Value = vbcRateCard.Min
    imSettingValue = False
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitRif                        *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize a record            *
'*                                                     *
'*******************************************************
Private Sub mInitRif(llIndex As Long)
    Dim slDate As String
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilLoop As Integer
    Dim llLkUpper As Long
    If imRifStartYear <= 0 Then 'Done defined
        gUnpackDate tgRcfI.iStartDate(0), tgRcfI.iStartDate(1), slDate
        slDate = gObtainEndStd(slDate)
        gObtainMonthYear 0, slDate, ilMonth, ilYear
        imRifStartYear = ilYear
        imRifNoYears = 1
        imPdYear = ilYear
        imPdStartWk = 1
    End If
    tmRifRec(llIndex).tRif.iYear = tgRcfI.iYear
    tmRifRec(llIndex).lLkYear = 0
    tmRifRec(llIndex).iType = 0
    tmRifRec(llIndex).iStatus = 0
    tmRifRec(llIndex).lRecPos = 0
    'tmRifRec(llIndex).tRif.lRateIndex = 0
    'tmRifRec(llIndex).tRif.iROSPct = 0
    tmRifRec(llIndex).tRif.lAcquisitionCost = 0
    tmRifRec(llIndex).tRif.sBase = ""
    tmRifRec(llIndex).tRif.sRpt = ""
    tmRifRec(llIndex).tRif.iSort = -1
    For ilLoop = LBound(tmRifRec(llIndex).tRif.lRate) To UBound(tmRifRec(llIndex).tRif.lRate) Step 1
        'slStr = ""
        'gStrToPDN slStr, 2, 5, tmRifRec(llIndex).tRif.sRate(ilLoop)
        tmRifRec(llIndex).tRif.lRate(ilLoop) = 0
    Next ilLoop
    If imRifNoYears > 1 Then
        llLkUpper = UBound(tmLkRifRec)
        tmRifRec(llIndex).lLkYear = llLkUpper
        For ilLoop = 2 To imRifNoYears Step 1
            tmLkRifRec(llLkUpper) = tmRifRec(llIndex)
            tmLkRifRec(llLkUpper).tRif.iYear = tmRifRec(llIndex).tRif.iYear + ilLoop - 1
            tmLkRifRec(llLkUpper).lLkYear = 0
            If ilLoop > 2 Then
                tmLkRifRec(llLkUpper - 1).lLkYear = llLkUpper
            End If
            llLkUpper = llLkUpper + 1
            ReDim Preserve tmLkRifRec(0 To llLkUpper) As RIFREC
        Next ilLoop
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:6/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec iTest
'   Where:
'
    Dim llLoop As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim llLkYear As Long
    If (tgRcfI.iGridsUsed > 0) And (imCGDSelectedIndex >= 0) Then
        tgRcfI.iTodayGrid = imCGDSelectedIndex  ' + 1
    End If
    For llLoop = LBONE To UBound(tmRifRec) - 1 Step 1
        'L.Bianchi
        gFindMatch Trim$(smRCSave(VEHINDEX, llLoop)), 0, lbcVehicle
        ilIndex = gLastFound(lbcVehicle)
        If ilIndex = -1 Then
            Exit Sub
        End If
        'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
        slCode = 0
        slNameCode = ""
        If ilIndex >= 0 And ilIndex < UBound(tgTempRCUserVehicle) Then
            slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
        End If
        On Error GoTo mMoveCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", RateCard
        On Error GoTo 0
        slCode = Trim$(slCode)
        If tmRifRec(llLoop).tRif.iVefCode <> CInt(slCode) Then
            ilRet = ilRet
        End If
        tmRifRec(llLoop).tRif.iVefCode = 0
        On Error Resume Next
        tmRifRec(llLoop).tRif.iVefCode = CInt(slCode)
        On Error GoTo 0
        'Daypart
        gFindMatch Trim$(smRCSave(DAYPARTINDEX, llLoop)), 0, lbcDPName
        ilIndex = gLastFound(lbcDPName)
        slNameCode = lbcDPNameCode.List(ilIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveCtrlToRec (gParseItem field 2)", RateCard
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmRifRec(llLoop).tRif.iRdfCode = 0
        On Error Resume Next
        tmRifRec(llLoop).tRif.iRdfCode = CInt(slCode)
        On Error GoTo 0
        'slStr = smRCSave(BASEINDEX, llLoop)
        'tmRifRec(llLoop).tRif.lRateIndex = gStrDecToLong(slStr, 3)
        'If smDPShow(DPBASEINDEX, llLoop) = "Y" Then
        '    slStr = ""
        'Else
        '    slStr = smRCSave(RPTINDEX, llLoop)
        'End If
        'tmRifRec(llLoop).tRif.iROSPct = gStrDecToInt(slStr, 2)
        'gStrDecToLong(slStr, 3)
        If imRCSave(17, llLoop) = 1 And imRCSave(16, llLoop) = 0 Then
            If Trim$(smRCSave(CPMINDEX, llLoop)) <> "" Then
                tmRifRec(llLoop).tRif.lPodCPM = gStrDecToLong(Trim$(smRCSave(CPMINDEX, llLoop)), 2)
            End If
        End If
        tmRifRec(llLoop).tRif.lAcquisitionCost = 0
        If Trim$(smRCSave(ACQUISITIONINDEX, llLoop)) <> "Y" Then
            tmRifRec(llLoop).tRif.lAcquisitionCost = gStrDecToLong(Trim$(smRCSave(ACQUISITIONINDEX, llLoop)), 2)
        End If
        tmRifRec(llLoop).tRif.sBase = Trim$(smRCSave(BASEINDEX, llLoop))
        tmRifRec(llLoop).tRif.sRpt = Trim$(smRCSave(RPTINDEX, llLoop))
        tmRifRec(llLoop).tRif.iSort = Val(Trim$(smRCSave(SORTINDEX, llLoop)))
        If tmRifRec(llLoop).tRif.iSort < 0 Then
            tmRifRec(llLoop).tRif.iSort = 0
        End If
        llLkYear = tmRifRec(llLoop).lLkYear
        Do While llLkYear > 0
            If tmLkRifRec(llLkYear).tRif.iVefCode <> tmRifRec(llLoop).tRif.iVefCode Then
                ilRet = ilRet
            End If
            tmLkRifRec(llLkYear).tRif.iVefCode = tmRifRec(llLoop).tRif.iVefCode
            tmLkRifRec(llLkYear).tRif.iRdfCode = tmRifRec(llLoop).tRif.iRdfCode
            'tmLkRifRec(llLkYear).tRif.lRateIndex = tmRifRec(llLoop).tRif.lRateIndex
            'tmLkRifRec(llLkYear).tRif.iROSPct = tmRifRec(llLoop).tRif.iROSPct
            tmLkRifRec(llLkYear).tRif.sBase = tmRifRec(llLoop).tRif.sBase
            tmLkRifRec(llLkYear).tRif.sRpt = tmRifRec(llLoop).tRif.sRpt
            tmLkRifRec(llLkYear).tRif.iSort = tmRifRec(llLoop).tRif.iSort
            llLkYear = tmLkRifRec(llLkYear).lLkYear
        Loop
    Next llLoop
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim slStr As String
    Dim llRowNo As Long
    
    Dim llPvfCode As Long
    Dim ilRet As Integer
    
    slStr = ""
    If tgRcfI.iGridsUsed > 0 Then
        lbcGrid.ListIndex = tgRcfI.iTodayGrid   ' - 1
        imGDSelectedIndex = lbcGrid.ListIndex
        slStr = lbcGrid.List(imGDSelectedIndex)
    End If
    gSetShow pbcSP, slStr, tmSPCtrls(GRIDINDEX)
    slStr = ""
    If lbcLen.ListCount > 0 Then
        gFindMatch Str$(tgRcfI.iBaseLen), 0, lbcLen
        If gLastFound(lbcLen) >= 0 Then
            lbcLen.ListIndex = gLastFound(lbcLen)
        Else
            lbcLen.ListIndex = 0
        End If
        imLenSelectedIndex = lbcLen.ListIndex
        slStr = lbcLen.List(imLenSelectedIndex)
    End If
    gSetShow pbcSP, slStr, tmSPCtrls(LENGTHINDEX)
    slStr = ""
    If (tgRcfI.iGridsUsed > 0) And (tgRcfI.iTodayGrid > 0) And (tgRcfI.iTodayGrid <= tgRcfI.iGridsUsed) Then
        lbcGrid.ListIndex = tgRcfI.iTodayGrid   ' - 1
        imCGDSelectedIndex = lbcGrid.ListIndex
        slStr = lbcGrid.List(imCGDSelectedIndex)
    End If
    gSetShow pbcSP, slStr, tmSPCtrls(CURGRIDINDEX)
    'Set default vehicle
    smDefVehicle = Trim$(sgUserDefVehicleName)
    If (tgRcfI.iVefCode <> 0) And (tgRcfI.iVefCode <> -32000) Then
        'For ilTest = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
        '    If tmUserVeh(ilTest).iCode = tgRcfI.iVefCode Then
            ilTest = mBinarySearch(tgRcfI.iVefCode)
            If ilTest <> -1 Then
                smDefVehicle = tmUserVeh(ilTest).sName
            End If
        '        Exit For
        '    End If
        'Next ilTest
    End If
    smDefVehicle = Trim$(smDefVehicle)
    For llRowNo = LBONE To UBound(tmRifRec) - 1 Step 1
        smRCSave(VEHINDEX, llRowNo) = ""
        'For ilTest = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
        '    If tmUserVeh(ilTest).iCode = tmRifRec(llRowNo).tRif.iVefCode Then
            ilTest = mBinarySearch(tmRifRec(llRowNo).tRif.iVefCode)
            If ilTest <> -1 Then
                smRCSave(VEHINDEX, llRowNo) = tmUserVeh(ilTest).sName
            End If
         '       Exit For
        '    End If
        'Next ilTest
        imRCSave(9, llRowNo) = 0
        imRCSave(11, llRowNo) = 0
        imRCSave(16, llRowNo) = 0
        imRCSave(17, llRowNo) = 0
        imRCSave(12, llRowNo) = 0
        'For ilTest = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilTest).iCode = tmRifRec(llRowNo).tRif.iVefCode Then
            ilTest = gBinarySearchVef(tmRifRec(llRowNo).tRif.iVefCode)
            If ilTest <> -1 Then
                If tgMVef(ilTest).sState = "D" Then
                    imRCSave(9, llRowNo) = 1
                End If
                If tgMVef(ilTest).sType = "P" Then
                    imRCSave(11, llRowNo) = 1
                End If
                'L.Bianchi
                mSetPvfType llRowNo
                'TTP 10388 - Rate card screen: if "ad server" site option setting is not enabled, a "podcast" medium type vehicle does not appear on the rate card screen
                '02/23/22 - JW Fix issue on the Rate Card screen per Jason Email: Wed 2/16/22 10:46 AM
                mSetVehicleMediumType llRowNo
            End If
        '        Exit For
        '    End If
        'Next ilTest
        smRCSave(VEHINDEX, llRowNo) = Trim$(smRCSave(VEHINDEX, llRowNo))
        If Trim$(smRCSave(VEHINDEX, llRowNo)) = "" Then
            gFindMatch smDefVehicle, 0, lbcVehicle
            If gLastFound(lbcVehicle) >= 0 Then
                smRCSave(VEHINDEX, llRowNo) = smDefVehicle
            Else
                gFindMatch sgUserDefVehicleName, 0, lbcVehicle
                If gLastFound(lbcVehicle) >= 0 Then
                    smRCSave(VEHINDEX, llRowNo) = sgUserDefVehicleName
                Else
                    smRCSave(VEHINDEX, llRowNo) = ""
                End If
            End If
        End If
        slStr = Trim$(smRCSave(VEHINDEX, llRowNo))
        gSetShow pbcRateCard, slStr, tmRCCtrls(VEHINDEX)
        smRCShow(VEHINDEX, llRowNo) = tmRCCtrls(VEHINDEX).sShow
        slStr = Trim$(smRCSave(VEHINDEX, llRowNo))
        gSetShow pbcDaypart, slStr, tmDPCtrls(VEHINDEX)
        smDPShow(VEHINDEX, llRowNo) = tmDPCtrls(VEHINDEX).sShow
        imRCSave(10, llRowNo) = 0
        If tmRifRec(llRowNo).tRif.iRdfCode > 0 Then
            'tmRdfSrchKey.iCode = tmRifRec(llRowNo).tRif.iRdfCode
            'ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            'If ilRet = BTRV_ERR_NONE Then
            '    smRCSave(DAYPARTINDEX, llRowNo) = Trim$(tmRdf.sName)
            '    mSetDPShow llRowNo
            'End If
            For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                If tmRifRec(llRowNo).tRif.iRdfCode = tgMRdf(ilLoop).iCode Then
                    tmRdf = tgMRdf(ilLoop)
                    smRCSave(DAYPARTINDEX, llRowNo) = Trim$(tmRdf.sName)
                    If tmRdf.sState = "D" Then
                        imRCSave(10, llRowNo) = 1
                    End If
                    mSetDPShow llRowNo
                    Exit For
                End If
            Next ilLoop
        End If
        'CPMINDEX L.Bianchi
        If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER And imRCSave(17, llRowNo) = 1 And imRCSave(16, llRowNo) = 0 Then
            slStr = gLongToStrDec(tmRifRec(llRowNo).tRif.lPodCPM, 2)
            
            gSetShow pbcRateCard, slStr, tmRCCtrls(CPMINDEX)
            smRCSave(CPMINDEX, llRowNo) = tmRCCtrls(CPMINDEX).sShow
            smRCShow(CPMINDEX, llRowNo) = tmRCCtrls(CPMINDEX).sShow
            gSetShow pbcDaypart, slStr, tmDPCtrls(CPMINDEX)
            smDPShow(CPMINDEX, llRowNo) = tmDPCtrls(CPMINDEX).sShow
        End If
        'Acquisition
        smRCSave(ACQUISITIONINDEX, llRowNo) = "Y"
        slStr = ""
        If ilTest <> -1 Then
            If ((tgMVef(ilTest).sType = "C") Or (tgMVef(ilTest).sType = "S")) And ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
                If tmRifRec(llRowNo).tRif.lAcquisitionCost > 0 Then
                    smRCSave(ACQUISITIONINDEX, llRowNo) = gLongToStrDec(tmRifRec(llRowNo).tRif.lAcquisitionCost, 2)
                Else
                    smRCSave(ACQUISITIONINDEX, llRowNo) = ""
                End If
                slStr = smRCSave(ACQUISITIONINDEX, llRowNo)
            End If
        End If
        gSetShow pbcRateCard, slStr, tmRCCtrls(ACQUISITIONINDEX)
        smRCShow(ACQUISITIONINDEX, llRowNo) = tmRCCtrls(ACQUISITIONINDEX).sShow
        'slStr = gLongToStrDec(tmRifRec(llRowNo).tRif.lRateIndex, 3)
        'smRCSave(BASEINDEX, llRowNo) = slStr
        'gSetShow pbcRateCard, slStr, tmRCCtrls(DOLLARINDEX)
        'smRCShow(DOLLARINDEX, llRowNo) = tmRCCtrls(DOLLARINDEX).sShow
        'slStr = gIntToStrDec(tmRifRec(llRowNo).tRif.iROSPct, 2)
        'If smDPShow(DPBASEINDEX, llRowNo) = "Y" Then
        '    slStr = ""
        '    smRCSave(RPTINDEX, llRowNo) = slStr
        'Else
        '    smRCSave(RPTINDEX, llRowNo) = slStr
        '    gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
        'End If
        'gSetShow pbcRateCard, slStr, tmRCCtrls(PCTINVINDEX)
        'smRCShow(PCTINVINDEX, llRowNo) = tmRCCtrls(PCTINVINDEX).sShow
        'L.Bianchi
'        If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER And imRCSave(17, llRowNo) = 1 Then
'            tmRifRec(llRowNo).tRif.sBase = "N"
'            tmRifRec(llRowNo).tRif.sRpt = "N"
'        End If
' TTP 10901 JJB 2023-12-12 Commented out above code because Jason said the logic no longer applied.

        If (tmRifRec(llRowNo).tRif.sBase = "Y") Or (tmRifRec(llRowNo).tRif.sBase = "N") Then
            slStr = tmRifRec(llRowNo).tRif.sBase
            smRCSave(BASEINDEX, llRowNo) = slStr
            gSetShow pbcRateCard, slStr, tmRCCtrls(BASEINDEX)
            smRCShow(BASEINDEX, llRowNo) = tmRCCtrls(BASEINDEX).sShow
            slStr = tmRifRec(llRowNo).tRif.sRpt
            smRCSave(RPTINDEX, llRowNo) = slStr
            gSetShow pbcRateCard, slStr, tmRCCtrls(RPTINDEX)
            smRCShow(RPTINDEX, llRowNo) = tmRCCtrls(RPTINDEX).sShow
            slStr = Trim$(Str$(tmRifRec(llRowNo).tRif.iSort))
            smRCSave(SORTINDEX, llRowNo) = slStr
            gSetShow pbcRateCard, slStr, tmRCCtrls(SORTINDEX)
            smRCShow(SORTINDEX, llRowNo) = tmRCCtrls(SORTINDEX).sShow
        Else
            slStr = ""
            smRCSave(BASEINDEX, llRowNo) = slStr
            gSetShow pbcRateCard, slStr, tmRCCtrls(BASEINDEX)
            smRCShow(BASEINDEX, llRowNo) = tmRCCtrls(BASEINDEX).sShow
            slStr = ""
            smRCSave(RPTINDEX, llRowNo) = slStr
            gSetShow pbcRateCard, slStr, tmRCCtrls(RPTINDEX)
            smRCShow(RPTINDEX, llRowNo) = tmRCCtrls(RPTINDEX).sShow
            slStr = ""
            smRCSave(SORTINDEX, llRowNo) = slStr
            gSetShow pbcRateCard, slStr, tmRCCtrls(SORTINDEX)
            smRCShow(SORTINDEX, llRowNo) = tmRCCtrls(SORTINDEX).sShow
        End If
    Next llRowNo
    mSetDefInSave   'Set defaults for extra row
    mGetShowDates
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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

    imPopReqd = False
    'ilRet = gPopRateCardBox(RateCard, lmNowDate, cbcSelect, lbcRateCard, -1)
    'ilRet = gPopRateCardBox(RateCard, lmNowDate, cbcSelect, tmRateCard(), smRateCardTag, -1)
    ilRet = gPopRateCardBox(RateCard, 0, cbcSelect, tmRateCard(), smRateCardTag, -1)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopRateCardBox)", RateCard
        On Error GoTo 0
        If (tgUrf(0).iCode > 2) And ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB) And (tgUrf(0).iMnfHubCode > 0) Then
            imAdjIndex = 0
        Else
            imAdjIndex = 1
            cbcSelect.AddItem "[New]", 0  'Force as first item on list
        End If
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
'*      Procedure Name:mRCEnableBox                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mRCEnableBox(ilBoxNo As Integer)
'
'   mRCEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilCode As Integer
    Dim llRowNo As Long
    
    If (ilBoxNo < imLBRCCtrls) Or (ilBoxNo > UBound(tmRCCtrls)) Then
        Exit Sub
    End If

    If (lmRCRowNo < vbcRateCard.Value) Or (lmRCRowNo >= vbcRateCard.Value + vbcRateCard.LargeChange + 1) Then
        'mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacRCFrame.Visible = False
        lacDPFrame.Visible = False
        Exit Sub
    End If
    lacRCFrame.Move 0, tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
    lacRCFrame.Visible = True
    lacDPFrame.Move 0, tmDPCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
    lacDPFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcRateCard.Top + tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHINDEX 'Vehicle
            'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
            lbcVehicle.Clear
            mVehPop lbcVehicle
            lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 10)
            edcDropDown.Width = tmRCCtrls(VEHINDEX).fBoxW
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(VEHINDEX).fBoxX, tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            'Vehicle default should be set in mSetDefInSave for all lines except first
            smSvVehName = Trim$(smRCSave(VEHINDEX, lmRCRowNo))
            imChgMode = True
            If Trim$(smRCSave(VEHINDEX, lmRCRowNo)) = "" Then
                If lmRCRowNo > 1 Then
                    gFindMatch Trim$(smRCSave(VEHINDEX, lmRCRowNo - 1)), 0, lbcVehicle
                    If gLastFound(lbcVehicle) >= 0 Then
                        smRCSave(VEHINDEX, lmRCRowNo) = Trim$(smRCSave(VEHINDEX, lmRCRowNo - 1))
                    End If
                End If
                If Trim$(smRCSave(VEHINDEX, lmRCRowNo)) = "" Then
                    gFindMatch smDefVehicle, 0, lbcVehicle
                    If gLastFound(lbcVehicle) >= 0 Then
                        smRCSave(VEHINDEX, lmRCRowNo) = smDefVehicle
                    Else
                        gFindMatch sgUserDefVehicleName, 0, lbcVehicle
                        If gLastFound(lbcVehicle) >= 0 Then
                            smRCSave(VEHINDEX, lmRCRowNo) = sgUserDefVehicleName
                        Else
                            smRCSave(VEHINDEX, lmRCRowNo) = ""
                        End If
                    End If
                End If
            End If
            gFindMatch Trim$(smRCSave(VEHINDEX, lmRCRowNo)), 0, lbcVehicle
            If gLastFound(lbcVehicle) >= 0 Then
                lbcVehicle.ListIndex = gLastFound(lbcVehicle)
            Else
                lbcVehicle.ListIndex = -1
            End If
            imComboBoxIndex = lbcVehicle.ListIndex
            If lbcVehicle.ListIndex >= 0 Then
                edcDropDown.Text = lbcVehicle.List(lbcVehicle.ListIndex)
            Else
                edcDropDown.Text = ""
            End If
            imChgMode = False
            If lmRCRowNo - vbcRateCard.Value <= vbcRateCard.LargeChange \ 2 Then
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case DAYPARTINDEX 'Daypart name index
            mDPNameRowPop
            lbcDPNameRow.Height = gListBoxHeight(lbcDPNameRow.ListCount, 10)
            edcDropDown.Width = tmRCCtrls(DAYPARTINDEX).fBoxW
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(DAYPARTINDEX).fBoxX, tmRCCtrls(DAYPARTINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            'Name default should be set in mSetDefInSave
            imChgMode = True
            gFindMatch Trim$(smRCSave(DAYPARTINDEX, lmRCRowNo)), 0, lbcDPNameRow
            If gLastFound(lbcDPNameRow) >= 0 Then
                lbcDPNameRow.ListIndex = gLastFound(lbcDPNameRow)
                edcDropDown.Text = lbcDPNameRow.List(lbcDPNameRow.ListIndex)
            Else
                If Trim$(smRCSave(DAYPARTINDEX, lmRCRowNo)) = "" Then
                    If lbcDPNameRow.ListCount <= 1 Then
                        lbcDPNameRow.ListIndex = 0
                    Else
                        lbcDPNameRow.ListIndex = 1
                    End If
                    edcDropDown.Text = lbcDPNameRow.List(lbcDPNameRow.ListIndex)
                Else
                    lbcDPNameRow.ListIndex = -1
                    edcDropDown.Text = Trim$(smRCSave(DAYPARTINDEX, lmRCRowNo))
                End If
            End If
            imChgMode = False
            If lmRCRowNo - vbcRateCard.Value <= vbcRateCard.LargeChange \ 2 Then
                lbcDPNameRow.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcDPNameRow.Move edcDropDown.Left, edcDropDown.Top - lbcDPNameRow.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        'Case DOLLARINDEX
        '    edcDropDown.Width = tmRCCtrls(DOLLARINDEX).fBoxW
        '    edcDropDown.MaxLength = 6
        '    gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(DOLLARINDEX).fBoxX, tmRCCtrls(DOLLARINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
        '    edcDropDown.Text = smRCSave(BASEINDEX, lmRCRowNo)
        '    edcDropDown.Visible = True  'Set visibility
        '    edcDropDown.SetFocus
        'Case PCTINVINDEX
        '    edcDropDown.Width = tmRCCtrls(PCTINVINDEX).fBoxW
        '    edcDropDown.MaxLength = 6
        '    gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(PCTINVINDEX).fBoxX, tmRCCtrls(PCTINVINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
        '    edcDropDown.Text = smRCSave(RPTINDEX, lmRCRowNo)
        '    edcDropDown.Visible = True  'Set visibility
        '    edcDropDown.SetFocus
        Case CPMINDEX
            If imRCSave(17, lmRCRowNo) = 1 And imRCSave(16, lmRCRowNo) = 0 Then
                edcDropDown.Width = tmRCCtrls(CPMINDEX).fBoxW
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(CPMINDEX).fBoxX, tmRCCtrls(CPMINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(smRCSave(CPMINDEX, lmRCRowNo))
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
            End If
        Case ACQUISITIONINDEX
            edcDropDown.Width = tmRCCtrls(ACQUISITIONINDEX).fBoxW
            edcDropDown.MaxLength = 11
            gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(ACQUISITIONINDEX).fBoxX, tmRCCtrls(ACQUISITIONINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = Trim$(smRCSave(ACQUISITIONINDEX, lmRCRowNo))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case BASEINDEX
            ' LB 02/10/21
            'If imRCSave(17, lmRCRowNo) = 0 Then
                If (Trim$(smRCSave(BASEINDEX, lmRCRowNo)) <> "Y") And (Trim$(smRCSave(BASEINDEX, lmRCRowNo)) <> "N") Then
                    smRCSave(BASEINDEX, lmRCRowNo) = Trim$(smDPShow(DPBASEINDEX, lmRCRowNo))    'Yes
                    imRifChg = True
                End If
                pbcYN.Width = tmRCCtrls(ilBoxNo).fBoxW
                gMoveTableCtrl pbcRateCard, pbcYN, tmRCCtrls(BASEINDEX).fBoxX, tmRCCtrls(BASEINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
            'End If
        Case RPTINDEX
            ' LB 02/10/21
            ' If imRCSave(17, lmRCRowNo) = 0 Then
                If (Trim$(smRCSave(RPTINDEX, lmRCRowNo)) <> "Y") And (Trim$(smRCSave(RPTINDEX, lmRCRowNo)) <> "N") Then
                    gFindMatch Trim$(smRCSave(DAYPARTINDEX, lmRCRowNo)), 0, lbcDPName
                    ilIndex = gLastFound(lbcDPName)
                    slNameCode = lbcDPNameCode.List(ilIndex)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilCode = Val(Trim$(slCode))
                    For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        If ilCode = tgMRdf(ilLoop).iCode Then
                            smRCSave(RPTINDEX, lmRCRowNo) = tgMRdf(ilLoop).sReport    'Yes
                            imRifChg = True
                            Exit For
                        End If
                    Next ilLoop
                End If
                pbcYN.Width = tmRCCtrls(ilBoxNo).fBoxW
                gMoveTableCtrl pbcRateCard, pbcYN, tmRCCtrls(RPTINDEX).fBoxX, tmRCCtrls(RPTINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
                pbcYN_Paint
                pbcYN.Visible = True
                pbcYN.SetFocus
             ' End If
        Case SORTINDEX
            If Trim$(smRCSave(SORTINDEX, lmRCRowNo)) = "" Then
                gFindMatch Trim$(smRCSave(DAYPARTINDEX, lmRCRowNo)), 0, lbcDPName
                ilIndex = gLastFound(lbcDPName)
                slNameCode = lbcDPNameCode.List(ilIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilCode = Val(Trim$(slCode))
                For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    If ilCode = tgMRdf(ilLoop).iCode Then
                        smRCSave(SORTINDEX, lmRCRowNo) = Trim$(Str$(tgMRdf(ilLoop).iSortCode))    'Yes
                        imRifChg = True
                        Exit For
                    End If
                Next ilLoop
            End If
            edcDropDown.Width = tmRCCtrls(SORTINDEX).fBoxW
            edcDropDown.MaxLength = 2
            gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(SORTINDEX).fBoxX, tmRCCtrls(SORTINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
            edcDropDown.Text = Trim$(smRCSave(SORTINDEX, lmRCRowNo))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case DOLLAR1INDEX
            ' LB 02/11/21
            'mSetPvfType llRowNo
             
            If (imRCSave(16, lmRCRowNo) = 0) Then 'avoid for CPM packages
                bmInStdPrice = True
                mGetStdPkgPrice True
                edcDropDown.Width = tmRCCtrls(DOLLAR1INDEX).fBoxW
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(DOLLAR1INDEX).fBoxX, tmRCCtrls(DOLLAR1INDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(Str$(lmRCSave(DOLLAR1INDEX - DOLLAR1INDEX + 1, lmRCRowNo)))
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
            End If
        Case DOLLAR2INDEX
            ' LB 02/10/21
            If imRCSave(16, lmRCRowNo) = 0 And rbcType(3).Value = False Then
                edcDropDown.Width = tmRCCtrls(DOLLAR2INDEX).fBoxW
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(DOLLAR2INDEX).fBoxX, tmRCCtrls(DOLLAR2INDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(Str$(lmRCSave(DOLLAR2INDEX - DOLLAR1INDEX + 1, lmRCRowNo)))
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
            End If
        Case DOLLAR3INDEX
            ' LB 02/10/21
            If imRCSave(16, lmRCRowNo) = 0 And rbcType(3).Value = False Then
                edcDropDown.Width = tmRCCtrls(DOLLAR3INDEX).fBoxW
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(DOLLAR3INDEX).fBoxX, tmRCCtrls(DOLLAR3INDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(Str$(lmRCSave(DOLLAR3INDEX - DOLLAR1INDEX + 1, lmRCRowNo)))
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
             End If
        Case DOLLAR4INDEX
            ' LB 02/10/21
            If imRCSave(16, lmRCRowNo) = 0 And rbcType(3).Value = False Then
                edcDropDown.Width = tmRCCtrls(DOLLAR4INDEX).fBoxW
                edcDropDown.MaxLength = 10
                gMoveTableCtrl pbcRateCard, edcDropDown, tmRCCtrls(DOLLAR4INDEX).fBoxX, tmRCCtrls(DOLLAR4INDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15)
                edcDropDown.Text = Trim$(Str$(lmRCSave(DOLLAR4INDEX - DOLLAR1INDEX + 1, lmRCRowNo)))
                edcDropDown.Visible = True  'Set visibility
                edcDropDown.SetFocus
            End If
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mRCSetShow                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Public Sub mRCSetShow(ilBoxNo As Integer)
'
'   mRCSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim slStart As String
    Dim slEnd As String
    Dim slDate As String
    'Dim slNoWks As String
    Dim ilNoWks As Integer
    Dim ilIndex As Integer
    Dim slMultiTimes As String
    Dim ilDay As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDollar As String
    Dim llDollar As Long
    Dim ilCode As Integer
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    If (ilBoxNo < imLBRCCtrls) Or (ilBoxNo > UBound(tmRCCtrls)) Then
        Exit Sub
    End If
    
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHINDEX 'Vehicle
            lbcVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
            smRCShow(VEHINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            smDPShow(VEHINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            If smSvVehName <> edcDropDown.Text Then
                smRCSave(VEHINDEX, lmRCRowNo) = lbcVehicle.List(lbcVehicle.ListIndex)
                'Moved to Enable Dollar1 field
                'mGetStdPkgPrice True
                imRifChg = True
            End If
        Case DAYPARTINDEX 'Name
            lbcDPNameRow.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            'slStr = edcDropDown.Text
            'gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
            'smRCShow(DAYPARTINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            'smDPShow(DAYPARTINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            If Trim$(smRCSave(DAYPARTINDEX, lmRCRowNo)) <> edcDropDown.Text Then
                imRifChg = True
                smRCSave(DAYPARTINDEX, lmRCRowNo) = edcDropDown.Text
                gFindMatch Trim$(smRCSave(DAYPARTINDEX, lmRCRowNo)), 0, lbcDPName
                ilIndex = gLastFound(lbcDPName)
                If ilIndex >= 0 Then
                    slNameCode = lbcDPNameCode.List(ilIndex)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilCode = Val(slCode)
                    For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        If ilCode = tgMRdf(ilLoop).iCode Then
                            tmRdf = tgMRdf(ilLoop)
                            mSetDPShow lmRCRowNo
                            Exit For
                        End If
                    Next ilLoop
                Else
                    slStr = ""
                    gSetShow pbcRateCard, slStr, tmRCCtrls(DAYPARTINDEX)
                    smRCShow(DAYPARTINDEX, lmRCRowNo) = tmRCCtrls(DAYPARTINDEX).sShow
                    slStr = ""
                    gSetShow pbcDaypart, slStr, tmDPCtrls(DAYPARTINDEX)
                    smDPShow(DAYPARTINDEX, lmRCRowNo) = tmDPCtrls(DAYPARTINDEX).sShow
                    slStr = ""
                    gSetShow pbcDaypart, slStr, tmDPCtrls(TIMESINDEX)
                    smDPShow(TIMESINDEX, lmRCRowNo) = tmDPCtrls(TIMESINDEX).sShow & slMultiTimes
                    For ilDay = 1 To 7 Step 1
                        slStr = ""
                        gSetShow pbcDaypart, slStr, tmDPCtrls(DAYINDEX + ilDay - 1)
                        smDPShow(DAYINDEX + ilDay - 1, lmRCRowNo) = tmDPCtrls(DAYINDEX + ilDay - 1).sShow
                    Next ilDay
                    slStr = ""
                    gSetShow pbcDaypart, slStr, tmDPCtrls(AVAILINDEX)
                    smDPShow(AVAILINDEX, lmRCRowNo) = tmDPCtrls(AVAILINDEX).sShow
                    slStr = ""
                    gSetShow pbcDaypart, slStr, tmDPCtrls(HRSINDEX)
                    smDPShow(HRSINDEX, lmRCRowNo) = tmDPCtrls(HRSINDEX).sShow
                End If
            End If
        'Case DOLLARINDEX
        '    edcDropDown.Visible = False
        '    slStr = edcDropDown.Text
        '    gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
        '    smRCShow(DOLLARINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
        '    If smRCSave(BASEINDEX, lmRCRowNo) <> edcDropDown.Text Then
        '        imRifChg = True
        '    End If
        '    smRCSave(BASEINDEX, lmRCRowNo) = edcDropDown.Text
        'Case PCTINVINDEX
        '    edcDropDown.Visible = False
        '    slStr = edcDropDown.Text
        '    gFormatStr slStr, FMTLEAVEBLANK + FMTPERCENTSIGN, 2, slStr
        '    gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
        '    smRCShow(PCTINVINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
        '    If smRCSave(RPTINDEX, lmRCRowNo) <> edcDropDown.Text Then
        '        imRifChg = True
        '    End If
        '    smRCSave(RPTINDEX, lmRCRowNo) = edcDropDown.Text
        Case CPMINDEX
            If imRCSave(17, lmRCRowNo) = 1 And imRCSave(16, lmRCRowNo) = 0 Then
                edcDropDown.Visible = False
                slStr = gDivStr(gMulStr(edcDropDown.Text, "1000.00"), "1000.00")
                gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
                If Trim$(smRCSave(CPMINDEX, lmRCRowNo)) <> edcDropDown.Text Then
                    imRifChg = True
                End If
                smRCShow(CPMINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
                smDPShow(CPMINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
                smRCSave(CPMINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            End If
        Case ACQUISITIONINDEX
            edcDropDown.Visible = False
            slStr = Format(Val(edcDropDown.Text), "#.00") '11/4/21 - JW - OK'd with Jason as ACQ is a Dollars Cost.
            If Val(edcDropDown.Text) = 0 Then slStr = ""
            gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
            smRCShow(ACQUISITIONINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            If smRCSave(ACQUISITIONINDEX, lmRCRowNo) <> edcDropDown.Text Then
                imRifChg = True
            End If
            smRCSave(ACQUISITIONINDEX, lmRCRowNo) = edcDropDown.Text
        Case BASEINDEX
                ' LB 02/10/21
                'If imRCSave(17, lmRCRowNo) = 0 Then
                pbcYN.Visible = False
                If Trim$(smRCSave(BASEINDEX, lmRCRowNo)) = "Y" Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
                gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
                smRCShow(BASEINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
                'End If
        Case RPTINDEX
            ' LB 02/10/21
            'If imRCSave(17, lmRCRowNo) = 0 Then
                pbcYN.Visible = False
                If Trim$(smRCSave(RPTINDEX, lmRCRowNo)) = "Y" Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
                gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
                smRCShow(RPTINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            'End If
        Case SORTINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
            smRCShow(SORTINDEX, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            If Trim$(smRCSave(SORTINDEX, lmRCRowNo)) <> edcDropDown.Text Then
                imRifChg = True
            End If
            smRCSave(SORTINDEX, lmRCRowNo) = edcDropDown.Text
        Case DOLLAR1INDEX, DOLLAR2INDEX, DOLLAR3INDEX, DOLLAR4INDEX
            ' LB 02/10/21
            
            If rbcType(3).Value = True And ilBoxNo > DOLLAR1INDEX Then
             Exit Sub
            End If
            
            
            If imRCSave(16, lmRCRowNo) = 0 Then
            edcDropDown.Visible = False
            slDollar = edcDropDown.Text
            llDollar = Val(slDollar)
            gFormatStr slDollar, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcRateCard, slStr, tmRCCtrls(ilBoxNo)
            smRCShow(ilBoxNo, lmRCRowNo) = tmRCCtrls(ilBoxNo).sShow
            If (lmRCSave(ilBoxNo - DOLLAR1INDEX + 1, lmRCRowNo) <> llDollar) Then
                'Recompute average and set weeks
                If tmPdGroups(1).iYear = tmPdGroups(ilBoxNo - DOLLAR1INDEX + 1).iYear Then
                    'smRCSave(TOTALINDEX, lmRCRowNo) = gSubStr(smRCSave(TOTALINDEX, lmRCRowNo), smRCSave(ilBoxNo, lmRCRowNo))
                    If tmPdGroups(1).iYear = tmPdGroups(ilBoxNo - DOLLAR1INDEX + 1).iYear Then
                        lmRCSave(TOTALINDEX - SORTINDEX, lmRCRowNo) = lmRCSave(TOTALINDEX - SORTINDEX, lmRCRowNo) - lmRCSave(ilBoxNo - DOLLAR1INDEX + 1, lmRCRowNo) * tmPdGroups(ilBoxNo - DOLLAR1INDEX + 1).iTrueNoWks
                    End If
                    lmRCSave(ilBoxNo - DOLLAR1INDEX + 1, lmRCRowNo) = llDollar
                    If tmPdGroups(1).iYear = tmPdGroups(ilBoxNo - DOLLAR1INDEX + 1).iYear Then
                        lmRCSave(TOTALINDEX - SORTINDEX, lmRCRowNo) = lmRCSave(TOTALINDEX - SORTINDEX, lmRCRowNo) + llDollar * tmPdGroups(ilBoxNo - DOLLAR1INDEX + 1).iTrueNoWks
                        If rbcShow(0).Value Then    'Corporate
                            slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                            slStart = gObtainStartCorp(slDate, True)
                            'slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                            'slEnd = gObtainEndCorp(slDate, True)
                            slDate = slStart
                            For ilLoop = 1 To 12 Step 1
                                slEnd = gObtainEndCorp(slDate, True)
                                slDate = gIncOneDay(slEnd)
                            Next ilLoop
                        Else
                            slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                            slStart = gObtainStartStd(slDate)
                            slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                            slEnd = gObtainEndStd(slDate)
                        End If
                        'slNoWks = Trim$(Str$((gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7))
                        ilNoWks = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
                        If llDollar > 0 Then
                            imRCSave(ilBoxNo - DOLLAR1INDEX + 1, lmRCRowNo) = imRCSave(ilBoxNo - DOLLAR1INDEX + 5, lmRCRowNo)
                        Else
                            imRCSave(ilBoxNo - DOLLAR1INDEX + 1, lmRCRowNo) = 0
                        End If
                        ilNoWks = imRCSave(2, lmRCRowNo) + imRCSave(1, lmRCRowNo) + imRCSave(3, lmRCRowNo) + imRCSave(4, lmRCRowNo)
                        'smRCSave(AVGINDEX, lmRCRowNo) = gDivStr(smRCSave(TOTALINDEX, lmRCRowNo), slNoWks)
                        If ilNoWks > 0 Then
                            lmRCSave(AVGINDEX - SORTINDEX, lmRCRowNo) = lmRCSave(TOTALINDEX - SORTINDEX, lmRCRowNo) / ilNoWks
                        Else
                            lmRCSave(AVGINDEX - SORTINDEX, lmRCRowNo) = 0
                        End If
                        slStr = Trim$(Str$(lmRCSave(AVGINDEX - SORTINDEX, lmRCRowNo)))
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcRateCard, slStr, tmRCCtrls(AVGINDEX)
                        smRCShow(AVGINDEX, lmRCRowNo) = tmRCCtrls(AVGINDEX).sShow
                    End If
                Else
                    lmRCSave(ilBoxNo - DOLLAR1INDEX + 1, lmRCRowNo) = llDollar
                End If
                If (Not bmInDupicate) Then
                    mSetPrice ilBoxNo - DOLLAR1INDEX + 1, lmRCRowNo, slDollar
                End If
                mGetAvg tmPdGroups(1).iYear, lmRCRowNo
                mGetShowPrices lmRCRowNo    'Set color flag
                imRifChg = True
                'pbcRateCard.Cls    'Use LIGHTYELLOW instead of blinking
                If Not bmInImportPrice Then
                    pbcRateCard_Paint
                End If
            End If
            imTabDirection = 0
            bmInStdPrice = False
         End If
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer) As Integer
'
'   iRet = mReadRec(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    slNameCode = tmRateCard(ilSelectIndex - imAdjIndex).sKey    'lbcRateCard.List(ilSelectIndex - 1)
    ilRet = gParseItem(slNameCode, 3, "\", slCode)
    On Error GoTo mReadRecErr
    gCPErrorMsg ilRet, "mReadRec (gParseItem field 3)", RateCard
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmRcfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmRcf, tgRcfI, imRcfRecLen, tmRcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", RateCard
    On Error GoTo 0
    mTestCorpYear tgRcfI.iYear
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRifRec                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************

'**********************************************************
'Rules to use for active Podcast Medium Vehicles (1/20/22)
'**********************************************************
'Rate card Screen: (you are here)
'**********************************************************
'  Always show active podcast medium vehicles
'  sGMedium = "P"
'**********************************************************
'Std Pkg Screen:
'**********************************************************
'  Show podcast medium vehicles when vehicle has programming
'    Vpf.sGMedium = "P" = PodCast
'    LTF_Lbrary_Title WHERE LtfVefCode
'**********************************************************
'CPM Pkg Screen:
'**********************************************************
'  Show podcast medium vehicles when it has an ad server..
'  vendor defined in vehicle options
'    Vpf.sGMedium = "P" (PodCast)
'
'    pvfType="C" =Podcast Ad Server (CPM only)
'    Vff.iAvfCode <> 0 (has Ad Server)
'
'    CpmPkg button visible when sFeatures8=PODADSERVER
'**********************************************************
Private Function mReadRifRec(ilRcfCode As Integer, ilAllYears As Integer) As Integer
'
'   iRet = mReadRifRec (ilRcfCode, ilAllYears)
'   Where:
'       ilAllYears(I)-True get all years; False get last year only
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim llUpper As Long
    Dim llLkUpper As Long
    Dim ilLoop As Integer
    Dim llLoop As Long
    Dim llLp1 As Long
    Dim llTemp As Long
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llSec As Long
    Dim llMin As Long
    Dim llHour As Long
    Dim llTime As Long
    Dim ilIndex As Integer
    Dim ilTime As Integer
    Dim slTime As String
    Dim ilVef As Integer
    Dim ilVehFound As Integer
    Dim slStr1 As String
    Dim slGpSort As String
    Dim ilSortVehbyGroup As Integer 'True=Sort vehicle by groups; False=Sort by name only
    Dim ilRdf As Integer
    Dim llRif As Long
    Dim ilRifFound As Integer
    Dim tlRifRec As RIFREC
    ReDim tmRifRec(0 To 1) As RIFREC
    ReDim tmLkRifRec(0 To 1) As RIFREC
    ReDim tmTrashRifRec(0 To 1) As RIFREC

    ReDim lmAutoDelRif(0 To 0) As Long
    Dim ilRcfYear As Integer
    Dim llDupl As Long
    Dim ilRemoved As Integer
    Dim ilPrice1 As Integer
    Dim ilPrice2 As Integer
    Dim slKey As String
    Dim tlRif As RIF
    Dim llTest As Long
    Dim llIndex As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim ilFound As Integer
    Dim ilTest As Integer
    Dim ilClf As Integer
    Dim ilCheckModel As Integer
    Dim sVehType As String
    'Obtain year to retain
    ilRcfYear = -1
    For ilLoop = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
        If ilRcfCode = tgMRcf(ilLoop).iCode Then
            ilRcfYear = tgMRcf(ilLoop).iYear
            Exit For
        End If
    Next ilLoop

    ilSortVehbyGroup = False
    llUpper = UBound(tmRifRec)
    btrExtClear hmRif   'Clear any previous extend operation
    ilExtLen = Len(tmRifRec(1).tRif)  'Extract operation record size
    llUpper = 0
    ReDim tmTempRifRec(0 To 0) As RIFREC
'    tmRifSrchKey.iCode = ilRcfCode
'    ilRet = btrGetGreaterOrEqual(hmRif, tmTempRifRec(1).tRif, imRifRecLen, tmRifSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'    If (tmTempRifRec(1).tRif.iRcfCode = ilRcfCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then
'        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
'        Call btrExtSetBounds(hmRif, llNoRec, -1, "UC", "RIF", "") '"EG") 'Set extract limits (all records)
'        ilOffset = gFieldOffset("Rif", "RifRcfCode")
'        ilRet = btrExtAddLogicConst(hmRif, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tmRifSrchKey, 2)
'        On Error GoTo mReadRifRecErr
'        gBtrvErrorMsg ilRet, "mReadRifRec (btrExtAddLogicConst):" & "Rif.Btr", RateCard
'        On Error GoTo 0
'        ilRet = btrExtAddField(hmRif, 0, ilExtLen) 'Extract the whole record
'        On Error GoTo mReadRifRecErr
'        gBtrvErrorMsg ilRet, "mReadRifRec (btrExtAddField):" & "Rif.Btr", RateCard
'        On Error GoTo 0
'        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
'        ilRet = btrExtGetNext(hmRif, tmTempRifRec(llUpper).tRif, ilExtLen, tmTempRifRec(llUpper).lRecPos)
'        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
    ilCheckModel = False
    ReDim tmRCModelInfo(0 To 0) As RCMODELINFO
    If (igRCMode = 0) And (sgRCModelDate <> "") And (imSelectedIndex = 0) And (imAdjIndex = 1) Then
        slCntrStatus = "HO"
        slCntrType = ""
        ilHOType = 2
        slStartDate = sgRCModelDate
        slEndDate = ""
        sgCntrForDateStamp = ""
        ilRet = gObtainCntrForDate(RateCard, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
        If (ilRet = CP_MSG_NOPOPREQ) Or (ilRet = CP_MSG_NONE) Then
            ilCheckModel = True
            'List of allowed vehicles and dayparts
            For ilLoop = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
                ilRet = gObtainChfClf(hmCHF, hmClf, tmChfAdvtExt(ilLoop).lCode, False, tmChf, tmClfModel())
                If ilRet Then
                    For ilClf = LBound(tmClfModel) To UBound(tmClfModel) - 1 Step 1
                        ilFound = False
                        For ilTest = 0 To UBound(tmRCModelInfo) - 1 Step 1
                            If (tmRCModelInfo(ilTest).iRdfCode = tmClfModel(ilClf).ClfRec.iRdfCode) Then
                                If (tmRCModelInfo(ilTest).iVefCode = tmClfModel(ilClf).ClfRec.iVefCode) Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilTest
                        If Not ilFound Then
                            tmRCModelInfo(UBound(tmRCModelInfo)).iRdfCode = tmClfModel(ilClf).ClfRec.iRdfCode
                            tmRCModelInfo(UBound(tmRCModelInfo)).iVefCode = tmClfModel(ilClf).ClfRec.iVefCode
                            ReDim Preserve tmRCModelInfo(0 To UBound(tmRCModelInfo) + 1) As RCMODELINFO
                        End If
                    Next ilClf
                End If
            Next ilLoop
        End If
    End If
    
    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
        If tgMRif(llRif).iRcfCode = ilRcfCode Then
            tmTempRifRec(llUpper).tRif = tgMRif(llRif)
            If ilCheckModel Then
                ilFound = False
                For ilTest = 0 To UBound(tmRCModelInfo) - 1 Step 1
                    If (tmRCModelInfo(ilTest).iRdfCode = tgMRif(llRif).iRdfCode) Then
                        If (tmRCModelInfo(ilTest).iVefCode = tgMRif(llRif).iVefCode) Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilTest
            Else
                ilFound = True
            End If
            If ilFound And (tgUrf(0).iCode > 2) And (tgUrf(0).iMnfHubCode > 0) Then
                If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
                    ilTest = gBinarySearchVef(tgMRif(llRif).iVefCode)
                    If ilTest <> -1 Then
                        If tgMVef(ilTest).iMnfHubCode <> tgUrf(0).iMnfHubCode Then
                            ilFound = False
                        End If
                    End If
                End If
            End If
            If ilFound Then
                slStr = ""
                ilVehFound = False
''                slRecCode = Trim$(Str$(tmTempRifRec(llUpper).tRif.iVefCode))
''                If Val(slRecCode) <> tmVef.iCode Then
''                    tmVefSrchKey.iCode = Val(slRecCode) 'ilCode
''                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
''                Else
''                    ilRet = BTRV_ERR_NONE
''                End If
''                If (ilRet = BTRV_ERR_NONE) And (tmVef.sState <> "D") Then

                 ilVef = gBinarySearchVef(tmTempRifRec(llUpper).tRif.iVefCode)
                 If ilVef <> -1 Then
                    'Check if Item should be removed
                    'Two cases- Different year (only going to have items in same year
                    '           Duplicate records
                    '7/15/05: Jim- remove dormant vehicles
                    ilRemoved = False
                    sVehType = ""
    
                    'TTP 10388 - Rate card screen: if "ad server" site option setting is not enabled, a "podcast" medium type vehicle does not appear on the rate card screen
                    'sVehType = mGetVehicleMedium(tgMRif(llRif).iVefCode)
                    'If sVehType = "P" And (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER Then
                        'L.Bianchi
                        'lmAutoDelRif(UBound(lmAutoDelRif)) = tmTempRifRec(llUpper).tRif.lCode
                        'ReDim Preserve lmAutoDelRif(0 To UBound(lmAutoDelRif) + 1) As Long
                        'ilRemoved = True
                    'ElseIf
                    If ((ilRcfYear <> -1) And (tmTempRifRec(llUpper).tRif.iYear <> ilRcfYear)) Or (tgMVef(ilVef).sState = "D") Then
                        lmAutoDelRif(UBound(lmAutoDelRif)) = tmTempRifRec(llUpper).tRif.lCode
                        ReDim Preserve lmAutoDelRif(0 To UBound(lmAutoDelRif) + 1) As Long
                        ilRemoved = True
                    Else
                        'Check if duplicate
                        slStr = Trim$(Str$(tmTempRifRec(llUpper).tRif.iVefCode))
                        Do While Len(slStr) < 5
                            slStr = "0" & slStr
                        Loop
                        slStr1 = Trim$(Str$(tmTempRifRec(llUpper).tRif.iRdfCode))
                        Do While Len(slStr1) < 5
                            slStr1 = "0" & slStr1
                        Loop
                        slKey = slKey & slStr1
                        slStr1 = Trim$(Str$(tmTempRifRec(llUpper).tRif.iYear))
                        Do While Len(slStr1) < 4
                            slStr1 = "0" & slStr1
                        Loop
                        slKey = slKey & slStr1
                        For llDupl = 0 To UBound(tmTempRifRec) - 1 Step 1
                            If (tmTempRifRec(llUpper).tRif.iRdfCode = tmTempRifRec(llDupl).tRif.iRdfCode) And (tmTempRifRec(llUpper).tRif.iVefCode = tmTempRifRec(llDupl).tRif.iVefCode) Then
                                ilRemoved = True
                                'Remove one of the items (the one without price)
                                ilPrice1 = False
                                For ilIndex = LBound(tmTempRifRec(llUpper).tRif.lRate) To UBound(tmTempRifRec(llUpper).tRif.lRate) Step 1
                                    If tmTempRifRec(llUpper).tRif.lRate(ilIndex) > 0 Then
                                        ilPrice1 = True
                                    End If
                                Next ilIndex
                                ilPrice2 = False
                                For ilIndex = LBound(tmTempRifRec(llDupl).tRif.lRate) To UBound(tmTempRifRec(llDupl).tRif.lRate) Step 1
                                    If tmTempRifRec(llDupl).tRif.lRate(ilIndex) > 0 Then
                                        ilPrice2 = True
                                    End If
                                Next ilIndex
                                If (ilPrice1) And (Not ilPrice2) Then
                                    lmAutoDelRif(UBound(lmAutoDelRif)) = tmTempRifRec(llDupl).tRif.lCode
                                    ReDim Preserve lmAutoDelRif(0 To UBound(lmAutoDelRif) + 1) As Long
                                    tmTempRifRec(llDupl) = tmTempRifRec(llUpper)
                                    tmTempRifRec(llDupl).sKey = slKey
                                ElseIf (Not ilPrice1) And (ilPrice2) Then
                                    lmAutoDelRif(UBound(lmAutoDelRif)) = tmTempRifRec(llUpper).tRif.lCode
                                    ReDim Preserve lmAutoDelRif(0 To UBound(lmAutoDelRif) + 1) As Long
                                Else
                                    If tmTempRifRec(llUpper).tRif.lCode > tmTempRifRec(llDupl).tRif.lCode Then
                                        lmAutoDelRif(UBound(lmAutoDelRif)) = tmTempRifRec(llDupl).tRif.lCode
                                        ReDim Preserve lmAutoDelRif(0 To UBound(lmAutoDelRif) + 1) As Long
                                        tmTempRifRec(llDupl) = tmTempRifRec(llUpper)
                                        tmTempRifRec(llDupl).sKey = slKey
                                    Else
                                        lmAutoDelRif(UBound(lmAutoDelRif)) = tmTempRifRec(llUpper).tRif.lCode
                                        ReDim Preserve lmAutoDelRif(0 To UBound(lmAutoDelRif) + 1) As Long
                                    End If
                                End If
                                Exit For
                            End If
                        Next llDupl
                    End If

                    If (tgMVef(ilVef).sState <> "D") And (Not ilRemoved) Then
'Move to Duplicate test area
'                        slStr = Trim$(Str$(tmTempRifRec(llUpper).tRif.iVefCode))
'                        Do While Len(slStr) < 5
'                            slStr = "0" & slStr
'                        Loop
'                        slStr1 = Trim$(Str$(tmTempRifRec(llUpper).tRif.iRdfcode))
'                        Do While Len(slStr1) < 5
'                            slStr1 = "0" & slStr1
'                        Loop
'                        slStr = slStr & slStr1
'                        slStr1 = Trim$(Str$(tmTempRifRec(llUpper).tRif.iYear))
'                        Do While Len(slStr1) < 4
'                            slStr1 = "0" & slStr1
'                        Loop
'                        slStr = slStr & slStr1
                        tmTempRifRec(llUpper).sKey = slKey
                        tmTempRifRec(llUpper).iStatus = 1
                        tmTempRifRec(llUpper).lLkYear = 0
                        llUpper = llUpper + 1
                        ReDim Preserve tmTempRifRec(0 To llUpper) As RIFREC
                    End If
                End If
            End If
        End If
'    End If
    Next llRif
    'If llUpper > 1 Then
    '    ArraySortTyp fnAV(tmTempRifRec(), 1), UBound(tmTempRifRec) - 1, 0, LenB(tmTempRifRec(1)), 0, LenB(tmTempRifRec(1).sKey), 0
    'End If
    If llUpper > 1 Then
        ArraySortTyp fnAV(tmTempRifRec(), 0), UBound(tmTempRifRec), 0, LenB(tmTempRifRec(0)), 0, LenB(tmTempRifRec(0).sKey), 0
    End If
    'Build sort key
    For llTemp = LBound(tmTempRifRec) To UBound(tmTempRifRec) - 1 Step 1
        slStr = ""
       
        ilVehFound = False
        ilVef = gBinarySearchVef(tmTempRifRec(llTemp).tRif.iVefCode)
        If ilVef <> -1 Then
            ilVehFound = True
            If ilSortVehbyGroup Then
                slStr = Trim$(Str$(tgMVef(ilVef).iSort))    'tmVef.iSort))
                Do While Len(slStr) < 5
                    slStr = "0" & slStr
                Loop
            Else
                slStr = "00000"
                ' LB 02/10/21
                ' If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
                '    sVehType = mGetVehicleMedium(tmTempRifRec(llTemp).tRif.iVefCode)
                '    If sVehType = "P" Then
                '        slStr = "00001"
                '    Else
                '        slStr = "00000"
                '    End If
                ' End If
            End If
            slStr = slStr & tgMVef(ilVef).sName
            If ilSortVehbyGroup Then
                tmMnfSrchKey.iCode = tgMVef(ilVef).iOwnerMnfCode
                If tmMnf.iCode <> tmMnfSrchKey.iCode Then
                    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmMnf.iGroupNo = 999
                    End If
                End If
                slGpSort = Trim$(Str$(tmMnf.iGroupNo))
                Do While Len(slGpSort) < 3
                    slGpSort = "0" & slGpSort
                Loop
            Else
                slGpSort = "000"
            End If
            slStr = slGpSort & slStr
        End If
        If ilVehFound Then
            For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                If tmTempRifRec(llTemp).tRif.iRdfCode = tgMRdf(ilRdf).iCode Then
                    ilRet = BTRV_ERR_NONE
                    tmRdf = tgMRdf(ilRdf)
                    If (tmTempRifRec(llTemp).tRif.sBase = "Y") Or (tmTempRifRec(llTemp).tRif.sBase = "N") Then
                        slStr1 = Trim$(Str$(tmTempRifRec(llTemp).tRif.iSort))  'tmRdf.iSortCode))
                    Else
                        slStr1 = Trim$(Str$(tmRdf.iSortCode))
                    End If
                    Do While Len(slStr1) < 5
                        slStr1 = "0" & slStr1
                    Loop
                    slStr = slStr & slStr1
                    ilTime = 1
                    For ilLoop = 1 To imMaxTDRows + 1 Step 1 'Row
                        If (tmRdf.iStartTime(0, ilLoop - 1) <> 1) Or (tmRdf.iStartTime(1, ilLoop - 1) <> 0) Then
                            ilTime = ilLoop
                            For ilIndex = 1 To 7 Step 1
                                If tmRdf.sWkDays(ilTime - 1, ilIndex - 1) <> "Y" Then
                                    slStr = slStr & "B"
                                Else
                                    slStr = slStr & "A"
                                End If
                            Next ilIndex
                            Exit For
                        End If
                    Next ilLoop
                    llSec = tmRdf.iStartTime(0, ilTime - 1) \ 256 'Obtain seconds
                    llMin = tmRdf.iStartTime(1, ilTime - 1) And &HFF 'Obtain Minutes
                    llHour = tmRdf.iStartTime(1, ilTime - 1) \ 256 'Obtain month
                    llTime = 3600 * llHour + 60 * llMin + llSec
                    slTime = Trim$(Str$(llTime))
                    Do While (Len(slTime) < 5)
                        slTime = "0" & slTime
                    Loop
                    slStr = slStr & slTime
                    tmTempRifRec(llTemp).sKey = slStr
                    Exit For
                End If
            Next ilRdf
        End If
    Next llTemp
    If UBound(tmTempRifRec) >= 1 Then
        ReDim tmRifRec(0 To 2) As RIFREC
    End If
    tmRifRec(1) = tmTempRifRec(0)
    llLoop = 1
    llUpper = 2
    For llTemp = LBound(tmTempRifRec) + 1 To UBound(tmTempRifRec) - 1 Step 1
        tmRifRec(llUpper) = tmTempRifRec(llTemp)
        ilRifFound = False
        If (tmRifRec(llLoop).tRif.iVefCode = tmRifRec(llUpper).tRif.iVefCode) And (tmRifRec(llLoop).tRif.iRdfCode = tmRifRec(llUpper).tRif.iRdfCode) Then
            If ilAllYears Then
                llLkUpper = UBound(tmLkRifRec)
                tmRifRec(llUpper).iStatus = 1
                tmRifRec(llUpper).sKey = tmRifRec(llLoop).sKey
                'Get records in year order
                If tmRifRec(llUpper).tRif.iYear < tmRifRec(llLoop).tRif.iYear Then
                    'Swap records
                    tlRifRec = tmRifRec(llLoop)
                    tmRifRec(llLoop) = tmRifRec(llUpper)
                    tmRifRec(llLoop).lLkYear = tlRifRec.lLkYear
                    tmRifRec(llUpper) = tlRifRec
                    tmRifRec(llUpper).lLkYear = 0
                End If
                If tmRifRec(llLoop).lLkYear = 0 Then
                    tmLkRifRec(llLkUpper) = tmRifRec(llUpper)
                    tmLkRifRec(llLkUpper).lLkYear = 0
                    tmRifRec(llLoop).lLkYear = llLkUpper
                    ReDim Preserve tmLkRifRec(0 To llLkUpper + 1) As RIFREC
                Else
                    llLp1 = tmRifRec(llLoop).lLkYear
                    Do
                        If tmRifRec(llUpper).tRif.iYear < tmLkRifRec(llLp1).tRif.iYear Then
                            'Swap records
                            tlRifRec = tmLkRifRec(llLp1)
                            tmLkRifRec(llLp1) = tmRifRec(llUpper)
                            tmLkRifRec(llLp1).lLkYear = tlRifRec.lLkYear
                            tmRifRec(llUpper) = tlRifRec
                            tmRifRec(llUpper).lLkYear = 0
                        Else
                            If tmLkRifRec(llLp1).lLkYear = 0 Then
                                tmLkRifRec(llLkUpper) = tmRifRec(llUpper)
                                tmLkRifRec(llLkUpper).lLkYear = 0
                                tmLkRifRec(llLp1).lLkYear = llLkUpper
                                ReDim Preserve tmLkRifRec(0 To llLkUpper + 1) As RIFREC
                                Exit Do
                            Else
                                llLp1 = tmLkRifRec(llLp1).lLkYear
                            End If
                        End If
                    Loop
                End If
            Else
                If tmRifRec(llUpper).tRif.iYear > tmRifRec(llLoop).tRif.iYear Then
                    tlRifRec = tmRifRec(llLoop)
                    tmRifRec(llLoop) = tmRifRec(llUpper)
                    tmRifRec(llLoop).sKey = tlRifRec.sKey
                    tmRifRec(llLoop).iStatus = 1
                    tmRifRec(llLoop).lLkYear = 0
                End If
            End If
            ilRifFound = True
        End If
        If Not ilRifFound Then
            slStr = ""
            ilVehFound = False
''            slRecCode = Trim$(Str$(tmRifRec(llUpper).tRif.iVefCode))
''            If Val(slRecCode) <> tmVef.iCode Then
''                tmVefSrchKey.iCode = Val(slRecCode) 'ilCode
''                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
''            Else
''                ilRet = BTRV_ERR_NONE
''            End If
''            If (ilRet = BTRV_ERR_NONE) And (tmVef.sState <> "D") Then
            ilVef = gBinarySearchVef(tmRifRec(llUpper).tRif.iVefCode)
            If ilVef <> -1 Then
                ilVehFound = True
'                If ilSortVehbyGroup Then
'                    slStr = Trim$(Str$(tgMVef(ilVef).iSort))    'tmVef.iSort))
'                    Do While Len(slStr) < 5
'                        slStr = "0" & slStr
'                    Loop
'                Else
'                    slStr = "00000"
'                End If
'                slStr = slStr & tgMVef(ilVef).sName
'                If ilSortVehbyGroup Then
'                    tmMnfSrchKey.iCode = tgMVef(ilVef).iMnfGroup(1)
'                    If tmMnf.iCode <> tmMnfSrchKey.iCode Then
'                        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                        If ilRet <> BTRV_ERR_NONE Then
'                            tmMnf.iGroupNo = 999
'                        End If
'                    End If
'                    slGpSort = Trim$(Str$(tmMnf.iGroupNo))
'                    Do While Len(slGpSort) < 3
'                        slGpSort = "0" & slGpSort
'                    Loop
'                Else
'                    slGpSort = "000"
'                End If
'                slStr = slGpSort & slStr
            End If
            If ilVehFound Then
'                'tmRdfSrchKey.iCode = tmRifRec(llUpper).tRif.iRdfCode
'                'ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'                For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
'                    If tmRifRec(llUpper).tRif.iRdfcode = tgMRdf(ilRdf).iCode Then
'                        ilRet = BTRV_ERR_NONE
'                        tmRdf = tgMRdf(ilRdf)
'                        If (tmRifRec(llUpper).tRif.sBase = "Y") Or (tmRifRec(llUpper).tRif.sBase = "N") Then
'                            slStr1 = Trim$(Str$(tmRifRec(llUpper).tRif.iSort))  'tmRdf.iSortCode))
'                        Else
'                            slStr1 = Trim$(Str$(tmRdf.iSortCode))
'                        End If
'                        Do While Len(slStr1) < 5
'                            slStr1 = "0" & slStr1
'                        Loop
'                        slStr = slStr & slStr1
'                        ilTime = 1
'                        For ilLoop = 1 To imMaxTDRows Step 1  'Row
'                            If (tmRdf.iStartTime(0, ilLoop) <> 1) Or (tmRdf.iStartTime(1, ilLoop) <> 0) Then
'                                ilTime = ilLoop
'                                For ilIndex = 1 To 7 Step 1
'                                    If tmRdf.sWkDays(ilTime, ilIndex) <> "Y" Then
'                                        slStr = slStr & "B"
'                                    Else
'                                        slStr = slStr & "A"
'                                    End If
'                                Next ilIndex
'                                Exit For
'                            End If
'                        Next ilLoop
'                        llSec = tmRdf.iStartTime(0, ilTime) \ 256 'Obtain seconds
'                        llMin = tmRdf.iStartTime(1, ilTime) And &HFF 'Obtain Minutes
'                        llHour = tmRdf.iStartTime(1, ilTime) \ 256 'Obtain month
'                        llTime = 3600 * llHour + 60 * llMin + llSec
'                        slTime = Trim$(Str$(llTime))
'                        Do While (Len(slTime) < 5)
'                            slTime = "0" & slTime
'                        Loop
'                        slStr = slStr & slTime
'                        tmRifRec(llUpper).sKey = slStr
                        tmRifRec(llUpper).iStatus = 1
                        tmRifRec(llUpper).lLkYear = 0
                        llLoop = llUpper
                        llUpper = llUpper + 1
                        ReDim Preserve tmRifRec(0 To llUpper) As RIFREC
'                        Exit For
'                    End If
'                Next ilRdf
            End If
        End If
    Next llTemp
    If llUpper > 1 Then
        'ArraySortTyp fnAV(tmRifRec(), 1), UBound(tmRifRec) - 1, 0, LenB(tmRifRec(1)), 0, LenB(tmRifRec(1).sKey), 0
        For llLoop = LBound(tmRifRec) To UBound(tmRifRec) - 1 Step 1
            tmRifRec(llLoop) = tmRifRec(llLoop + 1)
        Next llLoop
        ReDim Preserve tmRifRec(0 To UBound(tmRifRec) - 1) As RIFREC
        ArraySortTyp fnAV(tmRifRec(), 0), UBound(tmRifRec), 0, LenB(tmRifRec(0)), 0, LenB(tmRifRec(0).sKey), 0
        ReDim Preserve tmRifRec(0 To UBound(tmRifRec) + 1) As RIFREC
        For llLoop = UBound(tmRifRec) - 1 To LBound(tmRifRec) Step -1
            tmRifRec(llLoop + 1) = tmRifRec(llLoop)
        Next llLoop
    End If
    mInitRif UBound(tmRifRec)
    'Clear duplicates
    For llLoop = LBound(lmAutoDelRif) To UBound(lmAutoDelRif) - 1 Step 1
        Do
            tmRifSrchKey1.lCode = lmAutoDelRif(llLoop)
            ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            ilRet = btrDelete(hmRif)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet = BTRV_ERR_NONE Then
            For llTest = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                If tgMRif(llTest).lCode = tlRif.lCode Then
                    For llIndex = llTest To UBound(tgMRif) - 1 Step 1
                        tgMRif(llIndex) = tgMRif(llIndex + 1)
                    Next llIndex
                    'ReDim Preserve tgMRif(1 To UBound(tgMRif) - 1) As RIF
                    ReDim Preserve tgMRif(0 To UBound(tgMRif) - 1) As RIF
                    Exit For
                End If
            Next llTest
        End If
    Next llLoop
    Erase lmAutoDelRif
    'mInitRateCardCtrls
    Erase tmTempRifRec
    mReadRifRec = True
    Exit Function

    On Error GoTo 0
    mReadRifRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mRemakeDPName                   *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set the daypart show values    *
'*                                                     *
'*******************************************************
Private Sub mRemakeDpName()
    Dim llRowNo As Long
    Dim ilLoop As Integer
    For llRowNo = LBONE To UBound(tmRifRec) - 1 Step 1
        If tmRifRec(llRowNo).tRif.iRdfCode > 0 Then
            'tmRdfSrchKey.iCode = tmRifRec(llRowNo).tRif.iRdfCode
            'ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            'If ilRet = BTRV_ERR_NONE Then
            '    smRCSave(DAYPARTINDEX, llRowNo) = Trim$(tmRdf.sName)
            '    mSetDPShow llRowNo
            'End If
            For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                If tmRifRec(llRowNo).tRif.iRdfCode = tgMRdf(ilLoop).iCode Then
                    tmRdf = tgMRdf(ilLoop)
                    smRCSave(DAYPARTINDEX, llRowNo) = Trim$(tmRdf.sName)
                    mSetDPShow llRowNo
                    Exit For
                End If
            Next ilLoop
        End If
    Next llRowNo
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mRifGetRate                     *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get price for date specified   *
'*                                                     *
'*******************************************************
Private Sub mRifGetRate(llRif As Long, slInDate As String, tlRifRec() As RIFREC, tlLkRifRec() As RIFREC, llRateAmount As Long)
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim ilWkNo As Integer
    Dim ilFirstLastWk As Integer
    Dim llLkYear As Long
    llRateAmount = 0
    gObtainMonthYear 0, slInDate, ilMonth, ilYear
    gObtainWkNo 0, slInDate, ilWkNo, ilFirstLastWk
    If (ilWkNo > 0) And (ilWkNo < 54) Then
        If tlRifRec(llRif).tRif.iYear = ilYear Then
            llRateAmount = tlRifRec(llRif).tRif.lRate(ilWkNo)
        Else
            llLkYear = tlRifRec(llRif).lLkYear
            Do While llLkYear > 0
                If tlLkRifRec(llLkYear).tRif.iYear = ilYear Then
                    llRateAmount = tlLkRifRec(llLkYear).tRif.lRate(ilWkNo)
                    Exit Do
                Else
                    llLkYear = tlLkRifRec(llLkYear).lLkYear
                End If
            Loop
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mRifSetRate                     *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set price for date specified   *
'*                                                     *
'*******************************************************
Private Sub mRifSetRate(llRif As Long, slInDate As String, llRateAmount As Long, tlRifRec() As RIFREC, tlLkRifRec() As RIFREC)
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim ilWkNo As Integer
    Dim ilFirstLastWk As Integer
    Dim llLkYear As Long
    gObtainMonthYear 0, slInDate, ilMonth, ilYear
    gObtainWkNo 0, slInDate, ilWkNo, ilFirstLastWk
    If (ilWkNo > 0) And (ilWkNo < 54) Then
        If tlRifRec(llRif).tRif.iYear = ilYear Then
            tlRifRec(llRif).tRif.lRate(ilWkNo) = llRateAmount
        Else
            llLkYear = tlRifRec(llRif).lLkYear
            Do While llLkYear > 0
                If tlLkRifRec(llLkYear).tRif.iYear = ilYear Then
                    tlLkRifRec(llLkYear).tRif.lRate(ilWkNo) = llRateAmount
                    Exit Do
                Else
                    llLkYear = tlLkRifRec(llLkYear).lLkYear
                End If
            Loop
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:6/29/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim llRowNo As Long
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim slMsg As String
    Dim slStamp As String   'Date/Time stamp for file
    Dim llRif As Long
    Dim ilRcf As Integer
    Dim ilNewRcf As Integer
    Dim ilNewRif As Integer
    Dim llLkYear As Long
    Dim ilSvLkYear As Integer
    Dim llRcfRecPos As Long
    Dim ilFound As Integer
    Dim llTest As Long
    Dim llLoop As Long
    Dim ilRifChgd As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim tlRcf As RCF
    Dim tlRif As RIF
    Dim tlRif1 As MOVEREC
    Dim tlRif2 As MOVEREC
    mRCSetShow imRCBoxNo
    For llRowNo = LBONE To UBound(smRCSave, 2) - 1 Step 1
        lmRCRowNo = llRowNo
        If mTestSaveFields() = NO Then
            Beep
            mRCEnableBox imRCBoxNo
            Exit Function
        End If
    Next llRowNo
    mMoveCtrlToRec
    If mTestFields() = NO Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        mSaveRec = False
        Exit Function
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    ilRet = btrBeginTrans(hmRcf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
        mSaveRec = False
        Exit Function
    End If
    If ((imSelectedIndex = 0) And (imAdjIndex = 1)) Or (tgRcfI.iCode = 0) Then 'New selected
        If Not mSetEndDates() Then
            ilRet = btrAbortTrans(hmRcf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
            mSaveRec = False
            Exit Function
        End If
    End If
    If (((imSelectedIndex > 0) And (imAdjIndex = 1)) Or ((imSelectedIndex >= 0) And (imAdjIndex = 0))) And (tgRcfI.iCode <> 0) Then
        ilRet = btrGetPosition(hmRcf, llRcfRecPos)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmRcf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
            mSaveRec = False
            Exit Function
        End If
        'tmRec = tgRcfI
        'ilRet = gGetByKeyForUpdate("RCF", hmRcf, tmRec)
        ''tgRcfI = tmRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    ilRet = btrAbortTrans(hmRcf)
        '    Screen.MousePointer = vbDefault    'Default
        '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
        '    mSaveRec = False
        '    Exit Function
        'End If
    End If
    Do  'Loop until record updated or added
        slStamp = gFileDateTime(sgDBPath & "Rcf.btr")
        'If Len(lbcRateCard.Tag) > Len(slStamp) Then
        '    slStamp = slStamp & Right$(lbcRateCard.Tag, Len(lbcRateCard.Tag) - Len(slStamp))
        'End If
        If Len(smRateCardTag) > Len(slStamp) Then
            slStamp = slStamp & right$(smRateCardTag, Len(smRateCardTag) - Len(slStamp))
        End If
        tgRcfI.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
        If ((imSelectedIndex = 0) And (imAdjIndex = 1)) Or (tgRcfI.iCode = 0) Then 'New selected
            ilNewRcf = True
            tgRcfI.iCode = 0  'Autoincrement
            tgRcfI.iRemoteID = tgUrf(0).iRemoteUserID
            tgRcfI.iAutoCode = tgRcfI.iCode
            ilRet = btrInsert(hmRcf, tgRcfI, imRcfRecLen, INDEXKEY0)
            'tgRcfI.iCode = 0  'Autoincrement
            'tgRcfI.iYear = tgRcfI.iYear + 1
            'ilRet = btrInsert(hmRcf, tgRcfI, imRcfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Rate Card)"
        Else 'Old record-Update
            ilNewRcf = False
            gPackDate slSyncDate, tgRcfI.iSyncDate(0), tgRcfI.iSyncDate(1)
            gPackTime slSyncTime, tgRcfI.iSyncTime(0), tgRcfI.iSyncTime(1)
            ilRet = btrUpdate(hmRcf, tgRcfI, imRcfRecLen)
            If ilRet = BTRV_ERR_CONFLICT Then
                ilCRet = btrGetDirect(hmRcf, tlRcf, imRcfRecLen, llRcfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilCRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hmRcf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                    mSaveRec = False
                    Exit Function
                End If
                'tmRec = tlRcf  'tgRcfI
                'ilCRet = gGetByKeyForUpdate("RCF", hmRcf, tmRec)
                'tlRcf = tmRec  'tgRcfI = tmRec
                'If ilCRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmRcf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
                '    mSaveRec = False
                '    Exit Function
                'End If
            End If
            slMsg = "mSaveRec (btrUpdate: Rate Card)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrAbortTrans(hmRcf)
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
        mSaveRec = False
        Exit Function
    End If
    If ilNewRcf Then
        Do
            'tmRcfSrchKey.iCode = tgRcfI.iCode
            'ilRet = btrGetEqual(hmRcf, tgRcfI, imRcfRecLen, tmRcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            'If ilRet <> BTRV_ERR_NONE Then
            '    ilRet = btrAbortTrans(hmRcf)
            '    Screen.MousePointer = vbDefault    'Default
            '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
            '    mSaveRec = False
            '    Exit Function
            'End If
            tgRcfI.iRemoteID = tgUrf(0).iRemoteUserID
            tgRcfI.iAutoCode = tgRcfI.iCode
            gPackDate slSyncDate, tgRcfI.iSyncDate(0), tgRcfI.iSyncDate(1)
            gPackTime slSyncTime, tgRcfI.iSyncTime(0), tgRcfI.iSyncTime(1)
            ilRet = btrUpdate(hmRcf, tgRcfI, imRcfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmRcf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
            mSaveRec = False
            Exit Function
        End If
        tgMRcf(UBound(tgMRcf)) = tgRcfI
        'ReDim Preserve tgMRcf(1 To UBound(tgMRcf)) As RCF
        ReDim Preserve tgMRcf(0 To UBound(tgMRcf)) As RCF
    Else
        For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
            If tgMRcf(ilRcf).iCode = tgRcfI.iCode Then
                tgMRcf(ilRcf) = tgRcfI
                Exit For
            End If
        Next ilRcf
    End If
    sgMRcfStamp = gFileDateTime(sgDBPath & "Rcf.Btr")
    For llRif = LBONE To UBound(tmRifRec) - 1 Step 1
        Do  'Loop until record updated or added
            If ((imSelectedIndex = 0) And (imAdjIndex = 1)) Or (tmRifRec(llRif).iStatus = 0) Or ilNewRcf Then 'New selected
                ilNewRif = True
                tmRifRec(llRif).tRif.lCode = 0
                tmRifRec(llRif).tRif.iRcfCode = tgRcfI.iCode
                tmRifRec(llRif).tRif.iRemoteID = tgUrf(0).iRemoteUserID
                tmRifRec(llRif).tRif.lAutoCode = tmRifRec(llRif).tRif.lCode
                'tmRifRec(ilRif).tRif.iYear = tgRcfI.iYear
                ilRet = btrInsert(hmRif, tmRifRec(llRif).tRif, imRifRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert: Rate Card Items)"
            Else 'Old record-Update
                ilNewRif = False
                slMsg = "mSaveRec (btrGetDirect: Rate Card Items)"
                ilRifChgd = False
'                ilRet = btrGetEqual(hmRpf, tlRpf, imRpfRecLen, tmRpf0SrchKey, 0, BTRV_LOCK_NONE)  'position record so update works
                'tmRec = tlRif
                'ilRet = gGetByKeyForUpdate("RIF", hmRif, tmRec)
                'tlRif = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmRcf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
                '    mSaveRec = False
                '    Exit Function
                'End If
                ilFound = False
                For llTest = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                    If tgMRif(llTest).lCode = tmRifRec(llRif).tRif.lCode Then
                        tlRif = tgMRif(llTest)
                        ilFound = True
                        Exit For
                    End If
                Next llTest
                If Not ilFound Then
                    'ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmRifRec(ilRif).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    tmRifSrchKey1.lCode = tmRifRec(llRif).tRif.lCode
                    ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilRet = btrAbortTrans(hmRcf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                        mSaveRec = False
                        Exit Function
                    End If
                End If
                LSet tlRif1 = tlRif
                LSet tlRif2 = tmRifRec(llRif).tRif
                If StrComp(tlRif1.sChar, tlRif2.sChar, 0) <> 0 Then
                    If ilFound Then
                        'ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmRifRec(ilRif).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        tmRifSrchKey1.lCode = tmRifRec(llRif).tRif.lCode
                        ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmRcf)
                            Screen.MousePointer = vbDefault    'Default
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                            mSaveRec = False
                            Exit Function
                        End If
                    End If
                    gPackDate slSyncDate, tmRifRec(llRif).tRif.iSyncDate(0), tmRifRec(llRif).tRif.iSyncDate(1)
                    gPackTime slSyncTime, tmRifRec(llRif).tRif.iSyncTime(0), tmRifRec(llRif).tRif.iSyncTime(1)
                    ilRet = btrUpdate(hmRif, tmRifRec(llRif).tRif, imRifRecLen)
                    ilRifChgd = True
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Rate Card Items)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrAbortTrans(hmRcf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
            mSaveRec = False
            Exit Function
        End If
        If ilNewRif Then
            Do
                'tmRifSrchKey1.lCode = tmRifRec(ilRif).tRif.lCode
                'ilRet = btrGetEqual(hmRif, tmRifRec(ilRif).tRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmRcf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
                '    mSaveRec = False
                '    Exit Function
                'End If
                tmRifRec(llRif).tRif.iRemoteID = tgUrf(0).iRemoteUserID
                tmRifRec(llRif).tRif.lAutoCode = tmRifRec(llRif).tRif.lCode
                gPackDate slSyncDate, tmRifRec(llRif).tRif.iSyncDate(0), tmRifRec(llRif).tRif.iSyncDate(1)
                gPackTime slSyncTime, tmRifRec(llRif).tRif.iSyncTime(0), tmRifRec(llRif).tRif.iSyncTime(1)
                ilRet = btrUpdate(hmRif, tmRifRec(llRif).tRif, imRifRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmRcf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                mSaveRec = False
                Exit Function
            End If
            tgMRif(UBound(tgMRif)) = tmRifRec(llRif).tRif
            'ReDim Preserve tgMRif(1 To UBound(tgMRif) + 1) As RIF
            ReDim Preserve tgMRif(0 To UBound(tgMRif) + 1) As RIF
        Else
            If ilRifChgd Then
                For llTest = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                    If tgMRif(llTest).lCode = tmRifRec(llRif).tRif.lCode Then
                        tgMRif(llTest) = tmRifRec(llRif).tRif
                        Exit For
                    End If
                Next llTest
            End If
        End If
    Next llRif
    sgMRifStamp = gFileDateTime(sgDBPath & "Rif.Btr")
    For llRif = LBONE To UBound(tmRifRec) - 1 Step 1
        llLkYear = tmRifRec(llRif).lLkYear
        Do While llLkYear > 0
            Do  'Loop until record updated or added
                If ((imSelectedIndex = 0) And (imAdjIndex = 1)) Or (tmLkRifRec(llLkYear).iStatus = 0) Or ilNewRcf Then 'New selected
                    ilNewRif = True
                    tmLkRifRec(llLkYear).tRif.lCode = 0
                    tmLkRifRec(llLkYear).tRif.iRcfCode = tgRcfI.iCode
                    tmLkRifRec(llLkYear).tRif.iRemoteID = tgUrf(0).iRemoteUserID
                    tmLkRifRec(llLkYear).tRif.lAutoCode = tmLkRifRec(llLkYear).tRif.lCode
                    ilRet = btrInsert(hmRif, tmLkRifRec(llLkYear).tRif, imRifRecLen, INDEXKEY0)
                    slMsg = "mSaveRec (btrInsert: Rate Card Items)"
                Else 'Old record-Update
                    ilNewRif = False
                    ilRifChgd = False
                    slMsg = "mSaveRec (btrGetDirect: Rate Card Items)"
    '                ilRet = btrGetEqual(hmRpf, tlRpf, imRpfRecLen, tmRpf0SrchKey, 0, BTRV_LOCK_NONE)  'position record so update works
                    ilFound = False
                    For llTest = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                        If tgMRif(llTest).lCode = tmLkRifRec(llLkYear).tRif.lCode Then
                            tlRif = tgMRif(llTest)
                            ilFound = True
                            Exit For
                        End If
                    Next llTest
                    If Not ilFound Then
                        'ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmLkRifRec(llLkYear).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        tmRifSrchKey1.lCode = tmLkRifRec(llLkYear).tRif.lCode
                        ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmRcf)
                            Screen.MousePointer = vbDefault    'Default
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                            mSaveRec = False
                            Exit Function
                        End If
                    End If
                    'tmRec = tlRif
                    'ilRet = gGetByKeyForUpdate("RIF", hmRif, tmRec)
                    'tlRif = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hmRcf)
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                    LSet tlRif1 = tlRif
                    LSet tlRif2 = tmLkRifRec(llLkYear).tRif
                    If StrComp(tlRif1.sChar, tlRif2.sChar, 0) <> 0 Then
                        If ilFound Then
                            'ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmLkRifRec(llLkYear).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            tmRifSrchKey1.lCode = tmLkRifRec(llLkYear).tRif.lCode
                            ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet <> BTRV_ERR_NONE Then
                                ilRet = btrAbortTrans(hmRcf)
                                Screen.MousePointer = vbDefault    'Default
                                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                                mSaveRec = False
                                Exit Function
                            End If
                        End If
                        gPackDate slSyncDate, tmLkRifRec(llLkYear).tRif.iSyncDate(0), tmLkRifRec(llLkYear).tRif.iSyncDate(1)
                        gPackTime slSyncTime, tmLkRifRec(llLkYear).tRif.iSyncTime(0), tmLkRifRec(llLkYear).tRif.iSyncTime(1)
                        ilRet = btrUpdate(hmRif, tmLkRifRec(llLkYear).tRif, imRifRecLen)
                        ilRifChgd = True
                    Else
                        ilRet = BTRV_ERR_NONE
                    End If
                    slMsg = "mSaveRec (btrUpdate: Rate Card Items)"
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmRcf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                mSaveRec = False
                Exit Function
            End If
            If ilNewRif Then
                Do
                    'tmRifSrchKey1.lCode = tmLkRifRec(llLkYear).tRif.lCode
                    'ilRet = btrGetEqual(hmRif, tmLkRifRec(llLkYear).tRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    ilRet = btrAbortTrans(hmRcf)
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                    tmLkRifRec(llLkYear).tRif.iRemoteID = tgUrf(0).iRemoteUserID
                    tmLkRifRec(llLkYear).tRif.lAutoCode = tmLkRifRec(llLkYear).tRif.lCode
                    gPackDate slSyncDate, tmLkRifRec(llLkYear).tRif.iSyncDate(0), tmLkRifRec(llLkYear).tRif.iSyncDate(1)
                    gPackTime slSyncTime, tmLkRifRec(llLkYear).tRif.iSyncTime(0), tmLkRifRec(llLkYear).tRif.iSyncTime(1)
                    ilRet = btrUpdate(hmRif, tmLkRifRec(llLkYear).tRif, imRifRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hmRcf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                    mSaveRec = False
                    Exit Function
                End If
                tgMRif(UBound(tgMRif)) = tmLkRifRec(llLkYear).tRif
                'ReDim Preserve tgMRif(1 To UBound(tgMRif) + 1) As RIF
                ReDim Preserve tgMRif(0 To UBound(tgMRif) + 1) As RIF
            Else
                If ilRifChgd Then
                    For llTest = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                        If tgMRif(llTest).lCode = tmLkRifRec(llLkYear).tRif.lCode Then
                            tgMRif(llTest) = tmLkRifRec(llLkYear).tRif
                            Exit For
                        End If
                    Next llTest
                End If
            End If
            llLkYear = tmLkRifRec(llLkYear).lLkYear
        Loop
    Next llRif
    sgMRifStamp = gFileDateTime(sgDBPath & "Rif.Btr")
    'Remove trash
    For llRif = LBONE To UBound(tmTrashRifRec) - 1 Step 1
        If tmTrashRifRec(llRif).iStatus = 1 Then
            Do  'Loop until record updated or added
                slMsg = "mDeleteRec (btrGetDirect: Rate Card Items)"
                'ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmTrashRifRec(ilRif).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                tmRifSrchKey1.lCode = tmTrashRifRec(llRif).tRif.lCode
                ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                llLkYear = tmTrashRifRec(llRif).lLkYear
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hmRcf)
                    Screen.MousePointer = vbDefault    'Default
                    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                    mSaveRec = False
                    Exit Function
                End If
                'tmRec = tlRif
                'ilRet = gGetByKeyForUpdate("RIF", hmRif, tmRec)
                'tlRif = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmRcf)
                '    Screen.MousePointer = vbDefault    'Default
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
                '    mSaveRec = False
                '    Exit Function
                'End If
                ilRet = btrDelete(hmRif)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmRcf)
                Screen.MousePointer = vbDefault    'Default
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                mSaveRec = False
                Exit Function
            End If
            For llTest = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                If tgMRif(llTest).lCode = tlRif.lCode Then
                    For llLoop = llTest To UBound(tgMRif) - 1 Step 1
                        tgMRif(llLoop) = tgMRif(llLoop + 1)
                    Next llLoop
                    'ReDim Preserve tgMRif(1 To UBound(tgMRif) - 1) As RIF
                    ReDim Preserve tgMRif(0 To UBound(tgMRif) - 1) As RIF
                    Exit For
                End If
            Next llTest

'            If tgSpf.sRemoteUsers = "Y" Then
'                tmDsf.lCode = 0
'                tmDsf.sFileName = "RIF"
'                gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'                gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'                tmDsf.iRemoteID = tlRif.iRemoteID
'                tmDsf.lAutoCode = tlRif.lAutoCode
'                tmDsf.lCntrNo = 0
'                ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'            End If
            Do While llLkYear > 0
                If tmLkRifRec(llLkYear).iStatus = 1 Then
                    ilSvLkYear = tmLkRifRec(llLkYear).lLkYear
                    Do  'Loop until record updated or added
                        slMsg = "mDeleteRec (btrGetDirect: Rate Card Items)"
                        'ilRet = btrGetDirect(hmRif, tlRif, imRifRecLen, tmLkRifRec(llLkYear).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        tmRifSrchKey1.lCode = tmLkRifRec(llLkYear).tRif.lCode
                        ilRet = btrGetEqual(hmRif, tlRif, imRifRecLen, tmRifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilRet = btrAbortTrans(hmRcf)
                            Screen.MousePointer = vbDefault    'Default
                            ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                            mSaveRec = False
                            Exit Function
                        End If
                        'tmRec = tlRif
                        'ilRet = gGetByKeyForUpdate("RIF", hmRif, tmRec)
                        'tlRif = tmRec
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    ilRet = btrAbortTrans(hmRcf)
                        '    Screen.MousePointer = vbDefault    'Default
                        '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Rate Card")
                        '    mSaveRec = False
                        '    Exit Function
                        'End If
                        ilRet = btrDelete(hmRif)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    If ilRet <> BTRV_ERR_NONE Then
                        ilRet = btrAbortTrans(hmRcf)
                        Screen.MousePointer = vbDefault    'Default
                        ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Rate Card")
                        mSaveRec = False
                        Exit Function
                    End If
                    For llTest = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                        If tgMRif(llTest).lCode = tlRif.lCode Then
                            For llLoop = llTest To UBound(tgMRif) - 1 Step 1
                                tgMRif(llLoop) = tgMRif(llLoop + 1)
                            Next llLoop
                            'ReDim Preserve tgMRif(1 To UBound(tgMRif) - 1) As RIF
                            ReDim Preserve tgMRif(0 To UBound(tgMRif) - 1) As RIF
                            Exit For
                        End If
                    Next llTest
'                    If tgSpf.sRemoteUsers = "Y" Then
'                        tmDsf.lCode = 0
'                        tmDsf.sFileName = "RIF"
'                        gPackDate slSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'                        gPackTime slSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'                        tmDsf.iRemoteID = tlRif.iRemoteID
'                        tmDsf.lAutoCode = tlRif.lAutoCode
'                        tmDsf.lCntrNo = 0
'                        ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'                    End If
                    llLkYear = ilSvLkYear
                Else
                    llLkYear = tmLkRifRec(llLkYear).lLkYear
                End If
            Loop
        End If
    Next llRif
    ilRet = btrEndTrans(hmRcf)
    ReDim tmTrashRifRec(0 To 1) As RIFREC
    sgMRcfStamp = gFileDateTime(sgDBPath & "Rcf.Btr")
    sgMRifStamp = gFileDateTime(sgDBPath & "Rif.Btr")
'    'If lbcRateCard.Tag <> "" Then
'    '    If slStamp = lbcRateCard.Tag Then
'    '        lbcRateCard.Tag = FileDateTime(sgDBPath & "Rcf.btr")
'    '        If Len(slStamp) > Len(lbcRateCard.Tag) Then
'    '            lbcRateCard.Tag = lbcRateCard.Tag & Right$(slStamp, Len(slStamp) - Len(lbcRateCard.Tag))
'    '        End If
'    '    End If
'    'End If
'    'If ilNewRcf <> False Then
'    '    slNameFac = cbcSelect.List(1)
'    '    gUnpackDateForSort tgRcfI.iStartDate(0), tgRcfI.iStartDate(1), slName
'    '    slName = slName & "\" & slNameFac
'    '    slName = slName + "\" + LTrim$(Str$(tgRcfI.iCode))
'    '    lbcRateCard.AddItem slName
'    '    ilNewIndex = lbcRateCard.NewIndex
'    '    cbcSelect.RemoveItem 1
'    '    cbcSelect.AddItem slNameFac, lbcRateCard.ListCount - ilNewIndex
'    '    cbcSelect.ListIndex = lbcRateCard.ListCount - ilNewIndex
'    'End If

    tmcDelay.Enabled = True
'    If imSelectedIndex > 0 Then
'        slNameFac = cbcSelect.List(imSelectedIndex)
'    Else
'        slNameFac = ""
'    End If
'    mPopulate
'    gFindMatch slNameFac, 0, cbcSelect
'    If gLastFound(cbcSelect) > 0 Then
'        cbcSelect.ListIndex = gLastFound(cbcSelect)
'    End If
    mSaveRec = True
'    Screen.MousePointer = vbDefault    'Default
    Exit Function

    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:6/29/93       By:D. LeVine      *
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
'    If (igRcfChg Or imRpfChg Or imRgfChg) And ((imRowNo < UBound(tgRpfI) + 1) Or ((imRowNo = UBound(tgRpfI) + 1) And (imBoxNo = 0))) Then
    If igRcfChg Or imRifChg Then
        If ilAsk Then
            If ((imSelectedIndex > 0) And (imAdjIndex = 1)) Or ((imSelectedIndex >= 0) And (imAdjIndex = 0)) Then
                slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
            Else
                slMess = "Add " & tgRcfI.sName
            End If
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                pbcRateCard_Paint
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
    mSaveRecChg = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    'Update button set if all mandatory fields have data and any field altered
    If ((igRcfChg) Or (imRifChg)) And (imUpdateAllowed) Then
        cbcSelect.Enabled = False
        If UBound(tmRifRec) > 1 Then
            cmcUpdate.Enabled = True
        Else
            cmcUpdate.Enabled = False
        End If
    Else
        cbcSelect.Enabled = True
        cmcUpdate.Enabled = False
    End If
    If (imSelectedIndex <= 0) And (imAdjIndex = 1) Then
        cmcTerms.Enabled = False
        'plcSP.Enabled = False
        pbcRateCard.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
    Else
        cmcTerms.Enabled = True
        'plcInfo.Enabled = True
        If (igWinStatus(RATECARDSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            pbcRateCard.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
        End If
    End If
    'Revert button set if any field changed
    If igRcfChg Or imRifChg Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'Erase button set if any field contains a value or change mode
    If ((imSelectedIndex > 0) And (imAdjIndex = 1)) Or ((imSelectedIndex >= 0) And (imAdjIndex = 0)) Or (UBound(tmRifRec) > LBONE) Then
        If (imView = 2) Or (imAdjIndex = 0) Then
            cmcErase.Enabled = False
        Else
            If (igWinStatus(RATECARDSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
                If imUpdateAllowed Then
                    cmcErase.Enabled = True
                Else
                    cmcErase.Enabled = False
                End If
            Else
                cmcErase.Enabled = False
            End If
        End If
        cmcStdPkg.Enabled = True
    Else
        cmcErase.Enabled = False
        cmcStdPkg.Enabled = False
    End If
    'added by L.Bianchi
    If ((Asc(tgSaf(0).sFeatures8) And PODAIRTIME) <> PODAIRTIME) And ((Asc(tgSaf(0).sFeatures8) And PODSPOTS) <> PODSPOTS) Then
        cmcStdPkg.Enabled = False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetDefInSave                   *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set default values in imSave   *
'*                                                     *
'*******************************************************
Private Sub mSetDefInSave()
    Dim llLast As Long
    'Dim ilRpfIndex As Integer
    Dim ilLoop As Integer
    llLast = UBound(tmRifRec)
    tmRifRec(llLast).tRif.iRcfCode = tgRcfI.iCode
    tmRifRec(llLast).tRif.iVefCode = 0
    tmRifRec(llLast).tRif.iRdfCode = 0
    'Use previous if defined, if not defined use term Vehicle
    smRCSave(VEHINDEX, llLast) = ""
'    If ilLast > 1 Then
'        smSave(1, ilLast) = smSave(1, ilLast - 1)
'    Else
'        smSave(1, ilLast) = smDefVehicle
'    End If
'    If smSave(1, ilLast) = "" Then
'        smSave(1, ilLast) = sgUserDefVehicleName
'    End If
    For ilLoop = LBONE + 1 To UBound(smRCSave, 1) Step 1
        smRCSave(ilLoop, llLast) = ""
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetDPShow                      *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set the daypart show values    *
'*                                                     *
'*******************************************************
Private Sub mSetDPShow(llRowNo As Long)
'
'   Where:
'       llRowNo(I)- Row No
'
'       tmRdf(I)- Rate Card Daypart record
'
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLib As Integer
    Dim slMultiTimes As String
    Dim flWidth As Single
    Dim flSvWidth As Single
    Dim ilCount As Integer
    Dim ilTime As Integer
    Dim ilDay As Integer
    Dim ilIndex As Integer
    Dim slStart As String
    Dim slEnd As String
    slStr = Trim$(smRCSave(DAYPARTINDEX, llRowNo))
    gSetShow pbcRateCard, slStr, tmRCCtrls(DAYPARTINDEX)
    smRCShow(DAYPARTINDEX, llRowNo) = tmRCCtrls(DAYPARTINDEX).sShow
    slStr = Trim$(smRCSave(DAYPARTINDEX, llRowNo))
    gSetShow pbcDaypart, slStr, tmDPCtrls(DAYPARTINDEX)
    smDPShow(DAYPARTINDEX, llRowNo) = tmDPCtrls(DAYPARTINDEX).sShow
    'Determine if by library or time by testing libCode
    If (tmRdf.iLtfCode(0) <> 0) Or (tmRdf.iLtfCode(0) <> 0) Or (tmRdf.iLtfCode(0) <> 0) Then
        flSvWidth = tmDPCtrls(TIMESINDEX).fBoxW
        ilCount = 0
        For ilLib = 0 To 2 Step 1
            If tmRdf.iLtfCode(ilLib) Then
                ilCount = ilCount + 1
            End If
        Next ilLib
        If ilCount > 0 Then
            flWidth = tmDPCtrls(TIMESINDEX).fBoxW / ilCount
        Else
            flWidth = tmDPCtrls(TIMESINDEX).fBoxW
        End If
        slStr = ""
        For ilLib = 0 To 2 Step 1
            If tmRdf.iLtfCode(ilLib) > 0 Then
                tmLtfSrchKey.iCode = tmRdf.iLtfCode(ilLib)
                ilRet = btrGetEqual(hmLtf, tmLtf, imLtfRecLen, tmLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If ilCount > 0 Then
                        slStr = slStr & "/"
                    End If
                    ilCount = ilCount + 1
                    slStr = slStr & Trim$(tmLtf.sName)
                    tmDPCtrls(TIMESINDEX).fBoxW = ilCount * flWidth
                    gSetShow pbcDaypart, slStr, tmDPCtrls(TIMESINDEX)
                End If
            End If
        Next ilLib
        tmDPCtrls(TIMESINDEX).fBoxW = flSvWidth
        smDPShow(TIMESINDEX, llRowNo) = tmDPCtrls(TIMESINDEX).sShow
    Else
        slStr = ""
        slMultiTimes = ""
        For ilTime = 1 To imMaxTDRows + 1 Step 1 'Row
            If (tmRdf.iStartTime(0, ilTime - 1) <> 1) Or (tmRdf.iStartTime(1, ilTime - 1) <> 0) Then
                gUnpackTime tmRdf.iStartTime(0, ilTime - 1), tmRdf.iStartTime(1, ilTime - 1), "A", "1", slStart
                gUnpackTime tmRdf.iEndTime(0, ilTime - 1), tmRdf.iEndTime(1, ilTime - 1), "A", "1", slEnd
                If slStart <> "" Then
                    slStr = slStart & "-" & slEnd
                    If ilTime < imMaxTDRows + 1 Then
                        'If (tmRdf.iStartTime(0, ilTime + 1) <> 1) Or (tmRdf.iStartTime(1, ilTime + 1) <> 0) Then
                        If (tmRdf.iStartTime(0, ilTime) <> 1) Or (tmRdf.iStartTime(1, ilTime) <> 0) Then
                            slMultiTimes = "+"
                        End If
                    End If
                End If
                ilIndex = ilTime    'llRowNo
                Exit For
            End If
        Next ilTime
        gSetShow pbcDaypart, slStr, tmDPCtrls(TIMESINDEX)
        smDPShow(TIMESINDEX, llRowNo) = tmDPCtrls(TIMESINDEX).sShow & slMultiTimes
        For ilDay = 1 To 7 Step 1
            If tmRdf.sWkDays(ilIndex - 1, ilDay - 1) = "Y" Then
                slStr = "Y"
            Else
                slStr = "N"
            End If
            gSetShow pbcDaypart, slStr, tmDPCtrls(DAYINDEX + ilDay - 1)
            smDPShow(DAYINDEX + ilDay - 1, llRowNo) = tmDPCtrls(DAYINDEX + ilDay - 1).sShow
        Next ilDay
        slStr = "All avails"
        If (tmRdf.sInOut = "I") Or (tmRdf.sInOut = "O") Then
            If tmRdf.ianfCode > 0 Then
                tmAnfSrchKey.iCode = tmRdf.ianfCode
                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If (tmRdf.sInOut = "I") Then
                        slStr = Trim$(tmAnf.sName)
                    ElseIf (tmRdf.sInOut = "O") Then
                        slStr = "�" & Trim$(tmAnf.sName)     'Alt + 0216 or 0248
                    End If
                End If
            End If
        End If
        gSetShow pbcDaypart, slStr, tmDPCtrls(AVAILINDEX)
        smDPShow(AVAILINDEX, llRowNo) = tmDPCtrls(AVAILINDEX).sShow
        slStr = ""
        If tmRdf.sTimeOver = "Y" Then
            slStr = "Y"
        Else
            slStr = "N"
        End If
        gSetShow pbcDaypart, slStr, tmDPCtrls(HRSINDEX)
        smDPShow(HRSINDEX, llRowNo) = tmDPCtrls(HRSINDEX).sShow
        smDPShow(DPBASEINDEX, llRowNo) = tmRdf.sBase
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetEndDates                    *
'*                                                     *
'*             Created:7/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set end date for all rate cards*
'*                                                     *
'*******************************************************
Private Function mSetEndDates()
    Dim tlRcf As RCF
    'Dim tlRpf As RPF
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    Dim ilCRet As Integer
    Dim ilLoop As Integer
    Dim slMsg As String
    Dim slStr As String
    ReDim iEndDate(0 To 1) As Integer
    'Remove the setting of the end dates- might want this back in once we add
    'a Rate Card Status (Adive-Default; Active; Dormant; Hidden)
    mSetEndDates = True
    Exit Function
    gUnpackDate tgRcfI.iStartDate(0), tgRcfI.iStartDate(1), slStr
    slMsg = gDecOneDay(slStr)
    gPackDate slMsg, iEndDate(0), iEndDate(1)
    For ilLoop = 0 To UBound(tmRateCard) - 1 Step 1 'lbcRateCard.ListCount - 1 Step 1
        slNameCode = tmRateCard(ilLoop).sKey   'lbcRateCard.List(ilLoop)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        On Error GoTo mSetEndDatesErr
        gCPErrorMsg ilRet, "mSetEndDates (gParseItem field 2)", RateCard
        On Error GoTo 0
        slCode = Trim$(slCode)
        tmRcfSrchKey.iCode = CInt(slCode)
        ilRet = btrGetEqual(hmRcf, tlRcf, imRcfRecLen, tmRcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            mSetEndDates = False
            Exit Function
        End If
        If tlRcf.iVefCode = tgRcfI.iVefCode Then
            If (tlRcf.iEndDate(0) = 0) And (tlRcf.iEndDate(1) = 0) Then
                tlRcf.iEndDate(0) = iEndDate(0)
                tlRcf.iEndDate(1) = iEndDate(1)
                Do  'Loop until record updated or added
                    ilRet = btrUpdate(hmRcf, tlRcf, imRcfRecLen)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        ilCRet = btrGetEqual(hmRcf, tlRcf, imRcfRecLen, tmRcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilCRet <> BTRV_ERR_NONE Then
                            mSetEndDates = False
                            Exit Function
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    mSetEndDates = False
                    Exit Function
                End If
    '            tmRpfSrchKey.iRcfCode = tlRcf.iCode
    '            tmRpfSrchKey.iVefCode = -32000
    '            ilRet = btrGetGreaterOrEqual(hmRpf, tlRpf, imRpfRecLen, tmRpfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    '            slMsg = "mSetEndDates (btrGetGreaterOrEqual: Rate Card Program)"
    '            Do While ((tlRpf.iRcfCode = tlRcf.iCode) And (ilRet <> BTRV_ERR_END_OF_FILE))
    '                If (tlRpf.iEndDate(0) = 0) And (tlRpf.iEndDate(1) = 0) Then
    '                    tlRpf.iStartDate(0) = tlRcf.iStartDate(0)
    '                    tlRpf.iStartDate(1) = tlRcf.iStartDate(1)
    '                    tlRpf.iEndDate(0) = tlRcf.iEndDate(0)
    '                    tlRpf.iEndDate(1) = tlRcf.iEndDate(1)
    '                    Do  'Loop until record updated or added
    '                        ilRet = btrUpdate(hmRpf, tlRpf, imRpfRecLen)
    '                        slMsg = "mSetEndDates (btrUpdate: Rate Card Program)"
    '                    Loop While ilRet = BTRV_ERR_CONFLICT
    '                    On Error GoTo mSetEndDatesErr
    '                    gBtrvErrorMsg ilRet, slMsg, RateCard
    '                    On Error GoTo 0
    '                End If
    '                ilRet = btrGetNext(hmRpf, tlRpf, imRpfRecLen, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    '                slMsg = "mSetEndDates (btrGetNext: Rate Card Program)"
    '            Loop
    '            If ilRet <> BTRV_ERR_END_OF_FILE Then
    '                On Error GoTo mSetEndDatesErr
    '                gBtrvErrorMsg ilRet, slMsg, RateCard
    '                On Error GoTo 0
    '            End If
            End If
        End If
    Next ilLoop
    mSetEndDates = True
    Exit Function
mSetEndDatesErr:
    On Error GoTo 0
    mSetEndDates = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    If (ilBoxNo < imLBRCCtrls) Or (ilBoxNo > UBound(tmRCCtrls)) Then
        Exit Sub
    End If

    If (lmRCRowNo < vbcRateCard.Value) Or (lmRCRowNo >= vbcRateCard.Value + vbcRateCard.LargeChange + 1) Then
        mRCSetShow ilBoxNo
        pbcArrow.Visible = False
        lacRCFrame.Visible = False
        lacDPFrame.Visible = False
        Exit Sub
    End If
    lacRCFrame.Move 0, tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
    lacRCFrame.Visible = True
    lacDPFrame.Move 0, tmDPCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
    lacDPFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcRateCard.Top + tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHINDEX 'Vehicle
            edcDropDown.SetFocus
        Case DAYPARTINDEX 'Program name index
            edcDropDown.SetFocus
        'Case DOLLARINDEX 'Vehicle
        '    edcDropDown.SetFocus
        'Case PCTINVINDEX 'Program name index
        '    edcDropDown.SetFocus
        Case BASEINDEX
            pbcYN.SetFocus
        Case RPTINDEX
            pbcYN.SetFocus
        Case SORTINDEX 'Program name index
            edcDropDown.SetFocus
        Case DOLLAR1INDEX 'Vehicle
            edcDropDown.SetFocus
        Case DOLLAR2INDEX 'Vehicle
            edcDropDown.SetFocus
        Case DOLLAR3INDEX 'Vehicle
            edcDropDown.SetFocus
        Case DOLLAR4INDEX 'Vehicle
            edcDropDown.SetFocus
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetPrice                       *
'*                                                     *
'*             Created:7/09/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move values from working area  *
'*                      to the record                  *
'*                                                     *
'*******************************************************
Private Sub mSetPrice(ilGroup As Integer, llRowNo As Long, slDollar As String)
    Dim ilWk As Integer
    Dim slStart As String
    Dim llLkYear As Long
    Dim llDollar As Long
    llDollar = Val(slDollar)    'The user specifies the average price per week
    If tmRifRec(llRowNo).tRif.iYear = tmPdGroups(ilGroup).iYear Then
        For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
            slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
            If rbcShow(0).Value Then
                slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                mRifSetRate CLng(llRowNo), slStart, llDollar, tmRifRec(), tmLkRifRec()
            Else
                tmRifRec(llRowNo).tRif.lRate(ilWk) = llDollar
            End If
            'If ilWk = 1 Then
            '    If rbcShow(0).Value Then    'Don't split dollars if input via corporate
            '        'slStr = ".00"
            '        'gStrToPDN slStr, 2, 5, tmRifRec(llRowNo).tRif.sRate(0)
            '        'gStrToPDN slDollar, 2, 5, tmRifRec(llRowNo).tRif.sRate(ilWk)
            '        tmRifRec(llRowNo).tRif.lRate(0) = 0
            '        tmRifRec(llRowNo).tRif.lRate(ilWk) = llDollar
            '    Else
            '        'slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
            '        'slStart = gObtainStartCorp(slDate, True)
            '        'ilDay = gWeekDayStr(slStart)
            '        'If ilDay = 0 Then
            '            'slStr = ".00"
            '            'gStrToPDN slStr, 2, 5, tmRifRec(llRowNo).tRif.sRate(0)
            '            'gStrToPDN slDollar, 2, 5, tmRifRec(llRowNo).tRif.sRate(ilWk)
            '            tmRifRec(llRowNo).tRif.lRate(0) = 0
            '            tmRifRec(llRowNo).tRif.lRate(ilWk) = llDollar
            '        'Else
            '        '    'slStr = gDivStr(gMulStr(slDollar, Trim$(Str$(ilDay))), "7")
            '        '    'gStrToPDN slStr, 2, 5, tmRifRec(llRowNo).tRif.sRate(0)
            '        '    tmRifRec(llRowNo).tRif.lRate(0) = (llDollar * ilDay) / 7
            '        '    'slStr = gSubStr(slDollar, slStr)
            '        '    'gStrToPDN slStr, 2, 5, tmRifRec(llRowNo).tRif.sRate(ilWk)
            '        '    tmRifRec(llRowNo).tRif.lRate(ilWk) = llDollar - tmRifRec(llRowNo).tRif.lRate(0)
            '        'End If
            '    End If
            'ElseIf ilWk = 52 Then
            '    If rbcShow(1).Value Then    'Don't split dollars if input via standard
            '        'slStr = ".00"
            '        'gStrToPDN slStr, 2, 5, tmRifRec(llRowNo).tRif.sRate(53)
            '        'gStrToPDN slDollar, 2, 5, tmRifRec(llRowNo).tRif.sRate(ilWk)
            '        tmRifRec(llRowNo).tRif.lRate(53) = 0
            '        tmRifRec(llRowNo).tRif.lRate(ilWk) = llDollar
            '    Else
            '        'slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
            '        'slStart = gObtainEndCorp(slDate, True)
            '        'ilDay = gWeekDayStr(slStart)
            '        'If ilDay = 6 Then
            '            'slStr = ".00"
            '            'gStrToPDN slStr, 2, 5, tmRifRec(llRowNo).tRif.sRate(53)
            '            'gStrToPDN slDollar, 2, 5, tmRifRec(llRowNo).tRif.sRate(ilWk)
            '            tmRifRec(llRowNo).tRif.lRate(53) = 0
            '            tmRifRec(llRowNo).tRif.lRate(ilWk) = llDollar
            '        'Else
            '        '    ilDay = 7 - ilDay - 1
            '        '    'slStr = gDivStr(gMulStr(slDollar, Trim$(Str$(ilDay))), Trim$(Str$(ilDay)))
            '        '    'gStrToPDN slStr, 2, 5, tmRifRec(llRowNo).tRif.sRate(52)
            '        '    'slStr = gSubStr(slDollar, slStr)
            '        '    'gStrToPDN slStr, 2, 5, tmRifRec(llRowNo).tRif.sRate(53)
            '        '    tmRifRec(llRowNo).tRif.lRate(52) = (llDollar * ilDay) / 7
            '        '    tmRifRec(llRowNo).tRif.lRate(53) = llDollar - tmRifRec(llRowNo).tRif.lRate(52)
            '        'End If
            '    End If
            'Else
            '    'gStrToPDN slDollar, 2, 5, tmRifRec(llRowNo).tRif.sRate(ilWk)
            '    tmRifRec(llRowNo).tRif.lRate(ilWk) = llDollar
            'End If
        Next ilWk
    Else
        llLkYear = tmRifRec(llRowNo).lLkYear
        Do While llLkYear > 0
            If tmLkRifRec(llLkYear).tRif.iYear = tmPdGroups(ilGroup).iYear Then
                For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                    slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                    If rbcShow(0).Value Then
                        slStart = Format$(gDateValue(tmPdGroups(ilGroup).sStartDate) + (ilWk - tmPdGroups(ilGroup).iStartWkNo) * 7, "m/d/yy")
                        mRifSetRate CLng(llRowNo), slStart, llDollar, tmRifRec(), tmLkRifRec()
                    Else
                        tmLkRifRec(llLkYear).tRif.lRate(ilWk) = llDollar
                    End If
                    'If ilWk = 1 Then
                    '    If rbcShow(0).Value Then    'Don't split dollars if input via corporate
                    '        'slStr = ".00"
                    '        'gStrToPDN slStr, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(0)
                    '        'gStrToPDN slDollar, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(ilWk)
                    '        tmLkRifRec(llLkYear).tRif.lRate(0) = 0
                    '        tmLkRifRec(llLkYear).tRif.lRate(ilWk) = llDollar
                    '    Else
                    '        slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                    '        slStart = gObtainStartCorp(slDate, True)
                    '        ilDay = gWeekDayStr(slStart)
                    '        'If ilDay = 0 Then
                    '        '    'slStr = ".00"
                    '        '    'gStrToPDN slStr, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(0)
                    '        '    'gStrToPDN slDollar, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(ilWk)
                    '            tmLkRifRec(llLkYear).tRif.lRate(0) = 0
                    '            tmLkRifRec(llLkYear).tRif.lRate(ilWk) = llDollar
                    '        'Else
                    '        '    'slStr = gDivStr(gMulStr(slDollar, Trim$(Str$(ilDay))), "7")
                    '        '    'gStrToPDN slStr, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(0)
                    '        '    'slStr = gSubStr(slDollar, slStr)
                    '        '    'gStrToPDN slStr, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(ilWk)
                    '        '    tmLkRifRec(llLkYear).tRif.lRate(0) = (llDollar * ilDay) / 7
                    '        '    tmLkRifRec(llLkYear).tRif.lRate(ilWk) = llDollar - tmLkRifRec(llLkYear).tRif.lRate(0)
                    '        'End If
                    '    End If
                    'ElseIf ilWk = 52 Then
                    '    If rbcShow(1).Value Then    'Don't split dollars if input via standard
                    '        'slStr = ".00"
                    '        'gStrToPDN slStr, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(53)
                    '        'gStrToPDN slDollar, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(ilWk)
                    '        tmLkRifRec(llLkYear).tRif.lRate(53) = 0
                    '        tmLkRifRec(llLkYear).tRif.lRate(ilWk) = llDollar
                    '    Else
                    '        'slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                    '        'slStart = gObtainEndCorp(slDate, True)
                    '        'ilDay = gWeekDayStr(slStart)
                    '        'If ilDay = 6 Then
                    '            'slStr = ".00"
                    '            'gStrToPDN slStr, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(53)
                    '            'gStrToPDN slDollar, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(ilWk)
                    '            tmLkRifRec(llLkYear).tRif.lRate(53) = 0
                    '            tmLkRifRec(llLkYear).tRif.lRate(ilWk) = llDollar
                    '        'Else
                    '        '    ilDay = 7 - ilDay - 1
                    '        '    'slStr = gDivStr(gMulStr(slDollar, Trim$(Str$(ilDay))), Trim$(Str$(ilDay)))
                    '        '    'gStrToPDN slStr, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(52)
                    '        '    'slStr = gSubStr(slDollar, slStr)
                    '        '    'gStrToPDN slStr, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(53)
                    '        '    tmLkRifRec(llLkYear).tRif.lRate(52) = (llDollar * ilDay) / 7
                    '        '    tmLkRifRec(llLkYear).tRif.lRate(53) = llDollar - tmLkRifRec(llLkYear).tRif.lRate(52)
                    '        'End If
                    '    End If
                    'Else
                    '    'gStrToPDN slDollar, 2, 5, tmLkRifRec(llLkYear).tRif.sRate(ilWk)
                    '    tmLkRifRec(llLkYear).tRif.lRate(ilWk) = llDollar
                    'End If
                Next ilWk
                Exit Do
            Else
                llLkYear = tmLkRifRec(llLkYear).lLkYear
            End If
        Loop
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mShowRCInfo                     *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show Rate Card information for *
'*                      right mouse                    *
'*                                                     *
'*******************************************************
Private Sub mShowRCInfo()
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilIndex As Integer
    Dim ilVsf As Integer
    Dim ilButtonIndex As Integer
    ilButtonIndex = imButtonIndex
    If (imButtonIndex < LBONE) Or (imButtonIndex > UBound(tmRifRec) - 1) Then
        plcRCInfo.Visible = False
        Exit Sub
    End If
    slStr = "Vehicle Name " & Trim$(smRCSave(VEHINDEX, imButtonIndex))
    'Read Rdf and show info
    gFindMatch Trim$(smRCSave(DAYPARTINDEX, imButtonIndex)), 0, lbcDPName
    ilIndex = gLastFound(lbcDPName)
    slNameCode = lbcDPNameCode.List(ilIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilCode = Val(Trim$(slCode))
    For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
        If ilCode = tgMRdf(ilLoop).iCode Then
            tmRdf = tgMRdf(ilLoop)
            Exit For
        End If
    Next ilLoop
    'ilCount = 0
    'For ilLoop = 0 To 2 Step 1
    '    If Trim$(smSave(2 + ilLoop, ilButtonIndex)) <> "" Then
    '        If ilCount > 0 Then
    '            slStr = slStr & "/"
    '        Else
    '            slStr = slStr & "  Library Names "
    '        End If
    '        ilCount = ilCount + 1
    '        slStr = slStr & Trim$(smSave(2 + ilLoop, ilButtonIndex))
    '    End If
    'Next ilLoop
    'slStr = slStr & "  Name " & Trim$(smSave(5, ilButtonIndex))
    lacRCInfo(0).Caption = slStr & "  Daypart Name " & Trim$(tmRdf.sName)
    Select Case tmRdf.sInOut
        Case "N"  'All avails
            slStr = "Book into: All avails"
        Case "I"  'Book into selected avail
            tmAnfSrchKey.iCode = tmRdf.ianfCode
            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            slStr = "Book into: " & Trim$(tmAnf.sName) '& Trim$(smSave(6, ilButtonIndex))
        Case "O"  'Exclused selected avail
            tmAnfSrchKey.iCode = tmRdf.ianfCode
            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            slStr = "Book into Except: " & Trim$(tmAnf.sName) '& Trim$(smSave(6, ilButtonIndex))    'Alt+0216 or 0248
        Case Else
            slStr = ""
    End Select
    If tmRdf.sBase = "Y" Then
        slStr = slStr & "   Base Daypart: Yes"
    Else
        slStr = slStr & "   Base Daypart: No"
    End If
    lacRCInfo(1).Caption = slStr
    gFindMatch Trim$(smRCSave(VEHINDEX, ilButtonIndex)), 0, lbcVehicle
    ilIndex = gLastFound(lbcVehicle)
    lacRCInfo(2).Visible = False
    'If ilIndex < Traffic!lbcUserVehicle.ListCount Then
    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
    If ilIndex >= 0 And ilIndex < UBound(tgTempRCUserVehicle) Then
        slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilCode = Val(Trim$(slCode))
        For ilLoop = LBound(igVirtVefCode) To UBound(igVirtVefCode) - 1 Step 1
            If ilCode = igVirtVefCode(ilLoop) Then
                tmVefSrchKey.iCode = ilCode
                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) And (tmVef.sType = "V") Then
                    slStr = ""
                    tmVsfSrchKey.lCode = tmVef.lVsfCode
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If (ilRet = BTRV_ERR_NONE) Then
                        For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                            If tmVsf.iFSCode(ilVsf) > 0 Then
                                tmVefSrchKey.iCode = tmVsf.iFSCode(ilVsf)
                                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If (ilRet = BTRV_ERR_NONE) Then
                                    If slStr = "" Then
                                        slStr = "Virtual Vehicles: " & Trim$(tmVef.sName)
                                    Else
                                        slStr = slStr & "/" & Trim$(tmVef.sName)
                                    End If
                                End If
                            End If
                        Next ilVsf
                    End If
                    lacRCInfo(2).Caption = slStr
                    lacRCInfo(2).Visible = True
                End If
                Exit For
            End If
        Next ilLoop
    End If
    If (imButtonIndex < LBONE) Or (imButtonIndex > UBound(tmRifRec) - 1) Then
        plcRCInfo.Visible = False
        Exit Sub
    End If
    plcRCInfo.ZOrder vbBringToFront
    plcRCInfo.Visible = True
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSPEnableBox                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSPEnableBox(ilBoxNo As Integer)
'
'   mRCEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBSPCtrls) Or (ilBoxNo > UBound(tmSPCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        'Case BUDGETINDEX 'Vehicle
        '    lbcBudget.Height = gListBoxHeight(lbcBudget.ListCount, 20)
        '    edcSPDropDown.Width = tmSPCtrls(BUDGETINDEX).fBoxW - cmcSPDropDown.Width
        '    edcSPDropDown.MaxLength = 20
        '    gMoveFormCtrl pbcSP, edcSPDropDown, tmSPCtrls(BUDGETINDEX).fBoxX, tmSPCtrls(BUDGETINDEX).fBoxY
        '    cmcSPDropDown.Move edcSPDropDown.Left + edcSPDropDown.Width, edcSPDropDown.Top
        '    imChgMode = True
        '    If lbcBudget.ListIndex >= 0 Then
        '        imComboBoxIndex = lbcBudget.ListIndex
        '        edcSPDropDown.Text = lbcBudget.List(lbcBudget.ListIndex)
        '    Else
        '        lbcBudget.ListIndex = 0
        '        imComboBoxIndex = lbcBudget.ListIndex
        '        edcSPDropDown.Text = lbcBudget.List(0)
        '    End If
        '    imChgMode = False
        '    lbcBudget.Move edcSPDropDown.Left, edcSPDropDown.Top + edcSPDropDown.Height
        '    edcSPDropDown.SelStart = 0
        '    edcSPDropDown.SelLength = Len(edcSPDropDown.Text)
        '    edcSPDropDown.Visible = True
        '    cmcSPDropDown.Visible = True
        '    edcSPDropDown.SetFocus
        Case GRIDINDEX 'Grid
            lbcGrid.Height = gListBoxHeight(lbcGrid.ListCount, 20)
            edcSPDropDown.Width = tmSPCtrls(GRIDINDEX).fBoxW - cmcSPDropDown.Width
            edcSPDropDown.MaxLength = 2
            gMoveFormCtrl pbcSP, edcSPDropDown, tmSPCtrls(GRIDINDEX).fBoxX, tmSPCtrls(GRIDINDEX).fBoxY
            cmcSPDropDown.Move edcSPDropDown.Left + edcSPDropDown.Width, edcSPDropDown.Top
            imChgMode = True
            If imGDSelectedIndex >= 0 Then
                lbcGrid.ListIndex = imGDSelectedIndex
                imComboBoxIndex = imGDSelectedIndex
                edcSPDropDown.Text = lbcGrid.List(imGDSelectedIndex)
            Else
                lbcGrid.ListIndex = 0
                imComboBoxIndex = lbcGrid.ListIndex
                edcSPDropDown.Text = lbcGrid.List(0)
            End If
            imChgMode = False
            lbcGrid.Move edcSPDropDown.Left, edcSPDropDown.Top + edcSPDropDown.Height
            edcSPDropDown.SelStart = 0
            edcSPDropDown.SelLength = Len(edcSPDropDown.Text)
            edcSPDropDown.Visible = True
            cmcSPDropDown.Visible = True
            edcSPDropDown.SetFocus
        Case LENGTHINDEX 'Length
            lbcLen.Height = gListBoxHeight(lbcLen.ListCount, 20)
            edcSPDropDown.Width = tmSPCtrls(GRIDINDEX).fBoxW
            edcSPDropDown.MaxLength = 3
            gMoveFormCtrl pbcSP, edcSPDropDown, tmSPCtrls(LENGTHINDEX).fBoxX, tmSPCtrls(LENGTHINDEX).fBoxY
            cmcSPDropDown.Move edcSPDropDown.Left + edcSPDropDown.Width, edcSPDropDown.Top
            imChgMode = True
            If imLenSelectedIndex >= 0 Then
                lbcLen.ListIndex = imLenSelectedIndex
                imComboBoxIndex = imLenSelectedIndex
                edcSPDropDown.Text = lbcLen.List(imLenSelectedIndex)
            Else
                gFindMatch Str$(tgRcfI.iBaseLen), 0, lbcLen
                If gLastFound(lbcLen) >= 0 Then
                    lbcLen.ListIndex = gLastFound(lbcLen)
                Else
                    lbcLen.ListIndex = 0
                End If
                imComboBoxIndex = lbcLen.ListIndex
                edcSPDropDown.Text = lbcLen.List(0)
            End If
            imChgMode = False
            lbcLen.Move edcSPDropDown.Left, edcSPDropDown.Top + edcSPDropDown.Height
            edcSPDropDown.SelStart = 0
            edcSPDropDown.SelLength = Len(edcSPDropDown.Text)
            edcSPDropDown.Visible = True
            cmcSPDropDown.Visible = True
            edcSPDropDown.SetFocus
        Case CURGRIDINDEX 'Vehicle
            lbcGrid.Height = gListBoxHeight(lbcGrid.ListCount, 20)
            edcSPDropDown.Width = tmSPCtrls(CURGRIDINDEX).fBoxW - cmcSPDropDown.Width
            edcSPDropDown.MaxLength = 2
            gMoveFormCtrl pbcSP, edcSPDropDown, tmSPCtrls(CURGRIDINDEX).fBoxX, tmSPCtrls(CURGRIDINDEX).fBoxY
            cmcSPDropDown.Move edcSPDropDown.Left + edcSPDropDown.Width, edcSPDropDown.Top
            imChgMode = True
            If imCGDSelectedIndex >= 0 Then
                lbcGrid.ListIndex = imCGDSelectedIndex
                imComboBoxIndex = imCGDSelectedIndex
                edcSPDropDown.Text = lbcGrid.List(imCGDSelectedIndex)
            Else
                lbcGrid.ListIndex = 0
                imComboBoxIndex = lbcGrid.ListIndex
                edcSPDropDown.Text = lbcGrid.List(0)
            End If
            imChgMode = False
            lbcGrid.Move edcSPDropDown.Left, edcSPDropDown.Top + edcSPDropDown.Height
            edcSPDropDown.SelStart = 0
            edcSPDropDown.SelLength = Len(edcSPDropDown.Text)
            edcSPDropDown.Visible = True
            cmcSPDropDown.Visible = True
            edcSPDropDown.SetFocus
    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSPSetFocus                     *
'*                                                     *
'*             Created:7/15/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus                      *
'*                                                     *
'*******************************************************
Private Sub mSPSetFocus()
    If imSPBoxNo >= imLBSPCtrls And imSPBoxNo <= UBound(tmSPCtrls) Then
        Select Case imSPBoxNo 'Branch on box type (control)
            'Case BUDGETINDEX 'Vehicle
            '    edcSPDropDown.SetFocus
            Case GRIDINDEX 'Grid
                edcSPDropDown.SetFocus
            Case LENGTHINDEX 'Length
                edcSPDropDown.SetFocus
            Case CURGRIDINDEX 'Grid
                edcSPDropDown.SetFocus
        End Select
        Exit Sub
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSPSetShow                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSPSetShow(ilBoxNo As Integer)
'
'   mSPSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBSPCtrls) Or (ilBoxNo > UBound(tmSPCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        'Case BUDGETINDEX 'Vehicle
        '    lbcBudget.Visible = False
        '    edcSPDropDown.Visible = False
        '    cmcSPDropDown.Visible = False
        '    slStr = edcSPDropDown.Text
        '    gSetShow pbcSP, slStr, tmSPCtrls(ilBoxNo)
        '    'Change prices
        Case GRIDINDEX 'Name
            lbcGrid.Visible = False
            edcSPDropDown.Visible = False
            cmcSPDropDown.Visible = False
            imGDSelectedIndex = lbcGrid.ListIndex
            slStr = edcSPDropDown.Text
            gSetShow pbcSP, slStr, tmSPCtrls(ilBoxNo)
        Case LENGTHINDEX
            lbcLen.Visible = False
            edcSPDropDown.Visible = False
            cmcSPDropDown.Visible = False
            imLenSelectedIndex = lbcLen.ListIndex
            slStr = edcSPDropDown.Text
            gSetShow pbcSP, slStr, tmSPCtrls(ilBoxNo)
        Case CURGRIDINDEX 'Name
            lbcGrid.Visible = False
            edcSPDropDown.Visible = False
            cmcSPDropDown.Visible = False
            imCGDSelectedIndex = lbcGrid.ListIndex
            slStr = edcSPDropDown.Text
            gSetShow pbcSP, slStr, tmSPCtrls(ilBoxNo)
    End Select
    mSetCommands
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
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim llLoop As Long
    Dim ilRet As Integer
    Dim slVehName As String
    Dim slName As String
    Dim ilTest As Integer
    Dim slDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilNoWks As Integer
    'If Not gWinRoom(igNoExeWinRes(RCTERMSEXE)) Then
    '    mStartNew = False
    '    Exit Function
    'End If
    If imAdjIndex = 0 Then
        igRcfModel = 0
        mStartNew = True
        imInNew = False
        Exit Function
    End If
    imInNew = True
    If (cbcSelect.ListCount > 1) Then
        RCModel.Show vbModal
        If igRCReturn = 0 Then    'Cancelled
            mStartNew = False
            imInNew = False
            Exit Function
        End If
    Else
        igRcfModel = 0
    End If
    'Ask Model question
    Screen.MousePointer = vbHourglass    '
    RCTerms.Show vbModal
    Screen.MousePointer = vbDefault    'Default
    If igRCReturn = 0 Then    'Cancelled
        mStartNew = False
        imInNew = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass    '
    'Build program images from newest
    ilRet = mReadRifRec(igRcfModel, False)   'Ok to pass zero ([None])
    If Not ilRet Then
        Screen.MousePointer = vbDefault   '
        mStartNew = False
        imInNew = False
        Exit Function
    End If
    tgRcfI.iCode = 0
    slVehName = ""
    If (tgRcfI.iVefCode <> 0) And (tgRcfI.iVefCode <> -32000) Then
        'If tgRcfI.iVefCode > 0 Then
            'For ilTest = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
            '    If tmUserVeh(ilTest).iCode = tgRcfI.iVefCode Then
                ilTest = mBinarySearch(tgRcfI.iVefCode)
                If ilTest <> -1 Then
                    slVehName = Trim$(tmUserVeh(ilTest).sName)
                End If
            '        Exit For
            '    End If
            'Next ilTest
        'ElseIf tgRcfI.iVefCode < 0 Then
        '    slRecCode = Trim$(Str$(-tgRcfI.iVefCode))
        '    For ilLoop = 0 To lbcCombo.ListCount - 1 Step 1
        '        slNameCode = lbcCombo.List(ilLoop)
        '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        '        On Error GoTo mStartNewErr
        '        gCPErrorMsg ilRet, "mStartNew (gParseItem field 2)", RateCard
        '        On Error GoTo 0
        '        If slRecCode = slCode Then
        '            ilRet = gParseItem(slNameCode, 1, "\", slVehName)
        '            On Error GoTo mStartNewErr
        '            gCPErrorMsg ilRet, "mStartNew (gParseItem field 1)", RateCard
        '            On Error GoTo 0
        '            Exit For
        '        End If
        '    Next ilLoop
        'End If
    ElseIf tgRcfI.iVefCode = 0 Then
        slVehName = "[All Vehicles]"
    End If
    slVehName = Trim$(slVehName)
    If slVehName = "" Then
        smDefVehicle = Trim$(sgUserDefVehicleName)
    Else
        smDefVehicle = slVehName
    End If
    For llLoop = LBONE To UBound(tmRifRec) - 1 Step 1
        tmRifRec(llLoop).tRif.iYear = tgRcfI.iYear
        tmRifRec(llLoop).lRecPos = 0
        tmRifRec(llLoop).iType = 0
        tmRifRec(llLoop).iStatus = 0
    Next llLoop
'    gUnpackDate tgRcfI.iStartDate(0), tgRcfI.iStartDate(1), slStr
'    If slStr <> "" Then
'        slStr = gAddDayToDate(slStr)
'    End If
'    plcStartDate.Caption = slStr
'    plcStartDate.ZOrder vbBringToFront
    cbcSelect.Visible = False
    plcSP.Visible = False
    plcRateCard.Visible = False
    mInitRateCardCtrls  'Initial arrays
    mTestCorpYear tgRcfI.iYear
    If UBound(tmRifRec) = 1 Then
        mSetDefInSave
        mGetShowDates
    End If
    mInitRif UBound(tmRifRec)
    'For ilLoop = LBound(tmRifRec) To UBound(tmRifRec) - 1 Step 1
        'tmRif(ilLoop).iCode = 0
        'tgRpfI(ilLoop).iStartDate(0) = tgRcfI.iStartDate(0)
        'tgRpfI(ilLoop).iStartDate(1) = tgRcfI.iStartDate(1)
        'tgRpfI(ilLoop).iEndDate(0) = tgRcfI.iEndDate(0)
        'tgRpfI(ilLoop).iEndDate(1) = tgRcfI.iEndDate(1)
    'Next ilLoop
    llLoop = LBONE  'LBound(tmRifRec)
    slDate = "1/15/" & Trim$(Str$(tmRifRec(llLoop).tRif.iYear))
    llStartDate = gDateValue(gObtainYearStartDate(0, slDate))
    slDate = "12/15/" & Trim$(Str$(tmRifRec(llLoop).tRif.iYear))
    llEndDate = gDateValue(gObtainYearEndDate(0, slDate))
    ilNoWks = (llEndDate - llStartDate + 1) \ 7
    If ilNoWks = 53 Then
        For llLoop = LBONE To UBound(tmRifRec) - 1 Step 1
            If tmRifRec(llLoop).tRif.lRate(53) = 0 Then
                tmRifRec(llLoop).tRif.lRate(53) = tmRifRec(llLoop).tRif.lRate(52)
            End If
        Next llLoop
    Else
        For llLoop = LBONE To UBound(tmRifRec) - 1 Step 1
            If tmRifRec(llLoop).tRif.lRate(53) > 0 Then
                tmRifRec(llLoop).tRif.lRate(53) = 0
            End If
        Next llLoop
    End If
    mMoveRecToCtrl
    'mInitShow
    imChgMode = True    'Set change mode to avoid infinite loop
    slName = Trim$(tgRcfI.sName) & Str$(tgRcfI.iYear) & "/" & Trim$(slVehName)
    cbcSelect.AddItem slName, 1
    cbcSelect.ListIndex = 1
    imSelectedIndex = 1
    cbcSelect.Visible = True
    'plcSP.Visible = True   'Temporary removed until spec is to show 7/10/97
    plcRateCard.Visible = True
    pbcRateCard_Paint
    imChgMode = False
    mStartNew = True
    mSetCommands
    Screen.MousePointer = vbDefault    '
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
'*             Created:6/30/93       By:D. LeVine      *
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
    mUrfUpdate RateCard, tgUrf()


    imTerminate = False

    Screen.MousePointer = vbDefault
    'Unload IconTraf
    igManUnload = YES
    Unload RateCard
    igManUnload = NO
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestCorpYear                   *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test corporate year            *
'*                                                     *
'*******************************************************
Private Sub mTestCorpYear(ilInYear As Integer)
    Dim ilYear As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    ilYear = ilInYear
    If ilYear < 100 Then
        If ilYear >= 70 Then
            ilYear = 1900 + ilYear
        Else
            ilYear = 2000 + ilYear
        End If
    End If
    If tgSpf.sRUseCorpCal = "Y" Then
        ilFound = False
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If tgMCof(ilLoop).iYear = ilYear Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            MsgBox "Corporate Year Missing for" & Str$(ilYear), vbOKOnly + vbExclamation, "Rate Card"
            imIgnoreSetting = True
            rbcShow(1).Value = True
            rbcShow(0).Enabled = False
        Else
            rbcShow(0).Enabled = True
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields() As Integer
'
'   iRet = mTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilIndex As Integer
    Dim llRif As Long
    Dim llLoop1 As Long
    Dim llLoop2 As Long
    Dim slVehicleName As String
    Dim slDayPartName As String

    For llRif = LBONE To UBound(tmRifRec) - 1 Step 1
        If tmRifRec(llRif).tRif.iVefCode = 0 Then
            ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
            lmRCRowNo = llRif
            imRCBoxNo = VEHINDEX
            mTestFields = NO
            Exit Function
        End If
        If tmRifRec(llRif).tRif.iRdfCode <= 0 Then
            ilRes = MsgBox("Daypart must be specified", vbOKOnly + vbExclamation, "Incomplete")
            lmRCRowNo = llRif
            imRCBoxNo = DAYPARTINDEX
            mTestFields = NO
            Exit Function
        End If
    Next llRif
    For llLoop1 = UBound(tmRifRec) - 1 To LBONE Step -1
        For llLoop2 = llLoop1 - 1 To LBONE Step -1
            If (tmRifRec(llLoop1).tRif.iRdfCode = tmRifRec(llLoop2).tRif.iRdfCode) And (tmRifRec(llLoop1).tRif.iVefCode = tmRifRec(llLoop2).tRif.iVefCode) Then
                ilIndex = gBinarySearchVef(tmRifRec(llLoop1).tRif.iVefCode)
                If ilIndex <> -1 Then
                    slVehicleName = Trim$(tgMVef(ilIndex).sName)
                Else
                    slVehicleName = ""
                End If
                ilIndex = gBinarySearchRdf(tmRifRec(llLoop1).tRif.iRdfCode)
                If ilIndex <> -1 Then
                    slDayPartName = Trim$(tgMRdf(ilIndex).sName)
                Else
                    slDayPartName = ""
                End If

                ilRes = MsgBox("Duplicate Dayparts not allowed, found for " & slVehicleName & " " & slDayPartName, vbOKOnly + vbExclamation, "Incomplete")
                lmRCRowNo = llLoop1
                imRCBoxNo = DAYPARTINDEX
                mTestFields = NO
                Exit Function
            End If
        Next llLoop2
    Next llLoop1
    mTestFields = YES
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields() As Integer
'
'   iRet = mTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    'Dim ilRpf As Integer
    If Trim$(smRCSave(VEHINDEX, lmRCRowNo)) = "" Then
        ilRes = MsgBox("Vehicle must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imRCBoxNo = VEHINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smRCSave(DAYPARTINDEX, lmRCRowNo)) = "" Then
        ilRes = MsgBox("Daypart must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imRCBoxNo = DAYPARTINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smRCSave(BASEINDEX, lmRCRowNo)) = "" Then
        ilRes = MsgBox("Base must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imRCBoxNo = BASEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If Trim$(smRCSave(RPTINDEX, lmRCRowNo)) = "" Then
        ilRes = MsgBox("Show on Report must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imRCBoxNo = RPTINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mUrfUpdate                      *
'*                                                     *
'*             Created:5/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update User calendar/calculator *
'*                     related fields in the records   *
'*                                                     *
'*******************************************************
Private Sub mUrfUpdate(frm As Form, tlUrf() As URF)
'
'   mUrfUpdate MainForm, tlUrf()
'   Where:
'       MainForm (I)- Name of Form to unload if error exists
'       tlUrf (O)- the updated user records
'                   Note: tlUrf must be defined as Dim tlUrf() as URF
'
    Dim ilRecLen As Integer     'URF record length
    Dim hlUrf As Integer        'User Option file handle
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim tlUrfSet As URF    'Position to record so it can be updated
    Dim tlSrchKey As INTKEY0    'URF key record image
    hlUrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mUrfUpdateErr
    gBtrvErrorMsg ilRet, "mUrfUpdate (btrOpen):" & "Urf.Btr", frm
    On Error GoTo 0
    On Error GoTo gUrfNoDefinedErr
    ilRecLen = Len(tlUrf(0))  'btrRecordLength(hlUrf)  'Get and save record length
    If tlUrf(0).iCode <= 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    For ilLoop = LBound(tlUrf) To UBound(tlUrf) Step 1
        tlSrchKey.iCode = tlUrf(ilLoop).iCode
        ilRet = btrGetEqual(hlUrf, tlUrfSet, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        gUrfDecrypt tlUrfSet
        On Error GoTo mUrfUpdateErr
        gBtrvErrorMsg ilRet, "mUrfUpdate (btrGetEqual)", frm
        On Error GoTo 0
        If imView = 1 Then
            tlUrf(ilLoop).sRCView = "D"  'Daypart view
        Else
            tlUrf(ilLoop).sRCView = "R"  'Rate View
        End If
        gUrfEncrypt tlUrf(ilLoop)
        ilRet = btrUpdate(hlUrf, tlUrf(ilLoop), ilRecLen)
        On Error GoTo mUrfUpdateErr
        gBtrvErrorMsg ilRet, "mUrfUpdate (btrUpdate)", frm
        On Error GoTo 0
        gUrfDecrypt tlUrf(ilLoop)
    Next ilLoop
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    Exit Sub
mUrfUpdateErr:
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    'Error ERRORCODEBASE
    Exit Sub
gUrfNoDefinedErr:
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    Exit Sub
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
Private Sub mVehPop(cbcCtrl As Control)
    Dim ilRet As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim slName As String
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    'ilRet = gPopUserVehicleBox(RateCard, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, cbcCtrl, Traffic!lbcUserVehicle)
    'ilRet = gPopUserVehicleBox(RateCard, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcCtrl, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(RateCard, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHSTDPKG + VEHCPMPKG + VEHSPORT + ACTIVEVEH, cbcCtrl, tgRCUserVehicle(), sgRCUserVehicleTag)
    'ilRet = gPopUserVehComboBox(RateCard, cbcCtrl, Traffic!lbcUserVehicle, lbcCombo)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gPopUserVehComboBox: Vehicle/Combo)", RateCard
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", RateCard
        On Error GoTo 0
    End If
    'ReDim tmUserVeh(1 To Traffic!lbcUserVehicle.ListCount + 1) As USERVEH
    ReDim tmUserVeh(0 To UBound(tgRCUserVehicle) - LBound(tgRCUserVehicle)) As USERVEH
    ilUpper = 0
    For ilLoop = LBound(tgRCUserVehicle) To UBound(tgRCUserVehicle) - 1 Step 1 'Traffic!lbcUserVehicle.ListCount - 1 Step 1
        slNameCode = tgRCUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", tmUserVeh(ilUpper).sName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmUserVeh(ilUpper).iCode = Val(slCode)
        ilUpper = ilUpper + 1
    Next ilLoop
    If UBound(tmUserVeh) > 1 Then
        'ArraySortTyp fnAV(tmUserVeh(), 1), UBound(tmUserVeh) - 1, 0, LenB(tmUserVeh(1)), 0, -1, 0
        ArraySortTyp fnAV(tmUserVeh(), 0), UBound(tmUserVeh), 0, LenB(tmUserVeh(0)), 0, -1, 0
    End If
    
    'TTP 10388 - Rate card screen: if "ad server" site option setting is not enabled, a "podcast" medium type vehicle does not appear on the rate card screen
    ''L.Bianchi
    'mRemoveCPMVehicle
    ReDim tgTempRCUserVehicle(UBound(tgRCUserVehicle))
    tgTempRCUserVehicle = tgRCUserVehicle
    
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub pbcArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcDaypart_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
End Sub

Private Sub pbcDaypart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim llRowNo As Long
    Dim slStr As String
    Dim ilLoop As Integer
    If imView = 1 Then
        Exit Sub
    End If
    If Button = 2 Then
        imButtonIndex = -1
        plcRCInfo.Visible = False
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    ilCompRow = vbcRateCard.LargeChange + 1
    If UBound(tmRifRec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tmRifRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBRCCtrls To UBound(tmRCCtrls) Step 1
            If (X >= tmRCCtrls(ilBox).fBoxX) And (X <= (tmRCCtrls(ilBox).fBoxX + tmRCCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(ilBox).fBoxY + tmRCCtrls(ilBox).fBoxH)) Then
                    If ilBox > DAYPARTINDEX Then
                        Beep
                        Exit Sub
                    End If
                    llRowNo = ilRow + vbcRateCard.Value - 1
                    If llRowNo > UBound(tmRifRec) Then
                        Beep
                        If imSPBoxNo > 0 Then
                            mSPSetFocus
                        Else
                            mSetFocus imRCBoxNo
                        End If
                    End If
                    If (lmRCRowNo = UBound(tmRifRec)) And (imRCBoxNo = VEHINDEX) And (llRowNo <> lmRCRowNo) Then
                        For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
                            slStr = ""
                            gSetShow pbcRateCard, slStr, tmRCCtrls(ilLoop)
                            smRCShow(ilLoop, lmRCRowNo) = tmRCCtrls(ilLoop).sShow
                        Next ilLoop
                        pbcRateCard_Paint
                        mSetDefInSave   'Set defaults for extra row
                    End If
                    If (Trim$(smRCSave(VEHINDEX, llRowNo)) = "") And (ilBox <> VEHINDEX) Then
                        Beep
                        ilBox = VEHINDEX
                    End If
                    mSPSetShow imSPBoxNo
                    imSPBoxNo = -1
                    mRCSetShow imRCBoxNo
                    lmRCRowNo = ilRow + vbcRateCard.Value - 1
                    imRCBoxNo = ilBox
                    mRCEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    If imSPBoxNo > 0 Then
        mSPSetFocus
    Else
        mSetFocus imRCBoxNo
    End If
End Sub

Private Sub pbcDaypart_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim slFont As String
    Dim llColor As Long

    mPaintDPTitle
    ilStartRow = vbcRateCard.Value '+ 1  'Top location
    ilEndRow = vbcRateCard.Value + vbcRateCard.LargeChange ' + 1
    If ilEndRow > UBound(tmRifRec) Then
        ilEndRow = UBound(tmRifRec) 'include blank row as it might have data
    End If
    llColor = pbcDaypart.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBDPCtrls To UBound(tmDPCtrls) - 1 Step 1
            If (ilBox = TIMESINDEX) And (right$(Trim$(smDPShow(ilBox, ilRow)), 1) = "+") Then
                If pbcDaypart.TextWidth(Trim$(smDPShow(ilBox, ilRow))) > tmDPCtrls(TIMESINDEX).fBoxW - fgBoxInsetX Then
                    slStr = Left$(Trim$(smDPShow(ilBox, ilRow)), Len(Trim$(smDPShow(ilBox, ilRow))) - 2) & "+"
                Else
                    slStr = Trim$(smDPShow(ilBox, ilRow))
                End If
            Else
                slStr = Trim$(smDPShow(ilBox, ilRow))
            End If
            If (ilBox >= DAYINDEX) And (ilBox <= DAYINDEX + 6) Then
                gPaintArea pbcDaypart, tmDPCtrls(ilBox).fBoxX, tmDPCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmDPCtrls(ilBox).fBoxW - 15, tmDPCtrls(ilBox).fBoxH - 15, WHITE
                pbcDaypart.CurrentX = tmDPCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcDaypart.CurrentY = tmDPCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '- 15'- 30'+ fgBoxInsetY
                slFont = pbcDaypart.FontName
                pbcDaypart.FontName = "Monotype Sorts"
                pbcDaypart.FontBold = False
                If slStr = "Y" Then
                    slStr = "4"
                Else
                    slStr = "  "
                End If
                pbcDaypart.Print slStr
                pbcDaypart.FontName = slFont
                pbcDaypart.FontBold = True
            Else
                'If imUsedByCntr(ilRow - 1) And (ilBox <> NAMEINDEX) And (ilBox <> HRSINDEX) And (ilBox <> PRICEINDEX) And (ilBox <> TIMEDAYINDEX) Then
                '    pbcDaypart.ForeColor = RED
                'Else
                '    pbcDaypart.ForeColor = llColor
                'End If
                If (ilBox = AVAILINDEX) And (Left$(slStr, 1) = "�") Then
                    gPaintArea pbcDaypart, tmDPCtrls(ilBox).fBoxX, tmDPCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmDPCtrls(ilBox).fBoxW - 15, tmDPCtrls(ilBox).fBoxH - 15, WHITE
                    pbcDaypart.CurrentX = tmDPCtrls(ilBox).fBoxX + fgBoxInsetX
                    pbcDaypart.CurrentY = tmDPCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 15 '- 30'+ fgBoxInsetY
                    slFont = pbcDaypart.FontName
                    pbcDaypart.FontName = lbcDPName.FontName '"MS Sans Serif"
                    pbcDaypart.FontBold = False
                    pbcDaypart.Print Left$(slStr, 1)
                    pbcDaypart.FontName = slFont
                    pbcDaypart.FontBold = True
                    pbcDaypart.CurrentX = tmDPCtrls(ilBox).fBoxX + fgBoxInsetX + pbcDaypart.TextWidth("� ")
                    pbcDaypart.CurrentY = tmDPCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    pbcDaypart.Print right$(slStr, Len(slStr) - 1)
                Else
                    If (ilBox = DAYPARTINDEX) Then  'And (smDPShow(DPBASEINDEX, ilRow) = "Y") Then
                        If Trim$(smRCSave(BASEINDEX, ilRow)) = "Y" Then
                            pbcDaypart.ForeColor = MAGENTA
                        ElseIf (Trim$(smRCSave(BASEINDEX, ilRow)) <> "N") And (Trim$(smDPShow(DPBASEINDEX, ilRow)) = "Y") Then
                            pbcDaypart.ForeColor = MAGENTA
                        End If
                    End If
                    If (ilBox = VEHINDEX) And (imRCSave(11, ilRow) > 0) Then
                        pbcDaypart.ForeColor = BLUE
                    End If
                    If (ilBox = VEHINDEX) And (imRCSave(9, ilRow) > 0) Then
                        pbcDaypart.ForeColor = Red
                    End If
                    If (ilBox = DAYPARTINDEX) And (imRCSave(10, ilRow) > 0) Then
                        pbcDaypart.ForeColor = Red
                    End If
                    'added by L. Bianchi
                    If (ilBox = VEHINDEX) And (imRCSave(16, ilRow) > 0 Or imRCSave(17, ilRow) = 1) Then
                         pbcDaypart.FontItalic = True
                    Else
                         pbcDaypart.FontItalic = False
                    End If
                    gPaintArea pbcDaypart, tmDPCtrls(ilBox).fBoxX, tmDPCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmDPCtrls(ilBox).fBoxW - 15, tmDPCtrls(ilBox).fBoxH - 15, WHITE
                    pbcDaypart.CurrentX = tmDPCtrls(ilBox).fBoxX + fgBoxInsetX
                    pbcDaypart.CurrentY = tmDPCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    pbcDaypart.Print slStr
                    pbcDaypart.ForeColor = llColor
                End If
                pbcDaypart.ForeColor = llColor
            End If
        Next ilBox
    Next ilRow
End Sub

Private Sub pbcRateCard_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
End Sub

Private Sub pbcRateCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilLoop As Integer
    Dim ilWk As Integer
    Dim ilMn As Integer
    Dim ilDone As Integer
    Dim ilStartWkNo As Integer
    Dim ilMaxFltNo As Integer
    ReDim ilStartWk(0 To 12) As Integer 'Index zero ignored
    ReDim ilNoWks(0 To 12) As Integer   'Index zero ignored
    If Button = 2 Then  'Right Mouse
        If imView = 2 Then
            Exit Sub
        End If
        ilCompRow = vbcRateCard.LargeChange + 1
        If UBound(smRCSave, 2) > ilCompRow Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(smRCSave, 2)
        End If
        For ilRow = 1 To ilMaxRow Step 1
            If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(VEHINDEX).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(VEHINDEX).fBoxY + tmRCCtrls(VEHINDEX).fBoxH)) Then
                imButtonIndex = ilRow + vbcRateCard.Value - 1
                imIgnoreRightMove = True
                mShowRCInfo
                imIgnoreRightMove = False
                Exit Sub
            End If
        Next ilRow
        Exit Sub
    End If
    'Check if hot spot
    If imInHotSpot Then
        Exit Sub
    End If
    If UBound(tmRifRec) > 1 Then
        For ilLoop = LBONE To UBound(imHotSpot, 1) Step 1
            If (X >= imHotSpot(ilLoop, 1)) And (X <= imHotSpot(ilLoop, 3)) And (Y >= imHotSpot(ilLoop, 2)) And (Y <= imHotSpot(ilLoop, 4)) Then
                Screen.MousePointer = vbHourglass
                mSPSetShow imSPBoxNo    'Remove focus
                imSPBoxNo = -1
                mRCSetShow imRCBoxNo
                imRCBoxNo = -1
                imInHotSpot = True
                Select Case ilLoop
                    Case 1  'Goto Start
                        imPdYear = imRifStartYear
                        imPdStartWk = 1
                    Case 2  'Reduce by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(1).iYear = imRifStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo < 3) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 9 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                If (tmPdGroups(1).iStartWkNo >= 12) And (tmPdGroups(1).iStartWkNo <= 14) Then
                                    imPdStartWk = ilStartWk(1)
                                ElseIf (tmPdGroups(1).iStartWkNo >= 25) And (tmPdGroups(1).iStartWkNo <= 27) Then
                                    imPdStartWk = ilStartWk(1)
                                    For ilWk = 1 To 3 Step 1
                                        imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                    Next ilWk
                                Else
                                    imPdStartWk = ilStartWk(1)
                                    For ilWk = 1 To 6 Step 1
                                        imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                    Next ilWk
                                End If
                            End If
                        ElseIf rbcType(1).Value Then    'Month
                            If (tmPdGroups(1).iYear = imRifStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo < 3) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 11 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                'For ilWk = 2 To 12 Step 1
                                '    If tmPdGroups(1).iStartWkNo = ilStartWk(ilWk) Then
                                '        Exit For
                                '    End If
                                '    imPdStartWk = imPdStartWk + ilNoWks(ilWk - 1)
                                'Next ilWk
                                ilStartWkNo = tmPdGroups(1).iStartWkNo
                                ilDone = False
                                For ilMn = 0 To 2 Step 1
                                    For ilWk = 2 To 12 Step 1
                                        If ilStartWkNo = ilStartWk(ilWk) Then
                                            If imPdStartWk <= 1 Then
                                                imPdStartWk = 1
                                                ilDone = True
                                            Else
                                                ilStartWkNo = imPdStartWk
                                            End If
                                            Exit For
                                        End If
                                        imPdStartWk = imPdStartWk + ilNoWks(ilWk - 1)
                                    Next ilWk
                                    If ilDone Or ilMn = 2 Then
                                        Exit For
                                    End If
                                    imPdStartWk = ilStartWk(1)
                                Next ilMn
                            End If
                        ElseIf rbcType(2).Value Then    'Week
                            If (tmPdGroups(1).iYear = imRifStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(1).iStartWkNo < 2) Then 'At end
                                imPdYear = tmPdGroups(1).iYear - 1
                                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                                imPdStartWk = ilStartWk(1)
                                For ilWk = 1 To 12 Step 1
                                    imPdStartWk = imPdStartWk + ilNoWks(ilWk)
                                Next ilWk
                                imPdStartWk = imPdStartWk - 3   '1
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                imPdStartWk = tmPdGroups(1).iStartWkNo - 3  '1
                                If imPdStartWk < 1 Then
                                    imPdStartWk = 1
                                End If
                            End If
                        Else                            'Flight
                            If (tmPdGroups(1).iYear = imRifStartYear) And (tmPdGroups(1).iFltNo < 2) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If tmPdGroups(1).iFltNo < 2 Then
                                imPdYear = tmPdGroups(1).iYear - 1
                                imPdStartFltNo = 1
                                For ilWk = LBound(tgRcfI.iFltNo) + 1 To UBound(tgRcfI.iFltNo) Step 1
                                    If tgRcfI.iFltNo(ilWk) > imPdStartFltNo Then
                                        imPdStartFltNo = tgRcfI.iFltNo(ilWk)
                                    End If
                                Next ilWk
                            Else
                                imPdYear = tmPdGroups(1).iYear
                                imPdStartFltNo = tmPdGroups(1).iFltNo - 1
                            End If
                        End If
                    Case 3  'Increase by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(4).iYear = imRifStartYear + imRifNoYears - 1) And (tmPdGroups(4).iStartWkNo > 39) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths tmPdGroups(4).iYear, ilStartWk(), ilNoWks()
                            If (tmPdGroups(4).iYear = imRifStartYear + imRifNoYears - 1) And (tmPdGroups(4).iStartWkNo >= ilStartWk(12)) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(2).iYear = imRifStartYear + imRifNoYears - 1) And (tmPdGroups(2).iStartWkNo >= ilStartWk(9)) Then 'At end
                                imPdYear = tmPdGroups(2).iYear
                                imPdStartWk = tmPdGroups(2).iStartWkNo
                            ElseIf (tmPdGroups(3).iYear = imRifStartYear + imRifNoYears - 1) And (tmPdGroups(3).iStartWkNo >= ilStartWk(9)) Then 'At end
                                imPdYear = tmPdGroups(3).iYear
                                imPdStartWk = tmPdGroups(3).iStartWkNo
                            Else
                                imPdYear = tmPdGroups(4).iYear
                                imPdStartWk = tmPdGroups(4).iStartWkNo
                            End If
                            'imPdYear = tmPdGroups(2).iYear
                            'imPdStartWk = tmPdGroups(2).iStartWkNo
                        ElseIf rbcType(2).Value Then    'Week
                            mCompMonths tmPdGroups(4).iYear, ilStartWk(), ilNoWks()
                            If (tmPdGroups(4).iYear = imRifStartYear + imRifNoYears - 1) And (tmPdGroups(4).iStartWkNo >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(2).iYear = imRifStartYear + imRifNoYears - 1) And (tmPdGroups(2).iStartWkNo + 3 >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imPdYear = tmPdGroups(2).iYear
                                imPdStartWk = tmPdGroups(2).iStartWkNo
                            ElseIf (tmPdGroups(3).iYear = imRifStartYear + imRifNoYears - 1) And (tmPdGroups(3).iStartWkNo + 3 >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imPdYear = tmPdGroups(3).iYear
                                imPdStartWk = tmPdGroups(3).iStartWkNo
                            Else
                                imPdYear = tmPdGroups(4).iYear
                                imPdStartWk = tmPdGroups(4).iStartWkNo
                            End If
                            'imPdYear = tmPdGroups(2).iYear
                            'imPdStartWk = tmPdGroups(2).iStartWkNo
                        Else 'Flight
                            ilMaxFltNo = 1
                            For ilWk = LBound(tgRcfI.iFltNo) + 1 To UBound(tgRcfI.iFltNo) Step 1
                                If tgRcfI.iFltNo(ilWk) > ilMaxFltNo Then
                                    ilMaxFltNo = tgRcfI.iFltNo(ilWk)
                                End If
                            Next ilWk
                            If (tmPdGroups(4).iYear = imRifStartYear + imRifNoYears - 1) And (tmPdGroups(4).iFltNo >= ilMaxFltNo - 3) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If tmPdGroups(2).iYear <> imRifStartYear Then
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartFltNo = tmPdGroups(2).iFltNo
                        End If
                    Case 4  'GoTo End
                        imPdYear = imRifStartYear + imRifNoYears - 1
                        If rbcType(0).Value Then    'Quarter
                            imPdStartWk = 1
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(9)  'At end
                        ElseIf rbcType(2).Value Then    'Week
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(12) + ilNoWks(12) - 4
                        Else                            'Flight
                            ilMaxFltNo = 1
                            For ilWk = LBound(tgRcfI.iFltNo) + 1 To UBound(tgRcfI.iFltNo) Step 1
                                If tgRcfI.iFltNo(ilWk) > ilMaxFltNo Then
                                    ilMaxFltNo = tgRcfI.iFltNo(ilWk)
                                End If
                            Next ilWk
                            If ilMaxFltNo > 4 Then
                                imPdStartFltNo = ilMaxFltNo - 3
                            Else
                                imPdStartFltNo = 1
                            End If
                        End If
                End Select
                pbcRateCard.Cls
                mGetShowDates
                pbcRateCard_Paint
                Screen.MousePointer = vbDefault
                imInHotSpot = False
                Exit Sub
            End If
        Next ilLoop
    End If
    If imView = 2 Then
        Exit Sub
    End If
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub

Private Sub pbcRateCard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If imView = 2 Then
            Exit Sub
        End If
        ilCompRow = vbcRateCard.LargeChange + 1
        If UBound(smRCSave, 2) > ilCompRow Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(smRCSave, 2)
        End If
        For ilRow = 1 To ilMaxRow Step 1
            If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(VEHINDEX).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(VEHINDEX).fBoxY + tmRCCtrls(VEHINDEX).fBoxH)) Then
                If (imButtonIndex = ilRow + vbcRateCard.Value - 1) And (plcRCInfo.Visible) Then
                    Exit Sub
                End If
                imButtonIndex = ilRow + vbcRateCard.Value - 1
                imIgnoreRightMove = True
                mShowRCInfo
                imIgnoreRightMove = False
                Exit Sub
            End If
        Next ilRow
        plcRCInfo.Visible = False
        Exit Sub
    End If
End Sub

Private Sub pbcRateCard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim llRowNo As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilChg As Integer
    Dim ilFlt As Integer
    If imView = 2 Then
        Exit Sub
    End If
    'Eliminate Daypart changes (12/11/03) as input is still for imView = 0 (Rate)
    If Button = 2 Then
        imButtonIndex = -1
        plcRCInfo.Visible = False
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    For ilBox = imLBNWCtrls To UBound(tmNWCtrls) - 1 Step 1
        If (X >= tmNWCtrls(ilBox).fBoxX) And (X <= (tmNWCtrls(ilBox).fBoxX + tmNWCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmNWCtrls(ilBox).fBoxY)) And (Y <= (tmNWCtrls(ilBox).fBoxY + tmNWCtrls(ilBox).fBoxH)) Then
                If imInHotSpot Then
                    Exit Sub
                End If
                mSPSetShow imSPBoxNo    'Remove focus
                imSPBoxNo = -1
                mRCSetShow imRCBoxNo
                imRCBoxNo = -1
                sgWkStartDate = tmPdGroups(ilBox).sStartDate
                sgWkEndDate = tmPdGroups(ilBox).sEndDate
                If rbcType(3).Value Then     'Check flight weeks
                    igCurNoWks = tmPdGroups(ilBox).iTrueNoWks
                Else
                    igCurNoWks = -1
                End If
                RCSplit.Show vbModal
                ilChg = False
                If igRCReturn = 1 Then  'Done
                    If rbcType(3).Value Then     'Check flight weeks
                        If igCurNoWks > igNewNoWks Then
                            'Split flight
                            ilChg = True
                            For ilFlt = LBound(tgRcfI.iFltNo) + 1 To UBound(tgRcfI.iFltNo) Step 1
                                If tmPdGroups(ilBox).iFltNo = tgRcfI.iFltNo(ilFlt) Then
                                    For ilLoop = ilFlt + igNewNoWks To UBound(tgRcfI.iFltNo) Step 1
                                        tgRcfI.iFltNo(ilLoop) = tgRcfI.iFltNo(ilLoop) + 1
                                    Next ilLoop
                                    igRcfChg = True
                                    Exit For
                                End If
                            Next ilFlt
                        ElseIf igCurNoWks < igNewNoWks Then
                            'Merge flights
                            ilChg = True
                            For ilFlt = LBound(tgRcfI.iFltNo) + 1 To UBound(tgRcfI.iFltNo) Step 1
                                If tmPdGroups(ilBox).iFltNo = tgRcfI.iFltNo(ilFlt) Then
                                    For ilLoop = ilFlt + 1 To ilFlt + igNewNoWks - 1 Step 1
                                        tgRcfI.iFltNo(ilLoop) = tgRcfI.iFltNo(ilFlt)
                                        If ilLoop = UBound(tgRcfI.iFltNo) Then
                                            Exit For
                                        End If
                                    Next ilLoop
                                    igRcfChg = True
                                    Exit For
                                End If
                            Next ilFlt
                        End If
                    End If
                    If ilChg Then
                        pbcRateCard.Cls
                        mGetShowDates
                        pbcRateCard_Paint
                    End If
                End If
                Exit Sub
            End If
        End If
    Next ilBox
    ilCompRow = vbcRateCard.LargeChange + 1
    If UBound(tmRifRec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tmRifRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBRCCtrls To UBound(tmRCCtrls) Step 1
            If (X >= tmRCCtrls(ilBox).fBoxX) And (X <= (tmRCCtrls(ilBox).fBoxX + tmRCCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(ilBox).fBoxY + tmRCCtrls(ilBox).fBoxH)) Then
                    If ilBox > AVGINDEX Then    'DAYPARTINDEX Then
                        Beep
                        Exit Sub
                    End If
                    llRowNo = ilRow + vbcRateCard.Value - 1
                    If llRowNo > UBound(tmRifRec) Then
                        Beep
                        If imSPBoxNo > 0 Then
                            mSPSetFocus
                        Else
                            mSetFocus imRCBoxNo
                        End If
                        Exit Sub
                    End If
                    'TTP 10340 - 11/4/21 - JW - Rate Card screen: Acquisition cost can't be entered or edited
                    If (ilBox = ACQUISITIONINDEX) And ((Trim$(smRCSave(ACQUISITIONINDEX, llRowNo)) = "Y") Or (tgUrf(0).sChgAcq <> "I")) Then     'DAYPARTINDEX Then
                        Beep
                        Exit Sub
                    End If
                    'If (smDPShow(DPBASEINDEX, llRowNo) = "Y") And (ilBox = PCTINVINDEX) Then
                    '    Beep
                    '    If imSPBoxNo > 0 Then
                    '        mSPSetFocus
                    '    Else
                    '        mSetFocus imRCBoxNo
                    '    End If
                    '    Exit Sub
                    'End If
                    If (lmRCRowNo = UBound(tmRifRec)) And (imRCBoxNo = VEHINDEX) And (llRowNo <> lmRCRowNo) Then
                        For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
                            slStr = ""
                            gSetShow pbcRateCard, slStr, tmRCCtrls(ilLoop)
                            smRCShow(ilLoop, lmRCRowNo) = tmRCCtrls(ilLoop).sShow
                        Next ilLoop
                        pbcRateCard_Paint
                        mSetDefInSave   'Set defaults for extra row
                    End If
                    If (Trim$(smRCSave(VEHINDEX, llRowNo)) = "") And (ilBox <> VEHINDEX) Then
                        Beep
                        ilBox = VEHINDEX
                    End If
                    If (Trim$(edcDropDown.Text) = "") And (imRCBoxNo = VEHINDEX) Then
                        Beep
                        ilBox = VEHINDEX
                    End If
                    mSPSetShow imSPBoxNo
                    imSPBoxNo = -1
                    mRCSetShow imRCBoxNo
                    lmRCRowNo = ilRow + vbcRateCard.Value - 1
                    imRCBoxNo = ilBox
                    mRCEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    If imSPBoxNo > 0 Then
        mSPSetFocus
    Else
        mSetFocus imRCBoxNo
    End If
End Sub

Private Sub pbcRateCard_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    
    mPaintRCTitle
    llColor = pbcRateCard.ForeColor
    slFontName = pbcRateCard.FontName
    flFontSize = pbcRateCard.FontSize
    pbcRateCard.ForeColor = BLUE
    pbcRateCard.FontBold = False
    pbcRateCard.FontSize = 7
    pbcRateCard.FontName = "Arial"
    pbcRateCard.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilBox = imLBWKCtrls To UBound(tmWKCtrls) Step 1
        gPaintArea pbcRateCard, tmWKCtrls(ilBox).fBoxX, tmWKCtrls(ilBox).fBoxY, tmWKCtrls(ilBox).fBoxW - 15, tmWKCtrls(ilBox).fBoxH - 15, WHITE
        pbcRateCard.CurrentX = tmWKCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcRateCard.CurrentY = tmWKCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcRateCard.Print tmWKCtrls(ilBox).sShow
    Next ilBox
    For ilBox = imLBNWCtrls To UBound(tmNWCtrls) Step 1
        'gPaintArea pbcRateCard, tmNWCtrls(ilBox).fBoxX, tmNWCtrls(ilBox).fBoxY, tmNWCtrls(ilBox).fBoxW - 15, tmNWCtrls(ilBox).fBoxH - 15, LIGHTBLUE
        pbcRateCard.CurrentX = tmNWCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcRateCard.CurrentY = tmNWCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcRateCard.Print tmNWCtrls(ilBox).sShow
    Next ilBox
    pbcRateCard.FontSize = flFontSize
    pbcRateCard.FontName = slFontName
    pbcRateCard.FontSize = flFontSize
    pbcRateCard.ForeColor = llColor
    pbcRateCard.FontBold = True
    ilStartRow = vbcRateCard.Value '+ 1  'Top location
    ilEndRow = vbcRateCard.Value + vbcRateCard.LargeChange ' + 1
    llColor = pbcRateCard.ForeColor
    
    If imView <> 2 Then
        If ilEndRow > UBound(tmRifRec) Then
            ilEndRow = UBound(tmRifRec) 'include blank row as it might have data
        End If
        For ilRow = ilStartRow To ilEndRow Step 1
            If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER And imRCSave(17, ilRow) = 1 Then
                    GoTo Loop_Next_Row
            End If
        
            For ilBox = imLBRCCtrls To UBound(tmRCCtrls) Step 1
                If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER And ilBox = CPMINDEX Then
                    GoTo Loop_Next_Box
                End If
                '11/4/21 - JW - Ok'd with Jason
                If ilBox = RPTINDEX Or ilBox = BASEINDEX Then
                    If Trim(smRCShow(ilBox, ilRow)) = "Y" Then
                        smRCShow(ilBox, ilRow) = "Yes"
                    End If
                    If Trim(smRCShow(ilBox, ilRow)) = "N" Then
                        smRCShow(ilBox, ilRow) = "No"
                    End If
                End If
                If ilBox <> ACQUISITIONINDEX Then
                    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
                        'L.Bianchi
                        ' LB 02/10/21
                         'If (ilBox > CPMINDEX And imRCSave(17, ilRow) = 1 And ilBox <> SORTINDEX) Or (ilBox = CPMINDEX And imRCSave(17, ilRow) = 0) Or (ilBox = CPMINDEX And imRCSave(16, ilRow) > 0) Then
                         '   If Trim$(smRCShow(VEHINDEX, ilRow)) <> "" Then
                         '       gPaintArea pbcRateCard, tmRCCtrls(ilBox).fBoxX, tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmRCCtrls(ilBox).fBoxW - 15, tmRCCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                         '   End If
                         'Else
                         If (ilBox = CPMINDEX And imRCSave(17, ilRow) = 0) Or (ilBox = CPMINDEX And imRCSave(16, ilRow) > 0) Then
                            If Trim$(smRCShow(VEHINDEX, ilRow)) <> "" Then
                                gPaintArea pbcRateCard, tmRCCtrls(ilBox).fBoxX, tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmRCCtrls(ilBox).fBoxW - 15, tmRCCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                            End If
                         Else
                            If ilBox <> AVGINDEX Then
                                gPaintArea pbcRateCard, tmRCCtrls(ilBox).fBoxX, tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmRCCtrls(ilBox).fBoxW - 15, tmRCCtrls(ilBox).fBoxH - 15, WHITE
                            Else
                                If Trim$(smRCShow(VEHINDEX, ilRow)) <> "" Then
                                    gPaintArea pbcRateCard, tmRCCtrls(ilBox).fBoxX, tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmRCCtrls(ilBox).fBoxW - 15, tmRCCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                                End If
                            End If
                         End If
                    Else
                        If ilBox <> AVGINDEX Then
                            gPaintArea pbcRateCard, tmRCCtrls(ilBox).fBoxX, tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmRCCtrls(ilBox).fBoxW - 15, tmRCCtrls(ilBox).fBoxH - 15, WHITE
                        Else
                            gPaintArea pbcRateCard, tmRCCtrls(ilBox).fBoxX, tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmRCCtrls(ilBox).fBoxW - 15, tmRCCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                        End If
                    End If
                Else
                    If (Trim$(smRCSave(ACQUISITIONINDEX, ilRow)) <> "Y") And ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
                        gPaintArea pbcRateCard, tmRCCtrls(ilBox).fBoxX, tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmRCCtrls(ilBox).fBoxW - 15, tmRCCtrls(ilBox).fBoxH - 15, WHITE
                    Else
                        'L.Bianchi
                        If Trim$(smRCShow(VEHINDEX, ilRow)) <> "" Then
                                gPaintArea pbcRateCard, tmRCCtrls(ilBox).fBoxX, tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmRCCtrls(ilBox).fBoxW - 15, tmRCCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                        End If
                    End If
                End If
                
                If (ilBox = DAYPARTINDEX) Then
                    If Trim$(smRCSave(BASEINDEX, ilRow)) = "Y" Then
                        pbcRateCard.ForeColor = MAGENTA
                    ElseIf (Trim$(smRCSave(BASEINDEX, ilRow)) <> "N") And (Trim$(smDPShow(DPBASEINDEX, ilRow)) = "Y") Then
                        pbcRateCard.ForeColor = MAGENTA
                    End If
                End If
                
                If (ilBox = VEHINDEX) And (imRCSave(11, ilRow) > 0 Or imRCSave(16, ilRow) > 0) Then
                    pbcRateCard.ForeColor = BLUE
                End If
                If (ilBox = VEHINDEX) And (imRCSave(16, ilRow) > 0 Or imRCSave(17, ilRow) = 1) Then
                    pbcRateCard.FontItalic = True
                Else
                    pbcRateCard.FontItalic = False
                End If
                If (ilBox = VEHINDEX) And (imRCSave(9, ilRow) > 0) Then
                    pbcRateCard.ForeColor = Red
                End If
                If (ilBox = DAYPARTINDEX) And (imRCSave(10, ilRow) > 0) Then
                    pbcRateCard.ForeColor = Red
                End If
                If (ilBox = DOLLAR1INDEX) And (imRCSave(12, ilRow) > 0) Then
                    pbcRateCard.ForeColor = Red
                End If
                If (ilBox = DOLLAR2INDEX) And (imRCSave(13, ilRow) > 0) Then
                    pbcRateCard.ForeColor = Red
                End If
                If (ilBox = DOLLAR3INDEX) And (imRCSave(14, ilRow) > 0) Then
                    pbcRateCard.ForeColor = Red
                End If
                If (ilBox = DOLLAR4INDEX) And (imRCSave(15, ilRow) > 0) Then
                    pbcRateCard.ForeColor = Red
                End If
                If ilRow = UBound(tmRifRec) Then
                    pbcRateCard.ForeColor = DARKPURPLE
                End If
                
                pbcRateCard.CurrentX = tmRCCtrls(ilBox).fBoxX + fgBoxInsetX
                pbcRateCard.CurrentY = tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                slStr = Trim$(smRCShow(ilBox, ilRow))
                pbcRateCard.Print slStr
                pbcRateCard.ForeColor = llColor
Loop_Next_Box:
            Next ilBox
Loop_Next_Row:
        Next ilRow
    Else
        If ilEndRow > UBound(smBdShow, 2) Then
            ilEndRow = UBound(smBdShow, 2) 'include blank row as it might have data
        End If
        For ilRow = ilStartRow To ilEndRow Step 1
            For ilBox = imLBRCCtrls To UBound(tmRCCtrls) Step 1
                If (ilBox = DAYPARTINDEX) Then
                    pbcRateCard.ForeColor = MAGENTA
                End If
                If (ilBox = imLBRCCtrls) And (Trim$(smBdShow(DAYPARTINDEX, ilRow)) = "") Then
                    pbcRateCard.ForeColor = BLUE
                    pbcRateCard.FontBold = False
                    pbcRateCard.FontSize = 7
                    pbcRateCard.FontName = "Arial"
                    pbcRateCard.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
                    pbcRateCard.CurrentY = tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) '- 30'+ fgBoxInsetY
                Else
                    pbcRateCard.CurrentY = tmRCCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                End If
                pbcRateCard.CurrentX = tmRCCtrls(ilBox).fBoxX + fgBoxInsetX
                slStr = Trim$(smBdShow(ilBox, ilRow))
                If (ilBox > DAYPARTINDEX) Then
                    If Left$(slStr, 1) = "R" Then
                        pbcRateCard.ForeColor = Red
                        slStr = right$(slStr, Len(slStr) - 1)
                    ElseIf Left$(slStr, 1) = "G" Then
                        pbcRateCard.ForeColor = DARKGREEN
                        slStr = right$(slStr, Len(slStr) - 1)
                    Else
                        pbcRateCard.ForeColor = llColor
                    End If
                End If
                If ilRow = UBound(smBdShow, 2) Then
                    pbcRateCard.ForeColor = DARKPURPLE
                End If
                pbcRateCard.Print slStr
                pbcRateCard.ForeColor = llColor
                If (ilBox = imLBRCCtrls) And (Trim$(smBdShow(DAYPARTINDEX, ilRow)) = "") Then
                    pbcRateCard.FontSize = flFontSize
                    pbcRateCard.FontName = slFontName
                    pbcRateCard.FontSize = flFontSize
                    pbcRateCard.ForeColor = llColor
                    pbcRateCard.FontBold = True
                End If
            Next ilBox
        Next ilRow
    End If
End Sub

Private Sub pbcSP_GotFocus()
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
End Sub

Private Sub pbcSP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBSPCtrls To UBound(tmSPCtrls) Step 1
        If (X >= tmSPCtrls(ilBox).fBoxX) And (X <= tmSPCtrls(ilBox).fBoxX + tmSPCtrls(ilBox).fBoxW) Then
            If (Y >= tmSPCtrls(ilBox).fBoxY) And (Y <= tmSPCtrls(ilBox).fBoxY + tmSPCtrls(ilBox).fBoxH) Then
                'If ilBox = BUDGETINDEX Then
                '    If lbcBudget.ListCount <= 0 Then
                '        Beep
                '        Exit Sub
                '    End If
                'End If
                If ilBox = GRIDINDEX Then
                    If lbcGrid.ListCount <= 0 Then
                        Beep
                        Exit Sub
                    End If
                End If
                If ilBox = LENGTHINDEX Then
                    If lbcLen.ListCount <= 0 Then
                        Beep
                        Exit Sub
                    End If
                End If
                If ilBox = CURGRIDINDEX Then
                    If lbcGrid.ListCount <= 0 Then
                        Beep
                        Exit Sub
                    End If
                End If
                mSPSetShow imSPBoxNo
                imSPBoxNo = ilBox
                mSPEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSPSetFocus
End Sub

Private Sub pbcSP_Paint()
    Dim ilBox As Integer
    For ilBox = imLBSPCtrls To UBound(tmSPCtrls) Step 1
        pbcSP.CurrentX = tmSPCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSP.CurrentY = tmSPCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcSP.Print tmSPCtrls(ilBox).sShow
    Next ilBox
End Sub

Private Sub pbcSPSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSPSTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    Select Case imSPBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If (igRCMode = 0) And (imFirstTimeSelect) Then
                Exit Sub
            End If
            'If lbcBudget.ListCount <= 0 Then
                If lbcGrid.ListCount <= 0 Then
                    ilBox = LENGTHINDEX
                Else
                    ilBox = GRIDINDEX
                End If
            'Else
            '    ilBox = BUDGETINDEX
            'End If
            imSPBoxNo = ilBox
            mSPEnableBox ilBox
            Exit Sub
        'Case BUDGETINDEX
        '    mSPSetShow imSPBoxNo
        '    imSPBoxNo = -1
        '    cmcDone.SetFocus
        '    Exit Sub
        Case LENGTHINDEX
            If lbcGrid.ListCount <= 0 Then
                'If lbcBudget.ListCount <= 0 Then
                    mSPSetShow imSPBoxNo
                    imSPBoxNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                'Else
                '    ilBox = BUDGETINDEX
                'End If
            Else
                ilBox = GRIDINDEX
            End If
        Case GRIDINDEX
            'If lbcBudget.ListCount <= 0 Then
                mSPSetShow imSPBoxNo
                imSPBoxNo = -1
                cmcDone.SetFocus
                Exit Sub
            'Else
            '    ilBox = BUDGETINDEX
            'End If
        Case CURGRIDINDEX
            If lbcLen.ListCount <= 0 Then
                ilBox = GRIDINDEX
            Else
                ilBox = LENGTHINDEX
            End If
        Case Else
            ilBox = imSPBoxNo - 1
    End Select
    mSPSetShow imSPBoxNo
    imSPBoxNo = ilBox
    mSPEnableBox ilBox
End Sub

Private Sub pbcSPTab_GotFocus()
    Dim ilBox As Integer

    If GetFocus() <> pbcSPTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    If imDirProcess >= 0 Then
        mDirection imDirProcess
        imDirProcess = -1
        Exit Sub
    End If
    Select Case imSPBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            If lbcGrid.ListCount <= 0 Then
                ilBox = LENGTHINDEX
            Else
                ilBox = CURGRIDINDEX
            End If
        Case CURGRIDINDEX
            mSPSetShow imSPBoxNo
            imSPBoxNo = -1
            pbcSTab.SetFocus
            Exit Sub
        'Case BUDGETINDEX
        '    If lbcGrid.ListCount <= 0 Then
        '        If lbcLen.ListCount <= 0 Then
        '            mSPSetShow imSPBoxNo
        '            imSPBoxNo = -1
        '            pbcSTab.SetFocus
        '            Exit Sub
        '        Else
        '            ilBox = LENGTHINDEX
        '        End If
        '    Else
        '        ilBox = GRIDINDEX
        '    End If
        Case GRIDINDEX
            If lbcLen.ListCount <= 0 Then
                mSPSetShow imSPBoxNo
                imSPBoxNo = -1
                pbcSTab.SetFocus
                Exit Sub
            Else
                ilBox = CURGRIDINDEX
            End If
        Case LENGTHINDEX
            If lbcGrid.ListCount <= 0 Then
                mSPSetShow imSPBoxNo
                imSPBoxNo = -1
                pbcSTab.SetFocus
                Exit Sub
            Else
                ilBox = CURGRIDINDEX
            End If
        Case Else
            ilBox = imSPBoxNo + 1
    End Select
    mSPSetShow imSPBoxNo
    imSPBoxNo = ilBox
    mSPEnableBox ilBox
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If imRetBranch = True Then 'second gotfocus-ignore
        'imRetBranch = False
        Exit Sub
    End If
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If imView = 2 Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Eliminate Daypart changes (12/11/03) as input is still for imView = 0 (Rate)
    If imView = 1 Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    If imRCBoxNo = DAYPARTINDEX Then
        If mDPBranch() Then
            Exit Sub
        End If
    End If
    Select Case imRCBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If (igRCMode = 0) And (imFirstTimeSelect) Then
                Exit Sub
            End If
'                ilRet = mStartNew()
'                If Not ilRet Then
'                    Unload RateCard
'                    Exit Sub
'                End If
'            End If
'            imFirstTimeSelect = False
            imSettingValue = True
            vbcRateCard.Value = 1
            imSettingValue = False
            lmRCRowNo = 1
            ilBox = 1
            imRCBoxNo = ilBox
            mRCEnableBox ilBox
            Exit Sub
        Case VEHINDEX 'Name (first control within header)
            mRCSetShow imRCBoxNo
            If lmRCRowNo <= 1 Then
                If (pbcSPTab.Enabled) And (pbcSPTab.Visible) Then
                    pbcSPTab.SetFocus
                    Exit Sub
                End If
                If cbcSelect.Enabled Then
                    imRCBoxNo = -1
                    cbcSelect.SetFocus
                    Exit Sub
                End If
                ilBox = 1
            Else
                If imView = 1 Then
                    ilBox = DAYPARTINDEX
                Else
                    'If smDPShow(DPBASEINDEX, lmRCRowNo - 1) = "Y" Then
                    '    ilBox = DAYPARTINDEX
                    'Else
                    '    ilBox = PCTINVINDEX
                    'End If
                    ilBox = SORTINDEX
                End If
                
                'TTP 10388 - Rate card screen: if "ad server" site option setting is not enabled, a "podcast" medium type vehicle does not appear on the rate card screen
                '02/23/22 - JW Fix issue on the Rate Card screen per Jason Email: Wed 2/16/22 10:46 AM
                mSetVehicleMediumType lmRCRowNo
                mSetPvfType lmRCRowNo
                If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
                    mResetColumn (lmRCRowNo)
                End If
                
                lmRCRowNo = lmRCRowNo - 1
                If lmRCRowNo < vbcRateCard.Value Then
                    imSettingValue = True
                    vbcRateCard.Value = vbcRateCard.Value - 1
                    imSettingValue = False
                End If
                imRCBoxNo = ilBox
                mRCEnableBox ilBox
                Exit Sub
            End If
        Case BASEINDEX
            ilBox = imRCBoxNo - 1
            If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
                    If imRCSave(16, lmRCRowNo) = 1 Then
                        ilBox = DAYPARTINDEX
                    ElseIf imRCSave(17, lmRCRowNo) = 1 Then
                        ilBox = CPMINDEX
                    Else
                        If (Trim$(smRCSave(ACQUISITIONINDEX, lmRCRowNo)) = "Y") Or (tgUrf(0).sChgAcq <> "I") Then
                            ilBox = ilBox - 1
                        End If
                        ilBox = ilBox - 2 'for cpm index
                    End If
            Else
                If (Trim$(smRCSave(ACQUISITIONINDEX, lmRCRowNo)) = "Y") Or (tgUrf(0).sChgAcq <> "I") Then
                    ilBox = ilBox - 2
                End If
            End If
            
        'TTP 10340 - 11/4/21 - JW - Rate Card screen: Acquisition cost can't be entered or edited
        Case ACQUISITIONINDEX
            mRCSetShow imRCBoxNo
            ilBox = DAYPARTINDEX
            imRCBoxNo = ilBox
            mRCEnableBox ilBox
            Exit Sub
            
        Case SORTINDEX
            'L.Bianchi
            If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
                    If imRCSave(16, lmRCRowNo) = 1 Then
                        ilBox = DAYPARTINDEX
                    ' LB 02/10/21
                    'ElseIf imRCSave(17, lmRCRowNo) = 1 Then
                    '    ilBox = CPMINDEX
                    Else
                        ilBox = imRCBoxNo - 1
                    End If
            Else
                ilBox = imRCBoxNo - 1
            End If
        Case Else
            ilBox = imRCBoxNo - 1
    End Select
    mRCSetShow imRCBoxNo
    imRCBoxNo = ilBox
    mRCEnableBox ilBox
End Sub

Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcStartNew_GotFocus()
    Dim ilRet As Integer
    If imInNew Then
        Exit Sub
    End If
    If (igRCMode = 0) And (imFirstTimeSelect) Then
        imFirstTimeSelect = False
        ilRet = mStartNew()
        If Not ilRet Then
            imTerminate = True
            mTerminate
            Exit Sub
        End If
    End If
    mSetCommands
    If pbcSTab.Enabled Then
        If bmInStdPrice Then
            Exit Sub
        End If
        pbcSTab.SetFocus
    End If
End Sub

Private Sub pbcStatic_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilEnd As Integer
    If imRetBranch = True Then 'second gotfocus-ignore
        'imRetBranch = False
        Exit Sub
    End If
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    If imView = 2 Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Eliminate Daypart changes (12/11/03) as input is still for imView = 0 (Rate)
    If imView = 1 Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    If imRCBoxNo = DAYPARTINDEX Then
        If mDPBranch() Then
            Exit Sub
        End If
    End If
    If imDirProcess >= 0 Then
        mDirection imDirProcess
        imDirProcess = -1
        Exit Sub
    End If
    
    Select Case imRCBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            lmRCRowNo = UBound(tmRifRec)
            imSettingValue = True
            If lmRCRowNo <= vbcRateCard.LargeChange + 1 Then
                vbcRateCard.Value = 1
            Else
                vbcRateCard.Value = lmRCRowNo - vbcRateCard.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = DAYPARTINDEX
        Case 0
            ilBox = VEHINDEX
        Case VEHINDEX
            mRCSetShow imRCBoxNo
            If (lmRCRowNo >= UBound(tmRifRec)) And (Trim$(smRCSave(VEHINDEX, lmRCRowNo)) = "") Then
                imRCBoxNo = -1
'                For ilLoop = LBound(tmCtrls) To UBound(tmCtrls) Step 1
'                    slStr = ""
'                    gSetShow pbcRateCard, slStr, tmCtrls(ilLoop)
'                    smShow(ilLoop, imRowNo) = tmCtrls(ilLoop).sShow
'                Next ilLoop
'                pbcRateCard_Paint
                If cmcUpdate.Enabled Then
                    cmcUpdate.SetFocus
                Else
                    cmcDone.SetFocus
                End If
                Exit Sub
               
            End If
            ilBox = DAYPARTINDEX
            imRCBoxNo = ilBox
            mRCEnableBox ilBox
            'TTP 10388 - Rate card screen: if "ad server" site option setting is not enabled, a "podcast" medium type vehicle does not appear on the rate card screen
            '02/23/22 - JW Fix issue on the Rate Card screen per Jason Email: Wed 2/16/22 10:46 AM
            'L.Bianchi
            mSetVehicleMediumType lmRCRowNo
            mSetPvfType lmRCRowNo
            If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
                mResetColumn (lmRCRowNo)
            End If
            Exit Sub
        Case DAYPARTINDEX 'Daypart name index
            If imView = 1 Then
                mRCSetShow imRCBoxNo
                If mTestSaveFields() = NO Then
                    mRCEnableBox imRCBoxNo
                    Exit Sub
                End If
                If lmRCRowNo >= UBound(tmRifRec) Then
                    'ReDim Preserve tmRifRec(0 To lmRCRowNo + 1) As RIFREC
                    'ReDim Preserve smRCSave(0 To SORTINDEX, 0 To lmRCRowNo + 1) As String * 40
                    'ReDim Preserve lmRCSave(0 To TOTALINDEX - SORTINDEX, 0 To lmRCRowNo + 1) As Long
                    'ReDim Preserve imRCSave(0 To 15, 0 To lmRCRowNo + 1) As Integer
                    'ReDim Preserve smRCShow(0 To AVGINDEX, 0 To lmRCRowNo + 1) As String * 40
                    'ReDim Preserve smDPShow(0 To DPBASEINDEX, 0 To lmRCRowNo + 1) As String * 40 'Values shown in program area
                    'mInitRif lmRCRowNo + 1
                    'If UBound(tmRifRec) <= vbcRateCard.LargeChange Then 'was <=
                    '    vbcRateCard.Max = LBONE 'LBound(tmRifRec)
                    'Else
                    '    vbcRateCard.Max = UBound(tmRifRec) - vbcRateCard.LargeChange '-1
                    'End If
                    'mSetDefInSave   'Set defaults for extra row
                    'mInitLastShow  'Init last row
                    mAddNewRow lmRCRowNo + 1
                End If
                lmRCRowNo = lmRCRowNo + 1
                If lmRCRowNo > vbcRateCard.Value + vbcRateCard.LargeChange Then '+ 1 Then
                    imSettingValue = True
                    vbcRateCard.Value = vbcRateCard.Value + 1
                    imSettingValue = False
                End If
                If lmRCRowNo >= UBound(tmRifRec) Then
                    mSetCommands
                    imRCBoxNo = 0
                    lacRCFrame.Move 0, tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
                    lacRCFrame.Visible = True
                    lacDPFrame.Move 0, tmDPCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
                    lacDPFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcRateCard.Top + tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                    Exit Sub
                Else
                    ilBox = 1
                    imRCBoxNo = ilBox
                    mRCEnableBox ilBox
                    Exit Sub
                End If
            Else
                ilBox = imRCBoxNo + 1
                'L.Bianchi
                If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
                    If imRCSave(16, lmRCRowNo) = 1 Then 'CPM Package
                        ilBox = ilBox + 1
                        'TTP 10340 - 11/4/21 - JW - Rate Card screen: Acquisition cost can't be entered or edited
                        If (Trim$(smRCSave(ACQUISITIONINDEX, lmRCRowNo)) = "Y") Or (tgUrf(0).sChgAcq <> "I") Then
                            ilBox = ilBox + (BASEINDEX - ACQUISITIONINDEX)
                        End If
                    ElseIf imRCSave(17, lmRCRowNo) = 0 Then
                        ilBox = ilBox + 1
                        If (Trim$(smRCSave(ACQUISITIONINDEX, lmRCRowNo)) = "Y") Or (tgUrf(0).sChgAcq <> "I") Then
                            ilBox = ilBox + 1
                        End If
                    End If
                Else
                    ilBox = ilBox + 1
                    If (Trim$(smRCSave(ACQUISITIONINDEX, lmRCRowNo)) = "Y") Or (tgUrf(0).sChgAcq <> "I") Then
                        ilBox = ilBox + 1
                    End If
                End If
            End If
        Case CPMINDEX
            'L.Bianchi
            If (Trim$(smRCSave(ACQUISITIONINDEX, lmRCRowNo)) = "Y") Or (tgUrf(0).sChgAcq <> "I") Then
                ' LB 02/10/21
                'ilBox = imRCBoxNo + (SORTINDEX - CPMINDEX)
                ilBox = imRCBoxNo + (BASEINDEX - CPMINDEX)
            Else
                ilBox = imRCBoxNo + 2
            End If
        Case ACQUISITIONINDEX
            'L.Bianchi
            ilBox = imRCBoxNo + 1
            If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
                    ' LB 02/10/21
                    'If imRCSave(17, lmRCRowNo) = 1 Or imRCSave(16, lmRCRowNo) = 1 Then
                    ' If imRCSave(16, lmRCRowNo) = 1 Then
                    '    tmRifRec(lmRCRowNo).tRif.sBase = "N"
                    '    tmRifRec(lmRCRowNo).tRif.sRpt = "N"
                    '    smRCSave(BASEINDEX, lmRCRowNo) = tmRifRec(lmRCRowNo).tRif.sBase
                    '    gSetShow pbcRateCard, tmRifRec(lmRCRowNo).tRif.sBase, tmRCCtrls(BASEINDEX)
                    '    smRCShow(BASEINDEX, lmRCRowNo) = tmRCCtrls(BASEINDEX).sShow
                    
                    '    smRCSave(RPTINDEX, lmRCRowNo) = tmRifRec(lmRCRowNo).tRif.sRpt
                    '    gSetShow pbcRateCard, tmRifRec(lmRCRowNo).tRif.sRpt, tmRCCtrls(RPTINDEX)
                    '    smRCShow(RPTINDEX, lmRCRowNo) = tmRCCtrls(RPTINDEX).sShow
                    '    ilBox = ilBox + 2
                    'Else
                    '    ilBox = ilBox + 2
                    ' End If
            End If
        Case DOLLAR1INDEX To DOLLAR4INDEX   'DOLLARINDEX, PCTINVINDEX 'Last control within header
            ilEnd = False
                If imRCBoxNo - DOLLAR1INDEX + 2 >= 5 Then
                ilEnd = True
            Else
                If tmPdGroups(imRCBoxNo - DOLLAR1INDEX + 2).iStartWkNo < 0 Then
                    ilEnd = True
                End If
            End If
            If ilEnd Then
            'If (imRCBoxNo = PCTINVINDEX) Or ((imRCBoxNo = DOLLARINDEX) And (smDPShow(BASEINDEX, lmRCRowNo) = "Y")) Then
                mRCSetShow imRCBoxNo
                If mTestSaveFields() = NO Then
                    mRCEnableBox imRCBoxNo
                    Exit Sub
                End If
                If lmRCRowNo >= UBound(tmRifRec) Then
                    'ReDim Preserve tmRifRec(0 To lmRCRowNo + 1) As RIFREC
                    'ReDim Preserve smRCSave(0 To SORTINDEX, 0 To lmRCRowNo + 1) As String * 40
                    'ReDim Preserve lmRCSave(0 To TOTALINDEX - SORTINDEX, 0 To lmRCRowNo + 1) As Long
                    'ReDim Preserve imRCSave(0 To 15, 0 To lmRCRowNo + 1) As Integer
                    'ReDim Preserve smRCShow(0 To AVGINDEX, 0 To lmRCRowNo + 1) As String * 40
                    'ReDim Preserve smDPShow(0 To DPBASEINDEX, 0 To lmRCRowNo + 1) As String * 40 'Values shown in program area
                    'mInitRif lmRCRowNo + 1
                    'If UBound(tmRifRec) <= vbcRateCard.LargeChange Then 'was <=
                    '    vbcRateCard.Max = LBONE 'LBound(tmRifRec)
                    'Else
                    '    vbcRateCard.Max = UBound(tmRifRec) - vbcRateCard.LargeChange '-1
                    'End If
                    'mSetDefInSave   'Set defaults for extra row
                    'mInitLastShow  'Init last row
                    mAddNewRow lmRCRowNo + 1
                End If
                lmRCRowNo = lmRCRowNo + 1
                If lmRCRowNo > vbcRateCard.Value + vbcRateCard.LargeChange Then ' + 1 Then
                    imSettingValue = True
                    vbcRateCard.Value = vbcRateCard.Value + 1
                    imSettingValue = False
                End If
                If lmRCRowNo >= UBound(tmRifRec) Then
                    mSetCommands
                    imRCBoxNo = 0
                    lacRCFrame.Move 0, tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
                    lacRCFrame.Visible = True
                    lacDPFrame.Move 0, tmDPCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
                    lacDPFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcRateCard.Top + tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                    Exit Sub
                Else
                    ilBox = 1
                    imRCBoxNo = ilBox
                    mRCEnableBox ilBox
                    Exit Sub
                End If
            Else
                ilBox = imRCBoxNo + 1
            End If
        Case SORTINDEX
            ' LB 02/10/21
            'If imRCSave(17, lmRCRowNo) = 1 Or imRCSave(16, lmRCRowNo) = 1 Then
            If imRCSave(16, lmRCRowNo) = 1 Then
                ilEnd = True
                'If (imRCBoxNo = PCTINVINDEX) Or ((imRCBoxNo = DOLLARINDEX) And (smDPShow(BASEINDEX, lmRCRowNo) = "Y")) Then
                mRCSetShow imRCBoxNo
                If mTestSaveFields() = NO Then
                    mRCEnableBox imRCBoxNo
                    Exit Sub
                End If
                If lmRCRowNo >= UBound(tmRifRec) Then
                    'ReDim Preserve tmRifRec(0 To lmRCRowNo + 1) As RIFREC
                    'ReDim Preserve smRCSave(0 To SORTINDEX, 0 To lmRCRowNo + 1) As String * 40
                    'ReDim Preserve lmRCSave(0 To TOTALINDEX - SORTINDEX, 0 To lmRCRowNo + 1) As Long
                    'ReDim Preserve imRCSave(0 To 15, 0 To lmRCRowNo + 1) As Integer
                    'ReDim Preserve smRCShow(0 To AVGINDEX, 0 To lmRCRowNo + 1) As String * 40
                    'ReDim Preserve smDPShow(0 To DPBASEINDEX, 0 To lmRCRowNo + 1) As String * 40 'Values shown in program area
                    'mInitRif lmRCRowNo + 1
                    'If UBound(tmRifRec) <= vbcRateCard.LargeChange Then 'was <=
                    '    vbcRateCard.Max = LBONE 'LBound(tmRifRec)
                    'Else
                    '    vbcRateCard.Max = UBound(tmRifRec) - vbcRateCard.LargeChange '-1
                    'End If
                    'mSetDefInSave   'Set defaults for extra row
                    'mInitLastShow  'Init last row
                    mAddNewRow lmRCRowNo + 1
                End If
                lmRCRowNo = lmRCRowNo + 1
                If lmRCRowNo > vbcRateCard.Value + vbcRateCard.LargeChange Then ' + 1 Then
                    imSettingValue = True
                    vbcRateCard.Value = vbcRateCard.Value + 1
                    imSettingValue = False
                End If
                If lmRCRowNo >= UBound(tmRifRec) Then
                    mSetCommands
                    imRCBoxNo = 0
                    lacRCFrame.Move 0, tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
                    lacRCFrame.Visible = True
                    lacDPFrame.Move 0, tmDPCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
                    lacDPFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcRateCard.Top + tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                    Exit Sub
                Else
                    ilBox = 1
                    imRCBoxNo = ilBox
                    mRCEnableBox ilBox
                    Exit Sub
                End If
            Else
                ilBox = imRCBoxNo + 1
            End If
            
        Case 0
            ilBox = imRCBoxNo + 1
        Case Else
            ilBox = imRCBoxNo + 1
    End Select
    mRCSetShow imRCBoxNo
    imRCBoxNo = ilBox
    mRCEnableBox ilBox
End Sub

Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub pbcView_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
End Sub

Private Sub pbcView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    If imView = 2 Then
        edcBdDropDown.Visible = False
        cmcBdDropDown.Visible = False
        lbcBudget.Visible = False
        lbcBudget2.Visible = False
        imView = 0
        pbcDaypart.Visible = False
        pbcRateCard.Visible = True
        imSettingValue = True
        vbcRateCard.Min = LBONE 'LBound(tmRifRec)
        imSettingValue = True
        If UBound(tmRifRec) <= vbcRateCard.LargeChange Then ' + 1 Then
            vbcRateCard.Max = LBONE 'LBound(tmRifRec)
        Else
            vbcRateCard.Max = UBound(tmRifRec) - vbcRateCard.LargeChange
        End If
        imSettingValue = True
        If vbcRateCard.Value <> vbcRateCard.Min Then
            vbcRateCard.Value = vbcRateCard.Min
        Else
            pbcRateCard.Cls
            pbcRateCard_Paint
        End If
        imSettingValue = False
    ElseIf imView = 0 Then
        imView = 1
        pbcRateCard.Visible = False
        pbcDaypart.Visible = True
    ElseIf imView = 1 Then
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sGUsePropSys <> "Y") Then
            imView = 0
            pbcRateCard.Visible = True
            pbcDaypart.Visible = False
        Else
            imView = 2
            pbcRateCard.Visible = True
            pbcDaypart.Visible = False
            If tgSpf.sRUseCorpCal <> "Y" Then
                edcBdDropDown.Visible = True
                cmcBdDropDown.Visible = True
                lbcBudget.Move pbcRateCard.Left + edcBdDropDown.Left, pbcRateCard.Top + edcBdDropDown.Top + edcBdDropDown.Height
                lbcBudget.ListIndex = -1
                edcBdDropDown.Text = ""
                imBSelectedIndex = -1
            Else
                lbcBudget2.Move pbcRateCard.Left + edcBdDropDown.Left, pbcRateCard.Top + edcBdDropDown.Top ' + edcBdDropDown.Height
                lbcBudget2.Visible = True
            End If
            ReDim tgImpactRec(0 To 1) As IMPACTREC
            ReDim tgDollarRec(0 To 1, 0 To 1) As DOLLARREC
            ReDim smBdShow(0 To AVGINDEX, 0 To 1) As String * 40
            For ilLoop = LBound(smBdShow, 1) To UBound(smBdShow, 1) Step 1
                For ilIndex = LBound(smBdShow, 2) To UBound(smBdShow, 2) Step 1
                    smBdShow(ilLoop, ilIndex) = ""
                Next ilIndex
            Next ilLoop
            imSettingValue = True
            vbcRateCard.Min = 1
            imSettingValue = True
            vbcRateCard.Max = 1
            imSettingValue = True
            If vbcRateCard.Value <> 1 Then
                vbcRateCard.Value = 1
            Else
                vbcRateCard_Change
            End If
        End If
    End If
    pbcView_Paint
End Sub

Private Sub pbcView_Paint()
    pbcView.Cls
    pbcView.CurrentX = fgBoxInsetX
    pbcView.CurrentY = 0 'fgBoxInsetY
    If imView = 0 Then
        pbcView.Print "Rate"
    ElseIf imView = 1 Then
        pbcView.Print "Daypart"
    ElseIf imView = 2 Then
        pbcView.Print "Comparisons"
    End If
End Sub

Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    If imRCBoxNo = BASEINDEX Then
        ilIndex = BASEINDEX
    ElseIf imRCBoxNo = RPTINDEX Then
        ilIndex = RPTINDEX
    Else
        Exit Sub
    End If
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        imRifChg = True
        smRCSave(ilIndex, lmRCRowNo) = "Y"
        pbcYN_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        imRifChg = True
        smRCSave(ilIndex, lmRCRowNo) = "N"
        pbcYN_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If Trim$(smRCSave(ilIndex, lmRCRowNo)) = "Y" Then
            imRifChg = True
            smRCSave(ilIndex, lmRCRowNo) = "N"
            pbcYN_Paint
        ElseIf Trim$(smRCSave(ilIndex, lmRCRowNo)) = "N" Then
            imRifChg = True
            smRCSave(ilIndex, lmRCRowNo) = "Y"
            pbcYN_Paint
        End If
    End If
End Sub

Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    If imRCBoxNo = BASEINDEX Then
        ilIndex = BASEINDEX
    ElseIf imRCBoxNo = RPTINDEX Then
        ilIndex = RPTINDEX
    Else
        Exit Sub
    End If
    If Trim$(smRCSave(ilIndex, lmRCRowNo)) = "Y" Then
        imRifChg = True
        smRCSave(ilIndex, lmRCRowNo) = "N"
    Else
        imRifChg = True
        smRCSave(ilIndex, lmRCRowNo) = "Y"
    End If
    pbcYN_Paint
End Sub

Private Sub pbcYN_Paint()
    Dim ilIndex As Integer
    If imRCBoxNo = BASEINDEX Then
        ilIndex = BASEINDEX
    ElseIf imRCBoxNo = RPTINDEX Then
        ilIndex = RPTINDEX
    Else
        Exit Sub
    End If
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If Trim$(smRCSave(ilIndex, lmRCRowNo)) = "Y" Then
        pbcYN.Print "Yes"
    Else
        pbcYN.Print "No"
    End If
End Sub

Private Sub plcRateCard_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcRateCard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub plcRCInfo_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub plcSP_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub plcSP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub plcStatic_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub rbcShow_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcShow(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilFound As Integer
    ReDim ilStartWk(0 To 12) As Integer 'Index zero ignored
    ReDim ilNoWks(0 To 12) As Integer

    If imIgnoreSetting Then
        imIgnoreSetting = False
        Exit Sub
    End If
    If Value Then
        Screen.MousePointer = vbHourglass
        pbcRateCard.Cls
        If imRifStartYear <> 0 Then
            imPdYear = tmPdGroups(1).iYear
            If imTypeIndex = 1 Then 'By month
                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                imPdStartWk = tmPdGroups(1).iStartWkNo
                ilFound = False
                Do
                    For ilLoop = 1 To 12 Step 1
                        If imPdStartWk = ilStartWk(ilLoop) Then
                            ilFound = True
                            Exit Do
                        End If
                    Next ilLoop
                    If imPdStartWk <= 1 Then
                        imPdStartWk = 1
                        ilFound = True
                        Exit Do
                    End If
                    imPdStartWk = imPdStartWk - 1
                Loop Until ilFound
            ElseIf imTypeIndex = 2 Then 'Weeks- make sure not pass end
                mCompMonths imPdYear, ilStartWk(), ilNoWks()
                imPdStartWk = tmPdGroups(4).iStartWkNo
                ilFound = False
                Do
                    If (imPdStartWk > ilStartWk(12) + ilNoWks(12) - 4) Then
                        If imPdStartWk <= 1 Then
                            imPdStartWk = 1
                            ilFound = True
                            Exit Do
                        End If
                        imPdStartWk = imPdStartWk - 1
                    Else
                        ilFound = True
                        Exit Do
                    End If
                Loop Until ilFound
            End If
        End If
        imShowIndex = Index
        mGetShowDates
        pbcRateCard_Paint
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub rbcShow_GotFocus(Index As Integer)
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
End Sub

Private Sub rbcType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcType(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilFound As Integer
    ReDim ilStartWk(0 To 12) As Integer 'Index zero ignored
    ReDim ilNoWks(0 To 12) As Integer

    If Value Then
        Screen.MousePointer = vbHourglass
        pbcRateCard.Cls
        If imRifStartYear <> 0 Then
            imPdYear = tmPdGroups(1).iYear
            If Index = 0 Then   'Change to Quarter
                imPdStartWk = 1
            ElseIf Index = 1 Then   'Month
                If (imTypeIndex = 0) Or (imTypeIndex = 3) Then
                    imPdStartWk = 1
                Else    'by week- back up to start of month
                    mCompMonths imPdYear, ilStartWk(), ilNoWks()
                    imPdStartWk = tmPdGroups(1).iStartWkNo
                    ilFound = False
                    Do
                        For ilLoop = 1 To 12 Step 1
                            If imPdStartWk = ilStartWk(ilLoop) Then
                                ilFound = True
                                Exit Do
                            End If
                        Next ilLoop
                        If imPdStartWk <= 1 Then
                            imPdStartWk = 1
                            ilFound = True
                            Exit Do
                        End If
                        imPdStartWk = imPdStartWk - 1
                    Loop Until ilFound
                End If
            ElseIf Index = 2 Then   'Week
                If (imTypeIndex = 0) Or (imTypeIndex = 3) Then
                    imPdStartWk = 1
                Else  'Month
                    imPdStartWk = tmPdGroups(1).iStartWkNo
                End If
            Else    'Flight
                imPdStartFltNo = 1
            End If
        End If
        imTypeIndex = Index
        mGetShowDates
        pbcRateCard_Paint
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub rbcType_GotFocus(Index As Integer)
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imRCBoxNo
        Case DAYPARTINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcDPNameRow, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub

Private Sub tmcDelay_Timer()
    Dim slNameFac As String
    tmcDelay.Enabled = False
    If sgMRcfStamp <> gFileDateTime(sgDBPath & "Rcf.Btr") Then
        sgMRcfStamp = gFileDateTime(sgDBPath & "Rcf.Btr")
        sgMRifStamp = gFileDateTime(sgDBPath & "Rif.Btr")
        tmcDelay.Enabled = True
        Exit Sub
    End If
    If sgMRifStamp <> gFileDateTime(sgDBPath & "Rif.Btr") Then
        sgMRcfStamp = gFileDateTime(sgDBPath & "Rcf.Btr")
        sgMRifStamp = gFileDateTime(sgDBPath & "Rif.Btr")
        tmcDelay.Enabled = True
        Exit Sub
    End If
    If ((imSelectedIndex > 0) And (imAdjIndex = 1)) Or ((imSelectedIndex >= 0) And (imAdjIndex = 0)) Then
        slNameFac = cbcSelect.List(imSelectedIndex)
    Else
        slNameFac = ""
    End If
    mPopulate
    gFindMatch slNameFac, 0, cbcSelect
    If (gLastFound(cbcSelect) > 0 And (imAdjIndex = 1)) Or (gLastFound(cbcSelect) >= 0 And (imAdjIndex = 0)) Then
        cbcSelect.ListIndex = gLastFound(cbcSelect)
    End If
    imRCBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    imFirstTimeSelect = True
    igRcfChg = False
    imRifChg = False
    'llRowNo = UBound(tmRifRec)
    'For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
    '    slStr = ""
    '    gSetShow pbcRateCard, slStr, tmRCCtrls(ilLoop)
    '    smShow(ilLoop, llRowNo) = tmRCCtrls(ilLoop).sShow
    'Next ilLoop
    pbcRateCard_Paint
    mSetDefInSave   'Set defaults for extra row
    mSetCommands
    If pbcSTab.Enabled Then
        pbcSTab.SetFocus
    Else
        'To avoid an Invalid procedure call or argument at Mobility.
        'cmcCancel.SetFocus
    End If
    Screen.MousePointer = vbDefault    'Default
End Sub

Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcRateCard.LargeChange + 1
            If UBound(smRCSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smRCSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(VEHINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmRCCtrls(VEHINDEX).fBoxY + tmRCCtrls(VEHINDEX).fBoxH)) Then
                    'If tgRpfI(ilRow + vbcRateCard.Value - 1).iCode <> 0 Then
                    '    Beep
                    '    Exit Sub
                    'End If
                    mSPSetShow imSPBoxNo
                    imSPBoxNo = -1
                    mRCSetShow imRCBoxNo
                    imRCBoxNo = -1
                    lmRCRowNo = -1
                    lmRCRowNo = ilRow + vbcRateCard.Value - 1
                    If imView = 0 Then
                        lacRCFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                        lacRCFrame.Move 0, tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacRCFrame.Visible = True
                    ElseIf imView = 1 Then
                        lacDPFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                        lacDPFrame.Move 0, tmDPCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) - 30
                        'If gInvertArea call then remove visible setting
                        lacDPFrame.Visible = True
                    End If
                    pbcArrow.Move pbcArrow.Left, plcRateCard.Top + tmRCCtrls(VEHINDEX).fBoxY + (lmRCRowNo - vbcRateCard.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    If imView = 0 Then
                        lacRCFrame.Drag vbBeginDrag
                        lacRCFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    ElseIf imView = 1 Then
                        lacDPFrame.Drag vbBeginDrag
                        lacDPFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    End If
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub

Private Sub vbcRateCard_Change()
    If imSettingValue Then
        If imView = 2 Then  'Comparison
            pbcRateCard.Cls
            pbcRateCard_Paint
            imSettingValue = False
        ElseIf imView = 1 Then  'Daypart
            pbcDaypart.Cls
            pbcDaypart_Paint
            imSettingValue = False
        Else
            pbcRateCard.Cls
            pbcRateCard_Paint
            imSettingValue = False
        End If
    Else
        mSPSetShow imSPBoxNo    'Remove focus
        imSPBoxNo = -1
        If imView = 2 Then  'Comparison
            pbcRateCard.Cls
            pbcRateCard_Paint
        ElseIf imView = 1 Then  'Daypart
            mRCSetShow imRCBoxNo
            imRCBoxNo = -1
            lmRCRowNo = -1
            'imRCBoxNo = -1
            'pbcArrow.Visible = False
            'lacRCFrame.Visible = False
            'lacDPFrame.Visible = False
            pbcDaypart.Cls
            pbcDaypart_Paint
            If (igWinStatus(RATECARDSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
                mRCEnableBox imRCBoxNo
            End If
        Else
            mRCSetShow imRCBoxNo
            imRCBoxNo = -1
            lmRCRowNo = -1
            'pbcArrow.Visible = False
            'lacRCFrame.Visible = False
            'lacDPFrame.Visible = False
            pbcRateCard.Cls
            pbcRateCard_Paint
            'If (igWinStatus(RATECARDSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            '    mRCEnableBox imRCBoxNo
            'End If
        End If
    End If
End Sub

Private Sub vbcRateCard_GotFocus()
    mSPSetShow imSPBoxNo    'Remove focus
    imSPBoxNo = -1
    mRCSetShow imRCBoxNo
    imRCBoxNo = -1
    pbcArrow.Visible = False
    lacRCFrame.Visible = False
    lacDPFrame.Visible = False
    gCtrlGotFocus vbcRateCard
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Rate Card"
End Sub

Private Sub plcShow_Paint()
    plcShow.CurrentX = 0
    plcShow.CurrentY = 0
    plcShow.Print "Show by"
End Sub

Private Function mBinarySearch(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tmUserVeh)
    ilMax = UBound(tmUserVeh) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tmUserVeh(ilMiddle).iCode Then
            'found the match
            mBinarySearch = ilMiddle
            Exit Function
        ElseIf ilCode < tmUserVeh(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearch = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mPaintTitle                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Paint Header Titles            *
'*                                                     *
'*******************************************************
Private Sub mPaintRCTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcRateCard.ForeColor
    slFontName = pbcRateCard.FontName
    flFontSize = pbcRateCard.FontSize
    ilFillStyle = pbcRateCard.FillStyle
    llFillColor = pbcRateCard.FillColor
    pbcRateCard.ForeColor = BLUE
    pbcRateCard.FontBold = False
    pbcRateCard.FontSize = 7
    pbcRateCard.FontName = "Arial"
    pbcRateCard.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmRCCtrls(VEHINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    
    pbcRateCard.Line (tmRCCtrls(VEHINDEX).fBoxX - 15, tmRCCtrls(VEHINDEX).fBoxH + 30)-Step(tmRCCtrls(VEHINDEX).fBoxW + 15, tmRCCtrls(VEHINDEX).fBoxH + 15), BLUE, B
    pbcRateCard.CurrentX = tmRCCtrls(VEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcRateCard.CurrentY = tmRCCtrls(VEHINDEX).fBoxH + 30 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcRateCard.Print "Vehicle"
    
    pbcRateCard.Line (tmRCCtrls(DAYPARTINDEX).fBoxX - 15, tmRCCtrls(DAYPARTINDEX).fBoxH + 30)-Step(tmRCCtrls(DAYPARTINDEX).fBoxW + 15, tmRCCtrls(DAYPARTINDEX).fBoxH + 15), BLUE, B
    pbcRateCard.CurrentX = tmRCCtrls(DAYPARTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcRateCard.CurrentY = tmRCCtrls(DAYPARTINDEX).fBoxH + 30
    pbcRateCard.Print "Daypart"
    
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
        'L.Bianchi
        pbcRateCard.Line (tmRCCtrls(CPMINDEX).fBoxX - 15, tmRCCtrls(CPMINDEX).fBoxH + 30)-Step(tmRCCtrls(CPMINDEX).fBoxW + 15, tmRCCtrls(CPMINDEX).fBoxH + 15), BLUE, B
        pbcRateCard.CurrentX = tmRCCtrls(CPMINDEX).fBoxX + 15  'fgBoxInsetX
        pbcRateCard.CurrentY = tmRCCtrls(CPMINDEX).fBoxH + 30
        pbcRateCard.Print "CPM"
    End If
    
    pbcRateCard.Line (tmRCCtrls(ACQUISITIONINDEX).fBoxX - 15, tmRCCtrls(ACQUISITIONINDEX).fBoxH + 30)-Step(tmRCCtrls(ACQUISITIONINDEX).fBoxW + 15, tmRCCtrls(ACQUISITIONINDEX).fBoxH + 15), BLUE, B
    pbcRateCard.CurrentX = tmRCCtrls(ACQUISITIONINDEX).fBoxX + 15  'fgBoxInsetX
    pbcRateCard.CurrentY = tmRCCtrls(ACQUISITIONINDEX).fBoxH + 30
    If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
        pbcRateCard.Print "Acq Cost"
    End If
    pbcRateCard.Line (tmRCCtrls(BASEINDEX).fBoxX - 15, tmRCCtrls(BASEINDEX).fBoxH + 30)-Step(tmRCCtrls(BASEINDEX).fBoxW + 15, tmRCCtrls(BASEINDEX).fBoxH + 15), BLUE, B
    pbcRateCard.CurrentX = tmRCCtrls(BASEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcRateCard.CurrentY = tmRCCtrls(BASEINDEX).fBoxH + 30
    pbcRateCard.Print "Base"
    pbcRateCard.Line (tmRCCtrls(RPTINDEX).fBoxX - 15, tmRCCtrls(RPTINDEX).fBoxH + 30)-Step(tmRCCtrls(RPTINDEX).fBoxW + 15, tmRCCtrls(RPTINDEX).fBoxH + 15), BLUE, B
    pbcRateCard.CurrentX = tmRCCtrls(RPTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcRateCard.CurrentY = tmRCCtrls(RPTINDEX).fBoxH + 30
    pbcRateCard.Print "Rpt"
    pbcRateCard.Line (tmRCCtrls(SORTINDEX).fBoxX - 15, tmRCCtrls(SORTINDEX).fBoxH + 30)-Step(tmRCCtrls(SORTINDEX).fBoxW + 15, tmRCCtrls(SORTINDEX).fBoxH + 15), BLUE, B
    pbcRateCard.CurrentX = tmRCCtrls(SORTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcRateCard.CurrentY = tmRCCtrls(SORTINDEX).fBoxH + 30
    pbcRateCard.Print "Sort"
    pbcRateCard.Line (tmWKCtrls(WK1INDEX).fBoxX - 15, 15)-Step(tmWKCtrls(WK1INDEX).fBoxW + 15, tmWKCtrls(WK1INDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(DOLLAR1INDEX).fBoxX - 15, tmRCCtrls(DOLLAR1INDEX).fBoxH + 30)-Step(tmRCCtrls(DOLLAR1INDEX).fBoxW + 15, tmRCCtrls(DOLLAR1INDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(DOLLAR1INDEX).fBoxX, tmRCCtrls(DOLLAR1INDEX).fBoxH + 45)-Step(tmRCCtrls(DOLLAR1INDEX).fBoxW - 15, tmRCCtrls(DOLLAR1INDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcRateCard.Line (tmWKCtrls(WK2INDEX).fBoxX - 15, 15)-Step(tmWKCtrls(WK2INDEX).fBoxW + 15, tmWKCtrls(WK2INDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(DOLLAR2INDEX).fBoxX - 15, tmRCCtrls(DOLLAR2INDEX).fBoxH + 30)-Step(tmRCCtrls(DOLLAR2INDEX).fBoxW + 15, tmRCCtrls(DOLLAR2INDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(DOLLAR2INDEX).fBoxX, tmRCCtrls(DOLLAR2INDEX).fBoxH + 45)-Step(tmRCCtrls(DOLLAR2INDEX).fBoxW - 15, tmRCCtrls(DOLLAR2INDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcRateCard.Line (tmWKCtrls(WK3INDEX).fBoxX - 15, 15)-Step(tmWKCtrls(WK3INDEX).fBoxW + 15, tmWKCtrls(WK3INDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(DOLLAR3INDEX).fBoxX - 15, tmRCCtrls(DOLLAR3INDEX).fBoxH + 30)-Step(tmRCCtrls(DOLLAR3INDEX).fBoxW + 15, tmRCCtrls(DOLLAR3INDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(DOLLAR3INDEX).fBoxX, tmRCCtrls(DOLLAR3INDEX).fBoxH + 45)-Step(tmRCCtrls(DOLLAR3INDEX).fBoxW - 15, tmRCCtrls(DOLLAR3INDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcRateCard.Line (tmWKCtrls(WK4INDEX).fBoxX - 15, 15)-Step(tmWKCtrls(WK4INDEX).fBoxW + 15, tmWKCtrls(WK4INDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(DOLLAR4INDEX).fBoxX - 15, tmRCCtrls(DOLLAR4INDEX).fBoxH + 30)-Step(tmRCCtrls(DOLLAR4INDEX).fBoxW + 15, tmRCCtrls(DOLLAR4INDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(DOLLAR4INDEX).fBoxX, tmRCCtrls(DOLLAR4INDEX).fBoxH + 45)-Step(tmRCCtrls(DOLLAR4INDEX).fBoxW - 15, tmRCCtrls(DOLLAR4INDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcRateCard.Line (tmRCCtrls(AVGINDEX).fBoxX - 15, tmRCCtrls(AVGINDEX).fBoxH + 30)-Step(tmRCCtrls(AVGINDEX).fBoxW + 15, tmRCCtrls(AVGINDEX).fBoxH + 15), BLUE, B
    pbcRateCard.Line (tmRCCtrls(AVGINDEX).fBoxX, tmRCCtrls(AVGINDEX).fBoxH + 45)-Step(tmRCCtrls(AVGINDEX).fBoxW - 15, tmRCCtrls(AVGINDEX).fBoxY - 60), LIGHTYELLOW, BF
    'pbcRateCard.CurrentX = tmRCCtrls(AVGINDEX).fBoxX + 15  'fgBoxInsetX
    'pbcRateCard.CurrentY = tmRCCtrls(AVGINDEX).fBoxH + 30
    'pbcRateCard.Print "Average"

    ilLineCount = 0
    llTop = tmRCCtrls(1).fBoxY
    Do
        For ilLoop = imLBRCCtrls To UBound(tmRCCtrls) Step 1
            If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER And ilLoop = CPMINDEX Then
                GoTo Skip_Loop
            End If
        pbcRateCard.Line (tmRCCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmRCCtrls(ilLoop).fBoxW + 15, tmRCCtrls(ilLoop).fBoxH + 15), BLUE, B
Skip_Loop:
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmRCCtrls(1).fBoxH + 15

    Loop While llTop + tmRCCtrls(1).fBoxH < pbcRateCard.Height
    vbcRateCard.LargeChange = ilLineCount - 1
    pbcRateCard.FontSize = flFontSize
    pbcRateCard.FontName = slFontName
    pbcRateCard.FontSize = flFontSize
    pbcRateCard.ForeColor = llColor
    pbcRateCard.FontBold = True
End Sub

Private Sub mPaintDPTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer
    Dim llWidth As Integer
    Dim ilDay As Integer

    llColor = pbcDaypart.ForeColor
    slFontName = pbcDaypart.FontName
    flFontSize = pbcDaypart.FontSize
    ilFillStyle = pbcDaypart.FillStyle
    llFillColor = pbcDaypart.FillColor
    pbcDaypart.ForeColor = BLUE
    pbcDaypart.FontBold = False
    pbcDaypart.FontSize = 7
    pbcDaypart.FontName = "Arial"
    pbcDaypart.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    ilHalfY = tmDPCtrls(VEHINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    pbcDaypart.Line (tmDPCtrls(VEHINDEX).fBoxX - 15, tmDPCtrls(VEHINDEX).fBoxH + 30)-Step(tmDPCtrls(VEHINDEX).fBoxW + 15, tmDPCtrls(VEHINDEX).fBoxH + 15), BLUE, B
    pbcDaypart.CurrentX = tmDPCtrls(VEHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcDaypart.CurrentY = tmDPCtrls(VEHINDEX).fBoxH + 30 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcDaypart.Print "Vehicle"
    
    pbcDaypart.Line (tmDPCtrls(DAYPARTINDEX).fBoxX - 15, tmDPCtrls(DAYPARTINDEX).fBoxH + 30)-Step(tmDPCtrls(DAYPARTINDEX).fBoxW + 15, tmDPCtrls(DAYPARTINDEX).fBoxH + 15), BLUE, B
    pbcDaypart.CurrentX = tmDPCtrls(DAYPARTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcDaypart.CurrentY = tmDPCtrls(DAYPARTINDEX).fBoxH + 30
    pbcDaypart.Print "Daypart"
    'added by L. Bianchi
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
        pbcDaypart.Line (tmDPCtrls(CPMINDEX).fBoxX - 15, tmDPCtrls(CPMINDEX).fBoxH + 30)-Step(tmDPCtrls(CPMINDEX).fBoxW + 15, tmDPCtrls(CPMINDEX).fBoxH + 15), BLUE, B
        pbcDaypart.CurrentX = tmDPCtrls(CPMINDEX).fBoxX + 15  'fgBoxInsetX
        pbcDaypart.CurrentY = tmDPCtrls(CPMINDEX).fBoxH + 30
        pbcDaypart.Print "CPM"
    End If
    pbcDaypart.Line (tmDPCtrls(TIMESINDEX).fBoxX - 15, tmDPCtrls(TIMESINDEX).fBoxH + 30)-Step(tmDPCtrls(TIMESINDEX).fBoxW + 15, tmDPCtrls(TIMESINDEX).fBoxH + 15), BLUE, B
    pbcDaypart.CurrentX = tmDPCtrls(TIMESINDEX).fBoxX + 15  'fgBoxInsetX
    pbcDaypart.CurrentY = tmDPCtrls(TIMESINDEX).fBoxH + 30
    pbcDaypart.Print "Times or Library Names"
    pbcDaypart.Line (tmDPCtrls(DAYINDEX).fBoxX - 15, tmDPCtrls(DAYINDEX).fBoxH + 30)-Step(tmDPCtrls(AVAILINDEX).fBoxX - tmDPCtrls(DAYINDEX).fBoxX, tmDPCtrls(DAYINDEX).fBoxH + 15), BLUE, B
    For ilDay = DAYINDEX To DAYINDEX + 6 Step 1
        pbcDaypart.CurrentX = tmDPCtrls(ilDay).fBoxX + 15  'fgBoxInsetX
        pbcDaypart.CurrentY = tmDPCtrls(ilDay).fBoxH + 30
        Select Case ilDay
            Case DAYINDEX
                pbcDaypart.Print "Mo"
            Case DAYINDEX + 1
                pbcDaypart.Print "Tu"
            Case DAYINDEX + 2
                pbcDaypart.Print "We"
            Case DAYINDEX + 3
                pbcDaypart.Print "Th"
            Case DAYINDEX + 4
                pbcDaypart.Print "Fr"
            Case DAYINDEX + 5
                pbcDaypart.Print "Sa"
            Case DAYINDEX + 6
                pbcDaypart.Print "Su"
        End Select
    Next ilDay
    pbcDaypart.Line (tmDPCtrls(AVAILINDEX).fBoxX - 15, tmDPCtrls(AVAILINDEX).fBoxH + 30)-Step(tmDPCtrls(AVAILINDEX).fBoxW + 15, tmDPCtrls(AVAILINDEX).fBoxH + 15), BLUE, B
    pbcDaypart.CurrentX = tmDPCtrls(AVAILINDEX).fBoxX + 15  'fgBoxInsetX
    pbcDaypart.CurrentY = tmDPCtrls(AVAILINDEX).fBoxH + 30
    pbcDaypart.Print "Avail Name"
    pbcDaypart.Line (tmDPCtrls(HRSINDEX).fBoxX - 15, tmDPCtrls(HRSINDEX).fBoxH + 30)-Step(tmDPCtrls(HRSINDEX).fBoxW + 15, tmDPCtrls(HRSINDEX).fBoxH + 15), BLUE, B
    pbcDaypart.CurrentX = tmDPCtrls(HRSINDEX).fBoxX + 15  'fgBoxInsetX
    pbcDaypart.CurrentY = tmDPCtrls(HRSINDEX).fBoxH + 30
    pbcDaypart.Print "Hrs"

    llWidth = tmDPCtrls(HRSINDEX).fBoxX + tmDPCtrls(HRSINDEX).fBoxW - tmDPCtrls(TIMESINDEX).fBoxX
    pbcDaypart.Line (tmDPCtrls(TIMESINDEX).fBoxX - 15, 15)-Step(llWidth + 15, tmDPCtrls(HRSINDEX).fBoxH + 15), BLUE, B
    pbcDaypart.CurrentX = tmDPCtrls(TIMESINDEX).fBoxX + llWidth / 2 - pbcDaypart.TextWidth("Daypart Information") / 2 + 15 'fgBoxInsetX
    pbcDaypart.CurrentY = 30
    pbcDaypart.Print "Daypart Information"

    ilLineCount = 0
    llTop = tmDPCtrls(1).fBoxY
    Do
        For ilLoop = imLBDPCtrls To HRSINDEX Step 1
            If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER And ilLoop = CPMINDEX Then
                    GoTo Skip_Loop
            End If
            pbcDaypart.Line (tmDPCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmDPCtrls(ilLoop).fBoxW + 15, tmDPCtrls(ilLoop).fBoxH + 15), BLUE, B
Skip_Loop:
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmDPCtrls(1).fBoxH + 15
    Loop While llTop + tmDPCtrls(1).fBoxH < pbcDaypart.Height
    vbcRateCard.LargeChange = ilLineCount - 1
    pbcDaypart.FontSize = flFontSize
    pbcDaypart.FontName = slFontName
    pbcDaypart.FontSize = flFontSize
    pbcDaypart.ForeColor = llColor
    pbcDaypart.FontBold = True
End Sub

Public Function mSetImportPrice(slType As String, slVehicle As String, slDayPart As String, slDollars As String) As Integer
    Dim llRow As Long
    Dim blMatchVehicle As Boolean
    Dim blMatchDaypart As Boolean
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim llDollars As Long
    Dim llLkYear As Long
    
    If slType = "D" Then    'Dollar setting
        llDollars = Val(slDollars)
        blMatchVehicle = False
        blMatchDaypart = False
        For llRow = LBONE To UBound(tmRifRec) - 1 Step 1
            If UCase$(Trim$(smRCSave(VEHINDEX, llRow))) = slVehicle Then
                blMatchVehicle = True
                If UCase$(Trim$(smRCSave(DAYPARTINDEX, llRow))) = slDayPart Then
                    blMatchDaypart = True
                    If tmRifRec(llRow).tRif.iYear = tgRcfI.iYear Then
                        For ilWk = 1 To 53 Step 1
                            'TTP 10359 - Rate Card Import Not Updating $0 Rates
                            'If tmRifRec(llRow).tRif.lRate(ilWk) > 0 Then
                                tmRifRec(llRow).tRif.lRate(ilWk) = llDollars
                            'End If
                        Next ilWk
                    Else
                        llLkYear = tmRifRec(llRow).lLkYear
                        Do While llLkYear > 0
                            If tmLkRifRec(llLkYear).tRif.iYear = tgRcfI.iYear Then
                                For ilWk = 1 To 53 Step 1
                                    'TTP 10359 - Rate Card Import Not Updating $0 Rates
                                    'If tmRifRec(llRow).tRif.lRate(ilWk) > 0 Then
                                        tmRifRec(llRow).tRif.lRate(ilWk) = llDollars
                                    'End If
                                Next ilWk
                            Else
                                llLkYear = tmLkRifRec(llLkYear).lLkYear
                            End If
                        Loop
                    End If
                End If
            End If
            If blMatchDaypart Then
                lmRCRowNo = llRow
                mGetShowPrices lmRCRowNo    'Set color flag
                imRifChg = True
                Exit For
            End If
        Next llRow
        If Not blMatchVehicle Then
            mSetImportPrice = 1
        ElseIf Not blMatchDaypart Then
            mSetImportPrice = 2
        Else
            mSetImportPrice = 0
        End If
    ElseIf slType = "R" Then    'Reset package rate
        mResetStdPrice
    ElseIf slType = "C" Then    'Completed
        pbcRateCard.Cls
        pbcRateCard_Paint
        mSetCommands
    End If
End Function

Private Sub mResetStdPrice()
    Dim ilRow As Integer
    For ilRow = LBONE To UBound(tmRifRec) - 1 Step 1
        lmRCRowNo = ilRow
        mGetStdPkgPrice False
    Next ilRow
End Sub

Public Sub mAddNewRow(llRowNo As Long)
    ReDim Preserve tmRifRec(0 To llRowNo) As RIFREC
    ReDim Preserve smRCSave(0 To SORTINDEX, 0 To llRowNo) As String * 40
    ReDim Preserve lmRCSave(0 To TOTALINDEX - SORTINDEX, 0 To llRowNo) As Long
    'updated by l.bianchi
    ReDim Preserve imRCSave(0 To 17, 0 To llRowNo) As Integer
    ReDim Preserve smRCShow(0 To AVGINDEX, 0 To llRowNo) As String * 40
    ReDim Preserve smDPShow(0 To DPBASEINDEX, 0 To llRowNo) As String * 40 'Values shown in program area
    mInitRif llRowNo
    If UBound(tmRifRec) <= vbcRateCard.LargeChange Then 'was <=
        vbcRateCard.Max = LBONE 'LBound(tmRifRec)
    Else
        vbcRateCard.Max = UBound(tmRifRec) - vbcRateCard.LargeChange '-1
    End If
    mSetDefInSave   'Set defaults for extra row
    mInitLastShow  'Init last row
End Sub

'TTP 10388 - Rate card screen: if "ad server" site option setting is not enabled, a "podcast" medium type vehicle does not appear on the rate card screen
''L.Bianchi
Private Sub mSetVehicleMediumType(llRowNo As Long)
    '02/23/22 - JW Fix issue on the Rate Card screen per Jason Email: Wed 2/16/22 10:46 AM
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER Then
        Exit Sub
    End If
    'L.Bianchi
    Dim ilRet As Integer
    Dim ilCode As Integer
    Dim ilIndex As Integer
    imRCSave(17, llRowNo) = 0
    ilCode = mGetVehicleCode(llRowNo)
    If (mGetVehicleMedium(ilCode) = "P") Then
        imRCSave(17, llRowNo) = 1
    End If
End Sub

'L.Bianchi
Private Function mGetVehicleMedium(ilVefCode As Integer) As String
    '02/23/22 - JW Fix issue on the Rate Card screen per Jason Email: Wed 2/16/22 10:46 AM
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER Then
        Exit Function
    End If
    Dim ilLoop As Integer
    Dim ilVpf As Integer
    If ilVefCode = 0 Then
        mGetVehicleMedium = ""
    Else
        ilLoop = gBinarySearchVef(ilVefCode)
        If ilLoop <> -1 Then
            ilVpf = gBinarySearchVpf(ilVefCode)
            If ilVpf <> -1 Then
                mGetVehicleMedium = tgVpf(ilVpf).sGMedium
            End If
        End If
    End If
End Function

'L.Bianchi
Private Sub mSetPvfType(llRowNo As Long)
    Dim ilLoop As Integer
    Dim vefMedium As String
    Dim ilVpf As Integer
    Dim ilRet As Integer
    Dim ilCode As Integer
    Dim llPvfCode As Long
    imRCSave(16, llRowNo) = 0
    ilCode = mGetVehicleCode(llRowNo)
    ilLoop = gBinarySearchVef(ilCode)
    If ilLoop >= 0 Then
        If (tgMVef(ilLoop).sType = "P") And (tgMVef(ilLoop).lPvfCode > 0) Then
            llPvfCode = tgMVef(ilLoop).lPvfCode
            ReDim tmPvf(0 To 0) As PVF
            Do While llPvfCode > 0
                tmPvfSrchKey.lCode = llPvfCode
                ilRet = btrGetEqual(hmPvf, tmPvf(UBound(tmPvf)), imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                llPvfCode = tmPvf(UBound(tmPvf)).lLkPvfCode
                ReDim Preserve tmPvf(0 To UBound(tmPvf) + 1) As PVF
                If tmPvf(0).sType = "C" Then
                    '(Package vehicle) "C"=Podcast Ad Server buy only
                    imRCSave(16, llRowNo) = 1
                End If
            Loop
        End If
    End If
End Sub

'TTP 10388 - Rate card screen: if "ad server" site option setting is not enabled, a "podcast" medium type vehicle does not appear on the rate card screen
''L.Bianchi
'Private Sub mRemoveCPMVehicle()
'    Dim ilLoop As Integer
'    Dim ilIndex As Integer
'    Dim slNameCode As String
'    Dim slCode As String
'    Dim ilCode As Integer
'    Dim ilVpf As Integer
'    Dim ilRet As Integer
'    Dim slStamp As String   'Current time stamp
'    Dim filterVehicle() As String
'
'    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
'        ReDim tgTempRCUserVehicle(UBound(tgRCUserVehicle))
'        tgTempRCUserVehicle = tgRCUserVehicle
'        Exit Sub
'    End If
'
'    slStamp = gFileDateTime(sgDBPath & "Vef.btr") & Trim$(str$(VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHSTDPKG + VEHCPMPKG + VEHSPORT + ACTIVEVEH))
'     If sgRCUserVehicleTag <> "" Then
'        If StrComp(slStamp, sgRCUserVehicleTag, 1) = 0 Then
'            If (Not tgTempRCUserVehicle) <> -1 Then
'                Exit Sub
'            End If
'        End If
'    End If
'    ilIndex = 0
'
'    For ilLoop = 0 To lbcVehicle.ListCount - 1
'        slNameCode = tgRCUserVehicle(ilLoop).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
'        ilRet = gParseItem(slNameCode, 2, "\", slCode)
'        ilCode = CInt(slCode)
'        ilVpf = gBinarySearchVpf(ilCode)
'        If ilVpf <> -1 Then
'            'TTP 10388 - Rate card screen: if "ad server" site option setting is not enabled, a "podcast" medium type vehicle does not appear on the rate card screen
'            ilRet = gParseItem(slNameCode, 3, "|", slCode)
'            ilRet = gParseItem(slCode, 1, "\", slCode)
'            ReDim Preserve tgTempRCUserVehicle(ilIndex)
'            ReDim Preserve filterVehicle(ilIndex)
'            filterVehicle(ilIndex) = slCode
'            tgTempRCUserVehicle(ilIndex) = tgRCUserVehicle(ilLoop)
'            ilIndex = ilIndex + 1
'            'End If
'        End If
'    Next ilLoop
'    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
'    lbcVehicle.Clear
'    For ilLoop = 0 To UBound(filterVehicle)
'        lbcVehicle.AddItem filterVehicle(ilLoop)
'    Next ilLoop
'End Sub

Private Sub mResetColumn(llRowNo As Long)
   Dim iVefCodeNew As Integer
   Dim iVefCodeOld As Integer
   Dim iVehTypeNew As String
   Dim iVehTypeOld As String
   iVefCodeOld = tmRifRec(llRowNo).tRif.iVefCode
   iVefCodeNew = mGetVehicleCode(llRowNo)
       If iVefCodeNew = 0 Then
        Exit Sub
   End If
   If iVefCodeOld = iVefCodeNew Then
        Exit Sub
   End If
   iVehTypeNew = mGetVehicleMedium(iVefCodeNew)
   iVehTypeOld = mGetVehicleMedium(iVefCodeOld)
   
   tmRifRec(llRowNo).tRif.iVefCode = iVefCodeNew
   If iVehTypeNew = "P" Then
       If imRCSave(16, llRowNo) = 1 Then
          If Trim$(smRCShow(CPMINDEX, llRowNo)) <> "" Then
              smRCShow(CPMINDEX, llRowNo) = ""
              mRCSetShow CPMINDEX
          End If
       End If
       ' LB 02/10/21
       'If iVehTypeOld <> "P" Then
       ' smRCShow(BASEINDEX, llRowNo) = "N"
       ' mRCSetShow BASEINDEX
       ' smRCShow(RPTINDEX, llRowNo) = "N"
       ' mRCSetShow RPTINDEX
       ' smRCShow(DOLLAR1INDEX, llRowNo) = ""
       ' mRCSetShow DOLLAR1INDEX
       ' smRCShow(DOLLAR2INDEX, llRowNo) = ""
       ' mRCSetShow DOLLAR2INDEX
       ' smRCShow(DOLLAR3INDEX, llRowNo) = ""
       ' mRCSetShow DOLLAR3INDEX
       ' smRCShow(DOLLAR4INDEX, llRowNo) = ""
       ' mRCSetShow DOLLAR4INDEX
       ' End If
    Else
       If Trim$(smRCShow(CPMINDEX, llRowNo)) <> "" Then
           smRCShow(CPMINDEX, llRowNo) = ""
           mRCSetShow CPMINDEX
       End If
    End If
End Sub

Private Function mGetVehicleCode(llRowNo As Long) As Long
    Dim slCode As String
    Dim slNameCode As String
    Dim ilCode As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    mGetVehicleCode = 0
    gFindMatch Trim$(smRCSave(VEHINDEX, llRowNo)), 0, lbcVehicle
    ilIndex = gLastFound(lbcVehicle)
    'TTP 10387  - Rate Card: subscript out of range error when switching rate cards
    If ilIndex >= 0 And ilIndex < UBound(tgTempRCUserVehicle) Then
        slNameCode = tgTempRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        mGetVehicleCode = CInt(slCode)
    End If
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
    Dim slStartIn As String
    Dim slCSIName As String
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer

    
    sgCommandStr = Command$
    slStartIn = CurDir$
    slCSIName = ""
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommandStr, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
    slCommand = sgCommandStr    'Command$
    lgCurrHRes = GetDeviceCaps(Traffic!pbcList.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(Traffic!pbcList.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(Traffic!pbcList.hdc, BITSPIXEL)
    mTestPervasive
    '4/2/11: Add setting of value
    lgUlfCode = 0
    'If (Trim$(sgCommandStr) = "") Or (Trim$(sgCommandStr) = "/UserInput") Or (Trim$(sgCommandStr) = "Debug") Then
    If InStr(1, sgCommandStr, "^", vbTextCompare) <= 0 Then
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If
        slStr = sgUserName
        sgCallAppName = "Traffic"
    Else
        igSportsSystem = 0
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        'ilRet = gParseItem(slCommand, 3, "\", slStr)
        'igRptCallType = Val(slStr)
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
        sgUrfStamp = "~" 'Clear time stamp incase same name
        sgUserName = Trim$(slStr)
        '6/20/09:  Jim requested that the Guide sign in be changed to CSI for internal Guide only
        If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
            slDate = Format$(Now(), "m/d/yy")
            slMonth = Month(slDate)
            slYear = Year(slDate)
            llValue = Val(slMonth) * Val(slYear)
            ilValue = Int(10000 * Rnd(-llValue) + 1)
            llValue = ilValue
            ilValue = Int(10000 * Rnd(-llValue) + 1)
            slStr = Trim$(Str$(ilValue))
            Do While Len(slStr) < 4
                slStr = "0" & slStr
            Loop
            sgSpecialPassword = slStr
            slCSIName = "CSI"
            sgUserName = "Guide"
        End If
        gUrfRead Signon, sgUserName, True, tgUrf(), False  'Obtain user records
        If StrComp(slCSIName, "CSI", vbTextCompare) = 0 Then
            gExpandGuideAsUser tgUrf(0)
        End If
        mGetUlfCode
    End If
    'End If
    DoEvents
'    gInitStdAlone ReportList, slStr, igTestSystem
    gInitStdAlone
    mCheckForDate
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "UserOpt.Frm"
    'If igWinStatus(INVOICESJOB) = 0 Then
    '    imTerminate = True
    'End If
End Sub

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
        'ilPos = InStr(slCommand, ":")
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
        'Dan M 9/20/10 problems with gGetCSIName("SYSDate") in v57 reports.exe... change to global variable
     '   ilRet = csiSetName("SYSDate", slDate)
        ilRet = gCsiSetName(slDate)
    End If
End Sub
Private Sub mGetUlfCode()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    ilPos = InStr(1, sgCommandStr, "/ULF:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5)))
        Else
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5, ilSpace - ilPos - 3)))
        End If
    End If
End Sub


