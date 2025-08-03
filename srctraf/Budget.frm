VERSION 5.00
Begin VB.Form Budget 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5940
   ClientLeft      =   915
   ClientTop       =   1740
   ClientWidth     =   9360
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
   Icon            =   "Budget.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   9360
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7950
      Top             =   5460
   End
   Begin VB.PictureBox pbcKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3660
      Picture         =   "Budget.frx":08CA
      ScaleHeight     =   525
      ScaleWidth      =   3375
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   3405
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
      Left            =   45
      Picture         =   "Budget.frx":6690
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.CommandButton cmc12Mos 
      Appearance      =   0  'Flat
      Caption         =   "12 &Mo's"
      Enabled         =   0   'False
      Height          =   270
      HelpContextID   =   5
      Left            =   5355
      TabIndex        =   47
      Top             =   5310
      Width           =   990
   End
   Begin VB.CommandButton cmcAddVeh 
      Appearance      =   0  'Flat
      Caption         =   "&Add Veh."
      Height          =   270
      HelpContextID   =   7
      Left            =   6360
      TabIndex        =   48
      Top             =   5310
      Width           =   990
   End
   Begin VB.CommandButton cmcAdvt 
      Appearance      =   0  'Flat
      Caption         =   "Ad&vertiser"
      Enabled         =   0   'False
      Height          =   270
      HelpContextID   =   7
      Left            =   1755
      TabIndex        =   49
      Top             =   5610
      Width           =   1080
   End
   Begin VB.CommandButton cmcDemo 
      Appearance      =   0  'Flat
      Caption         =   "R&esearch"
      Enabled         =   0   'False
      Height          =   270
      HelpContextID   =   7
      Left            =   4860
      TabIndex        =   52
      Top             =   5610
      Width           =   990
   End
   Begin VB.CommandButton cmcActuals 
      Appearance      =   0  'Flat
      Caption         =   "Act&uals"
      Enabled         =   0   'False
      Height          =   270
      HelpContextID   =   7
      Left            =   5865
      TabIndex        =   53
      Top             =   5610
      Width           =   990
   End
   Begin VB.CommandButton cmcTrend 
      Appearance      =   0  'Flat
      Caption         =   "&Trend"
      Enabled         =   0   'False
      Height          =   270
      HelpContextID   =   7
      Left            =   2850
      TabIndex        =   50
      Top             =   5610
      Width           =   990
   End
   Begin VB.CommandButton cmcScale 
      Appearance      =   0  'Flat
      Caption         =   "Sc&ale"
      Enabled         =   0   'False
      Height          =   270
      HelpContextID   =   7
      Left            =   3855
      TabIndex        =   51
      Top             =   5610
      Width           =   990
   End
   Begin VB.PictureBox pbcDirect 
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
      Left            =   2145
      ScaleHeight     =   210
      ScaleWidth      =   840
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox pbcYear 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      Picture         =   "Budget.frx":699A
      ScaleHeight     =   375
      ScaleWidth      =   570
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   570
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
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   4785
      ScaleHeight     =   435
      ScaleWidth      =   4470
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   510
      Width           =   4470
      Begin VB.OptionButton rbcSort 
         Caption         =   "Office within Vehicle"
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
         Index           =   3
         Left            =   2130
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   2190
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Vehicle within Office"
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
         Index           =   2
         Left            =   15
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   210
         Width           =   2130
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Vehicle"
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
         Left            =   15
         TabIndex        =   16
         Top             =   0
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.OptionButton rbcSort 
         Caption         =   "Office"
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
         Left            =   2130
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   2085
      End
   End
   Begin VB.PictureBox plcComparison 
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
      Height          =   420
      Left            =   9135
      ScaleHeight     =   420
      ScaleWidth      =   5115
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   405
      Visible         =   0   'False
      Width           =   5115
      Begin VB.ComboBox cbcComparison 
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
         Left            =   1260
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   60
         Width           =   3795
      End
   End
   Begin VB.ListBox lbcSalesOffice 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   885
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2850
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.TextBox edcOVDropDown 
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
      Left            =   285
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.PictureBox pbcOSTab 
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
      Left            =   -90
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   12
      Top             =   840
      Width           =   105
   End
   Begin VB.PictureBox pbcOSSTab 
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
      Left            =   -90
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   465
      Width           =   105
   End
   Begin VB.PictureBox pbcBudgetName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   225
      Picture         =   "Budget.frx":6E3C
      ScaleHeight     =   375
      ScaleWidth      =   3300
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   510
      Width           =   3300
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Enabled         =   0   'False
      Height          =   270
      HelpContextID   =   5
      Left            =   4350
      TabIndex        =   46
      Top             =   5310
      Width           =   990
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   270
      HelpContextID   =   3
      Left            =   3345
      TabIndex        =   45
      Top             =   5310
      Width           =   990
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   270
      HelpContextID   =   2
      Left            =   2250
      TabIndex        =   44
      Top             =   5310
      Width           =   1080
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   270
      HelpContextID   =   1
      Left            =   1245
      TabIndex        =   43
      Top             =   5310
      Width           =   990
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
      Left            =   1560
      Picture         =   "Budget.frx":8526
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2145
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
      Left            =   495
      MaxLength       =   20
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2355
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   885
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2430
      Visible         =   0   'False
      Width           =   2550
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
      Height          =   75
      Left            =   180
      ScaleHeight     =   75
      ScaleWidth      =   90
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   5310
      Width           =   90
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
      Left            =   9285
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   5
      Top             =   195
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
      Left            =   30
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   31
      Top             =   5490
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
      Left            =   -90
      ScaleHeight     =   135
      ScaleWidth      =   105
      TabIndex        =   20
      Top             =   1020
      Width           =   105
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   885
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   885
   End
   Begin VB.PictureBox plcOS 
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
      Left            =   180
      ScaleHeight     =   420
      ScaleWidth      =   3345
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   465
      Width           =   3405
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
      Height          =   390
      Left            =   945
      ScaleHeight     =   330
      ScaleWidth      =   8175
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   8235
      Begin VB.PictureBox pbcType 
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
         ScaleWidth      =   1395
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   60
         Width           =   1395
      End
      Begin VB.OptionButton rbcOS 
         Caption         =   "Salesperson"
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
         Left            =   1095
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Width           =   1440
      End
      Begin VB.OptionButton rbcOS 
         Caption         =   "Office"
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
         Left            =   180
         TabIndex        =   2
         Top             =   60
         Value           =   -1  'True
         Width           =   900
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
         Left            =   4380
         TabIndex        =   4
         Top             =   15
         Width           =   3795
      End
   End
   Begin VB.PictureBox pbcSalesperson 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   630
      Picture         =   "Budget.frx":8620
      ScaleHeight     =   3945
      ScaleWidth      =   8745
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   810
      Visible         =   0   'False
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
         Left            =   2160
         Picture         =   "Budget.frx":2DEAA
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   61
         Top             =   120
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
         Left            =   7500
         Picture         =   "Budget.frx":2E15C
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   60
         Top             =   120
         Width           =   270
      End
      Begin VB.Label lacSFrame 
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
         TabIndex        =   30
         Top             =   675
         Visible         =   0   'False
         Width           =   8760
      End
   End
   Begin VB.PictureBox pbcOffice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   210
      Picture         =   "Budget.frx":2E40E
      ScaleHeight     =   2775
      ScaleWidth      =   8745
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1035
      Width           =   8745
      Begin VB.PictureBox pbcLnWkArrow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2160
         Picture         =   "Budget.frx":48AA8
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   63
         Top             =   120
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
         Index           =   3
         Left            =   7500
         Picture         =   "Budget.frx":48D5A
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   62
         Top             =   120
         Width           =   270
      End
      Begin VB.Label lacOFrame 
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
         TabIndex        =   29
         Top             =   645
         Visible         =   0   'False
         Width           =   8760
      End
   End
   Begin VB.PictureBox plcBudget 
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
      Height          =   2880
      Left            =   165
      ScaleHeight     =   2820
      ScaleWidth      =   9045
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   975
      Width           =   9105
      Begin VB.VScrollBar vbcBudget 
         Height          =   2745
         LargeChange     =   11
         Left            =   8775
         TabIndex        =   32
         Top             =   15
         Width           =   240
      End
   End
   Begin VB.PictureBox pbcTotals 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   210
      Picture         =   "Budget.frx":4900C
      ScaleHeight     =   1035
      ScaleWidth      =   8745
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3915
      Width           =   8745
   End
   Begin VB.PictureBox plcTotals 
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
      Height          =   1170
      Left            =   165
      ScaleHeight     =   1110
      ScaleWidth      =   9060
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3870
      Width           =   9120
      Begin VB.VScrollBar vbcTotals 
         Height          =   1035
         LargeChange     =   3
         Left            =   8775
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   15
         Width           =   240
      End
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3735
      Top             =   210
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   165
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
      Left            =   225
      ScaleHeight     =   225
      ScaleWidth      =   3030
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3030
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
         Left            =   1980
         TabIndex        =   38
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   0
         Width           =   1170
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
      Left            =   6180
      ScaleHeight     =   225
      ScaleWidth      =   2820
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2820
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   0
         Value           =   -1  'True
         Width           =   990
      End
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
      Left            =   9045
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   9300
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4980
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
      Left            =   9180
      TabIndex        =   57
      Top             =   4830
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3675
      Picture         =   "Budget.frx":5304E
      Top             =   690
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8805
      Picture         =   "Budget.frx":53358
      Top             =   5355
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Budget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budget.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Budget.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Rate Card input screen code
Option Explicit
Option Compare Text
Dim smSalesOfficeCodeTag As String
'Office Areas
Dim tmOVCtrls(0 To 3)  As FIELDAREA
Dim imLBOVCtrls As Integer
Dim imOVBoxNo As Integer   'Current Rate Card Box
'Salesperson Areas
Dim tmSCtrls(0 To 1)  As FIELDAREA
Dim tmLBSCtrls As Integer
Dim imSBoxNo As Integer   'Current Rate Card Box
'Office Field Areas
Dim tmBDCtrls(0 To 6)  As FIELDAREA
Dim imLBBDCtrls As Integer
Dim tmWKCtrls(0 To 4)  As FIELDAREA
Dim imLBWKCtrls As Integer
Dim tmNWCtrls(0 To 6)  As FIELDAREA
Dim imLBNWCtrls As Integer
Dim imBDBoxNo As Integer   'Current Budget Box
Dim imBDRowNo As Integer      'Current row number in Program area (start at 0)
'Totals Field Areas
Dim tmTCtrls(0 To 6)  As FIELDAREA
Dim imlbTCtrls As Integer
'Grand Totals
Dim tmGTCtrls(0 To 5)  As FIELDAREA
Dim imLBGTCtrls As Integer
'Office
Dim hmBvf As Integer    'Rate Card file handle
Dim tmBvfSrchKey As BVFKEY0    'Rcf key record image
Dim imBvfRecLen As Integer        'Rcf record length
Dim smMnfName As String
Dim imMnfCode As Integer
Dim imYear As Integer
'Salesperson
Dim hmBsf As Integer    'Rate Card file handle
Dim tmBsfSrchKey As BSFKEY0    'Rcf key record image
Dim imBsfReclen As Integer        'Rcf record length
'Dim tmRec As LPOPREC
Dim smOVSave(0 To 1) As String  '1=Name
Dim imLBOVSave As Integer
'Dim imOVSave(1 To 2) As Integer '1=Year; 2=Direct/Split
Dim imOVSave(0 To 2) As Integer '1=Year; 2=Direct/Split
'Dim imSSave(1 To 1) As Integer  'Year
Dim imSSave(0 To 1) As Integer  'Year
Dim imLBSSave(0 To 1) As Integer
'Dim sgBDShow() As String  'Values shown in budget area (1=Name; 2-6=Dollars)
Dim smBDSave() As String  'Values saved (1=Name)
Dim imLBBDSave As Integer
Dim lmBDSave() As Long    'Value saved (1-5=Dollars)
Dim imBDSave() As Integer  'Values saved (1=tgBvfRec or tgBsfRec index)
                            '-1=Total line; 0=Name line
Dim smTShow() As String  'Values saved (1=Name; 2-6=Dollars)
Dim imLBTShow As Integer
Dim lmTSave() As Long    'Values (1-5=Dollars)
Dim imLBTSave As Integer
Dim imTSave() As Integer 'Values saves (1=Vehicle code; 2=Office code, 3=Index into lmTSave)
'Dim smGTShow(1 To 5) As String  'Values saved (1-5=Dollars)
Dim smGTShow(0 To 5) As String  'Values saved (1-5=Dollars)

'Dim lmGTSave(1 To 5) As Long
Dim lmGTSave(0 To 5) As Long


Dim imBDChg As Integer  'True=Vehicle or salesperon value changed; False=No changes
'Vehicle
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
Dim tmUserVeh() As BDUSERVEH
Dim imLBUserVeh As Integer
'Virtual Vehicle
Dim hmVsf As Integer        'Virtual Vehicle file handle
Dim tmVsf As VSF            'VSF record image
Dim imVsfRecLen As Integer  'VSF record length
'Budget Names
Dim hmMnf As Integer        'Multi-Name file handle
Dim tmMnf As MNF            'MNF record image
Dim imMnfRecLen As Integer  'MNSF record length
'Sales Office
Dim hmSof As Integer        'Sales Officee file handle
Dim tmSof As SOF            'SOF record image
Dim imSofRecLen As Integer  'SOF record length
Dim tmSaleOffice() As SALEOFFICE
Dim imLBSalesOffice As Integer
'Salesperson
Dim hmSlf As Integer        'Salesperson file handle
Dim tmSlf As SLF            'SLF record image
Dim tmSlfSrchKey As INTKEY0 'SLF key record image
Dim imSlfRecLen As Integer  'SLF record length
'Log Calendar
Dim hmLcf As Integer        'Log Calendar file handle
Dim tmLcf As LCF            'LCF record image
Dim imLcfRecLen As Integer  'LCF record length

Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim imSafRecLen As Integer

'Period (column) Information
Dim imPdYear As Integer
Dim imPdStartWk As Integer 'start week number
Dim imPdStartFltNo As Integer
Dim imBDStartYear As Integer
Dim imBDNoYears As Integer
'Dim tmPdGroups(1 To 4) As PDGROUPS
Dim tmPdGroups(0 To 4) As PDGROUPS
'Dim imHotSpot(1 To 4, 1 To 4) As Integer
Dim imHotSpot(0 To 4, 0 To 4) As Integer
Dim imInHotSpot As Integer
Dim imShowIndex As Integer
Dim imTypeIndex As Integer
Dim imFirstTime As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imInNew As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstTimeSelect As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imSettingValue As Integer   'True=Don't enable any box with change
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imBypassFocus As Integer
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imUpdateAllowed As Integer
Dim imIgnoreRightMove As Integer
Dim imButtonIndex As Integer
Dim imRetBranch As Integer
Dim lmNowDate As Long
Dim imNowYear As Integer
Dim imFirstActivate As Integer

Dim tmBudNameCode() As SORTCODE
Dim smBudNameCodeTag As String

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor


Const OSNAMEINDEX = 1          'Budget control/field
Const DOLLAR1INDEX = 2
Const DOLLAR2INDEX = 3
Const DOLLAR3INDEX = 4
Const DOLLAR4INDEX = 5
Const TOTALINDEX = 6
Const BDNAMEINDEX = 1
Const DIRECTINDEX = 2
Const YEARINDEX = 3
Const SYEARINDEX = 1
Const TNAMEINDEX = 1
Const TDOLLAR1INDEX = 2
Const TDOLLAR2INDEX = 3
Const TDOLLAR3INDEX = 4
Const TDOLLAR4INDEX = 5
Const TTOTALINDEX = 6
Const GTDOLLAR1INDEX = 1
Const GTDOLLAR2INDEX = 2
Const GTDOLLAR3INDEX = 3
Const GTDOLLAR4INDEX = 4
Const GTTOTALINDEX = 5
Const WK1INDEX = 1
Const WK2INDEX = 2
Const WK3INDEX = 3
Const WK4INDEX = 4
Const NWNAMEINDEX = 1
Const NW1INDEX = 2
Const NW2INDEX = 3
Const NW3INDEX = 4
Const NW4INDEX = 5
Const NWTOTALINDEX = 6
Private Sub cbcSelect_Change()
    Dim ilRet As Integer    'Return status
    Dim ilLen As Integer    'Length of current enter text
    Dim slStr As String     'Text entered
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim slNameYear As String
    Dim slYear As String
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
                If igBDView = 0 Then
                    slNameCode = tmBudNameCode(ilIndex - 1).sKey  'lbcBudget.List(ilIndex - 1)
                    ilLen = gParseItem(slNameCode, 1, "\", slNameYear)
                    ilLen = gParseItem(slNameYear, 2, "/", smMnfName)
                    ilLen = gParseItem(slNameYear, 1, "/", slYear)
                    slYear = gSubStr("9999", slYear)
                    ilLen = gParseItem(slNameCode, 2, "\", slCode)
                    imMnfCode = Val(slCode)
                    imYear = Val(slYear)
                    'Changed 2/3/98- This was causing Save to come back on after adding and saving a new budget
                    If Not mReadBvfRec(imMnfCode, imYear, imYear, True) Then    'False) Then
                        GoTo cbcSelectErr
                    End If
                Else
                    slYear = cbcSelect.List(ilIndex)
                    imYear = Val(slYear)
                    If Not mReadBsfRec(imYear, imYear) Then
                        GoTo cbcSelectErr
                    End If
                End If
                mInitBudgetCtrls
                igMode = 1    'Change
            Else
                If ilRet = 1 Then
                    cbcSelect.ListIndex = 0
                End If
                ilRet = 1   'Clear fields as no match name found
                igMode = 0    'New
            End If
            If igBDView = 0 Then
                pbcOffice.Cls
                pbcTotals.Cls
                pbcBudgetName.Cls
            Else
                pbcSalesperson.Cls
            End If
            If ilRet = 0 Then
                imSelectedIndex = cbcSelect.ListIndex
                mMoveRecToCtrl
                'mInitShow
            Else
                imSelectedIndex = 0
                mClearCtrlFields
            End If
            If igBDView = 0 Then
                pbcBudgetName_Paint
                pbcOffice_Paint
                pbcTotals_Paint
            Else
                pbcSalesperson_Paint
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
    'mPopulate
    'If imTerminate Then
    '    Exit Sub
    'End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If imFirstTime Then
        DoEvents
        imFirstTime = False
    End If
    If cbcSelect.ListCount <= 1 Then
        igMode = 0    'New
        imFirstTimeSelect = True
        pbcStartNew.SetFocus
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        'Remove auto select until faster
        'cbcSelect.ListIndex = 1 'Force to newest instead of [New]
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
Private Sub cmc12Mos_Click()
    If igBDView <> 0 Then
        Exit Sub
    End If
    sgBAName = cbcSelect.List(imSelectedIndex)
    If imBDRowNo >= 0 Then
        sgBvfVehName = Trim$(tgBvfRec(imBDSave(1, imBDRowNo)).sVehicle)
        sgBvfOffName = Trim$(tgBvfRec(imBDSave(1, imBDRowNo)).SOffice)
    Else
        sgBvfVehName = ""
        sgBvfOffName = ""
    End If
    Bud12Mo.Show vbModal
    'pbcOffice.Cls
    'pbcTotals.Cls
    'mGetShowPrices
    'pbcOffice_Paint
    'pbcTotals_Paint
    'imBDChg = True
    'mSetCommands
    mRefreshControls False
End Sub

Private Sub cmc12Mos_GotFocus()
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub

Private Sub cmcActuals_Click()
    If igBDView <> 0 Then
        Exit Sub
    End If
    sgBAName = cbcSelect.List(imSelectedIndex)
    If igBudgetType = 0 Then
        BudActB.Show vbModal
    Else
        BudActA.Show vbModal
    End If
    If igBDReturn = 1 Then
        'pbcOffice.Cls
        'pbcTotals.Cls
        'mGetShowPrices
        'pbcOffice_Paint
        'pbcTotals_Paint
        'imBDChg = True
        'mSetCommands
        mRefreshControls False
    Else
        mRefreshControls True
    End If
End Sub
Private Sub cmcActuals_GotFocus()
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub cmcAddVeh_Click()
    If igBDView <> 0 Then
        Exit Sub
    End If
    sgBAName = cbcSelect.List(imSelectedIndex)
    BudAdd.Show vbModal
    If igBDReturn = 1 Then
        Screen.MousePointer = vbHourglass
        If UBound(tgBvfRec) > 1 Then
            ArraySortTyp fnAV(tgBvfRec(), 1), UBound(tgBvfRec) - 1, 0, LenB(tgBvfRec(1)), 0, LenB(tgBvfRec(1).sKey), 0
        End If
        mMoveRecToCtrl
        'pbcOffice.Cls
        'pbcTotals.Cls
        'mGetShowPrices
        'pbcOffice_Paint
        'pbcTotals_Paint
        'imBDChg = True
        'mSetCommands
        mRefreshControls False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmcAddVeh_GotFocus()
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub

Private Sub cmcAdvt_Click()
    If igBDView <> 0 Then
        Exit Sub
    End If
    sgBAName = cbcSelect.List(imSelectedIndex)
    BudAdvt.Show vbModal
    If igBDReturn = 1 Then
        'pbcOffice.Cls
        'pbcTotals.Cls
        'mGetShowPrices
        'pbcOffice_Paint
        'pbcTotals_Paint
        'imBDChg = True
        'mSetCommands
        mRefreshControls False
    Else
        mRefreshControls True
    End If
End Sub
Private Sub cmcAdvt_GotFocus()
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDemo_Click()
    If igBDView <> 0 Then
        Exit Sub
    End If
    sgBAName = cbcSelect.List(imSelectedIndex)
    BudResch.Show vbModal
    If igBDReturn = 1 Then
        'pbcOffice.Cls
        'pbcTotals.Cls
        'mGetShowPrices
        'pbcOffice_Paint
        'pbcTotals_Paint
        'imBDChg = True
        'mSetCommands
        mRefreshControls False
    Else
        mRefreshControls True
    End If
End Sub
Private Sub cmcDemo_GotFocus()
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        If imBDBoxNo > 0 Then
            mBDEnableBox imBDBoxNo
        End If
        'If imSPBoxNo > 0 Then
        '    mSPEnableBox imSPBoxNo
        'End If
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcScale_Click()
    If igBDView <> 0 Then
        Exit Sub
    End If
    lgTotal = lmGTSave(GTTOTALINDEX)
    sgBAName = cbcSelect.List(imSelectedIndex)
    BudScale.Show vbModal
    If igBDReturn = 1 Then
        'pbcOffice.Cls
        'pbcTotals.Cls
        'mGetShowPrices
        'pbcOffice_Paint
        'pbcTotals_Paint
        'imBDChg = True
        'mSetCommands
        mRefreshControls False
    Else
        mRefreshControls True
    End If
End Sub
Private Sub cmcScale_GotFocus()
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub cmcTrend_Click()
    If igBDView <> 0 Then
        Exit Sub
    End If
    sgBAName = cbcSelect.List(imSelectedIndex)
    BudTrend.Show vbModal
    If igBDReturn = 1 Then
        'pbcOffice.Cls
        'pbcTotals.Cls
        'mGetShowPrices
        'pbcOffice_Paint
        'pbcTotals_Paint
        'imBDChg = True
        'mSetCommands
        mRefreshControls False
    Else
        mRefreshControls True
    End If
End Sub
Private Sub cmcTrend_GotFocus()
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub cmcUndo_Click()
    Dim ilIndex As Integer
    ilIndex = imSelectedIndex
    If (imSelectedIndex <> 0) And (igMode <> 0) Then 'Not New selected
        If ilIndex > 0 Then
            Screen.MousePointer = vbHourglass
            If igBDView = 0 Then
                pbcOffice.Cls
                pbcTotals.Cls
                If Not mReadBvfRec(imMnfCode, imYear, imYear, True) Then    'False) Then
                    GoTo cmcUndoErr
                    Exit Sub
                End If
            Else
                pbcSalesperson.Cls
                If Not mReadBsfRec(imYear, imYear) Then
                    GoTo cmcUndoErr
                    Exit Sub
                End If
            End If
            mInitBudgetCtrls
            mMoveRecToCtrl
            If igBDView = 0 Then
                pbcOffice_Paint
                pbcTotals_Paint
            Else
                pbcSalesperson_Paint
            End If
            imBDChg = False
            mSetCommands
            Screen.MousePointer = vbDefault
            imBDBoxNo = -1
            imBDRowNo = -1
            'imSPBoxNo = -1 'Initialize current Box to N/A
            'pbcSTab.SetFocus
            Exit Sub
        End If
    End If
    'Clear pictures as paint event will only override what exist
    'Paint event will be generated by change event via setting ListIndex = 0
    imBDChg = False
    imBDBoxNo = -1
    imOVBoxNo = -1 'Initialize current Box to N/A
    imSBoxNo = -1 'Initialize current Box to N/A
    imSelectedIndex = 0
    cbcSelect.RemoveItem 1
    pbcOffice.Cls
    pbcTotals.Cls
    pbcSalesperson.Cls
    cbcSelect.ListIndex = 0
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
    Exit Sub
cmcUndoErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub cmcUndo_GotFocus()
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
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
        mBDEnableBox imBDBoxNo
        Exit Sub
    End If
    imBDBoxNo = -1    'Have to be after mSaveRecChg as it test imBoxNo = 1
    imFirstTimeSelect = True
    imBDChg = False
    'ilRowNo = UBound(tmRifRec)
    'For ilLoop = imLBBDCtrls To UBound(tmBDCtrls) Step 1
    '    slStr = ""
    '    gSetShow pbcBudget, slStr, tmBDCtrls(ilLoop)
    '    smShow(ilLoop, ilRowNo) = tmBDCtrls(ilLoop).sShow
    'Next ilLoop
    'If igBDView = 0 Then
    '    pbcOffice_Paint
    '    pbcTotals_Paint
    'Else
    '    pbcSalesperson_Paint
    'End If
    mSetCommands
    'pbcSTab.SetFocus
End Sub
Private Sub cmcUpdate_GotFocus()
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
    gCtrlGotFocus cmcUpdate
End Sub
Private Sub cmcUpdate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus edcDropDown
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim slComp As String

    Select Case imBDBoxNo
        Case DOLLAR1INDEX To DOLLAR4INDEX
            ''ilPos = InStr(edcDropDown.SelText, ".")
            ''If ilPos = 0 Then
            ''    ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
            ''    If ilPos > 0 Then
            ''        If KeyAscii = KEYDECPOINT Then
            ''            Beep
            ''            KeyAscii = 0
            ''            Exit Sub
            ''        End If
            ''    End If
            ''End If
            ''Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            ''If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
            'If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            '    Beep
            '    KeyAscii = 0
            '    Exit Sub
            'End If
            If (KeyAscii = KEYNEG) And ((Len(edcDropDown.Text) = 0) Or (Len(edcDropDown.Text) = edcDropDown.SelLength)) Then
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYNEG) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            Else
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            slComp = "99999999"
            'If gCompNumberStr(slStr, slComp) > 0 Then
            If gCompAbsNumberStr(slStr, slComp) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcOVDropDown_Change()
    If igBDView = 0 Then  'Office
        Select Case imOVBoxNo
            Case BDNAMEINDEX
            Case DIRECTINDEX
            Case YEARINDEX
        End Select
    Else
        Select Case imSBoxNo
            Case SYEARINDEX
        End Select
    End If
    imLbcArrowSetting = False
End Sub
Private Sub edcOVDropDown_GotFocus()
    gCtrlGotFocus edcOVDropDown
End Sub
Private Sub edcOVDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcOVDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcOVDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub Form_Activate()
    If imInNew Then
        Exit Sub
    End If
    Me.KeyPreview = True  'To get Alt J and Alt L keys
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        'gShowBranner imUpdateAllowed
        Me.KeyPreview = True  'To get Alt J and Alt L keys
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(BUDGETSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcOffice.Enabled = False
        pbcSalesperson.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        pbcOSSTab.Enabled = False
        pbcOSTab.Enabled = False
        imUpdateAllowed = False
    Else
        pbcOffice.Enabled = True
        pbcSalesperson.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        pbcOSSTab.Enabled = True
        pbcOSTab.Enabled = True
        imUpdateAllowed = True
    End If
    'gShowBranner imUpdateAllowed
    mSetCommands
    'DoEvents    'Process events so pending keys are not sent to this
                'form when keypreview turn on
    'If Not imTerminate Then
    '    Budget.KeyPreview = True   'To get Alt J and Alt L keys
    'End If
    Me.KeyPreview = True
    Me.ZOrder 0 'Send to front
    Budget.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Deactivate()
    'Budget.KeyPreview = False
    Me.KeyPreview = False
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'If ((Shift And vbAltMask) > 0) And (KeyCode = 74) Then    'J=74
    '    Budget.KeyPreview = False
    '    Traffic!gpcBasicWnd.Value = True   'Button up and unload
    'End If
    'If ((Shift And vbAltMask) > 0) And (KeyCode = 76) Then    'L=76
    '    Budget.KeyPreview = False
    '    Traffic!gpcAuxWnd.Value = True   'Button up and unload
    'End If
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcSelect.Enabled) And (imBDBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBDBoxNo > 0 Then
            mBDEnableBox imBDBoxNo
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
    lgPercentAdjH = 95
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        'Only expand First Column
        'Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    tmcStart.Enabled = True
    'If imTerminate Then
    '    cmcCancel_Click
    'End If
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
        mBDSetShow imBDBoxNo
        imBDBoxNo = -1
        pbcArrow.Visible = False
        lacOFrame.Visible = False
        lacSFrame.Visible = False
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            If imBDBoxNo <> -1 Then
                mBDEnableBox imBDBoxNo
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    
    Erase tgSalesOfficeCode
    btrExtClear hmLcf   'Clear any previous extend operation
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    btrExtClear hmSlf   'Clear any previous extend operation
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    btrExtClear hmSof   'Clear any previous extend operation
    ilRet = btrClose(hmSof)
    btrDestroy hmSof
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    btrExtClear hmVsf   'Clear any previous extend operation
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    btrExtClear hmBsf   'Clear any previous extend operation
    ilRet = btrClose(hmBsf)
    btrDestroy hmBsf
    btrExtClear hmBvf   'Clear any previous extend operation
    ilRet = btrClose(hmBvf)
    btrDestroy hmBvf
    Erase sgBDShow
    Erase smBDSave
    Erase lmBDSave
    Erase imBDSave
    Erase smTShow
    Erase lmTSave
    Erase tgBvfRec
    Erase tgBsfRec
    Erase tmUserVeh
    Erase tmSaleOffice

    Erase tgBudUserVehicle
    igJobShowing(BUDGETSJOB) = False
        
    Set Budget = Nothing   'Remove data segment

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
    'Dim tlRif As RIF
    If (imBDRowNo < 1) Then
        Exit Sub
    End If
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
'    lacRCFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        If igBDView = 0 Then
            lacOFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        Else
            lacSFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        End If
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        If igBDView = 0 Then
            lacOFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        Else
            lacSFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBDEnableBox                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mBDEnableBox(ilBoxNo As Integer)
'
'   mBDEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBBDCtrls) Or (ilBoxNo > UBound(tmBDCtrls)) Then
        Exit Sub
    End If

    If (imBDRowNo < vbcBudget.Value) Or (imBDRowNo >= vbcBudget.Value + vbcBudget.LargeChange + 1) Then
        'mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacSFrame.Visible = False
        lacOFrame.Visible = False
        Exit Sub
    End If
    If igBDView = 0 Then
        lacOFrame.Move 0, tmBDCtrls(OSNAMEINDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15) - 30
        lacOFrame.Visible = True
    Else
        lacSFrame.Move 0, tmBDCtrls(OSNAMEINDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15) - 30
        lacSFrame.Visible = True
    End If
    pbcArrow.Move pbcArrow.Left, plcBudget.Top + tmBDCtrls(OSNAMEINDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case OSNAMEINDEX 'Vehicle
        Case DOLLAR1INDEX
            edcDropDown.Width = tmBDCtrls(DOLLAR1INDEX).fBoxW
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcOffice, edcDropDown, tmBDCtrls(DOLLAR1INDEX).fBoxX, tmBDCtrls(DOLLAR1INDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15)
            'edcDropDown.Text = smBDSave(DOLLAR1INDEX, imBDRowNo)
            edcDropDown.Text = Trim$(Str$(lmBDSave(DOLLAR1INDEX - DOLLAR1INDEX + 1, imBDRowNo)))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case DOLLAR2INDEX
            edcDropDown.Width = tmBDCtrls(DOLLAR2INDEX).fBoxW
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcOffice, edcDropDown, tmBDCtrls(DOLLAR2INDEX).fBoxX, tmBDCtrls(DOLLAR2INDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15)
            'edcDropDown.Text = smBDSave(DOLLAR2INDEX, imBDRowNo)
            edcDropDown.Text = Trim$(Str$(lmBDSave(DOLLAR2INDEX - DOLLAR1INDEX + 1, imBDRowNo)))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case DOLLAR3INDEX
            edcDropDown.Width = tmBDCtrls(DOLLAR3INDEX).fBoxW
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcOffice, edcDropDown, tmBDCtrls(DOLLAR3INDEX).fBoxX, tmBDCtrls(DOLLAR3INDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15)
            'edcDropDown.Text = smBDSave(DOLLAR3INDEX, imBDRowNo)
            edcDropDown.Text = Trim$(Str$(lmBDSave(DOLLAR3INDEX - DOLLAR1INDEX + 1, imBDRowNo)))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case DOLLAR4INDEX
            edcDropDown.Width = tmBDCtrls(DOLLAR4INDEX).fBoxW
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcOffice, edcDropDown, tmBDCtrls(DOLLAR4INDEX).fBoxX, tmBDCtrls(DOLLAR4INDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15)
            'edcDropDown.Text = smBDSave(DOLLAR4INDEX, imBDRowNo)
            edcDropDown.Text = Trim$(Str$(lmBDSave(DOLLAR4INDEX - DOLLAR1INDEX + 1, imBDRowNo)))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBDSetShow                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mBDSetShow(ilBoxNo As Integer)
'
'   mBDSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim ilNegIndex As Integer
    Dim ilNeg As Integer
    Dim ilTIndex As Integer
    Dim slDollar As String
    Dim llDollar As Long
    Dim llTDollar As Long
    Dim llOldDollar As Long
    Dim ilBvf As Integer
    Dim ilStIndex As Integer
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim llAvgDollar As Long
    Dim llTAvgDollar As Long
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
    If (ilBoxNo < imLBBDCtrls) Or (ilBoxNo > UBound(tmBDCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case DOLLAR1INDEX, DOLLAR2INDEX, DOLLAR3INDEX, DOLLAR4INDEX
            edcDropDown.Visible = False
            'slStr = edcDropDown.Text
            'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            'gSetShow pbcOffice, slStr, tmBDCtrls(ilBoxNo)
            'sgBDShow(ilBoxNo, imBDRowNo) = tmBDCtrls(ilBoxNo).sShow
            slDollar = edcDropDown.Text
            llDollar = Val(slDollar)
            If lmBDSave(ilBoxNo - 1, imBDRowNo) <> llDollar Then
                llOldDollar = lmBDSave(ilBoxNo - 1, imBDRowNo)
                'Recompute total and set weeks
                If tmPdGroups(1).iYear = tmPdGroups(ilBoxNo - DOLLAR1INDEX + 1).iYear Then
                    'Remove old values
                    'smBDSave(TOTALINDEX, imBDRowNo) = gSubStr(smBDSave(TOTALINDEX, imBDRowNo), smBDSave(ilBoxNo, imBDRowNo))
                    lmBDSave(TOTALINDEX - 1, imBDRowNo) = lmBDSave(TOTALINDEX - 1, imBDRowNo) - llOldDollar 'lmBDSave(ilBoxNo - 1, imBDRowNo)
                    If igBDView = 0 Then
                        If rbcSort(0).Value Or rbcSort(1).Value Then
                            lmGTSave(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) = lmGTSave(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) - llOldDollar 'lmBDSave(ilBoxNo - 1, imBDRowNo)
                            lmGTSave(GTTOTALINDEX) = lmGTSave(GTTOTALINDEX) - llOldDollar   'lmBDSave(ilBoxNo - 1, imBDRowNo)
                        Else
                            ilTIndex = 1
                            For ilLoop = imBDRowNo - 1 To LBound(smBDSave, 2) Step -1
                                If imBDSave(1, ilLoop) = 0 Then
                                    Exit For
                                Else
                                    ilTIndex = ilTIndex + 1
                                End If
                            Next ilLoop
                            lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilTIndex) = lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilTIndex) - llOldDollar 'lmBDSave(ilBoxNo - 1, imBDRowNo)
                            lmTSave(TTOTALINDEX - 1, ilTIndex) = lmTSave(TTOTALINDEX - 1, ilTIndex) - llOldDollar   'lmBDSave(ilBoxNo - 1, imBDRowNo)

                            ilNegIndex = -1
                            If igBDView = 0 Then
                                For ilNeg = imBDRowNo + 1 To UBound(smBDSave, 2) - 1 Step 1
                                    If imBDSave(1, ilNeg) = -1 Then
                                        ilNegIndex = ilNeg
                                        Exit For
                                    End If
                                Next ilNeg
                            End If
                            If ilNegIndex > 0 Then
                                'smBDSave(ilBoxNo, ilNegIndex) = gSubStr(smBDSave(ilBoxNo, ilNegIndex), smBDSave(ilBoxNo, imBDRowNo))
                                lmBDSave(ilBoxNo - 1, ilNegIndex) = lmBDSave(ilBoxNo - 1, ilNegIndex) - llOldDollar 'lmBDSave(ilBoxNo - 1, imBDRowNo)
                                lmBDSave(TOTALINDEX - 1, ilNegIndex) = lmBDSave(TOTALINDEX - 1, ilNegIndex) - llOldDollar   'lmBDSave(ilBoxNo - 1, imBDRowNo)

                            End If
                            'smGTShow(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) = gSubStr(smGTShow(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX), smBDSave(ilBoxNo, imBDRowNo))
                            lmGTSave(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) = lmGTSave(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) - llOldDollar 'lmBDSave(ilBoxNo - 1, imBDRowNo)
                            'smGTShow(GTTOTALINDEX) = gSubStr(smGTShow(GTTOTALINDEX), smBDSave(ilBoxNo, imBDRowNo))
                            lmGTSave(GTTOTALINDEX) = lmGTSave(GTTOTALINDEX) - llOldDollar   'lmBDSave(ilBoxNo - 1, imBDRowNo)
                        End If
                    End If
                    'Set new values into fields
                    lmBDSave(ilBoxNo - 1, imBDRowNo) = llDollar
                    lmBDSave(TOTALINDEX - 1, imBDRowNo) = lmBDSave(TOTALINDEX - 1, imBDRowNo) + llDollar
                    slStr = Trim$(Str$(lmBDSave(ilBoxNo - 1, imBDRowNo)))
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    gSetShow pbcOffice, slStr, tmBDCtrls(ilBoxNo)
                    sgBDShow(ilBoxNo, imBDRowNo) = tmBDCtrls(ilBoxNo).sShow

                    slStr = Trim$(Str$(lmBDSave(TOTALINDEX - 1, imBDRowNo)))
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    gSetShow pbcOffice, slStr, tmBDCtrls(ilBoxNo)
                    sgBDShow(TOTALINDEX, imBDRowNo) = tmBDCtrls(ilBoxNo).sShow
                    If igBDView = 0 Then
                        If rbcSort(0).Value Or rbcSort(1).Value Then
                            ilGroup = ilBoxNo - DOLLAR1INDEX + 1
                            ilBvf = imBDSave(1, imBDRowNo)
                            llTAvgDollar = 0
                            If tgBvfRec(ilBvf).tBvf.iYear = tmPdGroups(ilGroup).iYear Then
                                If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                                    For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                                        llTDollar = 0
                                        ilStIndex = 1
                                        ilBvf = imBDSave(1, imBDRowNo)
                                        Do
                                            llTDollar = llTDollar + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                            lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilStIndex) = lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilStIndex) - tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                            lmTSave(TTOTALINDEX - 1, ilStIndex) = lmTSave(TTOTALINDEX - 1, ilStIndex) - tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                            If rbcSort(0).Value Then
                                                If tgBvfRec(ilBvf).tBvf.iVefCode <> tgBvfRec(ilBvf + 1).tBvf.iVefCode Then
                                                    Exit Do
                                                Else
                                                    ilBvf = ilBvf + 1
                                                    ilStIndex = ilStIndex + 1
                                                End If
                                            Else
                                                If tgBvfRec(ilBvf).tBvf.iSofCode <> tgBvfRec(ilBvf + 1).tBvf.iSofCode Then
                                                    Exit Do
                                                Else
                                                    ilBvf = ilBvf + 1
                                                    ilStIndex = ilStIndex + 1
                                                End If
                                            End If
                                            If ilBvf = UBound(tgBvfRec) Then
                                                ilStIndex = 0
                                                Exit Do
                                            End If
                                        Loop
                                        ilBvf = imBDSave(1, imBDRowNo)
                                        ilStIndex = 1
                                        Do
                                            llAvgDollar = llDollar / tmPdGroups(ilGroup).iTrueNoWks
                                            If llTDollar = 0 Then
                                                tgBvfRec(ilBvf).tBvf.lGross(ilWk) = llAvgDollar / (UBound(lmTSave, 2) - 1)
                                                If (tgBvfRec(ilBvf).tBvf.lGross(ilWk) <= 0) And (llAvgDollar > 0) Then
                                                    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = 1
                                                ElseIf (tgBvfRec(ilBvf).tBvf.lGross(ilWk) >= 0) And (llAvgDollar < 0) Then
                                                    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = -1
                                                End If
                                            Else
                                                If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                                                    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = CLng((CDbl(llAvgDollar) * CDbl(tgBvfRec(ilBvf).tBvf.lGross(ilWk))) / CDbl(llTDollar))
                                                    If (tgBvfRec(ilBvf).tBvf.lGross(ilWk) <= 0) And (llAvgDollar > 0) Then
                                                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = 1
                                                    ElseIf (tgBvfRec(ilBvf).tBvf.lGross(ilWk) >= 0) And (llAvgDollar < 0) Then
                                                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = -1
                                                    End If
                                                End If
                                            End If
                                            lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilStIndex) = lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilStIndex) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                            lmTSave(TTOTALINDEX - 1, ilStIndex) = lmTSave(TTOTALINDEX - 1, ilStIndex) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                            llTAvgDollar = llTAvgDollar + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                            If rbcSort(0).Value Then
                                                If tgBvfRec(ilBvf).tBvf.iVefCode <> tgBvfRec(ilBvf + 1).tBvf.iVefCode Then
                                                    Exit Do
                                                Else
                                                    ilBvf = ilBvf + 1
                                                    ilStIndex = ilStIndex + 1
                                                End If
                                            Else
                                                If tgBvfRec(ilBvf).tBvf.iSofCode <> tgBvfRec(ilBvf + 1).tBvf.iSofCode Then
                                                    Exit Do
                                                Else
                                                    ilBvf = ilBvf + 1
                                                    ilStIndex = ilStIndex + 1
                                                End If
                                            End If
                                            If ilBvf = UBound(tgBvfRec) Then
                                                Exit Do
                                            End If
                                        Loop
                                        ilBvf = imBDSave(1, imBDRowNo)
                                    Next ilWk
                                End If
                            End If
                            If llTAvgDollar <> llDollar Then
                                ilGroup = ilBoxNo - DOLLAR1INDEX + 1
                                Do
                                    ilBvf = imBDSave(1, imBDRowNo)
                                    ilStIndex = 1
                                    If tgBvfRec(ilBvf).tBvf.iYear = tmPdGroups(ilGroup).iYear Then
                                        If (tmPdGroups(ilGroup).iTrueNoWks > 0) Then
                                            For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                                                Do
                                                    If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                                                        If llTAvgDollar > llDollar Then
                                                            tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) - 1
                                                            lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilStIndex) = lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilStIndex) - 1
                                                            lmTSave(TTOTALINDEX - 1, ilStIndex) = lmTSave(TTOTALINDEX - 1, ilStIndex) - 1
                                                            llTAvgDollar = llTAvgDollar - 1
                                                        ElseIf llTAvgDollar < llDollar Then
                                                            tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) + 1
                                                            lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilStIndex) = lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilStIndex) + 1
                                                            lmTSave(TTOTALINDEX - 1, ilStIndex) = lmTSave(TTOTALINDEX - 1, ilStIndex) + 1
                                                            llTAvgDollar = llTAvgDollar + 1
                                                        End If
                                                        If llTAvgDollar = llDollar Then
                                                            Exit For
                                                        End If
                                                    End If
                                                    If rbcSort(0).Value Then
                                                        If tgBvfRec(ilBvf).tBvf.iVefCode <> tgBvfRec(ilBvf + 1).tBvf.iVefCode Then
                                                            ilBvf = imBDSave(1, imBDRowNo)
                                                            ilStIndex = 1
                                                            Exit Do
                                                        Else
                                                            ilBvf = ilBvf + 1
                                                            ilStIndex = ilStIndex + 1
                                                        End If
                                                    Else
                                                        If tgBvfRec(ilBvf).tBvf.iSofCode <> tgBvfRec(ilBvf + 1).tBvf.iSofCode Then
                                                            ilBvf = imBDSave(1, imBDRowNo)
                                                            ilStIndex = 1
                                                            Exit Do
                                                        Else
                                                            ilBvf = ilBvf + 1
                                                            ilStIndex = ilStIndex + 1
                                                        End If
                                                    End If
                                                    If ilBvf = UBound(tgBvfRec) Then
                                                        ilBvf = imBDSave(1, imBDRowNo)
                                                        ilStIndex = 1
                                                        Exit Do
                                                    End If
                                                Loop
                                            Next ilWk
                                        End If
                                    End If
                                Loop While llTAvgDollar <> llDollar
                            End If
                            'llTDollar = 0
                            'For ilLoop = LBound(lmTSave, 2) To UBound(lmTSave, 2) - 1 Step 1
                            '    llTDollar = llTDollar + lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilLoop)
                            'Next ilLoop
                            '
                            'If llTDollar > llDollar Then
                            '    ilLoop = LBound(lmTSave, 2)
                            '    Do
                            '        If lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilLoop) > 0 Then
                            '            lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilLoop) = lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilLoop) - 1
                            '            lmTSave(TTOTALINDEX - 1, ilLoop) = lmTSave(TTOTALINDEX - 1, ilLoop) - 1
                            '            llTDollar = llTDollar - 1
                            '            If llTDollar <= llDollar Then
                            '                Exit Do
                            '            End If
                            '        End If
                            '        ilLoop = ilLoop + 1
                            '        If ilLoop >= UBound(lmTSave, 2) Then
                            '            ilLoop = LBound(lmTSave, 2)
                            '        End If
                            '    Loop
                            'ElseIf llTDollar < llDollar Then
                            '    ilLoop = LBound(lmTSave, 2)
                            '    Do
                            '        If lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilLoop) > 0 Then
                            '            lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilLoop) = lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilLoop) + 1
                            '            lmTSave(TTOTALINDEX - 1, ilLoop) = lmTSave(TTOTALINDEX - 1, ilLoop) + 1
                            '            llTDollar = llTDollar + 1
                            '            If llTDollar <= llDollar Then
                            '                Exit Do
                            '            End If
                            '        End If
                            '        ilLoop = ilLoop + 1
                            '        If ilLoop >= UBound(lmTSave, 2) Then
                            '            ilLoop = LBound(lmTSave, 2)
                            '        End If
                            '    Loop
                            'End If
                            'For ilLoop = LBound(lmTSave, 2) To UBound(lmTSave, 2) - 1 Step 1
                            For ilLoop = LBound(lmTSave, 2) + imLBTSave To UBound(lmTSave, 2) - 1 Step 1
                                slStr = Trim$(Str$(lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilLoop)))
                                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                                gSetShow pbcTotals, slStr, tmTCtrls(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX)
                                smTShow(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX, ilLoop) = tmTCtrls(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX).sShow

                                slStr = Trim$(Str$(lmTSave(TTOTALINDEX - 1, ilLoop)))
                                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                                gSetShow pbcOffice, slStr, tmTCtrls(TTOTALINDEX)
                                smTShow(TTOTALINDEX, ilLoop) = tmTCtrls(TTOTALINDEX).sShow
                            Next ilLoop
                        Else
                            'smTShow(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX, ilTIndex) = gAddStr(smTShow(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX, ilTIndex), smBDSave(ilBoxNo, imBDRowNo))
                            lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilTIndex) = lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilTIndex) + llDollar
                            slStr = Trim$(Str$(lmTSave(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX - 1, ilTIndex)))
                            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                            gSetShow pbcTotals, slStr, tmTCtrls(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX)
                            smTShow(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX, ilTIndex) = tmTCtrls(ilBoxNo - DOLLAR1INDEX + TDOLLAR1INDEX).sShow
                            'smTShow(TTOTALINDEX, ilTIndex) = gAddStr(smTShow(TTOTALINDEX, ilTIndex), smBDSave(ilBoxNo, imBDRowNo))
                            lmTSave(TTOTALINDEX - 1, ilTIndex) = lmTSave(TTOTALINDEX - 1, ilTIndex) + llDollar
                            slStr = Trim$(Str$(lmTSave(TTOTALINDEX - 1, ilTIndex)))
                            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                            gSetShow pbcOffice, slStr, tmTCtrls(TTOTALINDEX)
                            smTShow(TTOTALINDEX, ilTIndex) = tmTCtrls(TTOTALINDEX).sShow

                            If ilNegIndex > 0 Then
                                'smBDSave(ilBoxNo, ilNegIndex) = gAddStr(smBDSave(ilBoxNo, ilNegIndex), smBDSave(ilBoxNo, imBDRowNo))
                                lmBDSave(ilBoxNo - 1, ilNegIndex) = lmBDSave(ilBoxNo - 1, ilNegIndex) + llDollar
                                slStr = Trim$(Str$(lmBDSave(ilBoxNo - 1, ilNegIndex)))
                                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                                gSetShow pbcOffice, slStr, tmBDCtrls(ilBoxNo)
                                sgBDShow(ilBoxNo, ilNegIndex) = tmBDCtrls(ilBoxNo).sShow
                                lmBDSave(TOTALINDEX - 1, ilNegIndex) = lmBDSave(TOTALINDEX - 1, ilNegIndex) + llDollar
                                slStr = Trim$(Str$(lmBDSave(TOTALINDEX - 1, ilNegIndex)))
                                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                                gSetShow pbcOffice, slStr, tmBDCtrls(TOTALINDEX)
                                sgBDShow(TOTALINDEX, ilNegIndex) = tmBDCtrls(TOTALINDEX).sShow
                            End If
                        End If
                        'smGTShow(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) = gAddStr(smGTShow(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX), smBDSave(ilBoxNo, imBDRowNo))
                        lmGTSave(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) = lmGTSave(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) + llDollar
                        slStr = Trim$(Str$(lmGTSave(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX)))
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcTotals, slStr, tmGTCtrls(1)
                        smGTShow(ilBoxNo - DOLLAR1INDEX + GTDOLLAR1INDEX) = tmGTCtrls(1).sShow
                        'smGTShow(GTTOTALINDEX) = gAddStr(smGTShow(GTTOTALINDEX), smBDSave(ilBoxNo, imBDRowNo))
                        lmGTSave(GTTOTALINDEX) = lmGTSave(GTTOTALINDEX) + llDollar
                        slStr = Trim$(Str$(lmGTSave(GTTOTALINDEX)))
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcTotals, slStr, tmGTCtrls(GTTOTALINDEX)
                        smGTShow(GTTOTALINDEX) = tmGTCtrls(GTTOTALINDEX).sShow
                    End If
                End If
                slStr = edcDropDown.Text
                mSetPrice ilBoxNo - DOLLAR1INDEX + 1, imBDRowNo, slStr
                imBDChg = True
                If igBDView = 0 Then
                    pbcOffice.Cls
                    pbcTotals.Cls
                    pbcOffice_Paint
                    pbcTotals_Paint
                Else
                    pbcSalesperson.Cls
                    pbcSalesperson_Paint
                End If
            End If
    End Select
    mSetCommands
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

    'ReDim tgBvfRec(1 To 1) As BVFREC
    ReDim tgBvfRec(0 To 1) As BVFREC
    'ReDim tgBsfRec(1 To 1) As BSFREC
    ReDim tgBsfRec(0 To 1) As BSFREC
    
    If igBDView = 0 Then
        For ilLoop = imLBOVCtrls To UBound(tmOVCtrls) Step 1
            tmOVCtrls(ilLoop).iChg = False
        Next ilLoop
    Else
        For ilLoop = tmLBSCtrls To UBound(tmSCtrls) Step 1
            tmSCtrls(ilLoop).iChg = False
        Next ilLoop
    End If
    imBDChg = False
    mInitBudgetCtrls
    'ReDim sgBDShow(1 To TOTALINDEX, 1 To 1) As String * 40 'Values shown in program area
    ReDim sgBDShow(0 To TOTALINDEX, 0 To 1) As String * 40 'Values shown in program area
    mInitBDShow
    'ReDim smBDSave(1 To TOTALINDEX, 1 To 1) As String 'Values saved (program name) in program area
    'ReDim smBDSave(1 To 1, 1 To 1) As String 'Values saved (program name) in program area
    ReDim smBDSave(0 To 1, 0 To 1) As String 'Values saved (program name) in program area
    'ReDim lmBDSave(1 To TOTALINDEX - 1, 1 To 1) As Long 'Values saved (program name) in program area
    ReDim lmBDSave(0 To TOTALINDEX - 1, 0 To 1) As Long 'Values saved (program name) in program area
    'ReDim imBDSave(1 To 1, 1 To 1) As Integer 'Values saved (program name) in program area
    ReDim imBDSave(0 To 1, 0 To 1) As Integer 'Values saved (program name) in program area
    'ReDim smTShow(1 To TTOTALINDEX, 1 To 1) As String 'Values shown in program area
    ReDim smTShow(0 To TTOTALINDEX, 0 To 1) As String 'Values shown in program area
    'ReDim lmTSave(1 To TTOTALINDEX - 1, 1 To 1) As Long
    ReDim lmTSave(0 To TTOTALINDEX - 1, 0 To 1) As Long
    'ReDim imTSave(1 To 3, 1 To 1) As Integer
    ReDim imTSave(0 To 3, 0 To 1) As Integer
    For ilLoop = 1 To 5 Step 1
        smGTShow(ilLoop) = ""
        lmGTSave(ilLoop) = 0
    Next ilLoop
    imSettingValue = True
    If igBDView = 0 Then
        vbcBudget.LargeChange = 11
    Else
        vbcBudget.LargeChange = 17
    End If
    imSettingValue = True
    vbcBudget.Min = LBound(sgBDShow) + igLBBDShow
    vbcBudget.Min = igLBBDShow
    imSettingValue = True
    If UBound(sgBDShow, 2) - 1 <= vbcBudget.LargeChange + 1 Then ' + 1 Then
        vbcBudget.Max = LBound(sgBDShow, 2)
    Else
        vbcBudget.Max = UBound(sgBDShow, 2) - vbcBudget.LargeChange
    End If
    imSettingValue = True
    vbcBudget.Value = vbcBudget.Min
    imSettingValue = True
    'vbcTotals.Min = LBound(smTShow)
    vbcTotals.Min = imLBTShow
    imSettingValue = True
    If UBound(smTShow, 2) - 1 <= vbcTotals.LargeChange + 1 Then ' + 1 Then
        vbcTotals.Max = LBound(smTShow, 2) + imLBTShow
    Else
        vbcTotals.Max = UBound(smTShow, 2) - vbcTotals.LargeChange
    End If
    imSettingValue = True
    vbcTotals.Value = vbcTotals.Min
    imSettingValue = False
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
        If tgSpf.sRUseCorpCal = "Y" Then
            slDate = ""
            For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                If tgMCof(ilLoop).iYear = ilYear Then
                    slDate = Trim$(Str$(tgMCof(ilLoop).iStartMnthNo)) & "/15/" & Trim$(Str$(ilYear - 1))
                    slDate = gObtainYearStartDate(5, slDate)
                    Exit For
                End If
            Next ilLoop
            If slDate = "" Then
                MsgBox "Corporate Year Missing for" & Str$(ilYear), vbOKOnly + vbExclamation, "Budget"
                slDate = "1/15/" & Trim$(Str$(ilYear))
            End If
        Else
            slDate = "1/15/" & Trim$(Str$(ilYear))
        End If
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
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    Dim slWkEnd As String
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilYearOk As Integer
    Dim ilWkNo As Integer
    Dim ilWkCount As Integer
    'ReDim ilStartWk(1 To 12) As Integer
    ReDim ilStartWk(0 To 12) As Integer
    'ReDim ilNoWks(1 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer
    Dim ilLBNoWks As Integer
    Dim slFontName As String
    Dim flFontSize As Single
    If igBDView = 0 Then
        If UBound(tgBvfRec) <= 1 Then
            'For ilIndex = LBound(tmPdGroups) To UBound(tmPdGroups) Step 1
            For ilIndex = LBound(tmPdGroups) + 1 To UBound(tmPdGroups) Step 1
                tmPdGroups(ilIndex).iStartWkNo = -1
                tmPdGroups(ilIndex).iNoWks = 0
                tmPdGroups(ilIndex).iTrueNoWks = 0
                tmPdGroups(ilIndex).iFltNo = 0
                tmPdGroups(ilIndex).sStartDate = ""
                tmPdGroups(ilIndex).sEndDate = ""
                gSetShow pbcOffice, "", tmWKCtrls(ilIndex)
                gSetShow pbcOffice, "", tmNWCtrls(ilIndex + 1)
            Next ilIndex
            Exit Sub
        End If
    Else
        If UBound(tgBsfRec) <= 1 Then
            'For ilIndex = LBound(tmPdGroups) To UBound(tmPdGroups) Step 1
            For ilIndex = LBound(tmPdGroups) + 1 To UBound(tmPdGroups) Step 1
                tmPdGroups(ilIndex).iStartWkNo = -1
                tmPdGroups(ilIndex).iNoWks = 0
                tmPdGroups(ilIndex).iTrueNoWks = 0
                tmPdGroups(ilIndex).iFltNo = 0
                tmPdGroups(ilIndex).sStartDate = ""
                tmPdGroups(ilIndex).sEndDate = ""
                gSetShow pbcOffice, "", tmWKCtrls(ilIndex)
                gSetShow pbcOffice, "", tmNWCtrls(ilIndex + 1)
            Next ilIndex
            Exit Sub
        End If
    End If
    slFontName = pbcOffice.FontName
    flFontSize = pbcOffice.FontSize
    pbcOffice.FontBold = False
    pbcOffice.FontSize = 7
    pbcOffice.FontName = "Arial"
    pbcOffice.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
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
                tmPdGroups(ilIndex + 1).iYear = tmPdGroups(ilIndex).iYear    'imPdYear
            End If
            ilIndex = ilIndex + 1
        Else
            tmPdGroups(ilIndex).iYear = tmPdGroups(ilIndex).iYear + 1
            tmPdGroups(ilIndex).iStartWkNo = 1
            'Test if year exist
            If tmPdGroups(ilIndex).iYear > imBDStartYear + imBDNoYears - 1 Then
                For ilLoop = ilIndex To 4 Step 1
                    tmPdGroups(ilLoop).iStartWkNo = -1
                    tmPdGroups(ilLoop).iTrueNoWks = 0
                    tmPdGroups(ilLoop).iNoWks = 0
                Next ilLoop
                Exit Do
            End If
        End If
    Loop Until ilIndex > 4
    'Compute Start/End Date if groups
    'For ilIndex = LBound(tmPdGroups) To UBound(tmPdGroups) Step 1
    For ilIndex = LBound(tmPdGroups) + 1 To UBound(tmPdGroups) Step 1
        If tmPdGroups(ilIndex).iStartWkNo > 0 Then
            If rbcShow(0).Value Then    'Corporate
                'slDate = "1/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                'slStart = gObtainStartCorp(slDate, True)
                'slDate = "12/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                'slEnd = gObtainEndCorp(slDate, True)
                For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                    If tgMCof(ilLoop).iYear = tmPdGroups(ilIndex).iYear Then
                        If tgMCof(ilLoop).iStartMnthNo = 1 Then
                            slDate = Trim$(Str$(tgMCof(ilLoop).iStartMnthNo)) & "/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                        Else
                            slDate = Trim$(Str$(tgMCof(ilLoop).iStartMnthNo)) & "/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear - 1))
                        End If
                        slStart = gObtainStartCorp(slDate, True)
                        If tgMCof(ilLoop).iStartMnthNo = 1 Then
                            slDate = Trim$("12/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear)))
                        Else
                            slDate = Trim$(Str$(tgMCof(ilLoop).iStartMnthNo - 1)) & "/15/" & Trim$(Str$(tmPdGroups(ilIndex).iYear))
                        End If
                        slEnd = gObtainEndCorp(slDate, True)
                        Exit For
                    End If
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
                            gSetShow pbcOffice, slDate, tmWKCtrls(ilIndex)
                            slStr = Trim$(Str$(ilWkCount))
                            slStr = "# Weeks " & slStr
                            gSetShow pbcOffice, slStr, tmNWCtrls(ilIndex + 1)
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
                                gSetShow pbcOffice, slDate, tmWKCtrls(ilIndex)
                                slStr = Trim$(Str$(ilWkCount - 1))
                                slStr = "# Weeks " & slStr
                                gSetShow pbcOffice, slStr, tmNWCtrls(ilIndex + 1)
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
            gSetShow pbcOffice, "", tmWKCtrls(ilIndex)
            gSetShow pbcOffice, "", tmNWCtrls(ilIndex + 1)
        End If
    Next ilIndex
    pbcOffice.FontSize = flFontSize
    pbcOffice.FontName = slFontName
    pbcOffice.FontSize = flFontSize
    pbcOffice.FontBold = True
    mGetShowPrices
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
Private Sub mGetShowPrices()
    Dim ilLoop As Integer
    Dim ilNeg As Integer
    Dim ilBvf As Integer
    Dim ilNegIndex As Integer
    Dim ilBsf As Integer
    Dim ilIndex As Integer
    Dim ilStIndex As Integer
    Dim ilGroup As Integer
    Dim ilWk As Integer
    Dim ilFound As Integer
    Dim slStr As String
    Dim ilBox As Integer
    Dim ilCount As Integer
    'ReDim smTShow(1 To TTOTALINDEX, 1 To 1) As String 'Values shown in program area
    ReDim smTShow(0 To TTOTALINDEX, 0 To 1) As String 'Values shown in program area
    'ReDim lmTSave(1 To TTOTALINDEX - 1, 1 To 1) As Long 'Values shown in program area
    ReDim lmTSave(0 To TTOTALINDEX - 1, 0 To 1) As Long 'Values shown in program area
    'ReDim imTSave(1 To 3, 1 To 1) As Integer
    ReDim imTSave(0 To 3, 0 To 1) As Integer
    ilCount = 1
    'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
    For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
    
        ilFound = False
        If rbcSort(1).Value Or rbcSort(2).Value Then     'Vehicle within office
            'For ilLoop = LBound(imTSave, 2) To UBound(imTSave, 2) - 1 Step 1
            For ilLoop = LBound(imTSave, 2) + imLBTSave To UBound(imTSave, 2) - 1 Step 1
                    If (tgBvfRec(ilBvf).tBvf.iVefCode = imTSave(1, ilLoop)) Then
                        ilFound = True
                        Exit For
                    End If
            Next ilLoop
        ElseIf rbcSort(0).Value Or rbcSort(3).Value Then
            'For ilLoop = LBound(imTSave, 2) To UBound(imTSave, 2) - 1 Step 1
            For ilLoop = LBound(imTSave, 2) + imLBTSave To UBound(imTSave, 2) - 1 Step 1
                    If (tgBvfRec(ilBvf).tBvf.iSofCode = imTSave(2, ilLoop)) Then
                        ilFound = True
                        Exit For
                    End If
            Next ilLoop
        End If
        If Not ilFound Then
            If rbcSort(1).Value Or rbcSort(2).Value Then     'Vehicle within office
                slStr = Trim$(tgBvfRec(ilBvf).sVehicle)
                gSetShow pbcTotals, slStr, tmTCtrls(TNAMEINDEX)
                smTShow(1, ilCount) = tmTCtrls(TNAMEINDEX).sShow
            ElseIf rbcSort(0).Value Or rbcSort(3).Value Then
                slStr = Trim$(tgBvfRec(ilBvf).SOffice)
                gSetShow pbcTotals, slStr, tmTCtrls(TNAMEINDEX)
                smTShow(1, ilCount) = tmTCtrls(TNAMEINDEX).sShow
            End If
            imTSave(1, ilCount) = tgBvfRec(ilBvf).tBvf.iVefCode
            imTSave(2, ilCount) = tgBvfRec(ilBvf).tBvf.iSofCode
            imTSave(3, ilCount) = ilCount
            ilCount = ilCount + 1
            'ReDim Preserve imTSave(1 To 3, 1 To ilCount) As Integer
            ReDim Preserve imTSave(0 To 3, 0 To ilCount) As Integer
            'ReDim Preserve smTShow(1 To TTOTALINDEX, 1 To ilCount) As String 'Values shown in program area
            ReDim Preserve smTShow(0 To TTOTALINDEX, 0 To ilCount) As String 'Values shown in program area
            'ReDim Preserve lmTSave(1 To TTOTALINDEX - 1, 1 To ilCount) As Long 'Values shown in program area
            ReDim Preserve lmTSave(0 To TTOTALINDEX - 1, 0 To ilCount) As Long 'Values shown in program area
        End If
    Next ilBvf
    For ilLoop = LBound(sgBDShow, 2) + igLBBDShow To UBound(sgBDShow, 2) - 1 Step 1
        If imBDSave(1, ilLoop) <> 0 Then    'Bypass title record
            sgBDShow(DOLLAR1INDEX, ilLoop) = ""
            sgBDShow(DOLLAR2INDEX, ilLoop) = ""
            sgBDShow(DOLLAR3INDEX, ilLoop) = ""
            sgBDShow(DOLLAR4INDEX, ilLoop) = ""
            sgBDShow(TOTALINDEX, ilLoop) = ""    'Total for year
            lmBDSave(DOLLAR1INDEX - DOLLAR1INDEX + 1, ilLoop) = 0
            lmBDSave(DOLLAR2INDEX - DOLLAR1INDEX + 1, ilLoop) = 0
            lmBDSave(DOLLAR3INDEX - DOLLAR1INDEX + 1, ilLoop) = 0
            lmBDSave(DOLLAR4INDEX - DOLLAR1INDEX + 1, ilLoop) = 0
            lmBDSave(TOTALINDEX - DOLLAR1INDEX + 1, ilLoop) = 0
        End If
    Next ilLoop
    'For ilLoop = LBound(smTShow, 2)  To UBound(smTShow, 2) - 1 Step 1
    For ilLoop = LBound(smTShow, 2) + imLBTShow To UBound(smTShow, 2) - 1 Step 1
        smTShow(TDOLLAR1INDEX, ilLoop) = "0.00"
        smTShow(TDOLLAR2INDEX, ilLoop) = "0.00"
        smTShow(TDOLLAR3INDEX, ilLoop) = "0.00"
        smTShow(TDOLLAR4INDEX, ilLoop) = "0.00"
        smTShow(TTOTALINDEX, ilLoop) = "0.00"    'total for year
        lmTSave(TDOLLAR1INDEX - TDOLLAR1INDEX + 1, ilLoop) = 0
        lmTSave(TDOLLAR2INDEX - TDOLLAR1INDEX + 1, ilLoop) = 0
        lmTSave(TDOLLAR3INDEX - TDOLLAR1INDEX + 1, ilLoop) = 0
        lmTSave(TDOLLAR4INDEX - TDOLLAR1INDEX + 1, ilLoop) = 0
        lmTSave(TTOTALINDEX - TDOLLAR1INDEX + 1, ilLoop) = 0    'total for year
    Next ilLoop
    smGTShow(GTDOLLAR1INDEX) = "0.00"
    smGTShow(GTDOLLAR2INDEX) = "0.00"
    smGTShow(GTDOLLAR3INDEX) = "0.00"
    smGTShow(GTDOLLAR4INDEX) = "0.00"
    smGTShow(GTTOTALINDEX) = "0.00"    'total for year
    lmGTSave(GTDOLLAR1INDEX) = 0
    lmGTSave(GTDOLLAR2INDEX) = 0
    lmGTSave(GTDOLLAR3INDEX) = 0
    lmGTSave(GTDOLLAR4INDEX) = 0
    lmGTSave(GTTOTALINDEX) = 0    'total for year
    'Sum value
    ilStIndex = 0
    If igBDView = 0 Then
        'For ilGroup = LBound(tmPdGroups) To UBound(tmPdGroups) Step 1
        For ilGroup = LBound(tmPdGroups) + 1 To UBound(tmPdGroups) Step 1
            'For ilIndex = LBound(smBDSave, 2) To UBound(smBDSave, 2) - 1 Step 1
            For ilIndex = LBound(smBDSave, 2) + 1 To UBound(smBDSave, 2) - 1 Step 1
                ilBvf = imBDSave(1, ilIndex)
                If ilBvf > 0 Then
                    'ilStIndex = ilStIndex + 1
                    Do
                        'For ilLoop = LBound(imTSave, 2) To UBound(imTSave, 2) - 1 Step 1
                        For ilLoop = LBound(imTSave, 2) + 1 To UBound(imTSave, 2) - 1 Step 1
                            If rbcSort(1).Value Or rbcSort(2).Value Then     'Vehicle within office
                                If (tgBvfRec(ilBvf).tBvf.iVefCode = imTSave(1, ilLoop)) Then
                                    ilStIndex = imTSave(3, ilLoop)  'ilLoop
                                    Exit For
                                End If
                            ElseIf rbcSort(0).Value Or rbcSort(3).Value Then
                                If (tgBvfRec(ilBvf).tBvf.iSofCode = imTSave(2, ilLoop)) Then
                                    ilStIndex = imTSave(3, ilLoop)  'ilLoop
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                        If tgBvfRec(ilBvf).tBvf.iYear = tmPdGroups(ilGroup).iYear Then
                            If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                                'Add in the first part of the standard week
                                ilNegIndex = -1
                                For ilNeg = ilIndex + 1 To UBound(smBDSave, 2) - 1 Step 1
                                    If imBDSave(1, ilNeg) = -1 Then
                                        ilNegIndex = ilNeg
                                        Exit For
                                    End If
                                Next ilNeg
                                If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                                    lmBDSave(ilGroup, ilIndex) = lmBDSave(ilGroup, ilIndex) + tgBvfRec(ilBvf).tBvf.lGross(0)
                                    If ilNegIndex > 0 Then
                                        lmBDSave(ilGroup, ilNegIndex) = lmBDSave(ilGroup, ilNegIndex) + tgBvfRec(ilBvf).tBvf.lGross(0)
                                    End If
                                    lmTSave(ilGroup, ilStIndex) = lmTSave(ilGroup, ilStIndex) + tgBvfRec(ilBvf).tBvf.lGross(0)
                                    lmGTSave(ilGroup) = lmGTSave(ilGroup) + tgBvfRec(ilBvf).tBvf.lGross(0)
                                End If
                                For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                                    lmBDSave(ilGroup, ilIndex) = lmBDSave(ilGroup, ilIndex) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                    If ilNegIndex > 0 Then
                                        lmBDSave(ilGroup, ilNegIndex) = lmBDSave(ilGroup, ilNegIndex) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                    End If
                                    lmTSave(ilGroup, ilStIndex) = lmTSave(ilGroup, ilStIndex) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                    lmGTSave(ilGroup) = lmGTSave(ilGroup) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                Next ilWk
                                'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 = 52) And (rbcShow(0).Value) Then
                                '    lmBDSave(ilGroup, ilIndex) = lmBDSave(ilGroup, ilIndex) + tgBvfRec(ilBvf).tBvf.lGross(53)
                                '    If ilNegIndex > 0 Then
                                '        lmBDSave(ilGroup, ilNegIndex) = lmBDSave(ilGroup, ilNegIndex) + tgBvfRec(ilBvf).tBvf.lGross(53)
                                '    End If
                                '    lmTSave(ilGroup, ilStIndex) = lmTSave(ilGroup, ilStIndex) + tgBvfRec(ilBvf).tBvf.lGross(53)
                                '    lmGTSave(ilGroup) = lmGTSave(ilGroup) + tgBvfRec(ilBvf).tBvf.lGross(53)
                                'End If
                                'If ilGroup = LBound(tmPdGroups)  Then
                                If ilGroup = LBound(tmPdGroups) + 1 Then
                                    If (rbcShow(1).Value) Then
                                        lmBDSave(TOTALINDEX - 1, ilIndex) = lmBDSave(TOTALINDEX - 1, ilIndex) + tgBvfRec(ilBvf).tBvf.lGross(0)
                                        If ilNegIndex > 0 Then
                                            lmBDSave(TOTALINDEX - 1, ilNegIndex) = lmBDSave(TOTALINDEX - 1, ilNegIndex) + tgBvfRec(ilBvf).tBvf.lGross(0)
                                        End If
                                        lmTSave(TTOTALINDEX - 1, ilStIndex) = lmTSave(TTOTALINDEX - 1, ilStIndex) + tgBvfRec(ilBvf).tBvf.lGross(0)
                                        lmGTSave(GTTOTALINDEX) = lmGTSave(GTTOTALINDEX) + tgBvfRec(ilBvf).tBvf.lGross(0)
                                    End If
                                    For ilWk = 1 To 53 Step 1
                                        lmBDSave(TOTALINDEX - 1, ilIndex) = lmBDSave(TOTALINDEX - 1, ilIndex) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                        If ilNegIndex > 0 Then
                                            lmBDSave(TOTALINDEX - 1, ilNegIndex) = lmBDSave(TOTALINDEX - 1, ilNegIndex) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                        End If
                                        lmTSave(TTOTALINDEX - 1, ilStIndex) = lmTSave(TTOTALINDEX - 1, ilStIndex) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                        lmGTSave(GTTOTALINDEX) = lmGTSave(GTTOTALINDEX) + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                                    Next ilWk
                                End If
                            End If
                        End If
                        If rbcSort(2).Value Or rbcSort(3).Value Then
                            Exit Do
                        ElseIf rbcSort(0).Value Then    'Vehicle
                            If tgBvfRec(ilBvf).tBvf.iVefCode <> tgBvfRec(ilBvf + 1).tBvf.iVefCode Then
                                'ilStIndex = 0
                                Exit Do
                            Else
                                ilBvf = ilBvf + 1
                                'ilStIndex = ilStIndex + 1
                            End If
                        ElseIf rbcSort(1).Value Then    'Office
                            If tgBvfRec(ilBvf).tBvf.iSofCode <> tgBvfRec(ilBvf + 1).tBvf.iSofCode Then
                                'ilStIndex = 0
                                Exit Do
                            Else
                                ilBvf = ilBvf + 1
                                'ilStIndex = ilStIndex + 1
                            End If
                        End If
                        If ilBvf = UBound(tgBvfRec) Then
                            'ilStIndex = 0
                            Exit Do
                        End If
                    Loop
                Else
                    'ilStIndex = 0
                End If
            Next ilIndex
        Next ilGroup
    Else
        'For ilGroup = LBound(tmPdGroups) To UBound(tmPdGroups) Step 1
        For ilGroup = LBound(tmPdGroups) + 1 To UBound(tmPdGroups) Step 1
            For ilIndex = LBound(smBDSave, 2) + 1 To UBound(smBDSave, 2) - 1 Step 1
                ilBsf = imBDSave(1, ilIndex)
                If ilBsf > 0 Then
                    If tgBsfRec(ilBsf).tBsf.iYear = tmPdGroups(ilGroup).iYear Then
                        If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                            'Add in the first part of the standard week
                            If (tmPdGroups(ilGroup).iStartWkNo = 1) And (rbcShow(1).Value) Then
                                'gPDNToStr tgBsfRec(ilBsf).tBsf.sGross(0), 2, slDollar
                                'smBDSave(ilGroup + DOLLAR1INDEX - 1, ilIndex) = gAddStr(smBDSave(ilGroup + DOLLAR1INDEX - 1, ilIndex), slDollar)
                                lmBDSave(ilGroup, ilIndex) = lmBDSave(ilGroup, ilIndex) + tgBsfRec(ilBsf).tBsf.lGross(0)
                            End If
                            For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                                'gPDNToStr tgBsfRec(ilBsf).tBsf.sGross(ilWk), 2, slDollar
                                'smBDSave(ilGroup + DOLLAR1INDEX - 1, ilIndex) = gAddStr(smBDSave(ilGroup + DOLLAR1INDEX - 1, ilIndex), slDollar)
                                lmBDSave(ilGroup, ilIndex) = lmBDSave(ilGroup, ilIndex) + tgBsfRec(ilBsf).tBsf.lGross(ilWk)
                            Next ilWk
                            'If (tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1) = 52 And (rbcShow(0).Value) Then
                            '    'gPDNToStr tgBsfRec(ilBsf).tBsf.sGross(53), 2, slDollar
                            '    'smBDSave(ilGroup + DOLLAR1INDEX - 1, ilIndex) = gAddStr(smBDSave(ilGroup + DOLLAR1INDEX - 1, ilIndex), slDollar)
                            '    lmBDSave(ilGroup, ilIndex) = lmBDSave(ilGroup, ilIndex) + tgBsfRec(ilBsf).tBsf.lGross(53)
                            'End If
                            'If ilGroup = LBound(tmPdGroups) Then    'Average Dollar
                            If ilGroup = LBound(tmPdGroups) + 1 Then  'Average Dollar
                                If (rbcShow(1).Value) Then
                                    'gPDNToStr tgBsfRec(ilBsf).tBsf.sGross(0), 2, slDollar
                                    'smBDSave(TOTALINDEX, ilIndex) = gAddStr(smBDSave(TOTALINDEX, ilIndex), slDollar)
                                    lmBDSave(TOTALINDEX - 1, ilIndex) = lmBDSave(TOTALINDEX - 1, ilIndex) + tgBsfRec(ilBsf).tBsf.lGross(0)
                                End If
                                For ilWk = 1 To 53 Step 1
                                    'gPDNToStr tgBsfRec(ilBsf).tBsf.sGross(ilWk), 2, slDollar
                                    'smBDSave(TOTALINDEX, ilIndex) = gAddStr(smBDSave(TOTALINDEX, ilIndex), slDollar)
                                    lmBDSave(TOTALINDEX - 1, ilIndex) = lmBDSave(TOTALINDEX - 1, ilIndex) + tgBsfRec(ilBsf).tBsf.lGross(ilWk)
                                Next ilWk
                            End If
                        End If
                    End If
                End If
            Next ilIndex
        Next ilGroup
    End If
    'For ilGroup = LBound(tmPdGroups) To UBound(tmPdGroups) Step 1
    For ilGroup = LBound(tmPdGroups) + 1 To UBound(tmPdGroups) Step 1
        For ilIndex = LBound(lmBDSave, 2) + 1 To UBound(lmBDSave, 2) - 1 Step 1
            If imBDSave(1, ilIndex) <> 0 Then
                If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                    slStr = Trim$(Str$(lmBDSave(ilGroup, ilIndex)))
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    gSetShow pbcOffice, slStr, tmBDCtrls(DOLLAR1INDEX)
                    sgBDShow(ilGroup + DOLLAR1INDEX - 1, ilIndex) = tmBDCtrls(DOLLAR1INDEX).sShow
                    If ilGroup = LBound(lmBDSave, 2) + 1 Then
                        slStr = Trim$(Str$(lmBDSave(TOTALINDEX - 1, ilIndex)))
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                        gSetShow pbcOffice, slStr, tmBDCtrls(TOTALINDEX)
                        sgBDShow(TOTALINDEX, ilIndex) = tmBDCtrls(TOTALINDEX).sShow
                    End If
                Else
                    sgBDShow(ilGroup + DOLLAR1INDEX - 1, ilIndex) = ""
                    If ilGroup = LBound(lmBDSave, 2) + 1 Then
                        sgBDShow(TOTALINDEX, ilIndex) = ""
                    End If
                End If
            Else
                sgBDShow(ilGroup + DOLLAR1INDEX - 1, ilIndex) = ""
                If ilGroup = LBound(lmBDSave, 2) + 1 Then
                    sgBDShow(TOTALINDEX, ilIndex) = ""
                End If
            End If
        Next ilIndex
        If igBDView = 0 Then
            'For ilIndex = LBound(lmTSave, 2) To UBound(lmTSave, 2) - 1 Step 1
            For ilIndex = LBound(lmTSave, 2) + imLBTSave To UBound(lmTSave, 2) - 1 Step 1
                slStr = Trim$(Str$(lmTSave(ilGroup, ilIndex)))
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                gSetShow pbcTotals, slStr, tmTCtrls(ilGroup + 1)
                smTShow(ilGroup + 1, ilIndex) = tmTCtrls(ilGroup + 1).sShow
                'If ilGroup = LBound(lmTSave, 2) Then
                If ilGroup = LBound(lmTSave, 2) + imLBTSave Then
                    slStr = Trim$(Str$(lmTSave(TTOTALINDEX - 1, ilIndex)))
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
                    gSetShow pbcOffice, slStr, tmTCtrls(TTOTALINDEX)
                    smTShow(TTOTALINDEX, ilIndex) = tmTCtrls(TTOTALINDEX).sShow
                End If
            Next ilIndex
        End If
    Next ilGroup
    If igBDView = 0 Then
        For ilBox = imLBGTCtrls To UBound(tmGTCtrls) Step 1
            slStr = Trim$(Str$(lmGTSave(ilBox)))
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            gSetShow pbcTotals, slStr, tmGTCtrls(ilBox)
            smGTShow(ilBox) = tmGTCtrls(ilBox).sShow
        Next ilBox
    End If
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
    Dim ilLoop As Integer

    Screen.MousePointer = vbHourglass
    imLBOVCtrls = 1
    tmLBSCtrls = 1
    imLBBDCtrls = 1
    imLBWKCtrls = 1
    imLBNWCtrls = 1
    imlbTCtrls = 1
    imLBGTCtrls = 1
    imLBOVSave = 1
    igLBBvfRec = 1
    igLBBsfRec = 1
    igLBBsfRec = 1
    igLBBDShow = 1
    imLBBDSave = 1
    imLBTShow = 1
    imLBTSave = 1
    imLBUserVeh = 1
    imLBSalesOffice = 1
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    igJobShowing(BUDGETSJOB) = True
    imFirstActivate = True
    imFirstTime = True
    imPopReqd = False
    imInNew = False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    'Budget.Height = cmcScale.Top + 5 * cmcScale.Height / 3
    'gCenterForm Budget
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    imRetBranch = False
    imBDBoxNo = -1 'Initialize current Box to N/A
    imBDRowNo = 1 'Initialize current Box to N/A
    imOVBoxNo = -1 'Initialize current Box to N/A
    imSBoxNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imLbcMouseDown = False
    imBypassFocus = False
    imChgMode = False
    imBSMode = False
    imBDChg = False
    igBDView = 0
    imButtonIndex = -1
    imShowIndex = 1 'Std Month
    imTypeIndex = 0 'Quarter
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    slDate = gObtainEndStd(slDate)
    gObtainMonthYear 0, slDate, ilMonth, imNowYear
    imIgnoreRightMove = False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imFirstTimeSelect = True
    imSettingValue = False
    imLbcArrowSetting = False
    igBudgetType = 0
    
    'ReDim smBDSave(1 To 1, 1 To 1) As String 'Values saved (program name) in program area
    ReDim smBDSave(0 To 1, 0 To 1) As String 'Values saved (program name) in program area
    'ReDim lmBDSave(1 To TOTALINDEX - 1, 1 To 1) As Long 'Values saved (program name) in program area
    ReDim lmBDSave(0 To TOTALINDEX - 1, 0 To 1) As Long 'Values saved (program name) in program area
    'ReDim imBDSave(1 To 1, 1 To 1) As Integer 'Values saved (program name) in program area
    ReDim imBDSave(0 To 1, 0 To 1) As Integer 'Values saved (program name) in program area
    'ReDim smTShow(1 To TTOTALINDEX, 1 To 1) As String 'Values shown in program area
    ReDim smTShow(0 To TTOTALINDEX, 0 To 1) As String 'Values shown in program area
    'ReDim lmTSave(1 To TTOTALINDEX - 1, 1 To 1) As Long
    ReDim lmTSave(0 To TTOTALINDEX - 1, 0 To 1) As Long
    'ReDim imTSave(1 To 3, 1 To 1) As Integer
    ReDim imTSave(0 To 3, 0 To 1) As Integer
    For ilLoop = 1 To 5 Step 1
        smGTShow(ilLoop) = ""
        lmGTSave(ilLoop) = 0
    Next ilLoop
    
    hmBvf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bvf.Btr)", Budget
    On Error GoTo 0
    hmBsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmBsf, "", sgDBPath & "Bsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bsf.btr)", Budget
    On Error GoTo 0
    'ReDim tgBvfRec(1 To 1) As BVFREC
    ReDim tgBvfRec(0 To 1) As BVFREC
    'ReDim tgBsfRec(1 To 1) As BSFREC
    ReDim tgBsfRec(0 To 1) As BSFREC
    imBvfRecLen = Len(tgBvfRec(1).tBvf)
    imBsfReclen = Len(tgBsfRec(1).tBsf)
    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", Budget
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", Budget
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", Budget
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmSof = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sof.Btr)", Budget
    On Error GoTo 0
    imSofRecLen = Len(tmSof)
    hmSlf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Slf.Btr)", Budget
    On Error GoTo 0
    imSlfRecLen = Len(tmSlf)
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", Budget
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)
    ilRet = gObtainCorpCal()
    'Populate Vehicle and event type list boxes
    'lbcSalesOffice.Clear 'Force population
    smSalesOfficeCodeTag = ""
    mSaleOfficePop 'lbcSalesOffice   'Create tmSaleOffice
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lbcVehicle.Clear 'Force population
    mVehPop lbcVehicle  'Create tmUserVeh
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    'Budget.Height = cmcReport.Top + 5 * cmcReport.Height / 3
    'gCenterForm Budget
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
    gCenterStdAlone Budget
    cbcSelect.Clear  'Force list box to be populated
    mPopulate
    Screen.MousePointer = vbHourglass
    If Not imTerminate Then
        'Add clear since auto select removed
        mClearCtrlFields
        'Remove auto select until getting data is faster
        'If cbcSelect.ListCount <= 1 Then
        '    cbcSelect.ListIndex = 0 'This will generate a select_change event
        'Else
        '    cbcSelect.ListIndex = 1
        'End If
        'mSetCommands
    End If
    If tgSpf.sRUseCorpCal <> "Y" Then
        rbcShow(0).Enabled = False
        rbcShow(1).Value = True
    Else
        rbcShow(1).Enabled = False
        rbcShow(0).Value = True
    End If
    DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub mInitBDShow()
    Dim ilLp1 As Integer
    Dim ilLp2 As Integer
    ilLp2 = UBound(sgBDShow, 2)
    For ilLp1 = LBound(sgBDShow, 1) + igLBBDShow To UBound(sgBDShow, 1) Step 1
        sgBDShow(ilLp1, ilLp2) = ""
    Next ilLp1
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
    Dim llMax As Long
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long

    flTextHeight = pbcBudgetName.TextHeight("1") - 35
    plcSelect.Move 960, 30
    pbcKey.Move 3660, 990
    'Position panel and picture areas with panel
    plcOS.Move 120, 465, pbcBudgetName.Width + fgPanelAdj, pbcBudgetName.Height + fgPanelAdj - 15
    pbcBudgetName.Move plcOS.Left + fgBevelX, plcOS.Top + fgBevelY - 15
    'plcComparison.Move 4080, 465
    plcSort.Move 4715, 480  '150, 945
    plcBudget.Move 120, 975, pbcOffice.Width + vbcBudget.Width + fgPanelAdj, pbcOffice.Height + fgPanelAdj
    pbcOffice.Move plcBudget.Left + fgBevelX, plcBudget.Top + fgBevelY
    vbcBudget.Move pbcOffice.Width + 15, 15, vbcBudget.Width, pbcOffice.Height - 30
    pbcArrow.Move plcBudget.Left - pbcArrow.Width - 15
    plcTotals.Move 120, plcBudget.Top + plcBudget.Height + 15, pbcTotals.Width + vbcTotals.Width + fgPanelAdj, pbcTotals.Height + fgPanelAdj
    pbcTotals.Move plcTotals.Left + fgBevelX, plcTotals.Top + fgBevelY
    vbcTotals.Move pbcTotals.Width + 15, 15, vbcTotals.Width, pbcTotals.Height - 30
    'plcShow.Move 1185, plcTotals.Top + plcTotals.Height + 15
    'plcType.Move 5220, plcShow.Top
    'Office
    'Budget Name
    gSetCtrl tmOVCtrls(BDNAMEINDEX), 30, 30, 1845, fgBoxStH
    'Direct/Split
    gSetCtrl tmOVCtrls(DIRECTINDEX), 1890, tmOVCtrls(BDNAMEINDEX).fBoxY, 840, fgBoxStH
    'Year
    gSetCtrl tmOVCtrls(YEARINDEX), 2745, tmOVCtrls(BDNAMEINDEX).fBoxY, 525, fgBoxStH

    'Salesperson
    'Year
    gSetCtrl tmSCtrls(SYEARINDEX), 30, 30, 525, fgBoxStH
    'Name
    gSetCtrl tmBDCtrls(OSNAMEINDEX), 30, 420, 2400, fgBoxGridH
    'Dollar 1
    gSetCtrl tmBDCtrls(DOLLAR1INDEX), 2445, tmBDCtrls(OSNAMEINDEX).fBoxY, 1245, fgBoxGridH
    'Dollar 2
    gSetCtrl tmBDCtrls(DOLLAR2INDEX), 3705, tmBDCtrls(OSNAMEINDEX).fBoxY, 1245, fgBoxGridH
    'Dollar 3
    gSetCtrl tmBDCtrls(DOLLAR3INDEX), 4965, tmBDCtrls(OSNAMEINDEX).fBoxY, 1245, fgBoxGridH
    'Dollar 4
    gSetCtrl tmBDCtrls(DOLLAR4INDEX), 6225, tmBDCtrls(OSNAMEINDEX).fBoxY, 1245, fgBoxGridH
    'Total
    gSetCtrl tmBDCtrls(TOTALINDEX), 7485, tmBDCtrls(OSNAMEINDEX).fBoxY, 1245, fgBoxGridH

    'Totals
    'Name
    gSetCtrl tmTCtrls(TNAMEINDEX), 30, 30, 2385, fgBoxGridH
    'Dollar1
    gSetCtrl tmTCtrls(TDOLLAR1INDEX), 2445, tmTCtrls(TNAMEINDEX).fBoxY, 1245, fgBoxGridH
    'Dollar 2
    gSetCtrl tmTCtrls(TDOLLAR2INDEX), 3705, tmTCtrls(TNAMEINDEX).fBoxY, 1245, fgBoxGridH
    'Dollar 3
    gSetCtrl tmTCtrls(TDOLLAR3INDEX), 4965, tmTCtrls(TNAMEINDEX).fBoxY, 1245, fgBoxGridH
    'Dollar 4
    gSetCtrl tmTCtrls(TDOLLAR4INDEX), 6225, tmTCtrls(TNAMEINDEX).fBoxY, 1245, fgBoxGridH
    'Total
    gSetCtrl tmTCtrls(TTOTALINDEX), 7485, tmTCtrls(TNAMEINDEX).fBoxY, 1245, fgBoxGridH

    'Grand Totals
    'Dollar1
    gSetCtrl tmGTCtrls(GTDOLLAR1INDEX), 2445, 840, 1245, fgBoxGridH
    'Dollar 2
    gSetCtrl tmGTCtrls(GTDOLLAR2INDEX), 3705, tmGTCtrls(GTDOLLAR1INDEX).fBoxY, 1245, fgBoxGridH
    'Dollar 3
    gSetCtrl tmGTCtrls(GTDOLLAR3INDEX), 4965, tmGTCtrls(GTDOLLAR1INDEX).fBoxY, 1245, fgBoxGridH
    'Dollar 4
    gSetCtrl tmGTCtrls(GTDOLLAR4INDEX), 6225, tmGTCtrls(GTDOLLAR1INDEX).fBoxY, 1245, fgBoxGridH
    'Total
    gSetCtrl tmGTCtrls(GTTOTALINDEX), 7485, tmGTCtrls(GTDOLLAR1INDEX).fBoxY, 1245, fgBoxGridH

    'Week 1
    gSetCtrl tmWKCtrls(WK1INDEX), 2445, 30, 1245, fgBoxGridH
    'Week 2
    gSetCtrl tmWKCtrls(WK2INDEX), 3705, tmWKCtrls(WK1INDEX).fBoxY, 1245, fgBoxGridH
    'Week 3
    gSetCtrl tmWKCtrls(WK3INDEX), 4965, tmWKCtrls(WK1INDEX).fBoxY, 1245, fgBoxGridH
    'Week 4
    gSetCtrl tmWKCtrls(WK4INDEX), 6225, tmWKCtrls(WK1INDEX).fBoxY, 1245, fgBoxGridH
    '# Week 1
    gSetCtrl tmNWCtrls(NWNAMEINDEX), 30, 225, 2400, fgBoxGridH
    '# Week 1
    gSetCtrl tmNWCtrls(NW1INDEX), 2445, 225, 1245, fgBoxGridH
    '# Week 2
    gSetCtrl tmNWCtrls(NW2INDEX), 3705, tmNWCtrls(WK1INDEX).fBoxY, 1245, fgBoxGridH
    '# Week 3
    gSetCtrl tmNWCtrls(NW3INDEX), 4965, tmNWCtrls(WK1INDEX).fBoxY, 1245, fgBoxGridH
    '# Week 4
    gSetCtrl tmNWCtrls(NW4INDEX), 6225, tmNWCtrls(WK1INDEX).fBoxY, 1245, fgBoxGridH
    'Total
    gSetCtrl tmNWCtrls(NWTOTALINDEX), 7485, tmNWCtrls(WK1INDEX).fBoxY, 1245, fgBoxGridH
    tmNWCtrls(NWNAMEINDEX).sShow = "Vehicle"    '"Office/Vehicle"



    llMax = 0
    For ilLoop = imLBBDCtrls To UBound(tmBDCtrls) Step 1
        If ilLoop = OSNAMEINDEX Then
            tmBDCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmBDCtrls(ilLoop).fBoxW)
            Do While (tmBDCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmBDCtrls(ilLoop).fBoxW = tmBDCtrls(ilLoop).fBoxW + 1
            Loop
        Else
            Do
                If tmBDCtrls(ilLoop).fBoxX < tmBDCtrls(ilLoop - 1).fBoxX + tmBDCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmBDCtrls(ilLoop).fBoxX = tmBDCtrls(ilLoop).fBoxX + 15
                ElseIf tmBDCtrls(ilLoop).fBoxX > tmBDCtrls(ilLoop - 1).fBoxX + tmBDCtrls(ilLoop - 1).fBoxW + 15 Then
                    tmBDCtrls(ilLoop).fBoxX = tmBDCtrls(ilLoop).fBoxX - 15
                Else
                    Exit Do
                End If
            Loop
        End If
        If tmBDCtrls(ilLoop).fBoxX + tmBDCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmBDCtrls(ilLoop).fBoxX + tmBDCtrls(ilLoop).fBoxW + 15
        End If
        tmTCtrls(ilLoop).fBoxX = tmBDCtrls(ilLoop).fBoxX
        tmTCtrls(ilLoop).fBoxW = tmBDCtrls(ilLoop).fBoxW
        If ilLoop > imLBBDCtrls Then
            tmGTCtrls(ilLoop - 1).fBoxX = tmBDCtrls(ilLoop).fBoxX
            tmGTCtrls(ilLoop - 1).fBoxW = tmBDCtrls(ilLoop).fBoxW
            If ilLoop <= UBound(tmWKCtrls) + 1 Then
                tmWKCtrls(ilLoop - 1).fBoxX = tmBDCtrls(ilLoop).fBoxX
                tmWKCtrls(ilLoop - 1).fBoxW = tmBDCtrls(ilLoop).fBoxW
            End If
        End If
        tmNWCtrls(ilLoop).fBoxX = tmBDCtrls(ilLoop).fBoxX
        tmNWCtrls(ilLoop).fBoxW = tmBDCtrls(ilLoop).fBoxW
    Next ilLoop
    pbcOffice.Picture = LoadPicture("")
    pbcOffice.Width = llMax
    pbcSalesperson.Picture = LoadPicture("")
    pbcSalesperson.Width = llMax
    plcBudget.Width = llMax + vbcBudget.Width + 2 * fgBevelX + 15
    lacOFrame.Width = llMax - 15
    lacSFrame.Width = llMax - 15
    pbcTotals.Picture = LoadPicture("")
    pbcTotals.Width = llMax
    plcTotals.Width = llMax + vbcTotals.Width + 2 * fgBevelX + 15
    Me.Width = plcBudget.Width + 3 * plcBudget.Left
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    'cmcDone.Left = (Budget.Width - 7 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    'Me.Width = 9450
    'Me.Height = 6030
    cmcDone.Left = (Me.Width - 6 * (cmcDone.Width + ilSpaceBetweenButtons) + ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcUndo.Left = cmcUpdate.Left + cmcUpdate.Width + ilSpaceBetweenButtons
    'cmcReport.Left = cmcUndo.Left + cmcUndo.Width + ilSpaceBetweenButtons
    'cmc12Mos.Left = cmcReport.Left + cmcReport.Width + ilSpaceBetweenButtons
    cmc12Mos.Left = cmcUndo.Left + cmcUndo.Width + ilSpaceBetweenButtons
    cmcAddVeh.Left = cmc12Mos.Left + cmc12Mos.Width + ilSpaceBetweenButtons
    cmcDone.Top = Budget.Height - (3 * cmcDone.Height) / 2 - cmcDone.Height - 60
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    cmcUndo.Top = cmcDone.Top
    'cmcReport.Top = cmcDone.Top
    cmc12Mos.Top = cmcDone.Top
    cmcAddVeh.Top = cmcDone.Top

    cmcAdvt.Left = cmcDone.Left + cmcDemo.Width / 2
    cmcTrend.Left = cmcAdvt.Left + cmcAdvt.Width + ilSpaceBetweenButtons
    cmcScale.Left = cmcTrend.Left + cmcTrend.Width + ilSpaceBetweenButtons
    cmcDemo.Left = cmcScale.Left + cmcScale.Width + ilSpaceBetweenButtons
    cmcActuals.Left = cmcDemo.Left + cmcDemo.Width + ilSpaceBetweenButtons
    cmcAdvt.Top = cmcCancel.Top + cmcCancel.Height + 60
    cmcTrend.Top = cmcAdvt.Top
    cmcScale.Top = cmcAdvt.Top
    cmcDemo.Top = cmcAdvt.Top
    cmcActuals.Top = cmcAdvt.Top

    imcTrash.Top = cmcDone.Top + cmcDone.Height - imcTrash.Height
    imcTrash.Left = Budget.Width - (3 * imcTrash.Width) / 2
    plcShow.Top = cmcDone.Top - plcShow.Height - 60
    plcType.Top = plcShow.Top
    llAdjTop = plcShow.Top - plcBudget.Top - fgBevelY - 120
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    llAdjTop = llAdjTop
    Do While plcBudget.Top + llAdjTop + 2 * fgBevelY + 240 < plcShow.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    llAdjTop = llAdjTop + 60
    plcBudget.Height = llAdjTop + 2 * fgBevelY
    pbcSalesperson.Left = plcBudget.Left + fgBevelX
    pbcSalesperson.Top = plcBudget.Top + fgBevelY
    pbcSalesperson.Height = plcBudget.Height - 2 * fgBevelY

    llAdjTop = llAdjTop - plcTotals.Height - 240
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While plcBudget.Top + llAdjTop + 2 * fgBevelY + 240 < plcShow.Top - plcTotals.Height
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    plcBudget.Height = llAdjTop + 2 * fgBevelY
    pbcOffice.Left = plcBudget.Left + fgBevelX
    pbcOffice.Top = plcBudget.Top + fgBevelY
    pbcOffice.Height = plcBudget.Height - 2 * fgBevelY
    plcTotals.Move plcBudget.Left, plcBudget.Top + plcBudget.Height + 60, plcBudget.Width
    pbcTotals.Left = plcTotals.Left + fgBevelX
    pbcTotals.Top = plcTotals.Top + fgBevelY

    vbcBudget.Left = plcBudget.Width - vbcBudget.Width - fgBevelX - 30
    vbcBudget.Top = fgBevelY
    vbcBudget.Height = pbcOffice.Height

    vbcTotals.Left = plcTotals.Width - vbcTotals.Width - fgBevelX - 30
    vbcTotals.Top = fgBevelY
    vbcTotals.Height = pbcTotals.Height

    pbcLnWkArrow(0).Left = tmBDCtrls(DOLLAR1INDEX).fBoxX - pbcLnWkArrow(0).Width - 30
    pbcLnWkArrow(0).Top = 15
    pbcLnWkArrow(1).Left = tmBDCtrls(DOLLAR1INDEX + 3).fBoxX + tmBDCtrls(DOLLAR1INDEX + 3).fBoxW + 60
    pbcLnWkArrow(1).Top = 15
    pbcLnWkArrow(2).Left = tmBDCtrls(DOLLAR1INDEX).fBoxX - pbcLnWkArrow(0).Width - 30
    pbcLnWkArrow(2).Top = 15
    pbcLnWkArrow(3).Left = tmBDCtrls(DOLLAR1INDEX + 3).fBoxX + tmBDCtrls(DOLLAR1INDEX + 3).fBoxW + 60
    pbcLnWkArrow(3).Top = 15

    pbcTab.Top = Budget.Height
    pbcClickFocus.Top = Budget.Height
    plcSelect.Left = plcBudget.Left + plcBudget.Width - plcSelect.Width
    pbcStartNew.Left = plcSelect.Left + plcSelect.Width + 120

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBudgetCtrls                *
'*                                                     *
'*             Created:7/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Sub mInitBudgetCtrls()
    Dim ilUpperBound As Integer
    Dim ilLower As Integer
    imPdYear = imNowYear
    imPdStartWk = 1
    imPdStartFltNo = 1
    imBDStartYear = 0
    imBDNoYears = 0
    If igBDView = 0 Then
        If UBound(tgBvfRec) > 1 Then
            'ilLower = LBound(tgBvfRec)
            ilLower = igLBBvfRec
            imBDStartYear = tgBvfRec(ilLower).tBvf.iYear
            imBDNoYears = 1
            'Adjust Period to be viewed
            If imBDStartYear > imPdYear Then
                imPdYear = imBDStartYear
            ElseIf imBDStartYear + imBDNoYears - 1 < imPdYear Then
                imPdYear = imBDStartYear + imBDNoYears - 1
            End If
        End If
        ilUpperBound = UBound(tgBvfRec)
    Else
        If UBound(tgBsfRec) > 1 Then
            'ilLower = LBound(tgBsfRec)
            ilLower = igLBBsfRec
            imBDStartYear = tgBsfRec(ilLower).tBsf.iYear
            imBDNoYears = 1
            'Adjust Period to be viewed
            If imBDStartYear > imPdYear Then
                imPdYear = imBDStartYear
            ElseIf imBDStartYear + imBDNoYears - 1 < imPdYear Then
                imPdYear = imBDStartYear + imBDNoYears - 1
            End If
        End If
        ilUpperBound = UBound(tgBsfRec)
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
'   mMoveCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    If igBDView = 0 Then
        'For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilLoop = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
        
        Next ilLoop
    Else
        'For ilLoop = LBound(tgBsfRec) To UBound(tgBsfRec) - 1 Step 1
        For ilLoop = igLBBvfRec To UBound(tgBsfRec) - 1 Step 1
        
        Next ilLoop
    End If
    Exit Sub

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
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim slPOffice As String
    Dim slPVehicle As String
    Dim ilUpper As Integer
    If igBDView = 0 Then
        'ReDim sgBDShow(1 To TOTALINDEX, 1 To 1) As String * 40 'Values shown in program area
        ReDim sgBDShow(0 To TOTALINDEX, 0 To 1) As String * 40 'Values shown in program area
        mInitBDShow
        'ReDim smBDSave(1 To TOTALINDEX, 1 To 1) As String 'Values saved (program name) in program area
        'ReDim smBDSave(1 To 1, 1 To 1) As String 'Values saved (program name) in program area
        ReDim smBDSave(0 To 1, 0 To 1) As String 'Values saved (program name) in program area
        'ReDim lmBDSave(1 To TOTALINDEX - 1, 1 To 1) As Long 'Values saved (program name) in program area
        ReDim lmBDSave(0 To TOTALINDEX - 1, 0 To 1) As Long 'Values saved (program name) in program area
        'ReDim imBDSave(1 To 1, 1 To 1) As Integer 'Values saved (program name) in program area
        ReDim imBDSave(0 To 1, 0 To 1) As Integer 'Values saved (program name) in program area
        'Count the number of totals
        If UBound(tgBvfRec) > 1 Then
            'ilCount = 1
            'slPOffice = Trim$(tgBvfRec(LBound(tgBvfRec)).sOffice)
            'slPVehicle = Trim$(tgBvfRec(LBound(tgBvfRec)).sVehicle)
            'ReDim smTShow(1 To TTOTALINDEX, 1 To 2) As String'Values shown in program area
            'ReDim lmTSave(1 To TTOTALINDEX - 1, 1 To 2) As Long
            'If rbcSort(1).Value Or rbcSort(2).Value Then    'Vehicle within office
            '    slStr = slPVehicle
            'ElseIf rbcSort(0).Value Or rbcSort(3).Value Then
            '    slStr = slPOffice
            'End If
            'gSetShow pbcTotals, slStr, tmTCtrls(TNAMEINDEX)
            'smTShow(1, ilCount) = tmTCtrls(TNAMEINDEX).sShow
            'For ilRowNo = LBound(tgBvfRec) + 1 To UBound(tgBvfRec) - 1 Step 1
            '    If rbcSort(1).Value Or rbcSort(2).Value Then     'Vehicle within office
            '        If StrComp(slPOffice, Trim$(tgBvfRec(ilRowNo).sOffice), 1) <> 0 Then
            '            Exit For
            '        Else
            '            ilCount = ilCount + 1
            '            ReDim Preserve smTShow(1 To TTOTALINDEX, 1 To ilCount + 1) As String'Values shown in program area
            '            ReDim Preserve lmTSave(1 To TTOTALINDEX - 1, 1 To ilCount + 1) As Long'Values shown in program area
            '            slStr = Trim$(tgBvfRec(ilRowNo).sVehicle)
            '            gSetShow pbcTotals, slStr, tmTCtrls(TNAMEINDEX)
            '            smTShow(1, ilCount) = tmTCtrls(TNAMEINDEX).sShow
            '        End If
            '    ElseIf rbcSort(0).Value Or rbcSort(3).Value Then
            '        If StrComp(slPVehicle, Trim$(tgBvfRec(ilRowNo).sVehicle), 1) <> 0 Then
            '            Exit For
            '        Else
            '            ilCount = ilCount + 1
            '            ReDim Preserve smTShow(1 To TTOTALINDEX, 1 To ilCount + 1) As String'Values shown in program area
            '            ReDim Preserve lmTSave(1 To TTOTALINDEX - 1, 1 To ilCount + 1) As Long'Values shown in program area
            '            slStr = Trim$(tgBvfRec(ilRowNo).sOffice)
            '            gSetShow pbcTotals, slStr, tmTCtrls(TNAMEINDEX)
            '            smTShow(1, ilCount) = tmTCtrls(TNAMEINDEX).sShow
            '        End If
            '    End If
            'Next ilRowNo
            'Moved to mGetShowPrice as sort can change
            'ReDim smTShow(1 To TTOTALINDEX, 1 To 1) As String 'Values shown in program area
            ReDim smTShow(0 To TTOTALINDEX, 0 To 1) As String 'Values shown in program area
            'ReDim lmTSave(1 To TTOTALINDEX - 1, 1 To 1) As Long 'Values shown in program area
            ReDim lmTSave(0 To TTOTALINDEX - 1, 0 To 1) As Long 'Values shown in program area
            'ReDim imTSave(1 To 3, 1 To 1) As Integer
            ReDim imTSave(0 To 3, 0 To 1) As Integer
        Else
            'ReDim smTShow(1 To TTOTALINDEX, 1 To 1) As String 'Values shown in program area
            ReDim smTShow(0 To TTOTALINDEX, 0 To 1) As String 'Values shown in program area
            'ReDim lmTSave(1 To TTOTALINDEX - 1, 1 To 1) As Long 'Values shown in program area
            ReDim lmTSave(0 To TTOTALINDEX - 1, 0 To 1) As Long 'Values shown in program area
            'ReDim imTSave(1 To 3, 1 To 1) As Integer
            ReDim imTSave(0 To 3, 0 To 1) As Integer
        End If
        ilUpper = UBound(smBDSave, 2)
        slStr = smMnfName
        smOVSave(1) = smMnfName
        gSetShow pbcOffice, slStr, tmOVCtrls(BDNAMEINDEX)
        If UBound(tgBvfRec) > 1 Then
            If igBudgetType = 0 Then
                If tgBvfRec(1).tBvf.sSplit = "S" Then
                    slStr = "Split"
                    imOVSave(2) = 1
                Else
                    slStr = "Direct"
                    imOVSave(2) = 0
                End If
            Else
                imOVSave(2) = 0
                slStr = "Actuals"
            End If
        Else
            slStr = ""
            imOVSave(2) = -1
        End If
        gSetShow pbcOffice, slStr, tmOVCtrls(DIRECTINDEX)
        imOVSave(1) = imYear
        If imYear <> 0 Then
            slStr = Trim$(Str$(imYear))
        Else
            slStr = ""
        End If
        gSetShow pbcOffice, slStr, tmOVCtrls(YEARINDEX)
        slPOffice = ""
        slPVehicle = ""
        'For ilRowNo = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilRowNo = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If slPOffice = "" Then
                'Add row to smSave and smShow
                If rbcSort(1).Value Or rbcSort(2).Value Then    'Vehicle within office
                    smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).SOffice)
                ElseIf rbcSort(0).Value Or rbcSort(3).Value Then
                    smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).sVehicle)
                End If
                If rbcSort(0).Value Or rbcSort(1).Value Then
                    imBDSave(1, ilUpper) = ilRowNo
                    tgBvfRec(ilRowNo).iSaveIndex = ilUpper
                ElseIf rbcSort(2).Value Or rbcSort(3).Value Then
                    imBDSave(1, ilUpper) = 0
                End If
                slStr = smBDSave(1, ilUpper)
                gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                sgBDShow(1, ilUpper) = tmBDCtrls(OSNAMEINDEX).sShow
                ilUpper = ilUpper + 1
                'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                mInitBDShow
                'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                If rbcSort(2).Value Or rbcSort(3).Value Then
                    If rbcSort(2).Value Then    'Vehicle within office
                        smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).sVehicle)
                    ElseIf rbcSort(3).Value Then
                        smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).SOffice)
                    End If
                    imBDSave(1, ilUpper) = ilRowNo
                    tgBvfRec(ilRowNo).iSaveIndex = ilUpper
                    slStr = "__" & smBDSave(1, ilUpper)
                    gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                    sgBDShow(1, ilUpper) = "  " & right$(tmBDCtrls(OSNAMEINDEX).sShow, Len(tmBDCtrls(OSNAMEINDEX).sShow) - 2)
                    ilUpper = ilUpper + 1
                    'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                    ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                    mInitBDShow
                    'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                    'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                    ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                    'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                    ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                    'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                    ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                End If
                slPOffice = Trim$(tgBvfRec(ilRowNo).SOffice)
                slPVehicle = Trim$(tgBvfRec(ilRowNo).sVehicle)
            Else
                If rbcSort(1).Value Then
                    If StrComp(slPOffice, Trim$(tgBvfRec(ilRowNo).SOffice), 1) <> 0 Then
                        smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).SOffice)
                        imBDSave(1, ilUpper) = ilRowNo
                        tgBvfRec(ilRowNo).iSaveIndex = ilUpper
                        slStr = smBDSave(1, ilUpper)
                        gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                        sgBDShow(1, ilUpper) = tmBDCtrls(OSNAMEINDEX).sShow
                        ilUpper = ilUpper + 1
                        'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                        ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                        mInitBDShow
                        'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                        ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                        ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                        'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                        ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                    End If
                    slPOffice = Trim$(tgBvfRec(ilRowNo).SOffice)
                ElseIf rbcSort(2).Value Then    'Vehicle within office
                    If StrComp(slPOffice, Trim$(tgBvfRec(ilRowNo).SOffice), 1) <> 0 Then
                        smBDSave(1, ilUpper) = "Total: " & slPOffice
                        imBDSave(1, ilUpper) = -1
                        slStr = "__" & smBDSave(1, ilUpper)
                        gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                        sgBDShow(1, ilUpper) = "  " & right$(tmBDCtrls(OSNAMEINDEX).sShow, Len(tmBDCtrls(OSNAMEINDEX).sShow) - 2)
                        ilUpper = ilUpper + 1
                        'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                        ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                        mInitBDShow
                        'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                        ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                        ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                        'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                        ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                        smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).SOffice)
                        imBDSave(1, ilUpper) = 0
                        slStr = smBDSave(1, ilUpper)
                        gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                        sgBDShow(1, ilUpper) = tmBDCtrls(OSNAMEINDEX).sShow
                        ilUpper = ilUpper + 1
                        'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                        ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                        mInitBDShow
                        'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                        ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                        ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                        'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                        ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                    End If
                    smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).sVehicle)
                    imBDSave(1, ilUpper) = ilRowNo
                    tgBvfRec(ilRowNo).iSaveIndex = ilUpper
                    slStr = "__" & smBDSave(1, ilUpper)
                    gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                    sgBDShow(1, ilUpper) = "  " & right$(tmBDCtrls(OSNAMEINDEX).sShow, Len(tmBDCtrls(OSNAMEINDEX).sShow) - 2)
                    ilUpper = ilUpper + 1
                    'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                    ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                    mInitBDShow
                    'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                    'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                    ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                    'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                    ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                    'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                    ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                    slPOffice = Trim$(tgBvfRec(ilRowNo).SOffice)
                ElseIf rbcSort(0).Value Then
                    If StrComp(slPVehicle, Trim$(tgBvfRec(ilRowNo).sVehicle), 1) <> 0 Then
                        smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).sVehicle)
                        imBDSave(1, ilUpper) = ilRowNo
                        tgBvfRec(ilRowNo).iSaveIndex = ilUpper
                        slStr = smBDSave(1, ilUpper)
                        gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                        sgBDShow(1, ilUpper) = tmBDCtrls(OSNAMEINDEX).sShow
                        ilUpper = ilUpper + 1
                        'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                        ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                        mInitBDShow
                        'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                        ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                        ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                        'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                        ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                    End If
                    slPVehicle = Trim$(tgBvfRec(ilRowNo).sVehicle)
                ElseIf rbcSort(3).Value Then
                    If StrComp(slPVehicle, Trim$(tgBvfRec(ilRowNo).sVehicle), 1) <> 0 Then
                        smBDSave(1, ilUpper) = "Total: " & slPVehicle
                        imBDSave(1, ilUpper) = -1
                        slStr = "__" & smBDSave(1, ilUpper)
                        gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                        sgBDShow(1, ilUpper) = "  " & right$(tmBDCtrls(OSNAMEINDEX).sShow, Len(tmBDCtrls(OSNAMEINDEX).sShow) - 2)
                        ilUpper = ilUpper + 1
                        'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                        ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                        mInitBDShow
                        'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                        ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                        ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                        'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                        ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                        smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).sVehicle)
                        imBDSave(1, ilUpper) = 0
                        slStr = smBDSave(1, ilUpper)
                        gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                        sgBDShow(1, ilUpper) = tmBDCtrls(OSNAMEINDEX).sShow
                        ilUpper = ilUpper + 1
                        'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                        ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                        mInitBDShow
                        'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                        ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                        'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                        ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                        'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                        ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                    End If
                    smBDSave(1, ilUpper) = Trim$(tgBvfRec(ilRowNo).SOffice)
                    imBDSave(1, ilUpper) = ilRowNo
                    tgBvfRec(ilRowNo).iSaveIndex = ilUpper
                    slStr = "__" & smBDSave(1, ilUpper)
                    gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
                    sgBDShow(1, ilUpper) = "  " & right$(tmBDCtrls(OSNAMEINDEX).sShow, Len(tmBDCtrls(OSNAMEINDEX).sShow) - 2)
                    ilUpper = ilUpper + 1
                    'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
                    ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
                    mInitBDShow
                    'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
                    'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
                    ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
                    'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
                    ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
                    'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
                    ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
                    slPVehicle = Trim$(tgBvfRec(ilRowNo).sVehicle)
                End If
            End If
        Next ilRowNo
        If rbcSort(2).Value Then    'Vehicle within office
            smBDSave(1, ilUpper) = "Total: " & slPOffice
            imBDSave(1, ilUpper) = -1
            slStr = "__" & smBDSave(1, ilUpper)
            gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
            sgBDShow(1, ilUpper) = "  " & right$(tmBDCtrls(OSNAMEINDEX).sShow, Len(tmBDCtrls(OSNAMEINDEX).sShow) - 2)
            ilUpper = ilUpper + 1
            'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
            ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
            mInitBDShow
            'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
            'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
            ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
            'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
            ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
            'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
            ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
        ElseIf rbcSort(3).Value Then
            smBDSave(1, ilUpper) = "Total: " & slPVehicle
            imBDSave(1, ilUpper) = -1
            slStr = "__" & smBDSave(1, ilUpper)
            gSetShow pbcOffice, slStr, tmBDCtrls(OSNAMEINDEX)
            sgBDShow(1, ilUpper) = "  " & right$(tmBDCtrls(OSNAMEINDEX).sShow, Len(tmBDCtrls(OSNAMEINDEX).sShow) - 2)
            ilUpper = ilUpper + 1
            'ReDim Preserve sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
            ReDim Preserve sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
            mInitBDShow
            'ReDim Preserve smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
            'ReDim Preserve smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
            ReDim Preserve smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
            'ReDim Preserve lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
            ReDim Preserve lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
            'ReDim Preserve imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
            ReDim Preserve imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
        End If
    Else
        ilUpper = UBound(tgBsfRec)
        'ReDim sgBDShow(1 To TOTALINDEX, 1 To ilUpper) As String * 40 'Values shown in program area
        ReDim sgBDShow(0 To TOTALINDEX, 0 To ilUpper) As String * 40 'Values shown in program area
        mInitBDShow
        'ReDim smBDSave(1 To TOTALINDEX, 1 To ilUpper) As String 'Values saved (program name) in program area
        'ReDim smBDSave(1 To 1, 1 To ilUpper) As String 'Values saved (program name) in program area
        ReDim smBDSave(0 To 1, 0 To ilUpper) As String 'Values saved (program name) in program area
        'ReDim lmBDSave(1 To TOTALINDEX - 1, 1 To ilUpper) As Long 'Values saved (program name) in program area
        ReDim lmBDSave(0 To TOTALINDEX - 1, 0 To ilUpper) As Long 'Values saved (program name) in program area
        'ReDim imBDSave(1 To 1, 1 To ilUpper) As Integer 'Values saved (program name) in program area
        ReDim imBDSave(0 To 1, 0 To ilUpper) As Integer 'Values saved (program name) in program area
        'ReDim smTShow(1 To TTOTALINDEX, 1 To 1) As String 'Values shown in program area
        ReDim smTShow(0 To TTOTALINDEX, 0 To 1) As String 'Values shown in program area
        ReDim lmTSave(0 To TTOTALINDEX - 1, 0 To 1) As Long 'Values shown in program area
        'ReDim imTSave(1 To 3, 1 To 1) As Integer
        ReDim imTSave(0 To 3, 0 To 1) As Integer
        imSSave(1) = imYear
        If imYear <> 0 Then
            slStr = Trim$(Str$(imYear))
        Else
            slStr = ""
        End If
        gSetShow pbcSalesperson, slStr, tmSCtrls(SYEARINDEX)
        'For ilRowNo = LBound(tgBsfRec) To UBound(tgBsfRec) - 1 Step 1
        For ilRowNo = igLBBsfRec To UBound(tgBsfRec) - 1 Step 1
        
            'Get salesperson name
            smBDSave(1, ilRowNo) = Trim$(tgBsfRec(ilRowNo).sKey)
            imBDSave(1, ilRowNo) = ilRowNo
            tgBsfRec(ilRowNo).iSaveIndex = ilRowNo
            slStr = smBDSave(1, ilRowNo)
            gSetShow pbcSalesperson, slStr, tmBDCtrls(OSNAMEINDEX)
            sgBDShow(1, ilRowNo) = tmBDCtrls(OSNAMEINDEX).sShow
        Next ilRowNo
    End If
    mGetShowDates
    imSettingValue = True
    If igBDView = 0 Then
        vbcBudget.LargeChange = 11
    Else
        vbcBudget.LargeChange = 17
    End If
    imSettingValue = True
    vbcBudget.Min = LBound(sgBDShow) + igLBBDShow
    imSettingValue = True
    If UBound(sgBDShow, 2) - 1 <= vbcBudget.LargeChange + 1 Then ' + 1 Then
        vbcBudget.Max = LBound(sgBDShow, 2) + igLBBDShow
    Else
        vbcBudget.Max = UBound(sgBDShow, 2) - vbcBudget.LargeChange
    End If
    imSettingValue = True
    vbcBudget.Value = vbcBudget.Min
    imSettingValue = True
    'vbcTotals.Min = LBound(smTShow)
    vbcTotals.Min = LBound(smTShow) + imLBTShow
    imSettingValue = True
    If UBound(smTShow, 2) - 1 <= vbcTotals.LargeChange + 1 Then ' + 1 Then
        vbcTotals.Max = LBound(smTShow, 2) + imLBTShow
    Else
        vbcTotals.Max = UBound(smTShow, 2) - vbcTotals.LargeChange
    End If
    imSettingValue = True
    vbcTotals.Value = vbcTotals.Min
    imSettingValue = False
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

    If igBDView = 0 Then
        'Show Unique Budget names and year
        imPopReqd = False
        'ilRet = gPopVehBudgetBox(Budget, 0, 1, cbcSelect, lbcBudget)
        ilRet = gPopVehBudgetBox(Budget, igBudgetType, 0, 1, cbcSelect, tmBudNameCode(), smBudNameCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mPopulateErr
            gCPErrorMsg ilRet, "mPopulate (gPopBudgetBox)", Budget
            On Error GoTo 0
            cbcSelect.AddItem "[New]", 0  'Force as first item on list
            imPopReqd = True
        End If
    Else
        'Show unique years only
        imPopReqd = False
        'ilRet = gPopSlspBudgetBox(Budget, cbcSelect, lbcBudget)
        ilRet = gPopSlspBudgetBox(Budget, cbcSelect, tmBudNameCode(), smBudNameCodeTag)
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mPopulateErr
            gCPErrorMsg ilRet, "mPopulate (gPopBudgetBox)", Budget
            On Error GoTo 0
            cbcSelect.AddItem "[New]", 0  'Force as first item on list
            imPopReqd = True
        End If
    End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadBsfRec                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadBsfRec(ilYear As Integer, ilNewYear As Integer) As Integer
'
'   iRet = mReadBsfRec (ilYear, ilNewYear)
'   Where:
'       ilYears(I)-Year to retrieve
'       ilNewYear(I)- Year to get new records for
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim ilRecOK As Integer
    Dim ilFound As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    'ReDim tgBsfRec(1 To 1) As BSFREC
    ReDim tgBsfRec(0 To 1) As BSFREC

    ilUpper = UBound(tgBsfRec)
    btrExtClear hmBsf   'Clear any previous extend operation
    ilExtLen = Len(tgBsfRec(1).tBsf)  'Extract operation record size
    tmBsfSrchKey.iYear = ilYear
    tmBsfSrchKey.iSeqNo = 1
    tmBsfSrchKey.iSlfCode = 0
    ilRet = btrGetGreaterOrEqual(hmBsf, tgBsfRec(1).tBsf, imBsfReclen, tmBsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'ilRet = btrGetFirst(hmBsf, tgBsfRec(1).tBsf, imBsfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmBsf, llNoRec, -1, "UC", "BSF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Bsf", "BsfYear")
        tlIntTypeBuff.iType = ilYear
        ilRet = btrExtAddLogicConst(hmBsf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        On Error GoTo mReadBsfRecErr
        gBtrvErrorMsg ilRet, "mReadBsfRec (btrExtAddLogicConst):" & "Bsf.Btr", Budget
        On Error GoTo 0
        ilRet = btrExtAddField(hmBsf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mReadBsfRecErr
        gBtrvErrorMsg ilRet, "mReadBsfRec (btrExtAddField):" & "Bsf.Btr", Budget
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmBsf, tgBsfRec(ilUpper).tBsf, ilExtLen, tgBsfRec(ilUpper).lRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadBsfRecErr
            gBtrvErrorMsg ilRet, "mReadBsfRec (btrExtGetNextExt):" & "Bsf.Btr", Budget
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tgBsfRec(1).tBsf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmBsf, tgBsfRec(ilUpper).tBsf, ilExtLen, tgBsfRec(ilUpper).lRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""
                ilRecOK = True
                If tgBsfRec(ilUpper).tBsf.iSlfCode <> tmSlf.iCode Then
                    tmSlfSrchKey.iCode = tgBsfRec(ilUpper).tBsf.iSlfCode 'ilCode
                    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet = BTRV_ERR_NONE Then
                    If igSlfFirstNameFirst Then
                        slStr = Trim$(tmSlf.sFirstName) & " " & Trim$(tmSlf.sLastName)
                    Else
                        slStr = Trim$(tmSlf.sLastName) & ", " & Trim$(tmSlf.sFirstName)
                    End If
                Else
                    ilRecOK = False
                End If
                If ilRecOK Then
                    tgBsfRec(ilUpper).sKey = slStr
                    tgBsfRec(ilUpper).iStatus = 1
                    ilUpper = ilUpper + 1
                    'ReDim Preserve tgBsfRec(1 To ilUpper) As BSFREC
                    ReDim Preserve tgBsfRec(0 To ilUpper) As BSFREC
                End If
                ilRet = btrExtGetNext(hmBsf, tgBsfRec(ilUpper).tBsf, ilExtLen, tgBsfRec(ilUpper).lRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmBsf, tgBsfRec(ilUpper).tBsf, ilExtLen, tgBsfRec(ilUpper).lRecPos)
                Loop
            Loop
        End If
    End If
    'Test if records missing
    ilUpper = UBound(tgBsfRec)
    'If ilUpper > LBound(tgBsfRec) Then  'If no records exist, wait to model to create
        ilRet = btrGetFirst(hmSlf, tmSlf, imSlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            'For ilLoop = LBound(tgBsfRec) To UBound(tgBsfRec) Step 1
            For ilLoop = igLBBsfRec To UBound(tgBsfRec) Step 1
                If tgBsfRec(ilLoop).tBsf.iSlfCode = tmSlf.iCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                tgBsfRec(ilUpper).tBsf.iSlfCode = tmSlf.iCode
                tgBsfRec(ilUpper).tBsf.iYear = ilNewYear    'tgBsfRec(LBound(tgBsfRec)).tBsf.iYear
                'tgBsfRec(ilUpper).tBsf.iSeqNo = tgBsfRec(LBound(tgBsfRec)).tBsf.iSeqNo
                tgBsfRec(ilUpper).tBsf.iSeqNo = tgBsfRec(igLBBsfRec).tBsf.iSeqNo
                'tgBsfRec(ilUpper).tBsf.iStartDate(0) = tgBsfRec(LBound(tgBsfRec)).tBsf.iStartDate(0)
                tgBsfRec(ilUpper).tBsf.iStartDate(0) = tgBsfRec(igLBBsfRec).tBsf.iStartDate(0)
                'tgBsfRec(ilUpper).tBsf.iStartDate(1) = tgBsfRec(LBound(tgBsfRec)).tBsf.iStartDate(1)
                tgBsfRec(ilUpper).tBsf.iStartDate(1) = tgBsfRec(igLBBsfRec).tBsf.iStartDate(1)
                'For ilLoop = LBound(tgBsfRec(ilUpper).tBsf.lGross) To UBound(tgBsfRec(ilUpper).tBsf.lGross) Step 1
                For ilLoop = LBound(tgBsfRec(ilUpper).tBsf.lGross) To UBound(tgBsfRec(ilUpper).tBsf.lGross) Step 1
                    'slStr = ""
                    'gStrToPDN slStr, 2, 5, tgBsfRec(ilUpper).tBsf.sGross(ilLoop)
                    tgBsfRec(ilUpper).tBsf.lGross(ilLoop) = 0
                Next ilLoop
                If igSlfFirstNameFirst Then
                    slStr = Trim$(tmSlf.sFirstName) & " " & Trim$(tmSlf.sLastName)
                Else
                    slStr = Trim$(tmSlf.sLastName) & ", " & Trim$(tmSlf.sFirstName)
                End If
                tgBsfRec(ilUpper).sKey = slStr
                tgBsfRec(ilUpper).iStatus = 0
                ilUpper = ilUpper + 1
                'ReDim Preserve tgBsfRec(1 To ilUpper) As BSFREC
                ReDim Preserve tgBsfRec(0 To ilUpper) As BSFREC
            End If
            ilRet = btrGetNext(hmSlf, tmSlf, imSlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
    'End If
    If ilUpper > 1 Then
        ArraySortTyp fnAV(tgBsfRec(), 1), UBound(tgBsfRec) - 1, 0, LenB(tgBsfRec(1)), 0, LenB(tgBsfRec(1).sKey), 0
    End If
    'mInitBudgetCtrls
    mReadBsfRec = True
    Exit Function
mReadBsfRecErr:
    On Error GoTo 0
    mReadBsfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadBvfRec                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadBvfRec(ilMnfCode As Integer, ilYear As Integer, ilNewYear As Integer, ilTestDormant As Integer) As Integer
'
'   iRet = mReadBvfRec (iMnfCode, ilYears)
'   Where:
'       ilMnfCode(I)-Budget Name Code
'       ilYears(I)-Year to retrieve
'       ilYear(I)-New year to create
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim ilRecOK As Integer
    Dim ilVeh As Integer
    Dim ilFound As Integer
    Dim slDate As String
    Dim slStart As String
    Dim slEnd As String
    Dim llStart As Long
    Dim llEnd As Long
    Dim llEDate As Long
    Dim llLDate As Long
    Dim slSplit As String
    Dim ilSaleOffice As Integer
    Dim ilAddRec As Integer
    Dim ilVef As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    'ReDim tgBvfRec(1 To 1) As BVFREC
    ReDim tgBvfRec(0 To 1) As BVFREC

    ilUpper = UBound(tgBvfRec)
    btrExtClear hmBvf   'Clear any previous extend operation
    ilExtLen = Len(tgBvfRec(1).tBvf)  'Extract operation record size
    tmBvfSrchKey.iYear = ilYear
    tmBvfSrchKey.iSeqNo = 1
    tmBvfSrchKey.iMnfBudget = ilMnfCode
    ilRet = btrGetGreaterOrEqual(hmBvf, tgBvfRec(1).tBvf, imBvfRecLen, tmBvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'ilRet = btrGetFirst(hmBvf, tgBvfRec(1).tBvf, imBvfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmBvf, llNoRec, -1, "UC", "BVF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Bvf", "BvfMnfBudget")
        tlIntTypeBuff.iType = ilMnfCode
        ilRet = btrExtAddLogicConst(hmBvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        On Error GoTo mReadBvfRecErr
        gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddLogicConst):" & "Bvf.Btr", Budget
        On Error GoTo 0
        ilOffSet = gFieldOffset("Bvf", "BvfYear")
        tlIntTypeBuff.iType = ilYear
        ilRet = btrExtAddLogicConst(hmBvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        On Error GoTo mReadBvfRecErr
        gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddLogicConst):" & "Bvf.Btr", Budget
        On Error GoTo 0
        ilRet = btrExtAddField(hmBvf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mReadBvfRecErr
        gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddField):" & "Bvf.Btr", Budget
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmBvf, tgBvfRec(ilUpper).tBvf, ilExtLen, tgBvfRec(ilUpper).lRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadBvfRecErr
            gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtGetNextExt):" & "Bvf.Btr", Budget
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tgBvfRec(1).tBvf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmBvf, tgBvfRec(ilUpper).tBvf, ilExtLen, tgBvfRec(ilUpper).lRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'User allow to see vehicle
                ilRecOK = False
                'For ilVeh = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
                For ilVeh = imLBUserVeh To UBound(tmUserVeh) - 1 Step 1
                    If tmUserVeh(ilVeh).iCode = tgBvfRec(ilUpper).tBvf.iVefCode Then
                        ilRecOK = True
                        If tmVef.iCode <> tmUserVeh(ilVeh).iCode Then
                            'tmVefSrchKey.iCode = tmUserVeh(ilVeh).iCode
                            'ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    ilRecOk = False
                            'End If
                            tmVef.iCode = tmUserVeh(ilVeh).iCode
                            tmVef.sName = tmUserVeh(ilVeh).sName
                            tmVef.sState = tmUserVeh(ilVeh).sState
                            tmVef.iSort = tmUserVeh(ilVeh).iSort
                        End If
                        Exit For
                    End If
                Next ilVeh
                If (ilRecOK) And (ilTestDormant) And (tmVef.sState = "D") Then
                    'Test if any Dollars- if so include
                    ilRecOK = False
                    For ilLoop = LBound(tgBvfRec(ilUpper).tBvf.lGross) To UBound(tgBvfRec(ilUpper).tBvf.lGross) Step 1
                        If tgBvfRec(ilUpper).tBvf.lGross(ilLoop) <> 0 Then
                            ilRecOK = True
                            'For ilVeh = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
                            For ilVeh = imLBUserVeh To UBound(tmUserVeh) - 1 Step 1
                                If tmUserVeh(ilVeh).iCode = tmVef.iCode Then
                                    tmUserVeh(ilVeh).sState = "A"
                                    tmVef.sState = "A"
                                    Exit For
                                End If
                            Next
                            Exit For
                        End If
                    Next ilLoop
                End If
                If ilRecOK Then
                    slStr = ""
                    slStr = Trim$(Str$(tmVef.iSort))
                    Do While Len(slStr) < 5
                        slStr = "0" & slStr
                    Loop
                    tgBvfRec(ilUpper).sVehSort = slStr
                    tgBvfRec(ilUpper).sVehicle = tmVef.sName
                    ilRecOK = False
                    'For ilSaleOffice = LBound(tmSaleOffice) To UBound(tmSaleOffice) - 1 Step 1
                    For ilSaleOffice = imLBSalesOffice To UBound(tmSaleOffice) - 1 Step 1
                        If tmSaleOffice(ilSaleOffice).iCode = tgBvfRec(ilUpper).tBvf.iSofCode Then
                            If (ilTestDormant) And (tmSaleOffice(ilSaleOffice).sState = "D") Then
                                ilRecOK = False
                                For ilLoop = LBound(tgBvfRec(ilUpper).tBvf.lGross) To UBound(tgBvfRec(ilUpper).tBvf.lGross) Step 1
                                    If tgBvfRec(ilUpper).tBvf.lGross(ilLoop) <> 0 Then
                                        ilRecOK = True
                                        tmSaleOffice(ilSaleOffice).sState = "A"
                                        Exit For
                                    End If
                                Next ilLoop
                            Else
                                ilRecOK = True
                            End If
                            If ilRecOK Then
                                tgBvfRec(ilUpper).SOffice = tmSaleOffice(ilSaleOffice).sName
                                tgBvfRec(ilUpper).sMktRank = tmSaleOffice(ilSaleOffice).sMktRank
                            End If
                            Exit For
                        End If
                    Next ilSaleOffice
                End If
                If ilRecOK Then
                    If rbcSort(1).Value Or rbcSort(2).Value Then    'Vehicle within office
                        tgBvfRec(ilUpper).sKey = tgBvfRec(ilUpper).sMktRank & tgBvfRec(ilUpper).SOffice & tgBvfRec(ilUpper).sVehSort & tgBvfRec(ilUpper).sVehicle
                    ElseIf rbcSort(0).Value Or rbcSort(3).Value Then
                        tgBvfRec(ilUpper).sKey = tgBvfRec(ilUpper).sVehSort & tgBvfRec(ilUpper).sVehicle & tgBvfRec(ilUpper).sMktRank & tgBvfRec(ilUpper).SOffice
                    End If
                    tgBvfRec(ilUpper).iStatus = 1
                    ilUpper = ilUpper + 1
                    'ReDim Preserve tgBvfRec(1 To ilUpper) As BVFREC
                    ReDim Preserve tgBvfRec(0 To ilUpper) As BVFREC
                End If
                ilRet = btrExtGetNext(hmBvf, tgBvfRec(ilUpper).tBvf, ilExtLen, tgBvfRec(ilUpper).lRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmBvf, tgBvfRec(ilUpper).tBvf, ilExtLen, tgBvfRec(ilUpper).lRecPos)
                Loop
            Loop
        End If
    End If
    'Test if records missing
    ilUpper = UBound(tgBvfRec)
    'If ilUpper > LBound(tgBvfRec) Then
    If ilUpper > igLBBvfRec Then
        slSplit = tgBvfRec(1).tBvf.sSplit
    Else
        If igDirect = 0 Then
            slSplit = "D"
        Else
            slSplit = "S"
        End If
    End If
    'If ilUpper > LBound(tgBvfRec) Then  'If no records exist, wait to model to create

        If rbcShow(0).Value Then    'Corporate
            For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                If tgMCof(ilLoop).iYear = ilNewYear Then
                    If tgMCof(ilLoop).iStartMnthNo = 1 Then
                        slDate = Trim$(Str$(tgMCof(ilLoop).iStartMnthNo)) & "/15/" & Trim$(Str$(ilNewYear))
                    Else
                        slDate = Trim$(Str$(tgMCof(ilLoop).iStartMnthNo)) & "/15/" & Trim$(Str$(ilNewYear - 1))
                    End If
                    slStart = gObtainYearStartDate(5, slDate)
                    llStart = gDateValue(slStart)
                    If tgMCof(ilLoop).iStartMnthNo = 1 Then
                        slDate = Trim$("12/15/" & Trim$(Str$(ilYear)))
                    Else
                        slDate = Trim$(Str$(tgMCof(ilLoop).iStartMnthNo - 1)) & "/15/" & Trim$(Str$(ilNewYear))
                    End If
                    slEnd = gObtainEndCorp(slDate, True)
                    llEnd = gDateValue(slDate)
                    Exit For
                End If
            Next ilLoop
        Else
            slDate = "1/15/" & Trim$(Str$(ilNewYear))
            slStart = gObtainStartStd(slDate)
            llStart = gDateValue(slStart)
            slDate = "12/15/" & Trim$(Str$(ilNewYear))
            slEnd = gObtainEndStd(slDate)
            llEnd = gDateValue(slDate)
        End If
        'For ilVeh = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
        For ilVeh = imLBUserVeh To UBound(tmUserVeh) - 1 Step 1
            'For ilSaleOffice = LBound(tmSaleOffice) To UBound(tmSaleOffice) - 1 Step 1
            For ilSaleOffice = imLBSalesOffice To UBound(tmSaleOffice) - 1 Step 1
                If (ilTestDormant) And (tmSaleOffice(ilSaleOffice).sState = "D") Then
                    ilFound = True
                Else
                    ilFound = False
                    'For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                    For ilLoop = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                        If (tgBvfRec(ilLoop).tBvf.iSofCode = tmSaleOffice(ilSaleOffice).iCode) Then
                            If (tgBvfRec(ilLoop).tBvf.iVefCode = tmUserVeh(ilVeh).iCode) Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                End If
                If (Not ilFound) Then
                    'Test if vehicle has LCF within year
                    ilAddRec = True
                    'Exclude Sport date checking
                    ilVef = gBinarySearchVef(tmUserVeh(ilVeh).iCode)
                    If ilVef <> -1 Then
                        If tgMVef(ilVef).sType <> "R" Then
                            ilRet = gLCFTFNExist(hmLcf, "C", tmUserVeh(ilVeh).iCode)
                            If Not ilRet Then
                                ilAddRec = False
                                llEDate = gGetEarliestLCFDate(hmLcf, "C", tmUserVeh(ilVeh).iCode)
                                If llEDate > 0 Then
                                    llLDate = gGetLatestLCFDate(hmLcf, "C", tmUserVeh(ilVeh).iCode)
                                    If (llEDate <= llEnd) And (llLDate >= llStart) Then
                                        ilAddRec = True
                                    Else
                                        Exit For    'Don't need to test this vehicle for each sales office
                                    End If
                                End If
                            End If
                        End If
                    Else
                        ilAddRec = False
                    End If
                    If ilAddRec Then
                        tgBvfRec(ilUpper).tBvf.iSofCode = tmSaleOffice(ilSaleOffice).iCode
                        tgBvfRec(ilUpper).tBvf.iVefCode = tmUserVeh(ilVeh).iCode
                        tgBvfRec(ilUpper).tBvf.iYear = ilNewYear    'tgBvfRec(LBound(tgBvfRec)).tBvf.iYear
                        'tgBvfRec(ilUpper).tBvf.iSeqNo = tgBvfRec(LBound(tgBvfRec)).tBvf.iSeqNo
                        tgBvfRec(ilUpper).tBvf.iSeqNo = tgBvfRec(igLBBvfRec).tBvf.iSeqNo
                        'tgBvfRec(ilUpper).tBvf.iStartDate(0) = tgBvfRec(LBound(tgBvfRec)).tBvf.iStartDate(0)
                        tgBvfRec(ilUpper).tBvf.iStartDate(0) = tgBvfRec(igLBBvfRec).tBvf.iStartDate(0)
                        'tgBvfRec(ilUpper).tBvf.iStartDate(1) = tgBvfRec(LBound(tgBvfRec)).tBvf.iStartDate(1)
                        tgBvfRec(ilUpper).tBvf.iStartDate(1) = tgBvfRec(igLBBvfRec).tBvf.iStartDate(1)
                        tgBvfRec(ilUpper).tBvf.iMnfBudget = ilMnfCode    'tgBvfRec(LBound(tgBvfRec)).tBvf.iYear
                        tgBvfRec(ilUpper).tBvf.sSplit = slSplit
                        For ilLoop = LBound(tgBvfRec(ilUpper).tBvf.lGross) To UBound(tgBvfRec(ilUpper).tBvf.lGross) Step 1
                            'slStr = ""
                            'gStrToPDN slStr, 2, 5, tgBvfRec(ilUpper).tBvf.sGross(ilLoop)
                            tgBvfRec(ilUpper).tBvf.lGross(ilLoop) = 0
                        Next ilLoop
                        If tmVef.iCode <> tmUserVeh(ilVeh).iCode Then
                            'tmVefSrchKey.iCode = tmUserVeh(ilVeh).iCode
                            'ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            tmVef.iCode = tmUserVeh(ilVeh).iCode
                            tmVef.sName = tmUserVeh(ilVeh).sName
                            tmVef.sState = tmUserVeh(ilVeh).sState
                            tmVef.iSort = tmUserVeh(ilVeh).iSort
                        End If
                        If (Not ilTestDormant) Or (tmVef.sState <> "D") Then
                            slStr = ""
                            slStr = Trim$(Str$(tmVef.iSort))
                            Do While Len(slStr) < 5
                                slStr = "0" & slStr
                            Loop
                            tgBvfRec(ilUpper).sVehSort = slStr
                            tgBvfRec(ilUpper).sVehicle = tmVef.sName
                            tgBvfRec(ilUpper).SOffice = tmSaleOffice(ilSaleOffice).sName
                            tgBvfRec(ilUpper).sMktRank = tmSaleOffice(ilSaleOffice).sMktRank
                            If rbcSort(1).Value Or rbcSort(2).Value Then    'Vehicle within office
                                tgBvfRec(ilUpper).sKey = tgBvfRec(ilUpper).sMktRank & tgBvfRec(ilUpper).SOffice & tgBvfRec(ilUpper).sVehSort & tgBvfRec(ilUpper).sVehicle
                            ElseIf rbcSort(0).Value Or rbcSort(3).Value Then
                                tgBvfRec(ilUpper).sKey = tgBvfRec(ilUpper).sVehSort & tgBvfRec(ilUpper).sVehicle & tgBvfRec(ilUpper).sMktRank & tgBvfRec(ilUpper).SOffice
                            End If
                            tgBvfRec(ilUpper).iStatus = 0
                            imBDChg = True
                            ilUpper = ilUpper + 1
                            'ReDim Preserve tgBvfRec(1 To ilUpper) As BVFREC
                            ReDim Preserve tgBvfRec(0 To ilUpper) As BVFREC
                        End If
                    End If
                End If
            Next ilSaleOffice
        Next ilVeh
    'End If
    If ilUpper > 1 Then
        ArraySortTyp fnAV(tgBvfRec(), 1), UBound(tgBvfRec) - 1, 0, LenB(tgBvfRec(1)), 0, LenB(tgBvfRec(1).sKey), 0
    End If
    'mInitBudgetCtrls
    mReadBvfRec = True
    Exit Function
mReadBvfRecErr:
    On Error GoTo 0
    mReadBvfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaleOfficePop                  *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSaleOfficePop()
    Dim ilRet As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilLoop As Integer
    Dim slStr As String
    'ilRet = gPopOfficeSourceBox(Budget, cbcCtrl, lbcSalesOfficeCode)
    ilRet = gPopOfficePlusBox(Budget, tgSalesOfficeCode(), smSalesOfficeCodeTag)
    'ilRet = gPopUserVehComboBox(Budget, cbcCtrl, lbcSaleOfficeCode, lbcCombo)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSaleOfficePopErr
        'gCPErrorMsg ilRet, "mSaleOfficePop (gPopUserVehComboBox: Vehicle/Combo)", Budget
        gCPErrorMsg ilRet, "mSaleOfficePop (gPopOfficeSourceBox: Vehicle)", Budget
        On Error GoTo 0
    End If
    'ReDim tmSaleOffice(1 To lbcSalesOfficeCode.ListCount + 1) As SALEOFFICE
    'ReDim tmSaleOffice(1 To UBound(tgSalesOfficeCode) + 1) As SALEOFFICE
    ReDim tmSaleOffice(0 To UBound(tgSalesOfficeCode) + 1) As SALEOFFICE
    For ilLoop = 0 To UBound(tgSalesOfficeCode) - 1 Step 1 'lbcSalesOfficeCode.ListCount - 1 Step 1
        slNameCode = tgSalesOfficeCode(ilLoop).sKey    'lbcSalesOfficeCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slStr)   'tmSaleOffice(ilLoop + 1).sName)
        ilRet = gParseItem(slStr, 2, "|", tmSaleOffice(ilLoop + 1).sName)
        ilRet = gParseItem(slStr, 3, "|", tmSaleOffice(ilLoop + 1).sState)
        ilRet = gParseItem(slStr, 1, "|", tmSaleOffice(ilLoop + 1).sMktRank)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmSaleOffice(ilLoop + 1).iCode = Val(slCode)
    Next ilLoop
    Exit Sub
mSaleOfficePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilLoop As Integer   'For loop control
    Dim ilRet As Integer
    Dim slNameFac As String
    Dim slMsg As String
    Dim ilBvf As Integer
    Dim ilBsf As Integer
    Dim ilNewBsf As Integer
    Dim ilNewBvf As Integer
    Dim tlBvf As BVF
    Dim tlBsf As BSF
    Dim tlBvf1 As MOVEREC
    Dim tlBvf2 As MOVEREC
    Dim tlBsf1 As MOVEREC
    Dim tlBsf2 As MOVEREC
    mBDSetShow imBDBoxNo
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
    If igBDView = 0 Then
        'Check if Budget name exist or should be added
        ilLoop = 0
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            Do  'Loop until record updated or added
                'StartDate is currently not used
                gPackDate "", tgBvfRec(ilBvf).tBvf.iStartDate(0), tgBvfRec(ilBvf).tBvf.iStartDate(1)
                If (imSelectedIndex = 0) Or (tgBvfRec(ilBvf).iStatus = 0) Then  'New selected
                    ilNewBvf = True
                    tgBvfRec(ilBvf).tBvf.lCode = 0
                    tgBvfRec(ilBvf).tBvf.iUrfCode = tgUrf(0).iCode
                    ilRet = btrInsert(hmBvf, tgBvfRec(ilBvf).tBvf, imBvfRecLen, INDEXKEY0)
                    tgBvfRec(ilBvf).iStatus = 1
                    ilRet = btrGetPosition(hmBvf, tgBvfRec(ilBvf).lRecPos)
                    slMsg = "mSaveRec (btrInsert: Budget)"
                    ilLoop = ilLoop + 1
                Else 'Old record-Update
                    ilNewBvf = False
                    slMsg = "mSaveRec (btrGetDirect: Budget)"
                    ilRet = btrGetDirect(hmBvf, tlBvf, imBvfRecLen, tgBvfRec(ilBvf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, Budget
                    On Error GoTo 0
                    'tmRec = tlBvf
                    'ilRet = gGetByKeyForUpdate("BVF", hmBvf, tmRec)
                    'tlBvf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                    '    imTerminate = True
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                    LSet tlBvf1 = tlBvf
                    LSet tlBvf2 = tgBvfRec(ilBvf).tBvf
                    If StrComp(tlBvf1.sChar, tlBvf2.sChar, 0) <> 0 Then
                        tgBvfRec(ilBvf).tBvf.iUrfCode = tgUrf(0).iCode
                        ilRet = btrUpdate(hmBvf, tgBvfRec(ilBvf).tBvf, imBvfRecLen)
                    Else
                        ilRet = BTRV_ERR_NONE
                    End If
                    slMsg = "mSaveRec (btrUpdate: Budget)"
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Budget
            On Error GoTo 0
        Next ilBvf
    Else
        'For ilBsf = LBound(tgBsfRec) To UBound(tgBsfRec) - 1 Step 1
        For ilBsf = igLBBsfRec To UBound(tgBsfRec) - 1 Step 1
            Do  'Loop until record updated or added
                'StartDate is currently not used
                gPackDate "", tgBsfRec(ilBsf).tBsf.iStartDate(0), tgBsfRec(ilBsf).tBsf.iStartDate(1)
                If (imSelectedIndex = 0) Or (tgBsfRec(ilBsf).iStatus = 0) Then  'New selected
                    ilNewBsf = True
                    tgBsfRec(ilBsf).tBsf.iUrfCode = tgUrf(0).iCode
                    tgBsfRec(ilBsf).tBsf.lCode = 0
                    ilRet = btrInsert(hmBsf, tgBsfRec(ilBsf).tBsf, imBsfReclen, INDEXKEY1)
                    tgBsfRec(ilBsf).iStatus = 1
                    ilRet = btrGetPosition(hmBsf, tgBsfRec(ilBsf).lRecPos)
                    slMsg = "mSaveRec (btrInsert: Budget)"
                Else 'Old record-Update
                    ilNewBsf = False
                    slMsg = "mSaveRec (btrGetDirect: Budget)"
                    ilRet = btrGetDirect(hmBsf, tlBsf, imBsfReclen, tgBsfRec(ilBsf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, Budget
                    On Error GoTo 0
                    'tmRec = tlBsf
                    'ilRet = gGetByKeyForUpdate("BSF", hmBsf, tmRec)
                    'tlBsf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    Screen.MousePointer = vbDefault    'Default
                    '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                    '    imTerminate = True
                    '    mSaveRec = False
                    '    Exit Function
                    'End If
                    LSet tlBsf1 = tlBsf
                    LSet tlBsf2 = tgBsfRec(ilBsf).tBsf
                    If StrComp(tlBsf1.sChar, tlBsf2.sChar, 0) <> 0 Then
                        tgBsfRec(ilBsf).tBsf.iUrfCode = tgUrf(0).iCode
                        ilRet = btrUpdate(hmBsf, tgBsfRec(ilBsf).tBsf, imBsfReclen)
                    Else
                        ilRet = BTRV_ERR_NONE
                    End If
                    slMsg = "mSaveRec (btrUpdate: Budget)"
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Budget
            On Error GoTo 0
        Next ilBsf
    End If
    If imSelectedIndex > 0 Then
        slNameFac = cbcSelect.List(imSelectedIndex)
    Else
        slNameFac = ""
    End If
    mPopulate
    gFindMatch slNameFac, 0, cbcSelect
    If gLastFound(cbcSelect) > 0 Then
        If cbcSelect.ListIndex = gLastFound(cbcSelect) Then
            'cbcSelect_Change
        Else
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        End If
    End If
    mSaveRec = True
    Screen.MousePointer = vbDefault    'Default
    Exit Function
mSaveRecErr:
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
    If imBDChg Then
        If ilAsk Then
            If imSelectedIndex > 0 Then
                slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
            Else
                slMess = "Add " '& tgRcfI.sName
            End If
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                If igBDView = 0 Then
                    pbcOffice_Paint
                    pbcTotals_Paint
                Else
                    pbcSalesperson_Paint
                End If
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
    If imBDChg Then
        cbcSelect.Enabled = False
        rbcOS(0).Enabled = False
        rbcOS(1).Enabled = False
        pbcType.Enabled = False
        If igBDView = 0 Then
            If (UBound(tgBvfRec) > 1) And (imUpdateAllowed) Then
                cmcUpdate.Enabled = True
            Else
                cmcUpdate.Enabled = False
            End If
        Else
            If (UBound(tgBsfRec) > 1) And (imUpdateAllowed) Then
                cmcUpdate.Enabled = True
            Else
                cmcUpdate.Enabled = False
            End If
        End If
    Else
        cbcSelect.Enabled = True
        cmcUpdate.Enabled = False
        rbcOS(0).Enabled = True
        rbcOS(1).Enabled = True
        pbcType.Enabled = True
    End If
    If imSelectedIndex <= 0 Then
        'pbcYear.Enabled = False
        'pbcBudgetName.Enabled = False
        'pbcOffice.Enabled = False
        'pbcSalesperson.Enabled = False
        'pbcOSSTab.Enabled = False
        'pbcOSTab.Enabled = False
        'pbcSTab.Enabled = False
        'pbcTab.Enabled = False
    Else
        If (igWinStatus(BUDGETSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            pbcYear.Enabled = True
            pbcBudgetName.Enabled = True
            pbcOffice.Enabled = True
            pbcSalesperson.Enabled = True
            pbcOSSTab.Enabled = True
            pbcOSTab.Enabled = True
            pbcSTab.Enabled = True
            pbcTab.Enabled = True
        End If
    End If
    'Revert button set if any field changed
    If imBDChg Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
    'If imUpdateAllowed And rbcOS(0).Value And (UBound(tgBvfRec) > LBound(tgBvfRec)) Then
    If imUpdateAllowed And rbcOS(0).Value And (UBound(tgBvfRec) > igLBBvfRec) Then
        If igBudgetType = 0 Then
            cmcScale.Enabled = True
            cmcTrend.Enabled = True
            cmcActuals.Enabled = True
            cmcAdvt.Enabled = True
            cmcDemo.Enabled = True
            cmc12Mos.Enabled = True
            cmcAddVeh.Enabled = True
        Else
            cmcScale.Enabled = False
            cmcTrend.Enabled = False
            cmcActuals.Enabled = True
            cmcAdvt.Enabled = False
            cmcDemo.Enabled = False
            cmc12Mos.Enabled = True
            cmcAddVeh.Enabled = True
        End If
    Else
        cmcScale.Enabled = False
        cmcTrend.Enabled = False
        cmcActuals.Enabled = False
        cmcAdvt.Enabled = False
        cmcDemo.Enabled = False
        cmc12Mos.Enabled = False
        cmcAddVeh.Enabled = False
    End If
End Sub
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
    If (ilBoxNo < imLBBDCtrls) Or (ilBoxNo > UBound(tmBDCtrls)) Then
        Exit Sub
    End If

    If (imBDRowNo < vbcBudget.Value) Or (imBDRowNo >= vbcBudget.Value + vbcBudget.LargeChange + 1) Then
        mBDSetShow ilBoxNo
        pbcArrow.Visible = False
        lacSFrame.Visible = False
        lacOFrame.Visible = False
        Exit Sub
    End If
    If igBDView = 0 Then
        lacOFrame.Move 0, tmBDCtrls(OSNAMEINDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15) - 30
        lacOFrame.Visible = True
    Else
        lacSFrame.Move 0, tmBDCtrls(OSNAMEINDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15) - 30
        lacSFrame.Visible = True
    End If
    pbcArrow.Move pbcArrow.Left, plcBudget.Top + tmBDCtrls(OSNAMEINDEX).fBoxY + (imBDRowNo - vbcBudget.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case OSNAMEINDEX 'Name
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
Private Sub mSetPrice(ilGroup As Integer, ilRowNo As Integer, slInDollar As String)
    Dim ilWk As Integer
    Dim ilBvf As Integer
    Dim ilBsf As Integer
    'Dim slNoWks As String
    'Dim slNoWks1 As String
    'Dim slAvgDollar As String
    'Dim slEndDollar As String
    Dim llAvgDollar As Long
    Dim llEndDollar As Long
    Dim llDollar As Long
    Dim llInDollar As Long
    Dim ilStIndex As Integer
    If (igBDView = 0) And (rbcSort(0).Value Or rbcSort(1).Value) Then
        Exit Sub
    End If
    llInDollar = Val(slInDollar)
    If igBDView = 0 Then
        ilBvf = imBDSave(1, ilRowNo)
        If ilBvf <= 0 Then
            Exit Sub
        End If
        ilStIndex = 1
        If rbcSort(0).Value Then    'Vehicle
            llInDollar = lmTSave(ilGroup, ilStIndex)
        ElseIf rbcSort(1).Value Then    'Office
            llInDollar = lmTSave(ilGroup, ilStIndex)
        End If
        Do
            If tgBvfRec(ilBvf).tBvf.iYear = tmPdGroups(ilGroup).iYear Then
                If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                    'slNoWks = Trim$(Str$(tmPdGroups(ilGroup).iTrueNoWks))
                    'slNoWks1 = Trim$(Str$(tmPdGroups(ilGroup).iTrueNoWks - 1))
                    'slAvgDollar = gDivStr(slInDollar, slNoWks)
                    llAvgDollar = llInDollar / tmPdGroups(ilGroup).iTrueNoWks
                    'slEndDollar = gSubStr(slInDollar, gMulStr(slAvgDollar, slNoWks1))
                    llEndDollar = llInDollar - llAvgDollar * (tmPdGroups(ilGroup).iTrueNoWks - 1)
                    For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                        'slDollar = slAvgDollar
                        llDollar = llAvgDollar
                        If ilWk = tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Then
                            'slDollar = slEndDollar
                            llDollar = llEndDollar
                        End If
                        If ilWk = 1 Then
                            If rbcShow(0).Value Then    'Don't split if entered for corporate
                                'slStr = ""
                                'gStrToPDN slStr, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(0)
                                'gStrToPDN slDollar, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(ilWk)
                                tgBvfRec(ilBvf).tBvf.lGross(0) = 0
                                tgBvfRec(ilBvf).tBvf.lGross(ilWk) = llDollar
                            Else
                                'slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                                'slStart = gObtainStartCorp(slDate, True)
                                'ilDay = gWeekDayStr(slStart)
                                'If ilDay = 0 Then
                                    'slStr = ""
                                    'gStrToPDN slStr, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(0)
                                    'gStrToPDN slDollar, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(ilWk)
                                    tgBvfRec(ilBvf).tBvf.lGross(0) = 0
                                    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = llDollar
                                'Else
                                '    'slStr = gDivStr(gMulStr(slDollar, Trim$(Str$(ilDay))), "7")
                                '    'gStrToPDN slStr, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(0)
                                '    tgBvfRec(ilBvf).tBvf.lGross(0) = (llDollar * ilDay) / 7
                                '    'slStr = gSubStr(slDollar, slStr)
                                '    'gStrToPDN slStr, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(ilWk)
                                '    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = llDollar - tgBvfRec(ilBvf).tBvf.lGross(0)
                                'End If
                            End If
                        'ElseIf ilWk = 52 Then
                        '    If rbcShow(1).Value Then    'Don't split dollars if input via standard
                        '        'slStr = ".00"
                        '        'gStrToPDN slStr, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(53)
                        '        'gStrToPDN slDollar, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(ilWk)
                        '        tgBvfRec(ilBvf).tBvf.lGross(53) = 0
                        '        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = llDollar
                        '    Else
                        '        'slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                        '        'slStart = gObtainEndCorp(slDate, True)
                        '        'ilDay = gWeekDayStr(slStart)
                        '        'If ilDay = 6 Then
                        '            'slStr = ".00"
                        '            'gStrToPDN slStr, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(53)
                        '            'gStrToPDN slDollar, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(ilWk)
                        '            tgBvfRec(ilBvf).tBvf.lGross(53) = 0
                        '            tgBvfRec(ilBvf).tBvf.lGross(ilWk) = llDollar
                        '        'Else
                        '        '    ilDay = 7 - ilDay - 1
                        '        '    'slStr = gDivStr(gMulStr(slDollar, Trim$(Str$(ilDay))), Trim$(Str$(ilDay)))
                        '        '    'gStrToPDN slStr, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(52)
                        '        '    tgBvfRec(ilBvf).tBvf.lGross(52) = (llDollar * ilDay) / 7
                        '        '    'slStr = gSubStr(slDollar, slStr)
                        '        '    'gStrToPDN slStr, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(53)
                        '        '    tgBvfRec(ilBvf).tBvf.lGross(53) = llDollar - tgBvfRec(ilBvf).tBvf.lGross(52)
                        '        'End If
                        '    End If
                        Else
                            'gStrToPDN slDollar, 2, 5, tgBvfRec(ilBvf).tBvf.sGross(ilWk)
                            tgBvfRec(ilBvf).tBvf.lGross(ilWk) = llDollar
                        End If
                    Next ilWk
                End If
            End If
            If rbcSort(2).Value Or rbcSort(3).Value Then
                Exit Do
            ElseIf rbcSort(0).Value Then    'Vehicle
                If tgBvfRec(ilBvf).tBvf.iVefCode <> tgBvfRec(ilBvf + 1).tBvf.iVefCode Then
                    ilStIndex = 0
                    Exit Do
                Else
                    ilBvf = ilBvf + 1
                    ilStIndex = ilStIndex + 1
                    llInDollar = lmTSave(ilGroup, ilStIndex)
                End If
            ElseIf rbcSort(1).Value Then    'Office
                If tgBvfRec(ilBvf).tBvf.iSofCode <> tgBvfRec(ilBvf + 1).tBvf.iSofCode Then
                    ilStIndex = 0
                    Exit Do
                Else
                    ilBvf = ilBvf + 1
                    ilStIndex = ilStIndex + 1
                    llInDollar = lmTSave(ilGroup, ilStIndex)
                End If
            End If
            If ilBvf = UBound(tgBvfRec) Then
                ilStIndex = 0
                Exit Do
            End If
        Loop
    Else
        ilBsf = imBDSave(1, ilRowNo)
        If ilBsf <= 0 Then
            Exit Sub
        End If
        If tgBsfRec(ilBsf).tBsf.iYear = tmPdGroups(ilGroup).iYear Then
            If tmPdGroups(ilGroup).iTrueNoWks > 0 Then
                'slNoWks = Trim$(Str$(tmPdGroups(ilGroup).iTrueNoWks))
                'slNoWks1 = Trim$(Str$(tmPdGroups(ilGroup).iTrueNoWks - 1))
                'slAvgDollar = gDivStr(slInDollar, slNoWks)
                llAvgDollar = llInDollar / tmPdGroups(ilGroup).iTrueNoWks
                'slEndDollar = gSubStr(slInDollar, gMulStr(slAvgDollar, slNoWks1))
                llEndDollar = llInDollar - llAvgDollar * (tmPdGroups(ilGroup).iTrueNoWks - 1)
                For ilWk = tmPdGroups(ilGroup).iStartWkNo To tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Step 1
                    'slDollar = slAvgDollar
                    llDollar = llAvgDollar
                    If ilWk = tmPdGroups(ilGroup).iStartWkNo + tmPdGroups(ilGroup).iTrueNoWks - 1 Then
                        'slDollar = slEndDollar
                        llDollar = llEndDollar
                    End If
                    If ilWk = 1 Then
                        If rbcShow(0).Value Then    'If input by corporate, then don't split
                            'slStr = ".00"
                            'gStrToPDN slStr, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(0)
                            'gStrToPDN slDollar, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(ilWk)
                            tgBsfRec(ilBsf).tBsf.lGross(0) = 0
                            tgBsfRec(ilBsf).tBsf.lGross(ilWk) = llDollar
                        Else
                            'slDate = "1/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                            'slStart = gObtainStartCorp(slDate, True)
                            'ilDay = gWeekDayStr(slStart)
                            'If ilDay = 0 Then
                                'slStr = ".00"
                                'gStrToPDN slStr, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(0)
                                'gStrToPDN slDollar, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(ilWk)
                                tgBsfRec(ilBsf).tBsf.lGross(0) = 0
                                tgBsfRec(ilBsf).tBsf.lGross(ilWk) = llDollar
                            'Else
                            '    'slStr = gDivStr(gMulStr(slDollar, Trim$(Str$(ilDay))), "7")
                            '    'gStrToPDN slStr, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(0)
                            '    tgBsfRec(ilBsf).tBsf.lGross(0) = (llDollar * ilDay) / 7
                            '    'slStr = gSubStr(slDollar, slStr)
                            '    'gStrToPDN slStr, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(ilWk)
                            '    tgBsfRec(ilBsf).tBsf.lGross(ilWk) = llDollar - tgBsfRec(ilBsf).tBsf.lGross(0)
                            'End If
                        End If
                    'ElseIf ilWk = 52 Then
                    '    If rbcShow(1).Value Then    'Don't split dollars if input via standard
                    '        'slStr = ".00"
                    '        'gStrToPDN slStr, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(53)
                    '        'gStrToPDN slDollar, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(ilWk)
                    '        tgBsfRec(ilBsf).tBsf.lGross(53) = 0
                    '        tgBsfRec(ilBsf).tBsf.lGross(ilWk) = llDollar
                    '    Else
                    '        'slDate = "12/15/" & Trim$(Str$(tmPdGroups(1).iYear))
                    '        'slStart = gObtainEndCorp(slDate, True)
                    '        'ilDay = gWeekDayStr(slStart)
                    '        'If ilDay = 6 Then
                    '        '    'slStr = ".00"
                    '        '    'gStrToPDN slStr, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(53)
                    '        '    'gStrToPDN slDollar, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(ilWk)
                    '            tgBsfRec(ilBsf).tBsf.lGross(53) = 0
                    '            tgBsfRec(ilBsf).tBsf.lGross(ilWk) = llDollar
                    '        'Else
                    '        '    ilDay = 7 - ilDay - 1
                    '        '    'slStr = gDivStr(gMulStr(slDollar, Trim$(Str$(ilDay))), Trim$(Str$(ilDay)))
                    '        '    'gStrToPDN slStr, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(52)
                    '        '    tgBsfRec(ilBsf).tBsf.lGross(52) = (llDollar * ilDay) / 7
                    '        '    'slStr = gSubStr(slDollar, slStr)
                    '        '    'gStrToPDN slStr, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(53)
                    '        '    tgBsfRec(ilBsf).tBsf.lGross(53) = llDollar - tgBsfRec(ilBsf).tBsf.lGross(52)
                    '        'End If
                    '    End If
                    Else
                        'gStrToPDN slDollar, 2, 5, tgBsfRec(ilBsf).tBsf.sGross(ilWk)
                        tgBsfRec(ilBsf).tBsf.lGross(ilWk) = llDollar
                    End If
                Next ilWk
            End If
        End If
    End If
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
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slName As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilModelNoWks As Integer
    Dim ilNewNoWks As Integer
    'If Not gWinRoom(igNoExeWinRes(RCTERMSEXE)) Then
    '    mStartNew = False
    '    Exit Function
    'End If
    imInNew = True
    igNewMnfBudget = 0
    igNewYear = 0
    'If cbcSelect.ListCount > 1 Then
        BudModel.Show vbModal
        DoEvents
        If igBDReturn = 0 Then    'Cancelled
            mStartNew = False
            imInNew = False
            Exit Function
        End If
    'Else
    '    igRcfModel = 0
    'End If
    Screen.MousePointer = vbHourglass    '
    'Build program images from newest
    If igBDView = 0 Then
        ilRet = mReadBvfRec(igModelMnfBudget, igModelYear, igNewYear, True)   'Ok to pass zero ([None])
        If Not ilRet Then
            mStartNew = False
            imInNew = False
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        If igModelYear <> 0 Then
            If tgSpf.sRUseCorpCal = "Y" Then
                slStartDate = gObtainYearStartDate(4, "1/15/" & Trim$(Str$(igModelYear)))
                slEndDate = gObtainYearEndDate(4, "1/15/" & Trim$(Str$(igModelYear)))
                ilModelNoWks = gDateValue(slEndDate) - gDateValue(slStartDate)
                slStartDate = gObtainYearStartDate(4, "1/15/" & Trim$(Str$(igNewYear)))
                slEndDate = gObtainYearEndDate(4, "1/15/" & Trim$(Str$(igNewYear)))
                ilNewNoWks = gDateValue(slEndDate) - gDateValue(slStartDate)
            Else
                slStartDate = gObtainYearStartDate(0, "1/15/" & Trim$(Str$(igModelYear)))
                slEndDate = gObtainYearEndDate(0, "1/15/" & Trim$(Str$(igModelYear)))
                ilModelNoWks = gDateValue(slEndDate) - gDateValue(slStartDate)
                slStartDate = gObtainYearStartDate(0, "1/15/" & Trim$(Str$(igNewYear)))
                slEndDate = gObtainYearEndDate(0, "1/15/" & Trim$(Str$(igNewYear)))
                ilNewNoWks = gDateValue(slEndDate) - gDateValue(slStartDate)
            End If
        Else
            ilModelNoWks = 53
            ilNewNoWks = 53
        End If
        'For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilLoop = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            tgBvfRec(ilLoop).tBvf.iMnfBudget = igNewMnfBudget
            tgBvfRec(ilLoop).tBvf.iYear = igNewYear
            tgBvfRec(ilLoop).tBvf.iSeqNo = 1
            If igDirect = 0 Then
                tgBvfRec(ilLoop).tBvf.sSplit = "D"
            Else
                tgBvfRec(ilLoop).tBvf.sSplit = "S"
            End If
            If ilModelNoWks < ilNewNoWks Then
                tgBvfRec(ilLoop).tBvf.lGross(53) = tgBvfRec(ilLoop).tBvf.lGross(52)
            ElseIf ilModelNoWks > ilNewNoWks Then
                tgBvfRec(ilLoop).tBvf.lGross(53) = 0
            End If
            tgBvfRec(ilLoop).tBvf.iUrfCode = tgUrf(0).iCode
            tgBvfRec(ilLoop).lRecPos = 0
            tgBvfRec(ilLoop).iStatus = 0
        Next ilLoop
        imMnfCode = igNewMnfBudget
        imYear = igNewYear
        smMnfName = sgBudgetName
    Else
        ilRet = mReadBsfRec(igModelYear, igNewYear)   'Ok to pass zero ([None])
        If Not ilRet Then
            mStartNew = False
            imInNew = False
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        'For ilLoop = LBound(tgBsfRec) To UBound(tgBsfRec) - 1 Step 1
        For ilLoop = igLBBsfRec To UBound(tgBsfRec) - 1 Step 1
            tgBsfRec(ilLoop).tBsf.iYear = igNewYear
            tgBsfRec(ilLoop).tBsf.iSeqNo = 1
            tgBsfRec(ilLoop).tBsf.iUrfCode = tgUrf(0).iCode
            tgBsfRec(ilLoop).lRecPos = 0
            tgBsfRec(ilLoop).iStatus = 0
        Next ilLoop
        imYear = igNewYear
    End If

    'plcSelect.Visible = False
    'plcOS.Visible = False
    'plcBudget.Visible = False
    mInitBudgetCtrls  'Initial arrays
    mMoveRecToCtrl
    'mInitShow
    imChgMode = True    'Set change mode to avoid infinite loop
    If igBDView = 0 Then
        slName = smMnfName & "/" & Trim$(Str$(imYear))
    Else
        slName = Trim$(Str$(imYear))
    End If
    cbcSelect.AddItem slName, 1
    cbcSelect.ListIndex = 1
    imSelectedIndex = 1
    plcSelect.Visible = True
    plcOS.Visible = True
    pbcOffice.Cls
    pbcTotals.Cls
    pbcBudgetName.Cls
    pbcYear.Cls
    pbcSalesperson.Cls
    If igBDView = 0 Then
        plcBudget.Visible = True
        pbcBudgetName_Paint
        pbcOffice_Paint
        pbcTotals_Paint
    Else
        plcBudget.Visible = True
        pbcYear_Paint
        pbcSalesperson_Paint
    End If
    imChgMode = False
    mStartNew = True
    mSetCommands
    Screen.MousePointer = vbDefault
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

    sgBudUserVehicleTag = ""


    imTerminate = False

    Screen.MousePointer = vbDefault
    Unload IconTraf
    igManUnload = YES
    Unload Budget
    igManUnload = NO
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
    mTestSaveFields = YES
End Function
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
    'ilRet = gPopUserVehicleBox(Budget, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, cbcCtrl, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(Budget, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + VEHNTR + ACTIVEVEH + DORMANTVEH, cbcCtrl, tgBudUserVehicle(), sgBudUserVehicleTag)
    'ilRet = gPopUserVehComboBox(Budget, cbcCtrl, Traffic!lbcUserVehicle, lbcCombo)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        'gCPErrorMsg ilRet, "mVehPop (gPopUserVehComboBox: Vehicle/Combo)", Budget
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", Budget
        On Error GoTo 0
    End If
    'ReDim tmUserVeh(1 To Traffic!lbcUserVehicle.ListCount + 1) As USERVEH
    'ReDim tmUserVeh(1 To UBound(tgBudUserVehicle) + 1) As BDUSERVEH
    ReDim tmUserVeh(0 To UBound(tgBudUserVehicle) + 1) As BDUSERVEH
    For ilLoop = 0 To UBound(tgBudUserVehicle) - 1 Step 1  'Traffic!lbcUserVehicle.ListCount - 1 Step 1
        slNameCode = tgBudUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", tmUserVeh(ilLoop + 1).sName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        tmUserVeh(ilLoop + 1).iCode = Val(slCode)
        tmVefSrchKey.iCode = tmUserVeh(ilLoop + 1).iCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        tmUserVeh(ilLoop + 1).sName = tmVef.sName
        tmUserVeh(ilLoop + 1).sState = tmVef.sState
        tmUserVeh(ilLoop + 1).iSort = tmVef.iSort
    Next ilLoop
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
Private Sub pbcBudgetName_Paint()
    Dim ilBox As Integer
    For ilBox = imLBOVCtrls To UBound(tmOVCtrls) Step 1
        pbcBudgetName.CurrentX = tmOVCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcBudgetName.CurrentY = tmOVCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcBudgetName.Print tmOVCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcClickFocus_GotFocus()
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
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
Private Sub pbcOffice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    Dim ilWk As Integer
    Dim ilMn As Integer
    Dim ilDone As Integer
    Dim ilStartWkNo As Integer
    'ReDim ilStartWk(1 To 12) As Integer
    ReDim ilStartWk(0 To 12) As Integer
    Dim ilLBStartWk As Integer
    'ReDim ilNoWks(1 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer
    Dim ilLBNoWks As Integer
    
    ilLBNoWks = 1
    ilLBStartWk = 1
    If Button = 2 Then  'Right Mouse
        Exit Sub
    End If
    'Check if hot spot
    If imInHotSpot Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        Exit Sub
    End If
    If UBound(tgBvfRec) > 1 Then
        'For ilLoop = LBound(imHotSpot, 1) To UBound(imHotSpot, 1) Step 1
        For ilLoop = LBound(imHotSpot, 1) + 1 To UBound(imHotSpot, 1) Step 1
            If (X >= imHotSpot(ilLoop, 1)) And (X <= imHotSpot(ilLoop, 3)) And (Y >= imHotSpot(ilLoop, 2)) And (Y <= imHotSpot(ilLoop, 4)) Then
                Screen.MousePointer = vbHourglass
                mBDSetShow imBDBoxNo
                imBDBoxNo = -1
                imInHotSpot = True
                Select Case ilLoop
                    Case 1  'Goto Start
                        imPdYear = imBDStartYear
                        imPdStartWk = 1
                    Case 2  'Reduce by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(1).iYear = imBDStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
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
                            If (tmPdGroups(1).iYear = imBDStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
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
                            If (tmPdGroups(1).iYear = imBDStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
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
                        End If
                    Case 3  'Increase by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(4).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(4).iStartWkNo > 39) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths tmPdGroups(4).iYear, ilStartWk(), ilNoWks()
                            If (tmPdGroups(4).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(4).iStartWkNo >= ilStartWk(12)) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(2).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(2).iStartWkNo >= ilStartWk(9)) Then 'At end
                                imPdYear = tmPdGroups(2).iYear
                                imPdStartWk = tmPdGroups(2).iStartWkNo
                            ElseIf (tmPdGroups(3).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(3).iStartWkNo >= ilStartWk(9)) Then 'At end
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
                            If (tmPdGroups(4).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(4).iStartWkNo >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(2).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(2).iStartWkNo + 3 >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imPdYear = tmPdGroups(2).iYear
                                imPdStartWk = tmPdGroups(2).iStartWkNo
                            ElseIf (tmPdGroups(3).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(3).iStartWkNo + 3 >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imPdYear = tmPdGroups(3).iYear
                                imPdStartWk = tmPdGroups(3).iStartWkNo
                            Else
                                imPdYear = tmPdGroups(4).iYear
                                imPdStartWk = tmPdGroups(4).iStartWkNo
                            End If
                            'imPdYear = tmPdGroups(2).iYear
                            'imPdStartWk = tmPdGroups(2).iStartWkNo
                        End If
                    Case 4  'GoTo End
                        imPdYear = imBDStartYear + imBDNoYears - 1
                        If rbcType(0).Value Then    'Quarter
                            imPdStartWk = 1
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(9)  'At end
                        ElseIf rbcType(2).Value Then    'Week
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(12) + ilNoWks(12) - 4
                        End If
                End Select
                pbcOffice.Cls
                pbcTotals.Cls
                mGetShowDates
                pbcOffice_Paint
                pbcTotals_Paint
                Screen.MousePointer = vbDefault
                imInHotSpot = False
                Exit Sub
            End If
        Next ilLoop
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub pbcOffice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        imButtonIndex = -1
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If cbcSelect.ListCount <= 1 Then
        Exit Sub
    End If
    ilCompRow = vbcBudget.LargeChange + 1
    'If UBound(tgBvfRec) > ilCompRow Then
    If UBound(smBDSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smBDSave, 2) + 1  'UBound(tgBvfRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBBDCtrls To UBound(tmBDCtrls) Step 1
            If (X >= tmBDCtrls(ilBox).fBoxX) And (X <= (tmBDCtrls(ilBox).fBoxX + tmBDCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmBDCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmBDCtrls(ilBox).fBoxY + tmBDCtrls(ilBox).fBoxH)) Then
                    If (ilBox < DOLLAR1INDEX) And (ilBox > DOLLAR4INDEX) Then    'DAYPARTINDEX Then
                        Beep
                        Exit Sub
                    End If
                    ilRowNo = ilRow + vbcBudget.Value - 1
                    If ilRowNo >= UBound(smBDSave, 2) Then
                        Beep
                        mSetFocus imBDBoxNo
                        Exit Sub
                    End If
                    If imBDSave(1, ilRowNo) <= 0 Then
                        Beep
                        mSetFocus imBDBoxNo
                        Exit Sub
                    End If
                    mBDSetShow imBDBoxNo
                    imBDRowNo = ilRow + vbcBudget.Value - 1
                    imBDBoxNo = ilBox
                    mBDEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBDBoxNo
End Sub
Private Sub pbcOffice_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    If imTerminate Then
        Exit Sub
    End If
    mPaintOfficeTitle
    llColor = pbcOffice.ForeColor
    slFontName = pbcOffice.FontName
    flFontSize = pbcOffice.FontSize
    pbcOffice.ForeColor = BLUE
    pbcOffice.FontBold = False
    pbcOffice.FontSize = 7
    pbcOffice.FontName = "Arial"
    pbcOffice.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilBox = imLBWKCtrls To UBound(tmWKCtrls) Step 1
        'gPaintArea pbcOffice, tmWKCtrls(ilBox).fBoxX, tmWKCtrls(ilBox).fBoxY, tmWKCtrls(ilBox).fBoxW - 15, tmWKCtrls(ilBox).fBoxH - 15, WHITE
        pbcOffice.CurrentX = tmWKCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcOffice.CurrentY = tmWKCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcOffice.Print tmWKCtrls(ilBox).sShow
    Next ilBox
    For ilBox = imLBNWCtrls To UBound(tmNWCtrls) Step 1
        'gPaintArea pbcOffice, tmNWCtrls(ilBox).fBoxX, tmNWCtrls(ilBox).fBoxY, tmNWCtrls(ilBox).fBoxW - 15, tmNWCtrls(ilBox).fBoxH - 15, LIGHTBLUE
        pbcOffice.CurrentX = tmNWCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcOffice.CurrentY = tmNWCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcOffice.Print tmNWCtrls(ilBox).sShow
    Next ilBox
    pbcOffice.FontSize = flFontSize
    pbcOffice.FontName = slFontName
    pbcOffice.FontSize = flFontSize
    pbcOffice.ForeColor = llColor
    pbcOffice.FontBold = True
    ilStartRow = vbcBudget.Value '+ 1  'Top location
    ilEndRow = vbcBudget.Value + vbcBudget.LargeChange ' + 1
    If ilEndRow > UBound(smBDSave, 2) Then
        ilEndRow = UBound(smBDSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcOffice.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow = UBound(smBDSave, 2) Then
            pbcOffice.ForeColor = DARKPURPLE
        Else
            pbcOffice.ForeColor = llColor
        End If
        For ilBox = imLBBDCtrls To UBound(tmBDCtrls) Step 1
            'If ilBox <> TOTALINDEX Then
            '    gPaintArea pbcOffice, tmBDCtrls(ilBox).fBoxX, tmBDCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmBDCtrls(ilBox).fBoxW - 15, tmBDCtrls(ilBox).fBoxH - 15, WHITE
            'Else
            '    gPaintArea pbcOffice, tmBDCtrls(ilBox).fBoxX, tmBDCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmBDCtrls(ilBox).fBoxW - 15, tmBDCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            'End If
            pbcOffice.CurrentX = tmBDCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcOffice.CurrentY = tmBDCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = Trim$(sgBDShow(ilBox, ilRow))
            pbcOffice.Print slStr
        Next ilBox
    Next ilRow
    pbcOffice.ForeColor = llColor
End Sub

Private Sub pbcSalesperson_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilLoop As Integer
    Dim ilWk As Integer
    Dim ilMn As Integer
    Dim ilDone As Integer
    Dim ilStartWkNo As Integer
    'ReDim ilStartWk(1 To 12) As Integer
    ReDim ilStartWk(0 To 12) As Integer
    'ReDim ilNoWks(1 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer
    If Button = 2 Then  'Right Mouse
        Exit Sub
    End If
    'Check if hot spot
    If imInHotSpot Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        Exit Sub
    End If
    If UBound(tgBsfRec) > 1 Then
        'For ilLoop = LBound(imHotSpot, 1) To UBound(imHotSpot, 1) Step 1
        For ilLoop = LBound(imHotSpot, 1) + 1 To UBound(imHotSpot, 1) Step 1
            If (X >= imHotSpot(ilLoop, 1)) And (X <= imHotSpot(ilLoop, 3)) And (Y >= imHotSpot(ilLoop, 2)) And (Y <= imHotSpot(ilLoop, 4)) Then
                Screen.MousePointer = vbHourglass
                imInHotSpot = True
                Select Case ilLoop
                    Case 1  'Goto Start
                        imPdYear = imBDStartYear
                        imPdStartWk = 1
                    Case 2  'Reduce by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(1).iYear = imBDStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
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
                            If (tmPdGroups(1).iYear = imBDStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
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
                            If (tmPdGroups(1).iYear = imBDStartYear) And (tmPdGroups(1).iStartWkNo < 3) Then 'At end
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
                        End If
                    Case 3  'Increase by one
                        If rbcType(0).Value Then    'Quarter
                            If (tmPdGroups(4).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(4).iStartWkNo > 39) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            imPdYear = tmPdGroups(2).iYear
                            imPdStartWk = tmPdGroups(2).iStartWkNo
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths tmPdGroups(4).iYear, ilStartWk(), ilNoWks()
                            If (tmPdGroups(4).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(4).iStartWkNo >= ilStartWk(12)) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(2).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(2).iStartWkNo >= ilStartWk(9)) Then 'At end
                                imPdYear = tmPdGroups(2).iYear
                                imPdStartWk = tmPdGroups(2).iStartWkNo
                            ElseIf (tmPdGroups(3).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(3).iStartWkNo >= ilStartWk(9)) Then 'At end
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
                            If (tmPdGroups(4).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(4).iStartWkNo >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imInHotSpot = False
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            If (tmPdGroups(2).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(2).iStartWkNo + 3 >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imPdYear = tmPdGroups(2).iYear
                                imPdStartWk = tmPdGroups(2).iStartWkNo
                            ElseIf (tmPdGroups(3).iYear = imBDStartYear + imBDNoYears - 1) And (tmPdGroups(3).iStartWkNo + 3 >= ilStartWk(12) + ilNoWks(12) - 1) Then 'At end
                                imPdYear = tmPdGroups(3).iYear
                                imPdStartWk = tmPdGroups(3).iStartWkNo
                            Else
                                imPdYear = tmPdGroups(4).iYear
                                imPdStartWk = tmPdGroups(4).iStartWkNo
                            End If
                            'imPdYear = tmPdGroups(2).iYear
                            'imPdStartWk = tmPdGroups(2).iStartWkNo
                        End If
                    Case 4  'GoTo End
                        imPdYear = imBDStartYear + imBDNoYears - 1
                        If rbcType(0).Value Then    'Quarter
                            imPdStartWk = 1
                        ElseIf rbcType(1).Value Then    'Month
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(9)  'At end
                        ElseIf rbcType(2).Value Then    'Week
                            mCompMonths imPdYear, ilStartWk(), ilNoWks()
                            imPdStartWk = ilStartWk(12) + ilNoWks(12) - 4
                        End If
                End Select
                pbcSalesperson.Cls
                mGetShowDates
                pbcSalesperson_Paint
                Screen.MousePointer = vbDefault
                imInHotSpot = False
                Exit Sub
            End If
        Next ilLoop
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub pbcSalesperson_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        Exit Sub
    End If
End Sub
Private Sub pbcSalesperson_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If Button = 2 Then
        imButtonIndex = -1
        Exit Sub
    End If
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If cbcSelect.ListCount <= 1 Then
        Exit Sub
    End If
    ilCompRow = vbcBudget.LargeChange + 1
    If UBound(tgBsfRec) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgBsfRec) + 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBBDCtrls To UBound(tmBDCtrls) Step 1
            If (X >= tmBDCtrls(ilBox).fBoxX) And (X <= (tmBDCtrls(ilBox).fBoxX + tmBDCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmBDCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmBDCtrls(ilBox).fBoxY + tmBDCtrls(ilBox).fBoxH)) Then
                    If (ilBox < DOLLAR1INDEX) And (ilBox > DOLLAR4INDEX) Then    'DAYPARTINDEX Then
                        Beep
                        Exit Sub
                    End If
                    ilRowNo = ilRow + vbcBudget.Value - 1
                    If ilRowNo >= UBound(smBDSave, 2) Then
                        Beep
                        mSetFocus imBDBoxNo
                        Exit Sub
                    End If
                    mBDSetShow imBDBoxNo
                    imBDRowNo = ilRow + vbcBudget.Value - 1
                    imBDBoxNo = ilBox
                    mBDEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    mSetFocus imBDBoxNo
End Sub
Private Sub pbcSalesperson_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single

    mPaintSalespersonTitle
    llColor = pbcSalesperson.ForeColor
    slFontName = pbcSalesperson.FontName
    flFontSize = pbcSalesperson.FontSize
    pbcSalesperson.ForeColor = BLUE
    pbcSalesperson.FontBold = False
    pbcSalesperson.FontSize = 7
    pbcSalesperson.FontName = "Arial"
    pbcSalesperson.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilBox = imLBWKCtrls To UBound(tmWKCtrls) Step 1
        'gPaintArea pbcSalesperson, tmWKCtrls(ilBox).fBoxX, tmWKCtrls(ilBox).fBoxY, tmWKCtrls(ilBox).fBoxW - 15, tmWKCtrls(ilBox).fBoxH - 15, WHITE
        pbcSalesperson.CurrentX = tmWKCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSalesperson.CurrentY = tmWKCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcSalesperson.Print tmWKCtrls(ilBox).sShow
    Next ilBox
    For ilBox = imLBNWCtrls To UBound(tmNWCtrls) Step 1
        'gPaintArea pbcSalesperson, tmNWCtrls(ilBox).fBoxX, tmNWCtrls(ilBox).fBoxY, tmNWCtrls(ilBox).fBoxW - 15, tmNWCtrls(ilBox).fBoxH - 15, LIGHTBLUE
        pbcSalesperson.CurrentX = tmNWCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcSalesperson.CurrentY = tmNWCtrls(ilBox).fBoxY '- 30'+ fgBoxInsetY
        pbcSalesperson.Print tmNWCtrls(ilBox).sShow
    Next ilBox
    pbcSalesperson.FontSize = flFontSize
    pbcSalesperson.FontName = slFontName
    pbcSalesperson.FontSize = flFontSize
    pbcSalesperson.ForeColor = llColor
    pbcSalesperson.FontBold = True
    ilStartRow = vbcBudget.Value '+ 1  'Top location
    ilEndRow = vbcBudget.Value + vbcBudget.LargeChange ' + 1
    If ilEndRow > UBound(smBDSave, 2) Then
        ilEndRow = UBound(smBDSave, 2) 'include blank row as it might have data
    End If
    llColor = pbcSalesperson.ForeColor
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow = UBound(smBDSave, 2) Then
            pbcSalesperson.ForeColor = DARKPURPLE
        Else
            pbcSalesperson.ForeColor = llColor
        End If
        For ilBox = imLBBDCtrls To UBound(tmBDCtrls) Step 1
            'If ilBox <> TOTALINDEX Then
            '    gPaintArea pbcSalesperson, tmBDCtrls(ilBox).fBoxX, tmBDCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmBDCtrls(ilBox).fBoxW - 15, tmBDCtrls(ilBox).fBoxH - 15, WHITE
            'Else
            '    gPaintArea pbcSalesperson, tmBDCtrls(ilBox).fBoxX, tmBDCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmBDCtrls(ilBox).fBoxW - 15, tmBDCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            'End If
            pbcSalesperson.CurrentX = tmBDCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcSalesperson.CurrentY = tmBDCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = Trim$(sgBDShow(ilBox, ilRow))
            pbcSalesperson.Print slStr
            pbcSalesperson.ForeColor = llColor
        Next ilBox
    Next ilRow
    pbcSalesperson.ForeColor = llColor
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If imRetBranch = True Then 'second gotfocus-ignore
        'imRetBranch = False
        Exit Sub
    End If
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    imTabDirection = -1 'Set- Right to left
    Select Case imBDBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If (igMode = 0) And (imFirstTimeSelect) Then
                Exit Sub
            End If
'                ilRet = mStartNew()
'                If Not ilRet Then
'                    Unload Budget
'                    Exit Sub
'                End If
'            End If
'            imFirstTimeSelect = False
            imSettingValue = True
            vbcBudget.Value = 1
            imSettingValue = False
            If igBDView = 0 Then
                If rbcSort(0).Value Or rbcSort(1).Value Then
                    imBDRowNo = 1
                Else
                    imBDRowNo = 2   '1st row is name
                End If
            Else
                imBDRowNo = 1
            End If
            ilBox = DOLLAR1INDEX
            imBDBoxNo = ilBox
            mBDEnableBox ilBox
            Exit Sub
        Case DOLLAR1INDEX 'Name (first control within header)
            mBDSetShow imBDBoxNo
            If igBDView = 0 Then
                Do
                    If imBDRowNo <= 1 Then
                        If (plcSelect.Enabled) And (cbcSelect.Enabled) Then
                            imBDBoxNo = -1
                            cbcSelect.SetFocus
                            Exit Sub
                        End If
                        cmcDone.SetFocus
                    Else
                        ilBox = DOLLAR4INDEX
                        imBDRowNo = imBDRowNo - 1
                        If imBDRowNo < vbcBudget.Value Then
                            imSettingValue = True
                            vbcBudget.Value = vbcBudget.Value - 1
                            imSettingValue = False
                        End If
                    End If
                Loop While imBDSave(1, imBDRowNo) <= 0
                imBDBoxNo = ilBox
                mBDEnableBox ilBox
                Exit Sub
            Else
                If imBDRowNo <= 1 Then
                    If plcSelect.Enabled Then
                        imBDBoxNo = -1
                        If cbcSelect.Enabled Then
                            cbcSelect.SetFocus
                        Else
                            cmcDone.SetFocus
                        End If
                        Exit Sub
                    End If
                    ilBox = 1
                Else
                    ilBox = DOLLAR4INDEX
                    imBDRowNo = imBDRowNo - 1
                    If imBDRowNo < vbcBudget.Value Then
                        imSettingValue = True
                        vbcBudget.Value = vbcBudget.Value - 1
                        imSettingValue = False
                    End If
                    imBDBoxNo = ilBox
                    mBDEnableBox ilBox
                    Exit Sub
                End If
            End If
        Case Else
            ilBox = imBDBoxNo - 1
    End Select
    mBDSetShow imBDBoxNo
    imBDBoxNo = ilBox
    mBDEnableBox ilBox
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
    If (igMode = 0) And (imFirstTimeSelect) Then
        imFirstTimeSelect = False
        ilRet = mStartNew()
        Screen.MousePointer = vbDefault
        If Not ilRet Then
            If cbcSelect.ListCount <= 1 Then
                'imTerminate = True
                'mTerminate
                If rbcOS(0).Value Then
                    rbcOS(0).SetFocus
                Else
                    rbcOS(1).SetFocus
                End If
                Exit Sub
            End If
            cbcSelect.SetFocus
            Exit Sub
        End If
        imBDChg = True
    End If
    mSetCommands
    If pbcSTab.Enabled Then
        pbcSTab.SetFocus
    Else
        pbcClickFocus.SetFocus
    End If
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilEnd As Integer

    If imRetBranch = True Then 'second gotfocus-ignore
        'imRetBranch = False
        Exit Sub
    End If
    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        Exit Sub
    End If
    imTabDirection = 0 'Set- Left to right
    Select Case imBDBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            If igBDView = 0 Then
                imBDRowNo = UBound(sgBDShow, 2) - 1 'UBound(tgBvfRec) - 1
            Else
                imBDRowNo = UBound(tgBsfRec) - 1
            End If
            imSettingValue = True
            If imBDRowNo <= vbcBudget.LargeChange + 1 Then
                vbcBudget.Value = 1
            Else
                vbcBudget.Value = imBDRowNo - vbcBudget.LargeChange - 1
            End If
            imSettingValue = False
            ilBox = DOLLAR4INDEX
        'Case 0
        '    ilBox = DOLLAR1INDEX
        Case DOLLAR1INDEX To DOLLAR4INDEX   'DOLLARINDEX, PCTINVINDEX 'Last control within header
            ilEnd = False
            If imBDBoxNo - DOLLAR1INDEX + 2 >= 5 Then
                ilEnd = True
            Else
                If tmPdGroups(imBDBoxNo - DOLLAR1INDEX + 2).iStartWkNo < 0 Then
                    ilEnd = True
                End If
            End If
            If ilEnd Then
            'If (imBDBoxNo = PCTINVINDEX) Or ((imBDBoxNo = DOLLARINDEX) And (smDPShow(BASEINDEX, imBDRowNo) = "Y")) Then
                mBDSetShow imBDBoxNo
                If mTestSaveFields() = NO Then
                    mBDEnableBox imBDBoxNo
                    Exit Sub
                End If
                If igBDView = 0 Then
                    Do
                        If imBDRowNo + 1 >= UBound(smBDSave, 2) Then    'UBound(tgBvfRec) Then
                            cmcDone.SetFocus
                            Exit Sub
                        End If
                        imBDRowNo = imBDRowNo + 1
                        If imBDRowNo > vbcBudget.Value + vbcBudget.LargeChange Then
                            imSettingValue = True
                            vbcBudget.Value = vbcBudget.Value + 1
                            imSettingValue = False
                        End If
                    Loop While imBDSave(1, imBDRowNo) <= 0
                Else
                    If imBDRowNo + 1 >= UBound(tgBsfRec) Then
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    imBDRowNo = imBDRowNo + 1
                    If imBDRowNo > vbcBudget.Value + vbcBudget.LargeChange Then
                        imSettingValue = True
                        vbcBudget.Value = vbcBudget.Value + 1
                        imSettingValue = False
                    End If
                End If
                ilBox = DOLLAR1INDEX
                imBDBoxNo = ilBox
                mBDEnableBox ilBox
                Exit Sub
            Else
                ilBox = imBDBoxNo + 1
            End If
        Case Else
            ilBox = imBDBoxNo + 1
    End Select
    mBDSetShow imBDBoxNo
    imBDBoxNo = ilBox
    mBDEnableBox ilBox
End Sub
Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcTotals_GotFocus()
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub pbcTotals_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String

    If imTerminate Then
        Exit Sub
    End If
    mPaintTotalTitle
    ilStartRow = vbcTotals.Value '+ 1  'Top location
    ilEndRow = vbcTotals.Value + vbcTotals.LargeChange ' + 1
    If ilEndRow > UBound(smTShow, 2) Then
        ilEndRow = UBound(smTShow, 2) 'include blank row as it might have data
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imlbTCtrls To UBound(tmTCtrls) Step 1
            pbcTotals.CurrentX = tmTCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcTotals.CurrentY = tmTCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            slStr = smTShow(ilBox, ilRow)
            'If ilBox > TNAMEINDEX Then
            '    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
            '    gSetShow pbcTotals, slStr, tmTCtrls(ilBox)
            '    slStr = tmTCtrls(ilBox).sShow
            'End If
            pbcTotals.Print slStr
        Next ilBox
    Next ilRow
    For ilBox = imLBGTCtrls To UBound(tmGTCtrls) Step 1
        pbcTotals.CurrentX = tmGTCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcTotals.CurrentY = tmGTCtrls(ilBox).fBoxY - 30  '+ fgBoxInsetY
        slStr = smGTShow(ilBox)
        'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
        'gSetShow pbcTotals, slStr, tmGTCtrls(ilBox)
        'slStr = tmGTCtrls(ilBox).sShow
        pbcTotals.Print slStr
    Next ilBox
End Sub
Private Sub pbcType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("B") Or (KeyAscii = Asc("b")) Then
        igBudgetType = 0
        pbcType_Paint
    ElseIf KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        igBudgetType = 1
        pbcType_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If igBudgetType = 0 Then
            igBudgetType = 1
            pbcType_Paint
        ElseIf igBudgetType = 1 Then
            igBudgetType = 0
            pbcType_Paint
        End If
    End If
    mPopulate
End Sub
Private Sub pbcType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If igBudgetType = 0 Then
        igBudgetType = 1
    ElseIf igBudgetType = 1 Then
        igBudgetType = 0
    End If
    pbcType_Paint
    mPopulate
End Sub
Private Sub pbcType_Paint()
    pbcType.Cls
    pbcType.CurrentX = fgBoxInsetX
    pbcType.CurrentY = 0 'fgBoxInsetY
    If igBudgetType = 0 Then
        pbcType.Print "Budgets"
    ElseIf igBudgetType = 1 Then
        pbcType.Print "Actuals"
    Else
        pbcType.Print "   "
    End If
End Sub
Private Sub pbcYear_Paint()
    Dim ilBox As Integer
    For ilBox = tmLBSCtrls To UBound(tmSCtrls) Step 1
        pbcYear.CurrentX = tmSCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcYear.CurrentY = tmSCtrls(ilBox).fBoxY + fgBoxInsetY
        pbcYear.Print tmSCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub plcBudget_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcBudget_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
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
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcTotals_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcOS_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOS(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then
            igBDView = 0
            mClearCtrlFields
            plcOS.Move 180, 465, pbcBudgetName.Width + fgPanelAdj, pbcBudgetName.Height + fgPanelAdj - 15
            pbcBudgetName.Move plcOS.Left + fgBevelX, plcOS.Top + fgBevelY
            pbcBudgetName.Visible = True
            pbcYear.Visible = False
            'plcComparison.Visible = True
            plcSort.Visible = True
            pbcType.Visible = True
            pbcOffice.Visible = True
            pbcSalesperson.Visible = False
            plcBudget.Move 120, 975, pbcOffice.Width + vbcBudget.Width + fgPanelAdj, pbcOffice.Height + fgPanelAdj
            pbcOffice.Move plcBudget.Left + fgBevelX, plcBudget.Top + fgBevelY
            'vbcBudget.Move pbcOffice.Width + 60, 60, vbcBudget.Width, pbcOffice.Height - 30
            vbcBudget.Height = pbcOffice.Height
            plcTotals.Visible = True
            pbcTotals.Visible = True
            'vbcBudget.LargeChange = 11
            If rbcSort(0).Value Then
                tmNWCtrls(NWNAMEINDEX).sShow = "Vehicle"
            ElseIf rbcSort(1).Value Then
                tmNWCtrls(NWNAMEINDEX).sShow = "Office"
            ElseIf rbcSort(2).Value Then
                tmNWCtrls(NWNAMEINDEX).sShow = "Office/Vehicle"
            ElseIf rbcSort(3).Value Then
                tmNWCtrls(NWNAMEINDEX).sShow = "Vehicle/Office"
            End If
        Else
            igBDView = 1
            mClearCtrlFields
            plcOS.Move 180, 465, pbcYear.Width + fgPanelAdj, pbcYear.Height + fgPanelAdj - 15
            pbcYear.Move plcOS.Left + fgBevelX, plcOS.Top + fgBevelY
            pbcYear.Visible = True
            pbcBudgetName.Visible = False
            'plcComparison.Visible = False
            plcSort.Visible = False
            pbcType.Visible = False
            pbcSalesperson.Visible = True
            pbcOffice.Visible = False
            plcBudget.Move 120, 975, pbcSalesperson.Width + vbcBudget.Width + fgPanelAdj, pbcSalesperson.Height + fgPanelAdj
            pbcSalesperson.Move plcBudget.Left + fgBevelX, plcBudget.Top + fgBevelY
            'vbcBudget.Move pbcSalesperson.Width + 60, 60, vbcBudget.Width, pbcSalesperson.Height - 30
            vbcBudget.Height = pbcSalesperson.Height
            plcTotals.Visible = False
            pbcTotals.Visible = False
            'vbcBudget.LargeChange = 17
            tmNWCtrls(NWNAMEINDEX).sShow = "Salesperson"
        End If
        mPopulate
    End If
End Sub
Private Sub rbcOS_GotFocus(Index As Integer)
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
    If imFirstTime Then
        DoEvents
        imFirstTime = False
    End If
End Sub
Private Sub rbcShow_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcShow(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilFound As Integer
    'ReDim ilStartWk(1 To 12) As Integer
    ReDim ilStartWk(0 To 12) As Integer
    'ReDim ilNoWks(1 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer

    If Value Then
        Screen.MousePointer = vbHourglass
        pbcOffice.Cls
        pbcSalesperson.Cls
        pbcTotals.Cls
        If imBDStartYear <> 0 Then
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
                    If (imPdStartWk > ilStartWk(12) + ilNoWks(12) - 6) Then
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
        If igBDView = 0 Then
            pbcOffice_Paint
            pbcTotals_Paint
        Else
            pbcSalesperson_Paint
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcShow_GotFocus(Index As Integer)
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub rbcSort_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSort(Index).Value
    'End of coded added
    Dim ilBvf As Integer
    Dim ilUpper As Integer
    If Value Then
        Screen.MousePointer = vbHourglass
        If Index = 0 Then
            tmNWCtrls(NWNAMEINDEX).sShow = "Vehicle"
        ElseIf Index = 1 Then
            tmNWCtrls(NWNAMEINDEX).sShow = "Office"
        ElseIf Index = 2 Then
            tmNWCtrls(NWNAMEINDEX).sShow = "Office/Vehicle"
        ElseIf Index = 3 Then
            tmNWCtrls(NWNAMEINDEX).sShow = "Vehicle/Office"
        End If
        'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
            If Index = 0 Then
                tgBvfRec(ilBvf).sKey = tgBvfRec(ilBvf).sVehSort & tgBvfRec(ilBvf).sVehicle & tgBvfRec(ilBvf).sMktRank & tgBvfRec(ilBvf).SOffice
            ElseIf Index = 1 Then
                tgBvfRec(ilBvf).sKey = tgBvfRec(ilBvf).sMktRank & tgBvfRec(ilBvf).SOffice & tgBvfRec(ilBvf).sVehSort & tgBvfRec(ilBvf).sVehicle
            ElseIf Index = 2 Then  'Vehicle within office
                tgBvfRec(ilBvf).sKey = tgBvfRec(ilBvf).sMktRank & tgBvfRec(ilBvf).SOffice & tgBvfRec(ilBvf).sVehSort & tgBvfRec(ilBvf).sVehicle
            ElseIf Index = 3 Then
                tgBvfRec(ilBvf).sKey = tgBvfRec(ilBvf).sVehSort & tgBvfRec(ilBvf).sVehicle & tgBvfRec(ilBvf).sMktRank & tgBvfRec(ilBvf).SOffice
            End If
        Next ilBvf
        ilUpper = UBound(tgBvfRec)
        If ilUpper > 1 Then
            ArraySortTyp fnAV(tgBvfRec(), 1), UBound(tgBvfRec) - 1, 0, LenB(tgBvfRec(1)), 0, LenB(tgBvfRec(1).sKey), 0
        End If
        pbcOffice.Cls
        pbcTotals.Cls
        'mInitBudgetCtrls
        mMoveRecToCtrl
        pbcOffice_Paint
        pbcTotals_Paint
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcSort_GotFocus(Index As Integer)
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub rbcType_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcType(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    Dim ilFound As Integer
    'ReDim ilStartWk(1 To 12) As Integer
    ReDim ilStartWk(0 To 12) As Integer
    'ReDim ilNoWks(1 To 12) As Integer
    ReDim ilNoWks(0 To 12) As Integer

    If Value Then
        Screen.MousePointer = vbHourglass
        pbcOffice.Cls
        pbcSalesperson.Cls
        pbcTotals.Cls
        If imBDStartYear <> 0 Then
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
        If igBDView = 0 Then
            pbcOffice_Paint
            pbcTotals_Paint
        Else
            pbcSalesperson_Paint
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub rbcType_GotFocus(Index As Integer)
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub tmcDrag_Timer()
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mCheckGG
    If igGGFlag = 0 Then
        imTerminate = True
        cmcCancel_Click
        Exit Sub
    End If
End Sub

Private Sub vbcBudget_Change()
    If imSettingValue Then
        If igBDView = 0 Then
            pbcOffice.Cls
            pbcOffice_Paint
        Else
            pbcSalesperson.Cls
            pbcSalesperson_Paint
        End If
        imSettingValue = False
    Else
        mBDSetShow imBDBoxNo
        'imBDBoxNo = -1
        'pbcArrow.Visible = False
        'lacRCFrame.Visible = False
        'lacDPFrame.Visible = False
        If igBDView = 0 Then
            pbcOffice.Cls
            pbcOffice_Paint
        Else
            pbcSalesperson.Cls
            pbcSalesperson_Paint
        End If
        If (igWinStatus(BUDGETSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
            mBDEnableBox imBDBoxNo
        End If
    End If
End Sub
Private Sub vbcBudget_GotFocus()
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub vbcTotals_Change()
    If imSettingValue Then
        pbcTotals.Cls
        pbcTotals_Paint
        imSettingValue = False
    Else
        pbcTotals.Cls
        pbcTotals_Paint
        'If (igWinStatus(BUDGETSJOB) <> 1) Or (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        '    mBDEnableBox imBDBoxNo
        'End If
    End If
End Sub
Private Sub vbcTotals_GotFocus()
    'mOVSetShow imOVBoxNo    'Remove focus
    imOVBoxNo = -1
    'mSSetShow imSBoxNo    'Remove focus
    imSBoxNo = -1
    mBDSetShow imBDBoxNo
    imBDBoxNo = -1
    imBDRowNo = -1
    pbcArrow.Visible = False
    lacOFrame.Visible = False
    lacSFrame.Visible = False
End Sub
Private Sub plcComparison_Paint()
    plcComparison.CurrentX = 0
    plcComparison.CurrentY = 0
    plcComparison.Print "Comparison"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Budget"
End Sub
Private Sub plcShow_Paint()
    plcShow.CurrentX = 0
    plcShow.CurrentY = 0
    plcShow.Print "Show by"
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
Private Sub mPaintOfficeTitle()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llWidth                       ilRow                                                   *
'******************************************************************************************

    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer

    llColor = pbcOffice.ForeColor
    slFontName = pbcOffice.FontName
    flFontSize = pbcOffice.FontSize
    ilFillStyle = pbcOffice.FillStyle
    llFillColor = pbcOffice.FillColor
    pbcOffice.ForeColor = BLUE
    pbcOffice.FontBold = False
    pbcOffice.FontSize = 7
    pbcOffice.FontName = "Arial"
    pbcOffice.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    llTop = 15
    For ilLoop = 1 To 2 Step 1
        If ilLoop = 2 Then
            pbcOffice.Line (tmBDCtrls(OSNAMEINDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(OSNAMEINDEX).fBoxW + 15, tmBDCtrls(OSNAMEINDEX).fBoxH + 15), BLUE, B
            pbcOffice.Line (tmBDCtrls(OSNAMEINDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(OSNAMEINDEX).fBoxW - 15, tmBDCtrls(OSNAMEINDEX).fBoxH - 15), LIGHTYELLOW, BF
        End If
        pbcOffice.Line (tmBDCtrls(DOLLAR1INDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(DOLLAR1INDEX).fBoxW + 15, tmBDCtrls(DOLLAR1INDEX).fBoxH + 15), BLUE, B
        pbcOffice.Line (tmBDCtrls(DOLLAR1INDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(DOLLAR1INDEX).fBoxW - 15, tmBDCtrls(DOLLAR1INDEX).fBoxH - 15), LIGHTYELLOW, BF
        pbcOffice.Line (tmBDCtrls(DOLLAR2INDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(DOLLAR2INDEX).fBoxW + 15, tmBDCtrls(DOLLAR2INDEX).fBoxH + 15), BLUE, B
        pbcOffice.Line (tmBDCtrls(DOLLAR2INDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(DOLLAR2INDEX).fBoxW - 15, tmBDCtrls(DOLLAR2INDEX).fBoxH - 15), LIGHTYELLOW, BF
        pbcOffice.Line (tmBDCtrls(DOLLAR3INDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(DOLLAR3INDEX).fBoxW + 15, tmBDCtrls(DOLLAR3INDEX).fBoxH + 15), BLUE, B
        pbcOffice.Line (tmBDCtrls(DOLLAR3INDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(DOLLAR3INDEX).fBoxW - 15, tmBDCtrls(DOLLAR3INDEX).fBoxH - 15), LIGHTYELLOW, BF
        pbcOffice.Line (tmBDCtrls(DOLLAR4INDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(DOLLAR4INDEX).fBoxW + 15, tmBDCtrls(DOLLAR4INDEX).fBoxH + 15), BLUE, B
        pbcOffice.Line (tmBDCtrls(DOLLAR4INDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(DOLLAR4INDEX).fBoxW - 15, tmBDCtrls(DOLLAR4INDEX).fBoxH - 15), LIGHTYELLOW, BF
        If ilLoop = 2 Then
            pbcOffice.Line (tmBDCtrls(TOTALINDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(TOTALINDEX).fBoxW + 15, tmBDCtrls(TOTALINDEX).fBoxH + 15), BLUE, B
            pbcOffice.Line (tmBDCtrls(TOTALINDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(TOTALINDEX).fBoxW - 15, tmBDCtrls(TOTALINDEX).fBoxH - 15), LIGHTYELLOW, BF
            pbcOffice.CurrentX = tmBDCtrls(TOTALINDEX).fBoxX + 15  'fgBoxInsetX
            pbcOffice.CurrentY = llTop
            pbcOffice.Print "Total"
        End If
        llTop = llTop + tmBDCtrls(1).fBoxH + 15
    Next ilLoop


    ilLineCount = 0
    llTop = tmBDCtrls(1).fBoxY
    Do
        For ilLoop = imLBBDCtrls To UBound(tmBDCtrls) Step 1
            pbcOffice.Line (tmBDCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmBDCtrls(ilLoop).fBoxW + 15, tmBDCtrls(ilLoop).fBoxH + 15), BLUE, B
            If (ilLoop = imLBBDCtrls) Or (ilLoop = UBound(tmBDCtrls)) Then
                pbcOffice.Line (tmBDCtrls(ilLoop).fBoxX, llTop)-Step(tmBDCtrls(ilLoop).fBoxW - 15, tmBDCtrls(ilLoop).fBoxH - 15), LIGHTYELLOW, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmBDCtrls(1).fBoxH + 15
    Loop While llTop + tmBDCtrls(1).fBoxH < pbcOffice.Height
    vbcBudget.LargeChange = ilLineCount - 1


    pbcOffice.FontSize = flFontSize
    pbcOffice.FontName = slFontName
    pbcOffice.FontSize = flFontSize
    pbcOffice.ForeColor = llColor
    pbcOffice.FontBold = True
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
Private Sub mPaintTotalTitle()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLineCount                   llWidth                                                 *
'******************************************************************************************

    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilRow As Integer

    llColor = pbcOffice.ForeColor
    slFontName = pbcOffice.FontName
    flFontSize = pbcOffice.FontSize
    ilFillStyle = pbcOffice.FillStyle
    llFillColor = pbcOffice.FillColor
    pbcTotals.ForeColor = BLUE
    pbcTotals.FontBold = False
    pbcTotals.FontSize = 7
    pbcTotals.FontName = "Arial"
    pbcTotals.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    llTop = tmTCtrls(1).fBoxY
    For ilRow = 1 To 5 Step 1
        For ilLoop = imlbTCtrls To UBound(tmTCtrls) Step 1
            If (ilRow = 5) And (ilLoop = imlbTCtrls) Then
                pbcTotals.CurrentX = tmTCtrls(ilLoop).fBoxX + 15  'fgBoxInsetX
                pbcTotals.CurrentY = llTop + 30
                pbcTotals.Print "Total"
            ElseIf ilRow = 5 Then
                pbcTotals.Line (tmTCtrls(ilLoop).fBoxX - 15, llTop + 15)-Step(tmTCtrls(ilLoop).fBoxW + 15, tmTCtrls(ilLoop).fBoxH + 15), BLUE, B
                pbcTotals.Line (tmTCtrls(ilLoop).fBoxX, llTop + 45)-Step(tmTCtrls(ilLoop).fBoxW - 15, tmTCtrls(ilLoop).fBoxH - 45), LIGHTYELLOW, BF
            Else
                pbcTotals.Line (tmTCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmTCtrls(ilLoop).fBoxW + 15, tmTCtrls(ilLoop).fBoxH + 15), BLUE, B
                pbcTotals.Line (tmTCtrls(ilLoop).fBoxX, llTop)-Step(tmTCtrls(ilLoop).fBoxW - 15, tmTCtrls(ilLoop).fBoxH - 15), LIGHTYELLOW, BF
            End If
        Next ilLoop
        llTop = llTop + tmTCtrls(1).fBoxH + 15
    Next ilRow

    pbcTotals.FontSize = flFontSize
    pbcTotals.FontName = slFontName
    pbcTotals.FontSize = flFontSize
    pbcTotals.ForeColor = llColor
    pbcTotals.FontBold = True
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
Private Sub mPaintSalespersonTitle()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llWidth                       ilRow                                                   *
'******************************************************************************************

    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer

    llColor = pbcSalesperson.ForeColor
    slFontName = pbcSalesperson.FontName
    flFontSize = pbcSalesperson.FontSize
    ilFillStyle = pbcSalesperson.FillStyle
    llFillColor = pbcSalesperson.FillColor
    pbcSalesperson.ForeColor = BLUE
    pbcSalesperson.FontBold = False
    pbcSalesperson.FontSize = 7
    pbcSalesperson.FontName = "Arial"
    pbcSalesperson.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    llTop = 15
    For ilLoop = 1 To 2 Step 1
        If ilLoop = 2 Then
            pbcSalesperson.Line (tmBDCtrls(OSNAMEINDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(OSNAMEINDEX).fBoxW + 15, tmBDCtrls(OSNAMEINDEX).fBoxH + 15), BLUE, B
            pbcSalesperson.Line (tmBDCtrls(OSNAMEINDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(OSNAMEINDEX).fBoxW - 15, tmBDCtrls(OSNAMEINDEX).fBoxH - 15), LIGHTYELLOW, BF
        End If
        pbcSalesperson.Line (tmBDCtrls(DOLLAR1INDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(DOLLAR1INDEX).fBoxW + 15, tmBDCtrls(DOLLAR1INDEX).fBoxH + 15), BLUE, B
        pbcSalesperson.Line (tmBDCtrls(DOLLAR1INDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(DOLLAR1INDEX).fBoxW - 15, tmBDCtrls(DOLLAR1INDEX).fBoxH - 15), LIGHTYELLOW, BF
        pbcSalesperson.Line (tmBDCtrls(DOLLAR2INDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(DOLLAR2INDEX).fBoxW + 15, tmBDCtrls(DOLLAR2INDEX).fBoxH + 15), BLUE, B
        pbcSalesperson.Line (tmBDCtrls(DOLLAR2INDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(DOLLAR2INDEX).fBoxW - 15, tmBDCtrls(DOLLAR2INDEX).fBoxH - 15), LIGHTYELLOW, BF
        pbcSalesperson.Line (tmBDCtrls(DOLLAR3INDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(DOLLAR3INDEX).fBoxW + 15, tmBDCtrls(DOLLAR3INDEX).fBoxH + 15), BLUE, B
        pbcSalesperson.Line (tmBDCtrls(DOLLAR3INDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(DOLLAR3INDEX).fBoxW - 15, tmBDCtrls(DOLLAR3INDEX).fBoxH - 15), LIGHTYELLOW, BF
        pbcSalesperson.Line (tmBDCtrls(DOLLAR4INDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(DOLLAR4INDEX).fBoxW + 15, tmBDCtrls(DOLLAR4INDEX).fBoxH + 15), BLUE, B
        pbcSalesperson.Line (tmBDCtrls(DOLLAR4INDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(DOLLAR4INDEX).fBoxW - 15, tmBDCtrls(DOLLAR4INDEX).fBoxH - 15), LIGHTYELLOW, BF
        If ilLoop = 2 Then
            pbcSalesperson.Line (tmBDCtrls(TOTALINDEX).fBoxX - 15, llTop)-Step(tmBDCtrls(TOTALINDEX).fBoxW + 15, tmBDCtrls(TOTALINDEX).fBoxH + 15), BLUE, B
            pbcSalesperson.Line (tmBDCtrls(TOTALINDEX).fBoxX, llTop + 15)-Step(tmBDCtrls(TOTALINDEX).fBoxW - 15, tmBDCtrls(TOTALINDEX).fBoxH - 15), LIGHTYELLOW, BF
            pbcSalesperson.CurrentX = tmBDCtrls(TOTALINDEX).fBoxX + 15  'fgBoxInsetX
            pbcSalesperson.CurrentY = llTop
            pbcSalesperson.Print "Total"
        End If
        llTop = llTop + tmBDCtrls(1).fBoxH + 15
    Next ilLoop


    ilLineCount = 0
    llTop = tmBDCtrls(1).fBoxY
    Do
        For ilLoop = imLBBDCtrls To UBound(tmBDCtrls) Step 1
            pbcSalesperson.Line (tmBDCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmBDCtrls(ilLoop).fBoxW + 15, tmBDCtrls(ilLoop).fBoxH + 15), BLUE, B
            If (ilLoop = imLBBDCtrls) Or (ilLoop = UBound(tmBDCtrls)) Then
                pbcSalesperson.Line (tmBDCtrls(ilLoop).fBoxX, llTop)-Step(tmBDCtrls(ilLoop).fBoxW - 15, tmBDCtrls(ilLoop).fBoxH - 15), LIGHTYELLOW, BF
            End If
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmBDCtrls(1).fBoxH + 15
    Loop While llTop + tmBDCtrls(1).fBoxH < pbcSalesperson.Height
    vbcBudget.LargeChange = ilLineCount - 1


    pbcSalesperson.FontSize = flFontSize
    pbcSalesperson.FontName = slFontName
    pbcSalesperson.FontSize = flFontSize
    pbcSalesperson.ForeColor = llColor
    pbcSalesperson.FontBold = True
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
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slStartIn As String
    Dim slCSIName As String

    
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
    gUserActivityLog "L", "Budget.Frm"
    If tgSpf.sGUsePropSys = "Y" Then
        If igWinStatus(BUDGETSJOB) = 0 Then
            imTerminate = True
        End If
    Else
        imTerminate = True
    End If

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
'4/2/11: Add routine
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

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    
    On Error Resume Next
    
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    igGGFlag = 1
    hmSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    imSafRecLen = Len(tmSaf)
    ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    ilField1 = Asc(tmSaf.sName)
    slStr = Mid$(tmSaf.sName, 2, 5)
    llDate = Val(slStr)
    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
    llNow = gDateValue(Format$(Now, "m/d/yy"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        igGGFlag = 0
    End If
    'gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub



Private Sub mRefreshControls(blVisibleOnly As Boolean)
    plcSelect.Visible = False
    plcSort.Visible = False
    plcShow.Visible = False
    plcType.Visible = False
    If Not blVisibleOnly Then
        pbcOffice.Cls
        pbcTotals.Cls
        mGetShowPrices
        pbcOffice_Paint
        pbcTotals_Paint
        imBDChg = True
        mSetCommands
    End If
    plcSelect.Visible = True
    plcSort.Visible = True
    plcShow.Visible = True
    plcType.Visible = True
End Sub
