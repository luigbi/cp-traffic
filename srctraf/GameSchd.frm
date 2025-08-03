VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form GameSchd 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   9345
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   9345
   Begin VB.ListBox lbcSubtotal 
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
      Index           =   1
      ItemData        =   "GameSchd.frx":0000
      Left            =   5175
      List            =   "GameSchd.frx":0002
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4455
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcSubtotal 
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
      Index           =   0
      ItemData        =   "GameSchd.frx":0004
      Left            =   4065
      List            =   "GameSchd.frx":0006
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox pbcDefault 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2775
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmcSpec 
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
      Left            =   1185
      Picture         =   "GameSchd.frx":0008
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cbcSeason 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   6975
      TabIndex        =   2
      Top             =   60
      Width           =   2115
   End
   Begin VB.PictureBox pbcLiveLogMerge 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6675
      ScaleHeight     =   210
      ScaleWidth      =   1620
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   915
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox pbcStartNew 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   9270
      ScaleHeight     =   180
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   15
      Width           =   60
   End
   Begin VB.CommandButton cmcSyncGames 
      Appearance      =   0  'Flat
      Caption         =   "S&ync"
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
      Left            =   7530
      TabIndex        =   41
      Top             =   5355
      Width           =   1050
   End
   Begin VB.CheckBox ckcShowVersion 
      Caption         =   "Show Library Version Numbers"
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
      Height          =   195
      Left            =   195
      TabIndex        =   40
      Top             =   5040
      Width           =   2745
   End
   Begin VB.CommandButton cmcFormats 
      Appearance      =   0  'Flat
      Caption         =   "&Formats"
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
      Left            =   4875
      TabIndex        =   37
      Top             =   5355
      Width           =   1050
   End
   Begin VB.CommandButton cmcMultimedia 
      Appearance      =   0  'Flat
      Caption         =   "&MultiMedia"
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
      Left            =   6225
      TabIndex        =   38
      Top             =   5355
      Width           =   1050
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8895
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5370
      Width           =   45
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8775
      Top             =   5385
   End
   Begin VB.ListBox lbcStatus 
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
      ItemData        =   "GameSchd.frx":0102
      Left            =   7050
      List            =   "GameSchd.frx":0104
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox pbcFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   675
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.ListBox lbcLanguage 
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
      ItemData        =   "GameSchd.frx":0106
      Left            =   1170
      List            =   "GameSchd.frx":0108
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4140
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcLibrary 
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
      ItemData        =   "GameSchd.frx":010A
      Left            =   4185
      List            =   "GameSchd.frx":010C
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcTeam 
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
      ItemData        =   "GameSchd.frx":010E
      Left            =   2115
      List            =   "GameSchd.frx":0110
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.ListBox lbcAirVehicle 
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
      ItemData        =   "GameSchd.frx":0112
      Left            =   3630
      List            =   "GameSchd.frx":0114
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3675
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.TextBox edcSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   225
      MaxLength       =   10
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   45
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   9
      Top             =   825
      Width           =   45
   End
   Begin VB.PictureBox pbcSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   45
      ScaleHeight     =   180
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   375
      Width           =   60
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   75
      Width           =   3720
   End
   Begin VB.CommandButton cmcSave 
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
      Left            =   3525
      TabIndex        =   36
      Top             =   5355
      Width           =   1050
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
      Left            =   1935
      Picture         =   "GameSchd.frx":0116
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   990
      MaxLength       =   10
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1410
      Left            =   1875
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
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
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "GameSchd.frx":0210
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
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
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "GameSchd.frx":0ECE
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1575
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1185
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
         Left            =   45
         Picture         =   "GameSchd.frx":11D8
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   29
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
            TabIndex        =   30
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
         Left            =   45
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   45
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
         Left            =   1635
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   27
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "GameSchd.frx":3FF2
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   90
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   33
      Top             =   5640
      Width           =   75
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   32
      Top             =   5070
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   10
      Top             =   810
      Width           =   105
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
      Left            =   2190
      TabIndex        =   35
      Top             =   5355
      Width           =   1050
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
      Left            =   825
      TabIndex        =   34
      Top             =   5355
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDates 
      Height          =   4035
      Left            =   195
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7117
      _Version        =   393216
      Rows            =   41
      Cols            =   34
      FixedRows       =   5
      FixedCols       =   4
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   34
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpec 
      Height          =   420
      Left            =   210
      TabIndex        =   5
      Top             =   465
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   741
      _Version        =   393216
      Rows            =   5
      Cols            =   12
      FixedRows       =   2
      FixedCols       =   2
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
   End
   Begin VB.Label plcScreen 
      Caption         =   "Event Schedule"
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
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   1425
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   5250
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "GameSchd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GameSchd.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmLcfSrchKey0                 tmSsf                         tmSsfSrchKey1             *
'*  imSsfRecLen                   tmProg                        tmAvail                   *
'*  tmClfSrchKey0                 tmMnfSrchKey                  tmIhfSrchKey1             *
'*  tmIhfSrchKey2                                                                         *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: GameSchd.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imSelectedIndex As Integer
Dim imSeasonSelectedIndex As Integer
Dim imInNew As Integer
Dim imFirstTimeSelect As Integer
Dim imComboBoxIndex As Integer
Dim imSeasonComboBoxIndex As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Dim imVefCode As Integer
Dim imVpfIndex As Integer
Dim smNowDate As String
Dim lmNowDate As Long
Dim lmLLD As Long
Dim lmFirstAllowedChgDate As Long
Dim smFeedSource As String
Dim smLiveLogMerge As String
Dim imLastColSorted As Integer
Dim imLastSort As Integer
Dim lmLockRecCode As Long
Dim lmSeasonGhfCode As Long
Dim smDefault As String

Dim imGhfChg As Integer
Dim imGsfChg As Integer
Dim imNewGame As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmSpecEnableRow As Long
Dim lmSpecEnableCol As Long
Dim imSpecCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmGhfSrchKey0 As LONGKEY0    'GHF key record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length
Dim tmSeasonInfo() As SEASONINFO

Dim hmGsf As Integer
Dim tmGsf() As GSF        'GSF record image
Dim tmGsfSrchKey0 As LONGKEY0    'GSF key record image
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

Dim hmLcf As Integer
Dim tmLcf As LCF        'LCF record image
Dim tmDeletedLcf As LCF 'Image used when updaing LCF
Dim tmLcfSrchKey1 As LCFKEY1    'LCF key record image
Dim tmLcfSrchKey2 As LCFKEY2    'LCF key record image
Dim imLcfRecLen As Integer        'LCF record length

'Library version
Dim hmLvf As Integer            'Log library file handle
Dim tmLvf As LVF                'LVF record image
Dim tmLvfSrchKey As LONGKEY0     'LVF key record image
Dim imLvfRecLen As Integer         'LVF record length

Dim hmSsf As Integer
Dim tmSsf As SSF
'Dim tmDeletedSsf As SSF 'Image used when updating SSF
Dim hmRlf As Integer

Dim hmSxf As Integer

Dim hmCHF As Integer
Dim tmChf As CHF        'CHF record image
Dim tmChfSrchKey0 As LONGKEY0    'CHF key record image
Dim imCHFRecLen As Integer        'CHF record length
Dim hmClf As Integer
Dim tmClf As CLF        'CLF record image
Dim tmClfSrchKey0 As CLFKEY0    'CLF key record image
Dim tmClfSrchKey2 As LONGKEY0    'CLF key record image
Dim tmClfSrchKey3 As CLFKEY3    'CLF key record image
Dim imClfRecLen As Integer        'CLF record length
Dim lmClfCode() As Long
Dim hmCff As Integer
Dim tmCff As CFF        'CFF record image
Dim tmCffSrchKey0 As CFFKEY0    'CFF key record image
Dim imCffRecLen As Integer        'CFF record length

Dim hmCgf As Integer
Dim tmCgf As CGF        'CGF record image
Dim tmCgfSrchKey1 As CGFKEY1    'CGF key record image
Dim imCgfRecLen As Integer        'CGF record length

Dim hmSdf As Integer
Dim tmSdf As SDF        'SDF record image
Dim tmSdfSrchKey3 As LONGKEY0    'SDF key record image
Dim tmSdfSrchKey6 As SDFKEY6    'SDF key record image
Dim imSdfRecLen As Integer        'SDF record length
Dim lmSdfCode() As Long
Dim lmBBSdfCode() As Long

Dim hmSmf As Integer
Dim tmSmf As SMF        'SMF record image
Dim tmSmfSrchKey1 As LONGKEY0    'SMF key record image
Dim tmSmfSrchKey2 As LONGKEY0    'SMF key record image
Dim tmSmfSrchKey3 As SMFKEY3    'SMF key record image
Dim imSmfRecLen As Integer        'SMF record length
Dim lmSmfCode() As Long

Dim hmMsf As Integer
Dim tmMsf As MSF        'MSF record image
Dim tmMsfSrchKey1 As MSFKEY1    'MSF key record image
Dim imMsfRecLen As Integer        'MSF record length


Dim hmMgf As Integer
Dim tmMgf As MGF        'MGF record image
Dim tmMgfSrchKey1 As MGFKEY1    'MGF key record image
Dim imMgfRecLen As Integer        'MGF record length


Dim hmSbf As Integer
Dim tmSbf As SBF        'SBF record image
Dim tmSbfSrchKey0 As SBFKEY0    'SBF key record image
Dim tmSbfSrchKey1 As LONGKEY0    'SBF key record image
Dim imSbfRecLen As Integer        'SBF record length

Dim tmMnf As MNF        'Mnf record image
Dim hmMnf As Integer    'Multi-Name file handle
Dim imMnfRecLen As Integer        'MNF record length
Dim tmNTRMNF() As MNF
Dim smMnfStamp As String
Dim imNTRMnfCode As Integer
Dim imNTRSlspComm As Integer

Dim hmIhf As Integer
Dim tmIhf As IHF        'IHF record image
Dim tmIhfSrchKey0 As INTKEY0    'IHF key record image
Dim imIhfRecLen As Integer        'IHF record length

Dim hmItf As Integer
Dim tmItfSrchKey0 As INTKEY0    'ITF key record image
Dim tmItf As ITF
Dim imItfRecLen As Integer        'ITF record length

Dim hmIif As Integer
Dim tmIif As IIF        'IIF record image
Dim tmIifSrchKey0 As INTKEY0    'IIF key record image
Dim imIifRecLen As Integer        'IIF record length

Dim hmVff As Integer
Dim tmVff As VFF        'SMF record image
Dim tmVffSrchKey0 As INTKEY0    'SMF key record image
Dim tmVffSrchKey1 As INTKEY0    'SMF key record image
Dim imVffRecLen As Integer        'SMF record length

Dim hmEcf As Integer
Dim tmEcf As ECF        'SMF record image
Dim tmEcfSrchKey0 As LONGKEY0    'SMF key record image
Dim imEcfRecLen As Integer        'SMF record length

'Log Spot record
Dim hmLst As Integer        'Log Spots file
Dim tmLst As LST
Dim imLstRecLen As Integer
Dim tmLstSrchKey3 As LSTKEY3    'LST key record image

Dim tmGameVehicle() As SORTCODE
Dim smGameVehicleTag As String

Dim tmAirVehicle() As SORTCODE
Dim smAirVehicleTag As String

Dim tmTeamCode() As SORTCODE
Dim smTeamCodeTag As String

Dim tmLanguageCode() As SORTCODE
Dim smLanguageCodeTag As String

Dim tmSubtotal1Code() As SORTCODE
Dim smSubtotal1CodeTag As String

Dim tmSubtotal2Code() As SORTCODE
Dim smSubtotal2CodeTag As String

Dim tmSdfMdExt() As SDFMDEXT
Dim smSdfMdExtTag As String
Dim imLBSdfMdExt As Integer

Dim tmLibName() As SORTCODE
Dim smLibNameTag As String

Dim bmChgFlag() As Boolean

'6/9/14
Dim smEventTitle1 As String
Dim smEventTitle2 As String

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

'Mouse down
Const SPECROW3INDEX = 3
Const SEASONNAMEINDEX = 2   '1
Const SEASONSTARTINDEX = 4   '2
Const SEASONENDINDEX = 6 '3
Const DEFAULTINDEX = 8
Const NOGAMEINDEX = 10   '1

'If changed, also change GameLib
Const GAMENOINDEX = 2   '1
Const FEEDSOURCEINDEX = 4   '2
Const LANGUAGEINDEX = 6 '3
Const VISITTEAMINDEX = 8    '4
Const HOMETEAMINDEX = 10    '5
Const SUBTOTAL1INDEX = 12
Const SUBTOTAL2INDEX = 14
Const LIBRARYINDEX = 16 '12 '6
Const AIRDATEINDEX = 18 '14 '7
Const AIRTIMEINDEX = 20 '16 '8
Const AIRVEHICLEINDEX = 22  '18  '9
Const XDSPROGCODEINDEX = 24 '20
Const BUSINDEX = 26 '22
Const GAMESTATUSINDEX = 28  '24  '10
Const TMGSFINDEX = 30   '26
Const CHGFLAGINDEX = 31 '27
Const SORTINDEX = 32    '28
Const VERLIBINDEX = 33  '29




Private Sub cbcSeason_Change()
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        If cbcSeason.Text <> "" Then
            gManLookAhead cbcSeason, imBSMode, imSeasonComboBoxIndex
            mcbcSeasonChange
        End If
    End If
    Exit Sub
End Sub

Private Sub cbcSeason_Click()
    cbcSeason_Change
End Sub

Private Sub cbcSeason_GotFocus()
    Dim ilVff As Integer
    Dim ilLoop As Integer
    
    mSpecSetShow
    mSetShow
    If cbcSeason.Text = "" Then
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tgVff(ilVff).iVefCode = imVefCode Then
                lmSeasonGhfCode = tgVff(ilVff).lSeasonGhfCode
                Exit For
            End If
        Next ilVff
        For ilLoop = 1 To cbcSeason.ListCount - 1 Step 1
            If cbcSeason.ItemData(ilLoop) = lmSeasonGhfCode Then
                cbcSeason.ListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
        If cbcSeason.ListIndex < 0 Then
            If cbcSeason.ListCount >= 1 Then
                cbcSeason.ListIndex = 0
            End If
        End If
        imSeasonComboBoxIndex = cbcSeason.ListIndex
        imSeasonSelectedIndex = imComboBoxIndex
    End If
    imSeasonComboBoxIndex = imSeasonSelectedIndex
    gCtrlGotFocus cbcSeason
End Sub

Private Sub cbcSeason_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcSeason_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSeason.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcSelect_Change()
    If imStartMode Then
        imStartMode = False
        mcbcSelectChange
        Exit Sub
    End If
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        If cbcSelect.Text <> "" Then
            gManLookAhead cbcSelect, imBSMode, imComboBoxIndex
            mcbcSelectChange
        End If
    End If
    Exit Sub
End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcSelect_GotFocus()
    mSpecSetShow
    mSetShow
    If cbcSelect.Text = "" Then
        gFindMatch sgUserDefVehicleName, 0, cbcSelect
        If gLastFound(cbcSelect) >= 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        Else
            If cbcSelect.ListCount >= 1 Then
                cbcSelect.ListIndex = 0
            End If
        End If
        imComboBoxIndex = cbcSelect.ListIndex
        imSelectedIndex = imComboBoxIndex
    End If
    imComboBoxIndex = imSelectedIndex
    gCtrlGotFocus cbcSelect
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

Private Sub ckcShowVersion_Click()
    Dim ilRow As Integer
    Dim ilPos As Integer
    Dim slStr As String

    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If grdDates.TextMatrix(ilRow, GAMENOINDEX) <> "" Then
            slStr = grdDates.TextMatrix(ilRow, VERLIBINDEX)
            If ckcShowVersion.Value = vbUnchecked Then
                ilPos = InStr(1, slStr, "/", vbTextCompare)
                If ilPos > 0 Then
                    slStr = Mid$(slStr, ilPos + 1)
                End If
            End If
            grdDates.TextMatrix(ilRow, LIBRARYINDEX) = slStr
        End If
    Next ilRow
End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    If imSpecCtrlVisible Then
        edcSpec.SelStart = 0
        edcSpec.SelLength = Len(edcSpec.Text)
        edcSpec.SetFocus
    Else
        edcDropdown.SelStart = 0
        edcDropdown.SelLength = Len(edcDropdown.Text)
        edcDropdown.SetFocus
    End If
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    If imSpecCtrlVisible Then
        edcSpec.SelStart = 0
        edcSpec.SelLength = Len(edcSpec.Text)
        edcSpec.SetFocus
    Else
        edcDropdown.SelStart = 0
        edcDropdown.SelLength = Len(edcDropdown.Text)
        edcDropdown.SetFocus
    End If
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSpecSetShow
    mSetShow
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                                                                                 *
'******************************************************************************************

    Dim slMess As String
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If imGhfChg Or imGsfChg Then
        If Not imNewGame Then
            slMess = "Save Changes to " & cbcSelect.List(cbcSelect.ListIndex)
        Else
            slMess = "Add Event Definitions to " & cbcSelect.List(cbcSelect.ListIndex)
        End If
        ilRet = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If ilRet = vbYes Then
            ilRet = mSaveRec()
            Screen.MousePointer = vbDefault
            gSetMousePointer grdSpec, grdDates, vbDefault
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSpecSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case lmEnableCol
        Case LANGUAGEINDEX
            lbcLanguage.Visible = Not lbcLanguage.Visible
        Case VISITTEAMINDEX
            lbcTeam.Visible = Not lbcTeam.Visible
        Case HOMETEAMINDEX
            lbcTeam.Visible = Not lbcTeam.Visible
        Case SUBTOTAL1INDEX
            lbcSubtotal(0).Visible = Not lbcSubtotal(0).Visible
        Case SUBTOTAL2INDEX
            lbcSubtotal(1).Visible = Not lbcSubtotal(1).Visible
        Case LIBRARYINDEX
            lbcLibrary.Visible = Not lbcLibrary.Visible
        Case AIRDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case AIRTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case AIRVEHICLEINDEX
            lbcAirVehicle.Visible = Not lbcAirVehicle.Visible
        Case GAMESTATUSINDEX
            lbcStatus.Visible = Not lbcStatus.Visible
    End Select
    edcDropdown.SelStart = 0
    edcDropdown.SelLength = Len(edcDropdown.Text)
    edcDropdown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcFormats_Click()
    igGameSchdVefCode = imVefCode
    GameLib.Show vbModal
    If igGameLibReturn Then
        imGsfChg = True
        grdDates.Refresh
        mSetCommands
    End If
End Sub

Private Sub cmcFormats_GotFocus()
    mSpecSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcFormats
End Sub

Private Sub cmcMultimedia_Click()
    igGameSchdVefCode = imVefCode
    lgSeasonGhfCode = lmSeasonGhfCode
    GameInv.Show vbModal
End Sub

Private Sub cmcMultimedia_GotFocus()
    mSpecSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcMultimedia
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ilRet = mSaveRec()
    If Not ilRet Then
        DoEvents
        Screen.MousePointer = vbDefault
        gSetMousePointer grdSpec, grdDates, vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSpec, grdDates, vbHourglass
    DoEvents
    mClearCtrlFields
    mSeasonPop
    ilRet = mReadRec(imVefCode, lmSeasonGhfCode)
    mMoveRecToCtrl
    imNewGame = False
    imGhfChg = False
    imGsfChg = False
    imFirstTimeSelect = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdDates, vbDefault
    mSetCommands
End Sub

Private Sub cmcSave_GotFocus()
    mSpecSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcSave
End Sub

Private Sub cmcSpec_Click()
    Select Case lmSpecEnableCol
        Case SEASONSTARTINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case SEASONENDINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
    End Select
    edcSpec.SelStart = 0
    edcSpec.SelLength = Len(edcSpec.Text)
    edcSpec.SetFocus
End Sub

Private Sub cmcSpec_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcSyncGames_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilValue                       ilLang                    *
'*  ilTeam                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim llRow As Long
    Dim ilFound As Integer
    Dim llBlankRow As Long
    Dim tlGhf As GHF
    Dim tlGsf As GSF

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    igVpfType = 2
    VehModel.Show vbModal
    If (igVehReturn = 0) Or (igVefCodeModel = 0) Then    'Cancelled
        mSetCommands
    End If
    grdDates.Redraw = False
    tmGhfSrchKey1.iVefCode = igVefCodeModel
    ilRet = btrGetEqual(hmGhf, tlGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        tmGsfSrchKey1.lghfcode = tlGhf.lCode
        tmGsfSrchKey1.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hmGsf, tlGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tlGhf.lCode = tlGsf.lghfcode)
            ilFound = False
            llBlankRow = -1
            For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                    If tlGsf.iGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX)) Then
                        ilFound = True
                        grdDates.Row = llRow
                        mSetGridValues llRow, tlGsf, True, True
                    End If
                Else
                    If llBlankRow = -1 Then
                        llBlankRow = llRow
                    End If
                End If
            Next llRow
            'If Not ilFound Then
            '    imGhfChg = True
            '    imGsfChg = True
            '    grdSpec.TextMatrix(grdSpec.FixedRows + 1, NOGAMEINDEX) = gAddStr(grdSpec.TextMatrix(grdSpec.FixedRows + 1, NOGAMEINDEX), "1")
            '    ReDim Preserve tmGsf(0 To UBound(tmGsf) + 1) As GSF
            '    tmGsf(UBound(tmGsf) - 1) = tlGsf
            '    tmGsf(UBound(tmGsf) - 1).lCode = 0
            '    If llBlankRow = -1 Then
            '        llRow = grdDates.FixedRows + 2 * (UBound(tmGsf) - 1)
            '    Else
            '        llRow = llBlankRow
            '    End If
            '    If llRow >= grdDates.Rows Then
            '        grdDates.AddItem ""
            '        grdDates.RowHeight(grdDates.Rows - 1) = fgBoxGridH
            '        grdDates.AddItem ""
            '        grdDates.RowHeight(grdDates.Rows - 1) = 15
            '        mInitNew llRow
            '    End If
            '    grdDates.Row = llRow
            '    grdDates.TextMatrix(llRow, TMGSFINDEX) = Trim$(str$(UBound(tmGsf) - 1))
            '    mSetGridValues llRow, tlGsf, True, False
            'End If
            ilRet = btrGetNext(hmGsf, tlGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    End If
    tmGhfSrchKey1.iVefCode = igVefCodeModel
    ilRet = btrGetEqual(hmGhf, tlGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        tmGsfSrchKey1.lghfcode = tlGhf.lCode
        tmGsfSrchKey1.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hmGsf, tlGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tlGhf.lCode = tlGsf.lghfcode)
            ilFound = False
            llBlankRow = -1
            For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                    If tlGsf.iGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX)) Then
                        ilFound = True
                    End If
                Else
                    If llBlankRow = -1 Then
                        llBlankRow = llRow
                    End If
                End If
            Next llRow
            If Not ilFound Then
                imGhfChg = True
                imGsfChg = True
                grdSpec.TextMatrix(grdSpec.FixedRows + 1, NOGAMEINDEX) = gAddStr(grdSpec.TextMatrix(grdSpec.FixedRows + 1, NOGAMEINDEX), "1")
                ReDim Preserve tmGsf(0 To UBound(tmGsf) + 1) As GSF
                tmGsf(UBound(tmGsf) - 1) = tlGsf
                tmGsf(UBound(tmGsf) - 1).lCode = 0
                If llBlankRow = -1 Then
                    llRow = grdDates.FixedRows + 2 * (UBound(tmGsf) - 1)
                Else
                    llRow = llBlankRow
                End If
                If llRow >= grdDates.Rows Then
                    grdDates.AddItem ""
                    grdDates.RowHeight(grdDates.Rows - 1) = fgBoxGridH
                    grdDates.AddItem ""
                    grdDates.RowHeight(grdDates.Rows - 1) = 15
                    mInitNew llRow
                End If
                grdDates.Row = llRow
                grdDates.TextMatrix(llRow, TMGSFINDEX) = Trim$(str$(UBound(tmGsf) - 1))
                mSetGridValues llRow, tlGsf, True, False
            End If
            ilRet = btrGetNext(hmGsf, tlGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    End If
    grdDates.Redraw = True
    mSetCommands
End Sub

Private Sub cmcSyncGames_GotFocus()
    mSpecSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcFormats
End Sub

Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer

    Select Case lmEnableCol
        Case LANGUAGEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropdown, lbcLanguage, imBSMode, slStr)
            If ilRet = 1 Then
                If lbcLanguage.ListCount > 0 Then
                    lbcLanguage.ListIndex = 0
                End If
            End If
        Case VISITTEAMINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropdown, lbcTeam, imBSMode, slStr)
            If ilRet = 1 Then
                If lbcTeam.ListCount > 0 Then
                    lbcTeam.ListIndex = 0
                End If
            End If
        Case HOMETEAMINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropdown, lbcTeam, imBSMode, slStr)
            If ilRet = 1 Then
                If lbcTeam.ListCount > 0 Then
                    lbcTeam.ListIndex = 0
                End If
            End If
        Case SUBTOTAL1INDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropdown, lbcSubtotal(0), imBSMode, slStr)
            If ilRet = 1 Then
                If lbcSubtotal(0).ListCount > 0 Then
                    lbcSubtotal(0).ListIndex = 0
                End If
            End If
        Case SUBTOTAL2INDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropdown, lbcSubtotal(1), imBSMode, slStr)
            If ilRet = 1 Then
                If lbcSubtotal(1).ListCount > 0 Then
                    lbcSubtotal(1).ListIndex = 0
                End If
            End If
        Case LIBRARYINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropdown, lbcLibrary, imBSMode, imComboBoxIndex
        Case AIRDATEINDEX
            slStr = edcDropdown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case AIRTIMEINDEX
        Case AIRVEHICLEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropdown, lbcAirVehicle, imBSMode, imComboBoxIndex
        Case XDSPROGCODEINDEX
        Case BUSINDEX
        Case GAMESTATUSINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropdown, lbcStatus, imBSMode, imComboBoxIndex
    End Select
    grdDates.CellForeColor = vbBlack
    imLbcArrowSetting = False
End Sub

Private Sub edcDropDown_DblClick()
    Select Case lmEnableCol
        Case LANGUAGEINDEX
            imDoubleClickName = True
        Case VISITTEAMINDEX
            imDoubleClickName = True
        Case HOMETEAMINDEX
            imDoubleClickName = True
        Case SUBTOTAL1INDEX
            imDoubleClickName = True
        Case SUBTOTAL2INDEX
            imDoubleClickName = True
        Case LIBRARYINDEX
        Case AIRDATEINDEX
        Case AIRTIMEINDEX
        Case AIRVEHICLEINDEX
        Case XDSPROGCODEINDEX
        Case BUSINDEX
        Case GAMESTATUSINDEX
    End Select
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmEnableCol
        Case LANGUAGEINDEX
        Case VISITTEAMINDEX
        Case HOMETEAMINDEX
        Case SUBTOTAL1INDEX
        Case SUBTOTAL2INDEX
        Case LIBRARYINDEX
        Case AIRDATEINDEX
        Case AIRTIMEINDEX
        Case AIRVEHICLEINDEX
        Case XDSPROGCODEINDEX
        Case BUSINDEX
        Case GAMESTATUSINDEX
    End Select
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
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropdown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case lmEnableCol
        Case LANGUAGEINDEX
        Case VISITTEAMINDEX
        Case HOMETEAMINDEX
        Case SUBTOTAL1INDEX
        Case SUBTOTAL2INDEX
        Case LIBRARYINDEX
        Case AIRDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case AIRTIMEINDEX
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
        Case AIRVEHICLEINDEX
        Case XDSPROGCODEINDEX
        Case BUSINDEX
        Case GAMESTATUSINDEX
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmEnableCol
            Case LANGUAGEINDEX
                gProcessArrowKey Shift, KeyCode, lbcLanguage, imLbcArrowSetting
            Case VISITTEAMINDEX
                gProcessArrowKey Shift, KeyCode, lbcTeam, imLbcArrowSetting
            Case HOMETEAMINDEX
                gProcessArrowKey Shift, KeyCode, lbcTeam, imLbcArrowSetting
            Case SUBTOTAL1INDEX
                gProcessArrowKey Shift, KeyCode, lbcSubtotal(0), imLbcArrowSetting
            Case SUBTOTAL2INDEX
                gProcessArrowKey Shift, KeyCode, lbcSubtotal(1), imLbcArrowSetting
            Case LIBRARYINDEX
                gProcessArrowKey Shift, KeyCode, lbcLibrary, imLbcArrowSetting
            Case AIRDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropdown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropdown.Text = slDate
                    End If
                End If
            Case AIRTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case AIRVEHICLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcAirVehicle, imLbcArrowSetting
            Case XDSPROGCODEINDEX
            Case BUSINDEX
            Case GAMESTATUSINDEX
                gProcessArrowKey Shift, KeyCode, lbcStatus, imLbcArrowSetting
        End Select
        edcDropdown.SelStart = 0
        edcDropdown.SelLength = Len(edcDropdown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case lmEnableCol
            Case LANGUAGEINDEX
            Case VISITTEAMINDEX
            Case HOMETEAMINDEX
            Case SUBTOTAL1INDEX
            Case SUBTOTAL2INDEX
            Case LIBRARYINDEX
            Case AIRDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropdown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropdown.Text = slDate
                    End If
                End If
                edcDropdown.SelStart = 0
                edcDropdown.SelLength = Len(edcDropdown.Text)
            Case AIRTIMEINDEX
            Case AIRVEHICLEINDEX
            Case XDSPROGCODEINDEX
            Case BUSINDEX
            Case GAMESTATUSINDEX
        End Select
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case lmEnableCol
            Case LANGUAGEINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case VISITTEAMINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case HOMETEAMINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case SUBTOTAL1INDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case SUBTOTAL2INDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case LIBRARYINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
            Case AIRDATEINDEX
            Case AIRTIMEINDEX
            Case AIRVEHICLEINDEX
            Case XDSPROGCODEINDEX
            Case BUSINDEX
            Case GAMESTATUSINDEX
        End Select
        imDoubleClickName = False
    End If
End Sub

Private Sub edcSpec_Change()
    Dim slStr As String
    
    Select Case lmSpecEnableCol
        Case SEASONSTARTINDEX
            slStr = edcSpec.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case SEASONENDINDEX
            slStr = edcSpec.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case NOGAMEINDEX
    End Select

End Sub

Private Sub edcSpec_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSpec_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcSpec_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSpec.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case lmSpecEnableCol
        Case SEASONSTARTINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case SEASONENDINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case NOGAMEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcSpec.Text
            slStr = Left$(slStr, edcSpec.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSpec.SelStart - edcSpec.SelLength)
            If gCompNumberStr(slStr, "15000") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub edcSpec_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmSpecEnableCol
            Case SEASONSTARTINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcSpec.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpec.Text = slDate
                    End If
                End If
            Case SEASONENDINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcSpec.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpec.Text = slDate
                    End If
                End If
        End Select
        edcSpec.SelStart = 0
        edcSpec.SelLength = Len(edcSpec.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case lmEnableCol
            Case SEASONSTARTINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcSpec.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpec.Text = slDate
                    End If
                End If
                edcSpec.SelStart = 0
                edcSpec.SelLength = Len(edcSpec.Text)
            Case SEASONENDINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcSpec.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcSpec.Text = slDate
                    End If
                End If
                edcSpec.SelStart = 0
                edcSpec.SelLength = Len(edcSpec.Text)
        End Select
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
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        grdSpec.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        grdDates.Enabled = False
        imUpdateAllowed = False
    Else
        grdSpec.Enabled = True
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        grdDates.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_Initialize()
    'Use about 80% of the screen
    Me.Width = (CLng(80) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(80) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    gCenterStdAlone GameSchd
    DoEvents
    mSetControls
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        If lmEnableCol > 0 Then
            mEnableBox
        End If
    End If
End Sub

Private Sub Form_Load()
    'mSetControls
    imInitNoRows = grdDates.Rows
    mInit
End Sub

Private Sub Form_Resize()
    'mSetControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    
    Erase tmGsf
    Erase bmChgFlag
    Erase tmSdfMdExt
    Erase lmSdfCode
    Erase lmBBSdfCode
    Erase lmSmfCode
    Erase lmClfCode
    smGameVehicleTag = ""
    Erase tmGameVehicle
    smAirVehicleTag = ""
    Erase tmAirVehicle
    smLanguageCodeTag = ""
    Erase tmLanguageCode
    smTeamCodeTag = ""
    Erase tmTeamCode
    smLibNameTag = ""
    Erase tmLibName

    ilRet = btrClose(hmLst)
    btrDestroy hmLst
    ilRet = btrClose(hmEcf)
    btrDestroy hmEcf
    ilRet = btrClose(hmVff)
    btrDestroy hmVff
    ilRet = btrClose(hmIif)
    btrDestroy hmIif
    ilRet = btrClose(hmItf)
    btrDestroy hmItf
    ilRet = btrClose(hmIhf)
    btrDestroy hmIhf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    ilRet = btrClose(hmMgf)
    btrDestroy hmMgf
    ilRet = btrClose(hmMsf)
    btrDestroy hmMsf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmLvf)
    btrDestroy hmLvf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmRlf)
    btrDestroy hmRlf
    ilRet = btrClose(hmSxf)
    btrDestroy hmSxf
    
    Set GameSchd = Nothing   'Remove data segment

End Sub

Private Sub grdDates_EnterCell()
    mSpecSetShow
    mSetShow
End Sub

Private Sub grdDates_GotFocus()
    mSpecSetShow
End Sub

Private Sub grdDates_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdDates.TopRow
    grdDates.Redraw = False
End Sub

Private Sub grdDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine if in header
''    If y < grdDates.RowHeight(0) Then
''        mSortCol grdDates.Col
''        Exit Sub
''    End If
'    If Y < grdDates.RowHeight(0) + grdDates.RowHeight(1) + grdDates.RowHeight(2) + grdDates.RowHeight(3) Then
'        grdDates.Col = grdDates.MouseCol
'        mSortCol grdDates.Col
'        Exit Sub
'    End If
    'Determine if in header
    If Y < grdDates.RowHeight(0) + grdDates.RowHeight(1) + grdDates.RowHeight(2) + grdDates.RowHeight(3) Then
        mSpecSetShow
        mSetShow
        lmEnableRow = -1
        lmEnableCol = -1
        grdDates.Col = grdDates.MouseCol
        mSortCol grdDates.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    On Error GoTo grdDatesErr:
    ilCol = grdDates.MouseCol
    ilRow = grdDates.MouseRow
    If ilCol < grdDates.FixedCols Then
        grdDates.Redraw = True
        Exit Sub
    End If
    If ilRow < grdDates.FixedRows Then
        grdDates.Redraw = True
        Exit Sub
    End If
    If ilRow Mod 2 = 0 Then
        ilRow = ilRow + 1
    End If
    If grdDates.ColWidth(ilCol) <= 15 Then
        grdDates.Redraw = True
        Exit Sub
    End If
    If grdDates.RowHeight(ilRow) <= 15 Then
        grdDates.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdDates.TopRow
    DoEvents
    If grdDates.TextMatrix(ilRow, GAMENOINDEX) = "" Then
        grdDates.Redraw = True
        Exit Sub
    End If
    grdDates.Col = ilCol
    grdDates.Row = ilRow
    If Not mColOk() Then
        grdDates.Redraw = True
        Exit Sub
    End If
    grdDates.Redraw = True
    mEnableBox
    On Error GoTo 0
    Exit Sub
grdDatesErr:
    On Error GoTo 0
    If (lmEnableRow >= grdDates.FixedRows) And (lmEnableRow < grdDates.Rows) Then
        grdDates.Row = lmEnableRow
        grdDates.Col = lmEnableCol
        mSetFocus
    End If
    grdDates.Redraw = False
    grdDates.Redraw = True
    Exit Sub
End Sub

Private Sub grdDates_Scroll()
    If imSettingValue Then
        Exit Sub
    End If
    If grdDates.Redraw = False Then
        grdDates.Redraw = True
        If lmTopRow < grdDates.FixedRows Then
            grdDates.TopRow = grdDates.FixedRows
        Else
            grdDates.TopRow = lmTopRow
        End If
        grdDates.Refresh
        'grdDates.Redraw = False
    End If
    If grdDates.RowHeight(grdDates.TopRow) <= 15 Then
        If grdDates.TopRow > lmTopRow Then
            grdDates.TopRow = grdDates.TopRow + 1
        ElseIf grdDates.TopRow < lmTopRow Then
            grdDates.TopRow = grdDates.TopRow - 1
        End If
    End If
    If (imCtrlVisible) And (grdDates.Row >= grdDates.FixedRows) And (grdDates.Col >= grdDates.FixedCols) Then
        If grdDates.RowIsVisible(grdDates.Row) Then
            mSetFocus
        Else
            pbcSetFocus.SetFocus
            cmcDropDown.Visible = False
            edcDropdown.Visible = False
            pbcFeed.Visible = False
            lbcLanguage.Visible = False
            lbcTeam.Visible = False
            lbcLibrary.Visible = False
            plcCalendar.Visible = False
            plcTme.Visible = False
            lbcAirVehicle.Visible = False
            lbcStatus.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
    End If
    lmTopRow = grdDates.TopRow
End Sub

Private Sub grdSpec_EnterCell()
    mSpecSetShow
    mSetShow
End Sub

Private Sub grdSpec_GotFocus()
    mSetShow
End Sub

Private Sub grdSpec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'lmTopRow = grdSpec.TopRow
    grdSpec.Redraw = False
End Sub

Private Sub grdSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine if in header
'    If y < grdSpec.RowHeight(0) Then
'        mSortCol grdSpec.Col
'        Exit Sub
'    End If
    'Determine row and col mouse up onto
    On Error GoTo grdSpecErr:
    ilCol = grdSpec.MouseCol
    ilRow = grdSpec.MouseRow
    If ilCol < grdSpec.FixedCols Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    If ilRow < grdSpec.FixedRows Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    If ilRow Mod 2 = 0 Then
        ilRow = ilRow + 1
    End If
    If grdSpec.ColWidth(ilCol) <= 15 Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    If grdSpec.RowHeight(ilRow) <= 15 Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    'lmTopRow = grdSpec.TopRow
    DoEvents
    grdSpec.Col = ilCol
    grdSpec.Row = ilRow
    grdSpec.Redraw = True
    mSpecEnableBox
    On Error GoTo 0
    Exit Sub
grdSpecErr:
    On Error GoTo 0
    If (lmSpecEnableRow >= grdSpec.FixedRows) And (lmSpecEnableRow < grdSpec.Rows) Then
        grdSpec.Row = lmSpecEnableRow
        grdSpec.Col = lmSpecEnableCol
        mSpecSetFocus
    End If
    grdSpec.Redraw = False
    grdSpec.Redraw = True
    Exit Sub
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    slStr = edcDropdown.Text
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
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (grdDates.Row < grdDates.FixedRows) Or (grdDates.Row >= grdDates.Rows) Or (grdDates.Col < grdDates.FixedCols) Or (grdDates.Col >= grdDates.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdDates.Row
    lmEnableCol = grdDates.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdDates.Left - pbcArrow.Width - 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + (grdDates.RowHeight(grdDates.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    Select Case grdDates.Col
        Case FEEDSOURCEINDEX
            smFeedSource = grdDates.Text
            Select Case smFeedSource
                Case "V"
                    smFeedSource = Trim$(tgSaf(0).sEventTitle1)    '"Visiting"
                Case "H"
                    smFeedSource = Trim$(tgSaf(0).sEventTitle2)    '"Home"
                Case "N"
                    smFeedSource = "National"
            End Select
            If (smFeedSource = "") Or (smFeedSource = "Missing") Then
                smFeedSource = Trim$(tgSaf(0).sEventTitle2)    '"Home"
            End If
        Case LANGUAGEINDEX
            lbcLanguage.Height = gListBoxHeight(lbcLanguage.ListCount, 10)
            edcDropdown.MaxLength = 20
            imChgMode = True
            slStr = grdDates.Text
            gFindMatch slStr, 1, lbcLanguage
            If gLastFound(lbcLanguage) >= 1 Then
                lbcLanguage.ListIndex = gLastFound(lbcLanguage)
                edcDropdown.Text = lbcLanguage.List(lbcLanguage.ListIndex)
            Else
                If lbcLanguage.ListCount > 1 Then
                    If grdDates.Row > grdDates.FixedRows Then
                        slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                        gFindMatch slStr, 1, lbcLanguage
                        If gLastFound(lbcLanguage) >= 1 Then
                            lbcLanguage.ListIndex = gLastFound(lbcLanguage)
                        Else
                            lbcLanguage.ListIndex = 1
                        End If
                        edcDropdown.Text = lbcLanguage.List(lbcLanguage.ListIndex)
                    Else
                        lbcLanguage.ListIndex = 1
                    End If
                    edcDropdown.Text = lbcLanguage.List(lbcLanguage.ListIndex)
                End If
            End If
            imChgMode = False
        Case VISITTEAMINDEX
            lbcTeam.Height = gListBoxHeight(lbcTeam.ListCount, 10)
            edcDropdown.MaxLength = 20
            imChgMode = True
            slStr = grdDates.Text
            gFindMatch slStr, 1, lbcTeam
            If gLastFound(lbcTeam) >= 1 Then
                lbcTeam.ListIndex = gLastFound(lbcTeam)
                edcDropdown.Text = lbcTeam.List(lbcTeam.ListIndex)
            Else
                If lbcTeam.ListCount > 1 Then
                    If grdDates.Row > grdDates.FixedRows Then
                        slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                        gFindMatch slStr, 1, lbcTeam
                        If gLastFound(lbcTeam) >= 1 Then
                            lbcTeam.ListIndex = gLastFound(lbcTeam)
                        Else
                            lbcTeam.ListIndex = 1
                        End If
                        edcDropdown.Text = lbcTeam.List(lbcTeam.ListIndex)
                    Else
                        lbcTeam.ListIndex = 1
                    End If
                    edcDropdown.Text = lbcTeam.List(lbcTeam.ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imChgMode = False
        Case HOMETEAMINDEX
            lbcTeam.Height = gListBoxHeight(lbcTeam.ListCount, 10)
            edcDropdown.MaxLength = 20
            imChgMode = True
            slStr = grdDates.Text
            gFindMatch slStr, 1, lbcTeam
            If gLastFound(lbcTeam) >= 1 Then
                lbcTeam.ListIndex = gLastFound(lbcTeam)
                edcDropdown.Text = lbcTeam.List(lbcTeam.ListIndex)
            Else
                If lbcTeam.ListCount > 1 Then
                    If grdDates.Row > grdDates.FixedRows Then
                        slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                        gFindMatch slStr, 1, lbcTeam
                        If gLastFound(lbcTeam) >= 1 Then
                            lbcTeam.ListIndex = gLastFound(lbcTeam)
                        Else
                            lbcTeam.ListIndex = 1
                        End If
                        edcDropdown.Text = lbcTeam.List(lbcTeam.ListIndex)
                    Else
                        lbcTeam.ListIndex = 1
                    End If
                    edcDropdown.Text = lbcTeam.List(lbcTeam.ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imChgMode = False
        Case SUBTOTAL1INDEX
            lbcSubtotal(0).Height = gListBoxHeight(lbcSubtotal(0).ListCount, 10)
            edcDropdown.MaxLength = 20
            imChgMode = True
            slStr = grdDates.Text
            gFindMatch slStr, 1, lbcSubtotal(0)
            If gLastFound(lbcSubtotal(0)) >= 1 Then
                lbcSubtotal(0).ListIndex = gLastFound(lbcSubtotal(0))
                edcDropdown.Text = lbcSubtotal(0).List(lbcSubtotal(0).ListIndex)
            Else
                If lbcSubtotal(0).ListCount > 1 Then
                    If grdDates.Row > grdDates.FixedRows Then
                        slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                        gFindMatch slStr, 1, lbcSubtotal(0)
                        If gLastFound(lbcSubtotal(0)) >= 1 Then
                            lbcSubtotal(0).ListIndex = gLastFound(lbcSubtotal(0))
                        Else
                            lbcSubtotal(0).ListIndex = 1
                        End If
                        edcDropdown.Text = lbcSubtotal(0).List(lbcSubtotal(0).ListIndex)
                    Else
                        lbcSubtotal(0).ListIndex = 1
                    End If
                    edcDropdown.Text = lbcSubtotal(0).List(lbcSubtotal(0).ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imChgMode = False
        Case SUBTOTAL2INDEX
            lbcSubtotal(1).Height = gListBoxHeight(lbcSubtotal(1).ListCount, 10)
            edcDropdown.MaxLength = 20
            imChgMode = True
            slStr = grdDates.Text
            gFindMatch slStr, 1, lbcSubtotal(1)
            If gLastFound(lbcSubtotal(1)) >= 1 Then
                lbcSubtotal(1).ListIndex = gLastFound(lbcSubtotal(1))
                edcDropdown.Text = lbcSubtotal(1).List(lbcSubtotal(1).ListIndex)
            Else
                If lbcSubtotal(1).ListCount > 1 Then
                    If grdDates.Row > grdDates.FixedRows Then
                        slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                        gFindMatch slStr, 1, lbcSubtotal(1)
                        If gLastFound(lbcSubtotal(1)) >= 1 Then
                            lbcSubtotal(1).ListIndex = gLastFound(lbcSubtotal(1))
                        Else
                            lbcSubtotal(1).ListIndex = 1
                        End If
                        edcDropdown.Text = lbcSubtotal(1).List(lbcSubtotal(1).ListIndex)
                    Else
                        lbcSubtotal(1).ListIndex = 1
                    End If
                    edcDropdown.Text = lbcSubtotal(1).List(lbcSubtotal(1).ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imChgMode = False
        Case LIBRARYINDEX
            lbcLibrary.Height = gListBoxHeight(lbcLibrary.ListCount, 10)
            edcDropdown.MaxLength = 20
            imChgMode = True
            'slStr = grdDates.Text
            slStr = grdDates.TextMatrix(grdDates.Row, VERLIBINDEX)
            gFindMatch slStr, 0, lbcLibrary
            If gLastFound(lbcLibrary) >= 0 Then
                lbcLibrary.ListIndex = gLastFound(lbcLibrary)
                edcDropdown.Text = lbcLibrary.List(lbcLibrary.ListIndex)
            Else
                If lbcLibrary.ListCount > 0 Then
                    If grdDates.Row > grdDates.FixedRows Then
                        slStr = grdDates.TextMatrix(grdDates.Row - 2, VERLIBINDEX)
                        gFindMatch slStr, 0, lbcLibrary
                        If gLastFound(lbcLibrary) >= 0 Then
                            lbcLibrary.ListIndex = gLastFound(lbcLibrary)
                        Else
                            lbcLibrary.ListIndex = 0
                        End If
                        edcDropdown.Text = lbcLibrary.List(lbcLibrary.ListIndex)
                    Else
                        lbcLibrary.ListIndex = 0
                    End If
                    edcDropdown.Text = lbcLibrary.List(lbcLibrary.ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imChgMode = False
        Case AIRDATEINDEX
            edcDropdown.MaxLength = 10
            slStr = grdDates.Text
            If (slStr = "") Or (slStr = "Missing") Then
                If grdDates.Row > grdDates.FixedRows Then
                    slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                    If slStr = "" Then
                        slStr = gObtainMondayFromToday()
                    End If
                Else
                    slStr = grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX)
                End If
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropdown.Text = slStr
        Case AIRTIMEINDEX
            edcDropdown.MaxLength = 10
            slStr = grdDates.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdDates.Row > grdDates.FixedRows Then
                    slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                End If
            End If
            edcDropdown.Text = slStr
        Case AIRVEHICLEINDEX
            smLiveLogMerge = Left$(grdDates.Text, 2)
            Select Case smLiveLogMerge
                Case "L:"
                    smLiveLogMerge = "Live Log & Pre-empt"
                Case "M:"
                    smLiveLogMerge = "Merge & Pre-empt"
                Case Else
                    smLiveLogMerge = "Live Log & Pre-empt"
            End Select
            lbcAirVehicle.Height = gListBoxHeight(lbcAirVehicle.ListCount, 10)
            edcDropdown.MaxLength = 40
            imChgMode = True
            slStr = grdDates.Text
            gFindMatch slStr, 0, lbcAirVehicle
            If gLastFound(lbcAirVehicle) >= 0 Then
                lbcAirVehicle.ListIndex = gLastFound(lbcAirVehicle)
                edcDropdown.Text = lbcAirVehicle.List(lbcAirVehicle.ListIndex)
            Else
                If lbcAirVehicle.ListCount >= 1 Then
                    If grdDates.Row > grdDates.FixedRows Then
                        slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                        gFindMatch slStr, 0, lbcAirVehicle
                        If gLastFound(lbcAirVehicle) >= 0 Then
                            lbcAirVehicle.ListIndex = gLastFound(lbcAirVehicle)
                        Else
                            lbcAirVehicle.ListIndex = 0
                        End If
                        edcDropdown.Text = lbcAirVehicle.List(lbcAirVehicle.ListIndex)
                    Else
                        lbcAirVehicle.ListIndex = 0
                    End If
                    edcDropdown.Text = lbcAirVehicle.List(lbcAirVehicle.ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imChgMode = False
        Case XDSPROGCODEINDEX
            edcDropdown.MaxLength = 8
            slStr = grdDates.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdDates.Row > grdDates.FixedRows Then
                    slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                End If
            End If
            edcDropdown.Text = slStr
        Case BUSINDEX
            edcDropdown.MaxLength = 20
            slStr = grdDates.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdDates.Row > grdDates.FixedRows Then
                    slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                End If
            End If
            edcDropdown.Text = slStr
        Case GAMESTATUSINDEX
            lbcStatus.Height = gListBoxHeight(lbcStatus.ListCount, 10)
            edcDropdown.MaxLength = 9
            imChgMode = True
            slStr = grdDates.Text
            Select Case UCase(slStr)
                Case "C"
                    slStr = "Canceled"
                Case "F"
                    slStr = "Firm"
                Case "P"
                    slStr = "Postponed"
                Case "T"
                    slStr = "Tentative"
            End Select
            gFindMatch slStr, 0, lbcStatus
            If gLastFound(lbcStatus) >= 0 Then
                lbcStatus.ListIndex = gLastFound(lbcStatus)
                edcDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
            Else
                If lbcStatus.ListCount >= 1 Then
                    If grdDates.Row > grdDates.FixedRows Then
                        slStr = grdDates.TextMatrix(grdDates.Row - 2, grdDates.Col)
                        gFindMatch slStr, 0, lbcStatus
                        If gLastFound(lbcStatus) >= 0 Then
                            lbcStatus.ListIndex = gLastFound(lbcStatus)
                        Else
                            lbcStatus.ListIndex = 0
                        End If
                        edcDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                    Else
                        lbcStatus.ListIndex = 0
                    End If
                    edcDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                Else
                    edcDropdown.Text = ""
                End If
            End If
            imChgMode = False
    End Select
    mSetFocus
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
    Dim slName As String

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    gSetMousePointer grdSpec, grdDates, vbHourglass
    imLBSdfMdExt = 1
    imFirstActivate = True
    imTerminate = False
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    imBypassFocus = False
    imSettingValue = False
    imStartMode = True
    imChgMode = False
    imBSMode = False
    imLbcArrowSetting = False
    imLbcMouseDown = False
    imTabDirection = 0  'Left to right movement
    imDoubleClickName = False
    imCalType = 0   'Standard
    imCtrlVisible = False
    imGsfChg = False
    imGhfChg = False
    imNewGame = True
    imInNew = False
    lmSeasonGhfCode = 0
    imFirstTimeSelect = True
    imSpecCtrlVisible = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmLockRecCode = 0
    mInitBox
    hmGhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)  'Get and save ARF record length

    hmGsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    ReDim tmGsf(0 To 0) As GSF
    imGsfRecLen = Len(tmGsf(0))  'Get and save ARF record length

    hmLcf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)  'Get and save ARF record length

    hmLvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imLvfRecLen = Len(tmLvf)  'Get and save ARF record length

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save ARF record length

    hmClf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imClfRecLen = Len(tmClf)  'Get and save ARF record length

    hmCff = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imCffRecLen = Len(tmCff)  'Get and save ARF record length

    hmCgf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imCgfRecLen = Len(tmCgf)  'Get and save ARF record length

    hmSsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    hmSdf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)  'Get and save ARF record length
    hmSmf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)  'Get and save ARF record length

    hmRlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0

    hmSxf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0

    hmMsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMsf, "", sgDBPath & "Msf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imMsfRecLen = Len(tmMsf)  'Get and save ARF record length


    hmMgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMgf, "", sgDBPath & "Mgf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imMgfRecLen = Len(tmMgf)  'Get and save ARF record length


    hmSbf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf)  'Get and save ARF record length

    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)  'Get and save ARF record length

    hmIhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmIhf, "", sgDBPath & "Ihf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imIhfRecLen = Len(tmIhf)  'Get and save ARF record length


    hmItf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmItf, "", sgDBPath & "Itf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imItfRecLen = Len(tmItf)  'Get and save ARF record length


    hmIif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmIif, "", sgDBPath & "Iif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imIifRecLen = Len(tmIif)  'Get and save ARF record length

    hmVff = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmVff, "", sgDBPath & "Vff.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imVffRecLen = Len(tmVff)  'Get and save ARF record length

    hmEcf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmEcf, "", sgDBPath & "Ecf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imEcfRecLen = Len(tmEcf)  'Get and save LST record length
    
    hmLst = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmLst, "", sgDBPath & "Lst.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameSchd
    On Error GoTo 0
    imLstRecLen = Len(tmLst)  'Get and save LST record length

    mTeamPop
    mLanguagePop
    mLibraryPop
    mSubtotalPop -1
    lbcStatus.AddItem "Tentative"
    lbcStatus.AddItem "Firm"
    lbcStatus.AddItem "Postponed"
    lbcStatus.AddItem "Canceled"

    'mXFerRecToCtrl
'    GameSchd.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    mVehPop
    ilRet = gVffRead()
    If sgCommandStr <> "" Then
        ilRet = gParseItem(sgCommandStr, 1, "\", slName)
        gFindMatch slName, 0, cbcSelect
        If gLastFound(cbcSelect) >= 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        End If
    End If
    imNTRMnfCode = mAddMultiMediaNTR()
'    gCenterStdAlone GameSchd
    mSetCommands
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdDates, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdDates, vbDefault
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  flTextHeight                  ilRow                         ilCol                     *
'*                                                                                        *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer

'    grdSpec.Move 180, 435
'    mGridSpecLayout
'    mGridSpecColumnWidths
'    mGridSpecColumns
'
'    grdDates.Move grdSpec.Left, grdSpec.Top + grdSpec.Height + 120
'    imInitNoRows = grdDates.Rows
'    mGridLayout
'    mGridColumnWidths
'    mGridColumns
'    grdDates.Height = grdDates.RowPos(grdDates.Rows - 1) + grdDates.RowHeight(grdDates.Rows - 1) + fgPanelAdj - 15
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    pbcStartNew.Top = -pbcStartNew.Height - 90
    pbcSpecSTab.Left = -pbcSpecSTab.Width - 90
    pbcSpecTab.Left = -pbcSpecTab.Width - 90
    pbcSTab.Left = -pbcSTab.Width - 90
    pbcTab.Left = -pbcTab.Width - 90
    pbcClickFocus.Left = -pbcClickFocus.Width - 90
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize values              *
'*                                                     *
'*******************************************************
Private Sub mInitNew(llRowNo As Long)
    Dim ilCol As Integer
    Dim llRow As Long

    For ilCol = grdDates.FixedCols To grdDates.Cols - 1 Step 1
        grdDates.TextMatrix(llRowNo, ilCol) = ""
    Next ilCol
    'grdDates.Row = llRowNo
    'grdDates.Col = 1
    'grdDates.CellBackColor = vbWhite
    'Horizontal Line
    grdDates.Row = llRowNo + 1
    For ilCol = 1 To grdDates.Cols - 1 Step 1
        grdDates.Col = ilCol
        grdDates.CellBackColor = vbBlue
    Next ilCol
    'Vertical Lines
    grdDates.Col = 1
    For llRow = llRowNo To llRowNo + 1 Step 1
        grdDates.Row = llRow
        grdDates.CellBackColor = vbBlue
    Next llRow
    grdDates.Col = 3
    For llRow = llRowNo To llRowNo + 1 Step 1
        grdDates.Row = llRow
        grdDates.CellBackColor = vbBlue
    Next llRow
    For ilCol = grdDates.FixedCols + 1 To grdDates.Cols - 1 Step 2
        grdDates.Col = ilCol
        For llRow = llRowNo To llRowNo + 1 Step 1
            grdDates.Row = llRow
            grdDates.CellBackColor = vbBlue
        Next llRow
    Next ilCol
    'Set Fix area Column to white
    grdDates.Col = 2
    grdDates.Row = llRowNo
    grdDates.CellBackColor = vbWhite
    If grdDates.RowHeight(grdDates.TopRow) <= 15 Then
        grdDates.TopRow = grdDates.TopRow + 1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilSvRow As Integer
    Dim ilSvCol As Integer
    Dim ilCol As Integer
    Dim ilSvGsfChg As Integer
    Dim ilRet As Integer

    ilSvGsfChg = imGsfChg
    imGsfChg = False
    pbcArrow.Visible = False
    If (lmEnableRow >= grdDates.FixedRows) And (lmEnableRow < grdDates.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case FEEDSOURCEINDEX
                pbcFeed.Visible = False
                slStr = smFeedSource    'Left$(smFeedSource, 1)
                If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imGsfChg = True
                End If
                grdDates.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case LANGUAGEINDEX
                edcDropdown.Visible = False
                cmcDropDown.Visible = False
                lbcLanguage.Visible = False
                If lbcLanguage.ListIndex > 0 Then
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> lbcLanguage.List(lbcLanguage.ListIndex) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = lbcLanguage.List(lbcLanguage.ListIndex)
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case VISITTEAMINDEX
                edcDropdown.Visible = False
                cmcDropDown.Visible = False
                lbcTeam.Visible = False
                If lbcTeam.ListIndex > 0 Then
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> lbcTeam.List(lbcTeam.ListIndex) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = lbcTeam.List(lbcTeam.ListIndex)
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case HOMETEAMINDEX
                edcDropdown.Visible = False
                cmcDropDown.Visible = False
                lbcTeam.Visible = False
                If lbcTeam.ListIndex > 0 Then
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> lbcTeam.List(lbcTeam.ListIndex) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = lbcTeam.List(lbcTeam.ListIndex)
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case SUBTOTAL1INDEX
                edcDropdown.Visible = False
                cmcDropDown.Visible = False
                lbcSubtotal(0).Visible = False
                If lbcSubtotal(0).ListIndex > 0 Then
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> lbcSubtotal(0).List(lbcSubtotal(0).ListIndex) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = lbcSubtotal(0).List(lbcSubtotal(0).ListIndex)
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case SUBTOTAL2INDEX
                edcDropdown.Visible = False
                cmcDropDown.Visible = False
                lbcSubtotal(1).Visible = False
                If lbcSubtotal(1).ListIndex > 0 Then
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> lbcSubtotal(1).List(lbcSubtotal(1).ListIndex) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = lbcSubtotal(1).List(lbcSubtotal(1).ListIndex)
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case LIBRARYINDEX
                edcDropdown.Visible = False
                cmcDropDown.Visible = False
                lbcLibrary.Visible = False
                If lbcLibrary.ListIndex >= 0 Then
                    If grdDates.TextMatrix(lmEnableRow, VERLIBINDEX) <> lbcLibrary.List(lbcLibrary.ListIndex) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, VERLIBINDEX) = lbcLibrary.List(lbcLibrary.ListIndex)
                    If ckcShowVersion.Value = vbChecked Then
                        grdDates.TextMatrix(lmEnableRow, lmEnableCol) = lbcLibrary.List(lbcLibrary.ListIndex)
                    Else
                        slStr = lbcLibrary.List(lbcLibrary.ListIndex)
                        ilPos = InStr(1, slStr, "/", vbTextCompare)
                        If ilPos > 0 Then
                            slStr = Mid$(slStr, ilPos + 1)
                        End If
                        grdDates.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    End If
                Else
                    If grdDates.TextMatrix(lmEnableRow, VERLIBINDEX) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, VERLIBINDEX) = ""
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case AIRDATEINDEX
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                edcDropdown.Visible = False  'Set visibility
                slStr = edcDropdown.Text
                If gValidDate(slStr) Then
                    If gDateValue(grdDates.TextMatrix(lmEnableRow, lmEnableCol)) <> gDateValue(slStr) Then
                        imGsfChg = True
                    End If
                    slStr = gFormatDate(slStr)
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                Else
                    Beep
                End If
            Case AIRTIMEINDEX
                cmcDropDown.Visible = False
                plcTme.Visible = False
                edcDropdown.Visible = False  'Set visibility
                slStr = edcDropdown.Text
                If slStr <> "" Then
                    If gValidTime(slStr) Then
                        If gTimeToLong(grdDates.TextMatrix(lmEnableRow, lmEnableCol), False) <> gTimeToLong(slStr, False) Then
                            imGsfChg = True
                        End If
                        slStr = gFormatTime(slStr, "A", "1")
                        grdDates.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    Else
                        Beep
                    End If
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case AIRVEHICLEINDEX
                edcDropdown.Visible = False
                cmcDropDown.Visible = False
                lbcAirVehicle.Visible = False
                If pbcLiveLogMerge.Visible Then
                    pbcLiveLogMerge.Visible = False
                    slStr = Left$(smLiveLogMerge, 1)
                    slStr = slStr & ": "
                    If Left$(grdDates.TextMatrix(lmEnableRow, lmEnableCol), 2) <> slStr Then
                        imGsfChg = True
                    End If
                Else
                    slStr = ""
                End If
                If lbcAirVehicle.ListIndex > 0 Then
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> slStr & lbcAirVehicle.List(lbcAirVehicle.ListIndex) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = slStr & lbcAirVehicle.List(lbcAirVehicle.ListIndex)
                ElseIf lbcAirVehicle.ListIndex = 0 Then
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> lbcAirVehicle.List(lbcAirVehicle.ListIndex) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = lbcAirVehicle.List(lbcAirVehicle.ListIndex)
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case XDSPROGCODEINDEX
                edcDropdown.Visible = False  'Set visibility
                slStr = Trim$(edcDropdown.Text)
                If slStr <> "" Then
                    If UCase(grdDates.TextMatrix(lmEnableRow, lmEnableCol)) <> UCase(slStr) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case BUSINDEX
                edcDropdown.Visible = False  'Set visibility
                slStr = Trim$(edcDropdown.Text)
                If slStr <> "" Then
                    If UCase(grdDates.TextMatrix(lmEnableRow, lmEnableCol)) <> UCase(slStr) Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case GAMESTATUSINDEX
                edcDropdown.Visible = False
                cmcDropDown.Visible = False
                lbcStatus.Visible = False
                If lbcStatus.ListIndex >= 0 Then
                    slStr = Left$(lbcStatus.List(lbcStatus.ListIndex), 1)
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        ilRet = vbYes
                        If slStr = "C" Then
                            If gDateValue(grdDates.TextMatrix(lmEnableRow, AIRDATEINDEX)) <= lmNowDate Then
                                ilRet = MsgBox("Event has aired; responding YES will pre-empt all spots in Event.  Continue?", vbYesNo + vbQuestion, "Event Status")
                            End If
                        End If
                        If ilRet = vbYes Then
                            imGsfChg = True
                            grdDates.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                        End If
                    End If
                Else
                    If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        imGsfChg = True
                    End If
                    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
        End Select
    End If
    If imGsfChg Then
        grdDates.TextMatrix(lmEnableRow, CHGFLAGINDEX) = "Y"
    Else
        imGsfChg = ilSvGsfChg
    End If
    ilSvCol = grdDates.Col
    ilSvRow = grdDates.Row
    If lmEnableRow >= grdDates.FixedRows Then
        grdDates.Row = lmEnableRow
        For ilCol = 0 To grdDates.Cols - 1 Step 1
            grdDates.Col = ilCol
            If ilCol = GAMENOINDEX Then
                grdDates.CellBackColor = LIGHTYELLOW
            Else
                If grdDates.ColWidth(ilCol) > 15 Then
                    grdDates.CellBackColor = vbWhite
                End If
            End If
        Next ilCol
        grdDates.Col = ilSvCol
        grdDates.Row = ilSvRow
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
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
    gSetMousePointer grdSpec, grdDates, vbDefault
    igManUnload = YES
    Unload GameSchd
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGridFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mGridFieldsOk(ilRowNo As Integer) As Integer
'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim ilLibraryDefined As Integer
    Dim ilAirDateDefined As Integer
    Dim ilAirTimeDefined As Integer
    Dim ilValue As Integer
    Dim ilError As Integer

    ilError = False
    If Trim$(grdDates.TextMatrix(ilRowNo, GAMENOINDEX)) = "" Then
        mGridFieldsOk = True
        Exit Function
    End If
    slStr = grdDates.TextMatrix(ilRowNo, VERLIBINDEX)
    gFindMatch slStr, 0, lbcLibrary
    If gLastFound(lbcLibrary) >= 0 Then
        ilLibraryDefined = True
    Else
        ilLibraryDefined = False
    End If
    slStr = grdDates.TextMatrix(ilRowNo, AIRDATEINDEX)
    If (slStr <> "") And (slStr <> "Missing") Then
        ilAirDateDefined = True
    Else
        ilAirDateDefined = False
    End If
    slStr = grdDates.TextMatrix(ilRowNo, AIRTIMEINDEX)
    If (slStr <> "") And (slStr <> "Missing") Then
        ilAirTimeDefined = True
    Else
        ilAirTimeDefined = False
    End If
    'This was initially placed here to allow changing from 20 to 10 games and not saving game 11-20.
    'That has been fixed to remove 11-20.  allowing this caused an error when creating LCF as no date defined.
    'If we want this, then add code into mMoveCtrlToRec to bypass either moving in values into gsf or in mCreatelcfssf, bypass blank dates
    'If (Not ilLibraryDefined) And (Not ilAirDateDefined) And (Not ilAirTimeDefined) Then
    '    mGridFieldsOk = True
    '    Exit Function
    'End If
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    If (ilValue And USINGFEED) = USINGFEED Then
        slStr = grdDates.TextMatrix(ilRowNo, FEEDSOURCEINDEX)
        ''If (StrComp(slStr, "Visiting", vbTextCompare) <> 0) And (StrComp(slStr, "Home", vbTextCompare) <> 0) Then
        'If (StrComp(slStr, "V", vbTextCompare) <> 0) And (StrComp(slStr, "H", vbTextCompare) <> 0) And (StrComp(slStr, "N", vbTextCompare) <> 0) Then
        If (StrComp(slStr, smEventTitle1, vbTextCompare) <> 0) And (StrComp(slStr, smEventTitle2, vbTextCompare) <> 0) And (StrComp(slStr, "National", vbTextCompare) <> 0) Then
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdDates.TextMatrix(ilRowNo, FEEDSOURCEINDEX) = "Missing"
                grdDates.Row = ilRowNo
                grdDates.Col = FEEDSOURCEINDEX
                grdDates.CellForeColor = vbMagenta
            Else
                ilError = True
                grdDates.Row = ilRowNo
                grdDates.Col = FEEDSOURCEINDEX
                grdDates.CellForeColor = vbMagenta
            End If
        Else
            grdDates.Row = ilRowNo
            grdDates.Col = FEEDSOURCEINDEX
            grdDates.CellForeColor = vbBlack
        End If
    End If
    'Language
    If (ilValue And USINGLANG) = USINGLANG Then
        slStr = grdDates.TextMatrix(ilRowNo, LANGUAGEINDEX)
        gFindMatch slStr, 1, lbcLanguage
        If gLastFound(lbcLanguage) < 1 Then
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdDates.TextMatrix(ilRowNo, LANGUAGEINDEX) = "Missing"
                grdDates.Row = ilRowNo
                grdDates.Col = LANGUAGEINDEX
                grdDates.CellForeColor = vbMagenta
            Else
                ilError = True
                grdDates.Row = ilRowNo
                grdDates.Col = LANGUAGEINDEX
                grdDates.CellForeColor = vbMagenta
            End If
        Else
            grdDates.Row = ilRowNo
            grdDates.Col = LANGUAGEINDEX
            grdDates.CellForeColor = vbBlack
        End If
    End If
    'Visiting Team
    slStr = grdDates.TextMatrix(ilRowNo, VISITTEAMINDEX)
    gFindMatch slStr, 1, lbcTeam
    If gLastFound(lbcTeam) < 1 Then
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdDates.TextMatrix(ilRowNo, VISITTEAMINDEX) = "Missing"
            grdDates.Row = ilRowNo
            grdDates.Col = VISITTEAMINDEX
            grdDates.CellForeColor = vbMagenta
        Else
            ilError = True
            grdDates.Row = ilRowNo
            grdDates.Col = VISITTEAMINDEX
            grdDates.CellForeColor = vbMagenta
        End If
    Else
        grdDates.Row = ilRowNo
        grdDates.Col = VISITTEAMINDEX
        grdDates.CellForeColor = vbBlack
    End If
    'Home Team
    slStr = grdDates.TextMatrix(ilRowNo, HOMETEAMINDEX)
    gFindMatch slStr, 1, lbcTeam
    If gLastFound(lbcTeam) < 1 Then
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdDates.TextMatrix(ilRowNo, HOMETEAMINDEX) = "Missing"
            grdDates.Row = ilRowNo
            grdDates.Col = HOMETEAMINDEX
            grdDates.CellForeColor = vbMagenta
        Else
            ilError = True
            grdDates.Row = ilRowNo
            grdDates.Col = HOMETEAMINDEX
            grdDates.CellForeColor = vbMagenta
        End If
    Else
        grdDates.Row = ilRowNo
        grdDates.Col = HOMETEAMINDEX
        grdDates.CellForeColor = vbBlack
    End If
    
    'Subtotal 1
    If Trim$(tgSaf(0).sEventSubtotal1) <> "" Then
        slStr = grdDates.TextMatrix(ilRowNo, SUBTOTAL1INDEX)
        gFindMatch slStr, 1, lbcSubtotal(0)
        If gLastFound(lbcSubtotal(0)) < 1 Then
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdDates.TextMatrix(ilRowNo, SUBTOTAL1INDEX) = "Missing"
                grdDates.Row = ilRowNo
                grdDates.Col = SUBTOTAL1INDEX
                grdDates.CellForeColor = vbMagenta
            Else
                ilError = True
                grdDates.Row = ilRowNo
                grdDates.Col = SUBTOTAL1INDEX
                grdDates.CellForeColor = vbMagenta
            End If
        Else
            grdDates.Row = ilRowNo
            grdDates.Col = SUBTOTAL1INDEX
            grdDates.CellForeColor = vbBlack
        End If
    End If
    'Subtotal 2
    If Trim$(tgSaf(0).sEventSubtotal2) <> "" Then
        slStr = grdDates.TextMatrix(ilRowNo, SUBTOTAL2INDEX)
        gFindMatch slStr, 1, lbcSubtotal(1)
        If gLastFound(lbcSubtotal(1)) < 1 Then
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdDates.TextMatrix(ilRowNo, SUBTOTAL2INDEX) = "Missing"
                grdDates.Row = ilRowNo
                grdDates.Col = SUBTOTAL2INDEX
                grdDates.CellForeColor = vbMagenta
            Else
                ilError = True
                grdDates.Row = ilRowNo
                grdDates.Col = SUBTOTAL2INDEX
                grdDates.CellForeColor = vbMagenta
            End If
        Else
            grdDates.Row = ilRowNo
            grdDates.Col = SUBTOTAL2INDEX
            grdDates.CellForeColor = vbBlack
        End If
    End If
    
    If Not ilLibraryDefined Then
        slStr = grdDates.TextMatrix(ilRowNo, VERLIBINDEX)
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdDates.TextMatrix(ilRowNo, LIBRARYINDEX) = "Missing"
            grdDates.Row = ilRowNo
            grdDates.Col = LIBRARYINDEX
            grdDates.CellForeColor = vbMagenta
        Else
            ilError = True
            grdDates.Row = ilRowNo
            grdDates.Col = LIBRARYINDEX
            grdDates.CellForeColor = vbMagenta
        End If
    Else
        grdDates.Row = ilRowNo
        grdDates.Col = LIBRARYINDEX
        grdDates.CellForeColor = vbBlack
    End If
    'Air Date
    slStr = grdDates.TextMatrix(ilRowNo, AIRDATEINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        grdDates.TextMatrix(ilRowNo, AIRDATEINDEX) = "Missing"
        grdDates.Row = ilRowNo
        grdDates.Col = AIRDATEINDEX
        grdDates.CellForeColor = vbMagenta
    Else
        If Not gValidDate(slStr) Then
            ilError = True
            grdDates.Row = ilRowNo
            grdDates.Col = AIRDATEINDEX
            grdDates.CellForeColor = vbMagenta
        Else
            grdDates.Row = ilRowNo
            grdDates.Col = AIRDATEINDEX
            grdDates.CellForeColor = vbBlack
        End If
    End If
    'Air Time
    slStr = grdDates.TextMatrix(ilRowNo, AIRTIMEINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        grdDates.TextMatrix(ilRowNo, AIRTIMEINDEX) = "Missing"
        grdDates.Row = ilRowNo
        grdDates.Col = AIRTIMEINDEX
        grdDates.CellForeColor = vbMagenta
    Else
        If Not gValidTime(slStr) Then
            ilError = True
            grdDates.Row = ilRowNo
            grdDates.Col = AIRTIMEINDEX
            grdDates.CellForeColor = vbMagenta
        Else
            grdDates.Row = ilRowNo
            grdDates.Col = AIRTIMEINDEX
            grdDates.CellForeColor = vbBlack
        End If
    End If
    'Air Vehicle
    If (ilValue And PREEMPTREGPROG) = PREEMPTREGPROG Then
        slStr = grdDates.TextMatrix(ilRowNo, AIRVEHICLEINDEX)
        If Left$(slStr, 2) = "L:" Then
            slStr = Trim$(Mid$(slStr, 3))
        ElseIf Left$(slStr, 2) = "M:" Then
            slStr = Trim$(Mid$(slStr, 3))
        End If
        gFindMatch slStr, 0, lbcAirVehicle
        If gLastFound(lbcAirVehicle) < 0 Then
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                slStr = UCase$(grdDates.TextMatrix(ilRowNo, GAMESTATUSINDEX))
                'If (slStr = "FIRM") Then
                If (slStr = "F") Then
                    ilError = True
                    grdDates.TextMatrix(ilRowNo, AIRVEHICLEINDEX) = "Missing"
                    grdDates.Row = ilRowNo
                    grdDates.Col = AIRVEHICLEINDEX
                    grdDates.CellForeColor = vbMagenta
                End If
            Else
                ilError = True
                grdDates.Row = ilRowNo
                grdDates.Col = AIRVEHICLEINDEX
                grdDates.CellForeColor = vbMagenta
            End If
        Else
            slStr = UCase$(grdDates.TextMatrix(ilRowNo, GAMESTATUSINDEX))
            'If (slStr = "FIRM") And (gLastFound(lbcAirVehicle) = 0) Then
            If (slStr = "F") And (gLastFound(lbcAirVehicle) = 0) Then
                ilError = True
                grdDates.Row = ilRowNo
                grdDates.Col = AIRVEHICLEINDEX
                grdDates.CellForeColor = vbMagenta
            Else
                grdDates.Row = ilRowNo
                grdDates.Col = AIRVEHICLEINDEX
                grdDates.CellForeColor = vbBlack
            End If
        End If
    End If
    If grdDates.ColWidth(XDSPROGCODEINDEX) > 0 Then
        slStr = UCase$(grdDates.TextMatrix(ilRowNo, XDSPROGCODEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdDates.TextMatrix(ilRowNo, XDSPROGCODEINDEX) = "Missing"
            grdDates.Row = ilRowNo
            grdDates.Col = XDSPROGCODEINDEX
            grdDates.CellForeColor = vbMagenta
        Else
            grdDates.Row = ilRowNo
            grdDates.Col = XDSPROGCODEINDEX
            grdDates.CellForeColor = vbBlack
        End If
    End If
    If grdDates.ColWidth(BUSINDEX) > 0 Then
        slStr = UCase$(grdDates.TextMatrix(ilRowNo, BUSINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdDates.TextMatrix(ilRowNo, BUSINDEX) = "Missing"
            grdDates.Row = ilRowNo
            grdDates.Col = BUSINDEX
            grdDates.CellForeColor = vbMagenta
        Else
            grdDates.Row = ilRowNo
            grdDates.Col = BUSINDEX
            grdDates.CellForeColor = vbBlack
        End If
    End If
    slStr = UCase$(grdDates.TextMatrix(ilRowNo, GAMESTATUSINDEX))
    'If (slStr <> "CANCELED") And (slStr <> "FIRM") And (slStr <> "POSTPONED") And (slStr <> "TENTATIVE") Then
    If (slStr <> "C") And (slStr <> "F") And (slStr <> "P") And (slStr <> "T") Then
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdDates.TextMatrix(ilRowNo, GAMESTATUSINDEX) = "Missing"
            grdDates.Row = ilRowNo
            grdDates.Col = GAMESTATUSINDEX
            grdDates.CellForeColor = vbMagenta
        Else
            ilError = True
            grdDates.Row = ilRowNo
            grdDates.Col = GAMESTATUSINDEX
            grdDates.CellForeColor = vbMagenta
        End If
    Else
        grdDates.Row = ilRowNo
        grdDates.Col = GAMESTATUSINDEX
        grdDates.CellForeColor = vbBlack
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                       ilValue                       ilLang                    *
'*  ilTeam                        ilLib                         ilVeh                     *
'*  slNameCode                    slCode                        ilRet                     *
'*  ilPos                                                                                 *
'******************************************************************************************

'
'   mXFerRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llRow As Long
    Dim ilVff As Integer
    
    slStr = Trim$(tmGhf.sSeasonName)
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, SEASONNAMEINDEX) = slStr
    gUnpackDate tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), slStr
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, SEASONSTARTINDEX) = slStr
    gUnpackDate tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), slStr
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, SEASONENDINDEX) = slStr
    smDefault = "No"
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If tgVff(ilVff).iVefCode = imVefCode Then
            If tgVff(ilVff).lSeasonGhfCode = lmSeasonGhfCode Then
                smDefault = "Yes"
                Exit For
            End If
        End If
    Next ilVff
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, DEFAULTINDEX) = smDefault
    slStr = Trim$(str$(tmGhf.iNoGames))
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, NOGAMEINDEX) = slStr
    grdDates.Redraw = False
    llRow = grdDates.FixedRows
    For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
        If llRow + 1 > grdDates.Rows Then
            grdDates.AddItem ""
            grdDates.RowHeight(grdDates.Rows - 1) = fgBoxGridH
            grdDates.AddItem ""
            grdDates.RowHeight(grdDates.Rows - 1) = 15
            mInitNew llRow
        End If
        grdDates.Row = llRow
        grdDates.TextMatrix(llRow, TMGSFINDEX) = Trim$(str$(ilLoop))
        mSetGridValues llRow, tmGsf(ilLoop), False, False
        llRow = llRow + 2
    Next ilLoop
    grdDates.Redraw = True
    Exit Sub
End Sub

Private Sub lbcAirVehicle_Click()
    gProcessLbcClick lbcAirVehicle, edcDropdown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcAirVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcLanguage_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcLanguage, edcDropdown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcLanguage_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True
End Sub

Private Sub lbcLanguage_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcLanguage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcLanguage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcLanguage, edcDropdown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcLibrary_Click()
    gProcessLbcClick lbcLibrary, edcDropdown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcLibrary_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcStatus_Click()
    gProcessLbcClick lbcStatus, edcDropdown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcStatus_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcSubtotal_Click(Index As Integer)
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcSubtotal(Index), edcDropdown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcSubtotal_DblClick(Index As Integer)
    tmcClick.Enabled = False
    imDoubleClickName = True
End Sub

Private Sub lbcSubtotal_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcSubtotal_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcSubtotal_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcSubtotal(Index), edcDropdown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
    End If
End Sub

Private Sub lbcTeam_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcTeam, edcDropdown, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcTeam_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True
End Sub

Private Sub lbcTeam_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcTeam_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcTeam_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcTeam, edcDropdown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcSTab.SetFocus
        Else
            pbcTab.SetFocus
        End If
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
                If imSpecCtrlVisible Then
                    edcSpec.Text = Format$(llDate, "m/d/yy")
                    edcSpec.SelStart = 0
                    edcSpec.SelLength = Len(edcSpec.Text)
                    imBypassFocus = True
                    edcSpec.SetFocus
                Else
                    edcDropdown.Text = Format$(llDate, "m/d/yy")
                    edcDropdown.SelStart = 0
                    edcDropdown.SelLength = Len(edcDropdown.Text)
                    imBypassFocus = True
                    edcDropdown.SetFocus
                End If
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    If imSpecCtrlVisible Then
        edcSpec.SetFocus
    Else
        edcDropdown.SetFocus
    End If
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSpecSetShow
    mSetShow
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcDefault_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        smDefault = "Yes"
        pbcDefault_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        smDefault = "No"
        pbcDefault_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smDefault = "Yes" Then
            smDefault = "No"
            pbcDefault_Paint
        ElseIf smDefault = "No" Then
            smDefault = "Yes"
            pbcDefault_Paint
        End If
    End If
End Sub

Private Sub pbcDefault_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smDefault = "Yes" Then
        smDefault = "No"
        pbcDefault_Paint
    Else
        smDefault = "Yes"
        pbcDefault_Paint
    End If

End Sub

Private Sub pbcDefault_Paint()
    pbcDefault.Cls
    pbcDefault.CurrentX = fgBoxInsetX
    pbcDefault.CurrentY = 0 'fgBoxInsetY
    pbcDefault.Print smDefault
End Sub

Private Sub pbcFeed_KeyPress(KeyAscii As Integer)
    'If (KeyAscii = Asc("H")) Or (KeyAscii = Asc("h")) Then
    If (KeyAscii = Asc(UCase(Left(smEventTitle2, 1)))) Or (KeyAscii = Asc(LCase(Left(smEventTitle2, 1)))) Then
        smFeedSource = smEventTitle2    '"Home"
        pbcFeed_Paint
    'ElseIf KeyAscii = Asc("V") Or (KeyAscii = Asc("v")) Then
    ElseIf (KeyAscii = Asc(UCase(Left(smEventTitle1, 1)))) Or (KeyAscii = Asc(LCase(Left(smEventTitle1, 1)))) Then
        smFeedSource = smEventTitle1    '"Visiting"
        pbcFeed_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        smFeedSource = "National"
        pbcFeed_Paint
    End If
    If KeyAscii = Asc(" ") Then
        'If smFeedSource = "Home" Then
        If smFeedSource = smEventTitle2 Then
            smFeedSource = smEventTitle1    '"Visiting"
            pbcFeed_Paint
        'ElseIf smFeedSource = "Visiting" Then
        ElseIf smFeedSource = smEventTitle1 Then
            smFeedSource = "National"
            pbcFeed_Paint
        ElseIf smFeedSource = "National" Then
            smFeedSource = smEventTitle2    '"Home"
            pbcFeed_Paint
        End If
    End If
End Sub

Private Sub pbcFeed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If smFeedSource = "Home" Then
    If smFeedSource = smEventTitle2 Then
        smFeedSource = smEventTitle1 '"Visiting"
        pbcFeed_Paint
    'ElseIf smFeedSource = "Visiting" Then
    ElseIf smFeedSource = smEventTitle1 Then
        smFeedSource = "National"
        pbcFeed_Paint
    Else
        smFeedSource = smEventTitle2 '"Home"
        pbcFeed_Paint
    End If
End Sub

Private Sub pbcFeed_Paint()
    pbcFeed.Cls
    pbcFeed.CurrentX = fgBoxInsetX
    pbcFeed.CurrentY = 0 'fgBoxInsetY
    pbcFeed.Print smFeedSource
End Sub

Private Sub pbcLiveLogMerge_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("L")) Or (KeyAscii = Asc("l")) Then
        smLiveLogMerge = "Live Log & Pre-empt"
        pbcLiveLogMerge_Paint
    ElseIf KeyAscii = Asc("M") Or (KeyAscii = Asc("m")) Then
        smLiveLogMerge = "Merge & Pre-empt"
        pbcLiveLogMerge_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smLiveLogMerge = "Live Log & Pre-empt" Then
            smLiveLogMerge = "Merge & Pre-empt"
            pbcLiveLogMerge_Paint
        ElseIf smLiveLogMerge = "Merge & Pre-empt" Then
            smLiveLogMerge = "Live Log & Pre-empt"
            pbcLiveLogMerge_Paint
        End If
    End If
End Sub

Private Sub pbcLiveLogMerge_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smLiveLogMerge = "Live Log & Pre-empt" Then
        smLiveLogMerge = "Merge & Pre-empt"
        pbcLiveLogMerge_Paint
    ElseIf smLiveLogMerge = "Merge & Pre-empt" Then
        smLiveLogMerge = "Live Log & Pre-empt"
        pbcLiveLogMerge_Paint
    End If
End Sub

Private Sub pbcLiveLogMerge_Paint()
    pbcLiveLogMerge.Cls
    pbcLiveLogMerge.CurrentX = fgBoxInsetX
    pbcLiveLogMerge.CurrentY = 0 'fgBoxInsetY
    pbcLiveLogMerge.Print smLiveLogMerge
End Sub

Private Sub pbcSpecSTab_GotFocus()
    Dim ilNext As Integer
    Dim blSetShow As Boolean
    Dim slStr As String
    Dim ilTestValue As Integer
    If GetFocus() <> pbcSpecSTab.hWnd Then
        Exit Sub
    End If
    blSetShow = True
    If imSpecCtrlVisible Then
        mSpecSetShow
        cmcDone.SetFocus
        Exit Sub
        
        ilTestValue = True
        Do
            ilNext = False
            Select Case grdSpec.Col
                Case SEASONSTARTINDEX
                    If ilTestValue Then
                        slStr = edcSpec.Text
                        If slStr <> "" Then
                            If Not gValidDate(slStr) Then
                                Beep
                                edcSpec.SetFocus
                                Exit Sub
                            End If
                        Else
                            Beep
                            edcSpec.SetFocus
                            Exit Sub
                        End If
                    End If
                    mSpecSetShow
                    cmcDone.SetFocus
                    Exit Sub
                Case SEASONENDINDEX
                    If ilTestValue Then
                        slStr = edcSpec.Text
                        If slStr <> "" Then
                            If Not gValidDate(slStr) Then
                                Beep
                                edcSpec.SetFocus
                                Exit Sub
                            End If
                        Else
                            Beep
                            edcSpec.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdSpec.Col = grdSpec.Col - 2
                Case Else 'Last control within header
                    grdSpec.Col = grdSpec.Col - 2
            End Select
        Loop While ilNext
        If blSetShow Then
            mSpecSetShow
        End If
    Else
        grdSpec.Row = grdSpec.FixedRows + 1 '+1 to bypass title
        grdSpec.Col = grdSpec.FixedCols
    End If
    mSpecEnableBox
End Sub

Private Sub pbcSpecTab_GotFocus()
    Dim ilNext As Integer
    Dim blSetShow As Boolean
    Dim slStr As String
    Dim ilTestValue As Integer
    
    If GetFocus() <> pbcSpecTab.hWnd Then
        Exit Sub
    End If
    blSetShow = True
    If imSpecCtrlVisible Then
        ilTestValue = True
        Do
            ilNext = False
            Select Case grdSpec.Col
                Case SEASONSTARTINDEX
                    If ilTestValue Then
                        slStr = edcSpec.Text
                        If slStr <> "" Then
                            If Not gValidDate(slStr) Then
                                Beep
                                edcSpec.SetFocus
                                Exit Sub
                            End If
                        Else
                            Beep
                            edcSpec.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdSpec.Col = grdSpec.Col + 2
                Case SEASONENDINDEX
                    If ilTestValue Then
                        slStr = edcSpec.Text
                        If slStr <> "" Then
                            If Not gValidDate(slStr) Then
                                Beep
                                edcSpec.SetFocus
                                Exit Sub
                            End If
                        Else
                            Beep
                            edcSpec.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdSpec.Col = grdSpec.Col + 2
                Case NOGAMEINDEX
                    mSpecSetShow
                    pbcSTab.SetFocus
                    Exit Sub
                Case Else 'Last control within header
                    grdSpec.Col = grdSpec.Col + 2
            End Select
        Loop While ilNext
        If blSetShow Then
            mSpecSetShow
        End If
    Else
        grdSpec.Row = grdSpec.Rows - 1
        grdSpec.Col = grdSpec.FixedCols
    End If
    mSpecEnableBox
End Sub

Private Sub pbcSTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilBox                                                                                 *
'******************************************************************************************

    Dim slStr As String
    Dim ilTestValue As Integer
    Dim ilNext As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        imTabDirection = -1 'Set- Right to left
        If grdDates.Col = LANGUAGEINDEX Then
            If mLangBranch() Then
                Exit Sub
            End If
        End If
        If grdDates.Col = VISITTEAMINDEX Then
            If mTeamBranch() Then
                Exit Sub
            End If
        End If
        If grdDates.Col = HOMETEAMINDEX Then
            If mTeamBranch() Then
                Exit Sub
            End If
        End If
        If grdDates.Col = SUBTOTAL1INDEX Then
            If mSubtotalBranch(0) Then
                Exit Sub
            End If
        End If
        If grdDates.Col = SUBTOTAL2INDEX Then
            If mSubtotalBranch(1) Then
                Exit Sub
            End If
        End If
        ilTestValue = True
        Do
            ilNext = False
            Select Case grdDates.Col
                Case AIRDATEINDEX
                    If ilTestValue Then
                        slStr = edcDropdown.Text
                        If slStr <> "" Then
                            If Not gValidDate(slStr) Then
                                Beep
                                edcDropdown.SetFocus
                                Exit Sub
                            End If
                        Else
                            Beep
                            edcDropdown.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdDates.Col = grdDates.Col - 2
                Case AIRTIMEINDEX
                    If ilTestValue Then
                        slStr = edcDropdown.Text
                        If slStr <> "" Then
                            If Not gValidTime(slStr) Then
                                Beep
                                edcDropdown.SetFocus
                                Exit Sub
                            End If
                        Else
                            Beep
                            edcDropdown.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdDates.Col = grdDates.Col - 2
                Case FEEDSOURCEINDEX
                    If grdDates.Row = grdDates.FixedRows Then
                        mSetShow
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    grdDates.Row = grdDates.Row - 2
                    imSettingValue = True
                    If Not grdDates.RowIsVisible(grdDates.Row) Then
                        grdDates.TopRow = grdDates.TopRow + 1
                    End If
                    imSettingValue = False
                    grdDates.Col = GAMESTATUSINDEX
                Case Else
                    grdDates.Col = grdDates.Col - 2
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
                ilTestValue = False
            End If
        Loop While ilNext
        mSetShow
    Else
        imTabDirection = 0  'Set-Left to right
        lmTopRow = -1
        grdDates.TopRow = grdDates.FixedRows
        grdDates.Row = grdDates.FixedRows
        grdDates.Col = FEEDSOURCEINDEX
        Do
            If mColOk() Then
                Exit Do
            Else
                grdDates.Col = grdDates.Col + 2
            End If
        Loop
    End If
    lmTopRow = grdDates.TopRow
    mEnableBox
End Sub

Private Sub pbcStartNew_GotFocus()
    Dim ilRet As Integer
    If imInNew Then
        Exit Sub
    End If
    If (imVefCode > 0) And (imNewGame) And (imFirstTimeSelect) Then
        imFirstTimeSelect = False
        ilRet = mStartNew()
        If Not ilRet Then
            imTerminate = True
            mTerminate
            Exit Sub
        End If
    End If
    mSetCommands
    ilRet = 0
    On Error GoTo ErrHandle:
    pbcSpecSTab.SetFocus
    If ilRet <> 0 Then
        pbcSpecSTab_GotFocus
    End If
    Exit Sub
ErrHandle:
    ilRet = 1
    Resume Next
End Sub

Private Sub pbcTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilBox                         ilLoop                                                  *
'******************************************************************************************

    Dim slStr As String
    Dim ilNext As Integer
    Dim ilTestValue As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long
    Dim blSetShow As Boolean

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    blSetShow = True
    If imCtrlVisible Then
        'Branch
        imTabDirection = 0 'Set- Left to right
        If grdDates.Col = LANGUAGEINDEX Then
            If mLangBranch() Then
                Exit Sub
            End If
        End If
        If grdDates.Col = VISITTEAMINDEX Then
            If mTeamBranch() Then
                Exit Sub
            End If
        End If
        If grdDates.Col = HOMETEAMINDEX Then
            If mTeamBranch() Then
                Exit Sub
            End If
        End If
        If grdDates.Col = SUBTOTAL1INDEX Then
            If mSubtotalBranch(0) Then
                Exit Sub
            End If
        End If
        If grdDates.Col = SUBTOTAL2INDEX Then
            If mSubtotalBranch(1) Then
                Exit Sub
            End If
        End If
        ilTestValue = True
        Do
            ilNext = False
            Select Case grdDates.Col
                Case AIRDATEINDEX
                    If ilTestValue Then
                        slStr = edcDropdown.Text
                        If slStr <> "" Then
                            If Not gValidDate(slStr) Then
                                Beep
                                edcDropdown.SetFocus
                                Exit Sub
                            End If
                        Else
                            Beep
                            edcDropdown.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdDates.Col = grdDates.Col + 2
                Case AIRTIMEINDEX
                    If ilTestValue Then
                        slStr = edcDropdown.Text
                        If slStr <> "" Then
                            If Not gValidTime(slStr) Then
                                Beep
                                edcDropdown.SetFocus
                                Exit Sub
                            End If
                        Else
                            Beep
                            edcDropdown.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdDates.Col = grdDates.Col + 2
                Case GAMESTATUSINDEX
                    blSetShow = False
                    llEnableRow = lmEnableRow
                    llEnableCol = lmEnableCol
                    mSetShow
                    lmEnableRow = llEnableRow
                    lmEnableCol = llEnableCol
                    If mGridFieldsOk(CInt(lmEnableRow)) = False Then
                        mEnableBox
                        Exit Sub
                    End If
                    If grdDates.Row + 2 > grdDates.Rows - 1 Then
                        'mSetShow
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    grdDates.Row = grdDates.Row + 2
                    slStr = grdDates.TextMatrix(grdDates.Row, GAMENOINDEX)
                    If slStr = "" Then
                        'mSetShow
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    imSettingValue = True
                    'If Not grdDates.RowIsVisible(grdDates.Row) Then
                    Do While grdDates.RowPos(grdDates.Row) + grdDates.RowHeight(grdDates.Row) > grdDates.Height
                        grdDates.TopRow = grdDates.TopRow + 1
                    Loop
                    If grdDates.RowHeight(grdDates.TopRow) <= 15 Then
                        grdDates.TopRow = grdDates.TopRow + 1
                    End If
                    imSettingValue = False
                    grdDates.Col = FEEDSOURCEINDEX
    '                If imRowNo >= UBound(smSave, 2) Then
    '                    pbcArrow.Move grdDates.Left - pbcArrow.Width - 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + (grdDates.RowHeight(grdDates.Row) - pbcArrow.Height) / 2
    '                    pbcArrow.Visible = True
    '                    pbcArrow.SetFocus
    '                    Exit Sub
    '                End If
                Case Else 'Last control within header
                    grdDates.Col = grdDates.Col + 2
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
                ilTestValue = False
            End If
        Loop While ilNext
        If blSetShow Then
            mSetShow
        End If
    Else
        imTabDirection = -1  'Set-Right to left
        imSettingValue = True
        lmTopRow = -1
        grdDates.TopRow = grdDates.FixedRows
        grdDates.Row = grdDates.Rows - 2
        Do
            If Not grdDates.RowIsVisible(grdDates.Row) Then
                grdDates.TopRow = grdDates.TopRow + 1
            Else
                Exit Do
            End If
        Loop
        grdDates.Col = GAMESTATUSINDEX
        imSettingValue = False
        Do
            If mColOk() Then
                Exit Do
            Else
                grdDates.Col = grdDates.Col - 2
            End If
        Loop
    End If
    lmTopRow = grdDates.TopRow
    mEnableBox
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
                    Select Case lmEnableCol
                        Case AIRTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropdown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropdown, slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
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

    ilRet = gPopUserVehicleBox(GameSchd, VEHSPORT + ACTIVEVEH, cbcSelect, tmGameVehicle(), smGameVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", GameSchd
        On Error GoTo 0
    End If
    'Jim:  At this time disallow airing vehicles 10/12/05
    'If used, then this could cause problems with the selling vehicles.  What if the selling avail is linked to two or
    'more airing vehicles and only one of the airing vehicles is carring the game.  Do we preempt the spot or retain it?
    'ilRet = gPopUserVehicleBox(GameSchd, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcAirVehicle, tmAirVehicle(), smAirVehicleTag)
    ilRet = gPopUserVehicleBox(GameSchd, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHEXCLUDESPORT + ACTIVEVEH, lbcAirVehicle, tmAirVehicle(), smAirVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", GameSchd
        On Error GoTo 0
        lbcAirVehicle.AddItem "[None]", 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gPopMnfPlusFieldsBox(GameSchd, lbcTeam, tmTeamCode(), smTeamCodeTag, "Z")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mTeamPopErr
        gCPErrorMsg ilRet, "mTeamPop (gPopMnfPlusFieldsBox)", GameSchd
        On Error GoTo 0
        lbcTeam.AddItem "[New]", 0
    End If
    Exit Sub
mTeamPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mLanguagePop                    *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Language list box     *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mLanguagePop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    ilRet = gPopMnfPlusFieldsBox(GameSchd, lbcLanguage, tmLanguageCode(), smLanguageCodeTag, "L")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLanguagePopErr
        gCPErrorMsg ilRet, "mLanguagePop (gPopMnfPlusFieldsBox)", GameSchd
        On Error GoTo 0
        lbcLanguage.AddItem "[New]", 0
    End If
    Exit Sub
mLanguagePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mLibraryPop                     *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Language list box     *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mLibraryPop()
    Dim slType As String
    Dim ilVer As Integer
    Dim ilRet As Integer

    slType = "R"
    ilVer = ALLLIBFRONT
    ilRet = gPopProgLibBox(GameSchd, ilVer, slType, imVefCode, lbcLibrary, tmLibName(), smLibNameTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLibPopErr
        gCPErrorMsg ilRet, "mLibraryPop (gPopProgLibBox: Library)", GameSchd
        On Error GoTo 0
    End If
    Exit Sub
mLibPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mcbcSelectChange                   *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process vehicle change         *
'*                                                     *
'*******************************************************
Private Sub mcbcSelectChange()
    Dim ilLoopCount As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    If imChgMode = False Then
        imChgMode = True
        ilLoopCount = 0
        Do
            Screen.MousePointer = vbHourglass  'Wait
            gSetMousePointer grdSpec, grdDates, vbHourglass
            If ilLoopCount > 0 Then
                If cbcSelect.ListIndex >= 0 Then
                    cbcSelect.Text = cbcSelect.List(cbcSelect.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            If cbcSelect.Text <> "" Then
                gManLookAhead cbcSelect, imBSMode, imComboBoxIndex
            End If
            imSelectedIndex = cbcSelect.ListIndex
            slNameCode = tmGameVehicle(imSelectedIndex).sKey  'Traffic!lbcVehicle.List(igVehIndexViaPrg)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imVefCode = Val(slCode)
            imVpfIndex = gVpfFind(GameSchd, imVefCode)
            gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLLD
            lmFirstAllowedChgDate = lmNowDate + 1
            mClearCtrlFields
            mLibraryPop
            mSeasonPop
            ilRet = mReadRec(imVefCode, lmSeasonGhfCode)
            If ilRet Then
                mMoveRecToCtrl
                imNewGame = False
            Else
                imNewGame = True
            End If
            imFirstTimeSelect = True
            Screen.MousePointer = vbHourglass  'Wait
            gSetMousePointer grdSpec, grdDates, vbHourglass
        Loop While imSelectedIndex <> cbcSelect.ListIndex
        Screen.MousePointer = vbDefault    'Default
        gSetMousePointer grdSpec, grdDates, vbDefault
        imChgMode = False
        imGsfChg = False
        imGhfChg = False
        mSetCommands
    End If
End Sub
Private Sub mcbcSeasonChange()
    Dim ilLoopCount As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    If imChgMode = False Then
        imChgMode = True
        ilLoopCount = 0
        Do
            Screen.MousePointer = vbHourglass  'Wait
            gSetMousePointer grdSpec, grdDates, vbHourglass
            If ilLoopCount > 0 Then
                If cbcSeason.ListIndex >= 0 Then
                    cbcSeason.Text = cbcSeason.List(cbcSeason.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            If cbcSeason.Text <> "" Then
                gManLookAhead cbcSeason, imBSMode, imSeasonComboBoxIndex
            End If
            imSeasonSelectedIndex = cbcSeason.ListIndex
            lmSeasonGhfCode = cbcSeason.ItemData(imSeasonSelectedIndex)
            mClearCtrlFields
            ilRet = mReadRec(imVefCode, lmSeasonGhfCode)
            If ilRet Then
                mMoveRecToCtrl
                imNewGame = False
            Else
                imNewGame = True
            End If
            Screen.MousePointer = vbHourglass  'Wait
            gSetMousePointer grdSpec, grdDates, vbHourglass
        Loop While imSeasonSelectedIndex <> cbcSeason.ListIndex
        Screen.MousePointer = vbDefault    'Default
        gSetMousePointer grdSpec, grdDates, vbDefault
        imChgMode = False
        imGsfChg = False
        imGhfChg = False
        mSetCommands
    End If
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
Private Function mReadRec(ilVefCode As Integer, llInSeasonGhfCode As Long) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mReadRecErr                                                                           *
'******************************************************************************************

'
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim llSeasonGhfCode As Long
    Dim ilVff As Integer

    ReDim tmGsf(0 To 0) As GSF
    If llInSeasonGhfCode = -1 Then 'Model
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tgVff(ilVff).iVefCode = ilVefCode Then
                llSeasonGhfCode = tgVff(ilVff).lSeasonGhfCode
                Exit For
            End If
        Next ilVff
    Else
        llSeasonGhfCode = llInSeasonGhfCode
    End If
    ilUpper = 0
    tmGhfSrchKey0.lCode = llSeasonGhfCode
    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If (ilRet = BTRV_ERR_NONE) Then
        tmGsfSrchKey1.lghfcode = tmGhf.lCode
        tmGsfSrchKey1.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf(ilUpper), imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.lCode = tmGsf(ilUpper).lghfcode)
            If ilVefCode <> imVefCode Then
                tmGsf(ilUpper).lCode = 0
            End If
            ReDim Preserve tmGsf(0 To UBound(tmGsf) + 1) As GSF
            ilUpper = UBound(tmGsf)
            ilRet = btrGetNext(hmGsf, tmGsf(ilUpper), imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        If ilVefCode <> imVefCode Then
            tmGhf.lCode = 0
        End If
    Else
        tmGhf.lCode = 0
        mReadRec = False
        Exit Function
    End If
    mReadRec = True
    Exit Function
mReadRecErr: 'VBC NR
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function


Private Sub mClearCtrlFields()
    Dim ilRow As Integer
    Dim ilCol As Integer

    smDefault = ""
    ReDim tmGsf(0 To 0) As GSF
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, SEASONNAMEINDEX) = ""
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, SEASONSTARTINDEX) = ""
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, SEASONENDINDEX) = ""
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, DEFAULTINDEX) = ""
    grdSpec.TextMatrix(grdSpec.FixedRows + 1, NOGAMEINDEX) = ""

    If grdDates.Rows > imInitNoRows Then
        For ilRow = grdDates.Rows To imInitNoRows Step -1
            grdDates.RemoveItem (ilRow)
        Next ilRow
    End If
    mGridColumnWidths
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
        grdDates.TextMatrix(ilRow, GAMENOINDEX) = ""
        For ilCol = grdDates.FixedCols To grdDates.Cols - 1 Step 1
            grdDates.TextMatrix(ilRow, ilCol) = ""
        Next ilCol
    Next ilRow
    grdDates.Col = GAMENOINDEX
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        grdDates.Row = ilRow
        grdDates.CellBackColor = LIGHTYELLOW
    Next ilRow
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLangBranch                     *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
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
Private Function mLangBranch() As Integer
'
'   ilRet = mLangBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropdown, lbcLanguage, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mLangBranch = False
        Exit Function
    End If
    sgMnfCallType = "L"
    igMNmCallSource = CALLSOURCEGAME
    If edcDropdown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "GameSchd^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "GameScd^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "EName^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "EName^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'EName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'PEvent.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mLangBranch = True
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
        lbcLanguage.Clear
        smLanguageCodeTag = ""
        mLanguagePop
        If imTerminate Then
            mLangBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcLanguage
        If gLastFound(lbcLanguage) > 0 Then
            imChgMode = True
            lbcLanguage.ListIndex = gLastFound(lbcLanguage)
            edcDropdown.Text = lbcLanguage.List(lbcLanguage.ListIndex)
            imChgMode = False
            mLangBranch = False
        Else
            imChgMode = True
            lbcLanguage.ListIndex = 0
            edcDropdown.Text = lbcLanguage.List(0)
            imChgMode = False
            edcDropdown.SetFocus
            sgMNmName = ""
            Exit Function
        End If
        sgMNmName = ""
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox
        Exit Function
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTeamBranch                  *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
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
Private Function mTeamBranch() As Integer
'
'   ilRet = mTeamBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropdown, lbcTeam, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mTeamBranch = False
        Exit Function
    End If
    sgMnfCallType = "Z"
    igMNmCallSource = CALLSOURCEGAME
    If edcDropdown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    igGameSchdVefCode = imVefCode
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "GameSchd^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "GameScd^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "EName^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "EName^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'EName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'PEvent.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mTeamBranch = True
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
        lbcTeam.Clear
        smTeamCodeTag = ""
        mTeamPop
        If imTerminate Then
            mTeamBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcTeam
        If gLastFound(lbcTeam) > 0 Then
            imChgMode = True
            lbcTeam.ListIndex = gLastFound(lbcTeam)
            edcDropdown.Text = lbcTeam.List(lbcTeam.ListIndex)
            imChgMode = False
            mTeamBranch = False
        Else
            imChgMode = True
            lbcTeam.ListIndex = 0
            edcDropdown.Text = lbcTeam.List(0)
            imChgMode = False
            edcDropdown.SetFocus
            sgMNmName = ""
            Exit Function
        End If
        sgMNmName = ""
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox
        Exit Function
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
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
    Dim ilValue As Integer

    'Update button set if all mandatory fields have data and any field altered
    If imGhfChg Or imGsfChg Then
        cbcSelect.Enabled = False
        cbcSeason.Enabled = False
    Else
        cbcSelect.Enabled = True
        cbcSeason.Enabled = True
    End If
    If (imGhfChg Or imGsfChg) And (UBound(tmGsf) > 0) Then  'At least one event added
        If imUpdateAllowed Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
    End If
    If (grdDates.TextMatrix(grdDates.FixedRows, GAMENOINDEX) <> "") And (imUpdateAllowed) Then
        cmcFormats.Enabled = True
    Else
        cmcFormats.Enabled = False
    End If
    ilValue = Asc(tgSpf.sUsingFeatures)  'Option Fields in Orders/Proposals
    If (ilValue And MULTIMEDIA) = MULTIMEDIA Then 'Using Live Log
        If Not imGhfChg And Not imGsfChg Then
            If (UBound(tmGsf) > 0) Then
                cmcMultimedia.Enabled = True
            Else
                cmcMultimedia.Enabled = False
            End If
        Else
            cmcMultimedia.Enabled = False
        End If
    Else
        cmcMultimedia.Enabled = False
    End If
    If Not imUpdateAllowed Then
        cmcSyncGames.Enabled = False
    Else
        cmcSyncGames.Enabled = True
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mSpecEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    
    If (grdSpec.Row < grdSpec.FixedRows) Or (grdSpec.Row >= grdSpec.Rows) Or (grdSpec.Col < grdSpec.FixedCols) Or (grdSpec.Col >= grdSpec.Cols - 1) Then
        Exit Sub
    End If
    lmSpecEnableRow = grdSpec.Row
    lmSpecEnableCol = grdSpec.Col

    Select Case grdSpec.Row
        Case SPECROW3INDEX
            Select Case grdSpec.Col
                Case SEASONNAMEINDEX
                    edcSpec.MaxLength = 20
                    If grdSpec.Text = "Missing" Then
                        grdSpec.Text = ""
                    End If
                    edcSpec.Text = grdSpec.Text
                Case SEASONSTARTINDEX
                    edcSpec.MaxLength = 10
                    slStr = grdSpec.Text
                    If (slStr = "") Or (slStr = "Missing") Then
                        slStr = gObtainMondayFromToday()
                    End If
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint
                    edcSpec.Text = slStr
                Case SEASONENDINDEX
                    edcSpec.MaxLength = 10
                    slStr = grdSpec.Text
                    If (slStr = "") Or (slStr = "Missing") Then
                        slStr = grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX)
                    End If
                    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                    pbcCalendar_Paint
                    edcSpec.Text = slStr
                Case DEFAULTINDEX
                    smDefault = grdSpec.Text
                    If smDefault = "" Then
                        smDefault = "No"
                    End If
                Case NOGAMEINDEX 'Name
                    '6/30/12:  Allow 5 digit event #'s
                    If grdSpec.Text = "Missing" Then
                        grdSpec.Text = ""
                    End If
                    edcSpec.MaxLength = 5   '3
                    edcSpec.Text = grdSpec.Text
            End Select
    End Select
    mSpecSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSpecSetShow()
    Dim ilNoGames As Integer
    Dim slStr As String
    Dim ilOrigUpper As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilIndex As Integer
    Dim ilCol As Integer
    Dim ilRet As Integer

    pbcArrow.Visible = False
    If (lmSpecEnableRow >= grdSpec.FixedRows) And (lmSpecEnableRow < grdSpec.Rows) Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                    Case SEASONNAMEINDEX
                        edcSpec.Visible = False
                        If edcSpec.Text <> grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) Then
                            imGhfChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = edcSpec.Text
                    Case SEASONSTARTINDEX
                        edcSpec.Visible = False
                        cmcSpec.Visible = False
                        plcCalendar.Visible = False
                        edcSpec.Visible = False
                        If edcSpec.Text <> grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) Then
                            imGhfChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = edcSpec.Text
                    Case SEASONENDINDEX
                        edcSpec.Visible = False
                        cmcSpec.Visible = False
                        plcCalendar.Visible = False
                        edcSpec.Visible = False
                        If edcSpec.Text <> grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) Then
                            imGhfChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = edcSpec.Text
                    Case DEFAULTINDEX
                        pbcDefault.Visible = False
                        If smDefault <> grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) Then
                            imGhfChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smDefault
                    Case NOGAMEINDEX
                        edcSpec.Visible = False
                        imLastColSorted = -1
                        imLastSort = -1
                        grdDates.Col = GAMENOINDEX
                        mSortCol grdDates.Col
                        slStr = Trim$(edcSpec.Text)
                        ilNoGames = Val(slStr)
                        ilOrigUpper = UBound(tmGsf)
                        If ilNoGames < UBound(tmGsf) Then
                            'If any game in past, disallow change
                            For ilLoop = ilNoGames + 1 To ilOrigUpper Step 1
                                llRow = grdDates.FixedRows + 2 * (ilLoop - 1)
                                ilIndex = Val(grdDates.TextMatrix(llRow, TMGSFINDEX))
                                If (tmGsf(ilIndex).lCode > 0) And (Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX)) <> "") Then
                                    If gDateValue(grdDates.TextMatrix(llRow, AIRDATEINDEX)) <= lmNowDate Then
                                        ilRet = MsgBox("Some Events have aired, so mass delete is not allowed.  Cancel Events one by one.", vbOKOnly + vbCritical, "# Events Changed")
                                        lmSpecEnableRow = -1
                                        lmSpecEnableCol = -1
                                        imSpecCtrlVisible = False
                                        mSetCommands
                                        Exit Sub
                                    End If
                                End If
                            Next ilLoop
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = Trim$(edcSpec.Text)
                        slStr = Trim$(edcSpec.Text)
                        ilNoGames = Val(slStr)
                        ilOrigUpper = UBound(tmGsf)
                        If ilNoGames < UBound(tmGsf) Then
                            ilRet = MsgBox("Events from " & Trim$(str(ilNoGames + 1)) & " to " & Trim$(str$(ilOrigUpper)) & " will be Canceled.  Continue with Change?", vbYesNo + vbQuestion, "# Events Changed")
                            If ilRet = vbNo Then
                                grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = Trim$(Trim$(str$(ilOrigUpper)))
                                lmSpecEnableRow = -1
                                lmSpecEnableCol = -1
                                imSpecCtrlVisible = False
                                mSetCommands
                                Exit Sub
                            End If
                            Screen.MousePointer = vbHourglass
                            gSetMousePointer grdSpec, grdDates, vbHourglass
                            For ilLoop = ilNoGames + 1 To ilOrigUpper Step 1
                                llRow = grdDates.FixedRows + 2 * (ilLoop - 1)
                                ilIndex = Val(grdDates.TextMatrix(llRow, TMGSFINDEX))
                                If (tmGsf(ilIndex).lCode > 0) And (Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX)) <> "") Then
                                    grdDates.TextMatrix(llRow, GAMESTATUSINDEX) = "C"
                                    grdDates.TextMatrix(llRow, CHGFLAGINDEX) = "Y"
                                Else
                                    For ilCol = GAMENOINDEX To TMGSFINDEX - 1 Step 1
                                        grdDates.TextMatrix(llRow, ilCol) = ""
                                    Next ilCol
                                End If
                            Next ilLoop
                            For ilLoop = ilNoGames + 1 To ilOrigUpper Step 1
                                llRow = grdDates.FixedRows + 2 * (ilLoop - 1)
                                ilIndex = Val(grdDates.TextMatrix(llRow, TMGSFINDEX))
                                If (tmGsf(ilIndex).lCode <= 0) Or (Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX)) = "") Then
                                    ReDim Preserve tmGsf(0 To ilLoop - 1) As GSF
                                    Exit For
                                End If
                            Next ilLoop
                            Screen.MousePointer = vbDefault
                            gSetMousePointer grdSpec, grdDates, vbDefault
                            imGhfChg = True
                        ElseIf ilNoGames > UBound(tmGsf) Then
                            Screen.MousePointer = vbHourglass
                            gSetMousePointer grdSpec, grdDates, vbHourglass
                            ReDim Preserve tmGsf(0 To ilNoGames) As GSF
                            For ilLoop = ilOrigUpper + 1 To UBound(tmGsf) Step 1
                                llRow = grdDates.FixedRows + 2 * (CLng(ilLoop) - 1)
                                If llRow >= grdDates.Rows Then
                                    grdDates.AddItem ""
                                    grdDates.RowHeight(grdDates.Rows - 1) = fgBoxGridH
                                    grdDates.AddItem ""
                                    grdDates.RowHeight(grdDates.Rows - 1) = 15
                                    'Changed 5/117/06: Wrong index
                                    tmGsf(ilLoop - 1).lCode = 0
                                    mInitNew llRow
                                Else
                                    If Trim$(grdDates.TextMatrix(llRow, GAMESTATUSINDEX)) = "" Then
                                        'Changed 5/117/06: Wrong index
                                        tmGsf(ilLoop - 1).lCode = 0
                                    End If
                                End If
                                grdDates.Col = GAMENOINDEX
                                grdDates.Row = llRow
                                grdDates.CellBackColor = LIGHTYELLOW
                                grdDates.TextMatrix(llRow, GAMENOINDEX) = Trim$(str$(ilLoop))
                                grdDates.TextMatrix(llRow, TMGSFINDEX) = Trim$(str$(ilLoop - 1))
                            Next ilLoop
                            Screen.MousePointer = vbDefault
                            gSetMousePointer grdSpec, grdDates, vbDefault
                            imGhfChg = True
                        End If
                End Select
        End Select
    End If
    lmSpecEnableRow = -1
    lmSpecEnableCol = -1
    imSpecCtrlVisible = False
    mSetCommands
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
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim ilSvCol As Integer

    If (grdDates.Row < grdDates.FixedRows) Or (grdDates.Row >= grdDates.Rows) Or (grdDates.Col < grdDates.FixedCols) Or (grdDates.Col >= grdDates.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdDates.Left - pbcArrow.Width - 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + (grdDates.RowHeight(grdDates.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdDates.Col - 1 Step 1
        llColPos = llColPos + grdDates.ColWidth(ilCol)
    Next ilCol
    ilSvCol = grdDates.Col
    For ilCol = 0 To grdDates.Cols - 1 Step 1
        If grdDates.ColWidth(ilCol) > 15 Then
            grdDates.Col = ilCol
            grdDates.CellBackColor = vbButtonFace
        End If
    Next ilCol
    grdDates.Col = ilSvCol
    Select Case grdDates.Col
        Case FEEDSOURCEINDEX
            pbcFeed.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row)
            pbcFeed_Paint
            pbcFeed.Visible = True
            pbcFeed.SetFocus
        Case LANGUAGEINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 1.5 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcLanguage.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            If lbcLanguage.Top + lbcLanguage.Height > cmcCancel.Top Then
                lbcLanguage.Top = edcDropdown.Top - lbcLanguage.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
        Case VISITTEAMINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 1.5 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcTeam.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height ', edcDropDown.Width + cmcDropDown.Width
            If lbcTeam.Top + lbcTeam.Height > cmcCancel.Top Then
                lbcTeam.Top = edcDropdown.Top - lbcTeam.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
        Case HOMETEAMINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 1.5 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcTeam.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height ', edcDropDown.Width + cmcDropDown.Width
            If lbcTeam.Top + lbcTeam.Height > cmcCancel.Top Then
                lbcTeam.Top = edcDropdown.Top - lbcTeam.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
        Case SUBTOTAL1INDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 1.5 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcSubtotal(0).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height ', edcDropDown.Width + cmcDropDown.Width
            If lbcSubtotal(0).Top + lbcSubtotal(0).Height > cmcCancel.Top Then
                lbcSubtotal(0).Top = edcDropdown.Top - lbcSubtotal(0).Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
        Case SUBTOTAL2INDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 1.5 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcSubtotal(1).Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height ', edcDropDown.Width + cmcDropDown.Width
            If lbcSubtotal(1).Top + lbcSubtotal(1).Height > cmcCancel.Top Then
                lbcSubtotal(1).Top = edcDropdown.Top - lbcSubtotal(1).Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
        Case LIBRARYINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 2 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcLibrary.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height ', edcDropDown.Width + cmcDropDown.Width
            If lbcLibrary.Top + lbcLibrary.Height > cmcCancel.Top Then
                lbcLibrary.Top = edcDropdown.Top - lbcLibrary.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
        Case AIRDATEINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 1.2 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            plcCalendar.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height
            If plcCalendar.Top + plcCalendar.Height > cmcCancel.Top Then
                plcCalendar.Top = edcDropdown.Top - plcCalendar.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            If edcDropdown.Text = "" Then
                plcCalendar.Visible = True
            End If
            edcDropdown.SetFocus
        Case AIRTIMEINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 1.2 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            plcTme.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height
            If plcTme.Top + plcTme.Height > cmcCancel.Top Then
                plcTme.Top = edcDropdown.Top - plcTme.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            If edcDropdown.Text = "" Then
                plcTme.Visible = True
            End If
            edcDropdown.SetFocus
        Case AIRVEHICLEINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            'lbcAirVehicle.Move cmcDropDown.Left + cmcDropDown.Width - lbcAirVehicle.Width, edcDropDown.Top + edcDropDown.Height
            If tgVpf(imVpfIndex).sGenLog <> "A" Then
                lbcAirVehicle.Move cmcDropDown.Left + cmcDropDown.Width - lbcAirVehicle.Width, edcDropdown.Top + edcDropdown.Height
            Else
                pbcLiveLogMerge.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                lbcAirVehicle.Move cmcDropDown.Left + cmcDropDown.Width - lbcAirVehicle.Width, edcDropdown.Top + edcDropdown.Height + pbcLiveLogMerge.Height
            End If
            If lbcAirVehicle.Top + lbcAirVehicle.Height > cmcCancel.Top Then
                lbcAirVehicle.Top = edcDropdown.Top - lbcAirVehicle.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            If tgVpf(imVpfIndex).sGenLog = "A" Then
                pbcLiveLogMerge.Visible = True
            End If
            edcDropdown.SetFocus
        Case XDSPROGCODEINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, 1.2 * grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case BUSINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, grdDates.ColWidth(grdDates.Col) + grdDates.ColWidth(grdDates.Col + 1) - 90, grdDates.RowHeight(grdDates.Row) - 15
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case GAMESTATUSINDEX
            If grdDates.ColWidth(grdDates.Col) > lbcStatus.Width Then
                edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, lbcStatus.Width - cmcDropDown.Width, grdDates.RowHeight(grdDates.Row) - 15
            Else
                edcDropdown.Move grdDates.Left + llColPos + 30 + grdDates.ColWidth(grdDates.Col) - lbcStatus.Width, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, lbcStatus.Width - cmcDropDown.Width, grdDates.RowHeight(grdDates.Row) - 15
            End If
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcStatus.Move cmcDropDown.Left + cmcDropDown.Width - lbcStatus.Width, edcDropdown.Top + edcDropdown.Height
            If lbcStatus.Top + lbcStatus.Height > cmcCancel.Top Then
                lbcStatus.Top = edcDropdown.Top - lbcStatus.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
    End Select
    mSetCommands
End Sub

Private Function mColOk() As Integer
    Dim slStr As String
    Dim ilCol As Integer
    Dim ilIndex As Integer

    mColOk = True
    If grdDates.ColWidth(grdDates.Col) <= 15 Then
        mColOk = False
        Exit Function
    End If
    If grdDates.RowHeight(grdDates.Row) <= 15 Then
        mColOk = False
        Exit Function
    End If
    If grdDates.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
    slStr = Trim$(grdDates.TextMatrix(grdDates.Row, AIRDATEINDEX))
    If (slStr <> "") And (StrComp(slStr, "Missing", vbTextCompare) <> 0) Then
        If (gDateValue(slStr) < lmFirstAllowedChgDate) And (Trim$(grdDates.TextMatrix(grdDates.Row, GAMESTATUSINDEX)) <> "P") Then
            ilCol = grdDates.Col
            If (ilCol <> FEEDSOURCEINDEX) And (ilCol <> LANGUAGEINDEX) And (ilCol <> VISITTEAMINDEX) And (ilCol <> HOMETEAMINDEX) And (ilCol <> SUBTOTAL1INDEX) And (ilCol <> SUBTOTAL2INDEX) And (ilCol <> GAMESTATUSINDEX) Then
                ilIndex = Val(grdDates.TextMatrix(grdDates.Row, TMGSFINDEX))
                If (tmGsf(ilIndex).lCode > 0) And (Trim$(grdDates.TextMatrix(grdDates.Row, GAMENOINDEX)) <> "") Then
                    mColOk = False
                    Exit Function
                End If
            End If
        Else
            If (grdDates.Col = GAMESTATUSINDEX) Then
                If Val(grdDates.TextMatrix(grdDates.Row, GAMENOINDEX)) > Val(grdSpec.TextMatrix(SPECROW3INDEX, NOGAMEINDEX)) Then
                    mColOk = False
                    Exit Function
                End If
            End If
        End If
    End If
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
Private Sub mSpecSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdSpec.Row < grdSpec.FixedRows) Or (grdSpec.Row >= grdSpec.Rows) Or (grdSpec.Col < grdSpec.FixedCols) Or (grdSpec.Col >= grdSpec.Cols - 1) Then
        Exit Sub
    End If
    imSpecCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdSpec.Col - 1 Step 1
        llColPos = llColPos + grdSpec.ColWidth(ilCol)
    Next ilCol
    Select Case grdSpec.Row
        Case SPECROW3INDEX
            Select Case grdSpec.Col
                Case SEASONNAMEINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    edcSpec.Visible = True
                    edcSpec.SetFocus
                Case SEASONSTARTINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 30, 1.2 * grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    cmcSpec.Move edcSpec.Left + edcSpec.Width, edcSpec.Top, cmcSpec.Width, edcSpec.Height
                    plcCalendar.Move edcSpec.Left, edcSpec.Top + edcSpec.Height
                    If plcCalendar.Top + plcCalendar.Height > cmcCancel.Top Then
                        plcCalendar.Top = edcSpec.Top - plcCalendar.Height
                    End If
                    edcSpec.Visible = True
                    cmcSpec.Visible = True
                    If edcSpec.Text = "" Then
                        plcCalendar.Visible = True
                    End If
                    edcSpec.SetFocus
                Case SEASONENDINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 30, 1.2 * grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    cmcSpec.Move edcSpec.Left + edcSpec.Width, edcSpec.Top, cmcSpec.Width, edcSpec.Height
                    plcCalendar.Move edcSpec.Left, edcSpec.Top + edcSpec.Height
                    If plcCalendar.Top + plcCalendar.Height > cmcCancel.Top Then
                        plcCalendar.Top = edcSpec.Top - plcCalendar.Height
                    End If
                    edcSpec.Visible = True
                    cmcSpec.Visible = True
                    If edcSpec.Text = "" Then
                        plcCalendar.Visible = True
                    End If
                    edcSpec.SetFocus
                Case DEFAULTINDEX
                    pbcDefault.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    pbcDefault_Paint
                    pbcDefault.Visible = True
                    pbcDefault.SetFocus
                Case NOGAMEINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    edcSpec.Visible = True
                    edcSpec.SetFocus
            End Select
    End Select
End Sub

Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    'Layout Fixed Rows:0=>Edge; 1=>Blue border; 2=>Column Title 1; 3=Column Title 2; 4=>Blue border
    '       Rows: 5=>input; 6=>blue row line; 7=>Input; 8=>blue row line
    'Layout Fixed Columns: 0=>Edge; 1=Blue border; 2=>Row Title; 3=>Blue border    Note:  This was done this way to allow for horizontal scrolling:  It is not used
    '       Columns: 4=>Input; 5=>Blue column line; 6=>Input; 7=>Blue Column;....
    grdDates.RowHeight(0) = 15
    grdDates.RowHeight(1) = 15
    grdDates.RowHeight(2) = 180
    grdDates.RowHeight(3) = 180
    grdDates.RowHeight(4) = 15
    'On Error Resume Next
    ilRow = grdDates.FixedRows
    Do
        If ilRow + 1 > grdDates.Rows Then
            grdDates.AddItem ""
            grdDates.AddItem ""
        End If
        grdDates.RowHeight(ilRow) = fgBoxGridH
        grdDates.RowHeight(ilRow + 1) = 15
        ilRow = ilRow + 2
    Loop While grdDates.RowIsVisible(ilRow - 1)

    For ilCol = 0 To grdDates.Cols - 1 Step 1
        grdDates.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
    grdDates.ColWidth(0) = 15
    grdDates.ColWidth(1) = 15
    grdDates.ColWidth(3) = 15
    For ilCol = grdDates.FixedCols + 1 To grdDates.Cols - 1 Step 2
        grdDates.ColWidth(ilCol) = 15
    Next ilCol
    'Horizontal Blue Border Lines
    grdDates.Row = 1
    For ilCol = 1 To grdDates.Cols - 1 Step 1
        grdDates.Col = ilCol
        grdDates.CellBackColor = vbBlue
    Next ilCol
    grdDates.Row = 4
    For ilCol = 1 To grdDates.Cols - 1 Step 1
        grdDates.Col = ilCol
        grdDates.CellBackColor = vbBlue
    Next ilCol
    'Horizontal Blue lines
    For ilRow = grdDates.FixedRows + 1 To grdDates.Rows - 1 Step 2
        grdDates.Row = ilRow
        For ilCol = 1 To grdDates.Cols - 1 Step 1
            grdDates.Col = ilCol
            grdDates.CellBackColor = vbBlue
        Next ilCol
    Next ilRow
    'Vertical Border Lines
    grdDates.Col = 1
    For ilRow = 1 To grdDates.Rows - 1 Step 1
        grdDates.Row = ilRow
        grdDates.CellBackColor = vbBlue
    Next ilRow
    grdDates.Col = 3
    For ilRow = 1 To grdDates.Rows - 1 Step 1
        grdDates.Row = ilRow
        grdDates.CellBackColor = vbBlue
    Next ilRow
    'Set color in fix area to white
    grdDates.Col = 2
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        grdDates.Row = ilRow
        grdDates.CellBackColor = vbWhite
    Next ilRow

    'Vertical Blue Lines
    For ilCol = grdDates.FixedCols + 1 To grdDates.Cols - 1 Step 2
        grdDates.Col = ilCol
        For ilRow = 1 To grdDates.Rows - 1 Step 1
            grdDates.Row = ilRow
            grdDates.CellBackColor = vbBlue
        Next ilRow
    Next ilCol
End Sub

Private Sub mGridSpecLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilCol = 0 To grdSpec.Cols - 1 Step 1
        grdSpec.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
    grdSpec.RowHeight(0) = 15
    grdSpec.RowHeight(1) = 15
    grdSpec.RowHeight(2) = 135
    grdSpec.RowHeight(3) = fgBoxGridH
    grdSpec.RowHeight(4) = 15
    grdSpec.ColWidth(0) = 15
    grdSpec.ColWidth(1) = 15
    grdSpec.ColWidth(SEASONNAMEINDEX + 1) = 15
    grdSpec.ColWidth(SEASONSTARTINDEX + 1) = 15
    grdSpec.ColWidth(SEASONENDINDEX + 1) = 15
    grdSpec.ColWidth(DEFAULTINDEX + 1) = 15
    grdSpec.ColWidth(NOGAMEINDEX + 1) = 15
    'Vertical Line
    For ilRow = 1 To 4
        grdSpec.Row = ilRow
        grdSpec.Col = 1
        grdSpec.CellBackColor = vbBlue
    Next ilRow
    For ilRow = 1 To 4
        grdSpec.Row = ilRow
        grdSpec.Col = SEASONNAMEINDEX + 1
        grdSpec.CellBackColor = vbBlue
        grdSpec.Col = SEASONSTARTINDEX + 1
        grdSpec.CellBackColor = vbBlue
        grdSpec.Col = SEASONENDINDEX + 1
        grdSpec.CellBackColor = vbBlue
        grdSpec.Col = DEFAULTINDEX + 1
        grdSpec.CellBackColor = vbBlue
        grdSpec.Col = NOGAMEINDEX + 1
        grdSpec.CellBackColor = vbBlue
    Next ilRow
    'Horizontal
    For ilCol = 1 To NOGAMEINDEX + 1
        grdSpec.Row = 1
        grdSpec.Col = ilCol
        grdSpec.CellBackColor = vbBlue
    Next ilCol
    For ilCol = 1 To NOGAMEINDEX + 1
        grdSpec.Row = 4
        grdSpec.Col = ilCol
        grdSpec.CellBackColor = vbBlue
    Next ilCol
End Sub

Private Sub mGridColumns()
    Dim ilPos As Integer
    Dim slFirstTitle As String
    Dim slSecondTitle As String
    Dim slStr As String
    
    gGetEventTitles imVefCode, smEventTitle1, smEventTitle2
    
    grdDates.Row = 2
    grdDates.Col = GAMENOINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, GAMENOINDEX) = "Event"
    grdDates.Row = 3
    grdDates.Col = GAMENOINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, GAMENOINDEX) = "#"
    'Feed Source
    grdDates.Row = 2
    grdDates.Col = FEEDSOURCEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, FEEDSOURCEINDEX) = "Feed"
    grdDates.Row = 3
    grdDates.Col = FEEDSOURCEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, FEEDSOURCEINDEX) = "Source"
    'Language
    grdDates.Row = 2
    grdDates.Col = LANGUAGEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, LANGUAGEINDEX) = "Language"
    grdDates.Row = 3
    grdDates.Col = LANGUAGEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, LANGUAGEINDEX) = ""
    'Visiting Team
    ilPos = InStr(1, smEventTitle1, " ", vbTextCompare)
    If ilPos > 0 Then
        slFirstTitle = Left(smEventTitle1, ilPos)
        slSecondTitle = Trim$(Mid(smEventTitle1, ilPos + 1))
    Else
        slFirstTitle = smEventTitle1
        slSecondTitle = ""
    End If
    grdDates.Row = 2
    grdDates.Col = VISITTEAMINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, VISITTEAMINDEX) = slFirstTitle   '"Visiting"
    grdDates.Row = 3
    grdDates.Col = VISITTEAMINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, VISITTEAMINDEX) = slSecondTitle  '"Team"
    'Home Team
    ilPos = InStr(1, smEventTitle2, " ", vbTextCompare)
    If ilPos > 0 Then
        slFirstTitle = Left(smEventTitle2, ilPos)
        slSecondTitle = Trim$(Mid(smEventTitle2, ilPos + 1))
    Else
        slFirstTitle = smEventTitle2
        slSecondTitle = ""
    End If
    grdDates.Row = 2
    grdDates.Col = HOMETEAMINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, HOMETEAMINDEX) = slFirstTitle   '"Home"
    grdDates.Row = 3
    grdDates.Col = HOMETEAMINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, HOMETEAMINDEX) = slSecondTitle  '"Team"
    'Subtotal 1
    If Trim$(tgSaf(0).sEventSubtotal1) <> "" Then
        slStr = Trim$(tgSaf(0).sEventSubtotal1)
        ilPos = InStr(1, slStr, " ", vbTextCompare)
        If ilPos > 0 Then
            slFirstTitle = Left(slStr, ilPos)
            slSecondTitle = Trim$(Mid(slStr, ilPos + 1))
        Else
            slFirstTitle = slStr
            slSecondTitle = ""
        End If
        grdDates.Row = 2
        grdDates.Col = SUBTOTAL1INDEX
        grdDates.CellFontBold = False
        grdDates.CellFontName = "Arial"
        grdDates.CellFontSize = 6.75
        grdDates.CellForeColor = vbBlue
        grdDates.CellBackColor = LIGHTBLUE
        grdDates.TextMatrix(2, SUBTOTAL1INDEX) = slFirstTitle   '"Home"
        grdDates.Row = 3
        grdDates.Col = SUBTOTAL1INDEX
        grdDates.CellAlignment = flexAlignLeftCenter
        grdDates.CellFontBold = False
        grdDates.CellFontName = "Arial"
        grdDates.CellFontSize = 6.75
        grdDates.CellForeColor = vbBlue
        grdDates.CellBackColor = LIGHTBLUE
        grdDates.TextMatrix(3, SUBTOTAL1INDEX) = slSecondTitle  '"Team"
    End If
    'Subtotal 2
    If Trim$(tgSaf(0).sEventSubtotal2) <> "" Then
        slStr = Trim$(tgSaf(0).sEventSubtotal2)
        ilPos = InStr(1, slStr, " ", vbTextCompare)
        If ilPos > 0 Then
            slFirstTitle = Left(slStr, ilPos)
            slSecondTitle = Trim$(Mid(slStr, ilPos + 1))
        Else
            slFirstTitle = slStr
            slSecondTitle = ""
        End If
        grdDates.Row = 2
        grdDates.Col = SUBTOTAL2INDEX
        grdDates.CellFontBold = False
        grdDates.CellFontName = "Arial"
        grdDates.CellFontSize = 6.75
        grdDates.CellForeColor = vbBlue
        grdDates.CellBackColor = LIGHTBLUE
        grdDates.TextMatrix(2, SUBTOTAL2INDEX) = slFirstTitle   '"Home"
        grdDates.Row = 3
        grdDates.Col = SUBTOTAL2INDEX
        grdDates.CellAlignment = flexAlignLeftCenter
        grdDates.CellFontBold = False
        grdDates.CellFontName = "Arial"
        grdDates.CellFontSize = 6.75
        grdDates.CellForeColor = vbBlue
        grdDates.CellBackColor = LIGHTBLUE
        grdDates.TextMatrix(3, SUBTOTAL2INDEX) = slSecondTitle  '"Team"
    End If    'Library
    grdDates.Row = 2
    grdDates.Col = LIBRARYINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, LIBRARYINDEX) = "Library"
    grdDates.Row = 3
    grdDates.Col = LIBRARYINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, LIBRARYINDEX) = ""
    'Air Date
    grdDates.Row = 2
    grdDates.Col = AIRDATEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, AIRDATEINDEX) = "Air"
    grdDates.Row = 3
    grdDates.Col = AIRDATEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, AIRDATEINDEX) = "Date"
    'Air Time
    grdDates.Row = 2
    grdDates.Col = AIRTIMEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, AIRTIMEINDEX) = "Air"
    grdDates.Row = 3
    grdDates.Col = AIRTIMEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, AIRTIMEINDEX) = "Time"
    'Air Vehicle
    grdDates.Row = 2
    grdDates.Col = AIRVEHICLEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, AIRVEHICLEINDEX) = "Pre-empt"
    grdDates.Row = 3
    grdDates.Col = AIRVEHICLEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, AIRVEHICLEINDEX) = "Vehicle"
    'XDS Program Code ID
    grdDates.Row = 2
    grdDates.Col = XDSPROGCODEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, XDSPROGCODEINDEX) = "XDS Prog"
    grdDates.Row = 3
    grdDates.Col = XDSPROGCODEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, XDSPROGCODEINDEX) = "Code ID"
    'Engr Bus
    grdDates.Row = 2
    grdDates.Col = BUSINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, BUSINDEX) = "Bus"
    grdDates.Row = 3
    grdDates.Col = BUSINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, BUSINDEX) = ""
    'Status
    grdDates.Row = 2
    grdDates.Col = GAMESTATUSINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, GAMESTATUSINDEX) = "Event"
    grdDates.Row = 3
    grdDates.Col = GAMESTATUSINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, GAMESTATUSINDEX) = "Status"
End Sub

Private Sub mGridColumnWidths()
    Dim ilValue As Integer
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer
    Dim ilVff As Integer
    Dim ilInterfaceID As Integer
    Dim ilVpfIndex As Integer

    grdDates.ColWidth(TMGSFINDEX) = 0
    grdDates.ColWidth(CHGFLAGINDEX) = 0
    grdDates.ColWidth(SORTINDEX) = 0
    grdDates.ColWidth(VERLIBINDEX) = 0

    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    'Game number
    '6/30/12:  Allow 5 digit event #'s
    grdDates.ColWidth(GAMENOINDEX) = 0.057 * grdDates.Width
    'Feed Source
    If (ilValue And USINGFEED) = USINGFEED Then
        grdDates.ColWidth(FEEDSOURCEINDEX) = 0.037 * grdDates.Width
    Else
        grdDates.ColWidth(FEEDSOURCEINDEX) = 0
        grdDates.ColWidth(FEEDSOURCEINDEX + 1) = 0
    End If
    'Language
    If (ilValue And USINGLANG) = USINGLANG Then
        grdDates.ColWidth(LANGUAGEINDEX) = 0.07 * grdDates.Width
    Else
        grdDates.ColWidth(LANGUAGEINDEX) = 0
        grdDates.ColWidth(LANGUAGEINDEX + 1) = 0
    End If
    grdDates.ColWidth(XDSPROGCODEINDEX) = 0
    ilInterfaceID = 0
    ilVpfIndex = gVpfFind(GameSchd, imVefCode)
    If ilVpfIndex > 0 Then
        ilInterfaceID = tgVpf(ilVpfIndex).iInterfaceID
    End If
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If tgVff(ilVff).iVefCode = imVefCode Then
            '4/16/14: Add By ISCI as option
            'If UCase(Trim(tgVff(ilVff).sXDProgCodeID)) = "EVENT" Then
            If (UCase(Trim(tgVff(ilVff).sXDProgCodeID)) = "EVENT") Or (ilInterfaceID > 0) Then
                grdDates.ColWidth(XDSPROGCODEINDEX) = 0.057 * grdDates.Width
                grdDates.ColWidth(XDSPROGCODEINDEX + 1) = 15
                Exit For
            End If
        End If
    Next ilVff
    If grdDates.ColWidth(XDSPROGCODEINDEX) = 0 Then
        grdDates.ColWidth(XDSPROGCODEINDEX + 1) = 0
    End If
    'Bus
    ilVpfIndex = gVpfFind(GameSchd, imVefCode)
    grdDates.ColWidth(BUSINDEX) = 0
    If ilVpfIndex > 0 Then
        If Trim$(tgVpf(ilVpfIndex).sExpHiNYISCI) = "Y" Then
            grdDates.ColWidth(BUSINDEX) = 0.04 * grdDates.Width
            grdDates.ColWidth(BUSINDEX + 1) = 15
        End If
    End If
    If grdDates.ColWidth(BUSINDEX) = 0 Then
        grdDates.ColWidth(BUSINDEX + 1) = 0
    End If
    'Visiting Team
    grdDates.ColWidth(VISITTEAMINDEX) = 0.11 * grdDates.Width
    'Home Team
    grdDates.ColWidth(HOMETEAMINDEX) = 0.11 * grdDates.Width
    If Trim$(tgSaf(0).sEventSubtotal1) <> "" Then
        grdDates.ColWidth(SUBTOTAL1INDEX) = 0.07 * grdDates.Width
    Else
        grdDates.ColWidth(SUBTOTAL1INDEX) = 0
        grdDates.ColWidth(SUBTOTAL1INDEX + 1) = 0
    End If
    If Trim$(tgSaf(0).sEventSubtotal2) <> "" Then
        grdDates.ColWidth(SUBTOTAL2INDEX) = 0.07 * grdDates.Width
    Else
        grdDates.ColWidth(SUBTOTAL2INDEX) = 0
        grdDates.ColWidth(SUBTOTAL2INDEX + 1) = 0
    End If
    'Library
    grdDates.ColWidth(LIBRARYINDEX) = 0.11 * grdDates.Width
    'Air Date
    grdDates.ColWidth(AIRDATEINDEX) = 0.08 * grdDates.Width
    'Air Time
    grdDates.ColWidth(AIRTIMEINDEX) = 0.08 * grdDates.Width
    'Air Vehicle
    If (ilValue And PREEMPTREGPROG) = PREEMPTREGPROG Then
        If grdDates.ColWidth(XDSPROGCODEINDEX) = 0 Then
            grdDates.ColWidth(AIRVEHICLEINDEX) = 0.15 * grdDates.Width
        Else
            grdDates.ColWidth(AIRVEHICLEINDEX) = 0.12 * grdDates.Width
        End If
    Else
        grdDates.ColWidth(AIRVEHICLEINDEX) = 0
        grdDates.ColWidth(AIRVEHICLEINDEX + 1) = 0
    End If
    'Status
    grdDates.ColWidth(GAMESTATUSINDEX) = 0.06 * grdDates.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdDates.Width
    For ilCol = 0 To grdDates.Cols - 1 Step 1
        llWidth = llWidth + grdDates.ColWidth(ilCol)
        If (grdDates.ColWidth(ilCol) > 15) And (grdDates.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdDates.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdDates.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdDates.Width
            For ilCol = 0 To grdDates.Cols - 1 Step 1
                If (grdDates.ColWidth(ilCol) > 15) And (grdDates.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdDates.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdDates.FixedCols To grdDates.Cols - 1 Step 1
                If grdDates.ColWidth(ilCol) > 15 Then
                    ilColInc = grdDates.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdDates.ColWidth(ilCol) = grdDates.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
    mGridColumns
End Sub

Private Sub mGridSpecColumns()
    grdSpec.Row = 2
    grdSpec.Col = 2
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(2, 2) = "Number of Events"

    grdSpec.Row = 2
    grdSpec.Col = SEASONNAMEINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(2, SEASONNAMEINDEX) = "Season Name"
    grdSpec.Row = 2
    grdSpec.Col = SEASONSTARTINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(2, SEASONSTARTINDEX) = "Start Date"
    grdSpec.Row = 2
    grdSpec.Col = SEASONENDINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(2, SEASONENDINDEX) = "End Date"
    
    grdSpec.Row = 2
    grdSpec.Col = DEFAULTINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(2, DEFAULTINDEX) = "Default"
   
    grdSpec.Row = 2
    grdSpec.Col = NOGAMEINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(2, NOGAMEINDEX) = "# of Events"

End Sub

Private Sub mGridSpecColumnWidths()
    grdSpec.ColWidth(NOGAMEINDEX) = 0.15 * grdSpec.Width
    grdSpec.ColWidth(SEASONSTARTINDEX) = 0.15 * grdSpec.Width
    grdSpec.ColWidth(SEASONENDINDEX) = 0.15 * grdSpec.Width
    grdSpec.ColWidth(DEFAULTINDEX) = 0.15 * grdSpec.Width
    grdSpec.ColWidth(SEASONNAMEINDEX) = grdSpec.Width - grdSpec.ColWidth(0) - grdSpec.ColWidth(1) - grdSpec.ColWidth(NOGAMEINDEX) - grdSpec.ColWidth(SEASONSTARTINDEX) - grdSpec.ColWidth(SEASONENDINDEX) - grdSpec.ColWidth(DEFAULTINDEX) - grdSpec.ColWidth(11) - 125   '75
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case lmEnableCol
        Case LANGUAGEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcLanguage, edcDropdown, imChgMode, imLbcArrowSetting
        Case VISITTEAMINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcTeam, edcDropdown, imChgMode, imLbcArrowSetting
        Case HOMETEAMINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcTeam, edcDropdown, imChgMode, imLbcArrowSetting
        Case SUBTOTAL1INDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSubtotal(0), edcDropdown, imChgMode, imLbcArrowSetting
        Case SUBTOTAL2INDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcSubtotal(1), edcDropdown, imChgMode, imLbcArrowSetting
    End Select
End Sub

Private Function mSaveRec() As Integer
    Dim ilRow As Integer
    Dim slMsg As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilError As Integer
    Dim llOrigDate As Long
    Dim llNewDate As Long
    Dim llDate As Long
    Dim llOrigTime As Long
    Dim llNewTime As Long
    Dim llTime As Long
    Dim llOrigLength As Long
    Dim llNewLength As Long
    Dim llLength As Long
    Dim ilVff As Integer
    Dim ilDate As Integer
    Dim llOrigSeasonStartDate As Long
    Dim llOrigSeasonEndDate As Long
    ReDim llPreemptChkDate(0 To 0) As Long
    Dim tlGhf As GHF
    Dim tlGsf As GSF

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSpec, grdDates, vbHourglass
    ilError = Not mSpecGridFieldsOk()
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If mGridFieldsOk(ilRow) = False Then
            ilError = True
        End If
    Next ilRow
    If ilError Then
        Screen.MousePointer = vbDefault
        gSetMousePointer grdSpec, grdDates, vbDefault
        Beep
        mSaveRec = False
        Exit Function
    End If
    If Not mCheckDates() Then
        Screen.MousePointer = vbDefault
        gSetMousePointer grdSpec, grdDates, vbDefault
        Beep
        mSaveRec = False
        Exit Function
    End If
    If Not mCheckOverlap() Then
        Screen.MousePointer = vbDefault
        gSetMousePointer grdSpec, grdDates, vbDefault
        Beep
        mSaveRec = False
        Exit Function
    End If
    mMoveCtrlToRec
    '3/9/10: Check that Game number not previously added
    For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
        If (Not imNewGame) And (tmGsf(ilLoop).lCode <= 0) Then
            tmGsfSrchKey1.lghfcode = tmGhf.lCode
            tmGsfSrchKey1.iGameNo = tmGsf(ilLoop).iGameNo
            ilRet = btrGetEqual(hmGsf, tlGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = MsgBox("Event Number " & tmGsf(ilLoop).iGameNo & " previously added, Save Stopped to avoid duplicate Event Numbers", vbOKOnly + vbExclamation, "Erase")
                Screen.MousePointer = vbDefault
                gSetMousePointer grdSpec, grdDates, vbDefault
                Beep
                mSaveRec = False
                Exit Function
            End If
        End If
    Next ilLoop
    '1/26/12: Removed Trans to see if this eliminates the Deadlock issues and some of the slowness
    'ilRet = btrBeginTrans(hmGhf, 1000)
    llOrigSeasonStartDate = 0
    llOrigSeasonEndDate = 0
    If imNewGame Then
        tmGhf.lCode = 0
        tmGhf.iVefCode = imVefCode
        tmGhf.sSeasonName = grdSpec.TextMatrix(SPECROW3INDEX, SEASONNAMEINDEX)
        gPackDate grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX), tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1)
        gPackDate grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX), tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1)
        tmGhf.iNoGames = Val(grdSpec.TextMatrix(SPECROW3INDEX, NOGAMEINDEX))
        tmGhf.sUnused = ""
        ilRet = btrInsert(hmGhf, tmGhf, imGhfRecLen, INDEXKEY0)
        slMsg = "mSaveRec (btrInsert:Event Header)"
    Else
        Do
            tmGhfSrchKey0.lCode = tmGhf.lCode
            ilRet = btrGetEqual(hmGhf, tlGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            gUnpackDateLong tlGhf.iSeasonStartDate(0), tlGhf.iSeasonStartDate(1), llOrigSeasonStartDate
            gUnpackDateLong tlGhf.iSeasonEndDate(0), tlGhf.iSeasonEndDate(1), llOrigSeasonEndDate
            tmGhf.sSeasonName = grdSpec.TextMatrix(SPECROW3INDEX, SEASONNAMEINDEX)
            gPackDate grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX), tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1)
            gPackDate grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX), tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1)
            tmGhf.iNoGames = Val(grdSpec.TextMatrix(SPECROW3INDEX, NOGAMEINDEX))
            ilRet = btrUpdate(hmGhf, tmGhf, imGhfRecLen)
            slMsg = "mSaveRec (btrUpdate:Event Header)"
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, GameSchd
    On Error GoTo 0
    'Unlock SSF and adjust Lock records (alf) prior to handling new and changed definitions
    'This is required to aviod unlocking areas that should have been locked.
    'i.e. Game running on vehicle A at 2p-4p.  it is moved to Vehicle B and another game is placed onto vehicle A running 1p-3p.
    '     If the game that is running at 1p-3p was processed first, then 2p-3p would end up unlocked.
    For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
        If (Not imNewGame) And (tmGsf(ilLoop).lCode > 0) And (bmChgFlag(ilLoop)) Then
            tmGsfSrchKey0.lCode = tmGsf(ilLoop).lCode
            ilRet = btrGetEqual(hmGsf, tlGsf, imGsfRecLen, tmGsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tlGsf.iAirDate(0), tlGsf.iAirDate(1), llOrigDate
                gUnpackDateLong tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), llNewDate
                gUnpackTimeLong tlGsf.iAirTime(0), tlGsf.iAirTime(1), False, llOrigTime
                gUnpackTimeLong tmGsf(ilLoop).iAirTime(0), tmGsf(ilLoop).iAirTime(1), False, llNewTime
                tmLvfSrchKey.lCode = tlGsf.lLvfCode
                ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    gSetMousePointer grdSpec, grdDates, vbDefault
                    'ilRet = btrAbortTrans(hmGhf)
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, llOrigLength
                tmLvfSrchKey.lCode = tmGsf(ilLoop).lLvfCode
                ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    gSetMousePointer grdSpec, grdDates, vbDefault
                    'ilRet = btrAbortTrans(hmGhf)
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                If Not mAddEcfRecord(tmGsf(ilLoop).lCode, tmGsf(ilLoop).sGameStatus, llOrigDate, llNewDate) Then
                    Screen.MousePointer = vbDefault
                    gSetMousePointer grdSpec, grdDates, vbDefault
                    'ilRet = btrAbortTrans(hmGhf)
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, llNewLength
                If ((llOrigDate <> llNewDate) Or (llOrigTime <> llNewTime) Or (llOrigLength <> llNewLength) Or (tlGsf.iAirVefCode <> tmGsf(ilLoop).iAirVefCode) Or (tlGsf.sGameStatus <> tmGsf(ilLoop).sGameStatus)) And ((tlGsf.sGameStatus = "F") Or (tmGsf(ilLoop).sGameStatus = "F")) Then
                    ilRet = mPreemptAndLock(0, tlGsf.iAirVefCode, tmGsf(ilLoop).iAirVefCode, tlGsf.sGameStatus, tmGsf(ilLoop).sGameStatus, llOrigDate, llNewDate, llOrigTime, llNewTime, llOrigLength, llNewLength)
                    If ilRet = 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        mSaveRec = False
                        Exit Function
                    ElseIf ilRet > 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    If (tlGsf.sGameStatus = "F") And (tmGsf(ilLoop).sGameStatus <> "F") Then
                        llPreemptChkDate(UBound(llPreemptChkDate)) = llOrigDate
                        ReDim Preserve llPreemptChkDate(0 To UBound(llPreemptChkDate) + 1) As Long
                    End If
                End If
           Else
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, GameSchd
                On Error GoTo 0
            End If
        End If
    Next ilLoop
    'How handle new and changed definitions
    For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
        If (imNewGame) Or (tmGsf(ilLoop).lCode <= 0) Then
            '3/9/10:  Verify that game not added.  Two new games with the same number
            tmGsfSrchKey1.lghfcode = tmGhf.lCode
            tmGsfSrchKey1.iGameNo = tmGsf(ilLoop).iGameNo
            ilRet = btrGetEqual(hmGsf, tlGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = MsgBox("Event Number " & tmGsf(ilLoop).iGameNo & " previously added, Save Stopped to avoid duplicate Event Numbers", vbOKOnly + vbExclamation, "Erase")
                Screen.MousePointer = vbDefault
                gSetMousePointer grdSpec, grdDates, vbDefault
                'ilRet = btrAbortTrans(hmGhf)
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
            tmGsf(ilLoop).lCode = 0
            tmGsf(ilLoop).iVefCode = imVefCode
            tmGsf(ilLoop).lghfcode = tmGhf.lCode
            ilRet = btrInsert(hmGsf, tmGsf(ilLoop), imGsfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert:Event Schedule)"
            If ilRet = BTRV_ERR_NONE Then
                ilRet = mCreateLcfSsf(tmGsf(ilLoop), 0)
                If Not ilRet Then
                    Screen.MousePointer = vbDefault
                    gSetMousePointer grdSpec, grdDates, vbDefault
                    'ilRet = btrAbortTrans(hmGhf)
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                If tmGsf(ilLoop).sGameStatus = "F" Then
                    tmLvfSrchKey.lCode = tmGsf(ilLoop).lLvfCode
                    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                    If ilRet <> BTRV_ERR_NONE Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, llNewLength
                    gUnpackDateLong tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), llNewDate
                    gUnpackTimeLong tmGsf(ilLoop).iAirTime(0), tmGsf(ilLoop).iAirTime(1), False, llNewTime
                    ilRet = mPreemptAndLock(1, 0, tmGsf(ilLoop).iAirVefCode, "", tmGsf(ilLoop).sGameStatus, 0, llNewDate, 0, llNewTime, 0, llNewLength)
                    If ilRet = 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        mSaveRec = False
                        Exit Function
                    ElseIf ilRet > 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                End If
            Else
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, GameSchd
                On Error GoTo 0
            End If
        ElseIf (bmChgFlag(ilLoop)) Then
            Do
                tmGsfSrchKey0.lCode = tmGsf(ilLoop).lCode
                ilRet = btrGetEqual(hmGsf, tlGsf, imGsfRecLen, tmGsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                ilRet = btrUpdate(hmGsf, tmGsf(ilLoop), imGsfRecLen)
                slMsg = "mSaveRec (btrUpdate:Event Schedule)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tlGsf.iAirDate(0), tlGsf.iAirDate(1), llOrigDate
                gUnpackDateLong tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), llNewDate
                gUnpackTimeLong tlGsf.iAirTime(0), tlGsf.iAirTime(1), False, llOrigTime
                gUnpackTimeLong tmGsf(ilLoop).iAirTime(0), tmGsf(ilLoop).iAirTime(1), False, llNewTime
                If (llOrigDate <> llNewDate) Or (llOrigTime <> llNewTime) Or (tmGsf(ilLoop).lLvfCode <> tlGsf.lLvfCode) Then
                    ilRet = mUpdateLcfSsf(tmGsf(ilLoop), llOrigDate, llOrigSeasonStartDate, llOrigSeasonEndDate)
                    If Not ilRet Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                End If
                tmLvfSrchKey.lCode = tlGsf.lLvfCode
                ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    gSetMousePointer grdSpec, grdDates, vbDefault
                    'ilRet = btrAbortTrans(hmGhf)
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, llOrigLength
                tmLvfSrchKey.lCode = tmGsf(ilLoop).lLvfCode
                ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    gSetMousePointer grdSpec, grdDates, vbDefault
                    'ilRet = btrAbortTrans(hmGhf)
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, llNewLength
                If ((llOrigDate <> llNewDate) Or (llOrigTime <> llNewTime) Or (llOrigLength <> llNewLength) Or (tlGsf.iAirVefCode <> tmGsf(ilLoop).iAirVefCode) Or (tlGsf.sGameStatus <> tmGsf(ilLoop).sGameStatus)) And ((tlGsf.sGameStatus = "F") Or (tmGsf(ilLoop).sGameStatus = "F")) Then
                    ilRet = mPreemptAndLock(1, tlGsf.iAirVefCode, tmGsf(ilLoop).iAirVefCode, tlGsf.sGameStatus, tmGsf(ilLoop).sGameStatus, llOrigDate, llNewDate, llOrigTime, llNewTime, llOrigLength, llNewLength)
                    If ilRet = 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        mSaveRec = False
                        Exit Function
                    ElseIf ilRet > 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                End If
                If tmGsf(ilLoop).lLvfCode <> tlGsf.lLvfCode Then
                    mReschGameMissed tmGsf(ilLoop)
                End If
                 If (tlGsf.sGameStatus <> tmGsf(ilLoop).sGameStatus) And (tmGsf(ilLoop).sGameStatus = "C") Then
                    ilRet = mPreemptGame(tmGsf(ilLoop).iVefCode, tmGsf(ilLoop).iGameNo, llNewDate)
                    If ilRet = 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrEndTrans(hmGhf)
                        mSaveRec = False
                        Exit Function
                    ElseIf ilRet > 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                End If
                ilRet = mAdjustMMNtr(tlGsf.iGameNo, llOrigDate, llNewDate)
                If Not ilRet Then
                    Screen.MousePointer = vbDefault
                    gSetMousePointer grdSpec, grdDates, vbDefault
                    'ilRet = btrAbortTrans(hmGhf)
                    imTerminate = True
                    mSaveRec = False
                    Exit Function
                End If
                '11/23/12:  If date changed and LST created, remove lst and ast records
                
           Else
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, GameSchd
                On Error GoTo 0
            End If
        End If
    Next ilLoop
    For ilDate = 0 To UBound(llPreemptChkDate) - 1 Step 1
        For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
            If (tmGsf(ilLoop).sGameStatus = "F") And (bmChgFlag(ilLoop)) Then
                gUnpackDateLong tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), llDate
                If llPreemptChkDate(ilDate) = llDate Then
                    gUnpackTimeLong tmGsf(ilLoop).iAirTime(0), tmGsf(ilLoop).iAirTime(1), False, llTime
                    tmLvfSrchKey.lCode = tmGsf(ilLoop).lLvfCode
                    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                    If ilRet <> BTRV_ERR_NONE Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                    gUnpackTimeLong tmLvf.iLen(0), tmLvf.iLen(1), False, llLength
                    ilRet = mPreemptAndLock(1, 0, tmGsf(ilLoop).iAirVefCode, "", tmGsf(ilLoop).sGameStatus, 0, llDate, 0, llTime, 0, llLength)
                    If ilRet = 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        mSaveRec = False
                        Exit Function
                    ElseIf ilRet > 1 Then
                        Screen.MousePointer = vbDefault
                        gSetMousePointer grdSpec, grdDates, vbDefault
                        'ilRet = btrAbortTrans(hmGhf)
                        imTerminate = True
                        mSaveRec = False
                        Exit Function
                    End If
                End If
            End If
        Next ilLoop
    Next ilDate
    If smDefault = "Yes" Then
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tgVff(ilVff).iVefCode = imVefCode Then
                If tgVff(ilVff).lSeasonGhfCode <> tmGhf.lCode Then
                    tmVffSrchKey0.iCode = tgVff(ilVff).iCode
                    ilRet = btrGetEqual(hmVff, tgVff(ilVff), imVffRecLen, tmVffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If (ilRet = BTRV_ERR_NONE) Then
                        tgVff(ilVff).lSeasonGhfCode = tmGhf.lCode
                        ilRet = btrUpdate(hmVff, tgVff(ilVff), imVffRecLen)
                    End If
                End If
            End If
        Next ilVff
    Else
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tgVff(ilVff).iVefCode = imVefCode Then
                If tgVff(ilVff).lSeasonGhfCode = tmGhf.lCode Then
                    tmVffSrchKey0.iCode = tgVff(ilVff).iCode
                    ilRet = btrGetEqual(hmVff, tgVff(ilVff), imVffRecLen, tmVffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If (ilRet = BTRV_ERR_NONE) Then
                        tgVff(ilVff).lSeasonGhfCode = 0
                        ilRet = btrUpdate(hmVff, tgVff(ilVff), imVffRecLen)
                    End If
                End If
            End If
        Next ilVff
    End If
    'ilRet = btrEndTrans(hmGhf)
    imNewGame = False
    imGhfChg = False
    imGsfChg = False
    mSaveRec = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdDates, vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    'ilRet = btrAbortTrans(hmGhf)
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdDates, vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer control values to     *
'*                      records                        *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
    Dim ilRow As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilLoop As Integer

    ReDim bmChgFlag(0 To UBound(tmGsf)) As Boolean
    For ilLoop = 0 To UBound(bmChgFlag) Step 1
        bmChgFlag(ilLoop) = False
    Next ilLoop
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If (grdDates.TextMatrix(ilRow, GAMENOINDEX) <> "") Then
            ilIndex = Val(grdDates.TextMatrix(ilRow, TMGSFINDEX))
            If grdDates.TextMatrix(ilRow, CHGFLAGINDEX) = "Y" Then
                bmChgFlag(ilIndex) = True
            End If
            tmGsf(ilIndex).iGameNo = Val(grdDates.TextMatrix(ilRow, GAMENOINDEX))
            ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
            'Feed
            tmGsf(ilIndex).sFeedSource = ""
            If (ilValue And USINGFEED) = USINGFEED Then
                slStr = grdDates.TextMatrix(ilRow, FEEDSOURCEINDEX)
                ''If StrComp(slStr, "Visiting", vbTextCompare) = 0 Then
                ''    tmGsf(ilIndex).sFeedSource = "V"
                ''ElseIf StrComp(slStr, "Home", vbTextCompare) = 0 Then
                ''    tmGsf(ilIndex).sFeedSource = "H"
                ''End If
                'If StrComp(slStr, "V", vbTextCompare) = 0 Then
                If StrComp(slStr, smEventTitle1, vbTextCompare) = 0 Then
                    tmGsf(ilIndex).sFeedSource = "V"
                'ElseIf StrComp(slStr, "H", vbTextCompare) = 0 Then
                ElseIf StrComp(slStr, smEventTitle2, vbTextCompare) = 0 Then
                    tmGsf(ilIndex).sFeedSource = "H"
                ElseIf StrComp(slStr, "National", vbTextCompare) = 0 Then
                    tmGsf(ilIndex).sFeedSource = "N"
                End If
            End If
            'Language
            tmGsf(ilIndex).iLangMnfCode = 0
            If (ilValue And USINGLANG) = USINGLANG Then
                slStr = grdDates.TextMatrix(ilRow, LANGUAGEINDEX)
                gFindMatch slStr, 1, lbcLanguage
                If gLastFound(lbcLanguage) >= 1 Then
                    slNameCode = tmLanguageCode(gLastFound(lbcLanguage) - 1).sKey 'Traffic!lbcAgency.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmGsf(ilIndex).iLangMnfCode = Val(slCode)
                End If
            End If
            'Visiting Team
            tmGsf(ilIndex).iVisitMnfCode = 0
            slStr = grdDates.TextMatrix(ilRow, VISITTEAMINDEX)
            gFindMatch slStr, 1, lbcTeam
            If gLastFound(lbcTeam) >= 1 Then
                slNameCode = tmTeamCode(gLastFound(lbcTeam) - 1).sKey 'Traffic!lbcAgency.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmGsf(ilIndex).iVisitMnfCode = Val(slCode)
            End If
            'Home Team
            tmGsf(ilIndex).iHomeMnfCode = 0
            slStr = grdDates.TextMatrix(ilRow, HOMETEAMINDEX)
            gFindMatch slStr, 1, lbcTeam
            If gLastFound(lbcTeam) >= 1 Then
                slNameCode = tmTeamCode(gLastFound(lbcTeam) - 1).sKey 'Traffic!lbcAgency.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmGsf(ilIndex).iHomeMnfCode = Val(slCode)
            End If
            'Subtotal 1
            tmGsf(ilIndex).iSubtotal1MnfCode = 0
            If Trim$(tgSaf(0).sEventSubtotal1) <> "" Then
                slStr = grdDates.TextMatrix(ilRow, SUBTOTAL1INDEX)
                gFindMatch slStr, 1, lbcSubtotal(0)
                If gLastFound(lbcSubtotal(0)) >= 1 Then
                    slNameCode = tmSubtotal1Code(gLastFound(lbcSubtotal(0)) - 1).sKey 'Traffic!lbcAgency.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmGsf(ilIndex).iSubtotal1MnfCode = Val(slCode)
                End If
            End If
            'Subtotal 2
            tmGsf(ilIndex).iSubtotal2MnfCode = 0
            If Trim$(tgSaf(0).sEventSubtotal2) <> "" Then
                slStr = grdDates.TextMatrix(ilRow, SUBTOTAL2INDEX)
                gFindMatch slStr, 1, lbcSubtotal(1)
                If gLastFound(lbcSubtotal(1)) >= 1 Then
                    slNameCode = tmSubtotal2Code(gLastFound(lbcSubtotal(1)) - 1).sKey 'Traffic!lbcAgency.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmGsf(ilIndex).iSubtotal2MnfCode = Val(slCode)
                End If
            End If
            'Library
            tmGsf(ilIndex).lLvfCode = 0
            slStr = grdDates.TextMatrix(ilRow, VERLIBINDEX)
            gFindMatch slStr, 0, lbcLibrary
            If gLastFound(lbcLibrary) >= 0 Then
                slNameCode = tmLibName(gLastFound(lbcLibrary)).sKey 'Traffic!lbcAgency.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmGsf(ilIndex).lLvfCode = Val(slCode)
            End If
            'Date
            slStr = grdDates.TextMatrix(ilRow, AIRDATEINDEX)
            gPackDate slStr, tmGsf(ilIndex).iAirDate(0), tmGsf(ilIndex).iAirDate(1)
            'Time
            slStr = grdDates.TextMatrix(ilRow, AIRTIMEINDEX)
            gPackTime slStr, tmGsf(ilIndex).iAirTime(0), tmGsf(ilIndex).iAirTime(1)
            'Air Vehicle
            tmGsf(ilIndex).iAirVefCode = 0
            tmGsf(ilIndex).sLiveLogMerge = ""
            If (ilValue And PREEMPTREGPROG) = PREEMPTREGPROG Then
                slStr = grdDates.TextMatrix(ilRow, AIRVEHICLEINDEX)
                If Left$(slStr, 2) = "L:" Then
                    tmGsf(ilIndex).sLiveLogMerge = "L"
                    slStr = Trim$(Mid$(slStr, 3))
                ElseIf Left$(slStr, 2) = "M:" Then
                    tmGsf(ilIndex).sLiveLogMerge = "M"
                    slStr = Trim$(Mid$(slStr, 3))
                End If
                gFindMatch slStr, 1, lbcAirVehicle
                If gLastFound(lbcAirVehicle) >= 1 Then
                    slNameCode = tmAirVehicle(gLastFound(lbcAirVehicle) - 1).sKey 'Traffic!lbcAgency.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmGsf(ilIndex).iAirVefCode = Val(slCode)
                Else
                    tmGsf(ilIndex).sLiveLogMerge = ""
                End If
            End If
            'XDS Program Code ID
            If grdDates.ColWidth(XDSPROGCODEINDEX) > 0 Then
                tmGsf(ilIndex).sXDSProgCodeID = grdDates.TextMatrix(ilRow, XDSPROGCODEINDEX)
            Else
                tmGsf(ilIndex).sXDSProgCodeID = ""
            End If
            'Bus
            If grdDates.ColWidth(BUSINDEX) > 0 Then
                tmGsf(ilIndex).sBus = grdDates.TextMatrix(ilRow, BUSINDEX)
            Else
                tmGsf(ilIndex).sBus = ""
            End If
            'Game Status
            tmGsf(ilIndex).sGameStatus = ""
            slStr = grdDates.TextMatrix(ilRow, GAMESTATUSINDEX)
            Select Case UCase$(slStr)
                Case "C"    '"CANCELED"
                    tmGsf(ilIndex).sGameStatus = "C"
                Case "F"    '"FIRM"
                    tmGsf(ilIndex).sGameStatus = "F"
                Case "P"    '"POSTPONED"
                    tmGsf(ilIndex).sGameStatus = "P"
                Case "T"    '"TENTATIVE"
                    tmGsf(ilIndex).sGameStatus = "T"
            End Select
        End If
    Next ilRow
    Exit Sub
End Sub

Private Function mCreateLcfSsf(tlGsf As GSF, llOrigDate As Long) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slMsg As String

    tmLcf.iVefCode = imVefCode
    tmLcf.iLogDate(0) = tlGsf.iAirDate(0)
    tmLcf.iLogDate(1) = tlGsf.iAirDate(1)
    tmLcf.iSeqNo = 1
    tmLcf.iType = tlGsf.iGameNo
    tmLcf.sStatus = "C"
    tmLcf.sTiming = "N"
    tmLcf.sAffPost = "N"
    gPackTime "", tmLcf.iLastTime(0), tmLcf.iLastTime(1)
    For ilLoop = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
        If ilLoop = LBound(tmLcf.lLvfCode) Then
            tmLcf.lLvfCode(ilLoop) = tlGsf.lLvfCode
            tmLcf.iTime(0, ilLoop) = tlGsf.iAirTime(0)
            tmLcf.iTime(1, ilLoop) = tlGsf.iAirTime(1)
        Else
            gPackTime "", tmLcf.iTime(0, ilLoop), tmLcf.iTime(1, ilLoop)
        End If
    Next ilLoop
    tmLcf.iUrfCode = tgUrf(0).iCode
    tmLcf.lCode = 0
    ilRet = btrInsert(hmLcf, tmLcf, imLcfRecLen, INDEXKEY3)
    slMsg = "mCreateLcfSsf (btrInsert:Library Calendar)"
    On Error GoTo mCreateLcfSsfErr
    gBtrvErrorMsg ilRet, slMsg, GameSchd
    On Error GoTo 0
    ilRet = gMakeSSF(False, hmSsf, hmSdf, hmSmf, tmLcf.iType, imVefCode, tmLcf.iLogDate(0), tmLcf.iLogDate(1), 0, llOrigDate)
    If ilRet Then
        mCreateLcfSsf = True
    Else
        mCreateLcfSsf = False
    End If
    Exit Function
mCreateLcfSsfErr:
    On Error GoTo 0
    mCreateLcfSsf = False
    Exit Function

End Function

Private Function mUpdateLcfSsf(tlGsf As GSF, llOrigDate As Long, llOrigSeasonStartDate As Long, llOrigSeasonEndDate As Long) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llSdfDate                     llTime                                                  *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slMsg As String
    Dim llGsfDate As Long
    Dim llMoGsfDate As Long
    Dim llCgfDate As Long
    Dim llMoCgfDate As Long
    Dim llClfEndDate As Long
    Dim llNewClfEndDate As Long
    Dim ilClf As Integer
    Dim tlCff As CFF
    Dim tlGhf As GHF
    Dim ilSdf As Integer
    Dim ilSmf As Integer
    Dim llTimeAdj As Long
    Dim llOldTime As Long
    Dim llNewTime As Long
    Dim ilFound As Integer
    Dim llSdfDate As Long
    Dim llMoOrigGsfDate As Long
    Dim llSeasonStartDate As Long
    Dim llSeasonEndDate As Long
    Dim blWeekChgd As Boolean
    Dim ilGhf As Integer
    Dim slDate As String
    Dim blWeekGameFound As Boolean
    Dim blMoveSpot As Boolean
    Dim llSsfRecPos As Long
    Dim llSsfDate As Long
    Dim slSpotType As String
    Dim slSchStatus As String
    Dim blCgfFound As Boolean
    Dim tlWeekGsf As GSF
    Dim tlCgf As CGF

    gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStartDate
    gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEndDate
    gUnpackDateLong tlGsf.iAirDate(0), tlGsf.iAirDate(1), llGsfDate
    llMoGsfDate = llGsfDate
    Do While gWeekDayLong(llMoGsfDate) <> 0   '0=monday
        llMoGsfDate = llMoGsfDate - 1
    Loop
    llMoOrigGsfDate = llOrigDate
    Do While gWeekDayLong(llMoOrigGsfDate) <> 0   '0=monday
        llMoOrigGsfDate = llMoOrigGsfDate - 1
    Loop

    blWeekChgd = False
    blWeekGameFound = False
    If llMoGsfDate <> llMoOrigGsfDate Then
        blWeekChgd = True
        'Determine if another game exist in original week
        tmGsfSrchKey1.lghfcode = tmGhf.lCode
        tmGsfSrchKey1.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hmGsf, tlWeekGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.lCode = tlWeekGsf.lghfcode)
            gUnpackDate tlWeekGsf.iAirDate(0), tlWeekGsf.iAirDate(1), slDate
            slDate = gObtainPrevMonday(slDate)
            If (gDateValue(slDate) = llMoOrigGsfDate) And (tlWeekGsf.iGameNo <> tlGsf.iGameNo) Then
                blWeekGameFound = True
                Exit Do
            End If
            ilRet = btrGetNext(hmGsf, tlWeekGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    End If
    
    'tmLcfSrchKey1.iType = tlGsf.iGameNo
    'tmLcfSrchKey1.iVefCode = imVefCode
    'Do
    '    ilRet = btrGetEqual(hmLcf, tmDeletedLcf, imLcfRecLen, tmLcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    '    If ilRet <> BTRV_ERR_NONE Then
    '        '8/29/06: LCF missing, then create it first
    '        If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
    '            ilRet = mCreateLcfSsf(tlGsf)
    '            ilRet = BTRV_ERR_CONFLICT
    '        Else
    '            On Error GoTo mUpdateLcfSsfErr
    '            gBtrvErrorMsg ilRet, slMsg, GameSchd
    '            On Error GoTo 0
    '        End If
    '    Else
    '        gUnpackTimeLong tmDeletedLcf.iTime(0, LBound(tmLcf.lLvfCode)), tmDeletedLcf.iTime(1, LBound(tmLcf.lLvfCode)), False, llOldTime
    '        ilRet = btrDelete(hmLcf)
    '    End If
    'Loop While ilRet = BTRV_ERR_CONFLICT
    ilFound = False
    tmLcfSrchKey2.iVefCode = imVefCode
    gPackDateLong llOrigDate, tmLcfSrchKey2.iLogDate(0), tmLcfSrchKey2.iLogDate(1)
    Do
        ilRet = btrGetEqual(hmLcf, tmDeletedLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmDeletedLcf.iVefCode = imVefCode)
            If (tmLcfSrchKey2.iLogDate(0) = tmDeletedLcf.iLogDate(0)) And (tmLcfSrchKey2.iLogDate(1) = tmDeletedLcf.iLogDate(1)) Then
                If (tmDeletedLcf.iType = tlGsf.iGameNo) Then
                    ilFound = True
                    gUnpackTimeLong tmDeletedLcf.iTime(0, LBound(tmLcf.lLvfCode)), tmDeletedLcf.iTime(1, LBound(tmLcf.lLvfCode)), False, llOldTime
                    ilRet = btrDelete(hmLcf)
                    Exit Do
                End If
            Else
                Exit Do
            End If
            ilRet = btrGetNext(hmLcf, tmDeletedLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        If Not ilFound Then
            ilRet = mCreateLcfSsf(tlGsf, 0)
            ilRet = BTRV_ERR_CONFLICT
            tmLcfSrchKey2.iVefCode = imVefCode
            gPackDateLong llOrigDate, tmLcfSrchKey2.iLogDate(0), tmLcfSrchKey2.iLogDate(1)
        End If
        ilFound = False
    Loop While ilRet = BTRV_ERR_CONFLICT
    gUnpackTimeLong tlGsf.iAirTime(0), tlGsf.iAirTime(1), False, llNewTime
    llTimeAdj = llNewTime - llOldTime
    'gUnpackDateLong tlGsf.iAirDate(0), tlGsf.iAirDate(1), llGsfDate

    tmLcf.iVefCode = imVefCode
    tmLcf.iLogDate(0) = tlGsf.iAirDate(0)
    tmLcf.iLogDate(1) = tlGsf.iAirDate(1)
    tmLcf.iSeqNo = 1
    tmLcf.iType = tlGsf.iGameNo
    tmLcf.sStatus = "C"
    tmLcf.sTiming = "N"
    tmLcf.sAffPost = "N"
    gPackTime "", tmLcf.iLastTime(0), tmLcf.iLastTime(1)
    For ilLoop = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
        If ilLoop = LBound(tmLcf.lLvfCode) Then
            tmLcf.lLvfCode(ilLoop) = tlGsf.lLvfCode
            tmLcf.iTime(0, ilLoop) = tlGsf.iAirTime(0)
            tmLcf.iTime(1, ilLoop) = tlGsf.iAirTime(1)
        Else
            gPackTime "", tmLcf.iTime(0, ilLoop), tmLcf.iTime(1, ilLoop)
        End If
    Next ilLoop
    tmLcf.iUrfCode = tgUrf(0).iCode
    tmLcf.lCode = 0
    ilRet = btrInsert(hmLcf, tmLcf, imLcfRecLen, INDEXKEY3)
    slMsg = "mUpdateLcfSsf (btrInsert:Library Calendar)"
    On Error GoTo mUpdateLcfSsfErr
    gBtrvErrorMsg ilRet, slMsg, GameSchd
    On Error GoTo 0

    'tmSsfSrchKey1.iType = tlGsf.iGameNo
    'tmSsfSrchKey1.iVefCode = imVefCode
    'Do
    '    imSsfRecLen = Len(tmDeletedSsf)
    '    ilRet = btrGetEqual(hmSsf, tmDeletedSsf, imSsfRecLen, tmSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    '    If ilRet <> BTRV_ERR_NONE Then
    '        On Error GoTo mUpdateLcfSsfErr
    '        gBtrvErrorMsg ilRet, slMsg, GameSchd
    '        On Error GoTo 0
    '    End If
    '    ilRet = btrDelete(hmSsf)
    'Loop While ilRet = BTRV_ERR_CONFLICT
    'ReDim lgReschSdfCode(1 To 1) As Long
    ReDim lgReschSdfCode(0 To 0) As Long
    ilRet = gMakeSSF(False, hmSsf, hmSdf, hmSmf, tmLcf.iType, imVefCode, tmLcf.iLogDate(0), tmLcf.iLogDate(1), llTimeAdj, llOrigDate)
    If ilRet Then
        'ilRet = gReSchSpots(False, 0, "YYYYYYY", 0, 86400)
        If UBound(lgReschSdfCode) > LBound(lgReschSdfCode) Then
            gGetSchParameters
            'gObtainMissedReasonCode
            Randomize   'Remove this if same results are to be obtained
            If gOpenSchFiles() Then
                igReschNoPasses = 2
                igSetEarliestDateAsToday = 1
                ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
                igSetEarliestDateAsToday = 1
                igReschNoPasses = 1
                gCloseSchFiles
            End If
        End If
        mUpdateLcfSsf = True
    Else
        mUpdateLcfSsf = False
    End If
'    Do
'        tmSsfSrchKey1.iType = tlGsf.iGameNo
'        tmSsfSrchKey1.iVefCode = imVefCode
'        imSsfRecLen = Len(tmSsf)
'        ilRet = btrGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
'        If ilRet <> BTRV_ERR_NONE Then
'            Exit Do
'        End If
'        tmSsf.iDate(0) = tlGsf.iAirDate(0)
'        tmSsf.iDate(1) = tlGsf.iAirDate(1)
'        'Adjust Time of each event
'        If llTimeAdj <> 0 Then
'            For ilLoop = 1 To tmSsf.iCount Step 1
'               LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilLoop)
'                If tmProg.iRecType = 1 Then
'                    gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llTime
'                    llTime = llTime + llTimeAdj
'                    gPackTimeLong llTime, tmProg.iStartTime(0), tmProg.iStartTime(1)
'                    gUnpackTimeLong tmProg.iEndTime(0), tmProg.iEndTime(1), True, llTime
'                    llTime = llTime + llTimeAdj
'                    gPackTimeLong llTime, tmProg.iEndTime(0), tmProg.iEndTime(1)
'                    tmSsf.tPas(ADJSSFPASBZ + ilLoop) = tmProg
'                Else
'                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilLoop)
'                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
'                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
'                        llTime = llTime + llTimeAdj
'                        gPackTimeLong llTime, tmAvail.iTime(0), tmAvail.iTime(1)
'                        tmSsf.tPas(ADJSSFPASBZ + ilLoop) = tmAvail
'                    End If
'                End If
'            Next ilLoop
'        End If
'        ilRet = btrUpdate(hmSsf, tmSsf, imSsfRecLen)
'    Loop While ilRet = BTRV_ERR_CONFLICT
    'Update Scheduled spots
    ReDim lmSdfCode(0 To 0) As Long
    ReDim lmBBSdfCode(0 To 0) As Long
    ReDim lmSmfCode(0 To 0) As Long
    tmSdfSrchKey6.iVefCode = imVefCode
    tmSdfSrchKey6.iGameNo = tlGsf.iGameNo
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey6, INDEXKEY6, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = imVefCode) And (tmSdf.iGameNo = tlGsf.iGameNo)
        'Reset all Missed as date or start time or length of game could have changed
        'gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
        'If (llGsfDate <> llSdfDate) And (tmSdf.sSchStatus = "M") Then
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
        If ((llGsfDate = llSdfDate) And (tmSdf.sSchStatus <> "M")) Or ((llOrigDate = llSdfDate) And (tmSdf.sSchStatus = "M")) Then
            If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
                lmBBSdfCode(UBound(lmBBSdfCode)) = tmSdf.lCode
                ReDim Preserve lmBBSdfCode(0 To UBound(lmBBSdfCode) + 1) As Long
            Else
                If tmSdf.sSchStatus = "M" Then
                    lmSdfCode(UBound(lmSdfCode)) = tmSdf.lCode
                    ReDim Preserve lmSdfCode(0 To UBound(lmSdfCode) + 1) As Long
                End If
                If (blWeekChgd) And (blWeekGameFound) Then
                    'Unschedule spot
                    If (tmSdf.sSchStatus = "S") Then
                        lmSdfCode(UBound(lmSdfCode)) = tmSdf.lCode
                        ReDim Preserve lmSdfCode(0 To UBound(lmSdfCode) + 1) As Long
                    ElseIf (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                        lmSdfCode(UBound(lmSdfCode)) = tmSdf.lCode
                        ReDim Preserve lmSdfCode(0 To UBound(lmSdfCode) + 1) As Long
                    End If
                Else
                    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                        lmSmfCode(UBound(lmSmfCode)) = tmSdf.lSmfCode
                        ReDim Preserve lmSmfCode(0 To UBound(lmSmfCode) + 1) As Long
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop

    'ReDim lgReschSdfCode(1 To 1) As Long
    ReDim lgReschSdfCode(0 To 0) As Long
    For ilSdf = 0 To UBound(lmSdfCode) - 1 Step 1
        tmSdfSrchKey3.lCode = lmSdfCode(ilSdf)
        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            blMoveSpot = True
            If (blWeekChgd) And (blWeekGameFound) Then
                If mSportsByWeek(tmSdf.lChfCode, tmSdf.iLineNo) = "W" Then
                    blMoveSpot = False
                End If
            End If
            If Not blMoveSpot Then
                '12/19/12: Unschedule spot and change game number
                llSsfRecPos = 0
                slSpotType = tmSdf.sSpotType
                slSchStatus = tmSdf.sSchStatus
                'If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                '    tmSmfSrchKey2.lCode = tmSdf.lCode
                '    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                '    If ilRet <> BTRV_ERR_NONE Then
                '        tmSmf.iGameNo = tmSdf.iGameNo
                '    End If
                'Else
                '    tmSmf.iGameNo = tmSdf.iGameNo
                'End If
                If tmSdf.sSchStatus <> "M" Then
                    ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, tmSdf.iGameNo, tmSmf, hmSsf, tmSsf, llSsfDate, llSsfRecPos, hmSxf, hmGsf, hmGhf)
                    If slSpotType <> "X" Then
                        If (slSchStatus = "S") Then
                            tmSdf.iGameNo = tlWeekGsf.iGameNo
                            tmSdf.iDate(0) = tlWeekGsf.iAirDate(0)
                            tmSdf.iDate(1) = tlWeekGsf.iAirDate(1)
                            tmSdf.iTime(0) = tlWeekGsf.iAirTime(0)
                            tmSdf.iTime(1) = tlWeekGsf.iAirTime(1)
                            lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                            'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                            ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                        Else
                            tmSdf.iGameNo = tmSmf.iGameNo
                            tmSdf.iDate(0) = tmSmf.iMissedDate(0)
                            tmSdf.iDate(1) = tmSmf.iMissedDate(1)
                            tmSdf.iTime(0) = tmSmf.iMissedTime(0)
                            tmSdf.iTime(1) = tmSmf.iMissedTime(1)
                            tmSdf.iVefCode = tmSmf.iOrigSchVef
                        End If
                        tmSdf.sXCrossMidnight = "N"
                        ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                    End If
                End If
            ElseIf tmSdf.sSchStatus = "M" Then
                tmSdf.iDate(0) = tlGsf.iAirDate(0)
                tmSdf.iDate(1) = tlGsf.iAirDate(1)
                'Reset missed time to start of game time as length could have changed
                'If llTimeAdj <> 0 Then
                '    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                '    llTime = llTime + llTimeAdj
                '    gPackTimeLong llTime, tmSdf.iTime(0), tmSdf.iTime(1)
                'End If
                tmSdf.iTime(0) = tlGsf.iAirTime(0)
                tmSdf.iTime(1) = tlGsf.iAirTime(1)
                tmSdf.sXCrossMidnight = "N"
                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
            End If
        End If
    Next ilSdf
    '8/2/11: Find MG's scheduled into different games
    tmSmfSrchKey3.iOrigSchVef = imVefCode
    tmSmfSrchKey3.iGameNo = tlGsf.iGameNo
    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.iOrigSchVef = imVefCode) And (tmSmf.iGameNo = tlGsf.iGameNo)
        gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llSdfDate
        If (llOrigDate = llSdfDate) Then
            lmSmfCode(UBound(lmSmfCode)) = tmSmf.lCode
            ReDim Preserve lmSmfCode(0 To UBound(lmSmfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop

    For ilSmf = 0 To UBound(lmSmfCode) - 1 Step 1
        tmSmfSrchKey1.lCode = lmSmfCode(ilSmf)
        ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            blMoveSpot = True
            If (blWeekChgd) And (blWeekGameFound) Then
                If mSportsByWeek(tmSmf.lChfCode, tmSmf.iLineNo) = "W" Then
                    blMoveSpot = False
                End If
            End If
            If Not blMoveSpot Then
                tmSmf.iGameNo = tlWeekGsf.iGameNo
                tmSmf.iMissedDate(0) = tlWeekGsf.iAirDate(0)
                tmSmf.iMissedDate(1) = tlWeekGsf.iAirDate(1)
                tmSmf.iMissedTime(0) = tlWeekGsf.iAirTime(0)
                tmSmf.iMissedTime(1) = tlWeekGsf.iAirTime(1)
                '10/20/14: Update actual date/time. TTP 7197
                tmSmf.iActualDate(0) = tlWeekGsf.iAirDate(0)
                tmSmf.iActualDate(1) = tlWeekGsf.iAirDate(1)
                tmSmf.iActualTime(0) = tlWeekGsf.iAirTime(0)
                tmSmf.iActualTime(1) = tlWeekGsf.iAirTime(1)
            Else
                tmSmf.iMissedDate(0) = tlGsf.iAirDate(0)
                tmSmf.iMissedDate(1) = tlGsf.iAirDate(1)
                tmSmf.iMissedTime(0) = tlGsf.iAirTime(0)
                tmSmf.iMissedTime(1) = tlGsf.iAirTime(1)
                '10/20/14: Update actual date/time. TTP 7197
                tmSmf.iActualDate(0) = tlGsf.iAirDate(0)
                tmSmf.iActualDate(1) = tlGsf.iAirDate(1)
                tmSmf.iActualTime(0) = tlGsf.iAirTime(0)
                tmSmf.iActualTime(1) = tlGsf.iAirTime(1)
            End If
            ilRet = btrUpdate(hmSmf, tmSmf, imSmfRecLen)
        End If
    Next ilSmf

    '5/7/11: Remove Billboards, they will be created again when Logs or export or copy report run
    For ilSdf = 0 To UBound(lmBBSdfCode) - 1 Step 1
        tmSdfSrchKey3.lCode = lmBBSdfCode(ilSdf)
        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hmSdf)
        End If
    Next ilSdf

    'Update Contracts
    'llMoGsfDate = llGsfDate
    'Do While gWeekDayLong(llMoGsfDate) <> 0   '0=monday
    '    llMoGsfDate = llMoGsfDate - 1
    'Loop
    ReDim lmClfCode(0 To 0) As Long
    ReDim llGhfCode(0 To 0) As Long
    tmGhfSrchKey1.iVefCode = imVefCode
    ilRet = btrGetEqual(hmGhf, tlGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlGhf.iVefCode = imVefCode)
        llGhfCode(UBound(llGhfCode)) = tlGhf.lCode
        ReDim Preserve llGhfCode(0 To UBound(llGhfCode) + 1) As Long
        ilRet = btrGetNext(hmGhf, tlGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    '12/19/12: Fix flights including spot moved to another game
    For ilGhf = 0 To UBound(llGhfCode) - 1 Step 1
        ReDim lmClfCode(0 To 0) As Long
        tmClfSrchKey3.lghfcode = llGhfCode(ilGhf)   'tmGhf.lCode
        tmClfSrchKey3.iEndDate(0) = 0
        tmClfSrchKey3.iEndDate(1) = 0
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lghfcode = llGhfCode(ilGhf)) 'tmGhf.lCode)
            If tmClf.sDelete = "N" Then
                lmClfCode(UBound(lmClfCode)) = tmClf.lCode
                ReDim Preserve lmClfCode(0 To UBound(lmClfCode) + 1) As Long
            End If
            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
        For ilClf = 0 To UBound(lmClfCode) - 1 Step 1
            tmClfSrchKey2.lCode = lmClfCode(ilClf)
            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                blMoveSpot = True
                If (blWeekChgd) And (blWeekGameFound) Then
                    If tmClf.sSportsByWeek = "W" Then
                        blMoveSpot = False
                    End If
                End If
                gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llClfEndDate
                llNewClfEndDate = llClfEndDate
                tmCgfSrchKey1.lClfCode = lmClfCode(ilClf)
                ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmCgf.lClfCode = lmClfCode(ilClf))
                    If tlGsf.iGameNo = tmCgf.iGameNo Then
                        'Determine if airing is a different week
                        gUnpackDateLong tmCgf.iAirDate(0), tmCgf.iAirDate(1), llCgfDate
                        If (llCgfDate >= llOrigSeasonStartDate) And (llCgfDate <= llOrigSeasonEndDate) Then
                            If llGsfDate <> llCgfDate Then  'Game moved
                                'Adjust CGF
                                If Not blMoveSpot Then
                                    'Move spots
                                    ilRet = btrDelete(hmCgf)
                                    blCgfFound = False
                                    tmCgfSrchKey1.lClfCode = lmClfCode(ilClf)
                                    ilRet = btrGetEqual(hmCgf, tlCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                    Do While (ilRet = BTRV_ERR_NONE) And (tlCgf.lClfCode = lmClfCode(ilClf))
                                        If tlWeekGsf.iGameNo = tlCgf.iGameNo Then
                                            'Determine if airing is a different week
                                            gUnpackDateLong tlCgf.iAirDate(0), tlCgf.iAirDate(1), llCgfDate
                                            If (llCgfDate >= llOrigSeasonStartDate) And (llCgfDate <= llOrigSeasonEndDate) Then
                                                tlCgf.iNoSpots = tlCgf.iNoSpots + tmCgf.iNoSpots
                                                ilRet = btrUpdate(hmCgf, tlCgf, imCgfRecLen)
                                                blCgfFound = True
                                                Exit Do
                                            End If
                                        End If
                                        ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                    Loop
                                    If Not blCgfFound Then
                                        tmCgf.lCode = 0
                                        tmCgf.iGameNo = tlWeekGsf.iGameNo
                                        tmCgf.iAirDate(0) = tlWeekGsf.iAirDate(0)
                                        tmCgf.iAirDate(1) = tlWeekGsf.iAirDate(1)
                                        ilRet = btrInsert(hmCgf, tmCgf, imCgfRecLen, INDEXKEY0)
                                    End If
                                Else
                                    tmCgf.iAirDate(0) = tlGsf.iAirDate(0)
                                    tmCgf.iAirDate(1) = tlGsf.iAirDate(1)
                                    ilRet = btrUpdate(hmCgf, tmCgf, imCgfRecLen)
                                End If
                                llMoCgfDate = llCgfDate
                                Do While gWeekDayLong(llMoCgfDate) <> 0   '0=monday
                                    llMoCgfDate = llMoCgfDate - 1
                                Loop
                                If llMoGsfDate <> llMoCgfDate Then  'Game is different week
                                    'Adjust flight, Remove spots from week and add into new week
                                    tmCffSrchKey0.lChfCode = tmClf.lChfCode
                                    tmCffSrchKey0.iClfLine = tmClf.iLine
                                    tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
                                    tmCffSrchKey0.iPropVer = tmClf.iPropVer
                                    gPackDateLong llMoCgfDate, tmCffSrchKey0.iStartDate(0), tmCffSrchKey0.iStartDate(1)
                                    ilRet = btrGetEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                    If ilRet = BTRV_ERR_NONE Then
                                        'If tmCff.iSpotsWk = tmCgf.iNoSpots Then
                                        '    'Change Date, delete record and insert as key can't be modified
                                        '    ilRet = btrDelete(hmCff)
                                        '    gPackDateLong llMoGsfDate, tmCff.iStartDate(0), tmCff.iStartDate(1)
                                        '    'ilRet = btrInsert(hmCff, tmCff, imCffRecLen, INDEXKEY0)
                                        'Else
                                            'Reduce Spots
                                            tmCff.iSpotsWk = tmCff.iSpotsWk - tmCgf.iNoSpots
                                            If tmCff.iSpotsWk > 0 Then
                                                ilRet = btrUpdate(hmCff, tmCff, imCffRecLen)
                                            Else
                                                ilRet = btrDelete(hmCff)
                                            End If
                                            'Either add week or Increase number of spots
                                            tmCffSrchKey0.lChfCode = tmClf.lChfCode
                                            tmCffSrchKey0.iClfLine = tmClf.iLine
                                            tmCffSrchKey0.iCntRevNo = tmClf.iCntRevNo
                                            tmCffSrchKey0.iPropVer = tmClf.iPropVer
                                            If Not blMoveSpot Then
                                                gPackDateLong llMoOrigGsfDate, tmCffSrchKey0.iStartDate(0), tmCffSrchKey0.iStartDate(1)
                                            Else
                                                gPackDateLong llMoGsfDate, tmCffSrchKey0.iStartDate(0), tmCffSrchKey0.iStartDate(1)
                                            End If
                                            ilRet = btrGetEqual(hmCff, tlCff, imCffRecLen, tmCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                            If ilRet = BTRV_ERR_NONE Then
                                                tlCff.iSpotsWk = tlCff.iSpotsWk + tmCgf.iNoSpots
                                                ilRet = btrUpdate(hmCff, tlCff, imCffRecLen)
                                            Else
                                                tmCff.iSpotsWk = tmCgf.iNoSpots
                                                If Not blMoveSpot Then
                                                    gPackDateLong llMoOrigGsfDate, tmCff.iStartDate(0), tmCff.iStartDate(1)
                                                    gPackDateLong llMoOrigGsfDate + 6, tmCff.iEndDate(0), tmCff.iEndDate(1)
                                                Else
                                                    gPackDateLong llMoGsfDate, tmCff.iStartDate(0), tmCff.iStartDate(1)
                                                    gPackDateLong llMoGsfDate + 6, tmCff.iEndDate(0), tmCff.iEndDate(1)
                                                End If
                                                tmCff.lCode = 0
                                                ilRet = btrInsert(hmCff, tmCff, imCffRecLen, INDEXKEY1)
                                                If llMoGsfDate + 6 > llNewClfEndDate Then
                                                    llNewClfEndDate = llMoGsfDate + 6
                                                End If
                                            End If
                                        'End If
                                    End If
                                End If
                            End If
                            Exit Do
                        End If
                    End If
                    ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                Loop
                If llNewClfEndDate > llClfEndDate Then
                    gPackDateLong llNewClfEndDate, tmClf.iEndDate(0), tmClf.iEndDate(1)
                    ilRet = btrUpdate(hmClf, tmClf, imClfRecLen)
                End If
            End If
        Next ilClf
    Next ilGhf
    '12/20/12: Reschedule spots placed into missed because event moved to another week and lines
    '          defined as weekly buy.
    '          Placed here because the spots can't be reschedule until the orders are corrected.
    If (blWeekChgd) And (blWeekGameFound) And (UBound(lgReschSdfCode) > LBound(lgReschSdfCode)) Then
        gGetSchParameters
        'gObtainMissedReasonCode
        Randomize   'Remove this if same results are to be obtained
        If gOpenSchFiles() Then
            igReschNoPasses = 2
            igSetEarliestDateAsToday = 1
            ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
            igSetEarliestDateAsToday = 1
            igReschNoPasses = 1
            gCloseSchFiles
        End If
    End If
    
    mUpdateLcfSsf = True
    Exit Function
mUpdateLcfSsfErr:
    On Error GoTo 0
    mUpdateLcfSsf = False
    Exit Function
End Function


Private Function mPreemptAndLock(ilAvailLock As Integer, ilOrigVefCode As Integer, ilNewVefCode As Integer, slOrigGameStatus As String, slNewGameStatus As String, llOrigDate As Long, llNewDate As Long, llOrigStartTime As Long, llNewStartTime As Long, llOrigLength As Long, llNewLength As Long) As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilSpotLock As Integer
    Dim llLastLogDate As Long
    Dim ilVpfIndex As Integer
    Dim ilRet As Integer
    Dim llChfCode As Long
    Dim ilLoop As Integer
    Dim slMoStartDate As String
    Dim slSuEndDate As String
    Dim ilOrigXMid As Integer
    Dim ilNewXMid As Integer

    mPreemptAndLock = 0
    If (Asc(tgSpf.sSportInfo) And PREEMPTREGPROG) <> PREEMPTREGPROG Then
        Exit Function
    End If
    ilOrigXMid = False
    ilNewXMid = False
    If ilAvailLock = 0 Then
        If (ilOrigVefCode > 0) And (slOrigGameStatus = "F") And (llOrigDate >= lmFirstAllowedChgDate) Then
            If Not mLockPreemptVehicle(ilOrigVefCode, llOrigDate) Then
                mPreemptAndLock = 1
                Exit Function
            End If
            ilSpotLock = 0  'Unlock
            slStartDate = Format$(llOrigDate, "m/d/yy")
            slEndDate = slStartDate
            slStartTime = gFormatTimeLong(llOrigStartTime, "A", "1")
            If llOrigStartTime + llOrigLength <= 86400 Then
                slEndTime = gFormatTimeLong(llOrigStartTime + llOrigLength - 1, "A", "1")
                gSetLockStatus ilOrigVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slStartTime, slEndTime
            Else
                If gWeekDayLong(llOrigDate) = 6 Then
                    ilOrigXMid = True
                End If
                slEndTime = gFormatTimeLong(86399, "A", "1")
                gSetLockStatus ilOrigVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slStartTime, slEndTime
                slStartDate = Format$(llOrigDate + 1, "m/d/yy")
                slEndDate = slStartDate
                slStartTime = gFormatTimeLong(0, "A", "1")
                slEndTime = gFormatTimeLong(llOrigStartTime + llOrigLength - 86401, "A", "1")
                gSetLockStatus ilOrigVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slStartTime, slEndTime
            End If
            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
        End If
        Exit Function
    End If
    If (ilNewVefCode > 0) And (slNewGameStatus = "F") And (llNewDate >= lmFirstAllowedChgDate) Then
        If Not mLockPreemptVehicle(ilNewVefCode, llNewDate) Then
            mPreemptAndLock = 1
            Exit Function
        End If
        ilSpotLock = 0  'Unlock
        slStartDate = Format$(llNewDate, "m/d/yy")
        slEndDate = slStartDate
        slStartTime = gFormatTimeLong(llNewStartTime, "A", "1")
        If llNewStartTime + llNewLength <= 86400 Then
            slEndTime = gFormatTimeLong(llNewStartTime + llNewLength - 1, "A", "1")
            gSetLockStatus ilNewVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slStartTime, slEndTime
            ilVpfIndex = gVpfFind(GameSchd, ilNewVefCode)
            gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llLastLogDate
            llChfCode = -1
            ilRet = gUnschSpots(ilNewVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slStartTime, slEndTime, -1)
            If Not ilRet Then
                ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
                mPreemptAndLock = 2
                Exit Function
            End If
            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
        Else
            If gWeekDayLong(llNewDate) = 6 Then
                ilNewXMid = True
            End If
            slEndTime = gFormatTimeLong(86399, "A", "1")
            gSetLockStatus ilNewVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slStartTime, slEndTime
            ilVpfIndex = gVpfFind(GameSchd, ilNewVefCode)
            gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llLastLogDate
            llChfCode = -1
            ilRet = gUnschSpots(ilNewVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slStartTime, slEndTime, -1)
            If Not ilRet Then
                ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
                mPreemptAndLock = 2
                Exit Function
            End If
            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
            If Not mLockPreemptVehicle(ilNewVefCode, llNewDate + 1) Then
                mPreemptAndLock = 1
                Exit Function
            End If
            slStartDate = Format$(llNewDate + 1, "m/d/yy")
            slEndDate = slStartDate
            slStartTime = gFormatTimeLong(0, "A", "1")
            slEndTime = gFormatTimeLong(llNewStartTime + llNewLength - 86401, "A", "1")
            gSetLockStatus ilNewVefCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slStartTime, slEndTime
            ilVpfIndex = gVpfFind(GameSchd, ilNewVefCode)
            gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llLastLogDate
            llChfCode = -1
            ilRet = gUnschSpots(ilNewVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slStartTime, slEndTime, -1)
            If Not ilRet Then
                ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
                mPreemptAndLock = 2
                Exit Function
            End If
            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
        End If
    End If
    'Reschedule spots preempted
    If (ilOrigVefCode > 0) And (slOrigGameStatus = "F") And (llOrigDate >= lmFirstAllowedChgDate) Then
        sgMovePass = "N"
        sgCompPass = "N"
        'ReDim lgReschSdfCode(1 To 1) As Long
        ReDim lgReschSdfCode(0 To 0) As Long
        slStartDate = Format$(llOrigDate, "m/d/yy")
        slEndDate = slStartDate
        smSdfMdExtTag = ""
        slMoStartDate = gObtainPrevMonday(slStartDate)
        If Not ilOrigXMid Then
            slSuEndDate = gObtainNextSunday(slEndDate)
        Else
            slSuEndDate = gIncOneDay(gObtainNextSunday(slEndDate))
        End If
        ilRet = gObtainMissedSpot("M", ilOrigVefCode, -1, 0, slMoStartDate, slSuEndDate, 1, tmSdfMdExt(), smSdfMdExtTag)
        'For ilLoop = LBound(tmSdfMdExt) To UBound(tmSdfMdExt) - 1 Step 1
        For ilLoop = imLBSdfMdExt To UBound(tmSdfMdExt) - 1 Step 1
            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tmSdfMdExt(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                If (tmSdf.sSchStatus = "M") Then
                    lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                    'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                    ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                End If
            End If
        Next ilLoop
        'If UBound(lgReschSdfCode) > 1 Then
        If UBound(lgReschSdfCode) > LBound(lgReschSdfCode) Then
            gGetSchParameters
            'gObtainMissedReasonCode
            Randomize   'Remove this if same results are to be obtained
            If gOpenSchFiles() Then
                igReschNoPasses = 2
                igSetEarliestDateAsToday = 1
                ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
                igSetEarliestDateAsToday = 1
                igReschNoPasses = 1
                gCloseSchFiles
            End If
        End If
    End If
    If (ilNewVefCode > 0) And (slNewGameStatus = "F") And (llNewDate >= lmFirstAllowedChgDate) Then
        sgMovePass = "N"
        sgCompPass = "N"
        'ReDim lgReschSdfCode(1 To 1) As Long
        ReDim lgReschSdfCode(0 To 0) As Long
        slStartDate = Format$(llNewDate, "m/d/yy")
        slEndDate = slStartDate
        smSdfMdExtTag = ""
        slMoStartDate = gObtainPrevMonday(slStartDate)
        If Not ilNewXMid Then
            slSuEndDate = gObtainNextSunday(slEndDate)
        Else
            slSuEndDate = gIncOneDay(gObtainNextSunday(slEndDate))
        End If
        ilRet = gObtainMissedSpot("M", ilNewVefCode, -1, 0, slMoStartDate, slSuEndDate, 1, tmSdfMdExt(), smSdfMdExtTag)
        'For ilLoop = LBound(tmSdfMdExt) To UBound(tmSdfMdExt) - 1 Step 1
        For ilLoop = imLBSdfMdExt To UBound(tmSdfMdExt) - 1 Step 1
            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tmSdfMdExt(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                If (tmSdf.sSchStatus = "M") Then
                    lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                    'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                    ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                End If
            End If
        Next ilLoop
        'If UBound(lgReschSdfCode) > 1 Then
        If UBound(lgReschSdfCode) > LBound(lgReschSdfCode) Then
            gGetSchParameters
            'gObtainMissedReasonCode
            Randomize   'Remove this if same results are to be obtained
            If gOpenSchFiles() Then
                igReschNoPasses = 2
                igSetEarliestDateAsToday = 1
                ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
                igSetEarliestDateAsToday = 1
                igReschNoPasses = 1
                gCloseSchFiles
            End If
        End If
    End If
End Function

Private Sub mReschGameMissed(tlGsf As GSF)
    Dim ilRet As Integer
    Dim ilGameNo As Integer
    Dim slAirDate As String
    Dim slMoStartDate As String
    Dim slSuEndDate As String
    Dim ilLoop As Integer

    ilGameNo = tlGsf.iGameNo
    gUnpackDate tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), slAirDate
    If gDateValue(slAirDate) < lmFirstAllowedChgDate Then
        Exit Sub
    End If
    slMoStartDate = gObtainPrevMonday(slAirDate)
    slSuEndDate = gIncOneDay(gObtainNextSunday(slAirDate))
    sgMovePass = "N"
    sgCompPass = "N"
    'ReDim lgReschSdfCode(1 To 1) As Long
    ReDim lgReschSdfCode(0 To 0) As Long
    smSdfMdExtTag = ""
    ilRet = gObtainMissedSpot("M", imVefCode, -1, ilGameNo, slMoStartDate, slSuEndDate, 1, tmSdfMdExt(), smSdfMdExtTag)
    'For ilLoop = LBound(tmSdfMdExt) To UBound(tmSdfMdExt) - 1 Step 1
    For ilLoop = imLBSdfMdExt To UBound(tmSdfMdExt) - 1 Step 1
        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tmSdfMdExt(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            If (tmSdf.sSchStatus = "M") Then
                lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
            End If
        End If
    Next ilLoop
    'If UBound(lgReschSdfCode) > 1 Then
    If UBound(lgReschSdfCode) > LBound(lgReschSdfCode) Then
        gGetSchParameters
        'gObtainMissedReasonCode
        Randomize   'Remove this if same results are to be obtained
        If gOpenSchFiles() Then
            igReschNoPasses = 2
            igSetEarliestDateAsToday = 1
            ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
            igSetEarliestDateAsToday = 1
            igReschNoPasses = 1
            gCloseSchFiles
        End If
    End If
End Sub

Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        slStr = Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX))
        If slStr <> "" Then
            If ilCol = AIRDATEINDEX Then
                slSort = Trim$(str$(gDateValue(grdDates.TextMatrix(llRow, AIRDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = AIRTIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdDates.TextMatrix(llRow, AIRTIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = GAMENOINDEX) Then
                slSort = Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX))
                '6/30/12:  Allow 5 digit event #'s
                Do While Len(slSort) < 5    '4
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = LIBRARYINDEX) Then
                slSort = Trim$(grdDates.TextMatrix(llRow, VERLIBINDEX))
            Else
                slSort = UCase$(Trim$(grdDates.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdDates.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdDates.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
                slRow = Trim$(str$(llRow + 1))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdDates.TextMatrix(llRow + 1, SORTINDEX) = slSort & slStr & "|" & slRow
                'Fill-in test column
                grdDates.TextMatrix(llRow + 1, GAMENOINDEX) = grdDates.TextMatrix(llRow, GAMENOINDEX)
            Else
                slRow = Trim$(str$(llRow + 1))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdDates.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdDates.TextMatrix(llRow + 1, SORTINDEX) = slSort & slStr & "|" & slRow
                'Fill-in test column
                grdDates.TextMatrix(llRow + 1, GAMENOINDEX) = grdDates.TextMatrix(llRow, GAMENOINDEX)
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdDates, GAMENOINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

Private Function mPreemptGame(ilGameVefCode As Integer, ilGameNo As Integer, llGameDate As Long) As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llLastLogDate As Long
    Dim ilVpfIndex As Integer
    Dim ilRet As Integer
    Dim llChfCode As Long

    mPreemptGame = 0
    If (llGameDate >= lmFirstAllowedChgDate) Then
        If Not mLockPreemptVehicle(ilGameVefCode, llGameDate) Then
            mPreemptGame = 1
            Exit Function
        End If
        ilVpfIndex = gVpfFind(GameSchd, ilGameVefCode)
        gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llLastLogDate
        llChfCode = -1
        slStartDate = Format$(llGameDate, "m/d/yy")
        slEndDate = slStartDate
        slStartTime = gFormatTimeLong(0, "A", "1")
        slEndTime = gFormatTimeLong(86399, "A", "1")
        ilRet = gUnschSpots(ilGameVefCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slStartTime, slEndTime, ilGameNo)
        If Not ilRet Then
            ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
            mPreemptGame = 2
            Exit Function
        End If
        ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
    End If
End Function

Private Sub mSetControls()
    Dim ilGap As Integer
    Dim ilRow As Integer
    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    cmcDone.Top = Me.Height - cmcDone.Height - 120
    cmcCancel.Top = cmcDone.Top
    cmcSave.Top = cmcDone.Top
    cmcFormats.Top = cmcDone.Top
    cmcMultimedia.Top = cmcDone.Top
    cmcSyncGames.Top = cmcDone.Top

    cmcSave.Left = GameSchd.Width / 2 - cmcSave.Width - ilGap / 2
    cmcCancel.Left = cmcSave.Left - cmcCancel.Width - ilGap
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    cmcFormats.Left = cmcSave.Left + cmcSave.Width + ilGap
    cmcMultimedia.Left = cmcFormats.Left + cmcFormats.Width + ilGap
    cmcSyncGames.Left = cmcMultimedia.Left + cmcMultimedia.Width + ilGap

    grdSpec.Move 180, 435
    mGridSpecLayout
    mGridSpecColumnWidths
    mGridSpecColumns

    grdDates.Move grdSpec.Left, grdSpec.Top + grdSpec.Height + 120, GameSchd.Width - 2 * grdSpec.Left, cmcDone.Top - grdSpec.Top - grdSpec.Height - 240

    ''imInitNoRows = grdDates.Rows
    'DoEvents
    mGridLayout
    'DoEvents
    mGridColumnWidths
    gGrid_IntegralHeight grdDates, fgBoxGridH + 15
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
        If grdDates.RowHeight(ilRow) > 15 Then
            grdDates.Col = GAMENOINDEX
            grdDates.Row = ilRow
            grdDates.CellBackColor = LIGHTYELLOW
        End If
    Next ilRow

    ckcShowVersion.Top = grdDates.Top + grdDates.Height + 60
    ckcShowVersion.Left = grdSpec.Left

    
    cbcSeason.Top = 60
    cbcSeason.Left = grdDates.Left + grdDates.Width - cbcSeason.Width
    
    cbcSelect.Top = 60
    cbcSelect.Left = cbcSeason.Left - cbcSelect.Width - 120
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
'*  ilLoop                        slStr                         ilIndex                   *
'*  slName                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    imInNew = True
    igVpfType = 2
    VehModel.Show vbModal
    If (igVehReturn = 0) Or (igVefCodeModel = 0) Then    'Cancelled
        igVefCodeModel = 0
        mStartNew = True
        imInNew = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    ilRet = mReadRec(igVefCodeModel, -1)
    mMoveRecToCtrl
    imChgMode = False
    imGsfChg = True
    mStartNew = True
    Screen.MousePointer = vbDefault
    mSetCommands
    imInNew = False
    Exit Function
End Function

Private Sub mSetGridValues(llRow As Long, tlGsf As GSF, ilFromSync As Integer, ilCompare As Integer)
    Dim slStr As String
    Dim ilValue As Integer
    Dim ilTeam As Integer
    Dim ilLib As Integer
    Dim ilLang As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilVeh As Integer
    Dim llCol As Long
    Dim ilSubtotal As Integer
    
    For llCol = GAMENOINDEX To VERLIBINDEX Step 1
        grdDates.Col = llCol
        grdDates.CellForeColor = BLACK
    Next llCol
    slStr = Trim$(str$(tlGsf.iGameNo))
    grdDates.TextMatrix(llRow, GAMENOINDEX) = Trim$(slStr)
    grdDates.Col = GAMENOINDEX
    grdDates.CellBackColor = LIGHTYELLOW
    'Feed
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    slStr = ""
    If (ilValue And USINGFEED) = USINGFEED Then
        If tlGsf.sFeedSource = "V" Then
            slStr = smEventTitle1   '"V" '"Visiting"
        ElseIf tlGsf.sFeedSource = "H" Then
            slStr = smEventTitle2   '"H" '"Home"
        ElseIf tlGsf.sFeedSource = "N" Then
            slStr = "National" '"National"
        End If
    End If
    If ilCompare Then
        If grdDates.TextMatrix(llRow, FEEDSOURCEINDEX) <> Trim$(slStr) Then
            imGsfChg = True
        End If
    End If
    grdDates.TextMatrix(llRow, FEEDSOURCEINDEX) = Trim$(slStr)
    'Language
    slStr = ""
    If (ilValue And USINGLANG) = USINGLANG Then
        For ilLang = 0 To UBound(tmLanguageCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            slNameCode = tmLanguageCode(ilLang).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tlGsf.iLangMnfCode = Val(slCode) Then
                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                Exit For
            End If
        Next ilLang
    End If
    If ilCompare Then
        If grdDates.TextMatrix(llRow, LANGUAGEINDEX) <> Trim$(slStr) Then
            imGsfChg = True
        End If
    End If
    grdDates.TextMatrix(llRow, LANGUAGEINDEX) = Trim$(slStr)
    'Visiting Team
    slStr = ""
    For ilTeam = 0 To UBound(tmTeamCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
        slNameCode = tmTeamCode(ilTeam).sKey 'Traffic!lbcAgency.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tlGsf.iVisitMnfCode = Val(slCode) Then
            ilRet = gParseItem(slNameCode, 1, "\", slStr)
            Exit For
        End If
    Next ilTeam
    If ilCompare Then
        If grdDates.TextMatrix(llRow, VISITTEAMINDEX) <> Trim$(slStr) Then
            imGsfChg = True
        End If
    End If
    grdDates.TextMatrix(llRow, VISITTEAMINDEX) = Trim$(slStr)
    'Home Team
    slStr = ""
    For ilTeam = 0 To UBound(tmTeamCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
        slNameCode = tmTeamCode(ilTeam).sKey 'Traffic!lbcAgency.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tlGsf.iHomeMnfCode = Val(slCode) Then
            ilRet = gParseItem(slNameCode, 1, "\", slStr)
            Exit For
        End If
    Next ilTeam
    If ilCompare Then
        If grdDates.TextMatrix(llRow, HOMETEAMINDEX) <> Trim$(slStr) Then
            imGsfChg = True
        End If
    End If
    grdDates.TextMatrix(llRow, HOMETEAMINDEX) = Trim$(slStr)
    
    'Subtotal1
    slStr = ""
    If Trim$(tgSaf(0).sEventSubtotal1) <> "" Then
        For ilSubtotal = 0 To UBound(tmSubtotal1Code) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            slNameCode = tmSubtotal1Code(ilSubtotal).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tlGsf.iSubtotal1MnfCode = Val(slCode) Then
                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                Exit For
            End If
        Next ilSubtotal
        If ilCompare Then
            If grdDates.TextMatrix(llRow, SUBTOTAL1INDEX) <> Trim$(slStr) Then
                imGsfChg = True
            End If
        End If
    End If
    grdDates.TextMatrix(llRow, SUBTOTAL1INDEX) = Trim$(slStr)
    'Subtotal2
    slStr = ""
    If Trim$(tgSaf(0).sEventSubtotal2) <> "" Then
        For ilSubtotal = 0 To UBound(tmSubtotal2Code) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            slNameCode = tmSubtotal2Code(ilSubtotal).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tlGsf.iSubtotal2MnfCode = Val(slCode) Then
                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                Exit For
            End If
        Next ilSubtotal
        If ilCompare Then
            If grdDates.TextMatrix(llRow, SUBTOTAL2INDEX) <> Trim$(slStr) Then
                imGsfChg = True
            End If
        End If
    End If
    grdDates.TextMatrix(llRow, SUBTOTAL2INDEX) = Trim$(slStr)
    
    If Not ilFromSync Then
        'Library
        slStr = ""
        For ilLib = 0 To UBound(tmLibName) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            slNameCode = tmLibName(ilLib).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tlGsf.lLvfCode = Val(slCode) Then
                slStr = lbcLibrary.List(ilLib)
                Exit For
            End If
        Next ilLib
        If (slStr = "") And (UBound(tmLibName) = 1) Then
            slNameCode = tmLibName(0).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tlGsf.lLvfCode = Val(slCode)
            slStr = lbcLibrary.List(0)
        End If
        If ilCompare Then
            If grdDates.TextMatrix(llRow, VERLIBINDEX) <> Trim$(slStr) Then
                imGsfChg = True
            End If
        End If
        grdDates.TextMatrix(llRow, VERLIBINDEX) = Trim$(slStr)
        If ckcShowVersion.Value = vbUnchecked Then
            ilPos = InStr(1, slStr, "/", vbTextCompare)
            If ilPos > 0 Then
                slStr = Mid$(slStr, ilPos + 1)
            End If
        End If
        grdDates.TextMatrix(llRow, LIBRARYINDEX) = Trim$(slStr)
    End If
    'Air Date
    gUnpackDate tlGsf.iAirDate(0), tlGsf.iAirDate(1), slStr
    If gDateValue(slStr) < lmFirstAllowedChgDate Then
        If tlGsf.sGameStatus <> "P" Then
            grdDates.Row = llRow
            grdDates.Col = AIRDATEINDEX
            grdDates.CellForeColor = vbRed
            grdDates.Col = AIRTIMEINDEX
            grdDates.CellForeColor = vbRed
        End If
    End If
    If ilCompare Then
        If gDateValue(grdDates.TextMatrix(llRow, AIRDATEINDEX)) <> gDateValue(slStr) Then
            imGsfChg = True
        End If
    End If
    grdDates.TextMatrix(llRow, AIRDATEINDEX) = slStr
    'Air Time
    gUnpackTime tlGsf.iAirTime(0), tlGsf.iAirTime(1), "A", "1", slStr
    If ilCompare Then
        If gTimeToLong(grdDates.TextMatrix(llRow, AIRTIMEINDEX), False) <> gTimeToLong(slStr, False) Then
            imGsfChg = True
        End If
    End If
    grdDates.TextMatrix(llRow, AIRTIMEINDEX) = slStr
    'Air Vehicle
    slStr = ""
    If (ilValue And PREEMPTREGPROG) = PREEMPTREGPROG Then
        For ilVeh = 0 To UBound(tmAirVehicle) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            slNameCode = tmAirVehicle(ilVeh).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tlGsf.iAirVefCode = Val(slCode) Then
                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                ilRet = gParseItem(slStr, 3, "|", slStr)
                If tlGsf.sLiveLogMerge = "L" Then
                    slStr = "L: " & slStr
                ElseIf tlGsf.sLiveLogMerge = "M" Then
                    slStr = "M: " & slStr
                End If
                Exit For
            End If
        Next ilVeh
    End If
    If ilCompare Then
        If grdDates.TextMatrix(llRow, AIRVEHICLEINDEX) <> Trim$(slStr) Then
            imGsfChg = True
        End If
    End If
    grdDates.TextMatrix(llRow, AIRVEHICLEINDEX) = Trim$(slStr)
    grdDates.TextMatrix(llRow, XDSPROGCODEINDEX) = Trim$(tlGsf.sXDSProgCodeID)
    grdDates.TextMatrix(llRow, BUSINDEX) = Trim$(tlGsf.sBus)
    'Game Status
    slStr = ""
    If tlGsf.sGameStatus = "C" Then
        slStr = "C" '"Canceled"
    ElseIf tlGsf.sGameStatus = "F" Then
        slStr = "F" '"Firm"
    ElseIf tlGsf.sGameStatus = "P" Then
        slStr = "P" '"Postponed"
    ElseIf tlGsf.sGameStatus = "T" Then
        slStr = "T" '"Tentative"
    End If
    grdDates.TextMatrix(llRow, GAMESTATUSINDEX) = slStr
End Sub

Private Function mAdjustMMNtr(ilGameNo As Integer, llOrigDate As Long, llNewDate As Long) As Integer
    Dim ilRet As Integer
    Dim llOrigStdDate As Long
    Dim llNewStdDate As Long
    Dim slDate As String
    Dim llDate As Long
    Dim llRate As Long
    Dim llOrigSbfCode As Long
    Dim llNewSbfCode As Long
    Dim slInvType As String
    Dim slInvItem As String

    slDate = Format$(llOrigDate, "m/d/yy")
    slDate = gObtainEndStd(slDate)
    llOrigStdDate = gDateValue(slDate)
    slDate = Format$(llNewDate, "m/d/yy")
    slDate = gObtainEndStd(slDate)
    llNewStdDate = gDateValue(slDate)
    If llOrigStdDate = llNewStdDate Then
        mAdjustMMNtr = True
        Exit Function
    End If
    tmMsfSrchKey1.iVefCode = imVefCode
    ilRet = btrGetGreaterOrEqual(hmMsf, tmMsf, imMsfRecLen, tmMsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmMsf.iVefCode = imVefCode)
        tmMgfSrchKey1.lMsfCode = tmMsf.lCode
        tmMgfSrchKey1.iGameNo = ilGameNo
        ilRet = btrGetEqual(hmMgf, tmMgf, imMgfRecLen, tmMgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) Then
            'Compute Total Rate
            llRate = tmMgf.iNoUnits * tmMgf.lRate
            tmChfSrchKey0.lCode = tmMsf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                If tmChf.sDelete <> "Y" Then
                    llOrigSbfCode = -1
                    llNewSbfCode = -1
                    tmSbfSrchKey0.lChfCode = tmChf.lCode
                    tmSbfSrchKey0.iDate(0) = 0
                    tmSbfSrchKey0.iDate(1) = 0
                    tmSbfSrchKey0.sTranType = " "
                    ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCode = tmSbf.lChfCode)
                        If tmMsf.iIhfCode = tmSbf.iIhfCode Then
                            gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
                            If llDate = llOrigStdDate Then
                                llOrigSbfCode = tmSbf.lCode
                            End If
                            If llDate = llNewStdDate Then
                                llNewSbfCode = tmSbf.lCode
                            End If
                            If (llOrigSbfCode > 0) And (llNewSbfCode > 0) Then
                                Exit Do
                            End If
                        End If
                        ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If llOrigSbfCode > 0 Then
                        'Update SBF record
                        Do
                            tmSbfSrchKey1.lCode = llOrigSbfCode
                            ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet = BTRV_ERR_NONE Then
                                tmSbf.lGross = tmSbf.lGross - llRate
                                If tmSbf.lGross > 0 Then
                                    ilRet = btrUpdate(hmSbf, tmSbf, imSbfRecLen)
                                Else
                                    ilRet = btrDelete(hmSbf)
                                End If
                            Else
                                Exit Do
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            mAdjustMMNtr = False
                            Exit Function
                        End If
                    End If
                    If llNewSbfCode > 0 Then
                        'Update SBF Record
                        Do
                            tmSbfSrchKey1.lCode = llNewSbfCode
                            ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet = BTRV_ERR_NONE Then
                                tmSbf.lGross = tmSbf.lGross + llRate
                                ilRet = btrUpdate(hmSbf, tmSbf, imSbfRecLen)
                            Else
                                Exit Do
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            mAdjustMMNtr = False
                            Exit Function
                        End If
                    Else
                        'Create SBF Record
                        tmSbf.lCode = 0
                        tmSbf.lChfCode = tmChf.lCode
                        tmSbf.sTranType = "I"
                        tmSbf.iBillVefCode = imVefCode
                        tmSbf.iAirVefCode = imVefCode
                        gPackDateLong llNewStdDate, tmSbf.iDate(0), tmSbf.iDate(1)
                        gPackDateLong llNewStdDate, tmSbf.iPrintInvDate(0), tmSbf.iPrintInvDate(1)
                        tmSbf.sDescr = ""
                        slInvType = ""
                        slInvItem = ""
                        tmIhfSrchKey0.iCode = tmMsf.iIhfCode
                        ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            If tmIhf.iItfCode <> tmItf.iCode Then
                                tmItfSrchKey0.iCode = tmIhf.iItfCode
                                ilRet = btrGetEqual(hmItf, tmItf, imItfRecLen, tmItfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    slInvType = Trim$(tmItf.sName)
                                End If
                            Else
                                slInvType = Trim$(tmItf.sName)
                            End If
                            If tmIhf.iIifCode <> tmIif.iCode Then
                                tmIifSrchKey0.iCode = tmIhf.iIifCode
                                ilRet = btrGetEqual(hmIif, tmIif, imIifRecLen, tmIifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    slInvItem = Trim$(tmIif.sName)
                                End If
                            Else
                                slInvItem = Trim$(tmIif.sName)
                            End If
                        End If
                        tmSbf.sDescr = slInvType & "/" & slInvItem
                        tmSbf.iMnfItem = imNTRMnfCode
                        tmSbf.lGross = llRate
                        tmSbf.iNoItems = 1
                        If tmChf.iAgfCode <= 0 Then
                            tmSbf.sAgyComm = "N"
                        Else
                            tmSbf.sAgyComm = "Y"
                        End If
                        tmSbf.iCommPct = imNTRSlspComm
                        tmSbf.iTrfCode = 0
                        tmSbf.sBilled = "N"
                        tmSbf.lAcquisitionCost = 0
                        tmSbf.iIhfCode = tmMsf.iIhfCode
                        tmSbf.iLineNo = 0
                        tmSbf.iCalCarryBonus = 0
                        ilRet = btrInsert(hmSbf, tmSbf, imSbfRecLen, INDEXKEY0)
                        If ilRet <> BTRV_ERR_NONE Then
                            mAdjustMMNtr = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(hmMsf, tmMsf, imMsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    mAdjustMMNtr = True
End Function

Private Function mAddMultiMediaNTR() As Integer
    Dim ilRet As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilMnf As Integer
    Dim slSlspComm As String

    ilRet = gObtainMnfForType("I", smMnfStamp, tmNTRMNF())
    For ilMnf = LBound(tmNTRMNF) To UBound(tmNTRMNF) - 1 Step 1
        If StrComp(Trim$(tmNTRMNF(ilMnf).sName), "MultiMedia", vbTextCompare) = 0 Then
            mAddMultiMediaNTR = tmNTRMNF(ilMnf).iCode
            gPDNToStr tmNTRMNF(ilMnf).sSSComm, 4, slSlspComm
            imNTRSlspComm = gStrDecToInt(slSlspComm, 2)
            'gPDNToStr tmMnf.sRPU, 2, slStr
            'gPDNToStr tmMnf.sSSComm, 4, slStr
            Exit Function
        End If
    Next ilMnf
    gGetSyncDateTime slSyncDate, slSyncTime
    tmMnf.iCode = 0
    tmMnf.sType = "I"
    tmMnf.sName = "MultiMedia"
    gStrToPDN "0", 2, 5, tmMnf.sRPU     'Amount per Item
    tmMnf.sUnitType = ""
    tmMnf.iMerge = 0
    tmMnf.iGroupNo = 0  'Not Taxable
    tmMnf.sCodeStn = "N"    'Hard Cost
    tmMnf.sUnitsPer = ""
    gStrToPDN "0", 4, 4, tmMnf.sSSComm  'Salesperson Commission
    imNTRSlspComm = 0
    tmMnf.lCost = 0     'Acquisition Cost
    tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
    tmMnf.iAutoCode = tmMnf.iCode
    ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        imNTRMnfCode = 0
        Exit Function
    End If
    Do
        tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
        tmMnf.iAutoCode = tmMnf.iCode
        gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
        gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
        ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    mAddMultiMediaNTR = tmMnf.iCode
End Function


Private Function mLockPreemptVehicle(ilVefCode As Integer, llDate As Long) As Integer
    Dim slUserName As String
    Dim ilRet As Integer
    Dim ilVef As Integer
    
    mLockPreemptVehicle = True
    lmLockRecCode = 0
    If ilVefCode <= 0 Then
        Exit Function
    End If
    lmLockRecCode = gCreateLockRec(hmRlf, "S", "C", 65536 * ilVefCode + llDate, True, slUserName)
    If lmLockRecCode > 0 Then
        Exit Function
    End If
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        ilRet = MsgBox("Unable to perform requested task as " & slUserName & " is working on pre-empt vehicle: " & Trim$(tgMVef(ilVef).sName) & ". Press Save again", vbOKOnly + vbInformation, "Block")
    Else
        ilRet = MsgBox("Unable to perform requested task as " & slUserName & " is working on pre-empt vehicle. Press Save again", vbOKOnly + vbInformation, "Block")
    End If
    mLockPreemptVehicle = False
End Function

Private Sub mSeasonPop()
    Dim llStartDate As Long
    Dim slStartDate As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llSeasonGhfCode As Long
    Dim ilVff As Integer
    
    cbcSeason.Clear
    lmSeasonGhfCode = 0
    ReDim tmSeasonInfo(0 To 0) As SEASONINFO
    tmGhfSrchKey1.iVefCode = imVefCode
    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = imVefCode)
        gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llStartDate
        slStartDate = Trim$(str$(llStartDate))
        Do While Len(slStartDate) < 6
            slStartDate = "0" & slStartDate
        Loop
        tmSeasonInfo(UBound(tmSeasonInfo)).sKey = slStartDate
        tmSeasonInfo(UBound(tmSeasonInfo)).sSeasonName = tmGhf.sSeasonName
        tmSeasonInfo(UBound(tmSeasonInfo)).lCode = tmGhf.lCode
        ReDim Preserve tmSeasonInfo(0 To UBound(tmSeasonInfo) + 1) As SEASONINFO
        ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If UBound(tmSeasonInfo) > 1 Then
        'Sort descending
        ArraySortTyp fnAV(tmSeasonInfo(), 0), UBound(tmSeasonInfo), 1, LenB(tmSeasonInfo(0)), 0, LenB(tmSeasonInfo(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmSeasonInfo) - 1 Step 1
        cbcSeason.AddItem Trim$(tmSeasonInfo(ilLoop).sSeasonName)
        cbcSeason.ItemData(cbcSeason.NewIndex) = tmSeasonInfo(ilLoop).lCode
    Next ilLoop
    cbcSeason.AddItem "[New]", 0
    cbcSeason.ItemData(cbcSeason.NewIndex) = 0
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If tgVff(ilVff).iVefCode = imVefCode Then
            lmSeasonGhfCode = tgVff(ilVff).lSeasonGhfCode
            Exit For
        End If
    Next ilVff
    For ilLoop = 1 To cbcSeason.ListCount - 1 Step 1
        If cbcSeason.ItemData(ilLoop) = lmSeasonGhfCode Then
            cbcSeason.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    
End Sub


Private Function mSpecGridFieldsOk() As Integer
    Dim ilError As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    
    ilError = False
    
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, SEASONNAMEINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, SEASONNAMEINDEX) = "Missing"
    Else
        slStr = UCase(Trim$(grdSpec.TextMatrix(SPECROW3INDEX, SEASONNAMEINDEX)))
        For ilLoop = 0 To cbcSeason.ListCount - 1 Step 1
            If UCase(Trim(cbcSeason.List(ilLoop))) = slStr Then
                If (imNewGame) Or (tmGhf.lCode <> cbcSeason.ItemData(ilLoop)) Then
                    MsgBox "Season Name Previously Used", vbInformation, "Name Unsed"
                    ilError = True
                End If
            End If
        Next ilLoop
    End If
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX) = "Missing"
    End If
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX) = "Missing"
    End If
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, GAMENOINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, GAMENOINDEX) = "Missing"
    End If
    If ilError Then
        mSpecGridFieldsOk = False
    Else
        mSpecGridFieldsOk = True
    End If
End Function

Private Function mCheckDates() As Integer
    Dim slStr As String
    Dim ilRow As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilError As Integer
    Dim ilRet As Integer
    
    ilError = False
    llStartDate = 999999
    llEndDate = 0
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If (grdDates.TextMatrix(ilRow, GAMENOINDEX) <> "") Then
            '12/19/12: Include canceled
            'If grdDates.TextMatrix(ilRow, GAMESTATUSINDEX) <> "C" Then
                slStr = grdDates.TextMatrix(ilRow, AIRDATEINDEX)
                If (slStr <> "") And (slStr <> "Missing") Then
                    If gDateValue(slStr) < llStartDate Then
                        llStartDate = gDateValue(slStr)
                    End If
                    If gDateValue(slStr) > llEndDate Then
                        llEndDate = gDateValue(slStr)
                    End If
                End If
            'End If
        End If
    Next ilRow
    If llStartDate <> 999999 Then
        If llStartDate < gDateValue(grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX)) Then
            'slStr = Format$(llStartDate, "m/d/yy")
            'ilRet = MsgBox("Start Date of Season " & grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX) & " is greater then the Event Start Date " & slStr & ", Continue with Save?", vbQuestion + vbYesNo + vbDefaultButton2, "Date Warning")
            'If ilRet = vbNo Then
                ilError = True
            'End If
        End If
    End If
    If llEndDate <> 0 Then
        If llEndDate > gDateValue(grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX)) Then
            'slStr = Format$(llEndDate, "m/d/yy")
            'ilRet = MsgBox("End Date of Season " & grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX) & " is less then the Event End Date " & slStr & ", Continue with Save?", vbQuestion + vbYesNo + vbDefaultButton2, "Date Warning")
            'If ilRet = vbNo Then
                ilError = True
            'End If
        End If
    End If
    If ilError Then
        If (llStartDate <> 999999) And (llEndDate <> 0) Then
            sgGenMsg = "Event Dates (" & Format$(llStartDate, "m/d/yy") & "-" & Format$(llEndDate, "m/d/yy") & ") are Outside of Season Dates (" & grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX) & "-" & grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX) & ")"
        Else
            sgGenMsg = "Event Dates are Outside of Season Dates (" & grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX) & "-" & grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX) & ")"
        End If
        sgCMCTitle(0) = "&Adjust Season Dates"
        sgCMCTitle(1) = "&Cancel Save"
        sgCMCTitle(2) = ""
        sgCMCTitle(3) = ""
        igDefCMC = 0
        GenMsg.Show vbModal
        If igAnsCMC = 0 Then
            grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX) = Format$(llStartDate, "m/d/yy")
            grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX) = Format$(llEndDate, "m/d/yy")
            mCheckDates = True
        Else
            mCheckDates = False
        End If
    Else
        mCheckDates = True
    End If
End Function

Private Function mCheckOverlap() As Integer
    Dim ilRet As Integer
    Dim tlGhf As GHF
    Dim llMoSeasonStart As Long
    Dim llMoSeasonEnd As Long
    Dim llSeasonStart As Long
    Dim llSeasonEnd As Long
    Dim llTestStart As Long
    Dim llTestEnd As Long
    Dim llMoTestStart As Long
    Dim llMoTestEnd As Long
    
    llSeasonStart = gDateValue(grdSpec.TextMatrix(SPECROW3INDEX, SEASONSTARTINDEX))
    llMoSeasonStart = llSeasonStart
    Do While gWeekDayLong(llMoSeasonStart) <> 0   '0=monday
        llMoSeasonStart = llMoSeasonStart - 1
    Loop
    llSeasonEnd = gDateValue(grdSpec.TextMatrix(SPECROW3INDEX, SEASONENDINDEX))
    llMoSeasonEnd = llSeasonEnd
    Do While gWeekDayLong(llMoSeasonEnd) <> 0   '0=monday
        llMoSeasonEnd = llMoSeasonEnd - 1
    Loop
    tmGhfSrchKey1.iVefCode = imVefCode
    ilRet = btrGetEqual(hmGhf, tlGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlGhf.iVefCode = imVefCode)
        If tmGhf.lCode <> tlGhf.lCode Then
            gUnpackDateLong tlGhf.iSeasonStartDate(0), tlGhf.iSeasonStartDate(1), llTestStart
            gUnpackDateLong tlGhf.iSeasonEndDate(0), tlGhf.iSeasonEndDate(1), llTestEnd
            If (llTestEnd >= llSeasonStart) And (llTestStart <= llSeasonEnd) Then
                MsgBox "Season Definitions overlap, Save canceled", vbCritical + vbOKOnly, "Overlap Dates"
                mCheckOverlap = False
                Exit Function
            End If
            llMoTestStart = llTestStart
            Do While gWeekDayLong(llMoTestStart) <> 0   '0=monday
                llMoTestStart = llMoTestStart - 1
            Loop
            llMoTestEnd = llTestEnd
            Do While gWeekDayLong(llMoTestEnd) <> 0   '0=monday
                llMoTestEnd = llMoTestEnd - 1
            Loop
            If (llMoSeasonStart = llMoTestEnd) Then
                MsgBox "Season Definitions overlap because current season start date and previous season end date in same week, Save canceled", vbCritical + vbOKOnly, "Overlap Dates"
                mCheckOverlap = False
                Exit Function
            End If
            If (llMoSeasonEnd = llMoTestStart) Then
                MsgBox "Season Definitions overlap because previous season start date and current season end date in same week, Save canceled", vbCritical + vbOKOnly, "Overlap Dates"
                mCheckOverlap = False
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hmGhf, tlGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    mCheckOverlap = True
End Function

Private Function mSportsByWeek(llChfCode As Long, ilLineNo As Integer) As String
    Dim ilRet As Integer
    
    mSportsByWeek = "N"
    tmClfSrchKey0.lChfCode = llChfCode
    tmClfSrchKey0.iLine = ilLineNo
    tmClfSrchKey0.iCntRevNo = 32000 ' Plug with very high number
    tmClfSrchKey0.iPropVer = 32000 ' Plug with very high number
    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) Then
        mSportsByWeek = tmClf.sSportsByWeek
    End If


End Function

Private Function mAddEcfRecord(llGsfCode As Long, slGameStatus As String, llFromDate As Long, llToDate As Long) As Integer
    Dim ilRet As Integer
    Dim slMsg As String
    
    On Error GoTo mAddEcfRecordErr
    mAddEcfRecord = True
    If slGameStatus <> "C" Then
        If llFromDate = llToDate Then
            Exit Function
        End If
    End If
    tmLstSrchKey3.iLogVefCode = imVefCode
    tmLstSrchKey3.lGsfCode = llGsfCode
    ilRet = btrGetEqual(hmLst, tmLst, imLstRecLen, tmLstSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
    If (ilRet <> BTRV_ERR_NONE) Then
        Exit Function
    End If
    tmEcf.lCode = 0
    tmEcf.lGsfCode = llGsfCode
    gPackDateLong llFromDate, tmEcf.iFromDate(0), tmEcf.iFromDate(1)
    gPackDateLong llToDate, tmEcf.iToDate(0), tmEcf.iToDate(1)
    tmEcf.sWebCleared = "N"
    tmEcf.sMarketronCleared = "N"
    tmEcf.sUnivisionCleared = "N"
    gPackDate Format$(gNow(), "m/d/yy"), tmEcf.iEnteredDate(0), tmEcf.iEnteredDate(1)
    gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tmEcf.iEnteredTime(0), tmEcf.iEnteredTime(1)
    tmEcf.iUrfCode = tgUrf(0).iCode
    tmEcf.sUnused = ""
    ilRet = btrInsert(hmEcf, tmEcf, imEcfRecLen, INDEXKEY0)
    slMsg = "mSaveRec (btrInsert:Event Change Date)"
    gBtrvErrorMsg ilRet, slMsg, GameSchd
    On Error GoTo 0
    Exit Function
mAddEcfRecordErr:
    On Error GoTo 0
    mAddEcfRecord = False
    Exit Function
End Function

Private Sub mSubtotalPop(ilSubtotalNo As Integer)

'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    If (ilSubtotalNo < 0) Or (ilSubtotalNo = 1) Then
        ilRet = gPopMnfPlusFieldsBox(GameSchd, lbcSubtotal(0), tmSubtotal1Code(), smSubtotal1CodeTag, "1")
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mSubtotalPopErr
            gCPErrorMsg ilRet, "mSubtotalPop (gPopMnfPlusFieldsBox)", GameSchd
            On Error GoTo 0
            lbcSubtotal(0).AddItem "[New]", 0
        End If
    End If
    If (ilSubtotalNo < 0) Or (ilSubtotalNo = 2) Then
        ilRet = gPopMnfPlusFieldsBox(GameSchd, lbcSubtotal(1), tmSubtotal2Code(), smSubtotal2CodeTag, "2")
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mSubtotalPopErr
            gCPErrorMsg ilRet, "mSubtotalPop (gPopMnfPlusFieldsBox)", GameSchd
            On Error GoTo 0
            lbcSubtotal(1).AddItem "[New]", 0
        End If
    End If
    Exit Sub
mSubtotalPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Function mSubtotalBranch(ilIndex) As Integer
'
'   ilRet = mSubtotalBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropdown, lbcSubtotal(ilIndex), imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mSubtotalBranch = False
        Exit Function
    End If
    If ilIndex = 0 Then
        sgMnfCallType = "1"
    Else
        sgMnfCallType = "2"
    End If
    igMNmCallSource = CALLSOURCEGAME
    If edcDropdown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "GameSchd^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "GameScd^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "EName^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "EName^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'EName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'PEvent.Enabled = True
    'edcLinkSrceDoneMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mSubtotalBranch = True
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
        If ilIndex = 0 Then
            lbcSubtotal(ilIndex).Clear
            smSubtotal1CodeTag = ""
            mSubtotalPop ilIndex + 1
        Else
            lbcSubtotal(ilIndex).Clear
            smSubtotal2CodeTag = ""
            mSubtotalPop ilIndex + 1
        End If
        If imTerminate Then
            mSubtotalBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcSubtotal(ilIndex)
        If gLastFound(lbcSubtotal(ilIndex)) > 0 Then
            imChgMode = True
            lbcSubtotal(ilIndex).ListIndex = gLastFound(lbcSubtotal(ilIndex))
            edcDropdown.Text = lbcSubtotal(ilIndex).List(lbcSubtotal(ilIndex).ListIndex)
            imChgMode = False
            mSubtotalBranch = False
        Else
            imChgMode = True
            lbcSubtotal(ilIndex).ListIndex = 0
            edcDropdown.Text = lbcSubtotal(ilIndex).List(0)
            imChgMode = False
            edcDropdown.SetFocus
            sgMNmName = ""
            Exit Function
        End If
        sgMNmName = ""
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mEnableBox
        Exit Function
    End If
    Exit Function
End Function
