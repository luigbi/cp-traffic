VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form NetworkSplit 
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
   Begin VB.PictureBox pbcPrinting 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   2520
      ScaleHeight     =   1200
      ScaleWidth      =   3825
      TabIndex        =   44
      Top             =   2100
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox lbcCategory 
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
      ItemData        =   "NetworkSplit.frx":0000
      Left            =   4950
      List            =   "NetworkSplit.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   855
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox pbcNewTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   9180
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   165
      Width           =   60
   End
   Begin VB.Frame frcCategory 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3930
      Index           =   1
      Left            =   285
      TabIndex        =   28
      Top             =   870
      Visible         =   0   'False
      Width           =   8835
      Begin VB.CommandButton cmcClearStationSelection 
         Appearance      =   0  'Flat
         Caption         =   "C&lear Selection"
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
         Left            =   6885
         TabIndex        =   43
         Top             =   3555
         Width           =   1875
      End
      Begin VB.PictureBox pbcStationSelection 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         ScaleHeight     =   330
         ScaleWidth      =   3975
         TabIndex        =   39
         Top             =   3555
         Width           =   3975
         Begin VB.OptionButton rbcStationSelection 
            Caption         =   "Range"
            Height          =   195
            Index           =   1
            Left            =   2835
            TabIndex        =   42
            Top             =   0
            Width           =   885
         End
         Begin VB.OptionButton rbcStationSelection 
            Caption         =   "Single"
            Height          =   195
            Index           =   0
            Left            =   1785
            TabIndex        =   41
            Top             =   0
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.Label lacStationSelection 
            Caption         =   "Station selection"
            Height          =   210
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Width           =   1515
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStation 
         Height          =   3435
         Left            =   60
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   60
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   6059
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         GridColor       =   -2147483635
         GridColorFixed  =   -2147483635
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   8
      End
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4155
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Frame frcCategory 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4035
      Index           =   0
      Left            =   525
      TabIndex        =   29
      Top             =   1245
      Visible         =   0   'False
      Width           =   8730
      Begin VB.PictureBox pbcUpMove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3975
         Picture         =   "NetworkSplit.frx":0018
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2040
         Width           =   180
      End
      Begin VB.PictureBox pbcDnMove 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4530
         Picture         =   "NetworkSplit.frx":00F2
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1005
         Width           =   180
      End
      Begin VB.ListBox lbcTo 
         Height          =   3375
         Left            =   5145
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   3450
      End
      Begin VB.ListBox lbcFrom 
         Height          =   3375
         Left            =   0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   360
         Width           =   3450
      End
      Begin VB.CommandButton cmcMoveFrom 
         Appearance      =   0  'Flat
         Caption         =   "  Mo&ve"
         Height          =   300
         Left            =   3930
         TabIndex        =   35
         Top             =   1980
         Width           =   810
      End
      Begin VB.CommandButton cmcMoveTo 
         Appearance      =   0  'Flat
         Caption         =   "M&ove   "
         Height          =   300
         Left            =   3930
         TabIndex        =   33
         Top             =   945
         Width           =   810
      End
      Begin VB.Label lacTo 
         Alignment       =   2  'Center
         Height          =   180
         Left            =   5625
         TabIndex        =   37
         Top             =   30
         Width           =   2490
      End
      Begin VB.Label lacFrom 
         Alignment       =   2  'Center
         Height          =   180
         Left            =   525
         TabIndex        =   36
         Top             =   30
         Width           =   2490
      End
   End
   Begin VB.PictureBox pbcInclExcl 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2925
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   10
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
      Left            =   1695
      Picture         =   "NetworkSplit.frx":01CC
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   660
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox ckcIncludeDormant 
      Caption         =   "Include Dormant Names"
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
      Left            =   2880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   2070
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8895
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   27
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
   Begin VB.PictureBox pbcStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1815
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   810
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
      TabIndex        =   13
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
      Left            =   5295
      TabIndex        =   2
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
      Left            =   5490
      TabIndex        =   25
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
      Picture         =   "NetworkSplit.frx":02C6
      TabIndex        =   15
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   630
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
      TabIndex        =   14
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
         Picture         =   "NetworkSplit.frx":03C0
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   16
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
            Picture         =   "NetworkSplit.frx":107E
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
      TabIndex        =   17
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
         Picture         =   "NetworkSplit.frx":1388
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   21
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
            TabIndex        =   22
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
         TabIndex        =   18
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   90
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5640
      Width           =   75
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
      Left            =   4155
      TabIndex        =   24
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
      Left            =   2790
      TabIndex        =   23
      Top             =   5355
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpec 
      Height          =   450
      Left            =   210
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   465
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   794
      _Version        =   393216
      Rows            =   5
      Cols            =   22
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
      _Band(0).Cols   =   22
   End
   Begin VB.Label plcScreen 
      Caption         =   "Region Definition"
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
Attribute VB_Name = "NetworkSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of NetworkSplit.frm on Wed 6/17/09 @ 12:
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmRafSrchKey2                 tmArtt                        imVefCode                 *
'*  imVpfIndex                    lmLLD                         lmFirstAllowedChgDate     *
'*  imLastColSorted               imLastSort                    lmTopRow                  *
'*  imInitNoRows                                                                          *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mPopOwners                                                                            *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: NetworkSplit.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

'Region Area
Dim tmRaf As RAF            'RAF record image
Dim tmRafSrchKey As LONGKEY0  'RAF key record image
Dim hmRaf As Integer        'RAF Handle
Dim imRafRecLen As Integer      'RAF record length

'Split Entity
Dim tmSef() As SEF            'SEF record image
Dim tmSefSrchKey As LONGKEY0  'SEF key record image
Dim tmSefSrchKey1 As SEFKEY1  'SEF key record image
Dim hmSef As Integer        'SEF Handle
Dim imSefRecLen As Integer      'SEF record length

'Agreement
Dim tmAtt As ATT                'ATT record image
Dim tmAttSrchKey2 As INTKEY0     'ATT key 1 image
Dim imAttRecLen As Integer      'ATT record length
Dim hmAtt As Integer            'Agreement file handle

'ARTT- Get Owner names

'Market Names
Dim tmMkt As MKT            'MKT record image

'Stations
Dim tmSHTT As SHTT
Dim imStationPop As Integer

'Format Names
Dim tmFmt As FMT            'MKT record image

'State
Dim tmSnt As SNT

'Zone
Dim tmTzt As TZT

Dim tmRegionCode() As SORTCODE
Dim smRegionCodeTag As String

'Contract line
Dim hmClf As Integer        'Contract line file handle
Dim tmClf As CLF            'CLF record image
Dim imClfRecLen As Integer
Dim tmClfSrchKey4 As CLFKEY4

Dim hmSdf As Integer        'Spot detail file handle
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey5 As LONGKEY0
Dim imSdfRecLen As Integer  'SDF record length

Dim hmSmf As Integer
Dim tmSmf As SMF

Dim hmSsf As Integer
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

Dim hmStf As Integer
Dim hmGsf As Integer
Dim hmGhf As Integer

Dim hmSxf As Integer

'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imInNewTab As Integer
Dim imSettingValue As Integer
Dim imLbcArrowSetting As Integer
Dim imBypassFocus As Integer
Dim imSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Dim smNowDate As String
Dim lmNowDate As Long
Dim smInclExcl As String
Dim smStatus As String
Dim imSpecChg As Integer
Dim imCountChg As Integer
Dim smShowOnProp As String
Dim smShowOnOrder As String
Dim smShowOnInv As String

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmSpecEnableRow As Long
Dim lmSpecEnableCol As Long
Dim imSpecCtrlVisible As Integer

Dim imLastStationColSorted As Integer
Dim imLastStationSort As Integer
Dim lmStationRangeRow As Long

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
Const CATEGORYINDEX = 2
Const NAMEINDEX = 4
Const ABBRINDEX = 6
Const INCLEXCLINDEX = 8
Const AUDPCTINDEX = 10
Const STATUSINDEX = 12
Const SHOWONINDEX = 14
Const SHOWONPROPINDEX = 16
Const SHOWONORDERINDEX = 18
Const SHOWONINVINDEX = 20

Const STATIONINDEX = 0
Const MARKETINDEX = 1
Const STATEINDEX = 2
'Const ZIPCODEINDEX = 3
'Const OWNERINDEX = 4
Const ZONEINDEX = 3
Const FORMATINDEX = 4   '5
Const SHTTCODEINDEX = 5 '6
Const SORTINDEX = 6 '7
Const SELECTEDINDEX = 7 '8






Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box

    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    gSetMousePointer grdSpec, grdStation, vbHourglass
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    mClearCtrlFields
    If (ilRet = 0) And (cbcSelect.ListIndex > 1) Then
        ilIndex = cbcSelect.ListIndex
        If Not mReadRec(ilIndex, SETFORREADONLY, 0) Then
            GoTo cbcSelectErr
        End If
    Else
        If ilRet = 1 Then
            If cbcSelect.ListCount > 0 Then
                cbcSelect.ListIndex = 0
            Else
                cbcSelect.ListIndex = -1
            End If
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
        mMoveSEFRecToCtrl
        cmcMoveTo.Enabled = False
        For ilLoop = 0 To UBound(lgNewRafCntr) - 1 Step 1
            If lgNewRafCntr(ilLoop) = tmRaf.lCode Then
                cmcMoveTo.Enabled = True
                Exit For
            End If
        Next ilLoop
    Else
        imSelectedIndex = cbcSelect.ListIndex
        If (slStr <> "[New]") And (slStr <> "[Model]") Then
            grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = slStr
        End If
        cmcMoveTo.Enabled = True
    End If
    gSetMousePointer grdSpec, grdStation, vbDefault
    Screen.MousePointer = vbDefault
    imChgMode = False
    mSetCommands
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    gSetMousePointer grdSpec, grdStation, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcSelect_GotFocus()
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSpec, grdStation, vbHourglass
    lmSpecEnableRow = -1
    mSpecSetShow
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
    If imSelectedIndex <= 1 Then
        mClearCtrlFields
    End If
    gCtrlGotFocus cbcSelect
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdStation, vbDefault
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

Private Sub ckcIncludeDormant_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                         ilPos                         slStr                     *
'*                                                                                        *
'******************************************************************************************



End Sub

Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    sgDoneMsg = ""
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSpecSetShow
    gCtrlGotFocus cmcCancel
End Sub

Private Sub cmcClearStationSelection_Click()
    Dim llRow As Long
    Dim llCol As Long
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSpec, grdStation, vbHourglass
    For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
        If grdStation.TextMatrix(llRow, STATIONINDEX) <> "" Then
            grdStation.Row = llRow
            grdStation.TextMatrix(llRow, SELECTEDINDEX) = "F"
            For llCol = STATIONINDEX To FORMATINDEX Step 1
                grdStation.Col = llCol
                grdStation.CellBackColor = vbWhite
            Next llCol
        End If
    Next llRow
    lmStationRangeRow = -1
    gSetMousePointer grdSpec, grdStation, vbDefault
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcClearStationSelection_GotFocus()
    mSpecSetShow
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
    Dim slStr As String

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If imSpecChg Or imCountChg Then
        slStr = grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)
        If imSelectedIndex > 1 Then
            slMess = "Save Changes to " & slStr
        Else
            slMess = "Add " & slStr
        End If
        ilRet = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If ilRet = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
            igSplitChgd = True
        End If
    End If
    sgDoneMsg = grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSpecSetShow
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case CATEGORYINDEX
                    lbcCategory.Visible = Not lbcCategory.Visible
            End Select
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub cmcMoveFrom_Click()
    Dim ilLoop As Integer
    For ilLoop = lbcTo.ListCount - 1 To 0 Step -1
        If lbcTo.Selected(ilLoop) Then
            lbcFrom.AddItem lbcTo.List(ilLoop)
            lbcFrom.ItemData(lbcFrom.NewIndex) = lbcTo.ItemData(ilLoop)
            lbcTo.RemoveItem ilLoop
        End If
    Next ilLoop
    imCountChg = True
    mSetCommands
End Sub

Private Sub cmcMoveTo_Click()
    Dim ilLoop As Integer
    For ilLoop = lbcFrom.ListCount - 1 To 0 Step -1
        If lbcFrom.Selected(ilLoop) Then
            lbcTo.AddItem lbcFrom.List(ilLoop)
            lbcTo.ItemData(lbcTo.NewIndex) = lbcFrom.ItemData(ilLoop)
            lbcFrom.RemoveItem ilLoop
        End If
    Next ilLoop
    imCountChg = True
    mSetCommands
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ilRet = mSaveRec()
    If Not ilRet Then
        Exit Sub
    End If
    smRegionCodeTag = ""
    mPopulate igAdfCode
    mSetCommands
    igSplitChgd = True
End Sub

Private Sub cmcSave_GotFocus()
    mSpecSetShow
End Sub

Private Sub edcDropDown_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                                                   *
'******************************************************************************************


    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case CATEGORYINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcDropDown, lbcCategory, imBSMode, imComboBoxIndex
            End Select
    End Select
    imLbcArrowSetting = False
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub edcDropDown_DblClick()
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case CATEGORYINDEX
            End Select
    End Select
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case CATEGORYINDEX
            End Select
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFound                       ilLoop                                                  *
'******************************************************************************************

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case CATEGORYINDEX
            End Select
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                                                                                *
'******************************************************************************************

    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                    Case CATEGORYINDEX
                        gProcessArrowKey Shift, KeyCode, lbcCategory, imLbcArrowSetting
                End Select
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                    Case CATEGORYINDEX
                End Select
        End Select
    End If
End Sub

Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                    Case CATEGORYINDEX
                End Select
        End Select
        imDoubleClickName = False
    End If
End Sub

Private Sub edcSpec_Change()
    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case NAMEINDEX
                Case AUDPCTINDEX
            End Select
    End Select
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub edcSpec_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSpec_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False

End Sub

Private Sub edcSpec_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String

    Select Case lmSpecEnableRow
        Case SPECROW3INDEX
            Select Case lmSpecEnableCol
                Case NAMEINDEX
                Case AUDPCTINDEX
                    ilPos = InStr(edcSpec.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcSpec.Text, ".")    'Disallow multi-decimal points
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
                    slStr = edcSpec.Text
                    slStr = Left$(slStr, edcSpec.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSpec.SelStart - edcSpec.SelLength)
                    If gCompNumberStr(slStr, "100.00") > 0 Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If

            End Select
    End Select

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
    If (igWinStatus(SPLITNETSLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        grdSpec.Enabled = False
        grdStation.Enabled = False
        imUpdateAllowed = False
    Else
        grdSpec.Enabled = True
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        grdStation.Enabled = True
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        If lmEnableCol > 0 Then
        '    mEnableBox
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub


Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    Erase tmRegionCode
    
    btrDestroy hmRaf
    btrDestroy hmSef
    btrDestroy hmSdf
    btrDestroy hmSsf
    btrDestroy hmSmf
    btrDestroy hmStf
    btrDestroy hmGsf
    btrDestroy hmGhf
    btrDestroy hmSxf
    btrDestroy hmClf
    btrDestroy hmAtt
    
    Set NetworkSplit = Nothing   'Remove data segment

End Sub

Private Sub grdSpec_EnterCell()
    mSpecSetShow
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

Private Sub grdStation_Click()
    Dim llRow As Long
    Dim llCol As Long
    Dim llLoop As Long

    llRow = grdStation.Row
    llCol = grdStation.Col
    If rbcStationSelection(1).Value Then
        If lmStationRangeRow >= grdStation.FixedRows Then
            If lmStationRangeRow < llRow Then
                For llLoop = lmStationRangeRow To llRow Step 1
                    If grdStation.TextMatrix(llLoop, STATIONINDEX) <> "" Then
                        grdStation.Row = llLoop
                        grdStation.TextMatrix(llLoop, SELECTEDINDEX) = "T"
                        For llCol = STATIONINDEX To FORMATINDEX Step 1
                            grdStation.Col = llCol
                            grdStation.CellBackColor = GRAY
                            'grdStation.CellForeColor = vbWhite
                        Next llCol
                    End If
                Next llLoop
            Else
                For llLoop = llRow To lmStationRangeRow Step 1
                    If grdStation.TextMatrix(llLoop, STATIONINDEX) <> "" Then
                        grdStation.Row = llLoop
                        grdStation.TextMatrix(llLoop, SELECTEDINDEX) = "T"
                        For llCol = STATIONINDEX To FORMATINDEX Step 1
                            grdStation.Col = llCol
                            grdStation.CellBackColor = GRAY
                            'grdStation.CellForeColor = vbWhite
                        Next llCol
                    End If
                Next llLoop
            End If
            lmStationRangeRow = -1
        Else
            lmStationRangeRow = llRow
            grdStation.Row = llRow
            grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T"
            For llCol = STATIONINDEX To FORMATINDEX Step 1
                grdStation.Col = llCol
                grdStation.CellBackColor = GRAY
                'grdStation.CellForeColor = vbWhite
            Next llCol
        End If
    Else
        lmStationRangeRow = -1
        If llRow >= grdStation.FixedRows Then
            If grdStation.TextMatrix(llRow, STATIONINDEX) <> "" Then
                If grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T" Then
                    grdStation.Row = llRow
                    grdStation.TextMatrix(llRow, SELECTEDINDEX) = "F"
                    For llCol = STATIONINDEX To FORMATINDEX Step 1
                        grdStation.Col = llCol
                        grdStation.CellBackColor = vbWhite
                        'grdStation.CellForeColor = vbBlack
                    Next llCol
                Else
                    grdStation.Row = llRow
                    grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T"
                    For llCol = STATIONINDEX To FORMATINDEX Step 1
                        grdStation.Col = llCol
                        grdStation.CellBackColor = GRAY
                        'grdStation.CellForeColor = vbWhite
                    Next llCol
                End If
            End If
        End If
    End If
    imCountChg = True

    mSetCommands
End Sub

Private Sub grdStation_EnterCell()
    mSpecSetShow
End Sub

Private Sub grdStation_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Y < grdStation.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        gSetMousePointer grdSpec, grdStation, vbHourglass
        grdStation.Col = grdStation.MouseCol
        mStationSortCol grdStation.Col
        mSetCommands
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
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
    slStr = edcDropDown.Text
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
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slName                                                                                *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    gSetMousePointer grdSpec, grdStation, vbHourglass
    imFirstActivate = True
    imTerminate = False
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
    imSpecCtrlVisible = False
    imSpecChg = False
    imCountChg = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmStationRangeRow = -1
    imInNewTab = False
    imStationPop = False
    mInitBox

    If Not gRecLengthOk("Raf.btr", Len(tmRaf)) Then
        imTerminate = True
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hmRaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", CopyRegn
    On Error GoTo 0
    imRafRecLen = Len(tmRaf)

    ReDim tmSef(0 To 0) As SEF
    If Not gRecLengthOk("Sef.btr", Len(tmSef(0))) Then
        imTerminate = True
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hmSef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSef, "", sgDBPath & "Sef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sef.Btr)", CopyRegn
    On Error GoTo 0
    imSefRecLen = Len(tmSef(0))

    hmAtt = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAtt, "", sgDBPath & "Att.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Att.Mkd)", CopyRegn
    On Error GoTo 0
    imAttRecLen = Len(tmAtt)

    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", CopyRegn
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", CopyRegn
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)

    hmSmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", CopyRegn
    On Error GoTo 0

    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", CopyRegn
    On Error GoTo 0

    hmStf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Stf.Btr)", CopyRegn
    On Error GoTo 0

    hmGsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", CopyRegn
    On Error GoTo 0
    
    hmGhf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", CopyRegn
    On Error GoTo 0
    
    hmSxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sxf.Btr)", CopyRegn
    On Error GoTo 0
    'If Not gRecLengthOk("Artt.mkd", Len(tmArtt)) Then
    '    imTerminate = True
    '    gSetMousePointer grdSpec, grdStation, vbDefault
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'End If
    'mPopOwners


    If Not gRecLengthOk("Mkt.mkd", Len(tmMkt)) Then
        imTerminate = True
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mPopMarkets

    If Not gRecLengthOk("Snt.mkd", Len(tmSnt)) Then
        imTerminate = True
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mPopStates

    If Not gRecLengthOk("Tzt.mkd", Len(tmTzt)) Then
        imTerminate = True
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mPopTimeZones

    If (UBound(tgStates) <= LBound(tgStates)) Or (UBound(tgTimeZones) <= LBound(tgTimeZones)) Then
        MsgBox "Exit Traffic System and Sign-On to the Affiliate System to initialize tables required by Copy Splits", vbCritical + vbOKOnly, "Split Copy"
    End If


    If Not gRecLengthOk("Shtt.mkd", Len(tmSHTT)) Then
        imTerminate = True
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If Not gRecLengthOk("Fmt.mkd", Len(tmFmt)) Then
        imTerminate = True
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mPopFormats

    'Populate only if category selected
    'mPopStations


    mPopCategory

    mPopulate igAdfCode

    If Trim$(sgSplitNetworkName) <> "" Then
        gFindMatch sgSplitNetworkName, 0, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        Else
            cbcSelect.ListIndex = 0
            cbcSelect.Text = sgSplitNetworkName
        End If
    Else
        cbcSelect.ListIndex = 0
    End If
    'mXFerRecToCtrl
    NetworkSplit.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone NetworkSplit
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdStation, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdStation, vbDefault
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
'*  flTextHeight                  ilCol                         llRet                     *
'*                                                                                        *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRow As Integer
    'flTextHeight = pbcDates.TextHeight("1") - 35


    grdSpec.Move 180, 435
    frcCategory(0).Move 180, 1080
    lbcFrom.Move 0, 360, frcCategory(0).Width / 2 - cmcMoveFrom.Width, frcCategory(1).Height - 360
    lacFrom.Move lbcFrom.Left, 30, lbcFrom.Width
    lbcTo.Move frcCategory(0).Width - lbcFrom.Width, 360, lbcFrom.Width, lbcFrom.Height
    lacTo.Move lbcTo.Left, 30, lbcTo.Width
    frcCategory(1).Move 180, 1080, grdSpec.Width
    grdStation.Move 0, 0, frcCategory(1).Width, frcCategory(1).Height
    mGridSpecLayout
    mGridSpecColumnWidths
    mGridSpecColumns

    mGridStationLayout
    mGridStationColumnWidths
    mGridStationColumns
    ilRow = grdStation.FixedRows
    Do
        If ilRow + 1 > grdStation.Rows Then
            grdStation.AddItem ""
        End If
        grdStation.RowHeight(ilRow) = fgBoxGridH + 15
        ilRow = ilRow + 1
    Loop While grdStation.RowIsVisible(ilRow - 1)
    gGrid_IntegralHeight grdStation, CInt(fgBoxGridH + 30) ' + 15
    grdStation.Height = grdStation.Height - 30
    pbcStationSelection.Top = grdStation.Top + grdStation.Height + 60
    cmcClearStationSelection.Top = pbcStationSelection.Top
    frcCategory(1).Height = grdStation.Height + pbcStationSelection.Height + 60 '+ fgBevelX + fgBevelY

    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop

    pbcPrinting.Left = (NetworkSplit.Width - pbcPrinting.Width) \ 2
    pbcPrinting.Top = (NetworkSplit.Height - pbcPrinting.Height) \ 2
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
    gSetMousePointer grdSpec, grdStation, vbDefault
    igManUnload = YES
    Unload NetworkSplit
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mSpecFieldsOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

'
'   iRet = mSpecFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim ilError As Integer

    ilError = False
    grdSpec.Row = SPECROW3INDEX
    grdSpec.Col = NAMEINDEX
    grdSpec.CellForeColor = vbBlack
    grdSpec.Col = ABBRINDEX
    grdSpec.CellForeColor = vbBlack
    grdSpec.Col = CATEGORYINDEX
    grdSpec.CellForeColor = vbBlack
    grdSpec.Col = INCLEXCLINDEX
    grdSpec.CellForeColor = vbBlack
    grdSpec.Col = AUDPCTINDEX
    grdSpec.CellForeColor = vbBlack
    grdSpec.Col = STATUSINDEX
    grdSpec.CellForeColor = vbBlack
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = NAMEINDEX
        grdSpec.CellForeColor = vbMagenta
    End If
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, ABBRINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, ABBRINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = ABBRINDEX
        grdSpec.CellForeColor = vbMagenta
    End If
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, CATEGORYINDEX)
    gFindMatch slStr, 0, lbcCategory
    If gLastFound(lbcCategory) < 0 Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, CATEGORYINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = CATEGORYINDEX
        grdSpec.CellForeColor = vbMagenta
    End If
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, INCLEXCLINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, INCLEXCLINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = INCLEXCLINDEX
        grdSpec.CellForeColor = vbMagenta
    End If
    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, AUDPCTINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, AUDPCTINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = AUDPCTINDEX
        grdSpec.CellForeColor = vbMagenta
    End If

    If Trim$(grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX)) = "" Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = STATUSINDEX
        grdSpec.CellForeColor = vbMagenta
    End If
    If ilError Then
        mSpecFieldsOk = False
    Else
        mSpecFieldsOk = True
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
'
'   mXFerRecToCtrl
'   Where:
'
    Dim slStr As String

    grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = Trim$(tmRaf.sName)
    grdSpec.TextMatrix(SPECROW3INDEX, ABBRINDEX) = Trim$(tmRaf.sAbbr)
    If tmRaf.sInclExcl = "I" Then
        grdSpec.TextMatrix(SPECROW3INDEX, INCLEXCLINDEX) = "Include"
        lacFrom.Caption = "Exclude List"
        lacTo.Caption = "Include List"
    ElseIf tmRaf.sInclExcl = "E" Then
        grdSpec.TextMatrix(SPECROW3INDEX, INCLEXCLINDEX) = "Exclude"
        lacFrom.Caption = "Include List"
        lacTo.Caption = "Exclude List"
    End If
    Select Case Trim$(tmRaf.sCategory)
        Case "M"
            slStr = "Market"
        Case "N"
            slStr = "State Name"
        'Case "Z"
        '    slStr = "Zip Code"
        'Case "O"
        '    slStr = "Owner"
        Case "F"
            slStr = "Format"
        Case "S"
            slStr = "Station"
            mPopStations False
        Case "T"
            slStr = "Time Zone"
        Case Else
            slStr = ""
    End Select
    grdSpec.TextMatrix(SPECROW3INDEX, CATEGORYINDEX) = slStr
    If tmRaf.iAudPct <= 0 Then
        tmRaf.iAudPct = 10000
    End If
    grdSpec.TextMatrix(SPECROW3INDEX, AUDPCTINDEX) = gIntToStrDec(tmRaf.iAudPct, 2)

    If tmRaf.sState = "A" Then
        grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
    ElseIf tmRaf.sState = "D" Then
        grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Dormant"
    End If
    smShowOnProp = "No"
    If tmRaf.sShowNoProposal = "Y" Then
        smShowOnProp = "Yes"
    End If
    smShowOnOrder = "No"
    If tmRaf.sShowOnOrder = "Y" Then
        smShowOnOrder = "Yes"
    End If
    smShowOnInv = "No"
    If tmRaf.sShowOnInvoice = "Y" Then
        smShowOnInv = "Yes"
    End If
    grdSpec.TextMatrix(SPECROW3INDEX, SHOWONPROPINDEX) = smShowOnProp
    grdSpec.TextMatrix(SPECROW3INDEX, SHOWONORDERINDEX) = smShowOnOrder
    grdSpec.TextMatrix(SPECROW3INDEX, SHOWONINVINDEX) = smShowOnInv
End Sub



Private Sub lbcCategory_Click()
    gProcessLbcClick lbcCategory, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcCategory_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Sub lbcFrom_GotFocus()
    mSpecSetShow
End Sub


Private Sub lbcTo_GotFocus()
    mSpecSetShow
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
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDropDown.SetFocus
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
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcInclExcl_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("I")) Or (KeyAscii = Asc("i")) Then
        smInclExcl = "Include"
        pbcInclExcl_Paint
    ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
        smInclExcl = "Exclude"
        pbcInclExcl_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smInclExcl = "Include" Then
            smInclExcl = "Exclude"
            pbcInclExcl_Paint
        ElseIf smInclExcl = "Exclude" Then
            smInclExcl = "Include"
            pbcInclExcl_Paint
        End If
    End If
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub pbcInclExcl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smInclExcl = "Include" Then
        smInclExcl = "Exclude"
        pbcInclExcl_Paint
    Else
        smInclExcl = "Include"
        pbcInclExcl_Paint
    End If
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub pbcInclExcl_Paint()
    pbcInclExcl.Cls
    pbcInclExcl.CurrentX = fgBoxInsetX
    pbcInclExcl.CurrentY = 0 'fgBoxInsetY
    pbcInclExcl.Print smInclExcl
End Sub

Private Sub pbcNewTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilLoop                        slNameCode                *
'*  slCode                                                                                *
'******************************************************************************************


    If imInNewTab Then
        Exit Sub
    End If
    If imUpdateAllowed = False Then
        cmcCancel.SetFocus
        Exit Sub
    End If

    If imSelectedIndex > 1 Then
        pbcSpecSTab.SetFocus
        Exit Sub
    End If
    If imSelectedIndex = 0 Then
        grdSpec.TextMatrix(SPECROW3INDEX, INCLEXCLINDEX) = "Include"
        grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
        pbcSpecSTab.SetFocus
        Exit Sub
    End If
    If UBound(tmSef) = LBound(tmSef) Then
        imInNewTab = True
        igSplitType = 1 'Network flag
        igIncludeDormantSplits = False
        If ckcIncludeDormant.Value = vbChecked Then
            igIncludeDormantSplits = True
        End If
        SplitModel.Show vbModal
        DoEvents
        If (igSplitModelReturn = 1) And (lgSplitModelCodeRaf) > 0 Then
            Screen.MousePointer = vbHourglass
            gSetMousePointer grdSpec, grdStation, vbHourglass
            If mReadRec(imSelectedIndex, SETFORREADONLY, lgSplitModelCodeRaf) Then
                mMoveRecToCtrl
                mMoveSEFRecToCtrl
                grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = ""
            End If
            grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
            gSetMousePointer grdSpec, grdStation, vbDefault
            Screen.MousePointer = vbDefault
        Else
            grdSpec.TextMatrix(SPECROW3INDEX, INCLEXCLINDEX) = "Include"
            grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = "Active"
        End If
    End If
    imInNewTab = False
    pbcSpecSTab.SetFocus

End Sub

Private Sub pbcPrinting_Paint()
    pbcPrinting.CurrentX = (pbcPrinting.Width - pbcPrinting.TextWidth("Resolving Station Conflicts....")) / 2
    pbcPrinting.CurrentY = (pbcPrinting.Height - pbcPrinting.TextHeight("Resolving Station Conflicts....")) / 2 - 30
    pbcPrinting.Print "Resolving Station Conflicts...."
End Sub

Private Sub pbcSpecSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSpecSTab.hwnd Then
        Exit Sub
    End If
    If imSpecCtrlVisible Then
        Do
            ilNext = False
            Select Case grdSpec.Row
                Case SPECROW3INDEX
                    Select Case grdSpec.Col
                        Case CATEGORYINDEX
                            mSpecSetShow
                            cmcDone.SetFocus
                            Exit Sub
                        Case Else
                            If grdSpec.Col >= CATEGORYINDEX + 2 Then
                                grdSpec.Col = grdSpec.Col - 2
                            Else
                                mSpecSetShow
                                cmcDone.SetFocus
                                Exit Sub
                            End If
                    End Select
            End Select
            If mSpecColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSpecSetShow
    Else
        grdSpec.Row = SPECROW3INDEX '+1 to bypass title
        grdSpec.Col = grdSpec.FixedCols
        Do
            If mSpecColOk() Then
                Exit Do
            Else
                grdSpec.Col = grdSpec.Col + 2
            End If
        Loop
    End If
    mSpecEnableBox
End Sub

Private Sub pbcSpecTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSpecTab.hwnd Then
        Exit Sub
    End If
    If imSpecCtrlVisible Then
        Do
            ilNext = False
            Select Case grdSpec.Row
                Case SPECROW3INDEX
                    Select Case grdSpec.Col
                        Case STATUSINDEX
                            grdSpec.Col = grdSpec.Col + 2
                        Case SHOWONINVINDEX
                            mSpecSetShow
                            cmcDone.SetFocus
                            Exit Sub
                        Case Else
                            grdSpec.Col = grdSpec.Col + 2
                    End Select
            End Select
            If mSpecColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSpecSetShow
    Else
        grdSpec.Row = grdSpec.Rows - 2
        grdSpec.Col = SHOWONINVINDEX
        Do
            If mSpecColOk() Then
                Exit Do
            Else
                grdSpec.Col = grdSpec.Col - 2
            End If
        Loop
    End If
    mSpecEnableBox
End Sub





Private Sub pbcStatus_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
        smStatus = "Active"
        pbcStatus_Paint
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        smStatus = "Dormant"
        pbcStatus_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smStatus = "Active" Then
            smStatus = "Dormant"
            pbcStatus_Paint
        ElseIf smStatus = "Dormant" Then
            smStatus = "Active"
            pbcStatus_Paint
        End If
    End If
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub pbcStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smStatus = "Active" Then
        smStatus = "Dormant"
        pbcStatus_Paint
    Else
        smStatus = "Active"
        pbcStatus_Paint
    End If
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub pbcStatus_Paint()
    pbcStatus.Cls
    pbcStatus.CurrentX = fgBoxInsetX
    pbcStatus.CurrentY = 0 'fgBoxInsetY
    pbcStatus.Print smStatus
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
                    'Select Case lmEnableCol
                    '    Case AIRTIMEINDEX
                    '        imBypassFocus = True    'Don't change select text
                    '        edcDropDown.SetFocus
                    '        SendKeys slKey
                    'End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub

Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If lmSpecEnableCol = SHOWONPROPINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            smShowOnProp = "Yes"
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            smShowOnProp = "No"
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If smShowOnProp = "Yes" Then
                smShowOnProp = "No"
                pbcYN_Paint
            ElseIf smShowOnProp = "No" Then
                smShowOnProp = "Yes"
                pbcYN_Paint
            End If
        End If
    End If
    If lmSpecEnableCol = SHOWONORDERINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            smShowOnOrder = "Yes"
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            smShowOnOrder = "No"
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If smShowOnOrder = "Yes" Then
                smShowOnOrder = "No"
                pbcYN_Paint
            ElseIf smShowOnOrder = "No" Then
                smShowOnOrder = "Yes"
                pbcYN_Paint
            End If
        End If
    End If
    If lmSpecEnableCol = SHOWONINVINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            smShowOnInv = "Yes"
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            smShowOnInv = "No"
            pbcYN_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If smShowOnInv = "Yes" Then
                smShowOnInv = "No"
                pbcYN_Paint
            ElseIf smShowOnInv = "No" Then
                smShowOnInv = "Yes"
                pbcYN_Paint
            End If
        End If
    End If
   grdSpec.CellForeColor = vbBlack
End Sub

Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lmSpecEnableCol = SHOWONPROPINDEX Then
        If smShowOnProp = "Yes" Then
            smShowOnProp = "No"
            pbcYN_Paint
        Else
            smShowOnProp = "Yes"
            pbcYN_Paint
        End If
    End If
    If lmSpecEnableCol = SHOWONORDERINDEX Then
        If smShowOnOrder = "Yes" Then
            smShowOnOrder = "No"
            pbcYN_Paint
        Else
            smShowOnOrder = "Yes"
            pbcYN_Paint
        End If
    End If
    If lmSpecEnableCol = SHOWONINVINDEX Then
        If smShowOnInv = "Yes" Then
            smShowOnInv = "No"
            pbcYN_Paint
        Else
            smShowOnInv = "Yes"
            pbcYN_Paint
        End If
    End If
   grdSpec.CellForeColor = vbBlack
End Sub

Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If lmSpecEnableCol = SHOWONPROPINDEX Then
        pbcYN.Print smShowOnProp
    End If
    If lmSpecEnableCol = SHOWONORDERINDEX Then
        pbcYN.Print smShowOnOrder
    End If
    If lmSpecEnableCol = SHOWONINVINDEX Then
        pbcYN.Print smShowOnInv
    End If

End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
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
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer, llModelRafCode As Long) As Integer
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
    Dim tlSef As SEF

    If ilSelectIndex > 1 Then
        slNameCode = tmRegionCode(ilSelectIndex - 2).sKey    'lbcCopyRegnCode.List(ilSelectIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mReadRecErr
        gCPErrorMsg ilRet, "mReadRecErr (gParseItem field 2)", NetworkSplit
        On Error GoTo 0
        tmRafSrchKey.lCode = CLng(slCode)
    Else
        tmRafSrchKey.lCode = llModelRafCode
    End If
    ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
    On Error GoTo mReadRecErr
    gBtrvErrorMsg ilRet, "mReadRecErr (btrGetEqual: Region)", NetworkSplit
    On Error GoTo 0
    ReDim tmSef(0 To 0) As SEF
    tmSefSrchKey1.lRafCode = tmRaf.lCode
    tmSefSrchKey1.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hmSef, tlSef, imSefRecLen, tmSefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlSef.lRafCode = tmRaf.lCode)
        tmSef(UBound(tmSef)) = tlSef
        ReDim Preserve tmSef(0 To UBound(tmSef) + 1) As SEF
        ilRet = btrGetNext(hmSef, tlSef, imSefRecLen, BTRV_LOCK_NONE, ilForUpdate)
    Loop
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function


Private Sub mClearCtrlFields()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRow                         ilCol                                                   *
'******************************************************************************************



    ReDim tmSef(0 To 0) As SEF
    grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, ABBRINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, CATEGORYINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, INCLEXCLINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, AUDPCTINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, SHOWONPROPINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, SHOWONORDERINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, SHOWONINVINDEX) = ""
    lbcTo.Clear
    lbcFrom.Clear
    'For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
    '    grdStation.Row = llRow
    '    For ilCol = STATIONINDEX To FORMATINDEX Step 1
    '        grdStation.Col = ilCol
    '        grdStation.CellBackColor = vbWhite
    '    Next ilCol
    '    grdStation.TextMatrix(llRow, SELECTEDINDEX) = "F"
    'Next llRow
    If imStationPop Then
        imStationPop = False
        mPopStations False
    End If
    imSpecChg = False
    imCountChg = False
    lmEnableRow = -1
    lmEnableCol = -1
    lmSpecEnableRow = -1
    lmSpecEnableCol = -1
End Sub



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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

'
'   mSetCommands
'   Where:
'

    'Update button set if all mandatory fields have data and any field altered
    If imSpecChg Or imCountChg Then
        cbcSelect.Enabled = False
        ckcIncludeDormant.Enabled = False
    Else
        cbcSelect.Enabled = True
        ckcIncludeDormant.Enabled = True
    End If
    If imSpecChg Or imCountChg Then  'At least one event added
        If imUpdateAllowed Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
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
                Case NAMEINDEX 'Name
                    edcSpec.MaxLength = 80
                    edcSpec.Text = grdSpec.Text
                Case ABBRINDEX 'Name
                    edcSpec.MaxLength = 5
                    edcSpec.Text = grdSpec.Text
                Case CATEGORYINDEX
                    lbcCategory.Height = gListBoxHeight(lbcCategory.ListCount, 10)
                    edcDropDown.MaxLength = 10
                    imChgMode = True
                    slStr = grdSpec.Text
                    gFindMatch slStr, 0, lbcCategory
                    If gLastFound(lbcCategory) >= 0 Then
                        lbcCategory.ListIndex = gLastFound(lbcCategory)
                        edcDropDown.Text = lbcCategory.List(lbcCategory.ListIndex)
                    Else
                        If lbcCategory.ListCount >= 1 Then
                            lbcCategory.ListIndex = 0
                            edcDropDown.Text = lbcCategory.List(lbcCategory.ListIndex)
                        Else
                            edcDropDown.Text = ""
                        End If
                    End If
                    imChgMode = False
                Case INCLEXCLINDEX
                    smInclExcl = grdSpec.Text
                    If (smInclExcl = "") Or (smInclExcl = "Missing") Then
                        smInclExcl = "Include"
                    End If
                Case AUDPCTINDEX
                    edcSpec.MaxLength = 6
                    edcSpec.Text = grdSpec.Text

                Case STATUSINDEX
                    smStatus = grdSpec.Text
                    If (smStatus = "") Or (smStatus = "Missing") Then
                        smStatus = "Active"
                    End If
                Case SHOWONPROPINDEX
                    smShowOnProp = grdSpec.Text
                    If (smShowOnProp = "") Or (smShowOnProp = "Missing") Then
                        smShowOnProp = "No"
                    End If
                Case SHOWONORDERINDEX
                    smShowOnOrder = grdSpec.Text
                    If (smShowOnOrder = "") Or (smShowOnOrder = "Missing") Then
                        smShowOnOrder = "No"
                    End If
                Case SHOWONINVINDEX
                    smShowOnInv = grdSpec.Text
                    If (smShowOnInv = "") Or (smShowOnInv = "Missing") Then
                        smShowOnInv = "No"
                    End If
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilNoGames                     ilOrigUpper                   ilLoop                    *
'*  llRow                         ilIndex                       ilCol                     *
'*                                                                                        *
'******************************************************************************************

    Dim slStr As String
    Dim ilCatChg As Integer

    If (lmSpecEnableRow >= grdSpec.FixedRows) And (lmSpecEnableRow < grdSpec.Rows) Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                    Case NAMEINDEX
                        slStr = edcSpec.Text
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> slStr Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = slStr
                    Case ABBRINDEX
                        slStr = edcSpec.Text
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> slStr Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = slStr
                    Case CATEGORYINDEX
                        slStr = edcDropDown.Text
                        ilCatChg = False
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> slStr Then
                            imSpecChg = True
                            ilCatChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = slStr
                        If StrComp(slStr, "Station", vbTextCompare) <> 0 Then
                            If ilCatChg Then
                                mPopFrom slStr
                            End If
                            frcCategory(0).Visible = True
                            frcCategory(1).Visible = False
                        Else
                            mPopStations True
                            frcCategory(1).Visible = True
                            frcCategory(0).Visible = False
                        End If
                    Case INCLEXCLINDEX
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> smInclExcl Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smInclExcl
                        If StrComp(smInclExcl, "Include", vbTextCompare) = 0 Then
                            lacFrom.Caption = "Exclude List"
                            lacTo.Caption = "Include List"
                        ElseIf StrComp(smInclExcl, "Exclude", vbTextCompare) = 0 Then
                            lacFrom.Caption = "Include List"
                            lacTo.Caption = "Exclude List"
                        End If
                    Case AUDPCTINDEX
                        slStr = edcSpec.Text
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> slStr Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = slStr

                    Case STATUSINDEX
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> smStatus Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smStatus
                    Case SHOWONPROPINDEX
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> smShowOnProp Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smShowOnProp
                    Case SHOWONORDERINDEX
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> smShowOnOrder Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smShowOnOrder
                    Case SHOWONINVINDEX
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> smShowOnInv Then
                            imSpecChg = True
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smShowOnInv
                End Select
        End Select
    End If
    lmSpecEnableRow = -1
    lmSpecEnableCol = -1
    imSpecCtrlVisible = False
    edcSpec.Visible = False
    edcDropDown.Visible = False
    cmcDropDown.Visible = False
    lbcCategory.Visible = False
    pbcInclExcl.Visible = False
    pbcStatus.Visible = False
    pbcYN.Visible = False
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
                Case NAMEINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    edcSpec.Visible = True
                    edcSpec.SetFocus
                Case ABBRINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    edcSpec.Visible = True
                    edcSpec.SetFocus
                Case CATEGORYINDEX
                    edcDropDown.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top, cmcDropDown.Width, cmcDropDown.Height
                    lbcCategory.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height, edcDropDown.Width + edcDropDown.Width
                    edcDropDown.Visible = True
                    cmcDropDown.Visible = True
                    edcDropDown.SetFocus
                Case INCLEXCLINDEX
                    pbcInclExcl.Move grdSpec.Left + llColPos + 45, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 45, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    pbcInclExcl_Paint
                    pbcInclExcl.Visible = True
                    pbcInclExcl.SetFocus
                Case AUDPCTINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, (3 * grdSpec.ColWidth(grdSpec.Col)) / 2 - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    edcSpec.Visible = True
                    edcSpec.SetFocus

                Case STATUSINDEX
                    pbcStatus.Move grdSpec.Left + llColPos + 45, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 45, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    pbcStatus_Paint
                    pbcStatus.Visible = True
                    pbcStatus.SetFocus
                Case SHOWONPROPINDEX
                    pbcYN.Move grdSpec.Left + llColPos + 45, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 45, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    pbcYN_Paint
                    pbcYN.Visible = True
                    pbcYN.SetFocus
                Case SHOWONORDERINDEX
                    pbcYN.Move grdSpec.Left + llColPos + 45, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 45, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    pbcYN_Paint
                    pbcYN.Visible = True
                    pbcYN.SetFocus
                Case SHOWONINVINDEX
                    pbcYN.Move grdSpec.Left + llColPos + 45, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 45, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                    pbcYN_Paint
                    pbcYN.Visible = True
                    pbcYN.SetFocus
            End Select
    End Select
End Sub



Private Sub mGridSpecLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilCol = 0 To grdSpec.Cols - 1 Step 1
        grdSpec.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
    grdSpec.RowHeight(0) = 15
    grdSpec.RowHeight(1) = 15
    grdSpec.RowHeight(2) = 150
    grdSpec.RowHeight(3) = fgBoxGridH
    grdSpec.RowHeight(4) = 15
    grdSpec.ColWidth(0) = 15
    grdSpec.ColWidth(1) = 15
    grdSpec.ColWidth(3) = 15
    grdSpec.ColWidth(5) = 15
    grdSpec.ColWidth(7) = 15
    grdSpec.ColWidth(9) = 15
    grdSpec.ColWidth(11) = 15
    grdSpec.ColWidth(13) = 15
    grdSpec.ColWidth(15) = 15
    grdSpec.ColWidth(17) = 15
    grdSpec.ColWidth(19) = 15

    'Horizontal
    For ilCol = 1 To grdSpec.Cols - 1 Step 1
        grdSpec.Row = 1
        grdSpec.Col = ilCol
        grdSpec.CellBackColor = vbBlue
    Next ilCol
    For ilRow = grdSpec.FixedRows + 2 To grdSpec.Rows - 1 Step 3
        For ilCol = 1 To grdSpec.Cols - 1 Step 1
            grdSpec.Row = ilRow
            grdSpec.Col = ilCol
            grdSpec.CellBackColor = vbBlue
        Next ilCol
    Next ilRow
    'Vertical Line
    For ilRow = 1 To grdSpec.Rows - 1 Step 1
        grdSpec.Row = ilRow
        grdSpec.Col = 1
        grdSpec.CellBackColor = vbBlue
    Next ilRow
    For ilCol = grdSpec.FixedCols + 1 To grdSpec.Cols - 1 Step 2
        For ilRow = 1 To grdSpec.Rows - 1 Step 1
            grdSpec.Row = ilRow
            grdSpec.Col = ilCol
            grdSpec.CellBackColor = vbBlue
        Next ilRow
    Next ilCol
End Sub




Private Sub mGridSpecColumns()
    grdSpec.Row = SPECROW3INDEX - 1
    grdSpec.Col = NAMEINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, NAMEINDEX) = "Name"
    grdSpec.Col = ABBRINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, ABBRINDEX) = "Abbr"
    grdSpec.Col = CATEGORYINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, CATEGORYINDEX) = "Category"
    grdSpec.Col = INCLEXCLINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, INCLEXCLINDEX) = "Include/Exclude"
    grdSpec.Col = AUDPCTINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, AUDPCTINDEX) = "Aud %"
    grdSpec.Col = STATUSINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, STATUSINDEX) = "Status"
    grdSpec.Col = SHOWONINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.CellBackColor = LIGHTYELLOW
    grdSpec.TextMatrix(SPECROW3INDEX - 1, SHOWONINDEX) = "Show"
    grdSpec.Row = SPECROW3INDEX
    grdSpec.Col = SHOWONINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.CellBackColor = LIGHTYELLOW
    grdSpec.TextMatrix(SPECROW3INDEX, SHOWONINDEX) = "Stations on:"
    grdSpec.Row = SPECROW3INDEX - 1
    grdSpec.Col = SHOWONPROPINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, SHOWONPROPINDEX) = "Proposal"
    grdSpec.Col = SHOWONORDERINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, SHOWONORDERINDEX) = "Order"
    grdSpec.Col = SHOWONINVINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(SPECROW3INDEX - 1, SHOWONINVINDEX) = "Invoice"

End Sub

Private Sub mGridSpecColumnWidths()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llMinWidth                    ilColInc                      ilLoop                    *
'*                                                                                        *
'******************************************************************************************

    Dim llWidth As Long
    Dim ilCol As Integer

    grdSpec.ColWidth(CATEGORYINDEX) = 0.115 * grdSpec.Width
    grdSpec.ColWidth(ABBRINDEX) = 0.08 * grdSpec.Width
    grdSpec.ColWidth(INCLEXCLINDEX) = 0.12 * grdSpec.Width
    grdSpec.ColWidth(STATUSINDEX) = 0.09 * grdSpec.Width
    grdSpec.ColWidth(AUDPCTINDEX) = 0.07 * grdSpec.Width
    grdSpec.ColWidth(SHOWONINDEX) = 0.1 * grdSpec.Width
    grdSpec.ColWidth(SHOWONPROPINDEX) = 0.07 * grdSpec.Width
    grdSpec.ColWidth(SHOWONORDERINDEX) = 0.05 * grdSpec.Width
    grdSpec.ColWidth(SHOWONINVINDEX) = 0.065 * grdSpec.Width
    llWidth = fgPanelAdj
    For ilCol = 0 To grdSpec.Cols - 1 Step 1
        If ilCol <> NAMEINDEX Then
            llWidth = llWidth + grdSpec.ColWidth(ilCol)
        End If
    Next ilCol
    grdSpec.ColWidth(NAMEINDEX) = grdSpec.Width - llWidth - 90
    llWidth = llWidth + grdSpec.ColWidth(NAMEINDEX)
    llWidth = grdSpec.Width - llWidth
    If llWidth >= 15 Then
        Do
            For ilCol = grdSpec.FixedCols To grdSpec.Cols - 1 Step 1
                If grdSpec.ColWidth(ilCol) > 15 Then
                    If ilCol = NAMEINDEX Then
                        grdSpec.ColWidth(ilCol) = grdSpec.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub

Private Sub rbcStationSelection_Click(Index As Integer)
    lmStationRangeRow = -1
End Sub

Private Sub rbcStationSelection_GotFocus(Index As Integer)
    mSpecSetShow
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    'Select Case lmEnableCol
    '    Case LANGUAGEINDEX
    '        imLbcArrowSetting = False
    '        gProcessLbcClick lbcLanguage, edcDropDown, imChgMode, imLbcArrowSetting
    '    Case VISITTEAMINDEX
    '        imLbcArrowSetting = False
    '        gProcessLbcClick lbcTeam, edcDropDown, imChgMode, imLbcArrowSetting
    '    Case HOMETEAMINDEX
    '        imLbcArrowSetting = False
    '        gProcessLbcClick lbcTeam, edcDropDown, imChgMode, imLbcArrowSetting
    'End Select
End Sub

Private Function mSaveRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFound                                                                               *
'******************************************************************************************

    Dim ilError As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim ilSef As Integer
    Dim ilCount As Integer
    Dim llRow As Long
    Dim ilCheckForConflicts As Integer
    Dim tlSef As SEF

    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSpec, grdStation, vbHourglass
    If mSpecFieldsOk() = False Then
        ilError = True
    End If
    If ilError Then
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        MsgBox "One or more fields not defined", vbOKOnly + vbExclamation, "Save"
        Beep
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        mSaveRec = False
        Exit Function
    End If
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, CATEGORYINDEX)
    If StrComp(slStr, "Station", vbTextCompare) <> 0 Then
        If lbcTo.ListCount <= 0 Then
            gSetMousePointer grdSpec, grdStation, vbDefault
            Screen.MousePointer = vbDefault
            Beep
            MsgBox "No Category Items in the Include/Exclude List", vbOKOnly + vbExclamation, "Save"
            mSaveRec = False
            Exit Function
        End If
    Else
        ilCount = 0
        For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
            If grdStation.TextMatrix(llRow, STATIONINDEX) <> "" Then
                If grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T" Then
                    ilCount = ilCount + 1
                End If
            End If
        Next llRow
        If ilCount <= 0 Then
            gSetMousePointer grdSpec, grdStation, vbDefault
            Screen.MousePointer = vbDefault
            Beep
            MsgBox "No Stations selected", vbOKOnly + vbExclamation, "Save"
            mSaveRec = False
            Exit Function
        End If
    End If
    ilCheckForConflicts = False
    If (imSelectedIndex > 1) Then
        ilRet = mCheckForAdditions()
        If ilRet = -1 Then
            mSaveRec = False
            Exit Function
        End If
        If ilRet <> 0 Then
            ilCheckForConflicts = True
        End If
    End If
    Do  'Loop until record updated or added
        If imSelectedIndex > 1 Then
            'Reread record in so lastest is obtained
            If Not mReadRec(imSelectedIndex, SETFORWRITE, 0) Then
                gSetMousePointer grdSpec, grdStation, vbDefault
                Screen.MousePointer = vbDefault
                MsgBox "Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save"
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec
        If imSelectedIndex <= 1 Then 'New selected
            tmRaf.lCode = 0
            tmRaf.iAdfCode = 0
            tmRaf.sAssigned = "N"
            slStr = Format$(gNow(), "m/d/yy")
            gPackDate slStr, tmRaf.iDateEntrd(0), tmRaf.iDateEntrd(1)
            slStr = ""
            gPackDate slStr, tmRaf.iDateDormant(0), tmRaf.iDateDormant(1)
            tmRaf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            ilRet = btrInsert(hmRaf, tmRaf, imRafRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Region RAF)"
            lgNewRafCntr(UBound(lgNewRafCntr)) = tmRaf.lCode
            ReDim Preserve lgNewRafCntr(0 To UBound(lgNewRafCntr) + 1) As Long
        Else 'Old record-Update
            tmRaf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            ilRet = btrUpdate(hmRaf, tmRaf, imRafRecLen)
            slMsg = "mSaveRec (btrUpdate: Region RAF)"
        End If
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, NetworkSplit
        On Error GoTo 0
        'Remove SEF records
        If imSelectedIndex > 1 Then
            For ilSef = 0 To UBound(tmSef) - 1 Step 1
                tmSefSrchKey.lCode = tmSef(ilSef).lCode
                ilRet = btrGetEqual(hmSef, tlSef, imSefRecLen, tmSefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmSef)
                End If
            Next ilSef
        End If
        mMoveSefCtrlToRec
        slMsg = "mSaveRec (btrUpdate: Region SEF)"
        For ilSef = 0 To UBound(tmSef) - 1 Step 1
            ilRet = btrInsert(hmSef, tmSef(ilSef), imSefRecLen, INDEXKEY0)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, NetworkSplit
            On Error GoTo 0
        Next ilSef
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, NetworkSplit
    On Error GoTo 0
    If ilCheckForConflicts Then
        'Determine if in conflict with any of region.  If so, then check spots
        pbcPrinting.Visible = True
        ilRet = mUnschdAndSchd()
        pbcPrinting.Visible = False
    End If
    imSpecChg = False
    imCountChg = False
    mSaveRec = True
    gSetMousePointer grdSpec, grdStation, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    gSetMousePointer grdSpec, grdStation, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser regions    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mPopulate(ilAdfCode As Integer)
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer
    Dim ilIncludeDormant As Integer

    'Repopulate if required- if sales source changed by another user while in this screen
    'imPopReqd = False
    If ckcIncludeDormant.Value = vbChecked Then
        ilIncludeDormant = True
    Else
        ilIncludeDormant = False
    End If
    ilRet = gPopRegionBox(NetworkSplit, -1, "N", ilIncludeDormant, cbcSelect, tmRegionCode(), smRegionCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopRegionBox)", CopyRegn
        On Error GoTo 0
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
        cbcSelect.AddItem "[Model]", 1  'Force as first item on list
        'mPopReqd = True
    End If
    If cbcSelect.ListIndex <> 0 Then
        cbcSelect.ListIndex = 0
    Else
        cbcSelect_Change
    End If
    imCountChg = False
    imSpecChg = False
    frcCategory(1).Visible = False
    frcCategory(0).Visible = False
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Private Sub mPopCategory()
    lbcCategory.Clear
    lbcCategory.AddItem "Format", 0
    lbcCategory.ItemData(lbcCategory.NewIndex) = 4
    lbcCategory.AddItem "Market", 1
    lbcCategory.ItemData(lbcCategory.NewIndex) = 0
    lbcCategory.AddItem "State Name", 2
    lbcCategory.ItemData(lbcCategory.NewIndex) = 1
    'lbcCategory.AddItem "Zip Code", 3
    'lbcCategory.ItemData(lbcCategory.NewIndex) = 2
    'lbcCategory.AddItem "Owner", 4
    'lbcCategory.ItemData(lbcCategory.NewIndex) = 3
    lbcCategory.AddItem "Station", 3
    lbcCategory.ItemData(lbcCategory.NewIndex) = 5
    lbcCategory.AddItem "Time Zone", 4
    lbcCategory.ItemData(lbcCategory.NewIndex) = 6

End Sub



Private Sub mPopStations(ilSetMouse As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilOwner                       llNowDate                                               *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilShtt As Integer
    Dim ilMkt As Integer
    Dim ilFormat As Integer
    Dim ilVef As Integer
    Dim llDropDate As Long
    Dim llOffAir As Long
    Dim ilAddStation As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim llDate As Long

    If imStationPop Then
        Exit Sub
    End If
    If ilSetMouse Then
        Screen.MousePointer = vbHourglass
        gSetMousePointer grdSpec, grdStation, vbHourglass
    End If
    grdStation.Redraw = False
    llRow = grdStation.FixedRows
    ilRet = gObtainStations()
    For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
        If (UBound(igSplitLnVefCode) > LBound(igSplitLnVefCode)) Then
            'Only include stations that have agreements for ordered lines
            ilAddStation = False
            tmAttSrchKey2.iCode = tgStations(ilShtt).iCode
            ilRet = btrGetEqual(hmAtt, tmAtt, imAttRecLen, tmAttSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While (ilRet = BTRV_ERR_NONE) And (tmAtt.iShfCode = tgStations(ilShtt).iCode)
                For ilVef = LBound(igSplitLnVefCode) To UBound(igSplitLnVefCode) - 1 Step 1
                    If tmAtt.iVefCode = igSplitLnVefCode(ilVef) Then
                        gUnpackDateLong tmAtt.iDropDate(0), tmAtt.iDropDate(1), llDropDate
                        gUnpackDateLong tmAtt.iOffAir(0), tmAtt.iOffAir(1), llOffAir
                        If llDropDate < llOffAir Then
                            llDate = llDropDate
                        Else
                            llDate = llOffAir
                        End If
                        If (lmNowDate <= llDate) Then
                            ilAddStation = True
                            Exit Do
                        End If
                    End If
                Next ilVef
                ilRet = btrGetNext(hmAtt, tmAtt, imAttRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Else
            ilAddStation = True
        End If
        If ilAddStation Then
            If llRow >= grdStation.Rows Then
                grdStation.AddItem ""
                grdStation.RowHeight(llRow) = fgBoxGridH + 15
            End If


            grdStation.TextMatrix(llRow, STATIONINDEX) = Trim$(tgStations(ilShtt).sCallLetters)
            slStr = ""
            For ilMkt = LBound(tgMarkets) To UBound(tgMarkets) - 1 Step 1
                If tgMarkets(ilMkt).iCode = tgStations(ilShtt).iMktCode Then
                    slStr = Trim$(tgMarkets(ilMkt).sName)
                    Exit For
                End If
            Next ilMkt
            grdStation.TextMatrix(llRow, MARKETINDEX) = slStr
            grdStation.TextMatrix(llRow, STATEINDEX) = Trim$(tgStations(ilShtt).sState)
            grdStation.TextMatrix(llRow, ZONEINDEX) = Trim$(tgStations(ilShtt).sTimeZone)
            'grdStation.TextMatrix(llRow, ZIPCODEINDEX) = Trim$(tgStations(ilShtt).sZip)
            'slStr = ""
            'For ilOwner = LBound(tgOwners) To UBound(tgOwners) - 1 Step 1
            '    If tgOwners(ilOwner).iCode = tgStations(ilShtt).iOwnerArttCode Then
            '        slStr = Trim$(tgOwners(ilOwner).sLastName)
            '        Exit For
            '    End If
            'Next ilOwner
            'grdStation.TextMatrix(llRow, OWNERINDEX) = slStr
            slStr = ""
            For ilFormat = LBound(tgFormats) To UBound(tgFormats) - 1 Step 1
                If tgFormats(ilFormat).iCode = tgStations(ilShtt).iFmtCode Then
                    slStr = Trim$(tgFormats(ilFormat).sName)
                    Exit For
                End If
            Next ilFormat
            grdStation.TextMatrix(llRow, FORMATINDEX) = slStr

            'Code
            grdStation.TextMatrix(llRow, SHTTCODEINDEX) = Trim$(str$(tgStations(ilShtt).iCode))
            grdStation.TextMatrix(llRow, SORTINDEX) = ""
            grdStation.TextMatrix(llRow, SELECTEDINDEX) = "F"
            llRow = llRow + 1
        End If
    Next ilShtt
    imStationPop = True
    imLastStationColSorted = -1
    mStationSortCol 0
    grdStation.Row = 0
    grdStation.Col = SHTTCODEINDEX
    grdStation.Redraw = True
    If ilSetMouse Then
        Screen.MousePointer = vbDefault
        gSetMousePointer grdSpec, grdStation, vbDefault
    End If
End Sub

Private Sub mPopFormats()
    Dim ilRet As Integer

    ilRet = gObtainFormats()
End Sub

Private Sub mPopMarkets()
    Dim ilRet As Integer

    ilRet = gObtainMarkets()
End Sub



Private Sub mPopFrom(slCategory As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilOwner                       ilShtt                        slState                   *
'*  slZipCode                     slTimeZone                                              *
'******************************************************************************************

    Dim ilMkt As Integer
    Dim ilFmt As Integer
    Dim ilTzt As Integer
    Dim ilSnt As Integer

    If StrComp(slCategory, "Station", vbTextCompare) = 0 Then
        Exit Sub
    End If
    lbcFrom.Clear
    lbcTo.Clear
    If StrComp(slCategory, "Market", vbTextCompare) = 0 Then
        For ilMkt = LBound(tgMarkets) To UBound(tgMarkets) - 1 Step 1
            lbcFrom.AddItem Trim$(tgMarkets(ilMkt).sName)
            lbcFrom.ItemData(lbcFrom.NewIndex) = tgMarkets(ilMkt).iCode
        Next ilMkt
    ElseIf StrComp(slCategory, "State Name", vbTextCompare) = 0 Then
        'For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
        '    'State
        '    slState = Trim$(tgStations(ilShtt).sState)
        '    If slState <> "" Then
        '        gFindMatch slState, 0, lbcFrom
        '        If gLastFound(lbcFrom) < 0 Then
        '            lbcFrom.AddItem slState
        '            lbcFrom.ItemData(lbcFrom.NewIndex) = tgStations(ilShtt).iCode
        '        End If
        '    End If
        'Next ilShtt
        For ilSnt = LBound(tgStates) To UBound(tgStates) - 1 Step 1
            'lbcFrom.AddItem Trim$(tgStates(ilSnt).sPostalName) & ", " & Trim$(tgStates(ilSnt).sName)
            lbcFrom.AddItem Trim$(tgStates(ilSnt).sPostalName) & " (" & Trim$(tgStates(ilSnt).sName) & ")"
            lbcFrom.ItemData(lbcFrom.NewIndex) = ilSnt  'tgStates(ilSnt).iCode
        Next ilSnt
    'ElseIf StrComp(slCategory, "Zip Code", vbTextCompare) = 0 Then
    '    For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
    '        'Zip Code
    '        slZipCode = Trim$(tgStations(ilShtt).sZip)
    '        If slZipCode <> "" Then
    '            gFindMatch slZipCode, 0, lbcFrom
    '            If gLastFound(lbcFrom) < 0 Then
    '                lbcFrom.AddItem slZipCode
    '                lbcFrom.ItemData(lbcFrom.NewIndex) = tgStations(ilShtt).iCode
    '            End If
    '        End If
    '    Next ilShtt
    'ElseIf StrComp(slCategory, "Owner", vbTextCompare) = 0 Then
    '    For ilOwner = LBound(tgOwners) To UBound(tgOwners) - 1 Step 1
    '        lbcFrom.AddItem Trim$(tgOwners(ilOwner).sLastName)
    '        lbcFrom.ItemData(lbcFrom.NewIndex) = tgOwners(ilOwner).iCode
    '    Next ilOwner
    ElseIf StrComp(slCategory, "Format", vbTextCompare) = 0 Then
        For ilFmt = LBound(tgFormats) To UBound(tgFormats) - 1 Step 1
            lbcFrom.AddItem Trim$(tgFormats(ilFmt).sName)
            lbcFrom.ItemData(lbcFrom.NewIndex) = tgFormats(ilFmt).iCode
        Next ilFmt
    ElseIf StrComp(slCategory, "Time Zone", vbTextCompare) = 0 Then
        For ilTzt = LBound(tgTimeZones) To UBound(tgTimeZones) - 1 Step 1
            'lbcFrom.AddItem Trim$(tgTimeZones(ilTzt).sCSIName) & ", " & Trim$(tgTimeZones(ilTzt).sName)
            Select Case Left$(Trim$(tgTimeZones(ilTzt).sCSIName), 1)
                Case "E"
                    lbcFrom.AddItem Trim$(tgTimeZones(ilTzt).sName) & " (ETZ)"
                    lbcFrom.ItemData(lbcFrom.NewIndex) = tgTimeZones(ilTzt).iCode
                Case "C"
                    lbcFrom.AddItem Trim$(tgTimeZones(ilTzt).sName) & " (CTZ)"
                    lbcFrom.ItemData(lbcFrom.NewIndex) = tgTimeZones(ilTzt).iCode
                Case "M"
                    lbcFrom.AddItem Trim$(tgTimeZones(ilTzt).sName) & " (MTZ)"
                    lbcFrom.ItemData(lbcFrom.NewIndex) = tgTimeZones(ilTzt).iCode
                Case "P"
                    lbcFrom.AddItem Trim$(tgTimeZones(ilTzt).sName) & " (PTZ)"
                    lbcFrom.ItemData(lbcFrom.NewIndex) = tgTimeZones(ilTzt).iCode
            End Select
        Next ilTzt
    End If
End Sub

Private Function mSpecColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    mSpecColOk = True
    If grdSpec.ColWidth(grdSpec.Col) <= 15 Then
        mSpecColOk = False
        Exit Function
    End If
    If grdSpec.CellBackColor = LIGHTYELLOW Then
        mSpecColOk = False
        Exit Function
    End If

End Function

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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    Dim slStr As String
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)
    gFindMatch slStr, 0, cbcSelect    'Determine if name exist
    If gLastFound(cbcSelect) <> -1 Then   'Name found
        If gLastFound(cbcSelect) <> imSelectedIndex Then
            If Trim$(slStr) = cbcSelect.List(gLastFound(cbcSelect)) Then
                Beep
                MsgBox "Advertiser Region name already defined, enter a different name", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                mOKName = False
                Exit Function
            End If
        End If
    End If
    mOKName = True
End Function

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
Private Sub mMoveCtrlToRec()
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim slStr As String
    tmRaf.sName = grdSpec.TextMatrix(SPECROW3INDEX, NAMEINDEX)
    tmRaf.sAbbr = grdSpec.TextMatrix(SPECROW3INDEX, ABBRINDEX)
    tmRaf.sCustom = "N"
    tmRaf.sUnused = ""

    slStr = grdSpec.TextMatrix(SPECROW3INDEX, INCLEXCLINDEX)
    Select Case UCase$(slStr)
        Case "INCLUDE"
            tmRaf.sInclExcl = "I"
        Case "EXCLUDE"
            tmRaf.sInclExcl = "E"
        Case Else
            tmRaf.sInclExcl = "I"
    End Select

    slStr = grdSpec.TextMatrix(SPECROW3INDEX, CATEGORYINDEX)
    Select Case UCase$(slStr)
        Case "MARKET"
            tmRaf.sCategory = "M"
        Case "STATE NAME"
            tmRaf.sCategory = "N"
        'Case "ZIP CODE"
        '    tmRaf.sCategory = "Z"
        'Case "OWNER"
        '    tmRaf.sCategory = "O"
        Case "FORMAT"
            tmRaf.sCategory = "F"
        Case "STATION"
            tmRaf.sCategory = "S"
        Case "Time Zone"
            tmRaf.sCategory = "T"
    End Select
    tmRaf.iAudPct = gStrDecToInt(grdSpec.TextMatrix(SPECROW3INDEX, AUDPCTINDEX), 2)
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, STATUSINDEX)
    Select Case UCase$(slStr)
        Case "ACTIVE"
            tmRaf.sState = "A"
        Case "DORMANT"
            If tmRaf.sState <> "D" Then
                slStr = Format$(gNow(), "m/d/yy")
                gPackDate slStr, tmRaf.iDateDormant(0), tmRaf.iDateDormant(1)
            End If
            tmRaf.sState = "D"
        Case Else
            tmRaf.sState = "A"
    End Select
    tmRaf.lRegionCode = 0
    tmRaf.sType = "N"       'Split Network

    tmRaf.sShowNoProposal = "N"
    tmRaf.sShowOnOrder = "N"
    tmRaf.sShowOnInvoice = "N"
    If StrComp(smShowOnProp, "Yes", vbTextCompare) = 0 Then
        tmRaf.sShowNoProposal = "Y"
    End If
    If StrComp(smShowOnOrder, "Yes", vbTextCompare) = 0 Then
        tmRaf.sShowOnOrder = "Y"
    End If
    If StrComp(smShowOnInv, "Yes", vbTextCompare) = 0 Then
        tmRaf.sShowOnInvoice = "Y"
    End If
    'tmRaf.sUnused = ""

    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveSefCtrlToRec               *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                                                     *
'*******************************************************
Private Sub mMoveSefCtrlToRec()
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim llRow As Long
    Dim ilSeqNo As Integer

    ReDim tmSef(0 To 0) As SEF
    ilSeqNo = 1
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, CATEGORYINDEX)
    Select Case UCase$(slStr)
        Case "MARKET"
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                ilUpper = UBound(tmSef)
                tmSef(ilUpper).iIntCode = Val(lbcTo.ItemData(ilLoop))
                tmSef(ilUpper).lLongCode = 0
                tmSef(ilUpper).sName = ""
                tmSef(ilUpper).sCategory = ""
                tmSef(ilUpper).sInclExcl = ""
                tmSef(ilUpper).iSeqNo = ilSeqNo
                ilSeqNo = ilSeqNo + 1
                ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
            Next ilLoop
        Case "STATE NAME"
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                ilUpper = UBound(tmSef)
                tmSef(ilUpper).sName = tgStates(lbcTo.ItemData(ilLoop)).sPostalName 'lbcTo.List(ilLoop)
                tmSef(ilUpper).iIntCode = 0
                tmSef(ilUpper).lLongCode = 0
                tmSef(ilUpper).sCategory = ""
                tmSef(ilUpper).sInclExcl = ""
                tmSef(ilUpper).iSeqNo = ilSeqNo
                ilSeqNo = ilSeqNo + 1
                ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
            Next ilLoop
        'Case "ZIP CODE"
        '    For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
        '        ilUpper = UBound(tmSef)
        '        tmSef(ilUpper).sName = lbcTo.List(ilLoop)
        '        tmSef(ilUpper).iIntCode = 0
        '        tmSef(ilUpper).lLongCode = 0
        '        tmSef(ilUpper).sCategory = ""
        '        tmSef(ilUpper).sInclExcl = ""
        '        tmSef(ilUpper).iSeqNo = ilSeqNo
        '        ilSeqNo = ilSeqNo + 1
        '        ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
        '    Next ilLoop
        'Case "OWNER"
        '    For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
        '        ilUpper = UBound(tmSef)
        '        tmSef(ilUpper).iIntCode = Val(lbcTo.ItemData(ilLoop))
        '        tmSef(ilUpper).lLongCode = 0
        '        tmSef(ilUpper).sName = ""
        '        ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
        '    Next ilLoop
        Case "FORMAT"
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                ilUpper = UBound(tmSef)
                tmSef(ilUpper).iIntCode = Val(lbcTo.ItemData(ilLoop))
                tmSef(ilUpper).lLongCode = 0
                tmSef(ilUpper).sName = ""
                tmSef(ilUpper).sCategory = ""
                tmSef(ilUpper).sInclExcl = ""
                tmSef(ilUpper).iSeqNo = ilSeqNo
                ilSeqNo = ilSeqNo + 1
                ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
            Next ilLoop
        Case "STATION"
            For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
                If (grdStation.TextMatrix(llRow, STATIONINDEX) <> "") And (grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T") Then
                    ilUpper = UBound(tmSef)
                    tmSef(ilUpper).iIntCode = grdStation.TextMatrix(llRow, SHTTCODEINDEX)
                    tmSef(ilUpper).lLongCode = 0
                    tmSef(ilUpper).sName = ""
                    tmSef(ilUpper).sCategory = ""
                    tmSef(ilUpper).sInclExcl = ""
                    tmSef(ilUpper).iSeqNo = ilSeqNo
                    ilSeqNo = ilSeqNo + 1
                    ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
                End If
            Next llRow
        Case "Time Zone"
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                ilUpper = UBound(tmSef)
                    tmSef(ilUpper).iIntCode = Val(lbcTo.ItemData(ilLoop))
                    tmSef(ilUpper).lLongCode = 0
                    tmSef(ilUpper).sName = ""
                tmSef(ilUpper).sCategory = ""
                tmSef(ilUpper).sInclExcl = ""
                tmSef(ilUpper).iSeqNo = ilSeqNo
                ilSeqNo = ilSeqNo + 1
                ReDim Preserve tmSef(0 To ilUpper + 1) As SEF
            Next ilLoop
    End Select
    For ilLoop = 0 To UBound(tmSef) - 1 Step 1
        tmSef(ilLoop).lCode = 0
        tmSef(ilLoop).lRafCode = tmRaf.lCode
        tmSef(ilLoop).sUnused = ""
    Next ilLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveSefRecToCtrl               *
'*                                                     *
'*             Created:5/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record to Controls        *
'*                                                     *
'*******************************************************
Sub mMoveSEFRecToCtrl()

    Dim ilSef As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llRow As Long
    Dim ilCol As Integer

    If UCase$(tmRaf.sCategory) = "S" Then
        For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
            grdStation.Row = llRow
            For ilCol = STATIONINDEX To FORMATINDEX Step 1
                grdStation.Col = ilCol
                grdStation.CellBackColor = vbWhite
            Next ilCol
            grdStation.TextMatrix(llRow, SELECTEDINDEX) = "F"
        Next llRow
    End If
    Select Case UCase$(tmRaf.sCategory)
        Case "M"
            slStr = "Market"
        Case "N"
            slStr = "State Name"
        'Case "Z"
        '    slStr = "Zip Code"
        'Case "O"
        '    slStr = "Owner"
        Case "F"
            slStr = "Format"
        Case "S"
            slStr = "Station"
        Case "T"
            slStr = "Time Zone"
    End Select
    If tmRaf.sCategory <> "S" Then
        mPopFrom slStr
    End If
    For ilSef = 0 To UBound(tmSef) - 1 Step 1
        Select Case UCase$(tmRaf.sCategory)
            Case "M"
                For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcFrom.ItemData(ilLoop)) Then
                        lbcTo.AddItem lbcFrom.List(ilLoop)
                        lbcTo.ItemData(lbcTo.NewIndex) = lbcFrom.ItemData(ilLoop)
                        lbcFrom.RemoveItem ilLoop
                        Exit For
                    End If
                Next ilLoop
            Case "N"
                For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
                    If StrComp(Trim$(tmSef(ilSef).sName), Trim$(tgStates(lbcFrom.ItemData(ilLoop)).sPostalName), vbTextCompare) = 0 Then
                        lbcTo.AddItem lbcFrom.List(ilLoop)
                        lbcTo.ItemData(lbcTo.NewIndex) = lbcFrom.ItemData(ilLoop)
                        lbcFrom.RemoveItem ilLoop
                        Exit For
                    End If
                Next ilLoop
            'Case "Z"
            '    For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
            '        If StrComp(Trim$(tmSef(ilSef).sName), Trim$(lbcFrom.List(ilLoop)), vbTextCompare) = 0 Then
            '            lbcTo.AddItem lbcFrom.List(ilLoop)
            '            lbcTo.ItemData(lbcTo.NewIndex) = lbcFrom.ItemData(ilLoop)
            '            lbcFrom.RemoveItem ilLoop
            '            Exit For
            '        End If
            '    Next ilLoop
            'Case "O"
            '    For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
            '        If tmSef(ilSef).iIntCode = Val(lbcFrom.ItemData(ilLoop)) Then
            '            lbcTo.AddItem lbcFrom.List(ilLoop)
            '            lbcTo.ItemData(lbcTo.NewIndex) = lbcFrom.ItemData(ilLoop)
            '            lbcFrom.RemoveItem ilLoop
            '            Exit For
            '        End If
            '    Next ilLoop
            Case "F"
                For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcFrom.ItemData(ilLoop)) Then
                        lbcTo.AddItem lbcFrom.List(ilLoop)
                        lbcTo.ItemData(lbcTo.NewIndex) = lbcFrom.ItemData(ilLoop)
                        lbcFrom.RemoveItem ilLoop
                        Exit For
                    End If
                Next ilLoop
            Case "S"
                For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(grdStation.TextMatrix(llRow, SHTTCODEINDEX)) Then
                        grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T"
                        grdStation.Row = llRow
                        For ilCol = STATIONINDEX To FORMATINDEX Step 1
                            grdStation.Col = ilCol
                            grdStation.CellBackColor = GRAY
                        Next ilCol
                        Exit For
                    End If
                Next llRow
            Case "T"
                For ilLoop = 0 To lbcFrom.ListCount - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcFrom.ItemData(ilLoop)) Then
                        lbcTo.AddItem lbcFrom.List(ilLoop)
                        lbcTo.ItemData(lbcTo.NewIndex) = lbcFrom.ItemData(ilLoop)
                        lbcFrom.RemoveItem ilLoop
                        Exit For
                    End If
                Next ilLoop
        End Select
    Next ilSef
    If tmRaf.sCategory <> "S" Then
        frcCategory(0).Visible = True
        frcCategory(1).Visible = False
    Else
        frcCategory(1).Visible = True
        frcCategory(0).Visible = False
    End If
End Sub

Private Sub mStationSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
        slStr = Trim$(grdStation.TextMatrix(llRow, STATIONINDEX))
        If slStr <> "" Then
            slSort = UCase$(Trim$(grdStation.TextMatrix(llRow, ilCol)))
            If slSort = "" Then
                slSort = "~"
            End If
            slStr = grdStation.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastStationColSorted) Or ((ilCol = imLastStationColSorted) And (imLastStationSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdStation.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdStation.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastStationColSorted Then
        imLastStationColSorted = SORTINDEX
    Else
        imLastStationColSorted = -1
        imLastStationSort = -1
    End If
    gGrid_SortByCol grdStation, STATIONINDEX, SORTINDEX, imLastStationColSorted, imLastStationSort
    imLastStationColSorted = ilCol
End Sub
Private Sub mGridStationLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdStation.Rows - 1 Step 1
        grdStation.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdStation.Cols - 1 Step 1
        grdStation.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridStationColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdStation.Row = grdStation.FixedRows - 1
    grdStation.Col = STATIONINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Station"
    grdStation.Col = MARKETINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Market"
    grdStation.Col = STATEINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "State"
    grdStation.Col = ZONEINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Zone"
    'grdStation.Col = ZIPCODEINDEX
    'grdStation.CellFontBold = False
    'grdStation.CellFontName = "Arial"
    'grdStation.CellFontSize = 6.75
    'grdStation.CellForeColor = vbBlue
    'grdStation.CellBackColor = LIGHTBLUE
    'grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Zip Code"
    'grdStation.Col = OWNERINDEX
    'grdStation.CellFontBold = False
    'grdStation.CellFontName = "Arial"
    'grdStation.CellFontSize = 6.75
    'grdStation.CellForeColor = vbBlue
    'grdStation.CellBackColor = LIGHTBLUE
    'grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Owner"
    grdStation.Col = FORMATINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Format"
    grdStation.Col = SHTTCODEINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Sef Code"
    grdStation.Col = SORTINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Sort"
    grdStation.Col = SELECTEDINDEX
    grdStation.CellFontBold = False
    grdStation.CellFontName = "Arial"
    grdStation.CellFontSize = 6.75
    grdStation.CellForeColor = vbBlue
    grdStation.CellBackColor = LIGHTBLUE
    grdStation.TextMatrix(grdStation.Row, grdStation.Col) = "Selected"

End Sub

Private Sub mGridStationColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdStation.ColWidth(SHTTCODEINDEX) = 0
    grdStation.ColWidth(SORTINDEX) = 0
    grdStation.ColWidth(SELECTEDINDEX) = 0
    grdStation.ColWidth(STATIONINDEX) = 0.12 * grdStation.Width
    grdStation.ColWidth(MARKETINDEX) = 0.25 * grdStation.Width
    grdStation.ColWidth(STATEINDEX) = 0.072 * grdStation.Width
    grdStation.ColWidth(ZONEINDEX) = 0.1 * grdStation.Width
    'grdStation.ColWidth(ZIPCODEINDEX) = 0.12 * grdStation.Width
    'grdStation.ColWidth(OWNERINDEX) = 0.2 * grdStation.Width
    grdStation.ColWidth(FORMATINDEX) = 0.2 * grdStation.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdStation.Width
    For ilCol = 0 To grdStation.Cols - 1 Step 1
        llWidth = llWidth + grdStation.ColWidth(ilCol)
        If (grdStation.ColWidth(ilCol) > 15) And (grdStation.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdStation.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdStation.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdStation.Width
            For ilCol = 0 To grdStation.Cols - 1 Step 1
                If (grdStation.ColWidth(ilCol) > 15) And (grdStation.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdStation.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdStation.FixedCols To grdStation.Cols - 1 Step 1
                If grdStation.ColWidth(ilCol) > 15 Then
                    ilColInc = grdStation.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdStation.ColWidth(ilCol) = grdStation.ColWidth(ilCol) + 15
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
End Sub


Private Function mUnschdAndSchd() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilClf                                                   *
'******************************************************************************************

    Dim ilRet As Integer
    '4/3/15: TTP 7439
    Dim llChfCode As Long   'Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim ilGameNo As Integer
    Dim llSdfRecPos As Long
    Dim llLastLogDate As Long
    Dim ilVpfIndex As Integer
    Dim ilSdf As Integer

    'ReDim lgReschSdfCode(1 To 1) As Long
    ReDim lgReschSdfCode(0 To 0) As Long
    ReDim llPrevSdfCode(0 To 0) As Long
    tmClfSrchKey4.lRafCode = tmRaf.lCode
    tmClfSrchKey4.iEndDate(0) = 0
    tmClfSrchKey4.iEndDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lRafCode = tmRaf.lCode)
        If (tmClf.sSchStatus = "F") And (tmClf.sDelete <> "Y") Then
            llChfCode = tmClf.lChfCode
            'Unschedule and Schedule contract
            tmSdfSrchKey5.lCode = llChfCode
            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.lChfCode = llChfCode)
                ilFound = False
                For ilSdf = 0 To UBound(llPrevSdfCode) - 1 Step 1
                    If llPrevSdfCode(ilSdf) = tmSdf.lCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilSdf
                If Not ilFound Then
                    If tmSdf.iLineNo = tmClf.iLine Then
                        If ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G")) Then
                            If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                ilVpfIndex = gVpfFind(NetworkSplit, tmSdf.iVefCode)
                                If ilVpfIndex <> -1 Then
                                    gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llLastLogDate
                                    If llDate > llLastLogDate Then
                                        If mPreemptSpot(tmSdf) Then
                                            ilFound = True
                                            ilRet = btrGetPosition(hmSdf, llSdfRecPos)
                                            ilRet = gMakeTracer(hmSdf, tmSdf, llSdfRecPos, hmStf, llLastLogDate, "M", "U", tmSdf.iRotNo, hmGsf)
                                            ilGameNo = tmSdf.iGameNo
                                            ilRet = gChgSchSpot("M", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tgSsf(0), lgSsfDate(0), lgSsfRecPos(0), hmSxf, hmGsf, hmGhf)
                                            lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                            'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                                            ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        llPrevSdfCode(UBound(llPrevSdfCode)) = tmSdf.lCode
                        ReDim Preserve llPrevSdfCode(0 To UBound(llPrevSdfCode) + 1) As Long
                    End If
                    If ilFound Then
                        tmSdfSrchKey5.lCode = llChfCode
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
                    Else
                        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    End If
                Else
                    ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                End If
            Loop
        End If
        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If UBound(lgReschSdfCode) > LBound(lgReschSdfCode) Then
        gGetSchParameters
        If gOpenSchFiles() Then
            'If imSave(1) = 1 Then
            '    If (llStartDate + 6 <= llEndDate) And (llStartTime = 0) And (llEndTime >= 86399) Then
            '        igUsePreferred = True
            '    End If
            'End If
            ilRet = gReSchSpots(True, 0, "YYYYYYY", 0, 86400)
            igUsePreferred = False
            'sgPreemptPass = "N"
            gCloseSchFiles
        End If
    End If
    Erase llPrevSdfCode
    Erase lgReschSdfCode
    mUnschdAndSchd = True
End Function

Private Function mPreemptSpot(tlSdf As SDF) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim slTime As String
    Dim ilLoop As Integer
    Dim ilSpot As Integer

    mPreemptSpot = True
    gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llDate
    gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slTime
    ilRet = gObtainSsfForDateOrGame(tlSdf.iVefCode, llDate, slTime, tmSdf.iGameNo, hmSsf, tgSsf(0), lgSsfDate(0), lgSsfRecPos(0))
    If ilRet Then
        For ilLoop = 1 To tgSsf(0).iCount Step 1
           LSet tmAvail = tgSsf(0).tPas(ADJSSFPASBZ + ilLoop)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                For ilSpot = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                   LSet tmSpot = tgSsf(0).tPas(ADJSSFPASBZ + ilSpot)
                    If tlSdf.lCode = tmSpot.lSdfCode Then
                        If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                            If ilSpot + 1 <= ilLoop + tmAvail.iNoSpotsThis Then
                               LSet tmSpot = tgSsf(0).tPas(ADJSSFPASBZ + ilSpot + 1)
                                If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                    mPreemptSpot = False
                                End If
                            Else
                                mPreemptSpot = False
                            End If
                        End If
                        Exit Function
                    End If
                Next ilSpot
            End If
        Next ilLoop
    End If
End Function

Private Function mCheckForAdditions() As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilSef As Integer
    Dim ilFound As Integer
    Dim llRow As Long
    Dim ilRet As Integer
    Dim ilShowMgs As Integer

    ilShowMgs = False
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, CATEGORYINDEX)
    Select Case UCase$(slStr)
        Case "MARKET"
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                ilFound = False
                For ilSef = 0 To UBound(tmSef) - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcTo.ItemData(ilLoop)) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilSef
                If Not ilFound Then
                    ilShowMgs = True
                    Exit For
                End If
            Next ilLoop
        Case "STATE NAME"
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                ilFound = False
                For ilSef = 0 To UBound(tmSef) - 1 Step 1
                    If StrComp(Trim$(tmSef(ilSef).sName), Trim$(lbcTo.List(ilLoop)), vbTextCompare) = 0 Then
                        ilFound = True
                        Exit For
                    End If
                Next ilSef
                If Not ilFound Then
                    ilShowMgs = True
                    Exit For
                End If
            Next ilLoop
        Case "TIME ZONE"
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                ilFound = False
                For ilSef = 0 To UBound(tmSef) - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcTo.ItemData(ilLoop)) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilSef
                If Not ilFound Then
                    ilShowMgs = True
                    Exit For
                End If
            Next ilLoop
        'Case "ZIP CODE"
        '    For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
        '        ilFound = False
        '        For ilSef = 0 To UBound(tmSef) - 1 Step 1
        '            If StrComp(Trim$(tmSef(ilSef).sName), Trim$(lbcTo.List(ilLoop)), vbTextCompare) = 0 Then
        '                ilFound = True
        '                Exit For
        '            End If
        '        Next ilSef
        '        If Not ilFound Then
        '            ilShowMgs = True
        '            Exit For
        '        End If
        '    Next ilLoop
        'Case "OWNER"
        '    For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
        '        ilFound = False
        '        For ilSef = 0 To UBound(tmSef) - 1 Step 1
        '            If tmSef(ilSef).iIntCode = Val(lbcTo.ItemData(ilLoop)) Then
        '                ilFound = True
        '                Exit For
        '            End If
        '        Next ilSef
        '        If Not ilFound Then
        '            ilShowMgs = True
        '            Exit For
        '        End If
        '    Next ilLoop
        Case "FORMAT"
            For ilLoop = 0 To lbcTo.ListCount - 1 Step 1
                ilFound = False
                For ilSef = 0 To UBound(tmSef) - 1 Step 1
                    If tmSef(ilSef).iIntCode = Val(lbcTo.ItemData(ilLoop)) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilSef
                If Not ilFound Then
                    ilShowMgs = True
                    Exit For
                End If
            Next ilLoop
        Case "STATION"
            For llRow = grdStation.FixedRows To grdStation.Rows - 1 Step 1
                If (grdStation.TextMatrix(llRow, STATIONINDEX) <> "") And (grdStation.TextMatrix(llRow, SELECTEDINDEX) = "T") Then
                    ilFound = False
                    For ilSef = 0 To UBound(tmSef) - 1 Step 1
                        If tmSef(ilSef).iIntCode = grdStation.TextMatrix(llRow, SHTTCODEINDEX) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilSef
                    If Not ilFound Then
                        ilShowMgs = True
                        Exit For
                    End If
                End If
            Next llRow
    End Select
    If ilShowMgs Then 'Station added
        gSetMousePointer grdSpec, grdStation, vbDefault
        Screen.MousePointer = vbDefault
        Beep
        ilRet = MsgBox("All Contracts will be scanned to resolve station conflicts, Continue with save", vbInformation + vbYesNo, "Warning")
        If ilRet = vbNo Then
            mCheckForAdditions = -1
        Else
            mCheckForAdditions = 1
            Screen.MousePointer = vbHourglass
            gSetMousePointer grdSpec, grdStation, vbHourglass
        End If
    Else
        mCheckForAdditions = 0
    End If

End Function

Private Sub mPopTimeZones()
    Dim ilRet As Integer

    ilRet = gObtainTimeZones()
End Sub

Private Sub mPopStates()
    Dim ilRet As Integer

    ilRet = gObtainStates()
End Sub

