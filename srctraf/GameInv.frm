VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form GameInv 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   10815
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
   ScaleWidth      =   10815
   Begin VB.ComboBox cbcSeason 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   5310
      TabIndex        =   2
      Top             =   45
      Width           =   2115
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
      Left            =   6120
      TabIndex        =   33
      Top             =   5490
      Width           =   1050
   End
   Begin VB.ListBox lbcGameNoSort 
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
      ItemData        =   "GameInv.frx":0000
      Left            =   9480
      List            =   "GameInv.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1545
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox pbcIndependent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6105
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   495
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.ComboBox cbcGameVeh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   1980
      TabIndex        =   1
      Top             =   45
      Width           =   3240
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
      Left            =   4815
      TabIndex        =   30
      Top             =   5490
      Width           =   1050
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   7515
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   2985
   End
   Begin VB.PictureBox pbcFeedSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4380
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmcGet 
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
      Left            =   1260
      Picture         =   "GameInv.frx":0004
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   405
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcGet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   315
      MaxLength       =   10
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pbcGetTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   45
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   12
      Top             =   750
      Width           =   45
   End
   Begin VB.PictureBox pbcGetSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   45
      ScaleHeight     =   30
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   300
      Width           =   30
   End
   Begin VB.ListBox lbcInvItem 
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
      ItemData        =   "GameInv.frx":00FE
      Left            =   3390
      List            =   "GameInv.frx":0100
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   270
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox pbcOversell 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5925
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   315
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmcSetInv 
      Appearance      =   0  'Flat
      Caption         =   "&Set Inv."
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
      Left            =   9450
      TabIndex        =   29
      Top             =   1095
      Width           =   1050
   End
   Begin VB.ListBox lbcInvType 
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
      ItemData        =   "GameInv.frx":0102
      Left            =   765
      List            =   "GameInv.frx":0104
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   300
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8895
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5370
      Width           =   45
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8010
      Top             =   5385
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
      ItemData        =   "GameInv.frx":0106
      Left            =   1575
      List            =   "GameInv.frx":0108
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox edcSet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   420
      MaxLength       =   10
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1335
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pbcSetTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   19
      Top             =   1125
      Width           =   45
   End
   Begin VB.PictureBox pbcSetSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   30
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   13
      Top             =   795
      Width           =   45
   End
   Begin VB.CommandButton cmcSet 
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
      Left            =   1365
      Picture         =   "GameInv.frx":010A
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   990
      MaxLength       =   10
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2580
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "GameInv.frx":0204
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1905
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
      TabIndex        =   27
      TabStop         =   0   'False
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
      TabIndex        =   24
      Top             =   5070
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   15
      ScaleHeight     =   105
      ScaleWidth      =   90
      TabIndex        =   20
      Top             =   1680
      Width           =   90
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
      Left            =   3510
      TabIndex        =   26
      Top             =   5490
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
      Left            =   2205
      TabIndex        =   25
      Top             =   5490
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDates 
      Height          =   3360
      Left            =   195
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1890
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   5927
      _Version        =   393216
      Rows            =   31
      Cols            =   23
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
      _Band(0).Cols   =   23
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSet 
      Height          =   885
      Left            =   210
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   930
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1561
      _Version        =   393216
      Rows            =   11
      Cols            =   8
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
      _Band(0).Cols   =   8
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdGet 
      Height          =   465
      Left            =   210
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   390
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   820
      _Version        =   393216
      Rows            =   5
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin VB.Label lacIndMsg 
      Appearance      =   0  'Flat
      Caption         =   "Ind = Event-Independent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   225
      TabIndex        =   31
      Top             =   5250
      Width           =   1845
   End
   Begin VB.Label plcScreen 
      Caption         =   "Multimedia Inventory"
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
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   1950
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   5400
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "GameInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GameInv.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmGhfSrchKey0                 tmGsfSrchKey0                 tmIhfSrchKey1             *
'*  tmIsfSrchKey1                 tmIsfSrchKey2                                           *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: GameInv.Frm
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
Dim imPopReqd As Integer
Dim imBypassFocus As Integer
Dim imSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Dim imVefCode As Integer
Dim imVpfIndex As Integer
Dim lmSeasonGhfCode As Long
Dim smFeedSource As String
Dim smOversell As String
Dim smIndependent As String
Dim imMinGameNo As Integer
Dim imMaxGameNo As Integer
Dim imLastColSorted As Integer
Dim imLastSort As Integer
Dim imBypassSetting As Integer
Dim smLastType As String
Dim smDefaultGameNo As String

Dim smNowDate As String
Dim lmNowDate As Long
Dim lmLLD As Long
Dim lmFirstAllowedChgDate As Long

Dim tmGameVehicle() As SORTCODE
Dim smGameVehicleTag As String

Dim smTeamTag As String
Dim tmTeam() As MNF

Dim tmInvType() As SORTCODE
Dim smInvTypeTag As String

Dim tmInvItem() As SORTCODE
Dim smInvItemTag As String

Dim imIhfChg As Integer
Dim imIsfChg As Integer
Dim imNewInv As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmSetEnableRow As Long
Dim lmSetEnableCol As Long
Dim imSetCtrlVisible As Integer
Dim lmGetEnableRow As Long
Dim lmGetEnableCol As Long
Dim imGetCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer

Dim imItfCode As Integer
Dim imIhfCode As Integer
Dim imIifCode As Integer

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmGhfSrchKey0 As LONGKEY0    'ISF key record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length

Dim hmGsf As Integer
Dim tmGsf() As GSF        'GSF record image
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

Dim hmIhf As Integer
Dim tmIhf As IHF        'IHF record image
Dim tmIhfSrchKey0 As INTKEY0    'IHF key record image
Dim tmIhfSrchKey1 As IHFKEY1    'IHF key record image
Dim tmIhfSrchKey2 As IHFKEY2    'IHF key record image
Dim imIhfRecLen As Integer        'IHF record length

Dim hmIsf As Integer
Dim tmIsf() As ISF        'ISF record image
Dim tmIsfSrchKey0 As LONGKEY0    'ISF key record image
Dim tmIsfSrchKey3 As ISFKEY3    'ISF key record image
Dim imIsfRecLen As Integer        'ISF record length

Dim hmItf As Integer
Dim tmItf As ITF        'ITF record image
Dim tmItfSrchKey0 As INTKEY0    'ITF key record image
Dim imItfRecLen As Integer        'ITF record length
Dim smIgnoreMultiFeed As String

Dim hmIif As Integer
Dim tmIif As IIF        'IIF record image
Dim tmIifSrchKey0 As INTKEY0    'IIF key record image
Dim imIifRecLen As Integer        'IIF record length

Dim tmLanguageCode() As SORTCODE
Dim smLanguageCodeTag As String

'6/9/14
Dim smEventTitle1 As String
Dim smEventTitle2 As String

Dim imSeasonComboBoxIndex As Integer
Dim imSeasonSelectedIndex As Integer

'Mouse down
Const GETROW3INDEX = 3
Const GETTYPEINDEX = 2
Const GETITEMINDEX = 4
Const GETINDEPENDENTINDEX = 6
Const GETOVERSELLINDEX = 8

'Row 3
Const SETROW3INDEX = 3
Const SETUNITINDEX = 2
Const SETCOSTINDEX = 4
Const SETRATEINDEX = 6
'Row 6
Const SETROW6INDEX = 6
Const SETLANGUAGEINDEX = 2
Const SETFEEDSOURCEINDEX = 6
'Row 9
Const SETROW9INDEX = 9
Const SETGAMENOSINDEX = 2
Const SETGAMEININDEX = 4
Const SETGAMEOUTINDEX = 6

Const GAMENOINDEX = 2   '1
Const FEEDSOURCEINDEX = 4   '2
Const LANGUAGEINDEX = 6 '3
Const TEAMSINDEX = 8
Const AIRDATEINDEX = 10 '7
Const AIRTIMEINDEX = 12 '8
Const UNITINDEX = 14
Const COSTINDEX = 16
Const RATEINDEX = 18
Const TMISFINDEX = 20
Const SORTINDEX = 21
Const GAMESTATUSINDEX = 22





Private Sub cbcGameVeh_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilIndex                                                 *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  cbcGameVehErr                                                                         *
'******************************************************************************************

    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim slCode As String

    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    gSetMousePointer grdSet, grdGet, vbHourglass
    gSetMousePointer grdGet, grdDates, vbHourglass
    ilRet = gOptionLookAhead(cbcGameVeh, imBSMode, slStr)
    mClearCtrlFields
    mClearGrdDates True
    If ilRet = 0 Then
        slStr = tmGameVehicle(cbcGameVeh.ListIndex).sKey
        ilRet = gParseItem(slStr, 2, "\", slCode)
        imVefCode = Val(slCode)
        imVpfIndex = gVpfFind(GameInv, imVefCode)
        gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLLD
        mSeasonPop False
        mTeamPop
        mClearCtrlFields
        ilRet = mGhfGsfReadRec()
        mLanguagePop
        mInvTypePop
        mPopulate
    End If
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
    imChgMode = False
    imBypassSetting = False
    Exit Sub
cbcGameVehErr: 'VBC NR
    On Error GoTo 0
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub

Private Sub cbcGameVeh_Click()
    cbcGameVeh_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcGameVeh_GotFocus()
    If imTerminate Then
        Exit Sub
    End If
    mGetSetShow
    mSetSetShow
    mSetShow
    gCtrlGotFocus cbcGameVeh
End Sub

Private Sub cbcSeason_Change()
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        If cbcSeason.Text <> "" Then
            gManLookAhead cbcSeason, imBSMode, imSeasonComboBoxIndex
            mcbcSeasonChange
        End If
    End If
End Sub

Private Sub cbcSeason_Click()
    cbcSeason_Change
End Sub

Private Sub cbcSeason_GotFocus()
    Dim ilVff As Integer
    Dim ilLoop As Integer
    
    mGetSetShow
    mSetSetShow
    mSetShow
    pbcArrow.Visible = False
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilIndex                                                 *
'******************************************************************************************

    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    gSetMousePointer grdSet, grdGet, vbHourglass
    gSetMousePointer grdGet, grdDates, vbHourglass
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    mClearCtrlFields
    If ilRet = 0 Then
        imIhfCode = cbcSelect.ItemData(cbcSelect.ListIndex)
        If Not mReadRec() Then
            GoTo cbcSelectErr
        End If
        imNewInv = False
    Else
        imNewInv = True
        imIhfCode = 0
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    imIhfChg = False
    imIsfChg = False
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mGSFMoveRecToCtrl
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        smIgnoreMultiFeed = "N"
        mGSFMoveRecToCtrl
    End If
    mSetCommands
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
    DoEvents
    imChgMode = False
    imBypassSetting = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mGetSetShow
    mSetSetShow
    mSetShow
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    'mPopulate
    'If imTerminate Then
    '    Exit Sub
    'End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        On Error Resume Next
        pbcGetSTab.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    gCtrlGotFocus cbcSelect
    If (slSvText = "") Or (slSvText = "[New]") Then
        If slSvText <> "[New]" Then
            cbcSelect.ListIndex = 0
        Else
            cbcSelect_Change    'Call change so picture area repainted
        End If
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                'cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            'mClearCtrlFields
            'cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
End Sub

Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mGetSetShow
    mSetSetShow
    mSetShow
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                         ilError                                                 *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slMess As String

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If imIhfChg Or imIsfChg Then
        If imNewInv Then
            slMess = "Add Inventory for " & grdGet.TextMatrix(GETROW3INDEX, GETTYPEINDEX) & "/" & grdGet.TextMatrix(GETROW3INDEX, GETITEMINDEX)
        Else
            slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
        End If
        ilRet = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        If ilRet = vbNo Then
            cbcSelect.ListIndex = 0
            Exit Sub
        End If
        If ilRet = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mGetSetShow
    mSetSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcDone
End Sub

Private Sub cmcErase_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStamp                                                                               *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slMsg As String
    Dim slInvType As String
    Dim slInvItem As String
    Dim tlIsf As ISF

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imSelectedIndex > 0 Then
        slInvType = Trim$(grdGet.TextMatrix(GETROW3INDEX, GETTYPEINDEX))
        slInvItem = Trim$(grdGet.TextMatrix(GETROW3INDEX, GETITEMINDEX))
        Screen.MousePointer = vbHourglass
        'Check that record is not referenced-Code missing
        ilRet = gIICodeRefExist(GameInv, imIhfCode, "Msf.Btr", "MsfIhfCode")
        If ilRet Then
            Screen.MousePointer = vbDefault
            slMsg = "Cannot erase - an Contract references name"
            ilRet = MsgBox(slMsg, vbOKOnly + vbExclamation, "Erase")
            Exit Sub
        End If
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("OK to remove " & slInvType & "/" & slInvItem, vbOKCancel + vbQuestion, "Erase")
        If ilRet = vbCancel Then
            Exit Sub
        End If
        Do
            tmIhfSrchKey0.iCode = imIhfCode
            ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                Do
                    tmIsfSrchKey3.iIhfCode = tmIhf.iCode
                    tmIsfSrchKey3.iGameNo = 0
                    ilRet = btrGetGreaterOrEqual(hmIsf, tlIsf, imIsfRecLen, tmIsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                    If (ilRet <> BTRV_ERR_NONE) Or (tmIhf.iCode <> tlIsf.iIhfCode) Then
                        Exit Do
                    End If
                    ilRet = btrDelete(hmIsf)
                Loop
            End If
            ilRet = btrDelete(hmIhf)
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo cmcEraseErr
        gBtrvErrorMsg ilRet, "cmcErase_Click (btrDelete)", GameInv
        On Error GoTo 0
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mPopulate
    End If
    'Remove focus from control and make invisible
    mClearCtrlFields
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
    mGetSetShow
    mSetSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcErase
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    If imNewInv Then
        slName = grdGet.TextMatrix(GETROW3INDEX, GETTYPEINDEX) & "/" & grdGet.TextMatrix(GETROW3INDEX, GETITEMINDEX)
    Else
        slName = cbcSelect.List(imSelectedIndex)
    End If
    ilRet = mSaveRec()
    If ilRet Then
        imNewInv = False
        imIhfChg = False
        imIsfChg = False
        mPopulate
        mSetCommands
        '12/14/05 (Jim)- default to new after save
        'gFindMatch slName, 1, cbcSelect
        'If gLastFound(cbcSelect) > 0 Then
        '    cbcSelect.ListIndex = gLastFound(cbcSelect)
        'Else
            cbcSelect.ListIndex = 0
        'End If
        cbcSelect.SetFocus
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mGetSetShow
    mSetSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcSave
End Sub

Private Sub cmcSet_Click()
    Select Case grdSet.Row
        Case SETROW6INDEX
            Select Case grdSet.Col
                Case SETLANGUAGEINDEX
                    lbcLanguage.Visible = Not lbcLanguage.Visible
            End Select
    End Select
    edcSet.SelStart = 0
    edcSet.SelLength = Len(edcSet.Text)
    edcSet.SetFocus
End Sub

Private Sub cmcSet_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcSetInv_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slLineNo                      ilLineNo                      ilLine                    *
'*  llDate                        ilDateLoop                    ilCount                   *
'*  ilNoSpots                     ilNoSpotsPerWk                ilSpots                   *
'*  ilFound                                                                               *
'******************************************************************************************

    Dim slStr As String
    Dim llRow As Long
    Dim ilPos As Integer
    Dim ilGameNo As Integer
    Dim ilSearchFrom As Integer
    Dim ilNextComma As Integer
    Dim ilStartGame As Integer
    Dim ilEndGame As Integer
    Dim ilGamesOn As Integer
    Dim ilGamesOff As Integer
    Dim ilGameCount As Integer
    Dim ilGameTest As Integer
    Dim ilLoop As Integer
    Dim ilGame As Integer
    Dim ilLangOk As Integer
    Dim ilFeedOk As Integer
    Dim ilValue As Integer
    Dim ilGameOk As Integer
    Dim slGameStatus As String

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSet, grdGet, vbHourglass
    gSetMousePointer grdGet, grdDates, vbHourglass
    grdDates.Redraw = False
    'Replace Number of spots in selected weeks only
    'For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
    '    If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
    '        grdDates.TextMatrix(llRow, NOSPOTSINDEX) = ""
    '    End If
    'Next llRow
    imIsfChg = True
    slStr = grdSet.TextMatrix(SETROW9INDEX, SETGAMEININDEX)
    If slStr = "" Then
        ilGamesOn = 1000
        ilGamesOff = 0
    Else
        ilGamesOn = Val(slStr)
        slStr = grdSet.TextMatrix(SETROW9INDEX, SETGAMEOUTINDEX)
        If slStr = "" Then
            ilGamesOff = 0
        Else
            ilGamesOff = Val(slStr)
        End If
    End If
    slStr = grdSet.TextMatrix(SETROW9INDEX, SETGAMENOSINDEX)
    'Create an array of game numbers
    ReDim slFields(0 To 0) As String
    ReDim ilGameNos(0 To 0) As Integer
    ilSearchFrom = 1
    Do
        ilNextComma = InStr(ilSearchFrom, slStr, ",", vbTextCompare)
        If ilNextComma > 0 Then
            slFields(UBound(slFields)) = Mid$(slStr, ilSearchFrom, ilNextComma - ilSearchFrom)
            ilSearchFrom = ilNextComma + 1
            ReDim Preserve slFields(0 To UBound(slFields) + 1) As String
        Else
            slFields(UBound(slFields)) = Mid$(slStr, ilSearchFrom)
            ReDim Preserve slFields(0 To UBound(slFields) + 1) As String
            Exit Do
        End If
    Loop While ilSearchFrom <= Len(slStr)
    ilValue = Asc(tgSpf.sSportInfo)
    For ilLoop = 0 To UBound(slFields) - 1 Step 1
        ilPos = InStr(1, slFields(ilLoop), "-", vbTextCompare)
        If ilPos > 0 Then
            ilStartGame = Left$(slFields(ilLoop), ilPos - 1)
            ilEndGame = Mid$(slFields(ilLoop), ilPos + 1)
            ilGameCount = 0
            ilGameTest = 0
            For ilGameNo = ilStartGame To ilEndGame Step 1
                ilGameOk = True
                slGameStatus = ""
                For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                    If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                        If ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX)) Then
                            grdDates.Row = llRow
                            grdDates.Col = UNITINDEX
                            If grdDates.ForeColor = vbRed Then
                                ilGameOk = False
                            End If
                            slGameStatus = grdDates.TextMatrix(llRow, GAMESTATUSINDEX)
                            Exit For
                        End If
                    End If
                Next llRow
                If (ilValue And USINGFEED) = USINGFEED Then
                    'If (smFeedSource = "Home") Or (smFeedSource = "Visiting") Or (smFeedSource = "National") Then
                    If (smFeedSource = smEventTitle2) Or (smFeedSource = smEventTitle1) Or (smFeedSource = "National") Then
                        ilFeedOk = False
                        For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                            If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                                ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                                If ilGameNo = ilGame Then
                                    slStr = grdDates.TextMatrix(llRow, FEEDSOURCEINDEX)
                                    If StrComp(smFeedSource, slStr, vbTextCompare) = 0 Then
                                        ilFeedOk = True
                                    End If
                                    Exit For
                                End If
                            End If
                        Next llRow
                    Else
                        ilFeedOk = True
                    End If
                Else
                    ilFeedOk = True
                End If
                'Ignore Language not matching games
                If (ilValue And USINGLANG) = USINGLANG Then
                    If lbcLanguage.ListIndex > 0 Then
                        ilLangOk = False
                        For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                            If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                                ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                                If ilGameNo = ilGame Then
                                    slStr = grdDates.TextMatrix(llRow, LANGUAGEINDEX)
                                    gFindMatch slStr, 0, lbcLanguage
                                    If gLastFound(lbcLanguage) >= 0 Then
                                        If lbcLanguage.ListIndex = gLastFound(lbcLanguage) Then
                                            ilLangOk = True
                                        End If
                                    End If
                                    Exit For
                                End If
                            End If
                        Next llRow
                    Else
                        ilLangOk = True
                    End If
                Else
                    ilLangOk = True
                End If
                If ilGameOk And ilFeedOk And ilLangOk Then
                    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                            If ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX)) Then
                                grdDates.TextMatrix(llRow, UNITINDEX) = ""
                                grdDates.TextMatrix(llRow, COSTINDEX) = ""
                                grdDates.TextMatrix(llRow, RATEINDEX) = ""
                                Exit For
                            End If
                        End If
                    Next llRow
                    If slGameStatus <> "C" Then
                        If ilGameTest = 0 Then
                            ilGameNos(UBound(ilGameNos)) = ilGameNo
                            ReDim Preserve ilGameNos(0 To UBound(ilGameNos) + 1) As Integer
                            ilGameCount = ilGameCount + 1
                            If (ilGameCount >= ilGamesOn) And (ilGamesOff > 0) Then
                                ilGameTest = 1
                                ilGameCount = 0
                            End If
                        Else
                            ilGameCount = ilGameCount + 1
                            If ilGameCount >= ilGamesOff Then
                                ilGameTest = 0
                                ilGameCount = 0
                            End If
                        End If
                    End If
                End If
            Next ilGameNo
        Else
            'Check Language
            If (ilValue And &H8) = &H8 Then
                If lbcLanguage.ListIndex > 0 Then
                    ilLangOk = False
                    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                            ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                            If ilGameNo = Val(slFields(ilLoop)) Then
                                slStr = grdDates.TextMatrix(llRow, LANGUAGEINDEX)
                                gFindMatch slStr, 0, lbcLanguage
                                If gLastFound(lbcLanguage) >= 0 Then
                                    If lbcLanguage.Selected(gLastFound(lbcLanguage)) Then
                                        ilLangOk = True
                                    End If
                                End If
                                Exit For
                            End If
                        End If
                    Next llRow
                Else
                    ilLangOk = True
                End If
            Else
                ilLangOk = True
            End If
            If ilLangOk Then
                For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                    If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                        If Val(slFields(ilLoop)) = Val(grdDates.TextMatrix(llRow, GAMENOINDEX)) Then
                            grdDates.TextMatrix(llRow, UNITINDEX) = ""
                            grdDates.TextMatrix(llRow, COSTINDEX) = ""
                            grdDates.TextMatrix(llRow, RATEINDEX) = ""
                            Exit For
                        End If
                    End If
                Next llRow
                ilGameNos(UBound(ilGameNos)) = Val(slFields(ilLoop))
                ReDim Preserve ilGameNos(0 To UBound(ilGameNos) + 1) As Integer
            End If
        End If
    Next ilLoop
    'Distribute spots to games
    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
            ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
            For ilGame = 0 To UBound(ilGameNos) - 1 Step 1
                If ilGameNos(ilGame) = ilGameNo Then
                    grdDates.TextMatrix(llRow, UNITINDEX) = grdSet.TextMatrix(SETROW3INDEX, SETUNITINDEX)
                    grdDates.TextMatrix(llRow, COSTINDEX) = grdSet.TextMatrix(SETROW3INDEX, SETCOSTINDEX)
                    grdDates.TextMatrix(llRow, RATEINDEX) = grdSet.TextMatrix(SETROW3INDEX, SETRATEINDEX)
                    Exit For
                End If
            Next ilGame
        End If
    Next llRow
    grdDates.Redraw = True
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcSetInv_GotFocus()
    mGetSetShow
    mSetSetShow
    mSetShow
    If Not mSetFormFieldsOk() Then
        cmcCancel.SetFocus
    End If
End Sub

Private Sub cmcGet_Click()
    Select Case grdGet.Row
        Case SETROW3INDEX
            Select Case grdGet.Col
                Case GETTYPEINDEX
                    lbcInvType.Visible = Not lbcInvType.Visible
                Case GETITEMINDEX
                    lbcInvItem.Visible = Not lbcInvItem.Visible
            End Select
    End Select
    edcGet.SelStart = 0
    edcGet.SelLength = Len(edcGet.Text)
    edcGet.SetFocus
End Sub
Private Sub cmcGet_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Sub edcDropDown_Change()
    grdDates.CellForeColor = vbBlack
End Sub


Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim ilKey As Integer
    Dim ilPos As Integer

    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case lmEnableCol
        Case UNITINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropdown.Text
            slStr = Left$(slStr, edcDropdown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropdown.SelStart - edcDropdown.SelLength)
            If gCompNumberStr(slStr, "9999") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case COSTINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            ilPos = InStr(edcDropdown.SelText, ".")
            If ilPos = 0 Then
                ilPos = InStr(edcDropdown.Text, ".")    'Disallow multi-decimal points
                If ilPos > 0 Then
                    If KeyAscii = KEYDECPOINT Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropdown.Text
            slStr = Left$(slStr, edcDropdown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropdown.SelStart - edcDropdown.SelLength)
            If gCompNumberStr(slStr, "9999999.99") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case RATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            ilPos = InStr(edcDropdown.SelText, ".")
            If ilPos = 0 Then
                ilPos = InStr(edcDropdown.Text, ".")    'Disallow multi-decimal points
                If ilPos > 0 Then
                    If KeyAscii = KEYDECPOINT Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropdown.Text
            slStr = Left$(slStr, edcDropdown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropdown.SelStart - edcDropdown.SelLength)
            If gCompNumberStr(slStr, "9999999.99") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub edcGet_Change()
    Dim slStr As String
    Dim ilRet As Integer

    Select Case grdGet.Row
        Case GETROW3INDEX
            Select Case grdGet.Col
                Case GETTYPEINDEX
                    imLbcArrowSetting = True
                    ilRet = gOptionalLookAhead(edcGet, lbcInvType, imBSMode, slStr)
                    If ilRet = 1 Then
                        If lbcInvType.ListCount > 0 Then
                            lbcInvType.ListIndex = 0
                        End If
                    End If
                Case GETITEMINDEX
                    imLbcArrowSetting = True
                    ilRet = gOptionalLookAhead(edcGet, lbcInvItem, imBSMode, slStr)
                    If ilRet = 1 Then
                        If lbcInvItem.ListCount > 0 Then
                            lbcInvItem.ListIndex = 0
                        End If
                    End If
            End Select
    End Select
    grdGet.CellForeColor = vbBlack
    imLbcArrowSetting = False
End Sub

Private Sub edcGet_DblClick()
    Select Case grdGet.Row
        Case GETROW3INDEX
            Select Case lmGetEnableCol
                Case GETTYPEINDEX
                    imDoubleClickName = True
                Case GETITEMINDEX
                    imDoubleClickName = True
            End Select
    End Select
End Sub

Private Sub edcGet_GotFocus()
    Select Case grdGet.Row
        Case GETROW3INDEX
            Select Case lmGetEnableCol
                Case GETTYPEINDEX
                Case GETITEMINDEX
            End Select
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcGet_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcGet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcGet.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case grdGet.Row
        Case GETROW3INDEX
            Select Case lmGetEnableCol
                Case GETTYPEINDEX
                Case GETITEMINDEX
            End Select
    End Select
End Sub

Private Sub edcGet_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdGet.Row
            Case SETROW3INDEX
                Select Case grdGet.Col
                    Case GETTYPEINDEX
                        gProcessArrowKey Shift, KeyCode, lbcInvType, imLbcArrowSetting
                    Case GETITEMINDEX
                        gProcessArrowKey Shift, KeyCode, lbcInvItem, imLbcArrowSetting
                End Select
        End Select
        edcGet.SelStart = 0
        edcGet.SelLength = Len(edcGet.Text)
    End If
End Sub

Private Sub edcGet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case grdGet.Row
            Case GETROW3INDEX
                Select Case lmGetEnableCol
                    Case GETTYPEINDEX
                        On Error Resume Next
                        If imTabDirection = -1 Then  'Right To Left
                            pbcGetSTab.SetFocus
                        Else
                            pbcGetTab.SetFocus
                        End If
                        On Error GoTo 0
                        Exit Sub
                    Case GETITEMINDEX
                        On Error Resume Next
                        If imTabDirection = -1 Then  'Right To Left
                            pbcGetSTab.SetFocus
                        Else
                            pbcGetTab.SetFocus
                        End If
                        On Error GoTo 0
                        Exit Sub
                End Select
        End Select
        imDoubleClickName = False
    End If
End Sub

Private Sub edcSet_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                                                   *
'******************************************************************************************


    Select Case grdSet.Row
        Case SETROW3INDEX
            Select Case lmSetEnableCol
                Case SETUNITINDEX
                Case SETCOSTINDEX
                Case SETRATEINDEX
            End Select
        Case SETROW6INDEX
            Select Case lmSetEnableCol
                Case SETLANGUAGEINDEX
                    imLbcArrowSetting = True
                    gMatchLookAhead edcSet, lbcLanguage, imBSMode, imComboBoxIndex
            End Select
        Case SETROW9INDEX
            Select Case lmSetEnableCol
                Case SETGAMENOSINDEX
                Case SETGAMEININDEX
                Case SETGAMEOUTINDEX
            End Select
    End Select
    grdSet.CellForeColor = vbBlack
    imLbcArrowSetting = False
End Sub

Private Sub edcSet_DblClick()
    Select Case grdSet.Row
        Case SETROW3INDEX
            Select Case lmSetEnableCol
                Case SETUNITINDEX
                Case SETCOSTINDEX
                Case SETRATEINDEX
            End Select
        Case SETROW6INDEX
            Select Case lmSetEnableCol
                Case SETLANGUAGEINDEX
            End Select
        Case SETROW9INDEX
            Select Case lmSetEnableCol
                Case SETGAMENOSINDEX
                Case SETGAMEININDEX
                Case SETGAMEOUTINDEX
            End Select
    End Select
End Sub

Private Sub edcSet_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcSet_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcSet_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim ilKey As Integer
    Dim ilPos As Integer

    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSet.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case grdSet.Row
        Case SETROW3INDEX
            Select Case lmSetEnableCol
                Case SETUNITINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                    slStr = edcSet.Text
                    slStr = Left$(slStr, edcSet.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSet.SelStart - edcSet.SelLength)
                    If gCompNumberStr(slStr, "9999") > 0 Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SETCOSTINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    ilPos = InStr(edcSet.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcSet.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If KeyAscii = KEYDECPOINT Then
                                Beep
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                    slStr = edcSet.Text
                    slStr = Left$(slStr, edcSet.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSet.SelStart - edcSet.SelLength)
                    If gCompNumberStr(slStr, "9999999.99") > 0 Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SETRATEINDEX
                    'Filter characters (allow only BackSpace, numbers 0 thru 9
                    ilPos = InStr(edcSet.SelText, ".")
                    If ilPos = 0 Then
                        ilPos = InStr(edcSet.Text, ".")    'Disallow multi-decimal points
                        If ilPos > 0 Then
                            If KeyAscii = KEYDECPOINT Then
                                Beep
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                    slStr = edcSet.Text
                    slStr = Left$(slStr, edcSet.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSet.SelStart - edcSet.SelLength)
                    If gCompNumberStr(slStr, "9999999.99") > 0 Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
            End Select
        Case SETROW6INDEX
            Select Case lmSetEnableCol
                Case SETLANGUAGEINDEX
            End Select
        Case SETROW9INDEX
            Select Case lmSetEnableCol
                Case SETGAMENOSINDEX
                    If (KeyAscii = KEYBACKSPACE) Or (KeyAscii = KEYNEG) Or ((KeyAscii >= KEY0) And (KeyAscii <= KEY9)) Or (KeyAscii = KEYCOMMA) Then
                    Else
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SETGAMEININDEX
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                    slStr = edcSet.Text
                    slStr = Left$(slStr, edcSet.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSet.SelStart - edcSet.SelLength)
                    If gCompNumberStr(slStr, Trim$(str$(imMaxGameNo))) > 0 Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                Case SETGAMEOUTINDEX
                    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
                    slStr = edcSet.Text
                    slStr = Left$(slStr, edcSet.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSet.SelStart - edcSet.SelLength)
                    If gCompNumberStr(slStr, Trim$(str$(imMaxGameNo))) > 0 Then
                        Beep
                        KeyAscii = 0
                        Exit Sub
                    End If
            End Select
    End Select
End Sub

Private Sub edcSet_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdSet.Row
            Case SETROW6INDEX
                Select Case lmSetEnableCol
                    Case SETLANGUAGEINDEX
                        gProcessArrowKey Shift, KeyCode, lbcLanguage, imLbcArrowSetting
                End Select
        End Select
        edcSet.SelStart = 0
        edcSet.SelLength = Len(edcSet.Text)
    End If
End Sub

Private Sub edcSet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imDoubleClickName = False
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
        pbcSetSTab.Enabled = False
        pbcSetTab.Enabled = False
        grdSet.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        grdDates.Enabled = False
        imUpdateAllowed = False
    Else
        grdSet.Enabled = True
        pbcSetSTab.Enabled = True
        pbcSetTab.Enabled = True
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        If lmEnableCol > 0 Then
            mEnableBox
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    
    Erase tmGsf
    Erase tmIsf

    smGameVehicleTag = ""
    Erase tmGameVehicle

    smTeamTag = ""
    Erase tmTeam

    smLanguageCodeTag = ""
    Erase tmLanguageCode
    smInvItemTag = ""
    Erase tmInvItem
    smInvTypeTag = ""
    Erase tmInvType

    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf

    ilRet = btrClose(hmIsf)
    btrDestroy hmIsf
    ilRet = btrClose(hmIhf)
    btrDestroy hmIhf

    ilRet = btrClose(hmItf)
    btrDestroy hmItf
    ilRet = btrClose(hmIif)
    btrDestroy hmIif

    Set GameInv = Nothing   'Remove data segment

End Sub

Private Sub grdDates_EnterCell()
    mGetSetShow
    mSetSetShow
    mSetShow
End Sub

Private Sub grdDates_GotFocus()
    mGetSetShow
    mSetSetShow
End Sub

Private Sub grdDates_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdDates.TopRow
    grdDates.Redraw = False
End Sub

Private Sub grdDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llHeight As Long

    For llRow = 0 To grdDates.FixedRows - 1 Step 1
        llHeight = grdDates.RowHeight(llRow)
    Next llRow
    grdDates.ToolTipText = ""
    If Y <= llHeight Then
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdDates, X, Y, llRow, llCol)
    If (llCol = GAMENOINDEX) Then
        grdDates.ToolTipText = Trim$(grdDates.TextMatrix(llRow, llCol))
    End If
End Sub

Private Sub grdDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine if in header
    On Error GoTo grdDatesErr:
    If Y < grdDates.RowHeight(0) + grdDates.RowHeight(1) + grdDates.RowHeight(2) + grdDates.RowHeight(3) Then
        grdDates.Col = grdDates.MouseCol
        mSortCol grdDates.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
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
    On Error GoTo grdDates_ScrollErr:
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
            edcDropdown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
    End If
    lmTopRow = grdDates.TopRow
    Exit Sub
grdDates_ScrollErr:
    grdDates.Redraw = True
    lmTopRow = grdDates.TopRow
    Exit Sub
End Sub

Private Sub grdGet_EnterCell()
    mGetSetShow
    mSetSetShow
    mSetShow
End Sub

Private Sub grdGet_GotFocus()
    mSetSetShow
    mSetShow
End Sub

Private Sub grdGet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine row and col mouse up onto
    On Error GoTo grdGetErr:
    ilCol = grdGet.MouseCol
    ilRow = grdGet.MouseRow
    If ilCol < grdGet.FixedCols Then
        Exit Sub
    End If
    If ilRow < grdGet.FixedRows Then
        Exit Sub
    End If
    If (ilRow = SETROW6INDEX - 1) Then
        ilRow = ilRow + 1
    End If
    If grdGet.ColWidth(ilCol) <= 15 Then
        Exit Sub
    End If
    If grdGet.RowHeight(ilRow) <= 15 Then
        Exit Sub
    End If
    DoEvents
    grdGet.Col = ilCol
    grdGet.Row = ilRow
    If Not mGetColOk() Then
        Exit Sub
    End If
    mGetEnableBox
    On Error GoTo 0
    Exit Sub
grdGetErr:
    On Error GoTo 0
    If (lmGetEnableRow >= grdGet.FixedRows) And (lmGetEnableRow < grdGet.Rows) Then
        grdGet.Row = lmGetEnableRow
        grdGet.Col = lmGetEnableRow
        mGetSetFocus
    End If
    grdGet.Redraw = False
    grdGet.Redraw = True
    Exit Sub
End Sub

Private Sub grdSet_EnterCell()
    mGetSetShow
    mSetSetShow
    mSetShow
End Sub

Private Sub grdSet_GotFocus()
    mGetSetShow
    mSetShow
End Sub

Private Sub grdSet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine if in header
'    If y < grdSet.RowHeight(0) Then
'        mSortCol grdSet.Col
'        Exit Sub
'    End If
    'Determine row and col mouse up onto
    On Error GoTo grdSetErr:
    ilCol = grdSet.MouseCol
    ilRow = grdSet.MouseRow
    If ilCol < grdSet.FixedCols Then
        Exit Sub
    End If
    If ilRow < grdSet.FixedRows Then
        Exit Sub
    End If
    If (ilRow = SETROW3INDEX - 1) Or (ilRow = SETROW6INDEX - 1) Or (ilRow = SETROW9INDEX - 1) Then
        ilRow = ilRow + 1
    End If
    If grdSet.ColWidth(ilCol) <= 15 Then
        Exit Sub
    End If
    If grdSet.RowHeight(ilRow) <= 15 Then
        Exit Sub
    End If
    DoEvents
    grdSet.Col = ilCol
    grdSet.Row = ilRow
    If Not mSetColOk() Then
        Exit Sub
    End If
    mSetEnableBox
    On Error GoTo 0
    Exit Sub
grdSetErr:
    On Error GoTo 0
    If (lmSetEnableRow >= grdSet.FixedRows) And (lmSetEnableRow < grdSet.Rows) Then
        grdSet.Row = lmSetEnableRow
        grdSet.Col = lmSetEnableRow
        mSetSetFocus
    End If
    grdSet.Redraw = False
    grdSet.Redraw = True
    Exit Sub
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (grdDates.Row < grdDates.FixedRows) Or (grdDates.Row >= grdDates.Rows) Or (grdDates.Col < grdDates.FixedCols) Or (grdDates.Col >= grdDates.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdDates.Row
    lmEnableCol = grdDates.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdDates.Left - pbcArrow.Width - 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + (grdDates.RowHeight(grdDates.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    Select Case grdDates.Col
        Case UNITINDEX
            edcDropdown.MaxLength = 5
            edcDropdown.Text = grdDates.Text
        Case COSTINDEX
            edcDropdown.MaxLength = 10
            edcDropdown.Text = grdDates.Text
        Case RATEINDEX
            edcDropdown.MaxLength = 10
            edcDropdown.Text = grdDates.Text
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slNameCode                    slName                        slDaypart                 *
'*  slLineNo                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim slCode As String
    Dim ilLoop As Integer

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSet, grdGet, vbHourglass
    gSetMousePointer grdGet, grdDates, vbHourglass
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
    imCtrlVisible = False
    imGetCtrlVisible = False
    imSetCtrlVisible = False
    imIsfChg = False
    imNewInv = True
    lmTopRow = -1
    imLastColSorted = -1
    imLastSort = -1
    smLastType = ""
    smIgnoreMultiFeed = "N"
    imSetCtrlVisible = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmFirstAllowedChgDate = lmNowDate + 1
    mInitBox
    ReDim tmInvType(0 To 0) As SORTCODE
    ReDim tmInvItem(0 To 0) As SORTCODE
    hmGhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameInv
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)  'Get and save ARF record length

    hmGsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameInv
    On Error GoTo 0
    ReDim tmGsf(0 To 0) As GSF
    imGsfRecLen = Len(tmGsf(0))  'Get and save ARF record length

    hmIhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmIhf, "", sgDBPath & "Ihf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameInv
    On Error GoTo 0
    imIhfRecLen = Len(tmIhf)  'Get and save ARF record length

    hmIsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmIsf, "", sgDBPath & "Isf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameInv
    On Error GoTo 0
    ReDim tmIsf(0 To 0) As ISF
    imIsfRecLen = Len(tmIsf(0))  'Get and save ARF record length

    hmItf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmItf, "", sgDBPath & "Itf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameInv
    On Error GoTo 0
    imItfRecLen = Len(tmItf)  'Get and save ARF record length

    hmIif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmIif, "", sgDBPath & "Iif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GameInv
    On Error GoTo 0
    imIifRecLen = Len(tmIif)  'Get and save ARF record length

    imVefCode = igGameSchdVefCode
    lmSeasonGhfCode = lgSeasonGhfCode

    mVehPop
    For ilLoop = 0 To UBound(tmGameVehicle) - 1 Step 1
        slStr = tmGameVehicle(ilLoop).sKey
        ilRet = gParseItem(slStr, 2, "\", slCode)
        If Val(slCode) = igGameSchdVefCode Then
            cbcGameVeh.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    lmSeasonGhfCode = lgSeasonGhfCode
    mSeasonPop True
    

   ' mTeamPop
   ' 'mLanguagePop
   ' mClearCtrlFields
   ' ilRet = mGhfGsfReadRec()
   ' 'mMoveRecToCtrl
   ' If ilRet Then
   '     mLanguagePop
   '     'mGSFMoveRecToCtrl
   ' End If
   ' mInvTypePop
   ' mPopulate
    GameInv.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone GameInv
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
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
'*  flTextHeight                  ilLoop                        ilRow                     *
'*                                                                                        *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim ilCol As Integer
    'flTextHeight = pbcDates.TextHeight("1") - 35

    cbcGameVeh.Move 2040, 45
    cbcSeason.Move 5340, 45
    cbcSelect.Move 7515, 45
    mGridGetLayout
    mGridGetColumnWidths
    mGridGetColumns
    grdGet.Move 180, cbcSelect.Top + cbcSelect.Height + 120
    grdGet.Height = grdGet.RowPos(grdGet.Rows - 1) + grdGet.RowHeight(grdGet.Rows - 1) + fgPanelAdj - 15
    mGridSetLayout
    mGridSetColumnWidths
    'mGridSetColumns
    grdSet.Visible = False
    cmcSetInv.Visible = False
    pbcSetSTab.Visible = False
    pbcSetTab.Visible = False

    grdSet.Move grdGet.Left, grdGet.Top + grdGet.Height + 120
    grdSet.Height = grdSet.RowPos(grdSet.Rows - 1) + grdSet.RowHeight(grdSet.Rows - 1) + fgPanelAdj - 15
    cmcSetInv.Move grdSet.Left + grdSet.Width + 120, grdSet.Top + grdSet.Height - cmcSetInv.Height
    'Merge Columns
    grdSet.Row = SETROW6INDEX
    For ilCol = SETLANGUAGEINDEX To SETLANGUAGEINDEX + 2 Step 1
        grdSet.TextMatrix(grdSet.Row, ilCol) = " "
    Next ilCol
    grdSet.MergeRow(SETROW6INDEX) = True
    grdSet.MergeRow(SETROW6INDEX - 1) = True
    grdSet.MergeCells = 1  '2 work, 3 and 4 don't work

    grdDates.Move grdSet.Left, grdSet.Top + grdSet.Height + 120
    imInitNoRows = grdDates.Rows
    mGridLayout
    mGridColumnWidths
    'mGridColumns
    grdDates.Height = grdDates.RowPos(grdDates.Rows - 1) + grdDates.RowHeight(grdDates.Rows - 1) + fgPanelAdj - 15
    lacIndMsg.Move grdDates.Left, grdDates.Top + grdDates.Height + 60
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'

    pbcArrow.Visible = False
    If (lmEnableRow >= grdDates.FixedRows) And (lmEnableRow < grdDates.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case UNITINDEX
                edcDropdown.Visible = False
                If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropdown.Text Then
                    imIsfChg = True
                End If
                grdDates.TextMatrix(lmEnableRow, lmEnableCol) = edcDropdown.Text
            Case COSTINDEX
                edcDropdown.Visible = False
                If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropdown.Text Then
                    imIsfChg = True
                End If
                grdDates.TextMatrix(lmEnableRow, lmEnableCol) = edcDropdown.Text
            Case RATEINDEX
                edcDropdown.Visible = False
                If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropdown.Text Then
                    imIsfChg = True
                End If
                grdDates.TextMatrix(lmEnableRow, lmEnableCol) = edcDropdown.Text
        End Select
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

    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload GameInv
    igManUnload = NO
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilAirDateDefined              ilAirTimeDefined          *
'*  ilValue                                                                               *
'******************************************************************************************

'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilError As Integer

    ilError = False
'    slStr = grdDates.TextMatrix(ilRowNo, AIRDATEINDEX)
'    If (slStr <> "") And (slStr <> "Missing") Then
'        ilAirDateDefined = True
'    Else
'        ilAirDateDefined = False
'    End If
'    slStr = grdDates.TextMatrix(ilRowNo, AIRTIMEINDEX)
'    If (slStr <> "") And (slStr <> "Missing") Then
'        ilAirTimeDefined = True
'    Else
'        ilAirTimeDefined = False
'    End If
'    If (Not ilAirDateDefined) And (Not ilAirTimeDefined) Then
'        mGridFieldsOk = True
'        Exit Function
'    End If
'    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
'    If (ilValue And &H4) = &H4 Then
'        slStr = grdDates.TextMatrix(ilRowNo, FEEDSOURCEINDEX)
'        If (StrComp(slStr, "Visiting", vbTextCompare) <> 0) And (StrComp(slStr, "Home", vbTextCompare) <> 0) Then
'            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
'                ilError = True
'                grdDates.TextMatrix(ilRowNo, FEEDSOURCEINDEX) = "Missing"
'                grdDates.Row = ilRowNo
'                grdDates.Col = FEEDSOURCEINDEX
'                grdDates.CellForeColor = vbRed
'            Else
'                ilError = True
'                grdDates.Row = ilRowNo
'                grdDates.Col = FEEDSOURCEINDEX
'                grdDates.CellForeColor = vbRed
'            End If
'        Else
'            grdDates.Row = ilRowNo
'            grdDates.Col = FEEDSOURCEINDEX
'            grdDates.CellForeColor = vbBlack
'        End If
'    End If
'    'Language
'    If (ilValue And &H8) = &H8 Then
'        slStr = grdDates.TextMatrix(ilRowNo, LANGUAGEINDEX)
'        gFindMatch slStr, 0, lbcLanguage
'        If gLastFound(lbcLanguage) < 0 Then
'            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
'                ilError = True
'                grdDates.TextMatrix(ilRowNo, LANGUAGEINDEX) = "Missing"
'                grdDates.Row = ilRowNo
'                grdDates.Col = LANGUAGEINDEX
'                grdDates.CellForeColor = vbRed
'            Else
'                ilError = True
'                grdDates.Row = ilRowNo
'                grdDates.Col = LANGUAGEINDEX
'                grdDates.CellForeColor = vbRed
'            End If
'        Else
'            grdDates.Row = ilRowNo
'            grdDates.Col = LANGUAGEINDEX
'            grdDates.CellForeColor = vbBlack
'        End If
'    End If
'    'Visiting Team
'    slStr = grdDates.TextMatrix(ilRowNo, VISITTEAMINDEX)
'    gFindMatch slStr, 0, lbcTeam
'    If gLastFound(lbcTeam) < 0 Then
'        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
'            ilError = True
'            grdDates.TextMatrix(ilRowNo, VISITTEAMINDEX) = "Missing"
'            grdDates.Row = ilRowNo
'            grdDates.Col = VISITTEAMINDEX
'            grdDates.CellForeColor = vbRed
'        Else
'            ilError = True
'            grdDates.Row = ilRowNo
'            grdDates.Col = VISITTEAMINDEX
'            grdDates.CellForeColor = vbRed
'        End If
'    Else
'        grdDates.Row = ilRowNo
'        grdDates.Col = VISITTEAMINDEX
'        grdDates.CellForeColor = vbBlack
'    End If
'    'Home Team
'    slStr = grdDates.TextMatrix(ilRowNo, HOMETEAMINDEX)
'    gFindMatch slStr, 0, lbcTeam
'    If gLastFound(lbcTeam) < 0 Then
'        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
'            ilError = True
'            grdDates.TextMatrix(ilRowNo, HOMETEAMINDEX) = "Missing"
'            grdDates.Row = ilRowNo
'            grdDates.Col = HOMETEAMINDEX
'            grdDates.CellForeColor = vbRed
'        Else
'            ilError = True
'            grdDates.Row = ilRowNo
'            grdDates.Col = HOMETEAMINDEX
'            grdDates.CellForeColor = vbRed
'        End If
'    Else
'        grdDates.Row = ilRowNo
'        grdDates.Col = HOMETEAMINDEX
'        grdDates.CellForeColor = vbBlack
'    End If
'    'Air Date
'    slStr = grdDates.TextMatrix(ilRowNo, AIRDATEINDEX)
'    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
'        ilError = True
'        grdDates.TextMatrix(ilRowNo, AIRDATEINDEX) = "Missing"
'        grdDates.Row = ilRowNo
'        grdDates.Col = AIRDATEINDEX
'        grdDates.CellForeColor = vbRed
'    Else
'        If Not gValidDate(slStr) Then
'            ilError = True
'            grdDates.Row = ilRowNo
'            grdDates.Col = AIRDATEINDEX
'            grdDates.CellForeColor = vbRed
'        Else
'            grdDates.Row = ilRowNo
'            grdDates.Col = AIRDATEINDEX
'            grdDates.CellForeColor = vbBlack
'        End If
'    End If
'    'Air Time
'    slStr = grdDates.TextMatrix(ilRowNo, AIRTIMEINDEX)
'    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
'        ilError = True
'        grdDates.TextMatrix(ilRowNo, AIRTIMEINDEX) = "Missing"
'        grdDates.Row = ilRowNo
'        grdDates.Col = AIRTIMEINDEX
'        grdDates.CellForeColor = vbRed
'    Else
'        If Not gValidTime(slStr) Then
'            ilError = True
'            grdDates.Row = ilRowNo
'            grdDates.Col = AIRTIMEINDEX
'            grdDates.CellForeColor = vbRed
'        Else
'            grdDates.Row = ilRowNo
'            grdDates.Col = AIRTIMEINDEX
'            grdDates.CellForeColor = vbBlack
'        End If
'    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGSFMoveRecToCtrl               *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mGSFMoveRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                       ilLib                         ilVeh                     *
'*  ilPos                                                                                 *
'******************************************************************************************

'
'   mXFerRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilValue As Integer
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim llColor As Long
    Dim ilFirstLastWk As Integer
    Dim ilWkNo As Integer
    Dim ilWkNoMinus1 As Integer
    Dim ilColorSet As Integer
    Dim ilAddRow As Integer
    Dim ilTest As Integer
    Dim ilItemNo As Integer
    Dim slTestStr As String
    Dim llFind As Long
    Dim ilStartGameNo As Integer
    Dim ilRunningGameNo As Integer
    Dim slGameNo As String

    grdDates.Redraw = False
    lbcGameNoSort.Clear
    mClearGrdDates False
    ilColorSet = False
    llRow = grdDates.FixedRows
    imMinGameNo = 0
    imMaxGameNo = 0
    'Add All
    grdDates.Row = llRow
    grdDates.TextMatrix(llRow, TMISFINDEX) = "-1"
    grdDates.TextMatrix(llRow, GAMENOINDEX) = "Ind"
    grdDates.Col = GAMENOINDEX
    grdDates.CellBackColor = LIGHTYELLOW
    grdDates.CellAlignment = flexAlignLeftCenter
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    slStr = ""
    If (ilValue And USINGFEED) = USINGFEED Then
        grdDates.Col = FEEDSOURCEINDEX
        grdDates.CellBackColor = LIGHTYELLOW
    End If
    grdDates.TextMatrix(llRow, FEEDSOURCEINDEX) = Trim$(slStr)
    'Language
    slStr = ""
    If (ilValue And USINGLANG) = USINGLANG Then
        grdDates.Col = LANGUAGEINDEX
        grdDates.CellBackColor = LIGHTYELLOW
    End If
    grdDates.TextMatrix(llRow, LANGUAGEINDEX) = slStr
    'Teams
    grdDates.TextMatrix(llRow, TEAMSINDEX) = ""
    grdDates.Col = TEAMSINDEX
    grdDates.CellBackColor = LIGHTYELLOW
    'Air Date
    grdDates.TextMatrix(llRow, AIRDATEINDEX) = ""
    grdDates.Col = AIRDATEINDEX
    grdDates.CellBackColor = LIGHTYELLOW
    'Air Time
    grdDates.TextMatrix(llRow, AIRTIMEINDEX) = ""
    grdDates.Col = AIRTIMEINDEX
    grdDates.CellBackColor = LIGHTYELLOW

    grdDates.TextMatrix(llRow, UNITINDEX) = ""
    grdDates.Col = UNITINDEX
    grdDates.CellBackColor = vbWhite
    grdDates.TextMatrix(llRow, COSTINDEX) = ""
    grdDates.Col = COSTINDEX
    grdDates.CellBackColor = vbWhite
    grdDates.TextMatrix(llRow, RATEINDEX) = ""
    grdDates.Col = RATEINDEX
    grdDates.CellBackColor = vbWhite
    llRow = llRow + 2
    For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
        If smIgnoreMultiFeed = "Y" Then
            ilAddRow = True
            'Test if Row matched a previous row
            For ilTest = 0 To ilLoop - 1 Step 1
                If (tmGsf(ilTest).iVisitMnfCode = tmGsf(ilLoop).iVisitMnfCode) And (tmGsf(ilTest).iHomeMnfCode = tmGsf(ilLoop).iHomeMnfCode) Then
                    If (tmGsf(ilTest).iAirDate(0) = tmGsf(ilLoop).iAirDate(0)) And (tmGsf(ilTest).iAirDate(1) = tmGsf(ilLoop).iAirDate(1)) Then
                        If (tmGsf(ilTest).iAirTime(0) = tmGsf(ilLoop).iAirTime(0)) And (tmGsf(ilTest).iAirTime(1) = tmGsf(ilLoop).iAirTime(1)) Then
                            ilAddRow = False
                            For llFind = grdDates.FixedRows + 2 To grdDates.Rows - 1 Step 2
                                slStr = grdDates.TextMatrix(llFind, GAMENOINDEX)
                                ilItemNo = 1
                                Do
                                    ilRet = gParseItem(slStr, ilItemNo, ",", slTestStr)
                                    If slTestStr = "" Then
                                        Exit Do
                                    End If
                                    If Val(slTestStr) = tmGsf(ilTest).iGameNo Then
                                        slStr = slStr & "," & Trim$(str$(tmGsf(ilLoop).iGameNo))
                                        grdDates.TextMatrix(llFind, GAMENOINDEX) = slStr
                                        grdDates.TextMatrix(llFind, FEEDSOURCEINDEX) = "All"
                                        grdDates.TextMatrix(llFind, LANGUAGEINDEX) = "All"
                                        Exit For
                                    End If
                                    ilItemNo = ilItemNo + 1
                                Loop While slTestStr <> ""
                            Next llFind
                        End If
                    End If
                End If
            Next ilTest
        Else
            ilAddRow = True
        End If
        If ilAddRow Then
            If llRow + 1 > grdDates.Rows Then
                grdDates.AddItem ""
                grdDates.RowHeight(grdDates.Rows - 1) = fgBoxGridH
                grdDates.AddItem ""
                grdDates.RowHeight(grdDates.Rows - 1) = 15
                tmGsf(ilLoop).lCode = 0
                mInitNew llRow
            End If
            grdDates.Row = llRow
            grdDates.TextMatrix(llRow, TMISFINDEX) = "-1"
            If imMinGameNo = 0 Then
                imMinGameNo = tmGsf(ilLoop).iGameNo
                imMaxGameNo = tmGsf(ilLoop).iGameNo
            Else
                If tmGsf(ilLoop).iGameNo < imMinGameNo Then
                    imMinGameNo = tmGsf(ilLoop).iGameNo
                End If
                If tmGsf(ilLoop).iGameNo > imMaxGameNo Then
                    imMaxGameNo = tmGsf(ilLoop).iGameNo
                End If
            End If
            slStr = Trim$(str$(tmGsf(ilLoop).iGameNo))
            grdDates.TextMatrix(llRow, GAMENOINDEX) = Trim$(slStr)
            grdDates.Col = GAMENOINDEX
            grdDates.CellBackColor = LIGHTYELLOW
            grdDates.CellAlignment = flexAlignLeftCenter
            'Feed
            ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
            slStr = ""
            If (ilValue And USINGFEED) = USINGFEED Then
                If tmGsf(ilLoop).sFeedSource = "V" Then
                    slStr = smEventTitle1   '"Visiting"
                ElseIf tmGsf(ilLoop).sFeedSource = "H" Then
                    slStr = smEventTitle2   '"Home"
                ElseIf tmGsf(ilLoop).sFeedSource = "N" Then
                    slStr = "National"
                End If
                grdDates.Col = FEEDSOURCEINDEX
                grdDates.CellBackColor = LIGHTYELLOW
            End If
            grdDates.TextMatrix(llRow, FEEDSOURCEINDEX) = Trim$(slStr)
            'Language
            slStr = ""
            If (ilValue And USINGLANG) = USINGLANG Then
                For ilLang = 0 To UBound(tmLanguageCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                    slNameCode = tmLanguageCode(ilLang).sKey 'Traffic!lbcAgency.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If tmGsf(ilLoop).iLangMnfCode = Val(slCode) Then
                        ilRet = gParseItem(slNameCode, 1, "\", slStr)
                        Exit For
                    End If
                Next ilLang
                grdDates.Col = LANGUAGEINDEX
                grdDates.CellBackColor = LIGHTYELLOW
            End If
            grdDates.TextMatrix(llRow, LANGUAGEINDEX) = slStr
            'Teams
            slStr = ""
            'For ilTeam = 1 To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            For ilTeam = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                If tmGsf(ilLoop).iVisitMnfCode = tmTeam(ilTeam).iCode Then
                    slStr = Trim$(Left$(tmTeam(ilTeam).sName, 4))
                    If Trim$(tmTeam(ilTeam).sUnitType) <> "" Then
                        slStr = Trim$(tmTeam(ilTeam).sUnitType)
                    End If
                    Exit For
                End If
            Next ilTeam
            'For ilTeam = 1 To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            For ilTeam = LBound(tmTeam) To UBound(tmTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                If tmGsf(ilLoop).iHomeMnfCode = tmTeam(ilTeam).iCode Then
                    If Trim$(tmTeam(ilTeam).sUnitType) <> "" Then
                        slStr = slStr & "@" & Trim$(tmTeam(ilTeam).sUnitType)
                    Else
                        slStr = slStr & "@" & Trim$(Left(tmTeam(ilTeam).sName, 4))
                    End If
                    Exit For
                End If
            Next ilTeam
            grdDates.TextMatrix(llRow, TEAMSINDEX) = slStr
            grdDates.Col = TEAMSINDEX
            grdDates.CellBackColor = LIGHTYELLOW
            'Air Date
            gUnpackDate tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), slStr
            If gDateValue(slStr) < lmFirstAllowedChgDate Then
                grdDates.Row = llRow
                grdDates.Col = AIRDATEINDEX
                grdDates.CellForeColor = vbRed
                grdDates.Col = AIRTIMEINDEX
                grdDates.CellForeColor = vbRed
                grdDates.Col = UNITINDEX
                grdDates.CellForeColor = vbRed
            Else
                slGameNo = Trim$(str$(tmGsf(ilLoop).iGameNo))
                Do While Len(slGameNo) < 5
                    slGameNo = "0" & slGameNo
                Loop
                lbcGameNoSort.AddItem slGameNo
                lbcGameNoSort.ItemData(lbcGameNoSort.NewIndex) = tmGsf(ilLoop).iGameNo
            End If
            grdDates.TextMatrix(llRow, AIRDATEINDEX) = slStr
            grdDates.Col = AIRDATEINDEX
            grdDates.CellBackColor = LIGHTYELLOW
            'Air Time
            gUnpackTime tmGsf(ilLoop).iAirTime(0), tmGsf(ilLoop).iAirTime(1), "A", "1", slStr
            grdDates.TextMatrix(llRow, AIRTIMEINDEX) = slStr
            grdDates.Col = AIRTIMEINDEX
            grdDates.CellBackColor = LIGHTYELLOW

            grdDates.TextMatrix(llRow, UNITINDEX) = ""
            grdDates.Col = UNITINDEX
            grdDates.CellBackColor = vbWhite
            grdDates.TextMatrix(llRow, COSTINDEX) = ""
            grdDates.Col = COSTINDEX
            grdDates.CellBackColor = vbWhite
            grdDates.TextMatrix(llRow, RATEINDEX) = ""
            grdDates.Col = RATEINDEX
            grdDates.CellBackColor = vbWhite


            slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
            If gDateValue(slStr) <= lmNowDate Then
                llColor = vbRed
                ilColorSet = False
            Else
                If Not ilColorSet Then
                    ilColorSet = True
                    llColor = vbBlack
                Else
                    gObtainWkNo 0, grdDates.TextMatrix(llRow, AIRDATEINDEX), ilWkNo, ilFirstLastWk
                    gObtainWkNo 0, grdDates.TextMatrix(llRow - 2, AIRDATEINDEX), ilWkNoMinus1, ilFirstLastWk
                    If ilWkNo <> ilWkNoMinus1 Then
                        If llColor = vbBlack Then
                            llColor = DARKGREEN
                        Else
                            llColor = vbBlack
                        End If
                    End If
                End If
            End If
            For ilCol = grdDates.FixedCols To grdDates.Cols - 1 Step 2
                grdDates.Col = ilCol
                grdDates.CellForeColor = llColor
            Next ilCol
            grdDates.TextMatrix(llRow, GAMESTATUSINDEX) = tmGsf(ilLoop).sGameStatus
            If tmGsf(ilLoop).sGameStatus = "C" Then
                For ilCol = grdDates.FixedCols To grdDates.Cols - 1 Step 2
                    grdDates.Col = ilCol
                    If grdDates.CellForeColor <> vbRed Then
                        grdDates.CellForeColor = vbCyan
                    End If
                Next ilCol
            End If
            llRow = llRow + 2
        End If
    Next ilLoop

    smDefaultGameNo = ""
    ilLoop = 0
    Do While ilLoop < lbcGameNoSort.ListCount
        If smDefaultGameNo <> "" Then
            smDefaultGameNo = smDefaultGameNo & ","
        End If
        ilStartGameNo = lbcGameNoSort.ItemData(ilLoop)
        ilRunningGameNo = ilStartGameNo + 1
        smDefaultGameNo = smDefaultGameNo & Trim$(str$(lbcGameNoSort.ItemData(ilLoop)))
        ilLoop = ilLoop + 1
        Do While ilLoop < lbcGameNoSort.ListCount
            If ilRunningGameNo <> lbcGameNoSort.ItemData(ilLoop) Then
                If ilStartGameNo <> ilRunningGameNo - 1 Then
                    smDefaultGameNo = smDefaultGameNo & "-" & Trim$(str(ilRunningGameNo - 1))
                End If
                Exit Do
            End If
            ilLoop = ilLoop + 1
            If ilLoop >= lbcGameNoSort.ListCount Then
                smDefaultGameNo = smDefaultGameNo & "-" & Trim$(str(ilRunningGameNo))
                Exit Do
            End If
            ilRunningGameNo = ilRunningGameNo + 1
        Loop
    Loop

    grdDates.Redraw = True
    Exit Sub
End Sub




Private Sub lbcInvItem_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcInvItem, edcGet, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcInvItem_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True
End Sub

Private Sub lbcInvItem_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcInvItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcInvItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcInvItem, edcGet, imChgMode, imLbcArrowSetting
        On Error Resume Next
        If imTabDirection = -1 Then  'Right To Left
            pbcGetSTab.SetFocus
        Else
            pbcGetTab.SetFocus
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub lbcInvType_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcInvType, edcGet, imChgMode, imLbcArrowSetting
    End If
End Sub

Private Sub lbcInvType_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True
End Sub

Private Sub lbcInvType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcInvType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub

Private Sub lbcInvType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcInvType, edcGet, imChgMode, imLbcArrowSetting
        On Error Resume Next
        If imTabDirection = -1 Then  'Right To Left
            pbcGetSTab.SetFocus
        Else
            pbcGetTab.SetFocus
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub lbcLanguage_Click()
    gProcessLbcClick lbcLanguage, edcSet, imChgMode, imLbcArrowSetting
    grdSet.CellForeColor = vbBlack
End Sub


Private Sub lbcLanguage_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcClickFocus_GotFocus()
    mGetSetShow
    mSetSetShow
    mSetShow
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub





Private Sub pbcGetSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcGetSTab.hWnd Then
        Exit Sub
    End If
    If imGetCtrlVisible Then
        If grdGet.Col = GETTYPEINDEX Then
            If mInvTypeBranch() Then
                Exit Sub
            End If
        End If
        If grdGet.Col = GETITEMINDEX Then
            If mInvItemBranch() Then
                Exit Sub
            End If
        End If
        Do
            ilNext = False
            Select Case grdGet.Row
                Case GETROW3INDEX
                    Select Case grdGet.Col
                        Case GETTYPEINDEX
                            mGetSetShow
                            On Error Resume Next
                            pbcSetSTab.SetFocus
                            On Error GoTo 0
                            Exit Sub
                        Case Else
                            grdGet.Col = grdGet.Col - 2
                    End Select
            End Select
            If mGetColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mGetSetShow
    Else
        grdGet.Row = GETROW3INDEX '+1 to bypass title
        grdGet.Col = grdGet.FixedCols
        Do
            If mGetColOk() Then
                Exit Do
            Else
                grdGet.Col = grdGet.Col + 2
            End If
        Loop
    End If
    mGetEnableBox
End Sub

Private Sub pbcGetTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llSpecEnableRow               llSpecEnableCol                                         *
'******************************************************************************************

    Dim ilNext As Integer

    If GetFocus() <> pbcGetTab.hWnd Then
        Exit Sub
    End If
    If imGetCtrlVisible Then
        If grdGet.Col = GETTYPEINDEX Then
            If mInvTypeBranch() Then
                Exit Sub
            End If
        End If
        If grdGet.Col = GETITEMINDEX Then
            If mInvItemBranch() Then
                Exit Sub
            End If
        End If
        Do
            ilNext = False
            Select Case grdGet.Row
                Case GETROW3INDEX
                    Select Case grdGet.Col
                        Case GETOVERSELLINDEX
                            mGetSetShow
                            On Error Resume Next
                            If pbcSetSTab.Visible Then
                                pbcSetSTab.SetFocus
                            Else
                                pbcSTab.SetFocus
                            End If
                            On Error GoTo 0
                            Exit Sub
                        Case Else
                            grdGet.Col = grdGet.Col + 2
                    End Select
            End Select
            If mGetColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mGetSetShow
    Else
        grdGet.Row = grdGet.Rows - 2
        grdGet.Col = grdGet.FixedCols
        Do
            If mGetColOk() Then
                Exit Do
            Else
                grdGet.Col = grdGet.Col - 2
            End If
        Loop
    End If
    mGetEnableBox

End Sub

Private Sub pbcIndependent_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        smIndependent = "Yes"
        pbcIndependent_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        smIndependent = "No"
        pbcIndependent_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smIndependent = "Yes" Then
            smIndependent = "No"
            pbcIndependent_Paint
        ElseIf smIndependent = "No" Then
            smIndependent = "Yes"
            pbcIndependent_Paint
        End If
    End If
End Sub

Private Sub pbcIndependent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smIndependent = "Yes" Then
        smIndependent = "No"
        pbcIndependent_Paint
    ElseIf smIndependent = "No" Then
        smIndependent = "Yes"
        pbcIndependent_Paint
    End If
End Sub

Private Sub pbcIndependent_Paint()
    pbcIndependent.Cls
    pbcIndependent.CurrentX = fgBoxInsetX
    pbcIndependent.CurrentY = 0 'fgBoxInsetY
    pbcIndependent.Print smIndependent
End Sub

Private Sub pbcOversell_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        smOversell = "Yes"
        pbcOversell_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        smOversell = "No"
        pbcOversell_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smOversell = "Yes" Then
            smOversell = "No"
            pbcOversell_Paint
        ElseIf smOversell = "No" Then
            smOversell = "Yes"
            pbcOversell_Paint
        End If
    End If
End Sub

Private Sub pbcOversell_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smOversell = "Yes" Then
        smOversell = "No"
        pbcOversell_Paint
    ElseIf smOversell = "No" Then
        smOversell = "Yes"
        pbcOversell_Paint
    End If
End Sub

Private Sub pbcOversell_Paint()
    pbcOversell.Cls
    pbcOversell.CurrentX = fgBoxInsetX
    pbcOversell.CurrentY = 0 'fgBoxInsetY
    pbcOversell.Print smOversell
End Sub

Private Sub pbcSetSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSetSTab.hWnd Then
        Exit Sub
    End If
    If imSetCtrlVisible Then
        Do
            ilNext = False
            Select Case grdSet.Row
                Case SETROW3INDEX
                    Select Case grdSet.Col
                        Case SETUNITINDEX
                            mSetSetShow
                            cmcSetInv.SetFocus
                            Exit Sub
                        Case Else
                            grdSet.Col = grdSet.Col - 2
                    End Select
                Case SETROW6INDEX
                    Select Case grdSet.Col
                        Case SETLANGUAGEINDEX
                            grdSet.Row = SETROW3INDEX
                            grdSet.Col = SETRATEINDEX
                        Case SETFEEDSOURCEINDEX
                            grdSet.Col = SETLANGUAGEINDEX
                    End Select
                Case SETROW9INDEX
                    Select Case grdSet.Col
                        Case SETGAMENOSINDEX
                            grdSet.Row = SETROW6INDEX
                            grdSet.Col = SETFEEDSOURCEINDEX
                        Case Else
                            grdSet.Col = grdSet.Col - 2
                    End Select
            End Select
            If mSetColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetSetShow
    Else
        grdSet.Row = SETROW3INDEX '+1 to bypass title
        grdSet.Col = grdSet.FixedCols
        Do
            If mSetColOk() Then
                Exit Do
            Else
                grdSet.Col = grdSet.Col + 2
            End If
        Loop
    End If
    mSetEnableBox
End Sub

Private Sub pbcSetTab_GotFocus()
    Dim ilNext As Integer
    Dim llSpecEnableRow As Long
    Dim llSpecEnableCol As Long

    If GetFocus() <> pbcSetTab.hWnd Then
        Exit Sub
    End If
    If imSetCtrlVisible Then
        If (grdSet.Row = SETROW9INDEX) And (grdSet.Col = SETGAMENOSINDEX) Then
            llSpecEnableRow = lmSetEnableRow
            llSpecEnableCol = lmSetEnableCol
            mSetSetShow
            lmSetEnableRow = llSpecEnableRow
            lmSetEnableCol = llSpecEnableCol
        End If
        Do
            ilNext = False
            Select Case grdSet.Row
                Case SETROW3INDEX
                    Select Case grdSet.Col
                        Case SETRATEINDEX
                            grdSet.Row = SETROW6INDEX
                            grdSet.Col = SETLANGUAGEINDEX
                        Case Else
                            grdSet.Col = grdSet.Col + 2
                    End Select
                Case SETROW6INDEX
                    Select Case grdSet.Col
                        Case SETFEEDSOURCEINDEX
                            grdSet.Row = SETROW9INDEX
                            grdSet.Col = SETGAMENOSINDEX
                        Case SETLANGUAGEINDEX
                            grdSet.Col = SETFEEDSOURCEINDEX
                    End Select
                Case SETROW9INDEX
                    Select Case grdSet.Col
                        Case SETGAMEOUTINDEX
                            mSetSetShow
                            cmcSetInv.SetFocus
                            Exit Sub
                        Case Else
                            grdSet.Col = grdSet.Col + 2
                    End Select
            End Select
            If mSetColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetSetShow
    Else
        grdSet.Row = grdSet.Rows - 2
        grdSet.Col = grdSet.Cols - 2
        Do
            If mSetColOk() Then
                Exit Do
            Else
                grdSet.Col = grdSet.Col - 2
            End If
        Loop
    End If
    mSetEnableBox
End Sub

Private Sub pbcFeedSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        smFeedSource = "All"
        pbcFeedSource_Paint
    'ElseIf (KeyAscii = Asc("H")) Or (KeyAscii = Asc("h")) Then
    ElseIf (KeyAscii = Asc(UCase(Left(smEventTitle2, 1)))) Or (KeyAscii = Asc(LCase(Left(smEventTitle2, 1)))) Then
        smFeedSource = smEventTitle2    '"Home"
        pbcFeedSource_Paint
    'ElseIf KeyAscii = Asc("V") Or (KeyAscii = Asc("v")) Then
    ElseIf (KeyAscii = Asc(UCase(Left(smEventTitle1, 1)))) Or (KeyAscii = Asc(LCase(Left(smEventTitle1, 1)))) Then
        smFeedSource = smEventTitle1    '"Visiting"
        pbcFeedSource_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        smFeedSource = "National"
        pbcFeedSource_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smFeedSource = "All" Then
            smFeedSource = smEventTitle2    '"Home"
            pbcFeedSource_Paint
        'ElseIf smFeedSource = "Home" Then
        ElseIf smFeedSource = smEventTitle2 Then
            smFeedSource = smEventTitle1    '"Visiting"
            pbcFeedSource_Paint
        'ElseIf smFeedSource = "Visiting" Then
        ElseIf smFeedSource = smEventTitle1 Then
            smFeedSource = "National"
            pbcFeedSource_Paint
        ElseIf smFeedSource = "National" Then
            smFeedSource = "All"    '"Home"
            pbcFeedSource_Paint
        End If
    End If
End Sub

Private Sub pbcFeedSource_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smFeedSource = "All" Then
        smFeedSource = smEventTitle2    '"Home"
        pbcFeedSource_Paint
    'ElseIf smFeedSource = "Home" Then
    ElseIf smFeedSource = smEventTitle2 Then
        smFeedSource = smEventTitle1    '"Visiting"
        pbcFeedSource_Paint
    'ElseIf smFeedSource = "Visiting" Then
    ElseIf smFeedSource = smEventTitle1 Then
        smFeedSource = "National"
        pbcFeedSource_Paint
    ElseIf smFeedSource = "National" Then
        smFeedSource = "All"    '"Home"
        pbcFeedSource_Paint
    End If
End Sub

Private Sub pbcFeedSource_Paint()
    pbcFeedSource.Cls
    pbcFeedSource.CurrentX = fgBoxInsetX
    pbcFeedSource.CurrentY = 0 'fgBoxInsetY
    pbcFeedSource.Print smFeedSource
End Sub

Private Sub pbcSTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilBox                         slStr                                                   *
'******************************************************************************************

    Dim ilTestValue As Integer
    Dim ilNext As Integer

    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        imTabDirection = -1 'Set- Right to left
        ilTestValue = True
        Do
            ilNext = False
            Select Case grdDates.Col
                Case UNITINDEX
                    If grdDates.Row = grdDates.FixedRows Then
                        mSetShow
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    grdDates.Row = grdDates.Row - 2
                    imSettingValue = True
                    If Not grdDates.RowIsVisible(grdDates.Row) Then
                        grdDates.TopRow = grdDates.TopRow - 1
                    End If
                    If grdDates.RowHeight(grdDates.TopRow) <= 15 Then
                        grdDates.TopRow = grdDates.TopRow - 1
                    End If
                    imSettingValue = False
                    grdDates.Col = RATEINDEX
                Case Else
                    grdDates.Col = grdDates.Col - 2
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        imTabDirection = 0  'Set-Left to right
        lmTopRow = -1
        grdDates.TopRow = grdDates.FixedRows
        grdDates.Row = grdDates.FixedRows
        grdDates.Col = UNITINDEX
        Do
            If mColOk() Then
                Exit Do
            End If
            If grdDates.Row + 2 >= grdDates.Rows Then
                cmcDone.SetFocus
                Exit Sub
            End If
            grdDates.Row = grdDates.Row + 2
            Do
                If Not grdDates.RowIsVisible(grdDates.Row) Then
                    grdDates.TopRow = grdDates.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
        Loop
    End If
    lmTopRow = grdDates.TopRow
    mEnableBox
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

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        imTabDirection = 0 'Set- Left to right
        ilTestValue = True
        Do
            ilNext = False
            Select Case grdDates.Col
                Case RATEINDEX
                    llEnableRow = lmEnableRow
                    llEnableCol = lmEnableCol
                    mSetShow
                    lmEnableRow = llEnableRow
                    lmEnableCol = llEnableCol
                    If mGridFieldsOk(CInt(lmEnableRow)) = False Then
                        mEnableBox
                        Exit Sub
                    End If
                    If grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = "Yes" Then
                        mSetShow
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    If grdDates.Row + 2 > grdDates.Rows - 1 Then
                        mSetShow
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    slStr = grdDates.TextMatrix(grdDates.Row + 2, GAMENOINDEX)
                    If slStr = "" Then
                        mSetShow
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    grdDates.Row = grdDates.Row + 2
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
                    grdDates.Col = UNITINDEX
                Case Else
                    grdDates.Col = grdDates.Col + 2
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        imTabDirection = -1  'Set-Right to left
        imSettingValue = True
        lmTopRow = -1
        grdDates.TopRow = grdDates.FixedRows
        grdDates.Col = RATEINDEX
        Do
            If grdDates.Row <= grdDates.FixedRows Then
                cmcDone.SetFocus
                Exit Sub
            End If
            grdDates.Row = grdDates.Rows - 2
            Do
                If Not grdDates.RowIsVisible(grdDates.Row) Then
                    grdDates.TopRow = grdDates.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
            If mColOk() Then
                Exit Do
            End If
        Loop
        imSettingValue = False
    End If
    lmTopRow = grdDates.TopRow
    mEnableBox
End Sub


Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
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
'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilLang As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String

    ilRet = gPopMnfPlusFieldsBox(GameInv, lbcLanguage, tmLanguageCode(), smLanguageCodeTag, "L")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLanguagePopErr
        gCPErrorMsg ilRet, "mLanguagePop (gPopMnfPlusFieldsBox)", GameInv
        On Error GoTo 0
    End If
    lbcLanguage.Clear
    For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
        For ilLang = 0 To UBound(tmLanguageCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            slNameCode = tmLanguageCode(ilLang).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tmGsf(ilLoop).iLangMnfCode = Val(slCode) Then
                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                gFindMatch slStr, 0, lbcLanguage
                If gLastFound(lbcLanguage) < 0 Then
                    lbcLanguage.AddItem slStr
                    lbcLanguage.ItemData(lbcLanguage.NewIndex) = slCode
                End If
                Exit For
            End If
        Next ilLang
    Next ilLoop
    lbcLanguage.AddItem "[All]", 0
    lbcLanguage.ItemData(0) = "0"
    Exit Sub
mLanguagePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub




'*******************************************************
'*                                                     *
'*      Procedure Name:mGhfGsfReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mGhfGsfReadRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mGhfGsfReadRecErr                                                                     *
'******************************************************************************************

'
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer

    ReDim tmGsf(0 To 0) As GSF
    ilUpper = 0
    'tmGhfSrchKey1.iVefCode = imVefCode
    'ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    tmGhfSrchKey0.lCode = lmSeasonGhfCode
    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        tmGsfSrchKey1.lghfcode = tmGhf.lCode
        tmGsfSrchKey1.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf(ilUpper), imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.lCode = tmGsf(ilUpper).lghfcode)
            ReDim Preserve tmGsf(0 To UBound(tmGsf) + 1) As GSF
            ilUpper = UBound(tmGsf)
            ilRet = btrGetNext(hmGsf, tmGsf(ilUpper), imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        Loop
    Else
        mGhfGsfReadRec = False
        Exit Function
    End If
    mGhfGsfReadRec = True
    Exit Function
mGhfGsfReadRecErr: 'VBC NR
    On Error GoTo 0
    mGhfGsfReadRec = False
    Exit Function
End Function


Private Sub mClearCtrlFields()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                         ilCol                                                   *
'******************************************************************************************



    ReDim tmIsf(0 To 0) As ISF

    grdGet.TextMatrix(GETROW3INDEX, GETTYPEINDEX) = ""
    grdGet.TextMatrix(GETROW3INDEX, GETITEMINDEX) = ""
    grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = ""
    grdGet.TextMatrix(GETROW3INDEX, GETOVERSELLINDEX) = ""
    grdGet.Col = GETTYPEINDEX
    grdGet.Row = GETROW3INDEX
    grdGet.CellBackColor = vbWhite
    grdGet.Col = GETTYPEINDEX
    grdGet.Row = GETROW3INDEX - 1
    grdGet.CellBackColor = vbWhite
    grdGet.Col = GETITEMINDEX
    grdGet.Row = GETROW3INDEX
    grdGet.CellBackColor = vbWhite
    grdGet.Col = GETITEMINDEX
    grdGet.Row = GETROW3INDEX - 1
    grdGet.CellBackColor = vbWhite

    grdSet.TextMatrix(SETROW3INDEX, SETUNITINDEX) = ""
    grdSet.TextMatrix(SETROW3INDEX, SETCOSTINDEX) = ""
    grdSet.TextMatrix(SETROW3INDEX, SETRATEINDEX) = ""

    grdSet.TextMatrix(SETROW6INDEX, SETLANGUAGEINDEX) = " "
    grdSet.TextMatrix(SETROW6INDEX, SETLANGUAGEINDEX + 1) = " "
    grdSet.TextMatrix(SETROW6INDEX, SETLANGUAGEINDEX + 2) = " "
    grdSet.TextMatrix(SETROW6INDEX, SETFEEDSOURCEINDEX) = ""

    grdSet.TextMatrix(SETROW9INDEX, SETGAMENOSINDEX) = ""
    grdSet.TextMatrix(SETROW9INDEX, SETGAMEININDEX) = ""
    grdSet.TextMatrix(SETROW9INDEX, SETGAMEOUTINDEX) = ""

    mGridColumnWidths


'Moved to mClearGrdDates
'    If grdDates.Rows > 31 Then
'        For ilRow = grdDates.Rows - 1 To 31 Step -1
'            grdDates.RemoveItem ilRow
'        Next ilRow
'    End If
'    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
'        grdDates.TextMatrix(ilRow, TMISFINDEX) = "0"
'        grdDates.TextMatrix(ilRow, UNITINDEX) = ""
'        grdDates.TextMatrix(ilRow, COSTINDEX) = ""
'        grdDates.TextMatrix(ilRow, RATEINDEX) = ""
'    Next ilRow
'    For ilCol = GAMENOINDEX To AIRTIMEINDEX Step 2
'        grdDates.Col = ilCol
'        For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
'            grdDates.Row = ilRow
'            grdDates.CellBackColor = LIGHTYELLOW
'        Next ilRow
'    Next ilCol
'    For ilCol = UNITINDEX To RATEINDEX Step 2
'        grdDates.Col = ilCol
'        For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
'            grdDates.Row = ilRow
'            grdDates.CellBackColor = vbWhite
'        Next ilRow
'    Next ilCol
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
'
'   mSetCommands
'   Where:
'
    'Update button set if all mandatory fields have data and any field altered
    If imIhfChg Or imIsfChg Then
        cmcSave.Enabled = True
        cbcGameVeh.Enabled = False
        cbcSelect.Enabled = False
    Else
        cmcSave.Enabled = False
        cbcGameVeh.Enabled = True
        cbcSelect.Enabled = True
    End If
    'Erase button set if any field contains a value or change mode
    If (imSelectedIndex > 0) And (imUpdateAllowed) Then
        cmcErase.Enabled = True
    Else
        cmcErase.Enabled = False
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
Private Sub mSetEnableBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLang                        slNameCode                    slCode                    *
'*  ilCode                        ilRet                         ilLoop                    *
'*                                                                                        *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String

    If (grdSet.Row < grdSet.FixedRows) Or (grdSet.Row >= grdSet.Rows) Or (grdSet.Col < grdSet.FixedCols) Or (grdSet.Col >= grdSet.Cols - 1) Then
        Exit Sub
    End If
    lmSetEnableRow = grdSet.Row
    lmSetEnableCol = grdSet.Col

    Select Case grdSet.Row
        Case SETROW3INDEX
            Select Case grdSet.Col
                Case SETUNITINDEX
                    edcSet.MaxLength = 4
                    edcSet.Text = grdSet.Text
                Case SETCOSTINDEX
                    edcSet.MaxLength = 10
                    edcSet.Text = grdSet.Text
                Case SETRATEINDEX
                    edcSet.MaxLength = 10
                    edcSet.Text = grdSet.Text
            End Select
        Case SETROW6INDEX
            Select Case grdSet.Col
                Case SETLANGUAGEINDEX
                    lbcLanguage.Height = gListBoxHeight(lbcLanguage.ListCount, 10)
                    edcSet.MaxLength = 20
                    imChgMode = True
                    slStr = Trim$(grdSet.Text)
                    gFindMatch slStr, 0, lbcLanguage
                    If gLastFound(lbcLanguage) >= 0 Then
                        lbcLanguage.ListIndex = gLastFound(lbcLanguage)
                        edcSet.Text = lbcInvType.List(lbcLanguage.ListIndex)
                    Else
                        If lbcLanguage.ListCount > 0 Then
                            lbcLanguage.ListIndex = 0
                            edcSet.Text = lbcLanguage.List(lbcLanguage.ListIndex)
                        Else
                            edcSet.Text = ""
                        End If
                    End If
                    imChgMode = False
                Case SETFEEDSOURCEINDEX
                    smFeedSource = Trim$(grdSet.Text)
                    If (smFeedSource = "") Or (smFeedSource = "Missing") Then
                        smFeedSource = "All"
                    End If
                    pbcFeedSource_Paint
            End Select
        Case SETROW9INDEX
            Select Case grdSet.Col
                Case SETGAMENOSINDEX
                    edcSet.MaxLength = 0
                    If grdSet.Text = "" Then
                        edcSet.Text = smDefaultGameNo   'Trim$(str$(imMinGameNo)) & "-" & Trim$(str$(imMaxGameNo))
                    Else
                        edcSet.Text = grdSet.Text
                    End If
                Case SETGAMEININDEX
                    edcSet.MaxLength = 0
                    edcSet.Text = grdSet.Text
                Case SETGAMEOUTINDEX
                    edcSet.MaxLength = 0
                    edcSet.Text = grdSet.Text
            End Select
    End Select
    mSetSetFocus
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
Private Sub mSetSetShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilNoGames                     slStr                         ilOrigUpper               *
'*  ilLoop                        llRow                         llSvRow                   *
'*  llSvCol                                                                               *
'******************************************************************************************


    pbcArrow.Visible = False
    If (lmSetEnableRow >= grdSet.FixedRows) And (lmSetEnableRow < grdSet.Rows) Then
        Select Case lmSetEnableRow
            Case SETROW3INDEX
                Select Case lmSetEnableCol
                    Case SETUNITINDEX
                        edcSet.Visible = False
                        grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = edcSet.Text
                    Case SETCOSTINDEX
                        edcSet.Visible = False
                        grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = edcSet.Text
                    Case SETRATEINDEX
                        edcSet.Visible = False
                        grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = edcSet.Text
                End Select
            Case SETROW6INDEX
                Select Case lmSetEnableCol
                    Case SETLANGUAGEINDEX
                        edcSet.Visible = False
                        cmcSet.Visible = False
                        lbcLanguage.Visible = False
                        If lbcLanguage.ListIndex >= 0 Then
                            grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = lbcLanguage.List(lbcLanguage.ListIndex)
                            grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol + 1) = lbcLanguage.List(lbcLanguage.ListIndex)
                            grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol + 2) = lbcLanguage.List(lbcLanguage.ListIndex)
                        Else
                            grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = " "
                            grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol + 1) = " "
                            grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol + 2) = " "
                        End If
                    Case SETFEEDSOURCEINDEX
                        pbcFeedSource.Visible = False
                        grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = smFeedSource
                End Select
            Case SETROW9INDEX
                Select Case lmSetEnableCol
                    Case SETGAMENOSINDEX
                        edcSet.Visible = False
                        grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = edcSet.Text
                    Case SETGAMEININDEX
                        edcSet.Visible = False
                        grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = edcSet.Text
                    Case SETGAMEOUTINDEX
                        edcSet.Visible = False
                        grdSet.TextMatrix(lmSetEnableRow, lmSetEnableCol) = edcSet.Text
                End Select
        End Select
    End If
    lmSetEnableRow = -1
    lmSetEnableCol = -1
    imSetCtrlVisible = False
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
    Select Case grdDates.Col
        Case UNITINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case COSTINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case RATEINDEX
            edcDropdown.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            edcDropdown.Visible = True
            edcDropdown.SetFocus
    End Select
    mSetCommands
End Sub

Private Function mColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************


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
    If grdDates.CellForeColor = vbRed Then
        mColOk = False
        Exit Function
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
Private Sub mSetSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long

    If (grdSet.Row < grdSet.FixedRows) Or (grdSet.Row >= grdSet.Rows) Or (grdSet.Col < grdSet.FixedCols) Or (grdSet.Col >= grdSet.Cols - 1) Then
        Exit Sub
    End If
    imSetCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdSet.Col - 1 Step 1
        llColPos = llColPos + grdSet.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdSet.ColWidth(grdSet.Col)
    ilCol = grdSet.Col
    Do While ilCol < grdSet.Cols - 1
        If (Trim$(grdSet.TextMatrix(grdSet.Row - 1, grdSet.Col)) <> "") And (Trim$(grdSet.TextMatrix(grdSet.Row - 1, grdSet.Col)) = Trim$(grdSet.TextMatrix(grdSet.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdSet.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdSet.Row
        Case SETROW3INDEX
            Select Case grdSet.Col
                Case SETUNITINDEX
                    edcSet.Move grdSet.Left + llColPos + 30, grdSet.Top + grdSet.RowPos(grdSet.Row) + 15, grdSet.ColWidth(grdSet.Col), grdSet.RowHeight(grdSet.Row) - 15
                    edcSet.Visible = True
                    edcSet.SetFocus
                Case SETCOSTINDEX
                    edcSet.Move grdSet.Left + llColPos + 30, grdSet.Top + grdSet.RowPos(grdSet.Row) + 15, grdSet.ColWidth(grdSet.Col), grdSet.RowHeight(grdSet.Row) - 15
                    edcSet.Visible = True
                    edcSet.SetFocus
                Case SETRATEINDEX
                    edcSet.Move grdSet.Left + llColPos + 30, grdSet.Top + grdSet.RowPos(grdSet.Row) + 15, grdSet.ColWidth(grdSet.Col), grdSet.RowHeight(grdSet.Row) - 15
                    edcSet.Visible = True
                    edcSet.SetFocus
            End Select
        Case SETROW6INDEX
            Select Case grdSet.Col
                Case SETLANGUAGEINDEX
                    edcSet.Move grdSet.Left + llColPos + 30, grdSet.Top + grdSet.RowPos(grdSet.Row) + 15, grdSet.ColWidth(grdSet.Col) - cmcSet.Width, grdSet.RowHeight(grdSet.Row) - 15
                    cmcSet.Move edcSet.Left + edcSet.Width, edcSet.Top, cmcSet.Width, edcSet.Height
                    lbcLanguage.Move edcSet.Left, edcSet.Top + edcSet.Height, edcSet.Width + edcSet.Width
                    edcSet.Visible = True
                    cmcSet.Visible = True
                    edcSet.SetFocus
                Case SETFEEDSOURCEINDEX
                    pbcFeedSource.Move grdSet.Left + llColPos + 30, grdSet.Top + grdSet.RowPos(grdSet.Row) + 30, grdSet.ColWidth(grdSet.Col), grdSet.RowHeight(grdSet.Row)
                    pbcFeedSource_Paint
                    pbcFeedSource.Visible = True
                    pbcFeedSource.SetFocus
            End Select
        Case SETROW9INDEX
            Select Case grdSet.Col
                Case SETGAMENOSINDEX
                    edcSet.Move grdSet.Left + llColPos + 30, grdSet.Top + grdSet.RowPos(grdSet.Row) + 15, grdSet.ColWidth(grdSet.Col), grdSet.RowHeight(grdSet.Row) - 15
                    edcSet.Visible = True
                    edcSet.SetFocus
                Case SETGAMEININDEX
                    edcSet.Move grdSet.Left + llColPos + 30, grdSet.Top + grdSet.RowPos(grdSet.Row) + 15, grdSet.ColWidth(grdSet.Col), grdSet.RowHeight(grdSet.Row) - 15
                    edcSet.Visible = True
                    edcSet.SetFocus
                Case SETGAMEOUTINDEX
                    edcSet.Move grdSet.Left + llColPos + 30, grdSet.Top + grdSet.RowPos(grdSet.Row) + 15, grdSet.ColWidth(grdSet.Col), grdSet.RowHeight(grdSet.Row) - 15
                    edcSet.Visible = True
                    edcSet.SetFocus
            End Select
    End Select
End Sub

Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    'Layout Fixed Rows:0=>Edge; 1=>Blue border; 2=>Column Title 1; 3=Column Title 2; 4=>Blue border
    '       Rows: 5=>input; 6=>blue row line; 7=>Input; 8=>blue row line
    'Layout Fixed Columns: 0=>Edge; 1=Blue border; 2=>Row Title; 3=>Blue border   Note:  This was done this way to allow for horizontal scrolling:  It is not used
    '       Columns: 4=>Input; 5=>Blue column line; 6=>Input; 7=>Blue Column;....
    grdDates.RowHeight(0) = 15
    grdDates.RowHeight(1) = 15
    grdDates.RowHeight(2) = 180
    grdDates.RowHeight(3) = 180
    grdDates.RowHeight(4) = 15
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        grdDates.RowHeight(ilRow) = fgBoxGridH
        grdDates.RowHeight(ilRow + 1) = 15
    Next ilRow

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

Private Sub mGridSetLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilCol = 0 To grdSet.Cols - 1 Step 1
        grdSet.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
    grdSet.RowHeight(0) = 15
    grdSet.RowHeight(1) = 15
    grdSet.RowHeight(2) = 150
    grdSet.RowHeight(SETROW3INDEX) = fgBoxGridH
    grdSet.RowHeight(4) = 15
    grdSet.RowHeight(5) = 150
    grdSet.RowHeight(SETROW6INDEX) = fgBoxGridH
    grdSet.RowHeight(7) = 15
    grdSet.RowHeight(8) = 150
    grdSet.RowHeight(SETROW9INDEX) = fgBoxGridH
    grdSet.RowHeight(10) = 15
    grdSet.ColWidth(0) = 15
    grdSet.ColWidth(1) = 15
    grdSet.ColWidth(3) = 15
    grdSet.ColWidth(5) = 15
    grdSet.ColWidth(7) = 15
    'Horizontal
    For ilCol = 1 To grdSet.Cols - 1 Step 1
        grdSet.Row = 1
        grdSet.Col = ilCol
        grdSet.CellBackColor = vbBlue
    Next ilCol
    For ilRow = grdSet.FixedRows + 2 To grdSet.Rows - 1 Step 3
        For ilCol = 1 To grdSet.Cols - 1 Step 1
            grdSet.Row = ilRow
            grdSet.Col = ilCol
            grdSet.CellBackColor = vbBlue
        Next ilCol
    Next ilRow
    'Vertical Line
    For ilRow = 1 To grdSet.Rows - 1 Step 1
        grdSet.Row = ilRow
        grdSet.Col = 1
        grdSet.CellBackColor = vbBlue
    Next ilRow
    For ilCol = grdSet.FixedCols + 1 To grdSet.Cols - 1 Step 2
        For ilRow = 1 To grdSet.Rows - 1 Step 1
            grdSet.Row = ilRow
            grdSet.Col = ilCol
            grdSet.CellBackColor = vbBlue
        Next ilRow
    Next ilCol
End Sub

Private Sub mGridColumns()
        
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
    'Teams
    grdDates.Row = 2
    grdDates.Col = TEAMSINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, TEAMSINDEX) = "Event"  '"Teams"
    grdDates.Row = 3
    grdDates.Col = TEAMSINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, TEAMSINDEX) = "Name" '""
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
    'Units
    grdDates.Row = 2
    grdDates.Col = UNITINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, UNITINDEX) = "#"
    grdDates.Row = 3
    grdDates.Col = UNITINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, UNITINDEX) = "Units"
    'Cost
    grdDates.Row = 2
    grdDates.Col = COSTINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, COSTINDEX) = "Cost"
    grdDates.Row = 3
    grdDates.Col = COSTINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, COSTINDEX) = ""
    'Rate
    grdDates.Row = 2
    grdDates.Col = RATEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(2, RATEINDEX) = "Rate"
    grdDates.Row = 3
    grdDates.Col = RATEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE
    grdDates.TextMatrix(3, RATEINDEX) = ""
End Sub

Private Sub mGridColumnWidths()
    Dim ilValue As Integer
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdDates.ColWidth(TMISFINDEX) = 0
    grdDates.ColWidth(SORTINDEX) = 0
    grdDates.ColWidth(GAMESTATUSINDEX) = 0
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    'Game number
    '6/30/12:  Allow 5 digit event #'s
    grdDates.ColWidth(GAMENOINDEX) = 0.057 * grdDates.Width
    'Feed Source
    If (ilValue And USINGFEED) = USINGFEED Then
        grdDates.ColWidth(FEEDSOURCEINDEX) = 0.057 * grdDates.Width
    Else
        grdDates.ColWidth(FEEDSOURCEINDEX) = 0
        grdDates.ColWidth(FEEDSOURCEINDEX + 1) = 0
    End If
    'Language
    If (ilValue And USINGLANG) = USINGLANG Then
        grdDates.ColWidth(LANGUAGEINDEX) = 0.078 * grdDates.Width
    Else
        grdDates.ColWidth(LANGUAGEINDEX) = 0
        grdDates.ColWidth(LANGUAGEINDEX + 1) = 0
    End If
    'Teams
    grdDates.ColWidth(TEAMSINDEX) = 0.081 * grdDates.Width
    'Air Date
    grdDates.ColWidth(AIRDATEINDEX) = 0.063 * grdDates.Width
    'Air Time
    grdDates.ColWidth(AIRTIMEINDEX) = 0.083 * grdDates.Width
    'Units
    grdDates.ColWidth(UNITINDEX) = 0.051 * grdDates.Width
    'Cost
    grdDates.ColWidth(COSTINDEX) = 0.1 * grdDates.Width
    'Rate
    grdDates.ColWidth(RATEINDEX) = 0.1 * grdDates.Width
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

Private Sub mGridSetColumns()
    Dim ilCol As Integer
    Dim ilValue As Integer

    grdSet.Row = SETROW3INDEX - 1
    grdSet.Col = SETUNITINDEX
    grdSet.CellFontBold = False
    grdSet.CellFontName = "Arial"
    grdSet.CellFontSize = 6.75
    grdSet.CellForeColor = vbBlue
    grdSet.TextMatrix(grdSet.Row, grdSet.Col) = "Units"
    grdSet.Row = SETROW3INDEX - 1
    grdSet.Col = SETCOSTINDEX
    grdSet.CellFontBold = False
    grdSet.CellFontName = "Arial"
    grdSet.CellFontSize = 6.75
    grdSet.CellForeColor = vbBlue
    grdSet.TextMatrix(grdSet.Row, grdSet.Col) = "Cost"
    grdSet.Row = SETROW3INDEX - 1
    grdSet.Col = SETRATEINDEX
    grdSet.CellFontBold = False
    grdSet.CellFontName = "Arial"
    grdSet.CellFontSize = 6.75
    grdSet.CellForeColor = vbBlue
    grdSet.TextMatrix(grdSet.Row, grdSet.Col) = "Rate"

    grdSet.Row = SETROW6INDEX - 1
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    If (ilValue And USINGLANG) = USINGLANG Then
        For ilCol = SETLANGUAGEINDEX To SETLANGUAGEINDEX + 2 Step 1
            grdSet.Col = ilCol
            grdSet.CellFontName = "Arial"
            grdSet.CellFontSize = 6.75
            grdSet.CellForeColor = vbBlue
            grdSet.TextMatrix(grdSet.Row, ilCol) = "Language"
        Next ilCol
    Else
        For ilCol = SETLANGUAGEINDEX To SETLANGUAGEINDEX + 2 Step 1
            grdSet.Col = ilCol
            grdSet.CellFontName = "Arial"
            grdSet.CellFontSize = 6.75
            grdSet.CellForeColor = vbBlue
            grdSet.TextMatrix(grdSet.Row, ilCol) = " "
        Next ilCol
    End If

    grdSet.Row = SETROW6INDEX - 1
    grdSet.Col = SETFEEDSOURCEINDEX
    grdSet.CellFontName = "Arial"
    grdSet.CellFontSize = 6.75
    grdSet.CellForeColor = vbBlue
    If (ilValue And USINGFEED) = USINGFEED Then
        grdSet.TextMatrix(grdSet.Row, grdSet.Col) = "Feed Source"
    Else
        grdSet.TextMatrix(grdSet.Row, grdSet.Col) = ""
    End If

    grdSet.Row = SETROW9INDEX - 1
    grdSet.Col = SETGAMENOSINDEX
    grdSet.CellFontBold = False
    grdSet.CellFontName = "Arial"
    grdSet.CellFontSize = 6.75
    grdSet.CellForeColor = vbBlue
    grdSet.TextMatrix(grdSet.Row, grdSet.Col) = "Events by #"
    grdSet.Col = SETGAMEININDEX
    grdSet.CellFontBold = False
    grdSet.CellFontName = "Arial"
    grdSet.CellFontSize = 6.75
    grdSet.CellForeColor = vbBlue
    grdSet.TextMatrix(grdSet.Row, grdSet.Col) = "Events In"
    grdSet.Col = SETGAMEOUTINDEX
    grdSet.CellFontBold = False
    grdSet.CellFontName = "Arial"
    grdSet.CellFontSize = 6.75
    grdSet.CellForeColor = vbBlue
    grdSet.TextMatrix(grdSet.Row, grdSet.Col) = "Events Out"
End Sub

Private Sub mGridSetColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdSet.ColWidth(SETGAMENOSINDEX) = 0.29 * grdSet.Width
    grdSet.ColWidth(SETGAMEININDEX) = 0.29 * grdSet.Width
    grdSet.ColWidth(SETGAMEOUTINDEX) = 0.29 * grdSet.Width
    llWidth = fgPanelAdj
    llMinWidth = grdSet.Width
    For ilCol = 0 To grdSet.Cols - 1 Step 1
        llWidth = llWidth + grdSet.ColWidth(ilCol)
        If (grdSet.ColWidth(ilCol) > 15) And (grdSet.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdSet.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdSet.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdSet.Width
            For ilCol = 0 To grdSet.Cols - 1 Step 1
                If (grdSet.ColWidth(ilCol) > 15) And (grdSet.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdSet.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdSet.FixedCols To grdSet.Cols - 1 Step 1
                If grdSet.ColWidth(ilCol) > 15 Then
                    ilColInc = grdSet.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdSet.ColWidth(ilCol) = grdSet.ColWidth(ilCol) + 15
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

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case grdGet.Row
        Case GETROW3INDEX
            Select Case lmGetEnableCol
                Case GETTYPEINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcInvType, edcGet, imChgMode, imLbcArrowSetting
                Case GETITEMINDEX
                    imLbcArrowSetting = False
                    gProcessLbcClick lbcInvItem, edcGet, imChgMode, imLbcArrowSetting
            End Select
    End Select
End Sub


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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llFind                        ilPos                                                   *
'******************************************************************************************

    Dim llRow As Long
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilItemNo As Integer
    Dim slTestStr As String
    Dim ilRet As Integer
    Dim slIndex As String
    Dim slIsfIndex As String

    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        slStr = Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX))
        If slStr <> "" Then
            If slStr <> "Ind" Then
                ilItemNo = 1
                Do
                    slStr = Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX))
                    ilRet = gParseItem(slStr, ilItemNo, ",", slTestStr)
                    If slTestStr = "" Then
                        Exit Do
                    End If
                    slIndex = grdDates.TextMatrix(llRow, TMISFINDEX)
                    ilRet = gParseItem(slIndex, ilItemNo, ",", slIsfIndex)
                    If (slIsfIndex = "") Or (Trim$(slIsfIndex) = "-1") Then
                        ilIndex = -1
                    Else
                        ilIndex = Val(slIsfIndex)
                    End If
                    If ilIndex = -1 Then
                        ilIndex = UBound(tmIsf)
                        ReDim Preserve tmIsf(0 To UBound(tmIsf) + 1) As ISF
                        tmIsf(ilIndex).lCode = 0
                    End If
                    tmIsf(ilIndex).iGameNo = Val(slTestStr)
                    If ilItemNo = 1 Then
                        slStr = Trim$(grdDates.TextMatrix(llRow, UNITINDEX))
                        tmIsf(ilIndex).iNoUnits = Val(slStr)
                    Else
                        tmIsf(ilIndex).iNoUnits = 0
                    End If
                    slStr = Trim$(grdDates.TextMatrix(llRow, COSTINDEX))
                    tmIsf(ilIndex).lCost = gStrDecToLong(slStr, 2)
                    slStr = Trim$(grdDates.TextMatrix(llRow, RATEINDEX))
                    tmIsf(ilIndex).lRate = gStrDecToLong(slStr, 2)
                    ilItemNo = ilItemNo + 1
                Loop While slTestStr <> ""
            Else
                slIndex = grdDates.TextMatrix(llRow, TMISFINDEX)
                If (slIndex = "") Or (Trim$(slIndex) = "-1") Then
                    ilIndex = -1
                Else
                    ilIndex = Val(slIndex)
                End If
                If ilIndex = -1 Then
                    ilIndex = UBound(tmIsf)
                    ReDim Preserve tmIsf(0 To UBound(tmIsf) + 1) As ISF
                    tmIsf(ilIndex).lCode = 0
                End If
                tmIsf(ilIndex).iGameNo = 0
                slStr = Trim$(grdDates.TextMatrix(llRow, UNITINDEX))
                tmIsf(ilIndex).iNoUnits = Val(slStr)
                slStr = Trim$(grdDates.TextMatrix(llRow, COSTINDEX))
                tmIsf(ilIndex).lCost = gStrDecToLong(slStr, 2)
                slStr = Trim$(grdDates.TextMatrix(llRow, RATEINDEX))
                tmIsf(ilIndex).lRate = gStrDecToLong(slStr, 2)
            End If
        End If
    Next llRow
    Exit Sub
End Sub




Private Function mSetColOk() As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilValue As Integer

    mSetColOk = True
    If grdSet.ColWidth(grdSet.Col) <= 15 Then
        mSetColOk = False
        Exit Function
    End If
    If grdSet.CellBackColor = LIGHTYELLOW Then
        mSetColOk = False
        Exit Function
    End If

    If (grdSet.Row = SETROW9INDEX) And ((grdSet.Col = SETGAMEININDEX) Or (grdSet.Col = SETGAMEOUTINDEX)) Then
        slStr = grdSet.TextMatrix(SETROW9INDEX, SETGAMENOSINDEX)
        ilPos = InStr(1, slStr, "-", vbTextCompare)
        If ilPos <= 0 Then
            mSetColOk = False
            Exit Function
        End If
    End If
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    If (grdSet.Row = SETROW6INDEX) And (grdSet.Col = SETLANGUAGEINDEX) And ((ilValue And USINGLANG) <> USINGLANG) Then
        mSetColOk = False
        Exit Function
    End If
    If (grdSet.Row = SETROW6INDEX) And (grdSet.Col = SETFEEDSOURCEINDEX) And ((ilValue And USINGFEED) <> USINGFEED) Then
        mSetColOk = False
        Exit Function
    End If
End Function




'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFormFieldsOk               *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mSetFormFieldsOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slLineNo                      ilLineNo                      ilLine                    *
'*                                                                                        *
'******************************************************************************************

'
'   iRet = mSetFormFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim ilError As Integer
    Dim ilValue As Integer
    Dim ilGame As Integer
    Dim ilLoop As Integer
    Dim ilGameNo As Integer
    Dim ilSearchFrom As Integer
    Dim ilNextComma As Integer
    Dim ilStartGame As Integer
    Dim ilEndGame As Integer
    Dim llRow As Long
    Dim ilPos As Integer

    ilError = False
    slStr = Trim$(grdSet.TextMatrix(SETROW3INDEX, SETUNITINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        grdSet.TextMatrix(SETROW3INDEX, SETUNITINDEX) = "Missing"
        grdSet.Row = SETROW3INDEX
        grdSet.Col = SETUNITINDEX
        grdSet.CellForeColor = vbMagenta
    End If
    slStr = Trim$(grdSet.TextMatrix(SETROW3INDEX, SETCOSTINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        grdSet.TextMatrix(SETROW3INDEX, SETCOSTINDEX) = "Missing"
        grdSet.Row = SETROW3INDEX
        grdSet.Col = SETCOSTINDEX
        grdSet.CellForeColor = vbMagenta
    End If
    slStr = Trim$(grdSet.TextMatrix(SETROW3INDEX, SETRATEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        grdSet.TextMatrix(SETROW3INDEX, SETRATEINDEX) = "Missing"
        grdSet.Row = SETROW3INDEX
        grdSet.Col = SETRATEINDEX
        grdSet.CellForeColor = vbMagenta
    End If
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    If (ilValue And USINGLANG) = USINGLANG Then
        slStr = grdSet.TextMatrix(SETROW6INDEX, SETLANGUAGEINDEX)
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdSet.TextMatrix(SETROW6INDEX, SETLANGUAGEINDEX) = "Missing"
            grdSet.TextMatrix(SETROW6INDEX, SETLANGUAGEINDEX + 1) = "Missing"
            grdSet.TextMatrix(SETROW6INDEX, SETLANGUAGEINDEX + 2) = "Missing"
            grdSet.Row = SETROW6INDEX
            grdSet.Col = SETLANGUAGEINDEX
            grdSet.CellForeColor = vbMagenta
        End If
    End If
    If (ilValue And USINGFEED) = USINGFEED Then
        slStr = grdSet.TextMatrix(SETROW6INDEX, SETFEEDSOURCEINDEX)
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdSet.TextMatrix(SETROW6INDEX, SETFEEDSOURCEINDEX) = "Missing"
            grdSet.Row = SETROW6INDEX
            grdSet.Col = SETFEEDSOURCEINDEX
            grdSet.CellForeColor = vbMagenta
        End If
    End If
    slStr = grdSet.TextMatrix(SETROW9INDEX, SETGAMENOSINDEX)
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        grdSet.TextMatrix(SETROW9INDEX, SETGAMENOSINDEX) = "Missing"
        grdSet.Row = SETROW9INDEX
        grdSet.Col = SETGAMENOSINDEX
        grdSet.CellForeColor = vbMagenta
    End If
    grdSet.Row = SETROW9INDEX
    grdSet.Col = SETGAMENOSINDEX
    If grdSet.CellForeColor <> vbMagenta Then
        slStr = grdSet.TextMatrix(SETROW9INDEX, SETGAMENOSINDEX)
        ReDim slFields(0 To 0) As String
        ilSearchFrom = 1
        Do
            ilNextComma = InStr(ilSearchFrom, slStr, ",", vbTextCompare)
            If ilNextComma > 0 Then
                slFields(UBound(slFields)) = Mid$(slStr, ilSearchFrom, ilNextComma - ilSearchFrom)
                ilSearchFrom = ilNextComma + 1
                ReDim Preserve slFields(0 To UBound(slFields) + 1) As String
            Else
                slFields(UBound(slFields)) = Mid$(slStr, ilSearchFrom)
                ReDim Preserve slFields(0 To UBound(slFields) + 1) As String
                Exit Do
            End If
        Loop While ilSearchFrom <= Len(slStr)
        For ilLoop = 0 To UBound(slFields) - 1 Step 1
            ilPos = InStr(1, slFields(ilLoop), "-", vbTextCompare)
            If ilPos > 0 Then
                ilStartGame = Left$(slFields(ilLoop), ilPos - 1)
                ilEndGame = Mid$(slFields(ilLoop), ilPos + 1)
                For ilGameNo = ilStartGame To ilEndGame Step 1
                    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                            ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                            If ilGameNo = ilGame Then
                                slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                                If gDateValue(slStr) <= lmNowDate Then
                                    ilError = True
                                    grdSet.Row = SETROW9INDEX
                                    grdSet.Col = SETGAMENOSINDEX
                                    grdSet.CellForeColor = vbMagenta
                                End If
                                Exit For
                            End If
                        End If
                    Next llRow
                Next ilGameNo
            Else
                For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                    If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                        ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                        If Val(slFields(ilLoop)) = ilGame Then
                            slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                            If gDateValue(slStr) <= lmNowDate Then
                                ilError = True
                                grdSet.Row = SETROW9INDEX
                                grdSet.Col = SETGAMENOSINDEX
                                grdSet.CellForeColor = vbMagenta
                            End If
                            Exit For
                        End If
                    End If
                Next llRow
            End If
        Next ilLoop
    Else
        ilError = True
    End If
    If ilError Then
        mSetFormFieldsOk = False
    Else
        mSetFormFieldsOk = True
    End If
End Function

Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    mGetSetShow
    mSetSetShow
    mSetShow
    lmEnableRow = -1
    lmEnableCol = -1
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
            ElseIf (ilCol = UNITINDEX) Then
                slSort = Trim$(grdDates.TextMatrix(llRow, UNITINDEX))
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = COSTINDEX) Then
                slSort = Trim$(grdDates.TextMatrix(llRow, COSTINDEX))
                slSort = Trim$(str$(gStrDecToLong(slSort, 2)))
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = RATEINDEX) Then
                slSort = Trim$(grdDates.TextMatrix(llRow, RATEINDEX))
                slSort = Trim$(str$(gStrDecToLong(slSort, 2)))
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = GAMENOINDEX) Then
                slSort = Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX))
                '6/30/12:  Allow 5 digit event #'s
                Do While Len(slSort) < 5    '4
                    slSort = "0" & slSort
                Loop
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

'*******************************************************
'*                                                     *
'*      Procedure Name:mInvTypePop                     *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mInvTypePop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    imPopReqd = False
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    ilRet = gIMoveListBox(InvItem, lbcInvType, tmInvType(), smInvTypeTag, "Itf.Btr", gFieldOffset("Itf", "ItfName"), 50, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mInvTypePopErr
        gCPErrorMsg ilRet, "mInvTypePop (gIMoveListBox)", GameInv
        On Error GoTo 0
        lbcInvType.AddItem "[New]", 0  'Force as first item on list
        imPopReqd = True
    End If
    Exit Sub
mInvTypePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInvItemPop                       *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mInvItemPop()
'
'   mInvItemPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    imPopReqd = False
    If imItfCode > 0 Then
        ilfilter(0) = INTEGERFILTER
        slFilter(0) = imItfCode
        ilOffSet(0) = gFieldOffset("Iif", "IifItfCode") '2
        ilRet = gIMoveListBox(InvItem, lbcInvItem, tmInvItem(), smInvItemTag, "Iif.Btr", gFieldOffset("Iif", "IifName"), 60, ilfilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mInvItemPopErr
            gCPErrorMsg ilRet, "mInvItemPop (gIMoveListBox)", GameInv
            On Error GoTo 0
            lbcInvItem.AddItem "[New]", 0  'Force as first item on list
            imPopReqd = True
        End If
    Else
        lbcInvItem.Clear
        lbcInvItem.AddItem "[New]", 0  'Force as first item on list
    End If
    Exit Sub
mInvItemPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mGridGetColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdGet.Row = GETROW3INDEX - 1
    grdGet.Col = GETTYPEINDEX
    grdGet.CellFontBold = False
    grdGet.CellFontName = "Arial"
    grdGet.CellFontSize = 6.75
    grdGet.CellForeColor = vbBlue
    grdGet.TextMatrix(grdGet.Row, grdGet.Col) = "Type"
    grdGet.Row = GETROW3INDEX - 1
    grdGet.Col = GETITEMINDEX
    grdGet.CellFontBold = False
    grdGet.CellFontName = "Arial"
    grdGet.CellFontSize = 6.75
    grdGet.CellForeColor = vbBlue
    grdGet.TextMatrix(grdGet.Row, grdGet.Col) = "Item"
    grdGet.Row = GETROW3INDEX - 1
    grdGet.Col = GETOVERSELLINDEX
    grdGet.CellFontBold = False
    grdGet.CellFontName = "Arial"
    grdGet.CellFontSize = 6.75
    grdGet.CellForeColor = vbBlue
    grdGet.TextMatrix(grdGet.Row, grdGet.Col) = "Oversell"
    grdGet.Row = GETROW3INDEX - 1
    grdGet.Col = GETINDEPENDENTINDEX
    grdGet.CellFontBold = False
    grdGet.CellFontName = "Arial"
    grdGet.CellFontSize = 6.75
    grdGet.CellForeColor = vbBlue
    grdGet.TextMatrix(grdGet.Row, grdGet.Col) = "Event-Independent"
End Sub

Private Sub mGridGetColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdGet.ColWidth(GETTYPEINDEX) = 0.25 * grdGet.Width
    grdGet.ColWidth(GETITEMINDEX) = 0.25 * grdGet.Width
    grdGet.ColWidth(GETOVERSELLINDEX) = 0.2 * grdGet.Width
    grdGet.ColWidth(GETINDEPENDENTINDEX) = 0.25 * grdGet.Width
    llWidth = fgPanelAdj
    llMinWidth = grdGet.Width
    For ilCol = 0 To grdGet.Cols - 1 Step 1
        llWidth = llWidth + grdGet.ColWidth(ilCol)
        If (grdGet.ColWidth(ilCol) > 15) And (grdGet.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdGet.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdGet.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdGet.Width
            For ilCol = 0 To grdGet.Cols - 1 Step 1
                If (grdGet.ColWidth(ilCol) > 15) And (grdGet.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdGet.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdGet.FixedCols To grdGet.Cols - 1 Step 1
                If grdGet.ColWidth(ilCol) > 15 Then
                    ilColInc = grdGet.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdGet.ColWidth(ilCol) = grdGet.ColWidth(ilCol) + 15
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

Private Sub mGridGetLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilCol = 0 To grdGet.Cols - 1 Step 1
        grdGet.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
    grdGet.RowHeight(0) = 15
    grdGet.RowHeight(1) = 15
    grdGet.RowHeight(2) = 150
    grdGet.RowHeight(SETROW3INDEX) = fgBoxGridH
    grdGet.RowHeight(4) = 15
    grdGet.ColWidth(0) = 15
    grdGet.ColWidth(1) = 15
    grdGet.ColWidth(3) = 15
    grdGet.ColWidth(5) = 15
    grdGet.ColWidth(7) = 15
    'Horizontal
    For ilCol = 1 To grdGet.Cols - 1 Step 1
        grdGet.Row = 1
        grdGet.Col = ilCol
        grdGet.CellBackColor = vbBlue
    Next ilCol
    For ilRow = grdGet.FixedRows + 2 To grdGet.Rows - 1 Step 3
        For ilCol = 1 To grdGet.Cols - 1 Step 1
            grdGet.Row = ilRow
            grdGet.Col = ilCol
            grdGet.CellBackColor = vbBlue
        Next ilCol
    Next ilRow
    'Vertical Line
    For ilRow = 1 To grdGet.Rows - 1 Step 1
        grdGet.Row = ilRow
        grdGet.Col = 1
        grdGet.CellBackColor = vbBlue
    Next ilRow
    For ilCol = grdGet.FixedCols + 1 To grdGet.Cols - 1 Step 2
        For ilRow = 1 To grdGet.Rows - 1 Step 1
            grdGet.Row = ilRow
            grdGet.Col = ilCol
            grdGet.CellBackColor = vbBlue
        Next ilRow
    Next ilCol
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
Private Sub mGetEnableBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLang                        slNameCode                    slCode                    *
'*  ilCode                        ilRet                         ilLoop                    *
'*                                                                                        *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String

    If (grdGet.Row < grdGet.FixedRows) Or (grdGet.Row >= grdGet.Rows) Or (grdGet.Col < grdGet.FixedCols) Or (grdGet.Col >= grdGet.Cols - 1) Then
        Exit Sub
    End If
    lmGetEnableRow = grdGet.Row
    lmGetEnableCol = grdGet.Col

    Select Case grdGet.Row
        Case SETROW3INDEX
            Select Case grdGet.Col
                Case GETTYPEINDEX
                    lbcInvType.Height = gListBoxHeight(lbcInvType.ListCount, 10)
                    edcGet.MaxLength = 50
                    imChgMode = True
                    slStr = grdGet.Text
                    gFindMatch slStr, 1, lbcInvType
                    If gLastFound(lbcInvType) >= 1 Then
                        lbcInvType.ListIndex = gLastFound(lbcInvType)
                        edcGet.Text = lbcInvType.List(lbcInvType.ListIndex)
                    Else
                        If smLastType <> "" Then
                            gFindMatch smLastType, 1, lbcInvType
                            If gLastFound(lbcInvType) >= 1 Then
                                lbcInvType.ListIndex = gLastFound(lbcInvType)
                                edcGet.Text = lbcInvType.List(lbcInvType.ListIndex)
                            Else
                                If lbcInvType.ListCount >= 1 Then
                                    lbcInvType.ListIndex = 0
                                    edcGet.Text = lbcInvType.List(lbcInvType.ListIndex)
                                Else
                                    edcGet.Text = ""
                                End If
                            End If
                        Else
                            If lbcInvType.ListCount >= 1 Then
                                lbcInvType.ListIndex = 0
                                edcGet.Text = lbcInvType.List(lbcInvType.ListIndex)
                            Else
                                edcGet.Text = ""
                            End If
                        End If
                    End If
                    imChgMode = False
                Case GETITEMINDEX
                    lbcInvItem.Height = gListBoxHeight(lbcInvItem.ListCount, 10)
                    edcGet.MaxLength = 60
                    imChgMode = True
                    slStr = grdGet.Text
                    gFindMatch slStr, 1, lbcInvItem
                    If gLastFound(lbcInvItem) >= 1 Then
                        lbcInvItem.ListIndex = gLastFound(lbcInvItem)
                        edcGet.Text = lbcInvItem.List(lbcInvItem.ListIndex)
                    Else
                        If lbcInvItem.ListCount >= 1 Then
                            lbcInvItem.ListIndex = 0
                            edcGet.Text = lbcInvItem.List(lbcInvItem.ListIndex)
                        Else
                            edcGet.Text = ""
                        End If
                    End If
                    imChgMode = False
                Case GETINDEPENDENTINDEX
                    smIndependent = Trim$(grdGet.Text)
                    If (smIndependent = "") Or (smIndependent = "Missing") Then
                        smIndependent = "No"
                    End If
                    pbcIndependent_Paint
                Case GETOVERSELLINDEX
                    smOversell = Trim$(grdGet.Text)
                    If (smOversell = "") Or (smOversell = "Missing") Then
                        smOversell = "No"
                    End If
                    pbcOversell_Paint
            End Select
    End Select
    mGetSetFocus
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
Private Sub mGetSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long

    If (grdGet.Row < grdGet.FixedRows) Or (grdGet.Row >= grdGet.Rows) Or (grdGet.Col < grdGet.FixedCols) Or (grdGet.Col >= grdGet.Cols - 1) Then
        Exit Sub
    End If
    imGetCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdGet.Col - 1 Step 1
        llColPos = llColPos + grdGet.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdGet.ColWidth(grdGet.Col)
    ilCol = grdGet.Col
    Do While ilCol < grdGet.Cols - 1
        If (Trim$(grdGet.TextMatrix(grdGet.Row - 1, grdGet.Col)) <> "") And (Trim$(grdGet.TextMatrix(grdGet.Row - 1, grdGet.Col)) = Trim$(grdGet.TextMatrix(grdGet.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdGet.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdGet.Row
        Case SETROW3INDEX
            Select Case grdGet.Col
                Case GETTYPEINDEX
                    edcGet.Move grdGet.Left + llColPos + 30, grdGet.Top + grdGet.RowPos(grdGet.Row) + 15, grdGet.ColWidth(grdGet.Col) - cmcGet.Width, grdGet.RowHeight(grdGet.Row) - 15
                    cmcGet.Move edcGet.Left + edcGet.Width, edcGet.Top, cmcGet.Width, edcGet.Height
                    lbcInvType.Move edcGet.Left, edcGet.Top + edcGet.Height, edcGet.Width + edcGet.Width
                    edcGet.Visible = True
                    cmcGet.Visible = True
                    edcGet.SetFocus
                Case GETITEMINDEX
                    edcGet.Move grdGet.Left + llColPos + 30, grdGet.Top + grdGet.RowPos(grdGet.Row) + 15, grdGet.ColWidth(grdGet.Col), grdGet.RowHeight(grdGet.Row) - 15
                    cmcGet.Move edcGet.Left + edcGet.Width, edcGet.Top, cmcGet.Width, edcGet.Height
                    lbcInvItem.Move edcGet.Left, edcGet.Top + edcGet.Height, edcGet.Width + edcGet.Width
                    edcGet.Visible = True
                    cmcGet.Visible = True
                    edcGet.SetFocus
                Case GETINDEPENDENTINDEX
                    pbcIndependent.Move grdGet.Left + llColPos + 30, grdGet.Top + grdGet.RowPos(grdGet.Row) + 30, grdGet.ColWidth(grdGet.Col), grdGet.RowHeight(grdGet.Row)
                    pbcIndependent_Paint
                    pbcIndependent.Visible = True
                    pbcIndependent.SetFocus
                Case GETOVERSELLINDEX
                    pbcOversell.Move grdGet.Left + llColPos + 30, grdGet.Top + grdGet.RowPos(grdGet.Row) + 30, grdGet.ColWidth(grdGet.Col), grdGet.RowHeight(grdGet.Row)
                    pbcOversell_Paint
                    pbcOversell.Visible = True
                    pbcOversell.SetFocus
            End Select
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mGetSetShow()
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    pbcArrow.Visible = False
    If (lmGetEnableRow >= grdGet.FixedRows) And (lmGetEnableRow < grdGet.Rows) Then
        Select Case lmGetEnableRow
            Case GETROW3INDEX
                Select Case lmGetEnableCol
                    Case GETTYPEINDEX
                        edcGet.Visible = False
                        cmcGet.Visible = False
                        lbcInvType.Visible = False
                        imItfCode = 0
                        If lbcInvType.ListIndex > 0 Then
                            grdGet.TextMatrix(lmGetEnableRow, lmGetEnableCol) = lbcInvType.List(lbcInvType.ListIndex)
                            slNameCode = tmInvType(lbcInvType.ListIndex - 1).sKey
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If imItfCode <> Val(slCode) Then
                                imItfCode = Val(slCode)
                                tmItfSrchKey0.iCode = imItfCode
                                ilRet = btrGetEqual(hmItf, tmItf, imItfRecLen, tmItfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    If smIgnoreMultiFeed <> tmItf.sMultiFeed Then
                                        Screen.MousePointer = vbHourglass  'Wait
                                        gSetMousePointer grdSet, grdGet, vbHourglass
                                        gSetMousePointer grdGet, grdDates, vbHourglass
                                        smIgnoreMultiFeed = tmItf.sMultiFeed
                                        mGSFMoveRecToCtrl
                                        If imSelectedIndex <> 0 Then
                                            mMoveRecToCtrl
                                        End If
                                        gSetMousePointer grdSet, grdGet, vbDefault
                                        gSetMousePointer grdGet, grdDates, vbDefault
                                        Screen.MousePointer = vbDefault
                                    End If
                                End If
                            End If
                        Else
                            imItfCode = 0
                            grdGet.TextMatrix(lmGetEnableRow, lmGetEnableCol) = ""
                        End If
                        mInvItemPop
                    Case GETITEMINDEX
                        edcGet.Visible = False
                        cmcGet.Visible = False
                        lbcInvItem.Visible = False
                        imIifCode = 0
                        If lbcInvItem.ListIndex > 0 Then
                            grdGet.TextMatrix(lmGetEnableRow, lmGetEnableCol) = lbcInvItem.List(lbcInvItem.ListIndex)
                            slNameCode = tmInvItem(lbcInvItem.ListIndex - 1).sKey
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            imIifCode = Val(slCode)
                        Else
                            grdGet.TextMatrix(lmGetEnableRow, lmGetEnableCol) = ""
                        End If
                    Case GETINDEPENDENTINDEX
                        pbcIndependent.Visible = False
                        If grdGet.TextMatrix(lmGetEnableRow, lmGetEnableCol) <> smIndependent Then
                            imIhfChg = True
                        End If
                        grdGet.TextMatrix(lmGetEnableRow, lmGetEnableCol) = smIndependent
                        mSetIndependent
                    Case GETOVERSELLINDEX
                        pbcOversell.Visible = False
                        If grdGet.TextMatrix(lmGetEnableRow, lmGetEnableCol) <> smOversell Then
                            imIhfChg = True
                        End If
                        grdGet.TextMatrix(lmGetEnableRow, lmGetEnableCol) = smOversell
                End Select
        End Select
    End If
    lmGetEnableRow = -1
    lmGetEnableCol = -1
    imGetCtrlVisible = False
    mSetCommands
End Sub

Private Function mGetColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    mGetColOk = True
    If grdGet.ColWidth(grdGet.Col) <= 15 Then
        mGetColOk = False
        Exit Function
    End If
    If grdGet.CellBackColor = LIGHTYELLOW Then
        mGetColOk = False
        Exit Function
    End If

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mInvTypeBranch                  *
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
Private Function mInvTypeBranch() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slNameCode                    slCode                                                  *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInvTypeBranchErr                                                                     *
'******************************************************************************************

'
'   ilRet = mInvTypeBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcGet, lbcInvType, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mInvTypeBranch = False
        Exit Function
    End If
    igInvTypeCallSource = CALLSOURCEENAME
    If edcGet.Text = "[New]" Then
        sgInvTypeName = ""
    Else
        sgInvTypeName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "GameInv^Test\" & sgUserName & "\" & Trim$(str$(igInvTypeCallSource)) & "\" & sgInvTypeName
        Else
            slStr = "GameInv^Prod\" & sgUserName & "\" & Trim$(str$(igInvTypeCallSource)) & "\" & sgInvTypeName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "EName^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igInvTypeCallSource)) & "\" & sgTypeName
    '    Else
    '        slStr = "EName^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igInvTypeCallSource)) & "\" & sgTypeName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "Type.Exe " & slStr, 1)
    'EName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    InvType.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgInvTypeName)
    igInvTypeCallSource = Val(sgInvTypeName)
    ilParse = gParseItem(slStr, 2, "\", sgInvTypeName)
    imDoubleClickName = False
    mInvTypeBranch = True
    imUpdateAllowed = ilUpdateAllowed
'        gShowBranner
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igInvTypeCallSource = CALLDONE Then  'Done
        igInvTypeCallSource = CALLNONE
        smInvTypeTag = ""
        lbcInvType.Clear
        mInvTypePop
        If imTerminate Then
            mInvTypeBranch = False
            Exit Function
        End If
        gFindMatch sgInvTypeName, 1, lbcInvType
        sgInvTypeName = ""
        If gLastFound(lbcInvType) > 0 Then
            imChgMode = True
            lbcInvType.ListIndex = gLastFound(lbcInvType)
            edcGet.Text = lbcInvType.List(lbcInvType.ListIndex)
            imChgMode = False
            mInvTypeBranch = False
        Else
            imChgMode = True
            lbcInvType.ListIndex = 0
            edcGet.Text = lbcInvType.List(lbcInvType.ListIndex)
            imChgMode = False
            edcGet.SetFocus
            Exit Function
        End If
    End If
    If igInvTypeCallSource = CALLCANCELLED Then  'Cancelled
        igInvTypeCallSource = CALLNONE
        sgInvTypeName = ""
        mGetEnableBox
        Exit Function
    End If
    If igInvTypeCallSource = CALLTERMINATED Then
        igInvTypeCallSource = CALLNONE
        sgInvTypeName = ""
        mGetEnableBox
        Exit Function
    End If
    Exit Function
mInvTypeBranchErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mInvItemBranch                    *
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
Private Function mInvItemBranch() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slNameCode                    slCode                                                  *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInvItemBranchErr                                                                     *
'******************************************************************************************

'
'   ilRet = mInvItemBranch()
'   Where:
'       ilRet (O)- True = event type started
'                  False = Event type over or not started
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcGet, lbcInvItem, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mInvItemBranch = False
        Exit Function
    End If
    igInvItemCallSource = CALLSOURCEENAME
    If edcGet.Text = "[New]" Then
        sgInvItemName = ""
    Else
        sgInvItemName = slStr
    End If
    If lbcInvType.Text = "[New]" Then
        sgInvTypeName = ""
    Else
        sgInvTypeName = lbcInvType.Text
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "GameInv^Test\" & sgUserName & "\" & Trim$(str$(igInvItemCallSource)) & "\" & sgInvTypeName & "\" & sgInvItemName
        Else
            slStr = "GameInv^Prod\" & sgUserName & "\" & Trim$(str$(igInvItemCallSource)) & "\" & sgInvTypeName & "\" & sgInvItemName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "EName^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igInvItemCallSource)) & "\" & sgTypeName
    '    Else
    '        slStr = "EName^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igInvItemCallSource)) & "\" & sgTypeName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "Type.Exe " & slStr, 1)
    'EName.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    InvItem.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgInvItemName)
    igInvItemCallSource = Val(sgInvItemName)
    ilParse = gParseItem(slStr, 2, "\", sgInvItemName)
    imDoubleClickName = False
    mInvItemBranch = True
    imUpdateAllowed = ilUpdateAllowed
'        gShowBranner
    'If imUpdateAllowed = False Then
    '    mSendHelpMsg "BF"
    'Else
    '    mSendHelpMsg "BT"
    'End If
    gShowBranner imUpdateAllowed
    If igInvItemCallSource = CALLDONE Then  'Done
        igInvItemCallSource = CALLNONE
        smInvItemTag = ""
        lbcInvItem.Clear
        mInvItemPop
        If imTerminate Then
            mInvItemBranch = False
            Exit Function
        End If
        gFindMatch sgInvItemName, 1, lbcInvItem
        sgInvItemName = ""
        If gLastFound(lbcInvItem) > 0 Then
            imChgMode = True
            lbcInvItem.ListIndex = gLastFound(lbcInvItem)
            edcGet.Text = lbcInvItem.List(lbcInvItem.ListIndex)
            imChgMode = False
            mInvItemBranch = False
        Else
            imChgMode = True
            lbcInvItem.ListIndex = 0
            edcGet.Text = lbcInvItem.List(lbcInvItem.ListIndex)
            imChgMode = False
            If edcGet.Visible Then
                edcGet.SetFocus
            End If
            Exit Function
        End If
    End If
    If igInvItemCallSource = CALLCANCELLED Then  'Cancelled
        igInvItemCallSource = CALLNONE
        sgInvItemName = ""
        mGetEnableBox
        Exit Function
    End If
    If igInvItemCallSource = CALLTERMINATED Then
        igInvItemCallSource = CALLNONE
        sgInvItemName = ""
        mGetEnableBox
        Exit Function
    End If
    Exit Function
mInvItemBranchErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mPopulateErr                                                                          *
'******************************************************************************************

'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slInvType As String
    Dim slInvItem As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String

    cbcSelect.Clear
    'tmIhfSrchKey2.iVefCode = imVefCode
    'ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
    tmIhfSrchKey1.lghfcode = lmSeasonGhfCode
    ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmIhf.iVefCode = imVefCode)
        slInvType = ""
        For ilLoop = 0 To UBound(tmInvType) - 1 Step 1
            slNameCode = tmInvType(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tmIhf.iItfCode = Val(slCode) Then
                ilRet = gParseItem(slNameCode, 1, "\", slInvType)
                Exit For
            End If
        Next ilLoop
        slInvItem = ""
        If tmIhf.iIifCode <> tmIif.iCode Then
            tmIifSrchKey0.iCode = tmIhf.iIifCode
            ilRet = btrGetEqual(hmIif, tmIif, imIifRecLen, tmIifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slInvItem = Trim$(tmIif.sName)
            End If
        Else
            slInvItem = Trim$(tmIif.sName)
        End If
        If (slInvType <> "") And (slInvItem <> "") Then
            cbcSelect.AddItem slInvType & "/" & slInvItem
            cbcSelect.ItemData(cbcSelect.NewIndex) = tmIhf.iCode
        End If
        ilRet = btrGetNext(hmIhf, tmIhf, imIhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    cbcSelect.AddItem "[New]", 0
    Exit Sub
mPopulateErr: 'VBC NR
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
Private Function mReadRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIhfCode                                                                             *
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

    smIgnoreMultiFeed = "N"
    ReDim tmIsf(0 To 0) As ISF
    ilUpper = 0
    tmIhfSrchKey0.iCode = imIhfCode
    ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tmIsfSrchKey3.iIhfCode = tmIhf.iCode
        tmIsfSrchKey3.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hmIsf, tmIsf(ilUpper), imIsfRecLen, tmIsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmIhf.iCode = tmIsf(ilUpper).iIhfCode)
            ReDim Preserve tmIsf(0 To UBound(tmIsf) + 1) As ISF
            ilUpper = UBound(tmIsf)
            ilRet = btrGetNext(hmIsf, tmIsf(ilUpper), imIsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        tmItfSrchKey0.iCode = tmIhf.iItfCode
        ilRet = btrGetEqual(hmItf, tmItf, imItfRecLen, tmItfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smIgnoreMultiFeed = tmItf.sMultiFeed
        End If
    Else
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
'*                                                                                        *
'******************************************************************************************

'
'   mXFerRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llRow As Long
    Dim slInvType As String
    Dim slInvItem As String
    Dim ilItemNo As Integer
    Dim slTestStr As String
    Dim ilPos As Integer
    Dim ilPosPrev As Integer
    Dim ilPosLoop As Integer
    Dim ilFound As Integer
    Dim ilRow As Integer

    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If ilRow = grdDates.FixedRows Then
            grdDates.TextMatrix(ilRow, TMISFINDEX) = "-1"
        Else
            'Count Commas
            slStr = grdDates.TextMatrix(ilRow, GAMENOINDEX)
            ilItemNo = 1
            Do
                ilRet = gParseItem(slStr, ilItemNo, ",", slTestStr)
                If slTestStr = "" Then
                    Exit Do
                End If
                ilItemNo = ilItemNo + 1
            Loop While slTestStr <> ""
            If ilItemNo > 1 Then
                ilItemNo = ilItemNo - 1
            End If
            slStr = ""
            For ilLoop = 1 To ilItemNo Step 1
                If slStr = "" Then
                    slStr = "-1"
                Else
                    slStr = slStr & "," & "-1"
                End If
            Next ilLoop
            grdDates.TextMatrix(ilRow, TMISFINDEX) = slStr
        End If
        grdDates.TextMatrix(ilRow, UNITINDEX) = ""
        grdDates.TextMatrix(ilRow, COSTINDEX) = ""
        grdDates.TextMatrix(ilRow, RATEINDEX) = ""
    Next ilRow
    slInvType = ""
    imItfCode = 0
    For ilLoop = 0 To UBound(tmInvType) - 1 Step 1
        slNameCode = tmInvType(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tmIhf.iItfCode = Val(slCode) Then
            imItfCode = tmIhf.iItfCode
            ilRet = gParseItem(slNameCode, 1, "\", slInvType)
            Exit For
        End If
    Next ilLoop
    mInvItemPop
    slInvItem = ""
    For ilLoop = 0 To UBound(tmInvItem) - 1 Step 1
        slNameCode = tmInvItem(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tmIhf.iIifCode = Val(slCode) Then
            ilRet = gParseItem(slNameCode, 1, "\", slInvItem)
            Exit For
        End If
    Next ilLoop

    grdGet.Row = GETROW3INDEX
    grdGet.Col = GETTYPEINDEX
    grdGet.TextMatrix(GETROW3INDEX, GETTYPEINDEX) = slInvType
'    If slInvType <> "" Then
'        grdGet.CellBackColor = LIGHTYELLOW
'        grdGet.Row = GETROW3INDEX - 1
'        grdGet.Col = GETTYPEINDEX
'        grdGet.CellBackColor = LIGHTYELLOW
'    Else
'        grdGet.CellBackColor = vbWhite
'        grdGet.Row = GETROW3INDEX - 1
'        grdGet.Col = GETTYPEINDEX
'        grdGet.CellBackColor = vbWhite
'    End If
    grdGet.Row = GETROW3INDEX
    grdGet.Col = GETITEMINDEX
    grdGet.TextMatrix(GETROW3INDEX, GETITEMINDEX) = slInvItem
'    If slInvItem <> "" Then
'        grdGet.CellBackColor = LIGHTYELLOW
'        grdGet.Row = GETROW3INDEX - 1
'        grdGet.Col = GETITEMINDEX
'        grdGet.CellBackColor = LIGHTYELLOW
'    Else
'        grdGet.CellBackColor = vbWhite
'        grdGet.Row = GETROW3INDEX - 1
'        grdGet.Col = GETITEMINDEX
'        grdGet.CellBackColor = vbWhite
'    End If
    If tmIhf.sGameIndependent = "Y" Then
        grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = "Yes"
    Else
        grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = "No"
    End If
    mSetIndependent
    If tmIhf.sOversell = "Y" Then
        grdGet.TextMatrix(GETROW3INDEX, GETOVERSELLINDEX) = "Yes"
    Else
        grdGet.TextMatrix(GETROW3INDEX, GETOVERSELLINDEX) = "No"
    End If

    For ilLoop = 0 To UBound(tmIsf) - 1 Step 1
        For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
            slStr = grdDates.TextMatrix(llRow, GAMENOINDEX)
            If slStr <> "Ind" Then
                ilFound = False
                ilItemNo = 1
                Do
                    ilRet = gParseItem(slStr, ilItemNo, ",", slTestStr)
                    If slTestStr = "" Then
                        Exit Do
                    End If
                    If Val(slTestStr) = tmIsf(ilLoop).iGameNo Then
                        ilFound = True
                        Exit Do
                    End If
                    ilItemNo = ilItemNo + 1
                Loop While slTestStr <> ""
                If ilFound Then
                    If ilItemNo = 1 Then
                        slStr = grdDates.TextMatrix(llRow, TMISFINDEX)
                        ilPos = InStr(1, slStr, ",", vbTextCompare)
                        grdDates.Row = llRow
                        If ilPos = 0 Then
                            grdDates.TextMatrix(llRow, TMISFINDEX) = Trim$(str$(ilLoop))
                        Else
                            slStr = grdDates.TextMatrix(llRow, TMISFINDEX)
                            grdDates.TextMatrix(llRow, TMISFINDEX) = Trim$(str$(ilLoop)) & Mid(slStr, ilPos)
                        End If
                        If tmIsf(ilLoop).iNoUnits > 0 Then
                            grdDates.TextMatrix(llRow, UNITINDEX) = Trim$(str$(tmIsf(ilLoop).iNoUnits))
                            grdDates.TextMatrix(llRow, COSTINDEX) = gLongToStrDec(tmIsf(ilLoop).lCost, 2)
                            grdDates.TextMatrix(llRow, RATEINDEX) = gLongToStrDec(tmIsf(ilLoop).lRate, 2)
                        Else
                            grdDates.TextMatrix(llRow, UNITINDEX) = ""
                            grdDates.TextMatrix(llRow, COSTINDEX) = ""
                            grdDates.TextMatrix(llRow, RATEINDEX) = ""
                        End If
                        Exit For
                    Else
                        slStr = grdDates.TextMatrix(llRow, TMISFINDEX)
                        ilPos = 1
                        ilPosPrev = 1
                        For ilPosLoop = 1 To ilItemNo Step 1
                            ilPos = InStr(ilPos, slStr, ",", vbTextCompare)
                            If ilPos = 0 Then
                                Exit For
                            End If
                            ilPosPrev = ilPos
                            ilPos = ilPos + 1
                        Next ilPosLoop
                        If ilPos = 0 Then
                            grdDates.TextMatrix(llRow, TMISFINDEX) = Left$(slStr, ilPosPrev - 1) & "," & Trim$(str$(ilLoop))
                        Else
                            grdDates.TextMatrix(llRow, TMISFINDEX) = Left$(slStr, ilPosPrev - 1) & "," & Trim$(str$(ilLoop)) & Mid(slStr, ilPos)
                        End If
                        If tmIsf(ilLoop).iNoUnits <> 0 Then
                            tmIsf(ilLoop).iNoUnits = 0
                            imIsfChg = True
                        End If
                    End If
                End If
            Else
                If tmIsf(ilLoop).iGameNo = 0 Then
                    grdDates.Row = llRow
                    grdDates.TextMatrix(llRow, TMISFINDEX) = Trim$(str$(ilLoop))
                    If tmIsf(ilLoop).iNoUnits > 0 Then
                        grdDates.TextMatrix(llRow, UNITINDEX) = Trim$(str$(tmIsf(ilLoop).iNoUnits))
                        grdDates.TextMatrix(llRow, COSTINDEX) = gLongToStrDec(tmIsf(ilLoop).lCost, 2)
                        grdDates.TextMatrix(llRow, RATEINDEX) = gLongToStrDec(tmIsf(ilLoop).lRate, 2)
                    Else
                        grdDates.TextMatrix(llRow, UNITINDEX) = ""
                        grdDates.TextMatrix(llRow, COSTINDEX) = ""
                        grdDates.TextMatrix(llRow, RATEINDEX) = ""
                    End If
                    Exit For
                End If
            End If
        Next llRow
    Next ilLoop
    Exit Sub
End Sub

Private Function mSaveRec() As Integer
    Dim ilRow As Integer
    Dim slMsg As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilError As Integer
    Dim tlIhf As IHF
    Dim tlIsf As ISF

    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSet, grdGet, vbHourglass
    gSetMousePointer grdGet, grdDates, vbHourglass
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If mGridFieldsOk(ilRow) = False Then
            ilError = True
        End If
    Next ilRow
    If Not mOkTypeItem() Then
        ilError = True
    End If
    If ilError Then
        gSetMousePointer grdSet, grdGet, vbDefault
        gSetMousePointer grdGet, grdDates, vbDefault
        Screen.MousePointer = vbDefault
        Beep
        mSaveRec = False
        Exit Function
    End If
    mMoveCtrlToRec
    ilRet = btrBeginTrans(hmIhf, 1000)
    If imNewInv Then
        tmIhf.iCode = 0
        tmIhf.lghfcode = tmGhf.lCode
        tmIhf.iVefCode = imVefCode
        tmIhf.iItfCode = imItfCode
        tmIhf.iIifCode = imIifCode
        If grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = "Yes" Then
            tmIhf.sGameIndependent = "Y"
        Else
            tmIhf.sGameIndependent = "N"
        End If
        If grdGet.TextMatrix(GETROW3INDEX, GETOVERSELLINDEX) = "Yes" Then
            tmIhf.sOversell = "Y"
        Else
            tmIhf.sOversell = "N"
        End If
        tmIhf.sUnused = ""
        ilRet = btrInsert(hmIhf, tmIhf, imIhfRecLen, INDEXKEY0)
        slMsg = "mSaveRec (btrInsert:Inventory Header)"
    Else
        Do
            tmIhfSrchKey0.iCode = tmIhf.iCode
            ilRet = btrGetEqual(hmIhf, tlIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = "Yes" Then
                tmIhf.sGameIndependent = "Y"
            Else
                tmIhf.sGameIndependent = "N"
            End If
            If grdGet.TextMatrix(GETROW3INDEX, GETOVERSELLINDEX) = "Yes" Then
                tmIhf.sOversell = "Y"
            Else
                tmIhf.sOversell = "N"
            End If
            ilRet = btrUpdate(hmIhf, tmIhf, imIhfRecLen)
            slMsg = "mSaveRec (btrUpdate:Inventory Header)"
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, GameInv
    On Error GoTo 0
    For ilLoop = 0 To UBound(tmIsf) - 1 Step 1
        If imNewInv Or tmIsf(ilLoop).lCode <= 0 Then
            tmIsf(ilLoop).lCode = 0
            tmIsf(ilLoop).iVefCode = imVefCode
            tmIsf(ilLoop).lghfcode = tmGhf.lCode
            tmIsf(ilLoop).iIhfCode = tmIhf.iCode
            ilRet = btrInsert(hmIsf, tmIsf(ilLoop), imIsfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert:Inventory Schedule)"
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, GameInv
                On Error GoTo 0
            End If
        Else
            Do
                tmIsfSrchKey0.lCode = tmIsf(ilLoop).lCode
                ilRet = btrGetEqual(hmIsf, tlIsf, imIsfRecLen, tmIsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                ilRet = btrUpdate(hmIsf, tmIsf(ilLoop), imIsfRecLen)
                slMsg = "mSaveRec (btrUpdate:Inventory Schedule)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, GameInv
                On Error GoTo 0
            End If
        End If
    Next ilLoop
    ilRet = btrEndTrans(hmGhf)
    mSaveRec = True
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    ilRet = btrAbortTrans(hmGhf)
    gSetMousePointer grdSet, grdGet, vbDefault
    gSetMousePointer grdGet, grdDates, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function

Private Sub mClearGrdDates(ilClearAllFields As Integer)
    Dim ilRow As Integer
    Dim ilCol As Integer

    If grdDates.Rows > 31 Then
        For ilRow = grdDates.Rows - 1 To 31 Step -1
            grdDates.RemoveItem ilRow
        Next ilRow
    End If
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If ilClearAllFields Then
            For ilCol = GAMENOINDEX To SORTINDEX Step 2
                If (ilCol = TMISFINDEX) Then
                    grdDates.TextMatrix(ilRow, ilCol) = "-1"
                Else
                    grdDates.TextMatrix(ilRow, ilCol) = ""
                End If
            Next ilCol
        Else
            grdDates.TextMatrix(ilRow, TMISFINDEX) = "-1"
            grdDates.TextMatrix(ilRow, UNITINDEX) = ""
            grdDates.TextMatrix(ilRow, COSTINDEX) = ""
            grdDates.TextMatrix(ilRow, RATEINDEX) = ""
        End If
    Next ilRow
    For ilCol = GAMENOINDEX To AIRTIMEINDEX Step 2
        grdDates.Col = ilCol
        For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
            grdDates.Row = ilRow
            grdDates.CellBackColor = LIGHTYELLOW
        Next ilRow
    Next ilCol
    For ilCol = UNITINDEX To RATEINDEX Step 2
        grdDates.Col = ilCol
        For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
            grdDates.Row = ilRow
            grdDates.CellBackColor = vbWhite
        Next ilRow
    Next ilCol
    lmTopRow = grdDates.TopRow
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

    'ilRet = gPopUserVehicleBox(Program, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, cbcVeh, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBox(GameInv, VEHSPORT + ACTIVEVEH, cbcGameVeh, tmGameVehicle(), smGameVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", Program
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mSetIndependent()
    Dim llRow As Long
    Dim slStr As String

    If grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = "Yes" Then
        grdSet.Visible = False
        cmcSetInv.Visible = False
        pbcSetSTab.Visible = False
        pbcSetTab.Visible = False
    Else
        mGridSetColumns
        grdSet.Visible = True
        cmcSetInv.Visible = True
        pbcSetSTab.Visible = True
        pbcSetTab.Visible = True
    End If
    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        slStr = Trim$(grdDates.TextMatrix(llRow, GAMENOINDEX))
        If slStr <> "" Then
            grdDates.Row = llRow
            If slStr <> "Ind" Then
                If grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = "Yes" Then
                    grdDates.Col = UNITINDEX
                    grdDates.CellBackColor = LIGHTYELLOW
                    grdDates.Text = ""
                    grdDates.Col = COSTINDEX
                    grdDates.CellBackColor = LIGHTYELLOW
                    grdDates.Text = ""
                    grdDates.Col = RATEINDEX
                    grdDates.CellBackColor = LIGHTYELLOW
                    grdDates.Text = ""
                Else
                    grdDates.Col = UNITINDEX
                    grdDates.CellBackColor = vbWhite
                    grdDates.Col = COSTINDEX
                    grdDates.CellBackColor = vbWhite
                    grdDates.Col = RATEINDEX
                    grdDates.CellBackColor = vbWhite
                End If
            Else
                If (grdGet.TextMatrix(GETROW3INDEX, GETINDEPENDENTINDEX) = "Yes") Then
                    grdDates.Col = UNITINDEX
                    grdDates.CellBackColor = vbWhite
                    grdDates.Col = COSTINDEX
                    grdDates.CellBackColor = vbWhite
                    grdDates.Col = RATEINDEX
                    grdDates.CellBackColor = vbWhite
                Else
                    grdDates.Col = UNITINDEX
                    grdDates.CellBackColor = LIGHTYELLOW
                    grdDates.Text = ""
                    grdDates.Col = COSTINDEX
                    grdDates.CellBackColor = LIGHTYELLOW
                    grdDates.Text = ""
                    grdDates.Col = RATEINDEX
                    grdDates.CellBackColor = LIGHTYELLOW
                    grdDates.Text = ""
                End If
            End If
        End If
    Next llRow
End Sub

Private Function mOkTypeItem() As Integer
    Dim slInvType As String
    Dim slInvItem As String
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilIhfCode As Integer
    Dim slName As String

    slInvType = grdGet.TextMatrix(GETROW3INDEX, GETTYPEINDEX)
    slInvItem = grdGet.TextMatrix(GETROW3INDEX, GETITEMINDEX)
    slName = Trim$(slInvType) & "/" & Trim$(slInvItem)
    For ilLoop = 0 To cbcSelect.ListCount - 1 Step 1
        slStr = Trim$(cbcSelect.List(ilLoop))
        If StrComp(slStr, slName, vbTextCompare) = 0 Then
            ilIhfCode = cbcSelect.ItemData(ilLoop)
            If imIhfCode <> ilIhfCode Then
                Beep
                MsgBox slName & " already defined, enter a different Type/Item", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                mOkTypeItem = False
                Exit Function
            End If
        End If
    Next ilLoop
    mOkTypeItem = True
    Exit Function
End Function

Private Sub mSeasonPop(blFromInit As Boolean)
    Dim llStartDate As Long
    Dim slStartDate As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llSeasonGhfCode As Long
    Dim ilVff As Integer
    
    cbcSeason.Clear
    If Not blFromInit Then
        lmSeasonGhfCode = 0
    End If
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
    If (Not blFromInit) Then
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tgVff(ilVff).iVefCode = imVefCode Then
                lmSeasonGhfCode = tgVff(ilVff).lSeasonGhfCode
                Exit For
            End If
        Next ilVff
    End If
    For ilLoop = 0 To cbcSeason.ListCount - 1 Step 1
        If cbcSeason.ItemData(ilLoop) = lmSeasonGhfCode Then
            cbcSeason.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    
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
            gSetMousePointer grdSet, grdGet, vbHourglass
            gSetMousePointer grdGet, grdDates, vbHourglass
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
            ilRet = mGhfGsfReadRec()
            mLanguagePop
            mInvTypePop
            mPopulate
            Screen.MousePointer = vbHourglass  'Wait
            gSetMousePointer grdSet, grdGet, vbHourglass
            gSetMousePointer grdGet, grdDates, vbHourglass
        Loop While imSeasonSelectedIndex <> cbcSeason.ListIndex
        Screen.MousePointer = vbDefault    'Default
        gSetMousePointer grdSet, grdGet, vbDefault
        gSetMousePointer grdGet, grdDates, vbDefault
        imChgMode = False
        mSetCommands
    End If
End Sub
