VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CGameSch 
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
   Begin VB.TextBox edcSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3615
      MaxLength       =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   930
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdWS 
      Height          =   1425
      Left            =   4335
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   2514
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.ListBox lbcSeason 
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
      ItemData        =   "CGameSch.frx":0000
      Left            =   3570
      List            =   "CGameSch.frx":0002
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   2415
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
      ItemData        =   "CGameSch.frx":0004
      Left            =   8190
      List            =   "CGameSch.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   195
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox pbcSpotsBy 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1575
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmcSetSpots 
      Appearance      =   0  'Flat
      Caption         =   "&Set Spots"
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
      Left            =   8235
      TabIndex        =   23
      Top             =   810
      Width           =   1050
   End
   Begin VB.ListBox lbcLnModel 
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
      ItemData        =   "CGameSch.frx":0008
      Left            =   2535
      List            =   "CGameSch.frx":000A
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   525
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox pbcComment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   390
      ScaleHeight     =   210
      ScaleWidth      =   810
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   915
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8895
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   22
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
      ItemData        =   "CGameSch.frx":000C
      Left            =   7050
      List            =   "CGameSch.frx":000E
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   1200
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
      ItemData        =   "CGameSch.frx":0010
      Left            =   1260
      List            =   "CGameSch.frx":0012
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   2190
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
      ItemData        =   "CGameSch.frx":0014
      Left            =   1965
      List            =   "CGameSch.frx":0016
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4515
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   45
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   11
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
      TabIndex        =   1
      Top             =   375
      Width           =   60
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
      Left            =   1395
      Picture         =   "CGameSch.frx":0018
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcNoSpots 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   990
      MaxLength       =   10
      TabIndex        =   14
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
      Picture         =   "CGameSch.frx":0112
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   17
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
      TabIndex        =   21
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
      TabIndex        =   18
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
      TabIndex        =   12
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
      Left            =   4980
      TabIndex        =   20
      Top             =   5460
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
      Left            =   3270
      TabIndex        =   19
      Top             =   5460
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDates 
      Height          =   3825
      Left            =   195
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1290
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6747
      _Version        =   393216
      Rows            =   39
      Cols            =   30
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
      _Band(0).Cols   =   30
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpec 
      Height          =   885
      Left            =   210
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   255
      Width           =   7680
      _ExtentX        =   13547
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
   Begin VB.Label plcScreen 
      Caption         =   "Event Spots"
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
      Top             =   15
      Width           =   6930
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
Attribute VB_Name = "CGameSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CGameSch.Frm
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
Dim imComboBoxIndex As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Dim imLoadingForm As Integer
Dim imVefCode As Integer
Dim imVpfIndex As Integer
Dim imRdfCode As Integer
Dim lmOvStartTime As Long
Dim lmOvEndTime As Long
Dim imCntrLineNo As Integer
Dim imLineSpotLen As Integer
Dim smComment As String
Dim smSpotsBy As String
Dim bmSpotsBySet As Boolean
Dim bmSetSpotsPressed As Boolean
Dim imMinGameNo As Integer
Dim imMaxGameNo As Integer
Dim smDefaultGameNo As String
Dim imLastColSorted As Integer
Dim imLastSort As Integer
Dim imAvailColorLevel As Integer    'set in mInit as 90%

Dim smNowDate As String
Dim lmNowDate As Long
Dim lmLLD As Long
Dim lmFirstAllowedChgDate As Long

Dim imCgfChg As Integer
Dim imNewGame As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmSpecEnableRow As Long
Dim lmSpecEnableCol As Long
Dim imSpecCtrlVisible As Integer
Dim lmWSEnableRow As Long
Dim lmWSEnableCol As Long
Dim imWSCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmGhfSrchKey0 As LONGKEY0    'GHF key record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length
Dim lmSeasonGhfCode As Long
Dim tmSeasonInfo() As SEASONINFO

Dim hmGsf As Integer
Dim tmGsf() As GSF        'GSF record image
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

Dim hmCgf As Integer
Dim tmCgf As CGF        'GHF record image
Dim tmCgfSrchKey1 As CGFKEY1    'GHF key record image
Dim imCgfRecLen As Integer        'GHF record length

'Dim hmSsf As Integer        'Spot summary file handle
Dim hmSsf As Integer
Dim tmSsf As SSF               'SSF record image
Dim tmSsfSrchKey1 As SSFKEY1 'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2 'SSF key record image
Dim imSsfRecLen As Integer  'SSF record length
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

'Spot file
Dim hmSdf As Integer
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey3 As LONGKEY0 'SDF key record image (code)
Dim tmSdfSrchKey6 As SDFKEY6 'SDF key record image (advt)
Dim imSdfRecLen As Integer     'SDF record length

'MG Spot file
Dim tmSmf As SMF            'SMF record image
Dim tmSmfSrchKey3 As SMFKEY3 'SMF key record image (agency)
Dim hmSmf As Integer        'SMF Handle
Dim imSmfRecLen As Integer     'SMF record length

Dim hmCHF As Integer
Dim tmChf As CHF            'SDF record image
Dim tmChfSrchKey0 As LONGKEY0 'SDF key record image (agency)
Dim imCHFRecLen As Integer     'SDF record length

'Contract Line
Dim tmClfSrchKey As CLFKEY0  'CLF key record image
Dim tmClfSrchKey3 As CLFKEY3  'CLF key record image
Dim hmClf As Integer        'CLF Handle
Dim tmClf As CLF
Dim imClfRecLen As Integer      'CLF record length
Dim tmPropGameInfo() As PROPGAMEINFO

Dim tmTeamCode() As SORTCODE
Dim smTeamCodeTag As String

Dim tmLanguageCode() As SORTCODE
Dim smLanguageCodeTag As String

Dim tmSvCff() As CFFLIST
Dim tmSvCgf() As CGFLIST
Dim imSvFirstCff As Integer
Dim imSvFirstCgf As Integer

'6/9/14
Dim smEventTitle1 As String
Dim smEventTitle2 As String

'Mouse down
'Row 3
Const SPECROW3INDEX = 3
Const COMMENTINDEX = 2  '4
Const MODELLNINDEX = 4  '2
Const SEASONINDEX = 6
'Row 6
Const SPECROW6INDEX = 6
Const LANGUAGETYPEINDEX = 2
Const SPOTSBYINDEX = 4
Const NOSPOTSPERINDEX = 6
'Row 9
Const SPECROW9INDEX = 9
'Const NOSPOTSPERGAMEINDEX = 2
'Const NOSPOTSPERWEEKINDEX = 4
'Const LANGUAGETYPEINDEX = 6
Const GAMENOSINDEX = 2
Const GAMEININDEX = 4
Const GAMEOUTINDEX = 6

Const GAMENOINDEX = 2   '1
Const FEEDSOURCEINDEX = 4   '2
Const LANGUAGEINDEX = 6 '3
Const AVAILSORDERINDEX = 8
Const AVAILSPROPOSALINDEX = 10
Const NOSPOTSINDEX = 12
Const VISITTEAMINDEX = 14    '4
Const HOMETEAMINDEX = 16    '5
Const WEEKOFINDEX = 18
Const AIRDAYINDEX = 20
Const AIRDATEINDEX = 22 '7
Const AIRTIMEINDEX = 24 '8
Const INVINDEX = 26
Const TMCGFINDEX = 27
Const SORTINDEX = 28
Const GAMESTATUSINDEX = 29

Const WSDATESINDEX = 0
Const WSSPOTSINDEX = 1
Const WSGAMENUMBERSINDEX = 2
Const WSMODATEINDEX = 3







Private Sub cmcCancel_Click()
    igGameReturn = False
    mRestoreFlightInfo
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mWSSetShow
    mSpecSetShow
    mSetShow
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    Dim ilRow As Integer
    Dim ilError As Integer

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If smComment = "Yes" Then
        tgClfCntr(imLnRowNo - 1).sGameLayout = "Y"
    Else
        tgClfCntr(imLnRowNo - 1).sGameLayout = "N"
    End If
    If bmSetSpotsPressed Then
        If smSpotsBy <> "Week" Then
            tgClfCntr(imLnRowNo - 1).ClfRec.sSportsByWeek = "N"
        Else
            tgClfCntr(imLnRowNo - 1).ClfRec.sSportsByWeek = "W"
        End If
    End If
    If imCgfChg Then
        ilError = False
        For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
            If mGridFieldsOk(ilRow) = False Then
                ilError = True
            End If
        Next ilRow
        If ilError Then
            Beep
            Exit Sub
        End If
        mMoveCtrlToRec
    End If
    igGameReturn = True
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mWSSetShow
    mSpecSetShow
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcDone
End Sub


Private Sub cmcSetSpots_Click()
    Dim slStr As String
    Dim llRow As Long
    Dim ilPos As Integer
    Dim slLineNo As String
    Dim ilLineNo As Integer
    Dim ilLine As Integer
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
    Dim llDate As Long
    Dim ilDateLoop As Integer
    Dim ilCount As Integer
    Dim ilNoSpots As Integer
    Dim ilNoSpotsPerWk As Integer
    Dim ilSpots As Integer
    Dim ilCSpots As Integer
    Dim ilFound As Integer
    Dim ilGame As Integer
    Dim ilLangOk As Integer
    Dim ilValue As Integer
    Dim slSortAvails As String
    Dim slSortDate As String
    Dim slSortGameNo As String
    Dim llGame As Long
    Dim slRow As String
    Dim ilRet As Integer
    Dim slOrigNoSpots As String
    Dim ilOrigNoSpots As Integer
    Dim ilGameOk As Integer
    Dim slGameStatus As String

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSpec, grdDates, vbHourglass
    grdDates.Redraw = False
    'Replace Number of spots in selected weeks only
    'For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
    '    If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
    '        grdDates.TextMatrix(llRow, NOSPOTSINDEX) = ""
    '    End If
    'Next llRow
    imCgfChg = True
    bmSetSpotsPressed = True
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX)
    If slStr <> "" Then
        'Set values from previous defined line
        gFindMatch slStr, 0, lbcLnModel
        If gLastFound(lbcLnModel) >= 0 Then
            ilPos = InStr(1, slStr, "-", vbTextCompare)
            If ilPos > 0 Then
                slLineNo = Left$(slStr, ilPos - 1)
                ilLineNo = Val(slLineNo)
                'For ilLine = LBound(smLnSave, 2) To UBound(smLnSave, 2) Step 1
                For ilLine = imLB1Or2 To UBound(smLnSave, LINEBOUNDINDEX) Step 1
                    If tgClfCntr(ilLine - 1).ClfRec.iLine = ilLineNo Then
                        tgClfCntr(imLnRowNo - 1).ClfRec.lghfcode = tgClfCntr(ilLine - 1).ClfRec.lghfcode
                        '3/17/13: Add Loop of seasons
                        tgClfCntr(imLnRowNo - 1).ClfRec.sSportsByWeek = tgClfCntr(ilLine - 1).ClfRec.sSportsByWeek
                        If tgClfCntr(imLnRowNo - 1).ClfRec.sSportsByWeek = "W" Then
                            smSpotsBy = "Week"
                        Else
                            smSpotsBy = "Event"
                        End If
                        lmSpecEnableRow = SPECROW6INDEX
                        lmSpecEnableCol = SPOTSBYINDEX
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smSpotsBy
                        mSeasonPop False
                        'Loop on seasons
                        For ilLoop = 0 To lbcSeason.ListCount - 1 Step 1
                            'lbcSeason.ListIndex = ilLoop
                            lmSpecEnableRow = SPECROW3INDEX
                            lmSpecEnableCol = SEASONINDEX
                            grdSpec.Row = lmSpecEnableRow
                            grdSpec.Col = lmSpecEnableCol
                            edcSpec.Text = lbcSeason.List(ilLoop)
                            mSpecSetShow
                            For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                                If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                                    ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                                    grdDates.TextMatrix(llRow, NOSPOTSINDEX) = mGetSpotCounts(ilLine, ilGameNo, grdDates.TextMatrix(llRow, AIRDATEINDEX))
                                End If
                            Next llRow
                        Next ilLoop
                        Exit For
                    End If
                Next ilLine
            End If
        End If
    Else
        slStr = grdSpec.TextMatrix(SPECROW9INDEX, GAMEININDEX)
        If slStr = "" Then
            ilGamesOn = 1000
            ilGamesOff = 0
        Else
            ilGamesOn = Val(slStr)
            slStr = grdSpec.TextMatrix(SPECROW9INDEX, GAMEOUTINDEX)
            If slStr = "" Then
                ilGamesOff = 0
            Else
                ilGamesOff = Val(slStr)
            End If
        End If
        If grdSpec.TextMatrix(SPECROW6INDEX, SPOTSBYINDEX) <> "Week" Then
            slStr = grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX)
        Else
            slStr = ""
            For llRow = grdWS.FixedRows To grdWS.Rows - 1 Step 1
                If grdWS.TextMatrix(llRow, WSSPOTSINDEX) <> "" Then
                    If slStr = "" Then
                        slStr = grdWS.TextMatrix(llRow, WSGAMENUMBERSINDEX)
                    Else
                        slStr = slStr & "," & grdWS.TextMatrix(llRow, WSGAMENUMBERSINDEX)
                    End If
                End If
            Next llRow
        End If
        'Create an array of game numbers
        ReDim slFields(0 To 0) As String
        ReDim ilGameNos(0 To 0) As Integer
        ReDim ilGameNoOfSpots(0 To 0) As Integer
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
                    'Ignore Language not matching games
                    If (ilValue And USINGLANG) = USINGLANG Then
                        ilLangOk = False
                        For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                            If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                                ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                                If ilGameNo = ilGame Then
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
                    If ilLangOk Then
                        ilGameOk = True
                        slGameStatus = ""
                        For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                            If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                                ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                                If ilGameNo = ilGame Then
                                    grdDates.Row = llRow
                                    grdDates.Col = NOSPOTSINDEX
                                    If grdDates.ForeColor = vbRed Then
                                        ilGameOk = False
                                    End If
                                    slGameStatus = grdDates.TextMatrix(llRow, GAMESTATUSINDEX)
                                    Exit For
                                End If
                            End If
                        Next llRow
                    End If
                    If ilLangOk And ilGameOk Then
                        For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                            If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                                If ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX)) Then
                                    ilGameNoOfSpots(UBound(ilGameNoOfSpots)) = Val(grdDates.TextMatrix(llRow, NOSPOTSINDEX))
                                    grdDates.TextMatrix(llRow, NOSPOTSINDEX) = ""
                                    Exit For
                                End If
                            End If
                        Next llRow
                        If slGameStatus <> "C" Then
                            If ilGameTest = 0 Then
                                ReDim Preserve ilGameNoOfSpots(0 To UBound(ilGameNoOfSpots) + 1) As Integer
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
                If (ilValue And USINGLANG) = USINGLANG Then
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
                If ilLangOk Then
                    ilGameOk = True
                    slGameStatus = ""
                    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                            ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                            If Val(slFields(ilLoop)) = ilGame Then
                                grdDates.Row = llRow
                                grdDates.Col = NOSPOTSINDEX
                                If grdDates.ForeColor = vbRed Then
                                    ilGameOk = False
                                End If
                                slGameStatus = grdDates.TextMatrix(llRow, GAMESTATUSINDEX)
                                Exit For
                            End If
                        End If
                    Next llRow
                End If
                If ilLangOk And ilGameOk Then
                    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                            If Val(slFields(ilLoop)) = Val(grdDates.TextMatrix(llRow, GAMENOINDEX)) Then
                                ilGameNoOfSpots(UBound(ilGameNoOfSpots)) = Val(grdDates.TextMatrix(llRow, NOSPOTSINDEX))
                                grdDates.TextMatrix(llRow, NOSPOTSINDEX) = ""
                                Exit For
                            End If
                        End If
                    Next llRow
                    If slGameStatus <> "C" Then
                        ReDim Preserve ilGameNoOfSpots(0 To UBound(ilGameNoOfSpots) + 1) As Integer
                        ilGameNos(UBound(ilGameNos)) = Val(slFields(ilLoop))
                        ReDim Preserve ilGameNos(0 To UBound(ilGameNos) + 1) As Integer
                    End If
                End If
            End If
        Next ilLoop
        'Distribute spots to games
        slStr = grdSpec.TextMatrix(SPECROW6INDEX, SPOTSBYINDEX)
        If slStr <> "Week" Then
            slStr = grdSpec.TextMatrix(SPECROW6INDEX, NOSPOTSPERINDEX)
            For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                    ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                    For ilGame = 0 To UBound(ilGameNos) - 1 Step 1
                        If ilGameNos(ilGame) = ilGameNo Then
                            grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX) = Trim$(str$(Val(grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX)) + (ilGameNoOfSpots(ilGame) - Val(slStr))))
                            grdDates.TextMatrix(llRow, NOSPOTSINDEX) = slStr
                            Exit For
                        End If
                    Next ilGame
                End If
            Next llRow
        Else
            'Determine number of games running with spots in Week, then distribute spots to each game within week
            ReDim llWkDate(0 To 0) As Long
            ReDim ilWkCount(0 To 0) As Integer
            ReDim ilAvailCount(0 To 0) As Integer
            ReDim ilAvailFlag(0 To 0) As Integer
            ReDim tlGameSort(0 To 0) As SORTCODE
            ilAvailFlag(0) = -1
            For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
                If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                    ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                    slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                    For ilGame = 0 To UBound(ilGameNos) - 1 Step 1
                        If ilGameNos(ilGame) = ilGameNo Then
                            slSortAvails = gSubStr("9999", grdDates.TextMatrix(llRow, AVAILSORDERINDEX))
                            Do While Len(slSortAvails) < 4
                                slSortAvails = "0" & slSortAvails
                            Loop
                            slSortDate = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                            slSortDate = Trim$(str$(gDateValue(slSortDate)))
                            Do While Len(slSortDate) < 6
                                slSortDate = "0" & slSortDate
                            Loop
                            slSortGameNo = grdDates.TextMatrix(llRow, GAMENOINDEX)
                            Do While Len(slSortGameNo) < 4
                                slSortGameNo = "0" & slSortGameNo
                            Loop
                            tlGameSort(UBound(tlGameSort)).sKey = slSortAvails & "|" & slSortDate & "|" & slSortGameNo & "|" & Trim$(str$(llRow)) & "|" & Trim$(str$(ilGameNoOfSpots(ilGame)))
                            ReDim Preserve tlGameSort(0 To UBound(tlGameSort) + 1) As SORTCODE
                            llDate = gDateValue(gObtainPrevMonday(slStr))
                            ilFound = False
                            For ilLoop = 0 To UBound(llWkDate) - 1 Step 1
                                If llWkDate(ilLoop) = llDate Then
                                    ilFound = True
                                    ilWkCount(ilLoop) = ilWkCount(ilLoop) + 1
                                    If Trim$(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) <> "" Then
                                        ilAvailCount(ilLoop) = ilAvailCount(ilLoop) + Val(grdDates.TextMatrix(llRow, AVAILSORDERINDEX))
                                    End If
                                    If Trim$(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) <> "" Then
                                        If ilAvailFlag(ilLoop) <> Val(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) Then
                                            ilAvailFlag(ilLoop) = -1
                                        End If
                                    Else
                                        If ilAvailFlag(ilLoop) <> 0 Then
                                            ilAvailFlag(ilLoop) = -1
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                llWkDate(UBound(llWkDate)) = llDate
                                ilWkCount(UBound(ilWkCount)) = 1
                                If Trim$(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) <> "" Then
                                    ilAvailCount(UBound(ilAvailCount)) = Val(grdDates.TextMatrix(llRow, AVAILSORDERINDEX))
                                Else
                                    ilAvailCount(UBound(ilAvailCount)) = 0
                                End If
                                If Trim$(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) <> "" Then
                                    ilAvailFlag(UBound(ilAvailFlag)) = Val(grdDates.TextMatrix(llRow, AVAILSORDERINDEX))
                                Else
                                    ilAvailFlag(UBound(ilAvailFlag)) = 0
                                End If
                                ReDim Preserve llWkDate(0 To UBound(llWkDate) + 1) As Long
                                ReDim Preserve ilWkCount(0 To UBound(ilWkCount) + 1) As Integer
                                ReDim Preserve ilAvailCount(0 To UBound(ilAvailCount) + 1) As Integer
                                ReDim Preserve ilAvailFlag(0 To UBound(ilAvailFlag) + 1) As Integer
                            End If
                            Exit For
                        End If
                    Next ilGame
                End If
            Next llRow
            If UBound(tlGameSort) - 1 > 0 Then
                ArraySortTyp fnAV(tlGameSort(), 0), UBound(tlGameSort), 0, LenB(tlGameSort(0)), 0, LenB(tlGameSort(0).sKey), 0
            End If
            'Distribute the spots to games with avails that still have room for spots only
            'grdDates.TextMatrix(llRow, AVAILSORDERINDEX) content the # units available to be booked into
            'ilAvailCount() contents total number of units available for the week
            'ilWkCount() contents number of games within the week
            slStr = grdSpec.TextMatrix(SPECROW6INDEX, NOSPOTSPERINDEX)
            ilNoSpotsPerWk = Val(slStr)
            For ilDateLoop = 0 To UBound(llWkDate) - 1 Step 1
                llDate = llWkDate(ilDateLoop)
                slStr = ""
                ilNoSpotsPerWk = 0
                For llRow = grdWS.FixedRows To grdWS.Rows - 1 Step 1
                    If grdWS.TextMatrix(llRow, WSSPOTSINDEX) <> "" Then
                        If llDate = gDateValue(grdWS.TextMatrix(llRow, WSMODATEINDEX)) Then
                            slStr = grdWS.TextMatrix(llRow, WSSPOTSINDEX)
                            ilNoSpotsPerWk = Val(slStr)
                            Exit For
                        End If
                    End If
                Next llRow

                ilSpots = ilNoSpotsPerWk \ ilWkCount(ilDateLoop)
                If ilSpots <= 0 Then
                    ilSpots = 1
                End If
                llDate = llWkDate(ilDateLoop)
                ilCount = 0
                ilNoSpots = 0
                'For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
                For llGame = 0 To UBound(tlGameSort) - 1 Step 1
                    ilRet = gParseItem(tlGameSort(llGame).sKey, 4, "|", slRow)
                    llRow = Val(slRow)
                    ilRet = gParseItem(tlGameSort(llGame).sKey, 5, "|", slOrigNoSpots)
                    ilOrigNoSpots = Val(slOrigNoSpots)
                    If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                        slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                        If llDate = gDateValue(gObtainPrevMonday(slStr)) Then
                            If (ilAvailCount(ilDateLoop) > 0) And (ilAvailFlag(ilDateLoop) = -1) Then
                                If Trim$(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) <> "" Then
                                    If Val(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) > 0 Then
                                        ilSpots = (CLng(ilNoSpotsPerWk) * Val(grdDates.TextMatrix(llRow, AVAILSORDERINDEX))) \ ilAvailCount(ilDateLoop)
                                        If ilSpots <= 0 Then
                                            ilSpots = 1
                                        End If
                                    Else
                                        ilSpots = -1
                                    End If
                                Else
                                    ilSpots = -1
                                End If
                            End If
                            If ilSpots > 0 Then
                                If ilNoSpots + ilSpots >= ilNoSpotsPerWk Then
                                    If ilNoSpotsPerWk - ilNoSpots > 0 Then
                                        grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX) = Trim$(str$(Val(grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX)) + (ilOrigNoSpots - (ilNoSpotsPerWk - ilNoSpots))))
                                        grdDates.TextMatrix(llRow, NOSPOTSINDEX) = Trim$(str$(ilNoSpotsPerWk - ilNoSpots))
                                        ilNoSpots = ilNoSpotsPerWk
                                    End If
                                    Exit For
                                End If
                                grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX) = Trim$(str$(Val(grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX)) + (ilOrigNoSpots - ilSpots)))
                                grdDates.TextMatrix(llRow, NOSPOTSINDEX) = Trim$(str$(ilSpots))
                                ilNoSpots = ilNoSpots + ilSpots
                            End If
                            ilCount = ilCount + 1
                            If ilCount = ilWkCount(ilDateLoop) Then
                                Exit For
                            End If
                        End If
                    End If
                Next llGame
                'Place remaining spots into games with avails that have room only unless no room exist within the week
                Do While ilNoSpots < ilNoSpotsPerWk
                    ilSpots = 1
                    llDate = llWkDate(ilDateLoop)
                    ilCount = 0
                    'For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
                    For llGame = 0 To UBound(tlGameSort) - 1 Step 1
                        ilRet = gParseItem(tlGameSort(llGame).sKey, 4, "|", slRow)
                        llRow = Val(slRow)
                        ilRet = gParseItem(tlGameSort(llGame).sKey, 5, "|", slOrigNoSpots)
                        ilOrigNoSpots = Val(slOrigNoSpots)
                        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                            slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                            If llDate = gDateValue(gObtainPrevMonday(slStr)) Then
                                If Trim$(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) <> "" Then
                                    'If Val(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) > 0 Then
                                    If (Val(grdDates.TextMatrix(llRow, AVAILSORDERINDEX)) > 0) Or (ilAvailCount(ilDateLoop) = 0) Then
                                        If ilNoSpots + ilSpots > ilNoSpotsPerWk Then
                                            Exit For
                                        End If
                                        ilCount = ilCount + 1
                                        ilCSpots = Val(grdDates.TextMatrix(llRow, NOSPOTSINDEX)) + ilSpots
                                        grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX) = Trim$(str$(Val(grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX)) - ilSpots))
                                        grdDates.TextMatrix(llRow, NOSPOTSINDEX) = Trim$(str$(ilCSpots))
                                        ilNoSpots = ilNoSpots + ilSpots
                                        If (ilCount = ilWkCount(ilDateLoop)) Or (ilNoSpots = ilNoSpotsPerWk) Then
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next llGame
                Loop
            Next ilDateLoop
        End If
    End If
    bmSpotsBySet = True
    grdDates.Redraw = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdDates, vbDefault
End Sub

Private Sub cmcSetSpots_GotFocus()
    mWSSetShow
    mSpecSetShow
    mSetShow
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If Not mSpecFormFieldsOk() Then
        cmcCancel.SetFocus
    End If
End Sub

Private Sub cmcSpec_Click()
    Select Case grdSpec.Row
        Case SPECROW3INDEX
            Select Case grdSpec.Col
                Case MODELLNINDEX
                    lbcLnModel.Visible = Not lbcLnModel.Visible
                Case SEASONINDEX
                    lbcSeason.Visible = Not lbcSeason.Visible
            End Select
    End Select
    edcSpec.SelStart = 0
    edcSpec.SelLength = Len(edcSpec.Text)
    edcSpec.SetFocus
End Sub
Private Sub cmcSpec_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub



Private Sub edcNoSpots_Change()

    grdDates.CellForeColor = vbBlack
End Sub


Private Sub edcNoSpots_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcNoSpots_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case lmEnableCol
        Case NOSPOTSINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcNoSpots.Text
            slStr = Left$(slStr, edcNoSpots.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcNoSpots.SelStart - edcNoSpots.SelLength)
            If gCompNumberStr(slStr, "999") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub edcSpec_Change()
    If Not imWSCtrlVisible Then
        Select Case grdSpec.Row
            Case SPECROW3INDEX
                Select Case grdSpec.Col
                    Case MODELLNINDEX
                        imLbcArrowSetting = True
                        gMatchLookAhead edcSpec, lbcLnModel, imBSMode, imComboBoxIndex
                    Case SEASONINDEX
                        imLbcArrowSetting = True
                        gMatchLookAhead edcSpec, lbcSeason, imBSMode, imComboBoxIndex
                End Select
        End Select
    End If
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub edcSpec_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSpec_KeyPress(KeyAscii As Integer)
    Dim slStr As String

    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcSpec.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    If Not imWSCtrlVisible Then
        Select Case grdSpec.Row
            Case SPECROW6INDEX
                Select Case grdSpec.Col
                    Case NOSPOTSPERINDEX
                        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                End Select
            Case SPECROW9INDEX
                Select Case grdSpec.Col
                    Case GAMENOSINDEX
                        If (KeyAscii = KEYBACKSPACE) Or (KeyAscii = KEYNEG) Or ((KeyAscii >= KEY0) And (KeyAscii <= KEY9)) Or (KeyAscii = KEYCOMMA) Then
                        Else
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Case GAMEININDEX
                        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                        slStr = edcSpec.Text
                        slStr = Left$(slStr, edcSpec.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSpec.SelStart - edcSpec.SelLength)
                        If gCompNumberStr(slStr, Trim$(str$(imMaxGameNo))) > 0 Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Case GAMEOUTINDEX
                        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                        slStr = edcSpec.Text
                        slStr = Left$(slStr, edcSpec.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSpec.SelStart - edcSpec.SelLength)
                        If gCompNumberStr(slStr, Trim$(str$(imMaxGameNo))) > 0 Then
                            Beep
                            KeyAscii = 0
                            Exit Sub
                        End If
                End Select
        End Select
    Else
        Select Case grdWS.Row
            Case WSSPOTSINDEX
                If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub edcSpec_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If Not imWSCtrlVisible Then
            Select Case grdSpec.Row
                Case SPECROW3INDEX
                    Select Case grdSpec.Col
                        Case MODELLNINDEX
                            gProcessArrowKey Shift, KeyCode, lbcLnModel, imLbcArrowSetting
                        Case SEASONINDEX
                            gProcessArrowKey Shift, KeyCode, lbcSeason, imLbcArrowSetting
                    End Select
            End Select
        End If
        edcSpec.SelStart = 0
        edcSpec.SelLength = Len(edcSpec.Text)
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
    imUpdateAllowed = igUpdateAllowed
    'If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
    If Not imUpdateAllowed Then
        pbcSpecSTab.Enabled = False
        pbcSpecTab.Enabled = False
        '12/19/12: Required to switch seasons
        'grdSpec.Enabled = False
        grdSpec.Enabled = True
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        'grdDates.Enabled = False
        'Required to allow scrolling
        grdDates.Enabled = True
        'imUpdateAllowed = False
    Else
        grdSpec.Enabled = True
        pbcSpecSTab.Enabled = True
        pbcSpecTab.Enabled = True
        grdDates.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        'imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    Me.Refresh
    imLoadingForm = False
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_Initialize()
    imLoadingForm = True
    Me.Width = (CLng(80) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(80) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    gCenterStdAlone CGameSch
    'DoEvents
    mSetControls
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
    imLoadingForm = True
    imInitNoRows = grdDates.Rows
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    
    Erase tmGsf
    Erase tmPropGameInfo
    Erase tmSvCff
    Erase tmSvCgf

    smLanguageCodeTag = ""
    Erase tmLanguageCode
    smTeamCodeTag = ""
    Erase tmTeamCode
    Erase tmSeasonInfo
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    
    Set CGameSch = Nothing   'Remove data segment
    
End Sub

Private Sub grdDates_EnterCell()
    mWSSetShow
    mSpecSetShow
    mSetShow
End Sub

Private Sub grdDates_GotFocus()
    mWSSetShow
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
            edcNoSpots.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
    End If
    lmTopRow = grdDates.TopRow
End Sub

Private Sub grdSpec_EnterCell()
    mWSSetShow
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
    If (ilRow = SPECROW3INDEX - 1) Or (ilRow = SPECROW6INDEX - 1) Or (ilRow = SPECROW9INDEX - 1) Then
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
    If Not mSpecColOk() Then
        grdSpec.Redraw = True
        Exit Sub
    End If
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

Private Sub grdWS_EnterCell()
    mWSSetShow
End Sub

Private Sub grdWS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine if in header
    On Error GoTo grdWSErr:
    If Y < grdWS.RowHeight(0) Then
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilCol = grdWS.MouseCol
    ilRow = grdWS.MouseRow
    If ilCol < grdWS.FixedCols Then
        grdWS.Redraw = True
        Exit Sub
    End If
    If ilRow < grdWS.FixedRows Then
        grdWS.Redraw = True
        Exit Sub
    End If
    DoEvents
    grdWS.Col = ilCol
    grdWS.Row = ilRow
    If Not mWSColOk() Then
        grdWS.Redraw = True
        Exit Sub
    End If
    grdWS.Redraw = True
    mWSEnableBox
    On Error GoTo 0
    Exit Sub
grdWSErr:
    On Error GoTo 0
    If (lmWSEnableRow >= grdWS.FixedRows) And (lmWSEnableRow < grdWS.Rows) Then
        grdWS.Row = lmWSEnableRow
        grdWS.Col = lmWSEnableCol
        mSetFocus
    End If
    grdWS.Redraw = False
    grdWS.Redraw = True
    Exit Sub
End Sub

Private Sub grdWS_Scroll()
    mWSSetShow
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
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    lmEnableRow = grdDates.Row
    lmEnableCol = grdDates.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdDates.Left - pbcArrow.Width - 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + (grdDates.RowHeight(grdDates.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    Select Case grdDates.Col
        Case NOSPOTSINDEX
            edcNoSpots.MaxLength = 3
            edcNoSpots.Text = grdDates.Text
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
'*  ilValue                                                                               *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilLoop As Integer
    Dim slDayPart As String
    Dim slLineNo As String

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdSpec, grdDates, vbHourglass
    imFirstActivate = True
    imTerminate = False
    mSaveFlightInfo
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    lmWSEnableRow = -1
    lmWSEnableCol = -1
    imWSCtrlVisible = False
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
    imCgfChg = False
    imNewGame = True
    imLastColSorted = -1
    imLastSort = -1
    imSpecCtrlVisible = False
    imAvailColorLevel = 90
    smSpotsBy = ""
    bmSpotsBySet = False
    bmSetSpotsPressed = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmFirstAllowedChgDate = lmNowDate '+ 1
    If (tgUrf(0).sChgCntr = "I") Then
        'If (Asc(tgSpf.sUsingFeatures10) And REPLACEDELWKWITHFILLS) = REPLACEDELWKWITHFILLS Then
            If (tgClfCntr(imLnRowNo - 1).lUnbilledDate > 0) And (tgClfCntr(imLnRowNo - 1).lUnbilledDate < lmFirstAllowedChgDate) Then
                lmFirstAllowedChgDate = tgClfCntr(imLnRowNo - 1).lUnbilledDate
            End If
        'End If
    End If
    mLnModelPop
    mInitBox
    hmGhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameSch
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)  'Get and save ARF record length

    hmGsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameSch
    On Error GoTo 0
    ReDim tmGsf(0 To 0) As GSF
    imGsfRecLen = Len(tmGsf(0))  'Get and save ARF record length

    hmCgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameSch
    On Error GoTo 0
    imCgfRecLen = Len(tmCgf)  'Get and save ARF record length

    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameSch
    On Error GoTo 0

    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameSch
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)  'Get and save ARF record length

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameSch
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)  'Get and save ARF record length

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameSch
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save ARF record length

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameSch
    On Error GoTo 0
    imClfRecLen = Len(tmClf)  'Get and save ARF record length

    imVefCode = 0
    imVpfIndex = -1
    lmLLD = -1
    slLineNo = Trim$(str$(tgClfCntr(imLnRowNo - 1).ClfRec.iLine))
    slName = ""
    gFindMatch smLnSave(1, imLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
    If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
        slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet = CP_MSG_NONE Then
            imVefCode = CInt(slCode)
            imVpfIndex = gVpfFind(CGameSch, imVefCode)
            gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLLD
        End If
    End If
    imRdfCode = imLnSave(1, imLnRowNo)
    slDayPart = ""
    For ilLoop = 0 To Contract.lbcPrg(igTabMapIndex).ListCount - 1 Step 1
        If imLnSave(1, imLnRowNo) = Contract.lbcPrg(igTabMapIndex).ItemData(ilLoop) Then
            slDayPart = Contract.lbcLnProgram(igTabMapIndex).List(ilLoop)
            Exit For
        End If
    Next ilLoop
    imCntrLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine
    imLineSpotLen = Val(smLnSave(16, imLnRowNo))
    If imLnSave(4, imLnRowNo) = 1 Then
        lmOvStartTime = CLng(gTimeToCurrency(smLnSave(2, imLnRowNo), False))
        lmOvEndTime = CLng(gTimeToCurrency(smLnSave(3, imLnRowNo), True))
    Else
        lmOvStartTime = 0
        lmOvEndTime = 0
    End If
    plcScreen.Caption = "Event Spots- Line #: " & slLineNo & " Vehicle: " & slName & " Daypart: " & slDayPart
    mTeamPop
    'mLanguagePop
    mClearCtrlFields
    smSpotsBy = ""
    bmSpotsBySet = False
    If tgClfCntr(imLnRowNo - 1).iFirstCgf >= 0 Then
        If tgClfCntr(imLnRowNo - 1).ClfRec.sSportsByWeek = "N" Then
            smSpotsBy = "Event"
            bmSpotsBySet = True
            bmSetSpotsPressed = True
        ElseIf tgClfCntr(imLnRowNo - 1).ClfRec.sSportsByWeek = "W" Then
            smSpotsBy = "Week"
            bmSpotsBySet = True
            bmSetSpotsPressed = True
        End If
    End If
    grdSpec.TextMatrix(SPECROW6INDEX, SPOTSBYINDEX) = smSpotsBy
    mSeasonPop True
    ilRet = mReadRec()
    mMoveRecToCtrl
    If ilRet Then
        mLanguagePop
        mGSFMoveRecToCtrl
        imNewGame = False
    Else
        imNewGame = True
    End If
    
    gGetEventTitles imVefCode, smEventTitle1, smEventTitle2
    
    'CGameSch.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    'gCenterStdAlone CGameSch
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
'*  flTextHeight                  ilLoop                        ilRow                     *
'*  ilCol                                                                                 *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    'flTextHeight = pbcDates.TextHeight("1") - 35


'    grdSpec.Move 180, 255
'    mGridSpecLayout
'    mGridSpecColumnWidths
'    mGridSpecColumns
'    grdSpec.Height = grdSpec.RowPos(grdSpec.Rows - 1) + grdSpec.RowHeight(grdSpec.Rows - 1) + fgPanelAdj - 15
'    cmcSetSpots.Move grdSpec.Left + grdSpec.Width + 120, grdSpec.Top + grdSpec.Height - cmcSetSpots.Height
'    'Merge Columns
'    grdSpec.Row = 3
'    For ilCol = COMMENTINDEX To COMMENTINDEX + 2 Step 1
'        grdSpec.TextMatrix(grdSpec.Row, ilCol) = " "
'    Next ilCol
'    grdSpec.MergeRow(3) = True
'    grdSpec.MergeRow(2) = True
'    grdSpec.MergeCells = 1  '2 work, 3 and 4 don't work
'
'    grdDates.Move grdSpec.Left, grdSpec.Top + grdSpec.Height + 120
'    imInitNoRows = grdDates.Rows
'    mGridLayout
'    mGridColumnWidths
'    mGridColumns
'    grdDates.Height = grdDates.RowPos(grdDates.Rows - 1) + grdDates.RowHeight(grdDates.Rows - 1) + fgPanelAdj - 15
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
            Case NOSPOTSINDEX
                edcNoSpots.Visible = False
                If grdDates.TextMatrix(lmEnableRow, lmEnableCol) <> edcNoSpots.Text Then
                    'grdDates.TextMatrix(lmEnableRow, AVAILSPROPOSALINDEX) = Trim$(str$(Val(grdDates.TextMatrix(lmEnableRow, AVAILSPROPOSALINDEX)) + (Val(grdDates.TextMatrix(lmEnableRow, lmEnableCol)) - Val(edcNoSpots.Text))))
                    'imCgfChg = True
                    mCheckSpots
                End If
                'grdDates.TextMatrix(lmEnableRow, lmEnableCol) = edcNoSpots.Text
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


    Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdDates, vbDefault
    igManUnload = YES
    Unload CGameSch
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
'*                                                                                        *
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
    Dim ilAvail As Integer
    Dim ilInvCount As Integer
    Dim ilPropCount As Integer
    Dim ilStartGameNo As Integer
    Dim ilRunningGameNo As Integer
    Dim slGameNo As String
    Dim llWSRow As Long
    Dim slMoDate As String
    Dim llLoop As Long
    Dim blFound As Boolean
    Dim slDate As String

    grdDates.Redraw = False
    lbcGameNoSort.Clear
    ilColorSet = False
    llRow = grdDates.FixedRows
    imMinGameNo = 0
    imMaxGameNo = 0
    mBuildPropArrays
    For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
        If llRow + 1 > grdDates.Rows Then
            grdDates.AddItem ""
            grdDates.RowHeight(grdDates.Rows - 1) = fgBoxGridH
            grdDates.AddItem ""
            grdDates.RowHeight(grdDates.Rows - 1) = 15
            tmGsf(ilLoop).lCode = 0
            mInitNew llRow
        End If
        grdDates.Row = llRow
        grdDates.TextMatrix(llRow, TMCGFINDEX) = ""
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
        'Feed
        ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
        slStr = ""
        If (ilValue And USINGFEED) = USINGFEED Then
            If tmGsf(ilLoop).sFeedSource = "V" Then
                slStr = smEventTitle1    '"Visiting"
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

        'Avails
        slStr = ""
        gUnpackDate tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), slDate
        ilRet = mAvailCount(slDate, tmGsf(ilLoop).iGameNo, ilInvCount, ilAvail)
        If ilRet Then
            slStr = Trim$(str$(ilAvail))
        End If
        grdDates.TextMatrix(llRow, AVAILSORDERINDEX) = slStr
        grdDates.Col = AVAILSORDERINDEX
        grdDates.CellBackColor = LIGHTYELLOW
        If ilRet Then
            If ilAvail <= ((100 - imAvailColorLevel) * ilInvCount) \ 100 Then
                If ilAvail < 0 Then
                    grdDates.CellForeColor = vbMagenta
                Else
                    grdDates.CellForeColor = DARKYELLOW
                End If
            End If
        End If
        If ilRet Then
            slStr = Trim$(str$(ilInvCount))
        End If
        grdDates.TextMatrix(llRow, INVINDEX) = slStr

        slStr = ""
        If ilRet Then
            ilRet = mPropCount(tmGsf(ilLoop).iGameNo, ilAvail, ilPropCount)
        End If
        If ilRet Then
            slStr = Trim$(str$(ilPropCount))
        End If
        grdDates.TextMatrix(llRow, AVAILSPROPOSALINDEX) = slStr
        grdDates.Col = AVAILSPROPOSALINDEX
        grdDates.CellBackColor = LIGHTYELLOW
        If ilRet Then
            If ilPropCount <= ((100 - imAvailColorLevel) * ilInvCount) \ 100 Then
                If ilPropCount < 0 Then
                    grdDates.CellForeColor = vbMagenta
                Else
                    grdDates.CellForeColor = DARKYELLOW
                End If
            End If
        End If

        'Number of Spots- Obtained from CGF

        gUnpackDate tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), slStr
        grdDates.TextMatrix(llRow, NOSPOTSINDEX) = mGetSpotCounts(imLnRowNo, tmGsf(ilLoop).iGameNo, slStr)
        'Visiting Team
        slStr = ""
        For ilTeam = 0 To UBound(tmTeamCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            slNameCode = tmTeamCode(ilTeam).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tmGsf(ilLoop).iVisitMnfCode = Val(slCode) Then
                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                Exit For
            End If
        Next ilTeam
        grdDates.TextMatrix(llRow, VISITTEAMINDEX) = slStr
        grdDates.Col = VISITTEAMINDEX
        grdDates.CellBackColor = LIGHTYELLOW
        'Home Team
        slStr = ""
        For ilTeam = 0 To UBound(tmTeamCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            slNameCode = tmTeamCode(ilTeam).sKey 'Traffic!lbcAgency.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tmGsf(ilLoop).iHomeMnfCode = Val(slCode) Then
                ilRet = gParseItem(slNameCode, 1, "\", slStr)
                Exit For
            End If
        Next ilTeam
        grdDates.TextMatrix(llRow, HOMETEAMINDEX) = slStr
        grdDates.Col = HOMETEAMINDEX
        grdDates.CellBackColor = LIGHTYELLOW
        'Air Date
        gUnpackDate tmGsf(ilLoop).iAirDate(0), tmGsf(ilLoop).iAirDate(1), slStr
        If gDateValue(slStr) < lmFirstAllowedChgDate Then
            grdDates.Row = llRow
            grdDates.Col = AIRDATEINDEX
            grdDates.CellForeColor = vbRed
            grdDates.Col = AIRTIMEINDEX
            grdDates.CellForeColor = vbRed
            grdDates.Col = NOSPOTSINDEX
            grdDates.CellForeColor = vbRed
        ElseIf gDateValue(slStr) < lmNowDate + 1 Then
            grdDates.Row = llRow
            grdDates.Col = AIRDATEINDEX
            grdDates.CellForeColor = vbRed
            grdDates.Col = AIRTIMEINDEX
            grdDates.CellForeColor = vbRed
        Else
            slGameNo = Trim$(str$(tmGsf(ilLoop).iGameNo))
            Do While Len(slGameNo) < 5
                slGameNo = "0" & slGameNo
            Loop
            lbcGameNoSort.AddItem slGameNo
            lbcGameNoSort.ItemData(lbcGameNoSort.NewIndex) = tmGsf(ilLoop).iGameNo
        End If
        grdDates.TextMatrix(llRow, WEEKOFINDEX) = gObtainPrevMonday(slStr)
        grdDates.Col = WEEKOFINDEX
        grdDates.CellBackColor = LIGHTYELLOW
        grdDates.TextMatrix(llRow, AIRDAYINDEX) = Left(Format(slStr, "ddd"), 2)
        grdDates.Col = AIRDAYINDEX
        grdDates.CellBackColor = LIGHTYELLOW
        
        grdDates.TextMatrix(llRow, AIRDATEINDEX) = slStr
        grdDates.Col = AIRDATEINDEX
        grdDates.CellBackColor = LIGHTYELLOW
        'Air Time
        gUnpackTime tmGsf(ilLoop).iAirTime(0), tmGsf(ilLoop).iAirTime(1), "A", "1", slStr
        grdDates.TextMatrix(llRow, AIRTIMEINDEX) = slStr
        grdDates.Col = AIRTIMEINDEX
        grdDates.CellBackColor = LIGHTYELLOW
        slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
        If gDateValue(slStr) < lmFirstAllowedChgDate Then  'lmNowDate Then
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
        For ilCol = grdDates.FixedCols To grdDates.Cols - 1 Step 1
            grdDates.Col = ilCol
            grdDates.CellForeColor = llColor
        Next ilCol
        'Game Status
        'slStr = ""
        'If tmGsf(ilLoop).sGameStatus = "C" Then
        '    slStr = "Cancel"
        'ElseIf tmGsf(ilLoop).sGameStatus = "F" Then
        '    slStr = "Firm"
        'ElseIf tmGsf(ilLoop).sGameStatus = "P" Then
        '    slStr = "Postpone"
        'ElseIf tmGsf(ilLoop).sGameStatus = "T" Then
        '    slStr = "Tentative"
        'End If
        'grdDates.TextMatrix(llRow, GAMESTATUSINDEX) = slStr
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
    grdDates.Col = AIRTIMEINDEX
    mSortCol AIRTIMEINDEX
    grdDates.Col = AIRDATEINDEX
    mSortCol AIRDATEINDEX
    
    'Build Week Game grid
    grdWS.RowHeight(0) = fgBoxGridH
    llWSRow = grdWS.FixedRows
    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
            slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
            slGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
            If gDateValue(slStr) <= lmNowDate Then
                grdWS.Col = WSSPOTSINDEX
                grdWS.CellBackColor = LIGHTYELLOW
            End If
            slMoDate = gObtainPrevMonday(slStr)
            slStr = slMoDate & "-" & gObtainNextSunday(slMoDate)
            blFound = False
            For llLoop = grdWS.FixedRows To grdWS.Rows - 1 Step 1
                If slStr = grdWS.TextMatrix(llLoop, WSDATESINDEX) Then
                    blFound = True
                    grdWS.TextMatrix(llLoop, WSGAMENUMBERSINDEX) = grdWS.TextMatrix(llLoop, WSGAMENUMBERSINDEX) & "," & slGameNo
                    Exit For
                End If
            Next llLoop
            If Not blFound Then
                If llWSRow >= grdWS.Rows Then
                    grdWS.AddItem ""
                End If
                grdWS.Row = llWSRow
                grdWS.Col = WSDATESINDEX
                grdWS.RowHeight(grdWS.Row) = fgBoxGridH
                grdWS.CellBackColor = LIGHTYELLOW
                grdWS.TextMatrix(llWSRow, WSDATESINDEX) = slStr
                grdWS.TextMatrix(llWSRow, WSGAMENUMBERSINDEX) = slGameNo
                grdWS.TextMatrix(llWSRow, WSMODATEINDEX) = slMoDate
                slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                slStr = gObtainNextSunday(slStr)
                If gDateValue(slStr) <= lmNowDate Then
                    grdWS.Col = WSSPOTSINDEX
                    grdWS.CellBackColor = LIGHTYELLOW
                End If
                llWSRow = llWSRow + 1
            End If
        End If
    Next llRow
    grdDates.Redraw = True
    Exit Sub
End Sub




Private Sub lbcLanguage_Click()
    grdSpec.CellForeColor = vbBlack
End Sub

Private Sub lbcLnModel_Click()
    gProcessLbcClick lbcLnModel, edcSpec, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcLnModel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcSeason_Click()
    gProcessLbcClick lbcSeason, edcSpec, imChgMode, imLbcArrowSetting
End Sub

Private Sub lbcSeason_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcClickFocus_GotFocus()
    mWSSetShow
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



Private Sub pbcComment_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        smComment = "Yes"
        pbcComment_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        smComment = "No"
        pbcComment_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smComment = "Yes" Then
            smComment = "No"
            pbcComment_Paint
        ElseIf smComment = "No" Then
            smComment = "Yes"
            pbcComment_Paint
        End If
    End If
End Sub

Private Sub pbcComment_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smComment = "Yes" Then
        smComment = "No"
        pbcComment_Paint
    Else
        smComment = "Yes"
        pbcComment_Paint
    End If
End Sub

Private Sub pbcComment_Paint()
    pbcComment.Cls
    pbcComment.CurrentX = fgBoxInsetX
    pbcComment.CurrentY = 0 'fgBoxInsetY
    pbcComment.Print smComment
End Sub

Private Sub pbcSpecSTab_GotFocus()
    Dim ilNext As Integer

    If imLoadingForm Then
        Exit Sub
    End If
    If GetFocus() <> pbcSpecSTab.hWnd Then
        Exit Sub
    End If
    If imSpecCtrlVisible Then
        Do
            ilNext = False
            Select Case grdSpec.Row
                Case SPECROW3INDEX
                    Select Case grdSpec.Col
                        Case COMMENTINDEX
                            mSpecSetShow
                            cmcSetSpots.SetFocus
                            Exit Sub
                        Case Else
                            grdSpec.Col = grdSpec.Col - 2
                    End Select
                Case SPECROW6INDEX
                    Select Case grdSpec.Col
                        Case LANGUAGETYPEINDEX
                            grdSpec.Row = SPECROW3INDEX
                            grdSpec.Col = SEASONINDEX
                        Case Else
                            grdSpec.Col = grdSpec.Col - 2
                    End Select
                Case SPECROW9INDEX
                    Select Case grdSpec.Col
                        Case GAMENOSINDEX
                            grdSpec.Row = SPECROW6INDEX
                            grdSpec.Col = NOSPOTSPERINDEX
                        Case Else
                            grdSpec.Col = grdSpec.Col - 2
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
                If grdSpec.Col + 2 < grdSpec.Cols Then
                    grdSpec.Col = grdSpec.Col + 2
                Else
                    Exit Sub
                End If
            End If
        Loop
    End If
    mSpecEnableBox
End Sub

Private Sub pbcSpecTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

    Dim ilNext As Integer
    Dim llSpecEnableRow As Long
    Dim llSpecEnableCol As Long
    Dim llWSEnableRow As Long
    Dim llWSEnableCol As Long

    If GetFocus() <> pbcSpecTab.hWnd Then
        Exit Sub
    End If
    If imWSCtrlVisible Then
        llWSEnableRow = lmWSEnableRow
        llWSEnableCol = lmWSEnableCol
        mWSSetShow
        grdWS.Row = llWSEnableRow
        Do
            ilNext = False
            Select Case llWSEnableCol
                Case WSSPOTSINDEX
                    If grdWS.Row + 1 >= grdWS.Rows Then
                        Exit Do
                    End If
                    If Not grdWS.RowIsVisible(grdWS.Row + 1) Then
                        grdWS.TopRow = grdWS.TopRow + 1
                    End If
                    grdWS.Row = grdWS.Row + 1
                    grdWS.Col = WSSPOTSINDEX
            End Select
            If mWSColOk() Then
                mWSEnableBox
                Exit Sub
            End If
        Loop
    End If
    If imSpecCtrlVisible Then

        If (grdSpec.Row = SPECROW9INDEX) And (grdSpec.Col = GAMENOSINDEX) Then
            llSpecEnableRow = lmSpecEnableRow
            llSpecEnableCol = lmSpecEnableCol
            mSpecSetShow
            lmSpecEnableRow = llSpecEnableRow
            lmSpecEnableCol = llSpecEnableCol
        End If
        Do
            ilNext = False
            Select Case grdSpec.Row
                Case SPECROW3INDEX
                    Select Case grdSpec.Col
                        Case SEASONINDEX
                            grdSpec.Row = SPECROW6INDEX
                            grdSpec.Col = LANGUAGETYPEINDEX
                        Case MODELLNINDEX
                            If lbcLnModel.ListIndex > 0 Then
                                mSpecSetShow
                                cmcSetSpots.SetFocus
                                Exit Sub
                            End If
                            grdSpec.Col = grdSpec.Col + 2
                        Case Else
                            grdSpec.Col = grdSpec.Col + 2
                    End Select
                Case SPECROW6INDEX
                    Select Case grdSpec.Col
                        Case NOSPOTSPERINDEX
                            grdSpec.Row = SPECROW9INDEX
                            grdSpec.Col = GAMENOSINDEX
                        Case Else
                            grdSpec.Col = grdSpec.Col + 2
                    End Select
                Case SPECROW9INDEX
                    Select Case grdSpec.Col
                        Case GAMEOUTINDEX
                            mSpecSetShow
                            cmcSetSpots.SetFocus
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
        grdSpec.Row = grdSpec.Rows - 1
        grdSpec.Col = grdSpec.FixedCols
        Do
            If mSpecColOk() Then
                Exit Do
            Else
                grdDates.Col = grdDates.Col - 2
            End If
        Loop
    End If
    mSpecEnableBox
End Sub

Private Sub pbcSpotsBy_KeyPress(KeyAscii As Integer)
    If (KeyAscii = Asc("W")) Or (KeyAscii = Asc("w")) Then
        smSpotsBy = "Week"
        pbcSpotsBy_Paint
    ElseIf KeyAscii = Asc("E") Or (KeyAscii = Asc("e")) Then
        smSpotsBy = "Event"
        pbcSpotsBy_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smSpotsBy = "Week" Then
            smSpotsBy = "Event"
            pbcSpotsBy_Paint
        ElseIf smSpotsBy = "Event" Then
            smSpotsBy = "Week"
            pbcSpotsBy_Paint
        End If
    End If
End Sub

Private Sub pbcSpotsBy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smSpotsBy = "Week" Then
        smSpotsBy = "Event"
        pbcSpotsBy_Paint
    Else
        smSpotsBy = "Week"
        pbcSpotsBy_Paint
    End If
End Sub

Private Sub pbcSpotsBy_Paint()
    pbcSpotsBy.Cls
    pbcSpotsBy.CurrentX = fgBoxInsetX
    pbcSpotsBy.CurrentY = 0 'fgBoxInsetY
    pbcSpotsBy.Print smSpotsBy
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
                Case NOSPOTSINDEX
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
                    grdDates.Col = NOSPOTSINDEX
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
        grdDates.Col = NOSPOTSINDEX
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
    Dim blSetShowReq As Boolean

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        imTabDirection = 0 'Set- Left to right
        ilTestValue = True
        blSetShowReq = True
        Do
            ilNext = False
            Select Case grdDates.Col
                Case NOSPOTSINDEX
                    blSetShowReq = False
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
                    slStr = grdDates.TextMatrix(grdDates.Row + 2, GAMENOINDEX)
                    If slStr = "" Then
                        'mSetShow
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
                    grdDates.Col = NOSPOTSINDEX
            End Select
            If mColOk() Then
                Exit Do
            Else
                'ilNext = True
                If blSetShowReq Then
                    mSetShow
                End If
                cmcDone.SetFocus
                Exit Sub
            End If
        Loop While ilNext
        If blSetShowReq Then
            mSetShow
        End If
    Else
        imTabDirection = -1  'Set-Right to left
        imSettingValue = True
        lmTopRow = -1
        grdDates.TopRow = grdDates.FixedRows
        grdDates.Col = NOSPOTSINDEX
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
    ilRet = gPopMnfPlusFieldsBox(CGameSch, lbcTeam, tmTeamCode(), smTeamCodeTag, "Z")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mTeamPopErr
        gCPErrorMsg ilRet, "mTeamPop (gPopMnfPlusFieldsBox)", CGameSch
        On Error GoTo 0
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

    ilRet = gPopMnfPlusFieldsBox(CGameSch, lbcLanguage, tmLanguageCode(), smLanguageCodeTag, "L")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLanguagePopErr
        gCPErrorMsg ilRet, "mLanguagePop (gPopMnfPlusFieldsBox)", CGameSch
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
    Exit Sub
mLanguagePopErr:
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

    ReDim tmGsf(0 To 0) As GSF
    ilUpper = 0
    If lmSeasonGhfCode = 0 And lbcSeason.ListCount = 0 Then
        tmGhfSrchKey1.iVefCode = imVefCode
        ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Else
        If lmSeasonGhfCode = 0 Then
            lbcSeason.ListIndex = 0
            lmSeasonGhfCode = lbcSeason.ItemData(0)
        End If
        tmGhfSrchKey0.lCode = lmSeasonGhfCode
        ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    End If
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


    ReDim tmGsf(0 To 0) As GSF
    grdSpec.TextMatrix(SPECROW3INDEX, COMMENTINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX) = ""
    grdSpec.TextMatrix(SPECROW3INDEX, SEASONINDEX) = ""

    grdSpec.TextMatrix(SPECROW6INDEX, LANGUAGETYPEINDEX) = ""
    grdSpec.TextMatrix(SPECROW6INDEX, SPOTSBYINDEX) = ""
    grdSpec.TextMatrix(SPECROW6INDEX, NOSPOTSPERINDEX) = ""

    grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX) = ""
    grdSpec.TextMatrix(SPECROW9INDEX, GAMEININDEX) = ""
    grdSpec.TextMatrix(SPECROW9INDEX, GAMEOUTINDEX) = ""

'    If grdDates.Rows > imInitNoRows Then
'        For ilRow = grdDates.Rows To imInitNoRows Step -1
'            grdDates.RemoveItem (ilRow)
'        Next ilRow
'    End If
'    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
'        grdDates.TextMatrix(ilRow, GAMENOINDEX) = ""
'        For ilCol = grdDates.FixedCols To grdDates.Cols - 1 Step 1
'            grdDates.TextMatrix(ilRow, ilCol) = ""
'        Next ilCol
'    Next ilRow
'    For ilCol = GAMENOINDEX To AIRTIMEINDEX Step 2
'        For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
'            grdDates.Row = ilRow
'            grdDates.CellBackColor = vbWhite
'        Next ilRow
'    Next ilCol
    mClearDateGrid

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
'    ilValue = Asc(tgSpf.sUsingFeatures)  'Option Fields in Orders/Proposals
'    If (ilValue And MULTIMEDIA) = MULTIMEDIA Then 'Using Live Log
'        If Not imCgfChg Then
'            If (UBound(tmGsf) > 0) Then
'                cmcMultimedia.Enabled = True
'            Else
'                cmcMultimedia.Enabled = False
'            End If
'        Else
'            cmcMultimedia.Enabled = False
'        End If
'    Else
'        cmcMultimedia.Enabled = False
'    End If
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
    Dim ilLang As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer

    If (grdSpec.Row < grdSpec.FixedRows) Or (grdSpec.Row >= grdSpec.Rows) Or (grdSpec.Col < grdSpec.FixedCols) Or (grdSpec.Col >= grdSpec.Cols - 1) Then
        Exit Sub
    End If
    If Not imUpdateAllowed Then
        'Exit Sub
        '12/19/12: allow to look at different seasons only
        If (grdSpec.Row <> SPECROW3INDEX) Or (grdSpec.Col <> SEASONINDEX) Then
            Exit Sub
        End If
    End If
    lmSpecEnableRow = grdSpec.Row
    lmSpecEnableCol = grdSpec.Col

    Select Case grdSpec.Row
        Case SPECROW3INDEX
            Select Case grdSpec.Col
                Case COMMENTINDEX
                    smComment = Trim$(grdSpec.Text)
                    If (smComment = "") Or (smComment = "Missing") Then
                        smComment = "No"
                    End If
                    pbcComment_Paint
                Case MODELLNINDEX
                    lbcLnModel.Height = gListBoxHeight(lbcLnModel.ListCount, 10)
                    edcSpec.MaxLength = 0
                    imChgMode = True
                    slStr = grdSpec.Text
                    gFindMatch slStr, 0, lbcLnModel
                    If gLastFound(lbcLnModel) >= 0 Then
                        lbcLnModel.ListIndex = gLastFound(lbcLnModel)
                        edcSpec.Text = lbcLnModel.List(lbcLnModel.ListIndex)
                    Else
                        lbcLnModel.ListIndex = 0
                        edcSpec.Text = lbcLnModel.List(lbcLnModel.ListIndex)
                    End If
                    imChgMode = False
                Case SEASONINDEX
                    lbcSeason.Height = gListBoxHeight(lbcSeason.ListCount, 10)
                    edcSpec.MaxLength = 0
                    imChgMode = True
                    slStr = grdSpec.Text
                    gFindMatch slStr, 0, lbcSeason
                    If gLastFound(lbcSeason) >= 0 Then
                        lbcSeason.ListIndex = gLastFound(lbcSeason)
                        edcSpec.Text = lbcSeason.List(lbcSeason.ListIndex)
                    Else
                        If lbcSeason.ListCount >= 1 Then
                            lbcSeason.ListIndex = 0
                            edcSpec.Text = lbcSeason.List(lbcSeason.ListIndex)
                        End If
                    End If
                    imChgMode = False
            End Select
        Case SPECROW6INDEX
            Select Case grdSpec.Col
                Case LANGUAGETYPEINDEX
                    lbcLanguage.Height = gListBoxHeight(lbcLanguage.ListCount, 10)
                    If lbcLanguage.SelCount <= 0 Then
                        For ilLang = 0 To UBound(tmLanguageCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                            slNameCode = tmLanguageCode(ilLang).sKey 'Traffic!lbcAgency.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            ilCode = Val(slCode)
                            For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
                                If tmGsf(ilLoop).iLangMnfCode = ilCode Then
                                    ilRet = gParseItem(slNameCode, 1, "\", slStr)
                                    gFindMatch slStr, 0, lbcLanguage
                                    If gLastFound(lbcLanguage) >= 0 Then
                                        lbcLanguage.Selected(gLastFound(lbcLanguage)) = True
                                    End If
                                    Exit For
                                End If
                            Next ilLoop
                        Next ilLang
                    End If
                Case SPOTSBYINDEX
                    smSpotsBy = Trim$(grdSpec.Text)
                    If (smSpotsBy = "") Or (smSpotsBy = "Missing") Then
                        smSpotsBy = "Week"
                    End If
                    pbcSpotsBy_Paint
                Case NOSPOTSPERINDEX
                    edcSpec.MaxLength = 0
                    edcSpec.Text = grdSpec.Text
            End Select
        Case SPECROW9INDEX
            Select Case grdSpec.Col
                Case GAMENOSINDEX
                    edcSpec.MaxLength = 0
                    If grdSpec.Text = "" Then
                        edcSpec.Text = smDefaultGameNo  'Trim$(Str$(imMinGameNo)) & "-" & Trim$(Str$(imMaxGameNo))
                    Else
                        edcSpec.Text = grdSpec.Text
                    End If
                Case GAMEININDEX
                    edcSpec.MaxLength = 0
                    edcSpec.Text = grdSpec.Text
                Case GAMEOUTINDEX
                    edcSpec.MaxLength = 0
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilNoGames                     ilOrigUpper                   llRow                     *
'*  llSvRow                       llSvCol                                                 *
'******************************************************************************************

    Dim slStr As String
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilCount As Integer
    Dim ilRet As Integer
    Dim llSeasonGhfCode As Long

    pbcArrow.Visible = False
    If (lmSpecEnableRow >= grdSpec.FixedRows) And (lmSpecEnableRow < grdSpec.Rows) Then
        Select Case lmSpecEnableRow
            Case SPECROW3INDEX
                Select Case lmSpecEnableCol
                    Case COMMENTINDEX
                        pbcComment.Visible = False
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smComment
                    Case MODELLNINDEX
                        edcSpec.Visible = False
                        cmcSpec.Visible = False
                        lbcLnModel.Visible = False
                        If lbcLnModel.ListIndex > 0 Then
                            grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = edcSpec.Text
                            grdSpec.TextMatrix(SPECROW3INDEX, SEASONINDEX) = ""
                            grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX) = ""
                            grdSpec.TextMatrix(SPECROW9INDEX, GAMEININDEX) = ""
                            grdSpec.TextMatrix(SPECROW9INDEX, GAMEOUTINDEX) = ""
                            grdSpec.TextMatrix(SPECROW6INDEX, SPOTSBYINDEX) = ""
                            grdSpec.TextMatrix(SPECROW6INDEX, NOSPOTSPERINDEX) = ""
                            grdSpec.TextMatrix(SPECROW6INDEX, LANGUAGETYPEINDEX) = ""
                        Else
                            grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX) = ""
                        End If
                    Case SEASONINDEX
                        edcSpec.Visible = False
                        cmcSpec.Visible = False
                        lbcSeason.Visible = False
                        If lbcSeason.ListIndex >= 0 Then
                            llSeasonGhfCode = lbcSeason.ItemData(lbcSeason.ListIndex)
                        Else
                            llSeasonGhfCode = 0
                        End If
                        If lmSeasonGhfCode <> llSeasonGhfCode Then
                            lmSeasonGhfCode = llSeasonGhfCode
                            Screen.MousePointer = vbHourglass
                            gSetMousePointer grdSpec, grdDates, vbHourglass
                            mMoveCtrlToRec  'Save values
                            mClearDateGrid
                            ilRet = mReadRec
                            mGSFMoveRecToCtrl
                            Screen.MousePointer = vbDefault
                            gSetMousePointer grdSpec, grdDates, vbDefault
                        End If
                        If lbcSeason.ListIndex >= 0 Then
                            grdSpec.TextMatrix(SPECROW3INDEX, SEASONINDEX) = edcSpec.Text
                        Else
                            grdSpec.TextMatrix(SPECROW3INDEX, SEASONINDEX) = ""
                        End If
                    
                End Select
            Case SPECROW6INDEX
                Select Case lmSpecEnableCol
                    Case LANGUAGETYPEINDEX
                        lbcLanguage.Visible = False
                        If lbcLanguage.SelCount > 0 Then
                            grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX) = ""
                            slStr = ""
                            For ilLoop = 0 To lbcLanguage.ListCount - 1 Step 1
                                If lbcLanguage.Selected(ilLoop) Then
                                    If slStr = "" Then
                                        slStr = lbcLanguage.List(ilLoop)
                                    Else
                                        slStr = slStr & "; " & lbcLanguage.List(ilLoop)
                                    End If
                                End If
                            Next ilLoop
                            grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = slStr
                        Else
                            grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = ""
                        End If
                    Case SPOTSBYINDEX
                        pbcSpotsBy.Visible = False
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> smSpotsBy Then
                            grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX) = ""
                            grdSpec.TextMatrix(SPECROW9INDEX, GAMEININDEX) = ""
                            grdSpec.TextMatrix(SPECROW9INDEX, GAMEOUTINDEX) = ""
                        End If
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = smSpotsBy
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> "" Then
                            grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX) = ""
                            If smSpotsBy = "Week" Then
                                grdSpec.TextMatrix(SPECROW6INDEX - 1, NOSPOTSPERINDEX) = "# Spots/Week"
                                grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMENOSINDEX) = "Events by Week"
                                grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMEININDEX) = ""
                                grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMEOUTINDEX) = ""
                            Else
                                grdSpec.TextMatrix(SPECROW6INDEX - 1, NOSPOTSPERINDEX) = "# Spots/Event"
                                grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMENOSINDEX) = "Events by #"
                                grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMEININDEX) = "Events In"
                                grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMEOUTINDEX) = "Events Out"
                            End If
                        End If
                    Case NOSPOTSPERINDEX
                        edcSpec.Visible = False
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = edcSpec.Text
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> "" Then
                            grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX) = ""
                        End If
                End Select
            Case SPECROW9INDEX
                Select Case lmSpecEnableCol
                    Case GAMENOSINDEX
                        If smSpotsBy <> "Week" Then
                            edcSpec.Visible = False
                            grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = edcSpec.Text
                            If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> "" Then
                                grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX) = ""
                            End If
                        Else
                            grdWS.Visible = False
                            ilCount = 0
                            grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = ""
                            For llRow = grdWS.FixedRows To grdWS.Rows - 1 Step 1
                                If grdWS.TextMatrix(llRow, WSSPOTSINDEX) <> "" Then
                                    If Val(grdWS.TextMatrix(llRow, WSSPOTSINDEX)) > 0 Then
                                        ilCount = ilCount + 1
                                    End If
                                End If
                            Next llRow
                            grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = ilCount & " of " & Trim(str(grdWS.Rows - 1)) & " Selected"
                        End If
                    Case GAMEININDEX
                        edcSpec.Visible = False
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = edcSpec.Text
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> "" Then
                            grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX) = ""
                        End If
                    Case GAMEOUTINDEX
                        edcSpec.Visible = False
                        grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = edcSpec.Text
                        If grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) <> "" Then
                            grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX) = ""
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
        Case NOSPOTSINDEX
            edcNoSpots.Move grdDates.Left + llColPos + 30, grdDates.Top + grdDates.RowPos(grdDates.Row) + 30, grdDates.ColWidth(grdDates.Col) - 30, grdDates.RowHeight(grdDates.Row) - 15
            edcNoSpots.Visible = True
            edcNoSpots.SetFocus
    End Select
    mSetCommands
End Sub

Private Function mColOk() As Integer

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
    'Check if in past
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
Private Sub mSpecSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long
    Dim llRow As Long
    Dim blFound As Boolean

    If (grdSpec.Row < grdSpec.FixedRows) Or (grdSpec.Row >= grdSpec.Rows) Or (grdSpec.Col < grdSpec.FixedCols) Or (grdSpec.Col >= grdSpec.Cols - 1) Then
        Exit Sub
    End If
    imSpecCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdSpec.Col - 1 Step 1
        llColPos = llColPos + grdSpec.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdSpec.ColWidth(grdSpec.Col)
    ilCol = grdSpec.Col
    Do While ilCol < grdSpec.Cols - 1
        If (Trim$(grdSpec.TextMatrix(grdSpec.Row - 1, grdSpec.Col)) <> "") And (Trim$(grdSpec.TextMatrix(grdSpec.Row - 1, grdSpec.Col)) = Trim$(grdSpec.TextMatrix(grdSpec.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdSpec.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdSpec.Row
        Case SPECROW3INDEX
            Select Case grdSpec.Col
                Case COMMENTINDEX
                    pbcComment.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 30, llColWidth, grdSpec.RowHeight(grdSpec.Row)
                    pbcComment_Paint
                    pbcComment.Visible = True
                    pbcComment.SetFocus
                Case MODELLNINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - cmcSpec.Width, grdSpec.RowHeight(grdSpec.Row) - 15
                    cmcSpec.Move edcSpec.Left + edcSpec.Width, edcSpec.Top, cmcSpec.Width, edcSpec.Height
                    lbcLnModel.Move edcSpec.Left, edcSpec.Top + edcSpec.Height, edcSpec.Width + cmcSpec.Width
                    edcSpec.Visible = True
                    cmcSpec.Visible = True
                    edcSpec.SetFocus
                Case SEASONINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - cmcSpec.Width, grdSpec.RowHeight(grdSpec.Row) - 15
                    cmcSpec.Move edcSpec.Left + edcSpec.Width, edcSpec.Top, cmcSpec.Width, edcSpec.Height
                    lbcSeason.Move edcSpec.Left, edcSpec.Top + edcSpec.Height, edcSpec.Width + cmcSpec.Width
                    edcSpec.Visible = True
                    cmcSpec.Visible = True
                    edcSpec.SetFocus
            End Select
        Case SPECROW6INDEX
            Select Case grdSpec.Col
                Case LANGUAGETYPEINDEX
                    lbcLanguage.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col)
                    lbcLanguage.Visible = True
                    lbcLanguage.SetFocus
                Case SPOTSBYINDEX
                    pbcSpotsBy.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 30, grdSpec.ColWidth(grdSpec.Col), grdSpec.RowHeight(grdSpec.Row)
                    pbcSpotsBy_Paint
                    pbcSpotsBy.Visible = True
                    pbcSpotsBy.SetFocus
                Case NOSPOTSPERINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col), grdSpec.RowHeight(grdSpec.Row) - 15
                    edcSpec.Visible = True
                    edcSpec.SetFocus
            End Select
        Case SPECROW9INDEX
            Select Case grdSpec.Col
                Case GAMENOSINDEX
                    If smSpotsBy <> "Week" Then
                        edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col), grdSpec.RowHeight(grdSpec.Row) - 15
                        edcSpec.Visible = True
                        edcSpec.SetFocus
                    Else
                        grdWS.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15 ', grdSpec.ColWidth(grdSpec.Col), grdSpec.RowHeight(grdSpec.Row) - 15
                        If grdWS.Rows * (fgBoxGridH + 15) > grdDates.Height Then
                            grdWS.Height = grdDates.Height
                        Else
                            grdWS.Height = grdWS.Rows * (fgBoxGridH + 15)
                        End If
                        gGrid_IntegralHeight grdWS, fgBoxGridH + 15
                        'grdWS.Height = grdWS.Height - 30
                        grdWS.Visible = True
                        'blFound = False
                        'For llRow = grdWS.FixedRows To grdWS.Rows - 1 Step 1
                        '    grdWS.Row = llRow
                        '    grdWS.Col = WSSPOTSINDEX
                        '    If Not grdWS.RowIsVisible(grdWS.Row) Then
                        '        grdWS.TopRow = grdWS.TopRow
                        '    End If
                        '    If mWSColOk() Then
                        '        blFound = True
                        '        mWSEnableBox
                        '        Exit For
                        '    End If
                        'Next llRow
                        'If Not blFound Then
                        '    cmcCancel.SetFocus
                        'End If
                        grdWS.SetFocus
                    End If
                Case GAMEININDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col), grdSpec.RowHeight(grdSpec.Row) - 15
                    edcSpec.Visible = True
                    edcSpec.SetFocus
                Case GAMEOUTINDEX
                    edcSpec.Move grdSpec.Left + llColPos + 30, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col), grdSpec.RowHeight(grdSpec.Row) - 15
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
    'Layout Fixed Columns: 0=>Edge; 1=Blue border; 2=>Row Title; 3=>Blue border  Note:  This was done this way to allow for horizontal scrolling:  It is not used
    '       Columns: 4=>Input; 5=>Blue column line; 6=>Input; 7=>Blue Column;....
    grdDates.RowHeight(0) = 15
    grdDates.RowHeight(1) = 15
    grdDates.RowHeight(2) = 180
    grdDates.RowHeight(3) = 180
    grdDates.RowHeight(4) = 15
'    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
'        grdDates.RowHeight(ilRow) = fgBoxGridH
'        grdDates.RowHeight(ilRow + 1) = 15
'    Next ilRow
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
    grdSpec.RowHeight(2) = 150
    grdSpec.RowHeight(SPECROW3INDEX) = fgBoxGridH
    grdSpec.RowHeight(4) = 15
    grdSpec.RowHeight(5) = 150
    grdSpec.RowHeight(SPECROW6INDEX) = fgBoxGridH
    grdSpec.RowHeight(7) = 15
    grdSpec.RowHeight(8) = 150
    grdSpec.RowHeight(SPECROW9INDEX) = fgBoxGridH
    grdSpec.RowHeight(10) = 15
    grdSpec.ColWidth(0) = 15
    grdSpec.ColWidth(1) = 15
    grdSpec.ColWidth(3) = 15
    grdSpec.ColWidth(5) = 15
    grdSpec.ColWidth(7) = 15
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

Private Sub mGridColumns()
    Dim ilPos As Integer
    Dim slFirstTitle As String
    Dim slSecondTitle As String
    Dim slStr As String
    
    grdDates.Row = 2
    grdDates.Col = GAMENOINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, GAMENOINDEX) = "Event"
    grdDates.Row = 3
    grdDates.Col = GAMENOINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, GAMENOINDEX) = "#"
    'Feed Source
    grdDates.Row = 2
    grdDates.Col = FEEDSOURCEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, FEEDSOURCEINDEX) = "Feed"
    grdDates.Row = 3
    grdDates.Col = FEEDSOURCEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, FEEDSOURCEINDEX) = "Source"
    'Language
    grdDates.Row = 2
    grdDates.Col = LANGUAGEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, LANGUAGEINDEX) = "Language"
    grdDates.Row = 3
    grdDates.Col = LANGUAGEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, LANGUAGEINDEX) = ""
    'Avails-Ordered
    grdDates.Row = 2
    grdDates.Col = AVAILSORDERINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, AVAILSORDERINDEX) = "Avails"
    grdDates.Row = 3
    grdDates.Col = AVAILSORDERINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, AVAILSORDERINDEX) = "Ordered"
    'Avails-Proposal
    grdDates.Row = 2
    grdDates.Col = AVAILSPROPOSALINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, AVAILSPROPOSALINDEX) = "Avails"
    grdDates.Row = 3
    grdDates.Col = AVAILSPROPOSALINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, AVAILSPROPOSALINDEX) = "Proposal"
    '# Spots
    grdDates.Row = 2
    grdDates.Col = NOSPOTSINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, NOSPOTSINDEX) = "#"
    grdDates.Row = 3
    grdDates.Col = NOSPOTSINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, NOSPOTSINDEX) = "Spots"
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
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, VISITTEAMINDEX) = slFirstTitle   '"Visiting"
    grdDates.Row = 3
    grdDates.Col = VISITTEAMINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
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
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, HOMETEAMINDEX) = slFirstTitle   '"Home"
    grdDates.Row = 3
    grdDates.Col = HOMETEAMINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, HOMETEAMINDEX) = slSecondTitle  '"Team"
    
    'Week of
    grdDates.Row = 2
    grdDates.Col = WEEKOFINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, WEEKOFINDEX) = "Week"
    grdDates.Row = 3
    grdDates.Col = WEEKOFINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, WEEKOFINDEX) = "of"
    ' Air Day
    grdDates.Row = 2
    grdDates.Col = AIRDAYINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, AIRDAYINDEX) = "Air"
    grdDates.Row = 3
    grdDates.Col = AIRDAYINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, AIRDAYINDEX) = "Day"
    
    'Air Date
    grdDates.Row = 2
    grdDates.Col = AIRDATEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, AIRDATEINDEX) = "Air"
    grdDates.Row = 3
    grdDates.Col = AIRDATEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, AIRDATEINDEX) = "Date"
    'Air Time
    grdDates.Row = 2
    grdDates.Col = AIRTIMEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(2, AIRTIMEINDEX) = "Air"
    grdDates.Row = 3
    grdDates.Col = AIRTIMEINDEX
    grdDates.CellFontBold = False
    grdDates.CellFontName = "Arial"
    grdDates.CellFontSize = 6.75
    grdDates.CellForeColor = vbBlue
    grdDates.CellBackColor = LIGHTBLUE  'vbWhite
    grdDates.TextMatrix(3, AIRTIMEINDEX) = "Time"
End Sub

Private Sub mGridColumnWidths()
    Dim ilValue As Integer
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdDates.ColWidth(INVINDEX) = 0
    grdDates.ColWidth(TMCGFINDEX) = 0
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
        grdDates.ColWidth(LANGUAGEINDEX) = 0.07 * grdDates.Width
    Else
        grdDates.ColWidth(LANGUAGEINDEX) = 0
        grdDates.ColWidth(LANGUAGEINDEX + 1) = 0
    End If
    'Avails-Ordered
    grdDates.ColWidth(AVAILSORDERINDEX) = 0.051 * grdDates.Width
    'Avails-Proposal
    If tgSpf.sGUsePropSys = "Y" Then
        grdDates.ColWidth(AVAILSPROPOSALINDEX) = 0.055 * grdDates.Width
    Else
        grdDates.ColWidth(AVAILSPROPOSALINDEX) = 0
        grdDates.ColWidth(AVAILSPROPOSALINDEX + 1) = 0
    End If
    '# Spots
    grdDates.ColWidth(NOSPOTSINDEX) = 0.051 * grdDates.Width
    'Visiting Team
    grdDates.ColWidth(VISITTEAMINDEX) = 0.126 * grdDates.Width
    'Home Team
    grdDates.ColWidth(HOMETEAMINDEX) = 0.126 * grdDates.Width
    'Week of
    grdDates.ColWidth(WEEKOFINDEX) = 0.065 * grdDates.Width
    'Air Day
    grdDates.ColWidth(AIRDAYINDEX) = 0.035 * grdDates.Width
    'Air Date
    grdDates.ColWidth(AIRDATEINDEX) = 0.065 * grdDates.Width
    'Air Time
    grdDates.ColWidth(AIRTIMEINDEX) = 0.083 * grdDates.Width
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
End Sub

Private Sub mGridSpecColumns()
    Dim ilCol As Integer
    Dim ilValue As Integer

    grdSpec.Row = SPECROW3INDEX - 1
    grdSpec.Col = MODELLNINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    If lbcLnModel.ListCount > 1 Then
        grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = "Model from Line"
    Else
        grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = "Model from Line"
    End If
    grdSpec.Row = SPECROW3INDEX - 1
    grdSpec.Col = COMMENTINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(grdSpec.Row, COMMENTINDEX) = "Show Event List on Printout"
    grdSpec.Row = SPECROW3INDEX - 1
    grdSpec.Col = SEASONINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(grdSpec.Row, SEASONINDEX) = "Season"

    grdSpec.Row = SPECROW6INDEX - 1
    grdSpec.Col = LANGUAGETYPEINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    If (ilValue And USINGLANG) = USINGLANG Then
        grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = "Language"
    Else
        grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = ""
    End If
    grdSpec.Col = SPOTSBYINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = "# Spots by"
    grdSpec.Col = NOSPOTSPERINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = ""   '"# Spots/Week"

    grdSpec.Row = SPECROW9INDEX - 1
    grdSpec.Col = GAMENOSINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = ""   '"Events by #"
    grdSpec.Col = GAMEININDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = ""   '"Games In"
    grdSpec.Col = GAMEOUTINDEX
    grdSpec.CellFontBold = False
    grdSpec.CellFontName = "Arial"
    grdSpec.CellFontSize = 6.75
    grdSpec.CellForeColor = vbBlue
    grdSpec.TextMatrix(grdSpec.Row, grdSpec.Col) = ""   '"Games Out"
    If smSpotsBy = "Week" Then
        grdSpec.TextMatrix(SPECROW6INDEX - 1, NOSPOTSPERINDEX) = "# Spots/Week"
        grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMENOSINDEX) = "Events by Week"
        grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMEININDEX) = ""
        grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMEOUTINDEX) = ""
        grdSpec.TextMatrix(SPECROW6INDEX, SPOTSBYINDEX) = smSpotsBy
    ElseIf smSpotsBy = "Event" Then
        grdSpec.TextMatrix(SPECROW6INDEX - 1, NOSPOTSPERINDEX) = "# Spots/Event"
        grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMENOSINDEX) = "Events by #"
        grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMEININDEX) = "Events In"
        grdSpec.TextMatrix(SPECROW9INDEX - 1, GAMEOUTINDEX) = "Events Out"
        grdSpec.TextMatrix(SPECROW6INDEX, SPOTSBYINDEX) = smSpotsBy
    End If
End Sub

Private Sub mGridSpecColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdSpec.ColWidth(GAMENOSINDEX) = 0.29 * grdSpec.Width
    grdSpec.ColWidth(GAMEININDEX) = 0.29 * grdSpec.Width
    grdSpec.ColWidth(GAMEOUTINDEX) = 0.29 * grdSpec.Width
    llWidth = fgPanelAdj
    llMinWidth = grdSpec.Width
    For ilCol = 0 To grdSpec.Cols - 1 Step 1
        llWidth = llWidth + grdSpec.ColWidth(ilCol)
        If (grdSpec.ColWidth(ilCol) > 15) And (grdSpec.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdSpec.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdSpec.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdSpec.Width
            For ilCol = 0 To grdSpec.Cols - 1 Step 1
                If (grdSpec.ColWidth(ilCol) > 15) And (grdSpec.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdSpec.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdSpec.FixedCols To grdSpec.Cols - 1 Step 1
                If grdSpec.ColWidth(ilCol) > 15 Then
                    ilColInc = grdSpec.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdSpec.ColWidth(ilCol) = grdSpec.ColWidth(ilCol) + 15
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
    Select Case lmEnableCol
        Case LANGUAGETYPEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcLanguage, edcNoSpots, imChgMode, imLbcArrowSetting
        Case VISITTEAMINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcTeam, edcNoSpots, imChgMode, imLbcArrowSetting
        Case HOMETEAMINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcTeam, edcNoSpots, imChgMode, imLbcArrowSetting
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
    Dim ilRow As Integer
    Dim slStr As String
    Dim ilCff As Integer
    Dim ilPrevCff As Integer
    Dim slNoSpots As String
    Dim llDate As Long
    Dim llCffDate As Long
    Dim ilFound As Integer
    Dim ilCgf As Integer
    Dim ilPrevCgf As Integer
    Dim ilGameNo As Integer
    Dim ilUpper As Integer
    Dim ilDay As Integer
    Dim ilWkDay As Integer
    Dim slPriceType As String
    Dim llActPrice As Long
    Dim ilAirDate0 As Integer
    Dim ilAirDate1 As Integer
    Dim llSeasonStart As Long
    Dim llSeasonEnd As Long
    Dim llMoSeasonStart As Long
    Dim llSuSeasonEnd As Long
    Dim llStartDate As Long
    Dim llEndDate As Long

    If smComment = "Yes" Then
        tgClfCntr(imLnRowNo - 1).sGameLayout = "Y"
    Else
        tgClfCntr(imLnRowNo - 1).sGameLayout = "N"
    End If
    tgClfCntr(imLnRowNo - 1).ClfRec.lghfcode = tmGhf.lCode
    gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStart
    gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEnd
    llMoSeasonStart = gDateValue(gObtainPrevMonday(Format(llSeasonStart, "m/d/yy")))
    llSuSeasonEnd = gDateValue(gObtainNextSunday(Format(llSeasonEnd, "m/d/yy")))
    ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
    Do While ilCff <> -1
        gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate
        If (llStartDate >= llMoSeasonStart) And (llStartDate <= llSuSeasonEnd) Then
            tgCffCntr(ilCff).CffRec.iSpotsWk = 0
            For ilDay = 0 To 6 Step 1
                tgCffCntr(ilCff).CffRec.iDay(ilDay) = 0
            Next ilDay
        End If
        ilCff = tgCffCntr(ilCff).iNextCff
    Loop
    ilCgf = tgClfCntr(imLnRowNo - 1).iFirstCgf
    Do While ilCgf <> -1
        gUnpackDateLong tgCgfCntr(ilCgf).CgfRec.iAirDate(0), tgCgfCntr(ilCgf).CgfRec.iAirDate(1), llStartDate
        If (llStartDate >= llSeasonStart) And (llStartDate <= llSeasonEnd) Then
            tgCgfCntr(ilCgf).CgfRec.iNoSpots = 0
        End If
        ilCgf = tgCgfCntr(ilCgf).iNextCgf
    Loop
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
        slNoSpots = grdDates.TextMatrix(ilRow, NOSPOTSINDEX)
        If (grdDates.TextMatrix(ilRow, GAMENOINDEX) <> "") And (Val(slNoSpots) > 0) Then
            slPriceType = "*"
            llActPrice = 0
            ilGameNo = Val(grdDates.TextMatrix(ilRow, GAMENOINDEX))
            slStr = grdDates.TextMatrix(ilRow, AIRDATEINDEX)
            ilWkDay = gWeekDayStr(slStr)
            llDate = gDateValue(gObtainPrevMonday(slStr))
            ilFound = False
            ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
            Do While ilCff <> -1
                'Date
                If tgCffCntr(ilCff).iStatus <> 2 Then
                    gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llCffDate
                    If llDate = llCffDate Then
                        ilFound = True
                        tgCffCntr(ilCff).CffRec.iSpotsWk = tgCffCntr(ilCff).CffRec.iSpotsWk + Val(slNoSpots)
                        For ilDay = 0 To 6 Step 1
                            tgCffCntr(ilCff).CffRec.iDay(ilDay) = 1
                        Next ilDay
                        '11/26/14: Reset status if -1
                        If tgCffCntr(ilCff).iStatus = -1 Then
                            tgCffCntr(ilCff).iStatus = 0
                        End If
                        Exit Do
                    End If
                End If
                ilCff = tgCffCntr(ilCff).iNextCff
            Loop
            If Not ilFound Then
                'Create CFF record in date order
                ilUpper = UBound(tgCffCntr)
                ilPrevCff = -1
                ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                Do While ilCff <> -1
                    'Date
                    If tgCffCntr(ilCff).iStatus <> 2 Then
                        gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llCffDate
                        If llDate < llCffDate Then
                            Exit Do
                        End If
                        ilPrevCff = ilCff
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
                If ilPrevCff = -1 Then
                    ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                    tgClfCntr(imLnRowNo - 1).iFirstCff = ilUpper
                    tgCffCntr(ilUpper).iNextCff = ilCff
                    If ilCff <> -1 Then
                        slPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        llActPrice = tgCffCntr(ilCff).CffRec.lActPrice
                    End If
                Else
                    tgCffCntr(ilPrevCff).iNextCff = UBound(tgCffCntr)
                    tgCffCntr(ilUpper).iNextCff = ilCff
                    If ilCff = -1 Then
                        slPriceType = tgCffCntr(ilPrevCff).CffRec.sPriceType
                        llActPrice = tgCffCntr(ilPrevCff).CffRec.lActPrice
                    Else
                        slPriceType = tgCffCntr(ilCff).CffRec.sPriceType
                        llActPrice = tgCffCntr(ilCff).CffRec.lActPrice
                    End If
                End If
                'Set All CFF value
                tgCffCntr(ilUpper).lRecPos = 0
                tgCffCntr(ilUpper).iStatus = 0
                tgCffCntr(ilUpper).lStartDate = llDate
                tgCffCntr(ilUpper).lEndDate = llDate + 6
                tgCffCntr(ilUpper).lAvgAud = 0
                tgCffCntr(ilUpper).lPriDemoAvgAud = 0
                tgCffCntr(ilUpper).lPriDemoPop = 0
                tgCffCntr(ilUpper).CffRec.lChfCode = tgChfCntr.lCode
                tgCffCntr(ilUpper).CffRec.iClfLine = tgClfCntr(imLnRowNo - 1).ClfRec.iLine
                tgCffCntr(ilUpper).CffRec.iCntRevNo = tgClfCntr(imLnRowNo - 1).ClfRec.iCntRevNo
                tgCffCntr(ilUpper).CffRec.iPropVer = tgClfCntr(imLnRowNo - 1).ClfRec.iPropVer
                gPackDateLong tgCffCntr(ilUpper).lStartDate, tgCffCntr(ilUpper).CffRec.iStartDate(0), tgCffCntr(ilUpper).CffRec.iStartDate(1)
                gPackDateLong tgCffCntr(ilUpper).lEndDate, tgCffCntr(ilUpper).CffRec.iEndDate(0), tgCffCntr(ilUpper).CffRec.iEndDate(1)
                tgCffCntr(ilUpper).CffRec.sDyWk = "W"
                tgCffCntr(ilUpper).CffRec.iSpotsWk = Val(slNoSpots)
                For ilDay = 0 To 6 Step 1
                    tgCffCntr(ilUpper).CffRec.iDay(ilDay) = 1
                Next ilDay
                tgCffCntr(ilUpper).CffRec.sDelete = "N"
                tgCffCntr(ilUpper).CffRec.iXSpotsWk = 0
                tgCffCntr(ilUpper).CffRec.sPriceType = slPriceType   '* used to indicate price needs to be set
                tgCffCntr(ilUpper).CffRec.lActPrice = llActPrice  'Later- might want to store average package price
                tgCffCntr(ilUpper).CffRec.lPropPrice = 0
                ReDim Preserve tgCffCntr(0 To ilUpper + 1) As CFFLIST
                tgCffCntr(ilUpper + 1).iStatus = -1
                tgCffCntr(ilUpper + 1).iNextCff = -1
            End If

            ilFound = False
            slStr = grdDates.TextMatrix(ilRow, AIRDATEINDEX)
            gPackDate slStr, ilAirDate0, ilAirDate1
            ilCgf = tgClfCntr(imLnRowNo - 1).iFirstCgf
            Do While ilCgf <> -1
                'Date
                If tgCgfCntr(ilCgf).iStatus <> 2 Then
                    If (tgCgfCntr(ilCgf).CgfRec.iGameNo = ilGameNo) Then
                        If (tgCgfCntr(ilCgf).CgfRec.iAirDate(0) = ilAirDate0) And (tgCgfCntr(ilCgf).CgfRec.iAirDate(1) = ilAirDate1) Then
                            ilFound = True
                            tgCgfCntr(ilCgf).CgfRec.iNoSpots = Val(slNoSpots)
                            Exit Do
                        End If
                    End If
                End If
                ilCgf = tgCgfCntr(ilCgf).iNextCgf
            Loop
            If Not ilFound Then
                'Create CGF record in game number order within CGF
                ilUpper = UBound(tgCgfCntr)
                ilPrevCgf = -1
                ilCgf = tgClfCntr(imLnRowNo - 1).iFirstCgf
                Do While ilCgf <> -1
                    'Date
                    If tgCgfCntr(ilCgf).iStatus <> 2 Then
                        If ilGameNo < tgCgfCntr(ilCgf).CgfRec.iGameNo Then
                            Exit Do
                        End If
                        ilPrevCgf = ilCgf
                    End If
                    ilCgf = tgCgfCntr(ilCgf).iNextCgf
                Loop
                If ilPrevCgf = -1 Then
                    ilCgf = tgClfCntr(imLnRowNo - 1).iFirstCgf
                    tgClfCntr(imLnRowNo - 1).iFirstCgf = ilUpper
                    tgCgfCntr(ilUpper).iNextCgf = ilCgf
                Else
                    tgCgfCntr(ilPrevCgf).iNextCgf = UBound(tgCgfCntr)
                    tgCgfCntr(ilUpper).iNextCgf = ilCgf
                End If
                'Set All CFF value
                tgCgfCntr(ilUpper).iStatus = 0
                tgCgfCntr(ilUpper).lStartDate = llDate
                tgCgfCntr(ilUpper).lEndDate = llDate + 6
                tgCgfCntr(ilUpper).CgfRec.lCode = 0
                tgCgfCntr(ilUpper).CgfRec.lClfCode = 0
                tgCgfCntr(ilUpper).CgfRec.iGameNo = ilGameNo
                tgCgfCntr(ilUpper).CgfRec.iNoSpots = Val(slNoSpots)
                slStr = grdDates.TextMatrix(ilRow, AIRDATEINDEX)
                gPackDate slStr, tgCgfCntr(ilUpper).CgfRec.iAirDate(0), tgCgfCntr(ilUpper).CgfRec.iAirDate(1)
                ReDim Preserve tgCgfCntr(0 To ilUpper + 1) As CGFLIST
                tgCgfCntr(ilUpper + 1).iStatus = -1
                tgCgfCntr(ilUpper + 1).iNextCgf = -1
            End If
        End If
    Next ilRow
    Exit Sub
End Sub


Private Sub mLnModelPop()
    Dim ilRowNo As Integer
    Dim ilVefCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim slStr As String

    'For ilRowNo = LBound(smLnSave, 2) To UBound(smLnSave, 2) Step 1
    For ilRowNo = imLB1Or2 To UBound(smLnSave, LINEBOUNDINDEX) Step 1
        If ilRowNo <> imLnRowNo Then
            gFindMatch smLnSave(1, ilRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
            If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    ilVefCode = CInt(slCode)
                    ilVef = gBinarySearchVef(ilVefCode)
                    If ilVef <> -1 Then
                        If tgMVef(ilVef).sType = "G" Then
                            slName = Trim$(tgMVef(ilVef).sName)
                            If tgClfCntr(ilRowNo - 1).ClfRec.iLine > 0 Then
                                slStr = Trim$(str$(tgClfCntr(ilRowNo - 1).ClfRec.iLine)) & "-"
                                lbcLnModel.AddItem slStr & slName
                            Else
                                'gFindMatch slStr, 0, lbcLnModel
                                'If gLastFound(lbcLnModel) < 0 Then
                                '   lbcLnModel.AddItem slName
                                'End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next ilRowNo
    lbcLnModel.AddItem "[None]", 0
End Sub

Private Function mSpecColOk() As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilValue As Integer

    mSpecColOk = True
    If grdSpec.ColWidth(grdSpec.Col) <= 15 Then
        mSpecColOk = False
        Exit Function
    End If
    If grdSpec.CellBackColor = LIGHTYELLOW Then
        mSpecColOk = False
        Exit Function
    End If

    If (grdSpec.Row = SPECROW9INDEX) And ((grdSpec.Col = GAMEININDEX) Or (grdSpec.Col = GAMEOUTINDEX)) Then
        If smSpotsBy <> "Week" Then
            slStr = grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX)
            ilPos = InStr(1, slStr, "-", vbTextCompare)
            If ilPos <= 0 Then
                mSpecColOk = False
                Exit Function
            End If
        Else
            mSpecColOk = False
            Exit Function
        End If
    End If
    If (grdSpec.Row = SPECROW3INDEX) And (grdSpec.Col = MODELLNINDEX) And (lbcLnModel.ListCount <= 1) Then
        mSpecColOk = False
        Exit Function
    End If
    ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
    If (grdSpec.Row = SPECROW6INDEX) And (grdSpec.Col = LANGUAGETYPEINDEX) And ((ilValue And USINGLANG) <> USINGLANG) Then
        mSpecColOk = False
        Exit Function
    End If
    If (grdSpec.Row = SPECROW6INDEX) And (grdSpec.Col = SPOTSBYINDEX) And (bmSpotsBySet) Then
        mSpecColOk = False
        Exit Function
    End If
    If Not imUpdateAllowed Then
        If (grdSpec.Row <> SPECROW3INDEX) Or (grdSpec.Col <> SEASONINDEX) Then
            mSpecColOk = False
            Exit Function
        End If
    End If
End Function

Private Function mGetSpotCounts(ilLnRowNo As Integer, ilGameNo As Integer, slAirDate As String) As String
    Dim ilCgf As Integer
    Dim ilAirDate0 As Integer
    Dim ilAirDate1 As Integer

    mGetSpotCounts = ""
    gPackDate slAirDate, ilAirDate0, ilAirDate1
    ilCgf = tgClfCntr(ilLnRowNo - 1).iFirstCgf
    Do While ilCgf <> -1
        If tgCgfCntr(ilCgf).iStatus <> 2 Then
            
            If tgCgfCntr(ilCgf).CgfRec.iGameNo = ilGameNo Then
                If (tgCgfCntr(ilCgf).CgfRec.iAirDate(0) = ilAirDate0) And (tgCgfCntr(ilCgf).CgfRec.iAirDate(1) = ilAirDate1) Then
                    mGetSpotCounts = Trim$(str$(tgCgfCntr(ilCgf).CgfRec.iNoSpots))
                    Exit Function
                End If
            End If
        End If
        ilCgf = tgCgfCntr(ilCgf).iNextCgf
    Loop
End Function

Private Sub mMoveRecToCtrl()
    If tgClfCntr(imLnRowNo - 1).sGameLayout = "Y" Then
        smComment = "Yes"
    Else
        smComment = "No"
    End If
    grdSpec.TextMatrix(SPECROW3INDEX, COMMENTINDEX) = smComment
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSpecFormFieldsOk               *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mSpecFormFieldsOk() As Integer
'
'   iRet = mSpecFormFieldsOk()
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
    Dim slLineNo  As String
    Dim ilLineNo As Integer
    Dim ilLine As Integer

    ilError = False
    slStr = Trim$(grdSpec.TextMatrix(SPECROW3INDEX, COMMENTINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        ilError = True
        grdSpec.TextMatrix(SPECROW3INDEX, COMMENTINDEX) = "Missing"
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = COMMENTINDEX
        grdSpec.CellForeColor = vbMagenta
    End If
    slStr = grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX)
    If (slStr <> "") And (slStr <> "[None]") Then
        gFindMatch slStr, 0, lbcLnModel
        If gLastFound(lbcLnModel) >= 0 Then
            ilPos = InStr(1, slStr, "-", vbTextCompare)
            If ilPos > 0 Then
                slLineNo = Left$(slStr, ilPos - 1)
                ilLineNo = Val(slLineNo)
                'For ilLine = LBound(smLnSave, 2) To UBound(smLnSave, 2) Step 1
                For ilLine = imLB1Or2 To UBound(smLnSave, LINEBOUNDINDEX) Step 1
                    If tgClfCntr(ilLine - 1).ClfRec.iLine = ilLineNo Then
                        mSpecFormFieldsOk = True
                        Exit Function
                    End If
                Next ilLine
            End If
        End If
        grdSpec.Row = SPECROW3INDEX
        grdSpec.Col = MODELLNINDEX
        grdSpec.CellForeColor = vbMagenta
        mSpecFormFieldsOk = False
        Exit Function
    Else
        slStr = grdSpec.TextMatrix(SPECROW3INDEX, SEASONINDEX)
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdSpec.TextMatrix(SPECROW3INDEX, SEASONINDEX) = "Missing"
            grdSpec.Row = SPECROW3INDEX
            grdSpec.Col = SEASONINDEX
            grdSpec.CellForeColor = vbMagenta
        End If
        ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
        If (ilValue And USINGLANG) = USINGLANG Then
            slStr = grdSpec.TextMatrix(SPECROW6INDEX, LANGUAGETYPEINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdSpec.TextMatrix(SPECROW6INDEX, LANGUAGETYPEINDEX) = "Missing"
                grdSpec.Row = SPECROW6INDEX
                grdSpec.Col = LANGUAGETYPEINDEX
                grdSpec.CellForeColor = vbMagenta
            End If
        End If
        slStr = grdSpec.TextMatrix(SPECROW6INDEX, NOSPOTSPERINDEX)
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdSpec.TextMatrix(SPECROW6INDEX, NOSPOTSPERINDEX) = "Missing"
            grdSpec.Row = SPECROW6INDEX
            grdSpec.Col = NOSPOTSPERINDEX
            grdSpec.CellForeColor = vbMagenta
        End If
        slStr = grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX)
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            ilError = True
            grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX) = "Missing"
            grdSpec.Row = SPECROW9INDEX
            grdSpec.Col = GAMENOSINDEX
            grdSpec.CellForeColor = vbMagenta
        End If
    End If
    grdSpec.Row = SPECROW9INDEX
    grdSpec.Col = GAMENOSINDEX
    If grdSpec.CellForeColor <> vbMagenta Then
        slStr = grdSpec.TextMatrix(SPECROW3INDEX, MODELLNINDEX)
        If slStr <> "" Then
            gFindMatch slStr, 0, lbcLnModel
            If gLastFound(lbcLnModel) >= 0 Then
                ilPos = InStr(1, slStr, "-", vbTextCompare)
                If ilPos > 0 Then
                    slLineNo = Left$(slStr, ilPos - 1)
                    ilLineNo = Val(slLineNo)
                    'For ilLine = LBound(smLnSave, 2) To UBound(smLnSave, 2) Step 1
                    For ilLine = imLB1Or2 To UBound(smLnSave, LINEBOUNDINDEX) Step 1
                        If tgClfCntr(ilLine - 1).ClfRec.iLine = ilLineNo Then
                            For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
                                If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                                    ilGameNo = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                                    slStr = mGetSpotCounts(ilLine, ilGameNo, grdDates.TextMatrix(llRow, AIRDATEINDEX))
                                    If slStr <> "" Then
                                        slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                                        If gDateValue(slStr) <= lmNowDate Then
                                            ilError = True
                                            grdSpec.Row = SPECROW3INDEX
                                            grdSpec.Col = MODELLNINDEX
                                            grdSpec.CellForeColor = vbMagenta
                                        End If
                                    End If
                                End If
                            Next llRow
                            Exit For
                        End If
                    Next ilLine
                End If
            End If
        Else
            'slStr = grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX)
            If grdSpec.TextMatrix(SPECROW6INDEX, SPOTSBYINDEX) <> "Week" Then
                slStr = grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX)
            Else
                slStr = ""
                For llRow = grdWS.FixedRows To grdWS.Rows - 1 Step 1
                    If grdWS.TextMatrix(llRow, WSSPOTSINDEX) <> "" Then
                        If slStr = "" Then
                            slStr = grdWS.TextMatrix(llRow, WSGAMENUMBERSINDEX)
                        Else
                            slStr = slStr & "," & grdWS.TextMatrix(llRow, WSGAMENUMBERSINDEX)
                        End If
                    End If
                Next llRow
            End If
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
                        For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
                            If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                                ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                                If ilGameNo = ilGame Then
                                    slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                                    If gDateValue(slStr) <= lmNowDate Then
                                        ilError = True
                                        grdSpec.Row = SPECROW9INDEX
                                        grdSpec.Col = GAMENOSINDEX
                                        grdSpec.CellForeColor = vbMagenta
                                    End If
                                    Exit For
                                End If
                            End If
                        Next llRow
                    Next ilGameNo
                Else
                    For llRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
                        If (grdDates.TextMatrix(llRow, GAMENOINDEX) <> "") Then
                            ilGame = Val(grdDates.TextMatrix(llRow, GAMENOINDEX))
                            If Val(slFields(ilLoop)) = ilGame Then
                                slStr = grdDates.TextMatrix(llRow, AIRDATEINDEX)
                                If gDateValue(slStr) <= lmNowDate Then
                                    ilError = True
                                    grdSpec.Row = SPECROW9INDEX
                                    grdSpec.Col = GAMENOSINDEX
                                    grdSpec.CellForeColor = vbMagenta
                                End If
                                Exit For
                            End If
                        End If
                    Next llRow
                End If
            Next ilLoop
        End If
    Else
        ilError = True
    End If
    If ilError Then
        mSpecFormFieldsOk = False
    Else
        mSpecFormFieldsOk = True
    End If
End Function

Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    mSpecSetShow
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
            ElseIf ilCol = WEEKOFINDEX Then
                slStr = Trim$(str$(gDateValue(grdDates.TextMatrix(llRow, WEEKOFINDEX))))
                Do While Len(slStr) < 6
                    slStr = "0" & slStr
                Loop
                slSort = Trim$(str$(gDateValue(grdDates.TextMatrix(llRow, AIRDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
                slSort = slStr & slSort
            ElseIf ilCol = AIRDAYINDEX Then
                slStr = gWeekDayStr(grdDates.TextMatrix(llRow, AIRDATEINDEX))
                slSort = Trim$(str$(gDateValue(grdDates.TextMatrix(llRow, AIRDATEINDEX))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
                slSort = slStr & slSort
            ElseIf (ilCol = AIRTIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdDates.TextMatrix(llRow, AIRTIMEINDEX), False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = NOSPOTSINDEX) Then
                slSort = Trim$(grdDates.TextMatrix(llRow, NOSPOTSINDEX))
                Do While Len(slSort) < 4
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

Private Function mAvailCount(slDate As String, ilGameNo As Integer, ilInvCount As Integer, ilAvail As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPass                        slType                        slDate                    *
'*  ilDate0                       ilDate1                       llDate                    *
'*                                                                                        *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilEvt As Integer
    Dim llTime As Long
    Dim il30InvCount As Integer
    Dim il60InvCount As Integer
    Dim il30Avail As Integer
    Dim il60Avail As Integer
    Dim il30Count As Integer
    Dim il60Count As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilLtfCode As Integer
    Dim ilAvailOk As Integer
    Dim ilSpot As Integer
    Dim ilSpotOK As Integer
    Dim ilRdfIndex As Integer
    Dim ilLoop As Integer
    Dim ilDay As Integer
    Dim ilMissedRdfIndex As Integer
    Dim llStartDate As Long
    Dim ilSpotCount As Integer
    Dim llSdfDate As Long
    Dim slCntrType As String
    Dim ilPctTrade As Integer
    Dim ilAnf As Integer

    mAvailCount = False
    ilRdfIndex = gBinarySearchRdf(imRdfCode)
    If ilRdfIndex = -1 Then
        Exit Function
    End If
    If imVpfIndex < 0 Then
        Exit Function
    End If
    llStartDate = 0
    tmSsfSrchKey2.iVefCode = imVefCode
    gPackDate slDate, tmSsfSrchKey2.iDate(0), tmSsfSrchKey2.iDate(1)
    imSsfRecLen = Len(tmSsf)
    ilRet = gSSFGetEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = imVefCode)
        gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llStartDate
        If llStartDate <> gDateValue(slDate) Then
            Exit Do
        End If
        If tmSsf.iType = ilGameNo Then
            ilEvt = 1
            Do While ilEvt <= tmSsf.iCount
               LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                    ilLtfCode = tmProg.iLtfCode
                ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                    'Determine which rate card program this is associated with
                    ilAvailOk = False
                    If (lmOvEndTime > 0) Then
                        If (llTime >= lmOvStartTime) And (llTime < lmOvEndTime) Then
                            ilAvailOk = True
                            'Exit For
                        End If
                    Else
                        If (tgMRdf(ilRdfIndex).iLtfCode(0) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(1) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(2) <> 0) Then
                            If (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(0)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Then
                                ilAvailOk = False    'True- code later
                            End If
                        Else
                            For ilLoop = LBound(tgMRdf(ilRdfIndex).iStartTime, 2) To UBound(tgMRdf(ilRdfIndex).iStartTime, 2) Step 1 'Row
                                If (tgMRdf(ilRdfIndex).iStartTime(0, ilLoop) <> 1) Or (tgMRdf(ilRdfIndex).iStartTime(1, ilLoop) <> 0) Then
                                    gUnpackTimeLong tgMRdf(ilRdfIndex).iStartTime(0, ilLoop), tgMRdf(ilRdfIndex).iStartTime(1, ilLoop), False, llStartTime
                                    gUnpackTimeLong tgMRdf(ilRdfIndex).iEndTime(0, ilLoop), tgMRdf(ilRdfIndex).iEndTime(1, ilLoop), True, llEndTime
                                    'If (llTime >= llStartTime) And (llTime < llEndTime) And (tgMRdf(ilRdfIndex).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                    If (llTime >= llStartTime) And (llTime < llEndTime) And (tgMRdf(ilRdfIndex).sWkDays(ilLoop, ilDay) = "Y") Then
                                        ilAvailOk = True
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                        End If
                    End If
                    If ilAvailOk Then
                        If tgMRdf(ilRdfIndex).sInOut = "I" Then   'Book into
                            If tmAvail.ianfCode <> tgMRdf(ilRdfIndex).ianfCode Then
                                ilAvailOk = False
                            End If
                        ElseIf tgMRdf(ilRdfIndex).sInOut = "O" Then   'Exclude
                            If tmAvail.ianfCode = tgMRdf(ilRdfIndex).ianfCode Then
                                ilAvailOk = False
                            End If
                        End If
                    End If
                    If ilAvailOk Then
                        ilLen = tmAvail.iLen
                        ilUnits = tmAvail.iAvInfo And &H1F
                        ilNo30 = 0
                        ilNo60 = 0
                        If tgVpf(imVpfIndex).sSSellOut = "B" Then
                            'Convert inventory to number of 30's and 60's
                            Do While (ilLen >= 30) And (ilUnits > 0)
                                ilNo30 = ilNo30 + 1
                                ilLen = ilLen - 30
                                ilUnits = ilUnits - 1
                            Loop
                            il30InvCount = il30InvCount + ilNo30
                            ilNo30 = 0
                            ilNo60 = 0
                            ilLen = tmAvail.iLen
                            ilUnits = tmAvail.iAvInfo And &H1F
                            Do While (ilLen >= 60) And (ilUnits > 0)
                                ilNo60 = ilNo60 + 1
                                ilLen = ilLen - 60
                                ilUnits = ilUnits - 1
                            Loop
                            Do While (ilLen >= 30) And (ilUnits > 0)
                                ilNo30 = ilNo30 + 1
                                ilLen = ilLen - 30
                                ilUnits = ilUnits - 1
                            Loop
                            il60InvCount = il60InvCount + ilNo60
                        ElseIf tgVpf(imVpfIndex).sSSellOut = "U" Then
                            'Count 30 or 60 and set flag if neither
                            Do While (ilLen >= 30) And (ilUnits > 0)
                                ilNo30 = ilNo30 + 1
                                ilLen = ilLen - 30
                                ilUnits = ilUnits - 1
                            Loop
                            il30InvCount = il30InvCount + ilNo30
                            ilNo30 = 0
                            ilNo60 = 0
                            ilLen = tmAvail.iLen
                            ilUnits = tmAvail.iAvInfo And &H1F
                            Do While (ilLen >= 60) And (ilUnits > 0)
                                ilNo60 = ilNo60 + 1
                                ilLen = ilLen - 60
                                ilUnits = ilUnits - 1
                            Loop
                            Do While (ilLen >= 30) And (ilUnits > 0)
                                ilNo30 = ilNo30 + 1
                                ilLen = ilLen - 30
                                ilUnits = ilUnits - 1
                            Loop
                            il60InvCount = il60InvCount + ilNo60
                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                            'Count 30 or 60 and set flag if neither
                            If ilLen = 60 Then
                                ilNo60 = 1
                            ElseIf ilLen = 30 Then
                                ilNo30 = 1
                            Else
                            End If
                            il30InvCount = il30InvCount + ilNo30
                            il60InvCount = il60InvCount + ilNo60
                        ElseIf tgVpf(imVpfIndex).sSSellOut = "T" Then
                        End If
                        il30Avail = il30Avail + ilNo30
                        il60Avail = il60Avail + ilNo60
                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                            ilSpotOK = True                             'assume spot is OK to include
    
                            'If (tmSpot.iRank And RANKMASK) = 1010 Then     'DR
                            '    ilSpotOK = False
                            'End If
                            'If (tmSpot.iRank And RANKMASK) = 1020 Then     'Remnant
                            '    ilSpotOK = False
                            'End If
                            'If (tmSpot.iRank And RANKMASK) = 1030 Then     'PI
                            '    ilSpotOK = False
                            'End If
                            If (tmSpot.iRank And RANKMASK) = DIRECTRESPONSERANK Then     'DR
                                If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
                                    ilSpotOK = False
                                End If
                            End If
                            If (tmSpot.iRank And RANKMASK) = REMNANTRANK Then     'Remnant
                                If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEREMNANT) <> AVAILINCLUDEREMNANT Then
                                    ilSpotOK = False
                                End If
                            End If
                            If (tmSpot.iRank And RANKMASK) = PERINQUIRYRANK Then     'PI
                                If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEPERINQUIRY) <> AVAILINCLUDEPERINQUIRY Then
                                    ilSpotOK = False
                                End If
                            End If
                            If (tmSpot.iRank And RANKMASK) = EXTRARANK Then     'Extra
                                ilSpotOK = False
                            End If
                            If (tmSpot.iRank And RANKMASK) = PROMORANK Then     'Promo
                                ilSpotOK = False
                            End If
                            If (tmSpot.iRank And RANKMASK) = PSARANK Then     'PSA
                                ilSpotOK = False
                            End If
                            If (tmSpot.iRank And RANKMASK) = RESERVATIONRANK Then     'Reservation
                                If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
                                    ilSpotOK = False
                                End If
                            End If
                            If (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                ilSpotOK = False
                            End If
                            ilLen = tmSpot.iPosLen And &HFFF
                            If ilSpotOK Then 'continue testing other filters
                                'If (lgSchChfCode > 0) And (ilLineNo > 0) And ((smOrigStatus = "O") Or (smOrigStatus = "H") Or (tgChfCntr.iCntRevNo > 0)) Then
                                '    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                '    ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                '    If (tmSdf.lChfCode = lgSchChfCode) And (tmSdf.iLineNo = ilLineNo) Then
                                '        ilSpotOK = False
                                '    End If
                                'End If
                                tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                'If (tmSdf.lChfCode = tgClfCntr(imLnRowNo - 1).ClfRec.lChfCode) And (tmSdf.iLineNo = imCntrLineNo) Then
                                '    ilSpotOK = False
                                'End If
                                If ilSpotOK Then
                                    '4/12/18
                                    'gGetContractParameters tmSdf.lChfCode, slCntrType, ilPctTrade
                                    If tmChf.lCode <> tmSdf.lChfCode Then
                                        tmChfSrchKey0.lCode = tmSdf.lChfCode
                                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        If ilRet <> BTRV_ERR_NONE Then
                                            ilSpotOK = False
                                        Else
                                            slCntrType = tmChf.sType
                                            ilPctTrade = tmChf.iPctTrade
                                        End If
                                    End If
                                    If (slCntrType = "R") Then     'DR
                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
                                            ilSpotOK = False
                                        End If
                                    End If
                                    If (slCntrType = "T") Then     'Remnant
                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEREMNANT) <> AVAILINCLUDEREMNANT Then
                                            ilSpotOK = False
                                        End If
                                    End If
                                    If (slCntrType = "Q") Then     'PI
                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEPERINQUIRY) <> AVAILINCLUDEPERINQUIRY Then
                                            ilSpotOK = False
                                        End If
                                    End If
                                    If (tmSdf.sSpotType = "X") Then     'Extra
                                        ilSpotOK = False
                                    End If
                                    If (slCntrType = "M") Then     'Promo
                                        ilSpotOK = False
                                    End If
                                    If (slCntrType = "S") Then     'PSA
                                        ilSpotOK = False
                                    End If
                                    If (slCntrType = "V") Then     'Reservation
                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
                                            ilSpotOK = False
                                        End If
                                    End If
                                End If
                                If ilSpotOK Then
                                    ilNo30 = 0
                                    ilNo60 = 0
                                    'ilLen = tmSpot.iPosLen And &HFFF
                                    If tgVpf(imVpfIndex).sSSellOut = "B" Then                   'both units and seconds
                                    'Convert inventory to number of 30's and 60's
                                        Do While ilLen >= 60
                                            ilNo60 = ilNo60 + 1
                                            ilLen = ilLen - 60
                                        Loop
                                        Do While ilLen >= 30
                                            ilNo30 = ilNo30 + 1
                                            ilLen = ilLen - 30
                                        Loop
                                        '5/17/18: Handle case when spot less the 30 sec booked into avail
                                        If (ilLen > 0) And (ilLen < 30) Then
                                            ilNo30 = ilNo30 + 1
                                            ilLen = 0
                                        End If
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "U" Then               'units sold
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        Else
                                            If ilLen <= 30 Then
                                                ilNo30 = 1
                                            Else
                                                ilNo60 = 1
                                            End If
                                        End If
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then               'matching units
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        End If
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "T" Then
                                    End If
                                    il30Count = il30Count + ilNo30
                                    il60Count = il60Count + ilNo60
                                    il60Avail = il60Avail - ilNo60
                                    il30Avail = il30Avail - ilNo30
                                End If                              'ilspotOK
                            End If
                        Next ilSpot                             'loop from ssf file for # spots in avail
                    End If                                          'Avail OK
                    ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                End If
                ilEvt = ilEvt + 1   'Increment to next event
            Loop                                                        'do while ilEvt <= tmSsf.iCount
        End If
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (tgSpf.sCIncludeMissDB = "Y") And (llStartDate > 0) Then
        '5/7/18
        ilAnf = gBinarySearchAnf(tgMRdf(ilRdfIndex).ianfCode, tgAvailAnf())
    
        'Get missed
        tmSdfSrchKey6.iVefCode = imVefCode
        tmSdfSrchKey6.iGameNo = ilGameNo
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey6, INDEXKEY6, BTRV_LOCK_NONE)   'Get first record as starting point
        'This code added as replacement for Ext operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = imVefCode) And (tmSdf.iGameNo = ilGameNo)
            If (tmSdf.sSchStatus <> "S") And (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSdfDate
                If gDateValue(slDate) = llSdfDate Then
                    ilSpotOK = True
                Else
                    ilSpotOK = False
                End If
'                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
'                ilDay = gWeekDayLong(llDate)
'                gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
'                ilSpotOK = False
'                If (lmOvEndTime > 0) Then
'                    If (llTime >= lmOvStartTime) And (llTime < lmOvEndTime) Then
'                        ilSpotOK = True
'                        Exit For
'                    End If
'                Else
'                    If (tgMRdf(ilRdfIndex).iLtfCode(0) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(1) <> 0) Or (tgMRdf(ilRdfIndex).iLtfCode(2) <> 0) Then
'                        If (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(0)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Or (ilLtfCode = tgMRdf(ilRdfIndex).iLtfCode(1)) Then
'                            ilSpotOK = False    'True- code later
'                        End If
'                    Else
'                        For ilLoop = LBound(tgMRdf(ilRdfIndex).iStartTime, 2) To UBound(tgMRdf(ilRdfIndex).iStartTime, 2) Step 1 'Row
'                            If (tgMRdf(ilRdfIndex).iStartTime(0, ilLoop) <> 1) Or (tgMRdf(ilRdfIndex).iStartTime(1, ilLoop) <> 0) Then
'                                gUnpackTimeLong tgMRdf(ilRdfIndex).iStartTime(0, ilLoop), tgMRdf(ilRdfIndex).iStartTime(1, ilLoop), False, llStartTime
'                                gUnpackTimeLong tgMRdf(ilRdfIndex).iEndTime(0, ilLoop), tgMRdf(ilRdfIndex).iEndTime(1, ilLoop), True, llEndTime
'                                If (llTime >= llStartTime) And (llTime < llEndTime) And (tgMRdf(ilRdfIndex).sWkDays(ilLoop, ilDay + 1) = "Y") Then
'                                    ilSpotOK = True
'                                    Exit For
'                                End If
'                            End If
'                        Next ilLoop
'                    End If
'                End If
                If ilSpotOK Then
                    'If (lgSchChfCode > 0) And (ilLineNo > 0) Then
                    '    If (tmSdf.lChfCode = lgSchChfCode) And (tmSdf.iLineNo = ilLineNo) Then
                    '        ilSpotOK = False
                    '    End If
                    'End If

                    'Check if missed spot matches
                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
                    tmClfSrchKey.iLine = tmSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest Revision
                    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))
                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus = "M") Or (tmClf.sSchStatus = "F")) Then
                        ilMissedRdfIndex = gBinarySearchRdf(tmClf.iRdfCode)
                        If ilMissedRdfIndex <> -1 Then
                            If tgMRdf(ilRdfIndex).iCode <> tgMRdf(ilMissedRdfIndex).iCode Then
                                If (tgMRdf(ilMissedRdfIndex).sInOut = "I") And (tgMRdf(ilRdfIndex).sInOut = "I") Then    'Book into
                                    If tgMRdf(ilMissedRdfIndex).ianfCode <> tgMRdf(ilRdfIndex).ianfCode Then
                                        ilSpotOK = False
                                    End If
                                ElseIf (tgMRdf(ilMissedRdfIndex).sInOut = "O") And (tgMRdf(ilRdfIndex).sInOut = "O") Then
                                    If tgMRdf(ilMissedRdfIndex).ianfCode <> tgMRdf(ilRdfIndex).ianfCode Then
                                        ilSpotOK = False
                                    End If
                                Else
                                    '5/7/18
                                    If (tgMRdf(ilMissedRdfIndex).sInOut <> "I") And (tgMRdf(ilMissedRdfIndex).sInOut <> "O") Then    'Book into
                                        If ilAnf <> -1 Then
                                            If tgAvailAnf(ilAnf).sSustain <> "Y" Then
                                                ilSpotOK = False
                                            End If
                                        End If
                                    Else
                                        ilSpotOK = False
                                    End If
                                End If
                            End If
                        End If
                    Else
                        ilSpotOK = False
                    End If
                    If ilSpotOK Then
                        '4/12/18
                        'gGetContractParameters tmSdf.lChfCode, slCntrType, ilPctTrade
                        If tmChf.lCode <> tmSdf.lChfCode Then
                            tmChfSrchKey0.lCode = tmSdf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                ilSpotOK = False
                            Else
                                slCntrType = tmChf.sType
                                ilPctTrade = tmChf.iPctTrade
                            End If
                        End If
                        If (slCntrType = "R") Then     'DR
                            If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
                                ilSpotOK = False
                            End If
                        End If
                        If (slCntrType = "T") Then     'Remnant
                            If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEREMNANT) <> AVAILINCLUDEREMNANT Then
                                ilSpotOK = False
                            End If
                        End If
                        If (slCntrType = "Q") Then     'PI
                            If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEPERINQUIRY) <> AVAILINCLUDEPERINQUIRY Then
                                ilSpotOK = False
                            End If
                        End If
                        If (tmSdf.sSpotType = "X") Then     'Extra
                            ilSpotOK = False
                        End If
                        If (slCntrType = "M") Then     'Promo
                            ilSpotOK = False
                        End If
                        If (slCntrType = "S") Then     'PSA
                            ilSpotOK = False
                        End If
                        If (slCntrType = "V") Then     'Reservation
                            If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
                                ilSpotOK = False
                            End If
                        End If
                    End If
                    If ilSpotOK Then
                        'Determine if Avr created
                        ilNo30 = 0
                        ilNo60 = 0
                        ilLen = tmSdf.iLen
                        If tgVpf(imVpfIndex).sSSellOut = "B" Then
                        'Convert inventory to number of 30's and 60's
                            Do While ilLen >= 60
                                ilNo60 = ilNo60 + 1
                                ilLen = ilLen - 60
                            Loop
                            Do While ilLen >= 30
                                ilNo30 = ilNo30 + 1
                                ilLen = ilLen - 30
                            Loop
                        ElseIf tgVpf(imVpfIndex).sSSellOut = "U" Then
                            'Count 30 or 60 and set flag if neither
                            If ilLen = 60 Then
                                ilNo60 = 1
                            ElseIf ilLen = 30 Then
                                ilNo30 = 1
                            Else
                                If ilLen <= 30 Then
                                    ilNo30 = 1
                                Else
                                    ilNo60 = 1
                                End If
                            End If
                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                            'Count 30 or 60 and set flag if neither
                            If ilLen = 60 Then
                                ilNo60 = 1
                            ElseIf ilLen = 30 Then
                                ilNo30 = 1
                            End If
                        ElseIf tgVpf(imVpfIndex).sSSellOut = "T" Then
                        End If
                        il30Count = il30Count + ilNo30
                        il60Count = il60Count + ilNo60
                        il60Avail = il60Avail - ilNo60
                        il30Avail = il30Avail - ilNo30
                    End If
                End If
            End If
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    'Adjust counts
    If (tgVpf(imVpfIndex).sSSellOut = "B") Then 'And (ilLnLen <= 30) Then                   'both units and seconds
        If il30Avail < 0 Then
            Do While (il60Avail > 0) And (il30Avail < 0)
                il60Avail = il60Avail - 1
                il30Avail = il30Avail + 2
            Loop
        End If
        If il60Avail > 0 Then
            il30Avail = il30Avail + 2 * il60Avail
        End If
    End If
    If (tgVpf(imVpfIndex).sSSellOut = "U") Then 'And (ilLnLen <= 30) Then                   'both units and seconds
        If il30Avail < 0 Then
            Do While (il60Avail > 0) And (il30Avail < 0)
                il60Avail = il60Avail - 1
                il30Avail = il30Avail + 1
            Loop
        End If
    End If
    If (tgVpf(imVpfIndex).sSSellOut = "B") Then 'And (ilLnLen <= 30) Then                   'both units and seconds
        If imLineSpotLen <= 30 Then
            ilSpotCount = il30Count + 2 * il60Count
            ilAvail = il30Avail
            ilInvCount = il30InvCount   'this includes 60 inventory
        Else
            ilSpotCount = 2 * il60Count + il30Count
            ilAvail = il60Avail '- il30Count \ 2
            ilInvCount = 2 * il60InvCount
        End If
    ElseIf (tgVpf(imVpfIndex).sSSellOut = "U") Then 'And (ilLnLen <= 30) Then                   'both units and seconds
        ilSpotCount = 2 * il60Count + il30Count
        ilAvail = il60Avail
        ilInvCount = 2 * il60InvCount
    ElseIf (tgVpf(imVpfIndex).sSSellOut = "M") Then 'And (ilLnLen <= 30) Then                   'both units and seconds
        If imLineSpotLen = 30 Then
            ilSpotCount = il30Count
            ilAvail = il30Avail
            ilInvCount = il30InvCount   'this includes 60 inventory
        ElseIf imLineSpotLen = 60 Then
            ilSpotCount = il60Count
            ilAvail = il60Avail
            ilInvCount = il60InvCount
        End If
    End If
    mAvailCount = True
End Function

Private Function mPropCount(ilGameNo As Integer, ilInAvail As Integer, ilPropCount As Integer) As Integer
    Dim ilLoop As Integer

    ilPropCount = ilInAvail
    mPropCount = False
    If tgSpf.sGUsePropSys <> "Y" Then
        Exit Function
    End If
    For ilLoop = LBound(tmPropGameInfo) To UBound(tmPropGameInfo) - 1 Step 1
        If tmPropGameInfo(ilLoop).iGameNo = ilGameNo Then
            If (tmPropGameInfo(ilLoop).i30NoSpotsProp > 0) Or (tmPropGameInfo(ilLoop).i60NoSpotsProp > 0) Then
                If imLineSpotLen <= 30 Then
                    ilPropCount = ilPropCount + (tmPropGameInfo(ilLoop).i30NoSpotsOrdered - tmPropGameInfo(ilLoop).i30NoSpotsProp - 2 * tmPropGameInfo(ilLoop).i60NoSpotsProp)
                Else
                    ilPropCount = ilPropCount + (tmPropGameInfo(ilLoop).i60NoSpotsOrdered - tmPropGameInfo(ilLoop).i60NoSpotsProp - tmPropGameInfo(ilLoop).i30NoSpotsProp / 2)
                End If
            End If
        End If
    Next ilLoop
    mPropCount = True
End Function

Private Sub mBuildPropArrays()
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilCgf As Integer
    Dim ilGetGames As Integer
    Dim ilNoSpots As Integer
    Dim ilLen As Integer
    Dim ilNo60 As Integer
    Dim ilNo30 As Integer
    Dim ilReSet As Integer
    Dim ilReplace As Integer
    Dim ilSpotCount As Integer

    ReDim tmPropGameInfo(0 To 0) As PROPGAMEINFO
    If tgSpf.sGUsePropSys <> "Y" Then
        Exit Sub
    End If
    tmClfSrchKey3.lghfcode = tmGhf.lCode
    tmClfSrchKey3.iEndDate(0) = 0
    tmClfSrchKey3.iEndDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lghfcode = tmGhf.lCode)
        ilGetGames = 0
        If (tmClf.iRdfCode = imRdfCode) Then
            tmChfSrchKey0.lCode = tmClf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If (tmChf.sDelete <> "Y") And (tmClf.sDelete <> "Y") Then
                If tmChf.sSchStatus = "F" Then
                    ilGetGames = 1
                Else
                    'If (tmChf.sStatus = "C") Or (tmChf.sStatus = "G") Or (tmChf.sStatus = "N") Then
                    If (tmChf.sStatus = "C") Or (tmChf.sStatus = "G") Or (tmChf.sStatus = "N") Or ((tmClf.lCode = tgClfCntr(imLnRowNo - 1).ClfRec.lCode) And ((tmChf.sStatus = "W") Or (tmChf.sStatus = "I"))) Then
                        ilGetGames = 2
                    End If
                End If
            End If
        End If
        If ilGetGames > 0 Then
            ilLen = tmClf.iLen
            ilNo60 = 0
            ilNo30 = 0
            If tgVpf(imVpfIndex).sSSellOut = "B" Then                   'both units and seconds
                'Convert inventory to number of 30's and 60's
                Do While ilLen >= 60
                    ilNo60 = ilNo60 + 1
                    ilLen = ilLen - 60
                Loop
                Do While ilLen >= 30
                    ilNo30 = ilNo30 + 1
                    ilLen = ilLen - 30
                Loop
                If ilNo60 > 0 Then
                    ilNo30 = ilNo30 + 2 * ilNo60
                End If
                If tmClf.iLen <= 30 Then
                    ilNo60 = 0
                Else
                    ilNo30 = 0
                End If
            ElseIf tgVpf(imVpfIndex).sSSellOut = "U" Then               'units sold
                'Count 30 or 60 and set flag if neither
                If ilLen = 60 Then
                    ilNo60 = 1
                ElseIf ilLen = 30 Then
                    ilNo30 = 1
                Else
                    If ilLen <= 30 Then
                        ilNo30 = 1
                    Else
                        ilNo60 = 1
                    End If
                End If
                ilSpotCount = ilNo30 + ilNo60
            ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then               'matching units
                'Count 30 or 60 and set flag if neither
                If ilLen = 60 Then
                    ilNo60 = 1
                ElseIf ilLen = 30 Then
                    ilNo30 = 1
                End If
            ElseIf tgVpf(imVpfIndex).sSSellOut = "T" Then
            End If
            tmCgfSrchKey1.lClfCode = tmClf.lCode
            ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmCgf.lClfCode = tmClf.lCode)
                ilFound = False
                'ilNoSpots = (ilNo60 + ilNo30) * tmCgf.iNoSpots
                ilNoSpots = (ilSpotCount) * tmCgf.iNoSpots
                For ilCgf = 0 To UBound(tmPropGameInfo) - 1 Step 1
                    If (tmPropGameInfo(ilCgf).lCntrNo = tmChf.lCntrNo) And (tmPropGameInfo(ilCgf).iGameNo = tmCgf.iGameNo) Then
                        ilFound = True
                        If ilGetGames = 1 Then
                            tmPropGameInfo(ilCgf).i30NoSpotsOrdered = tmPropGameInfo(ilCgf).i30NoSpotsOrdered + ilNo30 * tmCgf.iNoSpots
                            tmPropGameInfo(ilCgf).i30NoSpotsOrdered = tmPropGameInfo(ilCgf).i60NoSpotsOrdered + ilNo60 * tmCgf.iNoSpots
                        Else
                            If (tmPropGameInfo(ilCgf).iCntRevNo > 0) Or (tmChf.iCntRevNo > 0) Then
                                If (tmPropGameInfo(ilCgf).iCntRevNo = tmChf.iCntRevNo) Then
                                    tmPropGameInfo(ilCgf).i30NoSpotsProp = tmPropGameInfo(ilCgf).i30NoSpotsProp + ilNo30 * tmCgf.iNoSpots
                                    tmPropGameInfo(ilCgf).i60NoSpotsProp = tmPropGameInfo(ilCgf).i60NoSpotsProp + ilNo60 * tmCgf.iNoSpots
                                    tmPropGameInfo(ilCgf).iMnfPotnType = tmChf.iMnfPotnType
                                    tmPropGameInfo(ilCgf).iCntRevNo = tmChf.iCntRevNo
                                    tmPropGameInfo(ilCgf).iPropVer = tmChf.iPropVer
                                ElseIf (tmChf.iCntRevNo > tmPropGameInfo(ilCgf).iCntRevNo) Then
                                    tmPropGameInfo(ilCgf).i30NoSpotsProp = ilNo30 * tmCgf.iNoSpots
                                    tmPropGameInfo(ilCgf).i60NoSpotsProp = ilNo60 * tmCgf.iNoSpots
                                    tmPropGameInfo(ilCgf).iMnfPotnType = tmChf.iMnfPotnType
                                    tmPropGameInfo(ilCgf).iCntRevNo = tmChf.iCntRevNo
                                    tmPropGameInfo(ilCgf).iPropVer = tmChf.iPropVer
                                End If
                            Else
                                If tmChf.iMnfPotnType > 0 Then
                                    If tmPropGameInfo(ilCgf).iMnfPotnType > 0 Then
                                        tmPropGameInfo(ilCgf).i30NoSpotsProp = tmPropGameInfo(ilCgf).i30NoSpotsProp + ilNo30 * tmCgf.iNoSpots
                                        tmPropGameInfo(ilCgf).i60NoSpotsProp = tmPropGameInfo(ilCgf).i60NoSpotsProp + ilNo60 * tmCgf.iNoSpots
                                        tmPropGameInfo(ilCgf).iMnfPotnType = tmChf.iMnfPotnType
                                        tmPropGameInfo(ilCgf).iCntRevNo = tmChf.iCntRevNo
                                        tmPropGameInfo(ilCgf).iPropVer = tmChf.iPropVer
                                    Else
                                        tmPropGameInfo(ilCgf).i30NoSpotsProp = ilNo30 * tmCgf.iNoSpots
                                        tmPropGameInfo(ilCgf).i60NoSpotsProp = ilNo60 * tmCgf.iNoSpots
                                        tmPropGameInfo(ilCgf).iMnfPotnType = tmChf.iMnfPotnType
                                        tmPropGameInfo(ilCgf).iCntRevNo = tmChf.iCntRevNo
                                        tmPropGameInfo(ilCgf).iPropVer = tmChf.iPropVer
                                    End If
                                Else
                                    If tmChf.iPropVer = tmPropGameInfo(ilCgf).iPropVer Then
                                        tmPropGameInfo(ilCgf).i30NoSpotsProp = tmPropGameInfo(ilCgf).i30NoSpotsProp + ilNo30 * tmCgf.iNoSpots
                                        tmPropGameInfo(ilCgf).i60NoSpotsProp = tmPropGameInfo(ilCgf).i60NoSpotsProp + ilNo60 * tmCgf.iNoSpots
                                        tmPropGameInfo(ilCgf).iMnfPotnType = tmChf.iMnfPotnType
                                        tmPropGameInfo(ilCgf).iCntRevNo = tmChf.iCntRevNo
                                        tmPropGameInfo(ilCgf).iPropVer = tmChf.iPropVer
                                    ElseIf tmChf.iPropVer > tmPropGameInfo(ilCgf).iPropVer Then
                                        tmPropGameInfo(ilCgf).i30NoSpotsProp = ilNo30 * tmCgf.iNoSpots
                                        tmPropGameInfo(ilCgf).i60NoSpotsProp = ilNo60 * tmCgf.iNoSpots
                                        tmPropGameInfo(ilCgf).iMnfPotnType = tmChf.iMnfPotnType
                                        tmPropGameInfo(ilCgf).iCntRevNo = tmChf.iCntRevNo
                                        tmPropGameInfo(ilCgf).iPropVer = tmChf.iPropVer
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next ilCgf
                If Not ilFound Then
                    For ilCgf = 0 To UBound(tmPropGameInfo) - 1 Step 1
                        If (tmPropGameInfo(ilCgf).lCntrNo = tmChf.lCntrNo) Then
                            If (tmPropGameInfo(ilCgf).i30NoSpotsOrdered <= 0) And (tmPropGameInfo(ilCgf).i60NoSpotsOrdered <= 0) Then
                                ilReplace = False
                                If ilGetGames = 1 Then
                                    ilReplace = True
                                Else
                                    If (tmPropGameInfo(ilCgf).iCntRevNo > 0) Or (tmChf.iCntRevNo > 0) Then
                                        If (tmChf.iCntRevNo > tmPropGameInfo(ilCgf).iCntRevNo) Then
                                            ilReplace = True
                                        End If
                                    Else
                                        If tmChf.iMnfPotnType > 0 Then
                                            If tmPropGameInfo(ilCgf).iMnfPotnType <= 0 Then
                                                ilReplace = True
                                            End If
                                        Else
                                            If tmChf.iPropVer > tmPropGameInfo(ilCgf).iPropVer Then
                                                ilReplace = True
                                            End If
                                        End If
                                    End If
                                End If
                                'Remove all other game values
                                If ilReplace Then
                                    For ilReSet = 0 To UBound(tmPropGameInfo) - 1 Step 1
                                        If (tmPropGameInfo(ilReSet).lCntrNo = tmChf.lCntrNo) Then
                                            tmPropGameInfo(ilReSet).i30NoSpotsOrdered = 0
                                            tmPropGameInfo(ilReSet).i60NoSpotsProp = 0
                                            tmPropGameInfo(ilReSet).iMnfPotnType = tmChf.iMnfPotnType
                                            tmPropGameInfo(ilReSet).iCntRevNo = tmChf.iCntRevNo
                                            tmPropGameInfo(ilReSet).iPropVer = tmChf.iPropVer
                                        End If
                                    Next ilReSet
                                End If
                            End If
                        End If
                    Next ilCgf
                End If
                If Not ilFound Then
                    tmPropGameInfo(UBound(tmPropGameInfo)).lCntrNo = tmChf.lCntrNo
                    tmPropGameInfo(UBound(tmPropGameInfo)).iGameNo = tmCgf.iGameNo
                    If ilGetGames = 1 Then
                        tmPropGameInfo(UBound(tmPropGameInfo)).i30NoSpotsOrdered = ilNo30 * tmCgf.iNoSpots
                        tmPropGameInfo(UBound(tmPropGameInfo)).i60NoSpotsOrdered = ilNo60 * tmCgf.iNoSpots
                        tmPropGameInfo(UBound(tmPropGameInfo)).i30NoSpotsProp = 0
                        tmPropGameInfo(UBound(tmPropGameInfo)).i60NoSpotsProp = 0
                        tmPropGameInfo(UBound(tmPropGameInfo)).iMnfPotnType = tmChf.iMnfPotnType
                        tmPropGameInfo(UBound(tmPropGameInfo)).iCntRevNo = tmChf.iCntRevNo
                        tmPropGameInfo(UBound(tmPropGameInfo)).iPropVer = tmChf.iPropVer
                    Else
                        tmPropGameInfo(UBound(tmPropGameInfo)).i30NoSpotsOrdered = 0
                        tmPropGameInfo(UBound(tmPropGameInfo)).i60NoSpotsOrdered = 0
                        tmPropGameInfo(UBound(tmPropGameInfo)).i30NoSpotsProp = ilNo30 * tmCgf.iNoSpots
                        tmPropGameInfo(UBound(tmPropGameInfo)).i60NoSpotsProp = ilNo60 * tmCgf.iNoSpots
                        tmPropGameInfo(UBound(tmPropGameInfo)).iMnfPotnType = tmChf.iMnfPotnType
                        tmPropGameInfo(UBound(tmPropGameInfo)).iCntRevNo = tmChf.iCntRevNo
                        tmPropGameInfo(UBound(tmPropGameInfo)).iPropVer = tmChf.iPropVer
                    End If
                    ReDim Preserve tmPropGameInfo(0 To UBound(tmPropGameInfo) + 1) As PROPGAMEINFO
                End If
                ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop

End Sub

Private Sub mSetControls()
    Dim ilGap As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer

    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    cmcDone.Top = Me.Height - cmcDone.Height - 120
    cmcCancel.Top = cmcDone.Top

    cmcCancel.Left = CGameSch.Width / 2 + ilGap / 2
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap

    'grdSpec.Move 180, 255
    'mGridSpecLayout
    'mGridSpecColumnWidths
    'mGridSpecColumns
    grdSpec.Move 180, 255
    mGridSpecLayout
    mGridSpecColumnWidths
    mGridSpecColumns
    grdSpec.Height = grdSpec.RowPos(grdSpec.Rows - 1) + grdSpec.RowHeight(grdSpec.Rows - 1) + fgPanelAdj - 15
    cmcSetSpots.Move grdSpec.Left + grdSpec.Width + 120, grdSpec.Top + grdSpec.Height - cmcSetSpots.Height
    ''Merge Columns
    'grdSpec.Row = 3
    'For ilCol = COMMENTINDEX To COMMENTINDEX + 2 Step 1
    '    grdSpec.TextMatrix(grdSpec.Row, ilCol) = " "
    'Next ilCol
    'grdSpec.MergeRow(3) = True
    'grdSpec.MergeRow(2) = True
    'grdSpec.MergeCells = 1  '2 work, 3 and 4 don't work

    mGridWSLayout
    mGridWSColumnWidths
    mGridWSColumns

    grdDates.Move grdSpec.Left, grdSpec.Top + grdSpec.Height + 120, CGameSch.Width - 2 * grdSpec.Left, cmcDone.Top - grdSpec.Top - grdSpec.Height - 240

    ''imInitNoRows = grdDates.Rows
    'DoEvents
    mGridLayout
    'DoEvents
    mGridColumnWidths
    mGridColumns
    gGrid_IntegralHeight grdDates, fgBoxGridH + 15
    imInitNoRows = grdDates.Rows
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
        If grdDates.RowHeight(ilRow) > 15 Then
            grdDates.Col = GAMENOINDEX
            grdDates.Row = ilRow
            grdDates.CellBackColor = LIGHTYELLOW
        End If
    Next ilRow

End Sub

Private Sub mCheckSpots()
    Dim ilNoSpots As Integer
    Dim ilGameNo As Integer
    Dim ilMGNoSpots As Integer
    Dim llSpotMGDate As Long
    Dim ilRet As Integer
    
    If Trim$(edcNoSpots.Text) <> "" Then
        ilNoSpots = Val(edcNoSpots.Text)
    Else
        ilNoSpots = 0
    End If
    ilMGNoSpots = 0
    If Val(grdDates.TextMatrix(lmEnableRow, lmEnableCol)) > ilNoSpots Then
        If (ilNoSpots <> 0) Or (Asc(tgSpf.sUsingFeatures10) And REPLACEDELWKWITHFILLS) <> REPLACEDELWKWITHFILLS Then
            If gDateValue(grdDates.TextMatrix(lmEnableRow, AIRDATEINDEX)) < lmFirstAllowedChgDate Then 'lmNowDate Then
                MsgBox "Can't be reduced the number of spots as date is in past"
                Exit Sub
            Else
                'Get MG/Outside count in past
                tmSmfSrchKey3.iOrigSchVef = imVefCode
                ilGameNo = Val(grdDates.TextMatrix(lmEnableRow, GAMENOINDEX))
                tmSmfSrchKey3.iGameNo = ilGameNo
                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.iOrigSchVef = imVefCode) And (tmSmf.iGameNo = ilGameNo)
                    If (tmSmf.lChfCode = tgChfCntr.lCode) And (tmSmf.iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine) Then
                        gUnpackDateLong tmSmf.iActualDate(0), tmSmf.iActualDate(1), llSpotMGDate
                        If llSpotMGDate < lmFirstAllowedChgDate Then
                            ilMGNoSpots = ilMGNoSpots + 1
                        End If
                    End If
                    ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If ilMGNoSpots > ilNoSpots Then
                    If Val(grdDates.TextMatrix(lmEnableRow, lmEnableCol)) <> ilMGNoSpots Then
                        MsgBox Val(grdDates.TextMatrix(lmEnableRow, lmEnableCol)) & " can only be reduced to " & ilMGNoSpots & " due to MG's already aired"
                    Else
                        MsgBox "The number of spots can't be reduced due to MG's already aired"
                    End If
                    ilNoSpots = ilMGNoSpots
                End If
            End If
        End If
    End If
    grdDates.TextMatrix(lmEnableRow, AVAILSPROPOSALINDEX) = Trim$(str$(Val(grdDates.TextMatrix(lmEnableRow, AVAILSPROPOSALINDEX)) + (Val(grdDates.TextMatrix(lmEnableRow, lmEnableCol)) - ilNoSpots)))
    imCgfChg = True
    edcNoSpots.Text = ilNoSpots
    grdDates.TextMatrix(lmEnableRow, lmEnableCol) = ilNoSpots
End Sub
Private Sub mGridWSLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdWS.Rows - 1 Step 1
        grdWS.RowHeight(ilRow) = fgBoxGridH 'fgFlexGridRowH
    Next ilRow
    For ilCol = 0 To grdWS.Cols - 1 Step 1
        grdWS.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridWSColumns()
    Dim ilCol As Integer

    grdWS.Row = grdWS.FixedRows - 1
    grdWS.RowHeight(grdWS.Row) = fgBoxGridH
    For ilCol = WSDATESINDEX To WSMODATEINDEX Step 1
        grdWS.Col = ilCol
        grdWS.CellFontBold = False
        grdWS.CellFontName = "Arial"
        grdWS.CellFontSize = 6.75
        'grdWS.CellForeColor = vbBlue
        'grdWS.CellBackColor = LIGHTBLUE
    Next ilCol
    grdWS.TextMatrix(grdWS.Row, WSDATESINDEX) = "Week Dates"
    grdWS.TextMatrix(grdWS.Row, WSSPOTSINDEX) = "# Spots"

End Sub

Private Sub mGridWSColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdSpec.Row = SPECROW9INDEX
    grdWS.Width = grdSpec.ColWidth(GAMENOSINDEX)
    grdWS.ColWidth(WSGAMENUMBERSINDEX) = 0
    grdWS.ColWidth(WSMODATEINDEX) = 0
    grdWS.ColWidth(WSSPOTSINDEX) = 0.3 * grdWS.Width
    grdWS.ColWidth(WSDATESINDEX) = grdWS.Width - grdWS.ColWidth(WSSPOTSINDEX) - GRIDSCROLLWIDTH
'    llWidth = GRIDSCROLLWIDTH + 45
'    llMinWidth = grdWS.Width
'    For ilCol = 0 To grdWS.Cols - 1 Step 1
'        llWidth = llWidth + grdWS.ColWidth(ilCol)
'        If (grdWS.ColWidth(ilCol) > 15) And (grdWS.ColWidth(ilCol) < llMinWidth) Then
'            llMinWidth = grdWS.ColWidth(ilCol)
'        End If
'    Next ilCol
'    llWidth = grdWS.Width - llWidth
'    If llWidth >= 15 Then
'        Do
'            llMinWidth = grdWS.Width
'            For ilCol = 0 To grdWS.Cols - 1 Step 1
'                If (grdWS.ColWidth(ilCol) > 15) And (grdWS.ColWidth(ilCol) < llMinWidth) Then
'                    llMinWidth = grdWS.ColWidth(ilCol)
'                End If
'            Next ilCol
'            For ilCol = grdWS.FixedCols To grdWS.Cols - 1 Step 1
'                If grdWS.ColWidth(ilCol) > 15 Then
'                    ilColInc = grdWS.ColWidth(ilCol) / llMinWidth
'                    For ilLoop = 1 To ilColInc Step 1
'                        grdWS.ColWidth(ilCol) = grdWS.ColWidth(ilCol) + 15
'                        llWidth = llWidth - 15
'                        If llWidth < 15 Then
'                            Exit For
'                        End If
'                    Next ilLoop
'                    If llWidth < 15 Then
'                        Exit For
'                    End If
'                End If
'            Next ilCol
'        Loop While llWidth >= 15
'    End If
End Sub

Private Sub mWSEnableBox()
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
    If (grdWS.Row < grdWS.FixedRows) Or (grdWS.Row >= grdWS.Rows) Or (grdWS.Col < grdWS.FixedCols) Or (grdWS.Col >= grdWS.Cols - 1) Then
        Exit Sub
    End If
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    lmWSEnableRow = grdWS.Row
    lmWSEnableCol = grdWS.Col
    Select Case grdWS.Col
        Case WSSPOTSINDEX
            edcSpec.MaxLength = 3
            If grdWS.Text = "" Then
                edcSpec.Text = grdSpec.TextMatrix(SPECROW6INDEX, NOSPOTSPERINDEX)
            Else
                edcSpec.Text = grdWS.Text
            End If
    End Select
    mWSSetFocus
End Sub

Private Sub mWSSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdWS.Row < grdWS.FixedRows) Or (grdWS.Row >= grdWS.Rows) Or (grdWS.Col < grdWS.FixedCols) Or (grdWS.Col >= grdWS.Cols - 1) Then
        Exit Sub
    End If
    imWSCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdWS.Col - 1 Step 1
        llColPos = llColPos + grdWS.ColWidth(ilCol)
    Next ilCol
    Select Case grdWS.Col
        Case WSSPOTSINDEX
            edcSpec.Move grdWS.Left + llColPos + 15, grdWS.Top + grdWS.RowPos(grdWS.Row) + 15, grdWS.ColWidth(grdWS.Col) - 15, grdWS.RowHeight(grdWS.Row) - 15
            edcSpec.ZOrder
            edcSpec.Visible = True
            edcSpec.SetFocus
    End Select
    'mSetCommands
End Sub

Private Sub mWSSetShow()
'
    Dim llRow As Long
    
    If (lmWSEnableRow >= grdWS.FixedRows) And (lmWSEnableRow < grdWS.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmWSEnableCol
            Case WSSPOTSINDEX
                edcSpec.Visible = False
                If grdWS.TextMatrix(lmWSEnableRow, lmWSEnableCol) <> edcSpec.Text Then
                    grdWS.TextMatrix(lmWSEnableRow, lmWSEnableCol) = edcSpec.Text
                End If
        End Select
    End If
    lmWSEnableRow = -1
    lmWSEnableCol = -1
    imWSCtrlVisible = False
    mSetCommands
End Sub
Private Function mWSColOk() As Integer

    mWSColOk = True
    If grdWS.CellBackColor = LIGHTYELLOW Then
        mWSColOk = False
        Exit Function
    End If
End Function

Private Sub mSeasonPop(blSetSeason As Boolean)
    Dim llStartDate As Long
    Dim slStartDate As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVff As Integer
    
    lbcSeason.Clear
    ReDim tmSeasonInfo(0 To 0) As SEASONINFO
    tmGhfSrchKey1.iVefCode = imVefCode
    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = imVefCode)
        gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llStartDate
        slStartDate = Trim$(str$(llStartDate))
        Do While Len(slStartDate) < 6
            slStartDate = "0" & slStartDate
        Loop
        'If (Not tgClfCntr(imLnRowNo - 1).iLineSchd) Or ((tgClfCntr(imLnRowNo - 1).iLineSchd) And (tgClfCntr(imLnRowNo - 1).ClfRec.lGhfCode = tmGhf.lCode)) Then
            tmSeasonInfo(UBound(tmSeasonInfo)).sKey = slStartDate
            tmSeasonInfo(UBound(tmSeasonInfo)).sSeasonName = tmGhf.sSeasonName
            tmSeasonInfo(UBound(tmSeasonInfo)).lCode = tmGhf.lCode
            ReDim Preserve tmSeasonInfo(0 To UBound(tmSeasonInfo) + 1) As SEASONINFO
        'End If
        ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If UBound(tmSeasonInfo) > 1 Then
        'Sort descending
        ArraySortTyp fnAV(tmSeasonInfo(), 0), UBound(tmSeasonInfo), 1, LenB(tmSeasonInfo(0)), 0, LenB(tmSeasonInfo(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmSeasonInfo) - 1 Step 1
        lbcSeason.AddItem Trim$(tmSeasonInfo(ilLoop).sSeasonName)
        lbcSeason.ItemData(lbcSeason.NewIndex) = tmSeasonInfo(ilLoop).lCode
    Next ilLoop
    
    lmSeasonGhfCode = 0
    If tgClfCntr(imLnRowNo - 1).ClfRec.lghfcode > 0 Then
        lmSeasonGhfCode = tgClfCntr(imLnRowNo - 1).ClfRec.lghfcode
    Else
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If tgVff(ilVff).iVefCode = imVefCode Then
                lmSeasonGhfCode = tgVff(ilVff).lSeasonGhfCode
                Exit For
            End If
        Next ilVff
    End If
    If blSetSeason Then
        For ilLoop = 0 To lbcSeason.ListCount - 1 Step 1
            If lbcSeason.ItemData(ilLoop) = lmSeasonGhfCode Then
                lbcSeason.ListIndex = ilLoop
                grdSpec.TextMatrix(SPECROW3INDEX, SEASONINDEX) = lbcSeason.List(ilLoop)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub mClearDateGrid()
    Dim ilRow As Integer
    Dim ilCol As Integer
    
    grdDates.Redraw = False
    For ilCol = GAMENOINDEX To AIRTIMEINDEX Step 2
        For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 2
            grdDates.Row = ilRow
            grdDates.Col = ilCol
            grdDates.CellBackColor = vbWhite
        Next ilRow
    Next ilCol
    For ilRow = grdDates.FixedRows To grdDates.Rows - 1 Step 1
        grdDates.TextMatrix(ilRow, GAMENOINDEX) = ""
        For ilCol = grdDates.FixedCols To grdDates.Cols - 1 Step 1
            grdDates.TextMatrix(ilRow, ilCol) = ""
        Next ilCol
    Next ilRow
    If grdDates.Rows > imInitNoRows Then
        For ilRow = grdDates.Rows To imInitNoRows Step -1
            grdDates.RemoveItem (ilRow)
        Next ilRow
    End If
    grdDates.Redraw = True
    grdWS.Redraw = False
    For ilCol = WSDATESINDEX To WSMODATEINDEX Step 1
        For ilRow = grdWS.FixedRows To grdWS.Rows - 1 Step 1
            grdWS.Row = ilRow
            grdWS.Col = ilCol
            grdWS.CellBackColor = vbWhite
        Next ilRow
    Next ilCol
    For ilRow = grdWS.FixedRows To grdWS.Rows - 1 Step 1
        For ilCol = grdWS.FixedCols To grdWS.Cols - 1 Step 1
            grdWS.TextMatrix(ilRow, ilCol) = ""
        Next ilCol
    Next ilRow
    If grdWS.Rows > grdWS.FixedRows + 1 Then
        For ilRow = grdWS.Rows To grdWS.FixedRows + 1 Step -1
            grdWS.RemoveItem ilRow
        Next ilRow
    End If
    grdWS.Redraw = True
    grdSpec.TextMatrix(SPECROW6INDEX, NOSPOTSPERINDEX) = ""
    grdSpec.TextMatrix(SPECROW9INDEX, GAMENOSINDEX) = ""
    grdSpec.TextMatrix(SPECROW9INDEX, GAMEININDEX) = ""
    grdSpec.TextMatrix(SPECROW9INDEX, GAMEOUTINDEX) = ""
    
End Sub

Private Sub mSaveFlightInfo()
    Dim ilCff As Integer
    Dim ilCgf As Integer
    
    ReDim tmSvCff(LBound(tgCffCntr) To UBound(tgCffCntr)) As CFFLIST
    For ilCff = LBound(tgCffCntr) To UBound(tgCffCntr) Step 1
        tmSvCff(ilCff) = tgCffCntr(ilCff)
    Next ilCff
    ReDim tmSvCgf(LBound(tgCgfCntr) To UBound(tgCgfCntr)) As CGFLIST
    For ilCgf = LBound(tgCgfCntr) To UBound(tgCgfCntr) Step 1
        tmSvCgf(ilCgf) = tgCgfCntr(ilCgf)
    Next ilCgf
    imSvFirstCff = tgClfCntr(imLnRowNo - 1).iFirstCff
    imSvFirstCgf = tgClfCntr(imLnRowNo - 1).iFirstCgf
End Sub

Private Sub mRestoreFlightInfo()
    Dim ilCff As Integer
    Dim ilCgf As Integer
    If imLnRowNo <= 0 Then
        Exit Sub
    End If
    ReDim tgCffCntr(LBound(tmSvCff) To UBound(tmSvCff)) As CFFLIST
    For ilCff = LBound(tmSvCff) To UBound(tmSvCff) Step 1
        tgCffCntr(ilCff) = tmSvCff(ilCff)
    Next ilCff
    ReDim tgCgfCntr(LBound(tmSvCgf) To UBound(tmSvCgf)) As CGFLIST
    For ilCgf = LBound(tmSvCgf) To UBound(tmSvCgf) Step 1
        tgCgfCntr(ilCgf) = tmSvCgf(ilCgf)
    Next ilCgf
    tgClfCntr(imLnRowNo - 1).iFirstCff = imSvFirstCff
    tgClfCntr(imLnRowNo - 1).iFirstCgf = imSvFirstCgf
End Sub
