VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl CGameInv 
   Appearance      =   0  'Flat
   ClientHeight    =   5895
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   9345
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   9345
   Begin VB.TextBox edcMultiMediaMsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   330
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2025
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "CGameInv.ctx":0000
      Left            =   6270
      List            =   "CGameInv.ctx":0002
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox edcComment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   540
      HelpContextID   =   8
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2295
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.PictureBox pbcComingSoon 
      Height          =   1125
      Left            =   1500
      Picture         =   "CGameInv.ctx":0004
      ScaleHeight     =   1065
      ScaleWidth      =   5445
      TabIndex        =   11
      Top             =   1050
      Visible         =   0   'False
      Width           =   5505
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8895
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   10
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
   Begin VB.PictureBox pbcMMTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   5
      Top             =   1125
      Width           =   45
   End
   Begin VB.PictureBox pbcMMSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   120
      ScaleHeight     =   45
      ScaleWidth      =   0
      TabIndex        =   1
      Top             =   330
      Width           =   0
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   990
      MaxLength       =   10
      TabIndex        =   3
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
      Picture         =   "CGameInv.ctx":129C6
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   6
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
      TabIndex        =   9
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
      Left            =   4935
      TabIndex        =   8
      Top             =   5460
      Visible         =   0   'False
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
      Left            =   3360
      TabIndex        =   7
      Top             =   5460
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMultiMedia 
      Height          =   3465
      Left            =   165
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1770
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   6112
      _Version        =   393216
      Rows            =   31
      Cols            =   29
      FixedRows       =   5
      FixedCols       =   2
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
      _Band(0).Cols   =   29
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSelect 
      Height          =   1620
      Left            =   180
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   60
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   2858
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      Appearance      =   0
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lacTotals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   240
      Left            =   9000
      TabIndex        =   12
      Top             =   5460
      Visible         =   0   'False
      Width           =   75
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
      Visible         =   0   'False
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
Attribute VB_Name = "CGameInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CGameInv.ctl on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imSelectedIndex               imComboBoxIndex               smFeedSource              *
'*  smOversell                    imMinGameNo                   imMaxGameNo               *
'*  imTypeRowNo                   imItfCode                     imIifCode                 *
'*  tmGhfSrchKey0                 tmGsfSrchKey0                 tmIhfSrchKey1             *
'*  tmIsfSrchKey0                 tmIsfSrchKey1                 tmIsfSrchKey2             *
'*  tmItfSrchKey0                 tmMsfSrchKey0                 tmMgfSrchKey0             *
'*  tmMnfSrchKey                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CGameInv.ctl
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text

Private lmOpenPreviouslyCompleted As Long

Public Event SetSave(ilStatus As Integer)
Public MultiMediaVefCode As Integer
Public MultiMediaTypeItem As String
Public MultiMediaSeasonGhfCode As Long

'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imLbcArrowSetting As Integer
Dim imPopReqd As Integer
Dim imBypassFocus As Integer
Dim imDoubleClickName As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imUpdateAllowed As Integer
Dim imStartMode As Integer
Dim imVefCode As Integer
Dim imVpfIndex As Integer
Dim lmSeasonGhfCode As Long
Dim imBypassSetting As Integer
Dim imAvailColorLevel As Integer    'set in mInit as 90%
Dim imInTab As Integer
Dim smNowDate As String
Dim lmNowDate As Long
Dim lmLLD As Long
Dim lmFirstAllowedChgDate As Long

Dim tmMediaInvInfo() As MEDIAINVINFO

Dim tmIhfSort() As SORTCODE

Dim tmGameVehicle() As SORTCODE
Dim smGameVehicleTag As String

Dim tmItf() As ITF
Dim smITFTag As String

Dim imMsfChg As Integer
Dim imNewInv As Integer

Dim imCtrlVisible As Integer
Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imSetCtrlVisible As Integer
Dim imGetCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer

'Dim tmMsf() As MSFLIST
'Dim tmMgf() As MGFLIST

Dim imIhfCode As Integer

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
Dim tmIsfSrchKey3 As ISFKEY3    'ISF key record image
Dim imIsfRecLen As Integer        'ISF record length

Dim hmItf As Integer
Dim tmMFItf As ITF
Dim imItfRecLen As Integer        'ITF record length
Dim smIgnoreMultiFeed As String

Dim hmIif As Integer
Dim tmIif As IIF        'IIF record image
Dim tmIifSrchKey0 As INTKEY0    'IIF key record image
Dim imIifRecLen As Integer        'IIF record length

Dim hmMsf As Integer
Dim tmMsf As MSF        'MSF record image
Dim tmMsfSrchKey1 As MSFKEY1    'MSF key record image
Dim imMsfRecLen As Integer        'MSF record length

Dim hmMgf As Integer
Dim tmMgf As MGF        'MGF record image
Dim tmMgfSrchKey1 As MGFKEY1    'MGF key record image
Dim imMgfRecLen As Integer        'MGF record length

Dim hmCHF As Integer
Dim tmChf As CHF        'CHF record image
Dim tmChfSrchKey0 As LONGKEY0    'CHF key record image
Dim imCHFRecLen As Integer        'CHF record length

Dim tmMnf As MNF        'Mnf record image
Dim hmMnf As Integer    'Multi-Name file handle
Dim imMnfRecLen As Integer        'MNF record length
Dim tmNTRMNF() As MNF
Dim smMnfStamp As String
Dim imNTRMnfCode As Integer
Dim imNTRSlspComm As Integer

Private imLastSelectColSorted As Integer
Private imLastSelectSort As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long

Const VEHICLEINDEX = 0
Const SEASONINDEX = 1
Const GAMEDOLLARSINDEX = 2
Const SELECTEDINDEX = 3
Const SELECTSORTINDEX = 4
Const SEASONGHFCODEINDEX = 5
Const VEFCODEINDEX = 6

Const TYPEINDEX = 2   '1
Const ITEMINDEX = 4   '2
Const INVINDEX = 6 '3
Const AVAILSORDERINDEX = 8
Const AVAILSPROPOSALINDEX = 10 '7
Const UNITSINDEX = 12 '8
Const NOGAMESINDEX = 14
Const AVGCOSTINDEX = 16
Const AVGRATEINDEX = 18
Const TOTALRATEINDEX = 20
Const COMMENTINDEX = 22
Const MSFCODEINDEX = 24
Const IHFCODEINDEX = 25
Const STATUSINDEX = 26
Const INDEPENDENTINDEX = 27
Const SORTINDEX = 28



'8/5
'Private Sub cbcGameVeh_Change()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilLoop                        ilIndex                                                 *
''*                                                                                        *
''* Local Labels (Marked)                                                                  *
''*  cbcGameVehErr                                                                         *
''******************************************************************************************
'
'    Dim ilRet As Integer    'Return status
'    Dim slStr As String     'Text entered
'    Dim slCode As String
'
'    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
'        Exit Sub
'    End If
'    imChgMode = True    'Set change mode to avoid infinite loop
'    imBypassSetting = True
'    Screen.MousePointer = vbHourglass  'Wait
'    gSetMousePointer grdMultiMedia, grdSelect, vbHourglass
'    ilRet = gOptionLookAhead(cbcGameVeh, imBSMode, slStr)
'    mClearCtrlFields
'    If ilRet = 0 Then
'        slStr = tmGameVehicle(cbcGameVeh.ListIndex).sKey
'        ilRet = gParseItem(slStr, 2, "\", slCode)
'        imVefCode = Val(slCode)
'        imVpfIndex = gBinarySearchVpf(imVefCode)    'gVpfFind(CGameInv, imVefCode)
'        gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLLD
'        mBuildSoldInv True
'        ilRet = mGhfGsfReadRec()
'        mInvTypePop
'        mPopulate
'        MultiMediaVefCode = imVefCode
'        If cbcSelect.ListCount = 1 Then
'            imChgMode = False
'            cbcSelect.ListIndex = 0
'        End If
'    End If
'    Screen.MousePointer = vbDefault
'    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
'    imChgMode = False
'    imBypassSetting = False
'    Exit Sub
'cbcGameVehErr: 'VBC NR
'    On Error GoTo 0
'    Screen.MousePointer = vbDefault
'    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
'    imTerminate = True
'    imChgMode = False
'    Exit Sub
'End Sub
'
'Private Sub cbcGameVeh_Click()
'    cbcGameVeh_Change    'Process change as change event is not generated by VB
'End Sub
'
'Private Sub cbcGameVeh_GotFocus()
'    If imTerminate Then
'        Exit Sub
'    End If
'    mSetShow
'    gCtrlGotFocus cbcGameVeh
'End Sub

'Private Sub cbcSelect_Change()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilLoop                        ilIndex                                                 *
''*                                                                                        *
''* Local Labels (Marked)                                                                  *
''*  cbcSelectErr                                                                          *
''******************************************************************************************
'
'    Dim ilRet As Integer    'Return status
'    Dim slStr As String     'Text entered
'    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
'        Exit Sub
'    End If
'    imChgMode = True    'Set change mode to avoid infinite loop
'    imBypassSetting = True
'    Screen.MousePointer = vbHourglass  'Wait
'    gSetMousePointer grdMultiMedia, grdSelect, vbHourglass
'    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
'    mClearCtrlFields
'    If ilRet = 0 Then
'        If cbcSelect.ItemData(cbcSelect.ListIndex) = 0 Then
'            mTypeItemPop
'        Else
'            'Raed in Package
'        End If
'        MultiMediaTypeItem = cbcSelect.List(cbcSelect.ListIndex)
'    End If
'    Screen.MousePointer = vbDefault
'    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
'    imChgMode = False
'    imBypassSetting = False
'    Exit Sub
'cbcSelectErr: 'VBC NR
'    On Error GoTo 0
'    Screen.MousePointer = vbDefault
'    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
'    imTerminate = True
'    imChgMode = False
'    Exit Sub
'End Sub
'
'Private Sub cbcSelect_Click()
'    cbcSelect_Change    'Process change as change event is not generated by VB
'End Sub
'
'Private Sub cbcSelect_DropDown()
'    'mPopulate
'    'If imTerminate Then
'    '    Exit Sub
'    'End If
'End Sub
'
'Private Sub cbcSelect_GotFocus()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilLoop                                                                                *
''******************************************************************************************
'
'    Dim slSvText As String   'Save so list box can be reset
'    If imTerminate Then
'        Exit Sub
'    End If
'    mSetShow
'    slSvText = cbcSelect.Text
''    ilSvIndex = cbcSelect.ListIndex
'    'mPopulate
'    'If imTerminate Then
'    '    Exit Sub
'    'End If
'    If cbcSelect.ListCount <= 1 Then
'        If cbcSelect.ListCount > 0 Then
'            cbcSelect.ListIndex = 0
'        'mClearCtrlFields
'        'pbcGetSTab.SetFocus
'        End If
'        Exit Sub
'    End If
'    gCtrlGotFocus cbcSelect
'    If (slSvText = "") Or (slSvText = "[New]") Then
'        cbcSelect.ListIndex = 0
'        cbcSelect_Change    'Call change so picture area repainted
'    Else
'        gFindMatch slSvText, 1, cbcSelect
'        If gLastFound(cbcSelect) > 0 Then
''            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
'            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
'                cbcSelect.ListIndex = gLastFound(cbcSelect)
'                cbcSelect_Change    'Call change so picture area repainted
'                imPopReqd = False
'            End If
'        Else
'            cbcSelect.ListIndex = 0
'            mClearCtrlFields
'            cbcSelect_Change    'Call change so picture area repainted
'        End If
'    End If
'End Sub
'
'Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
'    'Delete key causes the charact to the right of the cursor to be deleted
'    imBSMode = False
'End Sub
'
'Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
'    'Backspace character cause selected test to be deleted or
'    'the first character to the lEtf of the cursor if no text selected
'    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
'        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
'            imBSMode = True 'Force deletion of character prior to selected text
'        End If
'    End If
'End Sub

Private Sub cmcCancel_Click()
    igGameInvReturn = False
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************


    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    mSaveRec
    igGameInvReturn = True
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow
    pbcArrow.Visible = False
    gCtrlGotFocus cmcDone
End Sub






Private Sub edcComment_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcComment_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    ilKey = KeyAscii
    If Not gCheckKeyAscii(ilKey) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcDropDown_Change()
    grdMultiMedia.CellForeColor = vbBlack
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
        Case UNITSINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "9999") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case AVGRATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            ilPos = InStr(edcDropDown.SelText, ".")
            If ilPos = 0 Then
                ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
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
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "9999999.99") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case TOTALRATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            ilPos = InStr(edcDropDown.SelText, ".")
            If ilPos = 0 Then
                ilPos = InStr(edcDropDown.Text, ".")    'Disallow multi-decimal points
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
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "9999999.99") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case COMMENTINDEX
    End Select
End Sub







Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Exit Sub
    End If
    imFirstActivate = False
    imUpdateAllowed = igUpdateAllowed
    'If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
    If Not imUpdateAllowed Then
        grdMultiMedia.Enabled = False
    Else
        grdMultiMedia.Enabled = True
    End If
    gShowBranner imUpdateAllowed
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
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


Private Sub grdMultiMedia_EnterCell()
    mSetShow
End Sub

Private Sub grdMultiMedia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim slStr As String

    'Determine if in header
'    If y < grdMultiMedia.RowHeight(0) Then
'        mSortCol grdMultiMedia.Col
'        Exit Sub
'    End If
    'Determine row and col mouse up onto
    On Error GoTo grdMultiMediaErr
    pbcArrow.Visible = False
    ilCol = grdMultiMedia.MouseCol
    ilRow = grdMultiMedia.MouseRow
    If ilCol < grdMultiMedia.FixedCols Then
        grdMultiMedia.Redraw = True
        Exit Sub
    End If
    If ilRow < grdMultiMedia.FixedRows Then
        grdMultiMedia.Redraw = True
        Exit Sub
    End If
    If ilRow Mod 2 = 0 Then
        ilRow = ilRow + 1
    End If
    If grdMultiMedia.ColWidth(ilCol) <= 15 Then
        grdMultiMedia.Redraw = True
        Exit Sub
    End If
    If grdMultiMedia.rowHeight(ilRow) <= 15 Then
        grdMultiMedia.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdMultiMedia.TopRow
    DoEvents
    If grdMultiMedia.TextMatrix(ilRow, TYPEINDEX) = "" Then
        grdMultiMedia.Redraw = True
        Exit Sub
    End If
    If (grdMultiMedia.TextMatrix(ilRow, UNITSINDEX) = "") And (ilCol <> UNITSINDEX) Then
        grdMultiMedia.Redraw = True
        Exit Sub
    End If
    grdMultiMedia.Col = ilCol
    grdMultiMedia.Row = ilRow
    If Not mColOk() Then
        grdMultiMedia.Redraw = True
        Exit Sub
    End If
    grdMultiMedia.Redraw = True
    If ilCol = NOGAMESINDEX Then
        lmEnableRow = ilRow
        lmEnableCol = ilCol
        mInitGetGames lmEnableRow, False
        igGetGameVefCode = imVefCode
        igGetGameGhfCode = lmSeasonGhfCode
        igGetGameIhfCode = Val(grdMultiMedia.TextMatrix(lmEnableRow, IHFCODEINDEX))
        GetGames.Show vbModal
        mSetNoGames lmEnableRow
        lmEnableRow = -1
        lmEnableCol = -1
    ElseIf ilCol = UNITSINDEX Then
        slStr = grdMultiMedia.TextMatrix(ilRow, ilCol)
        If InStr(1, slStr, ".", vbTextCompare) > 0 Then
            lmEnableRow = ilRow
            lmEnableCol = ilCol
            mInitGetGames lmEnableRow, False
            igGetGameVefCode = imVefCode
            igGetGameGhfCode = lmSeasonGhfCode
            igGetGameIhfCode = Val(grdMultiMedia.TextMatrix(lmEnableRow, IHFCODEINDEX))
            GetGames.Show vbModal
            mSetNoGames lmEnableRow
            lmEnableRow = -1
            lmEnableCol = -1
        Else
            mEnableBox
        End If
    Else
        mEnableBox
    End If
    On Error GoTo 0
    Exit Sub
grdMultiMediaErr:
    On Error GoTo 0
    If (lmEnableRow >= grdMultiMedia.FixedRows) And (lmEnableRow < grdMultiMedia.Rows) Then
        grdMultiMedia.Row = lmEnableRow
        grdMultiMedia.Col = lmEnableCol
        mSetFocus
    End If
    grdMultiMedia.Redraw = False
    grdMultiMedia.Redraw = True
    Exit Sub
End Sub

Private Sub grdMultiMedia_Scroll()
    mSetShow
    pbcArrow.Visible = False
    If grdMultiMedia.RowIsVisible(grdMultiMedia.Row) Then
        pbcArrow.Move grdMultiMedia.Left - pbcArrow.Width - 30, grdMultiMedia.Top + grdMultiMedia.RowPos(grdMultiMedia.Row) + (grdMultiMedia.rowHeight(grdMultiMedia.Row) - pbcArrow.height) / 2
        pbcArrow.Visible = True
    End If
End Sub

Private Sub grdSelect_EnterCell()
    mSetShow
End Sub

Private Sub grdSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String
    Dim llSelectedRow As Long

    If Y < grdSelect.rowHeight(0) Then
        grdSelect.Col = grdSelect.MouseCol
        mSelectSortCol grdSelect.Col
        grdSelect.Row = 0
        grdSelect.Col = VEFCODEINDEX
        Exit Sub
    End If
    llSelectedRow = -1
    ilFound = gGrid_GetRowCol(grdSelect, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdSelect.FixedRows Then
        Exit Sub
    End If
    For llRow = grdSelect.FixedRows To grdSelect.Rows - 1 Step 1
        If grdSelect.TextMatrix(llRow, VEHICLEINDEX) <> "" Then
            grdSelect.TextMatrix(llRow, SELECTEDINDEX) = "N"
            If llRow = llCurrentRow Then
                grdSelect.TextMatrix(llRow, SELECTEDINDEX) = "Y"
                llSelectedRow = llRow
            End If
        End If
        mPaintSelect llRow
    Next llRow
    grdSelect.Row = llCurrentRow
    If llSelectedRow <> -1 Then
        mGameVehicleSelected llSelectedRow
    End If
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInitErr                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slCode As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llRow As Long

    Screen.MousePointer = vbHourglass
    gSetMousePointer grdMultiMedia, grdSelect, vbHourglass
    imFirstActivate = True
    imTerminate = False
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.height = 165
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
    imMsfChg = False
    imNewInv = True
    imInTab = False
    smIgnoreMultiFeed = "N"
    lmEnableRow = -1
    imSetCtrlVisible = False
    imAvailColorLevel = 90
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmFirstAllowedChgDate = lmNowDate + 1
    lmLastClickedRow = -1
    lmScrollTop = grdSelect.FixedRows
    imLastSelectColSorted = -1
    imLastSelectSort = -1
    mInitBox
    smITFTag = ""
    ReDim tmItf(0 To 0) As ITF
    ReDim tmMediaInvInfo(0 To 0) As MEDIAINVINFO
    
    If lmOpenPreviouslyCompleted <> 123456789 Then
        hmGhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
        hmGsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
        hmIhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmIhf, "", sgDBPath & "Ihf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
        hmIsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmIsf, "", sgDBPath & "Isf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
        hmItf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmItf, "", sgDBPath & "Itf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
        hmIif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmIif, "", sgDBPath & "Iif.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
    
        hmMsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMsf, "", sgDBPath & "Msf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
        hmMgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMgf, "", sgDBPath & "Mgf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
        hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
    
    
        hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen)", CGameInv
        'On Error GoTo 0
        
        lmOpenPreviouslyCompleted = 123456789
    End If
    imGhfRecLen = Len(tmGhf)  'Get and save ARF record length
    ReDim tmGsf(0 To 0) As GSF
    imGsfRecLen = Len(tmGsf(0))  'Get and save ARF record length
    imIhfRecLen = Len(tmIhf)  'Get and save ARF record length
    ReDim tmIsf(0 To 0) As ISF
    imIsfRecLen = Len(tmIsf(0))  'Get and save ARF record length
    imItfRecLen = Len(tmMFItf)  'Get and save ARF record length
    imIifRecLen = Len(tmIif)  'Get and save ARF record length
    imMsfRecLen = Len(tmMsf)  'Get and save ARF record length
    imMgfRecLen = Len(tmMgf)  'Get and save ARF record length
    imCHFRecLen = Len(tmChf)  'Get and save ARF record length
    imMnfRecLen = Len(tmMnf)  'Get and save ARF record length

    imVefCode = igGameSchdVefCode
    mVehPop
    If imVefCode > 0 Then
        For llRow = grdSelect.FixedRows To grdSelect.Rows - 1 Step 1
            slCode = grdSelect.TextMatrix(llRow, VEFCODEINDEX)
            If Val(slCode) = imVefCode Then
                grdSelect.TextMatrix(llRow, SELECTEDINDEX) = "Y"
                mPaintSelect llRow
                mGameVehicleSelected llRow
                Exit For
            End If
        Next llRow
    End If
    imNTRMnfCode = mAddMultiMediaNTR()

    'ReDim tmMsf(0 To UBound(tgMsfCntr)) As MSFLIST
    'For ilLoop = 0 To UBound(tgMsfCntr) - 1 Step 1
    '    LSet tmMsf(ilLoop) = tgMsfCntr(ilLoop)
    'Next ilLoop

    'ReDim tmMgf(0 To UBound(tgMgfCntr)) As MGFLIST
    'For ilLoop = 0 To UBound(tgMgfCntr) - 1 Step 1
    '    LSet tmMgf(ilLoop) = tgMgfCntr(ilLoop)
    'Next ilLoop

    'mBuildSoldInv
    'mTeamPop
    'mClearCtrlFields
    'ilRet = mGhfGsfReadRec()
    'mLanguagePop
    'mInvTypePop
    'mPopulate
    'CGameInv.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    'gCenterStdAlone CGameInv
    mSetTotal
    Screen.MousePointer = vbDefault
    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
    Exit Sub
mInitErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
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
'*  flTextHeight                  ilLoop                        ilCol                     *
'*                                                                                        *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    Dim llRow As Long
    'flTextHeight = pbcDates.TextHeight("1") - 35
'8/5
    'cbcGameVeh.Move 2160, 45
    'cbcSelect.Move 5580, 45
    grdSelect.rowHeight(0) = fgBoxGridH + 15
    'grdSelect.Move 180, 60, Width / 3, 6 * grdSelect.RowHeight(0)
    If Screen.height <= 9000 Then
        grdSelect.Move 180, 60, Width / 3, 4 * grdSelect.rowHeight(0)
    Else
        grdSelect.Move 180, 60, Width / 3, 6 * grdSelect.rowHeight(0)
    End If
    mGridSelectColumns
    mGridSelectTitles
    gGrid_IntegralHeight grdSelect, fgBoxGridH + 15
    gGrid_FillWithRows grdSelect, fgBoxGridH + 15
    For llRow = grdSelect.FixedRows To grdSelect.Rows - 1 Step 1
        grdSelect.rowHeight(llRow) = fgBoxGridH + 15
    Next llRow
    mClearSelectGrid
    lmLastClickedRow = -1
    lmScrollTop = grdSelect.FixedRows
    imLastSelectColSorted = -1
    imLastSelectSort = -1
    
    grdMultiMedia.Move 180, grdSelect.Top + grdSelect.height + 120, Width - pbcArrow.Width - 120
    grdMultiMedia.height = height - grdSelect.height - grdSelect.Top - 240 - lacTotals.height
    grdMultiMedia.Redraw = False
    grdMultiMedia.rowHeight(0) = 2 * fgBoxGridH + 15
    grdMultiMedia.Rows = grdMultiMedia.FixedRows + 2
    llRow = grdMultiMedia.FixedRows
    Do
        If llRow + 1 > grdMultiMedia.Rows Then
            grdMultiMedia.AddItem ""
            grdMultiMedia.rowHeight(grdMultiMedia.Rows - 1) = fgBoxGridH
            grdMultiMedia.AddItem ""
            grdMultiMedia.rowHeight(grdMultiMedia.Rows - 1) = 15
            mInitNew llRow
        End If
        llRow = llRow + 2
    Loop While grdMultiMedia.RowIsVisible(llRow - 2)
    imInitNoRows = grdMultiMedia.Rows
    mGridMultiMediaLayout
    mGridMultiMediaColumnWidths
    mGridMultiMediaColumns
    mClearCtrlFields
    pbcComingSoon.Left = Width / 2 - pbcComingSoon.Width / 2
    'pbcComingSoon.Top = Height / 3 - pbcComingSoon.Height / 2
    pbcComingSoon.Top = grdMultiMedia.Top + pbcComingSoon.height / 2
    'gGrid_IntegralHeight grdMultiMedia, CInt(fgBoxGridH) + 15
    'grdMultiMedia.Height = grdMultiMedia.Height + 45
    lacTotals.Move grdMultiMedia.Left + grdMultiMedia.Width - lacTotals.Width, grdMultiMedia.Top + grdMultiMedia.height + 60
    pbcMMSTab.Left = -200
    pbcMMTab.Left = -200
    edcMultiMediaMsg.Left = Width / 2 - edcMultiMediaMsg.Width / 2
    edcMultiMediaMsg.Top = grdMultiMedia.Top + edcMultiMediaMsg.height / 2
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
    Dim ilRet As Integer

    Erase tmGsf
    Erase tmIsf

    Erase tmMediaInvInfo

    Erase tmGameVehicle
    Erase tmIhfSort

    smITFTag = ""
    Erase tmItf

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

    ilRet = btrClose(hmMsf)
    btrDestroy hmMsf

    ilRet = btrClose(hmMgf)
    btrDestroy hmMgf

    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF

    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf

    Screen.MousePointer = vbDefault
    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
    igManUnload = YES
    'Unload CGameInv
    'Set CGameInv = Nothing   'Remove data segment
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcMMSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcMMSTab.HWnd Then
        Exit Sub
    End If
    If imSetCtrlVisible Then
        Do
            ilNext = False
            Select Case grdMultiMedia.Col
                Case UNITSINDEX
                    mSetShow
                    Exit Sub
                Case AVGRATEINDEX
                    grdMultiMedia.Col = UNITSINDEX
                Case Else
                    grdMultiMedia.Col = grdMultiMedia.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdMultiMedia.Row = grdMultiMedia.FixedRows
        grdMultiMedia.Col = grdMultiMedia.FixedCols
        Do
            If mColOk() Then
                Exit Do
            Else
                '1/20/10:  Add test to verify limit not exceeded
                'grdMultiMedia.Col = grdMultiMedia.Col + 1
                If grdMultiMedia.Col + 1 <= COMMENTINDEX Then
                    grdMultiMedia.Col = grdMultiMedia.Col + 1
                Else
                    Exit Sub
                End If
            End If
        Loop
    End If
    mEnableBox
End Sub

Private Sub pbcMMTab_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llSpecEnableRow               llSpecEnableCol               llMsfIndex                *
'*                                                                                        *
'******************************************************************************************

    Dim ilNext As Integer
    Dim llRow As Long
    Dim llEnableRow As Long
    Dim llEnableCol As Long

    If GetFocus() <> pbcMMTab.HWnd Then
        Exit Sub
    End If
    If imInTab Then
        Exit Sub
    End If
    imInTab = True
    If imSetCtrlVisible Then
        Do
            ilNext = False
            Select Case grdMultiMedia.Col
                Case UNITSINDEX
                    'Test if first row, if so call the game form.  If not,
                    llEnableRow = lmEnableRow
                    llEnableCol = lmEnableCol
                    mSetShow
                    lmEnableRow = llEnableRow
                    lmEnableCol = llEnableCol
                    If grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        If grdMultiMedia.TextMatrix(grdMultiMedia.Row, INDEPENDENTINDEX) = "N" Then
                            If (grdMultiMedia.TextMatrix(grdMultiMedia.Row, NOGAMESINDEX) = "") Then
                                llRow = grdMultiMedia.Row - 2
                                Do While llRow >= grdMultiMedia.FixedRows
                                    If (grdMultiMedia.TextMatrix(llRow, NOGAMESINDEX) <> "") And (grdMultiMedia.TextMatrix(llRow, INDEPENDENTINDEX) = "N") Then
                                        grdMultiMedia.TextMatrix(grdMultiMedia.Row, NOGAMESINDEX) = grdMultiMedia.TextMatrix(llRow, NOGAMESINDEX)
                                        mInitGetGames llRow, True
                                        mSetRatesForGames lmEnableRow
                                        mSetNoGames lmEnableRow
                                        llEnableRow = lmEnableRow
                                        llEnableCol = lmEnableCol
                                        mSetShow
                                        lmEnableRow = llEnableRow
                                        lmEnableCol = llEnableCol
                                        Exit Do
                                    End If
                                    llRow = llRow - 2
                                Loop
                                If grdMultiMedia.TextMatrix(lmEnableRow, NOGAMESINDEX) = "" Then
                                    igGetGameVefCode = imVefCode
                                    igGetGameGhfCode = lmSeasonGhfCode
                                    igGetGameDefaultUnits = Val(edcDropDown.Text)
                                    igGetGameIhfCode = Val(grdMultiMedia.TextMatrix(lmEnableRow, IHFCODEINDEX))
                                    ReDim tgGetGameReturn(0 To 0) As GETGAMERETURN
                                    GetGames.Show vbModal
                                    mSetNoGames lmEnableRow
                                    If grdMultiMedia.TextMatrix(lmEnableRow, UNITSINDEX) = "" Then
                                        imInTab = False
                                        lmEnableRow = -1
                                        lmEnableCol = -1
                                        pbcClickFocus.SetFocus
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        DoEvents
                        grdMultiMedia.Col = AVGRATEINDEX
                        mEnableBox
                        imInTab = False
                        Exit Sub
                    Else
                        If grdMultiMedia.TextMatrix(grdMultiMedia.Row + 2, TYPEINDEX) = "" Then
                            imInTab = False
                            Exit Sub
                        End If
                        grdMultiMedia.Row = grdMultiMedia.Row + 2
                        grdMultiMedia.Col = UNITSINDEX
                        If Not grdMultiMedia.RowIsVisible(grdMultiMedia.Row) Then
                            Do
                                If Not grdMultiMedia.RowIsVisible(grdMultiMedia.Row) Then
                                    grdMultiMedia.TopRow = grdMultiMedia.TopRow + 1
                                Else
                                    Exit Do
                                End If
                            Loop
                        End If
                    End If
                Case COMMENTINDEX
                    If grdMultiMedia.Row + 2 >= grdMultiMedia.Rows Then
                        mSetShow
                        imInTab = False
                        Exit Sub
                    End If
                    If grdMultiMedia.TextMatrix(grdMultiMedia.Row + 2, TYPEINDEX) = "" Then
                        mSetShow
                        imInTab = False
                        Exit Sub
                    End If
                    grdMultiMedia.Row = grdMultiMedia.Row + 2
                    grdMultiMedia.Col = UNITSINDEX
                    If Not grdMultiMedia.RowIsVisible(grdMultiMedia.Row) Then
                        Do
                            If Not grdMultiMedia.RowIsVisible(grdMultiMedia.Row) Then
                                grdMultiMedia.TopRow = grdMultiMedia.TopRow + 1
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                Case Else
                    grdMultiMedia.Col = grdMultiMedia.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdMultiMedia.Row = grdMultiMedia.FixedRows
        grdMultiMedia.Col = grdMultiMedia.FixedCols
        Do
            If mColOk() Then
                Exit Do
            Else
                '1/20/10:  Add test to verify limit not exceeded
                'grdMultiMedia.Col = grdMultiMedia.Col + 1
                If grdMultiMedia.Col + 1 <= COMMENTINDEX Then
                    grdMultiMedia.Col = grdMultiMedia.Col + 1
                Else
                    Exit Sub
                End If
            End If
        Loop
    End If
    mEnableBox
    imInTab = False
End Sub





Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub








'*******************************************************
'*                                                     *
'*      Procedure Name:mGhfGsfReadRec                  *
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
    Dim ilCol As Integer
    Dim llRow As Long


    'ReDim tmGsf(0 To 0) As GSF
    ReDim tmIsf(0 To 0) As ISF
    
    lmEnableRow = -1
    grdMultiMedia.Redraw = False
    If grdMultiMedia.Rows > imInitNoRows Then
        For llRow = grdMultiMedia.Rows - 1 To imInitNoRows Step -1
            grdMultiMedia.RemoveItem llRow
        Next llRow
    End If
    For llRow = grdMultiMedia.FixedRows To grdMultiMedia.Rows - 1 Step 2
        grdMultiMedia.Row = llRow
        For ilCol = 0 To grdMultiMedia.Cols - 1 Step 1
            grdMultiMedia.TextMatrix(llRow, ilCol) = ""
            If ilCol = NOGAMESINDEX Then
                grdMultiMedia.Col = ilCol
                grdMultiMedia.CellBackColor = vbWhite
            End If
        Next ilCol
    Next llRow
    grdMultiMedia.Redraw = True

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
    If imMsfChg Then
        'cbcGameVeh.Enabled = False
        'cbcSelect.Enabled = False
        RaiseEvent SetSave(True)
    Else
        'cbcGameVeh.Enabled = True
        'cbcSelect.Enabled = True
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
Private Sub mEnableBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilLang                        slNameCode                *
'*  slCode                        ilCode                        ilLoop                    *
'*                                                                                        *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilRet As Integer
    Dim llMsfIndex As Long
    Dim llMgf As Long
    Dim tlIsf As ISF

    If (grdMultiMedia.Row < grdMultiMedia.FixedRows) Or (grdMultiMedia.Row >= grdMultiMedia.Rows) Or (grdMultiMedia.Col < grdMultiMedia.FixedCols) Or (grdMultiMedia.Col >= grdMultiMedia.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdMultiMedia.Row
    lmEnableCol = grdMultiMedia.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdMultiMedia.Left - pbcArrow.Width - 30, grdMultiMedia.Top + grdMultiMedia.RowPos(grdMultiMedia.Row) + (grdMultiMedia.rowHeight(grdMultiMedia.Row) - pbcArrow.height) / 2
    pbcArrow.Visible = True

    Select Case grdMultiMedia.Col
        Case UNITSINDEX
            edcDropDown.MaxLength = 4
            edcDropDown.Text = grdMultiMedia.Text
        Case AVGRATEINDEX
            edcDropDown.MaxLength = 10
            If (grdMultiMedia.TextMatrix(lmEnableRow, INDEPENDENTINDEX) = "Y") And (grdMultiMedia.Text = "") Then
                imIhfCode = Val(grdMultiMedia.TextMatrix(lmEnableRow, IHFCODEINDEX))
                tmIsfSrchKey3.iIhfCode = imIhfCode
                tmIsfSrchKey3.iGameNo = 0
                ilRet = btrGetGreaterOrEqual(hmIsf, tlIsf, imIsfRecLen, tmIsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (imIhfCode = tlIsf.iIhfCode) Then
                    edcDropDown.Text = gLongToStrDec(tlIsf.lRate, 2)
                    grdMultiMedia.TextMatrix(lmEnableRow, AVGCOSTINDEX) = gLongToStrDec(tlIsf.lCost, 2)
                    llMsfIndex = mAddMsfIfRequired(lmEnableRow)
                    If llMsfIndex <> -1 Then
                        If tgMsfCntr(llMsfIndex).iFirstMgf = -1 Then
                            ReDim Preserve tgMgfCntr(0 To UBound(tgMgfCntr) + 1) As MGFLIST
                            llMgf = UBound(tgMgfCntr) - 1
                            tgMsfCntr(llMsfIndex).iFirstMgf = llMgf
                            tgMgfCntr(llMgf).iStatus = 0
                            tgMgfCntr(llMgf).iNextMgf = -1
                            tgMgfCntr(llMgf).MgfRec.lCode = 0
                            tgMgfCntr(llMgf).MgfRec.iGameNo = 0
                            tgMgfCntr(llMgf).MgfRec.iNoUnits = grdMultiMedia.TextMatrix(lmEnableRow, UNITSINDEX)
                            tgMgfCntr(llMgf).MgfRec.lRate = tlIsf.lRate
                            tgMgfCntr(llMgf).MgfRec.lCost = tlIsf.lCost
                            tgMgfCntr(llMgf).MgfRec.lIsfCode = tlIsf.lCode
                        End If
                    End If
                End If
            Else
                edcDropDown.Text = grdMultiMedia.Text
            End If
        Case TOTALRATEINDEX
            edcDropDown.MaxLength = 10
            If (grdMultiMedia.TextMatrix(lmEnableRow, INDEPENDENTINDEX) = "Y") And (grdMultiMedia.Text = "") Then
                imIhfCode = Val(grdMultiMedia.TextMatrix(lmEnableRow, IHFCODEINDEX))
                tmIsfSrchKey3.iIhfCode = imIhfCode
                tmIsfSrchKey3.iGameNo = 0
                ilRet = btrGetGreaterOrEqual(hmIsf, tlIsf, imIsfRecLen, tmIsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (imIhfCode = tlIsf.iIhfCode) Then
                    edcDropDown.Text = gLongToStrDec(tlIsf.lRate, 2)
                    grdMultiMedia.TextMatrix(lmEnableRow, AVGCOSTINDEX) = gLongToStrDec(tlIsf.lCost, 2)
                    llMsfIndex = mAddMsfIfRequired(lmEnableRow)
                    If llMsfIndex <> -1 Then
                        If tgMsfCntr(llMsfIndex).iFirstMgf = -1 Then
                            ReDim Preserve tgMgfCntr(0 To UBound(tgMgfCntr) + 1) As MGFLIST
                            llMgf = UBound(tgMgfCntr) - 1
                            tgMsfCntr(llMsfIndex).iFirstMgf = llMgf
                            tgMgfCntr(llMgf).iStatus = 0
                            tgMgfCntr(llMgf).iNextMgf = -1
                            tgMgfCntr(llMgf).MgfRec.lCode = 0
                            tgMgfCntr(llMgf).MgfRec.iGameNo = 0
                            tgMgfCntr(llMgf).MgfRec.iNoUnits = grdMultiMedia.TextMatrix(lmEnableRow, UNITSINDEX)
                            tgMgfCntr(llMgf).MgfRec.lRate = tlIsf.lRate
                            tgMgfCntr(llMgf).MgfRec.lCost = tlIsf.lCost
                            tgMgfCntr(llMgf).MgfRec.lIsfCode = tlIsf.lCode
                        End If
                    End If
                End If
            Else
                edcDropDown.Text = grdMultiMedia.Text
            End If
        Case COMMENTINDEX
            edcComment.MaxLength = 0
            edcComment.Text = grdMultiMedia.Text
    End Select
    mSetFocus
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
'*  ilNoGames                     ilOrigUpper                   ilLoop                    *
'*  llRow                         llSvRow                       llSvCol                   *
'*                                                                                        *
'******************************************************************************************

    Dim slStr As String
    Dim llMsf As Long
    Dim slToStr As String
    Dim llMgf As Long
    Dim llInfo As Long
    Dim llUnitTotal As Long
    Dim llTypeUnitsOrdered As Long
    Dim llTypeUnitsProp As Long
    Dim slShort As String
    Dim slFuture As String
    Dim slBilled As String
    Dim slOldRate As String
    Dim slNewRate As String
    Dim ilCxfIndex As Integer
    Dim ilRet As Integer
    Dim llAdj As Long
    Dim llRem As Long
    Dim ilFutureNoUnits As Integer
    Dim ilNoUnits As Integer
    Dim llRate As Long
    Dim ilCheckBalance As Integer
    Dim tlIsf As ISF

    pbcArrow.Visible = False
    If (lmEnableRow >= grdMultiMedia.FixedRows) And (lmEnableRow < grdMultiMedia.Rows) Then
        Select Case lmEnableCol
            Case UNITSINDEX
                edcDropDown.Visible = False
                If grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol) <> edcDropDown.Text Then
                    llMsf = mAddMsfIfRequired(lmEnableRow)
                    igGetGameDefaultUnits = Val(edcDropDown.Text)
                    imIhfCode = Val(grdMultiMedia.TextMatrix(lmEnableRow, IHFCODEINDEX))
                    llMgf = tgMsfCntr(llMsf).iFirstMgf
                    If (llMgf = -1) And (grdMultiMedia.TextMatrix(lmEnableRow, INDEPENDENTINDEX) = "Y") Then
                        tmIsfSrchKey3.iIhfCode = imIhfCode
                        tmIsfSrchKey3.iGameNo = 0
                        ilRet = btrGetGreaterOrEqual(hmIsf, tlIsf, imIsfRecLen, tmIsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                        If (ilRet = BTRV_ERR_NONE) And (imIhfCode = tlIsf.iIhfCode) Then
                            ReDim Preserve tgMgfCntr(0 To UBound(tgMgfCntr) + 1) As MGFLIST
                            llMgf = UBound(tgMgfCntr) - 1
                            tgMsfCntr(llMsf).iFirstMgf = llMgf
                            tgMgfCntr(llMgf).iStatus = 0
                            tgMgfCntr(llMgf).iNextMgf = -1
                            tgMgfCntr(llMgf).MgfRec.lCode = 0
                            tgMgfCntr(llMgf).MgfRec.iGameNo = 0
                            tgMgfCntr(llMgf).MgfRec.iNoUnits = Val(edcDropDown.Text)
                            tgMgfCntr(llMgf).MgfRec.lRate = tlIsf.lRate
                            tgMgfCntr(llMgf).MgfRec.lCost = tlIsf.lCost
                            tgMgfCntr(llMgf).MgfRec.lIsfCode = tlIsf.lCode
                        End If
                    End If
                    Do While llMgf <> -1
                        If tgMgfCntr(llMgf).MgfRec.sBilled <> "Y" Then
                            For llInfo = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                                If (tmMediaInvInfo(llInfo).iIhfCode = imIhfCode) And (tgMgfCntr(llMgf).MgfRec.iGameNo = tmMediaInvInfo(llInfo).iGameNo) Then
                                    tmMediaInvInfo(llInfo).iNoUnitsProp = tmMediaInvInfo(llInfo).iNoUnitsProp - tgMgfCntr(llMgf).MgfRec.iNoUnits + igGetGameDefaultUnits
                                    Exit For
                                End If
                            Next llInfo
                            tgMgfCntr(llMgf).MgfRec.iNoUnits = igGetGameDefaultUnits
                        End If
                        llMgf = tgMgfCntr(llMgf).iNextMgf
                    Loop
                    llTypeUnitsOrdered = 0
                    llTypeUnitsProp = 0
                    For llInfo = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                        If tmMediaInvInfo(llInfo).iIhfCode = imIhfCode Then
                            llTypeUnitsOrdered = llTypeUnitsOrdered + tmMediaInvInfo(llInfo).iNoUnitsOrdered
                            If tmMediaInvInfo(llInfo).iNoUnitsProp > 0 Then
                                llTypeUnitsProp = llTypeUnitsProp + (tmMediaInvInfo(llInfo).iNoUnitsOrdered - tmMediaInvInfo(llInfo).iNoUnitsProp)
                            End If
                        End If
                    Next llInfo
                    llTypeUnitsProp = llTypeUnitsOrdered - llTypeUnitsProp
                    llUnitTotal = grdMultiMedia.TextMatrix(lmEnableRow, INVINDEX)
                    If llUnitTotal < llTypeUnitsProp Then
                        grdMultiMedia.CellForeColor = vbMagenta
                    ElseIf (llUnitTotal * imAvailColorLevel) \ 100 < llTypeUnitsProp Then
                        grdMultiMedia.CellForeColor = DARKYELLOW
                    Else
                        grdMultiMedia.CellForeColor = vbBlack
                    End If
                    grdMultiMedia.TextMatrix(lmEnableRow, AVAILSPROPOSALINDEX) = llUnitTotal - llTypeUnitsProp
                    imMsfChg = True
                End If
                If (grdMultiMedia.TextMatrix(lmEnableRow, NOGAMESINDEX) = "") Or (grdMultiMedia.TextMatrix(lmEnableRow, NOGAMESINDEX) = "Event-Ind") Then
                    grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol) = edcDropDown.Text
                Else
                    mSetAvgs lmEnableRow
                End If
            Case AVGRATEINDEX
                edcDropDown.Visible = False
                If gStrDecToLong(grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol), 2) <> gStrDecToLong(edcDropDown.Text, 2) Then
                    imMsfChg = True
                    'Distributed dollars to games
                    'Formula:  (NewAvgRate/OldAvgRate) + (DollarShortFromBilledGames/CurrentDollarsInUnbillGames)
                    '           where:
                    '           DollarHortFromBilledGames = (GameRate*(NewAvgRate/OldAvgRate))
                    slStr = edcDropDown.Text
                    'Add decimal point if required
                    slStr = gLongToStrDec(gStrDecToLong(slStr, 2), 2)
                    If grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        slToStr = gDivStr(gMulStr(slStr, "100"), grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol))
                    Else
                        slToStr = "100."
                    End If
                    slShort = "0.00"
                    slFuture = "0.00"
                    ilFutureNoUnits = 0
                    ilNoUnits = 0
                    llMsf = mAddMsfIfRequired(lmEnableRow)
                    llMgf = tgMsfCntr(llMsf).iFirstMgf
                    Do While llMgf <> -1
                        If tgMgfCntr(llMgf).MgfRec.sBilled = "Y" Then
                            slOldRate = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lRate, 2)
                            slNewRate = gMulStr(slToStr, slOldRate)
                            slNewRate = gRoundStr(slNewRate, "10.00", 2)
                            slStr = gSubStr(slNewRate, gMulStr(slOldRate, "100"))
                            slStr = gMulStr(slStr, Trim$(str$(tgMgfCntr(llMgf).MgfRec.iNoUnits)))
                            slShort = gAddStr(slShort, slStr)
                            ilNoUnits = ilNoUnits + tgMgfCntr(llMgf).MgfRec.iNoUnits
                        Else
                            slStr = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lRate, 2)
                            slStr = gMulStr(slStr, Trim$(str$(tgMgfCntr(llMgf).MgfRec.iNoUnits)))
                            slFuture = gAddStr(slFuture, slStr)
                            ilFutureNoUnits = ilFutureNoUnits + tgMgfCntr(llMgf).MgfRec.iNoUnits
                            ilNoUnits = ilNoUnits + tgMgfCntr(llMgf).MgfRec.iNoUnits
                        End If
                        llMgf = tgMgfCntr(llMgf).iNextMgf
                    Loop
                    If gStrDecToLong(slFuture, 2) = 0 Then
                        slToStr = "100.00"
                        If ilFutureNoUnits > 0 Then
                            slStr = edcDropDown.Text
                            slStr = gDivStr(gLongToStrDec(ilNoUnits * gStrDecToLong(slStr, 2) - gStrDecToLong(slShort, 2), 2), Trim$(str$(ilFutureNoUnits)))
                            llRate = gStrDecToLong(slStr, 2)
                            llMsf = mAddMsfIfRequired(lmEnableRow)
                            llMgf = tgMsfCntr(llMsf).iFirstMgf
                            Do While llMgf <> -1
                                If tgMgfCntr(llMgf).MgfRec.sBilled <> "Y" Then
                                    tgMgfCntr(llMgf).MgfRec.lRate = llRate
                                End If
                                llMgf = tgMgfCntr(llMgf).iNextMgf
                            Loop
                        End If
                    Else
                        slToStr = gAddStr(slToStr, gDivStr(slShort, slFuture))
                        llMsf = mAddMsfIfRequired(lmEnableRow)
                        llMgf = tgMsfCntr(llMsf).iFirstMgf
                        Do While llMgf <> -1
                            If tgMgfCntr(llMgf).MgfRec.sBilled <> "Y" Then
                                slStr = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lRate, 2)
                                slStr = gDivStr(gMulStr(slStr, slToStr), "100.00")
                                slStr = gRoundStr(slStr, ".10", 2)
                                tgMgfCntr(llMgf).MgfRec.lRate = gStrDecToLong(slStr, 2)
                            End If
                            llMgf = tgMgfCntr(llMgf).iNextMgf
                        Loop
                    End If
                End If
                'grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol) = edcDropdown.Text
                mSetAvgs lmEnableRow
            Case TOTALRATEINDEX
                edcDropDown.Visible = False
                If gStrDecToLong(grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol), 2) <> gStrDecToLong(edcDropDown.Text, 2) Then
                    imMsfChg = True
                    'Distributed dollars to games
                    slStr = edcDropDown.Text
                    slToStr = gLongToStrDec(gStrDecToLong(slStr, 2), 2)
                    slBilled = "0.00"
                    slFuture = "0.00"
                    ilFutureNoUnits = 0
                    llMsf = mAddMsfIfRequired(lmEnableRow)
                    llMgf = tgMsfCntr(llMsf).iFirstMgf
                    Do While llMgf <> -1
                        If tgMgfCntr(llMgf).MgfRec.sBilled = "Y" Then
                            slOldRate = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lRate, 2)
                            slStr = gMulStr(slOldRate, Trim$(str$(tgMgfCntr(llMgf).MgfRec.iNoUnits)))
                            slBilled = gAddStr(slBilled, slStr)
                        Else
                            slStr = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lRate, 2)
                            slStr = gMulStr(slStr, Trim$(str$(tgMgfCntr(llMgf).MgfRec.iNoUnits)))
                            slFuture = gAddStr(slFuture, slStr)
                            ilFutureNoUnits = ilFutureNoUnits + tgMgfCntr(llMgf).MgfRec.iNoUnits
                        End If
                        llMgf = tgMgfCntr(llMgf).iNextMgf
                    Loop
                    ilCheckBalance = True
                    If gStrDecToLong(slFuture, 2) = 0 Then
                        If ilFutureNoUnits > 0 Then
                            slStr = gDivStr(gSubStr(slToStr, slBilled), Trim$(str$(ilFutureNoUnits)))
                            llRate = gStrDecToLong(slStr, 2)
                            llMsf = mAddMsfIfRequired(lmEnableRow)
                            llMgf = tgMsfCntr(llMsf).iFirstMgf
                            Do While llMgf <> -1
                                If tgMgfCntr(llMgf).MgfRec.sBilled <> "Y" Then
                                    tgMgfCntr(llMgf).MgfRec.lRate = llRate
                                    slStr = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate, 2)
                                    slFuture = gAddStr(slFuture, slStr)
                                End If
                                llMgf = tgMgfCntr(llMgf).iNextMgf
                            Loop
                        Else
                            ilCheckBalance = False
                        End If
                    Else
                        slToStr = gDivStr(gMulStr(gSubStr(slToStr, slBilled), "100."), slFuture)
                        slFuture = "0.00"
                        llMsf = mAddMsfIfRequired(lmEnableRow)
                        llMgf = tgMsfCntr(llMsf).iFirstMgf
                        Do While llMgf <> -1
                            If tgMgfCntr(llMgf).MgfRec.sBilled <> "Y" Then
                                slStr = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lRate, 2)
                                slStr = gDivStr(gMulStr(slStr, slToStr), "100.00")
                                slStr = gRoundStr(slStr, ".10", 2)
                                tgMgfCntr(llMgf).MgfRec.lRate = gStrDecToLong(slStr, 2)
                                slStr = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate, 2)
                                slFuture = gAddStr(slFuture, slStr)
                            End If
                            llMgf = tgMgfCntr(llMgf).iNextMgf
                        Loop
                    End If
                    'Check balance
                    If ilCheckBalance Then
                        slStr = edcDropDown.Text
                        slToStr = gLongToStrDec(gStrDecToLong(slStr, 2), 2)
                        slStr = gAddStr(slBilled, slFuture)
                        If gStrDecToLong(slStr, 2) <> gStrDecToLong(slToStr, 2) Then
                            llAdj = gStrDecToLong(gSubStr(slToStr, slStr), 2)
                            llMsf = mAddMsfIfRequired(lmEnableRow)
                            llMgf = tgMsfCntr(llMsf).iFirstMgf
                            Do While llMgf <> -1
                                If tgMgfCntr(llMgf).MgfRec.sBilled <> "Y" Then
                                    If tgMgfCntr(llMgf).MgfRec.iNoUnits = 1 Then
                                        tgMgfCntr(llMgf).MgfRec.lRate = tgMgfCntr(llMgf).MgfRec.lRate + llAdj
                                        llAdj = 0
                                        Exit Do
                                    End If
                                End If
                                llMgf = tgMgfCntr(llMgf).iNextMgf
                            Loop
                            If llAdj <> 0 Then
                                llMsf = mAddMsfIfRequired(lmEnableRow)
                                llMgf = tgMsfCntr(llMsf).iFirstMgf
                                Do While llMgf <> -1
                                    If tgMgfCntr(llMgf).MgfRec.sBilled <> "Y" Then
                                        llRem = llAdj Mod tgMgfCntr(llMgf).MgfRec.iNoUnits
                                        If llRem = 0 Then
                                            tgMgfCntr(llMgf).MgfRec.lRate = tgMgfCntr(llMgf).MgfRec.lRate + llAdj / tgMgfCntr(llMgf).MgfRec.iNoUnits
                                            llAdj = 0
                                            Exit Do
                                        End If
                                    End If
                                    llMgf = tgMgfCntr(llMgf).iNextMgf
                                Loop
                            End If
                        End If
                    End If
                End If
                'grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol) = edcDropdown.Text
                mSetAvgs lmEnableRow
            Case COMMENTINDEX
                edcComment.Visible = False
                If grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol) <> edcComment.Text Then
                    imMsfChg = True
                End If
                grdMultiMedia.TextMatrix(lmEnableRow, lmEnableCol) = edcComment.Text
                llMsf = mAddMsfIfRequired(lmEnableRow)
                If tgMsfCntr(llMsf).iCxfIndex < 0 Then
                    ReDim Preserve sgMsfCntrCxf(0 To UBound(sgMsfCntrCxf) + 1) As String
                    tgMsfCntr(llMsf).iCxfIndex = UBound(sgMsfCntrCxf) - 1
                End If
                ilCxfIndex = tgMsfCntr(llMsf).iCxfIndex
                sgMsfCntrCxf(ilCxfIndex) = edcComment.Text
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    mSetTotal
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
    Dim llColWidth As Long

    If (grdMultiMedia.Row < grdMultiMedia.FixedRows) Or (grdMultiMedia.Row >= grdMultiMedia.Rows) Or (grdMultiMedia.Col < grdMultiMedia.FixedCols) Or (grdMultiMedia.Col >= grdMultiMedia.Cols - 1) Then
        Exit Sub
    End If
    imSetCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdMultiMedia.Col - 1 Step 1
        llColPos = llColPos + grdMultiMedia.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdMultiMedia.ColWidth(grdMultiMedia.Col)
    ilCol = grdMultiMedia.Col
    Do While ilCol < grdMultiMedia.Cols - 1
        If (Trim$(grdMultiMedia.TextMatrix(grdMultiMedia.Row - 1, grdMultiMedia.Col)) <> "") And (Trim$(grdMultiMedia.TextMatrix(grdMultiMedia.Row - 1, grdMultiMedia.Col)) = Trim$(grdMultiMedia.TextMatrix(grdMultiMedia.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdMultiMedia.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdMultiMedia.Col
        Case UNITSINDEX
            edcDropDown.Move grdMultiMedia.Left + llColPos + 30, grdMultiMedia.Top + grdMultiMedia.RowPos(grdMultiMedia.Row) + 15, grdMultiMedia.ColWidth(grdMultiMedia.Col), grdMultiMedia.rowHeight(grdMultiMedia.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case AVGRATEINDEX
            edcDropDown.Move grdMultiMedia.Left + llColPos + 30, grdMultiMedia.Top + grdMultiMedia.RowPos(grdMultiMedia.Row) + 30, grdMultiMedia.ColWidth(grdMultiMedia.Col) - 30, grdMultiMedia.rowHeight(grdMultiMedia.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TOTALRATEINDEX
            edcDropDown.Move grdMultiMedia.Left + llColPos + 30, grdMultiMedia.Top + grdMultiMedia.RowPos(grdMultiMedia.Row) + 30, grdMultiMedia.ColWidth(grdMultiMedia.Col) - 30, grdMultiMedia.rowHeight(grdMultiMedia.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case COMMENTINDEX
            edcComment.Move grdMultiMedia.Left + llColPos + grdMultiMedia.ColWidth(grdMultiMedia.Col) - edcComment.Width, grdMultiMedia.Top + grdMultiMedia.RowPos(grdMultiMedia.Row) + 30
            edcComment.Visible = True
            edcComment.SetFocus
    End Select
End Sub



Private Sub mGridMultiMediaLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    'Layout Fixed Rows:0=>Edge; 1=>Blue border; 2=>Column Title 1; 3=Column Title 2; 4=>Blue border
    '       Rows: 5=>input; 6=>blue row line; 7=>Input; 8=>blue row line
    'Layout Fixed Columns: 0=>Edge; 1=Blue border;
    '       Columns: 2=>Input; 3=>Blue column line; 4=>Input; 5=>Blue Column;....
    grdMultiMedia.rowHeight(0) = 15
    grdMultiMedia.rowHeight(1) = 15
    grdMultiMedia.rowHeight(2) = 180
    grdMultiMedia.rowHeight(3) = 180
    grdMultiMedia.rowHeight(4) = 15
    For ilRow = grdMultiMedia.FixedRows To grdMultiMedia.Rows - 1 Step 2
        grdMultiMedia.rowHeight(ilRow) = fgBoxGridH
        grdMultiMedia.Row = ilRow
        For ilCol = 0 To grdMultiMedia.Cols - 1 Step 1
            grdMultiMedia.ColAlignment(ilCol) = flexAlignLeftCenter
            If (ilCol <> UNITSINDEX) And (ilCol <> NOGAMESINDEX) And (ilCol <> AVGRATEINDEX) And (ilCol <> TOTALRATEINDEX) And (ilCol <> COMMENTINDEX) Then
                grdMultiMedia.Col = ilCol
                grdMultiMedia.CellBackColor = LIGHTYELLOW
            Else
                grdMultiMedia.Col = ilCol
                grdMultiMedia.CellBackColor = vbWhite
            End If
        Next ilCol
        grdMultiMedia.rowHeight(ilRow + 1) = 15
    Next ilRow

    'For ilCol = 0 To grdMultiMedia.Cols - 1 Step 1
    '    grdMultiMedia.ColAlignment(ilCol) = flexAlignLeftCenter
    'Next ilCol
    grdMultiMedia.ColWidth(0) = 15
    grdMultiMedia.ColWidth(1) = 15
    For ilCol = grdMultiMedia.FixedCols + 1 To grdMultiMedia.Cols - 1 Step 2
        grdMultiMedia.ColWidth(ilCol) = 15
    Next ilCol
    'Horizontal Blue Border Lines
    grdMultiMedia.Row = 1
    For ilCol = 1 To grdMultiMedia.Cols - 1 Step 1
        grdMultiMedia.Col = ilCol
        grdMultiMedia.CellBackColor = vbBlue
    Next ilCol
    grdMultiMedia.Row = 4
    For ilCol = 1 To grdMultiMedia.Cols - 1 Step 1
        grdMultiMedia.Col = ilCol
        grdMultiMedia.CellBackColor = vbBlue
    Next ilCol
    'Horizontal Blue lines
    For ilRow = grdMultiMedia.FixedRows + 1 To grdMultiMedia.Rows - 1 Step 2
        grdMultiMedia.Row = ilRow
        For ilCol = 1 To grdMultiMedia.Cols - 1 Step 1
            grdMultiMedia.Col = ilCol
            grdMultiMedia.CellBackColor = vbBlue
        Next ilCol
    Next ilRow
    'Vertical Border Lines
    grdMultiMedia.Col = 1
    For ilRow = 1 To grdMultiMedia.Rows - 1 Step 1
        grdMultiMedia.Row = ilRow
        grdMultiMedia.CellBackColor = vbBlue
    Next ilRow
    grdMultiMedia.Col = 1
    For ilRow = 1 To grdMultiMedia.Rows - 1 Step 1
        grdMultiMedia.Row = ilRow
        grdMultiMedia.CellBackColor = vbBlue
    Next ilRow
    ''Set color in fix area to white
    'grdMultiMedia.Col = 2
    'For ilRow = grdMultiMedia.FixedRows To grdMultiMedia.Rows - 1 Step 2
    '    grdMultiMedia.Row = ilRow
    '    grdMultiMedia.CellBackColor = vbWhite
    'Next ilRow

    'Vertical Blue Lines
    For ilCol = grdMultiMedia.FixedCols + 1 To grdMultiMedia.Cols - 1 Step 2
        grdMultiMedia.Col = ilCol
        For ilRow = 1 To grdMultiMedia.Rows - 1 Step 1
            grdMultiMedia.Row = ilRow
            grdMultiMedia.CellBackColor = vbBlue
        Next ilRow
    Next ilCol
End Sub



Private Sub mGridMultiMediaColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdMultiMedia.Row = 2
    grdMultiMedia.Col = TYPEINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Type"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = TYPEINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = ""
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = ITEMINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Item"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = ITEMINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = ""
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = INVINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Total"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = INVINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Inv"
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = AVAILSORDERINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Total Avails"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = AVAILSORDERINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Ordered"
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = AVAILSPROPOSALINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Total Avails"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = AVAILSPROPOSALINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Proposal"
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = UNITSINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Units/Event"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = UNITSINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "or Units"
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = NOGAMESINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "# Events"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = NOGAMESINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = ""
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = AVGCOSTINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Avg Cost"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = AVGCOSTINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = LIGHTYELLOW
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = ""
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = AVGRATEINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Avg Rate"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = AVGRATEINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = ""
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = TOTALRATEINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "Total Rate"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = TOTALRATEINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = ""
    grdMultiMedia.Row = 2
    grdMultiMedia.Col = COMMENTINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = "C"
    grdMultiMedia.Row = 3
    grdMultiMedia.Col = COMMENTINDEX
    grdMultiMedia.CellFontBold = False
    grdMultiMedia.CellFontName = "Arial"
    grdMultiMedia.CellFontSize = 6.75
    grdMultiMedia.CellForeColor = vbBlue
    grdMultiMedia.CellBackColor = vbWhite
    grdMultiMedia.TextMatrix(grdMultiMedia.Row, grdMultiMedia.Col) = ""

End Sub

Private Sub mGridMultiMediaColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdMultiMedia.ColWidth(MSFCODEINDEX) = 0
    grdMultiMedia.ColWidth(IHFCODEINDEX) = 0
    grdMultiMedia.ColWidth(STATUSINDEX) = 0
    grdMultiMedia.ColWidth(SORTINDEX) = 0
    grdMultiMedia.ColWidth(INDEPENDENTINDEX) = 0
    grdMultiMedia.ColWidth(TYPEINDEX) = 0.13 * grdMultiMedia.Width
    grdMultiMedia.ColWidth(ITEMINDEX) = 0.13 * grdMultiMedia.Width
    grdMultiMedia.ColWidth(INVINDEX) = 0.05 * grdMultiMedia.Width
    grdMultiMedia.ColWidth(AVAILSORDERINDEX) = 0.06 * grdMultiMedia.Width
    If tgSpf.sGUsePropSys = "Y" Then
        grdMultiMedia.ColWidth(AVAILSPROPOSALINDEX) = 0.06 * grdMultiMedia.Width
    Else
        grdMultiMedia.ColWidth(AVAILSPROPOSALINDEX) = 0
        grdMultiMedia.ColWidth(AVAILSPROPOSALINDEX + 1) = 0
    End If
    grdMultiMedia.ColWidth(UNITSINDEX) = 0.065 * grdMultiMedia.Width
    grdMultiMedia.ColWidth(NOGAMESINDEX) = 0.07 * grdMultiMedia.Width
    grdMultiMedia.ColWidth(AVGCOSTINDEX) = 0.07 * grdMultiMedia.Width
    grdMultiMedia.ColWidth(AVGRATEINDEX) = 0.07 * grdMultiMedia.Width
    grdMultiMedia.ColWidth(TOTALRATEINDEX) = 0.07 * grdMultiMedia.Width
    grdMultiMedia.ColWidth(COMMENTINDEX) = 0.015 * grdMultiMedia.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdMultiMedia.Width
    For ilCol = 0 To grdMultiMedia.Cols - 1 Step 1
        llWidth = llWidth + grdMultiMedia.ColWidth(ilCol)
        If (grdMultiMedia.ColWidth(ilCol) > 15) And (grdMultiMedia.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdMultiMedia.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdMultiMedia.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdMultiMedia.Width
            For ilCol = 0 To grdMultiMedia.Cols - 1 Step 1
                If (grdMultiMedia.ColWidth(ilCol) > 15) And (grdMultiMedia.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdMultiMedia.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdMultiMedia.FixedCols To grdMultiMedia.Cols - 1 Step 1
                If grdMultiMedia.ColWidth(ilCol) > 15 Then
                    ilColInc = grdMultiMedia.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdMultiMedia.ColWidth(ilCol) = grdMultiMedia.ColWidth(ilCol) + 15
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









Private Function mColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    mColOk = True
    If grdMultiMedia.ColWidth(grdMultiMedia.Col) <= 15 Then
        mColOk = False
        Exit Function
    End If
    If grdMultiMedia.rowHeight(grdMultiMedia.Row) <= 15 Then
        mColOk = False
        Exit Function
    End If
    If grdMultiMedia.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
    If grdMultiMedia.CellForeColor = vbRed Then
        mColOk = False
        Exit Function
    End If


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

'
'   mPopulate
'   Where:
'

'8/5
'    cbcSelect.Clear
'    'Get packages
'    cbcSelect.AddItem "[All]"
'    cbcSelect.ItemData(cbcSelect.NewIndex) = 0
    Exit Sub
mPopulateErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

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
Private Sub mTypeItemPop()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         slNameCode                                              *
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
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim slCode As String
    Dim llUnitTotal As Long
    Dim llTypeUnitsOrdered As Long
    Dim llTypeUnitsProp As Long
    Dim llMsf As Long
    Dim llMgf As Long
    Dim ilNoGames As Integer
    Dim ilNoUnits As Integer
    Dim llCost As Long
    Dim llRate As Long
    Dim llInfo As Long
    Dim ilIhf As Integer
    Dim slStr As String
    Dim llAvgCost As Long
    Dim llAvgRate As Long
    Dim llTotalRate As Long
    Dim ilNoAvg As Integer

    grdMultiMedia.Redraw = False
    'grdMultiMedia.Rows = grdMultiMedia.FixedRows + 1
    llRow = grdMultiMedia.FixedRows
    'grdMultiMedia.Row = llRow
    'For ilCol = 0 To grdMultiMedia.Cols - 1 Step 1
    '    If (ilCol <> UNITSINDEX) And (ilCol <> NOGAMESINDEX) And (ilCol <> AVGRATEINDEX) Then
    '        grdMultiMedia.Col = ilCol
    '        grdMultiMedia.CellBackColor = LIGHTYELLOW
    '    End If
    '    grdMultiMedia.TextMatrix(llRow, ilCol) = ""
    'Next ilCol
    ReDim tmIhfSort(0 To 0) As SORTCODE
    tmIhfSrchKey2.iVefCode = imVefCode
    ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmIhf.iVefCode = imVefCode)
        If tmIhf.lghfcode = lmSeasonGhfCode Then
            slInvType = ""
            For ilLoop = 0 To UBound(tmItf) - 1 Step 1
                If tmIhf.iItfCode = tmItf(ilLoop).iCode Then
                    slInvType = Trim$(tmItf(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
            Do While Len(slInvType) < 50
                slInvType = slInvType & " "
            Loop
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
            Do While Len(slInvItem) < 60
                slInvItem = slInvItem & " "
            Loop
            tmIhfSort(UBound(tmIhfSort)).sKey = slInvType & "\" & slInvItem & "\" & tmIhf.iCode
            ReDim Preserve tmIhfSort(0 To UBound(tmIhfSort) + 1) As SORTCODE
        End If
        ilRet = btrGetNext(hmIhf, tmIhf, imIhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If UBound(tmIhfSort) - 1 > 0 Then
        ArraySortTyp fnAV(tmIhfSort(), 0), UBound(tmIhfSort), 0, LenB(tmIhfSort(0)), 0, LenB(tmIhfSort(0).sKey), 0
    End If
    For ilIhf = 0 To UBound(tmIhfSort) - 1 Step 1
        slStr = tmIhfSort(ilIhf).sKey
        ilRet = gParseItem(slStr, 1, "\", slInvType)
        slInvType = Trim$(slInvType)
        ilRet = gParseItem(slStr, 2, "\", slInvItem)
        slInvItem = Trim$(slInvItem)
        ilRet = gParseItem(slStr, 3, "\", slCode)
        imIhfCode = Val(slCode)
        ilRet = mIhfReadRec()
        If llRow + 1 > grdMultiMedia.Rows Then
            grdMultiMedia.AddItem ""
            grdMultiMedia.rowHeight(grdMultiMedia.Rows - 1) = fgBoxGridH
            grdMultiMedia.AddItem ""
            grdMultiMedia.rowHeight(grdMultiMedia.Rows - 1) = 15
            mInitNew llRow
        End If
        llUnitTotal = 0
        llAvgCost = 0
        llAvgRate = 0
        llTotalRate = 0
        ilNoAvg = 0
        For ilLoop = 0 To UBound(tmIsf) - 1 Step 1
            If tmIhf.sGameIndependent <> "Y" Then
                If (tmIsf(ilLoop).iGameNo > 0) And (tmIsf(ilLoop).iNoUnits > 0) Then
                    ilNoAvg = ilNoAvg + 1
                    llUnitTotal = llUnitTotal + tmIsf(ilLoop).iNoUnits
                    llAvgCost = llAvgCost + tmIsf(ilLoop).lCost
                    llAvgRate = llAvgRate + tmIsf(ilLoop).lRate
                End If
            Else
                If (tmIsf(ilLoop).iGameNo = 0) And (tmIsf(ilLoop).iNoUnits > 0) Then
                    ilNoAvg = ilNoAvg + 1
                    llUnitTotal = llUnitTotal + tmIsf(ilLoop).iNoUnits
                    llAvgCost = llAvgCost + tmIsf(ilLoop).lCost
                    llAvgRate = llAvgRate + tmIsf(ilLoop).lRate
                End If
            End If
        Next ilLoop
        If (ilNoAvg > 1) Then
            llTotalRate = llAvgRate
            llAvgCost = llAvgCost / ilNoAvg
            llAvgRate = llAvgRate / ilNoAvg
        End If
        llTypeUnitsOrdered = 0
        llTypeUnitsProp = 0
        For llInfo = 0 To UBound(tmMediaInvInfo) - 1 Step 1
            If tmMediaInvInfo(llInfo).iIhfCode = imIhfCode Then
                llTypeUnitsOrdered = llTypeUnitsOrdered + tmMediaInvInfo(llInfo).iNoUnitsOrdered
                If tmMediaInvInfo(llInfo).iNoUnitsProp > 0 Then
                    llTypeUnitsProp = llTypeUnitsProp + (tmMediaInvInfo(llInfo).iNoUnitsProp - tmMediaInvInfo(llInfo).iNoUnitsOrdered)
                End If
            End If
        Next llInfo
        llTypeUnitsProp = llTypeUnitsOrdered + llTypeUnitsProp
        grdMultiMedia.Row = llRow
        grdMultiMedia.Col = TYPEINDEX
        grdMultiMedia.CellBackColor = LIGHTYELLOW
        grdMultiMedia.TextMatrix(llRow, TYPEINDEX) = slInvType
        grdMultiMedia.Col = ITEMINDEX
        grdMultiMedia.CellBackColor = LIGHTYELLOW
        grdMultiMedia.TextMatrix(llRow, ITEMINDEX) = slInvItem
        grdMultiMedia.Col = INVINDEX
        grdMultiMedia.CellBackColor = LIGHTYELLOW
        If llUnitTotal > 0 Then
            grdMultiMedia.TextMatrix(llRow, INVINDEX) = Trim$(str$(llUnitTotal))
        Else
            grdMultiMedia.TextMatrix(llRow, INVINDEX) = ""
        End If
        grdMultiMedia.Col = AVAILSORDERINDEX
        grdMultiMedia.CellBackColor = LIGHTYELLOW
        'If llTypeUnitsOrdered > 0 Then
            If llUnitTotal < llTypeUnitsOrdered Then
                grdMultiMedia.CellForeColor = vbMagenta
            ElseIf (llUnitTotal * imAvailColorLevel) \ 100 < llTypeUnitsOrdered Then
                grdMultiMedia.CellForeColor = DARKYELLOW
            Else
                grdMultiMedia.CellForeColor = vbBlack
            End If
            grdMultiMedia.TextMatrix(llRow, AVAILSORDERINDEX) = llUnitTotal - llTypeUnitsOrdered
        'Else
        '    grdMultiMedia.TextMatrix(llRow, TYPEAVAILSORDERINDEX) = ""
        'End If
        grdMultiMedia.Col = AVAILSPROPOSALINDEX
        grdMultiMedia.CellBackColor = LIGHTYELLOW
        'If llTypeUnitsProp <> 0 Then
            If llUnitTotal < llTypeUnitsProp Then
                grdMultiMedia.CellForeColor = vbMagenta
            ElseIf (llUnitTotal * imAvailColorLevel) \ 100 < llTypeUnitsProp Then
                grdMultiMedia.CellForeColor = DARKYELLOW
            Else
                grdMultiMedia.CellForeColor = vbBlack
            End If
            grdMultiMedia.TextMatrix(llRow, AVAILSPROPOSALINDEX) = llUnitTotal - llTypeUnitsProp
        'Else
        '    grdMultiMedia.TextMatrix(llRow, TYPEAVAILSPROPOSALINDEX) = ""
        'End If
        grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = ""
        grdMultiMedia.TextMatrix(llRow, NOGAMESINDEX) = ""
        If tmIhf.sGameIndependent = "Y" Then
            grdMultiMedia.Col = NOGAMESINDEX
            grdMultiMedia.CellBackColor = LIGHTYELLOW
            grdMultiMedia.TextMatrix(llRow, INDEPENDENTINDEX) = "Y"
            grdMultiMedia.TextMatrix(llRow, NOGAMESINDEX) = "Event-Ind"
        Else
            grdMultiMedia.Col = NOGAMESINDEX
            grdMultiMedia.CellBackColor = vbWhite
            grdMultiMedia.TextMatrix(llRow, INDEPENDENTINDEX) = "N"
        End If
        grdMultiMedia.Col = AVGCOSTINDEX
        grdMultiMedia.CellBackColor = LIGHTYELLOW
        grdMultiMedia.TextMatrix(llRow, AVGCOSTINDEX) = ""
        If llAvgCost > 0 Then
            grdMultiMedia.TextMatrix(llRow, AVGCOSTINDEX) = gLongToStrDec(llAvgCost, 2)
        End If
        grdMultiMedia.TextMatrix(llRow, AVGRATEINDEX) = ""
        If llAvgRate > 0 Then
            grdMultiMedia.TextMatrix(llRow, AVGRATEINDEX) = gLongToStrDec(llAvgRate, 2)
        End If
        grdMultiMedia.TextMatrix(llRow, TOTALRATEINDEX) = ""
        grdMultiMedia.TextMatrix(llRow, MSFCODEINDEX) = 0
        For llMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
            If (tgMsfCntr(llMsf).MsfRec.iIhfCode = tmIhf.iCode) And (imVefCode = tgMsfCntr(llMsf).MsfRec.iVefCode) Then
                'grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = tgMsfCntr(llMsf).MsfRec.iNoUnits
                'grdMultiMedia.TextMatrix(llRow, AVGRATEINDEX) = gLongToStrDec(tgMsfCntr(llMsf).MsfRec.lRate, 2)
                ilNoGames = 0
                ilNoUnits = 0
                llRate = 0
                llCost = 0
                If tmIhf.sGameIndependent <> "Y" Then
                    llMgf = tgMsfCntr(llMsf).iFirstMgf
                    Do While llMgf <> -1
                        If tgMgfCntr(llMgf).MgfRec.iGameNo > 0 Then
                            ilNoGames = ilNoGames + 1
                            ilNoUnits = ilNoUnits + tgMgfCntr(llMgf).MgfRec.iNoUnits
                            llRate = llRate + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate
                            llCost = llCost + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lCost
                        End If
                        llMgf = tgMgfCntr(llMgf).iNextMgf
                    Loop
                    If ilNoGames > 0 Then
                        grdMultiMedia.TextMatrix(llRow, NOGAMESINDEX) = ilNoGames
                        If (ilNoUnits Mod ilNoGames) = 0 Then
                            grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = ilNoUnits / ilNoGames
                        Else
                            grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = gIntToStrDec((10 * ilNoUnits) / ilNoGames, 1)
                        End If
                        If ilNoUnits > 0 Then
                            grdMultiMedia.TextMatrix(llRow, AVGCOSTINDEX) = gLongToStrDec(llCost / ilNoUnits, 2)
                            grdMultiMedia.TextMatrix(llRow, TOTALRATEINDEX) = gLongToStrDec(llRate, 2)
                            grdMultiMedia.TextMatrix(llRow, AVGRATEINDEX) = gLongToStrDec(llRate / ilNoUnits, 2)
                        End If
                    End If
                    grdMultiMedia.TextMatrix(llRow, MSFCODEINDEX) = tgMsfCntr(llMsf).MsfRec.lCode
                Else
                    llMgf = tgMsfCntr(llMsf).iFirstMgf
                    If llMgf <> -1 Then
                        grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = tgMgfCntr(llMgf).MgfRec.iNoUnits
                        grdMultiMedia.TextMatrix(llRow, AVGCOSTINDEX) = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lCost, 2)
                        grdMultiMedia.TextMatrix(llRow, AVGRATEINDEX) = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lRate, 2)
                        grdMultiMedia.TextMatrix(llRow, TOTALRATEINDEX) = gLongToStrDec(tgMgfCntr(llMgf).MgfRec.lRate * tgMgfCntr(llMgf).MgfRec.iNoUnits, 2)
                    End If
                    grdMultiMedia.TextMatrix(llRow, MSFCODEINDEX) = tgMsfCntr(llMsf).MsfRec.lCode
                End If
                grdMultiMedia.TextMatrix(llRow, COMMENTINDEX) = ""
                If tgMsfCntr(llMsf).iCxfIndex >= 0 Then
                    grdMultiMedia.TextMatrix(llRow, COMMENTINDEX) = sgMsfCntrCxf(tgMsfCntr(llMsf).iCxfIndex)
                End If
                Exit For
            End If
        Next llMsf
        grdMultiMedia.TextMatrix(llRow, IHFCODEINDEX) = tmIhf.iCode
        llRow = llRow + 2
        ilRet = btrGetNext(hmIhf, tmIhf, imIhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Next ilIhf
    grdMultiMedia.Redraw = True
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
Private Function mIhfReadRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIhfCode                                                                             *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mIhfReadRecErr                                                                        *
'******************************************************************************************

'
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer

    smIgnoreMultiFeed = "N"
    ReDim tmIsf(0 To 0) As ISF
    ilUpper = 0
    tmIhfSrchKey0.iCode = imIhfCode
    ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        ilRet = mIsfReadRec(tmIhf.iCode)
        'tmItfSrchKey0.iCode = tmIhf.iItfCode
        'ilRet = btrGetEqual(hmItf, tmMFItf, imItfRecLen, tmItfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        'If ilRet = BTRV_ERR_NONE Then
        '    smIgnoreMultiFeed = tmMFItf.sMultiFeed
        'End If
        For ilLoop = 0 To UBound(tmItf) - 1 Step 1
            If tmIhf.iItfCode = tmItf(ilLoop).iCode Then
                smIgnoreMultiFeed = tmItf(ilLoop).sMultiFeed
                Exit For
            End If
        Next ilLoop
    Else
        mIhfReadRec = False
        Exit Function
    End If
    mIhfReadRec = True
    Exit Function
mIhfReadRecErr: 'VBC NR
    On Error GoTo 0
    mIhfReadRec = False
    Exit Function
End Function

Private Sub mBuildSoldInv(ilIncludeCntr As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilUpper As Integer
    Dim ilGetGames As Integer
    Dim ilMsf As Integer
    Dim ilReplace As Integer
    Dim ilReSet As Integer
    Dim tlIhf As IHF
    Dim ilLegalGameNo As Integer
    ReDim tmMediaInvInfo(0 To 0) As MEDIAINVINFO

    tmMsfSrchKey1.iVefCode = imVefCode
    ilRet = btrGetGreaterOrEqual(hmMsf, tmMsf, imMsfRecLen, tmMsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmMsf.iVefCode = imVefCode)
        If tmMsf.lghfcode = lmSeasonGhfCode Then
            tmChfSrchKey0.lCode = tmMsf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilGetGames = 0
                If ilIncludeCntr Or ((Not ilIncludeCntr) And (tmChf.lCntrNo <> tgChfCntr.lCntrNo)) Then
                    If tmChf.sDelete <> "Y" Then
                        If tmChf.sSchStatus = "F" Then
                            ilGetGames = 1
                        Else
                            If (tmChf.sStatus = "C") Or (tmChf.sStatus = "G") Or (tmChf.sStatus = "N") Or ((tmChf.lCode = tgChfCntr.lCode) And ((tmChf.sStatus = "W") Or (tmChf.sStatus = "I"))) Then
                                ilGetGames = 2
                            End If
                        End If
                    End If
                End If
                If ilGetGames > 0 Then
                    For ilMsf = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                        If (tmMediaInvInfo(ilMsf).lCntrNo = tmChf.lCntrNo) Then
                            If tmMediaInvInfo(ilMsf).lChfCode <> tmChf.lCode Then
                                ilReplace = False
                                If ilGetGames = 1 Then
                                    If (tmChf.iCntRevNo > tmMediaInvInfo(ilMsf).iCntRevNo) Then
                                        ilReplace = True
                                    End If
                                Else
                                    If (tmMediaInvInfo(ilMsf).iCntRevNo > 0) Or (tmChf.iCntRevNo > 0) Then
                                        If (tmChf.iCntRevNo > tmMediaInvInfo(ilMsf).iCntRevNo) Then
                                            ilReplace = True
                                        End If
                                    Else
                                        If tmChf.iMnfPotnType > 0 Then
                                            If tmMediaInvInfo(ilMsf).iMnfPotnType <= 0 Then
                                                ilReplace = True
                                            End If
                                        Else
                                            If tmChf.iPropVer > tmMediaInvInfo(ilMsf).iPropVer Then
                                                ilReplace = True
                                            End If
                                        End If
                                    End If
                                End If
                                'Remove all other game values
                                If ilReplace Then
                                    For ilReSet = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                                        If (tmMediaInvInfo(ilReSet).lCntrNo = tmChf.lCntrNo) Then
                                            tmMediaInvInfo(ilReSet).lChfCode = tmChf.lCode
                                            tmMediaInvInfo(ilReSet).iNoUnitsOrdered = 0
                                            tmMediaInvInfo(ilReSet).iNoUnitsProp = 0
                                            tmMediaInvInfo(ilReSet).iMnfPotnType = tmChf.iMnfPotnType
                                            tmMediaInvInfo(ilReSet).iCntRevNo = tmChf.iCntRevNo
                                            tmMediaInvInfo(ilReSet).iPropVer = tmChf.iPropVer
                                        End If
                                    Next ilReSet
                                End If
                            End If
                            Exit For
                        End If
                    Next ilMsf
                    tmIhfSrchKey0.iCode = tmMsf.iIhfCode
                    ilRet = btrGetEqual(hmIhf, tlIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    tmMgfSrchKey1.lMsfCode = tmMsf.lCode
                    tmMgfSrchKey1.iGameNo = 0
                    ilRet = btrGetGreaterOrEqual(hmMgf, tmMgf, imMgfRecLen, tmMgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmMgf.lMsfCode = tmMsf.lCode)
                        If tlIhf.sGameIndependent = "Y" Then
                            If tmMgf.iGameNo = 0 Then
                                ilLegalGameNo = True
                            Else
                                ilLegalGameNo = False
                            End If
                        Else
                            If tmMgf.iGameNo > 0 Then
                                ilLegalGameNo = True
                            Else
                                ilLegalGameNo = False
                            End If
                        End If
                        ilFound = False
                        ilReplace = False
                        ilUpper = UBound(tmMediaInvInfo)
                        For ilMsf = 0 To ilUpper - 1 Step 1
                            If (tmMediaInvInfo(ilMsf).lCntrNo = tmChf.lCntrNo) And (ilLegalGameNo) And (tmMediaInvInfo(ilMsf).iGameNo = tmMgf.iGameNo) And (tmMediaInvInfo(ilMsf).iIhfCode = tmMsf.iIhfCode) Then
                                ilFound = True
                                If ilGetGames = 1 Then
                                    tmMediaInvInfo(ilMsf).iNoUnitsOrdered = tmMediaInvInfo(ilMsf).iNoUnitsOrdered + tmMgf.iNoUnits
                                Else
                                    If (tmMediaInvInfo(ilMsf).iCntRevNo > 0) Or (tmChf.iCntRevNo > 0) Then
                                        If (tmMediaInvInfo(ilMsf).iCntRevNo = tmChf.iCntRevNo) Then
                                            tmMediaInvInfo(ilMsf).iNoUnitsProp = tmMediaInvInfo(ilMsf).iNoUnitsProp + tmMgf.iNoUnits
                                            tmMediaInvInfo(ilMsf).iMnfPotnType = tmChf.iMnfPotnType
                                            tmMediaInvInfo(ilMsf).iCntRevNo = tmChf.iCntRevNo
                                            tmMediaInvInfo(ilMsf).iPropVer = tmChf.iPropVer
                                        ElseIf (tmChf.iCntRevNo > tmMediaInvInfo(ilMsf).iCntRevNo) Then
                                            tmMediaInvInfo(ilMsf).iNoUnitsProp = tmMgf.iNoUnits
                                            tmMediaInvInfo(ilMsf).iMnfPotnType = tmChf.iMnfPotnType
                                            tmMediaInvInfo(ilMsf).iCntRevNo = tmChf.iCntRevNo
                                            tmMediaInvInfo(ilMsf).iPropVer = tmChf.iPropVer
                                        End If
                                    Else
                                        If tmChf.iMnfPotnType > 0 Then
                                            If tmMediaInvInfo(ilMsf).iMnfPotnType > 0 Then
                                                tmMediaInvInfo(ilMsf).iNoUnitsProp = tmMediaInvInfo(ilMsf).iNoUnitsProp + tmMgf.iNoUnits
                                                tmMediaInvInfo(ilMsf).iMnfPotnType = tmChf.iMnfPotnType
                                                tmMediaInvInfo(ilMsf).iCntRevNo = tmChf.iCntRevNo
                                                tmMediaInvInfo(ilMsf).iPropVer = tmChf.iPropVer
                                            Else
                                                tmMediaInvInfo(ilMsf).iNoUnitsProp = tmMgf.iNoUnits
                                                tmMediaInvInfo(ilMsf).iMnfPotnType = tmChf.iMnfPotnType
                                                tmMediaInvInfo(ilMsf).iCntRevNo = tmChf.iCntRevNo
                                                tmMediaInvInfo(ilMsf).iPropVer = tmChf.iPropVer
                                            End If
                                        Else
                                            If tmChf.iPropVer = tmMediaInvInfo(ilMsf).iPropVer Then
                                                tmMediaInvInfo(ilMsf).iNoUnitsProp = tmMediaInvInfo(ilMsf).iNoUnitsProp + tmMgf.iNoUnits
                                                tmMediaInvInfo(ilMsf).iMnfPotnType = tmChf.iMnfPotnType
                                                tmMediaInvInfo(ilMsf).iCntRevNo = tmChf.iCntRevNo
                                                tmMediaInvInfo(ilMsf).iPropVer = tmChf.iPropVer
                                            ElseIf tmChf.iPropVer > tmMediaInvInfo(ilMsf).iPropVer Then
                                                tmMediaInvInfo(ilMsf).iNoUnitsProp = tmMgf.iNoUnits
                                                tmMediaInvInfo(ilMsf).iMnfPotnType = tmChf.iMnfPotnType
                                                tmMediaInvInfo(ilMsf).iCntRevNo = tmChf.iCntRevNo
                                                tmMediaInvInfo(ilMsf).iPropVer = tmChf.iPropVer
                                            End If
                                        End If
                                    End If
                                End If
                                Exit For
                            End If
                        Next ilMsf
                        If (Not ilFound) And (Not ilReplace) Then
                            For ilMsf = 0 To ilUpper - 1 Step 1
                                If (tmMediaInvInfo(ilMsf).lCntrNo = tmChf.lCntrNo) And (tmMediaInvInfo(ilMsf).iIhfCode = tmMsf.iIhfCode) Then
                                    If tmMediaInvInfo(ilMsf).iNoUnitsOrdered <= 0 Then
                                        ilReplace = False
                                        If ilGetGames = 1 Then
                                            ilReplace = True
                                        Else
                                            If (tmMediaInvInfo(ilMsf).iCntRevNo > 0) Or (tmChf.iCntRevNo > 0) Then
                                                If (tmChf.iCntRevNo > tmMediaInvInfo(ilMsf).iCntRevNo) Then
                                                    ilReplace = True
                                                End If
                                            Else
                                                If tmChf.iMnfPotnType > 0 Then
                                                    If tmMediaInvInfo(ilMsf).iMnfPotnType <= 0 Then
                                                        ilReplace = True
                                                    End If
                                                Else
                                                    If tmChf.iPropVer > tmMediaInvInfo(ilMsf).iPropVer Then
                                                        ilReplace = True
                                                    End If
                                                End If
                                            End If
                                        End If
                                        'Remove all other game values
                                        If ilReplace Then
                                            For ilReSet = 0 To ilUpper - 1 Step 1
                                                If (tmMediaInvInfo(ilReSet).lCntrNo = tmChf.lCntrNo) Then
                                                    tmMediaInvInfo(ilReSet).iNoUnitsOrdered = 0
                                                    tmMediaInvInfo(ilReSet).iNoUnitsProp = 0
                                                    tmMediaInvInfo(ilReSet).iMnfPotnType = tmChf.iMnfPotnType
                                                    tmMediaInvInfo(ilReSet).iCntRevNo = tmChf.iCntRevNo
                                                    tmMediaInvInfo(ilReSet).iPropVer = tmChf.iPropVer
                                                End If
                                            Next ilReSet
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next ilMsf
                        End If
                        If (Not ilFound) And (ilLegalGameNo) Then
                            tmMediaInvInfo(ilUpper).iIhfCode = tmMsf.iIhfCode
                            tmMediaInvInfo(ilUpper).lCntrNo = tmChf.lCntrNo
                            tmMediaInvInfo(ilUpper).lChfCode = tmChf.lCode
                            tmMediaInvInfo(ilUpper).iGameNo = tmMgf.iGameNo
                            If ilGetGames = 1 Then
                                tmMediaInvInfo(ilUpper).iNoUnitsOrdered = tmMgf.iNoUnits
                                tmMediaInvInfo(ilUpper).iNoUnitsProp = 0
                                tmMediaInvInfo(ilUpper).iMnfPotnType = tmChf.iMnfPotnType
                                tmMediaInvInfo(ilUpper).iCntRevNo = tmChf.iCntRevNo
                                tmMediaInvInfo(ilUpper).iPropVer = tmChf.iPropVer
                            Else
                                tmMediaInvInfo(ilUpper).iNoUnitsOrdered = 0
                                tmMediaInvInfo(ilUpper).iNoUnitsProp = tmMgf.iNoUnits
                                tmMediaInvInfo(ilUpper).iMnfPotnType = tmChf.iMnfPotnType
                                tmMediaInvInfo(ilUpper).iCntRevNo = tmChf.iCntRevNo
                                tmMediaInvInfo(ilUpper).iPropVer = tmChf.iPropVer
                            End If
                            ReDim Preserve tmMediaInvInfo(0 To ilUpper + 1) As MEDIAINVINFO
                        End If
                        ilRet = btrGetNext(hmMgf, tmMgf, imMgfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    Loop
                End If
            End If
        End If
        ilRet = btrGetNext(hmMsf, tmMsf, imMsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
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
    Dim llRow As Long
    Dim slName As String
    Dim slCode As String
    Dim llCol As Long
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilVef As Integer
    Dim llGameTotal As Long
    Dim llNonGameTotal As Long
    Dim llGhfStartDate As Long
    Dim llGhfEndDate As Long
    Dim llChfStartDate As Long
    Dim llChfEndDate As Long

    ''ilRet = gPopUserVehicleBox(Program, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHAIRING + VEHVIRTUAL + ACTIVEVEH, cbcVeh, Traffic!lbcVehicle)
    ilRet = gPopUserVehicleBoxNoForm(VEHSPORT + ACTIVEVEH, lbcVehicle, tmGameVehicle(), smGameVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ And ilRet <> CP_MSG_POPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", Contract
        On Error GoTo 0
    Else
        grdSelect.Redraw = False
        grdSelect.Row = 0
        For llCol = VEHICLEINDEX To GAMEDOLLARSINDEX Step 1
            grdSelect.Col = llCol
            grdSelect.CellBackColor = LIGHTBLUE
        Next llCol
        grdSelect.rowHeight(0) = fgBoxGridH + 15
        If UBound(tmGameVehicle) <= 0 Then
            grdSelect.Redraw = True
            Exit Sub
        End If
        llRow = grdSelect.FixedRows
        For ilLoop = 0 To UBound(tmGameVehicle) - 1 Step 1
            slStr = tmGameVehicle(ilLoop).sKey
            ilRet = gParseItem(slStr, 2, "\", slCode)
            ilVef = gBinarySearchVef(Val(slCode))
            If ilVef <> -1 Then
                tmGhfSrchKey1.iVefCode = Val(slCode)
                ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = Val(slCode))
                    gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llGhfStartDate
                    gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llGhfEndDate
                    gUnpackDateLong tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), llChfStartDate
                    'If tgChfCntr.lCntrNo > 0 Then
                    '    gUnpackDateLong tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), llChfEndDate
                    'Else
                        llChfEndDate = gDateValue("12/31/2069")
                    'End If
                    If (llChfEndDate >= llGhfStartDate) And (llChfStartDate <= llGhfEndDate) Then
                        tmIhfSrchKey1.lghfcode = tmGhf.lCode
                        ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                        If (ilRet = BTRV_ERR_NONE) And (tmIhf.lghfcode = tmGhf.lCode) Then
                            If llRow >= grdSelect.Rows Then
                                grdSelect.AddItem ""
                            End If
                            grdSelect.rowHeight(llRow) = fgBoxGridH + 15
                            For llCol = VEHICLEINDEX To GAMEDOLLARSINDEX Step 1
                                grdSelect.Row = llRow
                                grdSelect.Col = llCol
                                grdSelect.CellBackColor = vbWhite
                                grdSelect.CellForeColor = vbBlue
                            Next llCol
                            grdSelect.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgMVef(ilVef).sName)
                            grdSelect.TextMatrix(llRow, SEASONINDEX) = Trim$(tmGhf.sSeasonName)
                            grdSelect.TextMatrix(llRow, GAMEDOLLARSINDEX) = ""
                            mTotalByVehicle tgMVef(ilVef).iCode, llGameTotal, llNonGameTotal
                            If llGameTotal + llNonGameTotal > 0 Then
                                grdSelect.TextMatrix(llRow, GAMEDOLLARSINDEX) = gLongToStrDec(llGameTotal + llNonGameTotal, 2)
                            End If
                            grdSelect.TextMatrix(llRow, SELECTEDINDEX) = "N"
                            grdSelect.TextMatrix(llRow, SEASONGHFCODEINDEX) = tmGhf.lCode
                            grdSelect.TextMatrix(llRow, VEFCODEINDEX) = Trim$(slCode)
                            llRow = llRow + 1
                        End If
                    End If
                    ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                Loop
            End If
        Next ilLoop
        mSelectSortCol VEHICLEINDEX
        mSelectSortCol GAMEDOLLARSINDEX
        grdSelect.Row = 0
        grdSelect.Col = VEFCODEINDEX
        grdSelect.Redraw = True

    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInvTypePopErr                                                                        *
'******************************************************************************************

'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ilRet = gObtainItf(tmItf(), smITFTag)
    Exit Sub
mInvTypePopErr: 'VBC NR
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub



Public Sub Action(ilType As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilTypeItem As Integer
    Dim slStr As String
    Dim slCode As String

    Select Case ilType
        Case 1  'Clear Focus
            mSetShow
            pbcArrow.Visible = False
        Case 2  'Init function
            'Test if unloading control
            ilRet = 0
            On Error GoTo UserControlErr:
            If ilRet = 0 Then
                'ilIndex = cbcGameVeh.ListIndex
                'ilLoop = cbcSelect.ListIndex
'8/5
'                cbcGameVeh.ListIndex = -1
                Form_Load
                Form_Activate
                edcMultiMediaMsg.Visible = False
                Select Case Contract.tscLine.SelectedItem.Index
                    Case imTabMap(TABMULTIMEDIA)    '1  'Multi-Media
                        mInit
                        If (Asc(tgSpf.sUsingFeatures) And MULTIMEDIA) <> MULTIMEDIA Then
                            edcMultiMediaMsg.Text = "Contact Sales@Counterpoint.net to activate this feature"
                            edcMultiMediaMsg.Visible = True
                        End If
                    'Case 2  'Digital
                    Case imTabMap(TABNTR)    '3  'NTR
                    Case imTabMap(TABAIRTIME)    '4  'Air TimeSpotCount
                    Case imTabMap(TABPODCASTCPM)    'Podcast CPM
                    Case imTabMap(TABMERCH)    '5  'Merchandising
                    Case imTabMap(TABPROMO)    '6  'Promotional
                    Case imTabMap(TABINSTALL)    '7  'Installment
                End Select
                'cbcGameVeh.ListIndex = ilIndex
                For ilVef = 0 To UBound(tmGameVehicle) - 1 Step 1
                    slStr = tmGameVehicle(ilVef).sKey
                    ilRet = gParseItem(slStr, 2, "\", slCode)
                    If Val(slCode) = MultiMediaVefCode Then
'8/5
'                        cbcGameVeh.ListIndex = ilVef
'                        For ilTypeItem = 0 To cbcSelect.ListCount - 1 Step 1
'                            If StrComp(cbcSelect.List(ilTypeItem), MultiMediaTypeItem, vbTextCompare) = 0 Then
'                                cbcSelect.ListIndex = ilTypeItem
'                                Exit For
'                            End If
'                        Next ilTypeItem
                        Exit For
                    End If
                Next ilVef
                'cbcSelect.ListIndex = ilLoop
            End If
        Case 3  'terminate function
            mSetShow
            pbcArrow.Visible = False
            cmcCancel_Click
        Case 4  'Clear
            mClearSelectGrid
            If imInitNoRows > 0 Then
                mClearCtrlFields
            End If
'8/5
'            cbcGameVeh.ListIndex = -1
'            cbcSelect.ListIndex = -1
            Screen.MousePointer = vbDefault
            gSetMousePointer grdMultiMedia, grdSelect, vbDefault
        Case 5  'Save
            mSetShow
            pbcArrow.Visible = False
            mSaveRec
    End Select
    Exit Sub
UserControlErr:
    ilRet = 1
    Resume Next
End Sub
Public Property Let Enabled(ilState As Integer)
    UserControl.Enabled = ilState
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Form_MouseUp Button, Shift, X, Y
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

    For ilCol = grdMultiMedia.FixedCols To grdMultiMedia.Cols - 1 Step 1
        grdMultiMedia.TextMatrix(llRowNo, ilCol) = ""
    Next ilCol
    'grdMultiMedia.Row = llRowNo
    'grdMultiMedia.Col = 1
    'grdMultiMedia.CellBackColor = vbWhite
    'Horizontal Line
    grdMultiMedia.Row = llRowNo + 1
    For ilCol = 1 To grdMultiMedia.Cols - 1 Step 1
        grdMultiMedia.Col = ilCol
        grdMultiMedia.CellBackColor = vbBlue
    Next ilCol
    'Vertical Lines
    grdMultiMedia.Col = 1
    For llRow = llRowNo To llRowNo + 1 Step 1
        grdMultiMedia.Row = llRow
        grdMultiMedia.CellBackColor = vbBlue
    Next llRow
    grdMultiMedia.Col = 3
    For llRow = llRowNo To llRowNo + 1 Step 1
        grdMultiMedia.Row = llRow
        grdMultiMedia.CellBackColor = vbBlue
    Next llRow
    For ilCol = grdMultiMedia.FixedCols + 1 To grdMultiMedia.Cols - 1 Step 2
        grdMultiMedia.Col = ilCol
        For llRow = llRowNo To llRowNo + 1 Step 1
            grdMultiMedia.Row = llRow
            grdMultiMedia.CellBackColor = vbBlue
        Next llRow
    Next ilCol
    'Set Fix area Column to white
    grdMultiMedia.Col = 2
    grdMultiMedia.Row = llRowNo
    grdMultiMedia.CellBackColor = vbWhite
    If grdMultiMedia.rowHeight(grdMultiMedia.TopRow) <= 15 Then
        grdMultiMedia.TopRow = grdMultiMedia.TopRow + 1
    End If
End Sub

Private Sub mSaveRec()
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim llMsf As Long
    Dim llMgf As Long
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilGsf As Integer
    Dim slDate As String
    Dim llDate As Long
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbfIndex As Integer
    Dim llSbfIndex As Long
    Dim llSbfDate As Long
    Dim ilItf As Integer
    Dim slInvType As String
    Dim slInvItem As String
    Dim ilMatch As Integer
    Dim llEarliestSbfDate As Long

    If imMsfChg Then
        'replace NTR for MultiMedia
        For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
            If tgIBSbf(llSbf).SbfRec.iIhfCode > 0 Then
                If tgIBSbf(llSbf).iStatus >= 0 Then
                    tgIBSbf(llSbf).SbfRec.lGross = -1
                End If
            End If
        Next llSbf
        llEarliestSbfDate = 99999999
        For llMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
            tmIhfSrchKey0.iCode = tgMsfCntr(llMsf).MsfRec.iIhfCode
            ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                imVefCode = tgMsfCntr(llMsf).MsfRec.iVefCode
                ilRet = mGhfGsfReadRec()
                llMgf = tgMsfCntr(llMsf).iFirstMgf
                Do While llMgf <> -1
                    slDate = ""
                    llDate = 0
                    If tgMgfCntr(llMgf).MgfRec.iGameNo > 0 Then
                        For ilGsf = 0 To UBound(tmGsf) - 1 Step 1
                            If tmGsf(ilGsf).iGameNo = tgMgfCntr(llMgf).MgfRec.iGameNo Then
                                gUnpackDate tmGsf(ilGsf).iAirDate(0), tmGsf(ilGsf).iAirDate(1), slStr
                                slDate = gObtainEndStd(slStr)
                                llDate = gDateValue(slDate)
                                Exit For
                            End If
                        Next ilGsf
                    End If
                    If ((slDate <> "") Or (tgMgfCntr(llMgf).MgfRec.iGameNo = 0)) Then
                        'Look for a match
                        ilFound = False
                        For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
                            If (tgIBSbf(llSbf).iStatus >= 0) And (tgIBSbf(llSbf).SbfRec.iBillVefCode = imVefCode) And (tgIBSbf(llSbf).SbfRec.iIhfCode = tgMsfCntr(llMsf).MsfRec.iIhfCode) Then
                                gUnpackDateLong tgIBSbf(llSbf).SbfRec.iDate(0), tgIBSbf(llSbf).SbfRec.iDate(1), llSbfDate
                                ilMatch = False
                                If tgMgfCntr(llMgf).MgfRec.iGameNo > 0 Then
                                    If (tgIBSbf(llSbf).SbfRec.iIhfCode = tgMsfCntr(llMsf).MsfRec.iIhfCode) And (llDate = llSbfDate) Then
                                        ilMatch = True
                                    End If
                                Else
                                    If (tgIBSbf(llSbf).SbfRec.iIhfCode = tgMsfCntr(llMsf).MsfRec.iIhfCode) Then
                                        ilMatch = True
                                    End If
                                End If
                                'If (tgIBSbf(llSbf).SbfRec.iIhfCode = tgMsfCntr(llMsf).MsfRec.iIhfCode) And (((llDate = llSbfDate) And (tgMgfCntr(llMgf).MgfRec.iGameNo > 0)) Or ((tgMgfCntr(llMgf).MgfRec.iGameNo = 0) And (llDate = 0))) Then
                                If ilMatch Then
                                    If tgIBSbf(llSbf).SbfRec.lGross = -1 Then
                                        tgIBSbf(llSbf).SbfRec.lGross = 0
                                        tgIBSbf(llSbf).SbfRec.lAcquisitionCost = 0
                                        If (llSbfDate > 0) And (llDate < llEarliestSbfDate) Then
                                            llEarliestSbfDate = llDate
                                        End If
                                    End If
                                    tgIBSbf(llSbf).SbfRec.lGross = tgIBSbf(llSbf).SbfRec.lGross + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate
                                    tgIBSbf(llSbf).SbfRec.lAcquisitionCost = tgIBSbf(llSbf).SbfRec.lAcquisitionCost + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lCost
                                    ilFound = True
                                End If
                            End If
                        Next llSbf
                        If Not ilFound Then
                            llSbfIndex = -1
                            For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
                                If tgIBSbf(llSbf).iStatus = -1 Then
                                    tgIBSbf(llSbf).iStatus = 0
                                    llSbfIndex = llSbf
                                    Exit For
                                ElseIf tgIBSbf(llSbf).iStatus = 2 Then
                                    tgIBSbf(llSbf).iStatus = 1
                                    llSbfIndex = llSbf
                                    Exit For
                                End If
                            Next llSbf
                            If llSbfIndex = -1 Then
                                ReDim Preserve tgIBSbf(0 To UBound(tgIBSbf) + 1) As SBFLIST
                                llSbfIndex = UBound(tgIBSbf) - 1
                                tgIBSbf(llSbfIndex).iStatus = 0
                            End If

                            tgIBSbf(llSbfIndex).SbfRec.lChfCode = tgChfCntr.lCode
                            tgIBSbf(llSbfIndex).SbfRec.sTranType = "I"
                            tgIBSbf(llSbfIndex).SbfRec.iBillVefCode = imVefCode
                            tgIBSbf(llSbfIndex).SbfRec.iAirVefCode = imVefCode
                            gPackDate slDate, tgIBSbf(llSbfIndex).SbfRec.iDate(0), tgIBSbf(llSbfIndex).SbfRec.iDate(1)
                            gPackDate slDate, tgIBSbf(llSbfIndex).SbfRec.iPrintInvDate(0), tgIBSbf(llSbfIndex).SbfRec.iPrintInvDate(1)
                            If (llDate > 0) And (llSbfDate < llEarliestSbfDate) Then
                                llEarliestSbfDate = llDate
                            End If
                            tgIBSbf(llSbfIndex).SbfRec.sDescr = ""
                            slInvType = ""
                            For ilItf = 0 To UBound(tmItf) - 1 Step 1
                                If tmIhf.iItfCode = tmItf(ilItf).iCode Then
                                    slInvType = Trim$(tmItf(ilItf).sName)
                                    Exit For
                                End If
                            Next ilItf
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
                            tgIBSbf(llSbfIndex).SbfRec.sDescr = slInvType & "/" & slInvItem
                            tgIBSbf(llSbfIndex).SbfRec.iMnfItem = imNTRMnfCode
                            tgIBSbf(llSbfIndex).SbfRec.lGross = tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate
                            tgIBSbf(llSbfIndex).SbfRec.lAcquisitionCost = tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lCost
                            tgIBSbf(llSbfIndex).SbfRec.iNoItems = 1
                            If igDirAdvt Then
                                tgIBSbf(llSbfIndex).SbfRec.sAgyComm = "N"
                            Else
                                tgIBSbf(llSbfIndex).SbfRec.sAgyComm = "Y"
                            End If
                            tgIBSbf(llSbfIndex).SbfRec.iCommPct = imNTRSlspComm
                            tgIBSbf(llSbfIndex).SbfRec.iTrfCode = 0
                            tgIBSbf(llSbfIndex).SbfRec.sBilled = "N"
                            tgIBSbf(llSbfIndex).SbfRec.iIhfCode = tgMsfCntr(llMsf).MsfRec.iIhfCode
                            tgIBSbf(llSbfIndex).SbfRec.iLineNo = 0
                        End If
                    End If
                    llMgf = tgMgfCntr(llMgf).iNextMgf
                Loop
            End If
        Next llMsf
        llDate = gDateValue(Format(gNow(), "m/d/yy"))
        If (llEarliestSbfDate <> 99999999) And (llDate <= llEarliestSbfDate) Then
            llDate = llEarliestSbfDate
        End If
        slStr = Format$(llDate, "m/d/yy")
        slDate = gObtainEndStd(slStr)
        llDate = gDateValue(slDate)
        For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
            If (tgIBSbf(llSbf).SbfRec.iIhfCode > 0) And (tgIBSbf(llSbf).SbfRec.lGross = -1) And (tgIBSbf(llSbf).iStatus >= 0) Then
                If tgIBSbf(llSbf).iStatus = 0 Then
                    tgIBSbf(llSbf).iStatus = -1
                ElseIf tgIBSbf(llSbf).iStatus = 1 Then
                    tgIBSbf(llSbf).iStatus = 2
                End If
                tgIBSbf(llSbf).SbfRec.sTranType = ""
                tgIBSbf(llSbf).SbfRec.iBillVefCode = 0
                tgIBSbf(llSbf).SbfRec.iDate(0) = 0
                tgIBSbf(llSbf).SbfRec.iDate(1) = 0
                'gStrToPDN "", 2, 5, tmIBSbf(ilIndex).SbfRec.sItemAmount
                tgIBSbf(llSbf).SbfRec.lGross = 0
                tgIBSbf(llSbf).SbfRec.lAcquisitionCost = 0
                tgIBSbf(llSbf).SbfRec.iMnfItem = 0
                tgIBSbf(llSbf).SbfRec.iNoItems = 0
                'tmIBSbf(llSbf).SbfRec.sUnitName = ""
                tgIBSbf(llSbf).SbfRec.sDescr = ""
                tgIBSbf(llSbf).SbfRec.sBilled = "N"
                tgIBSbf(llSbf).SbfRec.iPrintInvDate(0) = 0
                tgIBSbf(llSbf).SbfRec.iPrintInvDate(1) = 0
                tgIBSbf(llSbf).SbfRec.iIhfCode = 0
            ElseIf (tgIBSbf(llSbf).SbfRec.iIhfCode > 0) And (tgIBSbf(llSbf).SbfRec.lGross >= 0) And (tgIBSbf(llSbf).iStatus >= 0) And (tgIBSbf(llSbf).SbfRec.iDate(0) = 0) And (tgIBSbf(llSbf).SbfRec.iDate(1) = 0) Then
                gPackDateLong llDate, tgIBSbf(llSbf).SbfRec.iDate(0), tgIBSbf(llSbf).SbfRec.iDate(1)
            End If
        Next llSbf
    End If
End Sub

Public Property Get Verify() As Integer
    Dim llMsf As Long
    Dim llMgf As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llUnitTotal As Long
    Dim llTypeUnitsOrdered As Long
    Dim ilNoUnits As Integer
    Dim llInfo As Long
    Dim ilCheck As Integer
    Dim ilGame As Integer
    Dim slInvType As String
    Dim slInvItem As String
    Dim llDate As Long
    Dim ilInitValues As Integer

    pbcArrow.Visible = False
    If imUpdateAllowed Then
        mInvTypePop
        ReDim tlMediaInvInfo(0 To UBound(tmMediaInvInfo)) As MEDIAINVINFO
        For ilLoop = 0 To UBound(tmMediaInvInfo) - 1 Step 1
            LSet tlMediaInvInfo(ilLoop) = tmMediaInvInfo(ilLoop)
        Next ilLoop
        Verify = True
        'Check on Oversold
        For llMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
            If tgMsfCntr(llMsf).iStatus >= 0 Then
                ilInitValues = True
                imIhfCode = tgMsfCntr(llMsf).MsfRec.iIhfCode
                tmIhfSrchKey0.iCode = imIhfCode
                ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If tmIhf.sOversell <> "Y" Then
                        llMgf = tgMsfCntr(llMsf).iFirstMgf
                        If llMgf <> -1 Then
                            If ilInitValues Then
                                imVefCode = tgMsfCntr(llMsf).MsfRec.iVefCode
                                ilRet = mGhfGsfReadRec()
                                mBuildSoldInv False
                                ilInitValues = False
                            End If
                            ilRet = mIsfReadRec(imIhfCode)
                            Do While llMgf <> -1
                                If tgMgfCntr(llMgf).MgfRec.sBilled <> "Y" Then
                                    ilCheck = True
                                    For ilGame = 0 To UBound(tmGsf) - 1 Step 1
                                        If tgMgfCntr(llMgf).MgfRec.iGameNo = tmGsf(ilGame).iGameNo Then
                                            gUnpackDateLong tmGsf(ilGame).iAirDate(0), tmGsf(ilGame).iAirDate(1), llDate
                                            If llDate < lmFirstAllowedChgDate Then
                                                ilCheck = False
                                            End If
                                            Exit For
                                        End If
                                    Next ilGame
                                    If ilCheck Then
                                        llTypeUnitsOrdered = 0
                                        For llInfo = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                                            If (tmMediaInvInfo(llInfo).iIhfCode = imIhfCode) And (tgMgfCntr(llMgf).MgfRec.iGameNo = tmMediaInvInfo(llInfo).iGameNo) Then
                                                llTypeUnitsOrdered = tmMediaInvInfo(llInfo).iNoUnitsOrdered
                                                Exit For
                                            End If
                                        Next llInfo
                                        llUnitTotal = 0
                                        For ilLoop = 0 To UBound(tmIsf) - 1 Step 1
                                            If tmIhf.sGameIndependent <> "Y" Then
                                                If tgMgfCntr(llMgf).MgfRec.iGameNo = tmIsf(ilLoop).iGameNo Then
                                                    llUnitTotal = tmIsf(ilLoop).iNoUnits
                                                    Exit For
                                                End If
                                            Else
                                                If tmIsf(ilLoop).iGameNo = 0 Then
                                                    llUnitTotal = tmIsf(ilLoop).iNoUnits
                                                    Exit For
                                                End If
                                            End If
                                        Next ilLoop
                                        ilNoUnits = tgMgfCntr(llMgf).MgfRec.iNoUnits
                                        If llTypeUnitsOrdered + ilNoUnits > llUnitTotal Then
                                            ReDim tmMediaInvInfo(0 To UBound(tlMediaInvInfo)) As MEDIAINVINFO
                                            For ilLoop = 0 To UBound(tlMediaInvInfo) - 1 Step 1
                                                LSet tmMediaInvInfo(ilLoop) = tlMediaInvInfo(ilLoop)
                                            Next ilLoop
                                            slInvType = ""
                                            For ilLoop = 0 To UBound(tmItf) - 1 Step 1
                                                If tmIhf.iItfCode = tmItf(ilLoop).iCode Then
                                                    slInvType = Trim$(tmItf(ilLoop).sName)
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
                                            ilRet = MsgBox(slInvType & "/" & slInvItem & " Oversold, reduce number units purchased", vbOKOnly + vbExclamation, "Message")
                                            Verify = False
                                            Exit Property
                                        End If
                                    End If
                                End If
                                llMgf = tgMgfCntr(llMgf).iNextMgf
                            Loop
                        End If
                    End If
                End If
            End If
        Next llMsf
        ReDim tmMediaInvInfo(0 To UBound(tlMediaInvInfo)) As MEDIAINVINFO
        For ilLoop = 0 To UBound(tlMediaInvInfo) - 1 Step 1
            LSet tmMediaInvInfo(ilLoop) = tlMediaInvInfo(ilLoop)
        Next ilLoop
    Else
        Verify = True
    End If
End Property



Private Sub mSetNoGames(llRow As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIhfCode                     ilIsf                         tlIsf                     *
'*  ilMsfIndex                                                                            *
'******************************************************************************************

    Dim ilNoGames As Integer
    Dim ilGame As Integer
    Dim ilRet As Integer
    Dim ilNoUnits As Integer
    Dim llRate As Long
    Dim llCost As Long
    Dim llMgf As Long
    Dim llMsf As Long
    Dim ilChanged As Integer
    Dim llMsfIndex As Long
    Dim llMsfCode As Long
    Dim ilFound As Integer
    Dim llMgfIndex As Long
    Dim ilUnitsVary As Integer
    Dim ilUnits As Integer
    Dim llUnitTotal As Long
    Dim llTypeUnitsOrdered As Long
    Dim llTypeUnitsProp As Long
    Dim llInfo As Long
    Dim ilUpper As Integer

    ilNoGames = 0
    ilNoUnits = 0
    llRate = 0
    llCost = 0
    ilUnitsVary = False
    ilUnits = -1
    imIhfCode = Val(grdMultiMedia.TextMatrix(llRow, IHFCODEINDEX))
    'ilRet = mIhfReadRec()
    'Test if any changes
    llMsfIndex = -1
    ilChanged = True
    llMsfCode = Val(grdMultiMedia.TextMatrix(llRow, MSFCODEINDEX))
    For llMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
        If (llMsfCode = tgMsfCntr(llMsf).MsfRec.lCode) And (imVefCode = tgMsfCntr(llMsf).MsfRec.iVefCode) Then
            ilChanged = False
            llMsfIndex = llMsf
            ilFound = False
            llMgf = tgMsfCntr(llMsf).iFirstMgf
            If (llMgf = -1) And (UBound(tgGetGameReturn) <= 0) Then
                grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = ""
                grdMultiMedia.TextMatrix(llRow, NOGAMESINDEX) = ""
                mSetAvgs llRow
                Exit Sub
            End If
            Do While llMgf <> -1
                ilFound = False
                For ilGame = 0 To UBound(tgGetGameReturn) - 1 Step 1
                    If tgGetGameReturn(ilGame).iGameNo = tgMgfCntr(llMgf).MgfRec.iGameNo Then
                        ilFound = True
                        If tgGetGameReturn(ilGame).iNoUnits <> tgMgfCntr(llMgf).MgfRec.iNoUnits Then
                            ilChanged = True
                            Exit Do
                        End If
                        Exit For
                    End If
                Next ilGame
                If Not ilFound Then
                    ilChanged = True
                    Exit Do
                End If
                llMgf = tgMgfCntr(llMgf).iNextMgf
            Loop
            If ilChanged Then
                Exit For
            End If
            For ilGame = 0 To UBound(tgGetGameReturn) - 1 Step 1
                ilFound = False
                llMgf = tgMsfCntr(llMsf).iFirstMgf
                Do While llMgf <> -1
                    If tgGetGameReturn(ilGame).iGameNo = tgMgfCntr(llMgf).MgfRec.iGameNo Then
                        ilFound = True
                        Exit Do
                    End If
                    llMgf = tgMgfCntr(llMgf).iNextMgf
                Loop
                If Not ilFound Then
                    ilChanged = True
                    Exit For
                End If
            Next ilGame
            Exit For
        End If
    Next llMsf
    If Not ilChanged Then
        Exit Sub
    End If
    imMsfChg = True
    If llMsfIndex <> -1 Then
        llMgf = tgMsfCntr(llMsfIndex).iFirstMgf
        Do While llMgf <> -1
            For llInfo = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                If (tmMediaInvInfo(llInfo).iIhfCode = imIhfCode) And (tgMgfCntr(llMgf).MgfRec.iGameNo = tmMediaInvInfo(llInfo).iGameNo) Then
                    tmMediaInvInfo(llInfo).iNoUnitsProp = tmMediaInvInfo(llInfo).iNoUnitsProp - tgMgfCntr(llMgf).MgfRec.iNoUnits
                    Exit For
                End If
            Next llInfo
            tgMgfCntr(llMgf).iStatus = -1
            llMgf = tgMgfCntr(llMgf).iNextMgf
        Loop
    Else
        ReDim Preserve tgMsfCntr(0 To UBound(tgMsfCntr) + 1) As MSFLIST
        llMsfIndex = UBound(tgMsfCntr) - 1
        tgMsfCntr(llMsfIndex).iStatus = 0
        tgMsfCntr(llMsfIndex).iCxfIndex = -1
        tgMsfCntr(llMsfIndex).MsfRec.lCode = -(llMsfIndex + 1)
        tgMsfCntr(llMsfIndex).MsfRec.iIhfCode = Val(grdMultiMedia.TextMatrix(llRow, IHFCODEINDEX))
        tgMsfCntr(llMsfIndex).MsfRec.iVefCode = imVefCode
        tgMsfCntr(llMsfIndex).MsfRec.lghfcode = 0
        tgMsfCntr(llMsfIndex).MsfRec.iInvGameNo = 0
        tmIhfSrchKey0.iCode = tgMsfCntr(llMsfIndex).MsfRec.iIhfCode
        ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            tgMsfCntr(llMsfIndex).MsfRec.lghfcode = tmIhf.lghfcode
        End If
        tgMsfCntr(llMsfIndex).MsfRec.lChfCode = 0
        tgMsfCntr(llMsfIndex).MsfRec.lCxfCode = 0
        tgMsfCntr(llMsfIndex).MsfRec.sUnused = ""
        grdMultiMedia.TextMatrix(llRow, MSFCODEINDEX) = tgMsfCntr(llMsfIndex).MsfRec.lCode
    End If
    tgMsfCntr(llMsfIndex).iFirstMgf = -1
    For ilGame = 0 To UBound(tgGetGameReturn) - 1 Step 1
        'For ilIsf = 0 To UBound(tmIsf) - 1 Step 1
        '    If tgGetGameReturn(ilGame).iGameNo = tmIsf(ilIsf).iGameNo Then
                ilNoGames = ilNoGames + 1
                If ilUnits = -1 Then
                    ilUnits = tgGetGameReturn(ilGame).iNoUnits
                Else
                    If ilUnits <> tgGetGameReturn(ilGame).iNoUnits Then
                        ilUnitsVary = True
                    End If
                End If
                ilNoUnits = ilNoUnits + tgGetGameReturn(ilGame).iNoUnits
                llRate = llRate + tgGetGameReturn(ilGame).iNoUnits * tgGetGameReturn(ilGame).lRate
                llCost = llCost + tgGetGameReturn(ilGame).iNoUnits * tgGetGameReturn(ilGame).lCost
                llMgf = 0
                Do While llMgf < UBound(tgMgfCntr)
                    If tgMgfCntr(llMgf).iStatus = -1 Then
                        Exit Do
                    End If
                    llMgf = llMgf + 1
                Loop
                If llMgf = UBound(tgMgfCntr) Then
                    ReDim Preserve tgMgfCntr(0 To UBound(tgMgfCntr) + 1) As MGFLIST
                End If
                If tgMsfCntr(llMsfIndex).iFirstMgf = -1 Then
                    tgMsfCntr(llMsfIndex).iFirstMgf = llMgf
                Else
                    llMgfIndex = tgMsfCntr(llMsfIndex).iFirstMgf
                    Do While llMgfIndex <> -1
                        If tgMgfCntr(llMgfIndex).iNextMgf = -1 Then
                            tgMgfCntr(llMgfIndex).iNextMgf = llMgf
                            Exit Do
                        End If
                        llMgfIndex = tgMgfCntr(llMgfIndex).iNextMgf
                    Loop
                End If
                tgMgfCntr(llMgf).iStatus = 0
                tgMgfCntr(llMgf).iNextMgf = -1
                tgMgfCntr(llMgf).MgfRec.lCode = 0
                tgMgfCntr(llMgf).MgfRec.iGameNo = tgGetGameReturn(ilGame).iGameNo
                tgMgfCntr(llMgf).MgfRec.iNoUnits = tgGetGameReturn(ilGame).iNoUnits
                tgMgfCntr(llMgf).MgfRec.lRate = tgGetGameReturn(ilGame).lRate
                tgMgfCntr(llMgf).MgfRec.lCost = tgGetGameReturn(ilGame).lCost
                tgMgfCntr(llMgf).MgfRec.lIsfCode = tgGetGameReturn(ilGame).lIsfCode
                tgMgfCntr(llMgf).MgfRec.sBilled = tgGetGameReturn(ilGame).sBilled
                tgMgfCntr(llMgf).MgfRec.sUnused = ""
                ilFound = False
                For llInfo = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                    If (tmMediaInvInfo(llInfo).iIhfCode = imIhfCode) And (tgMgfCntr(llMgf).MgfRec.iGameNo = tmMediaInvInfo(llInfo).iGameNo) Then
                        ilFound = True
                        tmMediaInvInfo(llInfo).iNoUnitsProp = tmMediaInvInfo(llInfo).iNoUnitsProp + tgMgfCntr(llMgf).MgfRec.iNoUnits
                        Exit For
                    End If
                Next llInfo
                If Not ilFound Then
                    ilUpper = UBound(tmMediaInvInfo)
                    tmMediaInvInfo(ilUpper).iIhfCode = imIhfCode
                    tmMediaInvInfo(ilUpper).lCntrNo = tgChfCntr.lCntrNo
                    tmMediaInvInfo(ilUpper).lChfCode = tgChfCntr.lCode
                    tmMediaInvInfo(ilUpper).iGameNo = tgMgfCntr(llMgf).MgfRec.iGameNo
                    tmMediaInvInfo(ilUpper).iNoUnitsOrdered = 0
                    tmMediaInvInfo(ilUpper).iNoUnitsProp = tgMgfCntr(llMgf).MgfRec.iNoUnits
                    tmMediaInvInfo(ilUpper).iMnfPotnType = tgChfCntr.iMnfPotnType
                    tmMediaInvInfo(ilUpper).iCntRevNo = tgChfCntr.iCntRevNo
                    tmMediaInvInfo(ilUpper).iPropVer = tgChfCntr.iPropVer
                    ReDim Preserve tmMediaInvInfo(0 To ilUpper + 1) As MEDIAINVINFO
                End If
                'Exit For
        '    End If
        'Next ilIsf
    Next ilGame
    If ilNoGames <= 0 Then
        If grdMultiMedia.TextMatrix(llRow, INDEPENDENTINDEX) = "N" Then
            grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = ""
            grdMultiMedia.TextMatrix(llRow, NOGAMESINDEX) = ""
        End If
        mSetAvgs llRow
    Else
        If (ilNoGames > 0) And (grdMultiMedia.TextMatrix(llRow, INDEPENDENTINDEX) = "N") Then
            grdMultiMedia.TextMatrix(llRow, NOGAMESINDEX) = ilNoGames
            If Not ilUnitsVary Then
                grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = ilNoUnits / ilNoGames
            Else
                grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = gIntToStrDec((10 * ilNoUnits) / ilNoGames, 1)
            End If
            If ilNoUnits > 0 Then
                grdMultiMedia.TextMatrix(llRow, AVGCOSTINDEX) = gLongToStrDec(llCost / ilNoUnits, 2)
                grdMultiMedia.TextMatrix(llRow, AVGRATEINDEX) = gLongToStrDec(llRate / ilNoUnits, 2)
                grdMultiMedia.TextMatrix(llRow, TOTALRATEINDEX) = gLongToStrDec(llRate, 2)
            End If
        End If
        If grdMultiMedia.TextMatrix(llRow, INDEPENDENTINDEX) = "N" Then
            llTypeUnitsOrdered = 0
            llTypeUnitsProp = 0
            For llInfo = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                If tmMediaInvInfo(llInfo).iIhfCode = imIhfCode Then
                    llTypeUnitsOrdered = llTypeUnitsOrdered + tmMediaInvInfo(llInfo).iNoUnitsOrdered
                    If tmMediaInvInfo(llInfo).iNoUnitsProp > 0 Then
                        llTypeUnitsProp = llTypeUnitsProp + (tmMediaInvInfo(llInfo).iNoUnitsProp - tmMediaInvInfo(llInfo).iNoUnitsOrdered)
                    End If
                End If
            Next llInfo
            llTypeUnitsProp = llTypeUnitsOrdered + llTypeUnitsProp
            llUnitTotal = grdMultiMedia.TextMatrix(lmEnableRow, INVINDEX)
            If llUnitTotal < llTypeUnitsProp Then
                grdMultiMedia.CellForeColor = vbMagenta
            ElseIf (llUnitTotal * imAvailColorLevel) \ 100 < llTypeUnitsProp Then
                grdMultiMedia.CellForeColor = DARKYELLOW
            Else
                grdMultiMedia.CellForeColor = vbBlack
            End If
            grdMultiMedia.TextMatrix(lmEnableRow, AVAILSPROPOSALINDEX) = llUnitTotal - llTypeUnitsProp
        End If
    End If
    mSetTotal
    mSetCommands
End Sub

Private Sub mInitGetGames(llRow As Long, ilModelling As Integer)
    Dim llMsfCode As Long
    Dim llMsf As Long
    Dim llMgf As Long
    Dim ilIhfCode As Integer
    Dim llInfo As Long
    Dim ilLoop As Integer
    Dim ilMinUnits As Integer
    Dim ilGame As Integer
    Dim ilAdd As Integer
    Dim llDate As Long

    ReDim tgGetGameReturn(0 To 0) As GETGAMERETURN
    igGetGameDefaultUnits = -1
    ilMinUnits = -1
    ilIhfCode = Val(grdMultiMedia.TextMatrix(llRow, IHFCODEINDEX))
    llMsfCode = Val(grdMultiMedia.TextMatrix(llRow, MSFCODEINDEX))
    For llMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
        If (llMsfCode = tgMsfCntr(llMsf).MsfRec.lCode) And (imVefCode = tgMsfCntr(llMsf).MsfRec.iVefCode) Then
            llMgf = tgMsfCntr(llMsf).iFirstMgf
            Do While llMgf <> -1
                ilAdd = True
                If ilModelling Then
                    If (tgMgfCntr(llMgf).MgfRec.sBilled = "Y") Then
                        ilAdd = False
                    Else
                        For ilGame = 0 To UBound(tmGsf) - 1 Step 1
                            If tgMgfCntr(llMgf).MgfRec.iGameNo = tmGsf(ilGame).iGameNo Then
                                gUnpackDateLong tmGsf(ilGame).iAirDate(0), tmGsf(ilGame).iAirDate(1), llDate
                                If llDate < lmFirstAllowedChgDate Then
                                    ilAdd = False
                                End If
                                Exit For
                            End If
                        Next ilGame
                    End If
                End If
                If ilAdd Then
                    tgGetGameReturn(UBound(tgGetGameReturn)).iGameNo = tgMgfCntr(llMgf).MgfRec.iGameNo
                    tgGetGameReturn(UBound(tgGetGameReturn)).iNoUnits = tgMgfCntr(llMgf).MgfRec.iNoUnits
                    If igGetGameDefaultUnits = -1 Then
                        igGetGameDefaultUnits = tgMgfCntr(llMgf).MgfRec.iNoUnits
                        ilMinUnits = tgMgfCntr(llMgf).MgfRec.iNoUnits
                    Else
                        If igGetGameDefaultUnits <> tgMgfCntr(llMgf).MgfRec.iNoUnits Then
                            igGetGameDefaultUnits = -2
                        End If
                        If (tgMgfCntr(llMgf).MgfRec.iNoUnits < ilMinUnits) And (tgMgfCntr(llMgf).MgfRec.iNoUnits > 0) Then
                            ilMinUnits = tgMgfCntr(llMgf).MgfRec.iNoUnits
                        End If
                    End If
                    tgGetGameReturn(UBound(tgGetGameReturn)).lRate = tgMgfCntr(llMgf).MgfRec.lRate
                    tgGetGameReturn(UBound(tgGetGameReturn)).lCost = tgMgfCntr(llMgf).MgfRec.lCost
                    tgGetGameReturn(UBound(tgGetGameReturn)).lIsfCode = tgMgfCntr(llMgf).MgfRec.lIsfCode
                    tgGetGameReturn(UBound(tgGetGameReturn)).sBilled = tgMgfCntr(llMgf).MgfRec.sBilled
                    tgGetGameReturn(UBound(tgGetGameReturn)).iNoUnitsOrdered = 0
                    tgGetGameReturn(UBound(tgGetGameReturn)).iNoUnitsProp = 0
                    For llInfo = 0 To UBound(tmMediaInvInfo) - 1 Step 1
                        If (tmMediaInvInfo(llInfo).iIhfCode = ilIhfCode) And (tgMgfCntr(llMgf).MgfRec.iGameNo = tmMediaInvInfo(llInfo).iGameNo) Then
                            tgGetGameReturn(UBound(tgGetGameReturn)).iNoUnitsOrdered = tgGetGameReturn(UBound(tgGetGameReturn)).iNoUnitsOrdered + tmMediaInvInfo(llInfo).iNoUnitsOrdered
                            If tmMediaInvInfo(llInfo).iNoUnitsProp > 0 Then
                                tgGetGameReturn(UBound(tgGetGameReturn)).iNoUnitsProp = tgGetGameReturn(UBound(tgGetGameReturn)).iNoUnitsProp + tmMediaInvInfo(llInfo).iNoUnitsOrdered - tmMediaInvInfo(llInfo).iNoUnitsProp
                            End If
                            Exit For
                        End If
                    Next llInfo
                    ReDim Preserve tgGetGameReturn(0 To UBound(tgGetGameReturn) + 1) As GETGAMERETURN
                End If
                llMgf = tgMgfCntr(llMgf).iNextMgf
            Loop
            Exit For
        End If
    Next llMsf
    For ilLoop = 0 To UBound(tgGetGameReturn) - 1 Step 1
        tgGetGameReturn(ilLoop).iNoUnitsProp = tgGetGameReturn(ilLoop).iNoUnitsOrdered - tgGetGameReturn(ilLoop).iNoUnitsProp
    Next ilLoop
    If igGetGameDefaultUnits = -1 Then
        If grdMultiMedia.TextMatrix(llRow, UNITSINDEX) <> "" Then
            igGetGameDefaultUnits = Val(grdMultiMedia.TextMatrix(llRow, UNITSINDEX))
        Else
            igGetGameDefaultUnits = 1
        End If
    End If
    If igGetGameDefaultUnits = -2 Then
        igGetGameDefaultUnits = ilMinUnits
    End If
End Sub

Private Function mAddMsfIfRequired(llRow As Long) As Long
    Dim llMsfCode As Long
    Dim llMsfIndex As Long
    Dim llMsf As Long
    Dim ilRet As Integer

    llMsfCode = Val(grdMultiMedia.TextMatrix(llRow, MSFCODEINDEX))
    For llMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
        If (llMsfCode = tgMsfCntr(llMsf).MsfRec.lCode) And (imVefCode = tgMsfCntr(llMsf).MsfRec.iVefCode) Then
            mAddMsfIfRequired = llMsf
            Exit Function
        End If
    Next llMsf
    ReDim Preserve tgMsfCntr(0 To UBound(tgMsfCntr) + 1) As MSFLIST
    llMsfIndex = UBound(tgMsfCntr) - 1
    tgMsfCntr(llMsfIndex).iStatus = 0
    tgMsfCntr(llMsfIndex).iCxfIndex = -1
    tgMsfCntr(llMsfIndex).MsfRec.lCode = -(llMsfIndex + 1)
    tgMsfCntr(llMsfIndex).MsfRec.iIhfCode = Val(grdMultiMedia.TextMatrix(llRow, IHFCODEINDEX))
    tgMsfCntr(llMsfIndex).MsfRec.iVefCode = imVefCode
    tgMsfCntr(llMsfIndex).MsfRec.lghfcode = 0
    tgMsfCntr(llMsfIndex).MsfRec.iInvGameNo = 0
    tmIhfSrchKey0.iCode = tgMsfCntr(llMsfIndex).MsfRec.iIhfCode
    ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tgMsfCntr(llMsfIndex).MsfRec.lghfcode = tmIhf.lghfcode
    End If
    tgMsfCntr(llMsfIndex).MsfRec.lChfCode = 0
    tgMsfCntr(llMsfIndex).MsfRec.lCxfCode = 0
    tgMsfCntr(llMsfIndex).MsfRec.sUnused = ""
    grdMultiMedia.TextMatrix(llRow, MSFCODEINDEX) = tgMsfCntr(llMsfIndex).MsfRec.lCode
    tgMsfCntr(llMsfIndex).iFirstMgf = -1
    mAddMsfIfRequired = llMsfIndex
    Exit Function

End Function

Private Sub mSetRatesForGames(llRow As Long)
    Dim ilRet As Integer
    Dim ilGame As Integer
    Dim ilIsf As Integer

    imIhfCode = Val(grdMultiMedia.TextMatrix(llRow, IHFCODEINDEX))
    ilRet = mIhfReadRec()
    For ilGame = 0 To UBound(tgGetGameReturn) - 1 Step 1
        For ilIsf = 0 To UBound(tmIsf) - 1 Step 1
            If tgGetGameReturn(ilGame).iGameNo = tmIsf(ilIsf).iGameNo Then
                tgGetGameReturn(ilGame).lRate = tmIsf(ilIsf).lRate
                tgGetGameReturn(ilGame).lCost = tmIsf(ilIsf).lCost
                tgGetGameReturn(ilGame).lIsfCode = tmIsf(ilIsf).lCode
                Exit For
            End If
        Next ilIsf
    Next ilGame
End Sub

Private Function mIsfReadRec(ilIhfCode As Integer) As Integer
    Dim ilRet As Integer
    Dim ilUpper As Integer
    ReDim tmIsf(0 To 0) As ISF
    ilUpper = 0
    tmIsfSrchKey3.iIhfCode = ilIhfCode
    tmIsfSrchKey3.iGameNo = 0
    ilRet = btrGetGreaterOrEqual(hmIsf, tmIsf(ilUpper), imIsfRecLen, tmIsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (ilIhfCode = tmIsf(ilUpper).iIhfCode)
        ReDim Preserve tmIsf(0 To UBound(tmIsf) + 1) As ISF
        ilUpper = UBound(tmIsf)
        ilRet = btrGetNext(hmIsf, tmIsf(ilUpper), imIsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
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
    If (Asc(tgSpf.sUsingFeatures) And MULTIMEDIA) <> MULTIMEDIA Then
        imNTRMnfCode = 0
        mAddMultiMediaNTR = 0
        Exit Function
    End If
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

Private Sub mSetTotal()
    Dim llMsf As Long
    Dim llMgf As Long
    Dim llTotal As Long
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llVehTotal As Long
    
    On Error GoTo mSetTotalErr:
    lacTotals.Caption = ""

    llTotal = 0
    llVehTotal = 0
    For llMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
        If tgMsfCntr(llMsf).iStatus >= 0 Then
            tmIhfSrchKey0.iCode = tgMsfCntr(llMsf).MsfRec.iIhfCode
            ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                llMgf = tgMsfCntr(llMsf).iFirstMgf
                Do While llMgf <> -1
                    If tmIhf.sGameIndependent = "Y" Then
                        llTotal = llTotal + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate
                    Else
                        llTotal = llTotal + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate
                    End If
                    If (imVefCode = tgMsfCntr(llMsf).MsfRec.iVefCode) Then
                        llVehTotal = llVehTotal + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate
                    End If
                    llMgf = tgMgfCntr(llMgf).iNextMgf
                Loop
            End If
        End If
    Next llMsf

    If llTotal <> 0 Then
        lacTotals.Caption = " Multi-Media Total: " & gLongToStrDec(llTotal, 2)
        lacTotals.Visible = True
    Else
        lacTotals.Visible = False
    End If
    For llRow = grdSelect.FixedRows To grdSelect.Rows - 1 Step 1
        If grdSelect.TextMatrix(llRow, SELECTEDINDEX) = "Y" Then
            If imVefCode = Val(grdSelect.TextMatrix(llRow, VEFCODEINDEX)) Then
                grdSelect.TextMatrix(llRow, GAMEDOLLARSINDEX) = gLongToStrDec(llVehTotal, 2)
            End If
            Exit For
        End If
    Next llRow
    Exit Sub
mSetTotalErr:
    Exit Sub
End Sub



Private Sub mSetAvgs(llRow As Long)
    Dim llMgf As Long
    Dim llMsf As Long
    Dim ilNoGames As Integer
    Dim ilNoUnits As Integer
    Dim llRate As Long
    Dim llCost As Long
    Dim ilUnitsVary As Integer
    Dim ilUnits As Integer

    ilNoGames = 0
    ilNoUnits = 0
    llRate = 0
    llCost = 0
    ilUnits = -1
    ilUnitsVary = False
    llMsf = mAddMsfIfRequired(llRow)
    llMgf = tgMsfCntr(llMsf).iFirstMgf
    Do While llMgf <> -1
        ilNoGames = ilNoGames + 1
        If ilUnits = -1 Then
            ilUnits = tgMgfCntr(llMgf).MgfRec.iNoUnits
        Else
            If ilUnits <> tgMgfCntr(llMgf).MgfRec.iNoUnits Then
                ilUnitsVary = True
            End If
        End If
        ilNoUnits = ilNoUnits + tgMgfCntr(llMgf).MgfRec.iNoUnits
        llCost = llCost + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lCost
        llRate = llRate + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate
        llMgf = tgMgfCntr(llMgf).iNextMgf
    Loop
    If (ilNoGames > 0) And (grdMultiMedia.TextMatrix(llRow, INDEPENDENTINDEX) = "N") Then
        If Not ilUnitsVary Then
            grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = ilNoUnits / ilNoGames
        Else
            grdMultiMedia.TextMatrix(llRow, UNITSINDEX) = gIntToStrDec((10 * ilNoUnits) / ilNoGames, 1)
        End If
        If ilNoUnits > 0 Then
            grdMultiMedia.TextMatrix(llRow, AVGCOSTINDEX) = gLongToStrDec(llCost / ilNoUnits, 2)
            grdMultiMedia.TextMatrix(llRow, AVGRATEINDEX) = gLongToStrDec(llRate / ilNoUnits, 2)
            grdMultiMedia.TextMatrix(llRow, TOTALRATEINDEX) = gLongToStrDec(llRate, 2)
        End If
    ElseIf (ilNoGames > 0) And (grdMultiMedia.TextMatrix(llRow, INDEPENDENTINDEX) = "Y") Then
        If ilNoUnits > 0 Then
            grdMultiMedia.TextMatrix(llRow, AVGCOSTINDEX) = gLongToStrDec(llCost / ilNoUnits, 2)
            grdMultiMedia.TextMatrix(llRow, AVGRATEINDEX) = gLongToStrDec(llRate / ilNoUnits, 2)
            grdMultiMedia.TextMatrix(llRow, TOTALRATEINDEX) = gLongToStrDec(llRate, 2)
        End If
    End If

End Sub

Private Sub mGridSelectColumns()
    Dim ilCol As Integer
    
    grdSelect.ColWidth(VEFCODEINDEX) = 0
    grdSelect.ColWidth(SEASONGHFCODEINDEX) = 0
    grdSelect.ColWidth(SELECTSORTINDEX) = 0
    grdSelect.ColWidth(SELECTEDINDEX) = 0
    'grdSelect.ColWidth(VEHICLEINDEX) = grdSelect.Width * 0.07
    grdSelect.ColWidth(GAMEDOLLARSINDEX) = grdSelect.Width * 0.25
    grdSelect.ColWidth(SEASONINDEX) = grdSelect.Width * 0.25
    
    grdSelect.ColWidth(VEHICLEINDEX) = grdSelect.Width - GRIDSCROLLWIDTH - 15
    For ilCol = VEHICLEINDEX To GAMEDOLLARSINDEX Step 1
        If ilCol <> VEHICLEINDEX Then
            grdSelect.ColWidth(VEHICLEINDEX) = grdSelect.ColWidth(VEHICLEINDEX) - grdSelect.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdSelect
End Sub
Private Sub mGridSelectTitles()
    'Set column titles
    grdSelect.TextMatrix(0, VEHICLEINDEX) = "Vehicle"
    grdSelect.TextMatrix(0, SEASONINDEX) = "Season"
    grdSelect.TextMatrix(0, GAMEDOLLARSINDEX) = "Dollars"
End Sub

Private Sub mPaintSelect(llRow As Long)
    Dim llCol As Long
    
    grdSelect.Row = llRow
    For llCol = VEHICLEINDEX To GAMEDOLLARSINDEX Step 1
        grdSelect.Col = llCol
        If grdSelect.TextMatrix(llRow, SELECTEDINDEX) <> "Y" Then
            grdSelect.CellBackColor = vbWhite
        Else
            grdSelect.CellBackColor = GRAY    'vbBlue
        End If
    Next llCol

End Sub


Private Sub mGameVehicleSelected(llRow As Long)
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim slCode As String

    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    'If imVefCode <= 0 Then
    '    Exit Sub
    'End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    gSetMousePointer grdMultiMedia, grdSelect, vbHourglass
    mClearCtrlFields

    slCode = grdSelect.TextMatrix(llRow, VEFCODEINDEX)
    imVefCode = Val(slCode)
    lmSeasonGhfCode = Val(grdSelect.TextMatrix(llRow, SEASONGHFCODEINDEX))
    imVpfIndex = gBinarySearchVpf(imVefCode)    'gVpfFind(CGameInv, imVefCode)
    If imVpfIndex = -1 Then
        imChgMode = False
        Screen.MousePointer = vbDefault
        gSetMousePointer grdMultiMedia, grdSelect, vbDefault
        Exit Sub
    End If
    gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), lmLLD
    mBuildSoldInv True
    ilRet = mGhfGsfReadRec()
    mInvTypePop
    MultiMediaVefCode = imVefCode
    MultiMediaSeasonGhfCode = lmSeasonGhfCode
    mTypeItemPop
    Screen.MousePointer = vbDefault
    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
    imChgMode = False
    imBypassSetting = False
    Exit Sub
cbcGameVehErr: 'VBC NR
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    gSetMousePointer grdMultiMedia, grdSelect, vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub

Private Sub mTotalByVehicle(ilVefCode As Integer, llGameTotal As Long, llNonGameTotal As Long)
    Dim llMsf As Long
    Dim llMgf As Long
    
    llGameTotal = 0
    llNonGameTotal = 0
    For llMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
        If (ilVefCode = tgMsfCntr(llMsf).MsfRec.iVefCode) Then
            llMgf = tgMsfCntr(llMsf).iFirstMgf
            If llMgf <> -1 Then
                If tgMgfCntr(llMgf).MgfRec.iGameNo > 0 Then
                    Do While llMgf <> -1
                        If tgMgfCntr(llMgf).MgfRec.iGameNo > 0 Then
                            llGameTotal = llGameTotal + tgMgfCntr(llMgf).MgfRec.iNoUnits * tgMgfCntr(llMgf).MgfRec.lRate
                        End If
                        llMgf = tgMgfCntr(llMgf).iNextMgf
                    Loop
                Else
                    llNonGameTotal = llNonGameTotal + tgMgfCntr(llMgf).MgfRec.lRate * tgMgfCntr(llMgf).MgfRec.iNoUnits
                End If
            End If
        End If
    Next llMsf

End Sub

Private Sub mClearSelectGrid()
    Dim llRow As Long
    Dim llCol As Long

    'Blank rows within grid
'    gGrid_Clear grdUsersLog, True
    'Set color within cells
    grdSelect.rowHeight(0) = fgBoxGridH + 15
    For llRow = grdSelect.FixedRows To grdSelect.Rows - 1 Step 1
        For llCol = VEHICLEINDEX To VEFCODEINDEX Step 1
            grdSelect.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdSelect.rowHeight(llRow) = fgBoxGridH + 15
    Next llRow
    For llRow = grdSelect.FixedRows To grdSelect.Rows - 1 Step 1
        mPaintSelect llRow
    Next llRow
End Sub

Private Sub mSelectSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer
    
    
    For llRow = grdSelect.FixedRows To grdSelect.Rows - 1 Step 1
        slStr = Trim$(grdSelect.TextMatrix(llRow, VEHICLEINDEX))
        If slStr <> "" Then
            If ilCol = GAMEDOLLARSINDEX Then
                slStr = grdSelect.TextMatrix(llRow, GAMEDOLLARSINDEX)
                If (slStr = "") Or (gStrDecToLong(slStr, 2) = 0) Then
                    slSort = "9999999999"
                Else
                    slSort = Trim$(str$(gStrDecToLong(slStr, 2)))
                    Do While Len(slSort) < 10
                        slSort = "0" & slSort
                    Loop
                End If
            Else
                slSort = UCase$(Trim$(grdSelect.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdSelect.TextMatrix(llRow, SELECTSORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastSelectColSorted) Or ((ilCol = imLastSelectColSorted) And (imLastSelectSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdSelect.TextMatrix(llRow, SELECTSORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdSelect.TextMatrix(llRow, SELECTSORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastSelectColSorted Then
        imLastSelectColSorted = SELECTSORTINDEX
    Else
        imLastSelectColSorted = -1
        imLastSelectSort = -1
    End If
    gGrid_SortByCol grdSelect, VEHICLEINDEX, SELECTSORTINDEX, imLastSelectColSorted, imLastSelectSort
    imLastSelectColSorted = ilCol
End Sub

