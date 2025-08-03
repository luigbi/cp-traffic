VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form GetGames 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "GetGames.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frcSelect 
      Caption         =   "Selection"
      Height          =   480
      Left            =   165
      TabIndex        =   12
      Top             =   4425
      Width           =   3030
      Begin VB.OptionButton rbcSelect 
         Caption         =   "Clear"
         Height          =   255
         Index           =   2
         Left            =   2205
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   195
         Width           =   720
      End
      Begin VB.OptionButton rbcSelect 
         Caption         =   "From Air Time"
         Height          =   255
         Index           =   1
         Left            =   795
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   195
         Width           =   1275
      End
      Begin VB.OptionButton rbcSelect 
         Caption         =   "All"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   195
         Width           =   585
      End
   End
   Begin VB.TextBox edcUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   -30
      ScaleHeight     =   45
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   285
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   4860
      Width           =   45
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
      ItemData        =   "GetGames.frx":08CA
      Left            =   6465
      List            =   "GetGames.frx":08CC
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3525
      Visible         =   0   'False
      Width           =   2685
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
      ItemData        =   "GetGames.frx":08CE
      Left            =   6465
      List            =   "GetGames.frx":08D0
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4065
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      TabIndex        =   8
      Top             =   4530
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   60
      Picture         =   "GetGames.frx":08D2
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcGameFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   0
      Width           =   60
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   4515
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdGame 
      Height          =   3810
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   17
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
      _Band(0).Cols   =   17
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      Left            =   285
      TabIndex        =   0
      Top             =   30
      Width           =   6930
   End
End
Attribute VB_Name = "GetGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GetGames.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  lmRowSelected                 tmGhfSrchKey0                 tmGsfSrchKey0             *
'*  tmIhfSrchKey1                 tmIhfSrchKey2                 tmIsfSrchKey0             *
'*  tmIsfSrchKey1                 tmIsfSrchKey2                                           *
'******************************************************************************************

'******************************************************
'*  GetGames - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imCtrlKey As Integer
Private imShiftKey As Integer
Private imTerminate As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long
Private lmEnableRow As Long
Private lmEnableCol As Long
Private imSetCtrlVisible As Long
Private lmFirstAllowedChgDate As Long
Private imAvailColorLevel As Integer    'set in mInit as 90%
Private smNowDate As String
Private lmNowDate As Long

Private imLastGameColSorted As Integer
Private imLastGameSort As Integer

'Private rst_Gsf As ADODB.Recordset
'Private rst_Ast As ADODB.Recordset

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
Dim imIhfRecLen As Integer        'IHF record length

Dim hmIsf As Integer
Dim tmIsf() As ISF        'ISF record image
Dim tmIsfSrchKey3 As ISFKEY3    'ISF key record image
Dim imIsfRecLen As Integer        'ISF record length

Dim tmTeamCode() As SORTCODE
Dim smTeamCodeTag As String

Dim tmLanguageCode() As SORTCODE
Dim smLanguageCodeTag As String

'6/9/14
Dim smEventTitle1 As String
Dim smEventTitle2 As String

'Grid Controls

Const GAMENOINDEX = 0
Const FEEDSOURCEINDEX = 1
Const LANGUAGEINDEX = 2
Const VISITTEAMINDEX = 3
Const HOMETEAMINDEX = 4
Const AIRDATEINDEX = 5
Const AIRTIMEINDEX = 6
Const AVAILSORDEREDINDEX = 7
Const AVAILSPROPOSALINDEX = 8
Const UNITSINDEX = 9
Const ISFCODEINDEX = 10
Const RATEINDEX = 11
Const COSTINDEX = 12
Const SORTINDEX = 13
Const SELECTEDINDEX = 14
Const BILLEDINDEX = 15
Const INVUNITSINDEX = 16






Private Sub mClearGrid()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llCol                                                                                 *
'******************************************************************************************

    Dim llRow As Long

    'Blank rows within grid
'    gGrid_Clear grdGame, True
    'Set color within cells
    grdGame.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
        'For llCol = 0 To SORTINDEX Step 1
        '    grdGame.Row = llRow
        '    grdGame.Col = llCol
        '    grdGame.CellBackColor = LIGHTYELLOW
        'Next llCol
        grdGame.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub

Private Sub cmcCancel_Click()
    igGetGameReturn = False
    mTerminate
End Sub

Private Sub cmcDone_Click()
    Dim llRow As Long
    Dim slStr As String

    igGetGameReturn = True
    ReDim tgGetGameReturn(0 To 0) As GETGAMERETURN
    For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
        slStr = Trim$(grdGame.TextMatrix(llRow, GAMENOINDEX))
        If (slStr <> "") And (Val(slStr) > 0) Then
            If grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                If Val(Trim$(grdGame.TextMatrix(llRow, UNITSINDEX))) > 0 Then
                    tgGetGameReturn(UBound(tgGetGameReturn)).iGameNo = Val(slStr)
                    tgGetGameReturn(UBound(tgGetGameReturn)).iNoUnits = Val(Trim$(grdGame.TextMatrix(llRow, UNITSINDEX)))
                    tgGetGameReturn(UBound(tgGetGameReturn)).lRate = Val(Trim$(grdGame.TextMatrix(llRow, RATEINDEX)))
                    tgGetGameReturn(UBound(tgGetGameReturn)).lCost = Val(Trim$(grdGame.TextMatrix(llRow, COSTINDEX)))
                    tgGetGameReturn(UBound(tgGetGameReturn)).lIsfCode = Val(Trim$(grdGame.TextMatrix(llRow, ISFCODEINDEX)))
                    tgGetGameReturn(UBound(tgGetGameReturn)).sBilled = grdGame.TextMatrix(llRow, BILLEDINDEX)
                    ReDim Preserve tgGetGameReturn(0 To UBound(tgGetGameReturn) + 1) As GETGAMERETURN
                End If
            End If
        End If
    Next llRow
    mTerminate
End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub edcUnits_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcUnits_KeyPress(KeyAscii As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                                                                                 *
'******************************************************************************************

    Dim slStr As String
    Dim ilKey As Integer

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
            slStr = edcUnits.Text
            slStr = Left$(slStr, edcUnits.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcUnits.SelStart - edcUnits.SelLength)
            If gCompNumberStr(slStr, "9999") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub Form_Activate()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                                                                                 *
'******************************************************************************************


    If imFirstTime Then
        imFirstTime = False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (CLng(90) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(90) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    gCenterStdAlone GetGames
    DoEvents
    mSetControls
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    mInit
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tmGsf

    smLanguageCodeTag = ""
    Erase tmLanguageCode
    smTeamCodeTag = ""
    Erase tmTeamCode

    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmIhf)
    btrDestroy hmIhf
    ilRet = btrClose(hmIsf)
    btrDestroy hmIsf
    Set GetGames = Nothing
End Sub





Private Sub frcSelect_Click()
    mSetShow
End Sub

Private Sub grdGame_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
    If (Shift And SHIFTMASK) > 0 Then
        imShiftKey = True
    Else
        imShiftKey = False
    End If
End Sub

Private Sub grdGame_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
    imShiftKey = False
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim llVeh As Long

    llVeh = gBinarySearchVef(CLng(igGetGameVefCode))
    If llVeh <> -1 Then
        plcScreen.Caption = "Event Selection-" & Trim$(tgMVef(llVeh).sName)
    Else
        plcScreen.Caption = "Event Selection"
    End If
    
    gGetEventTitles igGetGameVefCode, smEventTitle1, smEventTitle2
    
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    imAvailColorLevel = 90
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmLastClickedRow = -1
    lmScrollTop = grdGame.FixedRows
    imLastGameColSorted = -1
    imLastGameSort = -1

    hmGhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetGames
    On Error GoTo 0
    imGhfRecLen = Len(tmGhf)  'Get and save ARF record length

    hmGsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetGames
    On Error GoTo 0
    ReDim tmGsf(0 To 0) As GSF
    imGsfRecLen = Len(tmGsf(0))  'Get and save ARF record length

    hmIhf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmIhf, "", sgDBPath & "Ihf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetGames
    On Error GoTo 0
    imIhfRecLen = Len(tmIhf)  'Get and save ARF record length

    hmIsf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmIsf, "", sgDBPath & "Isf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", GetGames
    On Error GoTo 0
    ReDim tmIsf(0 To 0) As ISF
    imIsfRecLen = Len(tmIsf(0))  'Get and save ARF record length

    'gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), lmFirstAllowedChgDate
    lmFirstAllowedChgDate = lmNowDate

    mTeamPop
    mLanguagePop


    mClearGrid
    ilRet = mReadRec()
    mPopulate

    Screen.MousePointer = vbDefault
    gSetMousePointer grdGame, grdGame, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdGame, grdGame, vbDefault
    Exit Sub

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilGsf As Integer
    Dim ilGsfIndex As Integer
    Dim ilGame As Integer
    Dim llUnitTotal As Long
    Dim llTypeUnitsOrdered As Long
    Dim llTypeUnitsProp As Long
    Dim ilCol As Integer

    On Error GoTo ErrHand:

    grdGame.Redraw = False
    grdGame.Row = 0
    For llCol = GAMENOINDEX To AVAILSPROPOSALINDEX Step 1
        grdGame.Col = llCol
        grdGame.CellBackColor = vbBlack
        grdGame.CellBackColor = LIGHTBLUE
    Next llCol
    grdGame.RowHeight(0) = fgBoxGridH + 15
    grdGame.Col = UNITSINDEX
    grdGame.CellBackColor = vbBlack
    grdGame.CellBackColor = vbWhite
    llRow = grdGame.FixedRows
    For ilLoop = 0 To UBound(tmIsf) - 1 Step 1
        If (tmIsf(ilLoop).iGameNo > 0) And (tmIsf(ilLoop).iNoUnits > 0) Then
            If llRow >= grdGame.Rows Then
                grdGame.AddItem ""
            End If
            grdGame.RowHeight(llRow) = fgBoxGridH + 15
            grdGame.TextMatrix(llRow, SELECTEDINDEX) = "0"
            grdGame.TextMatrix(llRow, UNITSINDEX) = ""
            grdGame.TextMatrix(llRow, RATEINDEX) = tmIsf(ilLoop).lRate
            grdGame.TextMatrix(llRow, COSTINDEX) = tmIsf(ilLoop).lCost
            grdGame.TextMatrix(llRow, ISFCODEINDEX) = tmIsf(ilLoop).lCode
            grdGame.TextMatrix(llRow, INVUNITSINDEX) = tmIsf(ilLoop).iNoUnits
            grdGame.TextMatrix(llRow, BILLEDINDEX) = "N"
            grdGame.TextMatrix(llRow, AVAILSORDEREDINDEX) = tmIsf(ilLoop).iNoUnits
            grdGame.TextMatrix(llRow, AVAILSPROPOSALINDEX) = tmIsf(ilLoop).iNoUnits
            For ilGame = 0 To UBound(tgGetGameReturn) - 1 Step 1
                If tgGetGameReturn(ilGame).iGameNo = tmIsf(ilLoop).iGameNo Then
                    grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1"
                    grdGame.TextMatrix(llRow, UNITSINDEX) = tgGetGameReturn(ilGame).iNoUnits
                    grdGame.TextMatrix(llRow, RATEINDEX) = tgGetGameReturn(ilGame).lRate
                    grdGame.TextMatrix(llRow, COSTINDEX) = tgGetGameReturn(ilGame).lCost
                    grdGame.TextMatrix(llRow, BILLEDINDEX) = tgGetGameReturn(ilGame).sBilled
                    llUnitTotal = tmIsf(ilLoop).iNoUnits
                    llTypeUnitsOrdered = tgGetGameReturn(ilGame).iNoUnitsOrdered
                    grdGame.Row = llRow
                    grdGame.Col = AVAILSORDEREDINDEX
                    If llUnitTotal < llTypeUnitsOrdered Then
                        grdGame.CellForeColor = vbMagenta
                    ElseIf (llUnitTotal * imAvailColorLevel) \ 100 < llTypeUnitsOrdered Then
                        grdGame.CellForeColor = DARKYELLOW
                    Else
                        grdGame.CellForeColor = vbBlack
                    End If
                    grdGame.TextMatrix(llRow, AVAILSORDEREDINDEX) = llUnitTotal - llTypeUnitsOrdered
                    grdGame.Row = llRow
                    grdGame.Col = AVAILSPROPOSALINDEX
                    llTypeUnitsProp = tgGetGameReturn(ilGame).iNoUnitsProp
                    If llUnitTotal < llTypeUnitsProp Then
                        grdGame.CellForeColor = vbMagenta
                    ElseIf (llUnitTotal * imAvailColorLevel) \ 100 < llTypeUnitsProp Then
                        grdGame.CellForeColor = DARKYELLOW
                    Else
                        grdGame.CellForeColor = vbBlack
                    End If
                    grdGame.TextMatrix(llRow, AVAILSPROPOSALINDEX) = llUnitTotal - llTypeUnitsProp
                    Exit For
                End If
            Next ilGame
            mPaintRowColor llRow
            'Game Number
            grdGame.TextMatrix(llRow, GAMENOINDEX) = tmIsf(ilLoop).iGameNo
            ilGsfIndex = -1
            For ilGsf = 0 To UBound(tmGsf) - 1 Step 1
                If tmGsf(ilGsf).iGameNo = tmIsf(ilLoop).iGameNo Then
                    ilGsfIndex = ilGsf
                    Exit For
                End If
            Next ilGsf
            'Feed Source
            If ilGsfIndex >= 0 Then
                If ((Asc(tgSpf.sSportInfo) And USINGFEED) = USINGFEED) Then
                    If tmGsf(ilGsfIndex).sFeedSource = "V" Then
                        grdGame.TextMatrix(llRow, FEEDSOURCEINDEX) = smEventTitle1  '"Visting"
                    ElseIf tmGsf(ilGsfIndex).sFeedSource = "N" Then
                        grdGame.TextMatrix(llRow, FEEDSOURCEINDEX) = "National"
                    Else
                        grdGame.TextMatrix(llRow, FEEDSOURCEINDEX) = smEventTitle2   '"Home"
                    End If
                End If
                'Language
                If ((Asc(tgSpf.sSportInfo) And USINGLANG) = USINGLANG) Then
                    For ilLang = 0 To UBound(tmLanguageCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                        slNameCode = tmLanguageCode(ilLang).sKey 'Traffic!lbcAgency.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If tmGsf(ilGsfIndex).iLangMnfCode = Val(slCode) Then
                            ilRet = gParseItem(slNameCode, 1, "\", slStr)
                            grdGame.TextMatrix(llRow, LANGUAGEINDEX) = Trim$(slStr)
                            Exit For
                        End If
                    Next ilLang
                End If
                'Visiting Team
                For ilTeam = 0 To UBound(tmTeamCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                    slNameCode = tmTeamCode(ilTeam).sKey 'Traffic!lbcAgency.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If tmGsf(ilGsfIndex).iVisitMnfCode = Val(slCode) Then
                        ilRet = gParseItem(slNameCode, 1, "\", slStr)
                        grdGame.TextMatrix(llRow, VISITTEAMINDEX) = Trim$(slStr)
                        Exit For
                    End If
                Next ilTeam
                'Home Team
                For ilTeam = 0 To UBound(tmTeamCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                    slNameCode = tmTeamCode(ilTeam).sKey 'Traffic!lbcAgency.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If tmGsf(ilGsfIndex).iHomeMnfCode = Val(slCode) Then
                        ilRet = gParseItem(slNameCode, 1, "\", slStr)
                        grdGame.TextMatrix(llRow, HOMETEAMINDEX) = Trim$(slStr)
                        Exit For
                    End If
                Next ilTeam
                'Air Date
                gUnpackDate tmGsf(ilGsfIndex).iAirDate(0), tmGsf(ilGsfIndex).iAirDate(1), slStr
                grdGame.TextMatrix(llRow, AIRDATEINDEX) = slStr
                If (gDateValue(slStr) > lmFirstAllowedChgDate) And (grdGame.TextMatrix(llRow, BILLEDINDEX) = "Y") Then
                    lmFirstAllowedChgDate = gDateValue(slStr)
                End If
                'Start Time
                gUnpackTime tmGsf(ilGsfIndex).iAirTime(0), tmGsf(ilGsfIndex).iAirTime(1), "A", "1", slStr
                grdGame.TextMatrix(llRow, AIRTIMEINDEX) = slStr
                If tmGsf(ilGsfIndex).sGameStatus = "C" Then
                    For ilCol = grdGame.FixedCols To grdGame.Cols - 1 Step 1
                        grdGame.Col = ilCol
                        If grdGame.CellForeColor <> vbRed Then
                            grdGame.CellForeColor = vbCyan
                        End If
                    Next ilCol
                End If
            End If

            llRow = llRow + 1
        End If
    Next ilLoop
    lmFirstAllowedChgDate = lmFirstAllowedChgDate + 1
    For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
        slStr = Trim$(grdGame.TextMatrix(llRow, GAMENOINDEX))
        If slStr <> "" Then
            slStr = grdGame.TextMatrix(llRow, AIRDATEINDEX)
            If gDateValue(slStr) < lmFirstAllowedChgDate Then
                grdGame.Row = llRow
                grdGame.Col = UNITSINDEX
                grdGame.CellBackColor = LIGHTYELLOW
                'grdGame.CellForeColor = vbBlue
            End If
        End If
    Next llRow
    'rst_Gsf.Close
    mGameSortCol AIRTIMEINDEX
    mGameSortCol AIRDATEINDEX
    grdGame.Redraw = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    'gMsg = ""
    'For Each gErrSQL In cnn.Errors
    '    If gErrSQL.NativeError <> 0 Then             'SQLSetConnectAttr vs. SQLSetOpenConnection
    '        gMsg = "A SQL error has occured in Get Game-mPopulate: "
    '        gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
    '    End If
    'Next gErrSQL
    'If (Err.Number <> 0) And (gMsg = "") Then
    '    gMsg = "A general error has occured in Get Game-mPopulate: "
    '    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    'End If
    On Error GoTo 0

End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer

    grdGame.ColWidth(ISFCODEINDEX) = 0
    grdGame.ColWidth(RATEINDEX) = 0
    grdGame.ColWidth(COSTINDEX) = 0
    grdGame.ColWidth(SORTINDEX) = 0
    grdGame.ColWidth(SELECTEDINDEX) = 0
    grdGame.ColWidth(BILLEDINDEX) = 0
    grdGame.ColWidth(INVUNITSINDEX) = 0
    grdGame.ColWidth(GAMENOINDEX) = grdGame.Width * 0.07
    If ((Asc(tgSpf.sSportInfo) And USINGFEED) = USINGFEED) Then
        grdGame.ColWidth(FEEDSOURCEINDEX) = grdGame.Width * 0.08
    Else
        grdGame.ColWidth(FEEDSOURCEINDEX) = 0
    End If
    If ((Asc(tgSpf.sSportInfo) And USINGLANG) = USINGLANG) Then
        grdGame.ColWidth(LANGUAGEINDEX) = grdGame.Width * 0.06
    Else
        grdGame.ColWidth(LANGUAGEINDEX) = 0
    End If
    grdGame.ColWidth(VISITTEAMINDEX) = grdGame.Width * 0.16
    grdGame.ColWidth(HOMETEAMINDEX) = grdGame.Width * 0.16
    grdGame.ColWidth(AIRDATEINDEX) = grdGame.Width * 0.08
    grdGame.ColWidth(AIRTIMEINDEX) = grdGame.Width * 0.12
    grdGame.ColWidth(AVAILSORDEREDINDEX) = grdGame.Width * 0.11
    grdGame.ColWidth(AVAILSPROPOSALINDEX) = grdGame.Width * 0.11
    grdGame.ColWidth(UNITSINDEX) = grdGame.Width * 0.05

    grdGame.ColWidth(VISITTEAMINDEX) = grdGame.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To UNITSINDEX Step 1
        If ilCol <> VISITTEAMINDEX Then
            grdGame.ColWidth(VISITTEAMINDEX) = grdGame.ColWidth(VISITTEAMINDEX) - grdGame.ColWidth(ilCol)
        End If
    Next ilCol
    grdGame.ColWidth(VISITTEAMINDEX) = (grdGame.ColWidth(HOMETEAMINDEX) + grdGame.ColWidth(VISITTEAMINDEX)) \ 2
    grdGame.ColWidth(HOMETEAMINDEX) = grdGame.ColWidth(VISITTEAMINDEX)
    'Align columns to left
    'gGrid_AlignAllColsLeft grdGame
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdGame.TextMatrix(0, GAMENOINDEX) = "Event #"
    grdGame.TextMatrix(0, FEEDSOURCEINDEX) = "Feed Source"
    grdGame.TextMatrix(0, LANGUAGEINDEX) = "Language"
    grdGame.TextMatrix(0, VISITTEAMINDEX) = smEventTitle1   '"Visiting Team"
    grdGame.TextMatrix(0, HOMETEAMINDEX) = smEventTitle2    '"Home Team"
    grdGame.TextMatrix(0, AIRDATEINDEX) = "Air Date"
    grdGame.TextMatrix(0, AIRTIMEINDEX) = "Start Time"
    grdGame.TextMatrix(0, AVAILSORDEREDINDEX) = "Avails Ordered"
    grdGame.TextMatrix(0, AVAILSPROPOSALINDEX) = "Avails Proposal"
    grdGame.TextMatrix(0, UNITSINDEX) = "Units"
    'Set height of grid

End Sub

Private Sub mGameSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
        slStr = Trim$(grdGame.TextMatrix(llRow, GAMENOINDEX))
        If slStr <> "" Then
            If ilCol = AIRDATEINDEX Then
                slStr = grdGame.TextMatrix(llRow, AIRDATEINDEX)
                If slStr <> "" Then
                    slSort = Trim$(str$(gDateValue(slStr)))
                Else
                    slSort = "0"
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = AIRTIMEINDEX) Then
                slStr = grdGame.TextMatrix(llRow, AIRTIMEINDEX)
                If slStr <> "" Then
                    slSort = Trim$(str$(gTimeToLong(slStr, False)))
                Else
                    slSort = "0"
                End If
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = GAMENOINDEX) Then
                slSort = Trim$(grdGame.TextMatrix(llRow, GAMENOINDEX))
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = AVAILSORDEREDINDEX) Then
                slSort = Trim$(grdGame.TextMatrix(llRow, ilCol))
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = AVAILSPROPOSALINDEX) Then
                slSort = Trim$(grdGame.TextMatrix(llRow, ilCol))
                Do While Len(slSort) < 8
                    slSort = "0" & slSort
                Loop
            Else
                slSort = UCase$(Trim$(grdGame.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdGame.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastGameColSorted) Or ((ilCol = imLastGameColSorted) And (imLastGameSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdGame.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdGame.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastGameColSorted Then
        imLastGameColSorted = SORTINDEX
    Else
        imLastGameColSorted = -1
        imLastGameSort = -1
    End If
    gGrid_SortByCol grdGame, GAMENOINDEX, SORTINDEX, imLastGameColSorted, imLastGameSort
    imLastGameColSorted = ilCol
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
    ilRet = gPopMnfPlusFieldsBox(GetGames, lbcTeam, tmTeamCode(), smTeamCodeTag, "Z")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mTeamPopErr
        gCPErrorMsg ilRet, "mTeamPop (gPopMnfPlusFieldsBox)", GetGames
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilLang                        slNameCode                *
'*  slCode                        slStr                                                   *
'******************************************************************************************

'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, lbcCtrl, sType)
'
'   mdemoPop
'   Where:
'
    Dim ilRet As Integer

    ilRet = gPopMnfPlusFieldsBox(GetGames, lbcLanguage, tmLanguageCode(), smLanguageCodeTag, "L")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mLanguagePopErr
        gCPErrorMsg ilRet, "mLanguagePop (gPopMnfPlusFieldsBox)", GetGames
        On Error GoTo 0
    End If
    'lbcLanguage.Clear
    'For ilLoop = 0 To UBound(tmGsf) - 1 Step 1
    '    For ilLang = 0 To UBound(tmLanguageCode) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
    '        slNameCode = tmLanguageCode(ilLang).sKey 'Traffic!lbcAgency.List(ilLoop)
    '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '        If tmGsf(ilLoop).iLangMnfCode = Val(slCode) Then
    '            ilRet = gParseItem(slNameCode, 1, "\", slStr)
    '            gFindMatch slStr, 0, lbcLanguage
    '            If gLastFound(lbcLanguage) < 0 Then
    '                lbcLanguage.AddItem slStr
    '                lbcLanguage.ItemData(lbcLanguage.NewIndex) = slCode
    '            End If
    '            Exit For
    '        End If
    '    Next ilLang
    'Next ilLoop
    Exit Sub
mLanguagePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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



    Screen.MousePointer = vbDefault
    gSetMousePointer grdGame, grdGame, vbDefault
    igManUnload = YES
    Unload GetGames
    igManUnload = NO
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
    'tmGhfSrchKey1.iVefCode = igGetGameVefCode
    'ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    tmGhfSrchKey0.lCode = igGetGameGhfCode
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
        mReadRec = False
        Exit Function
    End If
    ReDim tmIsf(0 To 0) As ISF
    ilUpper = 0
    tmIhfSrchKey0.iCode = igGetGameIhfCode
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

Private Sub mSetControls()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                         ilCol                                                   *
'******************************************************************************************

    Dim ilGap As Integer

    ilGap = cmcCancel.Left - (cmcDone.Left + cmcDone.Width)
    cmcDone.Top = Me.Height - cmcDone.Height - 120
    cmcCancel.Top = cmcDone.Top
    frcSelect.Top = cmcDone.Top + cmcDone.Height - frcSelect.Height
    frcSelect.Left = 180
    cmcCancel.Left = GetGames.Width / 2 + ilGap / 2
    cmcDone.Left = cmcCancel.Left - cmcDone.Width - ilGap
    'cmcDone.Left = frcSelect.Left + frcSelect.Width + ilGap
    'cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilGap
    grdGame.Move 180, 255, GetGames.Width - 360, cmcDone.Top - 255 - 120
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdGame, fgBoxGridH + 15
    'grdGame.Height = grdGame.Height + 15
    'For ilRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
    '    If grdGame.RowHeight(ilRow) > 15 Then
    '        grdGame.Col = GAMENOINDEX
    '        grdGame.Row = ilRow
    '        grdGame.CellBackColor = LIGHTYELLOW
    '    End If
    'Next ilRow

End Sub


Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long

    grdGame.Row = llRow
    For llCol = GAMENOINDEX To AVAILSPROPOSALINDEX Step 1
        grdGame.Col = llCol
        If grdGame.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            grdGame.CellBackColor = LIGHTYELLOW
            'grdGame.CellForeColor = vbBlue
        Else
            grdGame.CellBackColor = GRAY    'vbBlue
            'grdGame.CellForeColor = vbWhite
        End If
    Next llCol

End Sub

Private Sub grdGame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mSetShow
End Sub

Private Sub grdGame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdGame.RowHeight(0) Then
        grdGame.Col = grdGame.MouseCol
        mGameSortCol grdGame.Col
        grdGame.Row = 0
        grdGame.Col = ISFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdGame, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdGame.FixedRows Then
        Exit Sub
    End If
    slStr = grdGame.TextMatrix(llCurrentRow, AIRDATEINDEX)
    If gDateValue(slStr) < lmFirstAllowedChgDate Then
        Exit Sub
    End If
    If llCurrentRow >= grdGame.FixedRows Then
        If grdGame.TextMatrix(llCurrentRow, GAMENOINDEX) <> "" Then
            If llCol = UNITSINDEX Then
                grdGame.Row = llCurrentRow
                grdGame.Col = llCol
                mEnableBox
                Exit Sub
            End If
        End If
    End If
    If llCurrentRow >= grdGame.FixedRows Then
        If grdGame.TextMatrix(llCurrentRow, GAMENOINDEX) <> "" Then
            grdGame.TopRow = lmScrollTop
            llTopRow = grdGame.TopRow
            If (Shift And CTRLMASK) > 0 Then
                If grdGame.TextMatrix(grdGame.Row, SELECTEDINDEX) <> 1 Then
                    grdGame.TextMatrix(grdGame.Row, SELECTEDINDEX) = 1
                    mSetPropAvails igGetGameDefaultUnits, grdGame.Row
                    If igGetGameDefaultUnits > 0 Then
                        grdGame.TextMatrix(grdGame.Row, UNITSINDEX) = igGetGameDefaultUnits
                    End If
                Else
                    mSetPropAvails 0, grdGame.Row
                    grdGame.TextMatrix(grdGame.Row, SELECTEDINDEX) = 0
                    grdGame.TextMatrix(grdGame.Row, UNITSINDEX) = ""
                End If
                mPaintRowColor grdGame.Row
            Else
                For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
                    If grdGame.TextMatrix(llRow, GAMENOINDEX) <> "" Then
                        grdGame.TextMatrix(llRow, SELECTEDINDEX) = "0"
                        If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                            If llRow = llCurrentRow Then
                                grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1"
                            Else
                                grdGame.TextMatrix(llRow, SELECTEDINDEX) = "0"
                            End If
                        ElseIf lmLastClickedRow < llCurrentRow Then
                            If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                                grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1"
                            End If
                        Else
                            If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                                grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1"
                            End If
                        End If
                        If grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                            mSetPropAvails igGetGameDefaultUnits, llRow
                            If igGetGameDefaultUnits > 0 Then
                                grdGame.TextMatrix(llRow, UNITSINDEX) = igGetGameDefaultUnits
                            End If
                        ElseIf grdGame.TextMatrix(llRow, SELECTEDINDEX) = "0" Then
                            mSetPropAvails 0, llRow
                            grdGame.TextMatrix(llRow, UNITSINDEX) = ""
                        End If
                        mPaintRowColor llRow
                    End If
                Next llRow
                grdGame.TopRow = llTopRow
                grdGame.Row = llCurrentRow
            End If
            lmLastClickedRow = llCurrentRow
        End If
    End If

End Sub

Private Sub grdGame_Scroll()
    lmScrollTop = grdGame.TopRow
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
'*  slCode                        ilCode                        ilRet                     *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'

    If (grdGame.Row < grdGame.FixedRows) Or (grdGame.Row >= grdGame.Rows) Or (grdGame.Col < grdGame.FixedCols) Or (grdGame.Col >= grdGame.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdGame.Row
    lmEnableCol = grdGame.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdGame.Left - pbcArrow.Width - 30, grdGame.Top + grdGame.RowPos(grdGame.Row) + (grdGame.RowHeight(grdGame.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True

    Select Case grdGame.Col
        Case UNITSINDEX
            edcUnits.MaxLength = 4
            If grdGame.Text = "" Then
                If igGetGameDefaultUnits > 0 Then
                    edcUnits.Text = igGetGameDefaultUnits
                Else
                    edcUnits.Text = grdGame.Text
                End If
            Else
                edcUnits.Text = grdGame.Text
            End If

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
    Dim ilUnits As Integer

    pbcArrow.Visible = False
    If (lmEnableRow >= grdGame.FixedRows) And (lmEnableRow < grdGame.Rows) Then
        Select Case lmEnableCol
            Case UNITSINDEX
                edcUnits.Visible = False
                ilUnits = Val(edcUnits.Text)
                If ilUnits > 0 Then
                    grdGame.TextMatrix(lmEnableRow, SELECTEDINDEX) = "1"
                Else
                    grdGame.TextMatrix(lmEnableRow, SELECTEDINDEX) = "0"
                End If
                mSetPropAvails ilUnits, lmEnableRow
                If ilUnits > 0 Then
                    grdGame.TextMatrix(lmEnableRow, lmEnableCol) = edcUnits.Text
                Else
                    grdGame.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
                mPaintRowColor lmEnableRow
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imSetCtrlVisible = False
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

    If (grdGame.Row < grdGame.FixedRows) Or (grdGame.Row >= grdGame.Rows) Or (grdGame.Col < grdGame.FixedCols) Or (grdGame.Col >= grdGame.Cols - 1) Then
        Exit Sub
    End If
    imSetCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdGame.Col - 1 Step 1
        llColPos = llColPos + grdGame.ColWidth(ilCol)
    Next ilCol
    llColWidth = grdGame.ColWidth(grdGame.Col)
    ilCol = grdGame.Col
    Do While ilCol < grdGame.Cols - 1
        If (Trim$(grdGame.TextMatrix(grdGame.Row - 1, grdGame.Col)) <> "") And (Trim$(grdGame.TextMatrix(grdGame.Row - 1, grdGame.Col)) = Trim$(grdGame.TextMatrix(grdGame.Row - 1, ilCol + 1))) Then
            llColWidth = llColWidth + grdGame.ColWidth(ilCol + 1)
            ilCol = ilCol + 1
        Else
            Exit Do
        End If
    Loop
    Select Case grdGame.Col
        Case UNITSINDEX
            edcUnits.Move grdGame.Left + llColPos + 30, grdGame.Top + grdGame.RowPos(grdGame.Row) + 15, grdGame.ColWidth(grdGame.Col), grdGame.RowHeight(grdGame.Row) - 15
            edcUnits.Visible = True
            edcUnits.SetFocus
    End Select
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
End Sub

Private Sub pbcSTab_GotFocus()
    pbcClickFocus.SetFocus
End Sub

Private Sub pbcTab_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub mSetPropAvails(ilNewUnits As Integer, llRow As Long)
    Dim llSvRow As Long
    Dim llSvCol As Long
    Dim llTypeUnitsProp As Long
    Dim llUnitTotal As Long

    llSvRow = grdGame.Row
    llSvCol = grdGame.Col
    llUnitTotal = Val(grdGame.TextMatrix(llRow, UNITSINDEX))
    llTypeUnitsProp = llUnitTotal - Val(grdGame.TextMatrix(llRow, AVAILSPROPOSALINDEX))
    llTypeUnitsProp = llTypeUnitsProp + ilNewUnits - Val(grdGame.TextMatrix(llRow, UNITSINDEX))
    grdGame.Row = llRow
    grdGame.Col = AVAILSPROPOSALINDEX
    If llUnitTotal < llTypeUnitsProp Then
        grdGame.CellForeColor = vbMagenta
    ElseIf (llUnitTotal * imAvailColorLevel) \ 100 < llTypeUnitsProp Then
        grdGame.CellForeColor = DARKYELLOW
    Else
        grdGame.CellForeColor = vbBlack
    End If
    grdGame.TextMatrix(llRow, AVAILSPROPOSALINDEX) = llUnitTotal - llTypeUnitsProp
    grdGame.Row = llSvRow
    grdGame.Col = llSvCol
End Sub

Private Sub rbcSelect_Click(Index As Integer)
    Dim slStr As String
    Dim ilClf As Integer
    Dim ilCgf As Integer
    Dim ilGameNo As Integer
    Dim llRow As Long

    If Index = 0 Then   'Select All
        For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
            If grdGame.TextMatrix(llRow, GAMENOINDEX) <> "" Then
                slStr = grdGame.TextMatrix(llRow, AIRDATEINDEX)
                If (gDateValue(slStr) >= lmFirstAllowedChgDate) And (grdGame.TextMatrix(llRow, SELECTEDINDEX) = "0") Then
                    grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1"
                    If grdGame.TextMatrix(llRow, UNITSINDEX) = "" Then
                        mSetPropAvails igGetGameDefaultUnits, llRow
                        grdGame.TextMatrix(llRow, UNITSINDEX) = igGetGameDefaultUnits
                    End If
                    mPaintRowColor llRow
                End If
            End If
        Next llRow
    ElseIf Index = 1 Then   'Set same games as Air Time
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            If tgClfCntr(ilClf).ClfRec.iVefCode = igGetGameVefCode Then
                For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
                    If grdGame.TextMatrix(llRow, GAMENOINDEX) <> "" Then
                        ilGameNo = Val(grdGame.TextMatrix(llRow, GAMENOINDEX))
                        slStr = grdGame.TextMatrix(llRow, AIRDATEINDEX)
                        If (gDateValue(slStr) >= lmFirstAllowedChgDate) And (grdGame.TextMatrix(llRow, SELECTEDINDEX) = "0") Then
                            ilCgf = tgClfCntr(ilClf).iFirstCgf
                            Do While ilCgf <> -1
                                If tgCgfCntr(ilCgf).CgfRec.iGameNo = ilGameNo Then
                                    If tgCgfCntr(ilCgf).CgfRec.iNoSpots > 0 Then
                                        grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                        If grdGame.TextMatrix(llRow, UNITSINDEX) = "" Then
                                            mSetPropAvails igGetGameDefaultUnits, llRow
                                            grdGame.TextMatrix(llRow, UNITSINDEX) = igGetGameDefaultUnits
                                        End If
                                        grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                        mPaintRowColor llRow
                                    End If
                                    Exit Do
                                End If
                                ilCgf = tgCgfCntr(ilCgf).iNextCgf
                            Loop
                        End If
                    End If
                Next llRow
            End If
        Next ilClf
    ElseIf Index = 2 Then   'Clear all
        For llRow = grdGame.FixedRows To grdGame.Rows - 1 Step 1
            If grdGame.TextMatrix(llRow, GAMENOINDEX) <> "" Then
                slStr = grdGame.TextMatrix(llRow, AIRDATEINDEX)
                If (gDateValue(slStr) >= lmFirstAllowedChgDate) And (grdGame.TextMatrix(llRow, SELECTEDINDEX) = "1") Then
                    If grdGame.TextMatrix(llRow, UNITSINDEX) <> "" Then
                        mSetPropAvails 0, llRow
                        grdGame.TextMatrix(llRow, UNITSINDEX) = ""
                    End If
                    grdGame.TextMatrix(llRow, SELECTEDINDEX) = "0"
                    mPaintRowColor llRow
                End If
            End If
        Next llRow
    End If
End Sub

Private Sub rbcSelect_GotFocus(Index As Integer)
    mSetShow
End Sub
