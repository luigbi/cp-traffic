VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form GameLib 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   9315
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   9315
   Begin VB.CheckBox ckcAllGames 
      Caption         =   "All Events"
      Height          =   210
      Left            =   150
      TabIndex        =   6
      Top             =   435
      Width           =   1785
   End
   Begin VB.ListBox lbcLibrary 
      Height          =   2790
      ItemData        =   "GameLib.frx":0000
      Left            =   6645
      List            =   "GameLib.frx":0002
      TabIndex        =   5
      Top             =   705
      Width           =   2490
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4800
      TabIndex        =   1
      Top             =   3720
      Width           =   945
   End
   Begin ComctlLib.ListView lbcGames 
      Height          =   2790
      Left            =   105
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   705
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   4921
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Game #"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Team"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Library"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox pbcClickFocus 
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
      Height          =   165
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   1770
      Width           =   75
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Change"
      Height          =   285
      Left            =   3495
      TabIndex        =   0
      Top             =   3720
      Width           =   945
   End
   Begin VB.Label lbcScreen 
      Caption         =   "Formats"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   45
      Width           =   1965
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   3630
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "GameLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GameLib.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Constants (Removed)                                                            *
'*  FEEDSOURCEINDEX               LANGUAGEINDEX                 AIRTIMEINDEX              *
'*  AIRVEHICLEINDEX               GAMESTATUSINDEX               TMGSFINDEX                *
'*  SORTINDEX                                                                             *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imPopReqd                     imBypassSetting               imShowHelpMsg             *
'*                                                                                        *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: GameLib.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer

Dim imVefCode As Integer

Dim tmLibName() As SORTCODE
Dim smLibNameTag As String

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassAll As Integer
Dim imFirstFocus As Integer
Dim smNowDate As String
Dim lmNowDate As Long
Dim lmFirstAllowedChgDate As Long

'Taken from GameSchd
''Const GAMENOINDEX = 2   '1
''Const VISITTEAMINDEX = 8    '4
''Const HOMETEAMINDEX = 10    '5
''Const LIBRARYINDEX = 12 '6
''Const AIRDATEINDEX = 14 '7
''Const VERLIBINDEX = 24
'
'Const GAMENOINDEX = 2   '1
'Const FEEDSOURCEINDEX = 4   '2
'Const LANGUAGEINDEX = 6 '3
'Const VISITTEAMINDEX = 8    '4
'Const HOMETEAMINDEX = 10    '5
'Const LIBRARYINDEX = 12 '6
'Const AIRDATEINDEX = 14 '7
'Const AIRTIMEINDEX = 16 '8
'Const AIRVEHICLEINDEX = 18  '9
'Const XDSPROGCODEINDEX = 20
'Const GAMESTATUSINDEX = 22  '10
'Const TMGSFINDEX = 24
'Const CHGFLAGINDEX = 25
'Const SORTINDEX = 26
'Const VERLIBINDEX = 27
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







Private Sub ckcAllGames_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRet                         llRg                                                    *
'******************************************************************************************

    Dim ilValue As Integer
    Dim ilLoop As Integer

    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilValue = False
    If ckcAllGames.Value = vbChecked Then
        ilValue = True
    End If
    'If lbcToGame.ListItems.Count > 0 Then
    '    llRg = CLng(lbcToGame.ListItems.Count - 1) * &H10000 Or 0
    '    llRet = SendMessageByNum(lbcToGame.hwnd, LB_SELITEMRANGE, ilValue, llRg)
    'End If
    For ilLoop = 0 To lbcGames.ListItems.Count - 1 Step 1
        lbcGames.ListItems(ilLoop + 1).Selected = ilValue
    Next ilLoop
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcUpdate_Click()
    Dim ilGridRow As Integer
    Dim ilLoop As Integer
    Dim slLibrary As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim ilGameNo As Integer

    If lbcLibrary.ListIndex < 0 Then
        Beep
        Exit Sub
    End If
    slLibrary = lbcLibrary.List(lbcLibrary.ListIndex)
    For ilLoop = 0 To lbcGames.ListItems.Count - 1 Step 1
        If lbcGames.ListItems(ilLoop + 1).Selected Then
            ilGameNo = Val(lbcGames.ListItems(ilLoop + 1).Text)
            For ilGridRow = GameSchd!grdDates.FixedRows To GameSchd!grdDates.Rows - 1 Step 1
                If GameSchd!grdDates.TextMatrix(ilGridRow, GAMENOINDEX) <> "" Then
                    If ilGameNo = Val(GameSchd!grdDates.TextMatrix(ilGridRow, GAMENOINDEX)) Then
                        GameSchd!grdDates.TextMatrix(ilGridRow, VERLIBINDEX) = slLibrary
                        GameSchd!grdDates.TextMatrix(ilGridRow, CHGFLAGINDEX) = "Y"
                        lbcGames.ListItems(ilLoop + 1).SubItems(2) = slLibrary
                        If GameSchd!ckcShowVersion.Value = vbChecked Then
                            GameSchd!grdDates.TextMatrix(ilGridRow, LIBRARYINDEX) = slLibrary
                        Else
                            slStr = slLibrary
                            ilPos = InStr(1, slStr, "/", vbTextCompare)
                            If ilPos > 0 Then
                                slStr = Mid$(slStr, ilPos + 1)
                            End If
                            GameSchd!grdDates.TextMatrix(ilGridRow, LIBRARYINDEX) = slStr
                        End If
                        Exit For
                    End If
                End If
            Next ilGridRow
        End If
    Next ilLoop
    cmcCancel.Caption = "&Done"
    igGameLibReturn = True
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    GameLib.Refresh
    Me.KeyPreview = True
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilReSet                                                                               *
'******************************************************************************************


    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
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
        mTerminate
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    'Reset used instead of Close to cause # Clients on network to be decrement
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    'btrStopAppl
    'End
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
'*  ilRet                                                                                 *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mInitErr                                                                              *
'******************************************************************************************

'
'   mInit
'   Where:
'
    imFirstActivate = True
    imTerminate = False

    Screen.MousePointer = vbHourglass
    imVefCode = igGameSchdVefCode
    igGameLibReturn = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmFirstAllowedChgDate = lmNowDate + 1

    mInitBox
    'mParseCmmdLine
    GameLib.Height = cmcUpdate.Top + 5 * cmcUpdate.Height / 3
    gCenterStdAlone GameLib
    'GameLib.Show
    Screen.MousePointer = vbHourglass
    imFirstFocus = True
    imBypassAll = False
    mLibraryPop
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
'    gCenterModalForm GameLib
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr: 'VBC NR
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

'
'   mTerminate
'   Where:
'

    Erase tmLibName

    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload GameLib
    Set GameLib = Nothing   'Remove data segment
    igManUnload = NO
End Sub



Private Sub lbcGames_Click()
    imBypassAll = True
    ckcAllGames.Value = vbUnchecked
    imBypassAll = False
End Sub

Private Sub pbcClickFocus_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCode                                                                                *
'******************************************************************************************

    If imFirstFocus Then
        imFirstFocus = False
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub mPopulate()
    Dim ilRow As Integer
    Dim ilGameNo As Integer
    Dim slVisitName As String
    Dim slHomeName As String
    Dim slLibrary As String
    Dim slDate As String
    Dim mItem As ListItem

    For ilRow = GameSchd!grdDates.FixedRows To GameSchd!grdDates.Rows - 1 Step 1
        If GameSchd!grdDates.TextMatrix(ilRow, GAMENOINDEX) <> "" Then
            ilGameNo = Val(GameSchd!grdDates.TextMatrix(ilRow, GAMENOINDEX))
            slVisitName = GameSchd!grdDates.TextMatrix(ilRow, VISITTEAMINDEX)
            'Home Team
            slHomeName = GameSchd!grdDates.TextMatrix(ilRow, HOMETEAMINDEX)
            'Library
            slLibrary = GameSchd!grdDates.TextMatrix(ilRow, LIBRARYINDEX)
            'Date
            slDate = GameSchd!grdDates.TextMatrix(ilRow, AIRDATEINDEX)
            If gDateValue(slDate) >= lmFirstAllowedChgDate Then
                Set mItem = lbcGames.ListItems.Add()
                mItem.Text = Trim$(str$(ilGameNo))
                mItem.SubItems(1) = Trim$(Left$(slVisitName, 4)) & " @" & Trim$(Left$(slHomeName, 4))
                mItem.SubItems(2) = slLibrary
                mItem.SubItems(3) = slDate
            End If
        End If
    Next ilRow
End Sub


Private Sub mListColumnWidths()
    Dim ilCol As Integer
    Dim llWidth As Long

    lbcGames.ColumnHeaders.Item(1).Width = lbcGames.Width / 7
    lbcGames.ColumnHeaders.Item(2).Width = lbcGames.Width / 6
    lbcGames.ColumnHeaders.Item(4).Width = lbcGames.Width / 6
    For ilCol = 1 To 4 Step 1
        If ilCol <> 3 Then
            llWidth = llWidth + lbcGames.ColumnHeaders.Item(ilCol).Width
        End If
    Next ilCol
    lbcGames.ColumnHeaders.Item(3).Width = lbcGames.Width - llWidth - GRIDSCROLLWIDTH - 5 * 240
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
    Dim llRet As Long

    'flTextHeight = pbcDates.TextHeight("1") - 35
    mListColumnWidths
    llRet = SendMessageByNum(lbcGames.hwnd, LV_SETEXTENDEDLISTVIEWSTYLE, 0, LV_FULLROWSSELECT + LV_GRIDLINES)
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
    ilRet = gPopProgLibBox(GameLib, ilVer, slType, imVefCode, lbcLibrary, tmLibName(), smLibNameTag)
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
