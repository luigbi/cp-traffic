VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form RegionModel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6165
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   11550
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
   ScaleHeight     =   6165
   ScaleWidth      =   11550
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Model"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   285
      Left            =   4515
      TabIndex        =   11
      Top             =   5745
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Don't Model"
      Height          =   285
      Left            =   5955
      TabIndex        =   10
      Top             =   5745
      Width           =   1215
   End
   Begin VB.CommandButton cmcRegion 
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
      Height          =   300
      Left            =   7080
      Picture         =   "RegionModel.frx":0000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   195
   End
   Begin VB.CommandButton cmcAdvt 
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
      Height          =   300
      Left            =   2910
      Picture         =   "RegionModel.frx":00FA
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox pbcTabStop 
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
      Left            =   6510
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   45
      Width           =   75
   End
   Begin VB.Timer tmcRegion 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10290
      Top             =   3915
   End
   Begin VB.Timer tmcAdvt 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10260
      Top             =   3360
   End
   Begin VB.ListBox lbcRegion 
      Appearance      =   0  'Flat
      Height          =   4650
      ItemData        =   "RegionModel.frx":01F4
      Left            =   3255
      List            =   "RegionModel.frx":01F6
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   795
      Width           =   3825
   End
   Begin VB.ListBox lbcAdvt 
      Appearance      =   0  'Flat
      Height          =   4650
      ItemData        =   "RegionModel.frx":01F8
      Left            =   135
      List            =   "RegionModel.frx":01FA
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   795
      Width           =   2775
   End
   Begin VB.TextBox edcRegion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3255
      MaxLength       =   80
      TabIndex        =   4
      Top             =   480
      Width           =   3825
   End
   Begin VB.TextBox edcAdvt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   135
      MaxLength       =   30
      TabIndex        =   1
      Top             =   480
      Width           =   2775
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
      TabIndex        =   9
      Top             =   1770
      Width           =   75
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRegionDef 
      Height          =   4950
      Left            =   7380
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   8731
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   12632256
      ForeColorSel    =   -2147483635
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
      _Band(0).Cols   =   4
   End
   Begin VB.Label lacScreen 
      Caption         =   "Model Region from:"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   5925
   End
End
Attribute VB_Name = "RegionModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RegionModel.frm on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RegionModel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
'Region Area
Dim tmRaf As RAF            'RAF record image
Dim hmRaf As Integer        'RAF Handle
Dim imRafRecLen As Integer      'RAF record length
Dim tmRafSrchKey As LONGKEY0  'SEF key record image
Dim tmRafSrchKey1 As RAFKEY1

'Split Entity
Dim tmSef As SEF            'SEF record image
Dim tmSefSrchKey As LONGKEY0  'SEF key record image
Dim tmSefSrchKey1 As SEFKEY1  'SEF key record image
Dim hmSef As Integer        'SEF Handle
Dim imSefRecLen As Integer      'SEF record length

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imCtrlKey As Integer
Dim lmSplitRowSelected As Long
Dim imLastSplitColSorted As Integer
Dim imLastSplitSort As Integer

Private imLbcArrowSetting As Integer
Private imComboBoxIndex As Integer
Private imBSMode As Integer
Private imBypassFocus As Integer
Private imChgMode As Integer
Private imStationPop As Integer
Private imAdvtIndex As Integer
Private imRegionIndex As Integer

Const INCLEXCLINDEX = 0
Const NAMEINDEX = 1
Const CATEGORYINDEX = 2
Const SEFCODEINDEX = 3



Private Sub mDone()
    If lbcRegion.ListIndex >= 0 Then
        igSplitModelReturn = 1
        lgSplitModelCodeRaf = lbcRegion.ItemData(imRegionIndex)
    Else
        igSplitModelReturn = 0
        lgSplitModelCodeRaf = 0
    End If
    mTerminate
End Sub

Private Sub cmcAdvt_Click()
    lbcAdvt.Visible = Not lbcAdvt.Visible
End Sub

Private Sub cmcCancel_Click()
    mCancel
End Sub

Private Sub cmcDone_Click()
    mDone
End Sub

Private Sub cmcRegion_Click()
    lbcRegion.Visible = Not lbcRegion.Visible
End Sub

Private Sub edcAdvt_Change()
    tmcAdvt.Enabled = False
    tmcRegion.Enabled = False
    imLbcArrowSetting = True
    gMatchLookAhead edcAdvt, lbcAdvt, imBSMode, imComboBoxIndex
    imLbcArrowSetting = True
    imAdvtIndex = lbcAdvt.ListIndex
    lbcRegion.Clear
    edcRegion.Text = ""
    mClearGrid
    imRegionIndex = -1
    tmcAdvt.Enabled = True
    mSetCommands
End Sub

Private Sub edcAdvt_GotFocus()
    imComboBoxIndex = lbcAdvt.ListIndex
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
    If lbcAdvt.ListCount = 1 Then
        If imAdvtIndex < 0 Then
            lbcAdvt.ListIndex = 0
            edcAdvt.Text = lbcAdvt.List(0)
        End If
        edcRegion.SetFocus
    End If
End Sub

Private Sub edcAdvt_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcAdvt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcAdvt.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub edcAdvt_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        imLbcArrowSetting = True
        gProcessArrowKey Shift, KeyCode, lbcAdvt, imLbcArrowSetting
        imLbcArrowSetting = True
        edcAdvt.SelStart = 0
        edcAdvt.SelLength = Len(edcAdvt.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        imLbcArrowSetting = True
        gProcessArrowKey Shift, KeyCode, lbcAdvt, imLbcArrowSetting
        imLbcArrowSetting = True
        edcAdvt.SelStart = 0
        edcAdvt.SelLength = Len(edcAdvt.Text)
    End If

End Sub

Private Sub edcAdvt_LostFocus()
    If tmcAdvt.Enabled Then
        tmcAdvt.Enabled = False
        tmcAdvt_Timer
    End If
End Sub

Private Sub edcRegion_Change()
    tmcRegion.Enabled = False
    imLbcArrowSetting = True
    gMatchLookAhead edcRegion, lbcRegion, imBSMode, imComboBoxIndex
    imLbcArrowSetting = True
    imRegionIndex = lbcRegion.ListIndex
    tmcRegion.Enabled = True
    mSetCommands
End Sub

Private Sub edcRegion_GotFocus()
    imComboBoxIndex = lbcRegion.ListIndex
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcRegion_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcRegion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcRegion.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub edcRegion_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        imLbcArrowSetting = True
        gProcessArrowKey Shift, KeyCode, lbcRegion, imLbcArrowSetting
        imLbcArrowSetting = True
        edcRegion.SelStart = 0
        edcRegion.SelLength = Len(edcRegion.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        imLbcArrowSetting = True
        gProcessArrowKey Shift, KeyCode, lbcRegion, imLbcArrowSetting
        imLbcArrowSetting = True
        edcRegion.SelStart = 0
        edcRegion.SelLength = Len(edcRegion.Text)
    End If
End Sub

Private Sub edcRegion_LostFocus()
    If tmcRegion.Enabled Then
        tmcRegion.Enabled = False
        tmcRegion_Timer
    End If
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
    RegionModel.Refresh
    Me.KeyPreview = True
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
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
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
'*  ilLoop                        slNameCode                    slCode                    *
'*                                                                                        *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilAdf As Integer

    imFirstActivate = True
    imTerminate = False

    Screen.MousePointer = vbHourglass
    'mParseCmmdLine
    'RegionModel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone RegionModel
    'RegionModel.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    lmSplitRowSelected = -1
    imAdvtIndex = -1
    imRegionIndex = -1
    imFirstFocus = True
    imCtrlKey = False
    lmSplitRowSelected = -1
    imChgMode = False
    imLbcArrowSetting = True
    imBypassFocus = False
    imStationPop = False
    hmRaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", RegionModel
    On Error GoTo 0
    imRafRecLen = Len(tmRaf)  'Get and save CHF record length

    hmSef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSef, "", sgDBPath & "Sef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sef.Btr)", RegionModel
    On Error GoTo 0
    imSefRecLen = Len(tmSef)  'Get and save CHF record length

    mInitBox

    mPopAdvt
    mPopFormats
    mPopDMAMarkets
    mPopMSAMarkets
    mPopStates
    mPopTimeZones
    If igAdfCode > 0 Then
        ilAdf = gBinarySearchAdf(igAdfCode)
    Else
        ilAdf = -1
    End If
    If (lgSplitModelCodeRaf > 0) Then
        cmcDone.Caption = "&Ok"
        cmcCancel.Caption = "&Cancel"
        If ilAdf <> -1 Then
            lacScreen.Caption = Trim$(tgCommAdf(ilAdf).sName) & ", Region Definition:"
        Else
            lacScreen.Caption = "Region Definition:"
        End If
    Else
        cmcDone.Caption = "&Model"
        cmcCancel.Caption = "&Don't Model"
        If ilAdf <> -1 Then
            lacScreen.Caption = Trim$(tgCommAdf(ilAdf).sName) & ", Model Region From:"
        Else
            lacScreen.Caption = "Model Region From:"
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RegionModel
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    btrExtClear hmRaf   'Clear any previous extend operation
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    btrExtClear hmSef   'Clear any previous extend operation
    ilRet = btrClose(hmSef)
    btrDestroy hmSef

    Set RegionModel = Nothing   'Remove data segment
End Sub

Private Sub lbcAdvt_Click()
    imLbcArrowSetting = True
    gProcessLbcClick lbcAdvt, edcAdvt, imChgMode, imLbcArrowSetting
    imLbcArrowSetting = True
End Sub

Private Sub lbcRegion_Click()
    imLbcArrowSetting = True
    gProcessLbcClick lbcRegion, edcRegion, imChgMode, imLbcArrowSetting
    imLbcArrowSetting = True
End Sub

Private Sub pbcClickFocus_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
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
    Dim ilRow As Integer
    'flTextHeight = pbcDates.TextHeight("1") - 35

    mGridLayout
    mGridColumnWidths
    mGridColumns
    'grdRegionModel.Height = cmcDone.Top - grdRegionModel.Top - 90
    gGrid_IntegralHeight grdRegionDef, CInt(fgBoxGridH + 30) ' + 15
    gGrid_FillWithRows grdRegionDef, fgBoxGridH + 15
    grdRegionDef.Height = grdRegionDef.Height + 15
    mClearGrid
    
    pbcTabStop.Left = -pbcTabStop.Width
    pbcClickFocus.Left = -pbcClickFocus.Left
End Sub




Private Sub mGridColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdRegionDef.Row = grdRegionDef.FixedRows - 1
    grdRegionDef.Col = INCLEXCLINDEX
    grdRegionDef.CellFontBold = False
    grdRegionDef.CellFontName = "Arial"
    grdRegionDef.CellFontSize = 6.75
    grdRegionDef.CellForeColor = vbBlue
    grdRegionDef.CellBackColor = vbWhite   'LIGHTBLUE
    grdRegionDef.TextMatrix(grdRegionDef.Row, grdRegionDef.Col) = "I/E"
    grdRegionDef.Col = NAMEINDEX
    grdRegionDef.CellFontBold = False
    grdRegionDef.CellFontName = "Arial"
    grdRegionDef.CellFontSize = 6.75
    grdRegionDef.CellForeColor = vbBlue
    grdRegionDef.CellBackColor = vbWhite   'LIGHTBLUE
    grdRegionDef.TextMatrix(grdRegionDef.Row, grdRegionDef.Col) = "Name"
    grdRegionDef.Col = CATEGORYINDEX
    grdRegionDef.CellFontBold = False
    grdRegionDef.CellFontName = "Arial"
    grdRegionDef.CellFontSize = 6.75
    grdRegionDef.CellForeColor = vbBlue
    grdRegionDef.CellBackColor = vbWhite   'LIGHTBLUE
    grdRegionDef.TextMatrix(grdRegionDef.Row, grdRegionDef.Col) = "Category"

End Sub

Private Sub mGridColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdRegionDef.ColWidth(SEFCODEINDEX) = 0
    grdRegionDef.ColWidth(INCLEXCLINDEX) = 0.12 * grdRegionDef.Width
    grdRegionDef.ColWidth(NAMEINDEX) = 0.6 * grdRegionDef.Width
    grdRegionDef.ColWidth(CATEGORYINDEX) = 0.2 * grdRegionDef.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdRegionDef.Width
    For ilCol = 0 To grdRegionDef.Cols - 1 Step 1
        llWidth = llWidth + grdRegionDef.ColWidth(ilCol)
        If (grdRegionDef.ColWidth(ilCol) > 15) And (grdRegionDef.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdRegionDef.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdRegionDef.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdRegionDef.Width
            For ilCol = 0 To grdRegionDef.Cols - 1 Step 1
                If (grdRegionDef.ColWidth(ilCol) > 15) And (grdRegionDef.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdRegionDef.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdRegionDef.FixedCols To grdRegionDef.Cols - 1 Step 1
                If grdRegionDef.ColWidth(ilCol) > 15 Then
                    ilColInc = grdRegionDef.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdRegionDef.ColWidth(ilCol) = grdRegionDef.ColWidth(ilCol) + 15
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

Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdRegionDef.Rows - 1 Step 1
        grdRegionDef.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdRegionDef.Cols - 1 Step 1
        grdRegionDef.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub


Private Sub mPopAdvt()
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim blFound As Boolean
    Dim ilLoop As Integer

    imRafRecLen = Len(tmRaf)

    lbcAdvt.Clear
    If lgSplitModelCodeRaf = 0 Then
        ilRet = btrGetFirst(hmRaf, tmRaf, imRafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While (ilRet = BTRV_ERR_NONE)
            If (igIncludeDormantSplits) Or ((Not igIncludeDormantSplits) And (tmRaf.sState <> "D")) Then
                '6/25/13: Include advertiser
                'If (tmRaf.sType = "C") And (tmRaf.iAdfCode <> igAdfCode) Then
                If (tmRaf.sType = "C") Then
                    ilAdf = gBinarySearchAdf(tmRaf.iAdfCode)
                    If ilAdf <> -1 Then
                        blFound = False
                        For ilLoop = 0 To lbcAdvt.ListCount - 1 Step 1
                            If lbcAdvt.ItemData(ilLoop) = tmRaf.iAdfCode Then
                                blFound = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not blFound Then
                            lbcAdvt.AddItem Trim$(tgCommAdf(ilAdf).sName)
                            lbcAdvt.ItemData(lbcAdvt.NewIndex) = tmRaf.iAdfCode
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hmRaf, tmRaf, imRafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Else
        ilAdf = gBinarySearchAdf(igAdfCode)
        If ilAdf <> -1 Then
            lbcAdvt.AddItem Trim$(tgCommAdf(ilAdf).sName)
            lbcAdvt.ItemData(lbcAdvt.NewIndex) = igAdfCode
        End If
    End If
    If lbcAdvt.ListCount = 1 Then
        lbcAdvt.ListIndex = 0
        tmcAdvt.Enabled = False
        tmcAdvt_Timer
    End If
End Sub

Private Sub mPopRegion()

    Dim ilRet As Integer
    Dim ilAdfCode As Integer
    Dim ilLoop As Integer

    lbcRegion.Clear
    If lbcAdvt.ListIndex < 0 Then
        Exit Sub
    End If
    ilAdfCode = lbcAdvt.ItemData(lbcAdvt.ListIndex)
    tmRafSrchKey1.iAdfCode = ilAdfCode
    tmRafSrchKey1.sType = "C"
    ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmRaf.iAdfCode = ilAdfCode) And (tmRaf.sType = "C")
        If (igIncludeDormantSplits) Or ((Not igIncludeDormantSplits) And (tmRaf.sState <> "D")) Then
            lbcRegion.AddItem Trim$(tmRaf.sName)
            lbcRegion.ItemData(lbcRegion.NewIndex) = tmRaf.lCode
        End If
        ilRet = btrGetNext(hmRaf, tmRaf, imRafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If lbcRegion.ListCount = 1 Then
        lbcRegion.ListIndex = 0
        tmcRegion.Enabled = False
        tmcRegion_Timer
    Else
        If (lgSplitModelCodeRaf > 0) And (imFirstFocus) Then
            For ilLoop = 0 To lbcRegion.ListCount - 1 Step 1
                If lgSplitModelCodeRaf = lbcRegion.ItemData(ilLoop) Then
                    lbcRegion.ListIndex = ilLoop
                    tmcRegion.Enabled = False
                    tmcRegion_Timer
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    imFirstFocus = False
End Sub

Private Sub pbcTabStop_GotFocus()
    If lbcAdvt.ListIndex >= 0 Then
        mDone
    Else
        mCancel
    End If
End Sub

Private Sub tmcAdvt_Timer()
    tmcAdvt.Enabled = False
    mPopRegion
    mSetCommands
End Sub

Private Sub mPopRegionDef()
    Dim ilRet As Integer
    Dim llRafCode As Long
    Dim llRow As Long
    Dim ilCol As Integer
    Dim slCategory As String
    Dim slCategoryName As String
    Dim slInclExcl As String
    Dim slName As String
    Dim ilMkt As Integer
    Dim ilSnt As Integer
    Dim ilTzt As Integer
    Dim ilFmt As Integer
    Dim ilShtt As Integer
    
    grdRegionDef.Redraw = False
    mClearGrid
    If lbcRegion.ListIndex < 0 Then
        grdRegionDef.Redraw = True
        Exit Sub
    End If
    llRow = grdRegionDef.FixedRows
    llRafCode = lbcRegion.ItemData(lbcRegion.ListIndex)
    tmRafSrchKey.lCode = llRafCode
    ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    tmSefSrchKey1.lRafCode = llRafCode
    tmSefSrchKey1.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hmSef, tmSef, imSefRecLen, tmSefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSef.lRafCode = llRafCode)
        slCategory = Trim$(tmSef.sCategory)
        If slCategory = "" Then
            slCategory = tmRaf.sCategory
        End If
        slInclExcl = Trim$(tmSef.sInclExcl)
        If slInclExcl = "" Then
            slInclExcl = tmRaf.sInclExcl
        End If
        If slInclExcl = "E" Then
            slInclExcl = "Excl"
        Else
            slInclExcl = "Incl"
        End If
        Select Case UCase$(slCategory)
            Case "M"
                slCategoryName = "DMA Market"
                For ilMkt = LBound(tgMarkets) To UBound(tgMarkets) - 1 Step 1
                    If tmSef.iIntCode = tgMarkets(ilMkt).iCode Then
                        slName = Trim$(tgMarkets(ilMkt).sName)
                        Exit For
                    End If
                Next ilMkt
            Case "A"
                slCategoryName = "MSA Market"
                For ilMkt = LBound(tgMSAMarkets) To UBound(tgMSAMarkets) - 1 Step 1
                    If tmSef.iIntCode = tgMSAMarkets(ilMkt).iCode Then
                        slName = Trim$(tgMSAMarkets(ilMkt).sName)
                        Exit For
                    End If
                Next ilMkt
            Case "N"
                slCategoryName = "State"
                For ilSnt = LBound(tgStates) To UBound(tgStates) - 1 Step 1
                    If StrComp(Trim$(tmSef.sName), Trim$(tgStates(ilSnt).sPostalName), vbTextCompare) = 0 Then
                        slName = Trim$(tgStates(ilSnt).sPostalName) & " (" & Trim$(tgStates(ilSnt).sName) & ")"
                        Exit For
                    End If
                Next ilSnt
            Case "T"
                slCategoryName = "Time Zone"
                For ilTzt = LBound(tgTimeZones) To UBound(tgTimeZones) - 1 Step 1
                    If tmSef.iIntCode = tgTimeZones(ilTzt).iCode Then
                        Select Case Left$(Trim$(tgTimeZones(ilTzt).sCSIName), 1)
                            Case "E"
                                slName = Trim$(tgTimeZones(ilTzt).sName) & " (ETZ)"
                            Case "C"
                                slName = Trim$(tgTimeZones(ilTzt).sName) & " (CTZ)"
                            Case "M"
                                slName = Trim$(tgTimeZones(ilTzt).sName) & " (MTZ)"
                            Case "P"
                                slName = Trim$(tgTimeZones(ilTzt).sName) & " (PTZ)"
                        End Select
                        Exit For
                    End If
                Next ilTzt
            Case "F"
                slCategoryName = "Format"
                For ilFmt = LBound(tgFormats) To UBound(tgFormats) - 1 Step 1
                    If tmSef.iIntCode = tgFormats(ilFmt).iCode Then
                        slName = Trim$(tgFormats(ilFmt).sName)
                        Exit For
                    End If
                Next ilFmt
            Case "S"
                slCategoryName = "Station"
                mPopStations
                For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
                    If tmSef.iIntCode = tgStations(ilShtt).iCode Then
                        slName = Trim$(tgStations(ilShtt).sCallLetters)
                        Exit For
                    End If
                Next ilShtt
        End Select
        If llRow >= grdRegionDef.Rows Then
            grdRegionDef.AddItem ""
            grdRegionDef.RowHeight(llRow) = fgFlexGridRowH
            grdRegionDef.Row = llRow
            For ilCol = 0 To grdRegionDef.Cols - 1 Step 1
                grdRegionDef.ColAlignment(ilCol) = flexAlignLeftCenter
                grdRegionDef.Col = ilCol
                grdRegionDef.CellBackColor = LIGHTYELLOW
            Next ilCol
        End If
        grdRegionDef.Row = llRow
        grdRegionDef.TextMatrix(llRow, INCLEXCLINDEX) = slInclExcl
        grdRegionDef.TextMatrix(llRow, NAMEINDEX) = slName
        grdRegionDef.TextMatrix(llRow, CATEGORYINDEX) = slCategoryName
        grdRegionDef.TextMatrix(llRow, SEFCODEINDEX) = tmSef.lCode
        ilRet = btrGetNext(hmSef, tmSef, imSefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        llRow = llRow + 1
    Loop
    grdRegionDef.Redraw = True

End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim ilCol As Integer
    
    For llRow = grdRegionDef.FixedRows To grdRegionDef.Rows - 1 Step 1
        grdRegionDef.Row = llRow
        For ilCol = grdRegionDef.FixedCols To grdRegionDef.Cols - 1 Step 1
            grdRegionDef.Col = ilCol
            grdRegionDef.CellBackColor = LIGHTYELLOW
            grdRegionDef.TextMatrix(llRow, ilCol) = ""
        Next ilCol
    Next llRow
    
End Sub

Private Sub mPopFormats()
    Dim ilRet As Integer

    ilRet = gObtainFormats()
End Sub

Private Sub mPopDMAMarkets()
    Dim ilRet As Integer

    ilRet = gObtainMarkets()
End Sub

Private Sub mPopMSAMarkets()
    Dim ilRet As Integer
    
    ilRet = gObtainMSAMarkets()
End Sub
Private Sub mPopStates()
    Dim ilRet As Integer

    ilRet = gObtainStates()
End Sub

Private Sub mPopTimeZones()
    Dim ilRet As Integer

    ilRet = gObtainTimeZones()
End Sub

Private Sub mPopStations()
    Dim ilRet As Integer
    If imStationPop Then
        Exit Sub
    End If
    ilRet = gObtainStations()
End Sub

Private Sub tmcRegion_Timer()
    tmcRegion.Enabled = False
    mPopRegionDef
    mSetCommands
End Sub

Private Sub mSetCommands()
    If (imAdvtIndex >= 0) And (imRegionIndex >= 0) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
    End If
End Sub


Private Sub mCancel()
    igSplitModelReturn = 0
    lgSplitModelCodeRaf = 0
    mTerminate
End Sub
