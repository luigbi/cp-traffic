VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form TaxTable 
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
   Begin VB.PictureBox pbcGrossNet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2325
      ScaleHeight     =   210
      ScaleWidth      =   795
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   915
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "TaxTable.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   3480
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   345
      Width           =   30
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1245
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5730
      TabIndex        =   9
      Top             =   3720
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4005
      TabIndex        =   8
      Top             =   3720
      Width           =   945
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
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
      Left            =   450
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3600
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
      Left            =   705
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3600
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
      Left            =   1005
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   525
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
      Left            =   9045
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   3885
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2235
      TabIndex        =   7
      Top             =   3720
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTax 
      Height          =   3090
      Left            =   210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5450
      _Version        =   393216
      Rows            =   10
      Cols            =   7
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
      _Band(0).Cols   =   7
   End
   Begin VB.Label lacScreen 
      Caption         =   "Tax Table"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   30
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
Attribute VB_Name = "TaxTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of TaxTable.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: TaxTable.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim hmTrf As Integer
Dim tmTrf As TRF        'Rvf record image
Dim tmTrfSrchKey0 As INTKEY0
Dim imTrfRecLen As Integer        'RvF record length

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imLastSelectRow As Integer
Dim imCtrlKey As Integer
Dim imLastTaxColSorted As Integer
Dim imLastTaxSort As Integer
Dim lmTaxRowSelected As Long
Dim imTaxChg As Integer
Dim imIgnoreScroll As Integer
Dim imFromArrow As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer
Dim smGrossNet As String


Const GROSSNETINDEX = 0
Const TAX1NAMEINDEX = 1
Const TAX1RATEINDEX = 2
Const TAX2NAMEINDEX = 3
Const TAX2RATEINDEX = 4
Const TRFCODEINDEX = 5
Const SORTINDEX = 6


Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If imTaxChg Then
        If MsgBox("Save all changes?", vbYesNo) = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUpdate_Click()
    Dim ilRet As Integer

    ilRet = mSaveRec()
    If ilRet Then
        mMoveRecToCtrl
    End If
End Sub

Private Sub cmcUpdate_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                                                   *
'******************************************************************************************


    Select Case lmEnableCol
        Case GROSSNETINDEX
        Case TAX1NAMEINDEX
        Case TAX1RATEINDEX
        Case TAX2NAMEINDEX
        Case TAX2RATEINDEX
    End Select
    grdTax.CellForeColor = vbBlack

End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmEnableCol
        Case GROSSNETINDEX
        Case TAX1NAMEINDEX
        Case TAX1RATEINDEX
        Case TAX2NAMEINDEX
        Case TAX2RATEINDEX
    End Select
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim ilPos As Integer
    Dim slStr As String

    ilKey = KeyAscii

    Select Case lmEnableCol
        Case GROSSNETINDEX
        Case TAX1NAMEINDEX
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case TAX1RATEINDEX
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
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "100.0000") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case TAX2NAMEINDEX
            If Not gCheckKeyAscii(ilKey) Then
                KeyAscii = 0
                Exit Sub
            End If
        Case TAX2RATEINDEX
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
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "100.0000") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select

End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    TaxTable.Refresh
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
    Dim ilRet As Integer
    
    On Error Resume Next
    btrExtClear hmTrf   'Clear any previous extend operation
    ilRet = btrClose(hmTrf)
    btrDestroy hmTrf
    Set TaxTable = Nothing   'Remove data segment
End Sub


Private Sub grdTax_EnterCell()
    mSetShow
End Sub

Private Sub grdTax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdTax.TopRow
    grdTax.Redraw = False
End Sub

Private Sub grdTax_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer

    imIgnoreScroll = False
    If Y < grdTax.RowHeight(0) Then
        grdTax.Col = grdTax.MouseCol
        mTaxSortCol grdTax.Col
        Exit Sub
    End If
    pbcArrow.Visible = False
    ilCol = grdTax.MouseCol
    ilRow = grdTax.MouseRow
    If ilCol < grdTax.FixedCols Then
        grdTax.Redraw = True
        Exit Sub
    End If
    If ilRow < grdTax.FixedRows Then
        grdTax.Redraw = True
        Exit Sub
    End If
    If grdTax.TextMatrix(ilRow, TAX1NAMEINDEX) = "" Then
        grdTax.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdTax.TextMatrix(ilRow, TAX1NAMEINDEX) = ""
        grdTax.Row = ilRow + 1
        grdTax.Col = GROSSNETINDEX
        grdTax.Redraw = True
    Else
        grdTax.Row = ilRow
        grdTax.Col = ilCol
    End If
    grdTax.Redraw = True
    lmTopRow = grdTax.TopRow
    'If Not mColOk() Then
    '    pbcArrow.Move grdTax.Left - pbcArrow.Width - 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + (grdTax.RowHeight(grdTax.Row) - pbcArrow.Height) / 2
    '    pbcArrow.Visible = True
    '    Exit Sub
    'End If

    mEnableBox
End Sub

Private Sub grdTax_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdTax.Redraw = False Then
        grdTax.Redraw = True
        If lmTopRow < grdTax.FixedRows Then
            grdTax.TopRow = grdTax.FixedRows
        Else
            grdTax.TopRow = lmTopRow
        End If
        grdTax.Refresh
        grdTax.Redraw = False
    End If
    If (imCtrlVisible) And (grdTax.Row >= grdTax.FixedRows) And (grdTax.Col >= 0) And (grdTax.Col < grdTax.Cols - 1) Then
        If grdTax.RowIsVisible(grdTax.Row) Then
            pbcArrow.Move grdTax.Left - pbcArrow.Width - 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + (grdTax.RowHeight(grdTax.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            pbcSetFocus.SetFocus
            edcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
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
'*  ilLoop                        slNameCode                    slCode                    *
'*                                                                                        *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer

    imFirstActivate = True
    imTerminate = False
    imIgnoreScroll = False
    imFromArrow = False
    imCtrlVisible = False

    Screen.MousePointer = vbHourglass
    'mParseCmmdLine
    TaxTable.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone TaxTable
    'TaxTable.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    lmTaxRowSelected = -1
    imTaxChg = False

    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    imTrfRecLen = Len(tmTrf)
    hmTrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmTrf, "", sgDBPath & "Trf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Trf.Btr)", TaxTable
    On Error GoTo 0
    imTrfRecLen = Len(tmTrf)  'Get and save CHF record length

    mInitBox

    mPopulate
    If imTerminate Then
        Exit Sub
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
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload TaxTable
    igManUnload = NO
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
    If grdTax.Visible Then
        lmTaxRowSelected = -1
        grdTax.Row = 0
        grdTax.Col = TRFCODEINDEX
        mSetCommands
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

    Dim ilRet As Integer

    ilRet = gObtainTrf()
    mMoveRecToCtrl
End Sub


Private Sub mSetCommands()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                                                                                 *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  cmcEraseErr                                                                           *
'******************************************************************************************

    Dim ilRet As Integer

    If imTaxChg Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    Exit Sub
cmcEraseErr: 'VBC NR
    ilRet = 1
    Resume Next
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

    mGridTaxLayout
    mGridTaxColumnWidths
    mGridTaxColumns
    grdTax.Move 180, lacScreen.Top + lacScreen.Height + 120, grdTax.Width, cmcDone.Top - (lacScreen.Top + lacScreen.Height) - 240
    'grdTax.Height = grdTax.RowPos(0) + 14 * grdTax.RowHeight(0) + fgPanelAdj - 15
    imInitNoRows = (cmcDone.Top - 120 - grdTax.Top) \ fgFlexGridRowH
    grdTax.Height = grdTax.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj - 15
End Sub

Private Sub mGridTaxLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdTax.Rows - 1 Step 1
        grdTax.RowHeight(ilRow) = fgFlexGridRowH
    Next ilRow
    For ilCol = 0 To grdTax.Cols - 1 Step 1
        grdTax.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridTaxColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdTax.Row = grdTax.FixedRows - 1
    grdTax.Col = GROSSNETINDEX
    grdTax.CellFontBold = False
    grdTax.CellFontName = "Arial"
    grdTax.CellFontSize = 6.75
    grdTax.CellForeColor = vbBlue
    grdTax.CellBackColor = LIGHTBLUE
    grdTax.TextMatrix(grdTax.Row, grdTax.Col) = "Gross/Net"
    grdTax.Col = TAX1NAMEINDEX
    grdTax.CellFontBold = False
    grdTax.CellFontName = "Arial"
    grdTax.CellFontSize = 6.75
    grdTax.CellForeColor = vbBlue
    grdTax.CellBackColor = LIGHTBLUE
    grdTax.TextMatrix(grdTax.Row, grdTax.Col) = "Tax 1 Name"
    grdTax.Col = TAX1RATEINDEX
    grdTax.CellFontBold = False
    grdTax.CellFontName = "Arial"
    grdTax.CellFontSize = 6.75
    grdTax.CellForeColor = vbBlue
    grdTax.CellBackColor = LIGHTBLUE
    grdTax.TextMatrix(grdTax.Row, grdTax.Col) = "Tax Rate"
    grdTax.Col = TAX2NAMEINDEX
    grdTax.CellFontBold = False
    grdTax.CellFontName = "Arial"
    grdTax.CellFontSize = 6.75
    grdTax.CellForeColor = vbBlue
    grdTax.CellBackColor = LIGHTBLUE
    grdTax.TextMatrix(grdTax.Row, grdTax.Col) = "Tax 2 Name"
    grdTax.Col = TAX2RATEINDEX
    grdTax.CellFontBold = False
    grdTax.CellFontName = "Arial"
    grdTax.CellFontSize = 6.75
    grdTax.CellForeColor = vbBlue
    grdTax.CellBackColor = LIGHTBLUE
    grdTax.TextMatrix(grdTax.Row, grdTax.Col) = "Tax Rate"
    grdTax.Col = TRFCODEINDEX
    grdTax.CellFontBold = False
    grdTax.CellFontName = "Arial"
    grdTax.CellFontSize = 6.75
    grdTax.CellForeColor = vbBlue
    grdTax.CellBackColor = LIGHTBLUE
    grdTax.TextMatrix(grdTax.Row, grdTax.Col) = "Trf Code"
    grdTax.Col = SORTINDEX
    grdTax.CellFontBold = False
    grdTax.CellFontName = "Arial"
    grdTax.CellFontSize = 6.75
    grdTax.CellForeColor = vbBlue
    grdTax.CellBackColor = LIGHTBLUE
    grdTax.TextMatrix(grdTax.Row, grdTax.Col) = "Sort"

End Sub

Private Sub mGridTaxColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdTax.ColWidth(TRFCODEINDEX) = 0
    grdTax.ColWidth(SORTINDEX) = 0
    grdTax.ColWidth(GROSSNETINDEX) = 0.1 * grdTax.Width
    grdTax.ColWidth(TAX1NAMEINDEX) = 0.34 * grdTax.Width
    grdTax.ColWidth(TAX1RATEINDEX) = 0.1 * grdTax.Width
    grdTax.ColWidth(TAX2NAMEINDEX) = 0.34 * grdTax.Width
    grdTax.ColWidth(TAX2RATEINDEX) = 0.1 * grdTax.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdTax.Width
    For ilCol = 0 To grdTax.Cols - 1 Step 1
        llWidth = llWidth + grdTax.ColWidth(ilCol)
        If (grdTax.ColWidth(ilCol) > 15) And (grdTax.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdTax.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdTax.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdTax.Width
            For ilCol = 0 To grdTax.Cols - 1 Step 1
                If (grdTax.ColWidth(ilCol) > 15) And (grdTax.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdTax.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdTax.FixedCols To grdTax.Cols - 1 Step 1
                If grdTax.ColWidth(ilCol) > 15 Then
                    ilColInc = grdTax.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdTax.ColWidth(ilCol) = grdTax.ColWidth(ilCol) + 15
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


Private Sub mTaxSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdTax.FixedRows To grdTax.Rows - 1 Step 1
        slStr = Trim$(grdTax.TextMatrix(llRow, TAX1NAMEINDEX))
        If slStr <> "" Then
            If ilCol = GROSSNETINDEX Then
                slSort = grdTax.TextMatrix(llRow, GROSSNETINDEX)
                Do While Len(slSort) < 5
                    slSort = slSort & " "
                Loop
            ElseIf ilCol = TAX1NAMEINDEX Then
                slSort = grdTax.TextMatrix(llRow, TAX1NAMEINDEX)
                Do While Len(slSort) < 30
                    slSort = slSort & " "
                Loop
            ElseIf (ilCol = TAX1RATEINDEX) Then
                slSort = Trim$(str$(gStrDecToLong(grdTax.TextMatrix(llRow, TAX1RATEINDEX), 4)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = TAX2NAMEINDEX Then
                slSort = grdTax.TextMatrix(llRow, TAX2NAMEINDEX)
                Do While Len(slSort) < 30
                    slSort = slSort & " "
                Loop
            ElseIf (ilCol = TAX2RATEINDEX) Then
                slSort = Trim$(str$(gStrDecToLong(grdTax.TextMatrix(llRow, TAX2RATEINDEX), 4)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            End If
            slStr = grdTax.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastTaxColSorted) Or ((ilCol = imLastTaxColSorted) And (imLastTaxSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdTax.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdTax.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastTaxColSorted Then
        imLastTaxColSorted = SORTINDEX
    Else
        imLastTaxColSorted = -1
        imLastTaxSort = -1
    End If
    gGrid_SortByCol grdTax, TAX1NAMEINDEX, SORTINDEX, imLastTaxColSorted, imLastTaxSort
    imLastTaxColSorted = ilCol
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
    If (grdTax.Row < grdTax.FixedRows) Or (grdTax.Row >= grdTax.Rows) Or (grdTax.Col < grdTax.FixedCols) Or (grdTax.Col >= grdTax.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdTax.Row
    lmEnableCol = grdTax.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdTax.Left - pbcArrow.Width - 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + (grdTax.RowHeight(grdTax.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    Select Case grdTax.Col
        Case GROSSNETINDEX
            slStr = grdTax.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                imTaxChg = True
                If grdTax.Row > grdTax.FixedRows Then
                    slStr = grdTax.TextMatrix(grdTax.Row - 1, grdTax.Col)
                Else
                    If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                        slStr = "Net"   '"Gross"
                    Else
                        slStr = "Net"
                    End If
                End If
            End If
            smGrossNet = slStr
        Case TAX1NAMEINDEX
            edcDropDown.MaxLength = 30
            slStr = grdTax.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdTax.Row > grdTax.FixedRows Then
                    slStr = grdTax.TextMatrix(grdTax.Row - 1, grdTax.Col)
                End If
            End If
            edcDropDown.Text = slStr
        Case TAX1RATEINDEX
            edcDropDown.MaxLength = 7
            slStr = grdTax.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdTax.Row > grdTax.FixedRows Then
                    slStr = grdTax.TextMatrix(grdTax.Row - 1, grdTax.Col)
                End If
            End If
            edcDropDown.Text = slStr
        Case TAX2NAMEINDEX
            edcDropDown.MaxLength = 30
            slStr = grdTax.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            edcDropDown.Text = slStr
        Case TAX2RATEINDEX
            edcDropDown.MaxLength = 7
            slStr = grdTax.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdTax.Row > grdTax.FixedRows Then
                    slStr = grdTax.TextMatrix(grdTax.Row - 1, grdTax.Col)
                End If
            End If
            edcDropDown.Text = slStr
    End Select
    mSetFocus
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

    If (grdTax.Row < grdTax.FixedRows) Or (grdTax.Row >= grdTax.Rows) Or (grdTax.Col < grdTax.FixedCols) Or (grdTax.Col >= grdTax.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdTax.Left - pbcArrow.Width - 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + (grdTax.RowHeight(grdTax.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdTax.Col - 1 Step 1
        llColPos = llColPos + grdTax.ColWidth(ilCol)
    Next ilCol
    Select Case grdTax.Col
        Case GROSSNETINDEX
            pbcGrossNet.Move grdTax.Left + llColPos + 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + 30, grdTax.ColWidth(grdTax.Col) - 30, grdTax.RowHeight(grdTax.Row) - 15
            pbcGrossNet.Visible = True
            pbcGrossNet.SetFocus
        Case TAX1NAMEINDEX
            edcDropDown.Move grdTax.Left + llColPos + 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + 30, grdTax.ColWidth(grdTax.Col) - 30, grdTax.RowHeight(grdTax.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TAX1RATEINDEX
            edcDropDown.Move grdTax.Left + llColPos + 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + 30, grdTax.ColWidth(grdTax.Col) - 30, grdTax.RowHeight(grdTax.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TAX2NAMEINDEX
            edcDropDown.Move grdTax.Left + llColPos + 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + 30, grdTax.ColWidth(grdTax.Col) - 30, grdTax.RowHeight(grdTax.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case TAX2RATEINDEX
            edcDropDown.Move grdTax.Left + llColPos + 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + 30, grdTax.ColWidth(grdTax.Col) - 30, grdTax.RowHeight(grdTax.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus

    End Select
    mSetCommands
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
'*  ilPos                                                                                 *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String

    pbcArrow.Visible = False
    If (lmEnableRow >= grdTax.FixedRows) And (lmEnableRow < grdTax.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case GROSSNETINDEX
                pbcGrossNet.Visible = False
                slStr = smGrossNet
                If StrComp(grdTax.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imTaxChg = True
                End If
                grdTax.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case TAX1NAMEINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If StrComp(grdTax.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imTaxChg = True
                End If
                grdTax.TextMatrix(lmEnableRow, lmEnableCol) = slStr
             Case TAX1RATEINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If gStrDecToLong(grdTax.TextMatrix(lmEnableRow, lmEnableCol), 4) <> gStrDecToLong(slStr, 4) Then
                    imTaxChg = True
                End If
                grdTax.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case TAX2NAMEINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If StrComp(grdTax.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imTaxChg = True
                End If
                grdTax.TextMatrix(lmEnableRow, lmEnableCol) = slStr
             Case TAX2RATEINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If gStrDecToLong(grdTax.TextMatrix(lmEnableRow, lmEnableCol), 4) <> gStrDecToLong(slStr, 4) Then
                    imTaxChg = True
                End If
                grdTax.TextMatrix(lmEnableRow, lmEnableCol) = slStr
        End Select
    End If
    pbcArrow.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        mSetShow
        Do
            ilPrev = False
            If grdTax.Col = GROSSNETINDEX Then
                If grdTax.Row > grdTax.FixedRows Then
                    lmTopRow = -1
                    grdTax.Row = grdTax.Row - 1
                    If Not grdTax.RowIsVisible(grdTax.Row) Then
                        grdTax.TopRow = grdTax.TopRow - 1
                    End If
                    grdTax.Col = TAX2RATEINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdTax.Col = grdTax.Col - 1
                'If gColOk(grdTax, grdTax.Row, grdTax.Col) Then
                    mEnableBox
                'Else
                '    ilPrev = True
                'End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdTax.TopRow = grdTax.FixedRows
        grdTax.Col = GROSSNETINDEX
        grdTax.Row = grdTax.FixedRows
        'If gColOk(grdTax, grdTax.Row, grdTax.Col) Then
            mEnableBox
        'Else
        '    cmcCancel.SetFocus
        'End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdTax.Col = TAX2RATEINDEX Then
                llRow = grdTax.Rows
                Do
                    llRow = llRow - 1
                Loop While grdTax.TextMatrix(llRow, TAX1NAMEINDEX) = ""
                llRow = llRow + 1
                If (grdTax.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdTax.Row = grdTax.Row + 1
                    If Not grdTax.RowIsVisible(grdTax.Row) Or (grdTax.Row - (grdTax.TopRow - grdTax.FixedRows) >= imInitNoRows) Then
                        imIgnoreScroll = True
                        grdTax.TopRow = grdTax.TopRow + 1
                    End If
                    grdTax.Col = GROSSNETINDEX
                    'grdTax.TextMatrix(grdTax.Row, CODEINDEX) = 0
                    If Trim$(grdTax.TextMatrix(grdTax.Row, GROSSNETINDEX)) <> "" Then
                        'If gColOk(grdTax, grdTax.Row, grdTax.Col) Then
                            mEnableBox
                        'Else
                        '    cmcCancel.SetFocus
                        'End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdTax.Left - pbcArrow.Width - 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + (grdTax.RowHeight(grdTax.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdTax.TextMatrix(llEnableRow, GROSSNETINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdTax.Row + 1 >= grdTax.Rows Then
                            grdTax.AddItem ""
                            grdTax.RowHeight(grdTax.Row + 1) = fgFlexGridRowH
                            grdTax.TextMatrix(grdTax.Row + 1, TRFCODEINDEX) = 0
                        End If
                        grdTax.Row = grdTax.Row + 1
                        If (Not grdTax.RowIsVisible(grdTax.Row)) Or (grdTax.Row - (grdTax.TopRow - grdTax.FixedRows) >= imInitNoRows) Then
                            imIgnoreScroll = True
                            grdTax.TopRow = grdTax.TopRow + 1
                        End If
                        grdTax.Col = GROSSNETINDEX
                        grdTax.TextMatrix(grdTax.Row, TRFCODEINDEX) = 0
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdTax.Left - pbcArrow.Width - 30, grdTax.Top + grdTax.RowPos(grdTax.Row) + (grdTax.RowHeight(grdTax.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdTax.Col = grdTax.Col + 1
                'If gColOk(grdTax, grdTax.Row, grdTax.Col) Then
                    mEnableBox
                'Else
                '    ilNext = True
                'End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdTax.TopRow = grdTax.FixedRows
        grdTax.Col = GROSSNETINDEX
        grdTax.Row = grdTax.FixedRows
        'If gColOk(grdTax, grdTax.Row, grdTax.Col) Then
            mEnableBox
        'Else
        '    cmcCancel.SetFocus
        'End If
    End If
End Sub

Private Function mSaveRec() As Integer
    Dim ilRow As Integer
    Dim slMsg As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilError As Integer
    Dim tlTrf As TRF

    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdTax, grdTax, vbHourglass
    For ilRow = grdTax.FixedRows To grdTax.Rows - 1 Step 1
        If mGridFieldsOk(ilRow) = False Then
            ilError = True
        End If
    Next ilRow
    If ilError Then
        gSetMousePointer grdTax, grdTax, vbDefault
        Screen.MousePointer = vbDefault
        Beep
        mSaveRec = False
        Exit Function
    End If
    mMoveCtrlToRec
    ilRet = btrBeginTrans(hmTrf, 1000)
    For ilLoop = 0 To UBound(tgTrf) - 1 Step 1
        If tgTrf(ilLoop).iCode <= 0 Then
            tgTrf(ilLoop).iCode = 0
            ilRet = btrInsert(hmTrf, tgTrf(ilLoop), imTrfRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert:Tax Table)"
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, TaxTable
                On Error GoTo 0
            End If
        Else
            Do
                tmTrfSrchKey0.iCode = tgTrf(ilLoop).iCode
                ilRet = btrGetEqual(hmTrf, tlTrf, imTrfRecLen, tmTrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                ilRet = btrUpdate(hmTrf, tgTrf(ilLoop), imTrfRecLen)
                slMsg = "mSaveRec (btrUpdate:Inventory Schedule)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, TaxTable
                On Error GoTo 0
            End If
        End If
    Next ilLoop
    imTaxChg = False
    ilRet = btrEndTrans(hmTrf)
    mSaveRec = True
    gSetMousePointer grdTax, grdTax, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    ilRet = btrAbortTrans(hmTrf)
    gSetMousePointer grdTax, grdTax, vbDefault
    Screen.MousePointer = vbDefault
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilTrf                                                   *
'******************************************************************************************

    Dim llRow As Long
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilCode As Integer

    For llRow = grdTax.FixedRows To grdTax.Rows - 1 Step 1
        slStr = Trim$(grdTax.TextMatrix(llRow, TAX1NAMEINDEX))
        If slStr <> "" Then
            ilCode = Val(grdTax.TextMatrix(llRow, TRFCODEINDEX))
            If ilCode <> 0 Then
                ilIndex = gBinarySearchTrf(ilCode)
                If ilIndex = -1 Then
                    ilCode = 0
                End If
            End If
            If ilCode = 0 Then
                ilIndex = UBound(tgTrf)
                ReDim Preserve tgTrf(0 To UBound(tgTrf) + 1) As TRF
                tgTrf(ilIndex).iCode = 0
                tgTrf(ilIndex).iUrfCode = tgUrf(0).iCode
            End If
            slStr = Left$(grdTax.TextMatrix(llRow, GROSSNETINDEX), 1)
            If StrComp(Trim$(tgTrf(ilIndex).sGrossNet), slStr, vbTextCompare) <> 0 Then
                tgTrf(ilIndex).iUrfCode = tgUrf(0).iCode
            End If
            tgTrf(ilIndex).sGrossNet = slStr
            tgTrf(ilIndex).sTax1Name = Trim$(grdTax.TextMatrix(llRow, TAX1NAMEINDEX))
            If StrComp(Trim$(tgTrf(ilIndex).sTax1Name), Trim$(grdTax.TextMatrix(llRow, TAX1NAMEINDEX)), vbTextCompare) <> 0 Then
                tgTrf(ilIndex).iUrfCode = tgUrf(0).iCode
            End If
            tgTrf(ilIndex).sTax1Name = Trim$(grdTax.TextMatrix(llRow, TAX1NAMEINDEX))
            slStr = Trim$(grdTax.TextMatrix(llRow, TAX1RATEINDEX))
            If tgTrf(ilIndex).lTax1Rate <> gStrDecToLong(slStr, 4) Then
                tgTrf(ilIndex).iUrfCode = tgUrf(0).iCode
            End If
            tgTrf(ilIndex).lTax1Rate = gStrDecToLong(slStr, 4)
            If StrComp(Trim$(tgTrf(ilIndex).sTax2Name), Trim$(grdTax.TextMatrix(llRow, TAX2NAMEINDEX)), vbTextCompare) <> 0 Then
                tgTrf(ilIndex).iUrfCode = tgUrf(0).iCode
            End If
            tgTrf(ilIndex).sTax2Name = Trim$(grdTax.TextMatrix(llRow, TAX2NAMEINDEX))
            slStr = Trim$(grdTax.TextMatrix(llRow, TAX2RATEINDEX))
            If tgTrf(ilIndex).lTax2Rate <> gStrDecToLong(slStr, 4) Then
                tgTrf(ilIndex).iUrfCode = tgUrf(0).iCode
            End If
            tgTrf(ilIndex).lTax2Rate = gStrDecToLong(slStr, 4)
        End If
    Next llRow
    Exit Sub
End Sub

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
'*  llRet                                                                                 *
'******************************************************************************************

    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilLoop As Integer


    grdTax.Redraw = False
    grdTax.Rows = imInitNoRows
    For llRow = grdTax.FixedRows To grdTax.Rows - 1 Step 1
        grdTax.RowHeight(llRow) = fgFlexGridRowH
        For ilCol = 0 To grdTax.Cols - 1 Step 1
            If ilCol = TRFCODEINDEX Then
                grdTax.TextMatrix(llRow, ilCol) = 0
            Else
                grdTax.TextMatrix(llRow, ilCol) = ""
            End If
        Next ilCol
    Next llRow
    llRow = grdTax.FixedRows

    For ilLoop = 0 To UBound(tgTrf) - 1 Step 1
        If llRow >= grdTax.Rows Then
            grdTax.AddItem ""
            grdTax.RowHeight(llRow) = fgFlexGridRowH
        End If
        If tgTrf(ilLoop).sGrossNet = "G" Then
            grdTax.TextMatrix(llRow, GROSSNETINDEX) = "Gross"
        Else
            grdTax.TextMatrix(llRow, GROSSNETINDEX) = "Net"
        End If
        grdTax.TextMatrix(llRow, TAX1NAMEINDEX) = Trim$(tgTrf(ilLoop).sTax1Name)
        grdTax.TextMatrix(llRow, TAX1RATEINDEX) = gLongToStrDec(tgTrf(ilLoop).lTax1Rate, 4)
        grdTax.TextMatrix(llRow, TAX2NAMEINDEX) = Trim$(tgTrf(ilLoop).sTax2Name)
        grdTax.TextMatrix(llRow, TAX2RATEINDEX) = gLongToStrDec(tgTrf(ilLoop).lTax2Rate, 4)
        grdTax.TextMatrix(llRow, TRFCODEINDEX) = tgTrf(ilLoop).iCode
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdTax.Rows Then
        grdTax.AddItem ""
        grdTax.RowHeight(llRow) = fgFlexGridRowH
        grdTax.TextMatrix(llRow, TRFCODEINDEX) = 0
    End If

    'Remove highlight
    mTaxSortCol TAX2NAMEINDEX
    mTaxSortCol TAX1NAMEINDEX
    grdTax.Row = 0
    grdTax.Col = TRFCODEINDEX
    grdTax.Redraw = True
    mSetCommands

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
    Dim ilError As Integer

    ilError = False
    slStr = Trim$(grdTax.TextMatrix(ilRowNo, TAX1NAMEINDEX))
    If slStr <> "" Then
        slStr = Trim$(grdTax.TextMatrix(ilRowNo, TAX1RATEINDEX))
        If gStrDecToLong(slStr, 4) > 0 Then
            If Trim$(grdTax.TextMatrix(ilRowNo, TAX1NAMEINDEX)) = "" Then
                grdTax.TextMatrix(ilRowNo, TAX1NAMEINDEX) = "Missing"
                ilError = True
                grdTax.Row = ilRowNo
                grdTax.Col = TAX1NAMEINDEX
                grdTax.CellForeColor = vbRed
            End If
        End If
        slStr = Trim$(grdTax.TextMatrix(ilRowNo, TAX2RATEINDEX))
        If gStrDecToLong(slStr, 4) > 0 Then
            If Trim$(grdTax.TextMatrix(ilRowNo, TAX2NAMEINDEX)) = "" Then
                grdTax.TextMatrix(ilRowNo, TAX2NAMEINDEX) = "Missing"
                ilError = True
                grdTax.Row = ilRowNo
                grdTax.Col = TAX2NAMEINDEX
                grdTax.CellForeColor = vbRed
            End If
        End If
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function
Private Sub pbcGrossNet_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcGrossNet_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("G") Or (KeyAscii = Asc("g")) Then
        If smGrossNet <> "Gross" Then
            imTaxChg = True
        End If
        smGrossNet = "Gross"
        pbcGrossNet_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If smGrossNet <> "Net" Then
            imTaxChg = True
        End If
        smGrossNet = "Net"
        pbcGrossNet_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If smGrossNet = "Gross" Then
            imTaxChg = True
            smGrossNet = "Net"
            pbcGrossNet_Paint
        ElseIf smGrossNet = "Net" Then
            imTaxChg = True
            smGrossNet = "Gross"
            pbcGrossNet_Paint
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcGrossNet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If smGrossNet = "Gross" Then
        imTaxChg = True
        smGrossNet = "Net"
    ElseIf smGrossNet = "Net" Then
        imTaxChg = True
        smGrossNet = "Gross"
    End If
    pbcGrossNet_Paint
    mSetCommands
End Sub

Private Sub pbcGrossNet_Paint()
    pbcGrossNet.Cls
    pbcGrossNet.CurrentX = fgBoxInsetX
    pbcGrossNet.CurrentY = 0 'fgBoxInsetY
    pbcGrossNet.Print smGrossNet
End Sub

