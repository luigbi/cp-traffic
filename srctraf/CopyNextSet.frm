VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CopyNextSet 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3780
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
   ScaleHeight     =   3780
   ScaleWidth      =   9315
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8055
      Top             =   3300
   End
   Begin VB.TextBox edcSetAll 
      Height          =   315
      Left            =   1710
      TabIndex        =   2
      Top             =   255
      Width           =   810
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8670
      Top             =   3300
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   11
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
      Picture         =   "CopyNextSet.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   3
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
      TabIndex        =   7
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
      TabIndex        =   4
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5025
      TabIndex        =   9
      Top             =   3315
      Width           =   945
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
      Left            =   3255
      TabIndex        =   8
      Top             =   3315
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdNext 
      Height          =   2610
      Left            =   210
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   4604
      _Version        =   393216
      Rows            =   10
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdInst 
      Height          =   2640
      Left            =   5790
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   4657
      _Version        =   393216
      Rows            =   10
      Cols            =   5
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
      _Band(0).Cols   =   5
   End
   Begin VB.Label lacSetAll 
      Caption         =   "Set All 'Next' to"
      Height          =   210
      Left            =   225
      TabIndex        =   1
      Top             =   300
      Width           =   1440
   End
   Begin VB.Label lacScreen 
      Caption         =   "Set Next Copy Assignment"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   2385
   End
End
Attribute VB_Name = "CopyNextSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CopyNextSet.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CopyNextSet.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim hmCvf As Integer
Dim tmCvf As CVF        'Rvf record image
Dim tmCvfSrchKey As LONGKEY0
Dim imCvfRecLen As Integer        'RvF record length
'Inventory
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0  'CIF key record image
Dim tmCifSrchKey1 As CIFKEY1  'CIF key record image
Dim tmCifSrchKey4 As CIFKEY4  'CIF key record image - used for vCreative
Dim hmCif As Integer        'CIF Handle
Dim imCifRecLen As Integer      'CIF record length
'Product
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0  'CPF key record image
Dim hmCpf As Integer        'CPF Handle
Dim imCpfRecLen As Integer      'CPF record length
'  Media code File
Dim hmMcf As Integer        'Media file handle
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer     'MCF record length

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Height adjustment factor

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imLastSelectRow As Integer
Dim imCtrlKey As Integer
Dim imLastColSorted As Integer
Dim imLastSort As Integer
Dim lmRowSelected As Long
Dim imChg As Integer
Dim imIgnoreScroll As Integer
Dim imFromArrow As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer

Private smMaxRotNumber As String


'GrdNext
Const VEHICLENAMEINDEX = 0
Const TYPEINDEX = 1
Const NEXTINDEX = 2
Const SORTINDEX = 3
Const CVFCODEINDEX = 4
Const CVFINDEXINDEX = 5

'GrdInst
Const NUMBERINDEX = 0
Const CARTINDEX = 1
Const ISCIINDEX = 2
Const CIFCODEINDEX = 3
Const CPFCODEINDEX = 4




Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If imChg Then
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



Private Sub edcDropDown_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                                                   *
'******************************************************************************************


    Select Case lmEnableCol
        Case NEXTINDEX
    End Select
    grdNext.CellForeColor = vbBlack

End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmEnableCol
        Case NEXTINDEX
    End Select
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim ilPos As Integer
    Dim slStr As String

    ilKey = KeyAscii

    Select Case lmEnableCol
        Case NEXTINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "999") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select

End Sub

Private Sub edcSetAll_Change()
    tmcDelay.Enabled = False
    tmcDelay.Enabled = True
End Sub

Private Sub edcSetAll_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSetAll_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcSetAll.Text
    slStr = Left$(slStr, edcSetAll.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSetAll.SelStart - edcSetAll.SelLength)
    If gCompNumberStr(slStr, smMaxRotNumber) > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub edcSetAll_LostFocus()
    If tmcDelay.Enabled Then
        tmcDelay_Timer
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
    CopyNextSet.Refresh
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
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = (((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        If fmAdjFactorW < 1# Then
            fmAdjFactorW = 1#
        Else
            Me.Width = ((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        End If
        fmAdjFactorH = (((lgPercentAdjH) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        If fmAdjFactorH < 1# Then
            fmAdjFactorH = 1#
        Else
            Me.Height = ((lgPercentAdjH) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
        End If
    End If
    mInit
    If imTerminate Then
        'mTerminate
        tmcTerminate.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    btrExtClear hmCvf   'Clear any previous extend operation
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    btrExtClear hmMcf   'Clear any previous extend operation
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    Set CopyNextSet = Nothing   'Remove data segment
End Sub

Private Sub grdNext_EnterCell()
    mSetShow
End Sub

Private Sub grdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdNext.TopRow
    grdNext.Redraw = False
End Sub

Private Sub grdNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer

    imIgnoreScroll = False
    If Y < grdNext.RowHeight(0) Then
        'grdNext.Col = grdNext.MouseCol
        'mNextSortCol grdNext.Col
        Exit Sub
    End If
    pbcArrow.Visible = False
    ilCol = grdNext.MouseCol
    ilRow = grdNext.MouseRow
    If ilCol < grdNext.FixedCols Then
        grdNext.Redraw = True
        Exit Sub
    End If
    If ilRow < grdNext.FixedRows Then
        grdNext.Redraw = True
        Exit Sub
    End If
    If grdNext.TextMatrix(ilRow, VEHICLENAMEINDEX) = "" Then
        grdNext.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdNext.TextMatrix(ilRow, VEHICLENAMEINDEX) = ""
        grdNext.Row = ilRow
        grdNext.Col = NEXTINDEX
    Else
        grdNext.Row = ilRow
        grdNext.Col = NEXTINDEX
    End If
    grdNext.Redraw = True
    lmTopRow = grdNext.TopRow
    mEnableBox
End Sub

Private Sub grdNext_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdNext.Redraw = False Then
        grdNext.Redraw = True
        If lmTopRow < grdNext.FixedRows Then
            grdNext.TopRow = grdNext.FixedRows
        Else
            grdNext.TopRow = lmTopRow
        End If
        grdNext.Refresh
        grdNext.Redraw = False
    End If
    If (imCtrlVisible) And (grdNext.Row >= grdNext.FixedRows) And (grdNext.Col >= 0) And (grdNext.Col < grdNext.Cols - 1) Then
        If grdNext.RowIsVisible(grdNext.Row) Then
            pbcArrow.Move grdNext.Left - pbcArrow.Width - 30, grdNext.Top + grdNext.RowPos(grdNext.Row) + (grdNext.RowHeight(grdNext.Row) - pbcArrow.Height) / 2
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
    'CopyNextSet.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone CopyNextSet
    'CopyNextSet.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    lmRowSelected = -1
    imChg = False

    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    hmCvf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cvf.Btr)", CopyNextSet
    On Error GoTo 0
    imCvfRecLen = Len(tmCvf)  'Get and save CHF record length

    hmCif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", CopyNextSet
    On Error GoTo 0
    imCifRecLen = Len(tmCif)  'Get and save CHF record length

    hmCpf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", CopyNextSet
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)  'Get and save CHF record length

    hmMcf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", CopyNextSet
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)  'Get and save CHF record length

    mInitBox

    mPopNext
    mPopInst
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
    Unload CopyNextSet
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
    If grdNext.Visible Then
        lmRowSelected = -1
        grdNext.Row = 0
        grdNext.Col = CVFCODEINDEX
        mSetCommands
    End If
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

    If imChg Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
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

    If fmAdjFactorW > 1 Then
        cmcDone.Left = Me.Width / 2 - cmcDone.Width / 2 - cmcDone.Width
        cmcCancel.Left = Me.Width / 2 + cmcDone.Width / 2
        cmcDone.Top = Me.Height - (3 * cmcDone.Height) / 2
        cmcCancel.Top = cmcDone.Top
        grdNext.Width = fmAdjFactorW * grdNext.Width
        grdInst.Width = fmAdjFactorW * grdInst.Width
        grdNext.Height = Me.Height - grdNext.Top - (2 * cmcDone.Height)
        grdInst.Height = grdNext.Height
        grdInst.Left = grdNext.Left + grdNext.Width + (Me.Width - (grdNext.Left + grdNext.Width)) / 2 - grdInst.Width / 2
    End If
        

    mGridNextLayout
    mGridNextColumnWidths
    mGridNextColumns
    mGridInstLayout
    mGridInstColumnWidths
    mGridInstColumns
    'grdNext.Height = grdNext.RowPos(0) + 14 * grdNext.RowHeight(0) + fgPanelAdj - 15
    imInitNoRows = (cmcDone.Top - 120 - grdNext.Top) \ fgFlexGridRowH
    grdNext.Height = grdNext.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj - 15
    grdInst.Height = grdInst.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj - 15
End Sub

Private Sub mGridNextLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdNext.Rows - 1 Step 1
        grdNext.RowHeight(ilRow) = fgFlexGridRowH
    Next ilRow
    For ilCol = 0 To grdNext.Cols - 1 Step 1
        grdNext.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridNextColumns()

    grdNext.Row = grdNext.FixedRows - 1
    grdNext.Col = VEHICLENAMEINDEX
    grdNext.CellFontBold = False
    grdNext.CellFontName = "Arial"
    grdNext.CellFontSize = 6.75
    grdNext.CellForeColor = vbBlue
    'grdNext.CellBackColor = LIGHTBLUE
    grdNext.TextMatrix(grdNext.Row, grdNext.Col) = "Vehicle Name"
    grdNext.Col = TYPEINDEX
    grdNext.CellFontBold = False
    grdNext.CellFontName = "Arial"
    grdNext.CellFontSize = 6.75
    grdNext.CellForeColor = vbBlue
    'grdNext.CellBackColor = LIGHTBLUE
    grdNext.TextMatrix(grdNext.Row, grdNext.Col) = "Type"
    grdNext.Col = NEXTINDEX
    grdNext.CellFontBold = False
    grdNext.CellFontName = "Arial"
    grdNext.CellFontSize = 6.75
    grdNext.CellForeColor = vbBlue
    'grdNext.CellBackColor = LIGHTBLUE
    grdNext.TextMatrix(grdNext.Row, grdNext.Col) = "Next"
    grdNext.Col = SORTINDEX
    grdNext.CellFontBold = False
    grdNext.CellFontName = "Arial"
    grdNext.CellFontSize = 6.75
    grdNext.CellForeColor = vbBlue
    'grdNext.CellBackColor = LIGHTBLUE
    grdNext.TextMatrix(grdNext.Row, grdNext.Col) = "Sort"
    grdNext.Col = CVFCODEINDEX
    grdNext.CellFontBold = False
    grdNext.CellFontName = "Arial"
    grdNext.CellFontSize = 6.75
    grdNext.CellForeColor = vbBlue
    'grdNext.CellBackColor = LIGHTBLUE
    grdNext.TextMatrix(grdNext.Row, grdNext.Col) = "CvfCode"
    grdNext.Col = CVFINDEXINDEX
    grdNext.CellFontBold = False
    grdNext.CellFontName = "Arial"
    grdNext.CellFontSize = 6.75
    grdNext.CellForeColor = vbBlue
    'grdNext.CellBackColor = LIGHTBLUE
    grdNext.TextMatrix(grdNext.Row, grdNext.Col) = "Index"

End Sub

Private Sub mGridNextColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdNext.ColWidth(CVFINDEXINDEX) = 0
    grdNext.ColWidth(CVFCODEINDEX) = 0
    grdNext.ColWidth(SORTINDEX) = 0
    grdNext.ColWidth(VEHICLENAMEINDEX) = 0.6 * grdNext.Width
    grdNext.ColWidth(TYPEINDEX) = 0.2 * grdNext.Width
    grdNext.ColWidth(NEXTINDEX) = 0.1 * grdNext.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdNext.Width
    For ilCol = 0 To grdNext.Cols - 1 Step 1
        llWidth = llWidth + grdNext.ColWidth(ilCol)
        If (grdNext.ColWidth(ilCol) > 15) And (grdNext.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdNext.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdNext.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdNext.Width
            For ilCol = 0 To grdNext.Cols - 1 Step 1
                If (grdNext.ColWidth(ilCol) > 15) And (grdNext.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdNext.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdNext.FixedCols To grdNext.Cols - 1 Step 1
                If grdNext.ColWidth(ilCol) > 15 Then
                    ilColInc = grdNext.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdNext.ColWidth(ilCol) = grdNext.ColWidth(ilCol) + 15
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


Private Sub mNextSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdNext.FixedRows To grdNext.Rows - 1 Step 1
        slStr = Trim$(grdNext.TextMatrix(llRow, VEHICLENAMEINDEX))
        If slStr <> "" Then
            If ilCol = VEHICLENAMEINDEX Then
                slSort = grdNext.TextMatrix(llRow, VEHICLENAMEINDEX)
                Do While Len(slSort) < 40
                    slSort = slSort & " "
                Loop
            ElseIf ilCol = TYPEINDEX Then
                slSort = grdNext.TextMatrix(llRow, TYPEINDEX)
                Do While Len(slSort) < 10
                    slSort = slSort & " "
                Loop
            ElseIf (ilCol = NEXTINDEX) Then
                slSort = grdNext.TextMatrix(llRow, NEXTINDEX)
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            End If
            slStr = grdNext.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdNext.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdNext.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdNext, VEHICLENAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
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
    If (grdNext.Row < grdNext.FixedRows) Or (grdNext.Row >= grdNext.Rows) Or (grdNext.Col < grdNext.FixedCols) Or (grdNext.Col >= grdNext.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdNext.Row
    lmEnableCol = grdNext.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdNext.Left - pbcArrow.Width - 30, grdNext.Top + grdNext.RowPos(grdNext.Row) + (grdNext.RowHeight(grdNext.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    Select Case grdNext.Col
        Case NEXTINDEX
            edcDropDown.MaxLength = 4
            slStr = grdNext.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdNext.Row > grdNext.FixedRows Then
                    slStr = grdNext.TextMatrix(grdNext.Row - 1, grdNext.Col)
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

    If (grdNext.Row < grdNext.FixedRows) Or (grdNext.Row >= grdNext.Rows) Or (grdNext.Col < grdNext.FixedCols) Or (grdNext.Col >= grdNext.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdNext.Left - pbcArrow.Width - 30, grdNext.Top + grdNext.RowPos(grdNext.Row) + (grdNext.RowHeight(grdNext.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdNext.Col - 1 Step 1
        llColPos = llColPos + grdNext.ColWidth(ilCol)
    Next ilCol
    Select Case grdNext.Col
        Case NEXTINDEX
            edcDropDown.Move grdNext.Left + llColPos + 30, grdNext.Top + grdNext.RowPos(grdNext.Row) + 30, grdNext.ColWidth(grdNext.Col) - 30, grdNext.RowHeight(grdNext.Row) - 15
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
    If (lmEnableRow >= grdNext.FixedRows) And (lmEnableRow < grdNext.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NEXTINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If StrComp(grdNext.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imChg = True
                End If
                grdNext.TextMatrix(lmEnableRow, lmEnableCol) = slStr
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
            If grdNext.Col = NEXTINDEX Then
                If grdNext.Row > grdNext.FixedRows Then
                    lmTopRow = -1
                    grdNext.Row = grdNext.Row - 1
                    If Not grdNext.RowIsVisible(grdNext.Row) Then
                        grdNext.TopRow = grdNext.TopRow - 1
                    End If
                    grdNext.Col = NEXTINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                cmcCancel.SetFocus
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdNext.TopRow = grdNext.FixedRows
        grdNext.Col = NEXTINDEX
        grdNext.Row = grdNext.FixedRows
        mEnableBox
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
            If grdNext.Col = NEXTINDEX Then
                llRow = grdNext.Rows
                Do
                    llRow = llRow - 1
                Loop While grdNext.TextMatrix(llRow, VEHICLENAMEINDEX) = ""
                llRow = llRow + 1
                If (grdNext.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdNext.Row = grdNext.Row + 1
                    If Not grdNext.RowIsVisible(grdNext.Row) Or (grdNext.Row - (grdNext.TopRow - grdNext.FixedRows) >= imInitNoRows) Then
                        imIgnoreScroll = True
                        grdNext.TopRow = grdNext.TopRow + 1
                    End If
                    grdNext.Col = NEXTINDEX
                    'grdNext.TextMatrix(grdNext.Row, CODEINDEX) = 0
                    If Trim$(grdNext.TextMatrix(grdNext.Row, VEHICLENAMEINDEX)) <> "" Then
                        'If gColOk(grdNext, grdNext.Row, grdNext.Col) Then
                            mEnableBox
                        'Else
                        '    cmcCancel.SetFocus
                        'End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdNext.Left - pbcArrow.Width - 30, grdNext.Top + grdNext.RowPos(grdNext.Row) + (grdNext.RowHeight(grdNext.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                pbcClickFocus.SetFocus
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdNext.TopRow = grdNext.FixedRows
        grdNext.Col = NEXTINDEX
        grdNext.Row = grdNext.FixedRows
        mEnableBox
    End If
End Sub

Private Function mSaveRec() As Integer
    Dim ilRow As Integer
    Dim slMsg As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilError As Integer
    Dim ilIndex As Integer
    
    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdNext, grdNext, vbHourglass
    For ilRow = grdNext.FixedRows To grdNext.Rows - 1 Step 1
        If mGridFieldsOk(ilRow) = False Then
            ilError = True
        End If
    Next ilRow
    If ilError Then
        gSetMousePointer grdNext, grdNext, vbDefault
        Screen.MousePointer = vbDefault
        Beep
        mSaveRec = False
        Exit Function
    End If
    ilRet = btrBeginTrans(hmCvf, 1000)
'    For ilLoop = 0 To UBound(tgTrf) - 1 Step 1
'        If tgTrf(ilLoop).iCode <= 0 Then
'            tgTrf(ilLoop).iCode = 0
'            ilRet = btrInsert(hmCvf, tgTrf(ilLoop), imCvfRecLen, INDEXKEY0)
'            slMsg = "mSaveRec (btrInsert:Tax Table)"
'            If ilRet <> BTRV_ERR_NONE Then
'                On Error GoTo mSaveRecErr
'                gBtrvErrorMsg ilRet, slMsg, CopyNextSet
'                On Error GoTo 0
'            End If
'        Else
'            Do
'                tmCvfSrchKey0.iCode = tgTrf(ilLoop).iCode
'                ilRet = btrGetEqual(hmCvf, tlTrf, imCvfRecLen, tmCvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                ilRet = btrUpdate(hmCvf, tgTrf(ilLoop), imCvfRecLen)
'                slMsg = "mSaveRec (btrUpdate:Inventory Schedule)"
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            If ilRet <> BTRV_ERR_NONE Then
'                On Error GoTo mSaveRecErr
'                gBtrvErrorMsg ilRet, slMsg, CopyNextSet
'                On Error GoTo 0
'            End If
'        End If
'    Next ilLoop
    igFLValue = -1
    For ilRow = grdNext.FixedRows To grdNext.Rows - 1 Step 1
        tmCvfSrchKey.lCode = Trim$(grdNext.TextMatrix(ilRow, CVFCODEINDEX))
        If (tmCvfSrchKey.lCode > 0) And (Trim$(grdNext.TextMatrix(ilRow, VEHICLENAMEINDEX)) <> "") Then
            ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            ilIndex = Trim$(grdNext.TextMatrix(ilRow, CVFINDEXINDEX))
            tmCvf.iNextFinal(ilIndex) = Trim$(grdNext.TextMatrix(ilRow, NEXTINDEX))
            tmCvf.iNextPrelim(ilIndex) = Trim$(grdNext.TextMatrix(ilRow, NEXTINDEX))
            If igFLValue = -1 Then
                igFLValue = tmCvf.iNextFinal(ilIndex)
            Else
                If tmCvf.iNextFinal(ilIndex) <> igFLValue Then
                    igFLValue = -2
                End If
            End If
            ilRet = btrUpdate(hmCvf, tmCvf, imCvfRecLen)
        End If
    Next ilRow
    imChg = False
    ilRet = btrEndTrans(hmCvf)
    mSaveRec = True
    gSetMousePointer grdNext, grdNext, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    ilRet = btrAbortTrans(hmCvf)
    gSetMousePointer grdNext, grdNext, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function




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
    slStr = Trim$(grdNext.TextMatrix(ilRowNo, VEHICLENAMEINDEX))
    If slStr <> "" Then
        If Trim$(grdNext.TextMatrix(ilRowNo, NEXTINDEX)) = "" Then
            grdNext.TextMatrix(ilRowNo, NEXTINDEX) = "Missing"
            ilError = True
            grdNext.Row = ilRowNo
            grdNext.Col = NEXTINDEX
            grdNext.CellForeColor = vbRed
        End If
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function

Private Sub mGridInstLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdInst.Rows - 1 Step 1
        grdInst.RowHeight(ilRow) = fgFlexGridRowH
    Next ilRow
    For ilCol = 0 To grdInst.Cols - 1 Step 1
        grdInst.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridInstColumns()

    grdInst.Row = grdInst.FixedRows - 1
    grdInst.Col = NUMBERINDEX
    grdInst.CellFontBold = False
    grdInst.CellFontName = "Arial"
    grdInst.CellFontSize = 6.75
    grdInst.CellForeColor = vbBlue
    'grdInst.CellBackColor = LIGHTBLUE
    grdInst.TextMatrix(grdInst.Row, grdInst.Col) = "Number"
    grdInst.Col = CARTINDEX
    grdInst.CellFontBold = False
    grdInst.CellFontName = "Arial"
    grdInst.CellFontSize = 6.75
    grdInst.CellForeColor = vbBlue
    'grdInst.CellBackColor = LIGHTBLUE
    grdInst.TextMatrix(grdInst.Row, grdInst.Col) = "Cart"
    grdInst.Col = ISCIINDEX
    grdInst.CellFontBold = False
    grdInst.CellFontName = "Arial"
    grdInst.CellFontSize = 6.75
    grdInst.CellForeColor = vbBlue
    'grdInst.CellBackColor = LIGHTBLUE
    grdInst.TextMatrix(grdInst.Row, grdInst.Col) = "ISCI"
    grdInst.Col = CIFCODEINDEX
    grdInst.CellFontBold = False
    grdInst.CellFontName = "Arial"
    grdInst.CellFontSize = 6.75
    grdInst.CellForeColor = vbBlue
    'grdInst.CellBackColor = LIGHTBLUE
    grdInst.TextMatrix(grdInst.Row, grdInst.Col) = "CifCode"
    grdInst.Col = CPFCODEINDEX
    grdInst.CellFontBold = False
    grdInst.CellFontName = "Arial"
    grdInst.CellFontSize = 6.75
    grdInst.CellForeColor = vbBlue
    'grdInst.CellBackColor = LIGHTBLUE
    grdInst.TextMatrix(grdInst.Row, grdInst.Col) = "CpfCode"

End Sub

Private Sub mGridInstColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdInst.ColWidth(CPFCODEINDEX) = 0
    grdInst.ColWidth(CIFCODEINDEX) = 0
    grdInst.ColWidth(NUMBERINDEX) = 0.15 * grdInst.Width
    grdInst.ColWidth(CARTINDEX) = 0.4 * grdInst.Width
    grdInst.ColWidth(ISCIINDEX) = 0.4 * grdInst.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdInst.Width
    For ilCol = 0 To grdInst.Cols - 1 Step 1
        llWidth = llWidth + grdInst.ColWidth(ilCol)
        If (grdInst.ColWidth(ilCol) > 15) And (grdInst.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdInst.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdInst.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdInst.Width
            For ilCol = 0 To grdInst.Cols - 1 Step 1
                If (grdInst.ColWidth(ilCol) > 15) And (grdInst.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdInst.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdInst.FixedCols To grdInst.Cols - 1 Step 1
                If grdInst.ColWidth(ilCol) > 15 Then
                    ilColInc = grdInst.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdInst.ColWidth(ilCol) = grdInst.ColWidth(ilCol) + 15
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

Private Sub mPopNext()
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slCart As String
    Dim ilCvf As Integer
    Dim ilVef As Integer
    
    For llRow = grdNext.FixedRows To grdNext.Rows - 1 Step 1
        grdNext.RowHeight(llRow) = fgFlexGridRowH
        For ilCol = 0 To grdNext.Cols - 1 Step 1
            If ilCol = CVFCODEINDEX Then
                grdNext.TextMatrix(llRow, ilCol) = 0
            Else
                grdNext.TextMatrix(llRow, ilCol) = ""
            End If
        Next ilCol
        grdNext.Row = llRow
        For ilCol = VEHICLENAMEINDEX To TYPEINDEX Step 1
            grdNext.Col = ilCol
            grdNext.CellBackColor = LIGHTYELLOW
        Next ilCol
    Next llRow
    llRow = grdNext.FixedRows
    tmCvfSrchKey.lCode = lgSetNextCvfCode
    ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        For ilCvf = 0 To 99 Step 1
            If tmCvf.iVefCode(ilCvf) > 0 Then
                ilVef = gBinarySearchVef(tmCvf.iVefCode(ilCvf))
                If ilVef <> -1 Then
                    If llRow >= grdNext.Rows Then
                        grdNext.AddItem ""
                        grdNext.RowHeight(llRow) = fgFlexGridRowH
                    End If
                    grdNext.Row = llRow
                    For ilCol = VEHICLENAMEINDEX To TYPEINDEX Step 1
                        grdNext.Col = ilCol
                        grdNext.CellBackColor = LIGHTYELLOW
                    Next ilCol
                    grdNext.TextMatrix(llRow, VEHICLENAMEINDEX) = Trim$(tgMVef(ilVef).sName)
                    Select Case Trim$(tgMVef(ilVef).sType)
                        Case "S"
                            grdNext.TextMatrix(llRow, TYPEINDEX) = "Selling"
                        Case "A"
                            grdNext.TextMatrix(llRow, TYPEINDEX) = "Airing"
                        Case "P"
                            grdNext.TextMatrix(llRow, TYPEINDEX) = "Package"
                        Case "L"
                            grdNext.TextMatrix(llRow, TYPEINDEX) = "Log"
                        Case Else
                            grdNext.TextMatrix(llRow, TYPEINDEX) = "Conventional"
                    End Select
                    grdNext.TextMatrix(llRow, NEXTINDEX) = tmCvf.iNextFinal(ilCvf)
                    grdNext.TextMatrix(llRow, CVFCODEINDEX) = tmCvf.lCode
                    grdNext.TextMatrix(llRow, CVFINDEXINDEX) = ilCvf
                    llRow = llRow + 1
                End If
            End If
        Next ilCvf
        If tmCvf.lLkCvfCode <= 0 Then
            Exit Do
        End If
        tmCvfSrchKey.lCode = tmCvf.lLkCvfCode
        ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    mNextSortCol VEHICLENAMEINDEX
    grdNext.Row = 0
    grdNext.Col = CVFCODEINDEX
    grdNext.Redraw = True
    mSetCommands

End Sub

Private Sub mPopInst()
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slCart As String
    
    For llRow = grdInst.FixedRows To grdInst.Rows - 1 Step 1
        grdInst.Row = llRow
        grdInst.RowHeight(llRow) = fgFlexGridRowH
        For ilCol = 0 To grdInst.Cols - 1 Step 1
            If ilCol = CVFCODEINDEX Then
                grdInst.TextMatrix(llRow, ilCol) = 0
            Else
                grdInst.TextMatrix(llRow, ilCol) = ""
            End If
            grdInst.Col = ilCol
            grdInst.CellBackColor = LIGHTYELLOW
        Next ilCol
    Next llRow
    llRow = grdInst.FixedRows
    
    For ilLoop = 0 To Copy.lbcInst.ListCount - 2 Step 1
        slName = Copy.lbcInst.List(ilLoop)
        gFindMatch slName, 0, Copy.lbcExpandActive
        If ((gLastFound(Copy.lbcExpandActive) > 0) And (tgSpf.sUseCartNo <> "B")) Or ((gLastFound(Copy.lbcExpandActive) > 1) And (tgSpf.sUseCartNo = "B")) Then
            If tgSpf.sUseCartNo <> "B" Then
                slNameCode = tgActiveCode(gLastFound(Copy.lbcExpandActive) - 1).sKey
            Else
                slNameCode = tgActiveCode(gLastFound(Copy.lbcExpandActive) - 2).sKey
            End If
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmCifSrchKey.lCode = slCode
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If llRow >= grdInst.Rows Then
                    grdInst.AddItem ""
                    grdInst.RowHeight(llRow) = fgFlexGridRowH
                End If
                grdInst.Row = llRow
                For ilCol = 0 To grdInst.Cols - 1 Step 1
                    grdInst.Col = ilCol
                    grdInst.CellBackColor = LIGHTYELLOW
                Next ilCol
                grdInst.TextMatrix(llRow, NUMBERINDEX) = llRow
                If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                    If tmCif.iMcfCode <> tmMcf.iCode Then
                        tmMcfSrchKey.iCode = tmCif.iMcfCode
                        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmMcf.sName = ""
                        End If
                        slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                    Else
                        slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                    End If
                Else
                    slCart = " "
                End If
                grdInst.TextMatrix(llRow, CARTINDEX) = Trim$(slCart)
                grdInst.TextMatrix(llRow, CPFCODEINDEX) = 0
                If tmCif.lcpfCode > 0 Then
                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        grdInst.TextMatrix(llRow, CPFCODEINDEX) = tmCpf.lCode
                        grdInst.TextMatrix(llRow, ISCIINDEX) = Trim$(tmCpf.sISCI)
                    End If
                End If
                grdInst.TextMatrix(llRow, CIFCODEINDEX) = tmCif.lCode
                llRow = llRow + 1
            End If
        End If
    Next ilLoop
    smMaxRotNumber = llRow - 1
End Sub

Private Sub tmcDelay_Timer()
    Dim ilValue As Integer
    Dim ilRow As Integer
    
    tmcDelay.Enabled = False
    If edcSetAll.Text = "" Then
        Exit Sub
    End If
    ilValue = edcSetAll.Text
    If ilValue <= 0 Then
        Exit Sub
    End If
    For ilRow = grdNext.FixedRows To grdNext.Rows - 1 Step 1
        If (Trim$(grdNext.TextMatrix(ilRow, VEHICLENAMEINDEX)) <> "") Then
            If grdNext.TextMatrix(ilRow, NEXTINDEX) <> ilValue Then
                imChg = True
            End If
            grdNext.TextMatrix(ilRow, NEXTINDEX) = ilValue
        End If
    Next ilRow
    mSetCommands
    
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    If imTerminate Then
        mTerminate
    End If
End Sub
