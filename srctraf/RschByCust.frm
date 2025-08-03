VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form RschByCust 
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
   Icon            =   "RschByCust.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   9315
   Begin VB.ListBox lbcCustDemo 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "RschByCust.frx":08CA
      Left            =   7965
      List            =   "RschByCust.frx":08CC
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3705
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7500
      Top             =   3600
   End
   Begin VB.TextBox edcBookName 
      Height          =   315
      Left            =   6060
      MaxLength       =   30
      TabIndex        =   15
      Top             =   45
      Width           =   3045
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   13
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
      Picture         =   "RschByCust.frx":08CE
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
      TabIndex        =   5
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
      TabIndex        =   8
      Top             =   3720
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4005
      TabIndex        =   7
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   3885
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2235
      TabIndex        =   6
      Top             =   3720
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRsch 
      Height          =   2955
      Left            =   210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   495
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5212
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
   Begin VB.Label lacBookName 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
      Height          =   195
      Left            =   4785
      TabIndex        =   14
      Top             =   45
      Width           =   1020
   End
   Begin VB.Label lacScreen 
      Caption         =   "Single Custom Demo Only"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   2445
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
Attribute VB_Name = "RschByCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RschByCust.frm on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmPopDrf                                                                              *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RschByCust.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim hmDnf As Integer    'Demo Book Name file handle
Dim tmDnf As DNF
Dim tmDnfSrchKey As INTKEY0    'Dnf key record image
Dim imDnfRecLen As Integer        'DNF record length
Dim hmDrf As Integer
Dim tmDrf As DRF        'Rvf record image
Dim tmDrfSrchKey2 As LONGKEY0
Dim imDrfRecLen As Integer        'RvF record length
Dim tmRschByCust() As RSCHBYCUSTINFO
Dim hmMnf As Integer    'file handle
Dim tmMnf As MNF        'Record structure
Dim imMnfRecLen As Integer  'Record length
Dim tmMnfSrchKey As INTKEY0     'MNF key record image
Dim tmCDemoCode() As SORTCODE
Dim smCDemoCodeTag As String
Dim hmVef As Integer    'file handle
Dim tmVef As VEF        'Record structure
Dim imVefRecLen As Integer  'Record length
Dim tmVefSrchKey As INTKEY0     'MNF key record image

Dim imVefCode As Integer

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
Dim imUpdateAllowed As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer


Const POPULATIONINDEX = 0
Const RATINGINDEX = 1
Const AUDIENCEINDEX = 2
Const DATEINDEX = 3
Const DRFCODEINDEX = 4
Const SORTINDEX = 5
Const CHGINDEX = 6


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

Private Sub cmcUpdate_Click()
    Dim ilRet As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    ilRet = mSaveRec()
    If ilRet Then
        mPopulate
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
        Case POPULATIONINDEX
        Case RATINGINDEX
        Case AUDIENCEINDEX
        Case DATEINDEX
    End Select
    grdRsch.CellForeColor = vbBlack

End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmEnableCol
        Case POPULATIONINDEX
        Case RATINGINDEX
        Case AUDIENCEINDEX
        Case DATEINDEX
    End Select
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim ilPos As Integer
    Dim slStr As String

    ilKey = KeyAscii

    Select Case lmEnableCol
        Case POPULATIONINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case RATINGINDEX
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
            If gCompNumberStr(slStr, "100.00") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case AUDIENCEINDEX
        Case DATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
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
    RschByCust.Refresh
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
    tmcStart.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tmRschByCust
    Erase tmCDemoCode

    btrExtClear hmDnf   'Clear any previous extend operation
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    btrExtClear hmDrf   'Clear any previous extend operation
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    Set RschByCust = Nothing   'Remove data segment
End Sub


Private Sub grdRsch_EnterCell()
    mSetShow
End Sub

Private Sub grdRsch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdRsch.TopRow
    grdRsch.Redraw = False
End Sub

Private Sub grdRsch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer

    If Not imUpdateAllowed Then
        Exit Sub
    End If
    imIgnoreScroll = False
    If Y < grdRsch.RowHeight(0) Then
        grdRsch.Col = grdRsch.MouseCol
        mSortCol grdRsch.Col
        Exit Sub
    End If
    pbcArrow.Visible = False
    ilCol = grdRsch.MouseCol
    ilRow = grdRsch.MouseRow
    If ilCol < grdRsch.FixedCols Then
        grdRsch.Redraw = True
        Exit Sub
    End If
    If ilRow < grdRsch.FixedRows Then
        grdRsch.Redraw = True
        Exit Sub
    End If
    If grdRsch.TextMatrix(ilRow, POPULATIONINDEX) = "" Then
        grdRsch.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdRsch.TextMatrix(ilRow, POPULATIONINDEX) = ""
        grdRsch.Row = ilRow + 1
        grdRsch.Col = POPULATIONINDEX
        grdRsch.Redraw = True
    Else
        grdRsch.Row = ilRow
        grdRsch.Col = ilCol
    End If
    grdRsch.Redraw = True
    lmTopRow = grdRsch.TopRow
    If Not mColOk() Then
        pbcArrow.Move grdRsch.Left - pbcArrow.Width - 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + (grdRsch.RowHeight(grdRsch.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Exit Sub
    End If
    mEnableBox
End Sub

Private Sub grdRsch_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdRsch.Redraw = False Then
        grdRsch.Redraw = True
        If lmTopRow < grdRsch.FixedRows Then
            grdRsch.TopRow = grdRsch.FixedRows
        Else
            grdRsch.TopRow = lmTopRow
        End If
        grdRsch.Refresh
        grdRsch.Redraw = False
    End If
    If (imCtrlVisible) And (grdRsch.Row >= grdRsch.FixedRows) And (grdRsch.Col >= 0) And (grdRsch.Col < grdRsch.Cols - 1) Then
        If grdRsch.RowIsVisible(grdRsch.Row) Then
            pbcArrow.Move grdRsch.Left - pbcArrow.Width - 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + (grdRsch.RowHeight(grdRsch.Row) - pbcArrow.Height) / 2
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
    RschByCust.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone RschByCust
    'RschByCust.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    lmRowSelected = -1
    imChg = False

    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    hmDnf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dnf.Btr)", RschByCust
    On Error GoTo 0
    imDnfRecLen = Len(tmDnf)  'Get and save CHF record length
    hmDrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Drf.Btr)", RschByCust
    On Error GoTo 0
    imDrfRecLen = Len(tmDrf)  'Get and save CHF record length
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", RschByCust
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)  'Get and save CHF record length
    hmVef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", RschByCust
    On Error GoTo 0
    imVefRecLen = Len(tmVef)  'Get and save CHF record length

    mInitBox

    smCDemoCodeTag = ""
    ilRet = gPopMnfPlusFieldsBox(RschByCust, lbcCustDemo, tmCDemoCode(), smCDemoCodeTag, "DC")

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
    igExitTraffic = 1
    'Unload Traffic
    Unload RschByCust
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
    If grdRsch.Visible Then
        lmRowSelected = -1
        grdRsch.Row = 0
        grdRsch.Col = DRFCODEINDEX
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
    Dim llFound As Long
    Dim llLoop As Long
    Dim slDate As String
    ReDim tmRschByCust(0 To 0) As RSCHBYCUSTINFO
    edcBookName.Text = ""
    ilRet = btrGetFirst(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        Do While (ilRet = BTRV_ERR_NONE)
            If tmDrf.sDataType = "B" Then
                llFound = -1
                For llLoop = 0 To UBound(tmRschByCust) - 1 Step 1
                    If tmDrf.iDnfCode = tmRschByCust(llLoop).iDnfCode Then
                        llFound = llLoop
                        Exit For
                    End If
                Next llLoop
                If llFound = -1 Then
                    llFound = UBound(tmRschByCust)
                    ReDim Preserve tmRschByCust(0 To llFound + 1) As RSCHBYCUSTINFO
                    tmRschByCust(llFound).iDnfCode = tmDrf.iDnfCode
                    tmDnfSrchKey.iCode = tmDrf.iDnfCode
                    ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        gUnpackDateForSort tmDnf.iBookDate(0), tmDnf.iBookDate(1), slDate
                        tmRschByCust(llFound).sKey = slDate
                        gUnpackDate tmDnf.iBookDate(0), tmDnf.iBookDate(1), slDate
                        tmRschByCust(llFound).sDate = slDate
                        If edcBookName.Text = "" Then
                            edcBookName.Text = Trim$(tmDnf.sBookName)
                        End If
                    End If
                End If
                If (tmDrf.sDemoDataType = "P") Then
                    'tmRschByCust(llFound).lPopDemo = tmDrf.lDemo(1)
                    tmRschByCust(llFound).lPopDemo = tmDrf.lDemo(0)
                    tmRschByCust(llFound).lPopDrf = tmDrf.lCode
                ElseIf (tmDrf.sDemoDataType = "D") Then
                    'tmRschByCust(llFound).lCustDemo = tmDrf.lDemo(1)
                    tmRschByCust(llFound).lCustDemo = tmDrf.lDemo(0)
                    tmRschByCust(llFound).lCustDrf = tmDrf.lCode
                End If
            End If
            ilRet = btrGetNext(hmDrf, tmDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If UBound(tmRschByCust) - 1 > 0 Then
        ArraySortTyp fnAV(tmRschByCust(), 0), UBound(tmRschByCust), 0, LenB(tmRschByCust(0)), 0, LenB(tmRschByCust(0).sKey), 0
    End If
    mMoveRecToCtrl
End Sub


Private Sub mSetCommands()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilRow                                                   *
'******************************************************************************************


    If imChg Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
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

    mGridLayout
    mGridColumnWidths
    mGridColumns
    grdRsch.Move 180, edcBookName.Top + edcBookName.Height + 120, grdRsch.Width, cmcDone.Top - (lacScreen.Top + lacScreen.Height) - 240
    'grdRsch.Height = grdRsch.RowPos(0) + 14 * grdRsch.RowHeight(0) + fgPanelAdj - 15
    imInitNoRows = (cmcDone.Top - 120 - grdRsch.Top) \ fgBoxGridH
    grdRsch.Height = grdRsch.RowPos(0) + imInitNoRows * (fgBoxGridH) + fgPanelAdj - 15
End Sub

Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdRsch.Rows - 1 Step 1
        grdRsch.RowHeight(ilRow) = fgBoxGridH
    Next ilRow
    For ilCol = 0 To grdRsch.Cols - 1 Step 1
        grdRsch.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdRsch.Row = grdRsch.FixedRows - 1
    grdRsch.Col = POPULATIONINDEX
    grdRsch.CellFontBold = False
    grdRsch.CellFontName = "Arial"
    grdRsch.CellFontSize = 6.75
    grdRsch.CellForeColor = vbBlue
    grdRsch.CellBackColor = LIGHTBLUE
    grdRsch.TextMatrix(grdRsch.Row, grdRsch.Col) = "Population"
    grdRsch.Col = RATINGINDEX
    grdRsch.CellFontBold = False
    grdRsch.CellFontName = "Arial"
    grdRsch.CellFontSize = 6.75
    grdRsch.CellForeColor = vbBlue
    grdRsch.CellBackColor = LIGHTBLUE
    grdRsch.TextMatrix(grdRsch.Row, grdRsch.Col) = "Rating"
    grdRsch.Col = AUDIENCEINDEX
    grdRsch.CellFontBold = False
    grdRsch.CellFontName = "Arial"
    grdRsch.CellFontSize = 6.75
    grdRsch.CellForeColor = vbBlue
    grdRsch.CellBackColor = LIGHTBLUE
    grdRsch.TextMatrix(grdRsch.Row, grdRsch.Col) = "Audience"
    grdRsch.Col = DATEINDEX
    grdRsch.CellFontBold = False
    grdRsch.CellFontName = "Arial"
    grdRsch.CellFontSize = 6.75
    grdRsch.CellForeColor = vbBlue
    grdRsch.CellBackColor = LIGHTBLUE
    grdRsch.TextMatrix(grdRsch.Row, grdRsch.Col) = "Date"
    grdRsch.Col = DRFCODEINDEX
    grdRsch.CellFontBold = False
    grdRsch.CellFontName = "Arial"
    grdRsch.CellFontSize = 6.75
    grdRsch.CellForeColor = vbBlue
    grdRsch.CellBackColor = LIGHTBLUE
    grdRsch.TextMatrix(grdRsch.Row, grdRsch.Col) = "Drf Code"
    grdRsch.Col = SORTINDEX
    grdRsch.CellFontBold = False
    grdRsch.CellFontName = "Arial"
    grdRsch.CellFontSize = 6.75
    grdRsch.CellForeColor = vbBlue
    grdRsch.CellBackColor = LIGHTBLUE
    grdRsch.TextMatrix(grdRsch.Row, grdRsch.Col) = "Sort"
    grdRsch.Col = CHGINDEX
    grdRsch.CellFontBold = False
    grdRsch.CellFontName = "Arial"
    grdRsch.CellFontSize = 6.75
    grdRsch.CellForeColor = vbBlue
    grdRsch.CellBackColor = LIGHTBLUE
    grdRsch.TextMatrix(grdRsch.Row, grdRsch.Col) = "Changed"

End Sub

Private Sub mGridColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdRsch.ColWidth(DRFCODEINDEX) = 0
    grdRsch.ColWidth(SORTINDEX) = 0
    grdRsch.ColWidth(CHGINDEX) = 0
    grdRsch.ColWidth(POPULATIONINDEX) = 0.2 * grdRsch.Width
    grdRsch.ColWidth(RATINGINDEX) = 0.2 * grdRsch.Width
    grdRsch.ColWidth(AUDIENCEINDEX) = 0.2 * grdRsch.Width
    grdRsch.ColWidth(DATEINDEX) = 0.2 * grdRsch.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdRsch.Width
    For ilCol = 0 To grdRsch.Cols - 1 Step 1
        llWidth = llWidth + grdRsch.ColWidth(ilCol)
        If (grdRsch.ColWidth(ilCol) > 15) And (grdRsch.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdRsch.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdRsch.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdRsch.Width
            For ilCol = 0 To grdRsch.Cols - 1 Step 1
                If (grdRsch.ColWidth(ilCol) > 15) And (grdRsch.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdRsch.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdRsch.FixedCols To grdRsch.Cols - 1 Step 1
                If grdRsch.ColWidth(ilCol) > 15 Then
                    ilColInc = grdRsch.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdRsch.ColWidth(ilCol) = grdRsch.ColWidth(ilCol) + 15
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


Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdRsch.FixedRows To grdRsch.Rows - 1 Step 1
        slStr = Trim$(grdRsch.TextMatrix(llRow, POPULATIONINDEX))
        If slStr <> "" Then
            If ilCol = POPULATIONINDEX Then
                slSort = grdRsch.TextMatrix(llRow, POPULATIONINDEX)
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = RATINGINDEX Then
                slSort = Trim$(Str$(gStrDecToLong(grdRsch.TextMatrix(llRow, RATINGINDEX), 2)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = AUDIENCEINDEX) Then
                slSort = grdRsch.TextMatrix(llRow, AUDIENCEINDEX)
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = DATEINDEX Then
                slSort = Trim$(Str$(gDateValue(grdRsch.TextMatrix(llRow, DATEINDEX))))
                Do While Len(slSort) < 10
                    slSort = "0" & slSort
                Loop
            End If
            slStr = grdRsch.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdRsch.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdRsch.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdRsch, POPULATIONINDEX, SORTINDEX, imLastColSorted, imLastSort
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
    If (grdRsch.Row < grdRsch.FixedRows) Or (grdRsch.Row >= grdRsch.Rows) Or (grdRsch.Col < grdRsch.FixedCols) Or (grdRsch.Col >= grdRsch.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdRsch.Row
    lmEnableCol = grdRsch.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdRsch.Left - pbcArrow.Width - 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + (grdRsch.RowHeight(grdRsch.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    Select Case grdRsch.Col
        Case POPULATIONINDEX
            edcDropDown.MaxLength = 10
            slStr = grdRsch.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdRsch.Row > grdRsch.FixedRows Then
                    slStr = grdRsch.TextMatrix(grdRsch.Row - 1, grdRsch.Col)
                End If
            End If
            edcDropDown.Text = slStr
        Case RATINGINDEX
            edcDropDown.MaxLength = 6
            slStr = grdRsch.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            edcDropDown.Text = slStr
        Case DATEINDEX
            edcDropDown.MaxLength = 10
            slStr = grdRsch.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdRsch.Row > grdRsch.FixedRows Then
                    slStr = grdRsch.TextMatrix(grdRsch.Row - 1, grdRsch.Col)
                    Select Case gWeekDayStr(slStr)
                        Case 0
                            slStr = gIncOneDay(slStr)
                        Case 1
                            slStr = gIncOneDay(slStr)
                        Case 2
                            slStr = gIncOneDay(slStr)
                        Case 3
                            slStr = gIncOneDay(slStr)
                        Case 4
                            slStr = gIncOneDay(slStr)
                            slStr = gIncOneDay(slStr)
                            slStr = gIncOneDay(slStr)
                        Case 5
                            slStr = ""
                        Case 6
                            slStr = ""
                    End Select
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

    If (grdRsch.Row < grdRsch.FixedRows) Or (grdRsch.Row >= grdRsch.Rows) Or (grdRsch.Col < grdRsch.FixedCols) Or (grdRsch.Col >= grdRsch.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdRsch.Left - pbcArrow.Width - 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + (grdRsch.RowHeight(grdRsch.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdRsch.Col - 1 Step 1
        llColPos = llColPos + grdRsch.ColWidth(ilCol)
    Next ilCol
    Select Case grdRsch.Col
        Case POPULATIONINDEX
            edcDropDown.Move grdRsch.Left + llColPos + 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + 30, grdRsch.ColWidth(grdRsch.Col) - 30, grdRsch.RowHeight(grdRsch.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case RATINGINDEX
            edcDropDown.Move grdRsch.Left + llColPos + 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + 30, grdRsch.ColWidth(grdRsch.Col) - 30, grdRsch.RowHeight(grdRsch.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case DATEINDEX
            edcDropDown.Move grdRsch.Left + llColPos + 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + 30, grdRsch.ColWidth(grdRsch.Col) - 30, grdRsch.RowHeight(grdRsch.Row) - 15
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
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slAudience As String
    Dim slPopulation As String
    Dim slRating As String

    pbcArrow.Visible = False
    If (lmEnableRow >= grdRsch.FixedRows) And (lmEnableRow < grdRsch.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case POPULATIONINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If StrComp(grdRsch.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imChg = True
                    grdRsch.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                    If Trim$(grdRsch.TextMatrix(lmEnableRow, RATINGINDEX)) <> "" Then
                        slRating = Trim$(grdRsch.TextMatrix(lmEnableRow, RATINGINDEX))
                        slAudience = gDivStr(gMulStr(slStr, slRating), "100.00")
                        grdRsch.TextMatrix(lmEnableRow, AUDIENCEINDEX) = slAudience
                    End If
                End If
                grdRsch.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case RATINGINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If gStrDecToLong(grdRsch.TextMatrix(lmEnableRow, lmEnableCol), 2) <> gStrDecToLong(slStr, 2) Then
                    imChg = True
                    grdRsch.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                    slPopulation = grdRsch.TextMatrix(lmEnableRow, POPULATIONINDEX)
                    slAudience = gDivStr(gMulStr(slPopulation, slStr), "100.00")
                    grdRsch.TextMatrix(lmEnableRow, AUDIENCEINDEX) = slAudience
                End If
                grdRsch.TextMatrix(lmEnableRow, lmEnableCol) = slStr
             Case DATEINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If gDateValue(grdRsch.TextMatrix(lmEnableRow, lmEnableCol)) <> gDateValue(slStr) Then
                    imChg = True
                    grdRsch.TextMatrix(lmEnableRow, CHGINDEX) = "1"
                End If
                grdRsch.TextMatrix(lmEnableRow, lmEnableCol) = slStr
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

    If GetFocus() <> pbcSTab.hWnd Then
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
            If grdRsch.Col = POPULATIONINDEX Then
                If grdRsch.Row > grdRsch.FixedRows Then
                    lmTopRow = -1
                    grdRsch.Row = grdRsch.Row - 1
                    If Not grdRsch.RowIsVisible(grdRsch.Row) Then
                        grdRsch.TopRow = grdRsch.TopRow - 1
                    End If
                    grdRsch.Col = DATEINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdRsch.Col = grdRsch.Col - 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdRsch.TopRow = grdRsch.FixedRows
        grdRsch.Col = POPULATIONINDEX
        grdRsch.Row = grdRsch.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdRsch.Col = DATEINDEX Then
                llRow = grdRsch.Rows
                Do
                    llRow = llRow - 1
                Loop While grdRsch.TextMatrix(llRow, POPULATIONINDEX) = ""
                llRow = llRow + 1
                If (grdRsch.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdRsch.Row = grdRsch.Row + 1
                    If Not grdRsch.RowIsVisible(grdRsch.Row) Or (grdRsch.Row - (grdRsch.TopRow - grdRsch.FixedRows) >= imInitNoRows) Then
                        imIgnoreScroll = True
                        grdRsch.TopRow = grdRsch.TopRow + 1
                    End If
                    grdRsch.Col = POPULATIONINDEX
                    'grdRsch.TextMatrix(grdRsch.Row, CODEINDEX) = 0
                    If Trim$(grdRsch.TextMatrix(grdRsch.Row, POPULATIONINDEX)) <> "" Then
                        If mColOk() Then
                            mEnableBox
                        Else
                            cmcCancel.SetFocus
                        End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdRsch.Left - pbcArrow.Width - 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + (grdRsch.RowHeight(grdRsch.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdRsch.TextMatrix(llEnableRow, POPULATIONINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdRsch.Row + 1 >= grdRsch.Rows Then
                            grdRsch.AddItem ""
                            grdRsch.RowHeight(grdRsch.Row + 1) = fgBoxGridH
                            grdRsch.TextMatrix(grdRsch.Row + 1, DRFCODEINDEX) = -1
                            grdRsch.TextMatrix(grdRsch.Row + 1, CHGINDEX) = "0"
                        End If
                        grdRsch.Row = grdRsch.Row + 1
                        grdRsch.Col = AUDIENCEINDEX
                        grdRsch.CellBackColor = LIGHTYELLOW
                        If (Not grdRsch.RowIsVisible(grdRsch.Row)) Or (grdRsch.Row - (grdRsch.TopRow - grdRsch.FixedRows) >= imInitNoRows) Then
                            imIgnoreScroll = True
                            grdRsch.TopRow = grdRsch.TopRow + 1
                        End If
                        grdRsch.Col = POPULATIONINDEX
                        grdRsch.TextMatrix(grdRsch.Row, DRFCODEINDEX) = -1
                        grdRsch.TextMatrix(grdRsch.Row, CHGINDEX) = "0"
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdRsch.Left - pbcArrow.Width - 30, grdRsch.Top + grdRsch.RowPos(grdRsch.Row) + (grdRsch.RowHeight(grdRsch.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdRsch.Col = grdRsch.Col + 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdRsch.TopRow = grdRsch.FixedRows
        grdRsch.Col = POPULATIONINDEX
        grdRsch.Row = grdRsch.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            cmcCancel.SetFocus
        End If
    End If
End Sub

Private Function mSaveRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilBookDone                                                                            *
'******************************************************************************************

    Dim llRow As Long
    Dim slMsg As String
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim ilError As Integer
    Dim slStr As String
    Dim ilDnfCode As Integer
    Dim llPopDrf As Long
    Dim llCustDrf As Long
    Dim ilLatestDnfCode As Integer
    Dim llLatestDnfDate As Long
    Dim llTestDate As Long
    Dim ilVef As Integer

    If Not imUpdateAllowed Then
        mSaveRec = False
        Exit Function
    End If
    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdRsch, grdRsch, vbHourglass
    For llRow = grdRsch.FixedRows To grdRsch.Rows - 1 Step 1
        If mGridFieldsOk(llRow) = False Then
            ilError = True
        End If
    Next llRow
    'Check that dates don't overlap
    If Not ilError Then
        If Not mCheckDates() Then
            ilError = True
        End If
    End If
    If (ilError) Or (Trim$(edcBookName.Text) = "") Then
        gSetMousePointer grdRsch, grdRsch, vbDefault
        Screen.MousePointer = vbDefault
        If ilError Then
            MsgBox "Check input fields marked in Red as in error", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        Else
            MsgBox "Book Name must be defined", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        End If
        Beep
        mSaveRec = False
        Exit Function
    End If
    ilLatestDnfCode = -1
    llLatestDnfDate = -1
    ilRet = btrBeginTrans(hmDrf, 1000)
    For llRow = grdRsch.FixedRows To grdRsch.Rows - 1 Step 1
        slStr = Trim$(grdRsch.TextMatrix(llRow, POPULATIONINDEX))
        If (slStr <> "") And (Val(slStr) > 0) Then
            llLoop = grdRsch.TextMatrix(llRow, DRFCODEINDEX)
            If llLoop >= 0 Then
                ilDnfCode = tmRschByCust(llLoop).iDnfCode
            Else
                ilDnfCode = 0
            End If
            If ilDnfCode <= 0 Then
                mInitDnf
                ilRet = BTRV_ERR_NONE
            Else
                tmDnfSrchKey.iCode = ilDnfCode
                ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            End If
            tmDnf.sBookName = Trim$(edcBookName.Text)
            slStr = grdRsch.TextMatrix(llRow, DATEINDEX)
            gPackDate slStr, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
            tmDnf.iUrfCode = tgUrf(0).iCode
            If ilDnfCode <= 0 Then
                tmDnf.sExactTime = "N"
                tmDnf.sSource = "M"
                ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
                Do
                    tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmDnf.iAutoCode = tmDnf.iCode
                    ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
            Else
                ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
            End If
            If ilLatestDnfCode = -1 Then
                ilLatestDnfCode = tmDnf.iCode
                gUnpackDateLong tmDnf.iBookDate(0), tmDnf.iBookDate(1), llLatestDnfDate
            Else
                gUnpackDateLong tmDnf.iBookDate(0), tmDnf.iBookDate(1), llTestDate
                If llTestDate > llLatestDnfDate Then
                    ilLatestDnfCode = tmDnf.iCode
                    gUnpackDateLong tmDnf.iBookDate(0), tmDnf.iBookDate(1), llLatestDnfDate
                End If
            End If
            If grdRsch.TextMatrix(llRow, CHGINDEX) = "1" Then
                Do
                    If llLoop >= 0 Then
                        llPopDrf = tmRschByCust(llLoop).lPopDrf
                    Else
                        llPopDrf = 0
                    End If
                    If llPopDrf <= 0 Then
                        mInitDrf "P"
                        ilRet = BTRV_ERR_NONE
                    Else
                        tmDrfSrchKey2.lCode = llPopDrf
                        ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                    End If
                    tmDrf.iDnfCode = tmDnf.iCode
                    mMoveCtrlToRec llRow, "P"
                    If llPopDrf <= 0 Then
                        ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
                        slMsg = "mSaveRec (btrInsert:DRF)"
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        Do
                            tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
                            tmDrf.lAutoCode = tmDrf.lCode
                            ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                    Else
                        ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
                        slMsg = "mSaveRec (btrUpdate:DRF)"
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, RschByCust
                    On Error GoTo 0
                End If
                Do
                    If llLoop >= 0 Then
                        llCustDrf = tmRschByCust(llLoop).lCustDrf
                    Else
                        llCustDrf = 0
                    End If
                    If llCustDrf <= 0 Then
                        mInitDrf "D"
                        ilRet = BTRV_ERR_NONE
                    Else
                        tmDrfSrchKey2.lCode = llCustDrf
                        ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                    End If
                    tmDrf.iDnfCode = tmDnf.iCode
                    mMoveCtrlToRec llRow, "D"
                    If llCustDrf <= 0 Then
                        ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
                        slMsg = "mSaveRec (btrInsert:DRF)"
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        Do
                            tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
                            tmDrf.lAutoCode = tmDrf.lCode
                            ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                    Else
                        ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
                        slMsg = "mSaveRec (btrUpdate:DRF)"
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, RschByCust
                    On Error GoTo 0
                End If
            End If
        End If
    Next llRow
    If ilLatestDnfCode <> -1 Then
        Do
            tmVefSrchKey.iCode = imVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                tmVef.iDnfCode = ilLatestDnfCode
                ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet = BTRV_ERR_NONE Then
            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If tgMVef(ilVef).iCode = imVefCode Then
                    tgMVef(ilVef).iDnfCode = ilLatestDnfCode
                    Exit For
                End If
            Next ilVef
            '11/26/17
            gFileChgdUpdate "vef.btr", False
        End If
    End If
    imChg = False
    ilRet = btrEndTrans(hmDrf)
    mSaveRec = True
    gSetMousePointer grdRsch, grdRsch, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    ilRet = btrAbortTrans(hmDrf)
    gSetMousePointer grdRsch, grdRsch, vbDefault
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
Private Sub mMoveCtrlToRec(llRow As Long, slType As String)
    Dim slStr As String
    Dim ilLoop As Integer

    If slType = "P" Then
        'tmDrf.lDemo(1) = Val(grdRsch.TextMatrix(llRow, POPULATIONINDEX))
        tmDrf.lDemo(0) = Val(grdRsch.TextMatrix(llRow, POPULATIONINDEX))
    Else
        'tmDrf.lDemo(1) = Val(grdRsch.TextMatrix(llRow, AUDIENCEINDEX))
        tmDrf.lDemo(0) = Val(grdRsch.TextMatrix(llRow, AUDIENCEINDEX))
    End If
    slStr = grdRsch.TextMatrix(llRow, DATEINDEX)
    For ilLoop = 0 To 6 Step 1
        tmDrf.sDay(ilLoop) = "N"
    Next ilLoop
    tmDrf.sDay(gWeekDayStr(slStr)) = "Y"
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
'*  llRet                         slStr                                                   *
'******************************************************************************************

    Dim llRow As Long
    Dim ilCol As Integer
    Dim llLoop As Long

    grdRsch.Redraw = False
    grdRsch.Rows = imInitNoRows
    For llRow = grdRsch.FixedRows To grdRsch.Rows - 1 Step 1
        grdRsch.RowHeight(llRow) = fgBoxGridH
        For ilCol = 0 To grdRsch.Cols - 1 Step 1
            If ilCol = DRFCODEINDEX Then
                grdRsch.TextMatrix(llRow, ilCol) = -1
            Else
                grdRsch.TextMatrix(llRow, ilCol) = ""
            End If
        Next ilCol
    Next llRow
    llRow = grdRsch.FixedRows

    For llLoop = 0 To UBound(tmRschByCust) - 1 Step 1
        If llRow >= grdRsch.Rows Then
            grdRsch.AddItem ""
            grdRsch.RowHeight(llRow) = fgBoxGridH
        End If
        grdRsch.TextMatrix(llRow, POPULATIONINDEX) = Trim$(Str$(tmRschByCust(llLoop).lPopDemo))
        grdRsch.TextMatrix(llRow, RATINGINDEX) = gDivStr(gMulStr(Trim$(Str$(tmRschByCust(llLoop).lCustDemo)), "100.00"), Trim$(Str$(tmRschByCust(llLoop).lPopDemo)))
        grdRsch.TextMatrix(llRow, AUDIENCEINDEX) = Trim$(Str$(tmRschByCust(llLoop).lCustDemo))
        grdRsch.Col = AUDIENCEINDEX
        grdRsch.CellBackColor = LIGHTYELLOW

        grdRsch.TextMatrix(llRow, DATEINDEX) = Trim$(tmRschByCust(llLoop).sDate)
        grdRsch.TextMatrix(llRow, DRFCODEINDEX) = llLoop
        llRow = llRow + 1
    Next llLoop
    If llRow >= grdRsch.Rows Then
        grdRsch.AddItem ""
        grdRsch.RowHeight(llRow) = fgBoxGridH
        grdRsch.TextMatrix(llRow, DRFCODEINDEX) = -1
    End If
    For llRow = grdRsch.FixedRows To grdRsch.Rows - 1 Step 1
        grdRsch.Row = llRow
        grdRsch.Col = AUDIENCEINDEX
        grdRsch.CellBackColor = LIGHTYELLOW
        grdRsch.TextMatrix(llRow, CHGINDEX) = "0"
    Next llRow
    'Remove highlight
    imLastColSorted = -1
    mSortCol DATEINDEX
    grdRsch.Row = 0
    grdRsch.Col = DRFCODEINDEX
    grdRsch.Redraw = True
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
Private Function mGridFieldsOk(llRowNo As Long) As Integer
'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStr As String
    Dim ilError As Integer

    ilError = False
    slStr = Trim$(grdRsch.TextMatrix(llRowNo, POPULATIONINDEX))
    If slStr <> "" Then
        slStr = Trim$(grdRsch.TextMatrix(llRowNo, DATEINDEX))
        If slStr = "" Then
            grdRsch.TextMatrix(llRowNo, DATEINDEX) = "Missing"
            ilError = True
            grdRsch.Row = llRowNo
            grdRsch.Col = DATEINDEX
            grdRsch.CellForeColor = vbRed
        Else
            If Not gValidDate(slStr) Then
                ilError = True
                grdRsch.Row = llRowNo
                grdRsch.Col = DATEINDEX
                grdRsch.CellForeColor = vbRed
            End If
        End If
    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function


Private Function mColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    If Not imUpdateAllowed Then
        mColOk = False
    Else
        mColOk = True
        If grdRsch.CellBackColor = LIGHTYELLOW Then
            mColOk = False
            Exit Function
        End If
    End If
End Function


Private Sub mInitDnf()
    tmDnf.iCode = 0
    tmDnf.sBookName = ""
    tmDnf.iBookDate(0) = 0
    tmDnf.iBookDate(1) = 0
    gPackDate Format$(gNow(), "m/d/yy"), tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
    tmDnf.iUrfCode = tgUrf(0).iCode
    tmDnf.sType = "C"
    tmDnf.iRemoteID = 0
    tmDnf.iAutoCode = 0
    gPackDate Format$(gNow(), "m/d/yy"), tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
    gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
    tmDnf.sForm = "8"
    tmDnf.iPopDnfCode = 0
    tmDnf.iQualPopDnfCode = 0
    tmDnf.sEstListenerOrUSA = "L"
    tmDnf.sUnused = ""
End Sub

Private Sub mInitDrf(slType As String)

    Dim ilLoop As Integer

    tmDrf.lCode = 0
    tmDrf.iDnfCode = 0
    tmDrf.sDemoDataType = slType
    tmDrf.iMnfSocEco = 0
    If slType = "D" Then
        tmDrf.iVefCode = imVefCode
    Else
        tmDrf.iVefCode = 0
    End If
    tmDrf.sInfoType = ""
    If slType = "D" Then
        tmDrf.sInfoType = "V"
    End If
    tmDrf.iRdfCode = 0
    tmDrf.sProgCode = ""
    tmDrf.iStartTime(0) = 1
    tmDrf.iStartTime(1) = 0
    tmDrf.iEndTime(0) = 1
    tmDrf.iEndTime(1) = 0
    For ilLoop = 0 To 6 Step 1
        tmDrf.sDay(ilLoop) = "N"
    Next ilLoop
    tmDrf.iStartTime2(0) = 1
    tmDrf.iStartTime2(1) = 0
    tmDrf.iEndTime2(0) = 1
    tmDrf.iEndTime2(1) = 0
    tmDrf.iCount = 0
    tmDrf.sExStdDP = "N"
                                             ' standard daypart (Y or N)="N"                As String * 1      ' sInfoType = T, then Excluded from
                                             ' report (Y or N)
    tmDrf.sDataType = "B"
    For ilLoop = 1 To 18 Step 1
        tmDrf.lDemo(ilLoop - 1) = 0
    Next ilLoop
    tmDrf.iDemoChgTime(0) = 0
    tmDrf.iDemoChgTime(1) = 0
    tmDrf.iRemoteID = 0
    tmDrf.lAutoCode = 0
    gPackDate Format$(gNow(), "m/d/yy"), tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
    gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
    tmDrf.lPopDrfCode = 0
    tmDrf.sForm = 8
                                             ' eighteen buckets (test for 8)
    tmDrf.sUnused = ""

End Sub

Private Function mCheckDates() As Integer
    Dim llRow1 As Long
    Dim llRow2 As Long
    Dim slStr As String
    Dim llDate1 As Long
    Dim llDate2 As Long

    mCheckDates = True
    For llRow1 = grdRsch.FixedRows To grdRsch.Rows - 1 Step 1
        slStr = Trim$(grdRsch.TextMatrix(llRow1, POPULATIONINDEX))
        If (slStr <> "") And (Val(slStr) > 0) Then
            llDate1 = gDateValue(grdRsch.TextMatrix(llRow1, DATEINDEX))
            For llRow2 = llRow1 + 1 To grdRsch.Rows - 1 Step 1
                slStr = Trim$(grdRsch.TextMatrix(llRow2, POPULATIONINDEX))
                If (slStr <> "") And (Val(slStr) > 0) Then
                    llDate2 = gDateValue(grdRsch.TextMatrix(llRow2, DATEINDEX))
                    If llDate1 = llDate2 Then
                        mCheckDates = False
                        grdRsch.Row = llRow2
                        grdRsch.Col = DATEINDEX
                        grdRsch.CellForeColor = vbRed
                    End If
                End If
            Next llRow2
        End If
    Next llRow1
End Function

Private Sub tmcStart_Timer()
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
    End If
    'Check if allowed to defined input
    'One vehicle only
    imVefCode = -1
    imUpdateAllowed = True
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "R") Then
            If imVefCode <> -1 Then
                MsgBox "More than one vehicle defined, update disallowed", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                imUpdateAllowed = False
                Exit For
            Else
                imVefCode = tgMVef(ilVef).iCode
            End If
        End If
    Next ilVef
    If imVefCode = -1 Then
        MsgBox "One vehicle must be defined, update disallowed", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        imUpdateAllowed = False
    End If
    'One custom daypart only
    If UBound(tmCDemoCode) <= LBound(tmCDemoCode) Then
        imUpdateAllowed = False
        MsgBox "One Custom Demo must be defined with Sort Code # of 1, update disallowed", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
    Else
        If UBound(tmCDemoCode) - 1 > LBound(tmCDemoCode) Then
            imUpdateAllowed = False
            MsgBox "Only One Custom Demo must be defined, update disallowed", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
        Else
            slNameCode = tmCDemoCode(LBound(tmCDemoCode)).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmMnfSrchKey.iCode = Val(slCode)
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                If tmMnf.iGroupNo <> 1 Then
                    imUpdateAllowed = False
                    MsgBox "The Custom Demo Sort Code # must be defined as 1, Update disallowed", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                End If
            End If
        End If
    End If
    If (igWinStatus(RESEARCHLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    End If
    If Not imUpdateAllowed Then
        cmcUpdate.Enabled = False
        grdRsch.Enabled = False
        edcBookName.Enabled = False
    End If
End Sub
