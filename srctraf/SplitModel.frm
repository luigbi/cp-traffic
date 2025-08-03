VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SplitModel 
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
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Don't Model"
      Height          =   285
      Left            =   4710
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   1770
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Model"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3270
      TabIndex        =   0
      Top             =   3720
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSplitModel 
      Height          =   3030
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   300
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   5345
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
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
      _Band(0).Cols   =   7
   End
   Begin VB.Label lacScreen 
      Caption         =   "Model Split from:"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   45
      Width           =   1650
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
Attribute VB_Name = "SplitModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SplitModel.frm on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SplitModel.Frm
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

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imCtrlKey As Integer
Dim lmSplitRowSelected As Long
Dim imLastSplitColSorted As Integer
Dim imLastSplitSort As Integer


Const ADVERTISERINDEX = 0
Const REGIONNAMEINDEX = 1
Const CATEGORYINDEX = 2
Const INCLEXCLINDEX = 3
Const STATUSINDEX = 4
Const RAFCODEINDEX = 5
Const SORTINDEX = 6








Private Sub cmcCancel_Click()
    igSplitModelReturn = 0
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
    If lmSplitRowSelected >= grdSplitModel.FixedRows Then
        igSplitModelReturn = 1
        lgSplitModelCodeRaf = CInt(grdSplitModel.TextMatrix(lmSplitRowSelected, RAFCODEINDEX))
    Else
        igSplitModelReturn = 0
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
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
    SplitModel.Refresh
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

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    btrExtClear hmRaf   'Clear any previous extend operation
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    
    Set SplitModel = Nothing   'Remove data segment

End Sub

Private Sub grdSplitModel_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRow                                                                                 *
'******************************************************************************************


    If grdSplitModel.Row >= grdSplitModel.FixedRows Then
        If grdSplitModel.TextMatrix(grdSplitModel.Row, REGIONNAMEINDEX) <> "" Then
            If (lmSplitRowSelected = grdSplitModel.Row) Then
                If imCtrlKey Then
                    lmSplitRowSelected = -1
                    grdSplitModel.Row = 0
                    grdSplitModel.Col = RAFCODEINDEX
                End If
            Else
                lmSplitRowSelected = grdSplitModel.Row
            End If
        Else
            lmSplitRowSelected = -1
            grdSplitModel.Row = 0
            grdSplitModel.Col = RAFCODEINDEX
        End If
    End If
    If lmSplitRowSelected >= grdSplitModel.FixedRows Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
    End If
End Sub

Private Sub grdSplitModel_DblClick()
    'If cmcSchedule.Enabled Then
    '    mSchedule
    'ElseIf cmcChgCntr.Enabled Then
    '    cmcChgCntr_Click
    'ElseIf cmcViewCntr.Enabled Then
    '    cmcViewCntr_Click
    'End If
End Sub

Private Sub grdSplitModel_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdSplitModel_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdSplitModel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRafCode As Long
    Dim llRow As Long

    If Y < grdSplitModel.RowHeight(0) Then
        Screen.MousePointer = vbHourglass
        llRafCode = -1
        If lmSplitRowSelected >= grdSplitModel.FixedRows Then
            If Trim$(grdSplitModel.TextMatrix(lmSplitRowSelected, REGIONNAMEINDEX)) <> "" Then
                llRafCode = grdSplitModel.TextMatrix(lmSplitRowSelected, RAFCODEINDEX)
            End If
        End If
        grdSplitModel.Col = grdSplitModel.MouseCol
        mSplitSortCol grdSplitModel.Col
        grdSplitModel.Row = 0
        grdSplitModel.Col = RAFCODEINDEX
        lmSplitRowSelected = -1
        If llRafCode <> -1 Then
            For llRow = grdSplitModel.FixedRows To grdSplitModel.Rows - 1 Step 1
                If llRafCode = grdSplitModel.TextMatrix(llRow, RAFCODEINDEX) Then
                    grdSplitModel.Row = llRow
                    grdSplitModel.RowSel = llRow
                    grdSplitModel.Col = ADVERTISERINDEX
                    grdSplitModel.ColSel = STATUSINDEX
                    lmSplitRowSelected = llRow
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Next llRow
        End If
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

    Screen.MousePointer = vbHourglass
    'mParseCmmdLine
    SplitModel.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone SplitModel
    'SplitModel.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    lmSplitRowSelected = -1

    imFirstFocus = True
    imCtrlKey = False
    lmSplitRowSelected = -1
    hmRaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", SplitModel
    On Error GoTo 0
    imRafRecLen = Len(tmRaf)  'Get and save CHF record length

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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload SplitModel
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
    If grdSplitModel.Visible Then
        lmSplitRowSelected = -1
        grdSplitModel.Row = 0
        grdSplitModel.Col = RAFCODEINDEX
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        llRet                                                   *
'******************************************************************************************


    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim slStr As String

    imRafRecLen = Len(tmRaf)

    grdSplitModel.Redraw = False
    'grdSplitModel.Rows = 14
    For llRow = grdSplitModel.FixedRows To grdSplitModel.Rows - 1 Step 1
        grdSplitModel.RowHeight(llRow) = fgBoxGridH + 15
        For ilCol = 0 To grdSplitModel.Cols - 1 Step 1
            If ilCol = RAFCODEINDEX Then
                grdSplitModel.TextMatrix(llRow, ilCol) = 0
            Else
                grdSplitModel.TextMatrix(llRow, ilCol) = ""
            End If
        Next ilCol
    Next llRow

    llRow = grdSplitModel.FixedRows


    ilRet = btrGetFirst(hmRaf, tmRaf, imRafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        If (igIncludeDormantSplits) Or ((Not igIncludeDormantSplits) And (tmRaf.sState <> "D")) Then
            If ((igSplitType = 0) And (tmRaf.sType = "C")) Or ((igSplitType = 1) And (tmRaf.sType = "N")) Then
                'ilAdf = gBinarySearchAdf(tmRaf.iAdfCode)
                If igSplitType = 0 Then
                    ilAdf = gBinarySearchAdf(tmRaf.iAdfCode)
                Else
                    ilAdf = 0
                End If
                If ilAdf <> -1 Then
                    If llRow >= grdSplitModel.Rows Then
                        grdSplitModel.AddItem ""
                        grdSplitModel.RowHeight(llRow) = fgBoxGridH + 15
                    End If
                    'grdSplitModel.TextMatrix(llRow, ADVERTISERINDEX) = Trim$(tgCommAdf(ilAdf).sName)
                    If igSplitType = 0 Then
                        grdSplitModel.TextMatrix(llRow, ADVERTISERINDEX) = Trim$(tgCommAdf(ilAdf).sName)
                    End If
                    grdSplitModel.TextMatrix(llRow, REGIONNAMEINDEX) = Trim$(tmRaf.sName)
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
                        Case "T"
                            slStr = "Time Zone"
                        Case "S"
                            slStr = "Station"
                        Case Else
                            slStr = ""
                    End Select
                    grdSplitModel.TextMatrix(llRow, CATEGORYINDEX) = slStr
                    If tmRaf.sInclExcl = "I" Then
                        grdSplitModel.TextMatrix(llRow, INCLEXCLINDEX) = "Include"
                    ElseIf tmRaf.sInclExcl = "E" Then
                        grdSplitModel.TextMatrix(llRow, INCLEXCLINDEX) = "Exclude"
                    Else
                        grdSplitModel.TextMatrix(llRow, INCLEXCLINDEX) = ""
                    End If
                    If tmRaf.sState = "A" Then
                        grdSplitModel.TextMatrix(llRow, STATUSINDEX) = "Active"
                    ElseIf tmRaf.sState = "D" Then
                        grdSplitModel.TextMatrix(llRow, STATUSINDEX) = "Dormant"
                    End If
                    grdSplitModel.TextMatrix(llRow, RAFCODEINDEX) = tmRaf.lCode
                    llRow = llRow + 1
                End If
            End If
        End If
        ilRet = btrGetNext(hmRaf, tmRaf, imRafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If igSplitType = 0 Then
        mSplitSortCol ADVERTISERINDEX
    Else
        mSplitSortCol REGIONNAMEINDEX
    End If
    grdSplitModel.Row = 0
    grdSplitModel.Col = RAFCODEINDEX
    lmSplitRowSelected = -1
    grdSplitModel.Redraw = True
End Sub



Private Sub mGridSplitLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdSplitModel.Rows - 1 Step 1
        grdSplitModel.RowHeight(ilRow) = fgBoxGridH + 15
    Next ilRow
    For ilCol = 0 To grdSplitModel.Cols - 1 Step 1
        grdSplitModel.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridSplitColumns()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                         ilValue                                                 *
'******************************************************************************************


    grdSplitModel.Row = grdSplitModel.FixedRows - 1
    grdSplitModel.Col = ADVERTISERINDEX
    grdSplitModel.CellFontBold = False
    grdSplitModel.CellFontName = "Arial"
    grdSplitModel.CellFontSize = 6.75
    grdSplitModel.CellForeColor = vbBlue
    grdSplitModel.CellBackColor = LIGHTBLUE
    grdSplitModel.TextMatrix(grdSplitModel.Row, grdSplitModel.Col) = "Advertiser"
    grdSplitModel.Col = REGIONNAMEINDEX
    grdSplitModel.CellFontBold = False
    grdSplitModel.CellFontName = "Arial"
    grdSplitModel.CellFontSize = 6.75
    grdSplitModel.CellForeColor = vbBlue
    grdSplitModel.CellBackColor = LIGHTBLUE
    grdSplitModel.TextMatrix(grdSplitModel.Row, grdSplitModel.Col) = "Region Name"
    grdSplitModel.Col = CATEGORYINDEX
    grdSplitModel.CellFontBold = False
    grdSplitModel.CellFontName = "Arial"
    grdSplitModel.CellFontSize = 6.75
    grdSplitModel.CellForeColor = vbBlue
    grdSplitModel.CellBackColor = LIGHTBLUE
    grdSplitModel.TextMatrix(grdSplitModel.Row, grdSplitModel.Col) = "Category"
    grdSplitModel.Col = INCLEXCLINDEX
    grdSplitModel.CellFontBold = False
    grdSplitModel.CellFontName = "Arial"
    grdSplitModel.CellFontSize = 6.75
    grdSplitModel.CellForeColor = vbBlue
    grdSplitModel.CellBackColor = LIGHTBLUE
    grdSplitModel.TextMatrix(grdSplitModel.Row, grdSplitModel.Col) = "Incl/Excl"
    grdSplitModel.Col = STATUSINDEX
    grdSplitModel.CellFontBold = False
    grdSplitModel.CellFontName = "Arial"
    grdSplitModel.CellFontSize = 6.75
    grdSplitModel.CellForeColor = vbBlue
    grdSplitModel.CellBackColor = LIGHTBLUE
    grdSplitModel.TextMatrix(grdSplitModel.Row, grdSplitModel.Col) = "Status"
    grdSplitModel.Col = RAFCODEINDEX
    grdSplitModel.CellFontBold = False
    grdSplitModel.CellFontName = "Arial"
    grdSplitModel.CellFontSize = 6.75
    grdSplitModel.CellForeColor = vbBlue
    grdSplitModel.CellBackColor = LIGHTBLUE
    grdSplitModel.TextMatrix(grdSplitModel.Row, grdSplitModel.Col) = "Raf Code"
    grdSplitModel.Col = SORTINDEX
    grdSplitModel.CellFontBold = False
    grdSplitModel.CellFontName = "Arial"
    grdSplitModel.CellFontSize = 6.75
    grdSplitModel.CellForeColor = vbBlue
    grdSplitModel.CellBackColor = LIGHTBLUE
    grdSplitModel.TextMatrix(grdSplitModel.Row, grdSplitModel.Col) = "Sort"

End Sub

Private Sub mGridSplitColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdSplitModel.ColWidth(RAFCODEINDEX) = 0
    grdSplitModel.ColWidth(SORTINDEX) = 0
    If igSplitType = 0 Then
        grdSplitModel.ColWidth(ADVERTISERINDEX) = 0.3 * grdSplitModel.Width
    Else
        grdSplitModel.ColWidth(ADVERTISERINDEX) = 0
    End If
    grdSplitModel.ColWidth(REGIONNAMEINDEX) = 0.3 * grdSplitModel.Width
    If igSplitType = 0 Then
        grdSplitModel.ColWidth(CATEGORYINDEX) = 0
        grdSplitModel.ColWidth(INCLEXCLINDEX) = 0
    Else
        grdSplitModel.ColWidth(CATEGORYINDEX) = 0.1 * grdSplitModel.Width
        grdSplitModel.ColWidth(INCLEXCLINDEX) = 0.1 * grdSplitModel.Width
    End If
    grdSplitModel.ColWidth(STATUSINDEX) = 0.1 * grdSplitModel.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdSplitModel.Width
    For ilCol = 0 To grdSplitModel.Cols - 1 Step 1
        llWidth = llWidth + grdSplitModel.ColWidth(ilCol)
        If (grdSplitModel.ColWidth(ilCol) > 15) And (grdSplitModel.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdSplitModel.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdSplitModel.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdSplitModel.Width
            For ilCol = 0 To grdSplitModel.Cols - 1 Step 1
                If (grdSplitModel.ColWidth(ilCol) > 15) And (grdSplitModel.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdSplitModel.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdSplitModel.FixedCols To grdSplitModel.Cols - 1 Step 1
                If grdSplitModel.ColWidth(ilCol) > 15 Then
                    ilColInc = grdSplitModel.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdSplitModel.ColWidth(ilCol) = grdSplitModel.ColWidth(ilCol) + 15
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

    grdSplitModel.Move 120, lacScreen.Top + lacScreen.Height + 90
    mGridSplitLayout
    mGridSplitColumnWidths
    mGridSplitColumns
    ilRow = grdSplitModel.FixedRows
    Do
        If ilRow + 1 > grdSplitModel.Rows Then
            grdSplitModel.AddItem ""
        End If
        grdSplitModel.RowHeight(ilRow) = fgBoxGridH + 15
        ilRow = ilRow + 1
    Loop While grdSplitModel.RowIsVisible(ilRow - 1)
    'grdSplitModel.Height = grdSplitModel.RowPos(0) + 14 * grdSplitModel.RowHeight(0) + fgPanelAdj - 15
    grdSplitModel.Height = cmcDone.Top - grdSplitModel.Top - 90
    gGrid_IntegralHeight grdSplitModel, CInt(fgBoxGridH + 30) ' + 15
    grdSplitModel.Height = grdSplitModel.Height + 15

End Sub



Private Sub mSplitSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdSplitModel.FixedRows To grdSplitModel.Rows - 1 Step 1
        If igSplitType = 0 Then
            slStr = Trim$(grdSplitModel.TextMatrix(llRow, ADVERTISERINDEX))
        Else
            slStr = Trim$(grdSplitModel.TextMatrix(llRow, REGIONNAMEINDEX))
        End If
        If slStr <> "" Then
            slSort = UCase$(Trim$(grdSplitModel.TextMatrix(llRow, ilCol)))
            slStr = grdSplitModel.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastSplitColSorted) Or ((ilCol = imLastSplitColSorted) And (imLastSplitSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdSplitModel.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdSplitModel.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastSplitColSorted Then
        imLastSplitColSorted = SORTINDEX
    Else
        imLastSplitColSorted = -1
        imLastSplitSort = -1
    End If
    If igSplitType = 0 Then
        gGrid_SortByCol grdSplitModel, ADVERTISERINDEX, SORTINDEX, imLastSplitColSorted, imLastSplitSort
    Else
        gGrid_SortByCol grdSplitModel, REGIONNAMEINDEX, SORTINDEX, imLastSplitColSorted, imLastSplitSort
    End If
    imLastSplitColSorted = ilCol
End Sub

