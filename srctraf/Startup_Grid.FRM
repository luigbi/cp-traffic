VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Startup_Grid 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4455
   ClientLeft      =   360
   ClientTop       =   3600
   ClientWidth     =   9570
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
   Icon            =   "Startup_Grid.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   9570
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9420
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   11
      Top             =   90
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   75
      Picture         =   "Startup_Grid.frx":08CA
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   60
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1425
      Width           =   60
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      HelpContextID   =   8
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1365
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   2925
      Width           =   15
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   1170
      Width           =   45
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
      Left            =   2835
      TabIndex        =   5
      Top             =   4005
      Width           =   1050
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
      Left            =   4245
      TabIndex        =   6
      Top             =   3990
      Width           =   1050
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      Left            =   5715
      TabIndex        =   7
      Top             =   3990
      Width           =   1050
   End
   Begin VB.ComboBox cbcSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   6165
      TabIndex        =   1
      Top             =   195
      Width           =   3000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdName 
      Height          =   3090
      Left            =   285
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   675
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
   Begin VB.Label plcScreen 
      Caption         =   "Client Form Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2325
   End
End
Attribute VB_Name = "Startup_Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Startup_Grid.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Import Rate Card input screen code
Option Explicit
Option Compare Text

Dim tmFileName As ANF        'ANF record image
Dim tmFileNameSrchKey As INTKEY0    'ANF key record image
Dim imFileNameRecLen As Integer        'ANF record length
Dim hmFileName As Integer 'File name file handle
Dim lmCode As Long

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long

Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim lmNowDate As Long
Dim imBypassFocus As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imUserInput As Integer
Dim imLBCtrls As Integer
Dim imBypassSetting As Integer
Dim imUpdateAllowed As Integer
Dim imInitNoRows As Integer
Dim imChgMode As Integer
Dim imComboBoxIndex As Integer
Dim imChg As Integer
Dim imFromArrow As Integer
Dim imIgnoreScroll As Integer

Dim imLastColSorted As Integer
Dim imLastSort As Integer

Const FIELDINDEX1 = 0
Const FIELDINDEX2 = 1
Const FIELDINDEX3 = 2
Const FIELDINDEX4 = 3
Const FIELDINDEX5 = 4
Const CODEINDEX = 5
Const SORTINDEX = 6



Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub Form_Activate()
    If (igWinStatus(AVAILNAMESLIST) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        grdName.Enabled = False
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
    Else
        grdName.Enabled = True
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
    End If
    Startup_Grid.Refresh

End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    btrStopAppl
    End
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
'
'   mInit
'   Where:
'
    Dim llNoRec As Long 'Number of records
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    imTerminate = False
    mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imLastColSorted = -1
    imLastSort = -1
    imLBCtrls = 1
    imFirstFocus = True
    imBypassFocus = False
    'pbcArrow.Picture = IconTraf!imcArrow.Picture
    'pbcArrow.Width = 90
    'pbcArrow.Height = 165
    imChgMode = False
    imBSMode = False
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gCenterStdAlone Startup_Grid
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
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
'
'   mInitBox
'   Where:
'
    mGridLayout
    mGridColumnWidths
    mGridColumns
    grdName.Move 180, plcScreen.Top + plcScreen.Height + 120, grdName.Width, cmcDone.Top - (plcScreen.Top + plcScreen.Height) - 240
    'grdName.Height = grdName.RowPos(0) + 14 * grdName.RowHeight(0) + fgPanelAdj - 15
    imInitNoRows = (cmcDone.Top - 120 - grdName.Top) \ fgFlexGridRowH
    grdName.Height = grdName.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj - 15
    gGrid_IntegralHeight grdName, CInt(fgFlexGridRowH)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    Dim ilPos As Integer
    Dim slChar As String
    slCommand = Command$
    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    sgCallAppName = ""
    slStr = "Guide"
    ilTestSystem = False
    ilPos = InStr(1, slCommand, "/S:Test", 1)
    If ilPos > 0 Then
        ilTestSystem = True
    End If
    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    If ilPos > 0 Then
        ilTestSystem = False
    End If


    gInitGlobalVar

    
    ilPos = InStr(1, slCommand, "/UserInput", 1)
    If ilPos > 0 Then
        imUserInput = True
        igBkgdProg = 0
    Else
        imUserInput = False
        igBkgdProg = 22
        Me.Visible = False
    End If

    'If password required
    If imUserInput Then
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If

    End If
    gInitStdAlone Startup_Grid, slStr, ilTestSystem
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
    igParentRestarted = False
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Startup_Grid
    Set Startup_Grid = Nothing   'Remove data segment
    igManUnload = NO
End Sub

Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim llRow As Long
    Dim llCol As Long
    
    grdName.RowHeight(0) = fgFlexGridRowH + 15
    grdName.Row = 0
    For llCol = FIELDINDEX1 To FIELDINDEX5 Step 1
        grdName.Col = llCol
        grdName.CellBackColor = WHITE
        grdName.CellForeColor = vbBlue
    Next llCol
    For llRow = grdName.FixedRows To grdName.Rows - 1 Step 1
        For llCol = FIELDINDEX1 To SORTINDEX Step 1
            grdName.TextMatrix(llRow, llCol) = ""
            'If llCol = CPMCOST Then
            '    grdName.Row = llRow
            '    grdName.Col = llCol
            '    grdName.CellBackColor = LIGHTYELLOW
            '    grdName.CellForeColor = vbBlack
            'End If
        Next llCol
        grdName.RowHeight(llRow) = fgFlexGridRowH
    Next llRow
End Sub
Private Sub mEnableBox()

    Dim slStr As String
    If (grdName.Row < grdName.FixedRows) Or (grdName.Row >= grdName.Rows) Or (grdName.Col < grdName.FixedCols) Or (grdName.Col >= grdName.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdName.Row
    lmEnableCol = grdName.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdName.Left - pbcArrow.Width - 30, grdName.Top + grdName.RowPos(grdName.Row) + (grdName.RowHeight(grdName.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    Select Case grdName.Col
        Case FIELDINDEX1
            edcDropdown.MaxLength = 30
            slStr = grdName.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdName.Row > grdName.FixedRows Then
                    slStr = grdName.TextMatrix(grdName.Row - 1, grdName.Col)
                End If
            End If
            edcDropdown.Text = slStr
        Case FIELDINDEX2
            edcDropdown.MaxLength = 30
            slStr = grdName.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdName.Row > grdName.FixedRows Then
                    slStr = grdName.TextMatrix(grdName.Row - 1, grdName.Col)
                End If
            End If
            edcDropdown.Text = slStr

    End Select
    mSetFocus
End Sub

Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdName.Row < grdName.FixedRows) Or (grdName.Row >= grdName.Rows) Or (grdName.Col < grdName.FixedCols) Or (grdName.Col >= grdName.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdName.Left - pbcArrow.Width - 30, grdName.Top + grdName.RowPos(grdName.Row) + (grdName.RowHeight(grdName.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdName.Col - 1 Step 1
        llColPos = llColPos + grdName.ColWidth(ilCol)
    Next ilCol
    Select Case grdName.Col
        Case FIELDINDEX1
            edcDropdown.Move grdName.Left + llColPos + 30, grdName.Top + grdName.RowPos(grdName.Row) + 30, grdName.ColWidth(grdName.Col) - 30, grdName.RowHeight(grdName.Row) - 15
            edcDropdown.Visible = True
            edcDropdown.SetFocus
        Case FIELDINDEX2
            edcDropdown.Move grdName.Left + llColPos + 30, grdName.Top + grdName.RowPos(grdName.Row) + 30, grdName.ColWidth(grdName.Col) - 30, grdName.RowHeight(grdName.Row) - 15
            edcDropdown.Visible = True
            edcDropdown.SetFocus
    End Select
    mSetCommands
End Sub

Private Sub mSetShow()
    Dim slStr As String

    pbcArrow.Visible = False
    If (lmEnableRow >= grdName.FixedRows) And (lmEnableRow < grdName.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case FIELDINDEX1
                edcDropdown.Visible = False  'Set visibility
                slStr = edcDropdown.Text
                If StrComp(grdName.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imChg = True
                End If
                grdName.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case FIELDINDEX2
                edcDropdown.Visible = False  'Set visibility
                slStr = edcDropdown.Text
                If StrComp(grdName.TextMatrix(lmEnableRow, lmEnableCol), slStr, vbTextCompare) <> 0 Then
                    imChg = True
                End If
                grdName.TextMatrix(lmEnableRow, lmEnableCol) = slStr
        End Select
    End If
    pbcArrow.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
End Sub




Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLp As Integer
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer

    'imPopReqd = False
    'If igANmCallSource = CALLSOURCEFEED Then
    '    ilfilter(0) = CHARFILTER
    '    slFilter(0) = "F"
    '    ilOffSet(0) = gFieldOffset("Anf", "AnfBookLocalFeed") '2
    'Else
    '    ilfilter(0) = NOFILTER
    '    slFilter(0) = ""
    '    ilOffSet(0) = 0
    'End If
    'ilRet = gIMoveListBox(AName, cbcSelect, tgNameCode(), sgNameCodeTag, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    'If ilRet <> CP_MSG_NOPOPREQ Then
    '    On Error GoTo mPopulateErr
    '    gCPErrorMsg ilRet, "mPopulate (gIMoveListBox)", AName
    '    On Error GoTo 0
    '    cbcSelect.AddItem "[New]", 0  'Force as first item on list
    '    imPopReqd = True
    'End If
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilAltered As Integer
    If imBypassSetting Then
        Exit Sub
    End If
    'Update button set if all mandatory fields have data and any field altered
    If (mTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (imChg) And (imUpdateAllowed) Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    If Not imChg Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim slStr As String
    Dim llRow As Long
    Dim llCol As Long
    For llRow = grdName.FixedRows To grdName.Rows - 1 Step 1
        For llCol = FIELDINDEX1 To FIELDINDEX5 Step 1
            slStr = Trim$(grdName.TextMatrix(llRow, FIELDINDEX1))
            If slStr <> "" Then
                Select Case llCol 'Branch on box type (control)
                    Case FIELDINDEX1
                        slStr = Trim$(grdName.TextMatrix(llRow, llCol))
                        'Test field
                    Case FIELDINDEX2
                        slStr = Trim$(grdName.TextMatrix(llRow, llCol))
                End Select
            End If
        Next llCol
    Next llRow
    mTestFields = YES
End Function

Private Sub pbcClickFocus_GotFocus()
    mSetShow
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
            If grdName.Col = FIELDINDEX1 Then
                If grdName.Row > grdName.FixedRows Then
                    lmTopRow = -1
                    grdName.Row = grdName.Row - 1
                    If Not grdName.RowIsVisible(grdName.Row) Then
                        grdName.TopRow = grdName.TopRow - 1
                    End If
                    grdName.Col = FIELDINDEX5
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdName.Col = grdName.Col - 1
                'If gColOk(grdName, grdName.Row, grdName.Col) Then
                    mEnableBox
                'Else
                '    ilPrev = True
                'End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdName.TopRow = grdName.FixedRows
        grdName.Col = FIELDINDEX1
        grdName.Row = grdName.FixedRows
        'If gColOk(grdName, grdName.Row, grdName.Col) Then
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

    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdName.Col = FIELDINDEX5 Then
                llRow = grdName.Rows
                Do
                    llRow = llRow - 1
                Loop While grdName.TextMatrix(llRow, FIELDINDEX1) = ""
                llRow = llRow + 1
                If (grdName.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdName.Row = grdName.Row + 1
                    If Not grdName.RowIsVisible(grdName.Row) Or (grdName.Row - (grdName.TopRow - grdName.FixedRows) >= imInitNoRows) Then
                        imIgnoreScroll = True
                        grdName.TopRow = grdName.TopRow + 1
                    End If
                    grdName.Col = FIELDINDEX1
                    'grdName.TextMatrix(grdName.Row, CODEINDEX) = 0
                    If Trim$(grdName.TextMatrix(grdName.Row, FIELDINDEX1)) <> "" Then
                        'If gColOk(grdName, grdName.Row, grdName.Col) Then
                            mEnableBox
                        'Else
                        '    cmcCancel.SetFocus
                        'End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdName.Left - pbcArrow.Width - 30, grdName.Top + grdName.RowPos(grdName.Row) + (grdName.RowHeight(grdName.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdName.TextMatrix(llEnableRow, FIELDINDEX1)) <> "" Then
                        lmTopRow = -1
                        If grdName.Row + 1 >= grdName.Rows Then
                            grdName.AddItem ""
                            grdName.RowHeight(grdName.Row + 1) = fgFlexGridRowH
                            grdName.TextMatrix(grdName.Row + 1, CODEINDEX) = 0
                        End If
                        grdName.Row = grdName.Row + 1
                        If (Not grdName.RowIsVisible(grdName.Row)) Or (grdName.Row - (grdName.TopRow - grdName.FixedRows) >= imInitNoRows) Then
                            imIgnoreScroll = True
                            grdName.TopRow = grdName.TopRow + 1
                        End If
                        grdName.Col = FIELDINDEX1
                        grdName.TextMatrix(grdName.Row, CODEINDEX) = 0
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdName.Left - pbcArrow.Width - 30, grdName.Top + grdName.RowPos(grdName.Row) + (grdName.RowHeight(grdName.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdName.Col = grdName.Col + 1
                'If gColOk(grdName, grdName.Row, grdName.Col) Then
                    mEnableBox
                'Else
                '    ilNext = True
                'End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdName.TopRow = grdName.FixedRows
        grdName.Col = FIELDINDEX1
        grdName.Row = grdName.FixedRows
        'If gColOk(grdName, grdName.Row, grdName.Col) Then
            mEnableBox
        'Else
        '    cmcCancel.SetFocus
        'End If
    End If
End Sub


Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim llRow As Long
    Dim slStr As String
    Dim ilRet As Integer
    Dim llRet As Long
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slEntryDate As String
    Dim slEntryTime As String
    Dim slSQLQuery As String
    Dim tlPcf As PCF

    mSetShow
    If mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    
    For llRow = grdName.FixedRows To grdName.Rows - 1 Step 1
        If (grdName.TextMatrix(llRow, FIELDINDEX1) <> "") Then
            slStr = grdName.TextMatrix(llRow, FIELDINDEX1)
            If (grdName.TextMatrix(llRow, CODEINDEX) = "") Or (Val(grdName.TextMatrix(llRow, CODEINDEX)) = "" = 0) Then
                'Insert

                slSQLQuery = "Insert Into FileName ( "
                slSQLQuery = slSQLQuery & "Code, "
                slSQLQuery = slSQLQuery & "Unused "
                slSQLQuery = slSQLQuery & ") "
                slSQLQuery = slSQLQuery & "Values ( "
                slSQLQuery = slSQLQuery & 0 & ", "
                slSQLQuery = slSQLQuery & ") "
                llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            Else
                'Update
                slSQLQuery = "Update pcf_Pod_CPM_Cntr Set "
                slSQLQuery = slSQLQuery & "Where Code = " & grdName.TextMatrix(llRow, CODEINDEX) & ", "
                llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            End If
            tgChfCntr.sPodCPMDefined = "Y"  'Set in mSetHdVeh
        End If
    Next llRow
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function

Private Sub grdName_EnterCell()
    mSetShow
End Sub

Private Sub grdName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdName.TopRow
    grdName.Redraw = False
End Sub

Private Sub grdName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer

    imIgnoreScroll = False
    If Y < grdName.RowHeight(0) Then
        grdName.Col = grdName.MouseCol
        mSortCol grdName.Col
        Exit Sub
    End If
    pbcArrow.Visible = False
    ilCol = grdName.MouseCol
    ilRow = grdName.MouseRow
    If ilCol < grdName.FixedCols Then
        grdName.Redraw = True
        Exit Sub
    End If
    If ilRow < grdName.FixedRows Then
        grdName.Redraw = True
        Exit Sub
    End If
    If grdName.TextMatrix(ilRow, FIELDINDEX1) = "" Then
        grdName.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdName.TextMatrix(ilRow, FIELDINDEX1) = ""
        grdName.Row = ilRow + 1
        grdName.Col = FIELDINDEX1
        grdName.Redraw = True
    Else
        grdName.Row = ilRow
        grdName.Col = ilCol
    End If
    grdName.Redraw = True
    lmTopRow = grdName.TopRow
    'If Not mColOk() Then
    '    pbcArrow.Move grdName.Left - pbcArrow.Width - 30, grdName.Top + grdName.RowPos(grdName.Row) + (grdName.RowHeight(grdName.Row) - pbcArrow.Height) / 2
    '    pbcArrow.Visible = True
    '    Exit Sub
    'End If

    mEnableBox
End Sub

Private Sub grdName_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdName.Redraw = False Then
        grdName.Redraw = True
        If lmTopRow < grdName.FixedRows Then
            grdName.TopRow = grdName.FixedRows
        Else
            grdName.TopRow = lmTopRow
        End If
        grdName.Refresh
        grdName.Redraw = False
    End If
    If (imCtrlVisible) And (grdName.Row >= grdName.FixedRows) And (grdName.Col >= 0) And (grdName.Col < grdName.Cols - 1) Then
        If grdName.RowIsVisible(grdName.Row) Then
            pbcArrow.Move grdName.Left - pbcArrow.Width - 30, grdName.Top + grdName.RowPos(grdName.Row) + (grdName.RowHeight(grdName.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            pbcSetFocus.SetFocus
            edcDropdown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If

End Sub

Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdName.Rows - 1 Step 1
        grdName.RowHeight(ilRow) = fgFlexGridRowH
    Next ilRow
    For ilCol = 0 To grdName.Cols - 1 Step 1
        grdName.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridColumns()

    grdName.Row = grdName.FixedRows - 1
    grdName.Col = FIELDINDEX1
    grdName.CellFontBold = False
    grdName.CellFontName = "Arial"
    grdName.CellFontSize = 6.75
    grdName.CellForeColor = vbBlue
    grdName.CellBackColor = LIGHTBLUE
    grdName.TextMatrix(grdName.Row, grdName.Col) = "Title1"
    grdName.Col = FIELDINDEX2
    grdName.CellFontBold = False
    grdName.CellFontName = "Arial"
    grdName.CellFontSize = 6.75
    grdName.CellForeColor = vbBlue
    grdName.CellBackColor = LIGHTBLUE
    grdName.TextMatrix(grdName.Row, grdName.Col) = "Title2"
    grdName.TextMatrix(grdName.Row, grdName.Col) = "Sort"

End Sub

Private Sub mGridColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdName.ColWidth(CODEINDEX) = 0
    grdName.ColWidth(SORTINDEX) = 0
    grdName.ColWidth(FIELDINDEX1) = 0.1 * grdName.Width
    grdName.ColWidth(FIELDINDEX2) = 0.34 * grdName.Width
    grdName.ColWidth(FIELDINDEX3) = 0.1 * grdName.Width
    grdName.ColWidth(FIELDINDEX4) = 0.34 * grdName.Width
    grdName.ColWidth(FIELDINDEX5) = 0.1 * grdName.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdName.Width
    For ilCol = 0 To grdName.Cols - 1 Step 1
        llWidth = llWidth + grdName.ColWidth(ilCol)
        If (grdName.ColWidth(ilCol) > 15) And (grdName.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdName.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdName.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdName.Width
            For ilCol = 0 To grdName.Cols - 1 Step 1
                If (grdName.ColWidth(ilCol) > 15) And (grdName.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdName.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdName.FixedCols To grdName.Cols - 1 Step 1
                If grdName.ColWidth(ilCol) > 15 Then
                    ilColInc = grdName.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdName.ColWidth(ilCol) = grdName.ColWidth(ilCol) + 15
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

    For llRow = grdName.FixedRows To grdName.Rows - 1 Step 1
        slStr = Trim$(grdName.TextMatrix(llRow, FIELDINDEX1))
        If slStr <> "" Then
            If ilCol = FIELDINDEX1 Then
                slSort = grdName.TextMatrix(llRow, FIELDINDEX1)
                Do While Len(slSort) < 5
                    slSort = slSort & " "
                Loop
            ElseIf ilCol = FIELDINDEX2 Then
                slSort = grdName.TextMatrix(llRow, FIELDINDEX2)
                Do While Len(slSort) < 30
                    slSort = slSort & " "
                Loop
            ElseIf (ilCol = FIELDINDEX3) Then
                slSort = Trim$(Str$(gStrDecToLong(grdName.TextMatrix(llRow, FIELDINDEX3), 4)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = FIELDINDEX4 Then
                slSort = grdName.TextMatrix(llRow, FIELDINDEX4)
                Do While Len(slSort) < 30
                    slSort = slSort & " "
                Loop
            ElseIf (ilCol = FIELDINDEX5) Then
                slSort = Trim$(Str$(gStrDecToLong(grdName.TextMatrix(llRow, FIELDINDEX5), 4)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            End If
            slStr = grdName.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdName.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdName.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdName, FIELDINDEX2, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub
Function mPopGrid() As Boolean
    Dim slSQLQuery As String
    Dim ilCol As Integer
    Dim llRow As Long
    Dim file_rst As ADODB.Recordset
    mPopGrid = False
    mClearCtrlFields
    grdName.Row = grdName.FixedRows - 1
    slSQLQuery = "Select * From FileName Where Code = " & lmCode
    Set file_rst = gSQLSelectCall(slSQLQuery)
    Do While Not file_rst.EOF
        If grdName.Row + 1 >= grdName.Rows Then
            grdName.AddItem ""
            grdName.RowHeight(grdName.Row + 1) = fgFlexGridRowH
            grdName.Row = grdName.Row + 1
            For ilCol = 0 To grdName.Cols - 1 Step 1
                grdName.ColAlignment(ilCol) = flexAlignLeftCenter
            Next ilCol
        Else
            grdName.Row = grdName.Row + 1
        End If
        llRow = grdName.Row
        grdName.TextMatrix(llRow, CODEINDEX) = file_rst!Code
        grdName.TextMatrix(llRow, FIELDINDEX1) = file_rst!Name
        grdName.TextMatrix(llRow, SORTINDEX) = ""
        file_rst.MoveNext
    Loop
    On Error Resume Next
    file_rst.Close
    mPopGrid = True
    Exit Function
End Function

Private Sub mSelectChange()
    Dim ilLoopCount As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer

    If imChgMode = False Then
        imChgMode = True
        ilLoopCount = 0
        Do
            Screen.MousePointer = vbHourglass  'Wait
            gSetMousePointer grdName, grdName, vbHourglass
            If ilLoopCount > 0 Then
                If cbcSelect.ListIndex >= 0 Then
                    cbcSelect.Text = cbcSelect.List(cbcSelect.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            If cbcSelect.Text <> "" Then
                gManLookAhead cbcSelect, imBSMode, imComboBoxIndex
            End If
            imSelectedIndex = cbcSelect.ListIndex
            lmCode = cbcSelect.ItemData(imSelectedIndex)
            mPopGrid
            Screen.MousePointer = vbHourglass  'Wait
            gSetMousePointer grdName, grdName, vbHourglass
        Loop While imSelectedIndex <> cbcSelect.ListIndex
        Screen.MousePointer = vbDefault    'Default
        gSetMousePointer grdName, grdName, vbDefault
        imChgMode = False
        mSetCommands
    End If
End Sub

Private Sub cbcSelect_Change()
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        If cbcSelect.Text <> "" Then
            gManLookAhead cbcSelect, imBSMode, imComboBoxIndex
            mSelectChange
        End If
    End If
    Exit Sub
End Sub

Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub

Private Sub cbcSelect_GotFocus()
    mSetShow
    If cbcSelect.Text = "" Then
        gFindMatch sgUserDefVehicleName, 0, cbcSelect
        If gLastFound(cbcSelect) >= 0 Then
            cbcSelect.ListIndex = gLastFound(cbcSelect)
        Else
            If cbcSelect.ListCount >= 1 Then
                cbcSelect.ListIndex = 0
            End If
        End If
        imComboBoxIndex = cbcSelect.ListIndex
        imSelectedIndex = imComboBoxIndex
    End If
    imComboBoxIndex = imSelectedIndex
    gCtrlGotFocus cbcSelect
End Sub

Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

