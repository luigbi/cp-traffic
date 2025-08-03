VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form RCItemDupl 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5400
   ClientLeft      =   900
   ClientTop       =   2430
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
   ScaleHeight     =   5400
   ScaleWidth      =   9315
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1650
      ScaleHeight     =   210
      ScaleWidth      =   315
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1740
      Visible         =   0   'False
      Width           =   885
   End
   Begin V81RateCard.CSI_ComboBoxList cbcSelect 
      Height          =   300
      Left            =   4995
      TabIndex        =   2
      Top             =   45
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   529
      BackColor       =   -2147483643
      ForeColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin VB.ListBox lbcVehicle 
      Height          =   2790
      ItemData        =   "RCItemDupl.frx":0000
      Left            =   6060
      List            =   "RCItemDupl.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   720
      Width           =   2805
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7680
      Top             =   4605
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   9
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
      TabIndex        =   6
      Top             =   345
      Width           =   30
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Duplicate"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5730
      TabIndex        =   14
      Top             =   4890
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4005
      TabIndex        =   13
      Top             =   4890
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
      TabIndex        =   15
      Top             =   3885
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2235
      TabIndex        =   12
      Top             =   4890
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDupl 
      Height          =   3090
      Left            =   210
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   705
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5450
      _Version        =   393216
      Rows            =   15
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
   Begin VB.Label lacModel 
      Caption         =   "Model From:"
      Height          =   240
      Left            =   3795
      TabIndex        =   1
      Top             =   75
      Width           =   1125
   End
   Begin VB.Label lacDaypart 
      Caption         =   "Dayparts to be Duplicated:"
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   420
      Width           =   2925
   End
   Begin VB.Label lacVehicle 
      Caption         =   "Vehicle to Duplicate Into:"
      Height          =   210
      Left            =   6060
      TabIndex        =   10
      Top             =   480
      Width           =   2790
   End
   Begin VB.Label lacScreen 
      Caption         =   "Duplicate Rate Card Dayparts"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   2415
   End
End
Attribute VB_Name = "RCItemDupl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RCItemDupl.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer

Private imTerminate As Integer  'True = terminating task, False= OK
Private imFirstFocus As Integer
Private imCtrlKey As Integer
Private bmInTimer As Boolean

Private imBoxNo As Integer
Private lmRowNo As Long
Private imChg As Integer

Private lmTopRow As Long
Private imInitNoRows As Integer

Private bmInGrid As Boolean

Private Rif_rst As ADODB.Recordset
Private Rdf_rst As ADODB.Recordset

Const LBONE = 1

Const DAYPARTINDEX = 0
Const INCLUDEINDEX = 1
Const BASEINDEX = 2
Const RPTINDEX = 3
Const SORTINDEX = 4
Const DOLLARROWINDEX = 5
Const RDFCODEINDEX = 6

'Taken from Rate Card
Const RCVEHINDEX = 1          'Vehicle control/field
Const RCDAYPARTINDEX = 2      'Daypart name control/field
'Const DOLLARINDEX = 3
'Const PCTINVINDEX = 4
Const RCCPMIndex = 3 'TTP 10609 jjb 2023-03-30
Const RCACQUISITIONINDEX = 4
Const RCBASEINDEX = 5
Const RCRPTINDEX = 6
Const RCSORTINDEX = 7
Const RCDOLLAR1INDEX = 8  '5
Const RCDOLLAR2INDEX = 9  '6
Const RCDOLLAR3INDEX = 10  '7
Const RCDOLLAR4INDEX = 11  '8
Const RCAVGINDEX = RCAVGRATEINDEX '9      'Also in StdPkg.Frm
Const RCTOTALINDEX = 13   '10
Dim imReturn As Integer


Private Sub cbcSelect_GotFocus()
    mSetShow imBoxNo
    lmRowNo = -1
    imBoxNo = -1
End Sub

Private Sub cbcSelect_OnChange()
    tmcDelay.Enabled = False
    DoEvents
    tmcDelay.Enabled = True
End Sub

Private Sub cmcCancel_Click()
    imReturn = 0
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    lmRowNo = -1
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If imChg Then
        If MsgBox("Duplicate all changes?", vbYesNo) = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    imReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    lmRowNo = -1
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUpdate_Click()
    Dim ilRet As Integer

    ilRet = mSaveRec()
    If ilRet Then
        tmcDelay_Timer
        imChg = False
    End If
End Sub

Private Sub cmcUpdate_GotFocus()
    mSetShow imBoxNo
    lmRowNo = -1
    imBoxNo = -1
    gCtrlGotFocus ActiveControl
End Sub


Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim ilRet As Integer
    Select Case imBoxNo
    End Select
End Sub

Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilKey As Integer
    Dim slStr As String
    Dim slComp As String

    Select Case imBoxNo
        Case SORTINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            slComp = "99"
            If gCompNumberStr(slStr, slComp) > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select

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
    RCItemDupl.Refresh
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
    If Screen.Width * 15 = 640 Then
        'fmAdjFactorW = 1#
        'fmAdjFactorH = 1#
    Else
        ''fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (50 * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        'fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    Rif_rst.Close
    Rdf_rst.Close
        
    Set RCItemDupl = Nothing   'Remove data segment
End Sub


Private Sub grdDupl_EnterCell()
    mSetShow imBoxNo
    lmRowNo = -1
    imBoxNo = -1
End Sub

Private Sub grdDupl_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdDupl_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdDupl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdDupl.RowHeight(0) Then
        Exit Sub
    End If
    bmInGrid = True
    llCurrentRow = grdDupl.MouseRow
    llCol = grdDupl.MouseCol
    If llCurrentRow < grdDupl.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdDupl.FixedRows Then
        If grdDupl.TextMatrix(llCurrentRow, DAYPARTINDEX) <> "" Then
            grdDupl.Row = llCurrentRow
            grdDupl.Col = llCol
            If mColOk() Then
                If Trim(grdDupl.TextMatrix(llCurrentRow, INCLUDEINDEX)) = "" Then
                    grdDupl.Col = INCLUDEINDEX
                End If
                lmRowNo = grdDupl.Row
                imBoxNo = grdDupl.Col
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    grdDupl.Row = 0
    grdDupl.Col = RDFCODEINDEX
    bmInGrid = False
    mSetCommands
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
    Dim ilRet As Integer

    imFirstActivate = True
    imTerminate = False
    gSetMousePointer grdDupl, grdDupl, vbHourglass
    'mParseCmmdLine
    'RCItemDupl.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone RCItemDupl
    'RCItemDupl.Show
    gSetMousePointer grdDupl, grdDupl, vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    
    lmRowNo = -1
    imBoxNo = -1
    imFirstFocus = True
    imCtrlKey = False
    bmInTimer = False
    
    pbcTab.Left = -pbcTab.Width - 100
    pbcSTab.Left = -pbcSTab.Width - 100
    pbcClickFocus.Left = -pbcClickFocus.Width - 100
    pbcSetFocus.Left = -pbcSetFocus.Width - 100
    
    mInitBox

    mPopulate
    imChg = False
    If imTerminate Then
        Exit Sub
    End If
    gSetMousePointer grdDupl, grdDupl, vbDefault
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
    gSetMousePointer grdDupl, grdDupl, vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RCItemDupl
    igManUnload = NO
End Sub




Private Sub lbcVehicle_GotFocus()
    mSetShow imBoxNo
    lmRowNo = -1
    imBoxNo = -1
End Sub

Private Sub pbcClickFocus_GotFocus()

    If imFirstFocus Then
        imFirstFocus = False
    End If
    mSetShow imBoxNo
    lmRowNo = -1
    imBoxNo = -1
    grdDupl.Row = 0
    grdDupl.Col = RDFCODEINDEX
    mSetCommands
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub mSetCommands()

    Dim ilRet As Integer

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

'
'   mInitBox
'   Where:
'
    'flTextHeight = pbcDates.TextHeight("1") - 35
    cbcSelect.Left = Me.Width - cbcSelect.Width - 120
    cbcSelect.Top = 30
    lacModel.Left = cbcSelect.Left - lacModel.Width
    lacDaypart.Left = 120
    lacDaypart.Top = 405
    grdDupl.Top = lacDaypart.Top + lacDaypart.Height
    grdDupl.Left = 120
    grdDupl.Width = (2 * Me.Width) / 3
    lacVehicle.Top = lacDaypart.Top
    lacVehicle.Left = grdDupl.Left + grdDupl.Width + 120
    lbcVehicle.Top = grdDupl.Top
    lbcVehicle.Left = lacVehicle.Left
    lbcVehicle.Width = Me.Width - lbcVehicle.Left - 240
    
    mGridDuplLayout
    mGridDuplColumnWidths
    mGridDuplColumns
    cmcDone.Top = Me.Height - (3 * cmcDone.Height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcUpdate.Top = cmcDone.Top
    cmcCancel.Left = Me.Width / 2 - cmcCancel.Width / 2
    cmcDone.Left = cmcCancel.Left - cmcCancel.Width - cmcDone.Width
    cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + cmcCancel.Width
    grdDupl.Height = cmcDone.Top - grdDupl.Top - 120
    ''grdDupl.Height = grdDupl.RowPos(0) + 14 * grdDupl.RowHeight(0) + fgPanelAdj - 15
    'imInitNoRows = (cmcDone.Top - 120 - grdDupl.Top) \ fgFlexGridRowH
    'grdDupl.Height = grdDupl.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj - 15
    gGrid_IntegralHeight grdDupl, CInt(fgBoxGridH + 30) ' + 15
    gGrid_FillWithRows grdDupl, fgBoxGridH + 15
    grdDupl.Height = grdDupl.Height + 30
    lbcVehicle.Height = grdDupl.Height
End Sub

Private Sub mGridDuplLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdDupl.rows - 1 Step 1
        grdDupl.RowHeight(ilRow) = fgFlexGridRowH
    Next ilRow
    For ilCol = 0 To grdDupl.cols - 1 Step 1
        grdDupl.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridDuplColumns()

    grdDupl.Row = grdDupl.FixedRows - 1
    grdDupl.Col = DAYPARTINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    'grdDupl.CellForeColor = vbBlue
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Daypart"
    grdDupl.Col = BASEINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    'grdDupl.CellForeColor = vbBlue
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Base"
    grdDupl.Col = RPTINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    'grdDupl.CellForeColor = vbBlue
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Rpt"
    grdDupl.Col = SORTINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    'grdDupl.CellForeColor = vbBlue
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Sort"
    grdDupl.Col = INCLUDEINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    'grdDupl.CellForeColor = vbBlue
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Include"
    grdDupl.Col = DOLLARROWINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    'grdDupl.CellForeColor = vbBlue
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Dollar Row"
    grdDupl.Col = RDFCODEINDEX
    grdDupl.CellFontBold = False
    grdDupl.CellFontName = "Arial"
    grdDupl.CellFontSize = 6.75
    'grdDupl.CellForeColor = vbBlue
    grdDupl.TextMatrix(grdDupl.Row, grdDupl.Col) = "Rdf Code"

End Sub

Private Sub mGridDuplColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdDupl.ColWidth(RDFCODEINDEX) = 0
    grdDupl.ColWidth(DOLLARROWINDEX) = 0
    grdDupl.ColWidth(DAYPARTINDEX) = 0.4 * grdDupl.Width
    grdDupl.ColWidth(BASEINDEX) = 0.1 * grdDupl.Width
    grdDupl.ColWidth(RPTINDEX) = 0.1 * grdDupl.Width
    grdDupl.ColWidth(SORTINDEX) = 0.1 * grdDupl.Width
    grdDupl.ColWidth(INCLUDEINDEX) = 0.1 * grdDupl.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdDupl.Width
    For ilCol = 0 To grdDupl.cols - 1 Step 1
        llWidth = llWidth + grdDupl.ColWidth(ilCol)
        If (grdDupl.ColWidth(ilCol) > 15) And (grdDupl.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdDupl.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdDupl.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdDupl.Width
            For ilCol = 0 To grdDupl.cols - 1 Step 1
                If (grdDupl.ColWidth(ilCol) > 15) And (grdDupl.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdDupl.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdDupl.FixedCols To grdDupl.cols - 1 Step 1
                If grdDupl.ColWidth(ilCol) > 15 Then
                    ilColInc = grdDupl.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdDupl.ColWidth(ilCol) = grdDupl.ColWidth(ilCol) + 15
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




Private Sub pbcSetFocus_GotFocus()
    tmcDelay.Enabled = False
    tmcDelay_Timer
    pbcClickFocus.SetFocus
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer

    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If bmInTimer Then
        Exit Sub
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            lmRowNo = 1
            ilBox = 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case INCLUDEINDEX 'Name (first control within header)
            mSetShow imBoxNo
            If lmRowNo <= grdDupl.FixedCols Then
                pbcClickFocus.SetFocus
                Exit Sub
            Else
                ilBox = SORTINDEX
                If Not grdDupl.RowIsVisible(grdDupl.Row - 1) Then
                    grdDupl.TopRow = grdDupl.TopRow - 1
                End If
                lmRowNo = lmRowNo - 1
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
    
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long
    Dim ilBox As Integer
    Dim blFound As Boolean
    
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            lmRowNo = grdDupl.FixedRows
            ilBox = INCLUDEINDEX
        Case INCLUDEINDEX, SORTINDEX
            If (imBoxNo = INCLUDEINDEX) And (Trim$(grdDupl.TextMatrix(lmRowNo, INCLUDEINDEX)) = "Yes") Then
                ilBox = imBoxNo + 1
            Else
                mSetShow imBoxNo
                If (lmRowNo + 1 >= grdDupl.rows) Then
                    imBoxNo = -1
                    lmRowNo = -1
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
                If (Trim$(grdDupl.TextMatrix(lmRowNo + 1, DAYPARTINDEX)) = "") Then
                    imBoxNo = -1
                    lmRowNo = -1
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
                If Not grdDupl.RowIsVisible(grdDupl.Row + 1) Then
                    grdDupl.TopRow = grdDupl.TopRow + 1
                End If
                grdDupl.Row = grdDupl.Row + 1
                lmRowNo = lmRowNo + 1
                ilBox = INCLUDEINDEX
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub

Private Function mSaveRec() As Integer
    Dim llRow As Long
    Dim slVehName As String
    Dim slDPName As String
    Dim ilVef As Integer
    Dim blFound As Boolean
    Dim llRowNo As Long
    Dim ilLoop As Integer
    
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdDupl, grdDupl, vbHourglass
    RateCard.edcDropDown.MaxLength = 0
    For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilVef) Then
            slVehName = lbcVehicle.List(ilVef)
            For llRow = grdDupl.FixedRows To grdDupl.rows - 1 Step 1
                If grdDupl.TextMatrix(llRow, DAYPARTINDEX) <> "" Then
                    slDPName = grdDupl.TextMatrix(llRow, DAYPARTINDEX)
                    If grdDupl.TextMatrix(llRow, INCLUDEINDEX) = "Yes" Then
                        blFound = False
                        For ilLoop = LBONE To UBound(tmRifRec) - 1 Step 1
                            If smRCSave(1, ilLoop) = slVehName Then
                                If smRCSave(2, ilLoop) = grdDupl.TextMatrix(llRow, DAYPARTINDEX) Then
                                    blFound = True
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                        If Not blFound Then
                            llRowNo = UBound(tmRifRec)
                            RateCard.mAddNewRow llRowNo + 1
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCVEHINDEX
                            RateCard.edcDropDown.Text = slVehName
                            RateCard.mRCSetShow RCVEHINDEX
                        
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCDAYPARTINDEX
                            RateCard.edcDropDown.Text = slDPName
                            RateCard.mRCSetShow RCDAYPARTINDEX
                            
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCACQUISITIONINDEX
                            RateCard.edcDropDown.Text = ""
                            RateCard.mRCSetShow RCACQUISITIONINDEX
                            
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCBASEINDEX
                            If grdDupl.TextMatrix(llRow, BASEINDEX) = "Yes" Then
                                smRCSave(RCBASEINDEX, llRowNo) = "Y"
                            Else
                                smRCSave(RCBASEINDEX, llRowNo) = "N"
                            End If
                            RateCard.mRCSetShow RCBASEINDEX
                            
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCRPTINDEX
                            If grdDupl.TextMatrix(llRow, RPTINDEX) = "Yes" Then
                                smRCSave(RCRPTINDEX, llRowNo) = "Y"
                            Else
                                smRCSave(RCRPTINDEX, llRowNo) = "N"
                            End If
                            RateCard.mRCSetShow RCRPTINDEX
                            
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCSORTINDEX
                            RateCard.edcDropDown.Text = grdDupl.TextMatrix(llRow, SORTINDEX)
                            RateCard.mRCSetShow RCSORTINDEX
                                                        
                            If Trim$(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)) <> "" Then
                                mCopyRate Val(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)), llRowNo
                            End If
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCDOLLAR1INDEX
                            If Trim$(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)) = "" Then
                                RateCard.edcDropDown.Text = ""
                            Else
                                RateCard.edcDropDown.Text = Trim$(Str$(lmRCSave(RCDOLLAR1INDEX - RCDOLLAR1INDEX + 1, Val(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)))))
                            End If
                            RateCard.mRCSetShow RCDOLLAR1INDEX
                            
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCDOLLAR2INDEX
                            If (Trim$(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)) = "") Or (igRCNoDollarColumns < 2) Then
                                RateCard.edcDropDown.Text = ""
                            Else
                                RateCard.edcDropDown.Text = Trim$(Str$(lmRCSave(RCDOLLAR2INDEX - RCDOLLAR1INDEX + 1, Val(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)))))
                            End If
                            RateCard.mRCSetShow RCDOLLAR2INDEX
                            
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCDOLLAR3INDEX
                            If (Trim$(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)) = "") Or (igRCNoDollarColumns < 3) Then
                                RateCard.edcDropDown.Text = ""
                            Else
                                RateCard.edcDropDown.Text = Trim$(Str$(lmRCSave(RCDOLLAR3INDEX - RCDOLLAR1INDEX + 1, Val(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)))))
                            End If
                            RateCard.mRCSetShow RCDOLLAR3INDEX
                            
                            RateCard.lmRCRowNo = llRowNo
                            RateCard.imRCBoxNo = RCDOLLAR4INDEX
                            If (Trim$(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)) = "") Or (igRCNoDollarColumns < 4) Then
                                RateCard.edcDropDown.Text = ""
                            Else
                                RateCard.edcDropDown.Text = Trim$(Str$(lmRCSave(RCDOLLAR4INDEX - RCDOLLAR1INDEX + 1, Val(grdDupl.TextMatrix(llRow, DOLLARROWINDEX)))))
                            End If
                            RateCard.mRCSetShow RCDOLLAR4INDEX

                            
                        End If
                    End If
                End If
            Next llRow
        End If
    Next ilVef
    
    imChg = False
    mSaveRec = True
    gSetMousePointer grdDupl, grdDupl, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    gSetMousePointer grdDupl, grdDupl, vbDefault
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
Private Function mColOk() As Integer
'
'   iRet = mColOk()
'   Where:
'       iRet (O)- True if fields
'
'
    Dim slStr As String
    Dim ilError As Integer

    slStr = Trim$(grdDupl.TextMatrix(grdDupl.Row, DAYPARTINDEX))
    If slStr = "" Then
        mColOk = False
    Else
        If grdDupl.CellBackColor = LIGHTYELLOW Then
            mColOk = False
        Else
            mColOk = True
        End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slVehName As String
    Dim slCode As String
    Dim ilIndex As Integer
    Dim ilVefCode As Integer
    Dim ilPrevVefCode As Integer
    Dim ilBypass As Integer
    Dim ilVef As Integer
    
    cbcSelect.FontBold = True
    cbcSelect.Clear
    cbcSelect.SetDropDownWidth (cbcSelect.Width)
    ilPrevVefCode = -1
    For ilLoop = LBONE To UBound(tmRifRec) - 1 Step 1
        'Vehicle
        gFindMatch Trim$(smRCSave(1, ilLoop)), 0, RateCard!lbcVehicle
        ilIndex = gLastFound(RateCard!lbcVehicle)
        If ilIndex >= 0 Then
            slNameCode = tgRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
            ilRet = gParseItem(slNameCode, 1, "\", slVehName)
            ilRet = gParseItem(slVehName, 3, "|", slVehName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If ilVefCode = ilPrevVefCode Then
                ilBypass = True
            Else
                ilPrevVefCode = ilVefCode
                'Test if Package vehicle- if so, bypass
                ilBypass = False
                ilVef = gBinarySearchVef(ilVefCode)
                If ilVef <> -1 Then
                    If (tgMVef(ilVef).sType = "P") Then
                        ilBypass = True
                    End If
                Else
                    ilBypass = True
                End If
            End If
        Else
        End If
        If Not ilBypass Then
            cbcSelect.AddItem slVehName
            cbcSelect.SetItemData = ilVefCode
        End If
    Next ilLoop
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    
    grdDupl.Row = llRow
    For llCol = DAYPARTINDEX To INCLUDEINDEX Step 1
        grdDupl.Col = llCol
        If llCol = DAYPARTINDEX Then
            grdDupl.CellBackColor = LIGHTYELLOW
        End If
    Next llCol
End Sub

Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    Dim ilIndex As Integer
    If imBoxNo = BASEINDEX Then
        ilIndex = BASEINDEX
    ElseIf imBoxNo = RPTINDEX Then
        ilIndex = RPTINDEX
    ElseIf imBoxNo = INCLUDEINDEX Then
        ilIndex = INCLUDEINDEX
    Else
        Exit Sub
    End If
    If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
        imChg = True
        grdDupl.TextMatrix(lmRowNo, ilIndex) = "Yes"
        pbcYN_Paint
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        imChg = True
        grdDupl.TextMatrix(lmRowNo, ilIndex) = "No"
        pbcYN_Paint
    End If
    If KeyAscii = Asc(" ") Then
        If Trim$(grdDupl.TextMatrix(lmRowNo, ilIndex)) = "Yes" Then
            imChg = True
            grdDupl.TextMatrix(lmRowNo, ilIndex) = "No"
            pbcYN_Paint
        ElseIf Trim$(grdDupl.TextMatrix(lmRowNo, ilIndex)) = "No" Then
            imChg = True
            grdDupl.TextMatrix(lmRowNo, ilIndex) = "Yes"
            pbcYN_Paint
        End If
    End If

End Sub

Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilIndex As Integer
    If imBoxNo = BASEINDEX Then
        ilIndex = BASEINDEX
    ElseIf imBoxNo = RPTINDEX Then
        ilIndex = RPTINDEX
    ElseIf imBoxNo = INCLUDEINDEX Then
        ilIndex = INCLUDEINDEX
    Else
        Exit Sub
    End If
    If Trim$(grdDupl.TextMatrix(lmRowNo, ilIndex)) = "Yes" Then
        imChg = True
        grdDupl.TextMatrix(lmRowNo, ilIndex) = "No"
    Else
        imChg = True
        grdDupl.TextMatrix(lmRowNo, ilIndex) = "Yes"
    End If
    pbcYN_Paint

End Sub

Private Sub pbcYN_Paint()
    Dim ilIndex As Integer
    If imBoxNo = BASEINDEX Then
        ilIndex = BASEINDEX
    ElseIf imBoxNo = RPTINDEX Then
        ilIndex = RPTINDEX
    ElseIf imBoxNo = INCLUDEINDEX Then
        ilIndex = INCLUDEINDEX
    Else
        Exit Sub
    End If
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    pbcYN.Print grdDupl.TextMatrix(lmRowNo, ilIndex)
End Sub

Private Sub tmcDelay_Timer()
    Dim slRdfCode As String
    Dim slNameCode As String
    Dim slVehName As String
    Dim slDPName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVefCode As Integer
    Dim ilRdfCode As Integer
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilVef As Integer
    Dim slSQLQuery As String
    
    tmcDelay.Enabled = False
    bmInTimer = True
    gSetMousePointer grdDupl, grdDupl, vbHourglass
    grdDupl.Redraw = False
    mClearGrid
    If cbcSelect.ListIndex < 0 Then
        bmInTimer = False
        Exit Sub
    End If
    slRdfCode = ""
    llRow = grdDupl.FixedRows
    ilVefCode = cbcSelect.GetItemData(cbcSelect.ListIndex)
    For ilLoop = LBONE To UBound(tmRifRec) - 1 Step 1
        'Vehicle
        gFindMatch Trim$(smRCSave(1, ilLoop)), 0, RateCard!lbcVehicle
        ilIndex = gLastFound(RateCard!lbcVehicle)
        If ilIndex >= 0 Then
            slNameCode = tgRCUserVehicle(ilIndex).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
            ilRet = gParseItem(slNameCode, 1, "\", slVehName)
            ilRet = gParseItem(slVehName, 3, "|", slVehName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilVefCode = Val(slCode) Then
                gFindMatch Trim$(smRCSave(2, ilLoop)), 0, RateCard!lbcDPName
                ilIndex = gLastFound(RateCard!lbcDPName)
                If ilIndex >= 0 Then
                    imChg = True
                    slNameCode = RateCard!lbcDPNameCode.List(ilIndex)
                    ilRet = gParseItem(slNameCode, 1, "\", slDPName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilRdfCode = Val(slCode)
                    slRdfCode = slRdfCode & slCode & ","
                    'Values saved (1=Vehicle; 2=Daypart; 3=Acquisition, 4=Base; 5=Report; 6=Sort)
                    If llRow >= grdDupl.rows Then
                        grdDupl.AddItem ""
                        grdDupl.RowHeight(llRow) = fgFlexGridRowH
                    End If
                    grdDupl.TextMatrix(llRow, DAYPARTINDEX) = Trim$(smRCSave(2, ilLoop))
                    grdDupl.TextMatrix(llRow, INCLUDEINDEX) = "Yes"
                    If Trim$(smRCSave(4, ilLoop)) = "Y" Then
                        grdDupl.TextMatrix(llRow, BASEINDEX) = "Yes"
                    Else
                        grdDupl.TextMatrix(llRow, BASEINDEX) = "No"
                    End If
                    If Trim$(smRCSave(5, ilLoop)) = "Y" Then
                        grdDupl.TextMatrix(llRow, RPTINDEX) = "Yes"
                    Else
                        grdDupl.TextMatrix(llRow, RPTINDEX) = "No"
                    End If
                    grdDupl.TextMatrix(llRow, SORTINDEX) = Trim$(smRCSave(6, ilLoop))
                    grdDupl.TextMatrix(llRow, DOLLARROWINDEX) = ilLoop
                    grdDupl.TextMatrix(llRow, RDFCODEINDEX) = ilRdfCode
                    mPaintRowColor llRow
                    llRow = llRow + 1
                End If
            End If
        End If
    Next ilLoop
'    slSQLQuery = "Select * from RDF_Standard_Daypart Where rdfCode <> " & ilRdfCode & " Order By rdfName"
    If slRdfCode <> "" Then
        slRdfCode = Left(slRdfCode, Len(slRdfCode) - 1)
        slSQLQuery = "Select * from RDF_Standard_Daypart Where rdfCode Not In (" & slRdfCode & ")" & " Order By rdfName"
    Else
        slSQLQuery = "Select * from RDF_Standard_Daypart " & " Order By rdfName"
    End If
    Set Rdf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not Rdf_rst.EOF
        If llRow >= grdDupl.rows Then
            grdDupl.AddItem ""
            grdDupl.RowHeight(llRow) = fgFlexGridRowH
        End If
        grdDupl.TextMatrix(llRow, DAYPARTINDEX) = Rdf_rst!rdfName
        grdDupl.TextMatrix(llRow, INCLUDEINDEX) = ""
        grdDupl.TextMatrix(llRow, BASEINDEX) = ""
        grdDupl.TextMatrix(llRow, RPTINDEX) = ""
        grdDupl.TextMatrix(llRow, SORTINDEX) = ""
        grdDupl.TextMatrix(llRow, DOLLARROWINDEX) = ""
        grdDupl.TextMatrix(llRow, RDFCODEINDEX) = Rdf_rst!rdfCode
        mPaintRowColor llRow
        llRow = llRow + 1
        Rdf_rst.MoveNext
    Loop
    lbcVehicle.Clear
    For ilLoop = 0 To RateCard.lbcVehicle.ListCount - 1 Step 1
        slNameCode = tgRCUserVehicle(ilLoop).sKey   'Traffic!lbcUserVehicle.List(ilIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Val(slCode) <> ilVefCode Then
            ilVef = gBinarySearchVef(Val(slCode))
            If ilVef <> -1 Then
                If (tgMVef(ilVef).sType <> "P") Then
                    lbcVehicle.AddItem RateCard.lbcVehicle.List(ilLoop)
                    lbcVehicle.ItemData(lbcVehicle.NewIndex) = Val(slCode)
                End If
            End If
        End If
    Next ilLoop
    mSetCommands
    bmInTimer = False
    grdDupl.Redraw = True
    gSetMousePointer grdDupl, grdDupl, vbDefault
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long

    'Blank rows within grid
    grdDupl.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdDupl.FixedRows To grdDupl.rows - 1 Step 1
        grdDupl.Row = llRow
        For llCol = DAYPARTINDEX To RDFCODEINDEX Step 1
            grdDupl.Col = llCol
            If llCol = DAYPARTINDEX Then
                grdDupl.CellBackColor = LIGHTYELLOW
            Else
                grdDupl.CellBackColor = WHITE
            End If
            grdDupl.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdDupl.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mRCEnableBox                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mRCEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilCode As Integer
    Dim llRowNo As Long
    
    If (ilBoxNo < INCLUDEINDEX) Or (ilBoxNo > SORTINDEX) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case INCLUDEINDEX
            If (Trim$(grdDupl.TextMatrix(lmRowNo, INCLUDEINDEX)) <> "1") And (Trim$(grdDupl.TextMatrix(lmRowNo, INCLUDEINDEX)) <> "No") Then
                grdDupl.TextMatrix(lmRowNo, INCLUDEINDEX) = "No"
                imChg = True
            End If
            pbcYN.Move grdDupl.Left + grdDupl.ColPos(ilBoxNo) + 30, grdDupl.Top + grdDupl.RowPos(grdDupl.Row) + 30, grdDupl.ColWidth(ilBoxNo) - 30, grdDupl.RowHeight(grdDupl.Row) - 15
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case BASEINDEX
            If (Trim$(grdDupl.TextMatrix(lmRowNo, BASEINDEX)) <> "Yes") And (Trim$(grdDupl.TextMatrix(lmRowNo, BASEINDEX)) <> "No") Then
                grdDupl.TextMatrix(lmRowNo, BASEINDEX) = "No"
                imChg = True
            End If
            pbcYN.Move grdDupl.Left + grdDupl.ColPos(ilBoxNo) + 30, grdDupl.Top + grdDupl.RowPos(grdDupl.Row) + 30, grdDupl.ColWidth(ilBoxNo) - 30, grdDupl.RowHeight(grdDupl.Row) - 15
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case RPTINDEX
            If (Trim$(grdDupl.TextMatrix(lmRowNo, RPTINDEX)) <> "Yes") And (Trim$(grdDupl.TextMatrix(lmRowNo, RPTINDEX)) <> "No") Then
                grdDupl.TextMatrix(lmRowNo, RPTINDEX) = "Yes"
                imChg = True
            End If
            pbcYN.Move grdDupl.Left + grdDupl.ColPos(ilBoxNo) + 30, grdDupl.Top + grdDupl.RowPos(grdDupl.Row) + 30, grdDupl.ColWidth(ilBoxNo) - 30, grdDupl.RowHeight(grdDupl.Row) - 15
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case SORTINDEX
            edcDropDown.MaxLength = 2
            edcDropDown.Move grdDupl.Left + grdDupl.ColPos(ilBoxNo) + 30, grdDupl.Top + grdDupl.RowPos(grdDupl.Row) + 30, grdDupl.ColWidth(ilBoxNo) - 30, grdDupl.RowHeight(grdDupl.Row) - 15
            edcDropDown.Text = Trim$(grdDupl.TextMatrix(lmRowNo, SORTINDEX))
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRCSetShow                      *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mRCSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    
    If (ilBoxNo < INCLUDEINDEX) Or (ilBoxNo > SORTINDEX) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case INCLUDEINDEX
            pbcYN.Visible = False
            If grdDupl.TextMatrix(lmRowNo, ilBoxNo) = "No" Then
                grdDupl.TextMatrix(lmRowNo, ilBoxNo) = ""
                grdDupl.TextMatrix(lmRowNo, BASEINDEX) = ""
                grdDupl.TextMatrix(lmRowNo, RPTINDEX) = ""
                grdDupl.TextMatrix(lmRowNo, SORTINDEX) = ""
            End If
        Case BASEINDEX
            pbcYN.Visible = False
        Case RPTINDEX
            pbcYN.Visible = False
        Case SORTINDEX
            edcDropDown.Visible = False
            If Trim$(grdDupl.TextMatrix(lmRowNo, SORTINDEX)) <> edcDropDown.Text Then
                imChg = True
            End If
            grdDupl.TextMatrix(lmRowNo, SORTINDEX) = edcDropDown.Text
    End Select
    mSetCommands
End Sub

Private Sub mCopyRate(llFromIndex As Long, llToIndex As Long)
    Dim ilWk As Integer
    Dim llFromLkYear As Long
    Dim llToLkYear As Long
    
    For ilWk = LBound(tmRifRec(llFromIndex).tRif.lRate) To UBound(tmRifRec(llFromIndex).tRif.lRate) Step 1
        tmRifRec(llToIndex).tRif.lRate(ilWk) = tmRifRec(llFromIndex).tRif.lRate(ilWk)
    Next ilWk
    llFromLkYear = tmRifRec(llFromIndex).lLkYear
    llToLkYear = tmRifRec(llToIndex).lLkYear
    Do While llFromLkYear > 0
        For ilWk = LBound(tmRifRec(llFromLkYear).tRif.lRate) To UBound(tmRifRec(llFromLkYear).tRif.lRate) Step 1
            tmRifRec(llToLkYear).tRif.lRate(ilWk) = tmRifRec(llFromLkYear).tRif.lRate(ilWk)
        Next ilWk
        llFromLkYear = tmLkRifRec(llFromLkYear).lLkYear
        llToLkYear = tmLkRifRec(llToLkYear).lLkYear
    Loop
    
End Sub
