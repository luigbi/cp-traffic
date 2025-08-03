VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAddMG 
   Caption         =   "Add MG"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "AffAddMG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   9105
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5925
      TabIndex        =   15
      Top             =   4350
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7545
      TabIndex        =   16
      Top             =   4350
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.ListBox lbcStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffAddMG.frx":08CA
      Left            =   6315
      List            =   "AffAddMG.frx":08CC
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1110
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   60
      Picture         =   "AffAddMG.frx":08CE
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   645
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcMGFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   6
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox pbcMGTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   12
      Top             =   4305
      Width           =   60
   End
   Begin VB.PictureBox pbcMGSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   405
      Width           =   60
   End
   Begin VB.TextBox txtDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   4035
      TabIndex        =   9
      Top             =   1650
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmcDropDown 
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
      Height          =   210
      Left            =   4980
      Picture         =   "AffAddMG.frx":0BD8
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "3) Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   8295
      Begin VB.OptionButton optPeriod 
         Caption         =   "All"
         Height          =   255
         Index           =   2
         Left            =   5100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
      Begin VB.OptionButton optPeriod 
         Caption         =   "Current plus Previous Week"
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   2580
      End
      Begin VB.OptionButton optPeriod 
         Caption         =   "Current Month"
         Height          =   255
         Index           =   1
         Left            =   3585
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Missed from"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   990
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   195
      Top             =   4695
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4920
      FormDesignWidth =   9105
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   4350
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMG 
      Height          =   3810
      Left            =   165
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmAddMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmAddMG - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private lmAttCode As Long
Private imVefCode As Integer
Private imShttCode As Integer
Private lmSdfCode As Long
Private lmAstCode As Long
Private lmLstCode As Long
Private smPostSDate As String
Private smPostEDate As String
Private imIntegralSet As Integer
Private imFirstDrop As Integer
Private imFieldChgd As Integer
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private smZone As String
Private tmAstInfo() As ASTINFO

'Grid Controls
Private imShowGridBox As Integer    'True-Edit box displayed within Grid
Private imFromArrow As Integer      'Tab from Arrow
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on

Private imVehCol As Integer
Private imAdvtCol As Integer
Private imDateCol As Integer
Private imTimeCol As Integer
Const PLEDGEDAYINDEX = 4
Const PLEDGETIMEINDEX = 5
Const AIRDATEINDEX = 7
Const AIRTIMEINDEX = 8
Const STATUSINDEX = 6
Const ASTINDEX = 9

Private Function mTestGridValues()
    Dim iLoop As Integer
    Dim iPack As Integer
    Dim sDate As String
    Dim sTime As String
    Dim iDay As Integer
    Dim iIndex As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilError As Integer
    Dim llRowIndex As Long
    
    grdMG.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdMG.FixedRows To grdMG.Rows - 1 Step 1
        slStr = Trim$(grdMG.TextMatrix(llRow, STATUSINDEX))
        If slStr <> "" Then
            slStr = grdMG.TextMatrix(llRow, STATUSINDEX)
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                iIndex = lbcStatus.ItemData(llRowIndex)
                If tgStatusTypes(iIndex).iPledged <> 2 Then
                    sDate = grdMG.TextMatrix(llRow, AIRDATEINDEX)
                    If (gIsDate(sDate) = False) Or (Len(Trim$(sDate)) = 0) Then   'Date not valid.
                        ilError = True
                        If Len(Trim$(sTime)) = 0 Then
                            grdMG.TextMatrix(llRow, AIRDATEINDEX) = "Missing"
                        End If
                        grdMG.Row = llRow
                        grdMG.Col = AIRDATEINDEX
                        grdMG.CellForeColor = vbRed
                    End If
                    sTime = grdMG.TextMatrix(llRow, AIRTIMEINDEX)
                    If (gIsTime(sTime) = False) Or (Len(Trim$(sTime)) = 0) Then    'Time not valid.
                        ilError = True
                        If Len(Trim$(sTime)) = 0 Then
                            grdMG.TextMatrix(llRow, AIRTIMEINDEX) = "Missing"
                        End If
                        grdMG.Row = llRow
                        grdMG.Col = AIRTIMEINDEX
                        grdMG.CellForeColor = vbRed
                    End If
                End If
            End If
        End If
    Next llRow
    If ilError Then
        grdMG.Redraw = True
        mTestGridValues = False
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        mTestGridValues = True
        Exit Function
    End If
End Function

Private Sub mMGSetShow()
    Dim slStr As String
    Dim ilIndex As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim iIndex As Integer
    Dim ilRowIndex As Integer
    Dim llRow As Long
    Dim llCol As Long
    
    
    'Check if row is without rows defined
    If (lmEnableRow >= grdMG.FixedRows) And (lmEnableRow < grdMG.Rows) Then
        'Process which cell is activated
        Select Case lmEnableCol
            Case AIRDATEINDEX
                'Check if value is Ok
                slStr = txtDropdown.Text
                If (gIsDate(slStr)) And (slStr <> "") Then
                    'Check if value changed
                    If grdMG.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        'Set change flag
                        imFieldChgd = True
                        'Set value back into grid
                        grdMG.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                        'Check if other cell values should be changed
                        'In this case, if status is not aired, then change it
                        If slStr <> "" Then
                            For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                'If tgStatusTypes(gGetAirStatus(iIndex)).iStatus = 20 Then
                                If tgStatusTypes(iIndex).iStatus = ASTEXTENDED_MG Then
                                    grdMG.TextMatrix(lmEnableRow, STATUSINDEX) = Trim$(tgStatusTypes(iIndex).sName)
                                    llRow = grdMG.Row
                                    llCol = grdMG.Col
                                    grdMG.Row = lmEnableRow
                                    grdMG.Col = AIRDATEINDEX
                                    grdMG.CellBackColor = vbWhite
                                    grdMG.Col = AIRTIMEINDEX
                                    grdMG.CellBackColor = vbWhite
                                    grdMG.Row = llRow
                                    grdMG.Col = llCol
                                    Exit For
                                End If
                            Next iIndex
                        End If
                    End If
                End If
            Case AIRTIMEINDEX
                'Check if value is Ok
                slStr = txtDropdown.Text
                If (gIsTime(slStr)) And (slStr <> "") Then
                    'Check if value changed
                    slStr = gConvertTime(slStr)
                    If Second(slStr) = 0 Then
                        slStr = Format$(slStr, sgShowTimeWOSecForm)
                    Else
                        slStr = Format$(slStr, sgShowTimeWSecForm)
                    End If
                    If grdMG.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                        'Set change flag
                        imFieldChgd = True
                        'Check if other cell values should be changed
                        'In this case, if status is not aired, then change it
                        grdMG.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                        If slStr <> "" Then
                            For iIndex = 0 To UBound(tgStatusTypes) Step 1
                                'If tgStatusTypes(gGetAirStatus(iIndex)).iStatus = 20 Then
                                If tgStatusTypes(iIndex).iStatus = ASTEXTENDED_MG Then
                                    grdMG.TextMatrix(lmEnableRow, STATUSINDEX) = Trim$(tgStatusTypes(iIndex).sName)
                                    llRow = grdMG.Row
                                    llCol = grdMG.Col
                                    grdMG.Row = lmEnableRow
                                    grdMG.Col = AIRDATEINDEX
                                    grdMG.CellBackColor = vbWhite
                                    grdMG.Col = AIRTIMEINDEX
                                    grdMG.CellBackColor = vbWhite
                                    grdMG.Row = llRow
                                    grdMG.Col = llCol
                                    Exit For
                                End If
                            Next iIndex
                        End If
                    End If
                End If
            Case STATUSINDEX
                'Check if value changed
                slStr = txtDropdown.Text
                If grdMG.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    'Set change flag
                    imFieldChgd = True
                    'Check if other cell values should be changed
                    'In this case, if status is not aired, then change it
                    grdMG.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    iStatus = -1
                    sStatus = Trim$(grdMG.TextMatrix(lmEnableRow, STATUSINDEX))
                    For iIndex = 0 To UBound(tgStatusTypes) Step 1
                        If StrComp(sStatus, Trim$(tgStatusTypes(iIndex).sName), 1) = 0 Then
                            iStatus = tgStatusTypes(iIndex).iStatus
                            ilRowIndex = iIndex
                            Exit For
                        End If
                    Next iIndex
                    If iStatus <> -1 Then
                        If tgStatusTypes(ilRowIndex).iPledged = 2 Then
'                            grdMG.TextMatrix(lmEnableRow, AIRDATEINDEX) = ""
'                            grdMG.TextMatrix(lmEnableRow, AIRTIMEINDEX) = ""
                            llRow = grdMG.Row
                            llCol = grdMG.Col
                            grdMG.Row = lmEnableRow
                            grdMG.Col = AIRDATEINDEX
                            grdMG.CellBackColor = LIGHTYELLOW
                            grdMG.Col = AIRTIMEINDEX
                            grdMG.CellBackColor = LIGHTYELLOW
                            grdMG.Row = llRow
                            grdMG.Col = llCol
                        Else
                            llRow = grdMG.Row
                            llCol = grdMG.Col
                            grdMG.Row = lmEnableRow
                            grdMG.Col = AIRDATEINDEX
                            grdMG.CellBackColor = vbWhite
                            grdMG.Col = AIRTIMEINDEX
                            grdMG.CellBackColor = vbWhite
                            grdMG.Row = llRow
                            grdMG.Col = llCol
                        End If
                        If Trim$(grdMG.TextMatrix(lmEnableRow, AIRDATEINDEX)) = "" Then
                            iIndex = grdMG.TextMatrix(lmEnableRow, ASTINDEX)
                            grdMG.TextMatrix(lmEnableRow, AIRDATEINDEX) = Format$(tmAstInfo(iIndex).sAirDate, sgShowDateForm)
                        End If
                        If Trim$(grdMG.TextMatrix(lmEnableRow, AIRTIMEINDEX)) = "" Then
                            iIndex = grdMG.TextMatrix(lmEnableRow, ASTINDEX)
                            If Second(tmAstInfo(iIndex).sAirTime) <> 0 Then
                                grdMG.TextMatrix(lmEnableRow, AIRTIMEINDEX) = Format$(tmAstInfo(iIndex).sAirTime, sgShowTimeWSecForm)
                            Else
                                grdMG.TextMatrix(lmEnableRow, AIRTIMEINDEX) = Format$(tmAstInfo(iIndex).sAirTime, sgShowTimeWOSecForm)
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    'Set row and column to not selected
    lmEnableRow = -1
    lmEnableCol = -1
    'Set flag to indicate no control is displayed- used in scroll event
    imShowGridBox = False
    'Set all controls as bot visible
    pbcArrow.Visible = False
    txtDropdown.Visible = False
    lbcStatus.Visible = False
    cmcDropDown.Visible = False
    'Set Save button
    If imFieldChgd Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub

Private Sub mMGEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim iLoop As Integer
    
    'Check if row is without rows defined
    If (grdMG.Row >= grdMG.FixedRows) And (grdMG.Row < grdMG.Rows) And (grdMG.Col >= 0) And (grdMG.Col < grdMG.Cols - 1) Then
        'Save current row and column used in scroll event and mNGSetShow routine
        lmEnableRow = grdMG.Row
        lmEnableCol = grdMG.Col
        'Set flag that cell box is deplayed.  Used in scroll event
        imShowGridBox = True
        'Display arrow
        pbcArrow.Move grdMG.Left - pbcArrow.Width, grdMG.Top + grdMG.RowPos(grdMG.Row) + (grdMG.RowHeight(grdMG.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        'Process which cell is activated
        Select Case grdMG.Col
            Case AIRDATEINDEX  'Date
                'Move control
                txtDropdown.Move grdMG.Left + grdMG.ColPos(grdMG.Col) + 30, grdMG.Top + grdMG.RowPos(grdMG.Row) + 15, grdMG.ColWidth(grdMG.Col) - 30, grdMG.RowHeight(grdMG.Row) - 15
                'Get value to be displayed from cell
                If grdMG.Text <> "Missing" Then
                    txtDropdown.Text = grdMG.Text
                Else
                    txtDropdown.Text = ""
                End If
                'Check if font size needs to be changed
                If txtDropdown.Height > grdMG.RowHeight(grdMG.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdMG.RowHeight(grdMG.Row) - 15
                End If
                'Make control visible
                txtDropdown.Visible = True
                'Set focus from grid or tab control to text box
                txtDropdown.SetFocus
            Case AIRTIMEINDEX  'Time
                'Move control
                txtDropdown.Move grdMG.Left + grdMG.ColPos(grdMG.Col) + 30, grdMG.Top + grdMG.RowPos(grdMG.Row) + 15, grdMG.ColWidth(grdMG.Col) - 30, grdMG.RowHeight(grdMG.Row) - 15
                'Get value to be displayed from cell
                If grdMG.Text <> "Missing" Then
                    txtDropdown.Text = grdMG.Text
                Else
                    txtDropdown.Text = ""
                End If
                'Check if font size needs to be changed
                If txtDropdown.Height > grdMG.RowHeight(grdMG.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdMG.RowHeight(grdMG.Row) - 15
                End If
                'Make control visible
                txtDropdown.Visible = True
                'Set focus from grid or tab control to text box
                txtDropdown.SetFocus
            Case STATUSINDEX
                'Move control
                txtDropdown.Move grdMG.Left + grdMG.ColPos(STATUSINDEX) + 30, grdMG.Top + grdMG.RowPos(grdMG.Row) + 15, grdMG.ColWidth(STATUSINDEX) - cmcDropDown.Width - 30, grdMG.RowHeight(grdMG.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcStatus.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + 3 * txtDropdown.Width
                'Set number of rows to display in list box
                gSetListBoxHeight lbcStatus, 4
                'Get value to be displayed from cell
                slStr = grdMG.TextMatrix(grdMG.Row, STATUSINDEX)
                ilIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcStatus.ListIndex = ilIndex
                Else
                    lbcStatus.ListIndex = 0
                End If
                txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                'Check if font size needs to be changed
                If txtDropdown.Height > grdMG.RowHeight(grdMG.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdMG.RowHeight(grdMG.Row) - 15
                End If
                'Make control visible
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcStatus.Visible = True
                'Set focus from grid or tab control to text box
                txtDropdown.SetFocus
        End Select
    End If
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    gGrid_Clear grdMG, True
    'Set color within cells
    For llRow = grdMG.FixedRows To grdMG.Rows - 1 Step 1
        For llCol = 0 To PLEDGETIMEINDEX Step 1
            grdMG.Row = llRow
            grdMG.Col = llCol
            grdMG.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llRow
End Sub






Private Sub cmcDropDown_Click()
    Select Case grdMG.Col
        Case STATUSINDEX
            lbcStatus.Visible = Not lbcStatus.Visible
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload frmAddMG
End Sub

Private Sub cmdCancel_GotFocus()
    mMGSetShow
End Sub

Private Sub cmdDone_Click()
    Dim iLoop As Integer
    Dim iPostingStatus As Integer
    Dim iPosted As Integer
    Dim sFWkDate As String
    Dim sLWkDate As String
    Dim iRet As Integer
    Dim iStatus As Integer
    
    Screen.MousePointer = vbHourglass
    If imFieldChgd Then
        If Not mSave() Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Unload frmAddMG
    Exit Sub
   
End Sub






Private Sub cmdDone_GotFocus()
    mMGSetShow
End Sub

Private Sub cmdSave_Click()
    Screen.MousePointer = vbHourglass
    If imFieldChgd Then
        If Not mSave() Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    imFieldChgd = False
    cmdSave.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_GotFocus()
    mMGSetShow
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        If igTimes = 0 Then
            imAdvtCol = 0
            imVehCol = 1
            imDateCol = 2
            imTimeCol = 3
        Else
            imVehCol = 0
            imDateCol = 1
            imTimeCol = 2
            imAdvtCol = 3
        End If
        'Set Column widths
        grdMG.ColWidth(ASTINDEX) = 0
        grdMG.ColWidth(imVehCol) = grdMG.Width * 0.18
        grdMG.ColWidth(imDateCol) = grdMG.Width * 0.08
        grdMG.ColWidth(imTimeCol) = grdMG.Width * 0.08
        grdMG.ColWidth(PLEDGEDAYINDEX) = grdMG.Width * 0.06
        grdMG.ColWidth(PLEDGETIMEINDEX) = grdMG.Width * 0.16
        grdMG.ColWidth(AIRDATEINDEX) = grdMG.Width * 0.08
        grdMG.ColWidth(AIRTIMEINDEX) = grdMG.Width * 0.08
        grdMG.ColWidth(STATUSINDEX) = grdMG.Width * 0.08

        grdMG.ColWidth(imAdvtCol) = grdMG.Width - GRIDSCROLLWIDTH
        For ilCol = 0 To AIRTIMEINDEX Step 1
            If ilCol <> imAdvtCol Then
                grdMG.ColWidth(imAdvtCol) = grdMG.ColWidth(imAdvtCol) - grdMG.ColWidth(ilCol)
            End If
        Next ilCol
        'Align columns to left
        gGrid_AlignAllColsLeft grdMG
        'Set column titles
        grdMG.TextMatrix(0, imVehCol) = "Vehicle"
        grdMG.TextMatrix(0, imDateCol) = "Feed Date"
        grdMG.TextMatrix(0, imTimeCol) = "Feed Time"
        grdMG.TextMatrix(0, imAdvtCol) = "Advertiser/ Product"
        grdMG.TextMatrix(0, PLEDGEDAYINDEX) = "Pledge Days"
        grdMG.TextMatrix(0, PLEDGETIMEINDEX) = "Pledge Times"
        grdMG.TextMatrix(0, AIRDATEINDEX) = "Air Date"
        grdMG.TextMatrix(0, AIRTIMEINDEX) = "Aired Time"
        grdMG.TextMatrix(0, STATUSINDEX) = "Status"
        'Set height of grid
        gGrid_IntegralHeight grdMG
        'Clear and set rows in grid
        mClearGrid
        imFirstTime = False
    End If

End Sub

Private Sub Form_Click()
    mMGSetShow
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.15
    Me.Height = Screen.Height / 1.55
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    If igTimes = 0 Then
        'frmDateTimes.Caption = "Spots by Advertiser"
        optPeriod(0).Visible = False
        optPeriod(2).Left = optPeriod(1).Left + (optPeriod(2).Left - optPeriod(1))
        optPeriod(1).Left = optPeriod(0).Left
    Else
        'frmDateTimes.Caption = "Spots by Date"
    End If
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    
    Screen.MousePointer = vbHourglass
    
    frmAddMG.Caption = "Add MG - " & sgClientName
    imIntegralSet = False
    imFirstDrop = True
    imMouseDown = False
    imFirstTime = True
    imFieldChgd = False
    imShowGridBox = False
    lmTopRow = -1
    imFromArrow = False
    lmEnableRow = -1
    lmEnableCol = -1
    imBSMode = False
    ReDim tmAstInfo(0 To 0) As ASTINFO
    For iLoop = 0 To UBound(tgStatusTypes) Step 1
        'If tgStatusTypes(gGetAirStatus(iLoop)).iStatus = 20 Then
        If tgStatusTypes(iLoop).iStatus = ASTEXTENDED_MG Then
            lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
            lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
            Exit For
        End If
    Next iLoop
    For iLoop = 0 To UBound(tgStatusTypes) Step 1
        If tgStatusTypes(iLoop).iPledged = 2 Then
            lbcStatus.AddItem Trim$(tgStatusTypes(iLoop).sName)
            lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
        End If
    Next iLoop
    
    Screen.MousePointer = vbDefault
    Exit Sub
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmAstInfo
    Set frmAddMG = Nothing
End Sub




Private Function mSave() As Integer
    Dim sStr As String
    Dim sDate As String
    Dim sTime As String
    Dim sAirDate As String
    Dim sAirTime As String
'    Dim sLstDate As String
'    Dim sLstTime As String
'    Dim iTimeAdj As Integer
'    Dim lTime As Long
    Dim iVef As Integer
    Dim iZone As Integer
    Dim lCode As Long
    Dim iIndex As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim iRow As Integer
    Dim iChg As Integer
    Dim iRet As Integer
    Dim iLoop As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim rstDat As ADODB.Recordset
    Dim ilAdfCode As Integer
    Dim llDATCode As Long
    Dim llCpfCode As Long
    Dim llRsfCode As Long
    Dim slStationCompliant As String
    Dim slAgencyCompliant As String
    Dim slAffidavitSource As String
    Dim llCntrNo As Long
    Dim ilLen As Integer
    Dim llLkAstCode As Long
    
    On Error GoTo ErrHand
    
    mSave = True
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        Exit Function
    End If
    If sgUstWin(7) <> "I" Then
        Exit Function
    End If
    If Not mTestGridValues() Then
        mSave = False
        Exit Function
    End If
    grdMG.Redraw = False
    llRow = grdMG.FixedRows
    For iRow = 0 To UBound(tmAstInfo) - 1 Step 1
        'Test if value changed and if so, update
        sStr = Trim$(grdMG.TextMatrix(llRow, AIRDATEINDEX))
        sStatus = Trim$(grdMG.TextMatrix(llRow, STATUSINDEX))
        sAirDate = ""
        sAirTime = ""
        iStatus = -1
        For iLoop = 0 To UBound(tgStatusTypes) Step 1
            If StrComp(sStatus, Trim$(tgStatusTypes(iLoop).sName), 1) = 0 Then
                iStatus = tgStatusTypes(iLoop).iStatus
                Exit For
            End If
        Next iLoop
        'If iStatus = 20 Then
        If gIsAstStatus(iStatus, ASTEXTENDED_MG) Then
            sDate = Trim$(grdMG.TextMatrix(llRow, AIRDATEINDEX))
            sTime = Trim$(grdMG.TextMatrix(llRow, AIRTIMEINDEX))
            If ((gIsDate(sDate) = True) And (gIsTime(sTime) = True)) And (Len(Trim$(sTime)) <> 0) Then
                sAirDate = Format$(sDate, sgShowDateForm)
                sAirTime = Format$(sTime, sgShowTimeWSecForm)
            Else
                sAirDate = ""
                sAirTime = ""
            End If
            'iStatus = 20
            iStatus = ASTEXTENDED_MG
        End If
        'If ((iStatus >= 0) And (iStatus < 20)) Or ((iStatus = 20) And (sAirDate <> "")) Then
        If (gGetAirStatus(iStatus) < ASTEXTENDED_MG) Or ((gGetAirStatus(iStatus) = ASTEXTENDED_MG) And (sAirDate <> "")) Or (gGetAirStatus(iStatus) = ASTAIR_MISSED_MG_BYPASS) Then
            iIndex = Val(grdMG.TextMatrix(llRow, ASTINDEX))
            lCode = tmAstInfo(iIndex).lCode
            lmAttCode = tmAstInfo(iIndex).lAttCode
            imShttCode = tmAstInfo(iIndex).iShttCode
            imVefCode = tmAstInfo(iIndex).iVefCode
            lmSdfCode = tmAstInfo(iIndex).lSdfCode
            smZone = tmAstInfo(iIndex).sLstZone
            llCntrNo = tmAstInfo(iIndex).lCntrNo
            ilLen = tmAstInfo(iIndex).iLen
            llLkAstCode = 0
            iChg = False
            If iStatus = 0 Then
                If ((gIsDate(sAirDate) = True) And (gIsTime(sTime) = True)) Then
                    If (DateValue(gAdjYear(sAirDate)) <> DateValue(gAdjYear(tmAstInfo(iIndex).sAirDate))) Or (gTimeToLong(sTime, False) <> gTimeToLong(tmAstInfo(iIndex).sAirTime, False)) Or (tmAstInfo(iIndex).iStatus <> iStatus) Then
                        iChg = True
                    End If
                End If
            Else
                If (tmAstInfo(iIndex).iStatus <> iStatus) Then
                    iChg = True
                End If
            End If
            If iChg Then
                'If iStatus = 20 Then
                If gIsAstStatus(iStatus, ASTEXTENDED_MG) Then
'The adjustment is not required since the user is entering local time
'                    sLstDate = sAirDate
'                    sLstTime = sAirTime
'                    iTimeAdj = 0
'                    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'                        If tgStationInfo(iLoop).iCode = imShttCode Then
'                            iTimeAdj = 0
'                            For iVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
'                                If tgVehicleInfo(iVef).iCode = imVefCode Then
'                                    For iZone = LBound(tgVehicleInfo(iVef).sZone) To UBound(tgVehicleInfo(iVef).sZone) Step 1
'                                        If StrComp(tgStationInfo(iLoop).sZone, tgVehicleInfo(iVef).sZone(iZone), 1) = 0 Then
'                                            iTimeAdj = -tgVehicleInfo(iVef).iLocalAdj(iZone)
'                                            Exit For
'                                        End If
'                                    Next iZone
'                                    Exit For
'                                End If
'                            Next iVef
'                            Exit For
'                        End If
'                    Next iLoop
'                    SQLQuery = "SELECT datDACode FROM dat"
'                    SQLQuery = SQLQuery + " WHERE (datShfCode = " & imShttCode & ")"
'                    SQLQuery = SQLQuery + " And (datAtfCode = " & lmAttCode & ")"
'                    SQLQuery = SQLQuery + " And (datVefCode = " & imVefCode & ")"
'                    Set rstDat = gSQLSelectCall(SQLQuery)
'
'                    If rstDat!datDACode = 2 Then
'                        lTime = gTimeToLong(Format$(sLstTime, "hh:mm:ssam/pm"), False)
'                    Else
'                        lTime = gTimeToLong(Format$(sLstTime, "hh:mm:ssam/pm"), False) + 3600 * iTimeAdj
'                        If lTime < 0 Then
'                            lTime = lTime + 86400
'                            sLstDate = Format$(DateValue(sLstDate) - 1, sgShowDateForm)
'                        ElseIf lTime > 86400 Then
'                            lTime = lTime - 86400
'                            sLstDate = Format$(DateValue(sLstDate) + 1, sgShowDateForm)
'                        End If
'                    End If
'                    sLstTime = Format$(gLongToTime(lTime), sgShowTimeWSecForm)
                    iRet = mAddLst(tmAstInfo(iIndex).lLstCode, sAirDate, sAirTime)
                    If iRet Then
                        SQLQuery = "SELECT * FROM ast"
                        SQLQuery = SQLQuery + " WHERE (astCode = " & lCode & ")"
                        Set rst = gSQLSelectCall(SQLQuery)
                        If Not rst.EOF Then
                            ilAdfCode = rst!astAdfCode
                            llDATCode = rst!astDatCode
                            llCpfCode = rst!astCpfCode
                            llRsfCode = rst!astRsfCode
                            slStationCompliant = ""
                            slAgencyCompliant = ""
                            slAffidavitSource = ""
                            SQLQuery = "INSERT INTO ast (astAtfCode, astShfCode, astVefCode, "
                            SQLQuery = SQLQuery & "astSdfCode, astLsfCode, astAirDate, "
                            SQLQuery = SQLQuery & "astAirTime, astStatus, astCPStatus, "
                            '12/13/13: Support New AST layout
                            'SQLQuery = SQLQuery & "astFeedDate, astFeedTime, astPledgeDate, "
                            'SQLQuery = SQLQuery & "astPledgeStartTime, astPledgeEndTime, astPledgeStatus )"
                            SQLQuery = SQLQuery & "astFeedDate, astFeedTime, "
                            SQLQuery = SQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
                            SQLQuery = SQLQuery & " VALUES (" & rst!astAtfCode & "," & rst!astShfCode & "," & rst!astVefCode & ","
                            SQLQuery = SQLQuery & lCode & "," & tmAstInfo(iIndex).lLstCode & ",'" & Format$(sAirDate, sgSQLDateForm) & "',"
                            SQLQuery = SQLQuery & "'" & Format$(sAirTime, sgSQLTimeForm) & "'," & iStatus & "," & "1" & ","
                            ''SQLQuery = SQLQuery & "'" & rst!astFeedDate & "'," & Format$(rst!astFeedTime, "hh:mm:ss") & ",'" & rst!astPledgeDate & "',"
                            'SQLQuery = SQLQuery & "'" & Format$(sAirDate, sgSQLDateForm) & "','" & Format$(sAirTime, sgSQLTimeForm) & "','" & Format$(rst!astPledgeDate, sgSQLDateForm) & "',"
                            'SQLQuery = SQLQuery & "'" & Format$(rst!astPledgeStartTime, sgSQLTimeForm) & "','" & Format$(rst!astPledgeEndTime, sgSQLTimeForm) & "'," & rst!astPledgeStatus & ")"
                            SQLQuery = SQLQuery & "'" & Format$(sAirDate, sgSQLDateForm) & "','" & Format$(sAirTime, sgSQLTimeForm) & "',"
                            SQLQuery = SQLQuery & ilAdfCode & ", " & llDATCode & ", " & llCpfCode & ", " & llRsfCode & ", "
                            SQLQuery = SQLQuery & "'" & slStationCompliant & "', '" & slAgencyCompliant & "', '" & slAffidavitSource & "', " & llCntrNo & ", " & ilLen & ", " & llLkAstCode & ", " & 0 & ", " & igUstCode & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "AddMG-mSave"
                                cnn.RollbackTrans
                                mSave = False
                                Exit Function
                            End If
                            cnn.CommitTrans
                            SQLQuery = "Select MAX(astCode) from ast"
                            Set rst = gSQLSelectCall(SQLQuery)
                            tmAstInfo(iIndex).lCode = rst(0).Value
                            tmAstInfo(iIndex).sAirDate = sAirDate
                            tmAstInfo(iIndex).sAirTime = sAirTime
                            tmAstInfo(iIndex).iStatus = iStatus
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astStatus = " & ASTEXTENDED_BONUS & ", "
                            SQLQuery = SQLQuery + "astCPStatus = " & "1"
                            SQLQuery = SQLQuery + " WHERE (astCode = " & lCode & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "AddMG-mSave"
                                cnn.RollbackTrans
                                mSave = False
                                Exit Function
                            End If
                            cnn.CommitTrans
                            igUpdateDTGrid = True
                        End If
                    End If
                Else
                    SQLQuery = "UPDATE ast SET "
                    SQLQuery = SQLQuery + "astStatus = " & iStatus & ", "
                    SQLQuery = SQLQuery + "astCPStatus = " & "1"
                    SQLQuery = SQLQuery + " WHERE (astCode = " & lCode & ")"
                    cnn.BeginTrans
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "AddMG-mSave"
                        cnn.RollbackTrans
                        mSave = False
                        Exit Function
                    End If
                    cnn.CommitTrans
                    tmAstInfo(iIndex).iStatus = iStatus
                    igUpdateDTGrid = True
                End If
            End If
        End If
        llRow = llRow + 1
    Next iRow
    On Error GoTo 0
    grdMG.Redraw = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AddMG-mSave"
    mSave = False
    Exit Function
End Function

Private Sub mGetMissed()
    Dim sAdvertiser As String
    Dim sFdDate As String
    Dim sFdTime As String
    Dim sPdDate As String
    Dim sPdDays As String
    Dim sPdTime As String
    Dim sAirDate As String
    Dim sAirTime As String
    Dim iLoop As Integer
    Dim iUpper As Integer
    Dim sFWkDate As String
    Dim sLWkDate As String
    Dim lSDate As Long
    Dim lEDate As Long
    Dim lSTime As Long
    Dim lETime As Long
    Dim lTime As Long
    Dim iIndex As Integer
    Dim sStr As String
    Dim iAst As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim sTstDate As String
    Dim iSLoop As Integer
    Dim llRow As Long
    Dim llCol As Long
    Dim ilRet As Integer
    Dim tlDatPledgeInfo As DATPLEDGEINFO

    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    grdMG.Redraw = False
    mClearGrid
    llRow = grdMG.FixedRows
    
    
    ReDim tmAstInfo(0 To 0) As ASTINFO
    For iLoop = 0 To UBound(tgCPPosting) - 1 Step 1
        
        If optPeriod(0).Value Then
            sFWkDate = Format$(gObtainPrevMonday(tgCPPosting(iLoop).sDate), sgShowDateForm)
            sTstDate = Format$(gObtainStartStd(tgCPPosting(iLoop).sDate), sgShowDateForm)
            sFWkDate = Format$(gObtainPrevMonday(Format$(DateValue(gAdjYear(sFWkDate)) - 1, sgShowDateForm)), sgShowDateForm)
            If DateValue(gAdjYear(sFWkDate)) < DateValue(gAdjYear(sTstDate)) Then
                sFWkDate = Format$(gObtainPrevMonday(tgCPPosting(iLoop).sDate), sgShowDateForm)
            End If
        ElseIf optPeriod(1).Value Then
            sFWkDate = Format$(gObtainStartStd(tgCPPosting(iLoop).sDate), sgShowDateForm)
        Else
            sFWkDate = "1/1/1970"
        End If
        If igTimes = 0 Then
            sLWkDate = Format$(gObtainEndStd(tgCPPosting(iLoop).sDate), sgShowDateForm)
        Else
            sLWkDate = Format$(gObtainNextSunday(tgCPPosting(iLoop).sDate), sgShowDateForm)
        End If
        iUpper = UBound(tmAstInfo)
        '12/13/13
        'SQLQuery = "SELECT astlsfCode, astAirDate, astAirTime, astStatus, astFeedDate, astFeedTime, astPledgeDate, astPledgeStartTime, astPledgeEndTime, astCode, lstZone FROM ast, lst"
        SQLQuery = "SELECT astlsfCode, astAtfCode, astDatCode, astVefCode, astAirDate, astAirTime, astStatus, astFeedDate, astFeedTime, astCode, lstZone FROM ast, lst"
        SQLQuery = SQLQuery + " WHERE (astatfCode= " & tgCPPosting(iLoop).lAttCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " AND lstCode = astLsfCode" & ")"
        SQLQuery = SQLQuery + " ORDER BY astAirDate, astAirTime"
        
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            If (gGetAirStatus(rst!astStatus) >= 2) And (gGetAirStatus(rst!astStatus) <= 5) Or (gGetAirStatus(rst!astStatus) = 8) Then
                tmAstInfo(iUpper).lCode = rst!astCode
                tmAstInfo(iUpper).lLstCode = rst!astLsfCode
                tmAstInfo(iUpper).iStatus = rst!astStatus
                tmAstInfo(iUpper).sAirDate = Format$(rst!astAirDate, sgShowDateForm)
                If Second(rst!astAirTime) <> 0 Then
                    tmAstInfo(iUpper).sAirTime = Format$(rst!astAirTime, sgShowTimeWSecForm)
                Else
                    tmAstInfo(iUpper).sAirTime = Format$(rst!astAirTime, sgShowTimeWOSecForm)
                End If
                tmAstInfo(iUpper).sFeedDate = Format$(rst!astFeedDate, sgShowDateForm)
                If Second(rst!astFeedTime) <> 0 Then
                    tmAstInfo(iUpper).sFeedTime = Format$(rst!astFeedTime, sgShowTimeWSecForm)
                Else
                    tmAstInfo(iUpper).sFeedTime = Format$(rst!astFeedTime, sgShowTimeWOSecForm)
                End If
                'tmAstInfo(iUpper).sPledgeDate = Format$(rst!astPledgeDate, sgShowDateForm)
                'If Second(rst!astPledgeStartTime) <> 0 Then
                '    tmAstInfo(iUpper).sPledgeStartTime = Format$(rst!astPledgeStartTime, sgShowTimeWSecForm)
                'Else
                '    tmAstInfo(iUpper).sPledgeStartTime = Format$(rst!astPledgeStartTime, sgShowTimeWOSecForm)
                'End If
                'If IsNull(rst!astPledgeEndTime) = False Then
                '    If Second(rst!astPledgeEndTime) <> 0 Then
                '        tmAstInfo(iUpper).sPledgeEndTime = Format$(rst!astPledgeEndTime, sgShowTimeWSecForm)
                '    Else
                '        tmAstInfo(iUpper).sPledgeEndTime = Format$(rst!astPledgeEndTime, sgShowTimeWOSecForm)
                '    End If
                'Else
                '    tmAstInfo(iUpper).sPledgeEndTime = ""
                'End If
                '12/13/13: Obtain Pledge information from Dat
                tlDatPledgeInfo.lAttCode = rst!astAtfCode
                tlDatPledgeInfo.lDatCode = rst!astDatCode
                tlDatPledgeInfo.iVefCode = rst!astVefCode
                tlDatPledgeInfo.sFeedDate = Format(rst!astFeedDate, "m/d/yy")
                tlDatPledgeInfo.sFeedTime = Format(rst!astFeedTime, "hh:mm:ssam/pm")
                ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
                tmAstInfo(iUpper).sPledgeDate = tlDatPledgeInfo.sPledgeDate
                tmAstInfo(iUpper).sPledgeStartTime = tlDatPledgeInfo.sPledgeStartTime
                tmAstInfo(iUpper).sPledgeEndTime = tlDatPledgeInfo.sPledgeEndTime
                
                tmAstInfo(iUpper).lAttCode = tgCPPosting(iLoop).lAttCode 'tlCPDat(iIndex).lAtfCode
                tmAstInfo(iUpper).iShttCode = tgCPPosting(iLoop).iShttCode 'tlCPDat(iIndex).iShfCode
                tmAstInfo(iUpper).iVefCode = tgCPPosting(iLoop).iVefCode 'tlCPDat(iIndex).iVefCode
                tmAstInfo(iUpper).sLstZone = rst!lstZone
                iUpper = iUpper + 1
                ReDim Preserve tmAstInfo(0 To iUpper) As ASTINFO
            End If
            rst.MoveNext
        Wend
    Next iLoop
    For iAst = 0 To UBound(tmAstInfo) - 1 Step 1
        'SQLQuery = "SELECT lst.lstProd, lst.lstLogDate, lst.lstLogTime, lst.lstSdfCode, lst.lstCode, adf.adfName, vef.vefName FROM lst, ADF_Advertisers adf, VEF_Vehicles vef"
        SQLQuery = "SELECT lstProd, lstLogDate, lstLogTime, lstSdfCode, lstCode, adfName, vefType, vefName"
        SQLQuery = SQLQuery & " FROM lst, ADF_Advertisers, "
        SQLQuery = SQLQuery & "VEF_Vehicles"
        SQLQuery = SQLQuery + " WHERE (adfCode = lstAdfCode"
        SQLQuery = SQLQuery + " AND vefCode = lstLogVefCode"
        SQLQuery = SQLQuery + " AND lstCode = " & tmAstInfo(iAst).lLstCode & ")"
        
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            If IsNull(rst!lstProd) Then
                sAdvertiser = Trim$(rst!adfName)
            Else
                If Trim$(rst!lstProd) <> "" Then
                    sAdvertiser = Trim$(rst!adfName) & ", " & Trim$(rst!lstProd)
                Else
                    sAdvertiser = Trim$(rst!adfName)
                End If
            End If
            If Len(Trim$(tmAstInfo(iAst).sPledgeEndTime)) <> 0 Then
                sPdTime = Trim$(tmAstInfo(iAst).sPledgeStartTime) & "-" & Trim$(tmAstInfo(iAst).sPledgeEndTime)
            Else
                sPdTime = Trim$(tmAstInfo(iAst).sPledgeStartTime)
            End If
            For iSLoop = 0 To UBound(tgStatusTypes) Step 1
                If tmAstInfo(iAst).iStatus = tgStatusTypes(iSLoop).iStatus Then
                    sStatus = Trim$(tgStatusTypes(iSLoop).sName)
                    Exit For
                End If
            Next iSLoop
            sAirDate = ""
            sAirTime = ""
            'Select Case tmAstInfo(iAst).iStatus
            '    Case 1
            '        sStatus = "2: Not Aired"
            '        sAirDate = ""
            '        sAirTime = ""
            '    Case 2
            '        sStatus = "3: Not Aired-Tech Error"
            '        sAirDate = ""
            '        sAirTime = ""
            '    Case 3
            '        sStatus = "4: Not Aired-Blackout"
            '        sAirDate = ""
            '        sAirTime = ""
            '    Case 4
            '        sStatus = "5: Not Aired-Product"
            '        sAirDate = ""
            '        sAirTime = ""
            '    Case 5
            '        sStatus = "6: Not Aired-Off Air"
            '        sAirDate = ""
            '        sAirTime = ""
             '   Case 6
            '        sStatus = "7: Not Aired-Other"
            '        sAirDate = ""
            '        sAirTime = ""
            'End Select
            'This is where the grid gets loaded for the cp screen
            If llRow + 1 > grdMG.Rows Then
                grdMG.AddItem ""
            End If
            grdMG.Row = llRow
            For llCol = 0 To PLEDGETIMEINDEX Step 1
                grdMG.Row = llRow
                grdMG.Col = llCol
                grdMG.CellBackColor = LIGHTYELLOW
            Next llCol
            For llCol = AIRDATEINDEX To AIRTIMEINDEX Step 1
                grdMG.Row = llRow
                grdMG.Col = llCol
                grdMG.CellBackColor = LIGHTYELLOW
            Next llCol
            If igTimes = 0 Then
                grdMG.TextMatrix(llRow, 0) = Trim$(sAdvertiser)
                If sgShowByVehType = "Y" Then
                    grdMG.TextMatrix(llRow, 1) = Trim$(rst!vefType) & ":" & Trim$(rst!vefName)
                Else
                    grdMG.TextMatrix(llRow, 1) = Trim$(rst!vefName)
                End If
                grdMG.TextMatrix(llRow, 2) = Trim$(tmAstInfo(iAst).sFeedDate)
                grdMG.TextMatrix(llRow, 3) = Trim$(tmAstInfo(iAst).sFeedTime)
            Else
                If sgShowByVehType = "Y" Then
                    grdMG.TextMatrix(llRow, 0) = Trim$(rst!vefType) & ":" & Trim$(rst!vefName)
                Else
                    grdMG.TextMatrix(llRow, 0) = Trim$(rst!vefName)
                End If
                grdMG.TextMatrix(llRow, 1) = Trim$(tmAstInfo(iAst).sFeedDate)
                grdMG.TextMatrix(llRow, 2) = Trim$(tmAstInfo(iAst).sFeedTime)
                grdMG.TextMatrix(llRow, 3) = Trim$(sAdvertiser)
            End If
            grdMG.TextMatrix(llRow, PLEDGEDAYINDEX) = Trim$(tmAstInfo(iAst).sPledgeDate)
            grdMG.TextMatrix(llRow, PLEDGETIMEINDEX) = Trim$(sPdTime)
            grdMG.TextMatrix(llRow, AIRDATEINDEX) = Trim$(sAirDate)
            grdMG.TextMatrix(llRow, AIRTIMEINDEX) = Trim$(sAirTime)
            grdMG.TextMatrix(llRow, STATUSINDEX) = Trim$(sStatus)
            grdMG.TextMatrix(llRow, ASTINDEX) = iAst
            llRow = llRow + 1
        End If
    Next iAst
    'Don't add extra row
'    If llRow >= grdMG.Rows Then
'        grdMG.AddItem ""
'    End If
    grdMG.Redraw = True
    imFieldChgd = False
    cmdSave.Enabled = False
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Add MG-mGetMissed"
End Sub

Private Sub grdMG_Click()
    Dim llRow As Long
    Dim llCol As Long
    
    'Check if user allowed to alter values
    If sgUstWin(7) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Check if any row can be entered
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Check if column allowed to be altered
    If (grdMG.Col < STATUSINDEX) Or (grdMG.Col > AIRTIMEINDEX) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdMG_EnterCell()
    'Check if any row can be entered
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Remove any grid box that is showing from a previous mouseup or tab event
    mMGSetShow
    'Check if user allowed to alter values
    If sgUstWin(7) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdMG_GotFocus()
    'Check if any row can be entered
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdMG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked
    'out temporary if user dragged mouse to different cell
    lmTopRow = grdMG.TopRow
    'Turn off refreshing of grid until mouseup event
    grdMG.Redraw = False
End Sub

Private Sub grdMG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String
    Dim ilIndex As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim iIndex As Integer
    Dim ilRowIndex As Integer
    
    'Check if Ok for user to alter value
    If sgUstWin(7) <> "I" Then
        grdMG.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Check that values exist to be altered
    If UBound(tmAstInfo) <= LBound(tmAstInfo) Then
        grdMG.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdMG, X, Y)
    If Not ilFound Then
        grdMG.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Check that row is not beyond end of rows defined
    If grdMG.Row - 1 >= UBound(tmAstInfo) Then
        grdMG.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Check if column allowed to be altered
    If (grdMG.Col < STATUSINDEX) Then
        grdMG.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Check if column allowed to be altered
    If grdMG.Col > AIRTIMEINDEX Then
        grdMG.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If (grdMG.Col = AIRDATEINDEX) Or (grdMG.Col = AIRTIMEINDEX) Then
        iStatus = -1
        sStatus = Trim$(grdMG.TextMatrix(grdMG.Row, STATUSINDEX))
        For iIndex = 0 To UBound(tgStatusTypes) Step 1
            If StrComp(sStatus, Trim$(tgStatusTypes(iIndex).sName), 1) = 0 Then
                iStatus = tgStatusTypes(iIndex).iStatus
                ilRowIndex = iIndex
                Exit For
            End If
        Next iIndex
        If iStatus <> -1 Then
            If tgStatusTypes(ilRowIndex).iPledged = 2 Then
                grdMG.Redraw = True
                pbcClickFocus.SetFocus
                Exit Sub
            End If
        End If
    End If
    'Save toprow for scroll event
    lmTopRow = grdMG.TopRow
    'Reset that grid can be refreshed
    grdMG.Redraw = True
    'Show box within cell
    mMGEnableBox
End Sub

Private Sub grdMG_Scroll()
    'Handle mousedown and drag as it could cause a scroll event
    If grdMG.Redraw = False Then
        grdMG.Redraw = True
        grdMG.TopRow = lmTopRow
        grdMG.Refresh
        grdMG.Redraw = False
    End If
    'Check if cell control were displayed
    If (imShowGridBox) And (grdMG.Row >= grdMG.FixedRows) And (grdMG.Col >= 0) And (grdMG.Col < grdMG.Cols - 1) Then
        'Check if row is visible
        If grdMG.RowIsVisible(grdMG.Row) Then
            'Show arrow and controls
            pbcArrow.Move grdMG.Left - pbcArrow.Width, grdMG.Top + grdMG.RowPos(grdMG.Row) + (grdMG.RowHeight(grdMG.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
           If grdMG.Col = AIRDATEINDEX Then  'Date
                txtDropdown.Move grdMG.Left + grdMG.ColPos(grdMG.Col) + 30, grdMG.Top + grdMG.RowPos(grdMG.Row) + 15, grdMG.ColWidth(grdMG.Col) - cmcDropDown.Width - 30, grdMG.RowHeight(grdMG.Row) - 15
                txtDropdown.Visible = True
                txtDropdown.SetFocus
           ElseIf grdMG.Col = AIRTIMEINDEX Then  'Time
                txtDropdown.Move grdMG.Left + grdMG.ColPos(grdMG.Col) + 30, grdMG.Top + grdMG.RowPos(grdMG.Row) + 15, grdMG.ColWidth(grdMG.Col) - cmcDropDown.Width - 30, grdMG.RowHeight(grdMG.Row) - 15
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            ElseIf grdMG.Col = STATUSINDEX Then
                txtDropdown.Move grdMG.Left + grdMG.ColPos(grdMG.Col) + 30, grdMG.Top + grdMG.RowPos(grdMG.Row) + 15, grdMG.ColWidth(grdMG.Col) - cmcDropDown.Width - 30, grdMG.RowHeight(grdMG.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcStatus.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + 3 * txtDropdown.Width
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcStatus.Visible = True
                txtDropdown.SetFocus
            End If
        Else
            'Hide all controls
            pbcMGFocus.SetFocus
            txtDropdown.Visible = False
            lbcStatus.Visible = False
            cmcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        'Hide arrow
        pbcMGFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub lbcStatus_Click()
    txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
    If (txtDropdown.Visible) And (txtDropdown.Enabled) Then
        txtDropdown.SetFocus
        lbcStatus.Visible = False
    End If
End Sub

Private Sub optPeriod_Click(Index As Integer)
    
    If imFieldChgd Then
        If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
            If Not mSave() Then
                Exit Sub
            End If
        End If
    End If
    mGetMissed
End Sub


Private Function mAddLst(llLstCode As Long, slAirDate As String, slAirTime As String) As Integer
    Dim slProd As String
    Dim slCart As String
    Dim slISCI As String
    
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT * FROM lst"
    SQLQuery = SQLQuery + " WHERE (lstCode= " & llLstCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!lstType <> 2 Then
            If IsNull(rst!lstProd) Then
                slProd = ""
            Else
                slProd = gFixQuote(rst!lstProd)
            End If
            If IsNull(rst!lstCart) Or Left$(rst!lstCart, 1) = Chr$(0) Then
                slCart = ""
            Else
                slCart = gFixQuote(rst!lstCart)
            End If
            If IsNull(rst!lstISCI) Then
                slISCI = ""
            Else
                slISCI = gFixQuote(rst!lstISCI)
            End If
            SQLQuery = "INSERT INTO lst (lstType, lstSdfCode, lstCntrNo, "
            SQLQuery = SQLQuery & "lstAdfCode, lstAgfCode, lstProd, "
            SQLQuery = SQLQuery & "lstLineNo, lstLnVefCode, lstStartDate, "
            SQLQuery = SQLQuery & "lstEndDate, lstMon, lstTue, "
            SQLQuery = SQLQuery & "lstWed, lstThu, lstFri, "
            SQLQuery = SQLQuery & "lstSat, lstSun, lstSpotsWk, "
            SQLQuery = SQLQuery & "lstPriceType, lstPrice, lstSpotType, "
            SQLQuery = SQLQuery & "lstLogVefCode, lstLogDate, lstLogTime, "
            SQLQuery = SQLQuery & "lstDemo, lstAud, lstISCI, "
            SQLQuery = SQLQuery & "lstWkNo, lstBreakNo, lstPositionNo, "
            SQLQuery = SQLQuery & "lstSeqNo, lstZone, lstCart, "
            SQLQuery = SQLQuery & "lstCpfCode, lstCrfCsfCode, lstStatus, "
            SQLQuery = SQLQuery & "lstLen, lstUnits, lstCifCode, "
            '12/28/06
            'SQLQuery = SQLQuery & "lstAnfCode)"
            SQLQuery = SQLQuery & "lstAnfCode, lstEvtIDCefCode, lstSplitNetwork, "
            'SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, lstUnused)"
            SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, "
            SQLQuery = SQLQuery & "lstLnStartTime, lstLnEndTime, lstUnused)"
            SQLQuery = SQLQuery & " VALUES (" & 2 & ", " & 0 & ", " & rst!lstCntrNo & ", "
            SQLQuery = SQLQuery & rst!lstAdfCode & ", " & rst!lstAgfCode & ", '" & slProd & "', "
            SQLQuery = SQLQuery & rst!lstLineNo & ", " & rst!lstLnVefCode & ", '" & Format$(rst!lstStartDate, sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$(rst!lstEndDate, sgSQLDateForm) & "', " & rst!lstMon & ", " & rst!lstTue & ", "
            SQLQuery = SQLQuery & rst!lstWed & ", " & rst!lstThu & ", " & rst!lstFri & ", "
            SQLQuery = SQLQuery & rst!lstSat & ", " & rst!lstSun & ", " & rst!lstSpotsWk & ", "
            SQLQuery = SQLQuery & rst!lstPriceType & ", " & rst!lstPrice & ", " & 5 & ", "
            SQLQuery = SQLQuery & imVefCode & ", '" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & "'" & rst!lstDemo & "', " & rst!lstAud & ", '" & slISCI & "', "
            SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
            SQLQuery = SQLQuery & 0 & ", '" & smZone & "', '" & slCart & "', "
            'SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 20 & ", "
            SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & ASTEXTENDED_MG & ", "
            SQLQuery = SQLQuery & rst!lstLen & ", " & 0 & ", " & 0 & ", "
            '12/28/06
            'SQLQuery = SQLQuery & rst!lstAnfCode & ")"
            SQLQuery = SQLQuery & rst!lstAnfCode & ", " & 0 & ", '" & "N" & "', "
            SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", "
            SQLQuery = SQLQuery & "'" & Format("12am", sgSQLTimeForm) & "', '" & Format("12am", sgSQLTimeForm) & "', '" & "" & "'" & ")"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "AddMG-mAddLst"
                cnn.RollbackTrans
                mAddLst = False
                Exit Function
            End If
            cnn.CommitTrans
            SQLQuery = "Select MAX(lstCode) from lst"
            Set rst = gSQLSelectCall(SQLQuery)
            llLstCode = rst(0).Value
            mAddLst = True
        Else
            SQLQuery = "UPDATE lst SET "
            SQLQuery = SQLQuery + "lstLogDate = '" & Format$(slAirDate, sgSQLDateForm) & "', "
            SQLQuery = SQLQuery + "lstLogTime = '" & Format$(slAirTime, sgSQLTimeForm) & "', "
            'D.S. 07/05/01
            'SQLQuery = SQLQuery + "astStatus = " & 20
            SQLQuery = SQLQuery + "lstStatus = " & 20
            SQLQuery = SQLQuery + " WHERE (lstCode = " & llLstCode & ")"
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "AddMG-mAddLst"
                cnn.RollbackTrans
                mAddLst = False
                Exit Function
            End If
            cnn.CommitTrans
            mAddLst = True
        End If
    Else
        mAddLst = False
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AddMG-mAddLst"
    mAddLst = False
    Exit Function
End Function

Private Sub optPeriod_GotFocus(Index As Integer)
    mMGSetShow
End Sub

Private Sub pbcMGSTab_GotFocus()
    Dim slStr As String
    Dim llRowIndex As Long
    Dim iIndex As Integer
    
    'Check that event is being caused by user action
    If GetFocus() <> pbcMGSTab.hwnd Then
        Exit Sub
    End If
    'Check if user allowed to change values
    If sgUstWin(7) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Check if coming from arrow (new line only), then enable control in cell
    If imFromArrow Then
        imFromArrow = False
        mMGEnableBox
        Exit Sub
    End If
    'Check if coming from a control
    If txtDropdown.Visible Then
        'Set current value back into grid and hide controls
        mMGSetShow
        'Determine next step to process
        If grdMG.Col = AIRDATEINDEX Then
            'Check if date is valid
            slStr = Trim$(txtDropdown.Text)
            If slStr = "" Then
                Beep
                grdMG.Col = grdMG.Col
                mMGEnableBox
                Exit Sub
            Else
                If (Not gIsDate(slStr)) Then
                    Beep
                    grdMG.Col = grdMG.Col
                    mMGEnableBox
                    Exit Sub
                End If
            End If
            grdMG.Col = grdMG.Col - 1
            mMGEnableBox
        ElseIf grdMG.Col = AIRTIMEINDEX Then
            'Check if time is valid
            slStr = Trim$(txtDropdown.Text)
            If slStr = "" Then
                Beep
                grdMG.Col = grdMG.Col
                mMGEnableBox
                Exit Sub
            Else
                If (Not gIsTime(slStr)) Then
                    Beep
                    grdMG.Col = grdMG.Col
                    mMGEnableBox
                    Exit Sub
                End If
            End If
            grdMG.Col = grdMG.Col - 1
            mMGEnableBox
        ElseIf grdMG.Col = STATUSINDEX Then
            If grdMG.Row > grdMG.FixedRows Then
                lmTopRow = -1
                grdMG.Row = grdMG.Row - 1
                If Not grdMG.RowIsVisible(grdMG.Row) Then
                    grdMG.TopRow = grdMG.TopRow - 1
                End If
                grdMG.Col = STATUSINDEX
                mMGEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        Else
            grdMG.Col = grdMG.Col - 1
            mMGEnableBox
        End If
    Else
        'Show text box in first row on first allowed column
        lmTopRow = -1
        grdMG.TopRow = grdMG.FixedRows
        grdMG.Col = AIRDATEINDEX
        grdMG.Row = grdMG.FixedRows
        mMGEnableBox
    End If
End Sub

Private Sub pbcMGTab_GotFocus()
    Dim slStr As String
    Dim llRowIndex As Long
    Dim iIndex As Integer
    Dim llRow As Long
    
    'Check that event is being caused by user action
    If GetFocus() <> pbcMGTab.hwnd Then
        Exit Sub
    End If
    'Check if user allowed to change values
    If sgUstWin(7) <> "I" Then
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Check if coming from a control
    If txtDropdown.Visible Then
        'Set current value back into grid and hide controls
        mMGSetShow
        'Determine next step to process
        If grdMG.Col = STATUSINDEX Then
            'Procede to next row if allowed
'            If grdMG.Row + 1 < grdMG.Rows Then
            slStr = grdMG.TextMatrix(grdMG.Row, STATUSINDEX)
            llRowIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRowIndex >= 0 Then
                iIndex = lbcStatus.ItemData(llRowIndex)
                If tgStatusTypes(iIndex).iPledged = 3 Then
                    grdMG.Col = AIRDATEINDEX
                    mMGEnableBox
                    Exit Sub
                End If
            End If
            llRow = grdMG.Rows
            Do
                llRow = llRow - 1
            Loop While grdMG.TextMatrix(llRow, STATUSINDEX) = ""
            llRow = llRow + 1
            If (grdMG.Row + 1 < llRow) Then
                lmTopRow = -1
                grdMG.Row = grdMG.Row + 1
                If Not grdMG.RowIsVisible(grdMG.Row) Then
                    grdMG.TopRow = grdMG.TopRow + 1
                End If
                grdMG.Col = STATUSINDEX
                mMGEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        ElseIf grdMG.Col = AIRDATEINDEX Then
            'Check if date is valid
            slStr = Trim$(txtDropdown.Text)
            If slStr = "" Then
                Beep
                grdMG.Col = grdMG.Col
                mMGEnableBox
            Else
                If Not gIsDate(slStr) Then
                    Beep
                    grdMG.Col = grdMG.Col
                    mMGEnableBox
                End If
            End If
            grdMG.Col = grdMG.Col + 1
            mMGEnableBox
        ElseIf grdMG.Col = AIRTIMEINDEX Then
            'Check if time is valid
            slStr = Trim$(txtDropdown.Text)
            If slStr = "" Then
                Beep
                grdMG.Col = grdMG.Col
                mMGEnableBox
            Else
                If Not gIsTime(slStr) Then
                    Beep
                    grdMG.Col = grdMG.Col
                    mMGEnableBox
                End If
            End If
            llRow = grdMG.Rows
            Do
                llRow = llRow - 1
            Loop While grdMG.TextMatrix(llRow, STATUSINDEX) = ""
            llRow = llRow + 1
            If (grdMG.Row + 1 < llRow) Then
                lmTopRow = -1
                grdMG.Row = grdMG.Row + 1
                If Not grdMG.RowIsVisible(grdMG.Row) Then
                    grdMG.TopRow = grdMG.TopRow + 1
                End If
                grdMG.Col = STATUSINDEX
                mMGEnableBox
            Else
                pbcClickFocus.SetFocus
            End If
        Else
            grdMG.Col = grdMG.Col + 1
            mMGEnableBox
        End If
    Else
        'Show text box in first row on first allowed column
        grdMG.TopRow = grdMG.FixedRows
        grdMG.Col = AIRDATEINDEX
        grdMG.Row = grdMG.FixedRows
        mMGEnableBox
    End If
End Sub

Private Sub txtDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    'Handle back space with type head
    slStr = txtDropdown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    'Check values associated with each cell
    Select Case grdMG.Col
        Case AIRDATEINDEX
            'Change color if OK, value will be set into grid with mMGSetShow
            slStr = Trim$(txtDropdown.Text)
            If (gIsDate(slStr)) And (slStr <> "") Then
                grdMG.CellForeColor = vbBlack
            End If
        Case AIRTIMEINDEX
            'Change color if OK, value will be set into grid with mMGSetShow
            slStr = Trim$(txtDropdown.Text)
            If (gIsTime(slStr)) And (slStr <> "") Then
                grdMG.CellForeColor = vbBlack
            End If
        Case STATUSINDEX
            'Look for partial match so that type head works
            llRow = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRow >= 0 Then
                lbcStatus.ListIndex = llRow
                txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                txtDropdown.SelStart = ilLen
                txtDropdown.SelLength = Len(txtDropdown.Text)
            End If
    End Select
End Sub

Private Sub txtDropdown_GotFocus()
    'Select text within cell control
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub txtDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If txtDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub txtDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdMG.Col
            Case STATUSINDEX
                gProcessArrowKey Shift, KeyCode, lbcStatus, True
        End Select
    End If
End Sub
