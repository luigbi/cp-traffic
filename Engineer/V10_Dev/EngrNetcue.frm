VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrNetcue 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrNetcue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   11700
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   15
      Width           =   45
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   10935
      Top             =   5730
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2685
      TabIndex        =   7
      Top             =   2430
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
      Left            =   3630
      Picture         =   "EngrNetcue.frx":030A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcDNE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrNetcue.frx":0404
      Left            =   4170
      List            =   "EngrNetcue.frx":040B
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4350
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox edcGrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2475
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   195
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   10
      Top             =   6735
      Width           =   60
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   195
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   450
      Width           =   60
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   480
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   60
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   105
      Picture         =   "EngrNetcue.frx":0417
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6975
      TabIndex        =   13
      Top             =   6630
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   10710
      Top             =   6465
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7290
      FormDesignWidth =   11790
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5205
      TabIndex        =   12
      Top             =   6630
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3420
      TabIndex        =   11
      Top             =   6630
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdNetcue 
      Height          =   5835
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   10292
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmcSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10035
      TabIndex        =   15
      Top             =   75
      Width           =   795
   End
   Begin VB.TextBox edcSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8340
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2250
      Picture         =   "EngrNetcue.frx":0721
      Top             =   6540
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Netcue"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1350
      Picture         =   "EngrNetcue.frx":0A2B
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8985
      Picture         =   "EngrNetcue.frx":12F5
      Top             =   6540
      Width           =   480
   End
End
Attribute VB_Name = "EngrNetcue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrNetcue - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private smState As String
Private imInChg As Integer
Private imBSMode As Integer
Private ilNNECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private imDoubleClickName As Integer
Private imLbcMouseDown As Integer

Private tmNNE As NNE

Private imDeleteCodes() As Integer

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer

Const NAMEINDEX = 0
Const DESCRIPTIONINDEX = 1
Const LIBNAMEINDEX = 2
Const STATEINDEX = 3
Const CODEINDEX = 4
Const USEDFLAGINDEX = 5

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    slStr = Trim$(edcSearch.text)
    llRow = gGrid_Search(grdNetcue, NAMEINDEX, slStr)
    If llRow >= 0 Then
        mEnableBox
    End If
End Sub

Private Sub cmcSearch_GotFocus()
    mSetShow
End Sub

Private Sub mPopDNE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_DNE_DayName("C", "B", sgCurrDNEStamp, "EngrAudio-mPopulate Audio Names", tgCurrDNE())
    lbcDNE.Clear
    For ilLoop = 0 To UBound(tgCurrDNE) - 1 Step 1
        lbcDNE.AddItem Trim$(tgCurrDNE(ilLoop).sName)
        lbcDNE.ItemData(lbcDNE.NewIndex) = tgCurrDNE(ilLoop).lCode
    Next ilLoop
    lbcDNE.AddItem "[None]", 0
    lbcDNE.ItemData(lbcDNE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOLIST) = 2) Then
        lbcDNE.AddItem "[New]", 0
        lbcDNE.ItemData(lbcDNE.NewIndex) = 0
    Else
        lbcDNE.AddItem "[View]", 0
        lbcDNE.ItemData(lbcDNE.NewIndex) = 0
    End If
End Sub
Private Function mNameOk() As Integer
    Dim ilError As Integer
    Dim llRow As Long
    Dim llTestRow As Long
    Dim slStr As String
    Dim slTestStr As String
    
    grdNetcue.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdNetcue.FixedRows To grdNetcue.Rows - 1 Step 1
        slStr = Trim$(grdNetcue.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdNetcue.Rows - 1 Step 1
                slTestStr = Trim$(grdNetcue.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdNetcue.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdNetcue.Row = llRow
                        grdNetcue.Col = NAMEINDEX
                        grdNetcue.CellForeColor = vbRed
                    Else
                        grdNetcue.Row = llTestRow
                        grdNetcue.Col = NAMEINDEX
                        grdNetcue.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdNetcue.Redraw = True
    If ilError Then
        MsgBox "Duplicate Names Found, Save Stopped", vbOKOnly + vbExclamation
        mNameOk = False
        Exit Function
    Else
        mNameOk = True
        Exit Function
    End If
End Function

Private Sub mSortCol(ilCol As Integer)
    Dim llEndRow As Long
    mSetShow
    gGrid_SortByCol grdNetcue, NAMEINDEX, ilCol, imLastColSorted, imLastSort
End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    If imInChg Then
        Exit Sub
    End If
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
    If imFieldChgd Then
        'Check that all mandatory answered
        ilRet = mCheckFields(False)
        If ilRet Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
    End If
End Sub

Private Sub mEnableBox()
    Dim slStr As String
    Dim ilIndex As Integer
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIOLIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdNetcue.Row >= grdNetcue.FixedRows) And (grdNetcue.Row < grdNetcue.Rows) And (grdNetcue.Col >= 0) And (grdNetcue.Col < grdNetcue.Cols - 1) Then
        lmEnableRow = grdNetcue.Row
        lmEnableCol = grdNetcue.Col
        sgReturnCallName = grdNetcue.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdNetcue.Left - pbcArrow.Width - 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + (grdNetcue.RowHeight(grdNetcue.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdNetcue.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdNetcue.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdNetcue.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdNetcue.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 15, grdNetcue.ColWidth(grdNetcue.Col) - 30, grdNetcue.RowHeight(grdNetcue.Row) - 15
                'edcGrid.MaxLength = Len(tmNNE.sName)
                edcGrid.MaxLength = gGetAllowedChars("NETCUE1", Len(tmNNE.sName))
                edcGrid.text = grdNetcue.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 15, grdNetcue.ColWidth(grdNetcue.Col) - 30, grdNetcue.RowHeight(grdNetcue.Row) - 15
                edcGrid.MaxLength = Len(tmNNE.sDescription)
                edcGrid.text = grdNetcue.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case LIBNAMEINDEX
                edcDropdown.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 15, grdNetcue.ColWidth(grdNetcue.Col) - cmcDropDown.Width - 30, grdNetcue.RowHeight(grdNetcue.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcDNE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcDNE, CLng(grdNetcue.Height / 2)
                If lbcDNE.Top + lbcDNE.Height > cmcCancel.Top Then
                    lbcDNE.Top = edcDropdown.Top - lbcDNE.Height
                End If
                slStr = grdNetcue.text
                'ilIndex = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcDNE, slStr)
                If ilIndex >= 0 Then
                    lbcDNE.ListIndex = ilIndex
                    edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                Else
                    edcDropdown.text = ""
                    If lbcDNE.ListCount <= 0 Then
                        lbcDNE.ListIndex = -1
                        edcDropdown.text = ""
                    ElseIf lbcDNE.ListCount <= 1 Then
                        lbcDNE.ListIndex = 0
                        edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                    Else
                        lbcDNE.ListIndex = 1
                        edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                    End If
                End If
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcDNE.Visible = True
                edcDropdown.SetFocus
            Case STATEINDEX
                pbcState.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 15, grdNetcue.ColWidth(grdNetcue.Col) - 30, grdNetcue.RowHeight(grdNetcue.Row) - 15
                smState = grdNetcue.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdNetcue.text
    End If
End Sub
Private Sub mSetShow()
    Dim llRow As Long
    Dim slStr As String
    
    tmcClick.Enabled = False
    If (lmEnableRow >= grdNetcue.FixedRows) And (lmEnableRow < grdNetcue.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = grdNetcue.TextMatrix(lmEnableRow, lmEnableCol)
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
            Case LIBNAMEINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcDNE, slStr)
                If (llRow <= 0) Then
                    grdNetcue.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
                If (Trim$(grdNetcue.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdNetcue.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdNetcue.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdNetcue.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdNetcue.TextMatrix(lmEnableRow, NAMEINDEX)
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    lbcDNE.Visible = False
    cmcDropDown.Visible = False
    edcDropdown.Visible = False
    pbcArrow.Visible = False
    edcGrid.Visible = False
    pbcState.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub
Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llLbc As Long
    Dim llRow As Long
    
    grdNetcue.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdNetcue.FixedRows To grdNetcue.Rows - 1 Step 1
        slStr = Trim$(grdNetcue.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdNetcue.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdNetcue.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdNetcue.Row = llRow
                grdNetcue.Col = NAMEINDEX
                grdNetcue.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdNetcue.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdNetcue.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdNetcue.Row = llRow
                    grdNetcue.Col = STATEINDEX
                    grdNetcue.CellForeColor = vbRed
                End If
'                slStr = Trim$(grdNetcue.TextMatrix(llRow, LIBNAMEINDEX))
'                llLbc = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
'                If (llLbc <= 0) Then
'                    ilError = True
'                    If slStr = "" Then
'                        grdNetcue.TextMatrix(llRow, LIBNAMEINDEX) = "Missing"
'                    End If
'                    grdNetcue.Row = llRow
'                    grdNetcue.Col = LIBNAMEINDEX
'                    grdNetcue.CellForeColor = vbRed
'                End If
            End If
        End If
    Next llRow
    grdNetcue.Redraw = True
    If ilError Then
        mCheckFields = False
        Exit Function
    Else
        mCheckFields = True
        Exit Function
    End If
End Function


Private Sub mGridColumns()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    gGrid_AlignAllColsLeft grdNetcue
    mGridColumnWidth
    'Set Titles
    grdNetcue.TextMatrix(0, NAMEINDEX) = "Name"
    grdNetcue.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdNetcue.TextMatrix(0, LIBNAMEINDEX) = "Library"
    grdNetcue.TextMatrix(0, STATEINDEX) = "State"
    grdNetcue.Row = 1
    For ilCol = 0 To grdNetcue.Cols - 1 Step 1
        grdNetcue.Col = ilCol
        grdNetcue.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdNetcue.Height = cmcCancel.Top - grdNetcue.Top - 120    '8 * grdNetcue.RowHeight(0) + 30
    gGrid_IntegralHeight grdNetcue
    gGrid_Clear grdNetcue, True
    grdNetcue.Row = grdNetcue.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdNetcue.Width = EngrNetcue.Width - 2 * grdNetcue.Left
    grdNetcue.ColWidth(CODEINDEX) = 0
    grdNetcue.ColWidth(USEDFLAGINDEX) = 0
    grdNetcue.ColWidth(NAMEINDEX) = grdNetcue.Width / 9
    grdNetcue.ColWidth(LIBNAMEINDEX) = grdNetcue.Width / 6
    grdNetcue.ColWidth(STATEINDEX) = grdNetcue.Width / 15
    grdNetcue.ColWidth(DESCRIPTIONINDEX) = grdNetcue.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdNetcue.ColWidth(DESCRIPTIONINDEX) > grdNetcue.ColWidth(ilCol) Then
                grdNetcue.ColWidth(DESCRIPTIONINDEX) = grdNetcue.ColWidth(DESCRIPTIONINDEX) - grdNetcue.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdNetcue, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim ilDNE As Integer
    Dim ilCCE As Integer
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdNetcue.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdNetcue.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmNNE.iCode = Val(grdNetcue.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdNetcue.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmNNE.sName = ""
    Else
        tmNNE.sName = slStr
    End If
    tmNNE.sDescription = grdNetcue.TextMatrix(llRow, DESCRIPTIONINDEX)
    tmNNE.lDneCode = 0
    slStr = Trim$(grdNetcue.TextMatrix(llRow, LIBNAMEINDEX))
    For ilDNE = 0 To UBound(tgCurrDNE) - 1 Step 1
        If StrComp(Trim$(tgCurrDNE(ilDNE).sName), slStr, vbTextCompare) = 0 Then
            tmNNE.lDneCode = tgCurrDNE(ilDNE).lCode
            Exit For
        End If
    Next ilDNE
    If grdNetcue.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmNNE.sState = "D"
    Else
        tmNNE.sState = "A"
    End If
    If tmNNE.iCode <= 0 Then
        tmNNE.sUsedFlag = "N"
    Else
        tmNNE.sUsedFlag = grdNetcue.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmNNE.iVersion = 0
    tmNNE.iOrigNneCode = tmNNE.iCode
    tmNNE.sCurrent = "Y"
    'tmNNE.sEnteredDate = smNowDate
    'tmNNE.sEnteredTime = smNowTime
    tmNNE.sEnteredDate = Format(Now, sgShowDateForm)
    tmNNE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmNNE.iUieCode = tgUIE.iCode
    tmNNE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilDNE As Integer
    Dim ilCCE As Integer
    
    'gGrid_Clear grdNetcue, True
    llRow = grdNetcue.FixedRows
    For ilLoop = 0 To UBound(tgCurrNNE) - 1 Step 1
        If llRow + 1 > grdNetcue.Rows Then
            grdNetcue.AddItem ""
        End If
        grdNetcue.Row = llRow
        grdNetcue.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrNNE(ilLoop).sName)
        grdNetcue.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrNNE(ilLoop).sDescription)
        For ilDNE = 0 To UBound(tgCurrDNE) - 1 Step 1
            If tgCurrNNE(ilLoop).lDneCode = tgCurrDNE(ilDNE).lCode Then
                grdNetcue.TextMatrix(llRow, LIBNAMEINDEX) = Trim$(tgCurrDNE(ilDNE).sName)
                Exit For
            End If
        Next ilDNE
        If tgCurrNNE(ilLoop).sState = "A" Then
            grdNetcue.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdNetcue.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdNetcue.TextMatrix(llRow, CODEINDEX) = tgCurrNNE(ilLoop).iCode
        grdNetcue.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrNNE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdNetcue.Rows Then
        grdNetcue.AddItem ""
    End If
    grdNetcue.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrNetcue-mPopulate", tgCurrNNE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlNNE As NNE
    
    gSetMousePointer grdNetcue, grdNetcue, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdNetcue, grdNetcue, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdNetcue, grdNetcue, vbDefault
        MsgBox "Duplicated names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdNetcue.Redraw = False
    For llRow = grdNetcue.FixedRows To grdNetcue.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmNNE.sName) <> "" Then
            ilNNECode = tmNNE.iCode
            If tmNNE.iCode > 0 Then
                ilRet = gGetRec_NNE_NetcueName(ilNNECode, "Netcue-mSave: Get NNE", tlNNE)
                If ilRet Then
                    If mCompare(tmNNE, tlNNE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmNNE.iVersion = tlNNE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmNNE.iCode <= 0 Then
                    ilRet = gPutInsert_NNE_NetcueName(0, tmNNE, "Netcue-mSave: Insert NNE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_NNE_NetcueName(1, tmNNE, "Netcue-mSave: Update NNE")
                    ilRet = gPutDelete_NNE_NetcueName(tmNNE.iCode, "Netcue-mSave: Delete NNE")
                    ilRet = gPutInsert_NNE_NetcueName(1, tmNNE, "Netcue-mSave: Insert NNE")
                End If
                ilRet = gPutUpdate_DNE_UsedFlag(tmNNE.lDneCode, tgCurrDNE())
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_NNE_NetcueName(imDeleteCodes(ilLoop), "EngrNetcue- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdNetcue.Redraw = True
    sgCurrNNEStamp = ""
    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrNetcue-mPopulate", tgCurrNNE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrNetcue
End Sub

Private Sub cmcDone_Click()
    
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrNetcue
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdNetcue, grdNetcue, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdNetcue, grdNetcue, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdNetcue, grdNetcue, vbDefault
    Unload EngrNetcue
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdNetcue.Col
        Case LIBNAMEINDEX
            lbcDNE.Visible = Not lbcDNE.Visible
    End Select
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    Dim llTopRow As Long
    
    If imFieldChgd = True Then
        gSetMousePointer grdNetcue, grdNetcue, vbHourglass
        llTopRow = grdNetcue.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdNetcue, grdNetcue, vbDefault
            Exit Sub
        End If
        grdNetcue.Redraw = False
        mClearControls
        mMoveRecToCtrls
        If imLastColSorted >= 0 Then
            If imLastSort = flexSortStringNoCaseDescending Then
                imLastSort = flexSortStringNoCaseAscending
            Else
                imLastSort = flexSortStringNoCaseDescending
            End If
            mSortCol imLastColSorted
        Else
            imLastSort = -1
            mSortCol 0
        End If
        grdNetcue.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdNetcue, grdNetcue, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As String
    
    slStr = edcDropdown.text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdNetcue.Col
        Case LIBNAMEINDEX
            'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
            llRow = gListBoxFind(lbcDNE, slStr)
            If llRow >= 0 Then
                lbcDNE.ListIndex = llRow
                edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
    End Select
    If StrComp(grdNetcue.text, edcDropdown.text, vbTextCompare) <> 0 Then
        imFieldChgd = True
    End If
    If StrComp(Trim$(edcDropdown.text), "[None]", vbTextCompare) <> 0 Then
        grdNetcue.text = edcDropdown.text
    Else
        grdNetcue.text = ""
    End If
    grdNetcue.CellForeColor = vbBlack
    mSetCommands
End Sub

Private Sub edcDropdown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If edcDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdNetcue.Col
            Case LIBNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcDNE, True
        End Select
        tmcClick.Enabled = False
    End If
End Sub

Private Sub edcDropdown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilRet As Integer
    
    If imDoubleClickName Then
        ilRet = mBranch()
    End If
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdNetcue.Col
        Case NAMEINDEX
            If grdNetcue.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdNetcue.text = edcGrid.text
            grdNetcue.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdNetcue.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdNetcue.text = edcGrid.text
            grdNetcue.CellForeColor = vbBlack
        Case STATEINDEX
    End Select
    mSetCommands
End Sub

Private Sub edcGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcSearch_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
        mFindMatch True
    End If
    imFirstActivate = False
    Me.KeyPreview = True
End Sub

Private Sub Form_Click()
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrNetcue
    gCenterFormModal EngrNetcue
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdNetcue.FixedRows) And (lmEnableRow < grdNetcue.Rows) Then
            If (lmEnableCol >= grdNetcue.FixedCols) And (lmEnableCol < grdNetcue.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdNetcue.text = smESCValue
                End If
                mSetShow
                mEnableBox
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
End Sub

Private Sub Form_Resize()
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    grdNetcue.Height = cmcCancel.Top - grdNetcue.Top - 120    '8 * grdNetcue.RowHeight(0) + 30
    gGrid_IntegralHeight grdNetcue
    gGrid_FillWithRows grdNetcue
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrNetcue = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdNetcue, grdNetcue, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    imcInsert.Picture = EngrMain!imcInsert.Picture
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    ReDim imDeleteCodes(0 To 0) As Integer
    cmcSearch.Top = 30
    edcSearch.Top = cmcSearch.Top
    imIgnoreScroll = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    imFirstActivate = True
    imInChg = True
    mPopulate
    mPopDNE
    mMoveRecToCtrls
    mSortCol 0
    imInChg = False
    imFieldChgd = False
    imLbcMouseDown = False
    imDoubleClickName = False
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(NETCUELIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdNetcue, grdNetcue, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdNetcue, grdNetcue, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Relay Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Relay Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub grdNetcue_DblClick()
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIOLIST) <> 2) Then
        Select Case grdNetcue.Col
            Case LIBNAMEINDEX
                igInitCallInfo = 1
                sgInitCallName = grdNetcue.TextMatrix(grdNetcue.Row, grdNetcue.Col)
                EngrAudioType.Show vbModal
                cmcCancel.SetFocus
        End Select
    End If
End Sub

Private Sub imcInsert_Click()
    mSetShow
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    igRptIndex = NETCUE_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub lbcDNE_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        'lbcDNE.Visible = False
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcDNE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcDNE.Visible = False
End Sub

Private Sub lbcDNE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llCode As Long
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcDNE, y)
    If (llRow < lbcDNE.ListCount) And (lbcDNE.ListCount > 0) And (llRow <> -1) Then
        llCode = lbcDNE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrDNE) - 1 Step 1
            If llCode = tgCurrDNE(ilLoop).lCode Then
                lbcDNE.ToolTipText = Trim$(tgCurrDNE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub



Private Sub grdNetcue_Click()
    If grdNetcue.Col >= grdNetcue.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdNetcue_EnterCell()
    mSetShow
End Sub

Private Sub grdNetcue_GotFocus()
    If grdNetcue.Col >= grdNetcue.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdNetcue_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdNetcue.TopRow
    grdNetcue.Redraw = False
End Sub

Private Sub grdNetcue_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdNetcue.RowHeight(0) Then
        mSortCol grdNetcue.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdNetcue, x, y)
    If Not ilFound Then
        grdNetcue.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdNetcue.Col >= grdNetcue.Cols - 1 Then
        grdNetcue.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdNetcue.TopRow
    DoEvents
    llRow = grdNetcue.Row
    If grdNetcue.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdNetcue.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdNetcue.TextMatrix(llRow, NAMEINDEX) = ""
        grdNetcue.Row = llRow + 1
        grdNetcue.Col = NAMEINDEX
        grdNetcue.Redraw = True
    End If
    grdNetcue.Redraw = True
    mEnableBox
End Sub

Private Sub grdNetcue_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdNetcue.Redraw = False Then
        grdNetcue.Redraw = True
        If lmTopRow < grdNetcue.FixedRows Then
            grdNetcue.TopRow = grdNetcue.FixedRows
        Else
            grdNetcue.TopRow = lmTopRow
        End If
        grdNetcue.Refresh
        grdNetcue.Redraw = False
    End If
    If (imShowGridBox) And (grdNetcue.Row >= grdNetcue.FixedRows) And (grdNetcue.Col >= 0) And (grdNetcue.Col < grdNetcue.Cols - 1) Then
        If grdNetcue.RowIsVisible(grdNetcue.Row) Then
            'edcGrid.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 30, grdNetcue.ColWidth(grdNetcue.Col) - 30, grdNetcue.RowHeight(grdNetcue.Row) - 30
            pbcArrow.Move grdNetcue.Left - pbcArrow.Width - 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + (grdNetcue.RowHeight(grdNetcue.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            lbcDNE.Visible = False
            cmcDropDown.Visible = False
            edcDropdown.Visible = False
            pbcArrow.Visible = False
            edcGrid.Visible = False
            pbcState.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub pbcSTab_GotFocus()
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If edcGrid.Visible Or edcDropdown.Visible Or pbcState.Visible Then
        If Not mBranch() Then
            mEnableBox
            Exit Sub
        End If
        mSetShow
        If grdNetcue.Col = NAMEINDEX Then
            If grdNetcue.Row > grdNetcue.FixedRows Then
                lmTopRow = -1
                grdNetcue.Row = grdNetcue.Row - 1
                If Not grdNetcue.RowIsVisible(grdNetcue.Row) Then
                    grdNetcue.TopRow = grdNetcue.TopRow - 1
                End If
                grdNetcue.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdNetcue.Col = grdNetcue.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdNetcue.TopRow = grdNetcue.FixedRows
        grdNetcue.Col = NAMEINDEX
        grdNetcue.Row = grdNetcue.FixedRows
        mEnableBox
    End If
End Sub

Private Sub pbcState_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
        If smState <> "Active" Then
            imFieldChgd = True
        End If
        smState = "Active"
        pbcState_Paint
        grdNetcue.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdNetcue.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdNetcue.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdNetcue.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdNetcue.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdNetcue.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcState_Paint()
    pbcState.Cls
    pbcState.CurrentX = 30  'fgBoxInsetX
    pbcState.CurrentY = 0 'fgBoxInsetY
    pbcState.Print smState
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim llEnableRow As Long
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If edcGrid.Visible Or edcDropdown.Visible Or pbcState.Visible Then
        If Not mBranch() Then
            mEnableBox
            Exit Sub
        End If
        llEnableRow = lmEnableRow
        mSetShow
        If grdNetcue.Col = STATEINDEX Then
            llRow = grdNetcue.Rows
            Do
                llRow = llRow - 1
            Loop While grdNetcue.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdNetcue.Row + 1 < llRow) Then
                lmTopRow = -1
                grdNetcue.Row = grdNetcue.Row + 1
                If Not grdNetcue.RowIsVisible(grdNetcue.Row) Then
                    imIgnoreScroll = True
                    grdNetcue.TopRow = grdNetcue.TopRow + 1
                End If
                grdNetcue.Col = NAMEINDEX
                'grdNetcue.TextMatrix(grdNetcue.Row, CODEINDEX) = 0
                If Trim$(grdNetcue.TextMatrix(grdNetcue.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdNetcue.Left - pbcArrow.Width - 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + (grdNetcue.RowHeight(grdNetcue.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdNetcue.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdNetcue.Row + 1 >= grdNetcue.Rows Then
                        grdNetcue.AddItem ""
                    End If
                    grdNetcue.Row = grdNetcue.Row + 1
                    If Not grdNetcue.RowIsVisible(grdNetcue.Row) Then
                        imIgnoreScroll = True
                        grdNetcue.TopRow = grdNetcue.TopRow + 1
                    End If
                    grdNetcue.Col = NAMEINDEX
                    grdNetcue.TextMatrix(grdNetcue.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdNetcue.Left - pbcArrow.Width - 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + (grdNetcue.RowHeight(grdNetcue.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdNetcue.Col = grdNetcue.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdNetcue.TopRow = grdNetcue.FixedRows
        grdNetcue.Col = NAMEINDEX
        grdNetcue.Row = grdNetcue.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdNetcue.TopRow
    llRow = grdNetcue.Row
    slMsg = "Insert above " & Trim$(grdNetcue.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdNetcue.Redraw = False
    grdNetcue.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdNetcue.Row = llRow
    grdNetcue.Redraw = False
    grdNetcue.TopRow = llTRow
    grdNetcue.Redraw = True
    DoEvents
    grdNetcue.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdNetcue.TopRow
    llRow = grdNetcue.Row
    If (Val(grdNetcue.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdNetcue.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdNetcue.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdNetcue.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdNetcue.Redraw = False
    If (Val(grdNetcue.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdNetcue.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdNetcue.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdNetcue.AddItem ""
    grdNetcue.Redraw = False
    grdNetcue.TopRow = llTRow
    grdNetcue.Redraw = True
    DoEvents
    grdNetcue.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mBranch() As Integer
    Dim llRow As Long
    Dim slStr As String
    
    mBranch = True
    If (lmEnableRow >= grdNetcue.FixedRows) And (lmEnableRow < grdNetcue.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = Trim$(grdNetcue.TextMatrix(lmEnableRow, lmEnableCol))
        If (slStr <> "") And (StrComp(slStr, "[None]", vbTextCompare) <> 0) Then
            Select Case lmEnableCol
                Case NAMEINDEX
                Case LIBNAMEINDEX
                    'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcDNE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrDayName.Show vbModal
                        sgCurrDNEStamp = ""
                        mPopDNE
                        lbcDNE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcDNE, CLng(grdNetcue.Height / 2)
                        If lbcDNE.Top + lbcDNE.Height > cmcCancel.Top Then
                            lbcDNE.Top = edcDropdown.Top - lbcDNE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcDNE, slStr)
                            If llRow > 0 Then
                                lbcDNE.ListIndex = llRow
                                edcDropdown.text = lbcDNE.List(lbcDNE.ListIndex)
                                edcDropdown.SelStart = 0
                                edcDropdown.SelLength = Len(edcDropdown.text)
                            Else
                                mBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mBranch = False
                        End If
                    End If
                Case DESCRIPTIONINDEX
                Case STATEINDEX
            End Select
        End If
    End If
    imDoubleClickName = False
End Function

Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrNNE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdNetcue.FixedRows To grdNetcue.Rows - 1 Step 1
            slStr = Trim$(grdNetcue.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdNetcue.Row = llRow
                    Do While Not grdNetcue.RowIsVisible(grdNetcue.Row)
                        imIgnoreScroll = True
                        grdNetcue.TopRow = grdNetcue.TopRow + 1
                    Loop
                    grdNetcue.Col = NAMEINDEX
                    mEnableBox
                    Exit Sub
                End If
            End If
        Next llRow
    End If
    If (Not ilCreateNew) Or (Not cmcDone.Enabled) Then
        Exit Sub
    End If
    'Find first blank row
    For llRow = grdNetcue.FixedRows To grdNetcue.Rows - 1 Step 1
        slStr = Trim$(grdNetcue.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdNetcue.Row = llRow
            Do While Not grdNetcue.RowIsVisible(grdNetcue.Row)
                imIgnoreScroll = True
                grdNetcue.TopRow = grdNetcue.TopRow + 1
            Loop
            grdNetcue.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdNetcue.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Function mCompare(tlNew As NNE, tlOld As NNE) As Integer
    If StrComp(tlNew.sName, tlOld.sName, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sDescription, tlOld.sDescription, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sState, tlOld.sState, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.lDneCode <> tlOld.lDneCode) Then
        mCompare = False
        Exit Function
    End If
    mCompare = True
End Function

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case grdNetcue.Col
        Case LIBNAMEINDEX
            lbcDNE.Visible = False
    End Select
End Sub




Private Sub mSetFocus()
    Select Case grdNetcue.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 15, grdNetcue.ColWidth(grdNetcue.Col) - 30, grdNetcue.RowHeight(grdNetcue.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 15, grdNetcue.ColWidth(grdNetcue.Col) - 30, grdNetcue.RowHeight(grdNetcue.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case LIBNAMEINDEX
            edcDropdown.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 15, grdNetcue.ColWidth(grdNetcue.Col) - cmcDropDown.Width - 30, grdNetcue.RowHeight(grdNetcue.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcDNE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcDNE, CLng(grdNetcue.Height / 2)
            If lbcDNE.Top + lbcDNE.Height > cmcCancel.Top Then
                lbcDNE.Top = edcDropdown.Top - lbcDNE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcDNE.Visible = True
            edcDropdown.SetFocus
        Case STATEINDEX
            pbcState.Move grdNetcue.Left + grdNetcue.ColPos(grdNetcue.Col) + 30, grdNetcue.Top + grdNetcue.RowPos(grdNetcue.Row) + 15, grdNetcue.ColWidth(grdNetcue.Col) - 30, grdNetcue.RowHeight(grdNetcue.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
