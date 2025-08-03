VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrAudioType 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrAudioType.frx":0000
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
      Left            =   11715
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   45
   End
   Begin VB.PictureBox pbcYesNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5880
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1125
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
      TabIndex        =   7
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
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   6630
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
      Picture         =   "EngrAudioType.frx":030A
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
      Left            =   6990
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9645
      Top             =   6495
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
      Left            =   5220
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3435
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAudioType 
      Height          =   5790
      Left            =   345
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   10213
      _Version        =   393216
      Rows            =   3
      Cols            =   8
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
      _Band(0).Cols   =   8
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
      Left            =   9975
      TabIndex        =   13
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
      Left            =   8280
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2535
      Picture         =   "EngrAudioType.frx":0614
      Top             =   6510
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Audio Type"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1650
      Picture         =   "EngrAudioType.frx":091E
      Top             =   6510
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8790
      Picture         =   "EngrAudioType.frx":11E8
      Top             =   6510
      Width           =   480
   End
End
Attribute VB_Name = "EngrAudioType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrAudioType - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private smState As String
Private smTestItemID As String
Private imInChg As Integer
Private imBSMode As Integer
Private imATECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmATE As ATE

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
Const PREEVENTTIMEINDEX = 2
Const POSTEVENTTIMEINDEX = 3
Const TESTITEMIDINDEX = 4
Const STATEINDEX = 5
Const CODEINDEX = 6
Const USEDFLAGINDEX = 7

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    slStr = Trim$(edcSearch.text)
    llRow = gGrid_Search(grdAudioType, NAMEINDEX, slStr)
    If llRow >= 0 Then
        mEnableBox
    End If
End Sub

Private Sub cmcSearch_GotFocus()
    mSetShow
End Sub


Private Function mNameOk() As Integer
    Dim ilError As Integer
    Dim llRow As Long
    Dim llTestRow As Long
    Dim slStr As String
    Dim slTestStr As String
    
    grdAudioType.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdAudioType.FixedRows To grdAudioType.Rows - 1 Step 1
        slStr = Trim$(grdAudioType.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdAudioType.Rows - 1 Step 1
                slTestStr = Trim$(grdAudioType.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdAudioType.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdAudioType.Row = llRow
                        grdAudioType.Col = NAMEINDEX
                        grdAudioType.CellForeColor = vbRed
                    Else
                        grdAudioType.Row = llTestRow
                        grdAudioType.Col = NAMEINDEX
                        grdAudioType.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdAudioType.Redraw = True
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
    gGrid_SortByCol grdAudioType, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIOTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdAudioType.Row >= grdAudioType.FixedRows) And (grdAudioType.Row < grdAudioType.Rows) And (grdAudioType.Col >= 0) And (grdAudioType.Col < grdAudioType.Cols - 1) Then
        lmEnableRow = grdAudioType.Row
        lmEnableCol = grdAudioType.Col
        sgReturnCallName = grdAudioType.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdAudioType.Left - pbcArrow.Width - 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + (grdAudioType.RowHeight(grdAudioType.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdAudioType.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdAudioType.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdAudioType.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdAudioType.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
                edcGrid.MaxLength = Len(tmATE.sName)
                edcGrid.text = grdAudioType.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
                edcGrid.MaxLength = Len(tmATE.sDescription)
                edcGrid.text = grdAudioType.text
                edcGrid.Visible = True
                edcGrid.SetFocus
             Case PREEVENTTIMEINDEX  'Pre-Event Time
                edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
                edcGrid.MaxLength = 13
                edcGrid.text = grdAudioType.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case POSTEVENTTIMEINDEX  'Post Event Time
                edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
                edcGrid.MaxLength = 13
                edcGrid.text = grdAudioType.text
                edcGrid.Visible = True
                edcGrid.SetFocus
           Case TESTITEMIDINDEX
                pbcYesNo.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
                smTestItemID = grdAudioType.text
                If (Trim$(smTestItemID) = "") Or (smTestItemID = "Missing") Then
                    smTestItemID = "No"
                End If
                pbcYesNo.Visible = True
                pbcYesNo.SetFocus
           Case STATEINDEX
                pbcState.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
                smState = grdAudioType.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdAudioType.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdAudioType.FixedRows) And (lmEnableRow < grdAudioType.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
            Case PREEVENTTIMEINDEX
            Case POSTEVENTTIMEINDEX
            Case TESTITEMIDINDEX
                grdAudioType.TextMatrix(lmEnableRow, lmEnableCol) = smTestItemID
                If (Trim$(grdAudioType.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdAudioType.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdAudioType.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdAudioType.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdAudioType.TextMatrix(lmEnableRow, NAMEINDEX)
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    edcGrid.Visible = False
    pbcState.Visible = False
    pbcYesNo.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub
Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llRow As Long
    
    grdAudioType.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdAudioType.FixedRows To grdAudioType.Rows - 1 Step 1
        slStr = Trim$(grdAudioType.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdAudioType.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdAudioType.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdAudioType.Row = llRow
                grdAudioType.Col = NAMEINDEX
                grdAudioType.CellForeColor = vbRed
            End If
            slStr = grdAudioType.TextMatrix(llRow, PREEVENTTIMEINDEX)
            If slStr <> "" Then
                ilError = True
                grdAudioType.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdAudioType.Row = llRow
                grdAudioType.Col = NAMEINDEX
                grdAudioType.CellForeColor = vbRed
            End If
            slStr = grdAudioType.TextMatrix(llRow, POSTEVENTTIMEINDEX)
            If slStr <> "" Then
                ilError = True
                grdAudioType.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdAudioType.Row = llRow
                grdAudioType.Col = NAMEINDEX
                grdAudioType.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdAudioType.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdAudioType.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdAudioType.Row = llRow
                    grdAudioType.Col = STATEINDEX
                    grdAudioType.CellForeColor = vbRed
                End If
                slStr = grdAudioType.TextMatrix(llRow, PREEVENTTIMEINDEX)
                If Not gIsTimeTenths(slStr) Then
                    ilError = True
                    grdAudioType.Row = llRow
                    grdAudioType.Col = PREEVENTTIMEINDEX
                    grdAudioType.CellForeColor = vbRed
                End If
                slStr = grdAudioType.TextMatrix(llRow, POSTEVENTTIMEINDEX)
                If Not gIsTimeTenths(slStr) Then
                    ilError = True
                    grdAudioType.Row = llRow
                    grdAudioType.Col = POSTEVENTTIMEINDEX
                    grdAudioType.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdAudioType.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdAudioType
    mGridColumnWidth
    'Set Titles
    grdAudioType.TextMatrix(0, NAMEINDEX) = "Name"
    grdAudioType.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdAudioType.TextMatrix(0, PREEVENTTIMEINDEX) = "Pre-Event Time"
    grdAudioType.TextMatrix(0, POSTEVENTTIMEINDEX) = "Post-Event Time"
    grdAudioType.TextMatrix(0, TESTITEMIDINDEX) = "Test Item ID"
    grdAudioType.TextMatrix(0, STATEINDEX) = "State"
    grdAudioType.Row = 1
    For ilCol = 0 To grdAudioType.Cols - 1 Step 1
        grdAudioType.Col = ilCol
        grdAudioType.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdAudioType.Height = cmcCancel.Top - grdAudioType.Top - 120    '8 * grdAudioType.RowHeight(0) + 30
    gGrid_IntegralHeight grdAudioType
    gGrid_Clear grdAudioType, True
    grdAudioType.Row = grdAudioType.FixedRows

End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdAudioType.Width = EngrAudioType.Width - 2 * grdAudioType.Left
    grdAudioType.ColWidth(CODEINDEX) = 0
    grdAudioType.ColWidth(USEDFLAGINDEX) = 0
    grdAudioType.ColWidth(NAMEINDEX) = grdAudioType.Width / 6
    grdAudioType.ColWidth(PREEVENTTIMEINDEX) = grdAudioType.Width / 8
    grdAudioType.ColWidth(POSTEVENTTIMEINDEX) = grdAudioType.Width / 8
    grdAudioType.ColWidth(TESTITEMIDINDEX) = grdAudioType.Width / 8
    grdAudioType.ColWidth(STATEINDEX) = grdAudioType.Width / 15
    grdAudioType.ColWidth(DESCRIPTIONINDEX) = grdAudioType.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdAudioType.ColWidth(DESCRIPTIONINDEX) > grdAudioType.ColWidth(ilCol) Then
                grdAudioType.ColWidth(DESCRIPTIONINDEX) = grdAudioType.ColWidth(DESCRIPTIONINDEX) - grdAudioType.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdAudioType, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdAudioType.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdAudioType.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmATE.iCode = Val(grdAudioType.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdAudioType.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmATE.sName = ""
    Else
        tmATE.sName = slStr
    End If
    tmATE.sDescription = grdAudioType.TextMatrix(llRow, DESCRIPTIONINDEX)
    slStr = Trim$(grdAudioType.TextMatrix(llRow, PREEVENTTIMEINDEX))
    If slStr = "" Then
        slStr = "00:00:00"
    End If
    tmATE.lPreBufferTime = gStrTimeInTenthToLong(slStr, False)
    slStr = Trim$(grdAudioType.TextMatrix(llRow, POSTEVENTTIMEINDEX))
    If slStr = "" Then
        slStr = "00:00:00"
    End If
    tmATE.lPostBufferTime = gStrTimeInTenthToLong(slStr, False)
    If grdAudioType.TextMatrix(llRow, TESTITEMIDINDEX) = "Yes" Then
        tmATE.sTestItemID = "Y"
    Else
        tmATE.sTestItemID = "N"
    End If
    If grdAudioType.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmATE.sState = "D"
    Else
        tmATE.sState = "A"
    End If
    If tmATE.iCode <= 0 Then
        tmATE.sUsedFlag = "N"
    Else
        tmATE.sUsedFlag = grdAudioType.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmATE.iVersion = 0
    tmATE.iOrigAteCode = tmATE.iCode
    tmATE.sCurrent = "Y"
    'tmATE.sEnteredDate = smNowDate
    'tmATE.sEnteredTime = smNowTime
    tmATE.sEnteredDate = Format(Now, sgShowDateForm)
    tmATE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmATE.iUieCode = tgUIE.iCode
    tmATE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim slStr As String
    
    'gGrid_Clear grdAudioType, True
    llRow = grdAudioType.FixedRows
    For ilLoop = 0 To UBound(tgCurrATE) - 1 Step 1
        If llRow + 1 > grdAudioType.Rows Then
            grdAudioType.AddItem ""
        End If
        grdAudioType.Row = llRow
        grdAudioType.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrATE(ilLoop).sName)
        grdAudioType.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrATE(ilLoop).sDescription)
        If tgCurrATE(ilLoop).lPreBufferTime = 0 Then
            grdAudioType.TextMatrix(llRow, PREEVENTTIMEINDEX) = ""
        Else
            grdAudioType.TextMatrix(llRow, PREEVENTTIMEINDEX) = gLongToStrTimeInTenth(tgCurrATE(ilLoop).lPreBufferTime)
        End If
        If tgCurrATE(ilLoop).lPostBufferTime = 0 Then
            grdAudioType.TextMatrix(llRow, POSTEVENTTIMEINDEX) = ""
        Else
            grdAudioType.TextMatrix(llRow, POSTEVENTTIMEINDEX) = gLongToStrTimeInTenth(tgCurrATE(ilLoop).lPostBufferTime)
        End If
        If tgCurrATE(ilLoop).sTestItemID = "Y" Then
            grdAudioType.TextMatrix(llRow, TESTITEMIDINDEX) = "Yes"
        Else
            grdAudioType.TextMatrix(llRow, TESTITEMIDINDEX) = "No"
        End If
        If tgCurrATE(ilLoop).sState = "A" Then
            grdAudioType.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdAudioType.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdAudioType.TextMatrix(llRow, CODEINDEX) = tgCurrATE(ilLoop).iCode
        grdAudioType.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrATE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdAudioType.Rows Then
        grdAudioType.AddItem ""
    End If
    grdAudioType.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlATE As ATE
    
    gSetMousePointer grdAudioType, grdAudioType, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdAudioType, grdAudioType, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdAudioType, grdAudioType, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdAudioType.Redraw = False
    For llRow = grdAudioType.FixedRows To grdAudioType.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmATE.sName) <> "" Then
            imATECode = tmATE.iCode
            If tmATE.iCode > 0 Then
                ilRet = gGetRec_ATE_AudioType(imATECode, "Audio Types-mSave: Get ATE", tlATE)
                If ilRet Then
                    If mCompare(tmATE, tlATE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmATE.iVersion = tlATE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmATE.iCode <= 0 Then
                    ilRet = gPutInsert_ATE_AudioType(0, tmATE, "Audio Types-mSave: Insert ATE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_ATE_AudioType(1, tmATE, "Audio Types-mSave: Update ATE")
                    ilRet = gPutDelete_ATE_AudioType(tmATE.iCode, "Audio Types-mSave: Delete ATE")
                    ilRet = gPutInsert_ATE_AudioType(1, tmATE, "Audio Types-mSave: Insert ATE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_ATE_AudioType(imDeleteCodes(ilLoop), "EngrAudioType- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdAudioType.Redraw = True
    sgCurrATEStamp = ""
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrAudioType
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrAudioType
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdAudioType, grdAudioType, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdAudioType, grdAudioType, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdAudioType, grdAudioType, vbDefault
    Unload EngrAudioType
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    Dim llTopRow As Long
    
    If imFieldChgd = True Then
        gSetMousePointer grdAudioType, grdAudioType, vbHourglass
        llTopRow = grdAudioType.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdAudioType, grdAudioType, vbDefault
            Exit Sub
        End If
        grdAudioType.Redraw = False
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
        grdAudioType.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdAudioType, grdAudioType, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdAudioType.Col
        Case NAMEINDEX
            If grdAudioType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdAudioType.text = edcGrid.text
            grdAudioType.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdAudioType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdAudioType.text = edcGrid.text
            grdAudioType.CellForeColor = vbBlack
        Case PREEVENTTIMEINDEX
            If grdAudioType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdAudioType.text = edcGrid.text
            grdAudioType.CellForeColor = vbBlack
        Case POSTEVENTTIMEINDEX
            If grdAudioType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdAudioType.text = edcGrid.text
            grdAudioType.CellForeColor = vbBlack
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
    gSetFonts EngrAudioType
    gCenterFormModal EngrAudioType
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdAudioType.FixedRows) And (lmEnableRow < grdAudioType.Rows) Then
            If (lmEnableCol >= grdAudioType.FixedCols) And (lmEnableCol < grdAudioType.Cols) Then
                If lmEnableCol = TESTITEMIDINDEX Then
                    smTestItemID = smESCValue
                ElseIf lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdAudioType.text = smESCValue
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
    grdAudioType.Height = cmcCancel.Top - grdAudioType.Top - 120    '8 * grdAudioType.RowHeight(0) + 30
    gGrid_IntegralHeight grdAudioType
    gGrid_FillWithRows grdAudioType
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrAudioType = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdAudioType, grdAudioType, vbHourglass
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
    mMoveRecToCtrls
    mSortCol 0
    imInChg = False
    imFieldChgd = False
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOTYPELIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdAudioType, grdAudioType, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdAudioType, grdAudioType, vbDefault
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

Private Sub imcInsert_Click()
    mSetShow
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    igRptIndex = AUDIOTYPE_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub grdAudioType_Click()
    If grdAudioType.Col >= grdAudioType.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAudioType_EnterCell()
    mSetShow
End Sub

Private Sub grdAudioType_GotFocus()
    If grdAudioType.Col >= grdAudioType.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAudioType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdAudioType.TopRow
    grdAudioType.Redraw = False
End Sub

Private Sub grdAudioType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdAudioType.RowHeight(0) Then
        mSortCol grdAudioType.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdAudioType, x, y)
    If Not ilFound Then
        grdAudioType.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdAudioType.Col >= grdAudioType.Cols - 1 Then
        grdAudioType.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdAudioType.TopRow
    DoEvents
    llRow = grdAudioType.Row
    If grdAudioType.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdAudioType.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdAudioType.TextMatrix(llRow, NAMEINDEX) = ""
        grdAudioType.Row = llRow + 1
        grdAudioType.Col = NAMEINDEX
        grdAudioType.Redraw = True
    End If
    grdAudioType.Redraw = True
    mEnableBox
End Sub

Private Sub grdAudioType_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdAudioType.Redraw = False Then
        grdAudioType.Redraw = True
        If lmTopRow < grdAudioType.FixedRows Then
            grdAudioType.TopRow = grdAudioType.FixedRows
        Else
            grdAudioType.TopRow = lmTopRow
        End If
        grdAudioType.Refresh
        grdAudioType.Redraw = False
    End If
    If (imShowGridBox) And (grdAudioType.Row >= grdAudioType.FixedRows) And (grdAudioType.Col >= 0) And (grdAudioType.Col < grdAudioType.Cols - 1) Then
        If grdAudioType.RowIsVisible(grdAudioType.Row) Then
            'edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 30, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 30
            pbcArrow.Move grdAudioType.Left - pbcArrow.Width - 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + (grdAudioType.RowHeight(grdAudioType.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            pbcSetFocus.SetFocus
            pbcArrow.Visible = False
            edcGrid.Visible = False
            pbcState.Visible = False
            pbcYesNo.Visible = False
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
    If edcGrid.Visible Or pbcState.Visible Or pbcYesNo.Visible Then
        mSetShow
        If grdAudioType.Col = NAMEINDEX Then
            If grdAudioType.Row > grdAudioType.FixedRows Then
                lmTopRow = -1
                grdAudioType.Row = grdAudioType.Row - 1
                If Not grdAudioType.RowIsVisible(grdAudioType.Row) Then
                    grdAudioType.TopRow = grdAudioType.TopRow - 1
                End If
                grdAudioType.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdAudioType.Col = grdAudioType.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdAudioType.TopRow = grdAudioType.FixedRows
        grdAudioType.Col = NAMEINDEX
        grdAudioType.Row = grdAudioType.FixedRows
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
        grdAudioType.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdAudioType.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdAudioType.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdAudioType.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdAudioType.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdAudioType.CellForeColor = vbBlack
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
    If edcGrid.Visible Or pbcState.Visible Or pbcYesNo.Visible Then
        llEnableRow = lmEnableRow
        mSetShow
        If grdAudioType.Col = STATEINDEX Then
            llRow = grdAudioType.Rows
            Do
                llRow = llRow - 1
            Loop While grdAudioType.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdAudioType.Row + 1 < llRow) Then
                lmTopRow = -1
                grdAudioType.Row = grdAudioType.Row + 1
                If Not grdAudioType.RowIsVisible(grdAudioType.Row) Then
                    imIgnoreScroll = True
                    grdAudioType.TopRow = grdAudioType.TopRow + 1
                End If
                grdAudioType.Col = NAMEINDEX
                'grdAudioType.TextMatrix(grdAudioType.Row, CODEINDEX) = 0
                If Trim$(grdAudioType.TextMatrix(grdAudioType.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdAudioType.Left - pbcArrow.Width - 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + (grdAudioType.RowHeight(grdAudioType.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdAudioType.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdAudioType.Row + 1 >= grdAudioType.Rows Then
                        grdAudioType.AddItem ""
                    End If
                    grdAudioType.Row = grdAudioType.Row + 1
                    If Not grdAudioType.RowIsVisible(grdAudioType.Row) Then
                        imIgnoreScroll = True
                        grdAudioType.TopRow = grdAudioType.TopRow + 1
                    End If
                    grdAudioType.Col = NAMEINDEX
                    grdAudioType.TextMatrix(grdAudioType.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdAudioType.Left - pbcArrow.Width - 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + (grdAudioType.RowHeight(grdAudioType.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdAudioType.Col = grdAudioType.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdAudioType.TopRow = grdAudioType.FixedRows
        grdAudioType.Col = NAMEINDEX
        grdAudioType.Row = grdAudioType.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdAudioType.TopRow
    llRow = grdAudioType.Row
    slMsg = "Insert above " & Trim$(grdAudioType.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdAudioType.Redraw = False
    grdAudioType.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdAudioType.Row = llRow
    grdAudioType.Redraw = False
    grdAudioType.TopRow = llTRow
    grdAudioType.Redraw = True
    DoEvents
    grdAudioType.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdAudioType.TopRow
    llRow = grdAudioType.Row
    If (Val(grdAudioType.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdAudioType.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdAudioType.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdAudioType.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdAudioType.Redraw = False
    If (Val(grdAudioType.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdAudioType.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdAudioType.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdAudioType.AddItem ""
    grdAudioType.Redraw = False
    grdAudioType.TopRow = llTRow
    grdAudioType.Redraw = True
    DoEvents
    grdAudioType.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrATE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdAudioType.FixedRows To grdAudioType.Rows - 1 Step 1
            slStr = Trim$(grdAudioType.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdAudioType.Row = llRow
                    Do While Not grdAudioType.RowIsVisible(grdAudioType.Row)
                        imIgnoreScroll = True
                        grdAudioType.TopRow = grdAudioType.TopRow + 1
                    Loop
                    grdAudioType.Col = NAMEINDEX
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
    For llRow = grdAudioType.FixedRows To grdAudioType.Rows - 1 Step 1
        slStr = Trim$(grdAudioType.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdAudioType.Row = llRow
            Do While Not grdAudioType.RowIsVisible(grdAudioType.Row)
                imIgnoreScroll = True
                grdAudioType.TopRow = grdAudioType.TopRow + 1
            Loop
            grdAudioType.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdAudioType.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Function mCompare(tlNew As ATE, tlOld As ATE) As Integer
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
    If tlNew.lPreBufferTime <> tlOld.lPreBufferTime Then
        mCompare = False
        Exit Function
    End If
    If tlNew.lPostBufferTime <> tlOld.lPostBufferTime Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sTestItemID, tlOld.sTestItemID, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    mCompare = True
End Function

Private Sub pbcYesNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If smTestItemID <> "Yes" Then
            imFieldChgd = True
        End If
        smTestItemID = "Yes"
        pbcYesNo_Paint
        grdAudioType.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If smTestItemID <> "No" Then
            imFieldChgd = True
        End If
        smTestItemID = "No"
        pbcYesNo_Paint
        grdAudioType.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smTestItemID = "Yes" Then
            imFieldChgd = True
            smTestItemID = "No"
            pbcYesNo_Paint
            grdAudioType.CellForeColor = vbBlack
        ElseIf smTestItemID = "No" Then
            imFieldChgd = True
            smTestItemID = "Yes"
            pbcYesNo_Paint
            grdAudioType.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcYesNo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smTestItemID = "Yes" Then
        imFieldChgd = True
        smTestItemID = "No"
        pbcYesNo_Paint
        grdAudioType.CellForeColor = vbBlack
    ElseIf smTestItemID = "No" Then
        imFieldChgd = True
        smTestItemID = "Yes"
        pbcYesNo_Paint
        grdAudioType.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcYesNo_Paint()
    pbcYesNo.Cls
    pbcYesNo.CurrentX = 30  'fgBoxInsetX
    pbcYesNo.CurrentY = 0 'fgBoxInsetY
    pbcYesNo.Print smTestItemID
End Sub

Private Sub mSetFocus()
    Select Case grdAudioType.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
         Case PREEVENTTIMEINDEX  'Pre-Event Time
            edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case POSTEVENTTIMEINDEX  'Post Event Time
            edcGrid.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
       Case TESTITEMIDINDEX
            pbcYesNo.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
            pbcYesNo.Visible = True
            pbcYesNo.SetFocus
       Case STATEINDEX
            pbcState.Move grdAudioType.Left + grdAudioType.ColPos(grdAudioType.Col) + 30, grdAudioType.Top + grdAudioType.RowPos(grdAudioType.Row) + 15, grdAudioType.ColWidth(grdAudioType.Col) - 30, grdAudioType.RowHeight(grdAudioType.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
