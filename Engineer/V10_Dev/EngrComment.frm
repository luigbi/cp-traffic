VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrComment 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11775
   ControlBox      =   0   'False
   Icon            =   "EngrComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11775
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   11700
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   15
      Width           =   45
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
      TabIndex        =   6
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
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   6705
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
      Picture         =   "EngrComment.frx":030A
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
      Left            =   6915
      TabIndex        =   10
      Top             =   6690
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9570
      Top             =   6585
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7290
      FormDesignWidth =   11775
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5145
      TabIndex        =   9
      Top             =   6690
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   6690
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdComment 
      Height          =   5925
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   10451
      _Version        =   393216
      Rows            =   3
      Cols            =   5
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
      _Band(0).Cols   =   5
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
      Left            =   9990
      TabIndex        =   12
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
      Left            =   8295
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2460
      Picture         =   "EngrComment.frx":0614
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Comment"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1575
      Picture         =   "EngrComment.frx":091E
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8715
      Picture         =   "EngrComment.frx":11E8
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "EngrComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrComment - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private hmCTE As Integer

Private imFieldChgd As Integer
Private smState As String
Private imInChg As Integer
Private imBSMode As Integer
Private lmCTECode As Long
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmCTE As CTE

Private lmDeleteCodes() As Long

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
Const STATEINDEX = 2
Const CODEINDEX = 3
Const USEDFLAGINDEX = 4

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcSearch_Click()
    Dim llRow As Long
    Dim slStr As String
    slStr = Trim$(edcSearch.text)
    llRow = gGrid_Search(grdComment, NAMEINDEX, slStr)
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
    
    grdComment.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdComment.FixedRows To grdComment.Rows - 1 Step 1
        slStr = Trim$(grdComment.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdComment.Rows - 1 Step 1
                slTestStr = Trim$(grdComment.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdComment.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdComment.Row = llRow
                        grdComment.Col = NAMEINDEX
                        grdComment.CellForeColor = vbRed
                    Else
                        grdComment.Row = llTestRow
                        grdComment.Col = NAMEINDEX
                        grdComment.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdComment.Redraw = True
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
    gGrid_SortByCol grdComment, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(COMMENTLIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdComment.Row >= grdComment.FixedRows) And (grdComment.Row < grdComment.Rows) And (grdComment.Col >= 0) And (grdComment.Col < grdComment.Cols - 1) Then
        lmEnableRow = grdComment.Row
        lmEnableCol = grdComment.Col
        sgReturnCallName = grdComment.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdComment.Left - pbcArrow.Width - 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + (grdComment.RowHeight(grdComment.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdComment.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdComment.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdComment.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdComment.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
                edcGrid.MaxLength = Len(tmCTE.sName)
                edcGrid.text = grdComment.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
                If igInitCallInfo = 1 Then
                    edcGrid.MaxLength = 66
                ElseIf (igInitCallInfo = 3) Then
                    edcGrid.MaxLength = Len(tmCTE.sComment)
                Else
                    edcGrid.MaxLength = 90
                End If
                    
                edcGrid.text = grdComment.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
                smState = grdComment.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdComment.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdComment.FixedRows) And (lmEnableRow < grdComment.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
            Case STATEINDEX
                grdComment.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdComment.TextMatrix(lmEnableRow, NAMEINDEX)
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    edcGrid.Visible = False
    pbcState.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub
Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llRow As Long
    
    grdComment.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdComment.FixedRows To grdComment.Rows - 1 Step 1
        slStr = Trim$(grdComment.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdComment.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdComment.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdComment.Row = llRow
                grdComment.Col = NAMEINDEX
                grdComment.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdComment.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdComment.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdComment.Row = llRow
                    grdComment.Col = STATEINDEX
                    grdComment.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdComment.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdComment
    mGridColumnWidth
    'Set Titles
    grdComment.TextMatrix(0, NAMEINDEX) = "Name"
    grdComment.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdComment.TextMatrix(0, STATEINDEX) = "State"
    grdComment.Row = 1
    For ilCol = 0 To grdComment.Cols - 1 Step 1
        grdComment.Col = ilCol
        grdComment.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdComment.Height = cmcCancel.Top - grdComment.Top - 120    '8 * grdComment.RowHeight(0) + 30
    gGrid_IntegralHeight grdComment
    gGrid_Clear grdComment, True
    grdComment.Row = grdComment.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdComment.Width = EngrComment.Width - 2 * grdComment.Left
    grdComment.ColWidth(CODEINDEX) = 0
    grdComment.ColWidth(USEDFLAGINDEX) = 0
    grdComment.ColWidth(NAMEINDEX) = grdComment.Width / 9
    grdComment.ColWidth(STATEINDEX) = grdComment.Width / 15
    grdComment.ColWidth(DESCRIPTIONINDEX) = grdComment.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdComment.ColWidth(DESCRIPTIONINDEX) > grdComment.ColWidth(ilCol) Then
                grdComment.ColWidth(DESCRIPTIONINDEX) = grdComment.ColWidth(DESCRIPTIONINDEX) - grdComment.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdComment, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdComment.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdComment.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmCTE.lCode = Val(grdComment.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdComment.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmCTE.sName = ""
    Else
        tmCTE.sName = slStr
    End If
    tmCTE.sComment = grdComment.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdComment.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmCTE.sState = "D"
    Else
        tmCTE.sState = "A"
    End If
    If igInitCallInfo = 1 Then
        tmCTE.sType = "T1"
    ElseIf (igInitCallInfo = 3) Then
        tmCTE.sType = "DH"
    Else
        tmCTE.sType = "T2"
    End If
    If tmCTE.lCode <= 0 Then
        tmCTE.sUsedFlag = "N"
    Else
        tmCTE.sUsedFlag = grdComment.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmCTE.iVersion = 0
    tmCTE.lOrigCteCode = tmCTE.lCode
    tmCTE.sCurrent = "Y"
    'tmCTE.sEnteredDate = smNowDate
    'tmCTE.sEnteredTime = smNowTime
    tmCTE.sEnteredDate = Format(Now, sgShowDateForm)
    tmCTE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmCTE.iUieCode = tgUIE.iCode
    tmCTE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdComment, True
    llRow = grdComment.FixedRows
    For ilLoop = 0 To UBound(tgCurrCTE) - 1 Step 1
        If llRow + 1 > grdComment.Rows Then
            grdComment.AddItem ""
        End If
        grdComment.Row = llRow
        grdComment.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrCTE(ilLoop).sName)
        grdComment.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrCTE(ilLoop).sComment)
        If tgCurrCTE(ilLoop).sState = "A" Then
            grdComment.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdComment.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdComment.TextMatrix(llRow, CODEINDEX) = tgCurrCTE(ilLoop).lCode
        grdComment.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrCTE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdComment.Rows Then
        grdComment.AddItem ""
    End If
    grdComment.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    If igInitCallInfo = 1 Then
        ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T1", sgCurrCTEStamp, "EngrComment-mPopulate", tgCurrCTE())
    ElseIf (igInitCallInfo = 3) Then
        ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "DH", sgCurrCTEStamp, "EngrComment-mPopulate", tgCurrCTE())
    Else
        ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T2", sgCurrCTEStamp, "EngrComment-mPopulate", tgCurrCTE())
    End If
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlCTE As CTE
    
    gSetMousePointer grdComment, grdComment, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdComment, grdComment, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdComment, grdComment, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdComment.Redraw = False
    For llRow = grdComment.FixedRows To grdComment.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmCTE.sName) <> "" Then
            lmCTECode = tmCTE.lCode
            If tmCTE.lCode > 0 Then
                ilRet = gGetRec_CTE_CommtsTitle(lmCTECode, "Comment-mSave: Get CTE", tlCTE)
                If ilRet Then
                    If mCompare(tmCTE, tlCTE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmCTE.iVersion = tlCTE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmCTE.lCode <= 0 Then
                    ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "Comment-mSave: Insert CTE", hmCTE)
                Else
                    ilRet = gPutUpdate_CTE_CommtsTitle(1, tmCTE, "Comment-mSave: Update CTE", hmCTE)
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(lmDeleteCodes) To UBound(lmDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_CTE_CommtsTitle(lmDeleteCodes(ilLoop), "EngrComment- Delete")
    Next ilLoop
    ReDim lmDeleteCodes(0 To 0) As Long
    grdComment.Redraw = True
    sgCurrCTEStamp = ""
    If igInitCallInfo = 1 Then
        ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T1", sgCurrCTEStamp, "EngrComment-mPopulate", tgCurrCTE())
    ElseIf (igInitCallInfo = 3) Then
        ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "DH", sgCurrCTEStamp, "EngrComment-mPopulate", tgCurrCTE())
    Else
        ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T2", sgCurrCTEStamp, "EngrComment-mPopulate", tgCurrCTE())
    End If
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrComment
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrComment
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdComment, grdComment, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdComment, grdComment, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdComment, grdComment, vbDefault
    Unload EngrComment
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
        gSetMousePointer grdComment, grdComment, vbHourglass
        llTopRow = grdComment.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdComment, grdComment, vbDefault
            Exit Sub
        End If
        grdComment.Redraw = False
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
        grdComment.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdComment, grdComment, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdComment.Col
        Case NAMEINDEX
            If grdComment.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdComment.text = edcGrid.text
            grdComment.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdComment.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdComment.text = edcGrid.text
            grdComment.CellForeColor = vbBlack
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
    gSetFonts EngrComment
    gCenterFormModal EngrComment
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdComment.FixedRows) And (lmEnableRow < grdComment.Rows) Then
            If (lmEnableCol >= grdComment.FixedCols) And (lmEnableCol < grdComment.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdComment.text = smESCValue
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
    grdComment.Height = cmcCancel.Top - grdComment.Top - 120    '8 * grdComment.RowHeight(0) + 30
    gGrid_IntegralHeight grdComment
    gGrid_FillWithRows grdComment
End Sub

Private Sub Form_Unload(Cancel As Integer)
    btrDestroy hmCTE
    
    Erase lmDeleteCodes
    Set EngrComment = Nothing
End Sub





Private Sub mInit()
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    gSetMousePointer grdComment, grdComment, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    imcInsert.Picture = EngrMain!imcInsert.Picture
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    ReDim lmDeleteCodes(0 To 0) As Long
    hmCTE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCTE, "", sgDBPath & "CTE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(COMMENTLIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdComment, grdComment, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdComment, grdComment, vbDefault
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
    igRptIndex = COMMENT_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub grdComment_Click()
    If grdComment.Col >= grdComment.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdComment_EnterCell()
    mSetShow
End Sub

Private Sub grdComment_GotFocus()
    If grdComment.Col >= grdComment.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdComment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdComment.TopRow
    grdComment.Redraw = False
End Sub

Private Sub grdComment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdComment.RowHeight(0) Then
        mSortCol grdComment.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdComment, x, y)
    If Not ilFound Then
        grdComment.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdComment.Col >= grdComment.Cols - 1 Then
        grdComment.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdComment.TopRow
    DoEvents
    llRow = grdComment.Row
    If grdComment.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdComment.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdComment.TextMatrix(llRow, NAMEINDEX) = ""
        grdComment.Row = llRow + 1
        grdComment.Col = NAMEINDEX
        grdComment.Redraw = True
    End If
    grdComment.Redraw = True
    mEnableBox
End Sub

Private Sub grdComment_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdComment.Redraw = False Then
        grdComment.Redraw = True
        If lmTopRow < grdComment.FixedRows Then
            grdComment.TopRow = grdComment.FixedRows
        Else
            grdComment.TopRow = lmTopRow
        End If
        grdComment.Refresh
        grdComment.Redraw = False
    End If
    If (imShowGridBox) And (grdComment.Row >= grdComment.FixedRows) And (grdComment.Col >= 0) And (grdComment.Col < grdComment.Cols - 1) Then
        If grdComment.RowIsVisible(grdComment.Row) Then
            'edcGrid.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 30, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 30
            pbcArrow.Move grdComment.Left - pbcArrow.Width - 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + (grdComment.RowHeight(grdComment.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            edcGrid.Visible = False
            pbcArrow.Visible = False
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
    If edcGrid.Visible Or pbcState.Visible Then
        mSetShow
        If grdComment.Col = NAMEINDEX Then
            If grdComment.Row > grdComment.FixedRows Then
                lmTopRow = -1
                grdComment.Row = grdComment.Row - 1
                If Not grdComment.RowIsVisible(grdComment.Row) Then
                    grdComment.TopRow = grdComment.TopRow - 1
                End If
                grdComment.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdComment.Col = grdComment.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdComment.TopRow = grdComment.FixedRows
        grdComment.Col = NAMEINDEX
        grdComment.Row = grdComment.FixedRows
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
        grdComment.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdComment.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdComment.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdComment.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdComment.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdComment.CellForeColor = vbBlack
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
    If edcGrid.Visible Or pbcState.Visible Then
        llEnableRow = lmEnableRow
        mSetShow
        If grdComment.Col = STATEINDEX Then
            llRow = grdComment.Rows
            Do
                llRow = llRow - 1
            Loop While grdComment.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdComment.Row + 1 < llRow) Then
                lmTopRow = -1
                grdComment.Row = grdComment.Row + 1
                If Not grdComment.RowIsVisible(grdComment.Row) Then
                    imIgnoreScroll = True
                    grdComment.TopRow = grdComment.TopRow + 1
                End If
                grdComment.Col = NAMEINDEX
                'grdComment.TextMatrix(grdComment.Row, CODEINDEX) = 0
                If Trim$(grdComment.TextMatrix(grdComment.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdComment.Left - pbcArrow.Width - 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + (grdComment.RowHeight(grdComment.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdComment.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdComment.Row + 1 >= grdComment.Rows Then
                        grdComment.AddItem ""
                    End If
                    grdComment.Row = grdComment.Row + 1
                    If Not grdComment.RowIsVisible(grdComment.Row) Then
                        imIgnoreScroll = True
                        grdComment.TopRow = grdComment.TopRow + 1
                    End If
                    grdComment.Col = NAMEINDEX
                    grdComment.TextMatrix(grdComment.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdComment.Left - pbcArrow.Width - 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + (grdComment.RowHeight(grdComment.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdComment.Col = grdComment.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdComment.TopRow = grdComment.FixedRows
        grdComment.Col = NAMEINDEX
        grdComment.Row = grdComment.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdComment.TopRow
    llRow = grdComment.Row
    slMsg = "Insert above " & Trim$(grdComment.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdComment.Redraw = False
    grdComment.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdComment.Row = llRow
    grdComment.Redraw = False
    grdComment.TopRow = llTRow
    grdComment.Redraw = True
    DoEvents
    grdComment.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdComment.TopRow
    llRow = grdComment.Row
    If (Val(grdComment.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdComment.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdComment.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdComment.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdComment.Redraw = False
    If (Val(grdComment.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        lmDeleteCodes(UBound(lmDeleteCodes)) = Val(grdComment.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve lmDeleteCodes(0 To UBound(lmDeleteCodes) + 1) As Long
    End If
    grdComment.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdComment.AddItem ""
    grdComment.Redraw = False
    grdComment.TopRow = llTRow
    grdComment.Redraw = True
    DoEvents
    grdComment.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As CTE, tlOld As CTE) As Integer
    If StrComp(tlNew.sName, tlOld.sName, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sComment, tlOld.sComment, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
'    If StrComp(tlNew.sState, tlOld.sState, vbTextCompare) <> 0 Then
'        mCompare = False
'        Exit Function
'    End If
    mCompare = True
End Function

Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrCTE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdComment.FixedRows To grdComment.Rows - 1 Step 1
            slStr = Trim$(grdComment.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdComment.Row = llRow
                    Do While Not grdComment.RowIsVisible(grdComment.Row)
                        imIgnoreScroll = True
                        grdComment.TopRow = grdComment.TopRow + 1
                    Loop
                    grdComment.Col = NAMEINDEX
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
    For llRow = grdComment.FixedRows To grdComment.Rows - 1 Step 1
        slStr = Trim$(grdComment.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdComment.Row = llRow
            Do While Not grdComment.RowIsVisible(grdComment.Row)
                imIgnoreScroll = True
                grdComment.TopRow = grdComment.TopRow + 1
            Loop
            grdComment.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdComment.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mSetFocus()
    Select Case grdComment.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdComment.Left + grdComment.ColPos(grdComment.Col) + 30, grdComment.Top + grdComment.RowPos(grdComment.Row) + 15, grdComment.ColWidth(grdComment.Col) - 30, grdComment.RowHeight(grdComment.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
