VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrFollow 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrFollow.frx":0000
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
      Left            =   11730
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
      Left            =   165
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   6750
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
      Picture         =   "EngrFollow.frx":030A
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
      TabIndex        =   10
      Top             =   6630
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9630
      Top             =   6525
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
      TabIndex        =   9
      Top             =   6630
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3420
      TabIndex        =   8
      Top             =   6630
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFollow 
      Height          =   5850
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   10319
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
      Left            =   10035
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
      Left            =   8340
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2520
      Picture         =   "EngrFollow.frx":0614
      Top             =   6540
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Follow"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1635
      Picture         =   "EngrFollow.frx":091E
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8775
      Picture         =   "EngrFollow.frx":11E8
      Top             =   6540
      Width           =   480
   End
End
Attribute VB_Name = "EngrFollow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrFollow - enters affiliate representative information
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private smState As String
Private imInChg As Integer
Private imBSMode As Integer
Private imFNECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmFNE As FNE

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
    llRow = gGrid_Search(grdFollow, NAMEINDEX, slStr)
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
    
    grdFollow.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdFollow.FixedRows To grdFollow.Rows - 1 Step 1
        slStr = Trim$(grdFollow.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdFollow.Rows - 1 Step 1
                slTestStr = Trim$(grdFollow.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdFollow.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdFollow.Row = llRow
                        grdFollow.Col = NAMEINDEX
                        grdFollow.CellForeColor = vbRed
                    Else
                        grdFollow.Row = llTestRow
                        grdFollow.Col = NAMEINDEX
                        grdFollow.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdFollow.Redraw = True
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
    gGrid_SortByCol grdFollow, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(FOLLOWLIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdFollow.Row >= grdFollow.FixedRows) And (grdFollow.Row < grdFollow.Rows) And (grdFollow.Col >= 0) And (grdFollow.Col < grdFollow.Cols - 1) Then
        lmEnableRow = grdFollow.Row
        lmEnableCol = grdFollow.Col
        sgReturnCallName = grdFollow.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdFollow.Left - pbcArrow.Width - 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + (grdFollow.RowHeight(grdFollow.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdFollow.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdFollow.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdFollow.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdFollow.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdFollow.Left + grdFollow.ColPos(grdFollow.Col) + 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + 15, grdFollow.ColWidth(grdFollow.Col) - 30, grdFollow.RowHeight(grdFollow.Row) - 15
                'edcGrid.MaxLength = Len(tmFNE.sName)
                edcGrid.MaxLength = gGetAllowedChars("FOLLOW", Len(tmFNE.sName))
                edcGrid.text = grdFollow.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdFollow.Left + grdFollow.ColPos(grdFollow.Col) + 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + 15, grdFollow.ColWidth(grdFollow.Col) - 30, grdFollow.RowHeight(grdFollow.Row) - 15
                edcGrid.MaxLength = Len(tmFNE.sDescription)
                edcGrid.text = grdFollow.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdFollow.Left + grdFollow.ColPos(grdFollow.Col) + 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + 15, grdFollow.ColWidth(grdFollow.Col) - 30, grdFollow.RowHeight(grdFollow.Row) - 15
                smState = grdFollow.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdFollow.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdFollow.FixedRows) And (lmEnableRow < grdFollow.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdFollow.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdFollow.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdFollow.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdFollow.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdFollow.TextMatrix(lmEnableRow, NAMEINDEX)
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
    
    grdFollow.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdFollow.FixedRows To grdFollow.Rows - 1 Step 1
        slStr = Trim$(grdFollow.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdFollow.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdFollow.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdFollow.Row = llRow
                grdFollow.Col = NAMEINDEX
                grdFollow.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdFollow.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdFollow.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdFollow.Row = llRow
                    grdFollow.Col = STATEINDEX
                    grdFollow.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdFollow.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdFollow
    mGridColumnWidth
    'Set Titles
    grdFollow.TextMatrix(0, NAMEINDEX) = "Name"
    grdFollow.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdFollow.TextMatrix(0, STATEINDEX) = "State"
    grdFollow.Row = 1
    For ilCol = 0 To grdFollow.Cols - 1 Step 1
        grdFollow.Col = ilCol
        grdFollow.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdFollow.Height = cmcCancel.Top - grdFollow.Top - 120    '8 * grdFollow.RowHeight(0) + 30
    gGrid_IntegralHeight grdFollow
    gGrid_Clear grdFollow, True
    grdFollow.Row = grdFollow.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdFollow.Width = EngrFollow.Width - 2 * grdFollow.Left
    grdFollow.ColWidth(CODEINDEX) = 0
    grdFollow.ColWidth(USEDFLAGINDEX) = 0
    grdFollow.ColWidth(NAMEINDEX) = grdFollow.Width / 9
    grdFollow.ColWidth(STATEINDEX) = grdFollow.Width / 15
    grdFollow.ColWidth(DESCRIPTIONINDEX) = grdFollow.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdFollow.ColWidth(DESCRIPTIONINDEX) > grdFollow.ColWidth(ilCol) Then
                grdFollow.ColWidth(DESCRIPTIONINDEX) = grdFollow.ColWidth(DESCRIPTIONINDEX) - grdFollow.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdFollow, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdFollow.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdFollow.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmFNE.iCode = Val(grdFollow.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdFollow.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmFNE.sName = ""
    Else
        tmFNE.sName = slStr
    End If
    tmFNE.sDescription = grdFollow.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdFollow.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmFNE.sState = "D"
    Else
        tmFNE.sState = "A"
    End If
    If tmFNE.iCode <= 0 Then
        tmFNE.sUsedFlag = "N"
    Else
        tmFNE.sUsedFlag = grdFollow.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmFNE.iVersion = 0
    tmFNE.iOrigFneCode = tmFNE.iCode
    tmFNE.sCurrent = "Y"
    'tmFNE.sEnteredDate = smNowDate
    'tmFNE.sEnteredTime = smNowTime
    tmFNE.sEnteredDate = Format(Now, sgShowDateForm)
    tmFNE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmFNE.iUieCode = tgUIE.iCode
    tmFNE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdFollow, True
    llRow = grdFollow.FixedRows
    For ilLoop = 0 To UBound(tgCurrFNE) - 1 Step 1
        If llRow + 1 > grdFollow.Rows Then
            grdFollow.AddItem ""
        End If
        grdFollow.Row = llRow
        grdFollow.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrFNE(ilLoop).sName)
        grdFollow.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrFNE(ilLoop).sDescription)
        If tgCurrFNE(ilLoop).sState = "A" Then
            grdFollow.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdFollow.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdFollow.TextMatrix(llRow, CODEINDEX) = tgCurrFNE(ilLoop).iCode
        grdFollow.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrFNE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdFollow.Rows Then
        grdFollow.AddItem ""
    End If
    grdFollow.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrFollow-mPopulate", tgCurrFNE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlFNE As FNE
    
    gSetMousePointer grdFollow, grdFollow, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdFollow, grdFollow, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdFollow, grdFollow, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdFollow.Redraw = False
    For llRow = grdFollow.FixedRows To grdFollow.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmFNE.sName) <> "" Then
            imFNECode = tmFNE.iCode
            If tmFNE.iCode > 0 Then
                ilRet = gGetRec_FNE_FollowName(imFNECode, "Follow-mSave: Get FNE", tlFNE)
                If ilRet Then
                    If mCompare(tmFNE, tlFNE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmFNE.iVersion = tlFNE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmFNE.iCode <= 0 Then
                    ilRet = gPutInsert_FNE_FollowName(0, tmFNE, "Follow-mSave: Insert FNE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_FNE_FollowName(1, tmFNE, "Follow-mSave: Update FNE")
                    ilRet = gPutDelete_FNE_FollowName(tmFNE.iCode, "Follow-mSave: Delete FNE")
                    ilRet = gPutInsert_FNE_FollowName(1, tmFNE, "Follow-mSave: Insert FNE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_FNE_FollowName(imDeleteCodes(ilLoop), "EngrFollow- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdFollow.Redraw = True
    sgCurrFNEStamp = ""
    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrFollow-mPopulate", tgCurrFNE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrFollow
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrFollow
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdFollow, grdFollow, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdFollow, grdFollow, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdFollow, grdFollow, vbDefault
    Unload EngrFollow
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
        gSetMousePointer grdFollow, grdFollow, vbHourglass
        llTopRow = grdFollow.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdFollow, grdFollow, vbDefault
            Exit Sub
        End If
        grdFollow.Redraw = False
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
        grdFollow.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdFollow, grdFollow, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdFollow.Col
        Case NAMEINDEX
            If grdFollow.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdFollow.text = edcGrid.text
            grdFollow.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdFollow.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdFollow.text = edcGrid.text
            grdFollow.CellForeColor = vbBlack
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
    gSetFonts EngrFollow
    gCenterFormModal EngrFollow
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdFollow.FixedRows) And (lmEnableRow < grdFollow.Rows) Then
            If (lmEnableCol >= grdFollow.FixedCols) And (lmEnableCol < grdFollow.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdFollow.text = smESCValue
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
    grdFollow.Height = cmcCancel.Top - grdFollow.Top - 120    '8 * grdFollow.RowHeight(0) + 30
    gGrid_IntegralHeight grdFollow
    gGrid_FillWithRows grdFollow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrFollow = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdFollow, grdFollow, vbHourglass
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(FOLLOWLIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdFollow, grdFollow, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdFollow, grdFollow, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Follow Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Follow Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub imcInsert_Click()
    mSetShow
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    igRptIndex = FOLLOW_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub grdFollow_Click()
    If grdFollow.Col >= grdFollow.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdFollow_EnterCell()
    mSetShow
End Sub

Private Sub grdFollow_GotFocus()
    If grdFollow.Col >= grdFollow.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdFollow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdFollow.TopRow
    grdFollow.Redraw = False
End Sub

Private Sub grdFollow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdFollow.RowHeight(0) Then
        mSortCol grdFollow.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdFollow, x, y)
    If Not ilFound Then
        grdFollow.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdFollow.Col >= grdFollow.Cols - 1 Then
        grdFollow.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdFollow.TopRow
    DoEvents
    llRow = grdFollow.Row
    If grdFollow.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdFollow.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdFollow.TextMatrix(llRow, NAMEINDEX) = ""
        grdFollow.Row = llRow + 1
        grdFollow.Col = NAMEINDEX
        grdFollow.Redraw = True
    End If
    grdFollow.Redraw = True
    mEnableBox
End Sub

Private Sub grdFollow_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdFollow.Redraw = False Then
        grdFollow.Redraw = True
        If lmTopRow < grdFollow.FixedRows Then
            grdFollow.TopRow = grdFollow.FixedRows
        Else
            grdFollow.TopRow = lmTopRow
        End If
        grdFollow.Refresh
        grdFollow.Redraw = False
    End If
    If (imShowGridBox) And (grdFollow.Row >= grdFollow.FixedRows) And (grdFollow.Col >= 0) And (grdFollow.Col < grdFollow.Cols - 1) Then
        If grdFollow.RowIsVisible(grdFollow.Row) Then
            'edcGrid.Move grdFollow.Left + grdFollow.ColPos(grdFollow.Col) + 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + 30, grdFollow.ColWidth(grdFollow.Col) - 30, grdFollow.RowHeight(grdFollow.Row) - 30
            pbcArrow.Move grdFollow.Left - pbcArrow.Width - 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + (grdFollow.RowHeight(grdFollow.Row) - pbcArrow.Height) / 2
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
        If grdFollow.Col = NAMEINDEX Then
            If grdFollow.Row > grdFollow.FixedRows Then
                lmTopRow = -1
                grdFollow.Row = grdFollow.Row - 1
                If Not grdFollow.RowIsVisible(grdFollow.Row) Then
                    grdFollow.TopRow = grdFollow.TopRow - 1
                End If
                grdFollow.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdFollow.Col = grdFollow.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdFollow.TopRow = grdFollow.FixedRows
        grdFollow.Col = NAMEINDEX
        grdFollow.Row = grdFollow.FixedRows
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
        grdFollow.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdFollow.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdFollow.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdFollow.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdFollow.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdFollow.CellForeColor = vbBlack
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
        If grdFollow.Col = STATEINDEX Then
            llRow = grdFollow.Rows
            Do
                llRow = llRow - 1
            Loop While grdFollow.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdFollow.Row + 1 < llRow) Then
                lmTopRow = -1
                grdFollow.Row = grdFollow.Row + 1
                If Not grdFollow.RowIsVisible(grdFollow.Row) Then
                    imIgnoreScroll = True
                    grdFollow.TopRow = grdFollow.TopRow + 1
                End If
                grdFollow.Col = NAMEINDEX
                'grdFollow.TextMatrix(grdFollow.Row, CODEINDEX) = 0
                If Trim$(grdFollow.TextMatrix(grdFollow.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdFollow.Left - pbcArrow.Width - 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + (grdFollow.RowHeight(grdFollow.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdFollow.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdFollow.Row + 1 >= grdFollow.Rows Then
                        grdFollow.AddItem ""
                    End If
                    grdFollow.Row = grdFollow.Row + 1
                    If Not grdFollow.RowIsVisible(grdFollow.Row) Then
                        imIgnoreScroll = True
                        grdFollow.TopRow = grdFollow.TopRow + 1
                    End If
                    grdFollow.Col = NAMEINDEX
                    grdFollow.TextMatrix(grdFollow.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdFollow.Left - pbcArrow.Width - 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + (grdFollow.RowHeight(grdFollow.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdFollow.Col = grdFollow.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdFollow.TopRow = grdFollow.FixedRows
        grdFollow.Col = NAMEINDEX
        grdFollow.Row = grdFollow.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdFollow.TopRow
    llRow = grdFollow.Row
    slMsg = "Insert above " & Trim$(grdFollow.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdFollow.Redraw = False
    grdFollow.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdFollow.Row = llRow
    grdFollow.Redraw = False
    grdFollow.TopRow = llTRow
    grdFollow.Redraw = True
    DoEvents
    grdFollow.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdFollow.TopRow
    llRow = grdFollow.Row
    If (Val(grdFollow.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdFollow.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdFollow.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdFollow.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdFollow.Redraw = False
    If (Val(grdFollow.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdFollow.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdFollow.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdFollow.AddItem ""
    grdFollow.Redraw = False
    grdFollow.TopRow = llTRow
    grdFollow.Redraw = True
    DoEvents
    grdFollow.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As FNE, tlOld As FNE) As Integer
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
    mCompare = True
End Function


Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrFNE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdFollow.FixedRows To grdFollow.Rows - 1 Step 1
            slStr = Trim$(grdFollow.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdFollow.Row = llRow
                    Do While Not grdFollow.RowIsVisible(grdFollow.Row)
                        imIgnoreScroll = True
                        grdFollow.TopRow = grdFollow.TopRow + 1
                    Loop
                    grdFollow.Col = NAMEINDEX
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
    For llRow = grdFollow.FixedRows To grdFollow.Rows - 1 Step 1
        slStr = Trim$(grdFollow.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdFollow.Row = llRow
            Do While Not grdFollow.RowIsVisible(grdFollow.Row)
                imIgnoreScroll = True
                grdFollow.TopRow = grdFollow.TopRow + 1
            Loop
            grdFollow.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdFollow.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mSetFocus()
    Select Case grdFollow.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdFollow.Left + grdFollow.ColPos(grdFollow.Col) + 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + 15, grdFollow.ColWidth(grdFollow.Col) - 30, grdFollow.RowHeight(grdFollow.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdFollow.Left + grdFollow.ColPos(grdFollow.Col) + 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + 15, grdFollow.ColWidth(grdFollow.Col) - 30, grdFollow.RowHeight(grdFollow.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdFollow.Left + grdFollow.ColPos(grdFollow.Col) + 30, grdFollow.Top + grdFollow.RowPos(grdFollow.Row) + 15, grdFollow.ColWidth(grdFollow.Col) - 30, grdFollow.RowHeight(grdFollow.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
