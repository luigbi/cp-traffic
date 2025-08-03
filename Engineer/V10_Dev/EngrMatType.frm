VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrMatType 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrMatType.frx":0000
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
      Left            =   11685
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   45
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
      Left            =   135
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   6555
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
      Picture         =   "EngrMatType.frx":030A
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
      Left            =   7005
      TabIndex        =   10
      Top             =   6750
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9660
      Top             =   6645
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
      Left            =   5235
      TabIndex        =   9
      Top             =   6750
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3450
      TabIndex        =   8
      Top             =   6750
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMatType 
      Height          =   5760
      Left            =   405
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   10160
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
      Left            =   9930
      TabIndex        =   12
      Top             =   90
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
      Left            =   8235
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2550
      Picture         =   "EngrMatType.frx":0614
      Top             =   6660
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Material Type"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1665
      Picture         =   "EngrMatType.frx":091E
      Top             =   6660
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8805
      Picture         =   "EngrMatType.frx":11E8
      Top             =   6660
      Width           =   480
   End
End
Attribute VB_Name = "EngrMatType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrMatType - enters affiliate representative information
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
Private imMTECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmMTE As MTE

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
    llRow = gGrid_Search(grdMatType, NAMEINDEX, slStr)
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
    
    grdMatType.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdMatType.FixedRows To grdMatType.Rows - 1 Step 1
        slStr = Trim$(grdMatType.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdMatType.Rows - 1 Step 1
                slTestStr = Trim$(grdMatType.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdMatType.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdMatType.Row = llRow
                        grdMatType.Col = NAMEINDEX
                        grdMatType.CellForeColor = vbRed
                    Else
                        grdMatType.Row = llTestRow
                        grdMatType.Col = NAMEINDEX
                        grdMatType.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdMatType.Redraw = True
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
    gGrid_SortByCol grdMatType, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(MATERIALTYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdMatType.Row >= grdMatType.FixedRows) And (grdMatType.Row < grdMatType.Rows) And (grdMatType.Col >= 0) And (grdMatType.Col < grdMatType.Cols - 1) Then
        lmEnableRow = grdMatType.Row
        lmEnableCol = grdMatType.Col
        sgReturnCallName = grdMatType.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdMatType.Left - pbcArrow.Width - 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + (grdMatType.RowHeight(grdMatType.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdMatType.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdMatType.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdMatType.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdMatType.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdMatType.Left + grdMatType.ColPos(grdMatType.Col) + 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + 15, grdMatType.ColWidth(grdMatType.Col) - 30, grdMatType.RowHeight(grdMatType.Row) - 15
                'edcGrid.MaxLength = Len(tmMTE.sName)
                edcGrid.MaxLength = gGetAllowedChars("MATERIAL", Len(tmMTE.sName))
                edcGrid.text = grdMatType.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdMatType.Left + grdMatType.ColPos(grdMatType.Col) + 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + 15, grdMatType.ColWidth(grdMatType.Col) - 30, grdMatType.RowHeight(grdMatType.Row) - 15
                edcGrid.MaxLength = Len(tmMTE.sDescription)
                edcGrid.text = grdMatType.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdMatType.Left + grdMatType.ColPos(grdMatType.Col) + 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + 15, grdMatType.ColWidth(grdMatType.Col) - 30, grdMatType.RowHeight(grdMatType.Row) - 15
                smState = grdMatType.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdMatType.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdMatType.FixedRows) And (lmEnableRow < grdMatType.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdMatType.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdMatType.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdMatType.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdMatType.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdMatType.TextMatrix(lmEnableRow, NAMEINDEX)
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
    
    grdMatType.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdMatType.FixedRows To grdMatType.Rows - 1 Step 1
        slStr = Trim$(grdMatType.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdMatType.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdMatType.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdMatType.Row = llRow
                grdMatType.Col = NAMEINDEX
                grdMatType.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdMatType.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdMatType.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdMatType.Row = llRow
                    grdMatType.Col = STATEINDEX
                    grdMatType.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdMatType.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdMatType
    mGridColumnWidth
    'Set Titles
    grdMatType.TextMatrix(0, NAMEINDEX) = "Name"
    grdMatType.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdMatType.TextMatrix(0, STATEINDEX) = "State"
    grdMatType.Row = 1
    For ilCol = 0 To grdMatType.Cols - 1 Step 1
        grdMatType.Col = ilCol
        grdMatType.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdMatType.Height = cmcCancel.Top - grdMatType.Top - 120    '8 * grdMatType.RowHeight(0) + 30
    gGrid_IntegralHeight grdMatType
    gGrid_Clear grdMatType, True
    grdMatType.Row = grdMatType.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdMatType.Width = EngrMatType.Width - 2 * grdMatType.Left
    grdMatType.ColWidth(CODEINDEX) = 0
    grdMatType.ColWidth(USEDFLAGINDEX) = 0
    grdMatType.ColWidth(NAMEINDEX) = grdMatType.Width / 9
    grdMatType.ColWidth(STATEINDEX) = grdMatType.Width / 15
    grdMatType.ColWidth(DESCRIPTIONINDEX) = grdMatType.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdMatType.ColWidth(DESCRIPTIONINDEX) > grdMatType.ColWidth(ilCol) Then
                grdMatType.ColWidth(DESCRIPTIONINDEX) = grdMatType.ColWidth(DESCRIPTIONINDEX) - grdMatType.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdMatType, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdMatType.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdMatType.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmMTE.iCode = Val(grdMatType.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdMatType.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmMTE.sName = ""
    Else
        tmMTE.sName = slStr
    End If
    tmMTE.sDescription = grdMatType.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdMatType.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmMTE.sState = "D"
    Else
        tmMTE.sState = "A"
    End If
    If tmMTE.iCode <= 0 Then
        tmMTE.sUsedFlag = "N"
    Else
        tmMTE.sUsedFlag = grdMatType.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmMTE.iVersion = 0
    tmMTE.iOrigMteCode = tmMTE.iCode
    tmMTE.sCurrent = "Y"
    'tmMTE.sEnteredDate = smNowDate
    'tmMTE.sEnteredTime = smNowTime
    tmMTE.sEnteredDate = Format(Now, sgShowDateForm)
    tmMTE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmMTE.iUieCode = tgUIE.iCode
    tmMTE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdMatType, True
    llRow = grdMatType.FixedRows
    For ilLoop = 0 To UBound(tgCurrMTE) - 1 Step 1
        If llRow + 1 > grdMatType.Rows Then
            grdMatType.AddItem ""
        End If
        grdMatType.Row = llRow
        grdMatType.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrMTE(ilLoop).sName)
        grdMatType.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrMTE(ilLoop).sDescription)
        If tgCurrMTE(ilLoop).sState = "A" Then
            grdMatType.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdMatType.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdMatType.TextMatrix(llRow, CODEINDEX) = tgCurrMTE(ilLoop).iCode
        grdMatType.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrMTE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdMatType.Rows Then
        grdMatType.AddItem ""
    End If
    grdMatType.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrMatType-mPopulate", tgCurrMTE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlMTE As MTE
    
    gSetMousePointer grdMatType, grdMatType, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdMatType, grdMatType, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdMatType, grdMatType, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdMatType.Redraw = False
    For llRow = grdMatType.FixedRows To grdMatType.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmMTE.sName) <> "" Then
            imMTECode = tmMTE.iCode
            If tmMTE.iCode > 0 Then
                ilRet = gGetRec_MTE_MaterialType(imMTECode, "Material Type-mSave: Get MTE", tlMTE)
                If ilRet Then
                    If mCompare(tmMTE, tlMTE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmMTE.iVersion = tlMTE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmMTE.iCode <= 0 Then
                    ilRet = gPutInsert_MTE_MaterialType(0, tmMTE, "Material Type-mSave: Insert MTE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_MTE_MaterialType(1, tmMTE, "Material Type-mSave: Update MTE")
                    ilRet = gPutDelete_MTE_MaterialType(tmMTE.iCode, "Material Type-mSave: Delete MTE")
                    ilRet = gPutInsert_MTE_MaterialType(1, tmMTE, "Material Type-mSave: Insert MTE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_MTE_MaterialType(imDeleteCodes(ilLoop), "EngrMatType- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdMatType.Redraw = True
    sgCurrMTEStamp = ""
    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrMatType-mPopulate", tgCurrMTE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrMatType
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrMatType
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdMatType, grdMatType, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdMatType, grdMatType, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdMatType, grdMatType, vbDefault
    Unload EngrMatType
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
        gSetMousePointer grdMatType, grdMatType, vbHourglass
        llTopRow = grdMatType.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdMatType, grdMatType, vbDefault
            Exit Sub
        End If
        grdMatType.Redraw = False
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
        grdMatType.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdMatType, grdMatType, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdMatType.Col
        Case NAMEINDEX
            If grdMatType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdMatType.text = edcGrid.text
            grdMatType.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdMatType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdMatType.text = edcGrid.text
            grdMatType.CellForeColor = vbBlack
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
    gSetFonts EngrMatType
    gCenterFormModal EngrMatType
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdMatType.FixedRows) And (lmEnableRow < grdMatType.Rows) Then
            If (lmEnableCol >= grdMatType.FixedCols) And (lmEnableCol < grdMatType.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdMatType.text = smESCValue
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
    grdMatType.Height = cmcCancel.Top - grdMatType.Top - 120    '8 * grdMatType.RowHeight(0) + 30
    gGrid_IntegralHeight grdMatType
    gGrid_FillWithRows grdMatType
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrMatType = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdMatType, grdMatType, vbHourglass
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(MATERIALTYPELIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdMatType, grdMatType, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdMatType, grdMatType, vbDefault
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
    igRptIndex = MATTYPE_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub grdMatType_Click()
    If grdMatType.Col >= grdMatType.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdMatType_EnterCell()
    mSetShow
End Sub

Private Sub grdMatType_GotFocus()
    If grdMatType.Col >= grdMatType.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdMatType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdMatType.TopRow
    grdMatType.Redraw = False
End Sub

Private Sub grdMatType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdMatType.RowHeight(0) Then
        mSortCol grdMatType.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdMatType, x, y)
    If Not ilFound Then
        grdMatType.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdMatType.Col >= grdMatType.Cols - 1 Then
        grdMatType.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdMatType.TopRow
    DoEvents
    llRow = grdMatType.Row
    If grdMatType.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdMatType.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdMatType.TextMatrix(llRow, NAMEINDEX) = ""
        grdMatType.Row = llRow + 1
        grdMatType.Col = NAMEINDEX
        grdMatType.Redraw = True
    End If
    grdMatType.Redraw = True
    mEnableBox
End Sub

Private Sub grdMatType_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdMatType.Redraw = False Then
        grdMatType.Redraw = True
        If lmTopRow < grdMatType.FixedRows Then
            grdMatType.TopRow = grdMatType.FixedRows
        Else
            grdMatType.TopRow = lmTopRow
        End If
        grdMatType.Refresh
        grdMatType.Redraw = False
    End If
    If (imShowGridBox) And (grdMatType.Row >= grdMatType.FixedRows) And (grdMatType.Col >= 0) And (grdMatType.Col < grdMatType.Cols - 1) Then
        If grdMatType.RowIsVisible(grdMatType.Row) Then
            'edcGrid.Move grdMatType.Left + grdMatType.ColPos(grdMatType.Col) + 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + 30, grdMatType.ColWidth(grdMatType.Col) - 30, grdMatType.RowHeight(grdMatType.Row) - 30
            pbcArrow.Move grdMatType.Left - pbcArrow.Width - 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + (grdMatType.RowHeight(grdMatType.Row) - pbcArrow.Height) / 2
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
        If grdMatType.Col = NAMEINDEX Then
            If grdMatType.Row > grdMatType.FixedRows Then
                lmTopRow = -1
                grdMatType.Row = grdMatType.Row - 1
                If Not grdMatType.RowIsVisible(grdMatType.Row) Then
                    grdMatType.TopRow = grdMatType.TopRow - 1
                End If
                grdMatType.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdMatType.Col = grdMatType.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdMatType.TopRow = grdMatType.FixedRows
        grdMatType.Col = NAMEINDEX
        grdMatType.Row = grdMatType.FixedRows
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
        grdMatType.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdMatType.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdMatType.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdMatType.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdMatType.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdMatType.CellForeColor = vbBlack
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
        If grdMatType.Col = STATEINDEX Then
            llRow = grdMatType.Rows
            Do
                llRow = llRow - 1
            Loop While grdMatType.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdMatType.Row + 1 < llRow) Then
                lmTopRow = -1
                grdMatType.Row = grdMatType.Row + 1
                If Not grdMatType.RowIsVisible(grdMatType.Row) Then
                    imIgnoreScroll = True
                    grdMatType.TopRow = grdMatType.TopRow + 1
                End If
                grdMatType.Col = NAMEINDEX
                'grdMatType.TextMatrix(grdMatType.Row, CODEINDEX) = 0
                If Trim$(grdMatType.TextMatrix(grdMatType.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdMatType.Left - pbcArrow.Width - 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + (grdMatType.RowHeight(grdMatType.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdMatType.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdMatType.Row + 1 >= grdMatType.Rows Then
                        grdMatType.AddItem ""
                    End If
                    grdMatType.Row = grdMatType.Row + 1
                    If Not grdMatType.RowIsVisible(grdMatType.Row) Then
                        imIgnoreScroll = True
                        grdMatType.TopRow = grdMatType.TopRow + 1
                    End If
                    grdMatType.Col = NAMEINDEX
                    grdMatType.TextMatrix(grdMatType.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdMatType.Left - pbcArrow.Width - 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + (grdMatType.RowHeight(grdMatType.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdMatType.Col = grdMatType.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdMatType.TopRow = grdMatType.FixedRows
        grdMatType.Col = NAMEINDEX
        grdMatType.Row = grdMatType.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdMatType.TopRow
    llRow = grdMatType.Row
    slMsg = "Insert above " & Trim$(grdMatType.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdMatType.Redraw = False
    grdMatType.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdMatType.Row = llRow
    grdMatType.Redraw = False
    grdMatType.TopRow = llTRow
    grdMatType.Redraw = True
    DoEvents
    grdMatType.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdMatType.TopRow
    llRow = grdMatType.Row
    If (Val(grdMatType.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdMatType.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdMatType.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdMatType.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdMatType.Redraw = False
    If (Val(grdMatType.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdMatType.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdMatType.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdMatType.AddItem ""
    grdMatType.Redraw = False
    grdMatType.TopRow = llTRow
    grdMatType.Redraw = True
    DoEvents
    grdMatType.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As MTE, tlOld As MTE) As Integer
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
        If UBound(tgCurrMTE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdMatType.FixedRows To grdMatType.Rows - 1 Step 1
            slStr = Trim$(grdMatType.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdMatType.Row = llRow
                    Do While Not grdMatType.RowIsVisible(grdMatType.Row)
                        imIgnoreScroll = True
                        grdMatType.TopRow = grdMatType.TopRow + 1
                    Loop
                    grdMatType.Col = NAMEINDEX
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
    For llRow = grdMatType.FixedRows To grdMatType.Rows - 1 Step 1
        slStr = Trim$(grdMatType.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdMatType.Row = llRow
            Do While Not grdMatType.RowIsVisible(grdMatType.Row)
                imIgnoreScroll = True
                grdMatType.TopRow = grdMatType.TopRow + 1
            Loop
            grdMatType.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdMatType.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mSetFocus()
    Select Case grdMatType.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdMatType.Left + grdMatType.ColPos(grdMatType.Col) + 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + 15, grdMatType.ColWidth(grdMatType.Col) - 30, grdMatType.RowHeight(grdMatType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdMatType.Left + grdMatType.ColPos(grdMatType.Col) + 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + 15, grdMatType.ColWidth(grdMatType.Col) - 30, grdMatType.RowHeight(grdMatType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdMatType.Left + grdMatType.ColPos(grdMatType.Col) + 30, grdMatType.Top + grdMatType.RowPos(grdMatType.Row) + 15, grdMatType.ColWidth(grdMatType.Col) - 30, grdMatType.RowHeight(grdMatType.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
