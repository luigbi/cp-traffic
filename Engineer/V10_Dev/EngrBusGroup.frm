VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrBusGroup 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrBusGroup.frx":0000
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
      Left            =   165
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   6930
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
      Picture         =   "EngrBusGroup.frx":030A
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
      Left            =   6960
      TabIndex        =   10
      Top             =   6660
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9615
      Top             =   6555
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
      Left            =   5190
      TabIndex        =   9
      Top             =   6660
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3405
      TabIndex        =   8
      Top             =   6660
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBusGroup 
      Height          =   5805
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   10239
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
      Left            =   8235
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2505
      Picture         =   "EngrBusGroup.frx":0614
      Top             =   6570
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Bus Group"
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
      Picture         =   "EngrBusGroup.frx":091E
      Top             =   6555
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8760
      Picture         =   "EngrBusGroup.frx":11E8
      Top             =   6570
      Width           =   480
   End
End
Attribute VB_Name = "EngrBusGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrBusGroup - enters affiliate representative information
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
Private imBGECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmBGE As BGE

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
    llRow = gGrid_Search(grdBusGroup, NAMEINDEX, slStr)
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
    
    grdBusGroup.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdBusGroup.FixedRows To grdBusGroup.Rows - 1 Step 1
        slStr = Trim$(grdBusGroup.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdBusGroup.Rows - 1 Step 1
                slTestStr = Trim$(grdBusGroup.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdBusGroup.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdBusGroup.Row = llRow
                        grdBusGroup.Col = NAMEINDEX
                        grdBusGroup.CellForeColor = vbRed
                    Else
                        grdBusGroup.Row = llTestRow
                        grdBusGroup.Col = NAMEINDEX
                        grdBusGroup.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdBusGroup.Redraw = True
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
    gGrid_SortByCol grdBusGroup, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(BUSGROUPLIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdBusGroup.Row >= grdBusGroup.FixedRows) And (grdBusGroup.Row < grdBusGroup.Rows) And (grdBusGroup.Col >= 0) And (grdBusGroup.Col < grdBusGroup.Cols - 1) Then
        lmEnableRow = grdBusGroup.Row
        lmEnableCol = grdBusGroup.Col
        sgReturnCallName = grdBusGroup.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdBusGroup.Left - pbcArrow.Width - 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + (grdBusGroup.RowHeight(grdBusGroup.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdBusGroup.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdBusGroup.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdBusGroup.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdBusGroup.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdBusGroup.Left + grdBusGroup.ColPos(grdBusGroup.Col) + 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + 15, grdBusGroup.ColWidth(grdBusGroup.Col) - 30, grdBusGroup.RowHeight(grdBusGroup.Row) - 15
                edcGrid.MaxLength = Len(tmBGE.sName)
                edcGrid.text = grdBusGroup.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdBusGroup.Left + grdBusGroup.ColPos(grdBusGroup.Col) + 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + 15, grdBusGroup.ColWidth(grdBusGroup.Col) - 30, grdBusGroup.RowHeight(grdBusGroup.Row) - 15
                edcGrid.MaxLength = Len(tmBGE.sDescription)
                edcGrid.text = grdBusGroup.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdBusGroup.Left + grdBusGroup.ColPos(grdBusGroup.Col) + 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + 15, grdBusGroup.ColWidth(grdBusGroup.Col) - 30, grdBusGroup.RowHeight(grdBusGroup.Row) - 15
                smState = grdBusGroup.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdBusGroup.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdBusGroup.FixedRows) And (lmEnableRow < grdBusGroup.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdBusGroup.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdBusGroup.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdBusGroup.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdBusGroup.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdBusGroup.TextMatrix(lmEnableRow, NAMEINDEX)
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
    
    grdBusGroup.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdBusGroup.FixedRows To grdBusGroup.Rows - 1 Step 1
        slStr = Trim$(grdBusGroup.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdBusGroup.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdBusGroup.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdBusGroup.Row = llRow
                grdBusGroup.Col = NAMEINDEX
                grdBusGroup.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdBusGroup.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdBusGroup.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdBusGroup.Row = llRow
                    grdBusGroup.Col = STATEINDEX
                    grdBusGroup.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdBusGroup.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdBusGroup
    mGridColumnWidth
    'Set Titles
    grdBusGroup.TextMatrix(0, NAMEINDEX) = "Name"
    grdBusGroup.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdBusGroup.TextMatrix(0, STATEINDEX) = "State"
    grdBusGroup.Row = 1
    For ilCol = 0 To grdBusGroup.Cols - 1 Step 1
        grdBusGroup.Col = ilCol
        grdBusGroup.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdBusGroup.Height = cmcCancel.Top - grdBusGroup.Top - 120    '8 * grdBusGroup.RowHeight(0) + 30
    gGrid_IntegralHeight grdBusGroup
    gGrid_Clear grdBusGroup, True
    grdBusGroup.Row = grdBusGroup.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdBusGroup.Width = EngrBusGroup.Width - 2 * grdBusGroup.Left
    grdBusGroup.ColWidth(CODEINDEX) = 0
    grdBusGroup.ColWidth(USEDFLAGINDEX) = 0
    grdBusGroup.ColWidth(NAMEINDEX) = grdBusGroup.Width / 7
    grdBusGroup.ColWidth(STATEINDEX) = grdBusGroup.Width / 15
    grdBusGroup.ColWidth(DESCRIPTIONINDEX) = grdBusGroup.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdBusGroup.ColWidth(DESCRIPTIONINDEX) > grdBusGroup.ColWidth(ilCol) Then
                grdBusGroup.ColWidth(DESCRIPTIONINDEX) = grdBusGroup.ColWidth(DESCRIPTIONINDEX) - grdBusGroup.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdBusGroup, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdBusGroup.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdBusGroup.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmBGE.iCode = Val(grdBusGroup.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdBusGroup.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmBGE.sName = ""
    Else
        tmBGE.sName = slStr
    End If
    tmBGE.sDescription = grdBusGroup.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdBusGroup.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmBGE.sState = "D"
    Else
        tmBGE.sState = "A"
    End If
    If tmBGE.iCode <= 0 Then
        tmBGE.sUsedFlag = "N"
    Else
        tmBGE.sUsedFlag = grdBusGroup.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmBGE.iVersion = 0
    tmBGE.iOrigBgeCode = tmBGE.iCode
    tmBGE.sCurrent = "Y"
    'tmBGE.sEnteredDate = smNowDate
    'tmBGE.sEnteredTime = smNowTime
    tmBGE.sEnteredDate = Format(Now, sgShowDateForm)
    tmBGE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmBGE.iUieCode = tgUIE.iCode
    tmBGE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdBusGroup, True
    llRow = grdBusGroup.FixedRows
    For ilLoop = 0 To UBound(tgCurrBGE) - 1 Step 1
        If llRow + 1 > grdBusGroup.Rows Then
            grdBusGroup.AddItem ""
        End If
        grdBusGroup.Row = llRow
        grdBusGroup.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrBGE(ilLoop).sName)
        grdBusGroup.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrBGE(ilLoop).sDescription)
        If tgCurrBGE(ilLoop).sState = "A" Then
            grdBusGroup.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdBusGroup.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdBusGroup.TextMatrix(llRow, CODEINDEX) = tgCurrBGE(ilLoop).iCode
        grdBusGroup.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrBGE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdBusGroup.Rows Then
        grdBusGroup.AddItem ""
    End If
    grdBusGroup.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_BGE_BusGroup("C", sgCurrBGEStamp, "EngrBusGroup-mPopulate", tgCurrBGE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlBGE As BGE
    
    gSetMousePointer grdBusGroup, grdBusGroup, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdBusGroup, grdBusGroup, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdBusGroup, grdBusGroup, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdBusGroup.Redraw = False
    For llRow = grdBusGroup.FixedRows To grdBusGroup.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmBGE.sName) <> "" Then
            imBGECode = tmBGE.iCode
            If tmBGE.iCode > 0 Then
                ilRet = gGetRec_BGE_BusGroup(imBGECode, "Bus Group-mSave: Get BGE", tlBGE)
                If ilRet Then
                    If mCompare(tmBGE, tlBGE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmBGE.iVersion = tlBGE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmBGE.iCode <= 0 Then
                    ilRet = gPutInsert_BGE_BusGroup(0, tmBGE, "Bus Group-mSave: Insert BGE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_BGE_BusGroup(1, tmBGE, "Bus Group-mSave: Update BGE")
                    ilRet = gPutDelete_BGE_BusGroup(tmBGE.iCode, "Bus Group-mSave: Delete BGE")
                    ilRet = gPutInsert_BGE_BusGroup(1, tmBGE, "Bus Group-mSave: Insert BGE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_BGE_BusGroup(imDeleteCodes(ilLoop), "EngrBusGroup- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdBusGroup.Redraw = True
    sgCurrBGEStamp = ""
    ilRet = gGetTypeOfRecs_BGE_BusGroup("C", sgCurrBGEStamp, "EngrBusGroup-mPopulate", tgCurrBGE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrBusGroup
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrBusGroup
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdBusGroup, grdBusGroup, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdBusGroup, grdBusGroup, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdBusGroup, grdBusGroup, vbDefault
    Unload EngrBusGroup
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
        gSetMousePointer grdBusGroup, grdBusGroup, vbHourglass
        llTopRow = grdBusGroup.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdBusGroup, grdBusGroup, vbDefault
            Exit Sub
        End If
        grdBusGroup.Redraw = False
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
        grdBusGroup.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdBusGroup, grdBusGroup, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdBusGroup.Col
        Case NAMEINDEX
            If grdBusGroup.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdBusGroup.text = edcGrid.text
            grdBusGroup.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdBusGroup.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdBusGroup.text = edcGrid.text
            grdBusGroup.CellForeColor = vbBlack
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
    gSetFonts EngrBusGroup
    gCenterFormModal EngrBusGroup
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdBusGroup.FixedRows) And (lmEnableRow < grdBusGroup.Rows) Then
            If (lmEnableCol >= grdBusGroup.FixedCols) And (lmEnableCol < grdBusGroup.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdBusGroup.text = smESCValue
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
    grdBusGroup.Height = cmcCancel.Top - grdBusGroup.Top - 120    '8 * grdBusGroup.RowHeight(0) + 30
    gGrid_IntegralHeight grdBusGroup
    gGrid_FillWithRows grdBusGroup
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrBusGroup = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdBusGroup, grdBusGroup, vbHourglass
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSGROUPLIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdBusGroup, grdBusGroup, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdBusGroup, grdBusGroup, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors   'rdoErrors
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
    igRptIndex = BUSGROUP_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub grdBusGroup_Click()
    If grdBusGroup.Col >= grdBusGroup.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdBusGroup_EnterCell()
    mSetShow
End Sub

Private Sub grdBusGroup_GotFocus()
    If grdBusGroup.Col >= grdBusGroup.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdBusGroup_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdBusGroup.TopRow
    grdBusGroup.Redraw = False
End Sub

Private Sub grdBusGroup_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdBusGroup.RowHeight(0) Then
        mSortCol grdBusGroup.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdBusGroup, x, y)
    If Not ilFound Then
        grdBusGroup.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdBusGroup.Col >= grdBusGroup.Cols - 1 Then
        grdBusGroup.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdBusGroup.TopRow
    DoEvents
    llRow = grdBusGroup.Row
    If grdBusGroup.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdBusGroup.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdBusGroup.TextMatrix(llRow, NAMEINDEX) = ""
        grdBusGroup.Row = llRow + 1
        grdBusGroup.Col = NAMEINDEX
        grdBusGroup.Redraw = True
    End If
    grdBusGroup.Redraw = True
    mEnableBox
End Sub

Private Sub grdBusGroup_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdBusGroup.Redraw = False Then
        grdBusGroup.Redraw = True
        If lmTopRow < grdBusGroup.FixedRows Then
            grdBusGroup.TopRow = grdBusGroup.FixedRows
        Else
            grdBusGroup.TopRow = lmTopRow
        End If
        grdBusGroup.Refresh
        grdBusGroup.Redraw = False
    End If
    If (imShowGridBox) And (grdBusGroup.Row >= grdBusGroup.FixedRows) And (grdBusGroup.Col >= 0) And (grdBusGroup.Col < grdBusGroup.Cols - 1) Then
        If grdBusGroup.RowIsVisible(grdBusGroup.Row) Then
            'edcGrid.Move grdBusGroup.Left + grdBusGroup.ColPos(grdBusGroup.Col) + 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + 30, grdBusGroup.ColWidth(grdBusGroup.Col) - 30, grdBusGroup.RowHeight(grdBusGroup.Row) - 30
            pbcArrow.Move grdBusGroup.Left - pbcArrow.Width - 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + (grdBusGroup.RowHeight(grdBusGroup.Row) - pbcArrow.Height) / 2
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
        If grdBusGroup.Col = NAMEINDEX Then
            If grdBusGroup.Row > grdBusGroup.FixedRows Then
                lmTopRow = -1
                grdBusGroup.Row = grdBusGroup.Row - 1
                If Not grdBusGroup.RowIsVisible(grdBusGroup.Row) Then
                    grdBusGroup.TopRow = grdBusGroup.TopRow - 1
                End If
                grdBusGroup.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdBusGroup.Col = grdBusGroup.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdBusGroup.TopRow = grdBusGroup.FixedRows
        grdBusGroup.Col = NAMEINDEX
        grdBusGroup.Row = grdBusGroup.FixedRows
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
        grdBusGroup.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdBusGroup.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdBusGroup.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdBusGroup.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdBusGroup.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdBusGroup.CellForeColor = vbBlack
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
        If grdBusGroup.Col = STATEINDEX Then
            llRow = grdBusGroup.Rows
            Do
                llRow = llRow - 1
            Loop While grdBusGroup.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdBusGroup.Row + 1 < llRow) Then
                lmTopRow = -1
                grdBusGroup.Row = grdBusGroup.Row + 1
                If Not grdBusGroup.RowIsVisible(grdBusGroup.Row) Then
                    imIgnoreScroll = True
                    grdBusGroup.TopRow = grdBusGroup.TopRow + 1
                End If
                grdBusGroup.Col = NAMEINDEX
                'grdBusGroup.TextMatrix(grdBusGroup.Row, CODEINDEX) = 0
                If Trim$(grdBusGroup.TextMatrix(grdBusGroup.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdBusGroup.Left - pbcArrow.Width - 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + (grdBusGroup.RowHeight(grdBusGroup.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdBusGroup.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdBusGroup.Row + 1 >= grdBusGroup.Rows Then
                        grdBusGroup.AddItem ""
                    End If
                    grdBusGroup.Row = grdBusGroup.Row + 1
                    If Not grdBusGroup.RowIsVisible(grdBusGroup.Row) Then
                        imIgnoreScroll = True
                        grdBusGroup.TopRow = grdBusGroup.TopRow + 1
                    End If
                    grdBusGroup.Col = NAMEINDEX
                    grdBusGroup.TextMatrix(grdBusGroup.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdBusGroup.Left - pbcArrow.Width - 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + (grdBusGroup.RowHeight(grdBusGroup.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdBusGroup.Col = grdBusGroup.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdBusGroup.TopRow = grdBusGroup.FixedRows
        grdBusGroup.Col = NAMEINDEX
        grdBusGroup.Row = grdBusGroup.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdBusGroup.TopRow
    llRow = grdBusGroup.Row
    slMsg = "Insert above " & Trim$(grdBusGroup.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdBusGroup.Redraw = False
    grdBusGroup.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdBusGroup.Row = llRow
    grdBusGroup.Redraw = False
    grdBusGroup.TopRow = llTRow
    grdBusGroup.Redraw = True
    DoEvents
    grdBusGroup.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdBusGroup.TopRow
    llRow = grdBusGroup.Row
    If (Val(grdBusGroup.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdBusGroup.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdBusGroup.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdBusGroup.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdBusGroup.Redraw = False
    If (Val(grdBusGroup.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdBusGroup.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdBusGroup.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdBusGroup.AddItem ""
    grdBusGroup.Redraw = False
    grdBusGroup.TopRow = llTRow
    grdBusGroup.Redraw = True
    DoEvents
    grdBusGroup.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As BGE, tlOld As BGE) As Integer
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
        If UBound(tgCurrBGE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdBusGroup.FixedRows To grdBusGroup.Rows - 1 Step 1
            slStr = Trim$(grdBusGroup.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdBusGroup.Row = llRow
                    Do While Not grdBusGroup.RowIsVisible(grdBusGroup.Row)
                        imIgnoreScroll = True
                        grdBusGroup.TopRow = grdBusGroup.TopRow + 1
                    Loop
                    grdBusGroup.Col = NAMEINDEX
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
    For llRow = grdBusGroup.FixedRows To grdBusGroup.Rows - 1 Step 1
        slStr = Trim$(grdBusGroup.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdBusGroup.Row = llRow
            Do While Not grdBusGroup.RowIsVisible(grdBusGroup.Row)
                imIgnoreScroll = True
                grdBusGroup.TopRow = grdBusGroup.TopRow + 1
            Loop
            grdBusGroup.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdBusGroup.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mSetFocus()
    Select Case grdBusGroup.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdBusGroup.Left + grdBusGroup.ColPos(grdBusGroup.Col) + 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + 15, grdBusGroup.ColWidth(grdBusGroup.Col) - 30, grdBusGroup.RowHeight(grdBusGroup.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdBusGroup.Left + grdBusGroup.ColPos(grdBusGroup.Col) + 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + 15, grdBusGroup.ColWidth(grdBusGroup.Col) - 30, grdBusGroup.RowHeight(grdBusGroup.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdBusGroup.Left + grdBusGroup.ColPos(grdBusGroup.Col) + 30, grdBusGroup.Top + grdBusGroup.RowPos(grdBusGroup.Row) + 15, grdBusGroup.ColWidth(grdBusGroup.Col) - 30, grdBusGroup.RowHeight(grdBusGroup.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
