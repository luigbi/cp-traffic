VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrSilence 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrSilence.frx":0000
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
      Left            =   105
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
      Picture         =   "EngrSilence.frx":030A
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
      TabIndex        =   10
      Top             =   6690
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9645
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
      FormDesignWidth =   11790
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   6690
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3435
      TabIndex        =   8
      Top             =   6690
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSilence 
      Height          =   5880
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   10372
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
      Left            =   10170
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
      Left            =   8475
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2535
      Picture         =   "EngrSilence.frx":0614
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Silence"
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
      Picture         =   "EngrSilence.frx":091E
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8790
      Picture         =   "EngrSilence.frx":11E8
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "EngrSilence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrSilence - enters affiliate representative information
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
Private imSCECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmSCE As SCE

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
    llRow = gGrid_Search(grdSilence, NAMEINDEX, slStr)
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
    
    grdSilence.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdSilence.FixedRows To grdSilence.Rows - 1 Step 1
        slStr = Trim$(grdSilence.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdSilence.Rows - 1 Step 1
                slTestStr = Trim$(grdSilence.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdSilence.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdSilence.Row = llRow
                        grdSilence.Col = NAMEINDEX
                        grdSilence.CellForeColor = vbRed
                    Else
                        grdSilence.Row = llTestRow
                        grdSilence.Col = NAMEINDEX
                        grdSilence.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdSilence.Redraw = True
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
    gGrid_SortByCol grdSilence, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(SILENCELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdSilence.Row >= grdSilence.FixedRows) And (grdSilence.Row < grdSilence.Rows) And (grdSilence.Col >= 0) And (grdSilence.Col < grdSilence.Cols - 1) Then
        lmEnableRow = grdSilence.Row
        lmEnableCol = grdSilence.Col
        sgReturnCallName = grdSilence.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdSilence.Left - pbcArrow.Width - 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + (grdSilence.RowHeight(grdSilence.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdSilence.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdSilence.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdSilence.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdSilence.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdSilence.Left + grdSilence.ColPos(grdSilence.Col) + 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + 15, grdSilence.ColWidth(grdSilence.Col) - 30, grdSilence.RowHeight(grdSilence.Row) - 15
                edcGrid.MaxLength = Len(tmSCE.sAutoChar)
                edcGrid.text = grdSilence.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdSilence.Left + grdSilence.ColPos(grdSilence.Col) + 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + 15, grdSilence.ColWidth(grdSilence.Col) - 30, grdSilence.RowHeight(grdSilence.Row) - 15
                edcGrid.MaxLength = Len(tmSCE.sDescription)
                edcGrid.text = grdSilence.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdSilence.Left + grdSilence.ColPos(grdSilence.Col) + 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + 15, grdSilence.ColWidth(grdSilence.Col) - 30, grdSilence.RowHeight(grdSilence.Row) - 15
                smState = grdSilence.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdSilence.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdSilence.FixedRows) And (lmEnableRow < grdSilence.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdSilence.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdSilence.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdSilence.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdSilence.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdSilence.TextMatrix(lmEnableRow, NAMEINDEX)
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
    
    grdSilence.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdSilence.FixedRows To grdSilence.Rows - 1 Step 1
        slStr = Trim$(grdSilence.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdSilence.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdSilence.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdSilence.Row = llRow
                grdSilence.Col = NAMEINDEX
                grdSilence.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdSilence.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdSilence.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdSilence.Row = llRow
                    grdSilence.Col = STATEINDEX
                    grdSilence.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdSilence.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdSilence
    mGridColumnWidth
    'Set Titles
    grdSilence.TextMatrix(0, NAMEINDEX) = "Name"
    grdSilence.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdSilence.TextMatrix(0, STATEINDEX) = "State"
    grdSilence.Row = 1
    For ilCol = 0 To grdSilence.Cols - 1 Step 1
        grdSilence.Col = ilCol
        grdSilence.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdSilence.Height = cmcCancel.Top - grdSilence.Top - 120    '8 * grdSilence.RowHeight(0) + 30
    gGrid_IntegralHeight grdSilence
    gGrid_Clear grdSilence, True
    grdSilence.Row = grdSilence.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdSilence.Width = EngrSilence.Width - 2 * grdSilence.Left
    grdSilence.ColWidth(CODEINDEX) = 0
    grdSilence.ColWidth(USEDFLAGINDEX) = 0
    grdSilence.ColWidth(NAMEINDEX) = grdSilence.Width / 9
    grdSilence.ColWidth(STATEINDEX) = grdSilence.Width / 15
    grdSilence.ColWidth(DESCRIPTIONINDEX) = grdSilence.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdSilence.ColWidth(DESCRIPTIONINDEX) > grdSilence.ColWidth(ilCol) Then
                grdSilence.ColWidth(DESCRIPTIONINDEX) = grdSilence.ColWidth(DESCRIPTIONINDEX) - grdSilence.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdSilence, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdSilence.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdSilence.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmSCE.iCode = Val(grdSilence.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdSilence.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmSCE.sAutoChar = ""
    Else
        tmSCE.sAutoChar = slStr
    End If
    tmSCE.sDescription = grdSilence.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdSilence.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmSCE.sState = "D"
    Else
        tmSCE.sState = "A"
    End If
    If tmSCE.iCode <= 0 Then
        tmSCE.sUsedFlag = "N"
    Else
        tmSCE.sUsedFlag = grdSilence.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmSCE.iVersion = 0
    tmSCE.iOrigSceCode = tmSCE.iCode
    tmSCE.sCurrent = "Y"
    'tmSCE.sEnteredDate = smNowDate
    'tmSCE.sEnteredTime = smNowTime
    tmSCE.sEnteredDate = Format(Now, sgShowDateForm)
    tmSCE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmSCE.iUieCode = tgUIE.iCode
    tmSCE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdSilence, True
    llRow = grdSilence.FixedRows
    For ilLoop = 0 To UBound(tgCurrSCE) - 1 Step 1
        If llRow + 1 > grdSilence.Rows Then
            grdSilence.AddItem ""
        End If
        grdSilence.Row = llRow
        grdSilence.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrSCE(ilLoop).sAutoChar)
        grdSilence.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrSCE(ilLoop).sDescription)
        If tgCurrSCE(ilLoop).sState = "A" Then
            grdSilence.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdSilence.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdSilence.TextMatrix(llRow, CODEINDEX) = tgCurrSCE(ilLoop).iCode
        grdSilence.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrSCE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdSilence.Rows Then
        grdSilence.AddItem ""
    End If
    
    grdSilence.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrSilence-mPopulate", tgCurrSCE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlSCE As SCE
    
    gSetMousePointer grdSilence, grdSilence, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdSilence, grdSilence, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdSilence, grdSilence, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdSilence.Redraw = False
    For llRow = grdSilence.FixedRows To grdSilence.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmSCE.sAutoChar) <> "" Then
            imSCECode = tmSCE.iCode
            If tmSCE.iCode > 0 Then
                ilRet = gGetRec_SCE_SilenceChar(imSCECode, "Silence-mSave: Get SCE", tlSCE)
                If ilRet Then
                    If mCompare(tmSCE, tlSCE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmSCE.iVersion = tlSCE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmSCE.iCode <= 0 Then
                    ilRet = gPutInsert_SCE_SilenceChar(0, tmSCE, "Silence-mSave: Insert SCE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_SCE_SilenceChar(1, tmSCE, "Silence-mSave: Update SCE")
                    ilRet = gPutDelete_SCE_SilenceChar(tmSCE.iCode, "Silence-mSave: Delete SCE")
                    ilRet = gPutInsert_SCE_SilenceChar(1, tmSCE, "Silence-mSave: Insert SCE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_SCE_SilenceChar(imDeleteCodes(ilLoop), "EngrSilence- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdSilence.Redraw = True
    sgCurrSCEStamp = ""
    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrSilence-mPopulate", tgCurrSCE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrSilence
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrSilence
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdSilence, grdSilence, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdSilence, grdSilence, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdSilence, grdSilence, vbDefault
    Unload EngrSilence
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
        gSetMousePointer grdSilence, grdSilence, vbHourglass
        llTopRow = grdSilence.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdSilence, grdSilence, vbDefault
            Exit Sub
        End If
        grdSilence.Redraw = False
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
        grdSilence.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdSilence, grdSilence, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdSilence.Col
        Case NAMEINDEX
            If grdSilence.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdSilence.text = edcGrid.text
            grdSilence.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdSilence.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdSilence.text = edcGrid.text
            grdSilence.CellForeColor = vbBlack
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
    gSetFonts EngrSilence
    gCenterFormModal EngrSilence
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdSilence.FixedRows) And (lmEnableRow < grdSilence.Rows) Then
            If (lmEnableCol >= grdSilence.FixedCols) And (lmEnableCol < grdSilence.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdSilence.text = smESCValue
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
    grdSilence.Height = cmcCancel.Top - grdSilence.Top - 120    '8 * grdSilence.RowHeight(0) + 30
    gGrid_IntegralHeight grdSilence
    gGrid_FillWithRows grdSilence
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrSilence = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdSilence, grdSilence, vbHourglass
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(SILENCELIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdSilence, grdSilence, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdSilence, grdSilence, vbDefault
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
    igRptIndex = SILENCE_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub grdSilence_Click()
    If grdSilence.Col >= grdSilence.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdSilence_EnterCell()
    mSetShow
End Sub

Private Sub grdSilence_GotFocus()
    If grdSilence.Col >= grdSilence.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdSilence_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdSilence.TopRow
    grdSilence.Redraw = False
End Sub

Private Sub grdSilence_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdSilence.RowHeight(0) Then
        mSortCol grdSilence.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdSilence, x, y)
    If Not ilFound Then
        grdSilence.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdSilence.Col >= grdSilence.Cols - 1 Then
        grdSilence.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdSilence.TopRow
    DoEvents
    llRow = grdSilence.Row
    If grdSilence.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdSilence.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdSilence.TextMatrix(llRow, NAMEINDEX) = ""
        grdSilence.Row = llRow + 1
        grdSilence.Col = NAMEINDEX
        grdSilence.Redraw = True
    End If
    grdSilence.Redraw = True
    mEnableBox
End Sub

Private Sub grdSilence_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdSilence.Redraw = False Then
        grdSilence.Redraw = True
        If lmTopRow < grdSilence.FixedRows Then
            grdSilence.TopRow = grdSilence.FixedRows
        Else
            grdSilence.TopRow = lmTopRow
        End If
        grdSilence.Refresh
        grdSilence.Redraw = False
    End If
    If (imShowGridBox) And (grdSilence.Row >= grdSilence.FixedRows) And (grdSilence.Col >= 0) And (grdSilence.Col < grdSilence.Cols - 1) Then
        If grdSilence.RowIsVisible(grdSilence.Row) Then
            'edcGrid.Move grdSilence.Left + grdSilence.ColPos(grdSilence.Col) + 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + 30, grdSilence.ColWidth(grdSilence.Col) - 30, grdSilence.RowHeight(grdSilence.Row) - 30
            pbcArrow.Move grdSilence.Left - pbcArrow.Width - 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + (grdSilence.RowHeight(grdSilence.Row) - pbcArrow.Height) / 2
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
        If grdSilence.Col = NAMEINDEX Then
            If grdSilence.Row > grdSilence.FixedRows Then
                lmTopRow = -1
                grdSilence.Row = grdSilence.Row - 1
                If Not grdSilence.RowIsVisible(grdSilence.Row) Then
                    grdSilence.TopRow = grdSilence.TopRow - 1
                End If
                grdSilence.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdSilence.Col = grdSilence.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdSilence.TopRow = grdSilence.FixedRows
        grdSilence.Col = NAMEINDEX
        grdSilence.Row = grdSilence.FixedRows
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
        grdSilence.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdSilence.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdSilence.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdSilence.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdSilence.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdSilence.CellForeColor = vbBlack
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
        If grdSilence.Col = STATEINDEX Then
            llRow = grdSilence.Rows
            Do
                llRow = llRow - 1
            Loop While grdSilence.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdSilence.Row + 1 < llRow) Then
                lmTopRow = -1
                grdSilence.Row = grdSilence.Row + 1
                If Not grdSilence.RowIsVisible(grdSilence.Row) Then
                    imIgnoreScroll = True
                    grdSilence.TopRow = grdSilence.TopRow + 1
                End If
                grdSilence.Col = NAMEINDEX
                'grdSilence.TextMatrix(grdSilence.Row, CODEINDEX) = 0
                If Trim$(grdSilence.TextMatrix(grdSilence.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdSilence.Left - pbcArrow.Width - 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + (grdSilence.RowHeight(grdSilence.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdSilence.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdSilence.Row + 1 >= grdSilence.Rows Then
                        grdSilence.AddItem ""
                    End If
                    grdSilence.Row = grdSilence.Row + 1
                    If Not grdSilence.RowIsVisible(grdSilence.Row) Then
                        imIgnoreScroll = True
                        grdSilence.TopRow = grdSilence.TopRow + 1
                    End If
                    grdSilence.Col = NAMEINDEX
                    grdSilence.TextMatrix(grdSilence.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdSilence.Left - pbcArrow.Width - 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + (grdSilence.RowHeight(grdSilence.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdSilence.Col = grdSilence.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdSilence.TopRow = grdSilence.FixedRows
        grdSilence.Col = NAMEINDEX
        grdSilence.Row = grdSilence.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdSilence.TopRow
    llRow = grdSilence.Row
    slMsg = "Insert above " & Trim$(grdSilence.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdSilence.Redraw = False
    grdSilence.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdSilence.Row = llRow
    grdSilence.Redraw = False
    grdSilence.TopRow = llTRow
    grdSilence.Redraw = True
    DoEvents
    grdSilence.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdSilence.TopRow
    llRow = grdSilence.Row
    If (Val(grdSilence.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdSilence.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdSilence.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdSilence.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdSilence.Redraw = False
    If (Val(grdSilence.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdSilence.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdSilence.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdSilence.AddItem ""
    grdSilence.Redraw = False
    grdSilence.TopRow = llTRow
    grdSilence.Redraw = True
    DoEvents
    grdSilence.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As SCE, tlOld As SCE) As Integer
    If StrComp(tlNew.sAutoChar, tlOld.sAutoChar, vbTextCompare) <> 0 Then
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
        If UBound(tgCurrSCE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdSilence.FixedRows To grdSilence.Rows - 1 Step 1
            slStr = Trim$(grdSilence.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdSilence.Row = llRow
                    Do While Not grdSilence.RowIsVisible(grdSilence.Row)
                        imIgnoreScroll = True
                        grdSilence.TopRow = grdSilence.TopRow + 1
                    Loop
                    grdSilence.Col = NAMEINDEX
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
    For llRow = grdSilence.FixedRows To grdSilence.Rows - 1 Step 1
        slStr = Trim$(grdSilence.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdSilence.Row = llRow
            Do While Not grdSilence.RowIsVisible(grdSilence.Row)
                imIgnoreScroll = True
                grdSilence.TopRow = grdSilence.TopRow + 1
            Loop
            grdSilence.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdSilence.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mSetFocus()
    Select Case grdSilence.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdSilence.Left + grdSilence.ColPos(grdSilence.Col) + 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + 15, grdSilence.ColWidth(grdSilence.Col) - 30, grdSilence.RowHeight(grdSilence.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdSilence.Left + grdSilence.ColPos(grdSilence.Col) + 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + 15, grdSilence.ColWidth(grdSilence.Col) - 30, grdSilence.RowHeight(grdSilence.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdSilence.Left + grdSilence.ColPos(grdSilence.Col) + 30, grdSilence.Top + grdSilence.RowPos(grdSilence.Row) + 15, grdSilence.ColWidth(grdSilence.Col) - 30, grdSilence.RowHeight(grdSilence.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
