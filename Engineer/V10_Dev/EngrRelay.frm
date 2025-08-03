VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrRelay 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrRelay.frx":0000
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
      Left            =   150
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   6900
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
      Picture         =   "EngrRelay.frx":030A
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
      Top             =   6705
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9630
      Top             =   6600
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
      Top             =   6705
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3420
      TabIndex        =   8
      Top             =   6705
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRelay 
      Height          =   5940
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   10478
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
      Left            =   10020
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
      Left            =   8325
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2520
      Picture         =   "EngrRelay.frx":0614
      Top             =   6615
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Relay"
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
      Picture         =   "EngrRelay.frx":091E
      Top             =   6615
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8775
      Picture         =   "EngrRelay.frx":11E8
      Top             =   6615
      Width           =   480
   End
End
Attribute VB_Name = "EngrRelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrRelay - enters affiliate representative information
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
Private imRneCode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmRNE As RNE

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
    llRow = gGrid_Search(grdRelay, NAMEINDEX, slStr)
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
    
    grdRelay.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdRelay.FixedRows To grdRelay.Rows - 1 Step 1
        slStr = Trim$(grdRelay.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdRelay.Rows - 1 Step 1
                slTestStr = Trim$(grdRelay.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdRelay.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdRelay.Row = llRow
                        grdRelay.Col = NAMEINDEX
                        grdRelay.CellForeColor = vbRed
                    Else
                        grdRelay.Row = llTestRow
                        grdRelay.Col = NAMEINDEX
                        grdRelay.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdRelay.Redraw = True
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
    gGrid_SortByCol grdRelay, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(RELAYLIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdRelay.Row >= grdRelay.FixedRows) And (grdRelay.Row < grdRelay.Rows) And (grdRelay.Col >= 0) And (grdRelay.Col < grdRelay.Cols - 1) Then
        lmEnableRow = grdRelay.Row
        lmEnableCol = grdRelay.Col
        sgReturnCallName = grdRelay.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdRelay.Left - pbcArrow.Width - 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + (grdRelay.RowHeight(grdRelay.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdRelay.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdRelay.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdRelay.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdRelay.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdRelay.Left + grdRelay.ColPos(grdRelay.Col) + 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + 15, grdRelay.ColWidth(grdRelay.Col) - 30, grdRelay.RowHeight(grdRelay.Row) - 15
                'edcGrid.MaxLength = Len(tmRNE.sName)
                edcGrid.MaxLength = gGetAllowedChars("RELAY1", Len(tmRNE.sName))
                edcGrid.text = grdRelay.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdRelay.Left + grdRelay.ColPos(grdRelay.Col) + 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + 15, grdRelay.ColWidth(grdRelay.Col) - 30, grdRelay.RowHeight(grdRelay.Row) - 15
                edcGrid.MaxLength = Len(tmRNE.sDescription)
                edcGrid.text = grdRelay.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdRelay.Left + grdRelay.ColPos(grdRelay.Col) + 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + 15, grdRelay.ColWidth(grdRelay.Col) - 30, grdRelay.RowHeight(grdRelay.Row) - 15
                smState = grdRelay.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdRelay.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdRelay.FixedRows) And (lmEnableRow < grdRelay.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdRelay.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdRelay.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdRelay.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdRelay.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdRelay.TextMatrix(lmEnableRow, NAMEINDEX)
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
    
    grdRelay.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdRelay.FixedRows To grdRelay.Rows - 1 Step 1
        slStr = Trim$(grdRelay.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdRelay.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdRelay.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdRelay.Row = llRow
                grdRelay.Col = NAMEINDEX
                grdRelay.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdRelay.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdRelay.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdRelay.Row = llRow
                    grdRelay.Col = STATEINDEX
                    grdRelay.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdRelay.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdRelay
    mGridColumnWidth
    'Set Titles
    grdRelay.TextMatrix(0, NAMEINDEX) = "Name"
    grdRelay.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdRelay.TextMatrix(0, STATEINDEX) = "State"
    grdRelay.Row = 1
    For ilCol = 0 To grdRelay.Cols - 1 Step 1
        grdRelay.Col = ilCol
        grdRelay.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdRelay.Height = cmcCancel.Top - grdRelay.Top - 120    '8 * grdRelay.RowHeight(0) + 30
    gGrid_IntegralHeight grdRelay
    gGrid_Clear grdRelay, True
    grdRelay.Row = grdRelay.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdRelay.Width = EngrRelay.Width - 2 * grdRelay.Left
    grdRelay.ColWidth(CODEINDEX) = 0
    grdRelay.ColWidth(USEDFLAGINDEX) = 0
    grdRelay.ColWidth(NAMEINDEX) = grdRelay.Width / 9
    grdRelay.ColWidth(STATEINDEX) = grdRelay.Width / 15
    grdRelay.ColWidth(DESCRIPTIONINDEX) = grdRelay.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdRelay.ColWidth(DESCRIPTIONINDEX) > grdRelay.ColWidth(ilCol) Then
                grdRelay.ColWidth(DESCRIPTIONINDEX) = grdRelay.ColWidth(DESCRIPTIONINDEX) - grdRelay.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdRelay, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdRelay.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdRelay.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmRNE.iCode = Val(grdRelay.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdRelay.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmRNE.sName = ""
    Else
        tmRNE.sName = slStr
    End If
    tmRNE.sDescription = grdRelay.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdRelay.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmRNE.sState = "D"
    Else
        tmRNE.sState = "A"
    End If
    If tmRNE.iCode <= 0 Then
        tmRNE.sUsedFlag = "N"
    Else
        tmRNE.sUsedFlag = grdRelay.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmRNE.iVersion = 0
    tmRNE.iOrigRneCode = tmRNE.iCode
    tmRNE.sCurrent = "Y"
    'tmRNE.sEnteredDate = smNowDate
    'tmRNE.sEnteredTime = smNowTime
    tmRNE.sEnteredDate = Format(Now, sgShowDateForm)
    tmRNE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmRNE.iUieCode = tgUIE.iCode
    tmRNE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdRelay, True
    llRow = grdRelay.FixedRows
    For ilLoop = 0 To UBound(tgCurrRNE) - 1 Step 1
        If llRow + 1 > grdRelay.Rows Then
            grdRelay.AddItem ""
        End If
        grdRelay.Row = llRow
        grdRelay.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrRNE(ilLoop).sName)
        grdRelay.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrRNE(ilLoop).sDescription)
        If tgCurrRNE(ilLoop).sState = "A" Then
            grdRelay.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdRelay.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdRelay.TextMatrix(llRow, CODEINDEX) = tgCurrRNE(ilLoop).iCode
        grdRelay.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrRNE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdRelay.Rows Then
        grdRelay.AddItem ""
    End If
    grdRelay.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrRelay-mPopulate", tgCurrRNE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlRNE As RNE
    
    gSetMousePointer grdRelay, grdRelay, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdRelay, grdRelay, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdRelay, grdRelay, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdRelay.Redraw = False
    For llRow = grdRelay.FixedRows To grdRelay.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmRNE.sName) <> "" Then
            imRneCode = tmRNE.iCode
            If tmRNE.iCode > 0 Then
                ilRet = gGetRec_RNE_RelayName(imRneCode, "Relay-mSave: Get RNE", tlRNE)
                If ilRet Then
                    If mCompare(tmRNE, tlRNE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmRNE.iVersion = tlRNE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmRNE.iCode <= 0 Then
                    ilRet = gPutInsert_RNE_RelayName(0, tmRNE, "Relay-mSave: Insert RNE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_RNE_RelayName(1, tmRNE, "Relay-mSave: Update RNE")
                    ilRet = gPutDelete_RNE_RelayName(tmRNE.iCode, "Relay-mSave: Delete RNE")
                    ilRet = gPutInsert_RNE_RelayName(1, tmRNE, "Relay-mSave: Insert RNE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_RNE_RelayName(imDeleteCodes(ilLoop), "EngrRelay- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdRelay.Redraw = True
    sgCurrRNEStamp = ""
    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrRelay-mPopulate", tgCurrRNE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrRelay
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrRelay
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdRelay, grdRelay, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdRelay, grdRelay, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdRelay, grdRelay, vbDefault
    Unload EngrRelay
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
        gSetMousePointer grdRelay, grdRelay, vbHourglass
        llTopRow = grdRelay.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdRelay, grdRelay, vbDefault
            Exit Sub
        End If
        grdRelay.Redraw = False
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
        grdRelay.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdRelay, grdRelay, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdRelay.Col
        Case NAMEINDEX
            If grdRelay.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdRelay.text = edcGrid.text
            grdRelay.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdRelay.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdRelay.text = edcGrid.text
            grdRelay.CellForeColor = vbBlack
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
    gSetFonts EngrRelay
    gCenterFormModal EngrRelay
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdRelay.FixedRows) And (lmEnableRow < grdRelay.Rows) Then
            If (lmEnableCol >= grdRelay.FixedCols) And (lmEnableCol < grdRelay.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdRelay.text = smESCValue
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
    grdRelay.Height = cmcCancel.Top - grdRelay.Top - 120    '8 * grdRelay.RowHeight(0) + 30
    gGrid_IntegralHeight grdRelay
    gGrid_FillWithRows grdRelay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrRelay = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdRelay, grdRelay, vbHourglass
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
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(RELAYLIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdRelay, grdRelay, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdRelay, grdRelay, vbDefault
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
    igRptIndex = RELAY_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub grdRelay_Click()
    If grdRelay.Col >= grdRelay.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdRelay_EnterCell()
    mSetShow
End Sub

Private Sub grdRelay_GotFocus()
    If grdRelay.Col >= grdRelay.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdRelay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdRelay.TopRow
    grdRelay.Redraw = False
End Sub

Private Sub grdRelay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdRelay.RowHeight(0) Then
        mSortCol grdRelay.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdRelay, x, y)
    If Not ilFound Then
        grdRelay.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdRelay.Col >= grdRelay.Cols - 1 Then
        grdRelay.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdRelay.TopRow
    DoEvents
    llRow = grdRelay.Row
    If grdRelay.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdRelay.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdRelay.TextMatrix(llRow, NAMEINDEX) = ""
        grdRelay.Row = llRow + 1
        grdRelay.Col = NAMEINDEX
        grdRelay.Redraw = True
    End If
    grdRelay.Redraw = True
    mEnableBox
End Sub

Private Sub grdRelay_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdRelay.Redraw = False Then
        grdRelay.Redraw = True
        If lmTopRow < grdRelay.FixedRows Then
            grdRelay.TopRow = grdRelay.FixedRows
        Else
            grdRelay.TopRow = lmTopRow
        End If
        grdRelay.Refresh
        grdRelay.Redraw = False
    End If
    If (imShowGridBox) And (grdRelay.Row >= grdRelay.FixedRows) And (grdRelay.Col >= 0) And (grdRelay.Col < grdRelay.Cols - 1) Then
        If grdRelay.RowIsVisible(grdRelay.Row) Then
            'edcGrid.Move grdRelay.Left + grdRelay.ColPos(grdRelay.Col) + 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + 30, grdRelay.ColWidth(grdRelay.Col) - 30, grdRelay.RowHeight(grdRelay.Row) - 30
            pbcArrow.Move grdRelay.Left - pbcArrow.Width - 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + (grdRelay.RowHeight(grdRelay.Row) - pbcArrow.Height) / 2
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
        If grdRelay.Col = NAMEINDEX Then
            If grdRelay.Row > grdRelay.FixedRows Then
                lmTopRow = -1
                grdRelay.Row = grdRelay.Row - 1
                If Not grdRelay.RowIsVisible(grdRelay.Row) Then
                    grdRelay.TopRow = grdRelay.TopRow - 1
                End If
                grdRelay.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdRelay.Col = grdRelay.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdRelay.TopRow = grdRelay.FixedRows
        grdRelay.Col = NAMEINDEX
        grdRelay.Row = grdRelay.FixedRows
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
        grdRelay.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdRelay.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdRelay.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdRelay.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdRelay.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdRelay.CellForeColor = vbBlack
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
        If grdRelay.Col = STATEINDEX Then
            llRow = grdRelay.Rows
            Do
                llRow = llRow - 1
            Loop While grdRelay.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdRelay.Row + 1 < llRow) Then
                lmTopRow = -1
                grdRelay.Row = grdRelay.Row + 1
                If Not grdRelay.RowIsVisible(grdRelay.Row) Then
                    imIgnoreScroll = True
                    grdRelay.TopRow = grdRelay.TopRow + 1
                End If
                grdRelay.Col = NAMEINDEX
                'grdRelay.TextMatrix(grdRelay.Row, CODEINDEX) = 0
                If Trim$(grdRelay.TextMatrix(grdRelay.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdRelay.Left - pbcArrow.Width - 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + (grdRelay.RowHeight(grdRelay.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdRelay.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdRelay.Row + 1 >= grdRelay.Rows Then
                        grdRelay.AddItem ""
                    End If
                    grdRelay.Row = grdRelay.Row + 1
                    If Not grdRelay.RowIsVisible(grdRelay.Row) Then
                        imIgnoreScroll = True
                        grdRelay.TopRow = grdRelay.TopRow + 1
                    End If
                    grdRelay.Col = NAMEINDEX
                    grdRelay.TextMatrix(grdRelay.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdRelay.Left - pbcArrow.Width - 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + (grdRelay.RowHeight(grdRelay.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdRelay.Col = grdRelay.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdRelay.TopRow = grdRelay.FixedRows
        grdRelay.Col = NAMEINDEX
        grdRelay.Row = grdRelay.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdRelay.TopRow
    llRow = grdRelay.Row
    slMsg = "Insert above " & Trim$(grdRelay.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdRelay.Redraw = False
    grdRelay.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdRelay.Row = llRow
    grdRelay.Redraw = False
    grdRelay.TopRow = llTRow
    grdRelay.Redraw = True
    DoEvents
    grdRelay.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdRelay.TopRow
    llRow = grdRelay.Row
    If (Val(grdRelay.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdRelay.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdRelay.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdRelay.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdRelay.Redraw = False
    If (Val(grdRelay.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdRelay.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdRelay.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdRelay.AddItem ""
    grdRelay.Redraw = False
    grdRelay.TopRow = llTRow
    grdRelay.Redraw = True
    DoEvents
    grdRelay.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As RNE, tlOld As RNE) As Integer
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
        If UBound(tgCurrRNE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdRelay.FixedRows To grdRelay.Rows - 1 Step 1
            slStr = Trim$(grdRelay.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdRelay.Row = llRow
                    Do While Not grdRelay.RowIsVisible(grdRelay.Row)
                        imIgnoreScroll = True
                        grdRelay.TopRow = grdRelay.TopRow + 1
                    Loop
                    grdRelay.Col = NAMEINDEX
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
    For llRow = grdRelay.FixedRows To grdRelay.Rows - 1 Step 1
        slStr = Trim$(grdRelay.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdRelay.Row = llRow
            Do While Not grdRelay.RowIsVisible(grdRelay.Row)
                imIgnoreScroll = True
                grdRelay.TopRow = grdRelay.TopRow + 1
            Loop
            grdRelay.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdRelay.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mSetFocus()
    Select Case grdRelay.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdRelay.Left + grdRelay.ColPos(grdRelay.Col) + 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + 15, grdRelay.ColWidth(grdRelay.Col) - 30, grdRelay.RowHeight(grdRelay.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdRelay.Left + grdRelay.ColPos(grdRelay.Col) + 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + 15, grdRelay.ColWidth(grdRelay.Col) - 30, grdRelay.RowHeight(grdRelay.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdRelay.Left + grdRelay.ColPos(grdRelay.Col) + 30, grdRelay.Top + grdRelay.RowPos(grdRelay.Row) + 15, grdRelay.ColWidth(grdRelay.Col) - 30, grdRelay.RowHeight(grdRelay.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
