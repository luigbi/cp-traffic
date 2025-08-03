VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrDaySubName 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrDaySubName.frx":0000
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
      Top             =   0
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
      Picture         =   "EngrDaySubName.frx":030A
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDaySubName 
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
      Picture         =   "EngrDaySubName.frx":0614
      Top             =   6615
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Day SubName"
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
      Picture         =   "EngrDaySubName.frx":091E
      Top             =   6615
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8775
      Picture         =   "EngrDaySubName.frx":11E8
      Top             =   6615
      Width           =   480
   End
End
Attribute VB_Name = "EngrDaySubName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrDaySubName - enters affiliate representative information
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
Private lmDSECode As Long
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmDSE As DSE

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
    llRow = gGrid_Search(grdDaySubName, NAMEINDEX, slStr)
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
    
    grdDaySubName.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdDaySubName.FixedRows To grdDaySubName.Rows - 1 Step 1
        slStr = Trim$(grdDaySubName.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdDaySubName.Rows - 1 Step 1
                slTestStr = Trim$(grdDaySubName.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdDaySubName.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdDaySubName.Row = llRow
                        grdDaySubName.Col = NAMEINDEX
                        grdDaySubName.CellForeColor = vbRed
                    Else
                        grdDaySubName.Row = llTestRow
                        grdDaySubName.Col = NAMEINDEX
                        grdDaySubName.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdDaySubName.Redraw = True
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
    gGrid_SortByCol grdDaySubName, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If (grdDaySubName.Row >= grdDaySubName.FixedRows) And (grdDaySubName.Row < grdDaySubName.Rows) And (grdDaySubName.Col >= 0) And (grdDaySubName.Col < grdDaySubName.Cols - 1) Then
        lmEnableRow = grdDaySubName.Row
        lmEnableCol = grdDaySubName.Col
        sgReturnCallName = grdDaySubName.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdDaySubName.Left - pbcArrow.Width - 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + (grdDaySubName.RowHeight(grdDaySubName.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdDaySubName.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdDaySubName.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdDaySubName.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdDaySubName.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdDaySubName.Left + grdDaySubName.ColPos(grdDaySubName.Col) + 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + 15, grdDaySubName.ColWidth(grdDaySubName.Col) - 30, grdDaySubName.RowHeight(grdDaySubName.Row) - 15
                edcGrid.MaxLength = Len(tmDSE.sName)
                edcGrid.text = grdDaySubName.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdDaySubName.Left + grdDaySubName.ColPos(grdDaySubName.Col) + 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + 15, grdDaySubName.ColWidth(grdDaySubName.Col) - 30, grdDaySubName.RowHeight(grdDaySubName.Row) - 15
                edcGrid.MaxLength = Len(tmDSE.sDescription)
                edcGrid.text = grdDaySubName.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdDaySubName.Left + grdDaySubName.ColPos(grdDaySubName.Col) + 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + 15, grdDaySubName.ColWidth(grdDaySubName.Col) - 30, grdDaySubName.RowHeight(grdDaySubName.Row) - 15
                smState = grdDaySubName.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdDaySubName.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdDaySubName.FixedRows) And (lmEnableRow < grdDaySubName.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdDaySubName.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdDaySubName.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdDaySubName.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdDaySubName.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdDaySubName.TextMatrix(lmEnableRow, NAMEINDEX)
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
    
    grdDaySubName.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdDaySubName.FixedRows To grdDaySubName.Rows - 1 Step 1
        slStr = Trim$(grdDaySubName.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdDaySubName.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdDaySubName.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdDaySubName.Row = llRow
                grdDaySubName.Col = NAMEINDEX
                grdDaySubName.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdDaySubName.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdDaySubName.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdDaySubName.Row = llRow
                    grdDaySubName.Col = STATEINDEX
                    grdDaySubName.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdDaySubName.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdDaySubName
    mGridColumnWidth
    'Set Titles
    grdDaySubName.TextMatrix(0, NAMEINDEX) = "Name"
    grdDaySubName.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdDaySubName.TextMatrix(0, STATEINDEX) = "State"
    grdDaySubName.Row = 1
    For ilCol = 0 To grdDaySubName.Cols - 1 Step 1
        grdDaySubName.Col = ilCol
        grdDaySubName.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdDaySubName.Height = cmcCancel.Top - grdDaySubName.Top - 120    '8 * grdDaySubName.RowHeight(0) + 30
    gGrid_IntegralHeight grdDaySubName
    gGrid_Clear grdDaySubName, True
    grdDaySubName.Row = grdDaySubName.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdDaySubName.Width = EngrDaySubName.Width - 2 * grdDaySubName.Left
    grdDaySubName.ColWidth(CODEINDEX) = 0
    grdDaySubName.ColWidth(USEDFLAGINDEX) = 0
    grdDaySubName.ColWidth(NAMEINDEX) = grdDaySubName.Width / 5
    grdDaySubName.ColWidth(STATEINDEX) = grdDaySubName.Width / 15
    grdDaySubName.ColWidth(DESCRIPTIONINDEX) = grdDaySubName.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdDaySubName.ColWidth(DESCRIPTIONINDEX) > grdDaySubName.ColWidth(ilCol) Then
                grdDaySubName.ColWidth(DESCRIPTIONINDEX) = grdDaySubName.ColWidth(DESCRIPTIONINDEX) - grdDaySubName.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdDaySubName, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdDaySubName.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdDaySubName.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmDSE.lCode = Val(grdDaySubName.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdDaySubName.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmDSE.sName = ""
    Else
        tmDSE.sName = slStr
    End If
    tmDSE.sDescription = grdDaySubName.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdDaySubName.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmDSE.sState = "D"
    Else
        tmDSE.sState = "A"
    End If
    If tmDSE.lCode <= 0 Then
        tmDSE.sUsedFlag = "N"
    Else
        tmDSE.sUsedFlag = grdDaySubName.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmDSE.iVersion = 0
    tmDSE.lOrigDseCode = tmDSE.lCode
    tmDSE.sCurrent = "Y"
    'tmDSE.sEnteredDate = smNowDate
    'tmDSE.sEnteredTime = smNowTime
    tmDSE.sEnteredDate = Format(Now, sgShowDateForm)
    tmDSE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmDSE.iUieCode = tgUIE.iCode
    tmDSE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdDaySubName, True
    llRow = grdDaySubName.FixedRows
    For ilLoop = 0 To UBound(tgCurrDSE) - 1 Step 1
        If llRow + 1 > grdDaySubName.Rows Then
            grdDaySubName.AddItem ""
        End If
        grdDaySubName.Row = llRow
        grdDaySubName.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrDSE(ilLoop).sName)
        grdDaySubName.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrDSE(ilLoop).sDescription)
        If tgCurrDSE(ilLoop).sState = "A" Then
            grdDaySubName.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdDaySubName.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdDaySubName.TextMatrix(llRow, CODEINDEX) = tgCurrDSE(ilLoop).lCode
        grdDaySubName.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrDSE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdDaySubName.Rows Then
        grdDaySubName.AddItem ""
    End If
    grdDaySubName.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrDaySubName-mPopulate", tgCurrDSE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlDSE As DSE
    
    gSetMousePointer grdDaySubName, grdDaySubName, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdDaySubName, grdDaySubName, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdDaySubName, grdDaySubName, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdDaySubName.Redraw = False
    For llRow = grdDaySubName.FixedRows To grdDaySubName.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmDSE.sName) <> "" Then
            lmDSECode = tmDSE.lCode
            If tmDSE.lCode > 0 Then
                ilRet = gGetRec_DSE_DaySubName(lmDSECode, "Day SubName-mSave: Get DSE", tlDSE)
                If ilRet Then
                    If mCompare(tmDSE, tlDSE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmDSE.iVersion = tlDSE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmDSE.lCode <= 0 Then
                    ilRet = gPutInsert_DSE_DaySubName(0, tmDSE, "Day SubName-mSave: Insert DSE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_DSE_DaySubName(1, tmDSE, "Day SubName-mSave: Update DSE")
                    ilRet = gPutDelete_DSE_DaySubName(tmDSE.lCode, "Day SubName-mSave: Delete DSE")
                    ilRet = gPutInsert_DSE_DaySubName(1, tmDSE, "Day SubName-mSave: Insert DSE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(lmDeleteCodes) To UBound(lmDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_DSE_DaySubName(lmDeleteCodes(ilLoop), "EngrDaySubName- Delete")
    Next ilLoop
    ReDim lmDeleteCodes(0 To 0) As Long
    grdDaySubName.Redraw = True
    sgCurrDSEStamp = ""
    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrDaySubName-mPopulate", tgCurrDSE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrDaySubName
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrDaySubName
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdDaySubName, grdDaySubName, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdDaySubName, grdDaySubName, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdDaySubName, grdDaySubName, vbDefault
    Unload EngrDaySubName
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
        gSetMousePointer grdDaySubName, grdDaySubName, vbHourglass
        llTopRow = grdDaySubName.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdDaySubName, grdDaySubName, vbDefault
            Exit Sub
        End If
        grdDaySubName.Redraw = False
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
        grdDaySubName.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdDaySubName, grdDaySubName, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdDaySubName.Col
        Case NAMEINDEX
            If grdDaySubName.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdDaySubName.text = edcGrid.text
            grdDaySubName.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdDaySubName.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdDaySubName.text = edcGrid.text
            grdDaySubName.CellForeColor = vbBlack
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
    gSetFonts EngrDaySubName
    gCenterFormModal EngrDaySubName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdDaySubName.FixedRows) And (lmEnableRow < grdDaySubName.Rows) Then
            If (lmEnableCol >= grdDaySubName.FixedCols) And (lmEnableCol < grdDaySubName.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdDaySubName.text = smESCValue
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
    grdDaySubName.Height = cmcCancel.Top - grdDaySubName.Top - 120    '8 * grdDaySubName.RowHeight(0) + 30
    gGrid_IntegralHeight grdDaySubName
    gGrid_FillWithRows grdDaySubName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase lmDeleteCodes
    Set EngrDaySubName = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdDaySubName, grdDaySubName, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    imcInsert.Picture = EngrMain!imcInsert.Picture
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    ReDim lmDeleteCodes(0 To 0) As Long
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
    If igInitCallInfo = 1 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(LIBRARYJOB) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
            imcInsert.Enabled = False
            imcTrash.Enabled = False
        End If
    ElseIf igInitCallInfo = 2 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(TEMPLATEJOB) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
            imcInsert.Enabled = False
            imcTrash.Enabled = False
        End If
    Else
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(LIBRARYJOB) = 2) Or (igListStatus(TEMPLATEJOB) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
            imcInsert.Enabled = False
            imcTrash.Enabled = False
        End If
    End If
    gSetMousePointer grdDaySubName, grdDaySubName, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdDaySubName, grdDaySubName, vbDefault
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

Private Sub grdDaySubName_Click()
    If grdDaySubName.Col >= grdDaySubName.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdDaySubName_EnterCell()
    mSetShow
End Sub

Private Sub grdDaySubName_GotFocus()
    If grdDaySubName.Col >= grdDaySubName.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdDaySubName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdDaySubName.TopRow
    grdDaySubName.Redraw = False
End Sub

Private Sub grdDaySubName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdDaySubName.RowHeight(0) Then
        mSortCol grdDaySubName.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdDaySubName, x, y)
    If Not ilFound Then
        grdDaySubName.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdDaySubName.Col >= grdDaySubName.Cols - 1 Then
        grdDaySubName.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdDaySubName.TopRow
    DoEvents
    llRow = grdDaySubName.Row
    If grdDaySubName.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdDaySubName.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdDaySubName.TextMatrix(llRow, NAMEINDEX) = ""
        grdDaySubName.Row = llRow + 1
        grdDaySubName.Col = NAMEINDEX
        grdDaySubName.Redraw = True
    End If
    grdDaySubName.Redraw = True
    mEnableBox
End Sub

Private Sub grdDaySubName_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdDaySubName.Redraw = False Then
        grdDaySubName.Redraw = True
        If lmTopRow < grdDaySubName.FixedRows Then
            grdDaySubName.TopRow = grdDaySubName.FixedRows
        Else
            grdDaySubName.TopRow = lmTopRow
        End If
        grdDaySubName.Refresh
        grdDaySubName.Redraw = False
    End If
    If (imShowGridBox) And (grdDaySubName.Row >= grdDaySubName.FixedRows) And (grdDaySubName.Col >= 0) And (grdDaySubName.Col < grdDaySubName.Cols - 1) Then
        If grdDaySubName.RowIsVisible(grdDaySubName.Row) Then
            'edcGrid.Move grdDaySubName.Left + grdDaySubName.ColPos(grdDaySubName.Col) + 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + 30, grdDaySubName.ColWidth(grdDaySubName.Col) - 30, grdDaySubName.RowHeight(grdDaySubName.Row) - 30
            pbcArrow.Move grdDaySubName.Left - pbcArrow.Width - 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + (grdDaySubName.RowHeight(grdDaySubName.Row) - pbcArrow.Height) / 2
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
        If grdDaySubName.Col = NAMEINDEX Then
            If grdDaySubName.Row > grdDaySubName.FixedRows Then
                lmTopRow = -1
                grdDaySubName.Row = grdDaySubName.Row - 1
                If Not grdDaySubName.RowIsVisible(grdDaySubName.Row) Then
                    grdDaySubName.TopRow = grdDaySubName.TopRow - 1
                End If
                grdDaySubName.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdDaySubName.Col = grdDaySubName.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdDaySubName.TopRow = grdDaySubName.FixedRows
        grdDaySubName.Col = NAMEINDEX
        grdDaySubName.Row = grdDaySubName.FixedRows
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
        grdDaySubName.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdDaySubName.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdDaySubName.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdDaySubName.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdDaySubName.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdDaySubName.CellForeColor = vbBlack
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
        If grdDaySubName.Col = STATEINDEX Then
            llRow = grdDaySubName.Rows
            Do
                llRow = llRow - 1
            Loop While grdDaySubName.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdDaySubName.Row + 1 < llRow) Then
                lmTopRow = -1
                grdDaySubName.Row = grdDaySubName.Row + 1
                If Not grdDaySubName.RowIsVisible(grdDaySubName.Row) Then
                    imIgnoreScroll = True
                    grdDaySubName.TopRow = grdDaySubName.TopRow + 1
                End If
                grdDaySubName.Col = NAMEINDEX
                'grdDaySubName.TextMatrix(grdDaySubName.Row, CODEINDEX) = 0
                If Trim$(grdDaySubName.TextMatrix(grdDaySubName.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdDaySubName.Left - pbcArrow.Width - 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + (grdDaySubName.RowHeight(grdDaySubName.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdDaySubName.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdDaySubName.Row + 1 >= grdDaySubName.Rows Then
                        grdDaySubName.AddItem ""
                    End If
                    grdDaySubName.Row = grdDaySubName.Row + 1
                    If Not grdDaySubName.RowIsVisible(grdDaySubName.Row) Then
                        imIgnoreScroll = True
                        grdDaySubName.TopRow = grdDaySubName.TopRow + 1
                    End If
                    grdDaySubName.Col = NAMEINDEX
                    grdDaySubName.TextMatrix(grdDaySubName.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdDaySubName.Left - pbcArrow.Width - 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + (grdDaySubName.RowHeight(grdDaySubName.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdDaySubName.Col = grdDaySubName.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdDaySubName.TopRow = grdDaySubName.FixedRows
        grdDaySubName.Col = NAMEINDEX
        grdDaySubName.Row = grdDaySubName.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdDaySubName.TopRow
    llRow = grdDaySubName.Row
    slMsg = "Insert above " & Trim$(grdDaySubName.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdDaySubName.Redraw = False
    grdDaySubName.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdDaySubName.Row = llRow
    grdDaySubName.Redraw = False
    grdDaySubName.TopRow = llTRow
    grdDaySubName.Redraw = True
    DoEvents
    grdDaySubName.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdDaySubName.TopRow
    llRow = grdDaySubName.Row
    If (Val(grdDaySubName.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdDaySubName.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdDaySubName.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdDaySubName.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdDaySubName.Redraw = False
    If (Val(grdDaySubName.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        lmDeleteCodes(UBound(lmDeleteCodes)) = Val(grdDaySubName.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve lmDeleteCodes(0 To UBound(lmDeleteCodes) + 1) As Long
    End If
    grdDaySubName.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdDaySubName.AddItem ""
    grdDaySubName.Redraw = False
    grdDaySubName.TopRow = llTRow
    grdDaySubName.Redraw = True
    DoEvents
    grdDaySubName.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mCompare(tlNew As DSE, tlOld As DSE) As Integer
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
        If UBound(tgCurrDSE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdDaySubName.FixedRows To grdDaySubName.Rows - 1 Step 1
            slStr = Trim$(grdDaySubName.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdDaySubName.Row = llRow
                    Do While Not grdDaySubName.RowIsVisible(grdDaySubName.Row)
                        imIgnoreScroll = True
                        grdDaySubName.TopRow = grdDaySubName.TopRow + 1
                    Loop
                    grdDaySubName.Col = NAMEINDEX
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
    For llRow = grdDaySubName.FixedRows To grdDaySubName.Rows - 1 Step 1
        slStr = Trim$(grdDaySubName.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdDaySubName.Row = llRow
            Do While Not grdDaySubName.RowIsVisible(grdDaySubName.Row)
                imIgnoreScroll = True
                grdDaySubName.TopRow = grdDaySubName.TopRow + 1
            Loop
            grdDaySubName.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdDaySubName.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mSetFocus()
    Select Case grdDaySubName.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdDaySubName.Left + grdDaySubName.ColPos(grdDaySubName.Col) + 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + 15, grdDaySubName.ColWidth(grdDaySubName.Col) - 30, grdDaySubName.RowHeight(grdDaySubName.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdDaySubName.Left + grdDaySubName.ColPos(grdDaySubName.Col) + 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + 15, grdDaySubName.ColWidth(grdDaySubName.Col) - 30, grdDaySubName.RowHeight(grdDaySubName.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdDaySubName.Left + grdDaySubName.ColPos(grdDaySubName.Col) + 30, grdDaySubName.Top + grdDaySubName.RowPos(grdDaySubName.Row) + 15, grdDaySubName.ColWidth(grdDaySubName.Col) - 30, grdDaySubName.RowHeight(grdDaySubName.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
