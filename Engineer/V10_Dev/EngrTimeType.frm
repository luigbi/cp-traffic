VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrTimeType 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrTimeType.frx":0000
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
      Left            =   11745
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   15
      Width           =   45
   End
   Begin VB.Frame frcSelect 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4800
      TabIndex        =   11
      Top             =   105
      Width           =   2025
      Begin VB.OptionButton rbcType 
         Caption         =   "Start"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "End"
         Height          =   255
         Index           =   1
         Left            =   885
         TabIndex        =   12
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox pbcState 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4050
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
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
      Picture         =   "EngrTimeType.frx":030A
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
      Top             =   6690
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9630
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
      Left            =   5205
      TabIndex        =   9
      Top             =   6690
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3420
      TabIndex        =   8
      Top             =   6690
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTimeType 
      Height          =   5835
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   10292
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
      Left            =   10245
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
      Left            =   8550
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2520
      Picture         =   "EngrTimeType.frx":0614
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Time Type"
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
      Picture         =   "EngrTimeType.frx":091E
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8775
      Picture         =   "EngrTimeType.frx":11E8
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "EngrTimeType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrTimeType - enters affiliate representative information
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
Private imTTECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmTTE As TTE

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
    llRow = gGrid_Search(grdTimeType, NAMEINDEX, slStr)
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
    
    grdTimeType.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdTimeType.FixedRows To grdTimeType.Rows - 1 Step 1
        slStr = Trim$(grdTimeType.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdTimeType.Rows - 1 Step 1
                slTestStr = Trim$(grdTimeType.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdTimeType.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdTimeType.Row = llRow
                        grdTimeType.Col = NAMEINDEX
                        grdTimeType.CellForeColor = vbRed
                    Else
                        grdTimeType.Row = llTestRow
                        grdTimeType.Col = NAMEINDEX
                        grdTimeType.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdTimeType.Redraw = True
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
    gGrid_SortByCol grdTimeType, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
        frcSelect.Enabled = False
        'Check that all mandatory answered
        ilRet = mCheckFields(False)
        If ilRet Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    Else
        cmcSave.Enabled = False
        frcSelect.Enabled = True
    End If
End Sub

Private Sub mEnableBox()
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(TIMETYPELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdTimeType.Row >= grdTimeType.FixedRows) And (grdTimeType.Row < grdTimeType.Rows) And (grdTimeType.Col >= 0) And (grdTimeType.Col < grdTimeType.Cols - 1) Then
        lmEnableRow = grdTimeType.Row
        lmEnableCol = grdTimeType.Col
        sgReturnCallName = grdTimeType.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdTimeType.Left - pbcArrow.Width - 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + (grdTimeType.RowHeight(grdTimeType.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdTimeType.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdTimeType.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdTimeType.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdTimeType.Col
            Case NAMEINDEX  'Call Letters
                edcGrid.Move grdTimeType.Left + grdTimeType.ColPos(grdTimeType.Col) + 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + 15, grdTimeType.ColWidth(grdTimeType.Col) - 30, grdTimeType.RowHeight(grdTimeType.Row) - 15
                'edcGrid.MaxLength = Len(tmTTE.sName)
                edcGrid.MaxLength = gGetAllowedChars("STARTTYPE", Len(tmTTE.sName))
                edcGrid.text = grdTimeType.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdTimeType.Left + grdTimeType.ColPos(grdTimeType.Col) + 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + 15, grdTimeType.ColWidth(grdTimeType.Col) - 30, grdTimeType.RowHeight(grdTimeType.Row) - 15
                edcGrid.MaxLength = Len(tmTTE.sDescription)
                edcGrid.text = grdTimeType.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdTimeType.Left + grdTimeType.ColPos(grdTimeType.Col) + 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + 15, grdTimeType.ColWidth(grdTimeType.Col) - 30, grdTimeType.RowHeight(grdTimeType.Row) - 15
                smState = grdTimeType.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdTimeType.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdTimeType.FixedRows) And (lmEnableRow < grdTimeType.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdTimeType.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdTimeType.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdTimeType.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdTimeType.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdTimeType.TextMatrix(lmEnableRow, NAMEINDEX)
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
    
    grdTimeType.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdTimeType.FixedRows To grdTimeType.Rows - 1 Step 1
        slStr = Trim$(grdTimeType.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdTimeType.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdTimeType.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdTimeType.Row = llRow
                grdTimeType.Col = NAMEINDEX
                grdTimeType.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdTimeType.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdTimeType.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdTimeType.Row = llRow
                    grdTimeType.Col = STATEINDEX
                    grdTimeType.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdTimeType.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdTimeType
    mGridColumnWidth
    'Set Titles
    grdTimeType.TextMatrix(0, NAMEINDEX) = "Name"
    grdTimeType.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdTimeType.TextMatrix(0, STATEINDEX) = "State"
    grdTimeType.Row = 1
    For ilCol = 0 To grdTimeType.Cols - 1 Step 1
        grdTimeType.Col = ilCol
        grdTimeType.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdTimeType.Height = cmcCancel.Top - grdTimeType.Top - 120    '8 * grdTimeType.RowHeight(0) + 30
    gGrid_IntegralHeight grdTimeType
    gGrid_Clear grdTimeType, True
    grdTimeType.Row = grdTimeType.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdTimeType.Width = EngrTimeType.Width - 2 * grdTimeType.Left
    grdTimeType.ColWidth(CODEINDEX) = 0
    grdTimeType.ColWidth(USEDFLAGINDEX) = 0
    grdTimeType.ColWidth(NAMEINDEX) = grdTimeType.Width / 9
    grdTimeType.ColWidth(STATEINDEX) = grdTimeType.Width / 15
    grdTimeType.ColWidth(DESCRIPTIONINDEX) = grdTimeType.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdTimeType.ColWidth(DESCRIPTIONINDEX) > grdTimeType.ColWidth(ilCol) Then
                grdTimeType.ColWidth(DESCRIPTIONINDEX) = grdTimeType.ColWidth(DESCRIPTIONINDEX) - grdTimeType.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    edcGrid.text = ""
    gGrid_Clear grdTimeType, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdTimeType.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdTimeType.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmTTE.iCode = Val(grdTimeType.TextMatrix(llRow, CODEINDEX))
    If rbcType(0).Value Then
        tmTTE.sType = "S"
    Else
        tmTTE.sType = "E"
    End If
    slStr = Trim$(grdTimeType.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmTTE.sName = ""
    Else
        tmTTE.sName = slStr
    End If
    tmTTE.sDescription = grdTimeType.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdTimeType.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmTTE.sState = "D"
    Else
        tmTTE.sState = "A"
    End If
    If tmTTE.iCode <= 0 Then
        tmTTE.sUsedFlag = "N"
    Else
        tmTTE.sUsedFlag = grdTimeType.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmTTE.iVersion = 0
    tmTTE.iOrigTteCode = tmTTE.iCode
    tmTTE.sCurrent = "Y"
    'tmTTE.sEnteredDate = smNowDate
    'tmTTE.sEnteredTime = smNowTime
    tmTTE.sEnteredDate = Format(Now, sgShowDateForm)
    tmTTE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmTTE.iUieCode = tgUIE.iCode
    tmTTE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdTimeType, True
    llRow = grdTimeType.FixedRows
    For ilLoop = 0 To UBound(tgCurrTTE) - 1 Step 1
        If llRow + 1 > grdTimeType.Rows Then
            grdTimeType.AddItem ""
        End If
        grdTimeType.Row = llRow
        grdTimeType.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrTTE(ilLoop).sName)
        grdTimeType.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrTTE(ilLoop).sDescription)
        If tgCurrTTE(ilLoop).sState = "A" Then
            grdTimeType.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdTimeType.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdTimeType.TextMatrix(llRow, CODEINDEX) = tgCurrTTE(ilLoop).iCode
        grdTimeType.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrTTE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdTimeType.Rows Then
        grdTimeType.AddItem ""
    End If
    grdTimeType.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    If rbcType(0).Value Then
        ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrTTEStamp, "EngrTimeType-mPopulate", tgCurrTTE())
    Else
        ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrTTEStamp, "EngrTimeType-mPopulate", tgCurrTTE())
    End If
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlTTE As TTE
    
    gSetMousePointer grdTimeType, grdTimeType, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdTimeType, grdTimeType, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdTimeType, grdTimeType, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdTimeType.Redraw = False
    For llRow = grdTimeType.FixedRows To grdTimeType.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmTTE.sName) <> "" Then
            imTTECode = tmTTE.iCode
            If tmTTE.iCode > 0 Then
                ilRet = gGetRec_TTE_TimeType(imTTECode, "Time Type-mSave: Get TTE", tlTTE)
                If ilRet Then
                    If mCompare(tmTTE, tlTTE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmTTE.iVersion = tlTTE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmTTE.iCode <= 0 Then
                    ilRet = gPutInsert_TTE_TimeType(0, tmTTE, "Time Type-mSave: Insert TTE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_TTE_TimeType(1, tmTTE, "Time Type-mSave: Update TTE")
                    ilRet = gPutDelete_TTE_TimeType(tmTTE.iCode, "Time Type-mSave: Delete TTE")
                    ilRet = gPutInsert_TTE_TimeType(1, tmTTE, "Time Type-mSave: Insert TTE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_TTE_TimeType(imDeleteCodes(ilLoop), "EngrTimeType- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdTimeType.Redraw = True
    sgCurrTTEStamp = ""
    If rbcType(0).Value Then
        ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrTTEStamp, "EngrTimeType-mPopulate", tgCurrTTE())
    Else
        ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrTTEStamp, "EngrTimeType-mPopulate", tgCurrTTE())
    End If
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrTimeType
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrTimeType
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdTimeType, grdTimeType, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdTimeType, grdTimeType, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdTimeType, grdTimeType, vbDefault
    Unload EngrTimeType
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
        gSetMousePointer grdTimeType, grdTimeType, vbHourglass
        llTopRow = grdTimeType.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdTimeType, grdTimeType, vbDefault
            Exit Sub
        End If
        grdTimeType.Redraw = False
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
        grdTimeType.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdTimeType, grdTimeType, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdTimeType.Col
        Case NAMEINDEX
            If grdTimeType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdTimeType.text = edcGrid.text
            grdTimeType.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdTimeType.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdTimeType.text = edcGrid.text
            grdTimeType.CellForeColor = vbBlack
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
    gSetFonts EngrTimeType
    gCenterFormModal EngrTimeType
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdTimeType.FixedRows) And (lmEnableRow < grdTimeType.Rows) Then
            If (lmEnableCol >= grdTimeType.FixedCols) And (lmEnableCol < grdTimeType.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdTimeType.text = smESCValue
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
    grdTimeType.Height = cmcCancel.Top - grdTimeType.Top - 120    '8 * grdTimeType.RowHeight(0) + 30
    gGrid_IntegralHeight grdTimeType
    gGrid_FillWithRows grdTimeType
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrTimeType = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdTimeType, grdTimeType, vbHourglass
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
    If igInitCallInfo = 1 Then
        frcSelect.Visible = False
    ElseIf igInitCallInfo = 2 Then
        frcSelect.Visible = False
        rbcType(1).Value = True
    End If
    mPopulate
    mMoveRecToCtrls
    mSortCol 0
    imInChg = False
    imFieldChgd = False
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(TIMETYPELIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdTimeType, grdTimeType, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdTimeType, grdTimeType, vbDefault
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
    igRptIndex = TIMETYPE_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub grdTimeType_Click()
    If grdTimeType.Col >= grdTimeType.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdTimeType_EnterCell()
    mSetShow
End Sub

Private Sub grdTimeType_GotFocus()
    If grdTimeType.Col >= grdTimeType.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdTimeType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdTimeType.TopRow
    grdTimeType.Redraw = False
End Sub

Private Sub grdTimeType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdTimeType.RowHeight(0) Then
        mSortCol grdTimeType.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdTimeType, x, y)
    If Not ilFound Then
        grdTimeType.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdTimeType.Col >= grdTimeType.Cols - 1 Then
        grdTimeType.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdTimeType.TopRow
    DoEvents
    llRow = grdTimeType.Row
    If grdTimeType.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdTimeType.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdTimeType.TextMatrix(llRow, NAMEINDEX) = ""
        grdTimeType.Row = llRow + 1
        grdTimeType.Col = NAMEINDEX
        grdTimeType.Redraw = True
    End If
    grdTimeType.Redraw = True
    mEnableBox
End Sub

Private Sub grdTimeType_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdTimeType.Redraw = False Then
        grdTimeType.Redraw = True
        If lmTopRow < grdTimeType.FixedRows Then
            grdTimeType.TopRow = grdTimeType.FixedRows
        Else
            grdTimeType.TopRow = lmTopRow
        End If
        grdTimeType.Refresh
        grdTimeType.Redraw = False
    End If
    If (imShowGridBox) And (grdTimeType.Row >= grdTimeType.FixedRows) And (grdTimeType.Col >= 0) And (grdTimeType.Col < grdTimeType.Cols - 1) Then
        If grdTimeType.RowIsVisible(grdTimeType.Row) Then
            'edcGrid.Move grdTimeType.Left + grdTimeType.ColPos(grdTimeType.Col) + 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + 30, grdTimeType.ColWidth(grdTimeType.Col) - 30, grdTimeType.RowHeight(grdTimeType.Row) - 30
            pbcArrow.Move grdTimeType.Left - pbcArrow.Width - 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + (grdTimeType.RowHeight(grdTimeType.Row) - pbcArrow.Height) / 2
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
        If grdTimeType.Col = NAMEINDEX Then
            If grdTimeType.Row > grdTimeType.FixedRows Then
                lmTopRow = -1
                grdTimeType.Row = grdTimeType.Row - 1
                If Not grdTimeType.RowIsVisible(grdTimeType.Row) Then
                    grdTimeType.TopRow = grdTimeType.TopRow - 1
                End If
                grdTimeType.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdTimeType.Col = grdTimeType.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdTimeType.TopRow = grdTimeType.FixedRows
        grdTimeType.Col = NAMEINDEX
        grdTimeType.Row = grdTimeType.FixedRows
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
        grdTimeType.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdTimeType.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdTimeType.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdTimeType.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdTimeType.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdTimeType.CellForeColor = vbBlack
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
        If grdTimeType.Col = STATEINDEX Then
            llRow = grdTimeType.Rows
            Do
                llRow = llRow - 1
            Loop While grdTimeType.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdTimeType.Row + 1 < llRow) Then
                lmTopRow = -1
                grdTimeType.Row = grdTimeType.Row + 1
                If Not grdTimeType.RowIsVisible(grdTimeType.Row) Then
                    imIgnoreScroll = True
                    grdTimeType.TopRow = grdTimeType.TopRow + 1
                End If
                grdTimeType.Col = NAMEINDEX
                'grdTimeType.TextMatrix(grdTimeType.Row, CODEINDEX) = 0
                If Trim$(grdTimeType.TextMatrix(grdTimeType.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdTimeType.Left - pbcArrow.Width - 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + (grdTimeType.RowHeight(grdTimeType.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdTimeType.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdTimeType.Row + 1 >= grdTimeType.Rows Then
                        grdTimeType.AddItem ""
                    End If
                    grdTimeType.Row = grdTimeType.Row + 1
                    If Not grdTimeType.RowIsVisible(grdTimeType.Row) Then
                        imIgnoreScroll = True
                        grdTimeType.TopRow = grdTimeType.TopRow + 1
                    End If
                    grdTimeType.Col = NAMEINDEX
                    grdTimeType.TextMatrix(grdTimeType.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdTimeType.Left - pbcArrow.Width - 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + (grdTimeType.RowHeight(grdTimeType.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdTimeType.Col = grdTimeType.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdTimeType.TopRow = grdTimeType.FixedRows
        grdTimeType.Col = NAMEINDEX
        grdTimeType.Row = grdTimeType.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdTimeType.TopRow
    llRow = grdTimeType.Row
    slMsg = "Insert above " & Trim$(grdTimeType.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdTimeType.Redraw = False
    grdTimeType.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdTimeType.Row = llRow
    grdTimeType.Redraw = False
    grdTimeType.TopRow = llTRow
    grdTimeType.Redraw = True
    DoEvents
    grdTimeType.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdTimeType.TopRow
    llRow = grdTimeType.Row
    If (Val(grdTimeType.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdTimeType.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdTimeType.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdTimeType.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdTimeType.Redraw = False
    If (Val(grdTimeType.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdTimeType.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdTimeType.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdTimeType.AddItem ""
    grdTimeType.Redraw = False
    grdTimeType.TopRow = llTRow
    grdTimeType.Redraw = True
    DoEvents
    grdTimeType.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Sub rbcType_Click(Index As Integer)
    If igInitCallInfo <> 0 Then
        Exit Sub
    End If
    If rbcType(Index).Value Then
        imInChg = True
        mClearControls
        imLastColSorted = -1
        imLastSort = -1
        lmEnableRow = -1
        mPopulate
        mMoveRecToCtrls
        mSortCol 0
        imInChg = False
        imFieldChgd = False
    End If

End Sub

Private Sub rbcType_GotFocus(Index As Integer)
    mSetShow
End Sub

Private Function mCompare(tlNew As TTE, tlOld As TTE) As Integer
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
        If UBound(tgCurrTTE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdTimeType.FixedRows To grdTimeType.Rows - 1 Step 1
            slStr = Trim$(grdTimeType.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdTimeType.Row = llRow
                    Do While Not grdTimeType.RowIsVisible(grdTimeType.Row)
                        imIgnoreScroll = True
                        grdTimeType.TopRow = grdTimeType.TopRow + 1
                    Loop
                    grdTimeType.Col = NAMEINDEX
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
    For llRow = grdTimeType.FixedRows To grdTimeType.Rows - 1 Step 1
        slStr = Trim$(grdTimeType.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdTimeType.Row = llRow
            Do While Not grdTimeType.RowIsVisible(grdTimeType.Row)
                imIgnoreScroll = True
                grdTimeType.TopRow = grdTimeType.TopRow + 1
            Loop
            grdTimeType.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdTimeType.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Sub mSetFocus()
    Select Case grdTimeType.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdTimeType.Left + grdTimeType.ColPos(grdTimeType.Col) + 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + 15, grdTimeType.ColWidth(grdTimeType.Col) - 30, grdTimeType.RowHeight(grdTimeType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdTimeType.Left + grdTimeType.ColPos(grdTimeType.Col) + 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + 15, grdTimeType.ColWidth(grdTimeType.Col) - 30, grdTimeType.RowHeight(grdTimeType.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdTimeType.Left + grdTimeType.ColPos(grdTimeType.Col) + 30, grdTimeType.Top + grdTimeType.RowPos(grdTimeType.Row) + 15, grdTimeType.ColWidth(grdTimeType.Col) - 30, grdTimeType.RowHeight(grdTimeType.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
