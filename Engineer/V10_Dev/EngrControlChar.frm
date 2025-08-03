VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrControlChar 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrControlChar.frx":0000
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   60
      Width           =   45
   End
   Begin VB.Frame frcSelect 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3930
      TabIndex        =   11
      Top             =   90
      Width           =   2025
      Begin VB.OptionButton rbcType 
         Caption         =   "Audio"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton rbcType 
         Caption         =   "Bus"
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
      Left            =   6030
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   825
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
      Left            =   120
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   6780
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
      Picture         =   "EngrControlChar.frx":030A
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
      Top             =   6660
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9630
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
      Left            =   5205
      TabIndex        =   9
      Top             =   6660
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3420
      TabIndex        =   8
      Top             =   6660
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdControlChar 
      Height          =   5925
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10845
      _ExtentX        =   19129
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
      Left            =   8295
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2520
      Picture         =   "EngrControlChar.frx":0614
      Top             =   6570
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Control Characters"
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
      Picture         =   "EngrControlChar.frx":091E
      Top             =   6570
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8775
      Picture         =   "EngrControlChar.frx":11E8
      Top             =   6570
      Width           =   480
   End
End
Attribute VB_Name = "EngrControlChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrControlChar - enters affiliate representative information
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
Private imCCECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private tmCCE As CCE

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
    llRow = gGrid_Search(grdControlChar, NAMEINDEX, slStr)
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
    
    grdControlChar.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdControlChar.FixedRows To grdControlChar.Rows - 1 Step 1
        slStr = Trim$(grdControlChar.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdControlChar.Rows - 1 Step 1
                slTestStr = Trim$(grdControlChar.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdControlChar.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdControlChar.Row = llRow
                        grdControlChar.Col = NAMEINDEX
                        grdControlChar.CellForeColor = vbRed
                    Else
                        grdControlChar.Row = llTestRow
                        grdControlChar.Col = NAMEINDEX
                        grdControlChar.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdControlChar.Redraw = True
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
    mSetShow
    gGrid_SortByCol grdControlChar, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    If rbcType(0).Value Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIOLIST) <> 2) Then
            cmcCancel.SetFocus
            Exit Sub
        End If
    Else
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(BUSLIST) <> 2) Then
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    If (grdControlChar.Row >= grdControlChar.FixedRows) And (grdControlChar.Row < grdControlChar.Rows) And (grdControlChar.Col >= 0) And (grdControlChar.Col < grdControlChar.Cols - 1) Then
        lmEnableRow = grdControlChar.Row
        lmEnableCol = grdControlChar.Col
        sgReturnCallName = grdControlChar.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdControlChar.Left - pbcArrow.Width - 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + (grdControlChar.RowHeight(grdControlChar.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdControlChar.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdControlChar.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdControlChar.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdControlChar.Col
            Case NAMEINDEX
                edcGrid.Move grdControlChar.Left + grdControlChar.ColPos(grdControlChar.Col) + 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + 15, grdControlChar.ColWidth(grdControlChar.Col) - 30, grdControlChar.RowHeight(grdControlChar.Row) - 15
                edcGrid.MaxLength = Len(tmCCE.sAutoChar)
                edcGrid.text = grdControlChar.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
                edcGrid.Move grdControlChar.Left + grdControlChar.ColPos(grdControlChar.Col) + 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + 15, grdControlChar.ColWidth(grdControlChar.Col) - 30, grdControlChar.RowHeight(grdControlChar.Row) - 15
                edcGrid.MaxLength = Len(tmCCE.sDescription)
                edcGrid.text = grdControlChar.text
                edcGrid.Visible = True
                edcGrid.SetFocus
            Case STATEINDEX
                pbcState.Move grdControlChar.Left + grdControlChar.ColPos(grdControlChar.Col) + 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + 15, grdControlChar.ColWidth(grdControlChar.Col) - 30, grdControlChar.RowHeight(grdControlChar.Row) - 15
                smState = grdControlChar.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
                pbcState.Visible = True
                pbcState.SetFocus
        End Select
        smESCValue = grdControlChar.text
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdControlChar.FixedRows) And (lmEnableRow < grdControlChar.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdControlChar.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdControlChar.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdControlChar.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdControlChar.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdControlChar.TextMatrix(lmEnableRow, NAMEINDEX)
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
    
    grdControlChar.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdControlChar.FixedRows To grdControlChar.Rows - 1 Step 1
        slStr = Trim$(grdControlChar.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdControlChar.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdControlChar.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdControlChar.Row = llRow
                grdControlChar.Col = NAMEINDEX
                grdControlChar.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdControlChar.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdControlChar.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdControlChar.Row = llRow
                    grdControlChar.Col = STATEINDEX
                    grdControlChar.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdControlChar.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdControlChar
    mGridColumnWidth
    'Set Titles
    grdControlChar.TextMatrix(0, NAMEINDEX) = "Name"
    grdControlChar.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdControlChar.TextMatrix(0, STATEINDEX) = "State"
    grdControlChar.Row = 1
    For ilCol = 0 To grdControlChar.Cols - 1 Step 1
        grdControlChar.Col = ilCol
        grdControlChar.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdControlChar.Height = cmcCancel.Top - grdControlChar.Top - 120    '8 * grdControlChar.RowHeight(0) + 30
    gGrid_IntegralHeight grdControlChar
    gGrid_Clear grdControlChar, True
    grdControlChar.Row = grdControlChar.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdControlChar.Width = EngrControlChar.Width - 2 * grdControlChar.Left
    grdControlChar.ColWidth(CODEINDEX) = 0
    grdControlChar.ColWidth(USEDFLAGINDEX) = 0
    grdControlChar.ColWidth(NAMEINDEX) = grdControlChar.Width / 9
    grdControlChar.ColWidth(STATEINDEX) = grdControlChar.Width / 15
    grdControlChar.ColWidth(DESCRIPTIONINDEX) = grdControlChar.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdControlChar.ColWidth(DESCRIPTIONINDEX) > grdControlChar.ColWidth(ilCol) Then
                grdControlChar.ColWidth(DESCRIPTIONINDEX) = grdControlChar.ColWidth(DESCRIPTIONINDEX) - grdControlChar.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    edcGrid.text = ""
    gGrid_Clear grdControlChar, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdControlChar.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdControlChar.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmCCE.iCode = Val(grdControlChar.TextMatrix(llRow, CODEINDEX))
    If rbcType(0).Value Then
        tmCCE.sType = "A"
    Else
        tmCCE.sType = "B"
    End If
    slStr = Trim$(grdControlChar.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmCCE.sAutoChar = ""
    Else
        tmCCE.sAutoChar = slStr
    End If
    tmCCE.sDescription = grdControlChar.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdControlChar.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmCCE.sState = "D"
    Else
        tmCCE.sState = "A"
    End If
    If tmCCE.iCode <= 0 Then
        tmCCE.sUsedFlag = "N"
    Else
        tmCCE.sUsedFlag = grdControlChar.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmCCE.iVersion = 0
    tmCCE.iOrigCceCode = tmCCE.iCode
    tmCCE.sCurrent = "Y"
    'tmCCE.sEnteredDate = smNowDate
    'tmCCE.sEnteredTime = smNowTime
    tmCCE.sEnteredDate = Format(Now, sgShowDateForm)
    tmCCE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmCCE.iUieCode = tgUIE.iCode
    tmCCE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    
    'gGrid_Clear grdControlChar, True
    llRow = grdControlChar.FixedRows
    For ilLoop = 0 To UBound(tgCurrCCE) - 1 Step 1
        If llRow + 1 > grdControlChar.Rows Then
            grdControlChar.AddItem ""
        End If
        grdControlChar.Row = llRow
        grdControlChar.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrCCE(ilLoop).sAutoChar)
        grdControlChar.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrCCE(ilLoop).sDescription)
        If tgCurrCCE(ilLoop).sState = "A" Then
            grdControlChar.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdControlChar.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdControlChar.TextMatrix(llRow, CODEINDEX) = tgCurrCCE(ilLoop).iCode
        grdControlChar.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrCCE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdControlChar.Rows Then
        grdControlChar.AddItem ""
    End If
    grdControlChar.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    If rbcType(0).Value Then
        ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrCCEStamp, "EngrControlChar-mPopulate", tgCurrCCE())
    Else
        ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrCCEStamp, "EngrControlChar-mPopulate", tgCurrCCE())
    End If
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlCCE As CCE
    
    gSetMousePointer grdControlChar, grdControlChar, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdControlChar, grdControlChar, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdControlChar, grdControlChar, vbDefault
        MsgBox "Duplicate names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdControlChar.Redraw = False
    For llRow = grdControlChar.FixedRows To grdControlChar.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmCCE.sAutoChar) <> "" Then
            imCCECode = tmCCE.iCode
            If tmCCE.iCode > 0 Then
                ilRet = gGetRec_CCE_ControlChar(imCCECode, "Control Char-mSave: Get CCE", tlCCE)
                If ilRet Then
                    If mCompare(tmCCE, tlCCE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmCCE.iVersion = tlCCE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmCCE.iCode <= 0 Then
                    ilRet = gPutInsert_CCE_ControlChar(0, tmCCE, "Control Char-mSave: Insert CCE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_CCE_ControlChar(1, tmCCE, "Control Char-mSave: Update CCE")
                    ilRet = gPutDelete_CCE_ControlChar(tmCCE.iCode, "Control Char-mSave: Delete CCE")
                    ilRet = gPutInsert_CCE_ControlChar(1, tmCCE, "Control Char-mSave: Insert CCE")
                End If
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutDelete_CCE_ControlChar(imDeleteCodes(ilLoop), "EngrControlChar- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    grdControlChar.Redraw = True
    sgCurrCCEStamp = ""
    If rbcType(0).Value Then
        ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrCCEStamp, "EngrControlChar-mPopulate", tgCurrCCE())
    Else
        ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrCCEStamp, "EngrControlChar-mPopulate", tgCurrCCE())
    End If
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrControlChar
End Sub

Private Sub cmcDone_Click()
    
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrControlChar
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdControlChar, grdControlChar, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdControlChar, grdControlChar, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdControlChar, grdControlChar, vbDefault
    Unload EngrControlChar
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
        gSetMousePointer grdControlChar, grdControlChar, vbHourglass
        llTopRow = grdControlChar.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdControlChar, grdControlChar, vbDefault
            Exit Sub
        End If
        grdControlChar.Redraw = False
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
        grdControlChar.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdControlChar, grdControlChar, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdControlChar.Col
        Case NAMEINDEX
            If grdControlChar.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdControlChar.text = edcGrid.text
            grdControlChar.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdControlChar.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdControlChar.text = edcGrid.text
            grdControlChar.CellForeColor = vbBlack
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
    gSetFonts EngrControlChar
    gCenterFormModal EngrControlChar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdControlChar.FixedRows) And (lmEnableRow < grdControlChar.Rows) Then
            If (lmEnableCol >= grdControlChar.FixedCols) And (lmEnableCol < grdControlChar.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdControlChar.text = smESCValue
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
    grdControlChar.Height = cmcCancel.Top - grdControlChar.Top - 120    '8 * grdControlChar.RowHeight(0) + 30
    gGrid_IntegralHeight grdControlChar
    gGrid_FillWithRows grdControlChar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrControlChar = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdControlChar, grdControlChar, vbHourglass
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
    If igInitCallInfo = 1 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOCONTROLLIST) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
            imcInsert.Enabled = False
            imcTrash.Enabled = False
        End If
    ElseIf igInitCallInfo = 2 Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSCONTROLLIST) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
            imcInsert.Enabled = False
            imcTrash.Enabled = False
        End If
    Else
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIOCONTROLLIST) = 2) Or (igListStatus(BUSCONTROLLIST) = 2) Then
            cmcDone.Enabled = True
        Else
            cmcDone.Enabled = False
            imcInsert.Enabled = False
            imcTrash.Enabled = False
        End If
    End If
    gSetMousePointer grdControlChar, grdControlChar, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdControlChar, grdControlChar, vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Control Character Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Control Character Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub imcInsert_Click()
    If rbcType(0).Value Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIOLIST) <> 2) Then
            cmcCancel.SetFocus
            Exit Sub
        End If
    Else
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(BUSLIST) <> 2) Then
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    mSetShow
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    igRptIndex = CONTROL_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    If rbcType(0).Value Then
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIOLIST) <> 2) Then
            cmcCancel.SetFocus
            Exit Sub
        End If
    Else
        If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(BUSLIST) <> 2) Then
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    mSetShow
    mDeleteRow
End Sub

Private Sub grdControlChar_Click()
    If grdControlChar.Col >= grdControlChar.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdControlChar_EnterCell()
    mSetShow
End Sub

Private Sub grdControlChar_GotFocus()
    If grdControlChar.Col >= grdControlChar.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdControlChar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdControlChar.TopRow
    grdControlChar.Redraw = False
End Sub

Private Sub grdControlChar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdControlChar.RowHeight(0) Then
        mSortCol grdControlChar.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdControlChar, x, y)
    If Not ilFound Then
        grdControlChar.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdControlChar.Col >= grdControlChar.Cols - 1 Then
        grdControlChar.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdControlChar.TopRow
    DoEvents
    llRow = grdControlChar.Row
    If grdControlChar.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdControlChar.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdControlChar.TextMatrix(llRow, NAMEINDEX) = ""
        grdControlChar.Row = llRow + 1
        grdControlChar.Col = NAMEINDEX
        grdControlChar.Redraw = True
    End If
    grdControlChar.Redraw = True
    mEnableBox
End Sub

Private Sub grdControlChar_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdControlChar.Redraw = False Then
        grdControlChar.Redraw = True
        If lmTopRow < grdControlChar.FixedRows Then
            grdControlChar.TopRow = grdControlChar.FixedRows
        Else
            grdControlChar.TopRow = lmTopRow
        End If
        grdControlChar.Refresh
        grdControlChar.Redraw = False
    End If
    If (imShowGridBox) And (grdControlChar.Row >= grdControlChar.FixedRows) And (grdControlChar.Col >= 0) And (grdControlChar.Col < grdControlChar.Cols - 1) Then
        If grdControlChar.RowIsVisible(grdControlChar.Row) Then
            'edcGrid.Move grdControlChar.Left + grdControlChar.ColPos(grdControlChar.Col) + 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + 30, grdControlChar.ColWidth(grdControlChar.Col) - 30, grdControlChar.RowHeight(grdControlChar.Row) - 30
            pbcArrow.Move grdControlChar.Left - pbcArrow.Width - 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + (grdControlChar.RowHeight(grdControlChar.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
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
    If edcGrid.Visible Or pbcState.Visible Then
        mSetShow
        If grdControlChar.Col = NAMEINDEX Then
            If grdControlChar.Row > grdControlChar.FixedRows Then
                lmTopRow = -1
                grdControlChar.Row = grdControlChar.Row - 1
                If Not grdControlChar.RowIsVisible(grdControlChar.Row) Then
                    grdControlChar.TopRow = grdControlChar.TopRow - 1
                End If
                grdControlChar.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdControlChar.Col = grdControlChar.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdControlChar.TopRow = grdControlChar.FixedRows
        grdControlChar.Col = NAMEINDEX
        grdControlChar.Row = grdControlChar.FixedRows
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
        grdControlChar.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdControlChar.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdControlChar.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdControlChar.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdControlChar.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdControlChar.CellForeColor = vbBlack
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
        If grdControlChar.Col = STATEINDEX Then
            llRow = grdControlChar.Rows
            Do
                llRow = llRow - 1
            Loop While grdControlChar.TextMatrix(llRow, NAMEINDEX) = ""
            llRow = llRow + 1
            If (grdControlChar.Row + 1 < llRow) Then
                lmTopRow = -1
                grdControlChar.Row = grdControlChar.Row + 1
                If Not grdControlChar.RowIsVisible(grdControlChar.Row) Then
                    imIgnoreScroll = True
                    grdControlChar.TopRow = grdControlChar.TopRow + 1
                End If
                grdControlChar.Col = NAMEINDEX
                'grdControlChar.TextMatrix(grdControlChar.Row, CODEINDEX) = 0
                If Trim$(grdControlChar.TextMatrix(grdControlChar.Row, NAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdControlChar.Left - pbcArrow.Width - 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + (grdControlChar.RowHeight(grdControlChar.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdControlChar.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdControlChar.Row + 1 >= grdControlChar.Rows Then
                        grdControlChar.AddItem ""
                    End If
                    grdControlChar.Row = grdControlChar.Row + 1
                    If Not grdControlChar.RowIsVisible(grdControlChar.Row) Then
                        imIgnoreScroll = True
                        grdControlChar.TopRow = grdControlChar.TopRow + 1
                    End If
                    grdControlChar.Col = NAMEINDEX
                    grdControlChar.TextMatrix(grdControlChar.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdControlChar.Left - pbcArrow.Width - 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + (grdControlChar.RowHeight(grdControlChar.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdControlChar.Col = grdControlChar.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdControlChar.TopRow = grdControlChar.FixedRows
        grdControlChar.Col = NAMEINDEX
        grdControlChar.Row = grdControlChar.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdControlChar.TopRow
    llRow = grdControlChar.Row
    slMsg = "Insert above " & Trim$(grdControlChar.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdControlChar.Redraw = False
    grdControlChar.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdControlChar.Row = llRow
    grdControlChar.Redraw = False
    grdControlChar.TopRow = llTRow
    grdControlChar.Redraw = True
    DoEvents
    grdControlChar.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdControlChar.TopRow
    llRow = grdControlChar.Row
    If (Val(grdControlChar.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdControlChar.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdControlChar.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdControlChar.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdControlChar.Redraw = False
    If (Val(grdControlChar.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdControlChar.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdControlChar.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdControlChar.AddItem ""
    grdControlChar.Redraw = False
    grdControlChar.TopRow = llTRow
    grdControlChar.Redraw = True
    DoEvents
    grdControlChar.Col = NAMEINDEX
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

Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrCCE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdControlChar.FixedRows To grdControlChar.Rows - 1 Step 1
            slStr = Trim$(grdControlChar.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdControlChar.Row = llRow
                    Do While Not grdControlChar.RowIsVisible(grdControlChar.Row)
                        imIgnoreScroll = True
                        grdControlChar.TopRow = grdControlChar.TopRow + 1
                    Loop
                    grdControlChar.Col = NAMEINDEX
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
    For llRow = grdControlChar.FixedRows To grdControlChar.Rows - 1 Step 1
        slStr = Trim$(grdControlChar.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdControlChar.Row = llRow
            Do While Not grdControlChar.RowIsVisible(grdControlChar.Row)
                imIgnoreScroll = True
                grdControlChar.TopRow = grdControlChar.TopRow + 1
            Loop
            grdControlChar.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdControlChar.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Function mCompare(tlNew As CCE, tlOld As CCE) As Integer
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


Private Sub mSetFocus()
    Select Case grdControlChar.Col
        Case NAMEINDEX
            edcGrid.Move grdControlChar.Left + grdControlChar.ColPos(grdControlChar.Col) + 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + 15, grdControlChar.ColWidth(grdControlChar.Col) - 30, grdControlChar.RowHeight(grdControlChar.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdControlChar.Left + grdControlChar.ColPos(grdControlChar.Col) + 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + 15, grdControlChar.ColWidth(grdControlChar.Col) - 30, grdControlChar.RowHeight(grdControlChar.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdControlChar.Left + grdControlChar.ColPos(grdControlChar.Col) + 30, grdControlChar.Top + grdControlChar.RowPos(grdControlChar.Row) + 15, grdControlChar.ColWidth(grdControlChar.Col) - 30, grdControlChar.RowHeight(grdControlChar.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub
