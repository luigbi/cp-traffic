VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form EngrAudioName 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrAudioName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin VB.PictureBox pbcCheckConflict 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5910
      ScaleHeight     =   210
      ScaleWidth      =   1125
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   11685
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   45
      Width           =   45
   End
   Begin VB.ListBox lbcCCE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrAudioName.frx":030A
      Left            =   6090
      List            =   "EngrAudioName.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   11370
      Top             =   6720
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   2685
      TabIndex        =   8
      Top             =   2430
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmcDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3630
      Picture         =   "EngrAudioName.frx":030E
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcATE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrAudioName.frx":0408
      Left            =   4170
      List            =   "EngrAudioName.frx":040F
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   1410
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
      TabIndex        =   10
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
      Left            =   90
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   12
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
      Picture         =   "EngrAudioName.frx":041B
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
      Left            =   6945
      TabIndex        =   15
      Top             =   6615
      Width           =   1335
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   10710
      Top             =   6705
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
      Left            =   5175
      TabIndex        =   14
      Top             =   6615
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3390
      TabIndex        =   13
      Top             =   6615
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAudioName 
      Height          =   5880
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   10372
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
      Left            =   10035
      TabIndex        =   17
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2220
      Picture         =   "EngrAudioName.frx":0725
      Top             =   6525
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Audio Name"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2625
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   1335
      Picture         =   "EngrAudioName.frx":0A2F
      Top             =   6525
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8955
      Picture         =   "EngrAudioName.frx":12F9
      Top             =   6525
      Width           =   480
   End
End
Attribute VB_Name = "EngrAudioName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrAudioName - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private smState As String
Private smCheckConflict As String
Private imInChg As Integer
Private imBSMode As Integer
Private ilANECode As Integer
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer

Private smESCValue As String    'Value used if ESC pressed

Private imDoubleClickName As Integer
Private imLbcMouseDown As Integer

Private tmANE As ANE

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
Const CONTROLINDEX = 2
Const AUDIOTYPEINDEX = 3
Const CHECKCONFLICTINDEX = 4
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
    llRow = gGrid_Search(grdAudioName, NAMEINDEX, slStr)
    If llRow >= 0 Then
        mEnableBox
    End If
End Sub

Private Sub cmcSearch_GotFocus()
    mSetShow
End Sub


Private Sub mPopATE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudio-mPopulate Audio Names", tgCurrATE())
    lbcATE.Clear
    For ilLoop = 0 To UBound(tgCurrATE) - 1 Step 1
        lbcATE.AddItem Trim$(tgCurrATE(ilLoop).sName)
        lbcATE.ItemData(lbcATE.NewIndex) = tgCurrATE(ilLoop).iCode
    Next ilLoop
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIONAMELIST) = 2) Then
        lbcATE.AddItem "[New]", 0
        lbcATE.ItemData(lbcATE.NewIndex) = 0
    Else
        lbcATE.AddItem "[View]", 0
        lbcATE.ItemData(lbcATE.NewIndex) = 0
    End If
End Sub
Private Function mNameOk() As Integer
    Dim ilError As Integer
    Dim llRow As Long
    Dim llTestRow As Long
    Dim slStr As String
    Dim slTestStr As String
    
    grdAudioName.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdAudioName.FixedRows To grdAudioName.Rows - 1 Step 1
        slStr = Trim$(grdAudioName.TextMatrix(llRow, NAMEINDEX))
        If (slStr <> "") Then
            For llTestRow = llRow + 1 To grdAudioName.Rows - 1 Step 1
                slTestStr = Trim$(grdAudioName.TextMatrix(llTestRow, NAMEINDEX))
                If StrComp(slStr, slTestStr, vbTextCompare) = 0 Then
                    ilError = True
                    If Val(grdAudioName.TextMatrix(llRow, CODEINDEX)) = 0 Then
                        grdAudioName.Row = llRow
                        grdAudioName.Col = NAMEINDEX
                        grdAudioName.CellForeColor = vbRed
                    Else
                        grdAudioName.Row = llTestRow
                        grdAudioName.Col = NAMEINDEX
                        grdAudioName.CellForeColor = vbRed
                    End If
                End If
            Next llTestRow
        End If
    Next llRow
    grdAudioName.Redraw = True
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
    gGrid_SortByCol grdAudioName, NAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    Dim slStr As String
    Dim ilIndex As Integer
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIONAMELIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdAudioName.Row >= grdAudioName.FixedRows) And (grdAudioName.Row < grdAudioName.Rows) And (grdAudioName.Col >= 0) And (grdAudioName.Col < grdAudioName.Cols - 1) Then
        lmEnableRow = grdAudioName.Row
        lmEnableCol = grdAudioName.Col
        sgReturnCallName = grdAudioName.TextMatrix(lmEnableRow, NAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdAudioName.Left - pbcArrow.Width - 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + (grdAudioName.RowHeight(grdAudioName.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If ((Val(grdAudioName.TextMatrix(lmEnableRow, CODEINDEX)) = 0) Or (grdAudioName.TextMatrix(lmEnableRow, USEDFLAGINDEX) <> "Y")) And (Trim$(grdAudioName.TextMatrix(lmEnableRow, NAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdAudioName.Col
            Case NAMEINDEX  'Call Letters
'                edcGrid.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
                'edcGrid.MaxLength = Len(tmANE.sName)
                edcGrid.MaxLength = gGetAllowedChars("AUDIONAME", Len(tmANE.sName))
                edcGrid.text = grdAudioName.text
'                edcGrid.Visible = True
'                edcGrid.SetFocus
            Case DESCRIPTIONINDEX  'Date
'                edcGrid.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
                edcGrid.MaxLength = Len(tmANE.sDescription)
                edcGrid.text = grdAudioName.text
'                edcGrid.Visible = True
'                edcGrid.SetFocus
            Case CONTROLINDEX
'                edcDropdown.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - cmcDropdown.Width - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
'                cmcDropdown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropdown.Width, edcDropdown.Height
'                lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropdown.Width
'                gSetListBoxHeight lbcCCE, CLng(grdAudioName.Height / 2)
'                If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
'                    lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
'                End If
                slStr = grdAudioName.text
                'ilIndex = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcCCE, slStr)
                If ilIndex >= 0 Then
                    lbcCCE.ListIndex = ilIndex
                    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                Else
                    edcDropdown.text = ""
                    If lbcCCE.ListCount <= 0 Then
                        lbcCCE.ListIndex = -1
                        edcDropdown.text = ""
                    Else
                        lbcCCE.ListIndex = 1
                        edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                    End If
                End If
'                edcDropdown.Visible = True
'                cmcDropdown.Visible = True
'                lbcCCE.Visible = True
'                edcDropdown.SetFocus
            Case AUDIOTYPEINDEX
'                edcDropdown.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - cmcDropdown.Width - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
'                cmcDropdown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropdown.Width, edcDropdown.Height
'                lbcATE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropdown.Width
'                gSetListBoxHeight lbcATE, CLng(grdAudioName.Height / 2)
'                If lbcATE.Top + lbcATE.Height > cmcCancel.Top Then
'                    lbcATE.Top = edcDropdown.Top - lbcATE.Height
'                End If
                slStr = grdAudioName.text
                'ilIndex = SendMessageByString(lbcATE.hwnd, LB_FINDSTRING, -1, slStr)
                ilIndex = gListBoxFind(lbcATE, slStr)
                If ilIndex >= 0 Then
                    lbcATE.ListIndex = ilIndex
                    edcDropdown.text = lbcATE.List(lbcATE.ListIndex)
                Else
                    edcDropdown.text = ""
                    If lbcATE.ListCount <= 0 Then
                        lbcATE.ListIndex = -1
                        edcDropdown.text = ""
                    ElseIf lbcATE.ListCount <= 1 Then
                        lbcATE.ListIndex = 0
                        edcDropdown.text = lbcATE.List(lbcATE.ListIndex)
                    Else
                        lbcATE.ListIndex = 1
                        edcDropdown.text = lbcATE.List(lbcATE.ListIndex)
                    End If
                End If
'                edcDropdown.Visible = True
'                cmcDropdown.Visible = True
'                lbcATE.Visible = True
'                edcDropdown.SetFocus
            Case CHECKCONFLICTINDEX
                smCheckConflict = grdAudioName.text
                If (Trim$(smCheckConflict) = "") Or (smCheckConflict = "Missing") Then
                    smCheckConflict = "Yes"
                End If
            Case STATEINDEX
'                pbcState.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
                smState = grdAudioName.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
'                pbcState.Visible = True
'                pbcState.SetFocus
        End Select
        smESCValue = grdAudioName.text
        mSetFocus
    End If
End Sub
Private Sub mSetShow()
    Dim llRow As Long
    Dim slStr As String
    
    tmcClick.Enabled = False
    If (lmEnableRow >= grdAudioName.FixedRows) And (lmEnableRow < grdAudioName.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = grdAudioName.TextMatrix(lmEnableRow, lmEnableCol)
        Select Case lmEnableCol
            Case NAMEINDEX
            Case DESCRIPTIONINDEX
            Case AUDIOTYPEINDEX
                'Remove illegal values
                'llRow = SendMessageByString(lbcATE.hwnd, LB_FINDSTRING, -1, slStr)
                llRow = gListBoxFind(lbcATE, slStr)
                If (llRow <= 0) Then
                    grdAudioName.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case CHECKCONFLICTINDEX
                grdAudioName.TextMatrix(lmEnableRow, lmEnableCol) = smCheckConflict
                If (Trim$(grdAudioName.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdAudioName.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdAudioName.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdAudioName.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdAudioName.TextMatrix(lmEnableRow, NAMEINDEX)
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    lbcATE.Visible = False
    lbcCCE.Visible = False
    cmcDropDown.Visible = False
    edcDropdown.Visible = False
    pbcArrow.Visible = False
    edcGrid.Visible = False
    pbcCheckConflict.Visible = False
    pbcState.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub
Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llLbc As Long
    Dim llRow As Long
    
    grdAudioName.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdAudioName.FixedRows To grdAudioName.Rows - 1 Step 1
        slStr = Trim$(grdAudioName.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdAudioName.TextMatrix(llRow, DESCRIPTIONINDEX)
            If slStr <> "" Then
                ilError = True
                grdAudioName.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdAudioName.Row = llRow
                grdAudioName.Col = NAMEINDEX
                grdAudioName.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = grdAudioName.TextMatrix(llRow, STATEINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdAudioName.TextMatrix(llRow, STATEINDEX) = "Missing"
                    grdAudioName.Row = llRow
                    grdAudioName.Col = STATEINDEX
                    grdAudioName.CellForeColor = vbRed
                End If
                slStr = grdAudioName.TextMatrix(llRow, CHECKCONFLICTINDEX)
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdAudioName.TextMatrix(llRow, CHECKCONFLICTINDEX) = "Missing"
                    grdAudioName.Row = llRow
                    grdAudioName.Col = CHECKCONFLICTINDEX
                    grdAudioName.CellForeColor = vbRed
                End If
                slStr = Trim$(grdAudioName.TextMatrix(llRow, AUDIOTYPEINDEX))
                'llLbc = SendMessageByString(lbcATE.hwnd, LB_FINDSTRING, -1, slStr)
                llLbc = gListBoxFind(lbcATE, slStr)
                If (llLbc <= 0) Then
                    ilError = True
                    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                        grdAudioName.TextMatrix(llRow, AUDIOTYPEINDEX) = "Missing"
                    End If
                    grdAudioName.Row = llRow
                    grdAudioName.Col = AUDIOTYPEINDEX
                    grdAudioName.CellForeColor = vbRed
                End If
            End If
        End If
    Next llRow
    grdAudioName.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdAudioName
    mGridColumnWidth
    'Set Titles
    grdAudioName.TextMatrix(0, NAMEINDEX) = "Name"
    grdAudioName.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdAudioName.TextMatrix(0, CONTROLINDEX) = "Control"
    grdAudioName.TextMatrix(0, AUDIOTYPEINDEX) = "Audio Type"
    grdAudioName.TextMatrix(0, CHECKCONFLICTINDEX) = "Check Conflict"
    grdAudioName.TextMatrix(0, STATEINDEX) = "State"
    grdAudioName.Row = 1
    For ilCol = 0 To grdAudioName.Cols - 1 Step 1
        grdAudioName.Col = ilCol
        grdAudioName.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdAudioName.Height = cmcCancel.Top - grdAudioName.Top - 120    '8 * grdAudioName.RowHeight(0) + 30
    gGrid_IntegralHeight grdAudioName
    gGrid_Clear grdAudioName, True
    grdAudioName.Row = grdAudioName.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdAudioName.Width = EngrAudioName.Width - 2 * grdAudioName.Left
    grdAudioName.ColWidth(CODEINDEX) = 0
    grdAudioName.ColWidth(USEDFLAGINDEX) = 0
    grdAudioName.ColWidth(NAMEINDEX) = grdAudioName.Width / 6
    If (tgUsedSumEPE.sAudioControl <> "Y") And (tgUsedSumEPE.sBkupAudioControl <> "Y") And (tgUsedSumEPE.sProtAudioControl <> "Y") Then
        grdAudioName.ColWidth(CONTROLINDEX) = 0
    Else
        grdAudioName.ColWidth(CONTROLINDEX) = grdAudioName.Width / 12
    End If
    grdAudioName.ColWidth(AUDIOTYPEINDEX) = grdAudioName.Width / 5
    grdAudioName.ColWidth(CHECKCONFLICTINDEX) = grdAudioName.Width / 10
    grdAudioName.ColWidth(STATEINDEX) = grdAudioName.Width / 15
    grdAudioName.ColWidth(DESCRIPTIONINDEX) = grdAudioName.Width - GRIDSCROLLWIDTH
    For ilCol = NAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdAudioName.ColWidth(DESCRIPTIONINDEX) > grdAudioName.ColWidth(ilCol) Then
                grdAudioName.ColWidth(DESCRIPTIONINDEX) = grdAudioName.ColWidth(DESCRIPTIONINDEX) - grdAudioName.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdAudioName, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim ilATE As Integer
    Dim ilCCE As Integer
    Dim slStr As String
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdAudioName.TextMatrix(llRow, CODEINDEX)) = "" Then
        grdAudioName.TextMatrix(llRow, CODEINDEX) = "0"
    End If
    tmANE.iCode = Val(grdAudioName.TextMatrix(llRow, CODEINDEX))
    slStr = Trim$(grdAudioName.TextMatrix(llRow, NAMEINDEX))
    If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
        tmANE.sName = ""
    Else
        tmANE.sName = slStr
    End If
    tmANE.sDescription = grdAudioName.TextMatrix(llRow, DESCRIPTIONINDEX)
    tmANE.iCceCode = 0
    slStr = Trim$(grdAudioName.TextMatrix(llRow, CONTROLINDEX))
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If StrComp(Trim$(tgCurrAudioCCE(ilCCE).sAutoChar), slStr, vbTextCompare) = 0 Then
            tmANE.iCceCode = tgCurrAudioCCE(ilCCE).iCode
            Exit For
        End If
    Next ilCCE
    tmANE.iAteCode = 0
    slStr = Trim$(grdAudioName.TextMatrix(llRow, AUDIOTYPEINDEX))
    For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
        If StrComp(Trim$(tgCurrATE(ilATE).sName), slStr, vbTextCompare) = 0 Then
            tmANE.iAteCode = tgCurrATE(ilATE).iCode
            Exit For
        End If
    Next ilATE
    If grdAudioName.TextMatrix(llRow, CHECKCONFLICTINDEX) = "No" Then
        tmANE.sCheckConflicts = "N"
    Else
        tmANE.sCheckConflicts = "Y"
    End If
    If grdAudioName.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tmANE.sState = "D"
    Else
        tmANE.sState = "A"
    End If
    If tmANE.iCode <= 0 Then
        tmANE.sUsedFlag = "N"
    Else
        tmANE.sUsedFlag = grdAudioName.TextMatrix(llRow, USEDFLAGINDEX)
    End If
    tmANE.iVersion = 0
    tmANE.iOrigAneCode = tmANE.iCode
    tmANE.sCurrent = "Y"
    'tmANE.sEnteredDate = smNowDate
    'tmANE.sEnteredTime = smNowTime
    tmANE.sEnteredDate = Format(Now, sgShowDateForm)
    tmANE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmANE.iUieCode = tgUIE.iCode
    tmANE.sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilATE As Integer
    Dim ilCCE As Integer
    
    'gGrid_Clear grdAudioName, True
    llRow = grdAudioName.FixedRows
    For ilLoop = 0 To UBound(tgCurrANE) - 1 Step 1
        If llRow + 1 > grdAudioName.Rows Then
            grdAudioName.AddItem ""
        End If
        grdAudioName.Row = llRow
        grdAudioName.TextMatrix(llRow, NAMEINDEX) = Trim$(tgCurrANE(ilLoop).sName)
        grdAudioName.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgCurrANE(ilLoop).sDescription)
        For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If tgCurrANE(ilLoop).iCceCode = tgCurrAudioCCE(ilCCE).iCode Then
                grdAudioName.TextMatrix(llRow, CONTROLINDEX) = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
                Exit For
            End If
        Next ilCCE
        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
            If tgCurrANE(ilLoop).iAteCode = tgCurrATE(ilATE).iCode Then
                grdAudioName.TextMatrix(llRow, AUDIOTYPEINDEX) = Trim$(tgCurrATE(ilATE).sName)
                Exit For
            End If
        Next ilATE
        If tgCurrANE(ilLoop).sCheckConflicts = "N" Then
            grdAudioName.TextMatrix(llRow, CHECKCONFLICTINDEX) = "No"
        Else
            grdAudioName.TextMatrix(llRow, CHECKCONFLICTINDEX) = "Yes"
        End If
        If tgCurrANE(ilLoop).sState = "A" Then
            grdAudioName.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdAudioName.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdAudioName.TextMatrix(llRow, CODEINDEX) = tgCurrANE(ilLoop).iCode
        grdAudioName.TextMatrix(llRow, USEDFLAGINDEX) = tgCurrANE(ilLoop).sUsedFlag
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdAudioName.Rows Then
        grdAudioName.AddItem ""
    End If
    grdAudioName.Redraw = True
End Sub
Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    
    
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrAudioName-mPopulate", tgCurrANE())
    
    
End Sub
Private Function mSave() As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim ilSave As Integer
    Dim tlANE As ANE
    
    gSetMousePointer grdAudioName, grdAudioName, vbHourglass
    If Not mCheckFields(True) Then
        gSetMousePointer grdAudioName, grdAudioName, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    If Not mNameOk() Then
        gSetMousePointer grdAudioName, grdAudioName, vbDefault
        MsgBox "Duplicated names not allowed", vbCritical + vbOKOnly, "Save not Completed"
        mSave = False
        Exit Function
    End If
    grdAudioName.Redraw = False
    For llRow = grdAudioName.FixedRows To grdAudioName.Rows - 1 Step 1
        mMoveCtrlsToRec llRow
        If Trim$(tmANE.sName) <> "" Then
            ilANECode = tmANE.iCode
            If tmANE.iCode > 0 Then
                ilRet = gGetRec_ANE_AudioName(ilANECode, "Audio Name-mSave: Get ANE", tlANE)
                If ilRet Then
                    If mCompare(tmANE, tlANE) Then
                        ilSave = False
                    Else
                        ilSave = True
                        tmANE.iVersion = tlANE.iVersion + 1
                    End If
                Else
                    ilSave = False
                End If
            Else
                ilSave = True
            End If
            If ilSave Then
                If tmANE.iCode <= 0 Then
                    ilRet = gPutInsert_ANE_AudioName(0, tmANE, "Audio Name-mSave: Insert ANE")
                Else
                    '7/12/11: History no longer retained
                    'ilRet = gPutUpdate_ANE_AudioName(1, tmANE, "Audio Name-mSave: Update ANE")
                    ilRet = gPutDelete_ANE_AudioName(tmANE.iCode, "Audio Name-mSave: Delete ANE")
                    ilRet = gPutInsert_ANE_AudioName(1, tmANE, "Audio Name-mSave: Insert ANE")
                End If
                ilRet = gPutUpdate_ATE_UsedFlag(tmANE.iAteCode, tgCurrATE())
                ilRet = gPutUpdate_CCE_UsedFlag(tmANE.iCceCode, tgCurrAudioCCE())
            End If
        End If
    Next llRow
    For ilLoop = LBound(imDeleteCodes) To UBound(imDeleteCodes) - 1 Step 1
        ilRet = gPutReplace_ASE_BkupANECode(imDeleteCodes(ilLoop), 0, "EngrAudioBkup- Zero")
        ilRet = gPutReplace_ASE_ProtANECode(imDeleteCodes(ilLoop), 0, "EngrAudioProt- Zero")
        ilRet = gPutDelete_ASE_AudioSource(imDeleteCodes(ilLoop), "EngrAudioSource- Delete")
        ilRet = gPutDelete_ANE_AudioName(imDeleteCodes(ilLoop), "EngrAudioName- Delete")
    Next ilLoop
    ReDim imDeleteCodes(0 To 0) As Integer
    
    grdAudioName.Redraw = True
    sgCurrANEStamp = ""
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrAudioName-mPopulate", tgCurrANE())
    imFieldChgd = False
    mSetCommands
    mSave = True
End Function
Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrAudioName
End Sub

Private Sub cmcDone_Click()
    
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrAudioName
        Exit Sub
    End If
    If MsgBox("Save all changes?", vbYesNo) = vbYes Then
        gSetMousePointer grdAudioName, grdAudioName, vbHourglass
        ilRet = mSave()
        gSetMousePointer grdAudioName, grdAudioName, vbDefault
        If Not ilRet Then
            Exit Sub
        End If
    End If
    
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdAudioName, grdAudioName, vbDefault
    Unload EngrAudioName
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdAudioName.Col
        Case CONTROLINDEX
            lbcCCE.Visible = Not lbcCCE.Visible
        Case AUDIOTYPEINDEX
            lbcATE.Visible = Not lbcATE.Visible
    End Select
End Sub

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim slName As String
    Dim llTopRow As Long
    
    If imFieldChgd = True Then
        gSetMousePointer grdAudioName, grdAudioName, vbHourglass
        llTopRow = grdAudioName.TopRow
        ilRet = mSave()
        If Not ilRet Then
            gSetMousePointer grdAudioName, grdAudioName, vbDefault
            Exit Sub
        End If
        grdAudioName.Redraw = False
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
        grdAudioName.TopRow = llTopRow
        lmEnableRow = -1
        imFieldChgd = False
        mSetCommands
        gSetMousePointer grdAudioName, grdAudioName, vbDefault
    End If
End Sub

Private Sub cmcSave_GotFocus()
    mSetShow
End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As String
    
    slStr = edcDropdown.text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdAudioName.Col
        Case CONTROLINDEX
            'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
            llRow = gListBoxFind(lbcCCE, slStr)
            If llRow >= 0 Then
                lbcCCE.ListIndex = llRow
                edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case AUDIOTYPEINDEX
            'llRow = SendMessageByString(lbcATE.hwnd, LB_FINDSTRING, -1, slStr)
            llRow = gListBoxFind(lbcATE, slStr)
            If llRow >= 0 Then
                lbcATE.ListIndex = llRow
                edcDropdown.text = lbcATE.List(lbcATE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
    End Select
    If StrComp(grdAudioName.text, edcDropdown.text, vbTextCompare) <> 0 Then
        imFieldChgd = True
    End If
    If StrComp(Trim$(edcDropdown.text), "[None]", vbTextCompare) <> 0 Then
        grdAudioName.text = edcDropdown.text
    Else
        grdAudioName.text = ""
    End If
    grdAudioName.CellForeColor = vbBlack
    mSetCommands
End Sub

Private Sub edcDropdown_DblClick()
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub

Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If edcDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdAudioName.Col
            Case CONTROLINDEX
                gProcessArrowKey Shift, KeyCode, lbcCCE, True
            Case AUDIOTYPEINDEX
                gProcessArrowKey Shift, KeyCode, lbcATE, True
        End Select
        tmcClick.Enabled = False
    End If
End Sub

Private Sub edcDropdown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilRet As Integer
    
    If imDoubleClickName Then
        ilRet = mBranch()
    End If
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdAudioName.Col
        Case NAMEINDEX
            If grdAudioName.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdAudioName.text = edcGrid.text
            grdAudioName.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdAudioName.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdAudioName.text = edcGrid.text
            grdAudioName.CellForeColor = vbBlack
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
    gSetFonts EngrAudioName
    gCenterFormModal EngrAudioName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdAudioName.FixedRows) And (lmEnableRow < grdAudioName.Rows) Then
            If (lmEnableCol >= grdAudioName.FixedCols) And (lmEnableCol < grdAudioName.Cols) Then
                If lmEnableCol = CHECKCONFLICTINDEX Then
                    smCheckConflict = smESCValue
                ElseIf lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdAudioName.text = smESCValue
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
    grdAudioName.Height = cmcCancel.Top - grdAudioName.Top - 120    '8 * grdAudioName.RowHeight(0) + 30
    gGrid_IntegralHeight grdAudioName
    gGrid_FillWithRows grdAudioName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrAudioName = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdAudioName, grdAudioName, vbHourglass
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
    mPopATE
    mPopCCE
    mMoveRecToCtrls
    mSortCol 0
    imInChg = False
    imFieldChgd = False
    imLbcMouseDown = False
    imDoubleClickName = False
    mSetCommands
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIONAMELIST) = 2) Then
        cmcDone.Enabled = True
    Else
        cmcDone.Enabled = False
        imcInsert.Enabled = False
        imcTrash.Enabled = False
    End If
    gSetMousePointer grdAudioName, grdAudioName, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdAudioName, grdAudioName, vbDefault
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

Private Sub grdAudioName_DblClick()
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(AUDIONAMELIST) <> 2) Then
        Select Case grdAudioName.Col
            Case AUDIOTYPEINDEX
                igInitCallInfo = 1
                sgInitCallName = grdAudioName.TextMatrix(grdAudioName.Row, grdAudioName.Col)
                EngrAudioType.Show vbModal
                cmcCancel.SetFocus
        End Select
    End If
End Sub

Private Sub imcInsert_Click()
    mSetShow
    mInsertRow
End Sub

Private Sub imcPrint_Click()
    igRptIndex = AUDIONAME_RPT
    igRptSource = vbModal
    EngrUserRpt.Show vbModal
End Sub

Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub lbcATE_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcATE.List(lbcATE.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        'lbcATE.Visible = False
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcATE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcATE.Visible = False
End Sub

Private Sub lbcATE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcATE, y)
    If (llRow < lbcATE.ListCount) And (lbcATE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcATE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrATE) - 1 Step 1
            If ilCode = tgCurrATE(ilLoop).iCode Then
                lbcATE.ToolTipText = Trim$(tgCurrATE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub lbcCCE_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        'lbcCCE.Visible = False
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcCCE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcCCE.Visible = False
End Sub

Private Sub lbcCCE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcCCE, y)
    If (llRow < lbcCCE.ListCount) And (lbcCCE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcCCE.ItemData(llRow)
        For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
            If ilCode = tgCurrAudioCCE(ilLoop).iCode Then
                lbcCCE.ToolTipText = Trim$(tgCurrAudioCCE(ilLoop).sDescription)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Sub grdAudioName_Click()
    If grdAudioName.Col >= grdAudioName.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAudioName_EnterCell()
    mSetShow
End Sub

Private Sub grdAudioName_GotFocus()
    If grdAudioName.Col >= grdAudioName.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdAudioName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdAudioName.TopRow
    grdAudioName.Redraw = False
End Sub

Private Sub grdAudioName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdAudioName.RowHeight(0) Then
        mSortCol grdAudioName.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdAudioName, x, y)
    If Not ilFound Then
        grdAudioName.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdAudioName.Col >= grdAudioName.Cols - 1 Then
        grdAudioName.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdAudioName.TopRow
    DoEvents
    llRow = grdAudioName.Row
    If grdAudioName.TextMatrix(llRow, NAMEINDEX) = "" Then
        grdAudioName.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdAudioName.TextMatrix(llRow, NAMEINDEX) = ""
        grdAudioName.Row = llRow + 1
        grdAudioName.Col = NAMEINDEX
        grdAudioName.Redraw = True
    End If
    grdAudioName.Redraw = True
    mEnableBox
End Sub

Private Sub grdAudioName_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdAudioName.Redraw = False Then
        grdAudioName.Redraw = True
        If lmTopRow < grdAudioName.FixedRows Then
            grdAudioName.TopRow = grdAudioName.FixedRows
        Else
            grdAudioName.TopRow = lmTopRow
        End If
        grdAudioName.Refresh
        grdAudioName.Redraw = False
    End If
    If (imShowGridBox) And (grdAudioName.Row >= grdAudioName.FixedRows) And (grdAudioName.Col >= 0) And (grdAudioName.Col < grdAudioName.Cols - 1) Then
        If grdAudioName.RowIsVisible(grdAudioName.Row) Then
            'edcGrid.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 30, grdAudioName.ColWidth(grdAudioName.Col) - 30, grdAudioName.RowHeight(grdAudioName.Row) - 30
            pbcArrow.Move grdAudioName.Left - pbcArrow.Width - 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + (grdAudioName.RowHeight(grdAudioName.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            lbcATE.Visible = False
            lbcCCE.Visible = False
            cmcDropDown.Visible = False
            edcDropdown.Visible = False
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

Private Sub pbcCheckConflict_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("Y") Or (KeyAscii = Asc("y")) Then
        If smCheckConflict <> "Yes" Then
            imFieldChgd = True
        End If
        smCheckConflict = "Yes"
        pbcCheckConflict_Paint
        grdAudioName.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
        If smCheckConflict <> "No" Then
            imFieldChgd = True
        End If
        smCheckConflict = "No"
        pbcCheckConflict_Paint
        grdAudioName.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smCheckConflict = "Yes" Then
            imFieldChgd = True
            smCheckConflict = "No"
            pbcCheckConflict_Paint
            grdAudioName.CellForeColor = vbBlack
        ElseIf smCheckConflict = "No" Then
            imFieldChgd = True
            smCheckConflict = "Yes"
            pbcCheckConflict_Paint
            grdAudioName.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcCheckConflict_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smCheckConflict = "Yes" Then
        imFieldChgd = True
        smCheckConflict = "No"
        pbcCheckConflict_Paint
        grdAudioName.CellForeColor = vbBlack
    ElseIf smCheckConflict = "No" Then
        imFieldChgd = True
        smCheckConflict = "Yes"
        pbcCheckConflict_Paint
        grdAudioName.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub pbcCheckConflict_Paint()
    pbcCheckConflict.Cls
    pbcCheckConflict.CurrentX = 30  'fgBoxInsetX
    pbcCheckConflict.CurrentY = 0 'fgBoxInsetY
    pbcCheckConflict.Print smCheckConflict
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow
    lmEnableRow = -1
    lmEnableCol = -1
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer
    
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If edcGrid.Visible Or edcDropdown.Visible Or pbcCheckConflict.Visible Or pbcState.Visible Then
        If Not mBranch() Then
            mEnableBox
            Exit Sub
        End If
        mSetShow
        Do
            ilPrev = False
            If grdAudioName.Col = NAMEINDEX Then
                If grdAudioName.Row > grdAudioName.FixedRows Then
                    lmTopRow = -1
                    grdAudioName.Row = grdAudioName.Row - 1
                    If Not grdAudioName.RowIsVisible(grdAudioName.Row) Then
                        grdAudioName.TopRow = grdAudioName.TopRow - 1
                    End If
                    grdAudioName.Col = STATEINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdAudioName.Col = grdAudioName.Col - 1
                If mColOk(grdAudioName.Row, grdAudioName.Col) Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdAudioName.TopRow = grdAudioName.FixedRows
        grdAudioName.Col = NAMEINDEX
        grdAudioName.Row = grdAudioName.FixedRows
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
        grdAudioName.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdAudioName.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdAudioName.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdAudioName.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdAudioName.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdAudioName.CellForeColor = vbBlack
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
    Dim ilNext As Integer
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If edcGrid.Visible Or edcDropdown.Visible Or pbcCheckConflict.Visible Or pbcState.Visible Then
        If Not mBranch() Then
            mEnableBox
            Exit Sub
        End If
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdAudioName.Col = STATEINDEX Then
                llRow = grdAudioName.Rows
                Do
                    llRow = llRow - 1
                Loop While grdAudioName.TextMatrix(llRow, NAMEINDEX) = ""
                llRow = llRow + 1
                If (grdAudioName.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdAudioName.Row = grdAudioName.Row + 1
                    If Not grdAudioName.RowIsVisible(grdAudioName.Row) Then
                        imIgnoreScroll = True
                        grdAudioName.TopRow = grdAudioName.TopRow + 1
                    End If
                    grdAudioName.Col = NAMEINDEX
                    'grdAudioName.TextMatrix(grdAudioName.Row, CODEINDEX) = 0
                    If Trim$(grdAudioName.TextMatrix(grdAudioName.Row, NAMEINDEX)) <> "" Then
                        mEnableBox
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdAudioName.Left - pbcArrow.Width - 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + (grdAudioName.RowHeight(grdAudioName.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdAudioName.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdAudioName.Row + 1 >= grdAudioName.Rows Then
                            grdAudioName.AddItem ""
                        End If
                        grdAudioName.Row = grdAudioName.Row + 1
                        If Not grdAudioName.RowIsVisible(grdAudioName.Row) Then
                            imIgnoreScroll = True
                            grdAudioName.TopRow = grdAudioName.TopRow + 1
                        End If
                        grdAudioName.Col = NAMEINDEX
                        grdAudioName.TextMatrix(grdAudioName.Row, CODEINDEX) = 0
                        grdAudioName.TextMatrix(grdAudioName.Row, USEDFLAGINDEX) = "N"
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdAudioName.Left - pbcArrow.Width - 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + (grdAudioName.RowHeight(grdAudioName.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdAudioName.Col = grdAudioName.Col + 1
                If mColOk(grdAudioName.Row, grdAudioName.Col) Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdAudioName.TopRow = grdAudioName.FixedRows
        grdAudioName.Col = NAMEINDEX
        grdAudioName.Row = grdAudioName.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdAudioName.TopRow
    llRow = grdAudioName.Row
    slMsg = "Insert above " & Trim$(grdAudioName.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdAudioName.Redraw = False
    grdAudioName.AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdAudioName.Row = llRow
    grdAudioName.Redraw = False
    grdAudioName.TopRow = llTRow
    grdAudioName.Redraw = True
    DoEvents
    grdAudioName.Col = NAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdAudioName.TopRow
    llRow = grdAudioName.Row
    If (Val(grdAudioName.TextMatrix(llRow, CODEINDEX)) <> 0) And (grdAudioName.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
        MsgBox Trim$(grdAudioName.TextMatrix(llRow, NAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
        mDeleteRow = False
        Exit Function
    End If
    slMsg = "Delete " & Trim$(grdAudioName.TextMatrix(llRow, NAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdAudioName.Redraw = False
    If (Val(grdAudioName.TextMatrix(llRow, CODEINDEX)) <> 0) Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdAudioName.TextMatrix(llRow, CODEINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdAudioName.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdAudioName.AddItem ""
    grdAudioName.Redraw = False
    grdAudioName.TopRow = llTRow
    grdAudioName.Redraw = True
    DoEvents
    grdAudioName.Col = NAMEINDEX
    mEnableBox
    mSetCommands
    mDeleteRow = True
End Function

Private Function mBranch() As Integer
    Dim llRow As Long
    Dim slStr As String
    
    mBranch = True
    If (lmEnableRow >= grdAudioName.FixedRows) And (lmEnableRow < grdAudioName.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = Trim$(grdAudioName.TextMatrix(lmEnableRow, lmEnableCol))
        If (slStr <> "") And (StrComp(slStr, "[None]", vbTextCompare) <> 0) Then
            Select Case lmEnableCol
                Case NAMEINDEX
                Case CONTROLINDEX
                    'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcCCE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrControlChar.Show vbModal
                        sgCurrAudioCCEStamp = ""
                        mPopCCE
                        lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcCCE, CLng(grdAudioName.Height / 2)
                        If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
                            lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcCCE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcCCE, slStr)
                            If llRow > 0 Then
                                lbcCCE.ListIndex = llRow
                                edcDropdown.text = lbcCCE.List(lbcCCE.ListIndex)
                                edcDropdown.SelStart = 0
                                edcDropdown.SelLength = Len(edcDropdown.text)
                            Else
                                mBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mBranch = False
                        End If
                    End If
                Case AUDIOTYPEINDEX
                    'llRow = SendMessageByString(lbcATE.hwnd, LB_FINDSTRING, -1, slStr)
                    llRow = gListBoxFind(lbcATE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrAudioType.Show vbModal
                        sgCurrATEStamp = ""
                        mPopATE
                        lbcATE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcATE, CLng(grdAudioName.Height / 2)
                        If lbcATE.Top + lbcATE.Height > cmcCancel.Top Then
                            lbcATE.Top = edcDropdown.Top - lbcATE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcATE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcATE, slStr)
                            If llRow > 0 Then
                                lbcATE.ListIndex = llRow
                                edcDropdown.text = lbcATE.List(lbcATE.ListIndex)
                                edcDropdown.SelStart = 0
                                edcDropdown.SelLength = Len(edcDropdown.text)
                            Else
                                mBranch = False
                            End If
                        ElseIf igReturnCallStatus = CALLCANCELLED Then
                            mBranch = False
                        ElseIf igReturnCallStatus = CALLTERMINATED Then
                            mBranch = False
                        End If
                    End If
                Case DESCRIPTIONINDEX
                Case STATEINDEX
            End Select
        End If
    End If
    imDoubleClickName = False
End Function

Private Sub mFindMatch(ilCreateNew As Integer)
    Dim llRow As Long
    Dim slStr As String
    
    If igInitCallInfo = 0 Then
        If UBound(tgCurrANE) > 0 Then
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
        For llRow = grdAudioName.FixedRows To grdAudioName.Rows - 1 Step 1
            slStr = Trim$(grdAudioName.TextMatrix(llRow, NAMEINDEX))
            If (slStr <> "") Then
                If StrComp(slStr, sgInitCallName, vbTextCompare) = 0 Then
                    grdAudioName.Row = llRow
                    Do While Not grdAudioName.RowIsVisible(grdAudioName.Row)
                        imIgnoreScroll = True
                        grdAudioName.TopRow = grdAudioName.TopRow + 1
                    Loop
                    grdAudioName.Col = NAMEINDEX
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
    For llRow = grdAudioName.FixedRows To grdAudioName.Rows - 1 Step 1
        slStr = Trim$(grdAudioName.TextMatrix(llRow, NAMEINDEX))
        If (slStr = "") Then
            grdAudioName.Row = llRow
            Do While Not grdAudioName.RowIsVisible(grdAudioName.Row)
                imIgnoreScroll = True
                grdAudioName.TopRow = grdAudioName.TopRow + 1
            Loop
            grdAudioName.Col = NAMEINDEX
            If StrComp(sgInitCallName, "[New]", vbTextCompare) <> 0 Then
                grdAudioName.text = sgInitCallName
            End If
            mEnableBox
            Exit Sub
        End If
    Next llRow
    
End Sub

Private Function mCompare(tlNew As ANE, tlOld As ANE) As Integer
    If StrComp(tlNew.sName, tlOld.sName, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sDescription, tlOld.sDescription, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sCheckConflicts, tlOld.sCheckConflicts, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If StrComp(tlNew.sState, tlOld.sState, vbTextCompare) <> 0 Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iCceCode <> tlOld.iCceCode) Then
        mCompare = False
        Exit Function
    End If
    If (tlNew.iAteCode <> tlOld.iAteCode) Then
        mCompare = False
        Exit Function
    End If
    mCompare = True
End Function

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case grdAudioName.Col
        Case CONTROLINDEX
            lbcCCE.Visible = False
        Case AUDIOTYPEINDEX
            lbcATE.Visible = False
    End Select
End Sub

Private Sub mPopCCE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrControlChar-mPopulate Audio Controls", tgCurrAudioCCE())
    lbcCCE.Clear
    For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        lbcCCE.AddItem Trim$(tgCurrAudioCCE(ilLoop).sAutoChar)
        lbcCCE.ItemData(lbcCCE.NewIndex) = tgCurrAudioCCE(ilLoop).iCode
    Next ilLoop
    lbcCCE.AddItem "[None]", 0
    lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(AUDIONAMELIST) = 2) Then
        lbcCCE.AddItem "[New]", 0
        lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    Else
        lbcCCE.AddItem "[View]", 0
        lbcCCE.ItemData(lbcCCE.NewIndex) = 0
    End If
End Sub

Private Sub mSetFocus()
    Select Case grdAudioName.Col
        Case NAMEINDEX  'Call Letters
            edcGrid.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case CONTROLINDEX
            edcDropdown.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - cmcDropDown.Width - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcCCE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcCCE, CLng(grdAudioName.Height / 2)
            If lbcCCE.Top + lbcCCE.Height > cmcCancel.Top Then
                lbcCCE.Top = edcDropdown.Top - lbcCCE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcCCE.Visible = True
            edcDropdown.SetFocus
        Case AUDIOTYPEINDEX
            edcDropdown.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - cmcDropDown.Width - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcATE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcATE, CLng(grdAudioName.Height / 2)
            If lbcATE.Top + lbcATE.Height > cmcCancel.Top Then
                lbcATE.Top = edcDropdown.Top - lbcATE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcATE.Visible = True
            edcDropdown.SetFocus
        Case CHECKCONFLICTINDEX
            pbcCheckConflict.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
            pbcCheckConflict.Visible = True
            pbcCheckConflict.SetFocus
        Case STATEINDEX
            pbcState.Move grdAudioName.Left + grdAudioName.ColPos(grdAudioName.Col) + 30, grdAudioName.Top + grdAudioName.RowPos(grdAudioName.Row) + 15, grdAudioName.ColWidth(grdAudioName.Col) - 30, grdAudioName.RowHeight(grdAudioName.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub

Private Function mColOk(llRow As Long, llCol As Long) As Integer
    
    mColOk = True
    If grdAudioName.ColWidth(grdAudioName.Col) <= 0 Then
        mColOk = False
        Exit Function
    End If
End Function

