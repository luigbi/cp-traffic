VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCategoryMatching 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9360
   ControlBox      =   0   'False
   Icon            =   "AffCategoryMatching.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   540
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   15
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   4665
      Width           =   60
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
      Left            =   2280
      Picture         =   "AffCategoryMatching.frx":08CA
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1335
      TabIndex        =   5
      Top             =   2310
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcToName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffCategoryMatching.frx":09C4
      Left            =   990
      List            =   "AffCategoryMatching.frx":09C6
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Timer tmcFillGrid 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8865
      Top             =   4860
   End
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7605
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   5055
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Continue"
      Height          =   315
      Left            =   3285
      TabIndex        =   0
      Top             =   5115
      Width           =   1245
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4860
      TabIndex        =   1
      Top             =   5115
      Width           =   1245
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8370
      Top             =   4710
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5595
      FormDesignWidth =   9360
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCategoryMatching 
      Height          =   3990
      Left            =   300
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   7038
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      Rows            =   10
      Cols            =   3
      FixedCols       =   0
      BackColorSel    =   12632256
      ForeColorSel    =   16711680
      BackColorUnpopulated=   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   3
   End
   Begin VB.PictureBox pbcOwnerOptions 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   285
      ScaleHeight     =   420
      ScaleWidth      =   7980
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4575
      Visible         =   0   'False
      Width           =   7980
      Begin VB.OptionButton rbcOwnerOptions 
         Caption         =   "Automatically update Stations with Import Names"
         Height          =   345
         Index           =   0
         Left            =   1800
         TabIndex        =   12
         Top             =   30
         Value           =   -1  'True
         Width           =   2790
      End
      Begin VB.OptionButton rbcOwnerOptions 
         Caption         =   "Manually define replacement Owners with Import Names"
         Height          =   345
         Index           =   1
         Left            =   4815
         TabIndex        =   13
         Top             =   30
         Width           =   3015
      End
      Begin VB.Label lacOwnerOptions 
         Caption         =   "Station Owners"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   75
         Width           =   1530
      End
   End
   Begin VB.Label lacTitle 
      Alignment       =   2  'Center
      Caption         =   "Title"
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   75
      Width           =   8760
   End
End
Attribute VB_Name = "frmCategoryMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hmFrom As Integer
Private smCallLetters As String
Private smBand As String
Private smCallLettersPlusBand As String
Private smMarketName As String
Private smRank As String
Private smOwnerName As String
Private smFormat As String

Private imBSMode As Integer

'Grid Controls
Private imCtrlVisible As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on

Const FROMINDEX = 0
Const TOINDEX = 1
Const REFINDEX = 2




Private Sub cmcCancel_Click()
    igNewNamesImportedReturn = 0
    Unload frmCategoryMatching
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub cmcDone_Click()

    Dim slName As String
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim slStr As String

    If (igNewNamesImportedType = 3) And (rbcOwnerOptions(0).Value = True) Then
        igNewNamesImportedReturn = 2
    Else
        igNewNamesImportedReturn = 1
        'Move grid values back into array
        If Not mMoveGridToRec() Then
            Exit Sub
        End If
    End If
    Unload frmCategoryMatching
    
End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcDropDown_Click()
    lbcToName.Visible = Not lbcToName.Visible
End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    slStr = edcDropdown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    llRow = SendMessageByString(lbcToName.hwnd, LB_FINDSTRING, -1, slStr)
    If llRow >= 0 Then
        lbcToName.ListIndex = llRow
        edcDropdown.Text = lbcToName.List(lbcToName.ListIndex)
        edcDropdown.SelStart = ilLen
        edcDropdown.SelLength = Len(edcDropdown.Text)
    End If

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
        gProcessArrowKey Shift, KeyCode, lbcToName, True
    End If
End Sub

Private Sub Form_Click()
    mSetShow
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 1.7 '1.1
    Me.Height = (Screen.Height) / 1.3
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
    gCenterForm Me
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    'Load Grid
    imBSMode = False
    imCtrlVisible = False
    If igNewNamesImportedType = 3 Then
        pbcOwnerOptions.Visible = True
    Else
        pbcOwnerOptions.Visible = False
    End If
    tmcFillGrid.Enabled = True
    Screen.MousePointer = vbDefault

End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    mSetGridColumns
    mSetGridTitles
    mClearGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCategoryMatching = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdCategoryMatching.ColWidth(REFINDEX) = 0
    grdCategoryMatching.ColWidth(FROMINDEX) = (grdCategoryMatching.Width - GRIDSCROLLWIDTH - 15) / 2
    grdCategoryMatching.ColWidth(TOINDEX) = grdCategoryMatching.Width - GRIDSCROLLWIDTH - 15 - grdCategoryMatching.ColWidth(FROMINDEX)
    
    'Align columns to left
    gGrid_AlignAllColsLeft grdCategoryMatching
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    If igNewNamesImportedType <= 3 Then
        grdCategoryMatching.TextMatrix(0, FROMINDEX) = "Import(New) Name"
        grdCategoryMatching.TextMatrix(0, TOINDEX) = "Current(Old) Name"
    Else
        grdCategoryMatching.TextMatrix(0, FROMINDEX) = "Import Name"
        grdCategoryMatching.TextMatrix(0, TOINDEX) = "Current Name"
    End If
    Select Case igNewNamesImportedType
        Case 1  'DMA Market
            lacTitle.Caption = "DMA Market"
        Case 2  'MSA Market
            lacTitle.Caption = "MSA Market"
        Case 3  'Owner
            lacTitle.Caption = "Owner"
        Case 4, 5  'Format
            lacTitle.Caption = "Format"
    End Select
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    gGrid_Clear grdCategoryMatching, True
    'Set color within cells
    For llRow = grdCategoryMatching.FixedRows To grdCategoryMatching.Rows - 1 Step 1
        For llCol = FROMINDEX To REFINDEX Step 1
            If llCol = FROMINDEX Then
                grdCategoryMatching.Row = llRow
                grdCategoryMatching.Col = llCol
                grdCategoryMatching.CellBackColor = LIGHTYELLOW
            End If
            grdCategoryMatching.TextMatrix(llRow, llCol) = ""
        Next llCol
    Next llRow
End Sub

Private Sub grdCategoryMatching_Click()
    If UBound(tgNewNamesImported) <= LBound(tgNewNamesImported) Then
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdCategoryMatching_EnterCell()
    If UBound(tgNewNamesImported) <= LBound(tgNewNamesImported) Then
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    mSetShow
End Sub

Private Sub grdCategoryMatching_GotFocus()
    If UBound(tgNewNamesImported) <= LBound(tgNewNamesImported) Then
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    If grdCategoryMatching.Col >= grdCategoryMatching.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdCategoryMatching_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdCategoryMatching.TopRow
    grdCategoryMatching.Redraw = False
End Sub

Private Sub grdCategoryMatching_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim ilType As Integer
    
    If UBound(tgNewNamesImported) <= LBound(tgNewNamesImported) Then
        grdCategoryMatching.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdCategoryMatching, X, Y)
    If Not ilFound Then
        grdCategoryMatching.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    If grdCategoryMatching.Row - 1 >= UBound(tgNewNamesImported) Then
        grdCategoryMatching.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    
    If Trim$(grdCategoryMatching.TextMatrix(grdCategoryMatching.Row, FROMINDEX)) = "" Then
        grdCategoryMatching.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    If Not mColOk() Then
        grdCategoryMatching.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    If grdCategoryMatching.Col > TOINDEX Then
        Exit Sub
    End If
    lmTopRow = grdCategoryMatching.TopRow
    grdCategoryMatching.Redraw = True
    mEnableBox
End Sub

Private Sub grdCategoryMatching_Scroll()
    If grdCategoryMatching.Redraw = False Then
        grdCategoryMatching.Redraw = True
        grdCategoryMatching.TopRow = lmTopRow
        grdCategoryMatching.Refresh
        grdCategoryMatching.Redraw = False
    End If
    If (imCtrlVisible) And (grdCategoryMatching.Row >= grdCategoryMatching.FixedRows) And (grdCategoryMatching.Col >= TOINDEX) And (grdCategoryMatching.Col < grdCategoryMatching.Cols - 1) Then
        'If grdCategoryMatching.RowIsVisible(grdCategoryMatching.Row) Then
        '   If grdCategoryMatching.Col = TOINDEX Then
        '        edcDropdown.Move grdCategoryMatching.Left + grdCategoryMatching.ColPos(grdCategoryMatching.Col) + 30, grdCategoryMatching.Top + grdCategoryMatching.RowPos(grdCategoryMatching.Row) + 15, grdCategoryMatching.ColWidth(grdCategoryMatching.Col) - cmcDropDown.Width - 30, grdCategoryMatching.RowHeight(grdCategoryMatching.Row) - 15
        '        cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
        '        lbcToName.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
        '        edcDropdown.Visible = True
        '        cmcDropDown.Visible = True
        '        lbcToName.Visible = True
        '        edcDropdown.SetFocus
        '    End If
        'Else
        '    edcDropdown.Visible = False
        '    cmcDropDown.Visible = False
        '    lbcToName.Visible = False
        'End If
        mSetShow
        cmcCancel.SetFocus
    End If

End Sub

Private Sub lbcToName_Click()
    edcDropdown.Text = lbcToName.List(lbcToName.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcToName.Visible = False
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim slStr As String
    Dim ilNext As Integer
    Dim ilTestValue As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        ilTestValue = True
        Do
            ilNext = False
            Select Case grdCategoryMatching.Col
                Case TOINDEX
                    llEnableRow = lmEnableRow
                    llEnableCol = lmEnableCol
                    mSetShow
                    lmEnableRow = llEnableRow
                    lmEnableCol = llEnableCol
                    'If mGridFieldsOk(CInt(lmEnableRow)) = False Then
                    '    mEnableBox
                    '    Exit Sub
                    'End If
                    If (grdCategoryMatching.Row + 1 >= grdCategoryMatching.Rows) Then
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    If (grdCategoryMatching.Row + 1 < grdCategoryMatching.Rows) Then
                        If (Trim$(grdCategoryMatching.TextMatrix(grdCategoryMatching.Row + 1, FROMINDEX)) = "") Then
                            cmcDone.SetFocus
                            Exit Sub
                        End If
                    End If
                    grdCategoryMatching.Row = grdCategoryMatching.Row + 1
                    grdCategoryMatching.Col = FROMINDEX
                    If Not grdCategoryMatching.RowIsVisible(grdCategoryMatching.Row) Then
                        grdCategoryMatching.TopRow = grdCategoryMatching.TopRow + 1
                    End If
                Case Else
                    grdCategoryMatching.Col = grdCategoryMatching.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdCategoryMatching.TopRow = grdCategoryMatching.FixedRows
        grdCategoryMatching.Col = TOINDEX
        Do
            If grdCategoryMatching.Row <= grdCategoryMatching.FixedRows Then
                cmcDone.SetFocus
                Exit Sub
            End If
            grdCategoryMatching.Row = grdCategoryMatching.Rows - 1
            Do
                If Not grdCategoryMatching.RowIsVisible(grdCategoryMatching.Row) Then
                    grdCategoryMatching.TopRow = grdCategoryMatching.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
            If mColOk() Then
                Exit Do
            End If
        Loop
    End If
    lmTopRow = grdCategoryMatching.TopRow
    mEnableBox
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        Do
            ilNext = False
            Select Case grdCategoryMatching.Col
                Case TOINDEX
                    If grdCategoryMatching.Row = grdCategoryMatching.FixedRows Then
                        mSetShow
                        cmcDone.SetFocus
                        Exit Sub
                    End If
                    lmTopRow = -1
                    grdCategoryMatching.Row = grdCategoryMatching.Row - 1
                    If Not grdCategoryMatching.RowIsVisible(grdCategoryMatching.Row) Then
                        grdCategoryMatching.TopRow = grdCategoryMatching.TopRow - 1
                    End If
                    grdCategoryMatching.Col = TOINDEX
                Case Else
                    grdCategoryMatching.Col = grdCategoryMatching.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        lmTopRow = -1
        grdCategoryMatching.TopRow = grdCategoryMatching.FixedRows
        grdCategoryMatching.Row = grdCategoryMatching.FixedRows
        grdCategoryMatching.Col = TOINDEX
        Do
            If mColOk() Then
                Exit Do
            End If
            If grdCategoryMatching.Row + 1 >= grdCategoryMatching.Rows Then
                cmcDone.SetFocus
                Exit Sub
            End If
            grdCategoryMatching.Row = grdCategoryMatching.Row + 1
            Do
                If Not grdCategoryMatching.RowIsVisible(grdCategoryMatching.Row) Then
                    grdCategoryMatching.TopRow = grdCategoryMatching.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
        Loop
    End If
    lmTopRow = grdCategoryMatching.TopRow
    mEnableBox
End Sub

Private Sub rbcOwnerOptions_Click(Index As Integer)
    If rbcOwnerOptions(Index).Value = True Then
        If Index = 0 Then
            grdCategoryMatching.Enabled = False
        Else
            grdCategoryMatching.Enabled = True
        End If
    End If
End Sub

Private Sub tmcFillGrid_Timer()
    Dim llLoop As Long
    Dim llRow As Long
    Dim llCol As Long
    Dim ilTestCol As Integer
    Dim ilSortCol As Integer
    Dim ilPrevSortCol As Integer
    Dim ilPrevSortDirection As Integer
    Dim ilFlt As Integer
    
    tmcFillGrid.Enabled = False
    frmCategoryMatching.MousePointer = vbHourglass
    gSetMousePointer grdCategoryMatching, grdCategoryMatching, vbHourglass
    grdCategoryMatching.Redraw = False
    llRow = grdCategoryMatching.FixedRows
    For llLoop = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
        If llRow >= grdCategoryMatching.Rows Then
            grdCategoryMatching.AddItem ""
        End If
        For llCol = FROMINDEX To REFINDEX Step 1
            If llCol = FROMINDEX Then
                grdCategoryMatching.Row = llRow
                grdCategoryMatching.Col = llCol
                grdCategoryMatching.CellBackColor = LIGHTYELLOW
            End If
            grdCategoryMatching.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdCategoryMatching.TextMatrix(llRow, FROMINDEX) = Trim$(tgNewNamesImported(llLoop).sNewName)
        If igNewNamesImportedType = 5 Then
            For ilFlt = LBound(tgFormatInfo) To UBound(tgFormatInfo) - 1 Step 1
                If tgFormatInfo(ilFlt).lCode = CInt(tgNewNamesImported(llLoop).lReplaceCode) Then
                    grdCategoryMatching.TextMatrix(llRow, TOINDEX) = Trim$(tgFormatInfo(ilFlt).sName)
                    Exit For
                End If
            Next ilFlt
        End If
        grdCategoryMatching.TextMatrix(llRow, REFINDEX) = llLoop
        llRow = llRow + 1
    Next llLoop
    ilTestCol = FROMINDEX
    ilSortCol = FROMINDEX
    ilPrevSortCol = -1
    ilPrevSortDirection = -1
    gGrid_SortByCol grdCategoryMatching, ilTestCol, ilSortCol, ilPrevSortCol, ilPrevSortDirection
    grdCategoryMatching.Row = 0
    grdCategoryMatching.Col = REFINDEX
    grdCategoryMatching.Redraw = True
    lbcToName.Clear
    Select Case igNewNamesImportedType
        Case 1  'DMA Market
            For llLoop = LBound(tgMarketInfo) To UBound(tgMarketInfo) - 1 Step 1
                lbcToName.AddItem Trim$(tgMarketInfo(llLoop).sName)
                lbcToName.ItemData(lbcToName.NewIndex) = tgMarketInfo(llLoop).lCode
            Next llLoop
        Case 2  'MSA Market
            For llLoop = LBound(tgMSAMarketInfo) To UBound(tgMSAMarketInfo) - 1 Step 1
                lbcToName.AddItem Trim$(tgMSAMarketInfo(llLoop).sName)
                lbcToName.ItemData(lbcToName.NewIndex) = tgMSAMarketInfo(llLoop).lCode
            Next llLoop
        Case 3  'Owner
            For llLoop = LBound(tgOwnerInfo) To UBound(tgOwnerInfo) - 1 Step 1
                lbcToName.AddItem Trim$(tgOwnerInfo(llLoop).sName)
                lbcToName.ItemData(lbcToName.NewIndex) = tgOwnerInfo(llLoop).lCode
            Next llLoop
        Case 4, 5  'Format
            For llLoop = LBound(tgFormatInfo) To UBound(tgFormatInfo) - 1 Step 1
                lbcToName.AddItem Trim$(tgFormatInfo(llLoop).sName)
                lbcToName.ItemData(lbcToName.NewIndex) = tgFormatInfo(llLoop).lCode
            Next llLoop
    End Select
    If (igNewNamesImportedType <> 4) And (igNewNamesImportedType <> 5) Then
        lbcToName.AddItem "[Add as New Name]", 0
        lbcToName.ItemData(lbcToName.NewIndex) = 0
    End If
    If igNewNamesImportedType <> 5 Then
        grdCategoryMatching.Row = grdCategoryMatching.FixedRows
        grdCategoryMatching.Col = TOINDEX
        mEnableBox
    Else
        cmcCancel.SetFocus
    End If
    frmCategoryMatching.MousePointer = vbDefault
    gSetMousePointer grdCategoryMatching, grdCategoryMatching, vbDefault
End Sub

Private Function mColOk() As Integer
    mColOk = True
    If grdCategoryMatching.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function

Private Sub mEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    If (grdCategoryMatching.Row >= grdCategoryMatching.FixedRows) And (grdCategoryMatching.Row < grdCategoryMatching.Rows) And (grdCategoryMatching.Col >= TOINDEX) And (grdCategoryMatching.Col < grdCategoryMatching.Cols - 1) Then
        lmEnableRow = grdCategoryMatching.Row
        lmEnableCol = grdCategoryMatching.Col
        imCtrlVisible = True
        Select Case grdCategoryMatching.Col
            Case TOINDEX  'Advertiser
                edcDropdown.Move grdCategoryMatching.Left + grdCategoryMatching.ColPos(grdCategoryMatching.Col) + 30, grdCategoryMatching.Top + grdCategoryMatching.RowPos(grdCategoryMatching.Row) + 15, grdCategoryMatching.ColWidth(grdCategoryMatching.Col) - cmcDropDown.Width - 30, grdCategoryMatching.RowHeight(grdCategoryMatching.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                gSetListBoxHeight lbcToName, 12
                If edcDropdown.Top + edcDropdown.Height + lbcToName.Height <= cmcCancel.Top Then
                    lbcToName.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                Else
                    lbcToName.Move edcDropdown.Left, edcDropdown.Top - lbcToName.Height, edcDropdown.Width + cmcDropDown.Width
                End If
                slStr = grdCategoryMatching.Text
                ilIndex = SendMessageByString(lbcToName.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcToName.ListIndex = ilIndex
                    edcDropdown.Text = lbcToName.List(lbcToName.ListIndex)
                Else
                    lbcToName.ListIndex = -1
                    edcDropdown.Text = ""
                End If
                If edcDropdown.Height > grdCategoryMatching.RowHeight(grdCategoryMatching.Row) - 15 Then
                    edcDropdown.FontName = "Arial"
                    edcDropdown.Height = grdCategoryMatching.RowHeight(grdCategoryMatching.Row) - 15
                End If
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcToName.Visible = True
                edcDropdown.SetFocus
        End Select
    End If
End Sub

Private Sub mSetShow()
    Dim slStr As String
    
    If (lmEnableRow >= grdCategoryMatching.FixedRows) And (lmEnableRow < grdCategoryMatching.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case TOINDEX
                slStr = edcDropdown.Text
                If grdCategoryMatching.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    grdCategoryMatching.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcToName.Visible = False
End Sub

Private Function mMoveGridToRec() As Integer
    Dim llRow As Long
    Dim blError As Boolean
    Dim llIndex As Long
    Dim slStr As String
    Dim ilIndex As Integer
    
    blError = False
    For llRow = grdCategoryMatching.FixedRows To grdCategoryMatching.Rows - 1 Step 1
        If Trim$(grdCategoryMatching.TextMatrix(llRow, FROMINDEX)) <> "" Then
            If Trim$(grdCategoryMatching.TextMatrix(llRow, TOINDEX)) <> "" Then
                llIndex = grdCategoryMatching.TextMatrix(llRow, REFINDEX)
                slStr = Trim$(grdCategoryMatching.TextMatrix(llRow, TOINDEX))
                ilIndex = SendMessageByString(lbcToName.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    tgNewNamesImported(llIndex).lReplaceCode = lbcToName.ItemData(ilIndex)
                Else
                    grdCategoryMatching.TextMatrix(llRow, TOINDEX) = "Missing"
                    blError = True
                End If
            Else
                grdCategoryMatching.TextMatrix(llRow, TOINDEX) = "Missing"
                blError = True
            End If
        End If
    Next llRow
    If blError Then
        mMoveGridToRec = False
    Else
        mMoveGridToRec = True
    End If
End Function

