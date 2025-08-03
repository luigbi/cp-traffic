VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrReplaceSchd 
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   9375
   ControlBox      =   0   'False
   Icon            =   "EngrReplaceSchd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9375
   Begin VB.CheckBox ckcEventType 
      Caption         =   "Spots"
      Height          =   195
      Index           =   2
      Left            =   4320
      TabIndex        =   17
      Top             =   255
      Value           =   1  'Checked
      Width           =   1125
   End
   Begin VB.CheckBox ckcEventType 
      Caption         =   "Avails"
      Height          =   195
      Index           =   1
      Left            =   3300
      TabIndex        =   16
      Top             =   255
      Value           =   1  'Checked
      Width           =   1125
   End
   Begin VB.CheckBox ckcEventType 
      Caption         =   "Programs"
      Height          =   195
      Index           =   0
      Left            =   2145
      TabIndex        =   15
      Top             =   255
      Value           =   1  'Checked
      Width           =   1125
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9270
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   45
      Width           =   45
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   8220
      Top             =   5310
   End
   Begin VB.ListBox lbcFieldName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrReplaceSchd.frx":030A
      Left            =   5055
      List            =   "EngrReplaceSchd.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   1410
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
      Left            =   4875
      Picture         =   "EngrReplaceSchd.frx":030E
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2595
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   3930
      TabIndex        =   6
      Top             =   2625
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lbcFileList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrReplaceSchd.frx":0408
      Left            =   1665
      List            =   "EngrReplaceSchd.frx":040A
      Sorted          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   1410
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
      TabIndex        =   10
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
      Picture         =   "EngrReplaceSchd.frx":040C
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.CommandButton cmcClear 
      Caption         =   "&Clear All"
      Height          =   375
      Left            =   5670
      TabIndex        =   13
      Top             =   5370
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
      FormDesignHeight=   5895
      FormDesignWidth =   9375
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3900
      TabIndex        =   12
      Top             =   5370
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   2145
      TabIndex        =   11
      Top             =   5355
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdReplace 
      Height          =   4530
      Left            =   375
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   525
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   7990
      _Version        =   393216
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
   Begin VB.Label lacEventType 
      Caption         =   "Apply to Event Type(s):"
      Height          =   195
      Left            =   405
      TabIndex        =   18
      Top             =   255
      Width           =   1875
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   735
      Picture         =   "EngrReplaceSchd.frx":0716
      Top             =   5280
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Replace Scheduled Information"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2415
   End
End
Attribute VB_Name = "EngrReplaceSchd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrReplaceSchd - enters affiliate representative information
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
Private imVersion As Integer
Private smNowDate As String
Private smNowTime As String
Private imFirstActivate As Integer
Private imIgnoreScroll As Integer
Private imFieldType As Integer


'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer



Const FIELDNAMEINDEX = 0
Const OLDVALUEINDEX = 1
Const NEWVALUEINDEX = 2
Const OLDCODEINDEX = 3
Const NEWCODEINDEX = 4

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub




Private Sub mSortCol(ilCol As Integer)
    mSetShow
    gGrid_SortByCol grdReplace, FIELDNAMEINDEX, ilCol, imLastColSorted, imLastSort
End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    If imInChg Then
        Exit Sub
    End If
End Sub

Private Sub mEnableBox()
    Dim slStr As String
    Dim llListIndex As Long
    Dim ilIndex As Integer
    
    If (grdReplace.Row >= grdReplace.FixedRows) And (grdReplace.Row < grdReplace.Rows) And (grdReplace.Col >= 0) And (grdReplace.Col < grdReplace.Cols) Then
        lmEnableRow = grdReplace.Row
        lmEnableCol = grdReplace.Col
        sgReturnCallName = grdReplace.TextMatrix(lmEnableRow, FIELDNAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdReplace.Left - pbcArrow.Width - 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + (grdReplace.RowHeight(grdReplace.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If Trim$(grdReplace.TextMatrix(grdReplace.Row, FIELDNAMEINDEX)) <> "" Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdReplace.Col
            Case FIELDNAMEINDEX
                edcDropdown.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - cmcDropDown.Width - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                edcDropdown.MaxLength = 0
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcFieldName.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcFieldName, CLng(grdReplace.Height / 2)
                If lbcFieldName.Top + lbcFieldName.Height > cmcCancel.Top Then
                    lbcFieldName.Top = edcDropdown.Top - lbcFieldName.Height
                End If
                slStr = grdReplace.text
                ilIndex = gListBoxFind(lbcFieldName, slStr)
                If ilIndex >= 0 Then
                    lbcFieldName.ListIndex = ilIndex
                    edcDropdown.text = lbcFieldName.List(lbcFieldName.ListIndex)
                Else
                    edcDropdown.text = ""
                    lbcFieldName.ListIndex = -1
                End If
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcFieldName.Visible = True
                edcDropdown.SetFocus
            Case OLDVALUEINDEX  'Date
                slStr = Trim$(grdReplace.TextMatrix(grdReplace.Row, FIELDNAMEINDEX))
                If slStr <> "" Then
                    llListIndex = gListBoxFind(lbcFieldName, slStr)
                    If llListIndex >= 0 Then
                        ilIndex = lbcFieldName.ItemData(llListIndex)
                        imFieldType = tgReplaceFields(ilIndex).iFieldType
                        If (imFieldType = 5) Or (imFieldType = 9) Then  'List
                            mPopOldList tgReplaceFields(ilIndex).sListFile
                            edcDropdown.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - cmcDropDown.Width - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                            edcDropdown.MaxLength = tgReplaceFields(ilIndex).iMaxNoChar
                            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                            lbcFileList.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                            gSetListBoxHeight lbcFileList, CLng(grdReplace.Height / 2)
                            If lbcFileList.Top + lbcFileList.Height > cmcCancel.Top Then
                                lbcFileList.Top = edcDropdown.Top - lbcFileList.Height
                            End If
                            slStr = grdReplace.text
                            ilIndex = gListBoxFind(lbcFileList, slStr)
                            If ilIndex >= 0 Then
                                lbcFileList.ListIndex = ilIndex
                                edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
                            Else
                                If lbcFileList.ListCount = 1 Then
                                    lbcFileList.ListIndex = 0
                                    edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
                                Else
                                    edcDropdown.text = ""
                                    lbcFileList.ListIndex = -1
                                End If
                            End If
                            tmcClick.Enabled = False
                            edcDropdown.Visible = True
                            cmcDropDown.Visible = True
                            lbcFileList.Visible = True
                            edcDropdown.SetFocus
                        Else
                            edcGrid.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                            edcGrid.MaxLength = tgReplaceFields(ilIndex).iMaxNoChar
                            edcGrid.text = grdReplace.text
                            edcGrid.Visible = True
                            edcGrid.SetFocus
                        End If
                    End If
                End If
            Case NEWVALUEINDEX
                slStr = Trim$(grdReplace.TextMatrix(grdReplace.Row, FIELDNAMEINDEX))
                If slStr <> "" Then
                    llListIndex = gListBoxFind(lbcFieldName, slStr)
                    If llListIndex >= 0 Then
                        ilIndex = lbcFieldName.ItemData(llListIndex)
                        imFieldType = tgReplaceFields(ilIndex).iFieldType
                        If (imFieldType = 5) Or (imFieldType = 9) Then  'List
                            mPopNewList tgReplaceFields(ilIndex).sListFile
                            edcDropdown.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - cmcDropDown.Width - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                            edcDropdown.MaxLength = tgReplaceFields(ilIndex).iMaxNoChar
                            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                            lbcFileList.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                            gSetListBoxHeight lbcFileList, CLng(grdReplace.Height / 2)
                            If lbcFileList.Top + lbcFileList.Height > cmcCancel.Top Then
                                lbcFileList.Top = edcDropdown.Top - lbcFileList.Height
                            End If
                            slStr = grdReplace.text
                            ilIndex = gListBoxFind(lbcFileList, slStr)
                            If ilIndex >= 0 Then
                                lbcFileList.ListIndex = ilIndex
                                edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
                            Else
                                If lbcFileList.ListCount = 1 Then
                                    lbcFileList.ListIndex = 0
                                    edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
                                Else
                                    edcDropdown.text = ""
                                    lbcFileList.ListIndex = -1
                                End If
                            End If
                            tmcClick.Enabled = False
                            edcDropdown.Visible = True
                            cmcDropDown.Visible = True
                            lbcFileList.Visible = True
                            edcDropdown.SetFocus
                        Else
                            edcGrid.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                            edcGrid.MaxLength = tgReplaceFields(ilIndex).iMaxNoChar
                            edcGrid.text = grdReplace.text
                            edcGrid.Visible = True
                            edcGrid.SetFocus
                        End If
                    End If
                End If
                
        End Select
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdReplace.FixedRows) And (lmEnableRow < grdReplace.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case FIELDNAMEINDEX
            Case OLDVALUEINDEX
            Case NEWVALUEINDEX
        End Select
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    edcGrid.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcFieldName.Visible = False
    lbcFileList.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub
Private Function mCheckFields(ilTestState As Integer) As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llRow As Long
    Dim slFieldName As String
    Dim llListIndex As Long
    Dim ilIndex As Integer
    
    grdReplace.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdReplace.FixedRows To grdReplace.Rows - 1 Step 1
        slStr = Trim$(grdReplace.TextMatrix(llRow, FIELDNAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdReplace.TextMatrix(llRow, OLDVALUEINDEX)
            If slStr <> "" Then
                ilError = True
                grdReplace.TextMatrix(llRow, FIELDNAMEINDEX) = "Missing"
                grdReplace.Row = llRow
                grdReplace.Col = FIELDNAMEINDEX
                grdReplace.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = Trim$(grdReplace.TextMatrix(llRow, OLDVALUEINDEX))
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdReplace.TextMatrix(llRow, OLDVALUEINDEX) = "Missing"
                    grdReplace.Row = llRow
                    grdReplace.Col = OLDVALUEINDEX
                    grdReplace.CellForeColor = vbRed
                Else
                    slFieldName = Trim$(grdReplace.TextMatrix(llRow, FIELDNAMEINDEX))
                    llListIndex = gListBoxFind(lbcFieldName, slFieldName)
                    If llListIndex >= 0 Then
                        ilIndex = lbcFieldName.ItemData(llListIndex)
                        imFieldType = tgReplaceFields(ilIndex).iFieldType
                        If imFieldType = 3 Then
                            If Not gIsDate(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 4 Then
                            If Not gIsTime(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 6 Then
                            If Not gIsTimeTenths(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 7 Then
                            If Not gIsLength(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 8 Then
                            If Not gIsLengthTenths(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        End If
                    Else
                        ilError = True
                        grdReplace.Row = llRow
                        grdReplace.Col = FIELDNAMEINDEX
                        grdReplace.CellForeColor = vbRed
                    End If
                End If
                slStr = Trim$(grdReplace.TextMatrix(llRow, NEWVALUEINDEX))
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdReplace.TextMatrix(llRow, NEWVALUEINDEX) = "Missing"
                    grdReplace.Row = llRow
                    grdReplace.Col = NEWVALUEINDEX
                    grdReplace.CellForeColor = vbRed
                Else
                    slFieldName = Trim$(grdReplace.TextMatrix(llRow, FIELDNAMEINDEX))
                    llListIndex = gListBoxFind(lbcFieldName, slFieldName)
                    If llListIndex >= 0 Then
                        ilIndex = lbcFieldName.ItemData(llListIndex)
                        imFieldType = tgReplaceFields(ilIndex).iFieldType
                        If imFieldType = 3 Then
                            If Not gIsDate(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 4 Then
                            If Not gIsTime(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 6 Then
                            If Not gIsTimeTenths(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 7 Then
                            If Not gIsLength(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 8 Then
                            If Not gIsLengthTenths(slStr) Then
                                ilError = True
                                grdReplace.Row = llRow
                                grdReplace.Col = NEWVALUEINDEX
                                grdReplace.CellForeColor = vbRed
                            End If
                        End If
                    Else
                        ilError = True
                        grdReplace.Row = llRow
                        grdReplace.Col = FIELDNAMEINDEX
                        grdReplace.CellForeColor = vbRed
                    End If
                End If
            End If
        End If
    Next llRow
    grdReplace.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdReplace
    mGridColumnWidth
    'Set Titles
    grdReplace.TextMatrix(0, FIELDNAMEINDEX) = "Field Name"
    grdReplace.TextMatrix(0, OLDVALUEINDEX) = "Old Value"
    grdReplace.TextMatrix(0, NEWVALUEINDEX) = "New Value"
    grdReplace.Row = 1
    For ilCol = 0 To grdReplace.Cols - 1 Step 1
        grdReplace.Col = ilCol
        grdReplace.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdReplace.Height = cmcCancel.Top - grdReplace.Top - 120    '8 * grdReplace.RowHeight(0) + 30
    gGrid_IntegralHeight grdReplace
    gGrid_Clear grdReplace, True
    grdReplace.Row = grdReplace.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdReplace.Width = EngrReplaceSchd.Width - 2 * grdReplace.Left
    grdReplace.ColWidth(OLDCODEINDEX) = 0
    grdReplace.ColWidth(NEWCODEINDEX) = 0
    grdReplace.ColWidth(FIELDNAMEINDEX) = grdReplace.Width / 4
    grdReplace.ColWidth(OLDVALUEINDEX) = grdReplace.Width / 4
    grdReplace.ColWidth(NEWVALUEINDEX) = grdReplace.Width - GRIDSCROLLWIDTH
    For ilCol = FIELDNAMEINDEX To NEWVALUEINDEX Step 1
        If ilCol <> NEWVALUEINDEX Then
            If grdReplace.ColWidth(NEWVALUEINDEX) > grdReplace.ColWidth(ilCol) Then
                grdReplace.ColWidth(NEWVALUEINDEX) = grdReplace.ColWidth(NEWVALUEINDEX) - grdReplace.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    edcGrid.text = ""
    edcDropdown.text = ""
    
    gGrid_Clear grdReplace, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec()
    Dim ilRow As Integer
    Dim slStr As String
    Dim ilUpper As Integer
    Dim llListIndex As Long
    Dim ilLoop As Integer
    
    For ilLoop = 0 To 2 Step 1
        If ckcEventType(ilLoop).Value = vbChecked Then
            bgApplyToEventType(ilLoop) = True
        Else
            bgApplyToEventType(ilLoop) = False
        End If
    Next ilLoop
    ReDim tgSchdReplaceValues(0 To 0) As SCHDREPLACEVALUES
    For ilRow = grdReplace.FixedRows To grdReplace.Rows - 1 Step 1
        slStr = Trim$(grdReplace.TextMatrix(ilRow, FIELDNAMEINDEX))
        If slStr <> "" Then
            ilUpper = UBound(tgSchdReplaceValues)
            tgSchdReplaceValues(ilUpper).sFieldName = slStr
            slStr = Trim$(grdReplace.TextMatrix(ilRow, OLDVALUEINDEX))
            tgSchdReplaceValues(ilUpper).sOldValue = slStr
            slStr = Trim$(grdReplace.TextMatrix(ilRow, OLDCODEINDEX))
            tgSchdReplaceValues(ilUpper).lOldCode = Val(slStr)
            'Missing code
            slStr = Trim$(grdReplace.TextMatrix(ilRow, NEWVALUEINDEX))
            tgSchdReplaceValues(ilUpper).sNewValue = slStr
            slStr = Trim$(grdReplace.TextMatrix(ilRow, NEWCODEINDEX))
            tgSchdReplaceValues(ilUpper).lNewCode = Val(slStr)
            ReDim Preserve tgSchdReplaceValues(0 To ilUpper + 1) As SCHDREPLACEVALUES
        End If
    Next ilRow
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilList As Integer
    Dim slStr As String
    Dim llListRow As Long

    
    'gGrid_Clear grdReplace, True
    llRow = grdReplace.FixedRows
    For ilLoop = 0 To UBound(tgSchdReplaceValues) - 1 Step 1
        If llRow + 1 > grdReplace.Rows Then
            grdReplace.AddItem ""
        End If
        grdReplace.Row = llRow
        slStr = Trim$(tgSchdReplaceValues(ilLoop).sFieldName)
        grdReplace.TextMatrix(llRow, FIELDNAMEINDEX) = Trim$(tgSchdReplaceValues(ilLoop).sFieldName)
        grdReplace.TextMatrix(llRow, OLDVALUEINDEX) = Trim$(tgSchdReplaceValues(ilLoop).sOldValue)
        grdReplace.TextMatrix(llRow, OLDCODEINDEX) = Trim$(tgSchdReplaceValues(ilLoop).lOldCode)
        grdReplace.TextMatrix(llRow, NEWVALUEINDEX) = Trim$(tgSchdReplaceValues(ilLoop).sNewValue)
        grdReplace.TextMatrix(llRow, NEWCODEINDEX) = Trim$(tgSchdReplaceValues(ilLoop).lNewCode)
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdReplace.Rows Then
        grdReplace.AddItem ""
    End If
    grdReplace.Redraw = True
End Sub
Private Sub cmcCancel_Click()
    igAnsReplace = CALLCANCELLED
    Unload EngrReplaceSchd
End Sub

Private Sub cmcClear_Click()
    edcGrid.text = ""
    edcDropdown.text = ""
    lbcFileList.Clear
    gGrid_Clear grdReplace, True
    imFieldChgd = True
End Sub

Private Sub cmcClear_GotFocus()
    mSetShow
End Sub

Private Sub cmcDone_Click()
    Dim slStr As String
    Dim ilRet As Integer
    
    If imFieldChgd = False Then
        igAnsReplace = CALLCANCELLED
        Unload EngrReplaceSchd
        Exit Sub
    End If
    If mCheckFields(True) Then
        slStr = "This will Replace All Fields Specified in the Events Currently Displayed on the Schedule Screen, Continue with Replace"
        If MsgBox(slStr, vbYesNo) = vbYes Then
            mMoveCtrlsToRec
            igAnsReplace = CALLDONE
            Unload EngrReplaceSchd
            Exit Sub
        End If
    End If
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub



Private Sub cmcDropDown_Click()
    Select Case grdReplace.Col
        Case FIELDNAMEINDEX
            lbcFieldName.Visible = Not lbcFieldName.Visible
        Case OLDVALUEINDEX
            lbcFileList.Visible = Not lbcFileList.Visible
        Case NEWVALUEINDEX
            lbcFileList.Visible = Not lbcFileList.Visible
    End Select
End Sub

Private Sub edcDropdown_Change()
    Dim slStr As String
    Dim ilLen As Integer
    Dim llRow As Long
    
    slStr = edcDropdown.text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdReplace.Col
        Case FIELDNAMEINDEX
            llRow = gListBoxFind(lbcFieldName, slStr)
            If llRow >= 0 Then
                lbcFieldName.ListIndex = llRow
                edcDropdown.text = lbcFieldName.List(lbcFieldName.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case OLDVALUEINDEX
            llRow = gListBoxFind(lbcFileList, slStr)
            If llRow >= 0 Then
                lbcFileList.ListIndex = llRow
                edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case NEWVALUEINDEX
            llRow = gListBoxFind(lbcFileList, slStr)
            If llRow >= 0 Then
                lbcFileList.ListIndex = llRow
                edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
    End Select
    If (StrComp(grdReplace.text, edcDropdown.text, vbTextCompare) <> 0) Then
        imFieldChgd = True
        Select Case grdReplace.Col
            Case FIELDNAMEINDEX
            Case OLDVALUEINDEX
                If lbcFileList.ListIndex >= 0 Then
                    grdReplace.TextMatrix(grdReplace.Row, OLDCODEINDEX) = lbcFileList.ItemData(lbcFileList.ListIndex)
                End If
            Case NEWVALUEINDEX
                If lbcFileList.ListIndex >= 0 Then
                    grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = lbcFileList.ItemData(lbcFileList.ListIndex)
                End If
        End Select
        grdReplace.text = edcDropdown.text
        grdReplace.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    
    If KeyAscii = 8 Then
        If edcDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
'    If (imMaxColChars < edcDropdown.MaxLength) And (imMaxColChars > 0) And (KeyAscii <> 8) Then
'        slStr = edcEDropdown.Text
'        slStr = Left$(slStr, edcDropdown.SelStart) & Chr$(KeyAscii) & Right$(slStr, Len(slStr) - edcEDropdown.SelStart - edcEDropdown.SelLength)
'        If (Len(slStr) > imMaxColChars) And (Left$(slStr, 1) <> "[") Then
'            Beep
'            KeyAscii = 0
'            Exit Sub
'        End If
'    End If
End Sub

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdReplace.Col
            Case FIELDNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcFieldName, True
            Case OLDVALUEINDEX
                gProcessArrowKey Shift, KeyCode, lbcFileList, True
            Case NEWVALUEINDEX
                gProcessArrowKey Shift, KeyCode, lbcFileList, True
        End Select
        tmcClick.Enabled = False
    End If
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdReplace.Col
        Case FIELDNAMEINDEX
        Case OLDVALUEINDEX
            If grdReplace.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdReplace.CellForeColor = vbBlack
            slStr = edcGrid.text
            If imFieldType = 1 Then
                grdReplace.TextMatrix(grdReplace.Row, OLDCODEINDEX) = slStr
            ElseIf imFieldType = 3 Then
                If gIsDate(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, OLDCODEINDEX) = gDateValue(slStr)
                End If
            ElseIf imFieldType = 4 Then
                If gIsTime(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, OLDCODEINDEX) = gTimeToLong(slStr, False)
                End If
            ElseIf imFieldType = 6 Then
                If gIsTimeTenths(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, OLDCODEINDEX) = gStrTimeInTenthToLong(slStr, False)
                End If
            ElseIf imFieldType = 7 Then
                If gIsLength(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, OLDCODEINDEX) = gLengthToLong(slStr)
                End If
            ElseIf imFieldType = 8 Then
                If gIsLengthTenths(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, OLDCODEINDEX) = gStrLengthInTenthToLong(slStr)
                End If
            End If
        Case NEWVALUEINDEX
            If grdReplace.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdReplace.CellForeColor = vbBlack
            slStr = edcGrid.text
            If imFieldType = 1 Then
                grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = slStr
            ElseIf imFieldType = 3 Then
                If gIsDate(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = gDateValue(slStr)
                End If
            ElseIf imFieldType = 4 Then
                If gIsTime(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = gTimeToLong(slStr, False)
                End If
            ElseIf imFieldType = 6 Then
                If gIsTimeTenths(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = gStrTimeInTenthToLong(slStr, False)
                End If
            ElseIf imFieldType = 7 Then
                If gIsLength(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = gLengthToLong(slStr)
                End If
            ElseIf imFieldType = 8 Then
                If gIsLengthTenths(slStr) Then
                    grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = gStrLengthInTenthToLong(slStr)
                End If
            End If
    End Select
    mSetCommands
End Sub

Private Sub edcGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
    End If
    imFirstActivate = False
End Sub

Private Sub Form_Click()
    cmcCancel.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrReplaceSchd
    gCenterFormModal EngrReplaceSchd
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
End Sub

Private Sub Form_Resize()
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    grdReplace.Height = cmcCancel.Top - grdReplace.Top - 120    '8 * grdReplace.RowHeight(0) + 30
    gGrid_IntegralHeight grdReplace
    gGrid_FillWithRows grdReplace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Erase lmDeleteCodes
    Set EngrReplaceSchd = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdReplace, grdReplace, vbHourglass
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
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
    gSetMousePointer grdReplace, grdReplace, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdReplace, grdReplace, vbDefault
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



Private Sub imcTrash_Click()
    mSetShow
    mDeleteRow
End Sub

Private Sub lbcFileList_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcFieldName_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcFieldName.List(lbcFieldName.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub


Private Sub grdReplace_Click()
    If grdReplace.Col >= grdReplace.Cols Then
        Exit Sub
    End If
End Sub

Private Sub grdReplace_EnterCell()
    mSetShow
End Sub

Private Sub grdReplace_GotFocus()
    If grdReplace.Col >= grdReplace.Cols Then
        Exit Sub
    End If
End Sub

Private Sub grdReplace_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdReplace.TopRow
    grdReplace.Redraw = False
End Sub

Private Sub grdReplace_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdReplace.RowHeight(0) Then
        mSortCol grdReplace.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdReplace, x, y)
    If Not ilFound Then
        grdReplace.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdReplace.Col >= grdReplace.Cols Then
        grdReplace.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdReplace.TopRow
    DoEvents
    llRow = grdReplace.Row
    If grdReplace.TextMatrix(llRow, FIELDNAMEINDEX) = "" Then
        grdReplace.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdReplace.TextMatrix(llRow, FIELDNAMEINDEX) = ""
        grdReplace.Row = llRow + 1
        grdReplace.Col = FIELDNAMEINDEX
        grdReplace.Redraw = True
    End If
    grdReplace.Redraw = True
    mEnableBox
End Sub

Private Sub grdReplace_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdReplace.Redraw = False Then
        grdReplace.Redraw = True
        If lmTopRow < grdReplace.FixedRows Then
            grdReplace.TopRow = grdReplace.FixedRows
        Else
            grdReplace.TopRow = lmTopRow
        End If
        grdReplace.Refresh
        grdReplace.Redraw = False
    End If
    If (imShowGridBox) And (grdReplace.Row >= grdReplace.FixedRows) And (grdReplace.Col >= 0) And (grdReplace.Col < grdReplace.Cols) Then
        If grdReplace.RowIsVisible(grdReplace.Row) Then
            'edcGrid.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 30, grdReplace.ColWidth(grdReplace.Col) - 30, grdReplace.RowHeight(grdReplace.Row) - 30
            pbcArrow.Move grdReplace.Left - pbcArrow.Width - 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + (grdReplace.RowHeight(grdReplace.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            pbcArrow.Visible = False
            edcGrid.Visible = False
            edcDropdown.Visible = False
            cmcDropDown.Visible = False
            lbcFieldName.Visible = False
            lbcFileList.Visible = False
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
    Dim slStr As String
    
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If edcGrid.Visible Or edcDropdown.Visible Then
        If (grdReplace.Col = OLDVALUEINDEX) Or (grdReplace.Col = NEWVALUEINDEX) Then
            slStr = edcGrid.text
            If imFieldType = 3 Then
                If Not gIsDate(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 4 Then
                If Not gIsTime(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 6 Then
                If Not gIsTimeTenths(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 7 Then
                If Not gIsLength(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 8 Then
                If Not gIsLengthTenths(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            End If
        End If
        mSetShow
        If grdReplace.Col = FIELDNAMEINDEX Then
            If grdReplace.Row > grdReplace.FixedRows Then
                lmTopRow = -1
                grdReplace.Row = grdReplace.Row - 1
                If Not grdReplace.RowIsVisible(grdReplace.Row) Then
                    grdReplace.TopRow = grdReplace.TopRow - 1
                End If
                grdReplace.Col = NEWVALUEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdReplace.Col = grdReplace.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdReplace.TopRow = grdReplace.FixedRows
        grdReplace.Col = FIELDNAMEINDEX
        grdReplace.Row = grdReplace.FixedRows
        mEnableBox
    End If
End Sub



Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim slStr As String
    Dim llEnableRow As Long
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If edcGrid.Visible Or edcDropdown.Visible Then
        If (grdReplace.Col = OLDVALUEINDEX) Or (grdReplace.Col = NEWVALUEINDEX) Then
            slStr = edcGrid.text
            If imFieldType = 3 Then
                If Not gIsDate(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 4 Then
                If Not gIsTime(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 6 Then
                If Not gIsTimeTenths(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 7 Then
                If Not gIsLength(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 8 Then
                If Not gIsLengthTenths(slStr) Then
                    Beep
                    edcGrid.SetFocus
                    Exit Sub
                End If
            End If
        End If
        llEnableRow = lmEnableRow
        mSetShow
        If grdReplace.Col = NEWVALUEINDEX Then
            llRow = grdReplace.Rows
            Do
                llRow = llRow - 1
            Loop While grdReplace.TextMatrix(llRow, FIELDNAMEINDEX) = ""
            llRow = llRow + 1
            If (grdReplace.Row + 1 < llRow) Then
                lmTopRow = -1
                grdReplace.Row = grdReplace.Row + 1
                If Not grdReplace.RowIsVisible(grdReplace.Row) Then
                    imIgnoreScroll = True
                    grdReplace.TopRow = grdReplace.TopRow + 1
                End If
                grdReplace.Col = FIELDNAMEINDEX
                'grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = 0
                If Trim$(grdReplace.TextMatrix(grdReplace.Row, FIELDNAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdReplace.Left - pbcArrow.Width - 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + (grdReplace.RowHeight(grdReplace.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdReplace.TextMatrix(llEnableRow, FIELDNAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdReplace.Row + 1 >= grdReplace.Rows Then
                        grdReplace.AddItem ""
                    End If
                    grdReplace.Row = grdReplace.Row + 1
                    If Not grdReplace.RowIsVisible(grdReplace.Row) Then
                        imIgnoreScroll = True
                        grdReplace.TopRow = grdReplace.TopRow + 1
                    End If
                    grdReplace.Col = FIELDNAMEINDEX
                    'grdReplace.TextMatrix(grdReplace.Row, NEWCODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdReplace.Left - pbcArrow.Width - 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + (grdReplace.RowHeight(grdReplace.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdReplace.Col = grdReplace.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdReplace.TopRow = grdReplace.FixedRows
        grdReplace.Col = FIELDNAMEINDEX
        grdReplace.Row = grdReplace.FixedRows
        mEnableBox
    End If
End Sub



Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdReplace.TopRow
    llRow = grdReplace.Row
    slMsg = "Delete " & Trim$(grdReplace.TextMatrix(llRow, FIELDNAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdReplace.Redraw = False
    grdReplace.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdReplace.AddItem ""
    grdReplace.Redraw = False
    grdReplace.TopRow = llTRow
    grdReplace.Redraw = True
    DoEvents
    pbcClickFocus.SetFocus
    mSetCommands
    mDeleteRow = True
End Function





Private Sub mPopulate()
    Dim ilRow As Integer
    Dim ilLoop As Integer
    
    For ilLoop = 0 To UBound(tgReplaceFields) - 1 Step 1
        lbcFieldName.AddItem Trim$(tgReplaceFields(ilLoop).sFieldName)
        lbcFieldName.ItemData(lbcFieldName.NewIndex) = ilLoop
    Next ilLoop
End Sub

Private Sub mPopOldList(slFileName As String)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    lbcFileList.Clear
    Select Case UCase$(Trim$(slFileName))
        Case "ATE"
            For ilLoop = 0 To UBound(tgUsedATE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedATE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedATE(ilLoop).iCode
            Next ilLoop
        Case "ANE"
            For ilLoop = 0 To UBound(tgUsedANE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedANE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedANE(ilLoop).iCode
            Next ilLoop
        Case "BDE"
            For ilLoop = 0 To UBound(tgUsedBDE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedBDE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedBDE(ilLoop).iCode
            Next ilLoop
        Case "ETE"
            For ilLoop = 0 To UBound(tgUsedETE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedETE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedETE(ilLoop).iCode
            Next ilLoop
        Case "FNE"
            For ilLoop = 0 To UBound(tgUsedFNE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedFNE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedFNE(ilLoop).iCode
            Next ilLoop
        Case "MTE"
            For ilLoop = 0 To UBound(tgUsedMTE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedMTE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedMTE(ilLoop).iCode
            Next ilLoop
        Case "NNE"
            For ilLoop = 0 To UBound(tgUsedNNE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedNNE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedNNE(ilLoop).iCode
            Next ilLoop
        Case "RNE"
            For ilLoop = 0 To UBound(tgUsedRNE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedRNE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedRNE(ilLoop).iCode
            Next ilLoop
        Case "TTES"
            For ilLoop = 0 To UBound(tgUsedStartTTE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedStartTTE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedStartTTE(ilLoop).iCode
            Next ilLoop
        Case "TTEE"
            For ilLoop = 0 To UBound(tgUsedEndTTE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedEndTTE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedEndTTE(ilLoop).iCode
            Next ilLoop
        Case "CCEA"
            For ilLoop = 0 To UBound(tgUsedAudioCCE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedAudioCCE(ilLoop).sAutoChar)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedAudioCCE(ilLoop).iCode
            Next ilLoop
        Case "CCEB"
            For ilLoop = 0 To UBound(tgUsedBusCCE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedBusCCE(ilLoop).sAutoChar)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedBusCCE(ilLoop).iCode
            Next ilLoop
        Case "SCE"
            For ilLoop = 0 To UBound(tgUsedSCE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgUsedSCE(ilLoop).sAutoChar)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedSCE(ilLoop).iCode
            Next ilLoop
        '7/8/11: Make T2 work like T1
        Case "CTE2"
            'For ilLoop = 0 To UBound(tgUsedT2CTE) - 1 Step 1
            '    lbcFileList.AddItem Trim$(tgUsedT2CTE(ilLoop).sComment)
            '    lbcFileList.ItemData(lbcFileList.NewIndex) = tgUsedT2CTE(ilLoop).lCode
            'Next ilLoop
            For ilLoop = 0 To UBound(tgT2MatchList) - 1 Step 1
                lbcFileList.AddItem Trim$(tgT2MatchList(ilLoop).sValue)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgT2MatchList(ilLoop).lValue
            Next ilLoop
        Case "FTYN"
            For ilLoop = 0 To UBound(tgYNMatchList) - 1 Step 1
                lbcFileList.AddItem Trim$(tgYNMatchList(ilLoop).sValue)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgYNMatchList(ilLoop).lValue
            Next ilLoop
        Case "CTE1"
            For ilLoop = 0 To UBound(tgT1MatchList) - 1 Step 1
                lbcFileList.AddItem Trim$(tgT1MatchList(ilLoop).sValue)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgT1MatchList(ilLoop).lValue
            Next ilLoop
    End Select
    lbcFileList.AddItem "[None]", 0
    lbcFileList.ItemData(lbcFileList.NewIndex) = 0
End Sub

Private Sub mPopNewList(slFileName As String)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    lbcFileList.Clear
    Select Case UCase$(Trim$(slFileName))
        Case "ATE"
            For ilLoop = 0 To UBound(tgCurrATE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrATE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrATE(ilLoop).iCode
            Next ilLoop
        Case "ANE"
            For ilLoop = 0 To UBound(tgCurrANE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrANE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrANE(ilLoop).iCode
            Next ilLoop
        Case "BDE"
            For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrBDE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrBDE(ilLoop).iCode
            Next ilLoop
        Case "ETE"
            For ilLoop = 0 To UBound(tgCurrETE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrETE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrETE(ilLoop).iCode
            Next ilLoop
        Case "FNE"
            For ilLoop = 0 To UBound(tgCurrFNE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrFNE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrFNE(ilLoop).iCode
            Next ilLoop
        Case "MTE"
            For ilLoop = 0 To UBound(tgCurrMTE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrMTE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrMTE(ilLoop).iCode
            Next ilLoop
        Case "NNE"
            For ilLoop = 0 To UBound(tgCurrNNE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrNNE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrNNE(ilLoop).iCode
            Next ilLoop
        Case "RNE"
            For ilLoop = 0 To UBound(tgCurrRNE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrRNE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrRNE(ilLoop).iCode
            Next ilLoop
        Case "TTES"
            For ilLoop = 0 To UBound(tgCurrStartTTE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrStartTTE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrStartTTE(ilLoop).iCode
            Next ilLoop
        Case "TTEE"
            For ilLoop = 0 To UBound(tgCurrEndTTE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrEndTTE(ilLoop).sName)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrEndTTE(ilLoop).iCode
            Next ilLoop
        Case "CCEA"
            For ilLoop = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrAudioCCE(ilLoop).sAutoChar)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrAudioCCE(ilLoop).iCode
            Next ilLoop
        Case "CCEB"
            For ilLoop = 0 To UBound(tgCurrBusCCE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrBusCCE(ilLoop).sAutoChar)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrBusCCE(ilLoop).iCode
            Next ilLoop
        Case "SCE"
            For ilLoop = 0 To UBound(tgCurrSCE) - 1 Step 1
                lbcFileList.AddItem Trim$(tgCurrSCE(ilLoop).sAutoChar)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrSCE(ilLoop).iCode
            Next ilLoop
        '7/8/11: Make T2 work like T1
        Case "CTE2"
            'For ilLoop = 0 To UBound(tgCurrCTE) - 1 Step 1
            '    lbcFileList.AddItem Trim$(tgCurrCTE(ilLoop).sComment)
            '    lbcFileList.ItemData(lbcFileList.NewIndex) = tgCurrCTE(ilLoop).lCode
            'Next ilLoop
            For ilLoop = 0 To UBound(tgT2MatchList) - 1 Step 1
                lbcFileList.AddItem Trim$(tgT2MatchList(ilLoop).sValue)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgT2MatchList(ilLoop).lValue
            Next ilLoop
        Case "FTYN"
            For ilLoop = 0 To UBound(tgYNMatchList) - 1 Step 1
                lbcFileList.AddItem Trim$(tgYNMatchList(ilLoop).sValue)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgYNMatchList(ilLoop).lValue
            Next ilLoop
        Case "CTE1"
            For ilLoop = 0 To UBound(tgT1MatchList) - 1 Step 1
                lbcFileList.AddItem Trim$(tgT1MatchList(ilLoop).sValue)
                lbcFileList.ItemData(lbcFileList.NewIndex) = tgT1MatchList(ilLoop).lValue
            Next ilLoop
    End Select
    lbcFileList.AddItem "[None]", 0
    lbcFileList.ItemData(lbcFileList.NewIndex) = 0
End Sub

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If edcDropdown.Visible Then
        Select Case grdReplace.Col
            Case FIELDNAMEINDEX
                lbcFieldName.Visible = False
            Case OLDVALUEINDEX
                lbcFileList.Visible = False
            Case NEWVALUEINDEX
                lbcFileList.Visible = False
        End Select
    End If

End Sub

Private Sub mSetFocus()
    Select Case grdReplace.Col
        Case FIELDNAMEINDEX
            edcDropdown.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - cmcDropDown.Width - 30, grdReplace.RowHeight(grdReplace.Row) - 15
            edcDropdown.MaxLength = 0
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcFieldName.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcFieldName, CLng(grdReplace.Height / 2)
            If lbcFieldName.Top + lbcFieldName.Height > cmcCancel.Top Then
                lbcFieldName.Top = edcDropdown.Top - lbcFieldName.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcFieldName.Visible = True
            edcDropdown.SetFocus
        Case OLDVALUEINDEX  'Date
            If (imFieldType = 5) Or (imFieldType = 9) Then  'List
                edcDropdown.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - cmcDropDown.Width - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcFileList.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcFileList, CLng(grdReplace.Height / 2)
                If lbcFileList.Top + lbcFileList.Height > cmcCancel.Top Then
                    lbcFileList.Top = edcDropdown.Top - lbcFileList.Height
                End If
                tmcClick.Enabled = False
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcFileList.Visible = True
                edcDropdown.SetFocus
            Else
                edcGrid.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                edcGrid.Visible = True
                edcGrid.SetFocus
            End If
        Case NEWVALUEINDEX
            If (imFieldType = 5) Or (imFieldType = 9) Then  'List
                edcDropdown.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - cmcDropDown.Width - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcFileList.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcFileList, CLng(grdReplace.Height / 2)
                If lbcFileList.Top + lbcFileList.Height > cmcCancel.Top Then
                    lbcFileList.Top = edcDropdown.Top - lbcFileList.Height
                End If
                tmcClick.Enabled = False
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcFileList.Visible = True
                edcDropdown.SetFocus
            Else
                edcGrid.Move grdReplace.Left + grdReplace.ColPos(grdReplace.Col) + 30, grdReplace.Top + grdReplace.RowPos(grdReplace.Row) + 15, grdReplace.ColWidth(grdReplace.Col) - 30, grdReplace.RowHeight(grdReplace.Row) - 15
                edcGrid.Visible = True
                edcGrid.SetFocus
            End If
    End Select
End Sub
