VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrSchdFilter 
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   9375
   ControlBox      =   0   'False
   Icon            =   "EngrSchdFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9375
   Begin V10EngineeringDev.CSI_Calendar cccDate 
      Height          =   240
      Left            =   4710
      TabIndex        =   8
      Top             =   300
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   423
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin V10EngineeringDev.CSI_TimeLength ltcTime 
      Height          =   195
      Left            =   6210
      TabIndex        =   9
      Top             =   315
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   344
      Text            =   "00:00.0"
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_UseHours    =   0   'False
      CSI_UseTenths   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9285
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   30
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
      ItemData        =   "EngrSchdFilter.frx":030A
      Left            =   5055
      List            =   "EngrSchdFilter.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcOperator 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrSchdFilter.frx":030E
      Left            =   3240
      List            =   "EngrSchdFilter.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   11
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
      Picture         =   "EngrSchdFilter.frx":0312
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
      ItemData        =   "EngrSchdFilter.frx":040C
      Left            =   1665
      List            =   "EngrSchdFilter.frx":040E
      Sorted          =   -1  'True
      TabIndex        =   10
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
      TabIndex        =   13
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
      Picture         =   "EngrSchdFilter.frx":0410
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   5370
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   2145
      TabIndex        =   14
      Top             =   5355
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFilter 
      Height          =   4530
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   7990
      _Version        =   393216
      Cols            =   4
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
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   735
      Picture         =   "EngrSchdFilter.frx":071A
      Top             =   5280
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Schedule Filter"
      Height          =   270
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   1785
   End
End
Attribute VB_Name = "EngrSchdFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrSchdFilter - enters affiliate representative information
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
Const OPINDEX = 1
Const VALUEINDEX = 2
Const CODEINDEX = 3

Private Sub cccDate_CalendarChanged()
    Dim slStr As String
    
    slStr = cccDate.text
    If grdFilter.text <> slStr Then
        imFieldChgd = True
        grdFilter.text = slStr
        grdFilter.CellForeColor = vbBlack
        If gIsDate(slStr) Then
            grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gDateValue(slStr)
        End If
    End If
    mSetCommands
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub




Private Sub mSortCol(ilCol As Integer)
    mSetShow
    gGrid_SortByCol grdFilter, FIELDNAMEINDEX, ilCol, imLastColSorted, imLastSort
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
    
    If (grdFilter.Row >= grdFilter.FixedRows) And (grdFilter.Row < grdFilter.Rows) And (grdFilter.Col >= 0) And (grdFilter.Col < grdFilter.Cols) Then
        lmEnableRow = grdFilter.Row
        lmEnableCol = grdFilter.Col
        sgReturnCallName = grdFilter.TextMatrix(lmEnableRow, FIELDNAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdFilter.Left - pbcArrow.Width - 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + (grdFilter.RowHeight(grdFilter.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If Trim$(grdFilter.TextMatrix(grdFilter.Row, FIELDNAMEINDEX)) <> "" Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdFilter.Col
            Case FIELDNAMEINDEX
                edcDropdown.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - cmcDropDown.Width - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                edcDropdown.MaxLength = 0
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcFieldName.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcFieldName, CLng(grdFilter.Height / 2)
                If lbcFieldName.Top + lbcFieldName.Height > cmcCancel.Top Then
                    lbcFieldName.Top = edcDropdown.Top - lbcFieldName.Height
                End If
                slStr = grdFilter.text
                ilIndex = gListBoxFind(lbcFieldName, slStr)
                If ilIndex >= 0 Then
                    lbcFieldName.ListIndex = ilIndex
                    edcDropdown.text = lbcFieldName.List(lbcFieldName.ListIndex)
                Else
                    edcDropdown.text = ""
                    lbcFieldName.ListIndex = -1
                End If
                tmcClick.Enabled = False
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcFieldName.Visible = True
                edcDropdown.SetFocus
            Case OPINDEX  'Date
                slStr = Trim$(grdFilter.TextMatrix(grdFilter.Row, FIELDNAMEINDEX))
                If slStr <> "" Then
                    llListIndex = gListBoxFind(lbcFieldName, slStr)
                    If llListIndex >= 0 Then
                        lbcOperator.Clear
                        ilIndex = lbcFieldName.ItemData(llListIndex)
                        If (tgFilterFields(ilIndex).iFieldType = 5) Or (tgFilterFields(ilIndex).iFieldType = 2) Or (tgFilterFields(ilIndex).iFieldType = 9) Then   'List
                            lbcOperator.AddItem "Equal to"
                            lbcOperator.ItemData(lbcOperator.NewIndex) = 1
                            lbcOperator.AddItem "Not Equal to"
                            lbcOperator.ItemData(lbcOperator.NewIndex) = 2
                        Else
                            lbcOperator.AddItem "Equal to"
                            lbcOperator.ItemData(lbcOperator.NewIndex) = 1
                            lbcOperator.AddItem "Not Equal to"
                            lbcOperator.ItemData(lbcOperator.NewIndex) = 2
                            lbcOperator.AddItem "Greater than"
                            lbcOperator.ItemData(lbcOperator.NewIndex) = 3
                            lbcOperator.AddItem "Less than"
                            lbcOperator.ItemData(lbcOperator.NewIndex) = 4
                            lbcOperator.AddItem "Greater than or Equal to"
                            lbcOperator.ItemData(lbcOperator.NewIndex) = 5
                            lbcOperator.AddItem "Less than or Equal to"
                            lbcOperator.ItemData(lbcOperator.NewIndex) = 6
                        End If
                    End If
                End If
                edcDropdown.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - cmcDropDown.Width - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                edcDropdown.MaxLength = 0
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcOperator.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcOperator, CLng(grdFilter.Height / 2)
                If lbcOperator.Top + lbcOperator.Height > cmcCancel.Top Then
                    lbcOperator.Top = edcDropdown.Top - lbcOperator.Height
                End If
                slStr = grdFilter.text
                ilIndex = gListBoxFind(lbcOperator, slStr)
                If ilIndex >= 0 Then
                    lbcOperator.ListIndex = ilIndex
                    edcDropdown.text = lbcOperator.List(lbcOperator.ListIndex)
                Else
                    edcDropdown.text = ""
                    lbcOperator.ListIndex = gListBoxFind(lbcOperator, "Equal to")
                    If lbcOperator.ListIndex >= 0 Then
                        edcDropdown.text = lbcOperator.List(lbcOperator.ListIndex)
                    End If
                End If
                tmcClick.Enabled = False
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcOperator.Visible = True
                edcDropdown.SetFocus
            Case VALUEINDEX
                slStr = Trim$(grdFilter.TextMatrix(grdFilter.Row, FIELDNAMEINDEX))
                If slStr <> "" Then
                    llListIndex = gListBoxFind(lbcFieldName, slStr)
                    If llListIndex >= 0 Then
                        ilIndex = lbcFieldName.ItemData(llListIndex)
                        imFieldType = tgFilterFields(ilIndex).iFieldType
                        If (imFieldType = 5) Or (imFieldType = 9) Then  'List
                            mPopList tgFilterFields(ilIndex).sListFile
                            edcDropdown.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - cmcDropDown.Width - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                            edcDropdown.MaxLength = tgFilterFields(ilIndex).iMaxNoChar
                            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                            lbcFileList.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                            gSetListBoxHeight lbcFileList, CLng(grdFilter.Height / 2)
                            If lbcFileList.Top + lbcFileList.Height > cmcCancel.Top Then
                                lbcFileList.Top = edcDropdown.Top - lbcFileList.Height
                            End If
                            slStr = grdFilter.text
                            ilIndex = gListBoxFind(lbcFileList, slStr)
                            If ilIndex >= 0 Then
                                lbcFileList.ListIndex = ilIndex
                                edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
                            Else
                                edcDropdown.text = ""
                                lbcFileList.ListIndex = -1
                            End If
                            tmcClick.Enabled = False
                            edcDropdown.Visible = True
                            cmcDropDown.Visible = True
                            lbcFileList.Visible = True
                            edcDropdown.SetFocus
                        ElseIf imFieldType = 3 Then 'Date
                            cccDate.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                            slStr = grdFilter.text
                            If Not gIsDate(slStr) Then
                                cccDate.text = ""
                            Else
                                cccDate.text = ""
                                cccDate.text = slStr 'grdLibEvents.Text
                            End If
                            cccDate.text = grdFilter.text
                            cccDate.Visible = True
                            cccDate.SetFocus
                        ElseIf imFieldType = 4 Then 'time
                            slStr = grdFilter.text
                            ltcTime.CSI_UseHours = True
                            ltcTime.CSI_UseTenths = False
                            ltcTime.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                            'ltcGrid.MaxLength = tgFilterFields(ilIndex).iMaxNoChar
                            If Not gIsTime(slStr) Then
                                ltcTime.text = ""
                            Else
                                ltcTime.text = ""
                                ltcTime.text = slStr 'grdLibEvents.Text
                            End If
                            ltcTime.Visible = True
                            ltcTime.SetFocus
                        ElseIf imFieldType = 6 Then 'time in tenths
                            slStr = grdFilter.text
                            ltcTime.CSI_UseHours = True
                            ltcTime.CSI_UseTenths = True
                            ltcTime.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                            'ltcGrid.MaxLength = tgFilterFields(ilIndex).iMaxNoChar
                            If Not gIsTimeTenths(slStr) Then
                                ltcTime.text = ""
                            Else
                                ltcTime.text = ""
                                ltcTime.text = slStr 'grdLibEvents.Text
                            End If
                            ltcTime.Visible = True
                            ltcTime.SetFocus
                        ElseIf imFieldType = 7 Then 'length
                            slStr = grdFilter.text
                            ltcTime.CSI_UseHours = False
                            ltcTime.CSI_UseTenths = False
                            ltcTime.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                            'ltcGrid.MaxLength = tgFilterFields(ilIndex).iMaxNoChar
                            If Not gIsLength(slStr) Then
                                ltcTime.text = ""
                            Else
                                ltcTime.text = ""
                                ltcTime.text = slStr 'grdLibEvents.Text
                            End If
                            ltcTime.Visible = True
                            ltcTime.SetFocus
                        ElseIf imFieldType = 8 Then 'Length in Tenths
                            slStr = grdFilter.text
                            ltcTime.CSI_UseHours = False
                            ltcTime.CSI_UseTenths = True
                            ltcTime.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                            'ltcGrid.MaxLength = tgFilterFields(ilIndex).iMaxNoChar
                            If Not gIsLengthTenths(slStr) Then
                                ltcTime.text = ""
                            Else
                                ltcTime.text = ""
                                ltcTime.text = slStr 'grdLibEvents.Text
                            End If
                            ltcTime.Visible = True
                            ltcTime.SetFocus
                        Else
                            edcGrid.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                            edcGrid.MaxLength = tgFilterFields(ilIndex).iMaxNoChar
                            edcGrid.text = grdFilter.text
                            edcGrid.Visible = True
                            edcGrid.SetFocus
                        End If
                    End If
                End If
                
        End Select
    End If
End Sub
Private Sub mSetShow()
    If (lmEnableRow >= grdFilter.FixedRows) And (lmEnableRow < grdFilter.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case FIELDNAMEINDEX
            Case OPINDEX
            Case VALUEINDEX
        End Select
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    ltcTime.Visible = False
    cccDate.Visible = False
    edcGrid.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcFieldName.Visible = False
    lbcOperator.Visible = False
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
    
    grdFilter.Redraw = False
    'Test if fields defined
    If ilTestState Then
        lbcOperator.Clear
        lbcOperator.AddItem "Equal to"
        lbcOperator.ItemData(lbcOperator.NewIndex) = 1
        lbcOperator.AddItem "Not Equal to"
        lbcOperator.ItemData(lbcOperator.NewIndex) = 2
        lbcOperator.AddItem "Greater than"
        lbcOperator.ItemData(lbcOperator.NewIndex) = 3
        lbcOperator.AddItem "Less than"
        lbcOperator.ItemData(lbcOperator.NewIndex) = 4
        lbcOperator.AddItem "Greater than or Equal to"
        lbcOperator.ItemData(lbcOperator.NewIndex) = 5
        lbcOperator.AddItem "Less than or Equal to"
        lbcOperator.ItemData(lbcOperator.NewIndex) = 6
    End If
    ilError = False
    For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
        slStr = Trim$(grdFilter.TextMatrix(llRow, FIELDNAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdFilter.TextMatrix(llRow, OPINDEX)
            If slStr <> "" Then
                ilError = True
                grdFilter.TextMatrix(llRow, FIELDNAMEINDEX) = "Missing"
                grdFilter.Row = llRow
                grdFilter.Col = FIELDNAMEINDEX
                grdFilter.CellForeColor = vbRed
            End If
        Else
            If ilTestState Then
                slStr = Trim$(grdFilter.TextMatrix(llRow, OPINDEX))
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdFilter.TextMatrix(llRow, OPINDEX) = "Missing"
                    grdFilter.Row = llRow
                    grdFilter.Col = OPINDEX
                    grdFilter.CellForeColor = vbRed
                Else
                    llListIndex = gListBoxFind(lbcOperator, slStr)
                    If llListIndex < 0 Then
                        ilError = True
                        grdFilter.Row = llRow
                        grdFilter.Col = OPINDEX
                        grdFilter.CellForeColor = vbRed
                    End If
                End If
                slStr = Trim$(grdFilter.TextMatrix(llRow, VALUEINDEX))
                If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                    ilError = True
                    grdFilter.TextMatrix(llRow, VALUEINDEX) = "Missing"
                    grdFilter.Row = llRow
                    grdFilter.Col = VALUEINDEX
                    grdFilter.CellForeColor = vbRed
                Else
                    slFieldName = Trim$(grdFilter.TextMatrix(llRow, FIELDNAMEINDEX))
                    llListIndex = gListBoxFind(lbcFieldName, slFieldName)
                    If llListIndex >= 0 Then
                        ilIndex = lbcFieldName.ItemData(llListIndex)
                        imFieldType = tgFilterFields(ilIndex).iFieldType
                        If imFieldType = 3 Then
                            If Not gIsDate(slStr) Then
                                ilError = True
                                grdFilter.Row = llRow
                                grdFilter.Col = VALUEINDEX
                                grdFilter.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 4 Then
                            If Not gIsTime(slStr) Then
                                ilError = True
                                grdFilter.Row = llRow
                                grdFilter.Col = VALUEINDEX
                                grdFilter.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 6 Then
                            If Not gIsTimeTenths(slStr) Then
                                ilError = True
                                grdFilter.Row = llRow
                                grdFilter.Col = VALUEINDEX
                                grdFilter.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 7 Then
                            If Not gIsLength(slStr) Then
                                ilError = True
                                grdFilter.Row = llRow
                                grdFilter.Col = VALUEINDEX
                                grdFilter.CellForeColor = vbRed
                            End If
                        ElseIf imFieldType = 8 Then
                            If Not gIsLengthTenths(slStr) Then
                                ilError = True
                                grdFilter.Row = llRow
                                grdFilter.Col = VALUEINDEX
                                grdFilter.CellForeColor = vbRed
                            End If
                        End If
                    Else
                        ilError = True
                        grdFilter.Row = llRow
                        grdFilter.Col = FIELDNAMEINDEX
                        grdFilter.CellForeColor = vbRed
                    End If
                End If
            End If
        End If
    Next llRow
    grdFilter.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdFilter
    mGridColumnWidth
    'Set Titles
    grdFilter.TextMatrix(0, FIELDNAMEINDEX) = "Field Name"
    grdFilter.TextMatrix(0, OPINDEX) = "Operator"
    grdFilter.TextMatrix(0, VALUEINDEX) = "Filter Value"
    grdFilter.Row = 1
    For ilCol = 0 To grdFilter.Cols - 1 Step 1
        grdFilter.Col = ilCol
        grdFilter.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdFilter.Height = cmcCancel.Top - grdFilter.Top - 120    '8 * grdFilter.RowHeight(0) + 30
    gGrid_IntegralHeight grdFilter
    gGrid_Clear grdFilter, True
    grdFilter.Row = grdFilter.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdFilter.Width = EngrSchdFilter.Width - 2 * grdFilter.Left
    grdFilter.ColWidth(CODEINDEX) = 0
    grdFilter.ColWidth(FIELDNAMEINDEX) = grdFilter.Width / 4
    grdFilter.ColWidth(OPINDEX) = grdFilter.Width / 4
    grdFilter.ColWidth(VALUEINDEX) = grdFilter.Width - GRIDSCROLLWIDTH
    For ilCol = FIELDNAMEINDEX To VALUEINDEX Step 1
        If ilCol <> VALUEINDEX Then
            If grdFilter.ColWidth(VALUEINDEX) > grdFilter.ColWidth(ilCol) Then
                grdFilter.ColWidth(VALUEINDEX) = grdFilter.ColWidth(VALUEINDEX) - grdFilter.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    edcGrid.text = ""
    edcDropdown.text = ""
    
    gGrid_Clear grdFilter, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec()
    Dim ilRow As Integer
    Dim slStr As String
    Dim ilUpper As Integer
    Dim llListIndex As Long
    
    ReDim tgFilterValues(0 To 0) As FILTERVALUES
    lbcOperator.Clear
    lbcOperator.AddItem "Equal to"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 1
    lbcOperator.AddItem "Not Equal to"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 2
    lbcOperator.AddItem "Greater than"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 3
    lbcOperator.AddItem "Less than"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 4
    lbcOperator.AddItem "Greater than or Equal to"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 5
    lbcOperator.AddItem "Less than or Equal to"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 6
    For ilRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
        slStr = Trim$(grdFilter.TextMatrix(ilRow, FIELDNAMEINDEX))
        If slStr <> "" Then
            ilUpper = UBound(tgFilterValues)
            tgFilterValues(ilUpper).sFieldName = slStr
            slStr = Trim$(grdFilter.TextMatrix(ilRow, OPINDEX))
            llListIndex = gListBoxFind(lbcOperator, slStr)
            If llListIndex >= 0 Then
                tgFilterValues(ilUpper).iOperator = lbcOperator.ItemData(llListIndex)
                slStr = Trim$(grdFilter.TextMatrix(ilRow, VALUEINDEX))
                tgFilterValues(ilUpper).sValue = slStr
                slStr = Trim$(grdFilter.TextMatrix(ilRow, CODEINDEX))
                tgFilterValues(ilUpper).lCode = Val(slStr)
                ReDim Preserve tgFilterValues(0 To ilUpper + 1) As FILTERVALUES
            End If
        End If
    Next ilRow
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim ilLoop As Integer
    Dim ilList As Integer
    Dim slStr As String
    Dim llListRow As Long

    'gGrid_Clear grdFilter, True
    llRow = grdFilter.FixedRows
    For ilLoop = 0 To UBound(tgFilterValues) - 1 Step 1
        If llRow + 1 > grdFilter.Rows Then
            grdFilter.AddItem ""
        End If
        grdFilter.Row = llRow
        slStr = Trim$(tgFilterValues(ilLoop).sFieldName)
        grdFilter.TextMatrix(llRow, FIELDNAMEINDEX) = Trim$(tgFilterValues(ilLoop).sFieldName)
        For ilList = 0 To lbcOperator.ListCount - 1 Step 1
            If lbcOperator.ItemData(ilList) = tgFilterValues(ilLoop).iOperator Then
                grdFilter.TextMatrix(llRow, OPINDEX) = Trim$(lbcOperator.List(ilList))
                Exit For
            End If
        Next ilList
        grdFilter.TextMatrix(llRow, VALUEINDEX) = Trim$(tgFilterValues(ilLoop).sValue)
        grdFilter.TextMatrix(llRow, CODEINDEX) = Trim$(tgFilterValues(ilLoop).lCode)
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdFilter.Rows Then
        grdFilter.AddItem ""
    End If
    grdFilter.Redraw = True
End Sub
Private Sub cmcCancel_Click()
    igAnsFilter = CALLCANCELLED
    Unload EngrSchdFilter
End Sub

Private Sub cmcClear_Click()
    edcGrid.text = ""
    edcDropdown.text = ""
    lbcFileList.Clear
    gGrid_Clear grdFilter, True
    imFieldChgd = True
End Sub

Private Sub cmcClear_GotFocus()
    mSetShow
End Sub

Private Sub cmcDone_Click()
    
    Dim ilRet As Integer
    If imFieldChgd = False Then
        igAnsFilter = CALLDONE
        Unload EngrSchdFilter
        Exit Sub
    End If
    If mCheckFields(True) Then
        mMoveCtrlsToRec
        igAnsFilter = CALLDONE
        Unload EngrSchdFilter
        Exit Sub
    End If
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub



Private Sub cmcDropDown_Click()
    Select Case grdFilter.Col
        Case FIELDNAMEINDEX
            lbcFieldName.Visible = Not lbcFieldName.Visible
        Case OPINDEX
            lbcOperator.Visible = Not lbcOperator.Visible
        Case VALUEINDEX
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
    Select Case grdFilter.Col
        Case FIELDNAMEINDEX
            llRow = gListBoxFind(lbcFieldName, slStr)
            If llRow >= 0 Then
                lbcFieldName.ListIndex = llRow
                edcDropdown.text = lbcFieldName.List(lbcFieldName.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case OPINDEX
            llRow = gListBoxFind(lbcOperator, slStr)
            If llRow >= 0 Then
                lbcOperator.ListIndex = llRow
                edcDropdown.text = lbcOperator.List(lbcOperator.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
        Case VALUEINDEX
            llRow = gListBoxFind(lbcFileList, slStr)
            If llRow >= 0 Then
                lbcFileList.ListIndex = llRow
                edcDropdown.text = lbcFileList.List(lbcFileList.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            Else
                lbcFileList.ListIndex = -1
                edcDropdown.text = ""
            End If
    End Select
    If (StrComp(grdFilter.text, edcDropdown.text, vbTextCompare) <> 0) Then
        imFieldChgd = True
        Select Case grdFilter.Col
            Case FIELDNAMEINDEX
            Case OPINDEX
            Case VALUEINDEX
                If lbcFileList.ListIndex >= 0 Then
                    grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = lbcFileList.ItemData(lbcFileList.ListIndex)
                End If
        End Select
        grdFilter.text = edcDropdown.text
        grdFilter.CellForeColor = vbBlack
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
        Select Case grdFilter.Col
            Case FIELDNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcFieldName, True
            Case OPINDEX
                gProcessArrowKey Shift, KeyCode, lbcOperator, True
            Case VALUEINDEX
                gProcessArrowKey Shift, KeyCode, lbcFileList, True
        End Select
        tmcClick.Enabled = False
    End If
End Sub

Private Sub edcGrid_Change()
    Dim slStr As String
    
    Select Case grdFilter.Col
        Case FIELDNAMEINDEX
        Case OPINDEX
        Case VALUEINDEX
            If grdFilter.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdFilter.text = edcGrid.text
            grdFilter.CellForeColor = vbBlack
            slStr = edcGrid.text
            If imFieldType = 1 Then
                grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = slStr
            ElseIf imFieldType = 3 Then
                If gIsDate(slStr) Then
                    grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gDateValue(slStr)
                End If
            ElseIf imFieldType = 4 Then
                If gIsTime(slStr) Then
                    grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gTimeToLong(slStr, False)
                End If
            ElseIf imFieldType = 6 Then
                If gIsTimeTenths(slStr) Then
                    grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gStrTimeInTenthToLong(slStr, False)
                End If
            ElseIf imFieldType = 7 Then
                If gIsLength(slStr) Then
                    grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gLengthToLong(slStr)
                End If
            ElseIf imFieldType = 8 Then
                If gIsLengthTenths(slStr) Then
                    grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gStrLengthInTenthToLong(slStr)
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
        mSetToFirstBlankRow
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
    gSetFonts EngrSchdFilter
    gCenterFormModal EngrSchdFilter
End Sub

Private Sub Form_Load()
    mGridColumns
    mInit
End Sub

Private Sub Form_Resize()
    'These call are here and in form_Active (call to mGridColumns)
    'They are in mGridColumn in case the For_Initialize size chage does not cause a resize event
    mGridColumnWidth
    grdFilter.Height = cmcCancel.Top - grdFilter.Top - 120    '8 * grdFilter.RowHeight(0) + 30
    gGrid_IntegralHeight grdFilter
    gGrid_FillWithRows grdFilter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Erase lmDeleteCodes
    Set EngrSchdFilter = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdFilter, grdFilter, vbHourglass
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imIgnoreScroll = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    imFirstActivate = True
    imInChg = True
    lbcOperator.AddItem "Equal to"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 1
    lbcOperator.AddItem "Not Equal to"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 2
    lbcOperator.AddItem "Greater than"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 3
    lbcOperator.AddItem "Less than"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 4
    lbcOperator.AddItem "Greater than or Equal to"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 5
    lbcOperator.AddItem "Less than or Equal to"
    lbcOperator.ItemData(lbcOperator.NewIndex) = 6
    mPopulate
    mMoveRecToCtrls
    mSortCol 0
    imInChg = False
    imFieldChgd = False
    'mSetToFirstBlankRow
    mSetCommands
    gSetMousePointer grdFilter, grdFilter, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdFilter, grdFilter, vbDefault
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

Private Sub lbcOperator_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcOperator.List(lbcOperator.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub grdFilter_Click()
    If grdFilter.Col >= grdFilter.Cols Then
        Exit Sub
    End If
End Sub

Private Sub grdFilter_EnterCell()
    mSetShow
End Sub

Private Sub grdFilter_GotFocus()
    If grdFilter.Col >= grdFilter.Cols Then
        Exit Sub
    End If
End Sub

Private Sub grdFilter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdFilter.TopRow
    grdFilter.Redraw = False
End Sub

Private Sub grdFilter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdFilter.RowHeight(0) Then
        mSortCol grdFilter.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdFilter, x, y)
    If Not ilFound Then
        grdFilter.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdFilter.Col >= grdFilter.Cols Then
        grdFilter.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdFilter.TopRow
    DoEvents
    llRow = grdFilter.Row
    If grdFilter.TextMatrix(llRow, FIELDNAMEINDEX) = "" Then
        grdFilter.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdFilter.TextMatrix(llRow, FIELDNAMEINDEX) = ""
        grdFilter.Row = llRow + 1
        grdFilter.Col = FIELDNAMEINDEX
        grdFilter.Redraw = True
    End If
    grdFilter.Redraw = True
    mEnableBox
End Sub

Private Sub grdFilter_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdFilter.Redraw = False Then
        grdFilter.Redraw = True
        If lmTopRow < grdFilter.FixedRows Then
            grdFilter.TopRow = grdFilter.FixedRows
        Else
            grdFilter.TopRow = lmTopRow
        End If
        grdFilter.Refresh
        grdFilter.Redraw = False
    End If
    If (imShowGridBox) And (grdFilter.Row >= grdFilter.FixedRows) And (grdFilter.Col >= 0) And (grdFilter.Col < grdFilter.Cols) Then
        If grdFilter.RowIsVisible(grdFilter.Row) Then
            'edcGrid.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 30, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 30
            pbcArrow.Move grdFilter.Left - pbcArrow.Width - 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + (grdFilter.RowHeight(grdFilter.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            pbcArrow.Visible = False
            ltcTime.Visible = False
            cccDate.Visible = False
            edcGrid.Visible = False
            edcDropdown.Visible = False
            cmcDropDown.Visible = False
            lbcFieldName.Visible = False
            lbcOperator.Visible = False
            lbcFileList.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub

Private Sub ltcTime_OnChange()
    Dim slStr As String
    
    slStr = ltcTime.text
    If grdFilter.text <> slStr Then
        imFieldChgd = True
        grdFilter.text = slStr
        grdFilter.CellForeColor = vbBlack
        If imFieldType = 4 Then
            If gIsTime(slStr) Then
                grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gTimeToLong(slStr, False)
            End If
        ElseIf imFieldType = 6 Then
            If gIsTimeTenths(slStr) Then
                grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gStrTimeInTenthToLong(slStr, False)
            End If
        ElseIf imFieldType = 7 Then
            If gIsLength(slStr) Then
                grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gLengthToLong(slStr)
            End If
        ElseIf imFieldType = 8 Then
            If gIsLengthTenths(slStr) Then
                grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = gStrLengthInTenthToLong(slStr)
            End If
        End If
    End If
    mSetCommands

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
    If edcGrid.Visible Or edcDropdown.Visible Or cccDate.Visible Or ltcTime.Visible Then
        If grdFilter.Col = VALUEINDEX Then
            slStr = edcGrid.text
            If imFieldType = 3 Then
                slStr = cccDate.text
                If Not gIsDate(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    cccDate.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 4 Then
                slStr = ltcTime.text
                If Not gIsTime(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    ltcTime.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 6 Then
                slStr = ltcTime.text
                If Not gIsTimeTenths(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    ltcTime.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 7 Then
                slStr = ltcTime.text
                If Not gIsLength(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    ltcTime.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 8 Then
                slStr = ltcTime.text
                If Not gIsLengthTenths(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    ltcTime.SetFocus
                    Exit Sub
                End If
            End If
        End If
        mSetShow
        If grdFilter.Col = FIELDNAMEINDEX Then
            If grdFilter.Row > grdFilter.FixedRows Then
                lmTopRow = -1
                grdFilter.Row = grdFilter.Row - 1
                If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                    grdFilter.TopRow = grdFilter.TopRow - 1
                End If
                grdFilter.Col = VALUEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdFilter.Col = grdFilter.Col - 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdFilter.TopRow = grdFilter.FixedRows
        grdFilter.Col = FIELDNAMEINDEX
        grdFilter.Row = grdFilter.FixedRows
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
    If edcGrid.Visible Or edcDropdown.Visible Or cccDate.Visible Or ltcTime.Visible Then
        If grdFilter.Col = VALUEINDEX Then
            slStr = edcGrid.text
            If imFieldType = 3 Then
                slStr = cccDate.text
                If Not gIsDate(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    cccDate.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 4 Then
                slStr = ltcTime.text
                If Not gIsTime(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    ltcTime.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 6 Then
                slStr = ltcTime.text
                If Not gIsTimeTenths(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    ltcTime.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 7 Then
                slStr = ltcTime.text
                If Not gIsLength(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    ltcTime.SetFocus
                    Exit Sub
                End If
            ElseIf imFieldType = 8 Then
                slStr = ltcTime.text
                If Not gIsLengthTenths(slStr) Then
                    Beep
                    'edcGrid.SetFocus
                    ltcTime.SetFocus
                    Exit Sub
                End If
            End If
        End If
        llEnableRow = lmEnableRow
        mSetShow
        If grdFilter.Col = VALUEINDEX Then
            llRow = grdFilter.Rows
            Do
                llRow = llRow - 1
            Loop While grdFilter.TextMatrix(llRow, FIELDNAMEINDEX) = ""
            llRow = llRow + 1
            If (grdFilter.Row + 1 < llRow) Then
                lmTopRow = -1
                grdFilter.Row = grdFilter.Row + 1
                If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                    imIgnoreScroll = True
                    grdFilter.TopRow = grdFilter.TopRow + 1
                End If
                grdFilter.Col = FIELDNAMEINDEX
                'grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = 0
                If Trim$(grdFilter.TextMatrix(grdFilter.Row, FIELDNAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdFilter.Left - pbcArrow.Width - 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + (grdFilter.RowHeight(grdFilter.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdFilter.TextMatrix(llEnableRow, FIELDNAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdFilter.Row + 1 >= grdFilter.Rows Then
                        grdFilter.AddItem ""
                    End If
                    grdFilter.Row = grdFilter.Row + 1
                    If Not grdFilter.RowIsVisible(grdFilter.Row) Then
                        imIgnoreScroll = True
                        grdFilter.TopRow = grdFilter.TopRow + 1
                    End If
                    grdFilter.Col = FIELDNAMEINDEX
                    'grdFilter.TextMatrix(grdFilter.Row, CODEINDEX) = 0
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdFilter.Left - pbcArrow.Width - 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + (grdFilter.RowHeight(grdFilter.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdFilter.Col = grdFilter.Col + 1
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdFilter.TopRow = grdFilter.FixedRows
        grdFilter.Col = FIELDNAMEINDEX
        grdFilter.Row = grdFilter.FixedRows
        mEnableBox
    End If
End Sub



Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdFilter.TopRow
    llRow = grdFilter.Row
    slMsg = "Delete " & Trim$(grdFilter.TextMatrix(llRow, FIELDNAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdFilter.Redraw = False
    grdFilter.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdFilter.AddItem ""
    grdFilter.Redraw = False
    grdFilter.TopRow = llTRow
    grdFilter.Redraw = True
    DoEvents
    mSetCommands
    pbcClickFocus.SetFocus
    mDeleteRow = True
End Function





Private Sub mPopulate()
    Dim ilRow As Integer
    Dim ilLoop As Integer
    
    For ilLoop = 0 To UBound(tgFilterFields) - 1 Step 1
        lbcFieldName.AddItem Trim$(tgFilterFields(ilLoop).sFieldName)
        lbcFieldName.ItemData(lbcFieldName.NewIndex) = ilLoop
    Next ilLoop
End Sub

Private Sub mPopList(slFileName As String)
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

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If edcDropdown.Visible Then
        Select Case grdFilter.Col
            Case FIELDNAMEINDEX
                lbcFieldName.Visible = False
            Case OPINDEX
                lbcOperator.Visible = False
            Case VALUEINDEX
                lbcFileList.Visible = False
        End Select
    End If

End Sub

Private Sub mSetFocus()
    Select Case grdFilter.Col
        Case FIELDNAMEINDEX
            edcDropdown.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - cmcDropDown.Width - 30, grdFilter.RowHeight(grdFilter.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcFieldName.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcFieldName, CLng(grdFilter.Height / 2)
            If lbcFieldName.Top + lbcFieldName.Height > cmcCancel.Top Then
                lbcFieldName.Top = edcDropdown.Top - lbcFieldName.Height
            End If
            tmcClick.Enabled = False
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcFieldName.Visible = True
            edcDropdown.SetFocus
        Case OPINDEX  'Date
            edcDropdown.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - cmcDropDown.Width - 30, grdFilter.RowHeight(grdFilter.Row) - 15
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcOperator.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcOperator, CLng(grdFilter.Height / 2)
            If lbcOperator.Top + lbcOperator.Height > cmcCancel.Top Then
                lbcOperator.Top = edcDropdown.Top - lbcOperator.Height
            End If
            tmcClick.Enabled = False
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            lbcOperator.Visible = True
            edcDropdown.SetFocus
        Case VALUEINDEX
            If (imFieldType = 5) Or (imFieldType = 9) Then  'List
                edcDropdown.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - cmcDropDown.Width - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcFileList.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcFileList, CLng(grdFilter.Height / 2)
                If lbcFileList.Top + lbcFileList.Height > cmcCancel.Top Then
                    lbcFileList.Top = edcDropdown.Top - lbcFileList.Height
                End If
                tmcClick.Enabled = False
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcFileList.Visible = True
                edcDropdown.SetFocus
            ElseIf (imFieldType = 3) Then
                cccDate.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                cccDate.Visible = True
                cccDate.SetFocus
            ElseIf (imFieldType = 4) Or (imFieldType = 6) Or (imFieldType = 7) Or (imFieldType = 8) Then
                ltcTime.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                ltcTime.Visible = True
                ltcTime.SetFocus
            Else
                edcGrid.Move grdFilter.Left + grdFilter.ColPos(grdFilter.Col) + 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + 15, grdFilter.ColWidth(grdFilter.Col) - 30, grdFilter.RowHeight(grdFilter.Row) - 15
                edcGrid.Visible = True
                edcGrid.SetFocus
            End If
    End Select
End Sub
Private Sub mSetToFirstBlankRow()
    Dim llRow As Long
    Dim slStr As String
    
    'Find first blank row
    For llRow = grdFilter.FixedRows To grdFilter.Rows - 1 Step 1
        slStr = Trim$(grdFilter.TextMatrix(llRow, FIELDNAMEINDEX))
        If (slStr = "") Then
            grdFilter.Row = llRow
            Do While Not grdFilter.RowIsVisible(grdFilter.Row)
                imIgnoreScroll = True
                grdFilter.TopRow = grdFilter.TopRow + 1
            Loop
            grdFilter.Col = FIELDNAMEINDEX
            'mEnableBox
            imFromArrow = True
            pbcArrow.Move grdFilter.Left - pbcArrow.Width - 30, grdFilter.Top + grdFilter.RowPos(grdFilter.Row) + (grdFilter.RowHeight(grdFilter.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            pbcArrow.SetFocus
            Exit Sub
        End If
    Next llRow

End Sub
