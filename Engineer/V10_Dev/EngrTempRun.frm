VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Begin VB.Form EngrTempRun 
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11790
   ControlBox      =   0   'False
   Icon            =   "EngrTempRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   Begin V10EngineeringDev.CSI_Calendar_UP cccAirDate_Up 
      Height          =   1995
      Left            =   7485
      TabIndex        =   11
      Top             =   45
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   3519
      Text            =   "6/14/12"
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
   Begin V10EngineeringDev.CSI_TimeLength ltcAirTime 
      Height          =   195
      Left            =   8955
      TabIndex        =   12
      Top             =   345
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   344
      Text            =   "00:00:00"
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_UseHours    =   -1  'True
      CSI_UseTenths   =   0   'False
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
   Begin V10EngineeringDev.CSI_Calendar cccAirDate 
      Height          =   225
      Left            =   5925
      TabIndex        =   10
      Top             =   1740
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   397
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
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   11295
      Top             =   6570
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1590
      TabIndex        =   6
      Top             =   1650
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
      Left            =   2535
      Picture         =   "EngrTempRun.frx":030A
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ListBox lbcBDE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "EngrTempRun.frx":0404
      Left            =   3390
      List            =   "EngrTempRun.frx":0406
      Sorted          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   11700
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   16
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
      TabIndex        =   9
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
      TabIndex        =   13
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
      Picture         =   "EngrTempRun.frx":0408
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   90
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
      Left            =   6075
      TabIndex        =   15
      Top             =   6705
      Width           =   1335
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   4290
      TabIndex        =   14
      Top             =   6705
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTempRun 
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
      Cols            =   7
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
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2520
      Picture         =   "EngrTempRun.frx":0712
      Top             =   6615
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Template Airing Information"
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
      Picture         =   "EngrTempRun.frx":0A1C
      Top             =   6615
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8775
      Picture         =   "EngrTempRun.frx":12E6
      Top             =   6615
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "EngrTempRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'******************************************************
'*  EngrTempRun - enters affiliate representative information
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
Private imMaxColChars As Integer
Private imDoubleClickName As Integer
Private lmCharacterWidth As Long

Private smESCValue As String    'Value used if ESC pressed

Private imDeleteCodes() As Integer
Private tmTSE As TSE

'Grid Controls
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private imShowGridBox As Integer
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imLastColSorted As Integer
Private imLastSort As Integer

Const BUSNAMEINDEX = 0
Const DATEINDEX = 1
Const TIMEINDEX = 2
Const DESCRIPTIONINDEX = 3
Const STATEINDEX = 4
Const AIRINFOINDEX = 5
Const SORTINDEX = 6

Private Sub cccAirDate_Change()
    If StrComp(Trim$(grdTempRun.text), Trim$(cccAirDate.text), vbTextCompare) <> 0 Then
        grdTempRun.text = cccAirDate.text
        grdTempRun.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub cccAirDate_Up_Change()
    If StrComp(Trim$(grdTempRun.text), Trim$(cccAirDate_Up.text), vbTextCompare) <> 0 Then
        grdTempRun.text = cccAirDate_Up.text
        grdTempRun.CellForeColor = vbBlack
    End If
    mSetCommands
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
End Sub

Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slBus As String
    Dim slDate As String
    Dim slTime As String
    Dim ilLen As Integer
    
    mSetShow
    lmEnableRow = -1
    lmEnableCol = -1
    For llRow = grdTempRun.FixedRows To grdTempRun.Rows - 1 Step 1
        slStr = Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX))
        If slStr <> "" Then
            slStr = grdTempRun.TextMatrix(llRow, DATEINDEX)
            slDate = Trim$(Str$(gDateValue(slStr)))
            Do While Len(slDate) < 6
                slDate = "0" & slDate
            Loop
            slStr = grdTempRun.TextMatrix(llRow, TIMEINDEX)
            slTime = Trim$(Str$(gStrTimeInTenthToLong(slStr, False)))
            Do While Len(slTime) < 8
                slTime = "0" & slTime
            Loop
            slBus = grdTempRun.TextMatrix(llRow, BUSNAMEINDEX)
            ilLen = gSetMaxChars("BusName", 0)
            Do While Len(slBus) < ilLen
                slBus = slBus & " "
            Loop
            grdTempRun.TextMatrix(llRow, SORTINDEX) = slDate & slTime & slBus
        End If
    Next llRow
    gGrid_SortByCol grdTempRun, BUSNAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
End Sub

Private Sub mSetCommands()
    Dim ilRet As Integer
    If imInChg Then
        Exit Sub
    End If
    If cmcDone.Enabled = False Then
        Exit Sub
    End If
End Sub

Private Sub mEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) <> 0) And (StrComp(sgUserName, "Guide", vbTextCompare) <> 0) And (igListStatus(RELAYLIST) <> 2) Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    If (grdTempRun.Row >= grdTempRun.FixedRows) And (grdTempRun.Row < grdTempRun.Rows) And (grdTempRun.Col >= 0) And (grdTempRun.Col < grdTempRun.Cols - 1) Then
        lmEnableRow = grdTempRun.Row
        lmEnableCol = grdTempRun.Col
        sgReturnCallName = grdTempRun.TextMatrix(lmEnableRow, BUSNAMEINDEX)
        imShowGridBox = True
        pbcArrow.Move grdTempRun.Left - pbcArrow.Width - 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + (grdTempRun.RowHeight(grdTempRun.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        If (Trim$(grdTempRun.TextMatrix(lmEnableRow, AIRINFOINDEX)) = "") And (Trim$(grdTempRun.TextMatrix(lmEnableRow, BUSNAMEINDEX)) <> "") Then
            imcTrash.Picture = EngrMain!imcTrashOpened.Picture
        Else
            imcTrash.Picture = EngrMain!imcTrashClosed.Picture
        End If
        Select Case grdTempRun.Col
            Case BUSNAMEINDEX  'Call Letters
                edcDropdown.MaxLength = gSetMaxChars("BusName", 6)
                imMaxColChars = gGetMaxChars("BusName")
                slStr = grdTempRun.text
                ilIndex = gListBoxFind(lbcBDE, slStr)
                If ilIndex >= 0 Then
                    lbcBDE.ListIndex = ilIndex
                    edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                Else
                    edcDropdown.text = ""
                    If lbcBDE.ListCount <= 0 Then
                        lbcBDE.ListIndex = -1
                        edcDropdown.text = ""
                    Else
                        If lbcBDE.ListCount <= 1 Then
                            lbcBDE.ListIndex = 0
                            edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                        Else
                            If lmEnableRow <= grdTempRun.FixedRows Then
                                lbcBDE.ListIndex = 1
                                edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                            Else
                                slStr = grdTempRun.TextMatrix(lmEnableRow - 1, lmEnableCol)
                                ilIndex = gListBoxFind(lbcBDE, slStr)
                                If ilIndex >= 0 Then
                                    lbcBDE.ListIndex = ilIndex
                                    edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                                Else
                                    lbcBDE.ListIndex = 1
                                    edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                                End If
                            End If
                        End If
                    End If
                End If
            Case DATEINDEX  'Date
'                edcGrid.MaxLength = 10
'                edcGrid.Text = grdTempRun.Text
                cccAirDate.text = grdTempRun.text
                cccAirDate_Up.text = grdTempRun.text
            Case TIMEINDEX  'Date
'                edcGrid.MaxLength = 10
'                edcGrid.Text = grdTempRun.Text
                slStr = grdTempRun.text
                If Not gIsLength(slStr) Then
                    ltcAirTime.text = ""
                Else
                    ltcAirTime.text = slStr   'grdLibEvents.Text
                End If
            Case DESCRIPTIONINDEX  'Date
                edcGrid.MaxLength = Len(tmTSE.sDescription)
                If grdTempRun.text = "" Then
                    edcGrid.text = sgTempDescription
                    grdTempRun.text = sgTempDescription
                Else
                    edcGrid.text = grdTempRun.text
                End If
            Case STATEINDEX
                smState = grdTempRun.text
                If (Trim$(smState) = "") Or (smState = "Missing") Then
                    smState = "Active"
                End If
        End Select
        smESCValue = grdTempRun.text
        mSetFocus
    End If
End Sub
Private Sub mSetShow()
    Dim slStr As String
    
    If (lmEnableRow >= grdTempRun.FixedRows) And (lmEnableRow < grdTempRun.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case BUSNAMEINDEX
                slStr = Trim$(grdTempRun.TextMatrix(lmEnableRow, lmEnableCol))
                If (slStr = "") And (grdTempRun.TextMatrix(lmEnableRow, AIRINFOINDEX) = "") Then
                    grdTempRun.TextMatrix(lmEnableRow, DATEINDEX) = ""
                    grdTempRun.TextMatrix(lmEnableRow, TIMEINDEX) = ""
                    grdTempRun.TextMatrix(lmEnableRow, DESCRIPTIONINDEX) = ""
                    smState = ""
                    grdTempRun.TextMatrix(lmEnableRow, STATEINDEX) = ""
                End If
            Case DATEINDEX
            Case TIMEINDEX
            Case DESCRIPTIONINDEX
                If (Trim$(grdTempRun.TextMatrix(lmEnableRow, STATEINDEX)) = "") Or (grdTempRun.TextMatrix(lmEnableRow, STATEINDEX) = "Missing") Then
                    smState = "Active"
                    grdTempRun.TextMatrix(lmEnableRow, STATEINDEX) = smState
                End If
            Case STATEINDEX
                grdTempRun.TextMatrix(lmEnableRow, lmEnableCol) = smState
        End Select
        sgReturnCallName = grdTempRun.TextMatrix(lmEnableRow, BUSNAMEINDEX)
    End If
    imShowGridBox = False
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    pbcArrow.Visible = False
    cccAirDate.Visible = False
    cccAirDate_Up.Visible = False
    ltcAirTime.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcBDE.Visible = False
    edcGrid.Visible = False
    pbcState.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
End Sub
Private Function mCheckFields() As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim llRow As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    
    grdTempRun.Redraw = False
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    'Test if fields defined
    ilError = False
    For llRow = grdTempRun.FixedRows To grdTempRun.Rows - 1 Step 1
        slStr = Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX))
        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
            slStr = grdTempRun.TextMatrix(llRow, DATEINDEX)
            If slStr <> "" Then
                ilError = True
                grdTempRun.TextMatrix(llRow, BUSNAMEINDEX) = "Missing"
                grdTempRun.Row = llRow
                grdTempRun.Col = BUSNAMEINDEX
                grdTempRun.CellForeColor = vbRed
            End If
        Else
            slStr = grdTempRun.TextMatrix(llRow, DATEINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdTempRun.TextMatrix(llRow, DATEINDEX) = "Missing"
                grdTempRun.Row = llRow
                grdTempRun.Col = DATEINDEX
                grdTempRun.CellForeColor = vbRed
            Else
                If Not gIsDate(slStr) Then
                    ilError = True
                    grdTempRun.Row = llRow
                    grdTempRun.Col = DATEINDEX
                    grdTempRun.CellForeColor = vbRed
                Else
                    If gDateValue(slStr) < llNowDate Then
                        grdTempRun.Row = llRow
                        grdTempRun.Col = DATEINDEX
                        If grdTempRun.CellBackColor <> LIGHTYELLOW Then
                            ilError = True
                            grdTempRun.CellForeColor = vbRed
                        End If
                    End If
                End If
            End If
            slStr = grdTempRun.TextMatrix(llRow, TIMEINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdTempRun.TextMatrix(llRow, TIMEINDEX) = "Missing"
                grdTempRun.Row = llRow
                grdTempRun.Col = TIMEINDEX
                grdTempRun.CellForeColor = vbRed
            Else
                If Not gIsTime(slStr) Then
                    ilError = True
                    grdTempRun.Row = llRow
                    grdTempRun.Col = TIMEINDEX
                    grdTempRun.CellForeColor = vbRed
                End If
            End If
            slStr = grdTempRun.TextMatrix(llRow, STATEINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                ilError = True
                grdTempRun.TextMatrix(llRow, STATEINDEX) = "Missing"
                grdTempRun.Row = llRow
                grdTempRun.Col = STATEINDEX
                grdTempRun.CellForeColor = vbRed
            End If
        End If
    Next llRow
    grdTempRun.Redraw = True
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
    
    gGrid_AlignAllColsLeft grdTempRun
    mGridColumnWidth
    'Set Titles
    grdTempRun.TextMatrix(0, BUSNAMEINDEX) = "Bus"
    grdTempRun.TextMatrix(0, DATEINDEX) = "Date"
    grdTempRun.TextMatrix(0, TIMEINDEX) = "Time"
    grdTempRun.TextMatrix(0, DESCRIPTIONINDEX) = "Description"
    grdTempRun.TextMatrix(0, STATEINDEX) = "State"
    grdTempRun.Row = 1
    For ilCol = 0 To grdTempRun.Cols - 1 Step 1
        grdTempRun.Col = ilCol
        grdTempRun.CellAlignment = flexAlignLeftCenter
    Next ilCol
    grdTempRun.Height = cmcCancel.Top - grdTempRun.Top - 120    '8 * grdTempRun.RowHeight(0) + 30
    gGrid_IntegralHeight grdTempRun
    gGrid_Clear grdTempRun, True
    grdTempRun.Row = grdTempRun.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdTempRun.Width = EngrTempRun.Width - 2 * grdTempRun.Left
    grdTempRun.ColWidth(AIRINFOINDEX) = 0
    grdTempRun.ColWidth(SORTINDEX) = 0
    grdTempRun.ColWidth(BUSNAMEINDEX) = grdTempRun.Width / 9
    grdTempRun.ColWidth(DATEINDEX) = grdTempRun.Width / 9
    grdTempRun.ColWidth(TIMEINDEX) = grdTempRun.Width / 9
    grdTempRun.ColWidth(STATEINDEX) = grdTempRun.Width / 15
    grdTempRun.ColWidth(DESCRIPTIONINDEX) = grdTempRun.Width - GRIDSCROLLWIDTH
    For ilCol = BUSNAMEINDEX To STATEINDEX Step 1
        If ilCol <> DESCRIPTIONINDEX Then
            If grdTempRun.ColWidth(DESCRIPTIONINDEX) > grdTempRun.ColWidth(ilCol) Then
                grdTempRun.ColWidth(DESCRIPTIONINDEX) = grdTempRun.ColWidth(DESCRIPTIONINDEX) - grdTempRun.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mClearControls()
    gGrid_Clear grdTempRun, True
    imFieldChgd = False
End Sub
Private Sub mMoveCtrlsToRec(llRow As Long)
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilBDE As Integer
    
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    If Trim$(grdTempRun.TextMatrix(llRow, AIRINFOINDEX)) = "" Then
        grdTempRun.TextMatrix(llRow, AIRINFOINDEX) = "0"
        ReDim Preserve tgAirInfoTSE(0 To UBound(tgAirInfoTSE) + 1) As TSE
        ilIndex = UBound(tgAirInfoTSE) - 1
        tgAirInfoTSE(ilIndex).lCode = 0
        tgAirInfoTSE(ilIndex).iVersion = 0
        grdTempRun.TextMatrix(llRow, AIRINFOINDEX) = ilIndex
    End If
    ilIndex = grdTempRun.TextMatrix(llRow, AIRINFOINDEX)
    tgAirInfoTSE(ilIndex).iBdeCode = 0
    slStr = Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX))
    'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
    '    If StrComp(Trim$(tgCurrBDE(ilBDE).sName), slStr, vbTextCompare) = 0 Then
        ilBDE = gBinarySearchName(slStr, tgCurrBDE_Name())
        If ilBDE <> -1 Then
            tgAirInfoTSE(ilIndex).iBdeCode = tgCurrBDE_Name(ilBDE).iCode    'tgCurrBDE(ilBDE).iCode
    '        Exit For
        End If
    'Next ilBDE
    tgAirInfoTSE(ilIndex).sLogDate = grdTempRun.TextMatrix(llRow, DATEINDEX)
    tgAirInfoTSE(ilIndex).sStartTime = grdTempRun.TextMatrix(llRow, TIMEINDEX)
    tgAirInfoTSE(ilIndex).sDescription = grdTempRun.TextMatrix(llRow, DESCRIPTIONINDEX)
    If grdTempRun.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
        tgAirInfoTSE(ilIndex).sState = "D"
    Else
        tgAirInfoTSE(ilIndex).sState = "A"
    End If
    tgAirInfoTSE(ilIndex).lOrigTseCode = tgAirInfoTSE(ilIndex).lCode
    tgAirInfoTSE(ilIndex).sCurrent = "Y"
    tgAirInfoTSE(ilIndex).sEnteredDate = Format(Now, sgShowDateForm) 'smNowDate
    tgAirInfoTSE(ilIndex).sEnteredTime = Format(Now, sgShowTimeWSecForm) 'smNowTime
    tgAirInfoTSE(ilIndex).iUieCode = tgUIE.iCode
    tgAirInfoTSE(ilIndex).sUnused = ""
End Sub

Private Sub mMoveRecToCtrls()
    Dim llRow As Long
    Dim llCol As Long
    Dim ilLoop As Integer
    Dim ilBDE As Integer
    Dim slStr As String
    Dim tlSHE As SHE
    Dim ilRet As Integer
    Dim llLatestLoadDate As Long
    
    'gGrid_Clear grdTempRun, True
    llRow = grdTempRun.FixedRows
    llLatestLoadDate = gDateValue(gGetLatestLoadDate(True))
    For ilLoop = 0 To UBound(tgAirInfoTSE) - 1 Step 1
        If llRow + 1 > grdTempRun.Rows Then
            grdTempRun.AddItem ""
        End If
        grdTempRun.Row = llRow
        slStr = ""
        'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        '    If tgAirInfoTSE(ilLoop).iBdeCode = tgCurrBDE(ilBDE).iCode Then
            ilBDE = gBinarySearchBDE(tgAirInfoTSE(ilLoop).iBdeCode, tgCurrBDE())
            If ilBDE <> -1 Then
                slStr = Trim$(tgCurrBDE(ilBDE).sName)
        '        Exit For
            End If
        'Next ilBDE
        grdTempRun.TextMatrix(llRow, BUSNAMEINDEX) = slStr
        grdTempRun.TextMatrix(llRow, DATEINDEX) = Trim$(tgAirInfoTSE(ilLoop).sLogDate)
        If gDateValue(tgAirInfoTSE(ilLoop).sLogDate) <= llLatestLoadDate Then
            For llCol = 0 To grdTempRun.Cols - 1 Step 1
                grdTempRun.Row = llRow
                grdTempRun.Col = llCol
                grdTempRun.CellBackColor = LIGHTYELLOW
            Next llCol
        End If
        ilRet = gGetRec_SHE_ScheduleHeaderByDate(tgAirInfoTSE(ilLoop).sLogDate, "Template Air Info: mMoveRecToCtrls- Test Dates", tlSHE)
        If ilRet = True Then
            grdTempRun.Col = DATEINDEX
            grdTempRun.Row = llRow
            grdTempRun.CellBackColor = LIGHTYELLOW
        End If
        grdTempRun.TextMatrix(llRow, TIMEINDEX) = Trim$(tgAirInfoTSE(ilLoop).sStartTime)
        grdTempRun.TextMatrix(llRow, DESCRIPTIONINDEX) = Trim$(tgAirInfoTSE(ilLoop).sDescription)
        If tgAirInfoTSE(ilLoop).sState = "A" Then
            grdTempRun.TextMatrix(llRow, STATEINDEX) = "Active"
        Else
            grdTempRun.TextMatrix(llRow, STATEINDEX) = "Dormant"
        End If
        grdTempRun.TextMatrix(llRow, AIRINFOINDEX) = ilLoop
        llRow = llRow + 1
    Next ilLoop
    If llRow >= grdTempRun.Rows Then
        grdTempRun.AddItem ""
    End If
    grdTempRun.Redraw = True
End Sub


Private Sub cmcCancel_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrTempRun
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer
    Dim llRow As Long
    Dim slDates As String
    Dim ilIndex As Integer
    
    If imFieldChgd = False Then
        igReturnCallStatus = CALLDONE
        Unload EngrTempRun
        Exit Sub
    End If
    If Not mCheckFields() Then
        gSetMousePointer grdTempRun, grdTempRun, vbDefault
        MsgBox "One or more required fields are missing or defined incorrectly", vbCritical + vbOKOnly, "Save not Completed"
        Exit Sub
    End If
    gSetMousePointer grdTempRun, grdTempRun, vbHourglass
    slDates = ""
    For llRow = grdTempRun.FixedRows To grdTempRun.Rows - 1 Step 1
        If Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX)) <> "" Then
            If grdTempRun.TextMatrix(llRow, STATEINDEX) = "Dormant" Then
                If Trim$(grdTempRun.TextMatrix(llRow, AIRINFOINDEX)) <> "" Then
                    ilIndex = grdTempRun.TextMatrix(llRow, AIRINFOINDEX)
                    If tgAirInfoTSE(ilIndex).sState <> "D" Then
                        If slDates = "" Then
                            slDates = grdTempRun.TextMatrix(llRow, DATEINDEX)
                        Else
                            slDates = slDates & ", " & grdTempRun.TextMatrix(llRow, DATEINDEX)
                        End If
                    End If
                End If
            End If
        End If
    Next llRow
    If slDates <> "" Then
        gSetMousePointer grdTempRun, grdTempRun, vbDefault
        ilRet = MsgBox("Changing Tamplate Air Date status to 'Dormant' will result in deleting Template from Scheduled dates- " & slDates, vbOKCancel + vbQuestion, "Template Dates")
        If ilRet = vbCancel Then
            Exit Sub
        End If
    End If
    gSetMousePointer grdTempRun, grdTempRun, vbHourglass
    For llRow = grdTempRun.FixedRows To grdTempRun.Rows - 1 Step 1
        If Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX)) <> "" Then
            mMoveCtrlsToRec llRow
        End If
    Next llRow
    igReturnCallStatus = CALLDONE
    gSetMousePointer grdTempRun, grdTempRun, vbDefault
    Unload EngrTempRun
    Exit Sub

End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub


Private Sub cmcDropDown_Click()
    Select Case grdTempRun.Col
        Case BUSNAMEINDEX
            lbcBDE.Visible = Not lbcBDE.Visible
    End Select
End Sub

Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As String
    Dim ilANE As Integer
    Dim ilCCE As Integer
    Dim ilASE As Integer
    Dim ilANE2 As Integer
    
    slStr = edcDropdown.text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdTempRun.Col
        Case BUSNAMEINDEX
            llRow = gListBoxFind(lbcBDE, slStr)
            If llRow >= 0 Then
                lbcBDE.ListIndex = llRow
                edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.text)
            End If
    End Select
    If (StrComp(grdTempRun.text, edcDropdown.text, vbTextCompare) <> 0) Then
        imFieldChgd = True
        Select Case grdTempRun.Col
        End Select
        If StrComp(Trim$(edcDropdown.text), "[None]", vbTextCompare) <> 0 Then
            grdTempRun.text = edcDropdown.text
        Else
            grdTempRun.text = ""
        End If
        grdTempRun.CellForeColor = vbBlack
    End If
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
    Dim slStr As String
    
    If KeyAscii = 8 Then
        If edcDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
    If (imMaxColChars < edcDropdown.MaxLength) And (imMaxColChars > 0) And (KeyAscii <> 8) Then
        slStr = edcDropdown.text
        slStr = Left$(slStr, edcDropdown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropdown.SelStart - edcDropdown.SelLength)
        If (Len(slStr) > imMaxColChars) And (Left$(slStr, 1) <> "[") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdTempRun.Col
            Case BUSNAMEINDEX
                gProcessArrowKey Shift, KeyCode, lbcBDE, True
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
    
    Select Case grdTempRun.Col
        Case BUSNAMEINDEX
            If grdTempRun.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdTempRun.text = edcGrid.text
            grdTempRun.CellForeColor = vbBlack
        Case DATEINDEX
            If grdTempRun.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdTempRun.text = edcGrid.text
            grdTempRun.CellForeColor = vbBlack
        Case TIMEINDEX
            If grdTempRun.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdTempRun.text = edcGrid.text
            grdTempRun.CellForeColor = vbBlack
        Case DESCRIPTIONINDEX
            If grdTempRun.text <> edcGrid.text Then
                imFieldChgd = True
            End If
            grdTempRun.text = edcGrid.text
            grdTempRun.CellForeColor = vbBlack
        Case STATEINDEX
    End Select
    mSetCommands
End Sub

Private Sub edcGrid_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub edcGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilRet As Integer
    
    If imDoubleClickName Then
        ilRet = mBranch()
    End If
End Sub

Private Sub Form_Activate()
    'mGridColumns
    If imFirstActivate Then
        mSetToFirstBlankRow
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
    gSetFonts EngrTempRun
    gCenterFormModal EngrTempRun
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = KEYESC) Then

        If (lmEnableRow >= grdTempRun.FixedRows) And (lmEnableRow < grdTempRun.Rows) Then
            If (lmEnableCol >= grdTempRun.FixedCols) And (lmEnableCol < grdTempRun.Cols) Then
                If lmEnableCol = STATEINDEX Then
                    smState = smESCValue
                Else
                    grdTempRun.text = smESCValue
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
    grdTempRun.Height = cmcCancel.Top - grdTempRun.Top - 120    '8 * grdTempRun.RowHeight(0) + 30
    gGrid_IntegralHeight grdTempRun
    gGrid_FillWithRows grdTempRun
    lmCharacterWidth = CLng(pbcTab.TextWidth("n"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase imDeleteCodes
    Set EngrTempRun = Nothing
End Sub





Private Sub mInit()
    On Error GoTo ErrHand
    
    gSetMousePointer grdTempRun, grdTempRun, vbHourglass
    imcPrint.Picture = EngrMain!imcPrinter.Picture
    imcInsert.Picture = EngrMain!imcInsert.Picture
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    ReDim imDeleteCodes(0 To 0) As Integer
    imIgnoreScroll = False
    imLastColSorted = -1
    imLastSort = -1
    lmEnableRow = -1
    imFirstActivate = True
    imInChg = True
    mPopBDE
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
    gSetMousePointer grdTempRun, grdTempRun, vbDefault
    Exit Sub
ErrHand:
    gSetMousePointer grdTempRun, grdTempRun, vbDefault
    'gMsg = ""
    'For Each gErrSQL In cnn.Errors  'rdoErrors
    '    If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
    '        gMsg = "A SQL error has occured in Relay Definition-Form Load: "
    '        MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
    '    End If
    'Next gErrSQL
    'If (Err.Number <> 0) And (gMsg = "") Then
    '    gMsg = "A general error has occured in Relay Definition-Form Load: "
    '    MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    'End If
    gHandleError "EngrErrors.Txt", "Template Run-Form Load"
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

Private Sub grdTempRun_Click()
    If grdTempRun.Col >= grdTempRun.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdTempRun_EnterCell()
    mSetShow
End Sub

Private Sub grdTempRun_GotFocus()
    If grdTempRun.Col >= grdTempRun.Cols - 1 Then
        Exit Sub
    End If
End Sub

Private Sub grdTempRun_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdTempRun.TopRow
    grdTempRun.Redraw = False
End Sub

Private Sub grdTempRun_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim llEndRow As Long
    Dim ilFound As Integer
    
    'Determine if in header
    If y < grdTempRun.RowHeight(0) Then
        mSortCol grdTempRun.Col
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdTempRun, x, y)
    If Not ilFound Then
        grdTempRun.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdTempRun.CellBackColor = LIGHTYELLOW Then
        grdTempRun.Redraw = True
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    lmTopRow = grdTempRun.TopRow
    DoEvents
    llRow = grdTempRun.Row
    If grdTempRun.TextMatrix(llRow, BUSNAMEINDEX) = "" Then
        grdTempRun.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdTempRun.TextMatrix(llRow, BUSNAMEINDEX) = ""
        grdTempRun.Row = llRow + 1
        grdTempRun.Col = BUSNAMEINDEX
        grdTempRun.Redraw = True
    End If
    grdTempRun.Redraw = True
    mEnableBox
End Sub

Private Sub grdTempRun_Scroll()
    If imIgnoreScroll Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdTempRun.Redraw = False Then
        grdTempRun.Redraw = True
        If lmTopRow < grdTempRun.FixedRows Then
            grdTempRun.TopRow = grdTempRun.FixedRows
        Else
            grdTempRun.TopRow = lmTopRow
        End If
        grdTempRun.Refresh
        grdTempRun.Redraw = False
    End If
    If (imShowGridBox) And (grdTempRun.Row >= grdTempRun.FixedRows) And (grdTempRun.Col >= 0) And (grdTempRun.Col < grdTempRun.Cols - 1) Then
        If grdTempRun.RowIsVisible(grdTempRun.Row) Then
            'edcGrid.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + 30, grdTempRun.ColWidth(grdTempRun.Col) - 30, grdTempRun.RowHeight(grdTempRun.Row) - 30
            pbcArrow.Move grdTempRun.Left - pbcArrow.Width - 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + (grdTempRun.RowHeight(grdTempRun.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            'edcGrid.Visible = True
            'edcGrid.SetFocus
            mSetFocus
        Else
            'pbcClickFocus.SetFocus
            pbcSetFocus.SetFocus
            cccAirDate.Visible = False
            cccAirDate_Up.Visible = False
            ltcAirTime.Visible = False
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

Private Sub lbcBDE_Click()
    tmcClick.Enabled = False
    edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        tmcClick.Enabled = True
    End If
End Sub

Private Sub lbcBDE_DblClick()
    tmcClick.Enabled = False
    Sleep 300
    DoEvents
    edcDropdown.SetFocus
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
    edcDropdown_MouseUp 0, 0, 0, 0
    lbcBDE.Visible = False
End Sub

Private Sub lbcBDE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim llRow As Long
    Dim ilCode As Integer
    Dim ilLoop As Integer
    
    llRow = gGetListBoxRow(lbcBDE, y)
    If (llRow < lbcBDE.ListCount) And (lbcBDE.ListCount > 0) And (llRow <> -1) Then
        ilCode = lbcBDE.ItemData(llRow)
        'For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        '    If ilCode = tgCurrBDE(ilLoop).iCode Then
            ilLoop = gBinarySearchBDE(ilCode, tgCurrBDE())
            If ilLoop <> -1 Then
                lbcBDE.ToolTipText = Trim$(tgCurrBDE(ilLoop).sDescription)
        '        Exit For
            End If
        'Next ilLoop
    End If
End Sub

Private Sub ltcAirTime_OnChange()
    Dim slStr As String
    
    slStr = ltcAirTime.text
    If grdTempRun.text <> slStr Then
        grdTempRun.text = slStr
        grdTempRun.CellForeColor = vbBlack
    End If
    mSetCommands
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
    If edcGrid.Visible Or pbcState.Visible Or edcDropdown.Visible Or cccAirDate.Visible Or cccAirDate_Up.Visible Or ltcAirTime.Visible Then
        mSetShow
        If grdTempRun.Col = BUSNAMEINDEX Then
            If grdTempRun.Row > grdTempRun.FixedRows Then
                lmTopRow = -1
                grdTempRun.Row = grdTempRun.Row - 1
                If Not grdTempRun.RowIsVisible(grdTempRun.Row) Then
                    grdTempRun.TopRow = grdTempRun.TopRow - 1
                End If
                grdTempRun.Col = STATEINDEX
                mEnableBox
            Else
                cmcCancel.SetFocus
            End If
        Else
            grdTempRun.Col = grdTempRun.Col - 1
            If grdTempRun.CellBackColor = LIGHTYELLOW Then
                grdTempRun.Col = grdTempRun.Col - 1
            End If
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdTempRun.TopRow = grdTempRun.FixedRows
        grdTempRun.Col = BUSNAMEINDEX
        grdTempRun.Row = grdTempRun.FixedRows
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
        grdTempRun.CellForeColor = vbBlack
    ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
        If smState <> "Dormant" Then
            imFieldChgd = True
        End If
        smState = "Dormant"
        pbcState_Paint
        grdTempRun.CellForeColor = vbBlack
    End If
    If KeyAscii = Asc(" ") Then
        If smState = "Active" Then
            imFieldChgd = True
            smState = "Dormant"
            pbcState_Paint
            grdTempRun.CellForeColor = vbBlack
        ElseIf smState = "Dormant" Then
            imFieldChgd = True
            smState = "Active"
            pbcState_Paint
            grdTempRun.CellForeColor = vbBlack
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcState_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If smState = "Active" Then
        imFieldChgd = True
        smState = "Dormant"
        pbcState_Paint
        grdTempRun.CellForeColor = vbBlack
    ElseIf smState = "Dormant" Then
        imFieldChgd = True
        smState = "Active"
        pbcState_Paint
        grdTempRun.CellForeColor = vbBlack
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
    If edcGrid.Visible Or pbcState.Visible Or edcDropdown.Visible Or cccAirDate.Visible Or cccAirDate_Up.Visible Or ltcAirTime.Visible Then
        llEnableRow = lmEnableRow
        mSetShow
        If grdTempRun.Col = STATEINDEX Then
            llRow = grdTempRun.Rows
            Do
                llRow = llRow - 1
            Loop While grdTempRun.TextMatrix(llRow, BUSNAMEINDEX) = ""
            llRow = llRow + 1
            If (grdTempRun.Row + 1 < llRow) Then
                lmTopRow = -1
                grdTempRun.Row = grdTempRun.Row + 1
                If Not grdTempRun.RowIsVisible(grdTempRun.Row) Then
                    imIgnoreScroll = True
                    grdTempRun.TopRow = grdTempRun.TopRow + 1
                End If
                grdTempRun.Col = BUSNAMEINDEX
                'grdTempRun.TextMatrix(grdTempRun.Row, AIRINFOINDEX) = 0
                If Trim$(grdTempRun.TextMatrix(grdTempRun.Row, BUSNAMEINDEX)) <> "" Then
                    mEnableBox
                Else
                    imFromArrow = True
                    pbcArrow.Move grdTempRun.Left - pbcArrow.Width - 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + (grdTempRun.RowHeight(grdTempRun.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                End If
            Else
                If Trim$(grdTempRun.TextMatrix(llEnableRow, BUSNAMEINDEX)) <> "" Then
                    lmTopRow = -1
                    If grdTempRun.Row + 1 >= grdTempRun.Rows Then
                        grdTempRun.AddItem ""
                    End If
                    grdTempRun.Row = grdTempRun.Row + 1
                    If Not grdTempRun.RowIsVisible(grdTempRun.Row) Then
                        imIgnoreScroll = True
                        grdTempRun.TopRow = grdTempRun.TopRow + 1
                    End If
                    grdTempRun.Col = BUSNAMEINDEX
                    grdTempRun.TextMatrix(grdTempRun.Row, AIRINFOINDEX) = ""
                    'mEnableBox
                    imFromArrow = True
                    pbcArrow.Move grdTempRun.Left - pbcArrow.Width - 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + (grdTempRun.RowHeight(grdTempRun.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    pbcArrow.SetFocus
                Else
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdTempRun.Col = grdTempRun.Col + 1
            If grdTempRun.CellBackColor = LIGHTYELLOW Then
                grdTempRun.Col = grdTempRun.Col + 1
            End If
            mEnableBox
        End If
    Else
        lmTopRow = -1
        grdTempRun.TopRow = grdTempRun.FixedRows
        grdTempRun.Col = BUSNAMEINDEX
        grdTempRun.Row = grdTempRun.FixedRows
        mEnableBox
    End If
End Sub

Private Function mInsertRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdTempRun.TopRow
    llRow = grdTempRun.Row
    slMsg = "Insert above " & Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mInsertRow = False
        Exit Function
    End If
    grdTempRun.Redraw = False
    grdTempRun.AddItem "" & vbTab & "" & vbTab & "Active" & vbTab & "0" & vbTab & "N", llRow
    grdTempRun.Row = llRow
    grdTempRun.Redraw = False
    grdTempRun.TopRow = llTRow
    grdTempRun.Redraw = True
    DoEvents
    grdTempRun.Col = BUSNAMEINDEX
    mEnableBox
    mInsertRow = True
End Function

Private Function mDeleteRow() As Integer
    Dim slMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdTempRun.TopRow
    llRow = grdTempRun.Row
'    If (Val(grdTempRun.TextMatrix(llRow, AIRINFOINDEX)) <> 0) And (grdTempRun.TextMatrix(llRow, USEDFLAGINDEX) = "Y") Then
'        MsgBox Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX)) & " used or was used, unable to delete", vbInformation + vbOKCancel
'        mDeleteRow = False
'        Exit Function
'    End If
    slMsg = "Delete " & Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX))
    If MsgBox(slMsg, vbYesNo) = vbNo Then
        mDeleteRow = False
        Exit Function
    End If
    grdTempRun.Redraw = False
    If (Trim$(grdTempRun.TextMatrix(llRow, AIRINFOINDEX)) <> "") Then
        imDeleteCodes(UBound(imDeleteCodes)) = Val(grdTempRun.TextMatrix(llRow, AIRINFOINDEX))
        ReDim Preserve imDeleteCodes(0 To UBound(imDeleteCodes) + 1) As Integer
    End If
    grdTempRun.RemoveItem llRow
    imcTrash.Picture = EngrMain!imcTrashClosed.Picture
    imFieldChgd = True
    'Add row at bottom in case less rows showing then room in grid
    grdTempRun.AddItem ""
    grdTempRun.Redraw = False
    grdTempRun.TopRow = llTRow
    grdTempRun.Redraw = True
    DoEvents
    grdTempRun.Col = BUSNAMEINDEX
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


Private Sub mSetFocus()
    Select Case grdTempRun.Col
        Case BUSNAMEINDEX  'Call Letters
            edcDropdown.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + 15, grdTempRun.ColWidth(grdTempRun.Col) - 30, grdTempRun.RowHeight(grdTempRun.Row) - 15
            edcDropdown.Width = gSetCtrlWidth("BusName", lmCharacterWidth, edcDropdown.Width, 6)
            cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
            lbcBDE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
            gSetListBoxHeight lbcBDE, CLng(grdTempRun.Height / 2)
            If lbcBDE.Top + lbcBDE.Height > cmcCancel.Top Then
                lbcBDE.Top = edcDropdown.Top - lbcBDE.Height
            End If
            edcDropdown.Visible = True
            cmcDropDown.Visible = True
            edcDropdown.SetFocus
        Case DATEINDEX  'Date
'            edcGrid.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + 15, grdTempRun.ColWidth(grdTempRun.Col) - 30, grdTempRun.RowHeight(grdTempRun.Row) - 15
'            edcGrid.Visible = True
'            edcGrid.SetFocus
            cccAirDate_Up.SetEditBoxHeight = grdTempRun.RowHeight(grdTempRun.Row) - 15
            If grdTempRun.RowPos(grdTempRun.Row) + cccAirDate_Up.TrueHeight < grdTempRun.Height Then
                cccAirDate.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + 15, grdTempRun.ColWidth(grdTempRun.Col) - 30, grdTempRun.RowHeight(grdTempRun.Row) - 15
                cccAirDate.Visible = True
                cccAirDate.SetFocus
            Else
                cccAirDate_Up.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + grdTempRun.RowHeight(grdTempRun.Row) - cccAirDate_Up.TrueHeight, grdTempRun.ColWidth(grdTempRun.Col) - 30 ', grdTempRun.RowHeight(grdTempRun.Row) - 15
                cccAirDate_Up.Visible = True
                cccAirDate_Up.SetFocus
            End If
        Case TIMEINDEX  'Date
'            edcGrid.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + 15, grdTempRun.ColWidth(grdTempRun.Col) - 30, grdTempRun.RowHeight(grdTempRun.Row) - 15
'            edcGrid.Visible = True
'            edcGrid.SetFocus
            ltcAirTime.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + 15, grdTempRun.ColWidth(grdTempRun.Col) - 30, grdTempRun.RowHeight(grdTempRun.Row) - 15
            ltcAirTime.Width = gSetCtrlWidth("TIME", lmCharacterWidth, ltcAirTime.Width, 0)
            ltcAirTime.Visible = True
            ltcAirTime.SetFocus
        Case DESCRIPTIONINDEX  'Date
            edcGrid.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + 15, grdTempRun.ColWidth(grdTempRun.Col) - 30, grdTempRun.RowHeight(grdTempRun.Row) - 15
            edcGrid.Visible = True
            edcGrid.SetFocus
        Case STATEINDEX
            pbcState.Move grdTempRun.Left + grdTempRun.ColPos(grdTempRun.Col) + 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + 15, grdTempRun.ColWidth(grdTempRun.Col) - 30, grdTempRun.RowHeight(grdTempRun.Row) - 15
            pbcState.Visible = True
            pbcState.SetFocus
    End Select
End Sub

Private Sub mPopBDE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrTempDef-mPopBDE Bus Definition", tgCurrBDE())
    lbcBDE.Clear
    For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        If tgCurrBDE(ilLoop).sState = "A" Then
            lbcBDE.AddItem Trim$(tgCurrBDE(ilLoop).sName)
            lbcBDE.ItemData(lbcBDE.NewIndex) = tgCurrBDE(ilLoop).iCode
        End If
    Next ilLoop
    If (StrComp(sgUserName, "Counterpoint", vbTextCompare) = 0) Or (StrComp(sgUserName, "Guide", vbTextCompare) = 0) Or (igListStatus(BUSLIST) = 2) Then
        lbcBDE.AddItem "[New]", 0
        lbcBDE.ItemData(lbcBDE.NewIndex) = 0
    Else
        lbcBDE.AddItem "[View]", 0
        lbcBDE.ItemData(lbcBDE.NewIndex) = 0
    End If
End Sub

Private Function mBranch() As Integer
    Dim llRow As Long
    Dim slStr As String
    
    mBranch = True
    If (lmEnableRow >= grdTempRun.FixedRows) And (lmEnableRow < grdTempRun.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        slStr = Trim$(grdTempRun.TextMatrix(lmEnableRow, lmEnableCol))
        If (slStr <> "") And (StrComp(slStr, "[None]", vbTextCompare) <> 0) Then
            Select Case lmEnableCol
                Case BUSNAMEINDEX
                    llRow = gListBoxFind(lbcBDE, slStr)
                    If (llRow <= 0) Or (imDoubleClickName) Then
                        igInitCallInfo = 1
                        sgInitCallName = slStr
                        EngrBus.Show vbModal
                        sgCurrBDEStamp = ""
                        mPopBDE
                        lbcBDE.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                        gSetListBoxHeight lbcBDE, CLng(grdTempRun.Height / 2)
                        If lbcBDE.Top + lbcBDE.Height > cmcCancel.Top Then
                            lbcBDE.Top = edcDropdown.Top - lbcBDE.Height
                        End If
                        If igReturnCallStatus = CALLDONE Then
                            slStr = sgReturnCallName
                            'llRow = SendMessageByString(lbcDNE.hwnd, LB_FINDSTRING, -1, slStr)
                            llRow = gListBoxFind(lbcBDE, slStr)
                            If llRow > 0 Then
                                lbcBDE.ListIndex = llRow
                                edcDropdown.text = lbcBDE.List(lbcBDE.ListIndex)
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
            End Select
        End If
    End If
    imDoubleClickName = False
End Function

Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If edcGrid.Visible Or pbcState.Visible Or edcDropdown.Visible Then
        Select Case grdTempRun.Col
            Case BUSNAMEINDEX
                lbcBDE.Visible = False
        End Select
    End If
End Sub

Private Sub mSetToFirstBlankRow()
    Dim llRow As Long
    Dim slStr As String
    
    'Find first blank row
    For llRow = grdTempRun.FixedRows To grdTempRun.Rows - 1 Step 1
        slStr = Trim$(grdTempRun.TextMatrix(llRow, BUSNAMEINDEX))
        If (slStr = "") Then
            grdTempRun.Row = llRow
            Do While Not grdTempRun.RowIsVisible(grdTempRun.Row)
                imIgnoreScroll = True
                grdTempRun.TopRow = grdTempRun.TopRow + 1
            Loop
            grdTempRun.Col = BUSNAMEINDEX
            'mEnableBox
            imFromArrow = True
            pbcArrow.Move grdTempRun.Left - pbcArrow.Width - 30, grdTempRun.Top + grdTempRun.RowPos(grdTempRun.Row) + (grdTempRun.RowHeight(grdTempRun.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            pbcArrow.SetFocus
            Exit Sub
        End If
    Next llRow

End Sub


