VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmExportSpec 
   Caption         =   "Export Specification"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "AffExportSpec.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8910
   Begin VB.ListBox lbcType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffExportSpec.frx":08CA
      Left            =   6210
      List            =   "AffExportSpec.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3885
      Visible         =   0   'False
      Width           =   1410
   End
   Begin V81Affiliate.CSI_Calendar cbcDate 
      Height          =   195
      Left            =   1035
      TabIndex        =   8
      Top             =   2070
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   344
      Text            =   "11/12/2020"
      BackColor       =   16777088
      ForeColor       =   -2147483640
      BorderStyle     =   0
      CSI_ShowDropDownOnFocus=   0   'False
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   51200
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   5310
      TabIndex        =   5
      Top             =   3525
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
      Left            =   6255
      Picture         =   "AffExportSpec.frx":08CE
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3495
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcProgSchdFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   90
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   120
      Width           =   60
   End
   Begin VB.PictureBox pbcProgSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   1110
      Width           =   60
   End
   Begin VB.PictureBox pbcSpecSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   1275
      Width           =   60
   End
   Begin VB.PictureBox pbcSpecTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   9
      Top             =   5130
      Width           =   60
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   5445
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3750
      TabIndex        =   12
      Top             =   5445
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   135
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5730
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   60
      Picture         =   "AffExportSpec.frx":09C8
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   90
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   630
      Top             =   5400
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6060
      FormDesignWidth =   8910
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   2025
      TabIndex        =   11
      Top             =   5445
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpec 
      Height          =   4815
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   255
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   11
      Cols            =   10
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8085
      Picture         =   "AffExportSpec.frx":0CD2
      Top             =   5355
      Width           =   480
   End
End
Attribute VB_Name = "frmExportSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmExportSpec - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imCtrlKey As Integer
Private imFieldChgd As Integer
Private lmRowSelected As Long
Private lmTopRow As Long
Private lmSpecEnableRow As Long
Private lmSpecEnableCol As Long
Private imSpecShowGridBox As Integer
Private imInChg As Integer
Private imInModel As Integer
Private imFromArrow As Integer

Private imLastSpecColSorted As Integer
Private imLastSpecSort As Integer

Private imSpecColPos(0 To 9) As Integer

Private lmDelEhtCode() As Long

Private rst_Eht As ADODB.Recordset
Private rst_Evt As ADODB.Recordset
Private rst_Ect As ADODB.Recordset

'Grid Controls

Const TYPEINDEX = 0
Const NAMEINDEX = 1
Const LDEINDEX = 2
Const LEADINDEX = 3
Const CYCLEINDEX = 4
Const VEHICLEINDEX = 5
Const NAMECODEINDEX = 6
Const SORTINDEX = 7
Const REFROWNOINDEX = 8
Const EHTCODEINDEX = 9

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    grdSpec.Rows = grdSpec.FixedRows + 10
    gGrid_IntegralHeight grdSpec
    gGrid_FillWithRows grdSpec
    For llRow = grdSpec.FixedRows To grdSpec.Rows - 1 Step 1
        For llCol = TYPEINDEX To EHTCODEINDEX Step 1
            If (llCol <> EHTCODEINDEX) And (llCol <> REFROWNOINDEX) Then
                grdSpec.TextMatrix(llRow, llCol) = ""
            ElseIf llCol = REFROWNOINDEX Then
                grdSpec.TextMatrix(llRow, llCol) = llRow
            Else
                grdSpec.TextMatrix(llRow, llCol) = "0"
            End If
        Next llCol
    Next llRow
    
End Sub



Private Sub cbcDate_Change()
    imFieldChgd = True
End Sub

Private Sub cmcDropDown_Click()
    lbcType.Visible = Not lbcType.Visible
End Sub

Private Sub cmdCancel_Click()
    
    mDeleteNewEht
    Unload frmExportSpec
End Sub

Private Sub cmdCancel_GotFocus()
    mSpecSetShow
End Sub

Private Sub cmdDone_Click()
    Dim ilRet As Integer
    
    If imFieldChgd = True Then
        If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
            'Screen.MousePointer = vbHourglass
            gSetMousePointer grdSpec, grdSpec, vbHourglass
            ilRet = mSave()
            grdSpec.Redraw = True
            If Not ilRet Then
                'Screen.MousePointer = vbDefault
                gSetMousePointer grdSpec, grdSpec, vbDefault
                Exit Sub    ' Dont exit until user takes care of whatever fields are invalid or missing.
            End If
        Else
            mDeleteNewEht
        End If
    End If
    On Error GoTo 0
    'Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdSpec, vbDefault
    Unload frmExportSpec
    Exit Sub
   
End Sub
Private Sub cmdDone_GotFocus()
    mSpecSetShow
End Sub
Private Sub cmdSave_Click()
    Dim ilRet As Integer
    
    ilRet = mSave()
    grdSpec.Redraw = True
    
End Sub

Private Sub cmdSave_GotFocus()
    mSpecSetShow
End Sub

Private Sub edcDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        'Screen.MousePointer = vbHourglass
        gSetMousePointer grdSpec, grdSpec, vbHourglass
        mSetGridColumns
        mSetGridTitles
        gGrid_IntegralHeight grdSpec
        gGrid_FillWithRows grdSpec
        mPopulate
        imFirstTime = False
        'Screen.MousePointer = vbDefault
        gSetMousePointer grdSpec, grdSpec, vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.55   '1.15
    Me.Height = Screen.Height / 1.55
    Me.Top = (Screen.Height - Me.Height) / 1.75
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmExportSpec
    'gCenterForm frmExportSpec
End Sub

Private Sub Form_Load()
    
    'Screen.MousePointer = vbHourglass
    gSetMousePointer grdSpec, grdSpec, vbHourglass
    
    mInit
    'Screen.MousePointer = vbDefault
    gSetMousePointer grdSpec, grdSpec, vbDefault
    Exit Sub
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcTrash.Picture = frmDirectory!imcTrashClosed.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Eht.Close
    rst_Evt.Close
    rst_Ect.Close
    Erase lmDelEhtCode
    Set frmExportSpec = Nothing
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim ilSpec As Integer
    '8156
    Dim ilvehicle As Vendors
    Dim blInclude As Boolean
    
    igVehicleSpecChgFlag = False
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    imLastSpecColSorted = -1
    imLastSpecSort = -1
    lmSpecEnableRow = -1
    lmSpecEnableCol = -1
    imSpecShowGridBox = False
    imFromArrow = False
    imFieldChgd = False
    For ilSpec = 0 To UBound(tgSpecInfo) Step 1
        '8156
        blInclude = True
        Select Case tgSpecInfo(ilSpec).sType
            Case "X"
                ilvehicle = XDS_Break
            Case "W"
                ilvehicle = Wegener_Compel
            Case "P"
               ilvehicle = Vendors.Wegener_IPump
            Case "D"
               ilvehicle = Vendors.iDc
            Case Else
                ilvehicle = Vendors.None
        End Select
        If ilvehicle > Vendors.None Then
            blInclude = gAdjustAllowedExportsImports(ilvehicle, False)
            If Not blInclude And ilvehicle = Vendors.XDS_Break Then
                blInclude = gAdjustAllowedExportsImports(XDS_ISCI, False)
            End If
        '10023
        ElseIf tgSpecInfo(ilSpec).sType = "I" Then
                blInclude = gISCIExport
        End If
        If blInclude Then
            lbcType.AddItem tgSpecInfo(ilSpec).sName
        End If
    Next ilSpec
    ReDim tgEhtInfo(0 To 0) As EHTINFO
    ReDim tgEvtInfo(0 To 0) As EVTINFO
    ReDim tgEctInfo(0 To 0) As ECTINFO
    ReDim lmDelEhtCode(0 To 0) As Long
    imcTrash.Picture = frmDirectory!imcTrashClosed.Picture
    imcTrash.Enabled = False
    If (sgExptSpec <> "Y") Or (sgUstWin(14) = "V") Then
        cmdSave.Enabled = False
        cmdDone.Enabled = False
    End If
    
End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCode As Long
    Dim ilRet As Integer
    
    mClearGrid
    ilRet = mPopulateGrid()
    
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    

    grdSpec.ColWidth(NAMECODEINDEX) = 0
    grdSpec.ColWidth(SORTINDEX) = 0
    grdSpec.ColWidth(REFROWNOINDEX) = 0
    grdSpec.ColWidth(EHTCODEINDEX) = 0
    grdSpec.ColWidth(TYPEINDEX) = grdSpec.Width * 0.11
    grdSpec.ColWidth(LDEINDEX) = grdSpec.Width * 0.2
    grdSpec.ColWidth(LEADINDEX) = grdSpec.Width * 0.08
    grdSpec.ColWidth(CYCLEINDEX) = grdSpec.Width * 0.08
    grdSpec.ColWidth(VEHICLEINDEX) = grdSpec.Width * 0.15
    
    grdSpec.ColWidth(NAMEINDEX) = grdSpec.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To VEHICLEINDEX Step 1
        If ilCol <> NAMEINDEX Then
            grdSpec.ColWidth(NAMEINDEX) = grdSpec.ColWidth(NAMEINDEX) - grdSpec.ColWidth(ilCol)
        End If
    Next ilCol
    gGrid_AlignAllColsLeft grdSpec
    grdSpec.ColAlignment(VEHICLEINDEX) = flexAlignCenterCenter
    For ilCol = 0 To grdSpec.Cols - 1 Step 1
        imSpecColPos(ilCol) = grdSpec.ColPos(ilCol)
    Next ilCol
End Sub

Private Sub mSetGridTitles()
    Dim llCol As Long
    
    grdSpec.TextMatrix(0, TYPEINDEX) = "Export"
    grdSpec.TextMatrix(0, NAMEINDEX) = "Name"
    grdSpec.TextMatrix(0, LDEINDEX) = "Last Export Date"
    grdSpec.TextMatrix(0, LEADINDEX) = "Lead"
    grdSpec.TextMatrix(0, CYCLEINDEX) = "Cycle"
    grdSpec.TextMatrix(0, VEHICLEINDEX) = "Vehicle List"
    grdSpec.Row = 0
    For llCol = TYPEINDEX To VEHICLEINDEX Step 1
        grdSpec.Col = llCol
        grdSpec.CellBackColor = LIGHTBLUE
    Next llCol
End Sub

Private Sub mSpecSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    
    For llRow = grdSpec.FixedRows To grdSpec.Rows - 1 Step 1
        slStr = Trim$(grdSpec.TextMatrix(llRow, TYPEINDEX))
        If slStr <> "" Then
            If (ilCol = LDEINDEX) Then
                If grdSpec.TextMatrix(llRow, LDEINDEX) <> "" Then
                    slSort = Trim$(Str$(gDateValue(grdSpec.TextMatrix(llRow, LDEINDEX))))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                Else
                    slSort = "999999"
                End If
            Else
                slSort = UCase$(Trim$(grdSpec.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdSpec.TextMatrix(llRow, SORTINDEX)
            grdSpec.TextMatrix(llRow, SORTINDEX) = slSort & slStr
        End If
    Next llRow
    If ilCol = imLastSpecColSorted Then
        imLastSpecColSorted = SORTINDEX
    Else
        imLastSpecColSorted = -1
        imLastSpecSort = -1
    End If
    gGrid_SortByCol grdSpec, TYPEINDEX, SORTINDEX, imLastSpecColSorted, imLastSpecSort
    imLastSpecColSorted = ilCol
End Sub

Private Sub mSpecEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llCol As Long
    
    If (grdSpec.Row >= grdSpec.FixedRows) And (grdSpec.Row < grdSpec.Rows) And (grdSpec.Col >= TYPEINDEX) And (grdSpec.Col <= CYCLEINDEX) Then
        lmSpecEnableRow = grdSpec.Row
        lmSpecEnableCol = grdSpec.Col
        imSpecShowGridBox = True
        pbcArrow.Move grdSpec.Left - pbcArrow.Width, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + (grdSpec.RowHeight(grdSpec.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Select Case grdSpec.Col
            Case TYPEINDEX
                edcDropdown.MaxLength = 8
                edcDropdown.Move grdSpec.Left + imSpecColPos(grdSpec.Col) + 15, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
                lbcType.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcType, 10
                If lbcType.Top + lbcType.Height > cmdCancel.Top Then
                    lbcType.Move edcDropdown.Left, edcDropdown.Top - lbcType.Height, edcDropdown.Width + cmcDropDown.Width
                End If
                slStr = grdSpec.Text
                ilIndex = SendMessageByString(lbcType.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcType.ListIndex = ilIndex
                    edcDropdown.Text = lbcType.List(lbcType.ListIndex)
                Else
                    lbcType.ListIndex = -1
                    edcDropdown.Text = ""
                End If
                If edcDropdown.Height > grdSpec.RowHeight(grdSpec.Row) - 15 Then
                    edcDropdown.FontName = "Arial"
                    edcDropdown.Height = grdSpec.RowHeight(grdSpec.Row) - 15
                End If
                edcDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcType.Visible = True
                If edcDropdown.Enabled Then
                    edcDropdown.SetFocus
                End If
            Case NAMEINDEX
                edcDropdown.MaxLength = 50
                edcDropdown.Move grdSpec.Left + imSpecColPos(grdSpec.Col) + 15, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                If grdSpec.Text <> "Missing" Then
                    edcDropdown.Text = grdSpec.Text
                Else
                    edcDropdown.Text = ""
                End If
                If edcDropdown.Height > grdSpec.RowHeight(grdSpec.Row) - 15 Then
                    edcDropdown.FontName = "Arial"
                    edcDropdown.Height = grdSpec.RowHeight(grdSpec.Row) - 15
                End If
                edcDropdown.Visible = True
                If edcDropdown.Enabled Then
                    edcDropdown.SetFocus
                End If
            Case LDEINDEX
                cbcDate.Move grdSpec.Left + imSpecColPos(grdSpec.Col) + 15, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                cbcDate.Text = grdSpec.Text
                cbcDate.Visible = True
                cbcDate.SetFocus
            Case LEADINDEX, CYCLEINDEX
                edcDropdown.MaxLength = 3
                edcDropdown.Move grdSpec.Left + imSpecColPos(grdSpec.Col) + 15, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - 30, grdSpec.RowHeight(grdSpec.Row) - 15
                If grdSpec.Text <> "Missing" Then
                    edcDropdown.Text = grdSpec.Text
                Else
                    edcDropdown.Text = ""
                End If
                If edcDropdown.Height > grdSpec.RowHeight(grdSpec.Row) - 15 Then
                    edcDropdown.FontName = "Arial"
                    edcDropdown.Height = grdSpec.RowHeight(grdSpec.Row) - 15
                End If
                edcDropdown.Visible = True
                If edcDropdown.Enabled Then
                    edcDropdown.SetFocus
                End If
        End Select
    End If
    imcTrash.Enabled = True
End Sub

Private Sub mSpecSetShow()
    Dim slStr As String
    Dim ilSpec As Integer

    If (lmSpecEnableRow >= grdSpec.FixedRows) And (lmSpecEnableRow < grdSpec.Rows) Then
        Select Case lmSpecEnableCol
            Case TYPEINDEX
                For ilSpec = LBound(tgSpecInfo) To UBound(tgSpecInfo) Step 1
                    slStr = grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol)
                    If slStr = tgSpecInfo(ilSpec).sName Then
                        grdSpec.TextMatrix(lmSpecEnableRow, NAMECODEINDEX) = tgSpecInfo(ilSpec).sType
                        Exit For
                    End If
                Next ilSpec
            Case LDEINDEX
                grdSpec.TextMatrix(lmSpecEnableRow, lmSpecEnableCol) = cbcDate.Text
        End Select
    End If
    lmSpecEnableRow = -1
    lmSpecEnableCol = -1
    imSpecShowGridBox = False
    pbcArrow.Visible = False
    edcDropdown.Visible = False
    cmcDropDown.Visible = False
    cbcDate.Visible = False
    lbcType.Visible = False
    imcTrash.Enabled = False
End Sub

Private Sub grdSpec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilFound As Integer
    Dim llRow As Long
    Dim llEhtInfo As Long
    
    mSpecSetShow
    If Y < grdSpec.RowHeight(0) Then
        'Screen.MousePointer = vbHourglass
        gSetMousePointer grdSpec, grdSpec, vbHourglass
        grdSpec.Redraw = True
        grdSpec.Col = grdSpec.MouseCol
        mSpecSortCol grdSpec.Col
        grdSpec.Row = 0
        'Screen.MousePointer = vbDefault
        gSetMousePointer grdSpec, grdSpec, vbDefault
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdSpec, X, Y)
    If Not ilFound Then
        grdSpec.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If (grdSpec.MouseCol = VEHICLEINDEX) And (grdSpec.TextMatrix(grdSpec.MouseRow, TYPEINDEX) <> "") Then
        llRow = grdSpec.MouseRow
        mGetVehicles True, llRow
        grdSpec.Redraw = True
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If (sgExptSpec <> "Y") Or (sgUstWin(14) = "V") Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    If grdSpec.Col > CYCLEINDEX Then
        grdSpec.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdSpec.TopRow
    llRow = grdSpec.Row
    If grdSpec.TextMatrix(llRow, TYPEINDEX) = "" Then
        grdSpec.Redraw = False
        Do
            llRow = llRow - 1
        Loop While grdSpec.TextMatrix(llRow, TYPEINDEX) = ""
        grdSpec.Row = llRow + 1
        grdSpec.TextMatrix(grdSpec.Row, REFROWNOINDEX) = grdSpec.Row
        grdSpec.TextMatrix(grdSpec.Row, EHTCODEINDEX) = "0"
        llEhtInfo = UBound(tgEhtInfo)
        tgEhtInfo(llEhtInfo).iRefRowNo = grdSpec.Row
        tgEhtInfo(llEhtInfo).lFirstEvt = -1
        tgEhtInfo(llEhtInfo).lFirstEct = -1
        ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
        'grdSpec.Redraw = True
    End If
    grdSpec.Redraw = True
    mSpecEnableBox

End Sub

Private Sub grdSpec_Scroll()
    If grdSpec.Redraw = False Then
        grdSpec.Redraw = True
        grdSpec.TopRow = lmTopRow
        grdSpec.Refresh
        grdSpec.Redraw = False
    End If
    If (imSpecShowGridBox) And (grdSpec.Row >= grdSpec.FixedRows) And (grdSpec.Col >= LDEINDEX) And (grdSpec.Col <= CYCLEINDEX) Then
        If grdSpec.RowIsVisible(grdSpec.Row) Then
            pbcArrow.Move grdSpec.Left - pbcArrow.Width, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + (grdSpec.RowHeight(grdSpec.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            If (grdSpec.Col = LDEINDEX) Or (grdSpec.Col = LEADINDEX) Or (grdSpec.Col = CYCLEINDEX) Then
                edcDropdown.Move grdSpec.Left + imSpecColPos(grdSpec.Col) + 15, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + 15, grdSpec.ColWidth(grdSpec.Col) - cmcDropDown.Width + 30, grdSpec.RowHeight(grdSpec.Row) - 15
                edcDropdown.Visible = True
            End If
        Else
            pbcArrow.Visible = False
            edcDropdown.Visible = False
        End If
    Else
        pbcArrow.Visible = False
        imFromArrow = False
    End If

End Sub


Private Sub imcTrash_Click()
    Dim iLoop As Integer
    Dim llRow As Long
    Dim llRows As Long
    
    llRow = lmSpecEnableRow
    llRows = grdSpec.Rows
    mSpecSetShow
    If (llRow < grdSpec.FixedRows) Or (llRow > grdSpec.Rows - 1) Then
        Exit Sub
    End If
    lmTopRow = -1
    If grdSpec.TextMatrix(llRow, TYPEINDEX) <> "" Then
        If (Trim$(grdSpec.TextMatrix(llRow, EHTCODEINDEX)) <> "") Or (grdSpec.TextMatrix(llRow, EHTCODEINDEX) <> "0") Then
            imFieldChgd = True
            lmDelEhtCode(UBound(lmDelEhtCode)) = grdSpec.TextMatrix(llRow, EHTCODEINDEX)
            ReDim Preserve lmDelEhtCode(0 To UBound(lmDelEhtCode) + 1) As Long
        End If
        grdSpec.RemoveItem llRow
        gGrid_FillWithRows grdSpec
        If pbcClickFocus.Enabled Then
            pbcClickFocus.SetFocus
        End If
    End If
End Sub

Private Sub imcTrash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imcTrash.Enabled Then
        imcTrash.Picture = frmDirectory!imcTrashOpened.Picture
    End If
End Sub

Private Sub lbcType_Click()
    edcDropdown.Text = lbcType.List(lbcType.ListIndex)
    If (edcDropdown.Visible) And (edcDropdown.Enabled) Then
        edcDropdown.SetFocus
        lbcType.Visible = False
    End If
End Sub

Private Sub pbcClickFocus_GotFocus()
    mSpecSetShow
End Sub


Private Sub pbcSpecSTab_GotFocus()
    
    If GetFocus() <> pbcSpecSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mSpecEnableBox
        Exit Sub
    End If
    If imSpecShowGridBox Then
        mSpecSetShow
        If grdSpec.Col = TYPEINDEX Then
            If grdSpec.Row > grdSpec.FixedRows Then
                lmTopRow = -1
                grdSpec.Row = grdSpec.Row - 1
                If Not grdSpec.RowIsVisible(grdSpec.Row) Then
                    grdSpec.TopRow = grdSpec.TopRow - 1
                End If
                grdSpec.Col = CYCLEINDEX
                mSpecEnableBox
            Else
                If pbcClickFocus.Enabled Then
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdSpec.Col = grdSpec.Col - 1
            mSpecEnableBox
        End If
    Else
        lmTopRow = -1
        grdSpec.TopRow = grdSpec.FixedRows
        grdSpec.Col = TYPEINDEX
        grdSpec.Row = grdSpec.FixedRows
        mSpecEnableBox
    End If

End Sub

Private Sub pbcSpecTab_GotFocus()
    Dim llCol As Long
    Dim slStr As String
    Dim llRow As Long
    Dim llEhtInfo As Long
    Dim llSvRow As Long
    Dim llSvCol As Long
    
    If GetFocus() <> pbcSpecTab.hwnd Then
        Exit Sub
    End If
    If imSpecShowGridBox Then
        llSvRow = lmSpecEnableRow
        llSvCol = lmSpecEnableCol
        mSpecSetShow
        If (llSvCol = CYCLEINDEX) Then
            If grdSpec.TextMatrix(llSvRow, TYPEINDEX) <> "" Then
                llRow = llSvRow
                If grdSpec.TextMatrix(llRow, VEHICLEINDEX) = "" Then
                    mGetVehicles True, llRow
                End If
                lmTopRow = -1
                If llRow + 1 >= grdSpec.Rows Then
                    grdSpec.AddItem ""
                    grdSpec.TextMatrix(llRow + 1, REFROWNOINDEX) = grdSpec.Row + 1
                    grdSpec.TextMatrix(llRow + 1, EHTCODEINDEX) = "0"
                    llEhtInfo = UBound(tgEhtInfo)
                    tgEhtInfo(llEhtInfo).iRefRowNo = llRow + 1
                    tgEhtInfo(llEhtInfo).lFirstEvt = -1
                    tgEhtInfo(llEhtInfo).lFirstEct = -1
                    ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
                End If
                grdSpec.Row = llRow + 1
                If Not grdSpec.RowIsVisible(grdSpec.Row) Then
                    grdSpec.TopRow = grdSpec.TopRow + 1
                End If
                'If grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) <> "" Then
                    grdSpec.Col = TYPEINDEX
                    imFromArrow = True
                    pbcArrow.Move grdSpec.Left - pbcArrow.Width, grdSpec.Top + grdSpec.RowPos(grdSpec.Row) + (grdSpec.RowHeight(grdSpec.Row) - pbcArrow.Height) / 2
                    pbcArrow.Visible = True
                    If pbcArrow.Enabled Then
                        pbcArrow.SetFocus
                    End If
                'Else
                '    If pbcClickFocus.Enabled Then
                '        pbcClickFocus.SetFocus
                '    End If
                'End If
            Else
                If pbcClickFocus.Enabled Then
                    pbcClickFocus.SetFocus
                End If
            End If
        ElseIf (grdSpec.Col = TYPEINDEX) Then
            If grdSpec.TextMatrix(grdSpec.Row, TYPEINDEX) <> "" Then
                grdSpec.Col = grdSpec.Col + 1
                mSpecEnableBox
            Else
                mSpecEnableBox
                If pbcClickFocus.Enabled Then
                    pbcClickFocus.SetFocus
                End If
            End If
        Else
            grdSpec.Col = grdSpec.Col + 1
            mSpecEnableBox
        End If
    Else
        lmTopRow = -1
        grdSpec.TopRow = grdSpec.FixedRows
        grdSpec.Col = TYPEINDEX
        grdSpec.Row = grdSpec.FixedRows
        mSpecEnableBox
    End If

End Sub


Private Sub edcDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    Select Case grdSpec.Col
        Case TYPEINDEX
            slStr = edcDropdown.Text
            ilLen = Len(slStr)
            If imBSMode Then
                ilLen = ilLen - 1
                If ilLen > 0 Then
                    slStr = Left$(slStr, ilLen)
                End If
                imBSMode = False
            End If
            llRow = SendMessageByString(lbcType.hwnd, LB_FINDSTRING, -1, slStr)
            If llRow >= 0 Then
                lbcType.ListIndex = llRow
                edcDropdown.Text = lbcType.List(lbcType.ListIndex)
                edcDropdown.SelStart = ilLen
                edcDropdown.SelLength = Len(edcDropdown.Text)
                grdSpec.CellForeColor = vbBlack
                If grdSpec.Text <> slStr Then
                    imFieldChgd = True
                End If
                grdSpec.Text = lbcType.List(lbcType.ListIndex)
            Else
                lbcType.ListIndex = -1
                edcDropdown.Text = ""
            End If
        Case NAMEINDEX
            slStr = Trim$(edcDropdown.Text)
            If (slStr <> "") Then
                grdSpec.CellForeColor = vbBlack
                If grdSpec.Text <> slStr Then
                    imFieldChgd = True
                End If
                grdSpec.Text = slStr
            End If
        Case LEADINDEX
            slStr = Trim$(edcDropdown.Text)
            If (slStr <> "") Then
                grdSpec.CellForeColor = vbBlack
                If grdSpec.Text <> slStr Then
                    imFieldChgd = True
                End If
                grdSpec.Text = slStr
            End If
        Case CYCLEINDEX
            slStr = Trim$(edcDropdown.Text)
            If (slStr <> "") Then
                grdSpec.CellForeColor = vbBlack
                If grdSpec.Text <> slStr Then
                    imFieldChgd = True
                End If
                grdSpec.Text = slStr
            End If
    End Select
End Sub

Private Sub edcDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropdown_KeyPress(KeyAscii As Integer)
    Select Case grdSpec.Col
        Case TYPEINDEX
            If KeyAscii = 8 Then
                If edcDropdown.SelLength <> 0 Then
                    imBSMode = True
                End If
            End If
        Case NAMEINDEX
        Case LEADINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case CYCLEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub edcDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case grdSpec.Col
        Case TYPEINDEX
            If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
                gProcessArrowKey Shift, KeyCode, lbcType, True
            End If
    End Select
End Sub



Private Function mPopulateGrid() As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim llCode As Long
    Dim blFound As Boolean
    Dim ilLoop As Integer
    Dim ilSpec As Integer
    Dim llEhtInfo As Long
    Dim llEvtInfo As Long
    Dim llEctInfo As Long
    Dim llNext As Long
    
    On Error GoTo ErrHand
    mPopulateGrid = True
    grdSpec.Redraw = False
    SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtSubType = 'S'"
    Set rst_Eht = gSQLSelectCall(SQLQuery)
    If Not rst_Eht.EOF Then
        llRow = grdSpec.FixedRows
        Do While Not rst_Eht.EOF
            blFound = False
            For ilSpec = LBound(tgSpecInfo) To UBound(tgSpecInfo) Step 1
                If Trim$(rst_Eht!ehtExportType) = Trim$(tgSpecInfo(ilSpec).sType) Then
                    If llRow >= grdSpec.Rows Then
                        grdSpec.AddItem ""
                    End If
                    grdSpec.TextMatrix(llRow, TYPEINDEX) = Trim$(tgSpecInfo(ilSpec).sName)
                    grdSpec.TextMatrix(llRow, NAMECODEINDEX) = Trim$(tgSpecInfo(ilSpec).sType)
                    blFound = True
                    Exit For
                End If
            Next ilSpec
            If blFound Then
                grdSpec.TextMatrix(llRow, NAMEINDEX) = Trim$(rst_Eht!ehtExportName)
                If gDateValue(Format(rst_Eht!ehtLDE, sgShowDateForm)) <> gDateValue("1/1/1970") Then
                    grdSpec.TextMatrix(llRow, LDEINDEX) = rst_Eht!ehtLDE
                    grdSpec.TextMatrix(llRow, LEADINDEX) = rst_Eht!ehtLeadTime
                    grdSpec.TextMatrix(llRow, CYCLEINDEX) = rst_Eht!ehtCycle
                Else
                    grdSpec.TextMatrix(llRow, LDEINDEX) = ""
                    grdSpec.TextMatrix(llRow, LEADINDEX) = ""
                    grdSpec.TextMatrix(llRow, CYCLEINDEX) = ""
                End If
                SQLQuery = "SELECT count(evtCode) FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtCode
                Set rst_Evt = gSQLSelectCall(SQLQuery)
                If Not rst_Evt.EOF Then
                    grdSpec.TextMatrix(llRow, VEHICLEINDEX) = rst_Evt(0).Value
                Else
                    grdSpec.TextMatrix(llRow, VEHICLEINDEX) = "0"
                End If
                grdSpec.TextMatrix(llRow, REFROWNOINDEX) = llRow
                grdSpec.TextMatrix(llRow, EHTCODEINDEX) = rst_Eht!ehtCode
                llEhtInfo = UBound(tgEhtInfo)
                tgEhtInfo(llEhtInfo).iRefRowNo = llRow
                tgEhtInfo(llEhtInfo).lFirstEvt = -1
                tgEhtInfo(llEhtInfo).lFirstEct = -1
                SQLQuery = "SELECT * FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtCode
                Set rst_Evt = gSQLSelectCall(SQLQuery)
                Do While Not rst_Evt.EOF
                    If tgEhtInfo(llEhtInfo).lFirstEvt = -1 Then
                        llNext = -1
                        'llEvtInfo = UBound(tgEvtInfo)
                        'tgEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
                        'tgEvtInfo(llEvtInfo).lNextEvt = -1
                    Else
                        llNext = tgEhtInfo(llEhtInfo).lFirstEvt
                        'llEvtInfo = UBound(tgEvtInfo)
                        'tgEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
                        'tgEvtInfo(llEvtInfo).lNextEvt = llNext
                    End If
                    llEvtInfo = UBound(tgEvtInfo)
                    tgEhtInfo(llEhtInfo).lFirstEvt = llEvtInfo
                    tgEvtInfo(llEvtInfo).iVefCode = rst_Evt!evtVefCode
                    tgEvtInfo(llEvtInfo).lNextEvt = llNext
                    ReDim Preserve tgEvtInfo(0 To UBound(tgEvtInfo) + 1) As EVTINFO
                    rst_Evt.MoveNext
                Loop
                
                SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & rst_Eht!ehtCode
                Set rst_Ect = gSQLSelectCall(SQLQuery)
                Do While Not rst_Ect.EOF
                    If tgEhtInfo(llEhtInfo).lFirstEct = -1 Then
                        llNext = -1
                        'llEctInfo = UBound(tgEctInfo)
                        'tgEhtInfo(llEhtInfo).lFirstEct = llEctInfo
                        'tgEctInfo(llEctInfo).lNextEct = -1
                    Else
                        llNext = tgEhtInfo(llEhtInfo).lFirstEct
                        'llEctInfo = UBound(tgEctInfo)
                        'tgEhtInfo(llEhtInfo).lFirstEct = llEctInfo
                        'tgEctInfo(llEctInfo).lNextEct = llNext
                    End If
                    llEctInfo = UBound(tgEctInfo)
                    tgEhtInfo(llEhtInfo).lFirstEct = llEctInfo
                    tgEctInfo(llEctInfo).sLogType = rst_Ect!ectLogType
                    tgEctInfo(llEctInfo).sFieldType = rst_Ect!ectFieldType
                    tgEctInfo(llEctInfo).sFieldName = rst_Ect!ectFieldName
                    tgEctInfo(llEctInfo).lFieldValue = rst_Ect!ectFieldValue
                    tgEctInfo(llEctInfo).sFieldString = rst_Ect!ectFieldString
                    tgEctInfo(llEctInfo).lNextEct = llNext
                    ReDim Preserve tgEctInfo(0 To UBound(tgEctInfo) + 1) As ECTINFO
                    rst_Ect.MoveNext
                Loop
                
                ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
                llRow = llRow + 1
            End If
            rst_Eht.MoveNext
        Loop
    End If
    On Error Resume Next
    rst_Evt.Close
    rst_Eht.Close
    On Error GoTo ErrHand
    
    mSpecSortCol TYPEINDEX
    For llRow = grdSpec.FixedRows To grdSpec.Rows - 1 Step 1
        slStr = Trim$(grdSpec.TextMatrix(llRow, TYPEINDEX))
        If slStr <> "" Then
            grdSpec.Row = llRow
            grdSpec.Col = VEHICLEINDEX
            grdSpec.CellBackColor = LIGHTGREENCOLOR 'GRAY
        End If
    Next llRow
    grdSpec.Redraw = True
    On Error GoTo 0
    Exit Function
ErrHand:
    gSetMousePointer grdSpec, grdSpec, vbDefault
    gHandleError "AffErrorLog.txt", "frmExportSpec-mPopulateGrid"
    mPopulateGrid = False
End Function


Private Function mSave() As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim llCode As Long
    Dim llEht As Long
    Dim llEvtCode As Long
    Dim llEctCode As Long
    Dim llRefRowNo As Long
    Dim llNext As Long
    
    On Error GoTo ErrHand
    If (sgExptSpec <> "Y") Or (sgUstWin(14) = "V") Then
        mSave = False
        Exit Function
    End If
    If Not mTestGridValues() Then
        mSave = False
        Exit Function
    End If
    If Not mDateSpanOk() Then
        mSave = False
        Exit Function
    End If
    
    For llRow = grdSpec.FixedRows To grdSpec.Rows - 1 Step 1
        slStr = Trim$(grdSpec.TextMatrix(llRow, TYPEINDEX))
        If slStr <> "" Then
            If Val(grdSpec.TextMatrix(llRow, EHTCODEINDEX)) <= 0 Then
                llCode = mAddSpec(grdSpec.TextMatrix(llRow, NAMECODEINDEX), Trim$(grdSpec.TextMatrix(llRow, NAMEINDEX)))
                grdSpec.TextMatrix(llRow, EHTCODEINDEX) = llCode
            End If
            If Val(grdSpec.TextMatrix(llRow, EHTCODEINDEX)) > 0 Then
                
                SQLQuery = "Update eht_Export_Header Set "
                SQLQuery = SQLQuery & "ehtExportType = '" & gFixQuote(grdSpec.TextMatrix(llRow, NAMECODEINDEX)) & "', "
                SQLQuery = SQLQuery & "ehtSubType = '" & gFixQuote("S") & "', "
                SQLQuery = SQLQuery & "ehtExportName = '" & gFixQuote(grdSpec.TextMatrix(llRow, NAMEINDEX)) & "', "
                SQLQuery = SQLQuery & "ehtUstCode = " & igUstCode & ", "
                If Trim$(grdSpec.TextMatrix(llRow, LDEINDEX)) <> "" Then
                    SQLQuery = SQLQuery & "ehtLDE = '" & Format$(Trim$(grdSpec.TextMatrix(llRow, LDEINDEX)), sgSQLDateForm) & "', "
                Else
                    SQLQuery = SQLQuery & "ehtLDE = '" & Format$("1/1/1970", sgSQLDateForm) & "', "
                End If
                If Trim$(grdSpec.TextMatrix(llRow, LEADINDEX)) <> "" Then
                    SQLQuery = SQLQuery & "ehtLeadTime = " & grdSpec.TextMatrix(llRow, LEADINDEX) & ", "
                Else
                    SQLQuery = SQLQuery & "ehtLeadTime = " & "0" & ", "
                End If
                If Trim$(grdSpec.TextMatrix(llRow, CYCLEINDEX)) <> "" Then
                    SQLQuery = SQLQuery & "ehtCycle = " & grdSpec.TextMatrix(llRow, CYCLEINDEX) & ", "
                Else
                    SQLQuery = SQLQuery & "ehtCycle = " & "0" & ", "
                End If
                SQLQuery = SQLQuery & "ehtUnused = '" & "" & "' "
                SQLQuery = SQLQuery & " WHERE ehtCode = " & grdSpec.TextMatrix(llRow, EHTCODEINDEX)
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    gSetMousePointer grdSpec, grdSpec, vbDefault
                    gHandleError "AffErrorLog.txt", "frmExportSpec-mSave"
                    mSave = False
                    Exit Function
                End If
            End If
            SQLQuery = "DELETE FROM evt_Export_Vehicles"
            SQLQuery = SQLQuery & " WHERE (EvtEhtCode = " & grdSpec.TextMatrix(llRow, EHTCODEINDEX) & ")"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                gSetMousePointer grdSpec, grdSpec, vbDefault
                gHandleError "AffErrorLog.txt", "frmExportSpec-mSave"
                mSave = False
                Exit Function
            End If
            llRefRowNo = Val(grdSpec.TextMatrix(llRow, REFROWNOINDEX))
            For llEht = 0 To UBound(tgEhtInfo) - 1 Step 1
                If tgEhtInfo(llEht).iRefRowNo = llRefRowNo Then
                    lgExportEhtInfoIndex = llEht
                    llNext = tgEhtInfo(llEht).lFirstEvt
                    Do While llNext <> -1
                        SQLQuery = "Insert Into evt_Export_Vehicles ( "
                        SQLQuery = SQLQuery & "evtCode, "
                        SQLQuery = SQLQuery & "evtEhtCode, "
                        SQLQuery = SQLQuery & "evtVefCode, "
                        SQLQuery = SQLQuery & "evtUnused "
                        SQLQuery = SQLQuery & ") "
                        SQLQuery = SQLQuery & "Values ( "
                        SQLQuery = SQLQuery & "Replace" & ", "
                        SQLQuery = SQLQuery & grdSpec.TextMatrix(llRow, EHTCODEINDEX) & ", "
                        SQLQuery = SQLQuery & tgEvtInfo(llNext).iVefCode & ", "
                        SQLQuery = SQLQuery & "'" & "" & "' "
                        SQLQuery = SQLQuery & ") "
                        llEvtCode = gInsertAndReturnCode(SQLQuery, "evt_Export_Vehicles", "evtCode", "Replace")
                        llNext = tgEvtInfo(llNext).lNextEvt
                    Loop
                    Exit For
                End If
            Next llEht
            SQLQuery = "DELETE FROM ect_Export_Criteria"
            SQLQuery = SQLQuery & " WHERE (EctEhtCode = " & grdSpec.TextMatrix(llRow, EHTCODEINDEX) & ")"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                gSetMousePointer grdSpec, grdSpec, vbDefault
                gHandleError "AffErrorLog.txt", "frmExportSpec-mSave"
                mSave = False
                Exit Function
            End If
            llRefRowNo = Val(grdSpec.TextMatrix(llRow, REFROWNOINDEX))
            For llEht = 0 To UBound(tgEhtInfo) - 1 Step 1
                If tgEhtInfo(llEht).iRefRowNo = llRefRowNo Then
                    lgExportEhtInfoIndex = llEht
                    llNext = tgEhtInfo(llEht).lFirstEct
                    Do While llNext <> -1
                        
                        SQLQuery = "Insert Into ect_Export_Criteria ( "
                        SQLQuery = SQLQuery & "ectCode, "
                        SQLQuery = SQLQuery & "ectEhtCode, "
                        SQLQuery = SQLQuery & "ectLogType, "
                        SQLQuery = SQLQuery & "ectFieldType, "
                        SQLQuery = SQLQuery & "ectFieldName, "
                        SQLQuery = SQLQuery & "ectFieldValue, "
                        SQLQuery = SQLQuery & "ectFieldString, "
                        SQLQuery = SQLQuery & "ectUnused "
                        SQLQuery = SQLQuery & ") "
                        SQLQuery = SQLQuery & "Values ( "
                        SQLQuery = SQLQuery & "Replace" & ", "
                        SQLQuery = SQLQuery & grdSpec.TextMatrix(llRow, EHTCODEINDEX) & ", "
                        SQLQuery = SQLQuery & "'" & gFixQuote(tgEctInfo(llNext).sLogType) & "', "
                        SQLQuery = SQLQuery & "'" & gFixQuote(tgEctInfo(llNext).sFieldType) & "', "
                        SQLQuery = SQLQuery & "'" & gFixQuote(tgEctInfo(llNext).sFieldName) & "', "
                        SQLQuery = SQLQuery & tgEctInfo(llNext).lFieldValue & ", "
                        SQLQuery = SQLQuery & "'" & gFixQuote(tgEctInfo(llNext).sFieldString) & "', "
                        SQLQuery = SQLQuery & "'" & "" & "' "
                        SQLQuery = SQLQuery & ") "
                        
                        llEctCode = gInsertAndReturnCode(SQLQuery, "ect_Export_Criteria", "ectCode", "Replace")
                        llNext = tgEctInfo(llNext).lNextEct
                    Loop
                    Exit For
                End If
            Next llEht
        End If
    Next llRow
    For llRow = 0 To UBound(lmDelEhtCode) - 1 Step 1
        SQLQuery = "DELETE FROM evt_Export_Vehicles"
        SQLQuery = SQLQuery & " WHERE (EvtEhtCode = " & lmDelEhtCode(llRow) & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gSetMousePointer grdSpec, grdSpec, vbDefault
            gHandleError "AffErrorLog.txt", "frmExportSpec-mSave"
            mSave = False
            Exit Function
        End If
        SQLQuery = "DELETE FROM ect_Export_Criteria"
        SQLQuery = SQLQuery & " WHERE (EctEhtCode = " & lmDelEhtCode(llRow) & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gSetMousePointer grdSpec, grdSpec, vbDefault
            gHandleError "AffErrorLog.txt", "frmExportSpec-mSave"
            mSave = False
            Exit Function
        End If
        SQLQuery = "DELETE FROM eqt_Export_Queue"
        SQLQuery = SQLQuery & " WHERE (EqtEhtCode = " & lmDelEhtCode(llRow) & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gSetMousePointer grdSpec, grdSpec, vbDefault
            gHandleError "AffErrorLog.txt", "frmExportSpec-mSave"
            mSave = False
            Exit Function
        End If
        SQLQuery = "DELETE FROM eht_Export_Header"
        SQLQuery = SQLQuery & " WHERE (ehtCode = " & lmDelEhtCode(llRow) & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            gSetMousePointer grdSpec, grdSpec, vbDefault
            gHandleError "AffErrorLog.txt", "frmExportSpec-mSave"
            mSave = False
            Exit Function
        End If
    Next llRow
    'ReDim lmNewEhtCode(0 To 0) As Long
    ReDim lmDelEhtCode(0 To 0) As Long
    imFieldChgd = False
    igVehicleSpecChgFlag = False
    mSave = True
    On Error GoTo 0
    Exit Function
ErrHand:
    gSetMousePointer grdSpec, grdSpec, vbDefault
    gHandleError "AffErrorLog.txt", "frmExportSpec-mSave"
    mSave = False
End Function

Private Function mTestGridValues()
    Dim iLoop As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilStartTimeOk As Integer
    Dim ilEndTimeOk As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilError As Integer
    
    grdSpec.Redraw = False
    'Test if fields defined
    ilError = False
    For llRow = grdSpec.FixedRows To grdSpec.Rows - 1 Step 1
        slStr = Trim$(grdSpec.TextMatrix(llRow, TYPEINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdSpec.TextMatrix(llRow, NAMEINDEX))
            If slStr = "" Or slStr = "Missing" Then
                ilError = True
                grdSpec.TextMatrix(llRow, NAMEINDEX) = "Missing"
                grdSpec.Row = llRow
                grdSpec.Col = NAMEINDEX
                grdSpec.CellForeColor = vbRed
            End If
            slStr = Trim$(grdSpec.TextMatrix(llRow, LDEINDEX))
            If slStr = "" Or slStr = "Missing" Then
                ilError = True
                grdSpec.TextMatrix(llRow, LDEINDEX) = "Missing"
                grdSpec.Row = llRow
                grdSpec.Col = LDEINDEX
                grdSpec.CellForeColor = vbRed
            End If
            If Trim$(grdSpec.TextMatrix(llRow, LEADINDEX)) = "" Or Trim$(grdSpec.TextMatrix(llRow, LEADINDEX)) = "Missing" Then
                ilError = True
                grdSpec.TextMatrix(llRow, LEADINDEX) = "Missing"
                grdSpec.Row = llRow
                grdSpec.Col = LEADINDEX
                grdSpec.CellForeColor = vbRed
            End If
            If Trim$(grdSpec.TextMatrix(llRow, CYCLEINDEX)) = "" Or Trim$(grdSpec.TextMatrix(llRow, CYCLEINDEX)) = "Missing" Then
                ilError = True
                grdSpec.TextMatrix(llRow, CYCLEINDEX) = "Missing"
                grdSpec.Row = llRow
                grdSpec.Col = CYCLEINDEX
                grdSpec.CellForeColor = vbRed
            End If
            If Trim$(grdSpec.TextMatrix(llRow, VEHICLEINDEX)) = "" Or Trim$(grdSpec.TextMatrix(llRow, VEHICLEINDEX)) = "Missing" Then
                ilError = True
                grdSpec.TextMatrix(llRow, VEHICLEINDEX) = "Missing"
                grdSpec.Row = llRow
                grdSpec.Col = VEHICLEINDEX
                grdSpec.CellForeColor = vbRed
            ElseIf Val(Trim$(grdSpec.TextMatrix(llRow, VEHICLEINDEX))) <= 0 Then
                ilError = True
                grdSpec.TextMatrix(llRow, VEHICLEINDEX) = "Missing"
                grdSpec.Row = llRow
                grdSpec.Col = VEHICLEINDEX
                grdSpec.CellForeColor = vbRed
            End If
        End If
    Next llRow
    If ilError Then
        grdSpec.Redraw = True
        mTestGridValues = False
        'Screen.MousePointer = vbDefault
        gSetMousePointer grdSpec, grdSpec, vbDefault
        Exit Function
    Else
        grdSpec.Redraw = True
        mTestGridValues = True
        Exit Function
    End If
End Function

Private Function mColOk() As Integer
    mColOk = True
    If grdSpec.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function


Private Function mAddSpec(slExportType As String, slExportName As String) As Long
    Dim llEhtCode As Long
    
    On Error GoTo ErrHand
    SQLQuery = "Insert Into eht_Export_Header ( "
    SQLQuery = SQLQuery & "ehtCode, "
    SQLQuery = SQLQuery & "ehtExportType, "
    SQLQuery = SQLQuery & "ehtSubType, "
    SQLQuery = SQLQuery & "ehtStandardEhtCode, "
    SQLQuery = SQLQuery & "ehtExportName, "
    SQLQuery = SQLQuery & "ehtUstCode, "
    SQLQuery = SQLQuery & "ehtLDE, "
    SQLQuery = SQLQuery & "ehtLeadTime, "
    SQLQuery = SQLQuery & "ehtCycle, "
    SQLQuery = SQLQuery & "ehtUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slExportType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote("S") & "', "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slExportName) & "', "
    SQLQuery = SQLQuery & igUstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llEhtCode = gInsertAndReturnCode(SQLQuery, "eht_Export_Header", "ehtCode", "Replace")
    mAddSpec = llEhtCode
    Exit Function
ErrHand:
    gSetMousePointer grdSpec, grdSpec, vbDefault
    gHandleError "AffErrorLog.txt", "frmExportSpec-mAddSpec"
    mAddSpec = 0
End Function

Private Sub mDeleteNewEht()
    Dim ilLoop As Integer
    
    On Error GoTo ErrHand
    
    'For ilLoop = 0 To UBound(lmNewEhtCode) - 1 Step 1
    '    SQLQuery = "DELETE FROM evt_Export_Vehicles"
    '    SQLQuery = SQLQuery & " WHERE (EvtEhtCode = " & lmNewEhtCode(ilLoop) & ")"
    '    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
    '        GoSub ErrHand:
    '    End If
    '    SQLQuery = "DELETE FROM eht_Export_Header"
    '    SQLQuery = SQLQuery & " WHERE (EhtCode = " & lmNewEhtCode(ilLoop) & ")"
    '    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
    '        GoSub ErrHand:
    '    End If
    'Next ilLoop
    Exit Sub
ErrHand:
    gSetMousePointer grdSpec, grdSpec, vbDefault
    gHandleError "AffErrorLog.txt", "frmExportSpec-mDeleteNewEht"
End Sub

Private Sub mGetVehicles(blAlwaysBranch As Boolean, llRow As Long)
    Dim llEht As Long
    Dim llRefRowNo As Long
    Dim ilCount As Integer
    Dim llNext As Long
    
    On Error GoTo ErrHand
    
    sgExportTypeChar = grdSpec.TextMatrix(llRow, NAMECODEINDEX)
    sgExportName = grdSpec.TextMatrix(llRow, NAMEINDEX)
    lgExportEhtCode = grdSpec.TextMatrix(llRow, EHTCODEINDEX)
    If (blAlwaysBranch) Then
        'if (lgExportEhtCode <= 0) Then
        '    lgExportEhtCode = mAddSpec(sgExportTypeChar, Trim$(sgExportName))
        '    grdSpec.TextMatrix(grdSpec.Row, EHTCODEINDEX) = lgExportEhtCode
        '    lmNewEhtCode(UBound(lmNewEhtCode)) = lgExportEhtCode
        '    ReDim Preserve lmNewEhtCode(0 To UBound(lmNewEhtCode) + 1) As Long
        'End If
        llRefRowNo = Val(grdSpec.TextMatrix(llRow, REFROWNOINDEX))
        lgExportEhtInfoIndex = -1
        For llEht = 0 To UBound(tgEhtInfo) - 1 Step 1
            If tgEhtInfo(llEht).iRefRowNo = llRefRowNo Then
                lgExportEhtInfoIndex = llEht
                Exit For
            End If
        Next llEht
        If lgExportEhtInfoIndex = -1 Then
            lgExportEhtInfoIndex = UBound(tgEhtInfo)
            tgEhtInfo(lgExportEhtInfoIndex).iRefRowNo = llRefRowNo
            tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt = -1
            tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = -1
            ReDim Preserve tgEhtInfo(0 To UBound(tgEhtInfo) + 1) As EHTINFO
        End If
        frmVehicleSpec.Show vbModal
        If igVehicleSpecChgFlag Then
            imFieldChgd = True
        End If
        'SQLQuery = "SELECT Count(evtCode) FROM evt_Export_Vehicles WHERE evtEhtCode = " & lgExportEhtCode
        'Set rst_Evt = gSQLSelectCall(SQLQuery)
        'If Not rst_Evt.EOF Then
        '    grdSpec.TextMatrix(llRow, VEHICLEINDEX) = rst_Evt(0).Value
        'Else
        '    grdSpec.TextMatrix(llRow, VEHICLEINDEX) = "0"
        'End If
        'rst_Evt.Close
        grdSpec.Row = llRow
        grdSpec.Col = VEHICLEINDEX
        ilCount = 0
        llNext = tgEhtInfo(lgExportEhtInfoIndex).lFirstEvt
        Do While llNext <> -1
            ilCount = ilCount + 1
            llNext = tgEvtInfo(llNext).lNextEvt
        Loop
        grdSpec.Text = ilCount
        grdSpec.CellBackColor = LIGHTGREENCOLOR 'GRAY
    End If

    Exit Sub
ErrHand:
    gSetMousePointer grdSpec, grdSpec, vbDefault
    gHandleError "AffErrorLog.txt", "frmExportSpec-mGetVehicles"
End Sub

Private Function mDateSpanOk() As Boolean
    Dim llRow As Long
    Dim slStr As String
    Dim blError As Boolean
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilLoop As Integer
    Dim slError As String
    
    blError = True
    slError = ""
    For llRow = grdSpec.FixedRows To grdSpec.Rows - 1 Step 1
        slStr = Trim$(grdSpec.TextMatrix(llRow, TYPEINDEX))
        If slStr <> "" Then
            For ilLoop = LBound(tgSpecInfo) To UBound(tgSpecInfo) Step 1
                If tgSpecInfo(ilLoop).sType = grdSpec.TextMatrix(llRow, NAMECODEINDEX) Then
                    If tgSpecInfo(ilLoop).sCheckDateSpan = "Y" Then
                        slStartDate = DateAdd("D", 1, grdSpec.TextMatrix(llRow, LDEINDEX))
                        slEndDate = DateAdd("D", grdSpec.TextMatrix(llRow, CYCLEINDEX) - 1, slStartDate)
                        If gWeekDayLong(gDateValue(slEndDate)) <= gWeekDayLong(gDateValue(slStartDate)) Then
                            blError = False
                            If slError = "" Then
                                slError = Trim$(tgSpecInfo(ilLoop).sName)
                            Else
                                slError = slError & ", " & Trim$(tgSpecInfo(ilLoop).sName)
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next ilLoop
        End If
    Next llRow
    mDateSpanOk = blError
    If Not blError Then
        MsgBox slError & " date(s) can't span Sunday", vbCritical + vbOKOnly
    End If
End Function
