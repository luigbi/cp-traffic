VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAstBuildQueueStatus 
   Caption         =   "Station Spot Builder Status"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "AffAstBuildQueueStatus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   9360
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
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
      TabIndex        =   4
      Top             =   4665
      Width           =   60
   End
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1335
      TabIndex        =   3
      Top             =   2310
      Visible         =   0   'False
      Width           =   945
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
      TabIndex        =   5
      Top             =   5055
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      Height          =   315
      Left            =   4110
      TabIndex        =   0
      Top             =   5145
      Width           =   1245
   End
   Begin VB.CommandButton cmcRefresh 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   7995
      TabIndex        =   1
      Top             =   105
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
   Begin V81Affiliate.CSI_Calendar txtToDate 
      Height          =   300
      Left            =   5445
      TabIndex        =   7
      Top             =   105
      Width           =   1785
      _ExtentX        =   2143
      _ExtentY        =   661
      BorderStyle     =   1
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
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin V81Affiliate.CSI_Calendar txtFromDate 
      Height          =   300
      Left            =   1785
      TabIndex        =   8
      Top             =   105
      Width           =   1875
      _ExtentX        =   2143
      _ExtentY        =   661
      BorderStyle     =   1
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
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStatus 
      Height          =   4185
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   660
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7382
      _Version        =   393216
      Rows            =   4
      Cols            =   9
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
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblStartDate 
      Caption         =   "Entered From Date"
      Height          =   225
      Left            =   330
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblEndDate 
      Caption         =   "Entered To Date"
      Height          =   240
      Left            =   3990
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAstBuildQueueStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private imBSMode As Integer

'Grid Controls
Private imCtrlVisible As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on

Private imLastColSorted As Integer
Private imLastSort As Integer

Private lm1970 As Long

Private rst_abf As ADODB.Recordset
Private rst_site As ADODB.Recordset

Private Const VEHICLEINDEX = 0
Private Const STATIONINDEX = 1
Private Const STATUSINDEX = 2
Private Const SOURCEINDEX = 3
Private Const GENDATEINDEX = 4
Private Const ENTEREDINDEX = 5
Private Const COMPLETEDINDEX = 6
Private Const SORTINDEX = 7
Private Const ABFCODEINDEX = 8

Private Sub cmcDone_Click()
    Unload frmAstBuildQueueStatus
End Sub

Private Sub cmcDone_GotFocus()
    mSetShow
End Sub

Private Sub cmcRefresh_Click()
    imLastColSorted = -1
    imLastSort = -1
    mPopulate
End Sub

Private Sub Form_Click()
    mSetShow
    cmcDone.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 1.2
    Me.Height = (Screen.Height) / 1.5
    Me.Top = (Screen.Height - Me.Height) / 1.4
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
    gCenterForm Me
End Sub

Private Sub Form_Load()

    gSetMousePointer grdStatus, grdStatus, vbHourglass
    'Load Grid
    imBSMode = False
    imCtrlVisible = False
    imLastColSorted = -1
    imLastSort = -1
    tmcFillGrid.Enabled = True
    lm1970 = gDateValue("1/1/1970")
    pbcSTab.Left = -2 * pbcSTab.Width
    pbcTab.Left = -2 * pbcTab.Width
    gSetMousePointer grdStatus, grdStatus, vbDefault

End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    gSetMousePointer grdStatus, grdStatus, vbHourglass
    mSetGridColumns
    mSetGridTitles
    'mPopulate
    gSetMousePointer grdStatus, grdStatus, vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAstBuildQueueStatus = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdStatus.ColWidth(ABFCODEINDEX) = 0
    grdStatus.ColWidth(SORTINDEX) = 0
    grdStatus.ColWidth(STATIONINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(STATUSINDEX) = grdStatus.Width * 0.08
    grdStatus.ColWidth(SOURCEINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(ENTEREDINDEX) = grdStatus.Width * 0.13
    grdStatus.ColWidth(GENDATEINDEX) = grdStatus.Width * 0.13
    grdStatus.ColWidth(COMPLETEDINDEX) = grdStatus.Width * 0.13
    grdStatus.ColWidth(VEHICLEINDEX) = grdStatus.Width - GRIDSCROLLWIDTH - 15
    For ilCol = VEHICLEINDEX To SORTINDEX Step 1
        If ilCol <> VEHICLEINDEX Then
            grdStatus.ColWidth(VEHICLEINDEX) = grdStatus.ColWidth(VEHICLEINDEX) - grdStatus.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdStatus
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdStatus.TextMatrix(0, VEHICLEINDEX) = "Vehicle Name"
    grdStatus.TextMatrix(0, STATIONINDEX) = "Station"
    grdStatus.TextMatrix(0, STATUSINDEX) = "Status"
    grdStatus.TextMatrix(0, SOURCEINDEX) = "Source"
    grdStatus.TextMatrix(0, GENDATEINDEX) = "Spot Dates"
    grdStatus.TextMatrix(0, ENTEREDINDEX) = "Entered"
    grdStatus.TextMatrix(0, COMPLETEDINDEX) = "Completed"

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim llVef As Long
    Dim ilShtt As Integer
    
    On Error GoTo ErrHand
    gSetMousePointer grdStatus, grdStatus, vbHourglass
    gGrid_Clear grdStatus, True
    If Trim$(txtFromDate.Text) = "" Then
        txtFromDate.Text = gObtainPrevMonday(Format(Now, "m/d/yy"))
    End If
    If Trim$(txtToDate.Text) = "" Then
        txtToDate.Text = Format(Now, "m/d/yy")
    End If
    grdStatus.Row = 0
    For llCol = VEHICLEINDEX To COMPLETEDINDEX Step 1
        grdStatus.Col = llCol
        grdStatus.CellBackColor = LIGHTBLUE
    Next llCol
    llRow = grdStatus.FixedRows
    grdStatus.Redraw = False
    
    SQLQuery = "Select "
    SQLQuery = SQLQuery & "* "
    SQLQuery = SQLQuery & "From abf_AST_Build_Queue "
    SQLQuery = SQLQuery & "Where "
    SQLQuery = SQLQuery & " abfEnteredDate >= '" & Format(txtFromDate.Text, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " And abfEnteredDate <= '" & Format(txtToDate.Text, sgSQLDateForm) & "'"
    Set rst_abf = gSQLSelectCall(SQLQuery)
    Do While Not rst_abf.EOF
        If llRow >= grdStatus.Rows Then
            grdStatus.AddItem ""
        End If
        grdStatus.Row = llRow
        For llCol = VEHICLEINDEX To COMPLETEDINDEX Step 1
            grdStatus.Col = llCol
            grdStatus.CellBackColor = LIGHTYELLOW
        Next llCol
        llVef = gBinarySearchVef(CLng(rst_abf!abfVefCode))
        If llVef <> -1 Then
            grdStatus.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
        Else
            grdStatus.TextMatrix(llRow, VEHICLEINDEX) = "Vehicle Code = " & rst_abf!abfVefCode
        End If
        If rst_abf!abfShttCode > 0 Then
            ilShtt = gBinarySearchStationInfoByCode(rst_abf!abfShttCode)
            If ilShtt <> -1 Then
                grdStatus.TextMatrix(llRow, STATIONINDEX) = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
            Else
                grdStatus.TextMatrix(llRow, STATIONINDEX) = "Station Code = " & rst_abf!abfShttCode
            End If
        End If
        If rst_abf!abfSource = "A" Then
            grdStatus.TextMatrix(llRow, SOURCEINDEX) = "Agreement"
        ElseIf rst_abf!abfSource = "P" Then
            grdStatus.TextMatrix(llRow, SOURCEINDEX) = "Post Log"
        ElseIf rst_abf!abfSource = "F" Then
            grdStatus.TextMatrix(llRow, SOURCEINDEX) = "Fast Add"
        Else
            grdStatus.TextMatrix(llRow, SOURCEINDEX) = "Log"
        End If
        If rst_abf!abfStatus = "P" Then
            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Processing"
        ElseIf rst_abf!abfStatus = "H" Then
            grdStatus.TextMatrix(llRow, STATUSINDEX) = "On Hold"
        ElseIf rst_abf!abfStatus = "C" Then
            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Completed"
        ElseIf rst_abf!abfStatus = "G" Then
            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Ready"
        End If
        grdStatus.TextMatrix(llRow, GENDATEINDEX) = Format(rst_abf!abfGenStartDate, sgShowDateForm) & "-" & Format(rst_abf!abfGenEndDate, sgShowDateForm)
        grdStatus.TextMatrix(llRow, ENTEREDINDEX) = Format(rst_abf!abfEnteredDate, sgShowDateForm) & " " & Format(rst_abf!abfEnteredTime, sgShowTimeWSecForm)
        If gDateValue(Format(rst_abf!abfCompletedDate, sgShowDateForm)) <> gDateValue("12/31/2069") Then
            grdStatus.TextMatrix(llRow, COMPLETEDINDEX) = Format(rst_abf!abfCompletedDate, sgShowDateForm) & " " & Format(rst_abf!abfCompletedTime, sgShowTimeWSecForm)
        End If
        grdStatus.TextMatrix(llRow, ABFCODEINDEX) = rst_abf!abfCode
        llRow = llRow + 1
        rst_abf.MoveNext
    Loop
    mStatusSortCol ENTEREDINDEX
    mStatusSortCol ENTEREDINDEX
    'mSetStatusGridColor
    gSetMousePointer grdStatus, grdStatus, vbDefault
    grdStatus.Redraw = True
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "StationBuildQueueStatusLog.txt", "StationSpotBuilderQueue-mPopulate"
    grdStatus.Redraw = True
    Resume Next
End Sub


Private Sub mSetStatusGridColor()
    Dim llRow As Long
    Dim llCol As Long
    
    'gGrid_Clear grdStatus, True
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        For llCol = VEHICLEINDEX To COMPLETEDINDEX Step 1
            grdStatus.Row = llRow
            grdStatus.Col = llCol
            If llCol = STATUSINDEX Then
                If grdStatus.TextMatrix(llRow, llCol) = "Not Ready" Then
                    grdStatus.CellBackColor = vbWhite
                Else
                    grdStatus.CellBackColor = LIGHTYELLOW
                End If
            Else
                grdStatus.CellBackColor = LIGHTYELLOW
            End If
        Next llCol
    Next llRow
End Sub

Private Sub mClearGrid()
    gGrid_Clear grdStatus, True
End Sub

Private Sub grdStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'grdStatus.ToolTipText = ""
    If (grdStatus.MouseRow >= grdStatus.FixedRows) And (grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol)) <> "" Then
        grdStatus.ToolTipText = grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol)
    Else
        grdStatus.ToolTipText = ""
    End If
End Sub


Private Sub grdStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdStatus.TopRow
    grdStatus.Redraw = False
End Sub

Private Sub grdStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim ilType As Integer
    
    If Y < grdStatus.RowHeight(0) Then
        grdStatus.Row = 0   'grdStatus.MouseRow
        grdStatus.Col = grdStatus.MouseCol
        If grdStatus.CellBackColor = LIGHTBLUE Then
            gSetMousePointer grdStatus, grdStatus, vbHourglass
            mStatusSortCol grdStatus.Col
            grdStatus.Row = 0
            grdStatus.Col = ABFCODEINDEX
            gSetMousePointer grdStatus, grdStatus, vbDefault
        End If
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdStatus, X, Y)
    If Not ilFound Then
        grdStatus.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    
    If Trim$(grdStatus.TextMatrix(grdStatus.Row, VEHICLEINDEX)) = "" Then
        grdStatus.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    If Not mColOk() Then
        grdStatus.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    lmTopRow = grdStatus.TopRow
    grdStatus.Redraw = True
    mEnableBox
End Sub

Private Sub grdStatus_Scroll()
    If grdStatus.Redraw = False Then
        grdStatus.Redraw = True
        grdStatus.TopRow = lmTopRow
        grdStatus.Refresh
        grdStatus.Redraw = False
    End If
    'If (imCtrlVisible) And (grdStatus.Row >= grdStatus.FixedRows) And (grdStatus.Col >= TOINDEX) And (grdStatus.Col < grdStatus.Cols - 1) Then
        ''If grdStatus.RowIsVisible(grdStatus.Row) Then
        ''   If grdStatus.Col = TOINDEX Then
        ''        edcDropdown.Move grdStatus.Left + grdStatus.ColPos(grdStatus.Col) + 30, grdStatus.Top + grdStatus.RowPos(grdStatus.Row) + 15, grdStatus.ColWidth(grdStatus.Col) - cmcDropDown.Width - 30, grdStatus.RowHeight(grdStatus.Row) - 15
        ''        cmcDropDown.Move edcDropdown.Left + edcDropdown.Width, edcDropdown.Top, cmcDropDown.Width, edcDropdown.Height
        ''        lbcToName.Move edcDropdown.Left, edcDropdown.Top + edcDropdown.Height, edcDropdown.Width + cmcDropDown.Width
        ''        edcDropdown.Visible = True
        ''        cmcDropDown.Visible = True
        ''        lbcToName.Visible = True
        ''        edcDropdown.SetFocus
        ''    End If
        ''Else
        ''    edcDropdown.Visible = False
        ''    cmcDropDown.Visible = False
        ''    lbcToName.Visible = False
        ''End If
        mSetShow
        cmcDone.SetFocus
    'End If

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
            Select Case grdStatus.Col
                Case Else
                    grdStatus.Col = grdStatus.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        grdStatus.TopRow = grdStatus.FixedRows
        grdStatus.Col = STATUSINDEX
        Do
            If grdStatus.Row <= grdStatus.FixedRows Then
                cmcDone.SetFocus
                Exit Sub
            End If
            grdStatus.Row = grdStatus.Rows - 1
            Do
                If Not grdStatus.RowIsVisible(grdStatus.Row) Then
                    grdStatus.TopRow = grdStatus.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
            If mColOk() Then
                Exit Do
            End If
        Loop
    End If
    lmTopRow = grdStatus.TopRow
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
            Select Case grdStatus.Col
                Case Else
                    grdStatus.Col = grdStatus.Col - 1
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
        grdStatus.TopRow = grdStatus.FixedRows
        grdStatus.Row = grdStatus.FixedRows
        grdStatus.Col = STATUSINDEX
        Do
            If mColOk() Then
                Exit Do
            End If
            If grdStatus.Row + 1 >= grdStatus.Rows Then
                cmcDone.SetFocus
                Exit Sub
            End If
            grdStatus.Row = grdStatus.Row + 1
            Do
                If Not grdStatus.RowIsVisible(grdStatus.Row) Then
                    grdStatus.TopRow = grdStatus.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
        Loop
    End If
    lmTopRow = grdStatus.TopRow
    mEnableBox
End Sub

Private Function mColOk() As Integer
    mColOk = True
    If grdStatus.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function

Private Sub mEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    If (grdStatus.Row >= grdStatus.FixedRows) And (grdStatus.Row < grdStatus.Rows) And (grdStatus.Col = STATUSINDEX) Then
        lmEnableRow = grdStatus.Row
        lmEnableCol = grdStatus.Col
        imCtrlVisible = True
        Select Case grdStatus.Col
        End Select
    End If
End Sub

Private Sub mSetShow()
    Dim slStr As String
    
    If (lmEnableRow >= grdStatus.FixedRows) And (lmEnableRow < grdStatus.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    edcDropdown.Visible = False
End Sub


Private Sub mAdjustStatus(llRow As Long)
    On Error GoTo ErrHand
    SQLQuery = "UPDATE abf_AST_Build_Queue SET "
    If grdStatus.TextMatrix(llRow, STATUSINDEX) = "Ready" Then
        SQLQuery = SQLQuery & "abfStatus = " & "'G'"
    ElseIf grdStatus.TextMatrix(llRow, STATUSINDEX) = "Not Ready" Then
        SQLQuery = SQLQuery & "abfStatus = " & "'H'"
    End If
    SQLQuery = SQLQuery & "WHERE abfCode = " & grdStatus.TextMatrix(llRow, ABFCODEINDEX)
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError "StationBuildQueueStatusLog.txt", "StationSpotBuilderQueueStatus-mAdjustStatus"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    gHandleError "StationBuildQueueStatusLog.txt", "StationSpotBuilderQueueStatus-mAdjustStatus"
    Exit Sub
'ErrHand1:
'    gHandleError "StationBuildQueueStatusLog.txt", "StationSpotBuilderQueueStatus-mAdjustStatus"
    
End Sub


Private Sub mStatusSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        slStr = Trim$(grdStatus.TextMatrix(llRow, VEHICLEINDEX))
        If slStr <> "" Then
            If ilCol = GENDATEINDEX Then
                slStr = Trim$(grdStatus.TextMatrix(llRow, GENDATEINDEX))
                ilPos = InStr(1, slStr, "-", vbTextCompare)
                If ilPos > 0 Then
                    slDate = Left(slStr, ilPos - 1)
                    slSort = Trim$(Str$(gDateValue(slDate)))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                Else
                    slSort = "      "
                End If
            ElseIf ilCol = ENTEREDINDEX Then
                slStr = Trim$(grdStatus.TextMatrix(llRow, ENTEREDINDEX))
                ilPos = InStr(1, slStr, " ", vbTextCompare)
                If ilPos > 0 Then
                    slDate = Left(slStr, ilPos - 1)
                    slTime = Mid(slStr, ilPos + 1)
                    slSort = Trim$(Str$(gDateValue(slDate)))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                    slStr = Trim$(Str$(gTimeToLong(slTime, False)))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    slSort = slSort & slStr
                Else
                    slSort = "            "
                End If
            ElseIf ilCol = COMPLETEDINDEX Then
                slStr = Trim$(grdStatus.TextMatrix(llRow, COMPLETEDINDEX))
                ilPos = InStr(1, slStr, " ", vbTextCompare)
                If ilPos > 0 Then
                    slDate = Left(slStr, ilPos - 1)
                    slTime = Mid(slStr, ilPos + 1)
                    slSort = Trim$(Str$(gDateValue(slDate)))
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                    slStr = Trim$(Str$(gTimeToLong(slTime, False)))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    slSort = slSort & slStr
                Else
                    slSort = "            "
                End If
            Else
                slSort = UCase$(Trim$(grdStatus.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = Chr(32)
                End If
            End If
            slStr = grdStatus.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdStatus.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdStatus.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdStatus, VEHICLEINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub


Private Sub tmcFillGrid_Timer()
    tmcFillGrid.Enabled = False
    mPopulate
End Sub

