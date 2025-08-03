VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReportQueueStatus 
   Caption         =   "Report Queu Status"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "AffReportQueueStatus.frx":0000
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
      Left            =   3285
      TabIndex        =   0
      Top             =   5115
      Width           =   1245
   End
   Begin VB.CommandButton cmcRefresh 
      Caption         =   "Refresh"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStatus 
      Height          =   4605
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   165
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   8123
      _Version        =   393216
      Rows            =   4
      Cols            =   11
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmReportQueueStatus"
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

Private rst_Rct As ADODB.Recordset
Private rst_rqt As ADODB.Recordset
Private rst_Ust As ADODB.Recordset

Private Const REPORTNAMEINDEX = 0
Private Const USERNAMEINDEX = 1
Private Const PRIORITYINDEX = 2
Private Const STATUEINDEX = 3
Private Const REQUESTEDINDEX = 4
Private Const STARTEDINDEX = 5
Private Const COMPLETEDINDEX = 6
Private Const DESCRIPTIONINDEX = 7
Private Const DELETEINDEX = 8
Private Const SORTINDEX = 9
Private Const RQTCODEINDEX = 10






Private Sub cmcDone_Click()
    Unload frmReportQueueStatus
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
    gSetMousePointer grdStatus, grdStatus, vbDefault

End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    gSetMousePointer grdStatus, grdStatus, vbHourglass
    mSetGridColumns
    mSetGridTitles
    mPopulate
    gSetMousePointer grdStatus, grdStatus, vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Rct.Close
    rst_rqt.Close
    rst_Ust.Close
    Set frmReportQueueStatus = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdStatus.ColWidth(RQTCODEINDEX) = 0
    grdStatus.ColWidth(DESCRIPTIONINDEX) = 0
    grdStatus.ColWidth(SORTINDEX) = 0
    grdStatus.ColWidth(USERNAMEINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(PRIORITYINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(STATUEINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(REQUESTEDINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(STARTEDINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(COMPLETEDINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(DELETEINDEX) = grdStatus.Width * 0.05
    grdStatus.ColWidth(REPORTNAMEINDEX) = grdStatus.Width - GRIDSCROLLWIDTH - 15
    For ilCol = REPORTNAMEINDEX To DELETEINDEX Step 1
        If ilCol <> REPORTNAMEINDEX Then
            grdStatus.ColWidth(REPORTNAMEINDEX) = grdStatus.ColWidth(REPORTNAMEINDEX) - grdStatus.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdStatus
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdStatus.TextMatrix(0, REPORTNAMEINDEX) = "Report Name"
    grdStatus.TextMatrix(0, USERNAMEINDEX) = "User"
    grdStatus.TextMatrix(0, PRIORITYINDEX) = "Priority"
    grdStatus.TextMatrix(0, STATUEINDEX) = "Status"
    grdStatus.TextMatrix(0, REQUESTEDINDEX) = "Requested"
    grdStatus.TextMatrix(0, STARTEDINDEX) = "Started"
    grdStatus.TextMatrix(0, COMPLETEDINDEX) = "Completed"

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim slStr As String
    
    On Error GoTo ErrHand
    gSetMousePointer grdStatus, grdStatus, vbHourglass
    gGrid_Clear grdStatus, True
    llRow = grdStatus.FixedRows
    grdStatus.Redraw = False
    
    SQLQuery = "Select "
    SQLQuery = SQLQuery & "* "
    SQLQuery = SQLQuery & "From rqt "
    SQLQuery = SQLQuery & "Where "
    SQLQuery = SQLQuery & " rqtDateEntered >= '" & Format(gNow(), sgSQLDateForm) & "'"
    Set rst_rqt = gSQLSelectCall(SQLQuery)
    Do While Not rst_rqt.EOF
        If llRow >= grdStatus.Rows Then
            grdStatus.AddItem ""
        End If
        grdStatus.Row = llRow
    
        grdStatus.TextMatrix(llRow, REPORTNAMEINDEX) = Trim$(rst_rqt!rqtReportName)
        SQLQuery = "SELECT ustname, ustReportName, ustUserInitials FROM Ust Where ustCode = " & rst_rqt!rqtUstCode
        Set rst_Ust = gSQLSelectCall(SQLQuery)
        If Not rst_Ust.EOF Then
            'If Trim$(rst_Ust!ustUserInitials) <> "" Then
            '    grdStatus.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_Ust!ustUserInitials)
            'Else
                If Trim$(rst_Ust!ustReportName) <> "" Then
                    grdStatus.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_Ust!ustReportName)
                Else
                    grdStatus.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_Ust!ustname)
                End If
            'End If
        End If
        If rst_rqt!rqtStatus = "P" Then
            grdStatus.TextMatrix(llRow, STATUEINDEX) = "Processing"
            grdStatus.TextMatrix(llRow, PRIORITYINDEX) = "Running"
            grdStatus.TextMatrix(llRow, REQUESTEDINDEX) = Format(rst_rqt!rqtDateEntered, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeEntered, sgShowTimeWSecForm)
            grdStatus.TextMatrix(llRow, STARTEDINDEX) = Format(rst_rqt!rqtDateStarted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeStarted, sgShowTimeWSecForm)
        ElseIf rst_rqt!rqtStatus = "C" Then
            grdStatus.TextMatrix(llRow, STATUEINDEX) = "Completed"
            grdStatus.TextMatrix(llRow, PRIORITYINDEX) = ""
            grdStatus.TextMatrix(llRow, REQUESTEDINDEX) = Format(rst_rqt!rqtDateEntered, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeEntered, sgShowTimeWSecForm)
            grdStatus.TextMatrix(llRow, STARTEDINDEX) = Format(rst_rqt!rqtDateStarted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeStarted, sgShowTimeWSecForm)
            grdStatus.TextMatrix(llRow, COMPLETEDINDEX) = Format(rst_rqt!rqtDateCompleted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeCompleted, sgShowTimeWSecForm)
        ElseIf rst_rqt!rqtStatus = "E" Then
            grdStatus.TextMatrix(llRow, STATUEINDEX) = "Error"
            grdStatus.TextMatrix(llRow, PRIORITYINDEX) = ""
            grdStatus.TextMatrix(llRow, REQUESTEDINDEX) = Format(rst_rqt!rqtDateEntered, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeEntered, sgShowTimeWSecForm)
            If gDateValue(Format(rst_rqt!rqtDateStarted, sgShowDateForm)) <> gDateValue("1/1/1970") Then
                grdStatus.TextMatrix(llRow, STARTEDINDEX) = Format(rst_rqt!rqtDateStarted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeStarted, sgShowTimeWSecForm)
            End If
            If gDateValue(Format(rst_rqt!rqtDateCompleted, sgShowDateForm)) <> gDateValue("1/1/1970") Then
                grdStatus.TextMatrix(llRow, COMPLETEDINDEX) = Format(rst_rqt!rqtDateCompleted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeCompleted, sgShowTimeWSecForm)
            End If
        Else
            grdStatus.TextMatrix(llRow, STATUEINDEX) = "Requested"
            grdStatus.TextMatrix(llRow, PRIORITYINDEX) = rst_rqt!rqtCode
            grdStatus.TextMatrix(llRow, REQUESTEDINDEX) = Format(rst_rqt!rqtDateEntered, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeEntered, sgShowTimeWSecForm)
        End If
        grdStatus.TextMatrix(llRow, DESCRIPTIONINDEX) = rst_rqt!rqtDescription
        grdStatus.TextMatrix(llRow, RQTCODEINDEX) = rst_rqt!rqtCode
        llRow = llRow + 1
        rst_rqt.MoveNext
    Loop
    mReportSortCol PRIORITYINDEX
    mSetStatusGridColor
    gSetMousePointer grdStatus, grdStatus, vbDefault
    grdStatus.Redraw = True
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "ReportQueueStatusLog.txt", "ReportQueue-mPopulate"
    grdStatus.Redraw = True
    Resume Next
End Sub


Private Sub mSetStatusGridColor()
    Dim llRow As Long
    Dim llCol As Long
    
    'gGrid_Clear grdStatus, True
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        For llCol = REPORTNAMEINDEX To COMPLETEDINDEX Step 1
            grdStatus.Row = llRow
            grdStatus.Col = llCol
            If llCol = PRIORITYINDEX Then
                If grdStatus.TextMatrix(llRow, llCol) = "Requeted" Then
                    grdStatus.CellBackColor = vbWhite
                Else
                    grdStatus.CellBackColor = LIGHTYELLOW
                End If
            ElseIf llCol = DELETEINDEX Then
                If grdStatus.TextMatrix(llRow, llCol) = "Requeted" Then
                    grdStatus.CellBackColor = vbRed
                    grdStatus.CellForeColor = vbWhite
                    grdStatus.TextMatrix(llRow, DELETEINDEX) = "X"
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
        If grdStatus.MouseCol = REPORTNAMEINDEX Then
            grdStatus.ToolTipText = Trim$(grdStatus.TextMatrix(grdStatus.MouseRow, DESCRIPTIONINDEX))
        Else
            grdStatus.ToolTipText = grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol)
        End If
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
    
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdStatus, X, Y)
    If Not ilFound Then
        grdStatus.Redraw = True
        On Error Resume Next
        cmcDone.SetFocus
        Exit Sub
    End If
    
    If Trim$(grdStatus.TextMatrix(grdStatus.Row, REPORTNAMEINDEX)) = "" Then
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
        grdStatus.Col = PRIORITYINDEX
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
        grdStatus.Col = PRIORITYINDEX
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
    If (grdStatus.Row >= grdStatus.FixedRows) And (grdStatus.Row < grdStatus.Rows) And (grdStatus.Col = PRIORITYINDEX) Then
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


Private Sub mAdjustPriority()
    Dim ilMin As Integer
    On Error GoTo ErrHand
    SQLQuery = "SELECT Min(rqtPriority) FROM rqt WHERE (rqtPriority > 0) AND (rqtStatus = 'R' or rqtStatus = 'P')"
    Set rst_rqt = gSQLSelectCall(SQLQuery)
    If Not rst_rqt.EOF Then
        If rst_rqt(0).Value <> vbNull Then
            ilMin = rst_rqt(0).Value
            If ilMin > 1 Then
                ilMin = ilMin - 1
                SQLQuery = "UPDATE rqt SET "
                SQLQuery = SQLQuery & "rqtPriority = rqtPriority - " & ilMin
                SQLQuery = SQLQuery & "WHERE rqtPriority > 1 AND eqtStatus = 'R'"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand1:
                    gHandleError "AffErrorLog.txt", "ReportQueue-mAdjustPriority"
                    Exit Sub
                End If
            End If
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "ReportQueueStatusLog.txt", "ReportQueue-mAdjustPriority"
    Exit Sub
'ErrHand1:
'    gHandleError "ReportQueueStatusLog.txt", "ReportQueue-mAdjustPriority"
'
End Sub


Private Sub mReportSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        slStr = Trim$(grdStatus.TextMatrix(llRow, REPORTNAMEINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdStatus.TextMatrix(llRow, REQUESTEDINDEX))
            ilPos = InStr(1, slStr, " ", vbTextCompare)
            If ilPos > 0 Then
                slDate = Left(slStr, ilPos - 1)
                slTime = Mid(slStr, ilPos + 1)
            Else
                slDate = slStr
                slTime = "11:59:59pm"
            End If
            slDate = gDateValue(slDate)
            Do While Len(slDate) < 6
                slDate = "0" & slDate
            Loop
            slTime = Trim$(Str$(gTimeToLong(slTime, False)))
            Do While Len(slTime) < 6
                slTime = "0" & slTime
            Loop
            If ilCol = PRIORITYINDEX Then
                If grdStatus.TextMatrix(llRow, STATUEINDEX) = "Processing" Then
                    slSort = "A"
                ElseIf grdStatus.TextMatrix(llRow, STATUEINDEX) = "Completed" Then
                    slStr = "C" & slDate & slTime
                Else
                    slSort = (Trim$(grdStatus.TextMatrix(llRow, PRIORITYINDEX)))
                    Do While Len(slSort) < 3
                        slSort = "0" & slSort
                    Loop
                    slSort = "B" & slSort
                End If
            ElseIf ilCol = REQUESTEDINDEX Then
                slSort = slDate & slTime
            ElseIf ilCol = STATUEINDEX Then
                If grdStatus.TextMatrix(llRow, STATUEINDEX) = "Processing" Then
                    slSort = "P"
                ElseIf grdStatus.TextMatrix(llRow, STATUEINDEX) = "Completed" Then
                    slStr = "C" & slDate & slTime
                Else
                    slSort = (Trim$(grdStatus.TextMatrix(llRow, PRIORITYINDEX)))
                    Do While Len(slSort) < 3
                        slSort = "0" & slSort
                    Loop
                    slSort = "R" & slSort
                End If
            Else
                slSort = UCase$(Trim$(grdStatus.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = " "
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
    gGrid_SortByCol grdStatus, REPORTNAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub


