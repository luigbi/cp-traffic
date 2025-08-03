VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmViewBlocks 
   Caption         =   "View Blocks Status"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "AffViewBlocks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   9360
   Begin VB.CommandButton cmcErase 
      Caption         =   "Erase"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5745
      TabIndex        =   6
      Top             =   5115
      Width           =   1245
   End
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
      TabIndex        =   3
      Top             =   4665
      Width           =   60
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
      TabIndex        =   4
      Top             =   5055
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   5115
      Width           =   1245
   End
   Begin VB.CommandButton cmcRefresh 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   4095
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
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   8123
      _Version        =   393216
      Rows            =   4
      Cols            =   7
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
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
End
Attribute VB_Name = "frmViewBlocks"
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

Private lmRowSelected As Long
Private imCtrlKey As Integer

Private rst_Rlf As ADODB.Recordset
Private rst_Ust As ADODB.Recordset
Private rst_Cptt As ADODB.Recordset

Private Const DATEENTEREDINDEX = 0
Private Const TIMEENTEREDINDEX = 1
Private Const TASKINDEX = 2
Private Const USERNAMEINDEX = 3
Private Const INFOMATIONINDEX = 4
Private Const SORTINDEX = 5
Private Const RLFCODEINDEX = 6

Private Sub cmcDone_Click()
    Unload frmViewBlocks
End Sub

Private Sub cmcErase_Click()
    Dim ilRet As Integer
    If lmRowSelected < grdStatus.FixedCols Then
        Exit Sub
    End If
    ilRet = gDeleteLockRec_ByRlfCode(grdStatus.TextMatrix(lmRowSelected, RLFCODEINDEX))
    If ilRet Then
        grdStatus.RemoveItem lmRowSelected
    End If
    cmcErase.Enabled = False
    lmRowSelected = -1
    imLastColSorted = -1
    imLastSort = -1
    mPopulate
End Sub

Private Sub cmcRefresh_Click()
    cmcErase.Enabled = False
    lmRowSelected = -1
    imLastColSorted = -1
    imLastSort = -1
    mPopulate
End Sub

Private Sub Form_Click()
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
    lmRowSelected = -1
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
    rst_Rlf.Close
    rst_Ust.Close
    rst_Cptt.Close
    Set frmViewBlocks = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdStatus.ColWidth(RLFCODEINDEX) = 0
    grdStatus.ColWidth(SORTINDEX) = 0
    grdStatus.ColWidth(TASKINDEX) = 0
    grdStatus.ColWidth(DATEENTEREDINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(TIMEENTEREDINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(USERNAMEINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(INFOMATIONINDEX) = grdStatus.Width - GRIDSCROLLWIDTH - 15
    For ilCol = DATEENTEREDINDEX To INFOMATIONINDEX Step 1
        If ilCol <> INFOMATIONINDEX Then
            grdStatus.ColWidth(INFOMATIONINDEX) = grdStatus.ColWidth(INFOMATIONINDEX) - grdStatus.ColWidth(ilCol)
        End If
    Next ilCol
    grdStatus.ColWidth(TASKINDEX) = grdStatus.ColWidth(INFOMATIONINDEX) / 2
    grdStatus.ColWidth(INFOMATIONINDEX) = grdStatus.ColWidth(INFOMATIONINDEX) - grdStatus.ColWidth(TASKINDEX)
    'Align columns to left
    gGrid_AlignAllColsLeft grdStatus
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdStatus.TextMatrix(0, DATEENTEREDINDEX) = "Date"
    grdStatus.TextMatrix(0, TIMEENTEREDINDEX) = "Time"
    grdStatus.TextMatrix(0, TASKINDEX) = "Task"
    grdStatus.TextMatrix(0, USERNAMEINDEX) = "User"
    grdStatus.TextMatrix(0, INFOMATIONINDEX) = "Information"

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim slStr As String
    Dim ilShtt As Integer
    Dim llVef As Long
    
    On Error GoTo ErrHand
    gSetMousePointer grdStatus, grdStatus, vbHourglass
    gGrid_Clear grdStatus, True
    llRow = grdStatus.FixedRows
    grdStatus.Redraw = False
    
    SQLQuery = "Select "
    SQLQuery = SQLQuery & "* "
    SQLQuery = SQLQuery & "From RLF_Record_Lock "
    SQLQuery = SQLQuery & "Where "
    SQLQuery = SQLQuery & " rlfType = '" & "A" & "'"
    Set rst_Rlf = gSQLSelectCall(SQLQuery)
    Do While Not rst_Rlf.EOF
        If llRow >= grdStatus.Rows Then
            grdStatus.AddItem ""
        End If
        grdStatus.Row = llRow
    
        grdStatus.TextMatrix(llRow, DATEENTEREDINDEX) = Format(rst_Rlf!rlfEnteredDate, sgShowDateForm)
        grdStatus.TextMatrix(llRow, TIMEENTEREDINDEX) = Format(rst_Rlf!rlfEnteredTime, sgShowTimeWSecForm)
        If rst_Rlf!rlfSubType = "G" Then
            grdStatus.TextMatrix(llRow, TASKINDEX) = "Gather Affiliate Spots"
        End If
        SQLQuery = "SELECT ustname, ustReportName, ustUserInitials FROM Ust Where ustCode = " & rst_Rlf!rlfUrfCode
        Set rst_Ust = gSQLSelectCall(SQLQuery)
        If Not rst_Ust.EOF Then
            If Trim$(rst_Ust!ustReportName) <> "" Then
                grdStatus.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_Ust!ustReportName)
            Else
                grdStatus.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_Ust!ustname)
            End If
        End If
        slStr = ""
        If rst_Rlf!rlfSubType = "G" Then
            SQLQuery = "SELECT * FROM cptt Where cpttCode = " & rst_Rlf!rlfRecCode
            Set rst_Cptt = gSQLSelectCall(SQLQuery)
            If Not rst_Cptt.EOF Then
                ilShtt = gBinarySearchStationInfoByCode(rst_Cptt!cpttshfcode)
                If ilShtt <> -1 Then
                    slStr = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                Else
                    slStr = ""
                End If
                llVef = gBinarySearchVef(CLng(rst_Cptt!cpttvefcode))
                If llVef <> -1 Then
                    slStr = Trim$(slStr & " " & tgVehicleInfo(llVef).sVehicle)
                End If
                slStr = slStr & " Week " & Format(rst_Cptt!CpttStartDate, sgShowDateForm)
            End If
        End If
        grdStatus.TextMatrix(llRow, INFOMATIONINDEX) = slStr
        
        grdStatus.TextMatrix(llRow, RLFCODEINDEX) = rst_Rlf!rlfCode
        llRow = llRow + 1
        rst_Rlf.MoveNext
    Loop
    mSortCol TIMEENTEREDINDEX
    mSortCol DATEENTEREDINDEX
    gSetMousePointer grdStatus, grdStatus, vbDefault
    grdStatus.Redraw = True
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "ViewBlocksLog.txt", "ViewBlocks-mPopulate"
    grdStatus.Redraw = True
    Resume Next
End Sub



Private Sub mClearGrid()
    gGrid_Clear grdStatus, True
End Sub

Private Sub grdStatus_Click()
    Dim llRow As Long
    
    If grdStatus.Row >= grdStatus.FixedRows Then
        If grdStatus.TextMatrix(grdStatus.Row, DATEENTEREDINDEX) <> "" Then
            If (lmRowSelected = grdStatus.Row) Then
                If imCtrlKey Then
                    lmRowSelected = -1
                    grdStatus.Row = 0
                    grdStatus.Col = RLFCODEINDEX
                End If
            Else
                lmRowSelected = grdStatus.Row
            End If
        Else
            lmRowSelected = -1
            grdStatus.Row = 0
            grdStatus.Col = RLFCODEINDEX
        End If
    End If
    If lmRowSelected >= grdStatus.FixedRows Then
        cmcErase.Enabled = True
    Else
        cmcErase.Enabled = False
    End If
End Sub

Private Sub grdStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
End Sub

Private Sub grdStatus_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
End Sub

Private Sub grdStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'grdStatus.ToolTipText = ""
    If (grdStatus.MouseRow >= grdStatus.FixedRows) And (grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol)) <> "" Then
        grdStatus.ToolTipText = grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol)
    Else
        grdStatus.ToolTipText = ""
    End If
End Sub




Private Function mColOk() As Integer
    mColOk = True
    If grdStatus.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function




Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        slStr = Trim$(grdStatus.TextMatrix(llRow, DATEENTEREDINDEX))
        If slStr <> "" Then
            If ilCol = DATEENTEREDINDEX Then
                slDate = gDateValue(Trim$(grdStatus.TextMatrix(llRow, DATEENTEREDINDEX)))
                Do While Len(slDate) < 6
                    slDate = "0" & slDate
                Loop
                slSort = slDate
            ElseIf ilCol = TIMEENTEREDINDEX Then
                slTime = gTimeToLong(Trim$(grdStatus.TextMatrix(llRow, TIMEENTEREDINDEX)), False)
                Do While Len(slTime) < 6
                    slTime = "0" & slTime
                Loop
                slSort = slStr
            Else
                slSort = Trim$(grdStatus.TextMatrix(llRow, ilCol))
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
    gGrid_SortByCol grdStatus, DATEENTEREDINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub


