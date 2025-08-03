VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmComplianceTracer 
   Caption         =   "Compliance Tracer"
   ClientHeight    =   6975
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11475
   Icon            =   "AffComplianceTracer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11475
   Begin V81Affiliate.CSI_Calendar txtEndDate 
      Height          =   375
      Left            =   7395
      TabIndex        =   3
      Top             =   240
      Width           =   1635
      _ExtentX        =   2884
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
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   12600
      Top             =   5160
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6975
      FormDesignWidth =   11475
   End
   Begin V81Affiliate.CSI_Calendar txtStartDate 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1635
      _ExtentX        =   2884
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
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13200
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   9390
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.CheckBox chkZeroSpots 
      Caption         =   "Only show weeks with no spots"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCptt 
      Height          =   1905
      Left            =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   8000
      _ExtentX        =   14102
      _ExtentY        =   3360
      _Version        =   393216
      Rows            =   4
      Cols            =   14
      FixedRows       =   2
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   14
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSpots 
      Height          =   2025
      Left            =   480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3240
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   3572
      _Version        =   393216
      Rows            =   4
      Cols            =   22
      FixedRows       =   3
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   22
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblComment 
      Caption         =   "* indicates a replacement spot"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   6480
      Width           =   4215
   End
   Begin VB.Label lblEndDate 
      Caption         =   "End Date"
      Height          =   375
      Left            =   6435
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblStartDate 
      Caption         =   "Start Date"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmComplianceTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************************************************************
'*  frmCompliance Tracer - Created to do research work on CPTT and Spots statuses
'*
'*  Created June,2011 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2011
'************************************************************************************


Option Explicit
Option Compare Text

Private smUserName As String

Private smStartDate As String
Private smEndDate As String
Private smTxtFile As String
Private lmRow As Long
Private lmMaxRows As Long
Private hmToDetail As Integer
Private lmLstCode() As Long

Private mSpotGridHeight As Integer
Private mSpotGridWidth  As Integer
Private mCpttGridHeight As Integer
Private mCpttGridWidth As Integer

'CPTT grid constants
Const CPTTWEEKINDEX = 0
Const VEHICLEINDEX = 1
Const STATIONINDEX = 2
Const ZONEINDEX = 3
Const OFFSETINDEX = 4
Const DATERANGEINDEX = 5
Const CPSTATUSINDEX = 6
Const POSTINGSTATUSINDEX = 7
Const CPTTASTSTATUSINDEX = 8
Const NUMCOMPLIANTINDEX = 9
Const NUMAIREDINDEX = 10
Const NUMGENEDINDEX = 11
Const CPTTATTCODEINDEX = 12
Const CPTTCODEINDEX = 13

'Spots grid constants
Const DATFDDAYSINDEX = 0
Const DATFDTIMESINDEX = 1
Const DATFDSTATUSINDEX = 2
Const DATAIRPLAYNOINDEX = 3
Const DATPDDAYSINDEX = 4
Const DATPDTIMEINDEX = 5
Const DATCODEINDEX = 6
Const ASTCPSTATUSINDEX = 7
Const ASTPDSTATUSINDEX = 8
Const ASTSTATUSINDEX = 9
Const LSTADFNAMEINDEX = 10
Const ASTFDDATEINDEX = 11
Const ASTFDTIMEINDEX = 12
Const ASTLKASTCODEINDEX = 13
Const ASTLKDATEINDEX = 14
Const ASTCODEINDEX = 15
Const lSTLOGDATENDEX = 16
Const LSTLOGTIMEINDEX = 17
Const LSTSTATUSINDEX = 18
Const ISCICODEINDEX = 19
Const LSTCODEINDEX = 20
Const SDFCODEINDEX = 21
Const CPTTTOPROW = 2
Const SPOTSTOPROW = 3



Private Sub cmdCancel_Click()
    
    Unload frmComplianceTracer
    
End Sub

Private Sub cmdSearch_Click()
    
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim llCpttCode As Long

    grdCptt.Height = mCpttGridHeight
    ilRow = ilRow
    grdCptt.ScrollBars = flexScrollBarVertical
    grdSpots.Visible = False
    grdCptt.TopRow = CPTTTOPROW
    grdCptt.Redraw = True
    grdCptt.Clear
    grdSpots.Clear
    ilRet = mFillCpttTitles
    ilRet = mFillCpttGrid
    grdCptt.Redraw = True
End Sub

Private Sub Form_Load()

    smUserName = sgUserName
    grdCptt.Visible = False
    grdSpots.Visible = False
    cmdExport.Visible = False
    Call mFillCpttTitles
    
End Sub

Private Sub Form_Initialize()

    'D.S. 06/4/11

    Dim ilFoundRes As Integer
    
    Me.Width = Screen.Width / 1.02
    Me.Height = Screen.Height / 1.25
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 1.2
    
    ilFoundRes = False
'2/21/15: Generalized the sizing of the Grids
'    '800 X 600 resolution
'    If Screen.Width = 12000 And Screen.Height = 9000 Then
'        ilFoundRes = True
'        mSpotGridHeight = 4300
'        mSpotGridWidth = 11000
'        mCpttGridHeight = 4900
'        mCpttGridWidth = 11000
'    End If
'
'    '1024 X 768 resolution
'    If Screen.Width = 15360 And Screen.Height = 11520 Then
'        ilFoundRes = True
'        mSpotGridHeight = 4300 * 1.25
'        mSpotGridWidth = 11000 * 1.25
'        mCpttGridHeight = 4900 * 1.25
'        mCpttGridWidth = 11000 * 1.25
'    End If
'
'    '1280 X 960 resolution
'    If Screen.Width = 19200 And Screen.Height = 14400 Then
'        ilFoundRes = True
'        mSpotGridHeight = 4300 * 1.65
'        mSpotGridWidth = 11000 * 1.6
'        mCpttGridHeight = 4900 * 1.65
'        mCpttGridWidth = 11000 * 1.5
'    End If
'
'    '1280 X 1024 resolution
'    If Screen.Width = 19200 And Screen.Height = 15360 Then
'        ilFoundRes = True
'        mSpotGridHeight = 4300 * 1.8
'        mSpotGridWidth = 11000 * 1.6
'        mCpttGridHeight = 4900 * 1.8
'        mCpttGridWidth = 11000 * 1.6
'    End If
    
    If Not ilFoundRes Then
        ''We covered all the smaller resolution sizes and we didn't find the correct
        ''resolution so set it to our max res. 1280 X 1024
        'mSpotGridHeight = 4300 * 1.8
        'mSpotGridWidth = 11000 * 1.6
        mCpttGridHeight = (cmdCancel.Top - (txtStartDate.Top + txtStartDate.Height)) * 0.9 '4900 * 1.8
        mCpttGridWidth = Me.Width * 0.9 '11000 * 1.6
        mSpotGridHeight = mCpttGridHeight - (grdCptt.RowHeight(0) * 3)
        mSpotGridWidth = mCpttGridWidth
    End If
    
    gSetFonts frmComplianceTracer
    gCenterForm frmComplianceTracer
    
End Sub

Private Function mFillCpttGrid() As Boolean

    'D.S. 06/4/11
    
    Dim rst As ADODB.Recordset
    Dim llRow As Long
    Dim slTemp As String
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    Dim ilNumberRowsToAdd As Integer
    
    On Error GoTo ErrHand
    
    mFillCpttGrid = False
    
    'debug pre-populate the date field
    'txtStartDate.Text = "4/19/10"
    'txtEndDate.Text = "4/25/10"

    Call gPopShttInfo
    Call gPopVehicles
    
    smStartDate = Trim$(txtStartDate.Text)
    If Not gIsDate(smStartDate) Then
        MsgBox "Please Enter a Valid Start Date"
        txtStartDate.SetFocus
        Exit Function
    End If
    
    smStartDate = Format(smStartDate, sgSQLDateForm)
        
    smEndDate = Trim$(txtEndDate.Text)
    If UCase(smEndDate) = "TFN" Then
        smEndDate = "2069-12-31"
    End If
    
    If smEndDate = "" Then
        smEndDate = "2069-12-31"
        txtEndDate.Text = "TFN"
    Else
        If Not gIsDate(smEndDate) Then
            MsgBox "Please Enter a Valid End Date"
            txtEndDate.SetFocus
            Exit Function
        Else
            smEndDate = Format(smEndDate, sgSQLDateForm)
        End If
    End If
    
    '2/21/15: Replaced zero with vbUnchecked
    'If chkZeroSpots.Value = 0 Then
    If chkZeroSpots.Value = vbUnchecked Then
        'All weeks
        SQLQuery = "Select * from Cptt, Shtt, ATT, VEF_Vehicles where cpttStartDate >= " & "'" & smStartDate & "'" & " And cpttStartDate <= " & "'" & smEndDate & "'" & " And ShttCode = cpttShfCode AND AttCode = cpttAtfCode AND vefCode = cpttVefCode"
        Set rst = gSQLSelectCall(SQLQuery)
    Else
        'weeks with No spots generated
        SQLQuery = "Select * from Cptt, Shtt, ATT, VEF_Vehicles where cpttStartDate >= " & "'" & smStartDate & "'" & " And cpttStartDate <= " & "'" & smEndDate & "'" & " AND cpttNoSpotsGen = 0" & " And ShttCode = cpttShfCode AND AttCode = cpttAtfCode AND vefCode = cpttVefCode"
        Set rst = gSQLSelectCall(SQLQuery)
    End If
    
    llRow = grdCptt.FixedRows
    gSetMousePointer grdCptt, grdSpots, vbHourglass
    grdCptt.Visible = False
    While Not rst.EOF
        If llRow >= grdCptt.Rows Then
            grdCptt.AddItem ""
        End If
        grdCptt.TextMatrix(llRow, CPTTWEEKINDEX) = Trim$(rst!CpttStartDate)
        grdCptt.TextMatrix(llRow, VEHICLEINDEX) = Trim$(rst!vefName)
        grdCptt.TextMatrix(llRow, STATIONINDEX) = Trim$(rst!shttCallLetters)
        grdCptt.TextMatrix(llRow, ZONEINDEX) = Trim$(rst!shttTimeZone)
        slTemp = gGetTimeZoneOffset(rst!shttCode, rst!vefCode)
        grdCptt.TextMatrix(llRow, OFFSETINDEX) = slTemp
        slTemp = gGetAgreementDateRange(rst!attCode)
        grdCptt.TextMatrix(llRow, DATERANGEINDEX) = Trim$(slTemp)
        
        Select Case Trim$(rst!cpttStatus)
        Case 0
            grdCptt.TextMatrix(llRow, CPSTATUSINDEX) = "Not Posted or Partial"
        Case 1
            grdCptt.TextMatrix(llRow, CPSTATUSINDEX) = "Posting Complete"
        Case 2
            grdCptt.TextMatrix(llRow, CPSTATUSINDEX) = "Posting Complete as None Aired"
        Case Else
            grdCptt.TextMatrix(llRow, CPSTATUSINDEX) = "Unknown"
        End Select
        
        Select Case Trim$(rst!cpttPostingStatus)
        Case 0
            grdCptt.TextMatrix(llRow, POSTINGSTATUSINDEX) = "Not Posted"
        Case 1
            grdCptt.TextMatrix(llRow, POSTINGSTATUSINDEX) = "Partial"
        Case 2
            grdCptt.TextMatrix(llRow, POSTINGSTATUSINDEX) = "Complete"
        Case Else
            grdCptt.TextMatrix(llRow, POSTINGSTATUSINDEX) = "Unknown"
        End Select
        
        Select Case Trim$(rst!cpttAstStatus)
        Case "N"
            grdCptt.TextMatrix(llRow, CPTTASTSTATUSINDEX) = "Never"
        Case "R"
            grdCptt.TextMatrix(llRow, CPTTASTSTATUSINDEX) = "Recreated"
        Case "C"
            grdCptt.TextMatrix(llRow, CPTTASTSTATUSINDEX) = "Created"
        Case Else
            grdCptt.TextMatrix(llRow, CPTTASTSTATUSINDEX) = "Unknown"
        End Select
        
        grdCptt.TextMatrix(llRow, NUMCOMPLIANTINDEX) = Trim$(rst!cpttNoCompliant)
        grdCptt.TextMatrix(llRow, NUMAIREDINDEX) = Trim$(rst!cpttNoSpotsAired)
        grdCptt.TextMatrix(llRow, NUMGENEDINDEX) = Trim$(rst!cpttNoSpotsGen)
        'If grdCptt.ColWidth(CPTTATTCODEINDEX) <> 0 Then
            grdCptt.TextMatrix(llRow, CPTTATTCODEINDEX) = Trim$(rst!cpttatfCode)
        'End If
        'If grdCptt.ColWidth(CPTTCODEINDEX) <> 0 Then
            grdCptt.TextMatrix(llRow, CPTTCODEINDEX) = Trim$(rst!cpttCode)
        'End If
        
        rst.MoveNext
        llRow = llRow + 1
    Wend
    
    rst.Close
    ilNumberRowsToAdd = mCpttGridHeight / grdCptt.RowHeight(0)
    For llRow = 1 To ilNumberRowsToAdd Step 1
        grdCptt.AddItem ""
    Next llRow
    grdCptt.Height = mCpttGridHeight
    mResizeCPTTColumns
    grdCptt.Visible = True
    gSetMousePointer grdCptt, grdSpots, vbDefault
    gGrid_AlignAllColsLeft grdCptt
    mFillCpttGrid = True
    
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmComplianceTracer - mFillCpttGrid"
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mHideSpotGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gSetMousePointer grdCptt, grdSpots, vbDefault
    Erase lmLstCode
    Set frmComplianceTracer = Nothing
    
End Sub

Private Sub grdCptt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'grdCptt.ToolTipText = ""
    If (grdCptt.MouseRow >= grdCptt.FixedRows) And (grdCptt.MouseCol > 0) And (grdCptt.TextMatrix(grdCptt.MouseRow, grdCptt.MouseCol)) <> "" Then
        grdCptt.ToolTipText = grdCptt.TextMatrix(grdCptt.MouseRow, grdCptt.MouseCol)
    Else
        grdCptt.ToolTipText = ""
    End If

End Sub

Private Sub grdCptt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'D.S. 06/4/11
    
    Dim ilRet As Integer
    Dim llCpttCode As Long
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    On Error GoTo ErrHand
    
    'Detect if the area clicked it is valid or not
    ilCol = grdCptt.MouseCol
    ilRow = grdCptt.MouseRow
    
    If ilCol < grdCptt.FixedCols Then
        grdCptt.Redraw = True
        Exit Sub
    End If
    If ilRow < grdCptt.FixedRows Then
        grdCptt.Redraw = True
        Exit Sub
    End If
    
    lmRow = grdCptt.Row
    DoEvents
    
    'Get CPTT from the row that the user clicked in
    llCpttCode = grdCptt.TextMatrix(lmRow, CPTTCODEINDEX)
    
    'Set the top row to the one the user clicked in
    If grdCptt.Height = mCpttGridHeight Then
        grdCptt.TopRow = lmRow
        grdCptt.LeftCol = 1
        'Shrink the height of the Cptt grid
        grdCptt.ScrollBars = flexScrollBarNone
        grdCptt.ColWidth(VEHICLEINDEX) = grdCptt.ColWidth(VEHICLEINDEX) + 275
        grdCptt.Height = grdCptt.RowHeight(0) * 3
        grdCptt.Redraw = True
        ilRet = mFillSpotTitles
        ilRet = mFillSpotGrid(llCpttCode)
    Else
'        grdCptt.Width = mCpttGridWidth
'        grdCptt.ColWidth(VEHICLEINDEX) = grdCptt.ColWidth(VEHICLEINDEX) - 275
'        grdCptt.Height = mCpttGridHeight
'        lmRow = lmRow
'        grdCptt.ScrollBars = flexScrollBarVertical
'        grdSpots.Visible = False
'        cmdExport.Visible = False
'        grdCptt.TopRow = CPTTTOPROW
'        grdCptt.LeftCol = 1
'
'        grdCptt.Redraw = True
        mHideSpotGrid
    End If
    
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmCompliance Tracer - grdCptt_MouseUp: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
    gLogMsg " ", "AffErrorLog.txt", False
    
End Sub

Private Function mFillSpotGrid(lCpttCode As Long) As Boolean

    'D.S. 06/4/11

    Dim llRow As Long
    Dim llAttCode As Long
    Dim llVefCode As Long
    Dim rst_Ast As ADODB.Recordset
    Dim rst_Cptt As ADODB.Recordset
    Dim rst_DAT As ADODB.Recordset
    Dim rst_Lst As ADODB.Recordset
    Dim rst_Temp As ADODB.Recordset
    Dim rst_AstLk As ADODB.Recordset
    Dim slDatFdDays As String
    Dim slDatPdDays As String
    Dim slTemp As String
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    Dim ilAstFound As Integer
    Dim ilRet As Integer
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    
    mFillSpotGrid = False
    
    On Error GoTo ErrHand
    
    ilAstFound = False
    grdSpots.Visible = False
    gSetMousePointer grdCptt, grdSpots, vbHourglass
    
    SQLQuery = "Select cpttAtfCode, cpttVefCode from Cptt where cpttCode = " & lCpttCode
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    If Not rst_Cptt.EOF Then
        llAttCode = rst_Cptt!cpttatfCode
        llVefCode = rst_Cptt!cpttvefcode
        llRow = grdSpots.FixedRows
        
        'Is there a DAT record associated with the above agreement
        SQLQuery = "Select * from Dat where datAtfCode = " & llAttCode
        Set rst_DAT = gSQLSelectCall(SQLQuery)
        ReDim lmLstCode(0 To 0)
        If Not rst_DAT.EOF Then
            While Not rst_DAT.EOF
                slDatFdDays = mGetDatFdDays(rst_DAT)
                gSetMousePointer grdCptt, grdSpots, vbHourglass
                grdSpots.Visible = False
                
                If llRow >= grdSpots.Rows Then
                    grdSpots.AddItem ""
                End If
                grdSpots.TextMatrix(llRow, DATFDDAYSINDEX) = Trim$(slDatFdDays)
                
                slTemp = Format(rst_DAT!datFdStTime, "hh:mmA/P") & "-" & Format(rst_DAT!datFdEdTime, "hh:mmA/P")
                grdSpots.TextMatrix(llRow, DATFDTIMESINDEX) = Trim$(slTemp)
                
                Select Case UCase(rst_DAT!datFdStatus)
                Case 0
                    slTemp = "Carried"
                Case 1
                    slTemp = "Delay"
                Case 2, 3, 4, 5
                    slTemp = "Not Carried"
                Case 7
                    slTemp = "Special"
                Case 8
                    slTemp = "Off Air"
                Case 9
                    slTemp = "Delay Cmml/Prg"
                Case 10
                    slTemp = "Air Cmml Only"
                End Select
        
                grdSpots.TextMatrix(llRow, DATFDSTATUSINDEX) = Trim$(slTemp)
                grdSpots.TextMatrix(llRow, DATAIRPLAYNOINDEX) = Trim$(rst_DAT!datAirPlayNo)
                
                slDatPdDays = mGetDatPdDays(rst_DAT)
                
                grdSpots.TextMatrix(llRow, DATPDDAYSINDEX) = Trim$(slDatPdDays)
                
                slTemp = Format(rst_DAT!datPdStTime, "hh:mmA/P") & "-" & Format(rst_DAT!datPdEdTime, "hh:mmA/P")
                grdSpots.TextMatrix(llRow, DATPDTIMEINDEX) = Trim$(slTemp)
                grdSpots.TextMatrix(llRow, DATCODEINDEX) = Trim$(rst_DAT!datCode)
                'If grdSpots.ColWidth(DATCODEINDEX) <> 0 Then
                    grdSpots.TextMatrix(llRow, DATCODEINDEX) = Trim$(rst_DAT!datCode)
                'End If
    
                'Is there any AST records associated with the above agreement
                'SQLQuery = "Select * from AST where astatfcode = " & llAttCode & " and astFeedDate >= " & "'" & smStartDate & "'" & " AND astFeedDate <= " & "'" & smEndDate & "'" & " AND astfeedtime = " & "'" & Format(rst_DAT!datFdStTime, sgSQLTimeForm) & "'" & " order by astCode"
                SQLQuery = "Select * from AST where astatfcode = " & llAttCode & " and astFeedDate >= " & "'" & smStartDate & "'" & " AND astFeedDate <= " & "'" & smEndDate & "'" & " AND astDatCode = " & rst_DAT!datCode & " order by astCode"
                Set rst_Ast = gSQLSelectCall(SQLQuery)
                
                If Not rst_Ast.EOF Then
                    While Not rst_Ast.EOF
                        
                        '12/13/13: Obtain Pledge information from Dat
                        tlDatPledgeInfo.lAttCode = rst_Ast!astAtfCode
                        tlDatPledgeInfo.lDatCode = rst_Ast!astDatCode
                        tlDatPledgeInfo.iVefCode = rst_Ast!astVefCode
                        tlDatPledgeInfo.sFeedDate = Format(rst_Ast!astFeedDate, "m/d/yy")
                        tlDatPledgeInfo.sFeedTime = Format(rst_Ast!astFeedTime, "hh:mm:ssam/pm")
                        ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
                        
                        Select Case Trim$(rst_Ast!astCPStatus)
                        Case 0
                            grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "Not Received"
                        Case 1
                            grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "Received"
                        Case 2
                            grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "Not Aired"
                        Case Else
                            grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "Unknown"
                        End Select
                                                
                        '12/13/13: Obtain Pledge information from Dat
                        'Select Case gGetAirStatus(Trim$(rst_Ast!astPledgeStatus))
                        Select Case gGetAirStatus(Trim$(tlDatPledgeInfo.iPledgeStatus))
                        Case 0
                            grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Carry"
                        Case 1
                            grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Delay"
                        Case 2, 3, 4, 5
                            grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Not Carried"
                        Case 7
                            grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Special"
                        Case 8
                            grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Off Air"
                        Case 9
                            grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Delay Cmml/Prg"
                        Case 10
                            grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Air Cmml Only"
                        Case Else
                            grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Unknown"
                        End Select
                        
                        Select Case gGetAirStatus(Trim$(rst_Ast!astStatus))
                        Case 0
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Carry"
                        Case 1
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Delay"
                        Case 2, 3, 4, 5
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Not Carried"
                        Case 6
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Aired Outside Pledge"
                        Case 7
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Special"
                        Case 8
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Off Air"
                        Case 9
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Delay Cmml/Prg"
                        Case 10
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Air Cmml Only"
                        Case ASTEXTENDED_MG
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "MG"
                        Case ASTEXTENDED_REPLACEMENT
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Replacement"
                        Case ASTEXTENDED_BONUS
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Bonus"
                        Case ASTAIR_MISSED_MG_BYPASS
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Missed MG Bypassed"
                        Case Else
                            grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Unknown"
                        End Select
                        
                        SQLQuery = "SELECT adfName, lstBkoutLstCode"
                        SQLQuery = SQLQuery & " From Lst, ADF_Advertisers"
                        SQLQuery = SQLQuery & " WHERE lstCode = " & rst_Ast!astLsfCode
                        SQLQuery = SQLQuery & " AND adfCode = lstAdfCode"
                        Set rst_Temp = gSQLSelectCall(SQLQuery)
        
                        If Not rst_Temp.EOF Then
                            If rst_Temp!lstBkoutLstCode > 0 Then
                                'It's a replacement spot so add an asterick designator
                                grdSpots.TextMatrix(llRow, LSTADFNAMEINDEX) = "*" & Trim$(rst_Temp!adfName)
                            Else
                                grdSpots.TextMatrix(llRow, LSTADFNAMEINDEX) = Trim$(rst_Temp!adfName)
                            End If
                        Else
                            grdSpots.TextMatrix(llRow, LSTADFNAMEINDEX) = "No Adv."
                        End If
        
                        grdSpots.TextMatrix(llRow, ASTFDDATEINDEX) = Trim$(rst_Ast!astFeedDate)
                        grdSpots.TextMatrix(llRow, ASTFDTIMEINDEX) = Format(Trim$(rst_Ast!astFeedTime), "hh:mmA/P")
                        
                        If rst_Ast!astLkAstCode > 0 Then
                            grdSpots.TextMatrix(llRow, ASTLKASTCODEINDEX) = rst_Ast!astLkAstCode
                            SQLQuery = "Select astFeedDate from AST where astcode = " & rst_Ast!astLkAstCode
                            Set rst_AstLk = gSQLSelectCall(SQLQuery)
                            If Not rst_AstLk.EOF Then
                                grdSpots.TextMatrix(llRow, ASTLKDATEINDEX) = Trim$(rst_AstLk!astFeedDate)
                            Else
                                grdSpots.TextMatrix(llRow, ASTLKDATEINDEX) = "Lk Ast Missing"
                            End If
                        Else
                            grdSpots.TextMatrix(llRow, ASTLKDATEINDEX) = ""
                        End If
                        
                        'If grdSpots.ColWidth(ASTCODEINDEX) <> 0 Then
                            grdSpots.TextMatrix(llRow, ASTCODEINDEX) = Trim$(rst_Ast!astCode)
                        'End If
                        
                        
                        'Get the LST records
                        SQLQuery = "SELECT lstLogDate, lstLogTime, lstStatus, lstCode, lstSdfCode, lstISCI From LST Where lstCode = " & rst_Ast!astLsfCode
                        Set rst_Lst = gSQLSelectCall(SQLQuery)
                        If Not rst_Lst.EOF Then
                            grdSpots.TextMatrix(llRow, lSTLOGDATENDEX) = Trim$(rst_Lst!lstLogDate)
                            grdSpots.TextMatrix(llRow, LSTLOGTIMEINDEX) = Format(Trim$(rst_Lst!lstLogTime), "hh:mmA/P")
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = Trim$(rst_Lst!lstStatus)
                            Select Case Trim$(rst_Lst!lstStatus)
                            Case 0
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Carried"
                            Case 1
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Delay"
                            Case 2, 3, 4, 5
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Not Carried"
                            Case 7
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Special"
                            Case 8
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Off Air"
                            Case 9
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Delay Cmml/Prg"
                            Case 10
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Air Cmml Only"
                            Case ASTEXTENDED_MG
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "MG"
                            Case ASTEXTENDED_REPLACEMENT
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Replacement"
                            Case ASTEXTENDED_BONUS
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Bonus"
                            Case Else
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Unknown"
                            End Select
                            
                            'If grdSpots.ColWidth(LSTCODEINDEX) <> 0 Then
                                grdSpots.TextMatrix(llRow, LSTCODEINDEX) = Trim$(rst_Lst!lstCode)
                                lmLstCode(UBound(lmLstCode)) = CLng(rst_Lst!lstCode)
                                ReDim Preserve lmLstCode(0 To UBound(lmLstCode) + 1)
                            'End If
                            
                            'If grdSpots.ColWidth(ISCICODEINDEX) <> 0 Then
                                grdSpots.TextMatrix(llRow, ISCICODEINDEX) = Trim$(rst_Lst!lstISCI)
                            'End If
                            
                            'If grdSpots.ColWidth(SDFCODEINDEX) <> 0 Then
                                grdSpots.TextMatrix(llRow, SDFCODEINDEX) = Trim$(rst_Lst!lstSdfCode)
                            'End If
                        Else
                            'No LST Found
                            grdSpots.TextMatrix(llRow, lSTLOGDATENDEX) = "No Post Log Spots"
                        End If
                            
                        If Not rst_Ast.EOF Then
                            llRow = llRow + 1
                            If llRow >= grdSpots.Rows Then
                                grdSpots.AddItem ""
                            End If
                        End If
                        
                        rst_Ast.MoveNext
                    Wend
                    
                    If UBound(lmLstCode) > 1 Then
                        ArraySortTyp fnAV(lmLstCode(), 0), UBound(lmLstCode) - 1, 0, LenB(lmLstCode(1)), 0, -2, 0
                    End If

                Else
                    'No AST found
                    grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "No Station Spots"
                    'Get LST records
                    SQLQuery = "SELECT lstLogDate, lstLogTime, lstStatus, lstCode, lstSdfCode, lstISCI From LST"
                    SQLQuery = SQLQuery & " WHERE (lstLogVefCode = " & llVefCode & " AND lstBkoutLstCode = 0" 'AND lstStatus < 20"
                    ''3/9/16: Fix the filter
                    'SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
                    SQLQuery = SQLQuery & " AND (lstLogDate >= " & "'" & smStartDate & "'"
                    SQLQuery = SQLQuery & " AND  lstLogDate <= " & "'" & smEndDate & "'))"
                    SQLQuery = SQLQuery & "ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
                    Set rst_Lst = gSQLSelectCall(SQLQuery)
                    
                    If Not rst_Lst.EOF Then
                        While Not rst_Lst.EOF
                            If mBinarySearchLst(rst_Lst!lstCode) = -1 Then
                                grdSpots.TextMatrix(llRow, lSTLOGDATENDEX) = Trim$(rst_Lst!lstLogDate)
                                grdSpots.TextMatrix(llRow, LSTLOGTIMEINDEX) = Format(Trim$(rst_Lst!lstLogTime), "hh:mmA/P")
                                grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = Trim$(rst_Lst!lstStatus)
                                Select Case Trim$(rst_Lst!lstStatus)
                                Case 0
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Carried"
                                Case 1
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Delay"
                                Case 2, 3, 4, 5
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Not Carried"
                                Case 7
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Special"
                                Case 8
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Off Air"
                                Case 9
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Delay Cmml/Prg"
                                Case 10
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Air Cmml Only"
                                Case ASTEXTENDED_MG
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "MG"
                                Case ASTEXTENDED_REPLACEMENT
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Replacement"
                                Case ASTEXTENDED_BONUS
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Bonus"
                                Case Else
                                    grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Unknown"
                                End Select
                                
                                'If grdSpots.ColWidth(LSTCODEINDEX) <> 0 Then
                                    grdSpots.TextMatrix(llRow, LSTCODEINDEX) = Trim$(rst_Lst!lstCode)
                                'End If
                                
                                'If grdSpots.ColWidth(ISCICODEINDEX) <> 0 Then
                                    grdSpots.TextMatrix(llRow, ISCICODEINDEX) = Trim$(rst_Lst!lstISCI)
                                'End If
                                
                                'If grdSpots.ColWidth(SDFCODEINDEX) <> 0 Then
                                    grdSpots.TextMatrix(llRow, SDFCODEINDEX) = Trim$(rst_Lst!lstSdfCode)
                                'End If
                                
                                If Not rst_Lst.EOF Then
                                    llRow = llRow + 1
                                    If llRow >= grdSpots.Rows Then
                                        grdSpots.AddItem ""
                                    End If
                                End If
                            End If
                            rst_Lst.MoveNext
                        Wend
                    Else
                        grdSpots.TextMatrix(llRow, lSTLOGDATENDEX) = "No Post Log Spots"
                    End If
                End If
                
                If llRow >= grdSpots.Rows Then
                    grdSpots.AddItem ""
                End If
                
                rst_DAT.MoveNext
            Wend
        Else
            'NO DAT Found
             grdSpots.TextMatrix(llRow, DATFDDAYSINDEX) = "No DAT"
            
            'Is there any AST records associated with the above agreement
            'SQLQuery = "Select * from AST where astatfcode = " & llAttCode & " and astFeedDate >= " & "'" & smStartDate & "'" & " AND astFeedDate <= " & "'" & smEndDate & "'" & " AND astfeedtime = " & "'" & Format(rst_DAT!datFdStTime, sgSQLTimeForm) & "'" & " order by astCode"
            SQLQuery = "Select * from AST where astatfcode = " & llAttCode & " and astFeedDate >= " & "'" & smStartDate & "'" & " AND astFeedDate <= " & "'" & smEndDate & "'" & " order by astCode"
            Set rst_Ast = gSQLSelectCall(SQLQuery)
            If Not rst_Ast.EOF Then
                ilAstFound = True
                While Not rst_Ast.EOF

                    '12/13/13: Obtain Pledge information from Dat
                    tlDatPledgeInfo.lAttCode = rst_Ast!astAtfCode
                    tlDatPledgeInfo.lDatCode = rst_Ast!astDatCode
                    tlDatPledgeInfo.iVefCode = rst_Ast!astVefCode
                    tlDatPledgeInfo.sFeedDate = Format(rst_Ast!astFeedDate, "m/d/yy")
                    tlDatPledgeInfo.sFeedTime = Format(rst_Ast!astFeedTime, "hh:mm:ssam/pm")
                    ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)

                    Select Case Trim$(rst_Ast!astCPStatus)
                    Case 0
                        grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "Not Received"
                    Case 1
                        grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "Received"
                    Case 2
                        grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "Not Aired"
                    Case Else
                        grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "Unknown"
                    End Select
                    
                    '12/13/13: Obtain Pledge information from Dat
                    'Select Case gGetAirStatus(Trim$(rst_Ast!astPledgeStatus))
                    Select Case gGetAirStatus(Trim$(tlDatPledgeInfo.iPledgeStatus))
                    Case 0
                        grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Live"
                    Case 1
                        grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Delay"
                    Case 2, 3, 4, 5
                        grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Not Carried"
                    Case 7
                        grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Special"
                    Case 8
                        grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Off Air"
                    Case 9
                        grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Delay Cmml/Prg"
                    Case 10
                        grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Air Cmml Only"
                    Case Else
                        grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) = "Unknown"
                    End Select
                    
                    Select Case gGetAirStatus(Trim$(rst_Ast!astStatus))
                    Case 0
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Carry"
                    Case 1
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Delay"
                    Case 2, 3, 4, 5
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Not Carried"
                    Case 6
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Aired Outside Pledge"
                    Case 7
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Special"
                    Case 8
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Off Air"
                    Case 9
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Delay Cmml/Prg"
                    Case 10
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Air Cmml Only"
                    Case ASTEXTENDED_MG
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "MG"
                    Case ASTEXTENDED_REPLACEMENT
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Replacement"
                    Case ASTEXTENDED_BONUS
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Bonus"
                    Case ASTAIR_MISSED_MG_BYPASS
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Missed MG Bypassed"
                    Case Else
                        grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) = "Unknown"
                    End Select
                
                    SQLQuery = "SELECT adfName, lstBkoutLstCode"
                    SQLQuery = SQLQuery & " From Lst, ADF_Advertisers"
                    SQLQuery = SQLQuery & " WHERE lstCode = " & rst_Ast!astLsfCode
                    SQLQuery = SQLQuery & " AND adfCode = lstAdfCode"
                    Set rst_Temp = gSQLSelectCall(SQLQuery)
    
                    If Not rst_Temp.EOF Then
                        If rst_Temp!lstBkoutLstCode > 0 Then
                            'It's a replacement spot so add an asterick designator
                            grdSpots.TextMatrix(llRow, LSTADFNAMEINDEX) = "*" & Trim$(rst_Temp!adfName)
                        Else
                            grdSpots.TextMatrix(llRow, LSTADFNAMEINDEX) = Trim$(rst_Temp!adfName)
                        End If
                    Else
                        grdSpots.TextMatrix(llRow, LSTADFNAMEINDEX) = "No Adv."
                    End If
    
                    grdSpots.TextMatrix(llRow, ASTFDDATEINDEX) = Trim$(rst_Ast!astFeedDate)
                    grdSpots.TextMatrix(llRow, ASTFDTIMEINDEX) = Format(Trim$(rst_Ast!astFeedTime), "hh:mmA/P")
                    
                    If rst_Ast!astLkAstCode > 0 Then
                        grdSpots.TextMatrix(llRow, ASTLKASTCODEINDEX) = rst_Ast!astLkAstCode
                        SQLQuery = "Select astFeedDate from AST where astcode = " & rst_Ast!astLkAstCode
                        Set rst_AstLk = gSQLSelectCall(SQLQuery)
                        If Not rst_AstLk.EOF Then
                            grdSpots.TextMatrix(llRow, ASTLKDATEINDEX) = Trim$(rst_AstLk!astFeedDate)
                        Else
                            grdSpots.TextMatrix(llRow, ASTLKDATEINDEX) = "Lk Ast Missing"
                        End If
                    Else
                        grdSpots.TextMatrix(llRow, ASTLKDATEINDEX) = ""
                    End If
                    
                    'If grdSpots.ColWidth(ASTCODEINDEX) <> 0 Then
                        grdSpots.TextMatrix(llRow, ASTCODEINDEX) = Trim$(rst_Ast!astCode)
                    'End If
                    
                    
                    'Get the LST records
                    SQLQuery = "SELECT lstLogDate, lstLogTime, lstStatus, lstCode, lstSdfCode From LST Where lstCode = " & rst_Ast!astLsfCode
                    Set rst_Lst = gSQLSelectCall(SQLQuery)
                    If Not rst_Lst.EOF Then
                        grdSpots.TextMatrix(llRow, lSTLOGDATENDEX) = Trim$(rst_Lst!lstLogDate)
                        grdSpots.TextMatrix(llRow, LSTLOGTIMEINDEX) = Format(Trim$(rst_Lst!lstLogTime), "hh:mmA/P")
                        
                        grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = Trim$(rst_Lst!lstStatus)
                        Select Case Trim$(rst_Lst!lstStatus)
                        Case 0
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Carried"
                        Case 1
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Delay"
                        Case 2, 3, 4, 5
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Not Carried"
                        Case 7
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Special"
                        Case 8
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Off Air"
                        Case 9
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Delay Cmml/Prg"
                        Case 10
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Air Cmml Only"
                        Case ASTEXTENDED_MG
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "MG"
                        Case ASTEXTENDED_REPLACEMENT
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Replacement"
                        Case ASTEXTENDED_BONUS
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Bonus"
                        Case Else
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Unknown"
                        End Select
                        
                        'If grdSpots.ColWidth(LSTCODEINDEX) <> 0 Then
                            grdSpots.TextMatrix(llRow, LSTCODEINDEX) = Trim$(rst_Lst!lstCode)
                        'End If
                        
                        'If grdSpots.ColWidth(ISCICODEINDEX) <> 0 Then
                            grdSpots.TextMatrix(llRow, ISCICODEINDEX) = Trim$(rst_Lst!lstISCI)
                        'End If
                        
                        'If grdSpots.ColWidth(SDFCODEINDEX) <> 0 Then
                            grdSpots.TextMatrix(llRow, SDFCODEINDEX) = Trim$(rst_Lst!lstSdfCode)
                        'End If
                    Else
                        'No LST Found
                        grdSpots.TextMatrix(llRow, lSTLOGDATENDEX) = "No Post Log Spots"
                    End If
                        
                    If Not rst_Ast.EOF Then
                        llRow = llRow + 1
                        If llRow >= grdSpots.Rows Then
                            grdSpots.AddItem ""
                        End If
                    End If
                    
                    rst_Ast.MoveNext
                Wend
            Else
                'No AST found - Get LST records
                grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) = "No Station Spots"
                SQLQuery = "SELECT lstLogDate, lstLogTime, lstStatus, lstCode, lstSdfCode From LST"
                SQLQuery = SQLQuery & " WHERE (lstLogVefCode = " & llVefCode & " AND lstBkoutLstCode = 0" 'AND lstStatus < 20"
                ''3/9/16: Fix the filter
                'SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
                SQLQuery = SQLQuery & " AND (lstLogDate >= " & "'" & smStartDate & "'"
                SQLQuery = SQLQuery & " AND  lstLogDate <= " & "'" & smEndDate & "'))"
                SQLQuery = SQLQuery & "ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
                Set rst_Lst = gSQLSelectCall(SQLQuery)
                
                If Not rst_Lst.EOF Then
                    While Not rst_Lst.EOF
                        grdSpots.TextMatrix(llRow, lSTLOGDATENDEX) = Trim$(rst_Lst!lstLogDate)
                        grdSpots.TextMatrix(llRow, LSTLOGTIMEINDEX) = Format(Trim$(rst_Lst!lstLogTime), "hh:mmA/P")
                        grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = Trim$(rst_Lst!lstStatus)
                        Select Case Trim$(rst_Lst!lstStatus)
                        Case 0
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Carried"
                        Case 1
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Delay"
                        Case 2, 3, 4, 5
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Not Carried"
                        Case 7
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Special"
                        Case 8
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Off Air"
                        Case 9
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Delay Cmml/Prg"
                        Case 10
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Air Cmml Only"
                        Case ASTEXTENDED_MG
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "MG"
                        Case ASTEXTENDED_REPLACEMENT
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Replacement"
                        Case ASTEXTENDED_BONUS
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Bonus"
                        Case Else
                            grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) = "Unknown"
                        End Select
                        
                        'If grdSpots.ColWidth(LSTCODEINDEX) <> 0 Then
                            grdSpots.TextMatrix(llRow, LSTCODEINDEX) = Trim$(rst_Lst!lstCode)
                        'End If
                        
                        'If grdSpots.ColWidth(ISCICODEINDEX) <> 0 Then
                            grdSpots.TextMatrix(llRow, ISCICODEINDEX) = Trim$(rst_Lst!lstISCI)
                        'End If
                        
                        'If grdSpots.ColWidth(SDFCODEINDEX) <> 0 Then
                            grdSpots.TextMatrix(llRow, SDFCODEINDEX) = Trim$(rst_Lst!lstSdfCode)
                        'End If
                        
                        If Not rst_Lst.EOF Then
                            llRow = llRow + 1
                            If llRow >= grdSpots.Rows Then
                                grdSpots.AddItem ""
                            End If
                        End If
                        
                        rst_Lst.MoveNext
                    Wend
                Else
                    grdSpots.TextMatrix(llRow, lSTLOGDATENDEX) = "No Post Log Spots"
                End If
            End If
        End If
    Else
        If llRow >= grdSpots.Rows Then
            grdSpots.AddItem ""
        End If
    End If
    
    lmMaxRows = llRow
    grdSpots.Left = grdCptt.Left
    'grdSpots.LeftCol = 1
    grdSpots.Top = grdCptt.Top + grdCptt.Height
    grdSpots.Height = mSpotGridHeight
    grdSpots.TopRow = SPOTSTOPROW
    grdSpots.Visible = True
    grdSpots.Visible = True
    cmdExport.Visible = True
    gSetMousePointer grdCptt, grdSpots, vbDefault
    
    gGrid_AlignAllColsLeft grdSpots

    rst_Cptt.Close
    rst_DAT.Close
    rst_Ast.Close
    rst_Lst.Close
    If ilAstFound Then
        rst_Temp.Close
    End If
    
    mFillSpotGrid = True

    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCompliance Tracer - mFillSpotGrid"
End Function

Private Function mFillSpotTitles() As Boolean

    'D.S. 06/4/11
    
    Dim ilBannerRow As Long
    Dim ilTitle1Row As Integer
    Dim ilTitle2Row As Integer
    Dim slTemp As String
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    
    On Error GoTo ErrHand
    
    mFillSpotTitles = False
    ilBannerRow = 0
    ilTitle1Row = 1
    ilTitle2Row = 2
    
    grdSpots.Row = 0
    grdSpots.Width = mCpttGridWidth
    If ilBannerRow >= grdSpots.Rows Then
        grdSpots.AddItem ""
    End If
    grdSpots.Clear
    grdSpots.Row = 0
    grdSpots.TextMatrix(ilBannerRow, DATFDDAYSINDEX) = "       Fed"
    grdSpots.TextMatrix(ilTitle1Row, DATFDDAYSINDEX) = "Days"
    grdSpots.TextMatrix(ilTitle2Row, DATFDDAYSINDEX) = ""
    grdSpots.ColWidth(DATFDDAYSINDEX) = grdSpots.Width * 0.09
    grdSpots.TextMatrix(ilBannerRow, DATFDTIMESINDEX) = "       Fed"
    grdSpots.TextMatrix(ilTitle1Row, DATFDTIMESINDEX) = "Times"
    grdSpots.TextMatrix(ilTitle2Row, DATFDTIMESINDEX) = ""
    grdSpots.ColWidth(DATFDTIMESINDEX) = grdSpots.Width * 0.11
    
    grdSpots.MergeCells = 2
    grdSpots.MergeRow(0) = True

    grdSpots.TextMatrix(ilBannerRow, DATFDSTATUSINDEX) = "                            Pledged"
    grdSpots.TextMatrix(ilTitle1Row, DATFDSTATUSINDEX) = "Status"
    grdSpots.TextMatrix(ilTitle2Row, DATFDSTATUSINDEX) = ""
    grdSpots.ColWidth(DATFDSTATUSINDEX) = grdSpots.Width * 0.08
    grdSpots.TextMatrix(ilBannerRow, DATAIRPLAYNOINDEX) = "                            Pledged"
    grdSpots.TextMatrix(ilTitle1Row, DATAIRPLAYNOINDEX) = "Air"
    grdSpots.TextMatrix(ilTitle2Row, DATAIRPLAYNOINDEX) = "Play"
    grdSpots.ColWidth(DATAIRPLAYNOINDEX) = grdSpots.Width * 0.05
    grdSpots.TextMatrix(ilBannerRow, DATPDDAYSINDEX) = "                            Pledged"
    grdSpots.TextMatrix(ilTitle1Row, DATPDDAYSINDEX) = "Days"
    grdSpots.TextMatrix(ilTitle2Row, DATPDDAYSINDEX) = ""
    grdSpots.ColWidth(DATPDDAYSINDEX) = grdSpots.Width * 0.09
    grdSpots.TextMatrix(ilBannerRow, DATPDTIMEINDEX) = "                            Pledged"
    grdSpots.TextMatrix(ilTitle1Row, DATPDTIMEINDEX) = "Times"
    grdSpots.TextMatrix(ilTitle2Row, DATPDTIMEINDEX) = ""
    grdSpots.ColWidth(DATPDTIMEINDEX) = grdSpots.Width * 0.11
    grdSpots.TextMatrix(ilBannerRow, DATCODEINDEX) = "                            Pledged"
    grdSpots.TextMatrix(ilTitle1Row, DATCODEINDEX) = "DAT"
    grdSpots.TextMatrix(ilTitle2Row, DATCODEINDEX) = "Code"
    If smUserName <> "Guide" Then
        grdSpots.ColWidth(DATCODEINDEX) = 0
    End If

    grdSpots.MergeCells = 2
    grdSpots.MergeRow(0) = True

    grdSpots.TextMatrix(ilBannerRow, ASTCPSTATUSINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, ASTCPSTATUSINDEX) = "CPTT"
    grdSpots.TextMatrix(ilTitle2Row, ASTCPSTATUSINDEX) = "Status"
    grdSpots.TextMatrix(ilBannerRow, ASTPDSTATUSINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, ASTPDSTATUSINDEX) = "Pledge"
    grdSpots.TextMatrix(ilTitle2Row, ASTPDSTATUSINDEX) = "Status"
    grdSpots.TextMatrix(ilBannerRow, ASTSTATUSINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, ASTSTATUSINDEX) = "Ast"
    grdSpots.TextMatrix(ilTitle2Row, ASTSTATUSINDEX) = "Status"
    grdSpots.TextMatrix(ilBannerRow, LSTADFNAMEINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, LSTADFNAMEINDEX) = "Advertiser"
    grdSpots.TextMatrix(ilTitle2Row, LSTADFNAMEINDEX) = "Name"
    grdSpots.TextMatrix(ilBannerRow, ASTFDDATEINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, ASTFDDATEINDEX) = "Feed"
    grdSpots.TextMatrix(ilTitle2Row, ASTFDDATEINDEX) = "Date"
    grdSpots.TextMatrix(ilBannerRow, ASTFDTIMEINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, ASTFDTIMEINDEX) = "Feed"
    grdSpots.TextMatrix(ilTitle2Row, ASTFDTIMEINDEX) = "Time"
    grdSpots.TextMatrix(ilBannerRow, ASTLKASTCODEINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, ASTLKASTCODEINDEX) = "Missed or"
    grdSpots.TextMatrix(ilTitle2Row, ASTLKASTCODEINDEX) = "MG Link"
    grdSpots.TextMatrix(ilBannerRow, ASTLKDATEINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, ASTLKDATEINDEX) = "Missed or"
    grdSpots.TextMatrix(ilTitle2Row, ASTLKDATEINDEX) = "MG Date"
    grdSpots.TextMatrix(ilBannerRow, ASTCODEINDEX) = "             Station Spot Infomation"
    grdSpots.TextMatrix(ilTitle1Row, ASTCODEINDEX) = "AST"
    grdSpots.TextMatrix(ilTitle2Row, ASTCODEINDEX) = "Code"
    If smUserName <> "Guide" Then
        grdSpots.ColWidth(ASTLKASTCODEINDEX) = 0
        grdSpots.ColWidth(ASTLKDATEINDEX) = 0
        grdSpots.ColWidth(ASTCODEINDEX) = 0
    End If

    grdSpots.MergeCells = 2
    grdSpots.MergeRow(0) = True

           'Post Log

    grdSpots.TextMatrix(ilBannerRow, lSTLOGDATENDEX) = "               Post Log"
    grdSpots.TextMatrix(ilTitle1Row, lSTLOGDATENDEX) = "Log"
    grdSpots.TextMatrix(ilTitle2Row, lSTLOGDATENDEX) = "Date"
    grdSpots.TextMatrix(ilBannerRow, LSTLOGTIMEINDEX) = "               Post Log"
    grdSpots.TextMatrix(ilTitle1Row, LSTLOGTIMEINDEX) = "Log"
    grdSpots.TextMatrix(ilTitle2Row, LSTLOGTIMEINDEX) = "Time"
    grdSpots.TextMatrix(ilBannerRow, LSTSTATUSINDEX) = "               Post Log"
    grdSpots.TextMatrix(ilTitle1Row, LSTSTATUSINDEX) = "Log"
    grdSpots.TextMatrix(ilTitle2Row, LSTSTATUSINDEX) = "Status"
    grdSpots.TextMatrix(ilBannerRow, LSTCODEINDEX) = "               Post Log"
    grdSpots.TextMatrix(ilTitle1Row, LSTCODEINDEX) = "LST"
    grdSpots.TextMatrix(ilTitle2Row, LSTCODEINDEX) = "Code"
    If smUserName <> "Guide" Then
        grdSpots.ColWidth(LSTCODEINDEX) = 0
    End If
    grdSpots.TextMatrix(ilBannerRow, ISCICODEINDEX) = "               Post Log"
    grdSpots.TextMatrix(ilTitle1Row, ISCICODEINDEX) = "ISCI"
    grdSpots.TextMatrix(ilTitle2Row, ISCICODEINDEX) = "Code"
    grdSpots.TextMatrix(ilBannerRow, SDFCODEINDEX) = "               Post Log"
    grdSpots.TextMatrix(ilTitle1Row, SDFCODEINDEX) = "SDF"
    grdSpots.TextMatrix(ilTitle2Row, SDFCODEINDEX) = "Code"
    If smUserName <> "Guide" Then
        grdSpots.ColWidth(SDFCODEINDEX) = 0
    End If
    
    grdSpots.MergeCells = 2
    grdSpots.MergeRow(0) = True
    
    grdSpots.Width = mCpttGridWidth
    grdSpots.Width = 0
    grdSpots.Width = mCpttGridWidth
    
    mFillSpotTitles = True
    
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmComplianceTracer - mFillSpotTitles: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
    gLogMsg " ", "AffErrorLog.txt", False
    
End Function

Private Function mFillCpttTitles() As Boolean

    'D.S. 06/4/11
    
    Dim llRow As Long
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String

    On Error GoTo ErrHand
    
    mFillCpttTitles = False
    
    grdCptt.Clear
    grdSpots.Width = mCpttGridWidth
    grdSpots.Width = 0
    grdSpots.Width = mCpttGridWidth
    
    llRow = 0
    'llRow = grdCptt.FixedRows
    
    grdCptt.Width = mCpttGridWidth
    If llRow > grdCptt.Rows Then
        grdCptt.AddItem ""
    
    End If
    grdCptt.TextMatrix(llRow, CPTTWEEKINDEX) = "Week"
    grdCptt.TextMatrix(llRow + 1, CPTTWEEKINDEX) = ""
    grdCptt.ColWidth(CPTTWEEKINDEX) = grdCptt.Width * 0.06
    
    grdCptt.TextMatrix(llRow, VEHICLEINDEX) = "Vehicle"
    grdCptt.TextMatrix(llRow + 1, VEHICLEINDEX) = "Name"
    grdCptt.ColWidth(VEHICLEINDEX) = grdCptt.Width * 0.17
    
    grdCptt.TextMatrix(llRow, STATIONINDEX) = "Station"
    grdCptt.TextMatrix(llRow + 1, STATIONINDEX) = "Name"
    grdCptt.ColWidth(STATIONINDEX) = grdCptt.Width * 0.08
    
    grdCptt.TextMatrix(llRow, ZONEINDEX) = "Time"
    grdCptt.TextMatrix(llRow + 1, ZONEINDEX) = "Zone"
    grdCptt.ColWidth(ZONEINDEX) = grdCptt.Width * 0.04
    
    grdCptt.TextMatrix(llRow, OFFSETINDEX) = "Zone"
    grdCptt.TextMatrix(llRow + 1, OFFSETINDEX) = "Offset"
    grdCptt.ColWidth(OFFSETINDEX) = grdCptt.Width * 0.05
    
    grdCptt.TextMatrix(llRow, DATERANGEINDEX) = "Date"
    grdCptt.TextMatrix(llRow + 1, DATERANGEINDEX) = "Range"
    grdCptt.ColWidth(DATERANGEINDEX) = grdCptt.Width * 0.12
    
    grdCptt.TextMatrix(llRow, CPSTATUSINDEX) = "CP"
    grdCptt.TextMatrix(llRow + 1, CPSTATUSINDEX) = "Status"
    grdCptt.ColWidth(CPSTATUSINDEX) = grdCptt.Width * 0.05
    
    grdCptt.TextMatrix(llRow, POSTINGSTATUSINDEX) = "Posting"
    grdCptt.TextMatrix(llRow + 1, POSTINGSTATUSINDEX) = "Status"
    grdCptt.ColWidth(POSTINGSTATUSINDEX) = grdCptt.Width * 0.05
    
    grdCptt.TextMatrix(llRow, CPTTASTSTATUSINDEX) = "AST"
    grdCptt.TextMatrix(llRow + 1, CPTTASTSTATUSINDEX) = "Status"
    grdCptt.ColWidth(CPTTASTSTATUSINDEX) = grdCptt.Width * 0.07
    
    grdCptt.TextMatrix(llRow, NUMCOMPLIANTINDEX) = "Number"
    grdCptt.TextMatrix(llRow + 1, NUMCOMPLIANTINDEX) = "Compliant"
    grdCptt.ColWidth(NUMCOMPLIANTINDEX) = grdCptt.Width * 0.07
    
    grdCptt.TextMatrix(llRow, NUMAIREDINDEX) = "Number"
    grdCptt.TextMatrix(llRow + 1, NUMAIREDINDEX) = "Aired"
    grdCptt.ColWidth(NUMAIREDINDEX) = grdCptt.Width * 0.07
    
    grdCptt.TextMatrix(llRow, NUMGENEDINDEX) = "Number"
    grdCptt.TextMatrix(llRow + 1, NUMGENEDINDEX) = "Gened"
    grdCptt.ColWidth(NUMGENEDINDEX) = grdCptt.Width * 0.07
    
    grdCptt.TextMatrix(llRow, CPTTATTCODEINDEX) = "ATT"
    grdCptt.TextMatrix(llRow + 1, CPTTATTCODEINDEX) = "Code"
    grdCptt.ColWidth(CPTTATTCODEINDEX) = grdCptt.Width * 0.05
    If smUserName <> "Guide" Then
        grdCptt.ColWidth(CPTTATTCODEINDEX) = 0
    End If
    grdCptt.TextMatrix(llRow, CPTTCODEINDEX) = "CPTT"
    grdCptt.TextMatrix(llRow + 1, CPTTCODEINDEX) = "Code"
    grdCptt.ColWidth(CPTTCODEINDEX) = grdCptt.Width * 0.05
    If smUserName <> "Guide" Then
        grdCptt.ColWidth(CPTTCODEINDEX) = 0
    End If
    
    mFillCpttTitles = True

    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmComplianceTracer - mFillCpttTitles: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
    gLogMsg " ", "AffErrorLog.txt", False

End Function


Private Function mGetDatFdDays(rst_DAT As ADODB.Recordset) As String

    'D.S. 06/4/11
    
    Dim slDayStr As String
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    
    On Error GoTo ErrHand
    
    mGetDatFdDays = ""

    If rst_DAT!datFdMon = 1 Then
        slDayStr = "Y"
    Else
        slDayStr = "N"
    End If
    If rst_DAT!datFdTue = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datFdWed = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datFdThu = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datFdFri = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datFdSat = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datFdSun = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If

    mGetDatFdDays = slDayStr

    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmComplianceTracer - mGetDatFdDays: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

     ilLine = Erl
     ilErrNo = Err.Number
     slDesc = Err.Description
     gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
     gLogMsg " ", "AffErrorLog.txt", False

End Function

Private Function mGetDatPdDays(rst_DAT As ADODB.Recordset) As String

    'D.S. 06/4/11
    
    Dim slDayStr As String
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    
    On Error GoTo ErrHand
    
    mGetDatPdDays = ""

    If rst_DAT!datPdMon = 1 Then
        slDayStr = "Y"
    Else
        slDayStr = "N"
    End If
    If rst_DAT!datPdTue = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datPdWed = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datPdThu = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datPdFri = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datPdSat = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If
    If rst_DAT!datPdSun = 1 Then
        slDayStr = slDayStr & "Y"
    Else
        slDayStr = slDayStr & "N"
    End If

    mGetDatPdDays = slDayStr
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmComplianceTracer - mGetDatPdDays: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
    gLogMsg " ", "AffErrorLog.txt", False

End Function

Private Sub mSaveSpotsGridToFile()

    'D.S. 06/4/11
    
    Dim llRow As Long
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    Dim slTemp As String
    
    On Error GoTo ErrHand

    'Write Header Record
    slTemp = "Fd Days,"
    slTemp = slTemp & "Fd Times,"
    slTemp = slTemp & "Fd Status,"
    slTemp = slTemp & "Air Play #,"
    slTemp = slTemp & "Pd Days,"
    slTemp = slTemp & "Pd Times,"
    If smUserName = "Guide" Then
        slTemp = slTemp & "DAT Code,"
    End If
    slTemp = slTemp & "CP Status,"
    slTemp = slTemp & "Pledge Status,"
    slTemp = slTemp & "Ast Status,"
    slTemp = slTemp & "Advertiser,"
    slTemp = slTemp & "Feed Date,"
    slTemp = slTemp & "Feed Time,"
    If smUserName = "Guide" Then
        slTemp = slTemp & "Lk AST Code,"
        slTemp = slTemp & "Lk Date,"
        slTemp = slTemp & "AST Code,"
    End If
    slTemp = slTemp & "Log Date,"
    slTemp = slTemp & "Log Time,"
    slTemp = slTemp & "Status,"
    If smUserName = "Guide" Then
        slTemp = slTemp & "LST Code,"
    End If
    If smUserName = "Guide" Then
        slTemp = slTemp & "ISCI,"
        slTemp = slTemp & "SDF Code"
    Else
        slTemp = slTemp & "ISCI"
    End If

    'Print to file
    Print #hmToDetail, slTemp
    
    For llRow = grdCptt.FixedRows To lmMaxRows - 1
        slTemp = grdSpots.TextMatrix(llRow, DATFDDAYSINDEX) & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, DATFDTIMESINDEX) & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, DATFDSTATUSINDEX) & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, DATAIRPLAYNOINDEX) & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, DATPDDAYSINDEX) & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, DATPDTIMEINDEX) & ","

        If smUserName = "Guide" Then
            slTemp = slTemp & grdSpots.TextMatrix(llRow, DATCODEINDEX) & ","
        End If
        slTemp = slTemp & grdSpots.TextMatrix(llRow, ASTCPSTATUSINDEX) & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, ASTPDSTATUSINDEX) & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, ASTSTATUSINDEX) & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, LSTADFNAMEINDEX) & ","
        slTemp = slTemp & Format(grdSpots.TextMatrix(llRow, ASTFDDATEINDEX), "mm/dd/yy") & ","
        slTemp = slTemp & Format(grdSpots.TextMatrix(llRow, ASTFDTIMEINDEX), "hh:mm:ssA/P") & ","
        If smUserName = "Guide" Then
            slTemp = slTemp & grdSpots.TextMatrix(llRow, ASTLKASTCODEINDEX) & ","
            slTemp = slTemp & grdSpots.TextMatrix(llRow, ASTLKDATEINDEX) & ","
        End If
        If smUserName = "Guide" Then
            slTemp = slTemp & grdSpots.TextMatrix(llRow, ASTCODEINDEX) & ","
        End If
        slTemp = slTemp & Format(grdSpots.TextMatrix(llRow, lSTLOGDATENDEX), "mm/dd/yy") & ","
        slTemp = slTemp & Format(grdSpots.TextMatrix(llRow, LSTLOGTIMEINDEX), "hh:mm:ssA/P") & ","
        slTemp = slTemp & grdSpots.TextMatrix(llRow, LSTSTATUSINDEX) & ","
        If smUserName = "Guide" Then
            slTemp = slTemp & grdSpots.TextMatrix(llRow, LSTCODEINDEX) & ","
        End If
        If smUserName = "Guide" Then
            slTemp = slTemp & grdSpots.TextMatrix(llRow, ISCICODEINDEX) & ","
            slTemp = slTemp & grdSpots.TextMatrix(llRow, SDFCODEINDEX)
        Else
            slTemp = slTemp & grdSpots.TextMatrix(llRow, ISCICODEINDEX)
        End If
        
        'Print to file
        Print #hmToDetail, slTemp
    
    Next llRow
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmCompliance Tracer - mSaveSpotsGridToFile: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
     
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
    gLogMsg " ", "AffErrorLog.txt", False

    Exit Sub
End Sub


Private Sub mSaveCpttGridToFile()

    'D.S. 06/4/11
    
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String
    Dim slTemp As String

    On Error GoTo ErrHand
    'Write Header Record
    slTemp = "Week,"
    slTemp = slTemp & "Vehicle,"
    slTemp = slTemp & "Station,"
    slTemp = slTemp & "Zone,"
    slTemp = slTemp & "Offset,"
    slTemp = slTemp & "Date Range,"
    slTemp = slTemp & "CP Status,"
    slTemp = slTemp & "Posting Status,"
    slTemp = slTemp & "Ast Status,"
    slTemp = slTemp & "# Compliant,"
    slTemp = slTemp & "# Aired,"
    If smUserName = "Guide" Then
        slTemp = slTemp & "# Gen.,"
        slTemp = slTemp & "ATTCode,"
        slTemp = slTemp & "CPTT Code"
    Else
        slTemp = slTemp & "# Gen."
    End If
    'Print to file
    Print #hmToDetail, slTemp
    
    'Get the Row Values
    slTemp = grdCptt.TextMatrix(lmRow, CPTTWEEKINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, VEHICLEINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, STATIONINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, ZONEINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, OFFSETINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, DATERANGEINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, CPSTATUSINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, POSTINGSTATUSINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, CPTTASTSTATUSINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, NUMCOMPLIANTINDEX) & ","
    slTemp = slTemp & grdCptt.TextMatrix(lmRow, NUMAIREDINDEX) & ","
    If smUserName = "Guide" Then
        slTemp = slTemp & grdCptt.TextMatrix(lmRow, NUMGENEDINDEX) & ","
        slTemp = slTemp & grdCptt.TextMatrix(lmRow, CPTTATTCODEINDEX) & ","
        slTemp = slTemp & grdCptt.TextMatrix(lmRow, CPTTCODEINDEX)
    Else
        slTemp = slTemp & grdCptt.TextMatrix(lmRow, NUMGENEDINDEX)
    End If

    'Print to file
    Print #hmToDetail, slTemp
    Print #hmToDetail, ""
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmCompliance Tracer - mSaveSpotsGridToFile: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
     
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
    gLogMsg " ", "AffErrorLog.txt", False

    Exit Sub
    
    Exit Sub
End Sub


Private Function mOpenFiles() As Boolean

    'D.S. 06/4/11
    
    Dim ilRet As Integer

    mOpenFiles = False
    
    'ilRet = 0
    'hmToDetail = FreeFile
    'Open smTxtFile For Output Lock Write As hmToDetail
    ilRet = gFileOpen(smTxtFile, "Output Lock Write", hmToDetail)
    
    If ilRet <> 0 Then
        gLogMsg "** Terminated - " & smTxtFile & " failed to open. **", "AffErrorLog.Txt", False
        Close #hmToDetail
        gMsgBox "Open Error #" & Str$(Err.Numner) & smTxtFile, vbOKOnly, "Open Error"
        mOpenFiles = False
        Exit Function
    End If
    
    mOpenFiles = True
    Exit Function
'cmdExportErr:
'    ilRet = Err
'    Resume Next

End Function

Private Sub mCloseFiles()

    Close #hmToDetail
    
End Sub

Private Sub cmdExport_Click()
    
    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.InitDir = sgExportDirectory
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNAllowMultiselect
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowSave
    ' Display name of selected file

    smTxtFile = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    gChDrDir
    If mOpenFiles Then
        mSaveCpttGridToFile
        mSaveSpotsGridToFile
    End If
    smTxtFile = ""
    mCloseFiles
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    ChDir slCurDir
    gChDrDir
    Exit Sub

End Sub

'Private Sub mSaveCpttGrid()
'
'   'D.S. 06/4/11
'    Dim ilIdx As Integer
'    Dim ilLine As Integer
'    Dim ilErrNo As Integer
'    Dim slDesc As String
'
'    On Error GoTo ErrHand
'
'    For ilIdx = grdCptt.FixedRows To grdCptt.Rows - 1
'        If ilIdx < grdCptt.Row Or ilIdx > grdCptt.RowSel Then
'           grdCptt.RowData(ilIdx) = grdCptt.RowHeight(ilIdx)
'           'grdCptt.RowHeight(ilIdx) = 0
'        End If
'    Next
'    For ilIdx = grdCptt.Cols To grdCptt.Cols - 1
'       If ilIdx < grdCptt.Col Or ilIdx > grdCptt.ColSel Then
'          grdCptt.ColData(ilIdx) = grdCptt.ColWidth(ilIdx)
'          'grdCptt.ColWidth(ilIdx) = 0
'       End If
'    Next
'    grdCptt.Height = mCpttGridHeight
'    grdCptt.TOPROW = TOPROW
'    smFirstTime = False
'
'    Exit Sub
'ErrHand:
'    Screen.MousePointer = vbDefault
'
'    gMsg = ""
'    If (Err.Number <> 0) And (gMsg = "") Then
'        gMsg = "A general error has occured in frmCompliance Tracer - mSaveCpttGrid: "
'        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
'    End If
'
'    ilLine = Erl
'    ilErrNo = Err.Number
'    slDesc = Err.Description
'    gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
'    gLogMsg " ", "AffErrorLog.txt", False
'
'End Sub
'
'Private Sub mRestoreCpttGrid()
'
'    Dim ilIdx As Integer
'    Dim ilLine As Integer
'    Dim ilErrNo As Integer
'    Dim slDesc As String
'
'    On Error GoTo ErrHand
'
'    grdSpots.Visible = False
'    For ilIdx = grdCptt.FixedRows To grdCptt.Rows - 1
'       If ilIdx < grdCptt.Row Or ilIdx > grdCptt.RowSel Then
'          grdCptt.RowHeight(ilIdx) = grdCptt.RowData(ilIdx)
'       End If
'    Next
'    For ilIdx = grdCptt.FixedCols To grdCptt.Cols - 1
'       If ilIdx < grdCptt.Col Or ilIdx > grdCptt.ColSel Then
'          grdCptt.ColWidth(ilIdx) = grdCptt.ColData(ilIdx)
'       End If
'    Next
'
'    grdCptt.TopRow = TOPROW
'
'    grdCptt.Height = mCpttGridHeight
'    grdCptt.Redraw = True
'
'    Exit Sub
'
'ErrHand:
'    Screen.MousePointer = vbDefault
'
'    gMsg = ""
'    If (Err.Number <> 0) And (gMsg = "") Then
'        gMsg = "A general error has occured in frmCompliance Tracer - mRestoreCpttGrid: "
'        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
'    End If
'
'    ilLine = Erl
'    ilErrNo = Err.Number
'    slDesc = Err.Description
'    gLogMsg "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc, "AffErrorLog.txt", False
'    gLogMsg " ", "AffErrorLog.txt", False
'
'End Sub

Private Sub grdSpots_Click()

End Sub

Private Sub grdSpots_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'grdSpots.ToolTipText = ""
    If (grdSpots.MouseRow >= grdSpots.FixedRows) And (grdSpots.MouseCol > 0) And (grdSpots.TextMatrix(grdSpots.MouseRow, grdSpots.MouseCol)) <> "" Then
        grdSpots.ToolTipText = grdSpots.TextMatrix(grdSpots.MouseRow, grdSpots.MouseCol)
    Else
        grdSpots.ToolTipText = ""
    End If

End Sub


Public Function mBinarySearchLst(llCode As Long) As Long
    
    'D.S. 06/4/11
    
    'Returns the index number of lmLstCode that matches the lstCode that was passed in
    'Note: for this to work tglsttInfo was previously sorted
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long

    llMin = LBound(lmLstCode)
    llMax = UBound(lmLstCode) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = lmLstCode(llMiddle) Then
            'found the match
            mBinarySearchLst = llMiddle
            Exit Function
        ElseIf llCode < lmLstCode(llMiddle) Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchLst = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mBinarySearchLst: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    mBinarySearchLst = -1
    Exit Function
    
End Function
Private Sub mResizeCPTTColumns()
    Dim ilCol As Integer
    grdCptt.ColWidth(VEHICLEINDEX) = mCpttGridWidth - GRIDSCROLLWIDTH - 15
    For ilCol = CPTTWEEKINDEX To CPTTCODEINDEX Step 1
        If (ilCol <> VEHICLEINDEX) Then
            grdCptt.ColWidth(VEHICLEINDEX) = grdCptt.ColWidth(VEHICLEINDEX) - grdCptt.ColWidth(ilCol)
        End If
    Next ilCol
    
End Sub
Private Sub mHideSpotGrid()
    grdCptt.Width = mCpttGridWidth
    grdCptt.ColWidth(VEHICLEINDEX) = grdCptt.ColWidth(VEHICLEINDEX) - 275
    grdCptt.Height = mCpttGridHeight
    lmRow = lmRow
    grdCptt.ScrollBars = flexScrollBarVertical
    grdSpots.Visible = False
    cmdExport.Visible = False
    grdCptt.TopRow = CPTTTOPROW
    grdCptt.LeftCol = 1
    
    grdCptt.Redraw = True
End Sub

