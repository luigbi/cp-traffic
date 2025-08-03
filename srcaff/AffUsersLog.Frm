VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmUsersLog 
   Caption         =   "Users Log"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "AffUsersLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   Begin VB.CheckBox ckcActivity 
      Caption         =   "Show Activity Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7095
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Frame frcUsers 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3345
      TabIndex        =   10
      Top             =   30
      Width           =   2340
      Begin VB.OptionButton rbcUsers 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   12
         Top             =   15
         Width           =   615
      End
      Begin VB.OptionButton rbcUsers 
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   11
         Top             =   15
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.Label lacUsers 
         Caption         =   "Users"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   570
      End
   End
   Begin VB.CommandButton cmcRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7470
      TabIndex        =   9
      Top             =   4575
      Width           =   1335
   End
   Begin VB.CommandButton cmcSignOn 
      Caption         =   "Sign On &Ok Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   4575
      Width           =   1995
   End
   Begin VB.CommandButton cmcSignOff 
      Caption         =   "&Sign Off Alert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3495
      TabIndex        =   6
      Top             =   4575
      Width           =   1335
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   195
      TabIndex        =   4
      Top             =   4575
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.PictureBox pbcUsersLogFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   30
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   0
      Width           =   60
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   90
      Top             =   4635
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5025
      FormDesignWidth =   9105
   End
   Begin VB.CommandButton cmcMessage 
      Caption         =   "&Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   3
      Top             =   4575
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdUsersLog 
      Height          =   3720
      Left            =   165
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6562
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdActivity 
      Height          =   1590
      Left            =   915
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2805
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
   Begin VB.Label lacNoUsers 
      Alignment       =   1  'Right Justify
      Caption         =   "Active:   Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6060
      TabIndex        =   8
      Top             =   4215
      Width           =   2775
   End
   Begin VB.Label lacInfo 
      Caption         =   "Mouse Over: Name to see PC Name plus Network card ID; E-Mail to see Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   150
      TabIndex        =   5
      Top             =   4230
      Width           =   6045
   End
End
Attribute VB_Name = "frmUsersLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmUsersLog - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private lmScrollTop As Long
Private imCtrlKey As Integer
Private imShiftKey As Integer
Private lmLastClickedRow As Long
Private imAllowOkMsg As Integer

Private imLastUsersLogColSorted As Integer
Private imLastUsersLogSort As Integer

Private imLastActivityColSorted As Integer
Private imLastActivitySort As Integer

Private rst_Ulf As ADODB.Recordset
Private rst_Uaf As ADODB.Recordset
Private rst_User As ADODB.Recordset

Dim hmUlf As Integer
Dim tmUlf As ULF        'GSF record image
Dim tmUlfSrchKey0 As LONGKEY0    'GSF key record image
Dim tmUlfSrchKey1 As ULFKEY1    'GSF key record image
Dim imUlfRecLen As Integer        'GSF record length
'Dan M 11/06/09 keep track if selected has email--if messages selected know to enable email option
Dim bmEnableEmail As Boolean
'Grid Controls

Const SYSTEMINDEX = 0
Const USERNAMEINDEX = 1
Const CITYINDEX = 2
Const PHONEINDEX = 3
Const EMAILFLAGINDEX = 4
Const ZONEINDEX = 5
Const SIGNONINDEX = 6
Const ACTIVITYINDEX = 7
Const SIGNOFFINDEX = 8
Const DURATIONINDEX = 9
Const SORTINDEX = 10
Const EMAILVALUEINDEX = 11
Const PCMACINDEX = 12
Const SELECTEDINDEX = 13
Const SIGNONUSERINDEX = 14
Const ULFCODEINDEX = 15

Const AACTIVITYINDEX = 0
Const ASTARTDATEINDEX = 1
Const AENDDATEINDEX = 2
Const ASTARTTIMEINDEX = 3
Const AENDTIMEINDEX = 4
Const ADURATIONINDEX = 5
Const ASTATUSINDEX = 6
Const ACSIDATEINDEX = 7
Const ASORTINDEX = 8
Const AUAFCODEINDEX = 9






Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    
    'Blank rows within grid
    gGrid_Clear grdUsersLog, True
    'Set color within cells
    For llRow = grdUsersLog.FixedRows To grdUsersLog.Rows - 1 Step 1
        For llCol = 0 To ULFCODEINDEX Step 1
        '    grdUsersLog.Row = llRow
        '    grdUsersLog.Col = llCol
            grdUsersLog.TextMatrix(llRow, llCol) = ""
        Next llCol
    Next llRow
    mClearActivityGrid
End Sub
Private Sub mClearActivityGrid()
    Dim llRow As Long
    Dim llCol As Long

    'Blank rows within grid
    For llRow = grdActivity.FixedRows To grdActivity.Rows - 1 Step 1
        For llCol = 0 To AUAFCODEINDEX Step 1
            grdActivity.Row = llRow
            grdActivity.Col = llCol
            grdActivity.CellForeColor = vbBlack
            grdActivity.TextMatrix(llRow, llCol) = ""
        Next llCol
    Next llRow
End Sub

Private Sub ckcActivity_Click()
    Dim llRow As Long
    Dim ilCount As Integer
    Dim llSelectedRow As Long
    
    gSetMousePointer grdUsersLog, grdActivity, vbHourglass
    If ckcActivity.Value = vbChecked Then
        grdUsersLog.Height = grdUsersLog.Height / 2
        gGrid_IntegralHeight grdUsersLog
        grdActivity.Height = lacInfo.Top - (grdUsersLog.Top + grdUsersLog.Height) - 240
        grdActivity.Top = grdUsersLog.Top + grdUsersLog.Height + 120
        gGrid_IntegralHeight grdActivity
        grdActivity.Visible = True
    Else
        grdUsersLog.Height = lacInfo.Top - grdUsersLog.Top - 120
        gGrid_IntegralHeight grdUsersLog
        grdActivity.Visible = False
    End If
    ilCount = 0
    For llRow = grdUsersLog.FixedRows To grdUsersLog.Rows - 1 Step 1
        If grdUsersLog.TextMatrix(llRow, SYSTEMINDEX) <> "" Then
            If grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                ilCount = ilCount + 1
                llSelectedRow = llRow
                If ilCount > 1 Then
                    Exit For
                End If
            End If
        End If
    Next llRow
    If ilCount > 1 Then
        For llRow = grdUsersLog.FixedRows To grdUsersLog.Rows - 1 Step 1
            If grdUsersLog.TextMatrix(llRow, SYSTEMINDEX) <> "" Then
                If grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                    grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "0"
                    mPaintRowColor llRow
                End If
            End If
        Next llRow
    ElseIf (ilCount = 1) And (ckcActivity.Value = vbChecked) Then
        mPopActivityGrid llSelectedRow
    End If
    If ilCount <> 1 Then
        If (ckcActivity.Value = vbChecked) Then
            mClearActivityGrid
        End If
        lmLastClickedRow = -1
    End If
    mSetCommands
    gSetMousePointer grdUsersLog, grdActivity, vbDefault
End Sub

Private Sub cmcCancel_Click()
    Unload frmUsersLog
End Sub

Private Sub cmcMessage_Click()
    igUserLogButton = 1  '1=Message; 2=Send Alert; 3=Send Ok message
    mSetUsers
    'Dan M 11/06/09 don't allow message if no one selected
    If UBound(lgUserLogUlfCode) > 0 Then
        frmUserLogEMail.bmEnableEmail = bmEnableEmail
        frmUserLogEMail.Show vbModal
    Else
        gMsgBox "Please select user(s)"
    End If
End Sub

Private Sub cmcRefresh_Click()
    gSetMousePointer grdUsersLog, grdActivity, vbHourglass
    mClearGrid
    mPopulate
    mSetCommands
    lmLastClickedRow = -1
    gSetMousePointer grdUsersLog, grdActivity, vbDefault
End Sub

Private Sub cmcSignOff_Click()
    Dim llRow As Long
    Dim ilRet As Integer
    
    For llRow = grdUsersLog.FixedRows To grdUsersLog.Rows - 1 Step 1
        If Trim$(grdUsersLog.TextMatrix(llRow, SYSTEMINDEX)) <> "" Then
            If (Trim$(grdUsersLog.TextMatrix(llRow, SIGNONINDEX)) <> "") And (Trim$(grdUsersLog.TextMatrix(llRow, SIGNOFFINDEX)) = "") Then
                If (grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) <> "1") And (grdUsersLog.TextMatrix(llRow, SIGNONUSERINDEX) <> "Y") Then
                    ilRet = MsgBox("Some Active Users have not been selected for alert notification, Continue anyway?", vbApplicationModal + vbInformation + vbYesNo, "Question")
                    If ilRet = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            End If
        End If
    Next llRow
    igUserLogButton = 2  '1=Message; 2=Send Alert; 3=Send Ok message
    mSetUsers
    frmUserLogEMail.Show vbModal
    mSetOkMsg
    mSetCommands
End Sub

Private Sub cmcSignOn_Click()
    igUserLogButton = 3  '1=Message; 2=Send Alert; 3=Send Ok message
    mSetUsers
    frmUserLogEMail.Show vbModal
    mSetOkMsg
    mSetCommands
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        gSetMousePointer grdUsersLog, grdActivity, vbHourglass
        mSetGridColumns
        mSetGridTitles
        gGrid_IntegralHeight grdUsersLog
        gGrid_FillWithRows grdUsersLog
        mPopulate
        imFirstTime = False
        gSetMousePointer grdUsersLog, grdActivity, vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.15
    Me.Height = Screen.Height / 1.55
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmUsersLog
    gCenterForm frmUsersLog
End Sub

Private Sub Form_Load()
    
    gSetMousePointer grdUsersLog, grdActivity, vbHourglass
    
    mInit
    gSetMousePointer grdUsersLog, grdActivity, vbDefault
    Exit Sub
   
End Sub

Private Sub Form_Resize()
    'grdActivity.Width = (4 * grdUsersLog.Width) / 5
    grdActivity.Width = grdUsersLog.Width
    'grdActivity.Left = grdUsersLog.Left + (grdUsersLog.Width - grdActivity.Width) / 2
    grdActivity.Left = grdUsersLog.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim ilRet As Integer
    ilRet = mCloseCEFFile()
    rst_Ulf.Close
    rst_Uaf.Close
    rst_User.Close
    Set frmUsersLog = Nothing
End Sub

Private Sub grdActivity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim llHeight As Long

    For llRow = 0 To grdActivity.FixedRows - 1 Step 1
        llHeight = grdActivity.RowHeight(llRow)
    Next llRow
    'grdActivity.ToolTipText = ""
    If Y <= llHeight Then
        grdActivity.ToolTipText = ""
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdActivity, X, Y, llRow, llCol)
    grdActivity.ToolTipText = Trim$(grdActivity.TextMatrix(llRow, llCol))
End Sub

Private Sub grdActivity_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < grdActivity.RowHeight(0) Then
        grdActivity.Col = grdActivity.MouseCol
        mActivitySortCol grdActivity.Col
    End If
    grdActivity.Row = 0
    grdActivity.Col = AUAFCODEINDEX
End Sub

Private Sub grdUsersLog_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And CTRLMASK) > 0 Then
        imCtrlKey = True
    Else
        imCtrlKey = False
    End If
    If (Shift And SHIFTMASK) > 0 Then
        imShiftKey = True
    Else
        imShiftKey = False
    End If
End Sub

Private Sub grdUsersLog_KeyUp(KeyCode As Integer, Shift As Integer)
    imCtrlKey = False
    imShiftKey = False
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    lmLastClickedRow = -1
    lmScrollTop = grdUsersLog.FixedRows
    imLastUsersLogColSorted = -1
    imLastUsersLogSort = -1
    
    ilRet = mOpenCEFFile()
    
    mSetOkMsg
    mClearGrid

    If Trim$(sgUstActivityLog) = "V" Then
        ckcActivity.Visible = True
    End If

    mSetCommands
    On Error GoTo 0
    Exit Sub
ErrHand:
    gSetMousePointer grdUsersLog, grdActivity, vbDefault
    gHandleError "AffErrorLog.txt", "frmUserLog-mInit"
End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim ilRet As Integer
    Dim slSignOnDate As String
    Dim slSignOffDate As String
    Dim slSignOnTime As String
    Dim slSignOffTime As String
    Dim ilNoDays As Integer
    Dim ilNoHours As Integer
    Dim ilNoMinutes As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilShowUser As Integer
    Dim slActiveDate As String
    Dim slActiveTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llMinuteDiff As Long
    Dim slZone As String
    ReDim ilActiveUrfCode(0 To 0) As Integer
    ReDim ilActiveUstCode(0 To 0) As Integer
    ReDim ilTotalUrfCode(0 To 0) As Integer
    ReDim ilTotalUstCode(0 To 0) As Integer
    
    On Error GoTo ErrHand:
    grdUsersLog.Redraw = False
    grdUsersLog.Row = 0
    For llCol = SYSTEMINDEX To DURATIONINDEX Step 1
        grdUsersLog.Col = llCol
        grdUsersLog.CellBackColor = LIGHTBLUE
    Next llCol
    llRow = grdUsersLog.FixedRows
    
    hmUlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmUlf, "", sgDBPath & "Ulf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imUlfRecLen = Len(tmUlf)  'btrRecordLength(hlUrf)  'Get and save record length
    slNowDate = Format$(Now, "m/d/yy")
    slNowTime = Format$(Now, "h:mm:ssAM/PM")
    slStr = slNowDate & " " & slNowTime
    slZone = Left$(gGetLocalTZName(), 1)
    Select Case slZone
        Case "E"
            slStr = slNowDate & " " & slNowTime
        Case "C"
            slStr = DateAdd("h", 1, slNowDate & " " & slNowTime)
        Case "M"
            slStr = DateAdd("h", 2, slNowDate & " " & slNowTime)
        Case "P"
            slStr = DateAdd("h", 3, slNowDate & " " & slNowTime)
    End Select
    slNowDate = Format$(slStr, "m/d/yy")
    slNowTime = Format$(slStr, "h:mm:ssAM/PM")
    ilRet = btrGetFirst(hmUlf, tmUlf, imUlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        gUnpackDate tmUlf.iSignOffDate(0), tmUlf.iSignOffDate(1), slSignOffDate
        If DateValue(gAdjYear(slSignOffDate)) = DateValue("12/31/2069") Then
            gUnpackDate tmUlf.iActiveLogDate(0), tmUlf.iActiveLogDate(1), slActiveDate
            gUnpackTime tmUlf.iActiveLogTime(0), tmUlf.iActiveLogTime(1), "A", "1", slActiveTime
            If slActiveDate <> "" Then
                slStr = slActiveDate & " " & slActiveTime
                slZone = Trim$(tmUlf.sTimeZone)
                Select Case slZone
                    Case "E"
                        slStr = slActiveDate & " " & slActiveTime
                    Case "C"
                        slStr = DateAdd("h", 1, slActiveDate & " " & slActiveTime)
                    Case "M"
                        slStr = DateAdd("h", 2, slActiveDate & " " & slActiveTime)
                    Case "P"
                        slStr = DateAdd("h", 3, slActiveDate & " " & slActiveTime)
                End Select
                slActiveDate = Format$(slStr, "m/d/yy")
                slActiveTime = Format$(slStr, "h:mm:ssAM/PM")
                llMinuteDiff = DateDiff("n", slActiveDate & " " & slActiveTime, slNowDate & " " & slNowTime)
                If llMinuteDiff > 120 Then
                    tmUlf.iSignOffDate(0) = tmUlf.iActiveLogDate(0)
                    tmUlf.iSignOffDate(1) = tmUlf.iActiveLogDate(1)
                    tmUlf.iSignOffTime(0) = tmUlf.iActiveLogTime(0)
                    tmUlf.iSignOffTime(1) = tmUlf.iActiveLogTime(1)
                    ilRet = btrUpdate(hmUlf, tmUlf, imUlfRecLen)
                    tmUlfSrchKey0.lCode = tmUlf.lCode
                    ilRet = btrGetEqual(hmUlf, tmUlf, imUlfRecLen, tmUlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                End If
            Else
                'Has not signed on since this release added, set signoff
                tmUlf.iSignOffDate(0) = tmUlf.iSignOnDate(0)
                tmUlf.iSignOffDate(1) = tmUlf.iSignOnDate(1)
                tmUlf.iSignOffTime(0) = tmUlf.iSignOnTime(0)
                tmUlf.iSignOffTime(1) = tmUlf.iSignOnTime(1)
                ilRet = btrUpdate(hmUlf, tmUlf, imUlfRecLen)
                tmUlfSrchKey0.lCode = tmUlf.lCode
                ilRet = btrGetEqual(hmUlf, tmUlf, imUlfRecLen, tmUlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            End If
        End If
        ilRet = btrGetNext(hmUlf, tmUlf, imUlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hmUlf)
    btrDestroy hmUlf
    
    SQLQuery = "SELECT * FROM ULF_User_Log"
    If rbcUsers(0).Value Then
        SQLQuery = SQLQuery & " WHERE ulfSignOffDate = '2069-12-31'"
    End If
    Set rst_Ulf = gSQLSelectCall(SQLQuery)
    Do While Not rst_Ulf.EOF
        ilShowUser = False
        If rst_Ulf!ulfSystemType = "T" Then
            SQLQuery = "SELECT * FROM URF_User_Options Where urfCode = " & rst_Ulf!ulfUrfCode
            Set rst_User = gSQLSelectCall(SQLQuery)
            If Not rst_User.EOF Then
                If rst_User!urfDelete <> "Y" Then
                    ilShowUser = True
                End If
            End If
        Else
            SQLQuery = "SELECT * FROM Ust Where ustCode = " & rst_Ulf!ulfUstCode
            Set rst_User = gSQLSelectCall(SQLQuery)
            If Not rst_User.EOF Then
                If rst_User!ustState <> 1 Then
                    ilShowUser = True
                End If
            End If
        End If
        If rbcUsers(0).Value Then
            slSignOffDate = Format$(Trim$(rst_Ulf!ulfSignOffDate), sgShowDateForm)
            If DateValue(gAdjYear(slSignOffDate)) <> DateValue("12/31/2069") Then
                ilShowUser = False
            End If
        End If
        If ilShowUser Then
            If llRow >= grdUsersLog.Rows Then
                grdUsersLog.AddItem ""
            End If
            For llCol = SYSTEMINDEX To DURATIONINDEX Step 1
                grdUsersLog.Row = llRow
                grdUsersLog.Col = llCol
                grdUsersLog.CellBackColor = vbWhite 'LIGHTBLUE
                grdUsersLog.CellForeColor = vbBlue
            Next llCol
            ilFound = False
            slSignOffDate = Format$(Trim$(rst_Ulf!ulfSignOffDate), sgShowDateForm)
            If rst_Ulf!ulfSystemType = "T" Then
                If DateValue(gAdjYear(slSignOffDate)) = DateValue("12/31/2069") Then
                    For ilLoop = 0 To UBound(ilActiveUrfCode) - 1 Step 1
                        If ilActiveUrfCode(ilLoop) = rst_Ulf!ulfUrfCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        ilActiveUrfCode(UBound(ilActiveUrfCode)) = rst_Ulf!ulfUrfCode
                        ReDim Preserve ilActiveUrfCode(0 To UBound(ilActiveUrfCode) + 1) As Integer
                    End If
                End If
                ilFound = False
                For ilLoop = 0 To UBound(ilTotalUrfCode) - 1 Step 1
                    If ilTotalUrfCode(ilLoop) = rst_Ulf!ulfUrfCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    ilTotalUrfCode(UBound(ilTotalUrfCode)) = rst_Ulf!ulfUrfCode
                    ReDim Preserve ilTotalUrfCode(0 To UBound(ilTotalUrfCode) + 1) As Integer
                End If
            Else
                If DateValue(gAdjYear(slSignOffDate)) = DateValue("12/31/2069") Then
                    For ilLoop = 0 To UBound(ilActiveUstCode) - 1 Step 1
                        If ilActiveUstCode(ilLoop) = rst_Ulf!ulfUstCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        ilActiveUstCode(UBound(ilActiveUstCode)) = rst_Ulf!ulfUstCode
                        ReDim Preserve ilActiveUstCode(0 To UBound(ilActiveUstCode) + 1) As Integer
                    End If
                End If
                ilFound = False
                For ilLoop = 0 To UBound(ilTotalUstCode) - 1 Step 1
                    If ilTotalUstCode(ilLoop) = rst_Ulf!ulfUstCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    ilTotalUstCode(UBound(ilTotalUstCode)) = rst_Ulf!ulfUstCode
                    ReDim Preserve ilTotalUstCode(0 To UBound(ilTotalUstCode) + 1) As Integer
                End If
            End If
            'System
            grdUsersLog.TextMatrix(llRow, SIGNONUSERINDEX) = "N"
            If rst_Ulf!ulfSystemType = "T" Then
                grdUsersLog.TextMatrix(llRow, SYSTEMINDEX) = "Traffic"
                If Trim$(gDecryptField(rst_User!urfRept)) <> "" Then
                    grdUsersLog.TextMatrix(llRow, USERNAMEINDEX) = Trim$(gDecryptField(rst_User!urfRept))
                Else
                    grdUsersLog.TextMatrix(llRow, USERNAMEINDEX) = Trim$(gDecryptField(rst_User!urfName))
                End If
                grdUsersLog.TextMatrix(llRow, CITYINDEX) = Trim$(rst_User!urfCity)
                grdUsersLog.TextMatrix(llRow, PHONEINDEX) = Trim$(rst_User!urfPhoneNo)
                If rst_User!urfEMailCefCode > 0 Then
                    ilRet = mGetCefComment(rst_User!urfEMailCefCode, slStr)
                    grdUsersLog.TextMatrix(llRow, EMAILFLAGINDEX) = "Yes"
                    grdUsersLog.TextMatrix(llRow, EMAILVALUEINDEX) = slStr
                Else
                    grdUsersLog.TextMatrix(llRow, EMAILFLAGINDEX) = "No"
                    grdUsersLog.TextMatrix(llRow, EMAILVALUEINDEX) = ""
                End If
            Else
                grdUsersLog.TextMatrix(llRow, SYSTEMINDEX) = "Affiliate"
                If Trim$(rst_User!ustReportName) <> "" Then
                    grdUsersLog.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_User!ustReportName)
                Else
                    grdUsersLog.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_User!ustname)
                End If
                grdUsersLog.TextMatrix(llRow, CITYINDEX) = Trim$(rst_User!ustCity)
                grdUsersLog.TextMatrix(llRow, PHONEINDEX) = Trim$(rst_User!ustPhoneNo)
                If rst_User!ustEmailcefcode > 0 Then
                    ilRet = mGetCefComment(rst_User!ustEmailcefcode, slStr)
                    grdUsersLog.TextMatrix(llRow, EMAILFLAGINDEX) = "Yes"
                    grdUsersLog.TextMatrix(llRow, EMAILVALUEINDEX) = slStr
                Else
                    grdUsersLog.TextMatrix(llRow, EMAILFLAGINDEX) = "No"
                    grdUsersLog.TextMatrix(llRow, EMAILVALUEINDEX) = ""
                End If
                If rst_Ulf!ulfUstCode = igUstCode Then
                    grdUsersLog.TextMatrix(llRow, SIGNONUSERINDEX) = "Y"
                End If
            End If
            grdUsersLog.TextMatrix(llRow, ZONEINDEX) = Trim$(rst_Ulf!ulfTimeZone)
            slSignOnDate = Format$(Trim$(rst_Ulf!ulfSignOnDate), sgShowDateForm)
            slSignOffDate = Format$(Trim$(rst_Ulf!ulfSignOffDate), sgShowDateForm)
            slSignOnTime = Format$(Trim$(rst_Ulf!ulfSignOnTime), sgShowTimeWOSecForm)
            slSignOffTime = Format$(Trim$(rst_Ulf!ulfSignOffTime), sgShowTimeWOSecForm)
            grdUsersLog.TextMatrix(llRow, SIGNONINDEX) = slSignOnDate & " " & slSignOnTime
            If DateValue(gAdjYear(slSignOffDate)) <> DateValue("12/31/2069") Then
                grdUsersLog.TextMatrix(llRow, SIGNOFFINDEX) = slSignOffDate & " " & slSignOffTime
            Else
                grdUsersLog.TextMatrix(llRow, SIGNOFFINDEX) = ""
                slSignOffDate = Format$(Now, sgShowDateForm)
                slSignOffTime = Format$(Now, sgShowTimeWOSecForm)
            End If
            
            'Get current activity
            grdUsersLog.TextMatrix(llRow, ACTIVITYINDEX) = ""
            SQLQuery = "SELECT * FROM Uaf_User_Activity"
            SQLQuery = SQLQuery & " WHERE uafUlfCode = " & rst_Ulf!ulfCode
            SQLQuery = SQLQuery & " AND uafStatus = 'I'"
            SQLQuery = SQLQuery & " ORDER BY uafStartDate DESC, uafStartTime Desc"
            Set rst_Uaf = gSQLSelectCall(SQLQuery)
            Do While Not rst_Uaf.EOF
                If UCase(Trim$(rst_Uaf!uafName)) <> UCase("Users Log") Then
                    grdUsersLog.TextMatrix(llRow, ACTIVITYINDEX) = Trim$(rst_Uaf!uafName)
                    Exit Do
                End If
                rst_Uaf.MoveNext
            Loop
            
            'ilNoDays = DateValue(slSignOffDate) - DateValue(slSignOnDate)
            'If Hour(slSignOffTime) >= Hour(slSignOnTime) Then
            '    ilNoHours = Hour(slSignOffTime) - Hour(slSignOnTime)
            'Else
            '    ilNoHours = Hour(slSignOffTime) + (24 - Hour(slSignOnTime))
            '    ilNoDays = ilNoDays - 1
            'End If
            'If Minute(slSignOffTime) >= Minute(slSignOnTime) Then
            '    ilNoMinutes = Minute(slSignOffTime) - Minute(slSignOnTime)
            'Else
            '    ilNoMinutes = Minute(slSignOffTime) + (60 - Minute(slSignOnTime))
            '    ilNoHours = ilNoHours - 1
            'End If
            llMinuteDiff = DateDiff("n", slSignOnDate & " " & slSignOnTime, slSignOffDate & " " & slSignOffTime)
            ilNoDays = llMinuteDiff \ 1440
            llMinuteDiff = llMinuteDiff - CLng(ilNoDays) * 1440
            ilNoHours = llMinuteDiff \ 60
            ilNoMinutes = llMinuteDiff - CLng(ilNoHours) * 60
            grdUsersLog.TextMatrix(llRow, DURATIONINDEX) = Trim$(Str$(ilNoDays)) & " Days " & Trim$(Str$(ilNoHours)) & "H" & Trim$(Str$(ilNoMinutes)) & "M"
            grdUsersLog.TextMatrix(llRow, PCMACINDEX) = Trim$(rst_Ulf!ulfPCName) & " " & Trim$(rst_Ulf!ulfPCMacAddr)
            grdUsersLog.TextMatrix(llRow, ULFCODEINDEX) = rst_Ulf!ulfCode
            llRow = llRow + 1
        End If
        rst_Ulf.MoveNext
    Loop
    If rbcUsers(1).Value Then
        SQLQuery = "SELECT * FROM Ust"
        Set rst_User = gSQLSelectCall(SQLQuery)
        Do While Not rst_User.EOF
            If (rst_User!ustState <> 1) And (rst_User!ustCode > 1) Then
                ilFound = False
                For ilLoop = 0 To UBound(ilTotalUstCode) - 1 Step 1
                    If ilTotalUstCode(ilLoop) = rst_User!ustCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    ilTotalUstCode(UBound(ilTotalUstCode)) = rst_User!ustCode
                    ReDim Preserve ilTotalUstCode(0 To UBound(ilTotalUstCode) + 1) As Integer
                    If llRow >= grdUsersLog.Rows Then
                        grdUsersLog.AddItem ""
                    End If
                    For llCol = SYSTEMINDEX To DURATIONINDEX Step 1
                        grdUsersLog.Row = llRow
                        grdUsersLog.Col = llCol
                        grdUsersLog.CellBackColor = vbWhite 'LIGHTBLUE
                        grdUsersLog.CellForeColor = vbBlue
                    Next llCol
                    grdUsersLog.TextMatrix(llRow, SYSTEMINDEX) = "Affiliate"
                    If Trim$(rst_User!ustReportName) <> "" Then
                        grdUsersLog.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_User!ustReportName)
                    Else
                        grdUsersLog.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_User!ustname)
                    End If
                    grdUsersLog.TextMatrix(llRow, CITYINDEX) = Trim$(rst_User!ustCity)
                    grdUsersLog.TextMatrix(llRow, PHONEINDEX) = Trim$(rst_User!ustPhoneNo)
                    If rst_User!ustEmailcefcode > 0 Then
                        ilRet = mGetCefComment(rst_User!ustEmailcefcode, slStr)
                        grdUsersLog.TextMatrix(llRow, EMAILFLAGINDEX) = "Yes"
                        grdUsersLog.TextMatrix(llRow, EMAILVALUEINDEX) = slStr
                    Else
                        grdUsersLog.TextMatrix(llRow, EMAILFLAGINDEX) = "No"
                        grdUsersLog.TextMatrix(llRow, EMAILVALUEINDEX) = ""
                    End If
                    grdUsersLog.TextMatrix(llRow, ZONEINDEX) = ""
                    grdUsersLog.TextMatrix(llRow, SIGNONINDEX) = ""
                    grdUsersLog.TextMatrix(llRow, ACTIVITYINDEX) = ""
                    grdUsersLog.TextMatrix(llRow, SIGNOFFINDEX) = ""
                    grdUsersLog.TextMatrix(llRow, DURATIONINDEX) = ""
                    grdUsersLog.TextMatrix(llRow, PCMACINDEX) = ""
                    grdUsersLog.TextMatrix(llRow, SIGNONUSERINDEX) = "N"
                    grdUsersLog.TextMatrix(llRow, ULFCODEINDEX) = "0"
                    llRow = llRow + 1
                End If
            End If
            rst_User.MoveNext
        Loop
    End If
    On Error Resume Next
    rst_User.Close
    rst_Uaf.Close
    rst_Ulf.Close
    If rbcUsers(0).Value Then
        lacNoUsers.Caption = "Current Active: " & Trim$(Str$(UBound(ilActiveUrfCode) + UBound(ilActiveUstCode)))
    Else
        lacNoUsers.Caption = "Active: " & Trim$(Str$(UBound(ilActiveUrfCode) + UBound(ilActiveUstCode))) & " Total: " & Trim$(Str$(UBound(ilTotalUrfCode) + UBound(ilTotalUstCode)))
    End If
    mUsersLogSortCol SIGNONINDEX
    mUsersLogSortCol SIGNONINDEX
    mUsersLogSortCol SYSTEMINDEX
    mUsersLogSortCol SIGNOFFINDEX
    mUsersLogSortCol SIGNOFFINDEX
    grdUsersLog.Row = 0
    grdUsersLog.Col = ULFCODEINDEX
    grdUsersLog.Redraw = True
    Exit Sub
ErrHand:
    gSetMousePointer grdUsersLog, grdActivity, vbDefault
    gHandleError "AffErrorLog.txt", "frmUserLog-mPopulate"
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdUsersLog.ColWidth(ULFCODEINDEX) = 0
    grdUsersLog.ColWidth(SIGNONUSERINDEX) = 0
    grdUsersLog.ColWidth(SELECTEDINDEX) = 0
    grdUsersLog.ColWidth(PCMACINDEX) = 0
    grdUsersLog.ColWidth(EMAILVALUEINDEX) = 0
    grdUsersLog.ColWidth(SORTINDEX) = 0
    grdUsersLog.ColWidth(SYSTEMINDEX) = grdUsersLog.Width * 0.07
    grdUsersLog.ColWidth(USERNAMEINDEX) = grdUsersLog.Width * 0.16
    grdUsersLog.ColWidth(CITYINDEX) = grdUsersLog.Width * 0.1
    grdUsersLog.ColWidth(PHONEINDEX) = grdUsersLog.Width * 0.1
    grdUsersLog.ColWidth(EMAILFLAGINDEX) = grdUsersLog.Width * 0.06
    grdUsersLog.ColWidth(ZONEINDEX) = grdUsersLog.Width * 0.03
    grdUsersLog.ColWidth(SIGNONINDEX) = grdUsersLog.Width * 0.09
    grdUsersLog.ColWidth(ACTIVITYINDEX) = grdUsersLog.Width * 0.13
    grdUsersLog.ColWidth(SIGNOFFINDEX) = grdUsersLog.Width * 0.09
    grdUsersLog.ColWidth(DURATIONINDEX) = grdUsersLog.Width * 0.11
    
    grdUsersLog.ColWidth(USERNAMEINDEX) = grdUsersLog.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To DURATIONINDEX Step 1
        If ilCol <> USERNAMEINDEX Then
            grdUsersLog.ColWidth(USERNAMEINDEX) = grdUsersLog.ColWidth(USERNAMEINDEX) - grdUsersLog.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdUsersLog
    
    grdActivity.ColWidth(AUAFCODEINDEX) = 0
    grdActivity.ColWidth(ASTARTDATEINDEX) = grdUsersLog.ColWidth(SIGNONINDEX)
    grdActivity.ColWidth(AENDDATEINDEX) = grdActivity.ColWidth(ASTARTDATEINDEX)
    grdActivity.ColWidth(ASTARTTIMEINDEX) = grdActivity.ColWidth(ASTARTDATEINDEX)
    grdActivity.ColWidth(AENDTIMEINDEX) = grdActivity.ColWidth(ASTARTDATEINDEX)
    grdActivity.ColWidth(ADURATIONINDEX) = grdUsersLog.ColWidth(DURATIONINDEX)
    grdActivity.ColWidth(ASTATUSINDEX) = (2 * grdActivity.ColWidth(ASTARTDATEINDEX) / 3)
    grdActivity.ColWidth(ACSIDATEINDEX) = grdActivity.ColWidth(ASTARTDATEINDEX)
    
    grdActivity.ColWidth(AACTIVITYINDEX) = grdActivity.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To ACSIDATEINDEX Step 1
        If ilCol <> AACTIVITYINDEX Then
            grdActivity.ColWidth(AACTIVITYINDEX) = grdActivity.ColWidth(AACTIVITYINDEX) - grdActivity.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdActivity

End Sub

Private Sub mSetGridTitles()
    Dim llCol As Long
    
    'Set column titles
    grdUsersLog.TextMatrix(0, SYSTEMINDEX) = "System"
    grdUsersLog.TextMatrix(0, USERNAMEINDEX) = "User Name"
    grdUsersLog.TextMatrix(0, CITYINDEX) = "City"
    grdUsersLog.TextMatrix(0, PHONEINDEX) = "Phone #"
    grdUsersLog.TextMatrix(0, EMAILFLAGINDEX) = "E-Mail"
    grdUsersLog.TextMatrix(0, ZONEINDEX) = "Z"
    grdUsersLog.TextMatrix(0, SIGNONINDEX) = "Sign On"
    grdUsersLog.TextMatrix(0, ACTIVITYINDEX) = "Activity"
    grdUsersLog.TextMatrix(0, SIGNOFFINDEX) = "Sign Off"
    grdUsersLog.TextMatrix(0, DURATIONINDEX) = "Duration"

    grdActivity.TextMatrix(0, AACTIVITYINDEX) = "Activity"
    grdActivity.TextMatrix(0, ASTARTDATEINDEX) = "Start Date"
    grdActivity.TextMatrix(0, AENDDATEINDEX) = "End Date"
    grdActivity.TextMatrix(0, ASTARTTIMEINDEX) = "Start Time"
    grdActivity.TextMatrix(0, AENDTIMEINDEX) = "End Time"
    grdActivity.TextMatrix(0, ADURATIONINDEX) = "Duration"
    grdActivity.TextMatrix(0, ASTATUSINDEX) = "Status"
    grdActivity.TextMatrix(0, ACSIDATEINDEX) = "CSI Date"

    grdActivity.Row = 0
    For llCol = AACTIVITYINDEX To ACSIDATEINDEX Step 1
        grdActivity.Col = llCol
        grdActivity.CellBackColor = LIGHTBLUE
    Next llCol

End Sub

Private Sub mUsersLogSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer
    
    For llRow = grdUsersLog.FixedRows To grdUsersLog.Rows - 1 Step 1
        slStr = Trim$(grdUsersLog.TextMatrix(llRow, SYSTEMINDEX))
        If slStr <> "" Then
            If ilCol = SIGNONINDEX Then
                slStr = grdUsersLog.TextMatrix(llRow, SIGNONINDEX)
                If slStr <> "" Then
                    ilPos = InStr(1, slStr, " ", vbTextCompare)
                    If ilPos > 0 Then
                        slDate = Left$(slStr, ilPos - 1)
                        slTime = Mid$(slStr, ilPos + 1)
                    Else
                        slDate = slStr
                    End If
                    If slDate <> "" Then
                        slSort = Trim$(Str$(DateValue(slDate)))
                    Else
                        slSort = "0"
                    End If
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                    If slTime <> "" Then
                        slStr = Trim$(Str$(gTimeToLong(slStr, False)))
                    Else
                        slStr = ""
                    End If
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    slSort = slSort & slStr
                End If
            ElseIf (ilCol = SIGNOFFINDEX) Then
                slStr = grdUsersLog.TextMatrix(llRow, SIGNOFFINDEX)
                If slStr <> "" Then
                    ilPos = InStr(1, slStr, " ", vbTextCompare)
                    If ilPos > 0 Then
                        slDate = Left$(slStr, ilPos - 1)
                        slTime = Mid$(slStr, ilPos + 1)
                    Else
                        slDate = slStr
                    End If
                    If slDate <> "" Then
                        slSort = Trim$(Str$(DateValue(slDate)))
                    Else
                        slSort = "0"
                    End If
                    Do While Len(slSort) < 6
                        slSort = "0" & slSort
                    Loop
                    If slTime <> "" Then
                        slStr = Trim$(Str$(gTimeToLong(slStr, False)))
                    Else
                        slStr = ""
                    End If
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    slSort = slSort & slStr
                Else
                    slSort = "999999999999"
                End If
            ElseIf ilCol = DURATIONINDEX Then
                If grdUsersLog.TextMatrix(llRow, SIGNONINDEX) = "" Then
                    slStr = grdUsersLog.TextMatrix(llRow, DURATIONINDEX)
                    ilPos = InStr(1, slStr, "Days", vbTextCompare)
                    If ilPos > 0 Then
                        slDays = Left$(slStr, ilPos - 1)
                        slStr = Mid$(slStr, ilPos + 3)
                    Else
                        slDays = "0"
                    End If
                    ilPos = InStr(1, slStr, "H", vbTextCompare)
                    If ilPos > 0 Then
                        slHours = Left$(slStr, ilPos - 1)
                        slMinutes = Mid$(slStr, ilPos + 1)
                        ilPos = InStr(1, slMinutes, "M", vbTextCompare)
                        If ilPos > 0 Then
                            slMinutes = Left$(slMinutes, ilPos - 1)
                        End If
                    Else
                        slHours = "0"
                        slMinutes = "0"
                    End If
                    Do While Len(slDays) < 4
                        slDays = "0" & slDays
                    Loop
                    Do While Len(slHours) < 2
                        slHours = "0" & slHours
                    Loop
                    Do While Len(slMinutes) < 2
                        slMinutes = "0" & slMinutes
                    Loop
                    slSort = slDays & slHours & slMinutes
                End If
            Else
                slSort = UCase$(Trim$(grdUsersLog.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdUsersLog.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastUsersLogColSorted) Or ((ilCol = imLastUsersLogColSorted) And (imLastUsersLogSort = flexSortStringNoCaseDescending)) Then
                'Ascending
                If grdUsersLog.TextMatrix(llRow, SIGNONINDEX) = "" Then
                    slSort = "zzzzzzzzzzzz"
                    slStr = grdUsersLog.TextMatrix(llRow, USERNAMEINDEX)
                End If
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdUsersLog.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                'Descending
                If grdUsersLog.TextMatrix(llRow, SIGNONINDEX) = "" Then
                    slSort = "!!!!!!!!!!!!"
                    slStr = grdUsersLog.TextMatrix(llRow, USERNAMEINDEX)
                    For ilChar = 1 To Len(slStr) Step 1
                        If (Asc(Mid$(slStr, ilChar, 1)) >= 48) And (Asc(Mid$(slStr, ilChar, 1)) <= 57) Then
                            Mid$(slStr, ilChar, 1) = Chr$(57 - Asc(Mid$(slStr, ilChar, 1)) + 48)
                        ElseIf (Asc(Mid$(slStr, ilChar, 1)) >= 65) And (Asc(Mid$(slStr, ilChar, 1)) <= 90) Then
                            Mid$(slStr, ilChar, 1) = Chr$(90 - Asc(Mid$(slStr, ilChar, 1)) + 65)
                        ElseIf (Asc(Mid$(slStr, ilChar, 1)) >= 97) And (Asc(Mid$(slStr, ilChar, 1)) <= 122) Then
                            Mid$(slStr, ilChar, 1) = Chr$(122 - Asc(Mid$(slStr, ilChar, 1)) + 97)
                        End If
                    Next ilChar
                End If
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdUsersLog.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastUsersLogColSorted Then
        imLastUsersLogColSorted = SORTINDEX
    Else
        imLastUsersLogColSorted = -1
        imLastUsersLogSort = -1
    End If
    gGrid_SortByCol grdUsersLog, SYSTEMINDEX, SORTINDEX, imLastUsersLogColSorted, imLastUsersLogSort
    imLastUsersLogColSorted = ilCol
End Sub

Private Sub grdUsersLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String
    Dim llHeight As Long
    
    For llRow = 0 To grdUsersLog.FixedRows - 1 Step 1
        llHeight = grdUsersLog.RowHeight(llRow)
    Next llRow
    'grdUsersLog.ToolTipText = ""
    If Y <= llHeight Then
        grdUsersLog.ToolTipText = ""
        Exit Sub
    End If
    slStr = ""
    ilFound = gGrid_GetRowCol(grdUsersLog, X, Y, llRow, llCol)
    If (llCol = EMAILFLAGINDEX) Then
        'grdUsersLog.ToolTipText = Trim$(grdUsersLog.TextMatrix(llRow, EMAILVALUEINDEX))
        slStr = Trim$(grdUsersLog.TextMatrix(llRow, EMAILVALUEINDEX))
    End If
    If (llCol = USERNAMEINDEX) Then
        'grdUsersLog.ToolTipText = Trim$(grdUsersLog.TextMatrix(llRow, PCMACINDEX))
        slStr = Trim$(grdUsersLog.TextMatrix(llRow, PCMACINDEX))
    End If
    If (llCol = SIGNONINDEX) Or (llCol = ACTIVITYINDEX) Or (llCol = SIGNOFFINDEX) Or (llCol = DURATIONINDEX) Then
        'grdUsersLog.ToolTipText = Trim$(grdUsersLog.TextMatrix(llRow, llCol))
        slStr = Trim$(grdUsersLog.TextMatrix(llRow, llCol))
    End If
    grdUsersLog.ToolTipText = slStr
End Sub

Private Sub mActivitySortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer


    For llRow = grdActivity.FixedRows To grdActivity.Rows - 1 Step 1
        slStr = Trim$(grdActivity.TextMatrix(llRow, AACTIVITYINDEX))
        If slStr <> "" Then
            If (ilCol = ASTARTDATEINDEX) Or (ilCol = AENDDATEINDEX) Or (ilCol = ACSIDATEINDEX) Then
                slDate = grdActivity.TextMatrix(llRow, ilCol)
                If slDate = "" Then
                    slDate = "12/31/2069"
                End If
                slSort = Trim$(Str$(gDateValue(slDate)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = ASTARTTIMEINDEX) Or (ilCol = AENDTIMEINDEX) Then
                slTime = grdActivity.TextMatrix(llRow, ilCol)
                If slTime = "" Then
                    slTime = "11:59:59PM"
                End If
                slSort = Trim$(Str$(gTimeToLong(slTime, False)))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf ilCol = ADURATIONINDEX Then
                If grdActivity.TextMatrix(llRow, ADURATIONINDEX) <> "" Then
                    slStr = grdActivity.TextMatrix(llRow, ADURATIONINDEX)
                    ilPos = InStr(1, slStr, "Days", vbTextCompare)
                    If ilPos > 0 Then
                        slDays = Left$(slStr, ilPos - 1)
                        slStr = Mid$(slStr, ilPos + 3)
                    Else
                        slDays = "0"
                    End If
                    ilPos = InStr(1, slStr, "H", vbTextCompare)
                    If ilPos > 0 Then
                        slHours = Left$(slStr, ilPos - 1)
                        slMinutes = Mid$(slStr, ilPos + 1)
                        ilPos = InStr(1, slMinutes, "M", vbTextCompare)
                        If ilPos > 0 Then
                            slMinutes = Left$(slMinutes, ilPos - 1)
                        End If
                    Else
                        slHours = "0"
                        slMinutes = "0"
                    End If
                    Do While Len(slDays) < 4
                        slDays = "0" & slDays
                    Loop
                    Do While Len(slHours) < 2
                        slHours = "0" & slHours
                    Loop
                    Do While Len(slMinutes) < 2
                        slMinutes = "0" & slMinutes
                    Loop
                    slSort = slDays & slHours & slMinutes
                End If
            Else
                slSort = UCase$(Trim$(grdActivity.TextMatrix(llRow, ilCol)))
            End If
            slStr = grdActivity.TextMatrix(llRow, ASORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastActivityColSorted) Or ((ilCol = imLastActivityColSorted) And (imLastActivitySort = flexSortStringNoCaseDescending)) Then
                'Ascending
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdActivity.TextMatrix(llRow, ASORTINDEX) = slSort & slStr & "|" & slRow
            Else
                'Descending
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdActivity.TextMatrix(llRow, ASORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastActivityColSorted Then
        imLastActivityColSorted = ASORTINDEX
    Else
        imLastActivityColSorted = -1
        imLastActivitySort = -1
    End If
    gGrid_SortByCol grdActivity, AACTIVITYINDEX, ASORTINDEX, imLastActivityColSorted, imLastActivitySort
    imLastActivityColSorted = ilCol
End Sub

Private Sub grdUsersLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

    If Y < grdUsersLog.RowHeight(0) Then
        grdUsersLog.Col = grdUsersLog.MouseCol
        mUsersLogSortCol grdUsersLog.Col
        grdUsersLog.Row = 0
        grdUsersLog.Col = ULFCODEINDEX
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdUsersLog, X, Y, llCurrentRow, llCol)
    If llCurrentRow < grdUsersLog.FixedRows Then
        Exit Sub
    End If
    If llCurrentRow >= grdUsersLog.FixedRows Then
        If grdUsersLog.TextMatrix(llCurrentRow, SYSTEMINDEX) <> "" Then
            grdUsersLog.TopRow = lmScrollTop
            llTopRow = grdUsersLog.TopRow
            If (Not ckcActivity.Visible) Or (ckcActivity.Value = vbUnchecked) Then
                If (Shift And CTRLMASK) > 0 Then
                    If grdUsersLog.TextMatrix(grdUsersLog.Row, ULFCODEINDEX) <> "0" Then
                        If grdUsersLog.TextMatrix(grdUsersLog.Row, SELECTEDINDEX) <> 1 Then
                            grdUsersLog.TextMatrix(grdUsersLog.Row, SELECTEDINDEX) = 1
                        Else
                            grdUsersLog.TextMatrix(grdUsersLog.Row, SELECTEDINDEX) = 0
                        End If
                        mPaintRowColor grdUsersLog.Row
                    End If
                Else
                    For llRow = grdUsersLog.FixedRows To grdUsersLog.Rows - 1 Step 1
                        If grdUsersLog.TextMatrix(llRow, SYSTEMINDEX) <> "" Then
                            grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "0"
                            If grdUsersLog.TextMatrix(llRow, ULFCODEINDEX) <> "0" Then
                                If (lmLastClickedRow = -1) Or ((Shift And SHIFTMASK) <= 0) Then
                                    If llRow = llCurrentRow Then
                                        grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                    Else
                                        grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "0"
                                    End If
                                ElseIf lmLastClickedRow < llCurrentRow Then
                                    If (llRow >= lmLastClickedRow) And (llRow <= llCurrentRow) Then
                                        grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                    End If
                                Else
                                    If (llRow >= llCurrentRow) And (llRow <= lmLastClickedRow) Then
                                        grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "1"
                                    End If
                                End If
                                mPaintRowColor llRow
                            End If
                        End If
                    Next llRow
                    grdUsersLog.TopRow = llTopRow
                    grdUsersLog.Row = llCurrentRow
                End If
            Else
                For llRow = grdUsersLog.FixedRows To grdUsersLog.Rows - 1 Step 1
                    If grdUsersLog.TextMatrix(llRow, SYSTEMINDEX) <> "" Then
                        If grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                            grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "0"
                            mPaintRowColor llRow
                        End If
                    End If
                Next llRow
                If lmLastClickedRow <> llCurrentRow Then
                    grdUsersLog.TextMatrix(llCurrentRow, SELECTEDINDEX) = "1"
                    mPaintRowColor llCurrentRow
                    mPopActivityGrid llCurrentRow
                Else
                    mClearActivityGrid
                End If
                grdUsersLog.TopRow = llTopRow
                grdUsersLog.Row = llCurrentRow
                If lmLastClickedRow = llCurrentRow Then
                    llCurrentRow = -1
                End If
            End If
            lmLastClickedRow = llCurrentRow
        End If
    End If
End Sub

Private Sub grdUsersLog_Scroll()
    lmScrollTop = grdUsersLog.TopRow
End Sub

Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long
    
    grdUsersLog.Row = llRow
    For llCol = SYSTEMINDEX To DURATIONINDEX Step 1
        grdUsersLog.Col = llCol
        If grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
            grdUsersLog.CellBackColor = vbWhite
            'grdUsersLog.CellForeColor = vbBlue
        Else
            grdUsersLog.CellBackColor = GRAY    'vbBlue
            'grdUsersLog.CellForeColor = vbWhite
        End If
    Next llCol

End Sub

Private Sub mSetUsers()
    Dim llRow As Long
    ReDim lgUserLogUlfCode(0 To 0) As Long    'Traffic User to receive message or alert
    bmEnableEmail = False
    For llRow = grdUsersLog.FixedRows To grdUsersLog.Rows - 1 Step 1
        If (grdUsersLog.TextMatrix(llRow, SYSTEMINDEX) <> "") And (grdUsersLog.TextMatrix(llRow, ULFCODEINDEX) <> "0") Then
            If grdUsersLog.TextMatrix(llRow, SELECTEDINDEX) = "1" Then
                lgUserLogUlfCode(UBound(lgUserLogUlfCode)) = Val(grdUsersLog.TextMatrix(llRow, ULFCODEINDEX))
                ReDim Preserve lgUserLogUlfCode(0 To UBound(lgUserLogUlfCode) + 1) As Long
                'Dan M 11/06/09 added test of email to enable if messages selected
                If grdUsersLog.TextMatrix(llRow, EMAILFLAGINDEX) = "Yes" Then
                    bmEnableEmail = True
                End If
            End If
        End If
    Next llRow
End Sub


Private Sub mSetCommands()
    If (ckcActivity.Visible = False) Or (ckcActivity.Value = vbUnchecked) Then
        cmcMessage.Enabled = True
        If sgAllowedToBlock <> "Y" Then
            cmcSignOn.Enabled = False
            cmcSignOff.Enabled = False
        Else
            If Not imAllowOkMsg Then
                cmcSignOn.Enabled = False
            Else
                cmcSignOn.Enabled = True
            End If
        End If
    Else
        cmcSignOn.Enabled = False
        cmcSignOff.Enabled = False
        cmcMessage.Enabled = False
    End If
    'If sgEMail = "" Then
    '    cmcMessage.Enabled = False
    'End If
End Sub

Private Sub mSetOkMsg()
    imAllowOkMsg = False
    SQLQuery = "SELECT * FROM AUF_Alert_User WHERE aufType = 'B' AND aufStatus = 'R'"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!aufCreateUstCode = igUstCode Then
            imAllowOkMsg = True
        End If
    End If
    rst.Close
End Sub

Private Sub rbcUsers_Click(Index As Integer)
    If rbcUsers(Index).Value Then
        cmcRefresh_Click
    End If
End Sub

Private Sub mPopActivityGrid(llSelectedRow As Long)
    Dim llUlfCode As Long
    Dim llRow As Long
    Dim slSDate As String
    Dim slSTime As String
    Dim slEDate As String
    Dim slETime As String
    Dim ilPass As Integer
    Dim slStatus As String
    Dim ilRet As Integer
    Dim ilNoDays As Integer
    Dim ilNoHours As Integer
    Dim ilNoMinutes As Integer
    Dim llMinuteDiff As Long
    Dim slCSIDate As String
    
    On Error GoTo ErrHand:
    grdActivity.Redraw = False
    mClearActivityGrid
    llUlfCode = Val(grdUsersLog.TextMatrix(llSelectedRow, ULFCODEINDEX))
    llRow = grdActivity.FixedRows
    For ilPass = 1 To 3 Step 1
        If ilPass = 1 Then
            slStatus = "I"
        ElseIf ilPass = 2 Then
            slStatus = "C"
        ElseIf ilPass = 3 Then
            slStatus = "A"
        End If
        
        SQLQuery = "SELECT * FROM Uaf_User_Activity"
        SQLQuery = SQLQuery & " WHERE uafUlfCode = " & llUlfCode
        SQLQuery = SQLQuery & " AND uafStatus = '" & slStatus & "'"
        SQLQuery = SQLQuery & " ORDER BY uafStartDate DESC, uafStartTime Desc"
        Set rst_Uaf = gSQLSelectCall(SQLQuery)
        Do While Not rst_Uaf.EOF
            If UCase(Trim$(rst_Uaf!uafName)) <> UCase("Users Log") Then
                If llRow >= grdActivity.Rows Then
                    grdActivity.AddItem ""
                End If
                grdActivity.TextMatrix(llRow, AACTIVITYINDEX) = Trim$(rst_Uaf!uafName)
                slSDate = Format(rst_Uaf!uafStartDate, sgShowDateForm)
                grdActivity.TextMatrix(llRow, ASTARTDATEINDEX) = slSDate
                slSTime = Format(rst_Uaf!uafStartTime, sgShowTimeWSecForm)
                grdActivity.TextMatrix(llRow, ASTARTTIMEINDEX) = slSTime
                slEDate = Format(rst_Uaf!uafEndDate, sgShowDateForm)
                If gDateValue(slEDate) = gDateValue("12/31/2069") Then
                    grdActivity.TextMatrix(llRow, AENDDATEINDEX) = ""
                    grdActivity.TextMatrix(llRow, AENDTIMEINDEX) = ""
                    slEDate = Format$(Now, "m/d/yy")
                    slETime = Format$(Now, "h:mm:ssAM/PM")
                Else
                    grdActivity.TextMatrix(llRow, AENDDATEINDEX) = slEDate
                    slETime = Format(rst_Uaf!uafEndTime, sgShowTimeWSecForm)
                    grdActivity.TextMatrix(llRow, AENDTIMEINDEX) = slETime
                End If
                llMinuteDiff = DateDiff("n", slSDate & " " & slSTime, slEDate & " " & slETime)
                ilNoDays = llMinuteDiff \ 1440
                llMinuteDiff = llMinuteDiff - CLng(ilNoDays) * 1440
                ilNoHours = llMinuteDiff \ 60
                ilNoMinutes = llMinuteDiff - CLng(ilNoHours) * 60
                grdActivity.TextMatrix(llRow, ADURATIONINDEX) = Trim$(Str$(ilNoDays)) & " Days " & Trim$(Str$(ilNoHours)) & "H" & Trim$(Str$(ilNoMinutes)) & "M"
                grdActivity.Row = llRow
                grdActivity.Col = ASTATUSINDEX
                If rst_Uaf!uafStatus = "C" Then
                    grdActivity.CellForeColor = DARKGREEN
                    grdActivity.TextMatrix(llRow, ASTATUSINDEX) = "Done"
                ElseIf rst_Uaf!uafStatus = "I" Then
                    grdActivity.TextMatrix(llRow, ASTATUSINDEX) = "Active"
                ElseIf rst_Uaf!uafStatus = "A" Then
                    grdActivity.CellForeColor = vbRed
                    grdActivity.TextMatrix(llRow, ASTATUSINDEX) = "Aborted"
                End If
                slCSIDate = ""
                On Error Resume Next
                slCSIDate = Format(rst_Uaf!uafCSIDate, sgShowDateForm)
                On Error GoTo ErrHand:
                grdActivity.TextMatrix(llRow, ACSIDATEINDEX) = slCSIDate
                grdActivity.TextMatrix(llRow, AUAFCODEINDEX) = rst_Uaf!uafCode
                llRow = llRow + 1
            End If
            rst_Uaf.MoveNext
        Loop
    Next ilPass
    grdActivity.Redraw = True
    Exit Sub
ErrHand:
    gSetMousePointer grdUsersLog, grdActivity, vbDefault
    gHandleError "AffErrorLog.txt", "frmUserLog-mPopActivityGrid"
End Sub
