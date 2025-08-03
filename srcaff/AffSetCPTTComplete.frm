VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmSetCPTTComplete 
   Caption         =   "Set Post CP Complete Flag"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "AffSetCPTTComplete.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9885
   Begin V81SetCPTTComplete.CSI_Calendar edcStartDate 
      Height          =   255
      Left            =   1605
      TabIndex        =   1
      Top             =   60
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   450
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
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
   Begin VB.CheckBox ckcNotAired 
      Caption         =   "Not Aired"
      Height          =   240
      Left            =   3180
      TabIndex        =   16
      Top             =   585
      Width           =   1065
   End
   Begin VB.CheckBox ckcPartial 
      Caption         =   "Partial"
      Height          =   240
      Left            =   2220
      TabIndex        =   15
      Top             =   600
      Width           =   870
   End
   Begin V81SetCPTTComplete.CSI_Calendar edcEndDate 
      Height          =   255
      Left            =   5340
      TabIndex        =   3
      Top             =   60
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   450
      Text            =   "9/22/2012"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BorderStyle     =   1
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
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   975
      Width           =   3825
   End
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5235
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Results"
      Top             =   975
      Width           =   3825
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9675
      Top             =   3300
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   4995
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3180
      ItemData        =   "AffSetCPTTComplete.frx":08CA
      Left            =   5070
      List            =   "AffSetCPTTComplete.frx":08CC
      TabIndex        =   9
      Top             =   1425
      Width           =   4455
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   3180
      ItemData        =   "AffSetCPTTComplete.frx":08CE
      Left            =   135
      List            =   "AffSetCPTTComplete.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9480
      Top             =   4740
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5970
      FormDesignWidth =   9885
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Fix"
      Height          =   375
      Left            =   5820
      TabIndex        =   7
      Top             =   5505
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7860
      TabIndex        =   8
      Top             =   5490
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Fix CPTT Discrepancy:"
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   630
      Width           =   2010
   End
   Begin VB.Label lacVehicle 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1185
      TabIndex        =   13
      Top             =   4755
      Width           =   6810
   End
   Begin VB.Label lacProgress 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1185
      TabIndex        =   12
      Top             =   5085
      Width           =   6810
   End
   Begin VB.Label lacWeekEndDate 
      Caption         =   "Week End Date"
      Height          =   255
      Left            =   3885
      TabIndex        =   2
      Top             =   75
      Width           =   1920
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   10
      Top             =   5445
      Width           =   5490
   End
   Begin VB.Label lacWeekStartDate 
      Caption         =   "Week Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1920
   End
End
Attribute VB_Name = "frmSetCPTTComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imGuideUstCode As Integer
Private bmNoPervasive As Boolean

Private smStartDate As String     'Export Date
Private smEndDate As String
Private imVefCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imChecking As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmMsg As Integer
Private hmTo As Integer
Private lmTotal As Long
Private lmRunningCount As Long
Private lmNumberChgd As Long
Private lmNumberPartials As Long
Private lmPart2Comp As Long
Private hmVehicles As Integer
Private bmErrorMsgLogged As Boolean
Private Const FORMNAME As String = "frmSetCPTTComplete"
Private cptt_rst As ADODB.Recordset
Private ast_rst As ADODB.Recordset
Private att_rst As ADODB.Recordset
Private webl_rst As ADODB.Recordset
'6777
Private lmNumberDidNotAirWrong As Long
Private imDiscrepancySearch As DiscrepancySearch
Private Enum DiscrepancySearch
    All = 2
    NotAired = 1
    Partial = 0
End Enum


'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile(sMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    'On Error GoTo mOpenMsgFileErr:
   ' slToFile = sgMsgDirectory & "SetCPTTComplete_" & sgClientName & "_" & Format$(Now, "mmddyy") & ".Csv"
        '6777
    If imDiscrepancySearch = Partial Then
        slToFile = sgMsgDirectory & "SetCPTTComplete_" & sgClientName & "_Partial_" & Format$(Now, "mmddyy") & ".Csv"
    ElseIf imDiscrepancySearch = NotAired Then
        slToFile = sgMsgDirectory & "SetCPTTComplete_" & sgClientName & "_Not Aired_" & Format$(Now, "mmddyy") & ".Csv"
    Else
        slToFile = sgMsgDirectory & "SetCPTTComplete_" & sgClientName & "_" & Format$(Now, "mmddyy") & ".Csv"
    End If
    slNowDate = Format$(gNow(), sgShowDateForm)
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, sgShowDateForm)
        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
        '    On Error GoTo 0
        '    ilRet = 0
        '    On Error GoTo mOpenMsgFileErr:
        '    hmMsg = FreeFile
        '    Open slToFile For Append As hmMsg
        '    If ilRet <> 0 Then
        '        Close hmMsg
        '        hmMsg = -1
        '        gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
        '        mOpenMsgFile = False
        '        Exit Function
        '    End If
        'Else
            On Error Resume Next
            'Kill slToFile
            'On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, " "
    Print #hmMsg, "** Setting Post CP Complete Flag: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, "Start Date: " & smStartDate & " End Date: " & smEndDate
    '6777
    If imDiscrepancySearch = Partial Then
        Print #hmMsg, "Discrepancy Fix: Partial Only"
    ElseIf imDiscrepancySearch = NotAired Then
        Print #hmMsg, "Discrepancy Search: Not Aired Only"
    Else
        Print #hmMsg, "Discrepancy Search: Partial and Not Aired"
    End If
    sMsgFileName = slToFile
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub mFillVehicle()
    Dim iLoop As Integer
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    mReadPreselectedVehicles
End Sub




Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If

End Sub

Private Sub cmdCheck_Click()
    Dim ilLoop As Integer
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    Dim sFileName As String
    Dim sLetter As String
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim ilRet As Integer
    Dim ilTotal As Integer
    Dim ilCount As Integer

    On Error GoTo ErrHand
    If imChecking = True Then
        Exit Sub
    End If
    imTerminate = False
    lacProgress.Caption = ""
    lacResult.Caption = ""
    lbcMsg.Clear
    If lbcVehicles.SelCount <= 0 Then
        gMsgBox "Vehicle must be specified.", vbOKOnly
        Exit Sub
    End If
    If edcStartDate.Text = "" Then
        gMsgBox "Start Date must be specified.", vbOKOnly
        edcStartDate.SetFocus
        Exit Sub
    End If
    If gIsDate(edcStartDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid start date (m/d/yy).", vbCritical
        edcStartDate.SetFocus
    Else
        smStartDate = Format(edcStartDate.Text, sgShowDateForm)
    End If
    smStartDate = gObtainPrevMonday(smStartDate)
    If (edcEndDate.Text <> "") And (edcEndDate.Text <> "TFN") Then
        If gIsDate(edcEndDate.Text) = False Then
            Beep
            gMsgBox "Please enter a valid end date (m/d/yy).", vbCritical
            edcEndDate.SetFocus
        Else
            smEndDate = Format(edcEndDate.Text, sgShowDateForm)
        End If
        smEndDate = gObtainNextSunday(smEndDate)
    Else
        smEndDate = ""
    End If
    '6777
    If Not (ckcPartial.Value = vbChecked Or ckcNotAired.Value = vbChecked) Then
        Beep
        gMsgBox "Please enter a valid CPTT Discrepancy to check.", vbCritical
        ckcPartial.SetFocus
        Exit Sub
    ElseIf ckcPartial.Value = vbChecked And ckcNotAired.Value = vbChecked Then
        imDiscrepancySearch = DiscrepancySearch.All
    ElseIf ckcPartial.Value = vbChecked Then
        imDiscrepancySearch = Partial
    Else
        imDiscrepancySearch = NotAired
    End If
    If smEndDate = "" Then
        smEndDate = gObtainNextSunday(Format(gNow(), "m/d/yy"))
    End If
    Screen.MousePointer = vbHourglass
    bmErrorMsgLogged = False
    If Not mOpenMsgFile(sMsgFileName) Then
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If

    lacProgress.Caption = "Determine Number Records to Check"
    lacVehicle.Caption = ""
    DoEvents
    
    imChecking = True
    lacResult.Caption = ""
    lmTotal = 0
    lmRunningCount = 0
    lmNumberChgd = 0
    lmNumberPartials = 0
    lmPart2Comp = 0
    '6777
    lmNumberDidNotAirWrong = 0
    For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilLoop) Then
            SQLQuery = "Select Count(*) FROM cptt WHERE cpttVefCode = " & lbcVehicles.ItemData(ilLoop)
            'dan 6777 "all" leaves it commented out
            'SQLQuery = SQLQuery & " AND cpttPostingStatus <> 2"
            If imDiscrepancySearch = Partial Then
                SQLQuery = SQLQuery & " AND cpttPostingStatus <> 2"
            ElseIf imDiscrepancySearch = NotAired Then
                SQLQuery = SQLQuery & " AND cpttPostingStatus = 2"
            End If
            SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(smStartDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(smEndDate, sgSQLDateForm) & "')"
            Set cptt_rst = cnn.Execute(SQLQuery)
            If Not cptt_rst.EOF Then
                lmTotal = lmTotal + cptt_rst(0).Value
            End If
        End If
    Next ilLoop
    lacProgress.Caption = ""
    '6777
   ' Print #hmMsg, "Type,attCode,Feed Date,Vehicle,Station"
    Print #hmMsg, "Station,Vehicle,attCode,Week Date,Type"
    For ilLoop = 0 To lbcVehicles.ListCount - 1
        DoEvents
        If lbcVehicles.Selected(ilLoop) Then
            'Get hmTo handle
            lacVehicle.Caption = "Checking: " & Trim$(lbcVehicles.List(ilLoop))
            DoEvents
            imVefCode = lbcVehicles.ItemData(ilLoop)
            cmdCancel.Caption = "&Cancel"
            ilRet = mUpdateCptt(imVefCode)
            If imTerminate Then
                Exit For
            End If
        End If
    Next ilLoop
    lacProgress.Caption = ""
    lacVehicle.Caption = ""
    If Not imTerminate Then
        mWritePreselectedVehicles
    End If
    imChecking = False
    If Not imTerminate Then
        Print #hmMsg, "** Completed Checking: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Else
        Print #hmMsg, "** Terminated Checking: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    End If
    Print #hmMsg, " "
    Print #hmMsg, lmRunningCount & " records checked"
    lbcMsg.AddItem lmRunningCount & " records checked"
    Print #hmMsg, lmNumberChgd & " records changed"
    lbcMsg.AddItem lmNumberChgd & " records changed"
    '6777 wrapped with if
    If imDiscrepancySearch <> NotAired Then
        Print #hmMsg, lmNumberPartials & " partially Posted records"
        lbcMsg.AddItem lmNumberPartials & " partially Posted records"
        Print #hmMsg, lmPart2Comp & " Changed from Partial to Complete"
        lbcMsg.AddItem lmPart2Comp & " Changed from Partial to Complete"
    End If
    '6777
    If imDiscrepancySearch <> Partial Then
        Print #hmMsg, lmNumberDidNotAirWrong & " Changed from Did Not Air to Complete"
        lbcMsg.AddItem lmNumberDidNotAirWrong & " Changed from Did Not Air to Complete"
    End If
    Print #hmMsg, "** End Setting Post CP Complete Flag: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Close #hmMsg
    lacResult.Caption = "See: " & sMsgFileName & " for Result Summary"
    cmdCheck.Enabled = False    'True
    cmdCancel.SetFocus
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmdCheckErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SetCPTTComplete: mCheck_Click"
End Sub

Private Sub cmdCancel_Click()
    If imChecking Then
        imTerminate = True
        Exit Sub
    End If
    edcStartDate.Text = ""
    Unload frmSetCPTTComplete
End Sub


Private Sub edcEndDate_Change()
    lbcMsg.Clear
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Fix"
    cmdCancel.Caption = "&Cancel"
End Sub

Private Sub edcEndDate_GotFocus()
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Fix"
    cmdCancel.Caption = "&Cancel"
End Sub

Private Sub edcStartDate_GotFocus()
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Fix"
    cmdCancel.Caption = "&Cancel"
End Sub

Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    If imFirstTime Then
        imFirstTime = False
    End If
End Sub

Private Sub Form_GotFocus()
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Fix"
    cmdCancel.Caption = "&Cancel"
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmSetCPTTComplete
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    Dim slAffToWebDate As String
    Dim slWebToAffDate As String
        
    Screen.MousePointer = vbHourglass
    smStartDate = ""
    imAllClick = False
    imTerminate = False
    imChecking = False
    imFirstTime = True
    
    mInit
    
    mFillVehicle
    chkAll.Value = vbChecked
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If imChecking Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    
    Erase tgCifCpfInfo1
    Erase tgCrfInfo1
    Erase lgUserLogUlfCode
    Erase tgCopyRotInfo
    Erase tgGameInfo
    Erase tgStationInfoByCode
    Erase tgCpfInfo
    Erase tgMarketInfo
    Erase tgMSAMarketInfo
    Erase tgTerritoryInfo
    Erase tgCityInfo
    Erase tgCountyInfo
    Erase tgAreaInfo
    Erase tgMonikerInfo
    Erase tgOperatorInfo
    Erase tgMarketRepInfo
    Erase tgServiceRepInfo
    Erase tgAffAEInfo
    Erase tgSellingVehicleInfo
    Erase tgVpfOptions
    Erase tgLstInfo
    Erase tgAttInfo1
    Erase tgShttInfo1
    Erase tgCpttInfo
    Erase sgAufsKey
    Erase tgRBofRec
    Erase tgSplitNetLastFill
    Erase tgAvailNamesInfo
    Erase tgMediaCodesInfo
    Erase tgTitleInfo
    Erase tgOwnerInfo
    Erase tgFormatInfo
    Erase tgVffInfo
    Erase tgTeamInfo
    Erase tgLangInfo
    Erase tgTimeZoneInfo
    Erase tgStateInfo
    Erase tgSubtotalGroupInfo
    Erase tgAttExpMon
    Erase tgReportNames
    Erase tgRff
    Erase tgRffExtended
    Erase tgUstInfo
    Erase tgDeptInfo
    
    
    Erase tgStationInfo
    Erase tgVehicleInfo
    Erase tgRnfInfo
    Erase tgAdvtInfo
    Erase sgStationImportTitles
    
    '9/11/06: Split Network stuff
    Erase tgRBofRec
    Erase tgSplitNetLastFill
    
    On Error Resume Next
    rstAlertUlf.Close
    On Error Resume Next
    rstAlert.Close
    'gLogMsg "Closing Pervasive API Engine. User: " & gGetComputerName(), "WebExportLog.Txt", False
    mClosePervasiveAPI
    cptt_rst.Close
    ast_rst.Close
    att_rst.Close
    webl_rst.Close
    cnn.Close
    
    Set frmSetCPTTComplete = Nothing
End Sub


Private Sub lbcVehicles_Click()
    lbcMsg.Clear
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Fix"
    cmdCancel.Caption = "&Cancel"
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0
        imAllClick = False
    End If
    
    
End Sub

Private Sub edcStartDate_Change()
    lbcMsg.Clear
    cmdCheck.Enabled = True
    cmdCheck.Caption = "&Fix"
    cmdCancel.Caption = "&Cancel"
End Sub




Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmSetCPTTComplete
End Sub

Private Sub mReadPreselectedVehicles()
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim llRow As Long
    
    ReDim smPreselectNames(0 To 0)
    ilRet = 0
    On Error GoTo mReadPreselectedVehicles:
    'hmVehicles = FreeFile
    'Open Trim$(sgImportDirectory) & "SetCPTTComplete.Txt" For Input Access Read As hmVehicles
    ilRet = gFileOpen(Trim$(sgImportDirectory) & "SetCPTTComplete.Txt", "Input Access Read", hmVehicles)
    If ilRet <> 0 Then
        Exit Sub
    End If
    Do
        'On Error GoTo mReadPreselectedVehicles:
        If EOF(hmVehicles) Then
            Exit Do
        End If
        Line Input #hmVehicles, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                For llRow = 0 To lbcVehicles.ListCount - 1 Step 1
                    If StrComp(lbcVehicles.List(llRow), slLine, vbTextCompare) = 0 Then
                        lbcVehicles.Selected(llRow) = True
                        Exit For
                    End If
                Next llRow
            End If
        End If
    Loop Until ilEof
    Close hmVehicles
    Exit Sub
mReadPreselectedVehicles:
    ilRet = Err.Number
    Resume Next
End Sub
Private Sub mWritePreselectedVehicles()
    Dim ilRet As Integer
    Dim llRow As Long
    
    On Error Resume Next
    Kill Trim$(sgImportDirectory) & "SetCPTTComplete.Txt"
    On Error GoTo 0
    
    ilRet = 0
    On Error GoTo mWritePreselectedVehiclesErr:
    'hmVehicles = FreeFile
    'Open Trim$(sgImportDirectory) & "SetCPTTComplete.Txt" For Output As hmVehicles
    ilRet = gFileOpen(Trim$(sgImportDirectory) & "SetCPTTComplete.Txt", "Output", hmVehicles)
    If ilRet = 0 Then
        For llRow = 0 To lbcVehicles.ListCount - 1 Step 1
            If lbcVehicles.Selected(llRow) Then
                Print #hmVehicles, lbcVehicles.List(llRow)
            End If
        Next llRow
        Close #hmVehicles
    End If
    Exit Sub
mWritePreselectedVehiclesErr:
    ilRet = Err.Number
    Resume Next
End Sub






Private Sub mInit()
    Dim sBuffer As String
    Dim lSize As Long
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim ilValue As Integer
    Dim ilValue8 As Integer
    Dim slDate As String
    Dim ilDatabase As Integer
    Dim ilLocation As Integer
    Dim ilSQL As Integer
    Dim ilForm As Integer
    Dim sMsg As String
    Dim iLoop As Integer
    Dim sCurDate As String
    Dim sAutoLogin As String
    Dim slTimeOut As String
    Dim slDSN As String
    Dim slStartIn As String
    Dim slStartStdMo As String
    Dim slTemp As String
    ReDim sWin(0 To 13) As String * 1
    Dim ilIsTntEmpty As Integer
    Dim ilIsShttEmpty As Integer
    Dim slDateTime1 As String
    Dim slDateTime2 As String
    Dim EmailExists_rst As ADODB.Recordset
    '5/11/11
    Dim blAddGuide As Boolean
    'dan 2/23/12 can't have error handler in error handler
    Dim blNeedToCloseCnn As Boolean
    Dim slXMLINIInputFile As String
    
    
    sgCommand = Command$
    blNeedToCloseCnn = False
    'Display gMsgBox
    'igShowMsgBox = True shows the gMsgBox.
    'igShowMsgBox = False does not show any gMsgBox
    
    'Warning: One thing to remember is that if you are expecting a return value from a gMsgBox
    'and you turn gMsgBox off then you need to make sure that you handle that case.
    'example:   ilRet = gMsgBox "xxxx"
    igShowMsgBox = True
    
    'To avoid web check
    'igDemoMode = False
    'If InStr(sgCommand, "Demo") Then
        igDemoMode = True
    'End If
    
    'Used to speed-up testing exports with multiple files reduce record count needed to create a new file
    igSmallFiles = False
    If InStr(sgCommand, "SmallFiles") Then
        igSmallFiles = True
    End If
    
    igAutoImport = False
    slStartIn = CurDir$
    sgCurDir = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommand, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
        
    sgBS = Chr$(8)  'Backspace
    sgTB = Chr$(9)  'Tab
    sgLF = Chr$(10) 'Line Feed (New Line)
    sgCR = Chr$(13) 'Carriage Return
    sgCRLF = sgCR + sgLF
   
   
    ilRet = 0
    ilLocation = False
    ilDatabase = False
    sgDatabaseName = ""
    sgReportDirectory = ""
    sgExportDirectory = ""
    sgImportDirectory = ""
    sgExeDirectory = ""
    sgLogoDirectory = ""
    sgPasswordAddition = ""
    sgSQLDateForm = "yyyy-mm-dd"
    sgCrystalDateForm = "yyyy,mm,dd"
    sgSQLTimeForm = "hh:mm:ss"
    igSQLSpec = 1               'Pervasive 2000
    sgShowDateForm = "m/d/yyyy"
    sgShowTimeWOSecForm = "h:mma/p"
    sgShowTimeWSecForm = "h:mm:ssa/p"
    igWaitCount = 10
    igTimeOut = -1
    sgWallpaper = ""
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
    sgLogoName = "rptlogo.bmp"
    sgNowDate = ""
    ilPos = InStr(1, sgCommand, "/D:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommand, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommand, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommand, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gIsDate(slDate) Then
            sgNowDate = slDate
        End If
    End If
    
    If Not gLoadOption("Locations", "Logo", sgLogoPath) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Unload frmSetCPTTComplete
        Exit Sub
    Else
        sgLogoPath = gSetPathEndSlash(sgLogoPath, True)
    End If
    
    
    If Not gLoadOption("Database", "Name", sgDatabaseName) Then
        gMsgBox "Affiliat.Ini [Database] 'Name' key is missing.", vbCritical
        Unload frmSetCPTTComplete
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Reports", sgReportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Reports' key is missing.", vbCritical
        Unload frmSetCPTTComplete
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Export", sgExportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Export' key is missing.", vbCritical
        Unload frmSetCPTTComplete
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Exe", sgExeDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Exe' key is missing.", vbCritical
        Unload frmSetCPTTComplete
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Logo", sgLogoDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Unload frmSetCPTTComplete
        Exit Sub
    End If
    
        
    'Import is optional
    If gLoadOption("Locations", "Import", sgImportDirectory) Then
        sgImportDirectory = gSetPathEndSlash(sgImportDirectory, True)
    Else
        sgImportDirectory = ""
    End If
    
    If gLoadOption("Locations", "ContractPDF", sgContractPDFPath) Then
        sgContractPDFPath = gSetPathEndSlash(sgContractPDFPath, True)
    Else
        sgContractPDFPath = ""
    End If
    
    
    'Commented out below because I can't see why you would need a backslash
    'on the end of a DSN name
    'sgDatabaseName = gSetPathEndSlash(sgDatabaseName)
    sgReportDirectory = gSetPathEndSlash(sgReportDirectory, True)
    sgExportDirectory = gSetPathEndSlash(sgExportDirectory, True)
    sgExeDirectory = gSetPathEndSlash(sgExeDirectory, True)
    sgLogoDirectory = gSetPathEndSlash(sgLogoDirectory, True)
    
    Call gLoadOption("SQLSpec", "Date", sgSQLDateForm)
    Call gLoadOption("SQLSpec", "Time", sgSQLTimeForm)
    If gLoadOption("SQLSpec", "System", sBuffer) Then
        If sBuffer = "P7" Then
            igSQLSpec = 0
        End If
    End If
    If gLoadOption("Locations", "TimeOut", slTimeOut) Then
        igTimeOut = Val(slTimeOut)
    End If
    Call gLoadOption("Locations", "Wallpaper", sgWallpaper)
    
    Call gLoadOption("Showform", "Date", sgShowDateForm)
    Call gLoadOption("Showform", "TimeWSec", sgShowTimeWSecForm)
    Call gLoadOption("Showform", "TimeWOSec", sgShowTimeWOSecForm)
    
    If Not gLoadOption("Locations", "DBPath", sgDBPath) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Unload frmSetCPTTComplete
        Exit Sub
    Else
        sgDBPath = gSetPathEndSlash(sgDBPath, True)
    End If
    
    'Set Message folder
    If Not gLoadOption("Locations", "DBPath", sgMsgDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Unload frmSetCPTTComplete
        Exit Sub
    Else
        sgMsgDirectory = gSetPathEndSlash(sgMsgDirectory, True) & "Messages\"
'        sgMsgDirectory = CurDir
'        If InStr(1, sgMsgDirectory, "Data", vbTextCompare) Then
'            sgMsgDirectory = gSetPathEndSlash(sgMsgDirectory) & "Messages\"
'        Else
'            sgMsgDirectory = sgExportDirectory
'        End If
    End If
    
    ' Not sure what section this next item is coming from. The original code did not specify.
    'Call gLoadOption("SQLSpec", "WaitCount", sBuffer)
    'igWaitCount = Val(sBuffer)
    
    On Error GoTo ErrHand
    Set cnn = New ADODB.Connection
   
    'Set env = rdoEnvironments(0)
    'cnn.CursorDriver = rdUseOdbc
    
    'Set cnn = cnn.OpenConnection(dsName:="Affiliate", Prompt:=rdDriverCompleteRequired)
    ' The default timeout is 15 seconds. This always fails on my PC the first time I run this program.


    slDSN = sgDatabaseName
    'ttp 4905.  Need to try connection. If it fails, try one more time, after sleeping.
    'cnn.Open "DSN=" & slDSN
    
    On Error GoTo ERRNOPERVASIVE
    ilRet = 0
    cnn.Open "DSN=" & slDSN
    
    On Error GoTo ErrHand
    If ilRet = 1 Then
        Sleep 2000
        cnn.Open "DSN=" & slDSN
    End If

    
    
    'Example of using a user name and password
    'cnn.Open "DSN=" & slDSN, "Master", "doug"
    Set rst = New ADODB.Recordset

    If igTimeOut >= 0 Then
        cnn.CommandTimeout = igTimeOut
    End If
 
    ' The sgDatabaseName may contain an ending backslash. Although this does not seem to have
    ' any effect, it does not seem like a good practice to let it stay like this here incase a later version of the RDO doesn't like it.
    If Mid(slDSN, Len(slDSN), 1) = "\" Then
        ' Yes it did end with a slash. Remove it.
        slDSN = Left(slDSN, Len(slDSN) - 1)
    End If
    'Set cnn = cnn.OpenConnection(dsName:=slDSN, Prompt:=rdDriverCompleteRequired)
    'If igTimeOut >= 0 Then
    '    cnn.QueryTimeout = igTimeOut
    'End If
    'Code modified for testing
    
    
    'Test for Guide- if not added- add
    'SQLQuery = "Select MAX(ustCode) from ust"
    'Set rst = cnn.Execute(SQLQuery)
    ''If rst(0).Value = 0 Then
    'If IsNull(rst(0).Value) Then
    ''5/11/11
    '    blAddGuide = True
    'Else
        SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
        Set rst = cnn.Execute(SQLQuery)
        If rst.EOF Then
            blAddGuide = True
        Else
            blAddGuide = False
            imGuideUstCode = rst!ustCode
        End If
    'End If
    If blAddGuide Then
    '5/11/11
        'SQLQuery = "INSERT INTO ust(ustName, ustPassword, ustState)"
        'SQLQuery = SQLQuery & "VALUES ('Guide', 'Guide', 0)"
        sCurDate = Format(Now, sgShowDateForm)
        For iLoop = 0 To 13 Step 1
            sWin(iLoop) = "I"
        Next iLoop
        '5/11/11
        'mResetGuideGlobals
        SQLQuery = "INSERT INTO ust(ustName, ustReportName, ustPassword, "
        SQLQuery = SQLQuery & "ustState, ustPassDate, ustActivityLog, ustWin1, "
        SQLQuery = SQLQuery & "ustWin2, ustWin3, ustWin4, "
        SQLQuery = SQLQuery & "ustWin5, ustWin6, ustWin7, "
        SQLQuery = SQLQuery & "ustWin8, ustWin9, ustPledge, "
        SQLQuery = SQLQuery & "ustExptSpotAlert, ustExptISCIAlert, ustTrafLogAlert, "
        SQLQuery = SQLQuery & "ustWin10, ustWin11, ustWin12, ustWin13, "
        SQLQuery = SQLQuery & "ustWin14, ustWin15, ustPhoneNo, ustCity, ustEMailCefCode, ustAllowedToBlock, "
        SQLQuery = SQLQuery & "ustWin16, "
        SQLQuery = SQLQuery & "ustUserInitials, "
        SQLQuery = SQLQuery & "ustDntCode, "
        SQLQuery = SQLQuery & "ustAllowCmmtChg, "
        SQLQuery = SQLQuery & "ustAllowCmmtDelete, "
        SQLQuery = SQLQuery & "ustUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "VALUES ('" & "Guide" & "', "
        SQLQuery = SQLQuery & "'" & "System" & "', '" & "Guide" & "', "
        SQLQuery = SQLQuery & 0 & ", '" & Format$(sCurDate, sgSQLDateForm) & "', '" & "V" & "', '" & sgUstWin(1) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(2) & "', '" & sgUstWin(3) & "', '" & sgUstWin(4) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(5) & "', '" & sgUstWin(6) & "', '" & sgUstWin(7) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(8) & "', '" & sgUstWin(9) & "', '" & sgUstPledge & "', "
        SQLQuery = SQLQuery & "'" & sgExptSpotAlert & "', '" & sgExptISCIAlert & "', '" & sgTrafLogAlert & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(10) & "', '" & sgUstWin(11) & "', '" & sgUstWin(12) & "', '" & sgUstWin(13) & "', "
        SQLQuery = SQLQuery & "'" & sgUstClear & "', '" & sgUstDelete & "', "
        SQLQuery = SQLQuery & "'" & "" & "', '" & "" & "', " & 0 & ", '" & "Y" & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(sgUstWin(0)) & "', "
        SQLQuery = SQLQuery & "'" & "G" & "', "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & sgUstAllowCmmtChg & "', "
        SQLQuery = SQLQuery & "'" & sgUstAllowCmmtDelete & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
        cnn.BeginTrans
        blNeedToCloseCnn = True
        'cnn.ConnectionTimeout = 30  ' Increase from the default of 15 to 30 seconds.
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHand:
        End If
        cnn.CommitTrans
        blNeedToCloseCnn = False
        SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
        Set rst = cnn.Execute(SQLQuery)
        If Not rst.EOF Then
            imGuideUstCode = rst!ustCode
        Else
            imGuideUstCode = 0
        End If
    End If
    
    gUsingCSIBackup = False
    gUsingXDigital = False
    gWegenerExport = False
    gOLAExport = False
    ' Dan M added spfusingFeatures2
    SQLQuery = "SELECT spfGClient, spfGAlertInterval, spfGUseAffSys, spfUsingFeatures7, spfUsingFeatures2, spfUsingFeatures8"
    SQLQuery = SQLQuery + " FROM SPF_Site_Options"
    Set rst = cnn.Execute(SQLQuery)
    
    If Not rst.EOF Then
        If UCase(rst!spfGUseAffSys) <> "Y" Then
            gMsgBox "The Affiliate system has not been activated.  Please call Counterpoint.", vbCritical
            Unload frmSetCPTTComplete
            Exit Sub
        End If
        ilValue8 = Asc(rst!spfUsingFeatures8)
        If (ilValue8 And ALLOWMSASPLITCOPY) <> ALLOWMSASPLITCOPY Then
            gUsingMSARegions = False
        Else
            gUsingMSARegions = True
        End If
        If (ilValue8 And ISCIEXPORT) <> ISCIEXPORT Then
            gISCIExport = False
        Else
            gISCIExport = True
        End If
        ilValue = Asc(rst!spfUsingFeatures7)
        If (ilValue And CSIBACKUP) <> CSIBACKUP Then
            gUsingCSIBackup = False
        Else
            gUsingCSIBackup = True
        End If
        
        If ((ilValue And XDIGITALISCIEXPORT) <> XDIGITALISCIEXPORT) And ((ilValue8 And XDIGITALBREAKEXPORT) <> XDIGITALBREAKEXPORT) Then
            gUsingXDigital = False
        Else
            gUsingXDigital = True
        End If
        If (ilValue And WEGENEREXPORT) <> WEGENEREXPORT Then
            gWegenerExport = False
        Else
            gWegenerExport = True
        End If
        If (ilValue And OLAEXPORT) <> OLAEXPORT Then
            gOLAExport = False
        Else
            gOLAExport = True
        End If
        ilValue = Asc(rst!spfusingfeatures2)
        If (ilValue And STRONGPASSWORD) <> STRONGPASSWORD Then
            bgStrongPassword = False
        Else
            bgStrongPassword = True
        End If
    End If
    
    If Not rst.EOF Then
        sgClientName = Trim$(rst!spfGClient)
        igAlertInterval = rst!spfGAlertInterval
    Else
        sgClientName = "Unknown"
        gMsgBox "Client name is not defined in Site Options"
        igAlertInterval = 0
    End If
    
    If InStr(1, sgCommand, "NoAlerts", vbTextCompare) > 0 Then
        'For Debug ONLY
        igAlertInterval = 0
    End If
    
    If Trim$(sgNowDate) = "" Then
        If InStr(1, sgClientName, "XYZ Broadcasting", vbTextCompare) > 0 Then
            sgNowDate = "12/15/1999"
        End If
    End If


    ilRet = gInitGlobals()
    If ilRet = 0 Then
        'While Not gVerifyWebIniSettings()
        '    frmWebIniOptions.Show vbModal
        '    If Not igWebIniOptionsOK Then
        '        Unload frmSetCPTTComplete
        '        Exit Sub
        '    End If
        'Wend
    End If
    
    Call gLoadOption("Database", "AutoLogin", sAutoLogin)
    
    
    On Error GoTo ErrHand
    'If Not igAutoImport Then
    '    ilRet = mInitAPIReport()      '4-19-04
    'End If
    
    
    ilRet = gTestWebVersion()
    'Move report logo to local C drice (c:\csi\rptlogo.bmp)
    ilRet = 0
    On Error GoTo mStartUpErr:
    'slDateTime1 = FileDateTime("C:\CSI\RptLogo.Bmp")
    'If ilRet <> 0 Then
    '    ilRet = 0
    '    MkDir "C:\CSI"
    '    If ilRet = 0 Then
    '        FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
    '    Else
    '        FileCopy sgDBPath & "RptLogo.Bmp", sgLogoPath & "RptLogo.Bmp"
    '    End If
    'Else
    '    ilRet = 0
    '    slDateTime2 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
    '    If ilRet = 0 Then
    '        If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
    '            FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
    '        End If
    '    End If
    'End If
     'ttp 5260
    'If Dir(sgLogoPath & "RptLogo.jpg") > "" Then
    '    If Dir("c:\csi\RptLogo.jpg") = "" Then
    '        FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
    '    'ok, both exist.  is logopath's more recent?
    '    Else
    '        slDateTime1 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
    '        slDateTime2 = FileDateTime("C:\CSI\RptLogo.jpg")
    '        If StrComp(slDateTime1, slDateTime2, vbBinaryCompare) <> 0 Then
     '           FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
    '        End If
    '    End If
    'End If
    'Determine number if X-Digital HeadEnds
    ReDim sgXDSSection(0 To 0) As String
    'slXMLINIInputFile = gXmlIniPath(True)
    'If LenB(slXMLINIInputFile) <> 0 Then
    '    ilRet = gSearchFile(slXMLINIInputFile, "[XDigital", True, 1, sgXDSSection())
    'End If
    'Test to see if this function has been ran before, if so don't run it again
    igEmailNeedsConv = False
    mCreateStatustype
    ilRet = gPopMarkets()
    ilRet = gPopMSAMarkets()         'MSA markets
    ilRet = gPopMntInfo("T", tgTerritoryInfo())
    ilRet = gPopMntInfo("C", tgCityInfo())
    ilRet = gPopOwnerNames()
    ilRet = gPopStations()
    ilRet = gPopVehicleOptions()
    ilRet = gPopVehicles()
    ilRet = gPopSellingVehicles()
    ilRet = gPopAdvertisers()
    ilRet = gPopReportNames()
    ilRet = gGetLatestRatecard()
    ilRet = gPopTimeZones()
    ilRet = gPopStates()
    ilRet = gPopFormats()
    ilRet = gPopAvailNames()
    ilRet = gPopMediaCodes()
    '6777
    ckcPartial.Value = vbChecked
    Exit Sub

mStartUpErr:
    ilRet = Err.Number
    Resume Next
ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
'    gMsg = ""
'    For Each gErrSQL In cnn.Errors
'        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
'            gMsg = "A SQL error has occured: "
'            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
'        End If
'    Next gErrSQL
'    On Error Resume Next
'    cnn.RollbackTrans
'    On Error GoTo 0
'    If gMsg = "" Then
'        gMsgBox "Error at Start-up " & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
'    End If
    'ttp 5217
    gHandleError "", FORMNAME & "-Form_Load"
    'ttp 4905 need to quit app
    bmNoPervasive = True
    If blNeedToCloseCnn Then
        cnn.RollbackTrans
    End If
    'unload affiliate  ttp 4905
    tmcTerminate.Enabled = True
End Sub
Private Sub mCreateStatustype()
    'Agreement only shows status- 1:; 2:; 5: and 9:
    'All other screens show all the status
    tgStatusTypes(0).sName = "1-Aired Live"        'In Agreement and Pre_Log use 'Air Live'
    tgStatusTypes(0).iPledged = 0
    tgStatusTypes(0).iStatus = 0
    tgStatusTypes(1).sName = "2-Aired Delay B'cast" '"2-Aired In Daypart"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(1).iPledged = 1
    tgStatusTypes(1).iStatus = 1
    tgStatusTypes(2).sName = "3-Not Aired Tech Diff"
    tgStatusTypes(2).iPledged = 2
    tgStatusTypes(2).iStatus = 2
    tgStatusTypes(3).sName = "4-Not Aired Blackout"
    tgStatusTypes(3).iPledged = 2
    tgStatusTypes(3).iStatus = 3
    tgStatusTypes(4).sName = "5-Not Aired Other"
    tgStatusTypes(4).iPledged = 2
    tgStatusTypes(4).iStatus = 4
    tgStatusTypes(5).sName = "6-Not Aired Product"
    tgStatusTypes(5).iPledged = 2
    tgStatusTypes(5).iStatus = 5
    tgStatusTypes(6).sName = "7-Aired Outside Pledge"  'In Pre-Log use 'Air-Outside Pledge'
    tgStatusTypes(6).iPledged = 3
    tgStatusTypes(6).iStatus = 6
    tgStatusTypes(7).sName = "8-Aired Not Pledged"  'in Pre-Log use 'Air-Not Pledged'
    tgStatusTypes(7).iPledged = 3
    tgStatusTypes(7).iStatus = 7
    'D.S. 11/6/08 remove the "or Aired" from the status 9 description
    'Affiliate Meeting Decisions item 5) f-iv
    'tgStatusTypes(8).sName = "9-Not Carried or Aired"
    tgStatusTypes(8).sName = "9-Not Carried"
    tgStatusTypes(8).iPledged = 2
    tgStatusTypes(8).iStatus = 8
    tgStatusTypes(9).sName = "10-Delay Cmml/Prg"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(9).iPledged = 1
    tgStatusTypes(9).iStatus = 9
    tgStatusTypes(10).sName = "11-Air Cmml Only"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(10).iPledged = 1
    tgStatusTypes(10).iStatus = 10
    tgStatusTypes(ASTEXTENDED_MG).sName = "MG"
    tgStatusTypes(ASTEXTENDED_MG).iPledged = 3
    tgStatusTypes(ASTEXTENDED_MG).iStatus = ASTEXTENDED_MG
    tgStatusTypes(ASTEXTENDED_BONUS).sName = "Bonus"
    tgStatusTypes(ASTEXTENDED_BONUS).iPledged = 3
    tgStatusTypes(ASTEXTENDED_BONUS).iStatus = ASTEXTENDED_BONUS
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).sName = "Replacement"
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).iPledged = 3
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).iStatus = ASTEXTENDED_REPLACEMENT
End Sub



Private Function mUpdateCptt(ilVefCode As Integer) As Boolean
    'Created by D.S. June 2007  Modified Dan M 11/02/10 v58 new values in cptt added 2/25/2011
    'Set the CPTT week's value

    Dim ilStatus As Integer
    Dim llVeh As Long
    'new values in cptt
    Dim slMondayFeedDate As String
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slSuDate As String
    Dim llAttCode As Long
    Dim llAstCount As Long
    Dim llNotAiredCount As Long
    Dim blNoSpotsAired As Boolean
    Dim slMsg As String
    Dim ilShtt As Integer
    Dim ilShttCode As Integer
    
    On Error GoTo ErrHand
    mUpdateCptt = False
    SQLQuery = "Select * FROM cptt WHERE cpttVefCode = " & ilVefCode
    '6777 'all' leaves the statement out
   ' SQLQuery = SQLQuery & " AND cpttPostingStatus <> 2"
    If imDiscrepancySearch = Partial Then
        SQLQuery = SQLQuery & " AND cpttPostingStatus <> 2"
    ElseIf imDiscrepancySearch = NotAired Then
        SQLQuery = SQLQuery & " AND cpttPostingStatus = 2"
    End If
    SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(smStartDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(smEndDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " Order By cpttShfCode, cpttStartDate"
    Set cptt_rst = cnn.Execute(SQLQuery)
    Do While Not cptt_rst.EOF
        If imTerminate Then
            Exit Function
        End If
        llAttCode = cptt_rst!cpttatfCode
        ilShttCode = cptt_rst!cpttShfCode
        slMondayFeedDate = Format(cptt_rst!CpttStartDate, "m/d/yy")
        slSuDate = DateAdd("d", 6, slMondayFeedDate)
        
        SQLQuery = "Select Count(*) FROM ast WHERE astAtfCode = " & llAttCode
        SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
        Set ast_rst = cnn.Execute(SQLQuery)
        If Not ast_rst.EOF Then
            llAstCount = ast_rst(0).Value
            If llAstCount > 0 Then
                '12/4/13: Test if partial week that was imported from Web
                If (cptt_rst!cpttPostingStatus = 1) And (cptt_rst!cpttStatus = 0) Then
                    mFixWebPartialWeekPosted llAttCode, slMondayFeedDate, slSuDate
                End If
                'Set any Not Aired to received as they are not exported
                For ilStatus = 0 To UBound(tgStatusTypes) Step 1
                    If (tgStatusTypes(ilStatus).iPledged = 2) Then
                        SQLQuery = "UPDATE ast SET "
                        SQLQuery = SQLQuery & "astCPStatus = " & "1"    'Received
                        SQLQuery = SQLQuery & " WHERE (astAtfCode = " & llAttCode
                        SQLQuery = SQLQuery & " AND astCPStatus = 0"
                        SQLQuery = SQLQuery & " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
                        SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            GoSub ErrHand:
                        End If
                    End If
                Next ilStatus
                llNotAiredCount = 0
                For ilStatus = 0 To UBound(tgStatusTypes) Step 1
                    If (tgStatusTypes(ilStatus).iPledged = 2) Then
                        SQLQuery = "Select Count(*) FROM ast WHERE astAtfCode = " & llAttCode
                        SQLQuery = SQLQuery & " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
                        SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                        Set ast_rst = cnn.Execute(SQLQuery)
                        If Not ast_rst.EOF Then
                            llNotAiredCount = llNotAiredCount + ast_rst(0).Value
                        End If
                    End If
                Next ilStatus
                If llAstCount <> llNotAiredCount Then
                    blNoSpotsAired = False
                Else
                    blNoSpotsAired = True
                End If
                If imTerminate Then
                    Exit Function
                End If
        
                'Determine if CPTTStatus should to set to 0=Partial or 1=Completed:  because of above code, will always be complete
                SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
                SQLQuery = SQLQuery & " AND astAtfCode = " & llAttCode
                SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                Set ast_rst = cnn.Execute(SQLQuery)
                If ast_rst.EOF Then
                    'Set CPTT as complete
                    SQLQuery = "UPDATE cptt SET "
                    llVeh = gBinarySearchVef(CLng(ilVefCode))
                    If llVeh <> -1 Then
                        If (tgVehicleInfo(llVeh).sVehType = "G") And (DateValue(slSuDate) > DateValue(Format$(gNow(), "m/d/yy"))) Then
                            SQLQuery = SQLQuery & "cpttStatus = 0" & ", " 'Partial
                            SQLQuery = SQLQuery & "cpttPostingStatus = 1" 'Partial
                        Else
                            If blNoSpotsAired Then
                                SQLQuery = SQLQuery & "cpttStatus = 2" & ", " 'Complete
                            Else
                                SQLQuery = SQLQuery & "cpttStatus = 1" & ", " 'Complete
                            End If
                            SQLQuery = SQLQuery & "cpttPostingStatus = 2"  'Complete
                        End If
                    Else
                        If blNoSpotsAired Then
                            SQLQuery = SQLQuery & "cpttStatus = 2" & ", " 'Complete
                        Else
                            SQLQuery = SQLQuery & "cpttStatus = 1" & ", " 'Complete
                        End If
                        SQLQuery = SQLQuery & "cpttPostingStatus = 2"  'Complete
                    End If
                    SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAttCode
                    SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHand:
                    End If
                    'Dan 6777
                    'lmNumberChgd = lmNumberChgd + 1
                    If cptt_rst!cpttPostingStatus <> 2 Then
                        lmNumberChgd = lmNumberChgd + 1
                    Else
                        If cptt_rst!cpttStatus = 2 And Not blNoSpotsAired Then
                            lmNumberChgd = lmNumberChgd + 1
                            lmNumberDidNotAirWrong = lmNumberDidNotAirWrong + 1
                            ilShtt = gBinarySearchStationInfoByCode(ilShttCode)
                            If ilShtt <> -1 Then
                                slMsg = Trim$(tgStationInfoByCode(ilShtt).sCallLetters) & ","
                            Else
                                slMsg = " ,"
                            End If
                            llVeh = gBinarySearchVef(CLng(ilVefCode))
                            If llVeh <> -1 Then
                                slMsg = slMsg & """" & Trim$(tgVehicleInfo(llVeh).sVehicle) & """" & ","
                            Else
                                slMsg = slMsg & " ,"
                            End If
                            slMsg = slMsg & Trim$(Str$(llAttCode))
                            slMsg = slMsg & "," & slMondayFeedDate & " to " & slSuDate & ",   Removed Did Not Air"
                            Print #hmMsg, slMsg
                        End If
                    End If
                Else
                    SQLQuery = "Select count(*) FROM ast WHERE astCPStatus = 1"
                    SQLQuery = SQLQuery & " AND astAtfCode = " & llAttCode
                    SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                    Set ast_rst = cnn.Execute(SQLQuery)
                    If Not ast_rst.EOF Then
                        If (ast_rst(0).Value > 0) And (ast_rst(0).Value <> llNotAiredCount) Then
                            lmNumberPartials = lmNumberPartials + 1
                            '6777
'                            slMsg = Trim$(Str$(llAttCode))
'                            slMsg = slMsg & "," & slMondayFeedDate
'                            llVeh = gBinarySearchVef(CLng(ilVefCode))
'                            If llVeh <> -1 Then
'                                slMsg = slMsg & "," & Trim$(tgVehicleInfo(llVeh).sVehicle)
'                            End If
'                            ilShtt = gBinarySearchStationInfoByCode(ilShttCode)
'                            If ilShtt <> -1 Then
'                                slMsg = slMsg & "," & Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
'                            End If
'                            Print #hmMsg, "Partially Posted:," & slMsg
                            ilShtt = gBinarySearchStationInfoByCode(ilShttCode)
                            If ilShtt <> -1 Then
                                slMsg = Trim$(tgStationInfoByCode(ilShtt).sCallLetters) & ","
                            Else
                                slMsg = " ,"
                            End If
                            llVeh = gBinarySearchVef(CLng(ilVefCode))
                            If llVeh <> -1 Then
                                slMsg = slMsg & """" & Trim$(tgVehicleInfo(llVeh).sVehicle) & """" & ","
                            Else
                                slMsg = slMsg & " ,"
                            End If
                            slMsg = slMsg & Trim$(Str$(llAttCode))
                            slMsg = slMsg & "," & slMondayFeedDate & " to " & slSuDate & ",   Partially Posted"
                            Print #hmMsg, slMsg
                        End If
                    End If
                End If
            End If
        End If
        lmRunningCount = lmRunningCount + 1
        lacProgress.Caption = "Processed " & lmRunningCount & " of " & lmTotal & ", Changed " & lmNumberChgd
        DoEvents
        cptt_rst.MoveNext
    Loop
    gFileChgdUpdate "cptt.mkd", True
    mUpdateCptt = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SetCPTTComplete: mUpdateCptt"
End Function


Private Sub mFixWebPartialWeekPosted(llAttCode As Long, slMoDate As String, slSuDate As String)
    Dim ilStatus As Integer
    Dim ilShttCode As Integer
    Dim slStaName As String
    Dim ilVehCode As Integer
    Dim slVehName As String
    Dim slStartDate As String
    Dim slEndDate As String
        
    
        
    On Error GoTo ErrHand
    slStartDate = Format$(slMoDate, "mm-dd-yyyy")
    slEndDate = Format$(slSuDate, "mm-dd-yyyy")
    
    SQLQuery = "Select * FROM att WHERE attCode = " & llAttCode
    Set att_rst = cnn.Execute(SQLQuery)
    If att_rst.EOF Then
        Exit Sub
    End If
    If att_rst!attExportType <> 1 Then
        Exit Sub
    End If
    If att_rst!attExportToWeb <> "Y" Then
        Exit Sub
    End If
    SQLQuery = "Select Count(*) FROM webl WHERE weblAttCode = " & llAttCode
    SQLQuery = SQLQuery & " AND (weblPostDay >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND weblPostDay <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
    Set webl_rst = cnn.Execute(SQLQuery)
    If webl_rst.EOF Then
        Exit Sub
    End If
    For ilStatus = 0 To UBound(tgStatusTypes) Step 1
        If (tgStatusTypes(ilStatus).iPledged <> 2) Then
            SQLQuery = "UPDATE ast SET "
            SQLQuery = SQLQuery & "astCPStatus = " & "1"    'Received
            SQLQuery = SQLQuery & " WHERE (astAtfCode = " & llAttCode
            SQLQuery = SQLQuery & " AND astCPStatus = 0"
            '12/13/13: Change to astStatus as that is how it is in affiliat2
            'SQLQuery = SQLQuery & " AND astPledgeStatus = " & tgStatusTypes(ilStatus).iStatus
            SQLQuery = SQLQuery & " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
            SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                GoSub ErrHand:
            End If
        End If
    Next ilStatus
    ilShttCode = gGetShttCodeFromAttCode(CStr(llAttCode))
    slStaName = gGetCallLettersByShttCode(ilShttCode)
    ilVehCode = gGetVehCodeFromAttCode(CStr(llAttCode))
    slVehName = gGetVehNameByVefCode(ilVehCode)
    lmPart2Comp = lmPart2Comp + 1
    '6777
    'Print #hmMsg, slStaName & ", " & slVehName & ", " & slStartDate & ", " & slEndDate & ",   Changed From Partial To Complete."
    Print #hmMsg, slStaName & ", " & """" & slVehName & """" & ", " & llAttCode & ", " & slStartDate & " to " & slEndDate & ",   Changed From Partial To Complete."

    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "SetCPTTComplete: mFixWebPartialWeekPosted"
End Sub
