VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRadarClrRpt 
   Caption         =   "Radar Clearance Report"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   7575
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frcSelection 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   7200
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6480
         Picture         =   "AffRadarClrRpt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Select Stations from File.."
         Top             =   2160
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   180
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         borderstyle     =   1
         csi_showdropdownonfocus=   -1
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   -1
         csi_defaultdatetype=   1
         csi_caldateformat=   1
         font            =   "AffRadarClrRpt.frx":056A
         csi_daynamefont =   "AffRadarClrRpt.frx":0596
         csi_monthnamefont=   "AffRadarClrRpt.frx":05C4
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   180
         Width           =   855
         _extentx        =   1508
         _extenty        =   503
         borderstyle     =   1
         csi_showdropdownonfocus=   -1
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0
         csi_allowblankdate=   -1
         csi_allowtfn    =   -1
         csi_defaultdatetype=   1
         csi_caldateformat=   1
         font            =   "AffRadarClrRpt.frx":05F2
         csi_daynamefont =   "AffRadarClrRpt.frx":061E
         csi_monthnamefont=   "AffRadarClrRpt.frx":064C
      End
      Begin VB.ListBox lbcStatus 
         Height          =   840
         ItemData        =   "AffRadarClrRpt.frx":067A
         Left            =   120
         List            =   "AffRadarClrRpt.frx":067C
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         Top             =   3000
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1620
         Index           =   1
         ItemData        =   "AffRadarClrRpt.frx":067E
         Left            =   3840
         List            =   "AffRadarClrRpt.frx":0680
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   2520
         Width           =   3075
      End
      Begin VB.ListBox lbcSelection 
         Height          =   1620
         Index           =   0
         ItemData        =   "AffRadarClrRpt.frx":0682
         Left            =   3840
         List            =   "AffRadarClrRpt.frx":0684
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   480
         Width           =   3075
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox CkcAll 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   180
         Width           =   1455
      End
      Begin VB.CheckBox ckcInclNotRecd 
         Caption         =   "Include Stations Not Reported"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox ckcIncludeMiss 
         Caption         =   "Include Not Aired"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Aired Dates- Start"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "End"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   210
         Width           =   465
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3240
      Top             =   960
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6300
      FormDesignWidth =   7575
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   20
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   19
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   18
      Top             =   225
      Width           =   2685
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1050
         TabIndex        =   4
         Top             =   765
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   1170
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   810
         Width           =   690
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   525
         Width           =   2130
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmRadarClrRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************************
'*  frmRadarClrRpt - Create a report of spots (from AST)generated for
'   Radar vehicles only.  The radar vehicles will be obtained from the
'   file RHT, RET which contains the radar network times and codes.
'   Create a prepass file in AFR which has a pointer to the AST file
'   The report will contain spot detail along with a radar network code,
'   determined by the time entry within the network table.  If a time
'   entry isnt found for the spot, the network code will be blank.
'
'   this form is a copy of frmPldgAirRpt.frm.  Status list box is populated,
'   and Not Carried status is always excluded.  Misses are options; as well
'   as Not Reported CPs.   Vehicle selectivity based on the radar vehicles
'   populated and hidden to include with common routine;
'   all stations are included without an option for selectivity
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imckcAllIgnore As Integer           '3-6-05
Private imckcAllStationsIgnore As Integer
Private hmAst As Integer
Private tmStatusOptions As STATUSOPTIONS
Private imListBoxWidth As Integer
'
'        mGetStationSelection - get all the selected stations from user selection
'       <input> ilCkcAll - 0 = selected station (not all)
'               lbcListBox - list box of station (lbcVehAff or lbcSelection(1)
'       <return> SQL string selected vehicles
Function mGetStationSelection(ilCkcAll As Integer, lbcListBox As control) As String
Dim i As Integer
Dim slStr As String
    slStr = ""
    If ilCkcAll = 0 Then    'User did NOT select all vehicles
        For i = 0 To lbcListBox.ListCount - 1 Step 1
            If lbcListBox.Selected(i) Then
                If Len(slStr) = 0 Then
                    slStr = "(shttCode = " & lbcListBox.ItemData(i) & ")"
                Else
                    slStr = slStr & " OR (shttCode = " & lbcListBox.ItemData(i) & ")"
                End If
            End If
        Next i
    End If
    mGetStationSelection = slStr
        
End Function

Private Sub CkcAll_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imckcAllIgnore Then
        Exit Sub
    End If
    If CkcAll.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcSelection(0).ListCount > 0 Then
        imckcAllIgnore = True
        lRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllStations_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imckcAllStationsIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcSelection(1).ListCount > 0 Then
        imckcAllStationsIgnore = True
        lRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSelection(1).hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllStationsIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload frmRadarClrRpt
End Sub

'       Radar Spot Clearance
'
Private Sub cmdReport_Click()
    'Dan 7/20/11 this creates 3 variants and 1 integer
'    Dim i, j, X, Y, iPos As Integer
    Dim i As Integer, j As Integer, X As Integer, Y As Integer, iPos As Integer
    Dim sCode As String
    'Dan 7/20/11 bm and sName not used.  Can't dim this way, first 4 in line become variants
    Dim sStatus As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim sStr As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    'Dim NewForm As New frmViewReport
    Dim slNow As String
    Dim ilSelected As Integer
    Dim ilNotSelected As Integer
    Dim slStatusSelected As String
    Dim slStatusNotSelected As String
    Dim slSelection As String
    'Dim sGenDate As String
    'Dim sGenTime As String
    Dim slInputStartDate As String
    Dim slInputEndDate As String
    Dim ilShowExact As Integer
    Dim ilIncludeNotCarried As Integer
    Dim ilIncludeNotReported As Integer
    Dim slAdjustedStartDate As String
    
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer           'required for common rtn, this report doesnt use
    Dim blUseAirDAte As Boolean         'true to use air date vs feed date
    Dim ilAdvtOption As Integer         'true if one or more advt selected, false if not advt option or ALL advt selected
    Dim blFilterAvailNames As Boolean   'avail names selectivity
    Dim blIncludePledgeInfo As Boolean
    Dim tlSpotRptOptions As SPOT_RPT_OPTIONS
    ReDim ilSelectedVehicles(0 To 0) As Integer     '5-30-18
    
    On Error GoTo ErrHand
    
    'sGenDate = Format$(gNow(), "m/d/yyyy")
    'sGenTime = Format$(gNow(), sgShowTimeWSecForm)
    sgGenDate = Format$(gNow(), "m/d/yyyy")         '7-10-13 use global gen date/time for crystal filtering
    sgGenTime = Format$(gNow(), sgShowTimeWSecForm)

    
    'debugging only for timing tests
    'Dim sGenStartTime As String
    'Dim sGenEndTime As String
    'sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)

    slInputStartDate = Trim$(CalOnAirDate.Text)
    
    slInputEndDate = Trim$(CalOffAirDate.Text)
     
    If gIsDate(slInputStartDate) = False Or (Len(Trim$(slInputStartDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    If gIsDate(slInputEndDate) = False Or (Len(Trim$(slInputEndDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOffAirDate.SetFocus
        Exit Sub
    End If
      
    slAdjustedStartDate = slInputStartDate
    Do While Weekday(slAdjustedStartDate, vbSunday) <> vbMonday
        slAdjustedStartDate = DateAdd("d", -1, slAdjustedStartDate)
    Loop
    '2-21-14 no need to backup the date to get the true spots aired with new design and keys
    'slAdjustedStartDate = DateAdd("d", -7, slAdjustedStartDate)       'using air dates, need to back up weekto process to get the spots moved to following week

    slInputStartDate = Format(slInputStartDate, "m/d/yyyy")
    slInputEndDate = Format(slInputEndDate, "m/d/yyyy")
      
    Screen.MousePointer = vbHourglass
  
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilExportType = cboFileType.ListIndex       '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    sStartDate = Format(slInputStartDate, "m/d/yyyy")
    sEndDate = Format(slInputEndDate, "m/d/yyyy")
    dFWeek = CDate(slInputStartDate)
    sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    dFWeek = CDate(slInputEndDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    
    sStatus = ""
    slStatusSelected = ""
    slStatusNotSelected = ""
    
    ' Detrmine what to sort by
    ilAdvtOption = False            'assume to get all advt
    
    'determine option to include non-reported stations
    If Not ckcInclNotRecd.Value = vbChecked Then    'exclude non-reported (or cp not received) stations
        ilIncludeNotReported = False
    Else
        ilIncludeNotReported = True
    End If
    
    'get the description of spot statuses (included/excluded) to show on report
    gGetSQLStatusForCrystal lbcStatus, sStatus, slSelection, ilIncludeNotCarried, ilIncludeNotReported
  
    gInitStatusSelections tmStatusOptions               '3-26-12 set all options to exclude
    'all aired spots are included
    tmStatusOptions.iInclLive0 = True
    tmStatusOptions.iInclDelay1 = True
    tmStatusOptions.iInclAirOutPledge6 = True
    tmStatusOptions.iInclAiredNotPledge7 = True
    tmStatusOptions.iInclDelayCmmlOnly9 = True
    tmStatusOptions.iInclAirCmmlOnly10 = True
    tmStatusOptions.iInclMG11 = True
    tmStatusOptions.iInclBonus12 = True
    tmStatusOptions.iInclRepl13 = True
    
    slStatusNotSelected = "Excluded: Not Carried"
    If ckcIncludeMiss.Value = vbChecked Then        'check if not aired to be included
        tmStatusOptions.iInclMissed2 = True
        tmStatusOptions.iInclMissed3 = True
        tmStatusOptions.iInclMissed4 = True
        tmStatusOptions.iInclMissed5 = True
        tmStatusOptions.iInclResolveMissed = True
        tmStatusOptions.iInclMissedMGBypass14 = True        '4-21-17 missed - MG bypassed
    Else
        slStatusNotSelected = slStatusNotSelected & ", Not Aired"
    End If
    
    'determine option to include non-reported stations
    If ckcInclNotRecd.Value = vbChecked Then     'include non-reported (or cp not received) stations
        tmStatusOptions.iNotReported = True
    Else
        slStatusNotSelected = slStatusNotSelected & ", Not Reported"
    End If
    sgCrystlFormula3 = Trim$(slStatusNotSelected)
        
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "STart time: " + slNow, vbOKOnly
    'if both missed spots and not reported are excluded, only get those spots already marked as aired
    
     
     'Always exclude spots not carried.
    ilShowExact = True
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
    'sStartDate = DateAdd("d", -7, sStartDate)       'using air dates, need to back up weekto process to get the spots moved to following week
    blUseAirDAte = True   'use air date vs feed date
                                
    ilIncludeNonRegionSpots = True          '7-22-10 include spots with/without regional copy
    ilFilterCatBy = -1             '7-22-10 no filering of categories in this report, required for common rtn
                                'send any list box as last parameter -gbuildaststnclr, but it wont be used
    blFilterAvailNames = False      'no selective testing for avail names, include all

    'send the options sent into structure for general rtn
    tlSpotRptOptions.sStartDate = sStartDate
    tlSpotRptOptions.sEndDate = sEndDate
    tlSpotRptOptions.bUseAirDAte = blUseAirDAte
    tlSpotRptOptions.iAdvtOption = ilAdvtOption
    tlSpotRptOptions.iCreateAstInfo = True
    tlSpotRptOptions.iShowExact = ilShowExact
    tlSpotRptOptions.iIncludeNonRegionSpots = ilIncludeNonRegionSpots
    tlSpotRptOptions.iFilterCatBy = ilFilterCatBy
    tlSpotRptOptions.bFilterAvailNames = blFilterAvailNames
    tlSpotRptOptions.bIncludePledgeInfo = blIncludePledgeInfo
    tlSpotRptOptions.lContractNumber = 0            '6-4-18 no single contract option in this report

    'parameter after blFilterAvailNames is unused (placed for filler) since no avail name filtering is required; include all names
    'gBuildAstSpotsByStatus hmAst, slAdjustedStartDate, sEndDate, blUseAirDAte, lbcSelection(0), lbcSelection(1), False, lbcSelection(0), True, ilShowExact, ilIncludeNonRegionSpots, ilFilterBy, lbcSelection(0), tmStatusOptions, blFilterAvailNames, lbcSelection(0)
    gCopySelectedVehicles lbcSelection(0), ilSelectedVehicles()         '5-30-18
'    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, lbcSelection(0), lbcSelection(1), lbcSelection(0), lbcSelection(0), lbcSelection(0)
    gBuildAstSpotsByStatus hmAst, tlSpotRptOptions, tmStatusOptions, ilSelectedVehicles(), lbcSelection(1), lbcSelection(0), lbcSelection(0), lbcSelection(0)
    On Error GoTo ErrHand
    
    SQLQuery = "Select afrastCode,afrISCI, "
    '12/11/13: Pledge information obtained from astInfo instead of ast
    'SQLQuery = SQLQuery & "astAtfCode, astLsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astPledgeDate, astPledgeStartTime, astPledgeEndTime, astPledgeStatus, "
    SQLQuery = SQLQuery & "astAtfCode, astLsfCode, astAirDate, astAirTime, astStatus, astCPStatus, afrPledgeDate, afrPledgeStartTime, afrPledgeEndTime, afrPledgeStatus, "
    SQLQuery = SQLQuery & "shttCallLetters, adfName, VefName,  "
    SQLQuery = SQLQuery & " cpfISCI, cpfName, mnfName, mnfcode "

    SQLQuery = SQLQuery & " FROM afr INNER JOIN ast on afrastcode = astcode "
    SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on astadfcode = adfcode "
    SQLQuery = SQLQuery & " INNER JOIN shtt on astshfcode = shttcode "
    SQLQuery = SQLQuery & " INNER JOIN VEF_Vehicles on astvefcode = vefcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on astcpfcode = cpfcode "
    SQLQuery = SQLQuery & " LEFT OUTER JOIN mnf_Multi_Names on AfrMissedMnfCode = mnfcode "

    'use air dates vs fed dates
    SQLQuery = SQLQuery & "WHERE (astAirDate >= '" & Format$(slInputStartDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(slInputEndDate, sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " AND (afrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
   
    
    'slNow = Format$(gNow(), "hh:mm:ss AMPM")
    'gMsgBox "End time: " + slNow, vbOKOnly
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
    
   frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfRadarClr.rpt", "AfRadarClr"
  
    
'    cmdReport.Enabled = True            'give user back control to gen, done buttons
'    cmdDone.Enabled = True
'    cmdReturn.Enabled = True
    
    'debugging only for timing tests
    'sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime
     gUserActivityLog "S", sgReportListName & ": Clear AFR"
   
    'remove all the records just printed
    SQLQuery = "DELETE FROM afr "
    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "RadarClrRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True

    gUserActivityLog "E", sgReportListName & ": Clear AFR"
    
    Screen.MousePointer = vbDefault

        
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarClrRpt-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmRadarClrRpt
End Sub

'TTP 9943 - Add ability to import stations for report selectivity
Private Sub cmdStationListFile_Click()
    Dim slCurDir As String
    slCurDir = CurDir
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    ' Import from the Selected File
    gSelectiveStationsFromImport lbcSelection(1), ckcAllStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Activate()
'    'grdVehAff.Columns(0).Width = grdVehAff.Width
   'mListBoxWidth = lbcVehAff.Width
       
End Sub

Private Sub Form_Initialize()
'    Me.Width = Screen.Width / 1.3
'    Me.Height = Screen.Height / 1.3
'    Me.Top = (Screen.Height - Me.Height) / 2
'    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Screen.Width / 1.3, Screen.Height / 1.3
    gSetFonts frmRadarClrRpt
    frmRadarClrRpt.Caption = "Radar Clearance Report- " & sgClientName
    
    gCenterForm frmRadarClrRpt
End Sub
Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim sVehicleStn As String           '4-9-04
    Dim lRg As Long
    Dim lRet As Long
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilFound As Integer
    Dim ilHideNotCarried As Integer

    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    igRptIndex = frmReports!lbcReports.ItemData(frmReports!lbcReports.ListIndex)

    imckcAllIgnore = False
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        sVehicleStn = rst!siteVehicleStn         'are vehicles also stations?
    End If
    

    gPopExportTypes cboFileType         '3-15-04 Populate all export types
    cboFileType.Enabled = False         'disable the export types since display mode is default
    
    ilHideNotCarried = True             'deselect Not Carried status (never include them_
    gPopSpotStatusCodesExt lbcStatus, ilHideNotCarried         '3-14-12 populate list box with hard-coded extended spot status codes
    gPopVff
    
    
    CkcAll.Caption = "All Vehicles"
    CkcAll.Value = Unchecked
    lbcSelection(0).Clear
    
    SQLQuery = "Select distinct rhtvefcode from RHT"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        ilVef = gBinarySearchVef(rst!rhtVefCode)
        If ilVef >= 0 Then
            lbcSelection(0).AddItem tgVehicleInfo(ilVef).sVehicleName
            lbcSelection(0).ItemData(lbcSelection(0).NewIndex) = rst!rhtVefCode
        End If
        rst.MoveNext
    Wend
       
    'lbcSelection(1) = stations
    'lbcSelection(0) = vehicles
    ckcAllStations.Caption = "All Stations"
    ckcAllStations.Value = vbUnchecked
    lbcSelection(1).Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcSelection(1).AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcSelection(1).ItemData(lbcSelection(1).NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop

    ckcAllStations.Value = vbChecked

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set frmRadarClrRpt = Nothing
End Sub



Private Sub lbcSelection_Click(Index As Integer)
 
'    If imckcAllIgnore Then
'        Exit Sub
'    End If
    If Index = 0 Then
        If imckcAllIgnore Then
            Exit Sub
        End If
        If CkcAll.Value = vbChecked Then
            imckcAllIgnore = True
            CkcAll.Value = vbUnchecked
            imckcAllIgnore = False
        End If
    Else
        If imckcAllStationsIgnore Then
            Exit Sub
        End If
        If ckcAllStations.Value = vbChecked Then
            imckcAllStationsIgnore = True
            ckcAllStations.Value = vbUnchecked
            imckcAllStationsIgnore = False
        End If
    End If
 
End Sub


Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       '3-15-04 default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub


'
'
'        mGetVehicleSelection - get all the selected vehicles from user selection
'       <input> ilCkcAll - 0 = selected vehicle (not all)
'               lbcListBox - list box of vehicles (lbcVehAff or lbcSelection(0)
'       <return> SQL string selected vehicles
Function mGetVehicleSelection(ilCkcAll As Integer, lbcListBox As control) As String
Dim i As Integer
Dim slStr As String
    slStr = ""
    If ilCkcAll = 0 Then    'User did NOT select all vehicles
        For i = 0 To lbcListBox.ListCount - 1 Step 1
            If lbcListBox.Selected(i) Then
                If Len(slStr) = 0 Then
                    slStr = "(vefCode = " & lbcListBox.ItemData(i) & ")"
                Else
                    slStr = slStr & " OR (vefCode = " & lbcListBox.ItemData(i) & ")"
                End If
            End If
        Next i
    End If
    mGetVehicleSelection = slStr
        
End Function

Private Sub optTimeSort_Click(Index As Integer)
    '4-17-08 regardless of sort, always leave ShowExactStationFeed as checked on and not shown
    '        this will exclude spots not carried
    'If optTimeSort(0).Value = True Then
    '    chkShowExact.Visible = True
    'Else
    '    chkShowExact.Visible = False
    '    chkShowExact.Value = vbUnchecked
    'End If
End Sub
