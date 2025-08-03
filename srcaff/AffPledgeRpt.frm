VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPledgeRpt 
   Caption         =   "Pledge Report"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "AffPledgeRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   7125
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      FormDesignWidth =   7125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   240
      TabIndex        =   6
      Top             =   1845
      Width           =   6705
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6000
         Picture         =   "AffPledgeRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Select Stations from File.."
         Top             =   120
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1350
         TabIndex        =   7
         Top             =   240
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
         font            =   "AffPledgeRpt.frx":0E34
         csi_daynamefont =   "AffPledgeRpt.frx":0E60
         csi_monthnamefont=   "AffPledgeRpt.frx":0E8E
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   285
         Left            =   1350
         TabIndex        =   8
         Top             =   675
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
         font            =   "AffPledgeRpt.frx":0EBC
         csi_daynamefont =   "AffPledgeRpt.frx":0EE8
         csi_monthnamefont=   "AffPledgeRpt.frx":0F16
      End
      Begin VB.CheckBox chkPageSkip 
         Caption         =   "Skip to new page each station"
         Height          =   255
         Left            =   225
         TabIndex        =   13
         Top             =   1905
         Width           =   2505
      End
      Begin VB.ListBox lbcStations 
         Height          =   2595
         ItemData        =   "AffPledgeRpt.frx":0F44
         Left            =   4755
         List            =   "AffPledgeRpt.frx":0F4B
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   420
         Width           =   1665
      End
      Begin VB.CheckBox chkAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   4770
         TabIndex        =   18
         Top             =   150
         Width           =   1215
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   2595
         ItemData        =   "AffPledgeRpt.frx":0F53
         Left            =   2940
         List            =   "AffPledgeRpt.frx":0F55
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   420
         Width           =   1665
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   810
         Left            =   180
         TabIndex        =   9
         Top             =   1080
         Width           =   2325
         Begin VB.OptionButton optSP 
            Caption         =   "Missing Pledges"
            Height          =   255
            Index           =   1
            Left            =   30
            TabIndex        =   11
            Top             =   240
            Width           =   1620
         End
         Begin VB.OptionButton optSP 
            Caption         =   "Non-Live Pledges"
            Height          =   255
            Index           =   0
            Left            =   30
            TabIndex        =   10
            Top             =   0
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton optSP 
            Caption         =   "All Pledges"
            Height          =   255
            Index           =   2
            Left            =   30
            TabIndex        =   12
            Top             =   480
            Width           =   1470
         End
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   2955
         TabIndex        =   16
         Top             =   150
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Start Date:"
         Height          =   225
         Left            =   225
         TabIndex        =   14
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Label4 
         Caption         =   "End Date:"
         Height          =   255
         Left            =   225
         TabIndex        =   15
         Top             =   735
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4590
      TabIndex        =   22
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4350
      TabIndex        =   21
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4155
      TabIndex        =   20
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
         ItemData        =   "AffPledgeRpt.frx":0F57
         Left            =   1050
         List            =   "AffPledgeRpt.frx":0F59
         TabIndex        =   4
         Top             =   765
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   135
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
         Left            =   135
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
Attribute VB_Name = "frmPledgeRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmPledgeRpt - List of pledge times by vehicle, station
'*                 Exceptions only (not Live) or All
'*
'*  Created 12/1/99 D Hosaka
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer
Private imChkAllStationsBoxIgnore As Integer



Private Sub chkAllStations_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkAllStationsBoxIgnore Then
        Exit Sub
    End If
    If chkAllStations.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStations.ListCount > 0 Then
        imChkAllStationsBoxIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkAllStationsBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkListBox_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehAff.ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcVehAff.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehAff.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdDone_Click()
    Unload frmPledgeRpt
End Sub
Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim sCode As String
    Dim bm As Variant
    Dim sName, sVehicles As String
    Dim sDateRange As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim AgreeRst As ADODB.Recordset
    Dim sCurDate As String
    Dim sCurTime As String
    Dim sStr As String
    Dim sFdDays As String
    Dim sFdSTime As String
    Dim sFdETime As String
    Dim sFdStatus As String
    Dim sPdDays As String
    Dim sPdSTime As String
    Dim sPdETime As String
    Dim lAtfCode As Long
    Dim iFdStatus As Integer
    Dim sFeedTime As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    'ReDim ilStationCodes(1 To 1) As Integer
    ReDim ilStationCodes(0 To 0) As Integer
    Dim ilIncludeCodes As Integer
    'ReDim ilVehicleCodes(1 To 1) As Integer
    ReDim ilVehicleCodes(0 To 0) As Integer
    Dim ilIncludeVehicleCodes As Integer
    Dim ilFoundStation As Integer
    Dim ilTemp As Integer
    Dim slStationSelectionQuery As String
    'Dim NewForm As New frmViewReport
    
    On Error GoTo ErrHand
    sStartDate = Trim$(CalOnAirDate.Text)
    If sStartDate = "" Then
        sStartDate = "1/1/1970"
    End If
    sEndDate = Trim$(CalOffAirDate.Text)
    If sEndDate = "" Then
        sEndDate = "12/31/2069"
    End If
    If gIsDate(sStartDate) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sEndDate) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalOffAirDate.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'CRpt1.Connect = "DSN = " & sgDatabaseName
  
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmPledgeRpt, "Pledge.rpt", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex    '3-15-04
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    sStartDate = Format(sStartDate, "m/d/yyyy")
    sEndDate = Format(sEndDate, "m/d/yyyy")
    sDateRange = "(attOffAir >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & ") And (attDropDate >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & ") And (attOnAir <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'" & ")"
    sVehicles = ""
    
'    If chkListBox.Value = 0 Then    '= 0 Then                        'User did NOT select all vehicles
'        For i = 0 To lbcVehAff.ListCount - 1 Step 1
'            If lbcVehAff.Selected(i) Then
'               If Len(sVehicles) = 0 Then
'                   sVehicles = "(vefCode = " & lbcVehAff.ItemData(i) & ")"
'               Else
'                   sVehicles = sVehicles & " OR (vefCode = " & lbcVehAff.ItemData(i) & ")"
'               End If
'            End If
'        Next i
'    End If

    gObtainCodes lbcVehAff, ilIncludeVehicleCodes, ilVehicleCodes()        'build array of which station codes to incl/excl
    
    sVehicles = gFormInclExclQuery("AttVefCode", ilIncludeVehicleCodes, ilVehicleCodes())

'
    '11-29-11 option for selection stations
    gObtainCodes lbcStations, ilIncludeCodes, ilStationCodes()        'build array of which station codes to incl/excl
    
    slStationSelectionQuery = gFormInclExclQuery("AttShfCode", ilIncludeCodes, ilStationCodes())
    If chkPageSkip.Value = vbChecked Then
        sgCrystlFormula4 = "Y"         'skip new page each station
    Else
        sgCrystlFormula4 = "N"          'only skip page for new vehicles
    End If
    
    'CRpt1.SQLQuery = SQLQuery
    'CRpt1.ReportFileName = sgReportDirectory + "afPledge.rpt"
    'Send flag indicating Exceptions Only
    If optSP(0).Value Then
        'CRpt1.Formulas(0) = "ExceptFlag = 'Y'"  'show only non-live pledges
        sgCrystlFormula1 = "Y" 'ExceptFlag
        iFdStatus = 1                    'show  statuses greater or equal to 1 (ignore 0 which is live)
    ElseIf optSP(1).Value Then           'discreps only( missing pledges defined)
        'CRpt1.Formulas(0) = "ExceptFlag =  'D'"
        sgCrystlFormula1 = "D" 'ExceptFlag
        iFdStatus = 32000                   'this really doesnt apply since were looking for no Pledges defined
    Else
        'CRpt1.Formulas(0) = "ExceptFlag = 'N'"  'show all pledge times
        sgCrystlFormula1 = "N" 'ExceptFlag
        iFdStatus = 0                           'show all status
    End If
    dFWeek = CDate(sStartDate)
    'CRpt1.Formulas(1) = "StartDate = Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")" 'StartDate
    dFWeek = CDate(sEndDate)
    'CRpt1.Formulas(2) = "EndDate = Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")" 'EndDate
    
    sCurDate = Format(gNow(), sgShowDateForm)    'current date and time used as key for prepass file to
                                              'access and clear
    sCurTime = Format(gNow(), sgShowTimeWSecForm)
    
    'Gather agreements and associated pledge times--create SVR prepass record
    'for Crystal
    
    SQLQuery = "SELECT attCode, attshfCode, attvefCode, vefcode "
    'SQLQuery = SQLQuery + " FROM VEF_Vehicles, shtt, att"
'    SQLQuery = SQLQuery + " WHERE (vefCode = attVefCode"
'    SQLQuery = SQLQuery + " AND attshfCode = shttCode"
    SQLQuery = SQLQuery + " From Att inner join vef_vehicles on attvefcode = vefcode "
    'SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y' and (" & sDateRange & ")"
    SQLQuery = SQLQuery + " where (attServiceAgreement <> 'Y' and (" & sDateRange & ")"
    If sVehicles <> "" Then
        SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
    End If
    If Trim$(slStationSelectionQuery) <> "" Then
        SQLQuery = SQLQuery + " and (" & slStationSelectionQuery & ")"
    End If
    SQLQuery = SQLQuery + ")" '+ "ORDER BY vef.vefName, shtt.shttCallLetters"
    Set AgreeRst = gSQLSelectCall(SQLQuery)
    If Not AgreeRst.EOF Then        'test for end of retrieval
        While Not AgreeRst.EOF      'loop and process agreements

            'gather all pledge information for this agreement
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM dat"
            SQLQuery = SQLQuery + " WHERE (datAtfCode = " & AgreeRst!attCode & ""
            SQLQuery = SQLQuery + " AND datShfCode= " & AgreeRst!attshfcode & ""
            SQLQuery = SQLQuery + " AND datVefCode = " & AgreeRst!attvefCode & ")"
            SQLQuery = SQLQuery & " ORDER BY datFdStTime"
            'dont filter out the live status yet, there will be an erroneous "no Pleges defined"
            'SQLQuery = SQLQuery + " AND dat.datFdStatus >= " & iFdStatus & ")"
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                While Not rst.EOF
                    ilFoundStation = False
                    If ilIncludeCodes Then      'include
                        'For ilTemp = 1 To UBound(ilStationCodes) - 1 Step 1
                        For ilTemp = LBound(ilStationCodes) To UBound(ilStationCodes) - 1 Step 1
                            If ilStationCodes(ilTemp) = AgreeRst!attshfcode Then
                                ilFoundStation = True
                                Exit For
                            End If
                        Next ilTemp
                    Else                            'exclude
                        ilFoundStation = True
                        'For ilTemp = 1 To UBound(ilStationCodes) - 1 Step 1
                        For ilTemp = LBound(ilStationCodes) To UBound(ilStationCodes) - 1 Step 1
                            If ilStationCodes(ilTemp) = AgreeRst!attshfcode Then
                                ilFoundStation = False
                                Exit For
                            End If
                        Next ilTemp
                    End If
                    
                    If (ilFoundStation) Then            'include this station
                        If rst!datFdStatus >= iFdStatus Then    'either include or exclude the Live status
                            sFdStatus = ""
                            '1-14-09 show all the statuses of the agreement
                            sFdStatus = tgStatusTypes(rst!datFdStatus).sName
    '                        Select Case rst!datFdStatus
    '                            Case 0
    '                                sFdStatus = "Live"
    '                            Case 1
    '                                sFdStatus = "Delayed"
    '                            Case 2
    '                                sFdStatus = "NA-Tech problem"
    '                            Case 3
    '                                sFdStatus = "NA-Blackout"
    '                            Case 4
    '                                sFdStatus = "NA-Other"
    '                            Case 5
    '                                sFdStatus = "NA-Product"
    '                            Case 6
    '
    '                            Case 7
    '                                sFdStatus = "Carried-No Pledge"
    '                            Case 8
    '                                sFdStatus = "NA-Off Air"
    '                        End Select
                    
                            If Second(rst!datFdStTime) = 0 Then
                                sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWOSecForm)
                                sFeedTime = Format$(CStr(rst!datFdStTime), sgShowTimeWOSecForm)
                            Else
                                sFeedTime = Format$(CStr(rst!datFdStTime), sgShowTimeWSecForm)
                                sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWSecForm)
                            End If
                            'sFeedTime = Format$(CStr(rst!datFdStTime), "hh:mm:ss")
                            If Second(rst!datFdEdTime) = 0 Then
                                sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWOSecForm)
                            Else
                                sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWSecForm)
                            End If
                            
                            'If rst!datFdStatus <= 1 Then
                            If rst!datFdStatus <> 8 Then        'not carried shouldnt show times
                                If Second(rst!datPdStTime) = 0 Then
                                    sPdSTime = Format$(CStr(rst!datPdStTime), sgShowTimeWOSecForm)
                                Else
                                    sPdSTime = Format$(CStr(rst!datPdStTime), sgShowTimeWSecForm)
                                End If
                                If Second(rst!datPdEdTime) = 0 Then
                                    sPdETime = Format$(CStr(rst!datPdEdTime), sgShowTimeWOSecForm)
                                Else
                                    sPdETime = Format$(CStr(rst!datPdEdTime), sgShowTimeWSecForm)
                                End If
                            Else
                                sPdSTime = ""
                                sPdETime = ""
                            End If
                            
                            sStr = ""        'setup string of fed days of week
                            If rst!datFdMon <> 0 Then
                                sStr = sStr + "Mo"
                            End If
                            If rst!datFdTue <> 0 Then
                                sStr = sStr + "Tu"
                            End If
                            If rst!datFdWed <> 0 Then
                                sStr = sStr + "We"
                            End If
                            If rst!datFdThu <> 0 Then
                                sStr = sStr + "Th"
                            End If
                            If rst!datFdFri <> 0 Then
                                sStr = sStr + "Fr"
                            End If
                            If rst!datFdSat <> 0 Then
                                sStr = sStr + "Sa"
                            End If
                            If rst!datFdSun <> 0 Then
                                sStr = sStr + "Su"
                            End If
                            sFdDays = gDayMap(sStr)
                            
                            sStr = ""        'setup string of pledged days of week
                            If rst!datPdMon <> 0 Then
                                sStr = sStr + "Mo"
                            End If
                            If rst!datPdTue <> 0 Then
                                sStr = sStr + "Tu"
                            End If
                            If rst!datPdWed <> 0 Then
                                sStr = sStr + "We"
                            End If
                            If rst!datPdThu <> 0 Then
                                sStr = sStr + "Th"
                            End If
                            If rst!datPdFri <> 0 Then
                                sStr = sStr + "Fr"
                            End If
                            If rst!datPdSat <> 0 Then
                                sStr = sStr + "Sa"
                            End If
                            If rst!datPdSun <> 0 Then
                                sStr = sStr + "Su"
                            End If
                            sPdDays = gDayMap(sStr)
                            
                            'Create record for Crystal
                            'svrVefCode = vehicle code
                            'svrSeq = Station code
                            'svrHd1CefCode = Agreement Code
                            'svrFt1CefCode = DAT code for estimated times
                            'svrProductMo = Feed Days of week
                            'svrProductTu = Feed Start Time
                            'svrproductWe = Feed End Time
                            'svrProductTh = Feed Status
                            'svrProductFr = Pledge days of week
                            'svrProductSa = Pledge Start Time
                            'svrProductSu = Pledge End Time
                            'svrBreakMo = Seq # for feed time within agreement
                            
                            sFdStatus = gRemoveIllegalChars(sFdStatus)      '1-27-09 remove illegal char
                            sFdStatus = gFixQuote(sFdStatus)
                            'SQLQuery = "INSERT INTO SVR_7Day_Report svr (svrSeq, svrPosition, svrvefCode,  "
                            SQLQuery = "INSERT INTO " & "SVR_7Day_Report"
                            'SQLQuery = SQLQuery & " (svrSeq, svrPosition, svrvefCode,  "
                            SQLQuery = SQLQuery & " (svrSeq, svrHd1CefCode, svrvefCode,  "
                            SQLQuery = SQLQuery & "svrFt1CefCode, "
                            SQLQuery = SQLQuery & "svrProductMo, svrProductTu, svrProductWe, "
                            SQLQuery = SQLQuery & "svrProductTh, svrProductFr, svrProductSa, "
                            SQLQuery = SQLQuery & "svrProductSu, svrType, svrZone, svrAirTime, svrGenDate, svrGenTime)"
                            
                            SQLQuery = SQLQuery & " VALUES (" & rst!datShfCode & "," & rst!datAtfCode & "," & rst!datVefCode & ","
                            SQLQuery = SQLQuery & rst!datCode & ","
                            '8-7-06 remove the military time, use AM/PM times
                            'SQLQuery = SQLQuery & "'" & sFdDays & "', '" & Format$(sFdSTime, sgSQLTimeForm) & "', '" & Format$(sFdETime, sgSQLTimeForm) & "',"
                            'SQLQuery = SQLQuery & "'" & sFdStatus & "', '" & sPdDays & "' , '" & Format$(sPdSTime, sgSQLTimeForm) & "',"
                            'SQLQuery = SQLQuery & "'" & Format$(sPdETime, sgSQLTimeForm) & "', 0, '" & Format$(sFeedTime, sgSQLTimeForm) & "'," & "'" & Format$(sCurDate, sgSQLDateForm) & "' , '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sCurTime, False))))) & "')"
                            SQLQuery = SQLQuery & "'" & sFdDays & "', '" & Trim$(sFdSTime) & "', '" & Trim$(sFdETime) & "',"
                            SQLQuery = SQLQuery & "'" & sFdStatus & "', '" & sPdDays & "' ,'" & Trim$(sPdSTime) & "',"
                            SQLQuery = SQLQuery & "'" & Trim$(sPdETime) & "', 0, '" & rst!datPdDayFed & "', '" & Format$(sFeedTime, sgSQLTimeForm) & "'," & "'" & Format$(sCurDate, sgSQLDateForm) & "' , '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sCurTime, False))))) & "')"
    
                            
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "PledgeRpt-cmdReport_Click"
                                cnn.RollbackTrans
                                Exit Sub
                            End If
                            If ilRet = 0 Then
                               cnn.CommitTrans
                            End If
                        End If                  'dat.datFdStatus >= iFdStatus
                    End If
                    rst.MoveNext
                Wend
            Else                        'agreement exists, but no pledges
                If Not optSP(0).Value Then  'only show the missing pledges for all or missing options
                    ilFoundStation = False
                    If ilIncludeCodes Then      'include
                        'For ilTemp = 1 To UBound(ilStationCodes) - 1 Step 1
                        For ilTemp = LBound(ilStationCodes) To UBound(ilStationCodes) - 1 Step 1
                            If ilStationCodes(ilTemp) = AgreeRst!attshfcode Then
                                ilFoundStation = True
                                Exit For
                            End If
                        Next ilTemp
                    Else                            'exclude
                        ilFoundStation = True
                        'For ilTemp = 1 To UBound(ilStationCodes) - 1 Step 1
                        For ilTemp = LBound(ilStationCodes) To UBound(ilStationCodes) - 1 Step 1
                            If ilStationCodes(ilTemp) = AgreeRst!attshfcode Then
                                ilFoundStation = False
                                Exit For
                            End If
                        Next ilTemp
                    End If
                    
                    If (ilFoundStation) Then
                        'write out dummy record to flag it to show NO Pledges for station on report
                        sFdDays = ""
                        sFdSTime = ""
                        sFdETime = ""
                        sFdStatus = ""
                        sPdDays = ""
                        sPdSTime = ""
                        sPdETime = ""
                        sFeedTime = "00:00:00"
                        'SQLQuery = "INSERT INTO SVR_7Day_Report svr (svrSeq, svrPosition, svrvefCode,  "
                        SQLQuery = "INSERT INTO " & "SVR_7Day_Report"
                        'SQLQuery = SQLQuery & " (svrSeq, svrPosition, svrvefCode,  "
                        SQLQuery = SQLQuery & " (svrSeq, svrHd1CefCode, svrvefCode,  "
                        SQLQuery = SQLQuery & "svrFt1CefCode, "
                        SQLQuery = SQLQuery & "svrProductMo, svrProductTu, svrProductWe, "
                        SQLQuery = SQLQuery & "svrProductTh, svrProductFr, svrProductSa, "
                        SQLQuery = SQLQuery & "svrProductSu, svrType, svrAirTime, svrGenDate, svrGenTime)"
                        
                        SQLQuery = SQLQuery & " VALUES (" & AgreeRst!attshfcode & "," & AgreeRst!attCode & "," & AgreeRst!attvefCode & ","
                        '5-23-11 rst!datcode doesnt exist since theres no pledges
                        'SQLQuery = SQLQuery & rst!datcode & ","
                        SQLQuery = SQLQuery & 0 & ","
                        SQLQuery = SQLQuery & "'" & sFdDays & "', '" & Format$(sFdSTime, sgSQLTimeForm) & "', '" & Format$(sFdETime, sgSQLTimeForm) & "',"
                        SQLQuery = SQLQuery & "'" & sFdStatus & "', '" & sPdDays & "' , '" & Format$(sPdSTime, sgSQLTimeForm) & "',"
                        SQLQuery = SQLQuery & "'" & Format$(sPdETime, sgSQLTimeForm) & "', 1, '" & Format$(sFeedTime, sgSQLTimeForm) & "'," & "'" & Format$(sCurDate, sgSQLDateForm) & "' , '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sCurTime, False))))) & "')"
                                            
                        cnn.BeginTrans
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/12/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "PledgeRpt-cmdReport_Click"
                            cnn.RollbackTrans
                            Exit Sub
                        End If
                        If ilRet = 0 Then
                           cnn.CommitTrans
                        End If
                    End If
                End If
            End If          'If Not rst.EOF
            AgreeRst.MoveNext
        Wend                                        'While Not AgreeRst.EOF      'loop and process agreements
    End If                                          'Not AgreeRst.EOF Then        'test for end of retrieval
    
    'Prepare records to pass to Crystal
'    SQLQuery = "SELECT *"
'    SQLQuery = SQLQuery & " FROM VEF_Vehicles, att, "
'    '5/10/07:  Removed Affiliate Rep from Station File and added it to agreement
'9-19-11  Dan M for cr11 rollback, fix sql call
'    'SQLQuery = SQLQuery & "SVR_7Day_Report, shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode Left Outer Join artt on shttarttcode = arttcode "
'    SQLQuery = SQLQuery & "SVR_7Day_Report, shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode "
'    SQLQuery = SQLQuery + " WHERE (vefCode = svrVefCode"
'    SQLQuery = SQLQuery + " AND shttCode = svrSeq"
'    SQLQuery = SQLQuery + " AND attCode = svrHd1CefCode "
'    SQLQuery = SQLQuery + " AND svrgenDate = '" & Format$(sCurDate, sgSQLDateForm) & "' AND svrgenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sCurTime, False))))) & "')"
    SQLQuery = "SELECT * FROM SVR_7Day_Report INNER JOIN  VEF_Vehicles on svrvefCode = vefCode INNER JOIN shtt on svrSeq = shttCode" _
    & " INNER JOIN att on svrHd1CefCode = attCode LEFT OUTER JOIN mkt on shttMktCode = mktCode  LEFT OUTER JOIN  dat on svrFt1CefCode = datCode" _
    & " LEFT OUTER JOIN ept on datcode = eptdatcode"
    SQLQuery = SQLQuery & " WHERE svrgenDate = '" & Format$(sCurDate, sgSQLDateForm) & "' AND svrgenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sCurTime, False))))) & "'" '& "')"
   
    
    
    
    'removed because unable to run on s.p.5 at mai.
    '5/5/00
    'SQLQuery = SQLQuery + Chr$(13) + Chr$(10) + " ORDER BY vef.vefSort, vef.vefName, shttCallLetters, attAgreeStart, svr.svrAirTime, svr.svrProductMo"
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfPledge.rpt", "AfPledge"
    'CRpt1.SQLQuery = SQLQuery
    'CRpt1.Action = 1           'call crystal
    'CRpt1.Formulas(0) = ""
    'CRpt1.Formulas(1) = ""
    'CRpt1.Formulas(2) = ""
         
    
    'SQLQuery = "DELETE FROM SVR_7Day_Report svr WHERE (svrGenDate = '" & sCurDate & "' " & "and svrGenTime = " & sCurTime & ")"
    SQLQuery = "DELETE FROM " & "SVR_7Day_Report"
    'SQLQuery = SQLQuery & " WHERE (svrGenDate = '" & Format$(sCurDate, sgSQLDateForm) & "' " & "and svrGenTime = '" & Format$(sCurTime, sgSQLTimeForm) & "')"
    SQLQuery = SQLQuery & " WHERE (svrGenDate = '" & Format$(sCurDate, sgSQLDateForm) & "' " & "and svrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sCurTime, False))))) & "')"
    
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "PledgeRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    
    Screen.MousePointer = vbDefault
    'If optRptDest(2).Value = True Then
    '    gMsgBox "Output Sent To: " & sOutput, vbInformation
    'End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "frmPledgeRpt" & "-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmPledgeRpt
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
    gSelectiveStationsFromImport lbcStations, chkAllStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub


Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmPledgeRpt
    gCenterForm frmPledgeRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    
    imChkListBoxIgnore = False
    frmPledgeRpt.Caption = "Pledge Report - " & sgClientName
    slDate = Format$(gNow(), "m/d/yyyy")
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    CalOnAirDate.Text = Format$(slDate, sgShowDateForm)
    CalOffAirDate.Text = Format$(DateAdd("d", 6, slDate), sgShowDateForm)
'    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
'        ''grdVehAff.AddItem "" & Trim$(tgVehicleInfo(iLoop).sVehicle) & "|" & tgVehicleInfo(iLoop).iCode
'        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
'            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
'            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
'        'End If
'    Next iLoop
    chkListBox.Value = 0    'chged from false to 0 10-22-99
    lbcVehAff.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    
    lbcStations.Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            'If tgStationInfo(iLoop).iType = 0 Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            'End If
        End If
    Next iLoop

    gPopExportTypes cboFileType     '3-15-04 populate export types
    cboFileType.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmPledgeRpt = Nothing
End Sub

Private Sub grdVehAff_Click()
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub lbcStations_Click()
  If imChkAllStationsBoxIgnore Then
        Exit Sub
    End If
    If chkAllStations.Value = 1 Then
        imChkAllStationsBoxIgnore = True
        'chkListBox.Value = False
        chkAllStations.Value = 0    'chged from false to 0 10-22-99
        imChkAllStationsBoxIgnore = False
    End If
End Sub

Private Sub lbcVehAff_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        'chkListBox.Value = False
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0           '3-15-04 default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub
