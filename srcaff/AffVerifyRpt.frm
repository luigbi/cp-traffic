VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVerifyRpt 
   Caption         =   "Feed Verification Report"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "AffVerifyRpt.frx":0000
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
      Height          =   4380
      Left            =   240
      TabIndex        =   6
      Top             =   1845
      Width           =   6705
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   6120
         Picture         =   "AffVerifyRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Select Stations from File.."
         Top             =   2280
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1350
         TabIndex        =   7
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Text            =   "12/13/2022"
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
      Begin VB.Frame frcShowBy 
         Caption         =   "ShowBy"
         Height          =   585
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   2100
         Begin VB.OptionButton rbcShowBy 
            Caption         =   "Detail"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton rbcShowBy 
            Caption         =   "Summary"
            Height          =   210
            Index           =   1
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.CheckBox chkAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ListBox lbcStations 
         Height          =   1620
         ItemData        =   "AffVerifyRpt.frx":0E34
         Left            =   3120
         List            =   "AffVerifyRpt.frx":0E36
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   2640
         Width           =   3435
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1620
         ItemData        =   "AffVerifyRpt.frx":0E38
         Left            =   3120
         List            =   "AffVerifyRpt.frx":0E3A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   480
         Width           =   3435
      End
      Begin VB.TextBox txtNoWeeks 
         Height          =   285
         Left            =   1350
         TabIndex        =   10
         Text            =   "1"
         Top             =   660
         Width           =   375
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Start Monday:"
         Height          =   225
         Left            =   225
         TabIndex        =   8
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lacWeeks 
         Caption         =   "# Weeks:"
         Height          =   255
         Left            =   225
         TabIndex        =   9
         Top             =   720
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4590
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4350
      TabIndex        =   14
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4155
      TabIndex        =   13
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
         ItemData        =   "AffVerifyRpt.frx":0E3C
         Left            =   1050
         List            =   "AffVerifyRpt.frx":0E3E
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
Attribute VB_Name = "frmVerifyRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmVerifyRpt - Compare all the spot feeds to the feeds
'*              specified in the pledges, to make sure that
'*              all feeds are accounted for.
'*  Created 8/18/06 D Hosaka
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer
Private imChkStationListIgnore As Integer

Private Type FEEDVERIFY
    iStatus As Integer                  '0 = N/a, 1 = avail missing, 2 = avail added
    iFeedDay As Integer                 'day missing or added
    sSoldTime As String                  'sold avail time
    sFeedTime As String                   'user defined fed time
    iDACode As Integer                  '0 = DP, 1 = avails, 2 = CD/tape
End Type
Private tmFeedVerify() As FEEDVERIFY

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
Private Sub chkAllStations_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkStationListIgnore Then
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
        imChkStationListIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkStationListIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload frmVerifyRpt
End Sub
Private Sub cmdReport_Click()
    Dim i As Integer
    Dim sCode As String
    Dim sDateRange As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim AgreeRst As ADODB.Recordset
    Dim sGenDate As String
    Dim sGenTime As String
    Dim sStr As String
    Dim sFdDays As String
    Dim sFdSTime As String
    Dim sFdStatus As String
    Dim sPdDays As String
    Dim sPdSTime As String
    Dim sPdETime As String
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim ilNoWeeks As Integer        '# weeks requested
    Dim ilVehLoop As Integer        'vehicle loop index
    Dim ilIncludeCodes As Integer   'true to include codes in array, false to exclude codes in array
    ''ReDim ilusecodes(1 To 1) As Integer   'array of codes to include or exclude
    'ReDim ilusecodes(0 To 0) As Integer   'array of codes to include or exclude
    ReDim ilusecodes(0 To 0) As Integer   'array of codes to include or exclude
    Dim slstations As String
    Dim ilLoopWeek As Integer       'week loop
    Dim ilVefCombo As Integer
    Dim VehCombo_rst As ADODB.Recordset
    Dim iTemp As Integer
    Dim ilSoldAvails As Integer
    Dim llSoldTime As Long          'avail time sold from traffic
    Dim llFeedTime As Long          'avail time that agreemnt is accepting feed
    Dim ilUpper As Integer
    Dim ilLoopOnDay As Integer
    Dim ilFeedDays(0 To 6) As Integer
    Dim sDiscrep As String
    Dim llStartWeek As Long
    Dim slDaysOfWk As String * 14
    Dim slFeedType As String * 1            'udpated in GRF to indicate avail or cd/tape feed
    Dim ilFoundMatch As Integer             'equal day and avail time found between traffic sold avail and affiliate fed avail
    Dim ilAnyAffFeedTimes As Integer        'flag to indicate if any affiliate DAT records found;  if not, no discrepancy, its a DP feed
    Dim llSortDate As Long                  'updated in GRF for crystal sorting
    Dim llSortTime As Long                  'updated in GRF for crystal sorting
    Dim ilTotalATTForVeh As Integer         'total agreements for a vehicle for debug only
    Dim slReptName As String
    'Dim NewForm As New frmViewReport
    
    Dim blDiscrepFound As Boolean              'for summary version; only write out one discrep record so the vehicle/station/week can be shown on report
    Dim blRstUsed As Boolean                    'flag to determine to close RST
    
    On Error GoTo ErrHand
    slDaysOfWk = "MoTuWeThFrSaSu"
    sStartDate = Trim$(CalOnAirDate.Text)
    If sStartDate = "" Then
        sStartDate = "1/1/1970"
    End If

    'date must be a monday
    If gIsDate(sStartDate) = False Or Weekday(sStartDate, vbSunday) <> vbMonday Then
        Beep
        gMsgBox "Please enter a valid Monday date (m/d/yy)", vbCritical
        CalOnAirDate.SetFocus
        Exit Sub
    End If
    
    sCode = Trim$(txtNoWeeks.Text)
    If Not IsNumeric(sCode) Then
        gMsgBox "Invalid # weeks", vbOKOnly
        txtNoWeeks.SetFocus
        Exit Sub
    End If
    ilNoWeeks = Val(txtNoWeeks.Text)
    
    sEndDate = DateAdd("d", (ilNoWeeks * 7) - 1, sStartDate)  'calculate end date
    sgCrystlFormula1 = sStartDate & "-" & sEndDate
    
    
    If lbcStations.SelCount = 0 Or lbcVehAff.SelCount = 0 Then
        Beep
        gMsgBox "At least 1 vehicle and station must be selected"
        Exit Sub
    End If
    
    '6-13-13 option to show summary or detail (summary version is 1 line per vehicle/station/week
    If rbcShowBy(0).Value Then          'detail
        sgCrystlFormula2 = "'D'"
        slReptName = "AfVerify"
    Else
        sgCrystlFormula2 = "'S'"
        slReptName = "AfVerifySum"
    End If
    
    Screen.MousePointer = vbHourglass
    
    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
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
               
    sGenDate = Format(gNow(), sgShowDateForm)    'current date and time used as key for prepass file to
                                              'access and clear
    sGenTime = Format(gNow(), sgShowTimeWSecForm)
    
    gObtainCodes lbcStations, ilIncludeCodes, ilusecodes()        'build array of which codes to incl/excl
    slstations = ""
    If ilIncludeCodes = True Then
        If Not chkAllStations.Value = vbChecked Then    'User did NOT select all vehicles
            For i = LBound(ilusecodes) To UBound(ilusecodes) - 1 Step 1
                If ilusecodes(i) Then
                    If Len(slstations) = 0 Then
                        slstations = " AND ((shttCode = " & ilusecodes(i) & ")"
                    Else
                        slstations = slstations & " OR (shttCode = " & ilusecodes(i) & ")"
                    End If
                End If
            Next i
            If Len(slstations) > 0 Then
                slstations = slstations & ")"
            End If
        End If
    Else
        For i = LBound(ilusecodes) To UBound(ilusecodes) - 1 Step 1
            If ilusecodes(i) Then
                If Len(slstations) = 0 Then
                    slstations = "AND ((shttCode <> " & ilusecodes(i) & ")"
                Else
                    slstations = slstations & " AND (shttCode <> " & ilusecodes(i) & ")"
                End If
            End If
        Next i
        If Len(slstations) > 0 Then
            slstations = slstations & ")"
        End If
    End If
    blRstUsed = False
    For ilLoopWeek = 1 To ilNoWeeks         'loop on # weeks requested
        sStartDate = Format(sStartDate, "m/d/yyyy")
        llStartWeek = DateValue(sStartDate)
        sEndDate = Format(sEndDate, "m/d/yyyy")
        sDateRange = "(attOffAir >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & ") And (attDropDate >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & ") And (attOnAir <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'" & ")"
        
        For ilVehLoop = 0 To lbcVehAff.ListCount - 1 Step 1     'for each week, process each vehicle
            If lbcVehAff.Selected(ilVehLoop) Then
                On Error GoTo ErrHand
                ilTotalATTForVeh = 0                'this is for debug only
                ilVefCombo = 0
                SQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & lbcVehAff.ItemData(ilVehLoop)
                Set VehCombo_rst = gSQLSelectCall(SQLQuery)
                If Not VehCombo_rst.EOF Then
                    ilVefCombo = VehCombo_rst!vefCombineVefCode
                End If
                'Gather agreements for the selected vehicle and stations(s)
                SQLQuery = "SELECT attCode, attshfCode, attvefCode, attPledgeType"
                SQLQuery = SQLQuery + " FROM VEF_Vehicles, shtt, att"
                SQLQuery = SQLQuery + " WHERE vefCode = attVefCode"
                SQLQuery = SQLQuery + " AND attshfCode = shttCode "
                SQLQuery = SQLQuery + " AND (" & sDateRange & ")"
                SQLQuery = SQLQuery + " AND vefcode = " & Str(lbcVehAff.ItemData(ilVehLoop)) & slstations
                Set AgreeRst = gSQLSelectCall(SQLQuery)
                'all agreements for vehicle and selected stations have been retreive.
                'cycle thru the agreements and obtain their sold avails to see if all stations
                'pledge times have been accounted for

                If Not AgreeRst.EOF Then        'test for end of retrieval
                    While Not AgreeRst.EOF      'loop and process agreements
                        blDiscrepFound = False
                        
                        ReDim tgDat(0 To 0) As DAT      'gGetAVails loads tgDat array with avails
                        'get the sold avails so it can be matched against the fed avails already defined and build
                        'into array TGDAT()
                        gGetAvails AgreeRst!attCode, AgreeRst!attshfcode, AgreeRst!attvefCode, ilVefCombo, sStartDate, True
                        
                        mBuildDailySoldAvails AgreeRst!attPledgeType, ilUpper       'build array of avails and days into TmFeedVerify()
                        
                        If AgreeRst!attPledgeType <> "D" Then
                            blRstUsed = True
                            'gather all pledge information for this agreement from DAT.mkd
                            mCreateSQLCallForDAT AgreeRst!attCode, AgreeRst!attshfcode, AgreeRst!attvefCode
                            On Error GoTo ErrHand
                            Set rst = gSQLSelectCall(SQLQuery)
                            
                            'are there any affiliate feed times?  (any DAT records?)
                            ilAnyAffFeedTimes = False
                            Do While Not rst.EOF
                                ilAnyAffFeedTimes = True
                                ''the DAT has the type of feed (DP, avails, CD/tape)
                                'If rst!datDACode = 1 Then           'avails
                                '    slFeedType = "A"
                                'ElseIf rst!datDACode = 2 Then       'cd/tape
                                '    slFeedType = "C"
                                'Else
                                '    slFeedType = Str$(rst!datDACode)
                                'End If
                                If AgreeRst!attPledgeType = "A" Then
                                    slFeedType = "A"
                                ElseIf AgreeRst!attPledgeType = "C" Then
                                    slFeedType = "C"
                                Else
                                    slFeedType = "0"
                                End If
                             
                                'put affiliate feed days in array for looping
                                'RST = avail times from DAT
                                ilFeedDays(0) = rst!datFdMon
                                ilFeedDays(1) = rst!datFdTue
                                ilFeedDays(2) = rst!datFdWed
                                ilFeedDays(3) = rst!datFdThu
                                ilFeedDays(4) = rst!datFdFri
                                ilFeedDays(5) = rst!datFdSat
                                ilFeedDays(6) = rst!datFdSun
                                  
                                'convert the feed time to string and truncate seconds if they dont exist
                                If Second(rst!datFdStTime) = 0 Then
                                    sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWOSecForm)
                                Else
                                    sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWSecForm)
                                End If
                                
                                llFeedTime = gTimeToLong((rst!datFdStTime), False)                      'feed time from affiliate DAT avails
                                
                                ilFoundMatch = False         'assume no matching day & time found
                                For ilLoopOnDay = 0 To 6    'loop of affiliate feed days of week, search for matching avail day & time
                                    If ilFeedDays(ilLoopOnDay) > 0 Then
                                        For ilSoldAvails = LBound(tmFeedVerify) To UBound(tmFeedVerify) - 1
                                            
                                            'convert times to long for time testing
                                            llSoldTime = gTimeToLong(tmFeedVerify(ilSoldAvails).sSoldTime, False)   'sold time from traffic avails
                                     
'                                            If llFeedTime = llSoldTime And ilFeedDays(ilLoopOnDay) > 0 And tmFeedVerify(ilSoldAvails).iFeedDay = ilLoopOnDay + 1 And tmFeedVerify(ilSoldAvails).iStatus <> 1 Then            'got a match
                                            If llFeedTime = llSoldTime And ilFeedDays(ilLoopOnDay) > 0 And tmFeedVerify(ilSoldAvails).iFeedDay = ilLoopOnDay + 1 Then                   '4-5-17 got a match
                                                If ilFeedDays(ilLoopOnDay) > 0 And tmFeedVerify(ilSoldAvails).iFeedDay = ilLoopOnDay + 1 Then    'Affiliate feed day must be an airing day,                                                                                      'and the traffic sold day must be an airing day
                                                    'time and day matches
                                                    If tmFeedVerify(ilSoldAvails).iStatus <> 2 Then     '4-5-17
                                                        tmFeedVerify(ilSoldAvails).iStatus = 1      'processed and matching
                                                    End If
                                                    ilFoundMatch = True
                                                    Exit For
                                                End If
                                            End If
                                        
                                        Next ilSoldAvails
                                        If Not ilFoundMatch Then
                                            tmFeedVerify(ilUpper).iStatus = 2          'flag as added avail
                                            tmFeedVerify(ilUpper).iFeedDay = ilLoopOnDay + 1      'day of week added
                                            tmFeedVerify(ilUpper).sFeedTime = sFdSTime      'affiliate feed time
                                            tmFeedVerify(ilUpper).sSoldTime = sFdSTime  '""               'no traffic sold time
                                            ReDim Preserve tmFeedVerify(0 To ilUpper + 1) As FEEDVERIFY
                                            blDiscrepFound = True
                                            ilUpper = ilUpper + 1
                                            If rbcShowBy(1).Value Then                'summary only, once a discrepancy is found, get out for speed.  only need to know the station, vehicle, week
                                                Exit Do
                                            End If
                                        End If
                                    End If              'ilfeeddays > 0
    
                                Next ilLoopOnDay
                                
                                rst.MoveNext
                            'Wend
                            Loop
                        Else
                            ilAnyAffFeedTimes = False
                        End If
                        If ilAnyAffFeedTimes Then          'if no feed times found, then its a DP feed or not defined yet.  No discreps to show
                            For ilSoldAvails = LBound(tmFeedVerify) To UBound(tmFeedVerify) - 1
                                If tmFeedVerify(ilSoldAvails).iStatus <> 1 Then      '1 = processed and avails/days match, 0 = day & avail missing, 2 = day & avail added
                                    sDiscrep = Mid$(slDaysOfWk, (tmFeedVerify(ilSoldAvails).iFeedDay - 1) * 2 + 1, 2)
                                    sDiscrep = sDiscrep & " " & Format$(llStartWeek + (tmFeedVerify(ilSoldAvails).iFeedDay - 1), "m/d/yy")
                                    llSortDate = llStartWeek + (tmFeedVerify(ilSoldAvails).iFeedDay - 1)        'update date as long for sorting in cyrstal
                                   
                                    If tmFeedVerify(ilSoldAvails).iStatus = 0 Then
                                        sDiscrep = sDiscrep & " @" & Trim$(tmFeedVerify(ilSoldAvails).sSoldTime) & " Missing"       'traffic sold avail missing
                                        llSortTime = gTimeToLong(Trim$(tmFeedVerify(ilSoldAvails).sSoldTime), False)        'update time as long for sorting in crystal
                                    Else
                                        sDiscrep = sDiscrep & " @" & Trim$(tmFeedVerify(ilSoldAvails).sFeedTime) & " Extra"         'affiliate fed time added
                                        llSortTime = gTimeToLong(Trim$(tmFeedVerify(ilSoldAvails).sFeedTime), False)      'update time as long for sorting in crystal
                                    End If

                                    ilRet = mInsertGRFDiscrep(AgreeRst!attvefCode, AgreeRst!attshfcode, sDiscrep, sStartDate, slFeedType, llSortDate, llSortTime, sGenDate, sGenTime)
                                    If ilRet <> 0 Then          'error from insert
                                        Exit Sub
                                    End If

                                End If
                                If (rbcShowBy(1).Value = True And blDiscrepFound) Then          'summary only and 1 discrepancy found, no need to continue since only showing vehicle/station /week on output
                                    Exit For
                                End If
                            Next ilSoldAvails
                        End If
                       
                        ilTotalATTForVeh = ilTotalATTForVeh + 1     'for debug only
                        AgreeRst.MoveNext       'next agreement
                    Wend
                End If
            End If                                  'vehicle selected
        Next ilVehLoop                              'next vehicle
        sStartDate = DateAdd("d", 7, sStartDate)      'calculate next week
        sEndDate = DateAdd("d", 6, sStartDate)      'new end week
    Next ilLoopWeek
    
    'Prepare records to pass to Crystal
    SQLQuery = "SELECT *"
    'SQLQuery = SQLQuery & " FROM VEF_Vehicles, shtt, "
    SQLQuery = SQLQuery & " From GRF_Generic_Report inner join vef_vehicles on grfvefcode = vefcode inner join shtt on grfsofcode = shttcode "
    'SQLQuery = SQLQuery + " WHERE (vefCode = grfvefCode"
    'SQLQuery = SQLQuery + " AND shttCode = grfsofCode"
    'SQLQuery = SQLQuery + " AND grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
    SQLQuery = SQLQuery + " where grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'"

    'SQLQuery = SQLQuery + "  Order by vefname, shttcallLetters, grfStartDate, grfPer1, grfPer2 "
    gUserActivityLog "E", sgReportListName & ": Prepass"
    
  
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slReptName & ".rpt", slReptName
         
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    
    SQLQuery = "DELETE FROM GRF_Generic_Report"
    SQLQuery = SQLQuery & " WHERE (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"

    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "VerifyRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans

    If blRstUsed Then
        rst.Close
    End If
    AgreeRst.Close
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmVerifyRpt-"
    Exit Sub            'terminate
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmVerifyRpt
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
Dim ilHalf As Integer
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    chkListBox.Top = 240
    
    ilHalf = (Frame2.Height - chkListBox.Height - chkAllStations.Height - 480) / 2
    lbcVehAff.Move chkListBox.Left, chkListBox.Top + chkListBox.Height
    lbcVehAff.Height = ilHalf
    lbcStations.Height = ilHalf
    chkAllStations.Top = lbcVehAff.Top + lbcVehAff.Height + 120
    'TTP 9943
    cmdStationListFile.Top = chkAllStations.Top - 50
    lbcStations.Top = chkAllStations.Top + chkAllStations.Height

    gSetFonts frmVerifyRpt
    gCenterForm frmVerifyRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    
    imChkListBoxIgnore = False
    imChkStationListIgnore = False
    frmVerifyRpt.Caption = "Feed Verification Report - " & sgClientName
    slDate = Format$(gNow(), "m/d/yyyy")
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    CalOnAirDate.Text = Format$(slDate, sgShowDateForm)
    'txtOffAirDate.Text = Format$(DateAdd("d", 6, slDate), sgShowDateForm)
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        ''grdVehAff.AddItem "" & Trim$(tgVehicleInfo(iLoop).sVehicle) & "|" & tgVehicleInfo(iLoop).iCode
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then        '4-27-09 chged to test OLA only
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    chkListBox.Value = 0    'chged from false to 0 10-22-99
    lbcVehAff.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then    '4-27-09 chged to test OLA only
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    
    'populate the Stations, Vehicles & Advertisers (currently only advertisers are selectable)
    lbcStations.Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop

    gPopExportTypes cboFileType     '3-15-04 populate export types
    cboFileType.Enabled = False
    ilRet = gPopAvailNames()
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmFeedVerify
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmVerifyRpt = Nothing
End Sub
Private Sub lbcStations_Click()
    If imChkStationListIgnore Then
        Exit Sub
    End If
    If chkAllStations.Value = 1 Then
        imChkStationListIgnore = True
        'chkListBox.Value = False
        chkAllStations.Value = 0    'chged from false to 0 10-22-99
        imChkStationListIgnore = False
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
'*        Build  array of sold avails with entry time defined with its own day
'*        Output will produce individual lines for each day/time missing or added
'*        i.e. m-f 6AM avail will be 5 events in the array
'*
Public Sub mBuildDailySoldAvails(slPledgeType As String, ilUpper As Integer)
Dim ilSoldAvails As Integer
Dim iTemp As Integer
        ilUpper = 0
        ReDim tmFeedVerify(0 To 0) As FEEDVERIFY
        For ilSoldAvails = LBound(tgDat) To UBound(tgDat) - 1
            For iTemp = 0 To 6
                If tgDat(ilSoldAvails).iFdDay(iTemp) > 0 Then
                    tmFeedVerify(ilUpper).iStatus = 0
                    tmFeedVerify(ilUpper).iFeedDay = iTemp + 1 'make 1 = mo, 2 = tu, 7 = su
                    tmFeedVerify(ilUpper).sSoldTime = tgDat(ilSoldAvails).sFdSTime
                    tmFeedVerify(ilUpper).sFeedTime = ""                               'time of defined fed time when its found from DAT
                    'tmFeedVerify(ilUpper).iDACode = tgDat(ilSoldAvails).iDACode     '0 = DP, 1 = avails, 2 = cd/tape
                    If slPledgeType = "A" Then
                        tmFeedVerify(ilUpper).iDACode = 1     '0 = DP, 1 = avails, 2 = cd/tape
                    ElseIf slPledgeType = "C" Then
                        tmFeedVerify(ilUpper).iDACode = 2     '0 = DP, 1 = avails, 2 = cd/tape
                    Else
                        tmFeedVerify(ilUpper).iDACode = 0     '0 = DP, 1 = avails, 2 = cd/tape
                    End If
                    ReDim Preserve tmFeedVerify(0 To ilUpper + 1) As FEEDVERIFY
                    ilUpper = ilUpper + 1
                End If
            Next iTemp
        Next ilSoldAvails
End Sub
'
'           Setup SQL Call to gather all DAT avails for the agreement
'           <input>  llAttCode - agreement code
'                    llShfCode - station code
'                    ilvefcode = vehicle code
'
Public Sub mCreateSQLCallForDAT(llAttCode As Long, llShfCode As Long, ilVefCode As Integer)
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM dat"
    SQLQuery = SQLQuery + " WHERE (datAtfCode = " & llAttCode
    SQLQuery = SQLQuery + " AND datShfCode= " & llShfCode '& " AND datDaCode > 0 "  'DACode moved to agreement (attPledgeType)
    SQLQuery = SQLQuery + " AND datVefCode = " & ilVefCode & ")"
    SQLQuery = SQLQuery & " ORDER BY datFdStTime"
End Sub
'
'        insert avails found to be discrepant into GRf for output
'       <input> ilvefcode - vehicle code
'               llshfcode - station code
'               sDiscrep - text of discrepancy (i.e. Mo 8/1/06 @ 6a Missing)
'               sStartDate - string indicating start monday processed
'               slFeedType - A = avails, C = cd/tape
'               llSortDate - date of discrepany used for sorting time within date incrystal
'               llSorttime - time of discrepancy used for sorting time within date in crystal
Public Function mInsertGRFDiscrep(ilVefCode As Integer, llShfCode As Long, sDiscrep As String, sStartDate As String, slFeedType As String, llSortDate As Long, llSortTime As Long, sGenDate As String, sGenTime As String) As Integer
Dim ilRet As Integer
Dim ilStartDate(0 To 1) As Integer

    On Error GoTo ErrHand
    mInsertGRFDiscrep = 0
    SQLQuery = "INSERT INTO " & "GRF_Generic_Report"
    SQLQuery = SQLQuery & " (grfvefcode, grfsofcode, grfgendesc, "
    SQLQuery = SQLQuery & " grfStartDate, grfBktType, grfPer1, grfPer2, "
    SQLQuery = SQLQuery & "  grfGendate, grfGenTime) "
    SQLQuery = SQLQuery & "VALUES ('" & ilVefCode & "', " & "'" & llShfCode & "', " & "'" & sDiscrep & "', "
    SQLQuery = SQLQuery & "'" & Format$(sStartDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & slFeedType & "', " & "'" & llSortDate & "', " & "'" & llSortTime & "', "
    SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"   '", "
    ilRet = 0
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "VehVisualRpt-mInsertGRFDiscrep"
        cnn.RollbackTrans
        mInsertGRFDiscrep = Err.Number
        Exit Function
    End If
    If ilRet = 0 Then
       cnn.CommitTrans
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmVerifyRpt-mInsertGRFDiscrep"
    mInsertGRFDiscrep = Err.Number
    Exit Function
End Function
