VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVehVisualRpt 
   Caption         =   "Vehicle Visual Summary Report"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "AffVehVisualRpt.frx":0000
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
         Picture         =   "AffVehVisualRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Select Stations from File.."
         Top             =   2280
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   360
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
      Begin VB.ComboBox cbcSortVG 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "AffVehVisualRpt.frx":0E34
         Left            =   1200
         List            =   "AffVehVisualRpt.frx":0E36
         TabIndex        =   9
         Top             =   840
         Width           =   1500
      End
      Begin VB.Frame frcShowBy 
         Caption         =   "ShowBy"
         Height          =   585
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Visible         =   0   'False
         Width           =   2100
         Begin VB.OptionButton rbcShowBy 
            Caption         =   "Detail"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton rbcShowBy 
            Caption         =   "Summary"
            Height          =   210
            Index           =   1
            Left            =   960
            TabIndex        =   17
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.CheckBox chkAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ListBox lbcStations 
         Height          =   1620
         ItemData        =   "AffVehVisualRpt.frx":0E38
         Left            =   3120
         List            =   "AffVehVisualRpt.frx":0E3A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   2640
         Width           =   3435
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   1620
         ItemData        =   "AffVehVisualRpt.frx":0E3C
         Left            =   3120
         List            =   "AffVehVisualRpt.frx":0E3E
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   3435
      End
      Begin VB.TextBox txtNoWeeks 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Text            =   "1"
         Top             =   1920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lacVG 
         Caption         =   "Vehicle Group"
         Height          =   165
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label3 
         Caption         =   "Start Monday:"
         Height          =   165
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lacWeeks 
         Caption         =   "# Weeks:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
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
         ItemData        =   "AffVehVisualRpt.frx":0E40
         Left            =   1050
         List            =   "AffVehVisualRpt.frx":0E42
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
Attribute VB_Name = "frmVehVisualRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmVehVisualRpt - Compare all the spot feeds to the feeds
'*              specified in the pledges, to make sure that
'*              all feeds are accounted for.
'*  Created 8/18/06 D Hosaka
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imChkListBoxIgnore As Integer
Private imChkStationListIgnore As Integer


Private Type VISUAL_INFO
    iVefCode As Integer
    imnfVGCode As Integer
    sType As String * 1                 'E = embedded, R = ROS
    iStartTime(0 To 1) As Integer
    iEndTime(0 To 1) As Integer
    lStartTime As Long
    lEndTime As Long
    lMaxSec As Long                     'total length of vehicle program
    iDay As Integer
    iESTCount As Integer
    iCSTCount As Integer
    iMSTCount As Integer
    iPSTCount As Integer
    'i24Hr(1 To 24) As Integer         'each integer represents an hour starting at 12m.  -1 : show not airing, 0 : show airing but no avails, > 0: 30" avail count
    'In this report the lower bounds is not used.  All that was needed was to change the indexes from 1 to 24 to 0 to 24
    i24Hr(0 To 24) As Integer         'each integer represents an hour starting at 12m.  -1 : show not airing, 0 : show airing but no avails, > 0: 30" avail count
End Type
Private tmVisual_Info() As VISUAL_INFO



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
    Unload frmVehVisualRpt
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
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim ilNoWeeks As Integer        '# weeks requested
    Dim ilVehLoop As Integer        'vehicle loop index
    Dim ilIncludeCodes As Integer   'true to include codes in array, false to exclude codes in array
    'ReDim ilusecodes(1 To 1) As Integer   'array of codes to include or exclude
    ReDim ilusecodes(0 To 0) As Integer   'array of codes to include or exclude
    Dim slstations As String
    Dim ilLoopWeek As Integer       'week loop
    Dim ilVefCombo As Integer
    Dim ilVefCode As Integer
    Dim VehCombo_rst As ADODB.Recordset
    Dim iTemp As Integer
    Dim ilUpper As Integer
    Dim ilLoopOnDay As Integer
    Dim llStartWeek As Long
    Dim slDaysOfWk As String * 14
    Dim ilTotalATTForVeh As Integer         'total agreements for a vehicle for debug only
    Dim slReptName As String
    Dim ilLoopOnPgm As Integer
    Dim ilLoopOnBreaks As Integer
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim il30SecUnitCount As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim blFound As Boolean
    Dim ilESTCount As Integer
    Dim ilCSTCount As Integer
    Dim ilMSTCount As Integer
    Dim ilPSTCount As Integer
    Dim llMaxSec As Long
    Dim llMaxAffiliates As Long
    Dim llEarliestStartTime As Long
    Dim slKey As String
    Dim llTemp As Long
    Dim ilLoopOnPgmInx As Integer
    Dim ilSortVGBy As Integer       '0 = n/a
    
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
    sgCrystlFormula1 = sStartDate
    
    '6-13-13 option to show summary or detail (summary version is 1 line per vehicle/station/week
    If rbcShowBy(0).Value Then          'detail
        sgCrystlFormula2 = "'D'"
        slReptName = "AfVehicleVisual"
    Else
        sgCrystlFormula2 = "'S'"
        slReptName = "AfVehicleVisualSum"
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
    
    iTemp = cbcSortVG.ListIndex
    ilSortVGBy = cbcSortVG.ItemData(iTemp)
    sgCrystlFormula3 = ilSortVGBy
    
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
    
    For ilLoopWeek = 1 To ilNoWeeks         'loop on # weeks requested
        sStartDate = Format(sStartDate, "m/d/yyyy")
        llStartWeek = DateValue(sStartDate)
        sEndDate = Format(sEndDate, "m/d/yyyy")
        sDateRange = "(attOffAir >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & ") And (attDropDate >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & ") And (attOnAir <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'" & ")"
        
        For ilVehLoop = 0 To lbcVehAff.ListCount - 1 Step 1     'for each week, process each vehicle
            If lbcVehAff.Selected(ilVehLoop) Then
                ReDim tmVisual_Info(0 To 0) As VISUAL_INFO          'init the Program times per vehicle
                ReDim tgDat(0 To 0) As DAT      'gGetAVails loads tgDat array with avails
                ReDim tgPrgTimes(0 To 0) As PRGTIMES
                llMaxSec = 0
                llMaxAffiliates = 0
                ilESTCount = 0
                ilCSTCount = 0
                ilMSTCount = 0
                ilPSTCount = 0
                ilVefCode = lbcVehAff.ItemData(ilVehLoop)
                On Error GoTo ErrHand
                ilTotalATTForVeh = 0                'this is for debug only
                ilVefCombo = 0
                SQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & lbcVehAff.ItemData(ilVehLoop)
                Set VehCombo_rst = gSQLSelectCall(SQLQuery)
                If Not VehCombo_rst.EOF Then
                    ilVefCombo = VehCombo_rst!vefCombineVefCode
                End If
                'Gather agreements for the selected vehicle and stations(s)
                SQLQuery = "SELECT attCode, attshfCode, attvefCode "
                SQLQuery = SQLQuery + " FROM VEF_Vehicles, shtt, att"
                SQLQuery = SQLQuery + " WHERE vefCode = attVefCode"
                SQLQuery = SQLQuery + " AND attshfCode = shttCode "
                SQLQuery = SQLQuery + " AND (" & sDateRange & ")"
                SQLQuery = SQLQuery + " AND vefcode = " & Str(ilVefCode) & slstations

                Set AgreeRst = gSQLSelectCall(SQLQuery)
                
                'general flow:
                'Process 1 vehicle at a time:  gather active agreements for a week
                '   Obtain the Program Library from traffic and create DAT image array (tgDat contains the times of the program libray avail events)
                '   For each program defined in the week, create an entry into array tmVisual_Info which contains the hours the vehicle is airing.
                '   Max time in seconds is stored because the .rpt will sort longest program first (If program airs twice in the day, the total times of the both shows are added together)
                'all agreements for vehicle and selected stations have been retreive.
                If Not AgreeRst.EOF Then
                    'one time only for the vehicle, but the array for the number of programs and the hours they represent
                    'If AgreeRst.Fields("AttThisVehicle").Value > 0 Then     'any agreements at all?
                        gGetAvails 0, 0, ilVefCode, ilVefCombo, sStartDate, True        'att & shtt codes not required, just get program library
                        If UBound(tgDat) > 0 Then                                       'vehicle library exists
                            For ilLoopOnPgm = LBound(tgPrgTimes) To UBound(tgPrgTimes) - 1
                                blFound = False
                                For ilUpper = LBound(tmVisual_Info) To UBound(tmVisual_Info) - 1
                                    If tmVisual_Info(ilUpper).iDay = tgPrgTimes(ilLoopOnPgm).iDay Then
                                        'set the valid program times without avail count
                                        ilLo = (tgPrgTimes(ilLoopOnPgm).lPrgStartTime \ 3600) + 1
                                        ilHi = (tgPrgTimes(ilLoopOnPgm).lPrgEndTime \ 3600)
                                        If ilHi < ilLo Then
                                            ilHi = ilLo
                                        End If
                                        For iTemp = ilLo To ilHi
                                            tmVisual_Info(ilUpper).i24Hr(iTemp) = 0
                                        Next iTemp
                                        'tmVisual_Info(ilUpper).lMaxSec = tmVisual_Info(ilUpper).lMaxSec + (tgPrgTimes(ilLoopOnPgm).lPrgEndTime - tgPrgTimes(ilLoopOnPgm).lPrgStartTime)        'length in seconds of all programs for this vehicle and day
                                        llMaxSec = llMaxSec + (tgPrgTimes(ilLoopOnPgm).lPrgEndTime - tgPrgTimes(ilLoopOnPgm).lPrgStartTime)
                                        
                                        'obtain the earliest start time of the program during the week
                                        If tgPrgTimes(ilLoopOnPgm).lPrgStartTime < llEarliestStartTime Then
                                            llEarliestStartTime = tgPrgTimes(ilLoopOnPgm).lPrgStartTime
                                        End If
                                        blFound = True
                                        Exit For
                                    End If
                                Next ilUpper
                                If Not blFound Then
                                    tmVisual_Info(ilUpper).iVefCode = ilVefCode
                                    gGetVehGrpSets ilVefCode, ilSortVGBy, tmVisual_Info(ilUpper).imnfVGCode

                                    tmVisual_Info(ilUpper).iDay = tgPrgTimes(ilLoopOnPgm).iDay
                                    'tmVisual_Info(ilLoopOnPgm).lEndTime = tgPrgTimes(ilLoopOnPgm).lPrgEndTime
                                    'tmVisual_Info(ilLoopOnPgm).lSTartTime = tgPrgTimes(ilLoopOnPgm).lPrgStartTime
                                    tmVisual_Info(ilUpper).lMaxSec = tgPrgTimes(ilLoopOnPgm).lPrgEndTime - tgPrgTimes(ilLoopOnPgm).lPrgStartTime        'length in seconds of all programs for this vehicle and day
                                    llMaxSec = llMaxSec + tmVisual_Info(ilUpper).lMaxSec

                                    'init the 24 hour counts
                                    '  -1 : show not airing, 0 : show airing but no avails, > 0: 30" avail count
                                    For iTemp = 1 To 24
                                        tmVisual_Info(ilUpper).i24Hr(iTemp) = -1
                                    Next iTemp
                                    'maintain earliest start time of program
                                    llEarliestStartTime = tgPrgTimes(ilLoopOnPgm).lPrgStartTime
                                    'set the valid program times without avail count
                                    ilLo = (tgPrgTimes(ilLoopOnPgm).lPrgStartTime \ 3600) + 1
                                    ilHi = tgPrgTimes(ilLoopOnPgm).lPrgEndTime \ 3600
                                    If ilHi < ilLo Then
                                        ilHi = ilLo
                                    End If
                                    For iTemp = ilLo To ilHi
                                        tmVisual_Info(ilUpper).i24Hr(iTemp) = 0
                                    Next iTemp
                                    ReDim Preserve tmVisual_Info(0 To UBound(tmVisual_Info) + 1) As VISUAL_INFO
                                End If
                            Next ilLoopOnPgm
                            
                            'tmVisual_info() has the program start/end times for each day valid day
                            'tgDat() has the avails for each day
                            For ilLoopOnBreaks = LBound(tgDat) To UBound(tgDat) - 1
                                For ilLoopOnDay = 0 To 6
                                    If tgDat(ilLoopOnBreaks).iFdDay(ilLoopOnDay) > 0 Then
                                            For ilLoopOnPgm = LBound(tmVisual_Info) To UBound(tmVisual_Info) - 1
                                                If tmVisual_Info(ilLoopOnPgm).iDay = ilLoopOnDay Then
                                                    'determine 30" unit
                                                    llEndTime = gTimeToLong(Format$(tgDat(ilLoopOnBreaks).sFdETime, "h:mm:ssAM/PM"), False)
                                                    llStartTime = gTimeToLong(Format$(tgDat(ilLoopOnBreaks).sFdSTime, "h:mm:ssAM/PM"), False)
                                                    il30SecUnitCount = ((llEndTime) - llStartTime) \ 30
                                                    If (llEndTime - llStartTime) Mod 30 > 0 Then
                                                        il30SecUnitCount = il30SecUnitCount + 1
                                                    End If
                                                    'determine the hour this break belongs in and accumulate counts
                                                    ilLo = llStartTime \ 3600 + 1
                                                    tmVisual_Info(ilLoopOnPgm).i24Hr(ilLo) = tmVisual_Info(ilLoopOnPgm).i24Hr(ilLo) + il30SecUnitCount
                                                    Exit For
                                                End If
                                            Next ilLoopOnPgm
                                        End If
                                Next ilLoopOnDay
                            Next ilLoopOnBreaks
                        End If
                    'End If
                End If

                While Not AgreeRst.EOF      'loop and process agreements
                    'process the counts for each agreement
                    ilRet = gBinarySearchShtt(AgreeRst!attshfcode)
                    If ilRet >= 0 Then
                        sStr = UCase$(Trim$(tgShttInfo1(ilRet).shttTimeZone))
                        'each day doesnt have individual count of affiliate carrying program; it is shown on once for the week
                        If sStr = "EST" Then
                            ilESTCount = ilESTCount + 1
                        ElseIf sStr = "CST" Then
                            ilCSTCount = ilCSTCount + 1
                        ElseIf sStr = "MST" Then
                            ilMSTCount = ilMSTCount + 1
                        ElseIf sStr = "PST" Then
                            ilPSTCount = ilPSTCount + 1
                        End If
                        llMaxAffiliates = llMaxAffiliates + 1
                    End If
                    AgreeRst.MoveNext       'next agreement
                Wend
            
                'update the prepass records for this vehicle
                For ilLoopOnPgm = LBound(tmVisual_Info) To UBound(tmVisual_Info) - 1
                    'key:  max length programs(desc- all program times, 6 char), max affiliates (5 char), earliest start time of program (time in Long, 5 char), vehiclename
                    llTemp = 999999 - llMaxSec            'desc sort
                    sStr = Trim$(Str$(llTemp))
                    Do While Len(sStr) < 6
                        sStr = "0" & Trim$(sStr)
                    Loop
                    slKey = Trim$(sStr)
                    llTemp = 9999 - llMaxAffiliates
                    sStr = Trim$(Str$(llTemp))
                    Do While Len(sStr) < 4
                        sStr = "0" & Trim$(sStr)
                    Loop
                    slKey = slKey + Trim$(sStr)
                    sStr = Trim$(Str$(llEarliestStartTime))
                    Do While Len(sStr) < 5
                        sStr = "0" & Trim$(sStr)
                    Loop
                    slKey = slKey & Trim$(sStr)
                    ilRet = gBinarySearchVef(CLng(ilVefCode))
                    sStr = ""
                    If ilRet <> -1 Then
                        'sStr = Trim$(tgVehicleInfo(CLng(ilRet)).sVehicle)
                        sStr = Trim$(gRemoveChar(tgVehicleInfo(CLng(ilRet)).sVehicle, "'"))         '7-25-17 remove special char causing mvrrinsert error
                    End If
                    slKey = slKey & Trim$(sStr)
                    'ilLoopOnPgmInx = ilLoopOnPgm
                    ilRet = mInsertVVR(ilLoopOnPgm, slKey, llMaxSec, ilESTCount, ilCSTCount, ilMSTCount, ilPSTCount, sGenDate, sGenTime)
                Next ilLoopOnPgm
                
             End If                                  'vehicle selected
           
        Next ilVehLoop                              'next vehicle
        sStartDate = DateAdd("d", 7, sStartDate)      'calculate next week
        sEndDate = DateAdd("d", 6, sStartDate)      'new end week
    Next ilLoopWeek
    
    'Prepare records to pass to Crystal
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery & " From VVR_Vehicle_Visual inner join vef_vehicles on vvrvefcode = vefcode "
    SQLQuery = SQLQuery & " inner join vpf_vehicle_options on vefcode = vpfvefkcode "
    SQLQuery = SQLQuery & " left outer join mnf_multi_Names on vvrvgmnfcode = mnfcode "
    SQLQuery = SQLQuery + " where vvrGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND vvrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "'"
    SQLQuery = SQLQuery + " order by vvrKey "

    'SQLQuery = SQLQuery + "  Order by vefname, shttcallLetters, grfStartDate, grfPer1, grfPer2 "
    gUserActivityLog "E", sgReportListName & ": Prepass"
    
  
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slReptName & ".rpt", slReptName
         
    
    SQLQuery = "DELETE FROM VVR_Vehicle_Visual "
    SQLQuery = SQLQuery & " WHERE (vvrGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and vvrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"

    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "VehVisualRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    rst.Close
    AgreeRst.Close

    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    Screen.MousePointer = vbDefault

    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "VehVisualRpt-cmdReport"
    Exit Sub            'terminate
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmVehVisualRpt
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
    lbcStations.Top = chkAllStations.Top + chkAllStations.Height
    cmdStationListFile.Top = chkAllStations.Top - 50

    gSetFonts frmVehVisualRpt
    gCenterForm frmVehVisualRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilUseNone As Integer
    
    imChkListBoxIgnore = False
    imChkStationListIgnore = False
    frmVehVisualRpt.Caption = "Vehicle Visual Summary Report - " & sgClientName
    slDate = Format$(gNow(), "m/d/yyyy")
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    CalOnAirDate.Text = Format$(slDate, sgShowDateForm)
    'txtOffAirDate.Text = Format$(DateAdd("d", 6, slDate), sgShowDateForm)
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    chkListBox.Value = 0    'chged from false to 0 10-22-99
    lbcVehAff.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
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
    
    'gPopVehicleGroups frmVehVisualRpt!cbcSortVG, tgVehicleSets1(), True
    ilRet = gPopShttInfo()
    ilUseNone = True                   'VGselection is optional
    gPopVehicleGroups cbcSortVG, ilUseNone
    gPopExportTypes cboFileType     '3-15-04 populate export types
    cboFileType.Enabled = False
    ilRet = gPopAvailNames()
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmVisual_Info
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmVehVisualRpt = Nothing
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
'
Public Function mInsertVVR(ilLoopOnPgm As Integer, slKey As String, llMaxSec As Long, ilESTCount As Integer, ilCSTCount As Integer, ilMSTCount As Integer, ilPSTCount As Integer, sGenDate As String, sGenTime As String) As Integer
Dim ilRet As Integer

    On Error GoTo ErrHand
    mInsertVVR = 0
    
    SQLQuery = "INSERT INTO " & "VVR_Vehicle_Visual "
    SQLQuery = SQLQuery & " (vvrcode, vvrvefcode, vvrKey, vvrMaxSec, vvrDay, vvrVGMnfCode, "
    SQLQuery = SQLQuery & " vvrESTCount, vvrCSTCount, vvrMSTCount, vvrPSTCount, "
    SQLQuery = SQLQuery & " vvr12m, vvr1a, vvr2a, vvr3a, vvr4a, vvr5a, "
    SQLQuery = SQLQuery & " vvr6a, vvr7a, vvr8a, vvr9a, vvr10a, vvr11a, "
    SQLQuery = SQLQuery & " vvr12n, vvr1p, vvr2p, vvr3p, vvr4p, vvr5p, "
    SQLQuery = SQLQuery & " vvr6p, vvr7p, vvr8p, vvr9p, vvr10p, vvr11p, "
    SQLQuery = SQLQuery & "  vvrGendate, vvrGenTime) "
    
    SQLQuery = SQLQuery & "VALUES (" & 0 & ", " & tmVisual_Info(ilLoopOnPgm).iVefCode & ", '" & slKey & "', " & llMaxSec & ", " & tmVisual_Info(ilLoopOnPgm).iDay & ", " & tmVisual_Info(ilLoopOnPgm).imnfVGCode & ", "
    SQLQuery = SQLQuery & ilESTCount & ", " & ilCSTCount & ", " & ilMSTCount & ", " & ilPSTCount & ", "
    SQLQuery = SQLQuery & tmVisual_Info(ilLoopOnPgm).i24Hr(1) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(2) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(3) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(4) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(5) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(6) & ", "
    SQLQuery = SQLQuery & tmVisual_Info(ilLoopOnPgm).i24Hr(7) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(8) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(9) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(10) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(11) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(12) & ", "
    SQLQuery = SQLQuery & tmVisual_Info(ilLoopOnPgm).i24Hr(13) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(14) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(15) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(16) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(17) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(18) & ", "
    SQLQuery = SQLQuery & tmVisual_Info(ilLoopOnPgm).i24Hr(19) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(20) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(21) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(22) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(23) & ", " & tmVisual_Info(ilLoopOnPgm).i24Hr(24) & ", "
    SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"   '", "
    ilRet = 0
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "VehVisualRpt-mInsertVVR"
        cnn.RollbackTrans
        mInsertVVR = Err.Number
        Exit Function
    End If
    If ilRet = 0 Then
       cnn.CommitTrans
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "VehVisualRpt-mInsertVVR"
    mInsertVVR = Err.Number
    Exit Function
End Function
