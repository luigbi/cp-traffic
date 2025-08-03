VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmExportOverdue 
   Caption         =   "Affidavit Overdue Export"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   Icon            =   "AffExportOverdue.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   8025
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   720
      Top             =   6360
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6855
      FormDesignWidth =   8025
   End
   Begin VB.Frame Frame2 
      Caption         =   "Affidavit Overdue Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   7545
      Begin VB.Frame frcPostType 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   840
         TabIndex        =   12
         Top             =   3800
         Width           =   2985
      End
      Begin V81Affiliate.CSI_Calendar CalFromDate 
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   510
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         Text            =   "8/19/2021"
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
         CSI_DefaultDateType=   0
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2010
         TabIndex        =   9
         Top             =   2440
         Width           =   1545
      End
      Begin VB.Frame frcPassword 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   615
         TabIndex        =   4
         Top             =   3525
         Width           =   2145
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   3960
         ItemData        =   "AffExportOverdue.frx":08CA
         Left            =   3840
         List            =   "AffExportOverdue.frx":08CC
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   690
         Width           =   3495
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
      Begin V81Affiliate.CSI_Calendar CalToDate 
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   510
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         Text            =   "8/19/2021"
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
         CSI_CurDayForeColor=   51200
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.Frame Frame5 
         Caption         =   "Dates"
         Height          =   840
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   3540
         Begin VB.Label LabTo 
            Caption         =   "End:"
            Height          =   240
            Left            =   1800
            TabIndex        =   2
            Top             =   360
            Width           =   450
         End
         Begin VB.Label labFrom 
            Caption         =   "Start"
            Height          =   240
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   570
         End
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Generate "
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lacResult 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "frmExportOverdue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmExportOverdue  - Export Overdue affidavits
'*
'*  Created July,1998 by Dick LeVine
'*  Modified May, 2000 by D. Smith
'*
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'Private Const cMaxDates = 11 ' max number of dates to print on a line in the report
'Private Const cMaxDates = 8 ' max number of dates to print on a line in the report
'Private Const cMaxDates = 7 ' max number of dates to print on a line in the report
Private Const cMaxDates = 6 ' 12-3-10 max number of dates to print on a line in the report

Private Type OVERDUE_AFFS
    iVefCode As Integer
    sVefName As String
    iShttCode As Integer
    sCallLetters As String
    sDMAName As String
    sDMARank As String
    lStationID As Long
    lCount As Integer           'count of outstanding affidavits
End Type

Dim tmOverdue_Affs() As OVERDUE_AFFS
Dim tmUpToDate_Affs() As OVERDUE_AFFS

Private imChkListBoxIgnore As Integer
Dim imNoDaysDelq As Integer         'from site - # weeks before CP overdue
Dim imConsecutiveWksNCR As Integer
Dim lmDefaultStartDate As Long
Dim lmDefaultEndDate As Long
Dim smDefaultStartDate As String
Dim smDefaultEndDate As String
Dim smClientName As String
Private smExportName As String
Private hmExport As Integer
Private bmFirstTime As Boolean
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imExporting As Integer  'True = exporting in progress
Dim ilLoaded As Integer    '0=loading, 1=loaded

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
Private Function OpenExportFile(sStartDate As String, sEndDate As String) As Integer
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim sLetter As String
    Dim sStartYear As String
    Dim sStartMonth As String
    Dim sStartDay As String
    Dim sEndYear As String
    Dim sEndMonth As String
    Dim sEndDay As String

    'On Error GoTo OpenExportFileErr:
    
    gObtainYearMonthDayStr sStartDate, True, sStartYear, sStartMonth, sStartDay
    gObtainYearMonthDayStr sEndDate, True, sEndYear, sEndMonth, sEndDay
   

    sLetter = ""
    Do
        ilRet = 0
        'smToFile = sgExportDirectory & "D" & Format$(gNow(), "mm") & Format$(gNow(), "dd") & Format$(gNow(), "yy") & sLetter & ".csv"
        'slDateTime = FileDateTime(smToFile)
        smExportName = sgExportDirectory & smClientName + " Overdue " + sStartMonth + sStartDay + sStartYear + "-" + sEndMonth + sEndDay + sEndYear + Trim$(sLetter) + ".csv"
        'slDateTime = FileDateTime(smExportName)
        ilRet = gFileExist(smExportName)            '7-11-17
        If ilRet = 0 Then
            If Trim$(sLetter) = "" Then
                sLetter = "A"
            Else
                sLetter = Chr$(Asc(sLetter) + 1)
            End If
        End If
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo OpenExportFileErr:
    'hmExport = FreeFile
    'Open smExportName For Output As hmExport
    ilRet = gFileOpen(smExportName, "Output", hmExport)
    If ilRet <> 0 Then
        Close hmExport
        hmExport = -1
        gMsgBox "Open File " & smExportName & " error#" & Str$(Err.Number), vbOKOnly
        OpenExportFile = False
        Exit Function
    End If
    lacResult.Caption = "Generating: " & Trim$(smExportName)
    'lacResult.Visible = False
    Exit Function
'OpenExportFileErr:
'
'    ilRet = 1
'    Resume Next
End Function

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
    mSetGenerate
    Screen.MousePointer = vbDefault

End Sub
Private Sub cmdDone_Click()
    Dim ilMsgbox As Integer
    If cmdDone.Caption = "Cancel" Then
        If imExporting = True Then
            ilMsgbox = MsgBox("Cancel Export?", vbQuestion + vbYesNo, "Export Overdue")
            If ilMsgbox = vbYes Then imTerminate = True
        Else
            Unload frmExportOverdue
            Exit Sub
        End If
    End If
    If cmdDone.Caption = "Done" Then Unload frmExportOverdue
    
End Sub

'10-21-16 Generate Export file of Overdue Affidavits with counts by vehicle and station
'Generate export of Overdue affidavits from user - selected vehicles and dates.
'Include all vehicle/stations as long as there overdue affidavits.  Include vehicle/stations if active but no overdue
'Count of overdue affidavits will be 0
'export will include 4 fields:  vehicle, station, station id, and count of overdue affidavits.
'Date selectivity could go from beginning of time to the current weeks previous Sunday date.
'
Private Sub cmdExport_Click()
    Dim iTtlUnqDates, iIsFirst, iNumDates As Integer
    'change iArrayIndx from integer to long
    Dim lArrayIdx As Long
    Dim i, iRet, iDateIdx, iTtlDatesInRange As Integer
    Dim sVehicles, sStations, sMail As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim sOutput As String
    Dim DelinRst As ADODB.Recordset
    Dim sGenDate As String
    Dim sGenTime As String
    Dim iPrevStnCode As Integer
    Dim iPrevVehCode As Integer
    
    Dim dFWeek As Date
    Dim lSDate, lEDate As Long
    Dim sCurDate As String
    Dim slExportName As String
    Dim ilRptDest As Integer
    Dim slRptName As String         'report name for Overdue CP
    Dim iGotAnyData As Integer
    Dim lLoop As Long
    Dim sOffAir As String
    Dim sDropDate As String
    Dim sAttEndDate As String
    Dim blOneFoundToOutput As Boolean
    Dim slMonday As String
    Dim slSunday As String
    Dim llCpttStartDate As Long
    Dim slCpttStartDate As String
    Dim ilDoe As Integer
    Dim llRecords As Long
    llRecords = 0
    On Error GoTo ErrHand
        imExporting = True
        Frame2.Enabled = False
       lacResult.Caption = "Reading records.."
       lacResult.Visible = True
       lacResult.Refresh
       cmdDone.Caption = "Cancel"
       cmdExport.Enabled = False
        sCurDate = "1/1/01"
        Screen.MousePointer = vbHourglass
    
        'Retrieve information from the list box
        If CalFromDate.Text = "" Then
            sStartDate = "1/1/1970"
        Else
            sStartDate = CalFromDate.Text
             'end date will always default to sunday
            Do While Weekday(sStartDate) <> vbMonday
                sStartDate = DateAdd("d", -1, sStartDate)
            Loop
        End If
        
        sEndDate = CalToDate.Text
        'end date will always default to sunday
        Do While Weekday(sEndDate) <> vbSunday
            sEndDate = DateAdd("d", 1, sEndDate)
        Loop

        sStartDate = gAdjYear(Format$(DateValue(sStartDate), "m/d/yyyy"))
        sEndDate = Format$(DateValue(gAdjYear(sEndDate)), "m/d/yyyy")
        
        lSDate = DateValue(gAdjYear(sStartDate))
        lEDate = DateValue(gAdjYear(sEndDate))
        
        bmFirstTime = True      'initalize for header information only once
        
        '7-8-09 take the entire date range so that report can get the total # of discrepant affidvits.
        sDateRange = "(cpttStartDate >= '" + Format$(sStartDate, sgSQLDateForm) & "')"
        sDateRange = sDateRange & " And (cpttStartDate <= '" + Format$(sEndDate, sgSQLDateForm) & "')"
        sDateRange = sDateRange & " And (cpttStartDate >= attOnAir)"
        sDateRange = sDateRange & " And (cpttStartDate <= attOffAir)"
        sDateRange = sDateRange & " And (cpttStartDate <= attDropDate)"
        sVehicles = ""
        sStations = ""
        ' Selecting by vehicles
        If chkListBox.Value = 0 Then    '= 0 Then                        'User did NOT select all vehicles
            For i = 0 To lbcVehAff.ListCount - 1 Step 1
                If lbcVehAff.Selected(i) Then
                    If Len(sVehicles) = 0 Then
                        sVehicles = "(cpttVefCode = " & lbcVehAff.ItemData(i) & ")"
                    Else
                        sVehicles = sVehicles & " OR (cpttVefCode = " & lbcVehAff.ItemData(i) & ")"
                    End If
                End If
            Next i
        End If
        
        ReDim tmOverdue_Affs(0 To 0) As OVERDUE_AFFS
        ReDim tmUpToDate_Affs(0 To 0) As OVERDUE_AFFS
        iRet = OpenExportFile(sStartDate, sEndDate)
        If iRet <> 0 Then
            MsgBox "Cannot create filename " + smExportName
        End If
        ' Select by Vehicle all unposted/partially posted affidavits between requested date span
        'build array of unique vehicle and station overdue affidavit counts
        'TTP 10138 - add DMA Name, DMA Rank
        SQLQuery = "SELECT cpttshfcode, cpttvefcode, cpttatfcode, cpttstartdate, cpttstatus, cpttpostingstatus, attonair, attoffair, attdropdate, attcode, attserviceagreement, "
        SQLQuery = SQLQuery + " vefcode, vefsort, vefname, shttcallletters, shttcode, shttPermStationID, shttMarket, shttRank "
        SQLQuery = SQLQuery + " FROM VEF_Vehicles, cptt, shtt, att "
        SQLQuery = SQLQuery + " WHERE (vefCode = cpttVefCode"
        SQLQuery = SQLQuery + " AND shttCode = cpttShfCode"
        SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
        
        'get partial and unposted
        SQLQuery = SQLQuery + " AND cpttStatus = 0 and attServiceAgreement <> 'Y' "
        If sVehicles <> "" Then
            SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
        End If
        SQLQuery = SQLQuery + " AND (" & sDateRange & ")"
        SQLQuery = SQLQuery + ")" + " ORDER BY  vefName, shttCallLetters, cpttStartDate"

        iGotAnyData = False     'assume no data found yet until something written to prepass file
        Set DelinRst = gSQLSelectCall(SQLQuery)
        If Not DelinRst.EOF Then
            sGenDate = Format$(gNow(), "m/d/yyyy")
            sGenTime = Format$(gNow(), sgShowTimeWSecForm)
            sStartDate = Format(sStartDate, "m/d/yyyy")
            While Not DelinRst.EOF
                slCpttStartDate = Format(DelinRst!CpttStartDate, "m/dd/yyyy")
                If ((DateValue(gAdjYear(slCpttStartDate)) >= lSDate) And (DateValue(gAdjYear(slCpttStartDate)) <= lEDate)) Then
                    'find the matching vehicle and station to keep count of overdue affidavits
                    blOneFoundToOutput = False
'If DelinRst!cpttatfCode = 3801 Then
'lLoop = lLoop
'End If
                     For lLoop = LBound(tmOverdue_Affs) To UBound(tmOverdue_Affs) - 1
                        If DelinRst!cpttshfcode = tmOverdue_Affs(lLoop).iShttCode And DelinRst!cpttvefcode = tmOverdue_Affs(lLoop).iVefCode Then
                            tmOverdue_Affs(lLoop).lCount = tmOverdue_Affs(lLoop).lCount + 1
'If tmOverdue_Affs(lLoop).lCount >= 30 Then
'lLoop = lLoop
'End If
                            blOneFoundToOutput = True
                            Exit For
                        End If
                    Next lLoop
                     
                    If Not blOneFoundToOutput Then
                        ReDim Preserve tmOverdue_Affs(0 To UBound(tmOverdue_Affs) + 1) As OVERDUE_AFFS
                        tmOverdue_Affs(lLoop).iVefCode = DelinRst!cpttvefcode
                        tmOverdue_Affs(lLoop).sVefName = DelinRst!vefName
                        tmOverdue_Affs(lLoop).iShttCode = DelinRst!cpttshfcode
                        tmOverdue_Affs(lLoop).sDMAName = DelinRst!shttMarket
                        tmOverdue_Affs(lLoop).sDMARank = DelinRst!shttRank
                        tmOverdue_Affs(lLoop).sCallLetters = DelinRst!shttCallLetters
                        tmOverdue_Affs(lLoop).lStationID = DelinRst!shttPermStationID
                        tmOverdue_Affs(lLoop).lCount = 1
                    End If
                End If
                If imTerminate = True Then GoTo ErrHand
                
                llRecords = llRecords + 1
                ilDoe = ilDoe + 1
                If ilDoe >= 100 Then
                    lacResult.Caption = "Processing " & llRecords & " records.."
                    lacResult.Refresh
                    DoEvents
                    ilDoe = 0
                End If
                
                DelinRst.MoveNext
            Wend
         
        End If
'        If Not iGotAnyData Then             'if not set, nothing was found to output
'            'gMsgBox "No Data Exists for Requested Period"
'            'put message in file
'        End If
        DelinRst.Close
    
        iPrevStnCode = 0
        iPrevVehCode = 0
        
        lacResult.Caption = "Reading records.."
        lacResult.Refresh

        'Select by Vehicle all posted affidavits between requested date span
        'Need to create entry in export file of those unique vehicle/stations that are uptodate in posting, and have no outstanding affidavits .
        'Create with count of 0
        'TTP 10138 - add DMA Name, DMA Rank
        SQLQuery = "SELECT cpttshfcode, cpttvefcode, cpttatfcode, cpttstartdate, cpttstatus, cpttpostingstatus, attonair, attoffair, attdropdate, attcode, attserviceagreement, "
        SQLQuery = SQLQuery + " vefcode, vefsort, vefname, shttcallletters, shttcode, shttPermStationID, shttMarket, shttRank "
        SQLQuery = SQLQuery + " FROM VEF_Vehicles, cptt, shtt, att "
        SQLQuery = SQLQuery + " WHERE (vefCode = cpttVefCode"
        SQLQuery = SQLQuery + " AND shttCode = cpttShfCode"
        SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"

        'get completely posted only
        SQLQuery = SQLQuery + " AND cpttPostingStatus = 2 and attServiceAgreement <> 'Y' "
        If sVehicles <> "" Then
            SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
        End If
        SQLQuery = SQLQuery + " AND (" & sDateRange & ")"

        SQLQuery = SQLQuery + ")" + " ORDER BY  vefName, shttCallLetters, cpttStartDate"
        Set DelinRst = gSQLSelectCall(SQLQuery)
        
        lacResult.Caption = "Processing records.."
        lacResult.Refresh
        llRecords = 0
        ilDoe = 0
        If Not DelinRst.EOF Then
            sStartDate = Format(sStartDate, "m/d/yyyy")
            While Not DelinRst.EOF
                slCpttStartDate = Format(DelinRst!CpttStartDate, "m/dd/yyyy")
                If ((DateValue(gAdjYear(slCpttStartDate)) >= lSDate) And (DateValue(gAdjYear(slCpttStartDate)) <= lEDate)) Then
    
                    If (iPrevStnCode <> DelinRst!cpttshfcode) Or (iPrevVehCode <> DelinRst!cpttvefcode) Then
                        'find the matching vehicle and station to keep count of overdue affidavits
                        blOneFoundToOutput = False
                         For lLoop = LBound(tmOverdue_Affs) To UBound(tmOverdue_Affs) - 1
                            If DelinRst!cpttshfcode = tmOverdue_Affs(lLoop).iShttCode And DelinRst!cpttvefcode = tmOverdue_Affs(lLoop).iVefCode Then
                                blOneFoundToOutput = True
                                iPrevStnCode = DelinRst!cpttshfcode
                                iPrevVehCode = DelinRst!cpttvefcode
                                Exit For
                            End If
                        Next lLoop
                         
                        If Not blOneFoundToOutput Then
                            tmUpToDate_Affs(UBound(tmUpToDate_Affs)).iVefCode = DelinRst!cpttvefcode
                            tmUpToDate_Affs(UBound(tmUpToDate_Affs)).sVefName = DelinRst!vefName
                            tmUpToDate_Affs(UBound(tmUpToDate_Affs)).iShttCode = DelinRst!cpttshfcode
                            tmUpToDate_Affs(UBound(tmUpToDate_Affs)).sCallLetters = DelinRst!shttCallLetters
                            tmUpToDate_Affs(UBound(tmUpToDate_Affs)).sDMAName = DelinRst!shttMarket
                            tmUpToDate_Affs(UBound(tmUpToDate_Affs)).sDMARank = DelinRst!shttRank
                            tmUpToDate_Affs(UBound(tmUpToDate_Affs)).lStationID = DelinRst!shttPermStationID
                            tmUpToDate_Affs(UBound(tmUpToDate_Affs)).lCount = 0
                            iPrevStnCode = DelinRst!cpttshfcode
                            iPrevVehCode = DelinRst!cpttvefcode
                            ReDim Preserve tmUpToDate_Affs(0 To UBound(tmUpToDate_Affs) + 1) As OVERDUE_AFFS
                        End If
                    End If
                End If
                If imTerminate = True Then GoTo ErrHand
                llRecords = llRecords + 1
                ilDoe = ilDoe + 1
                If ilDoe >= 1000 Then
                    lacResult.Caption = "Processing " & llRecords & " records.."
                    DoEvents
                    ilDoe = 0
                End If
                DelinRst.MoveNext
            Wend
            DelinRst.Close
        End If
        iRet = mWriteExportRec(tmOverdue_Affs())
        If iRet <> 0 Then   'error
            lacResult.Caption = "Error writing to filename : " & smExportName
            lacResult.Visible = True
            cmdExport.Enabled = True
            Exit Sub
        End If
        lacResult.Caption = "Writing Records... "
        iRet = mWriteExportRec(tmUpToDate_Affs())
        Close #hmExport
        lacResult.Visible = True
        cmdDone.Caption = "Done"
        cmdExport.Enabled = True
        Erase tmOverdue_Affs
        Erase tmUpToDate_Affs
        lacResult.Caption = "File Saved: " & Trim$(smExportName)
        imExporting = False
        Frame2.Enabled = True
        Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmExportOverdue-cmdExport"
    
    On Error Resume Next
    If hmExport <> -1 Then Close #hmExport
    Kill smExportName
    lacResult.Visible = True
    cmdExport.Enabled = True
    Frame2.Enabled = True
    Erase tmOverdue_Affs
    Erase tmUpToDate_Affs
    imExporting = False
    If imTerminate = True Then Unload Me
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
    If ilLoaded = False Then
        Frame2.Visible = True
        lacResult.Visible = True
        cmdDone.Visible = True
        cmdExport.Visible = True
        ilLoaded = True
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.2
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmExportOverdue
    gCenterForm frmExportOverdue
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim slSunday As String
    Dim rst_Temp As ADODB.Recordset
    
    imTerminate = False
    imExporting = False
    ilLoaded = False
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    

'    SQLQuery = "SELECT * From Site Where siteCode = 1"
'    Set rst = gSQLSelectCall(SQLQuery)
'    If Not rst.EOF Then
'        imNoDaysDelq = rst!siteOMNoWeeks         'determine from date by # weeks considered overdue
'    End If
'

    SQLQuery = "SELECT spfgclient, spfMnfClientAbbr, mnfName"
    SQLQuery = SQLQuery & " FROM SPF_Site_Options, MNF_Multi_Names"
    SQLQuery = SQLQuery & " WHERE spfCode = 1"
    SQLQuery = SQLQuery & " AND spfMnfClientAbbr = mnfCode"
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    'FTP site for audio to show on the log screen
    smClientName = ""
    If Not rst_Temp.EOF Then
        If rst_Temp!spfMnfClientAbbr > 0 Then
            smClientName = Trim$(rst_Temp!mnfName)
        Else
            smClientName = rst_Temp!spfgClient
        End If
    End If
    frmExportOverdue.Caption = "Export Overdue Affidavits - " & Trim(rst_Temp!spfgClient)
    rst_Temp.Close
    slSunday = Format$(gNow(), sgShowDateForm)
    Do While Weekday(slSunday) <> vbSunday
        slSunday = DateAdd("d", -1, slSunday)
    Loop

    CalToDate.Text = slSunday
    CalFromDate.Text = ""
    imChkListBoxIgnore = False
    
    CalFromDate.SetEnabled (True)
    CalToDate.SetEnabled (True)
    labFrom.Enabled = True

    'load vehicles
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    
    chkListBox.Caption = "All Vehicles"
    chkListBox.Value = 0    'False
    lbcVehAff.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmOverdue_Affs
    Erase tmUpToDate_Affs
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmExportOverdue = Nothing
End Sub


Private Sub lbcVehAff_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = 1 Then
        imChkListBoxIgnore = True
        chkListBox.Value = 0    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If
    mSetGenerate
End Sub

Private Sub CalFromDate_Change()
Dim ilLen As Integer
Dim slDate As String
Dim llFromDate As Long
Dim llToDate As Long

    If lmDefaultStartDate > 0 Then
        ilLen = Len(CalFromDate.Text)
        If ilLen >= 3 Then              'date entered may be x/x (no year)
            slDate = CalFromDate.Text          'retrieve jan thru dec year
            llFromDate = gDateValue(slDate)
            slDate = CalToDate.Text
            llToDate = gDateValue(slDate)
        End If
    End If
    mSetGenerate
End Sub

Private Sub CalFromDate_GotFocus()
    gCtrlGotFocus CalFromDate
    CalFromDate.ZOrder (vbBringToFront)
End Sub

Private Sub CalToDate_Change()
Dim ilLen As Integer
Dim slDate As String
Dim llFromDate As Long
Dim llToDate As Long

    If lmDefaultEndDate > 0 Then
        ilLen = Len(CalToDate.Text)
        If ilLen >= 3 Then                    'date entered may be x/x (no year)
            slDate = CalToDate.Text          'retrieve jan thru dec year
            llToDate = gDateValue(slDate)
            slDate = CalFromDate.Text
            llFromDate = gDateValue(slDate)
            
        End If
    End If
    mSetGenerate
End Sub

Private Sub CalToDate_GotFocus()
    gCtrlGotFocus CalToDate
    CalToDate.ZOrder (vbBringToFront)
End Sub
'
'   Turn on Generate button only if the end date is set and at least one vehicle selected
'
Public Sub mSetGenerate()
    
    cmdExport.Enabled = False
    If chkListBox.Value = vbChecked Or lbcVehAff.SelCount > 0 Then
        If CalToDate.Text <> "" Then
            cmdExport.Enabled = True
        End If
    End If
End Sub
Private Function mWriteExportRec(tlAffs() As OVERDUE_AFFS) As Integer
Dim llLoopOnAffs As Long
Dim slStr As String
Dim ilError As Integer
Dim llRecords As Long
Dim ilDoe As Integer

    ilError = False
    If bmFirstTime Then         'create the header record
        slStr = "As of " & Format$(gNow(), "mm/dd/yy") & " "
        slStr = slStr & Format$(gNow(), "h:mm:ssAM/PM")
        If chkListBox.Value = vbChecked Then                'all vehicles selected
            slStr = slStr & " All Vehicles"
        Else
            slStr = slStr & " Some Vehicles"
        End If
        On Error GoTo mWriteExportRecErr
        Print #hmExport, slStr        'write header description
        On Error GoTo 0

        'slStr = """" & "Station ID" & """" & "," & """" & "Call Letters-Band" & """" & "," & """" & "Weeks" & """" & "," & """" & "Vehicle" & """" & ","
        'TTP 10138
        slStr = """Station ID"",""Call Letters-Band"",""DMA Name"",""DMA Rank"",""Weeks"",""Vehicle"""
        
        On Error GoTo mWriteExportRecErr
        Print #hmExport, slStr     'write header description
        On Error GoTo 0

        bmFirstTime = False         'do the heading and time stamp only once
    End If
    
    For llLoopOnAffs = LBound(tlAffs) To UBound(tlAffs) - 1
'            slStr = """" & Trim$(Str$(tlAffs(llLoopOnAffs).lStationID)) & """" & ","
'            slStr = slStr & """" & Trim$(tlAffs(llLoopOnAffs).sCallLetters) & """" & ","
'            slStr = slStr & """" & Trim$(Str$(tlAffs(llLoopOnAffs).lCount)) & """" & ","
'            slStr = slStr & """" & Trim$(tlAffs(llLoopOnAffs).sVefName) & """"

            slStr = Trim$(Str$(tlAffs(llLoopOnAffs).lStationID)) & ","
            slStr = slStr & """" & Trim$(tlAffs(llLoopOnAffs).sCallLetters) & """" & ","
            'TTP 10138
            slStr = slStr & """" & Trim$(tlAffs(llLoopOnAffs).sDMAName) & """" & ","
            slStr = slStr & """" & Trim$(tlAffs(llLoopOnAffs).sDMARank) & """" & ","
            slStr = slStr & Trim$(Str$(tlAffs(llLoopOnAffs).lCount)) & ","
            slStr = slStr & """" & Trim$(tlAffs(llLoopOnAffs).sVefName) & """"
        
            llRecords = llRecords + 1
            ilDoe = ilDoe + 1
            If ilDoe >= 100 Then
                lacResult.Caption = "Writing " & llRecords & " records.."
                DoEvents
                ilDoe = 0
            End If
            
            On Error GoTo mWriteExportRecErr
            Print #hmExport, slStr
            On Error GoTo 0
    Next llLoopOnAffs

   
    mWriteExportRec = ilError
    Exit Function

mWriteExportRecErr:
    ilError = True
    Resume Next

    
End Function
