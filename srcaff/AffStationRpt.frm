VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStationRpt 
   Caption         =   "Station Information"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "AffStationRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frcSelection 
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
      Top             =   1860
      Width           =   8775
      Begin VB.CommandButton cmdStationListFile 
         Height          =   360
         Left            =   8040
         Picture         =   "AffStationRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Select Stations from File.."
         Top             =   210
         Width           =   360
      End
      Begin VB.CheckBox ckcInclWebPW 
         Caption         =   "Include Web Password"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3915
         Width           =   2070
      End
      Begin VB.Frame frcSubsort 
         Caption         =   "Multicast Subsort"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   3225
         Visible         =   0   'False
         Width           =   3480
         Begin VB.OptionButton optSubSort 
            Caption         =   "DMA Market"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   29
            Top             =   270
            Width           =   1320
         End
         Begin VB.OptionButton optSubSort 
            Caption         =   "Owner"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   270
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox ckcAll 
         Caption         =   "All "
         Height          =   255
         Left            =   4800
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.ListBox lbcTitles 
         Height          =   1620
         ItemData        =   "AffStationRpt.frx":0E34
         Left            =   4800
         List            =   "AffStationRpt.frx":0E3B
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   2640
         Visible         =   0   'False
         Width           =   3660
      End
      Begin VB.ListBox lbcSortList 
         Height          =   3570
         ItemData        =   "AffStationRpt.frx":0E42
         Left            =   4800
         List            =   "AffStationRpt.frx":0E49
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   3660
      End
      Begin VB.Frame frcContact 
         Caption         =   "Contact"
         Height          =   1095
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   3480
         Begin VB.OptionButton optContact 
            Caption         =   "Aff-Email"
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   33
            Top             =   510
            Width           =   1140
         End
         Begin VB.OptionButton optContact 
            Caption         =   "ISCI Export"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   510
            Width           =   1260
         End
         Begin VB.OptionButton optContact 
            Caption         =   "Affidavit Contact (Label)"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   270
            Value           =   -1  'True
            Width           =   2010
         End
         Begin VB.OptionButton optContact 
            Caption         =   "Specific Title"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   765
            Width           =   1275
         End
         Begin VB.OptionButton optContact 
            Caption         =   "All Contacts"
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   20
            Top             =   765
            Width           =   1260
         End
      End
      Begin VB.Frame frcSortBy 
         Caption         =   "Sort by"
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3480
         Begin VB.OptionButton optSortBy 
            Caption         =   "MSA Market"
            Height          =   255
            Index           =   4
            Left            =   1710
            TabIndex        =   31
            Top             =   510
            Width           =   1245
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "Multicast"
            Height          =   255
            Index           =   3
            Left            =   1710
            TabIndex        =   18
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "Owner"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   510
            Width           =   1140
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "DMA Market"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   765
            Width           =   1275
         End
         Begin VB.OptionButton optSortBy 
            Caption         =   "Call  Letters"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select"
         Height          =   570
         Left            =   120
         TabIndex        =   7
         Top             =   225
         Width           =   3480
         Begin VB.OptionButton optSP 
            Caption         =   "Both"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   10
            Top             =   270
            Width           =   735
         End
         Begin VB.OptionButton optSP 
            Caption         =   "Stations"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   270
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optSP 
            Caption         =   "People"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   9
            Top             =   270
            Width           =   975
         End
      End
      Begin VB.Label lacTitles 
         Caption         =   "Titles"
         Height          =   255
         Left            =   4920
         TabIndex        =   25
         Top             =   2335
         Visible         =   0   'False
         Width           =   975
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
      FormDesignHeight=   6480
      FormDesignWidth =   9360
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4455
      TabIndex        =   13
      Top             =   1200
      Width           =   1920
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4275
      TabIndex        =   12
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4050
      TabIndex        =   11
      Top             =   240
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
      Height          =   1485
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffStationRpt.frx":0E50
         Left            =   1065
         List            =   "AffStationRpt.frx":0E52
         TabIndex        =   4
         Top             =   690
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   480
         Width           =   2190
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2310
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Export"
         Height          =   255
         Index           =   4
         Left            =   135
         TabIndex        =   34
         Top             =   960
         Width           =   1200
      End
   End
   Begin VB.Label lblExportDestination 
      Caption         =   "Export Stored in-"
      Height          =   255
      Left            =   3360
      TabIndex        =   35
      Top             =   1600
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "frmStationRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim imckcAllIgnore As Integer
Dim tmStatInfoRst As ADODB.Recordset       'station records
Dim tmPersonInfoRst As ADODB.Recordset     'contact rcords
Dim tmEmailInfoRst As ADODB.Recordset
Dim tmStationArray() As STATIONARRAY    'Info formatted from station & contact records
'   sample array (entries 0 -3) :
'       1234 Any Street          John Doe, Affiliate Contact,P:800 111-1111,F: 800 222-2222 E: johndoe@en.com
'       Any Suite#               John Smith, Program Director, P:800 111-1112,SF: 800 222-2223 E: johnsmith@en.com
'       Any City, Any State, Zip (blank)
'       Country USA              (blank)
'
'******************************************************
'*  frmStationRpt - a report of station contact information.
'*
'*  Created January,1998 by Wade Bjerke
'
'       3-15-06 Modified D hosaka:  rewrite with prepass file
'           to work with all new file structures for personnel
'           and multicast.  Add more selectivity and sorts.
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************


Option Explicit

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
    If lbcSortList.ListCount > 0 Then
        imckcAllIgnore = True
        lRg = CLng(lbcSortList.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcSortList.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload frmStationRpt
End Sub


Private Sub cmdReport_Click()
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim lLoop As Long
    Dim slSortQuery As String       'SQL call for the selection items from list box (mkt, call ltrs, owner)
    Dim slStationTypeQuery As String    'sql call for type of station records to retrieve (station, people)
    Dim slContactQuery As String        'part of the sql call to get the personnel info from artt
    Dim slFullContactQuery As String    'complete call to get the station contacts
    Dim ilMarket As Integer     'true if selecting markets
    Dim ilOwner As Integer      'true if selecting owners
    Dim ilCallLetters As Integer    'true if selecting call letters
    Dim ilMSAMarket As Integer
    Dim sGenDate As String      'gen date for prepass filter
    Dim sGenTime As String      'gen time for prepass filter
    Dim llIndex As Long
    Dim ilFoundOne As Integer
    Dim slPersonInfo As String
    Dim llCount As Long         'debugging purposes only
    Dim slEmailQuery As String
    Dim slExportFilename As String
    Dim blNeedAnd As Boolean
    Dim ilDoe As Integer
    'Dim NewForm As New frmViewReport
                
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass

    If optRptDest(0).Value = True Then
       ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilRptDest = 2
        ilExportType = cboFileType.ListIndex        '3-12-04
    ElseIf optRptDest(4).Value = True Then
        ilRptDest = 4                               '10-08-20 - TTP 9985 - Export Station Information in CSV format
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If lbcSortList.SelCount <= 0 And CkcAll.Value = vbUnchecked Then
        gMsgBox "Select at least 1 station"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If optContact(1).Value = True Then 'specific title
        If lbcTitles.SelCount <= 0 Then
            gMsgBox "Select at least 1 contact title"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    If optRptDest(4).Value = True Then
        'make export filename = ExportPath + '\' + StationInfo mmddyy-hhmmssa/p.csv
        slExportFilename = IIf(right(sgExportDirectory, 1) <> "\", sgExportDirectory & "\", sgExportDirectory) & "StationInfo "
        slExportFilename = slExportFilename & IIf(Len(Month(Date)) = 2, Month(Date), "0" & Month(Date))
        slExportFilename = slExportFilename & IIf(Len(Day(Date)) = 2, Day(Date), "0" & Day(Date))
        slExportFilename = slExportFilename & right(Year(Date), 2)
        slExportFilename = slExportFilename & "-"
        slExportFilename = slExportFilename & IIf(Hour(TIME()) = 0, "12", IIf(Hour(TIME()) > 12, IIf(Hour(TIME()) - 12 < 10, "0" & Hour(TIME()) - 12, Hour(TIME()) - 12), IIf(Hour(TIME()) < 10, "0" & Hour(TIME()), Hour(TIME()))))
        slExportFilename = slExportFilename & IIf(Minute(TIME()) < 10, "0" & Minute(TIME()), Minute(TIME()))
        slExportFilename = slExportFilename & IIf(Second(TIME()) < 10, "0" & Second(TIME()), Second(TIME()))
        slExportFilename = slExportFilename & IIf(Hour(TIME()) >= 12, "p", "a")
        slExportFilename = slExportFilename & ".csv"
        
        'Export data
        mExportStationInfoCSV (slExportFilename)
        
        'Done with Export
        lblExportDestination.Caption = "Export Stored in- " & slExportFilename
        Screen.MousePointer = vbDefault
        lblExportDestination.Visible = True
        cmdReport.Enabled = True
        cmdReturn.Enabled = True
        cmdDone.Enabled = True
        Exit Sub
    End If
    
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    'set the appropriate sort type to determine SQL query
    ilMarket = False
    ilCallLetters = False
    ilOwner = False
    sgCrystlFormula2 = ""
    If optSortby(0).Value = True Then       'call letters sort
        ilCallLetters = True
        sgCrystlFormula1 = "'C'"
    ElseIf (optSortby(1).Value = True) Then        'market   multicast with market subsort
        ilMarket = True
        sgCrystlFormula1 = "'M'"
    ElseIf (optSortby(2).Value = True) Then        'owner
        ilOwner = True
        sgCrystlFormula1 = "'O'"
    ElseIf (optSortby(3).Value = True And optSubSort(1).Value = True) Then      'multicast with market subsort
        ilMarket = True
        sgCrystlFormula1 = "'U'"
        sgCrystlFormula2 = "'M'"                          'market subsort
    ElseIf (optSortby(3).Value = True And optSubSort(0).Value = True) Then     ' multicast with owner subsort
        ilOwner = True
        sgCrystlFormula1 = "'U'"
        sgCrystlFormula2 = "'O'"                          'owner subsort
    ElseIf optSortby(4).Value = True Then          'MSA market
        ilMSAMarket = True
        sgCrystlFormula1 = "'S'"                    'metro survey
    End If
    
    SQLQuery = "SELECT *  From shtt left outer JOIN mkt ON shttMktCode = mktCode "
    SQLQuery = SQLQuery + " left outer JOIN  artt ON shttOwnerArttCode = arttCode "
    'SQLQuery = SQLQuery + " left outer JOIN  mgt ON shttCode = mgtshfCode"
    SQLQuery = SQLQuery + " left outer JOIN  met ON shttmetCode = metCode"

    
    slSortQuery = ""
    slStationTypeQuery = ""
    slContactQuery = ""
    llCount = 0                 'count of stations retreived from result set
    
    'select station, people or both
    If optSP(0).Value Then
        slStationTypeQuery = " WHERE (shttType = 0)" ' ORDER BY shttState, shttCallLetters"
    ElseIf optSP(1).Value Then
        slStationTypeQuery = " WHERE (shttType = 1)"
    End If
    
    'sort by call letters (0), DMA market (1), owner (2), multicast(3), MSA Market (4)
      If CkcAll.Value = vbUnchecked Then          'all is not checked, format the codes of selected items
        
        'TTP 10638 - Station Information report: "maximum of 3000 predicates was exceeded" error
'        If slStationTypeQuery = "" Then
'            slStationTypeQuery = " WHERE ("
'        Else
'            slStationTypeQuery = slStationTypeQuery + " and ("
'        End If
        If slStationTypeQuery = "" Then
            slStationTypeQuery = " WHERE "
        Else
            slStationTypeQuery = slStationTypeQuery + " AND "
        End If
        If ilCallLetters Then
            slSortQuery = "shttCode IN ("
        ElseIf ilMarket Then            'filter out matching market
            slSortQuery = "mktCode IN ("
        ElseIf ilMSAMarket Then     'filter matching MSA markets
            slSortQuery = "metCode IN ("
        Else             'owners
            slSortQuery = "arttCode IN ("
        End If
'        For lLoop = 0 To lbcSortList.ListCount - 1 Step 1
'            If lbcSortList.Selected(lLoop) Then
'                If Len(slSortQuery) = 0 Then
'                    If ilCallLetters Then       'call ltrs
'                        slSortQuery = " (shttCode = " & lbcSortList.ItemData(lLoop) & ")"
'                    ElseIf ilMarket Then        'filter matching markets
'                        slSortQuery = " (mktCode = " & lbcSortList.ItemData(lLoop) & ")"
'                    ElseIf ilMSAMarket Then     'filter matching MSA markets
'                        slSortQuery = " (metCode = " & lbcSortList.ItemData(lLoop) & ")"
'                    Else                    'owner
'                        slSortQuery = " (arttCode = " & lbcSortList.ItemData(lLoop) & ")"
'                    End If
'                Else
'                    If ilCallLetters Then
'                        slSortQuery = slSortQuery & " OR (shttCode = " & lbcSortList.ItemData(lLoop) & ")"
'                    ElseIf ilMarket Then            'filter out matching market
'                        slSortQuery = slSortQuery & " OR (mktCode = " & lbcSortList.ItemData(lLoop) & ")"
'                    ElseIf ilMSAMarket Then     'filter matching MSA markets
'                        slSortQuery = " (metCode = " & lbcSortList.ItemData(lLoop) & ")"
'                   Else             'owners
'                        slSortQuery = slSortQuery & " OR (arttCode = " & lbcSortList.ItemData(lLoop) & ")"
'                   End If
'                End If
'            End If
'        Next lLoop
        For lLoop = 0 To lbcSortList.ListCount - 1 Step 1
            If lbcSortList.Selected(lLoop) Then
                If blNeedAnd Then
                    slSortQuery = slSortQuery & ","
                    blNeedAnd = False
                End If
                slSortQuery = slSortQuery & lbcSortList.ItemData(lLoop)
                blNeedAnd = True
            End If
        Next lLoop
        slSortQuery = slSortQuery + ")"
    End If
    
    slContactQuery = "Select * from ARTT, TNT where "
    slContactQuery = "Select * from ARTT left outer JOIN tnt ON artttntCode = tntCode where"
    'determine which contacts to show, Affiliate contact (0), specific contact(1), all contacts(2)
    If optContact(0).Value Then         'affiliate contact
        'slContactQuery = slContactQuery + " (arttAffContact = '1')"
        '3-8-11 change reference to get the affidavit contact
        'slContactQuery = slContactQuery + " (arttWebEMail = 'Y') and artttntCode = tntCode "
        
        '9-19-11 change back to using affContact
        '11-8-11 fix sql call, remove extra " and artttntCode = tntCode"
        slContactQuery = slContactQuery + " (arttaffContact = '1')  "
    ElseIf optContact(1).Value = True Then 'specific title
        For lLoop = 0 To lbcTitles.ListCount - 1 Step 1
            If lbcTitles.Selected(lLoop) Then
                slContactQuery = slContactQuery + " (artttntCode = " & lbcTitles.ItemData(lLoop) & ") and artttntcode = tntcode "
                Exit For        'only 1 selection allowed
            End If
        Next lLoop
    ElseIf optContact(2).Value = True Then            'get whatever titles are defined (all contacts)
        slContactQuery = "Select * from ARTT left outer JOIN tnt ON artttntCode = tntCode where"
        slContactQuery = slContactQuery + " (arttType <> 'A' and arttType <> 'O') "
    '11-6-13 implement selectivity for contacts
    ElseIf optContact(3).Value = True Then            'ISCI export contacts
        slContactQuery = slContactQuery + " (arttISCI2Contact = '1')  "
    ElseIf optContact(4).Value = True Then              'aff-email contacts
        slContactQuery = slContactQuery + " (arttWebEmail = 'Y')  "
    End If
    'slContactQuery = slContactQuery + " and artttntCode = tntcode "

    'obtain all stations with their associated titles, and markets, owners & groups (where applicable)
    SQLQuery = SQLQuery + slStationTypeQuery + slSortQuery
    
    Set tmStatInfoRst = gSQLSelectCall(SQLQuery)

    'get generation date and time for crystal report filter of records
    sGenDate = Format$(gNow(), "m/d/yyyy")
    sGenTime = Format$(gNow(), sgShowTimeWSecForm)

    'loop thru all the stations and build  array of one or more entries containing its address
    'Once address if built, find all the associated personnel records for the station and build
    'into same memory array (tmstationarray)
    While Not tmStatInfoRst.EOF
        llCount = llCount + 1               'debugging only
        ReDim tmStationArray(0 To 0) As STATIONARRAY
        'initalize first index of array
        tmStationArray(0).sAddress = ""
        tmStationArray(0).sPersonInfo = ""
        llIndex = 0
        'build the location of the station into max 5 lines of addresses
        mBuildLoc llIndex, Trim$(tmStatInfoRst!shttAddress1), ""
        mBuildLoc llIndex, Trim$(tmStatInfoRst!shttAddress2), ""
        If Trim$(tmStatInfoRst!shttCity) = "" Or Trim$(tmStatInfoRst!shttState) = "" Then       'no commas if not city & state
            mBuildLoc llIndex, Trim$(tmStatInfoRst!shttCity), Trim$(tmStatInfoRst!shttState)
        Else
            mBuildLoc llIndex, (Trim$(tmStatInfoRst!shttCity) & ","), Trim$(tmStatInfoRst!shttState) + " " + Trim$(tmStatInfoRst!shttZip)
        End If
        mBuildLoc llIndex, Trim$(tmStatInfoRst!shttCountry), ""
        
        '2-22-11
        'get the station main phone #, fax & email
        llIndex = 0
        slPersonInfo = "Main # "
        If Trim$(tmStatInfoRst!shttPhone) <> "" Then
            slPersonInfo = slPersonInfo + " SP:" + Trim$(tmStatInfoRst!shttPhone)     'no person phone, use station phone
        End If
        
        'get the station fax #
        If Trim$(tmStatInfoRst!shttFax) <> "" Then       'does station fax exist?
            slPersonInfo = slPersonInfo + ", SF:" + Trim$(tmStatInfoRst!shttFax)     'no person fax, use station phone
        End If
        
         'get the email
        '10-4-11 remove this item to show on the report.  No longer applies
'        If Trim$(tmStatInfoRst!shttEMail) <> "" Then        'does station email exist?
'            slPersonInfo = slPersonInfo + ", SE:" + Trim$(tmStatInfoRst!shttEMail)     'no person email, use station email
'        End If
        
        mBuildPerson llIndex, slPersonInfo

        
        slFullContactQuery = slContactQuery + " and (arttshttCode = " + Str(tmStatInfoRst!shttCode) + ") order by arttAffContact desc, tntTitle"
        
        'get all the matching station personnel information for this station
        Set tmPersonInfoRst = gSQLSelectCall(slFullContactQuery)

        'llIndex = 0
        ilFoundOne = False
        While Not tmPersonInfoRst.EOF
            ilFoundOne = True
            slPersonInfo = ""
            'Test for Affidavit contact person
             If tmPersonInfoRst!arttAffContact = "1" Then       'Aff Label is the contact
                slPersonInfo = "A:"
            End If
            
            '11-6-13 Show the types of email addresses
            If tmPersonInfoRst!arttWebEMail = "Y" Then      'Affiliate Email
                slPersonInfo = slPersonInfo + "AE:"
            End If
            If tmPersonInfoRst!arttISCI2Contact = "1" Then
                slPersonInfo = slPersonInfo + "IE:"         'ISCI Export
            End If
            
            If slPersonInfo = "" Then           'no affidavit contact, just show name & title
                If tmPersonInfoRst!arttTntCode = 0 Then
                    slPersonInfo = Trim$(tmPersonInfoRst!arttFirstName) + " " + Trim$(tmPersonInfoRst!arttLastName)
                Else
                    If IsNull(tmPersonInfoRst!tntTitle) Then             '8-11-16 some artt records pointing to non-existant records due to previous bug that was fixed
                        slPersonInfo = Trim$(tmPersonInfoRst!arttFirstName) + " " + Trim$(tmPersonInfoRst!arttLastName)
                    Else
                        If Trim$(tmPersonInfoRst!tntTitle) = "" Then
                            slPersonInfo = Trim$(tmPersonInfoRst!arttFirstName) + " " + Trim$(tmPersonInfoRst!arttLastName)
                        Else
                            slPersonInfo = Trim$(tmPersonInfoRst!arttFirstName) + " " + Trim$(tmPersonInfoRst!arttLastName) + ", " + Trim$(tmPersonInfoRst!tntTitle)
                        End If
                    End If
                End If
            Else            'this is an affidavit contact
                If tmPersonInfoRst!arttTntCode = 0 Then
                    slPersonInfo = slPersonInfo + Trim$(tmPersonInfoRst!arttFirstName) + " " + Trim$(tmPersonInfoRst!arttLastName)
                Else
                    If IsNull(tmPersonInfoRst!tntTitle) Then             '8-11-16 some artt records pointing to non-existant records due to previous bug that was fixed
                        slPersonInfo = slPersonInfo + Trim$(tmPersonInfoRst!arttFirstName) + " " + Trim$(tmPersonInfoRst!arttLastName)
                    Else
                        If Trim$(tmPersonInfoRst!tntTitle) = "" Then
                            slPersonInfo = slPersonInfo + Trim$(tmPersonInfoRst!arttFirstName) + " " + Trim$(tmPersonInfoRst!arttLastName)
                        Else
                            slPersonInfo = slPersonInfo + Trim$(tmPersonInfoRst!arttFirstName) + " " + Trim$(tmPersonInfoRst!arttLastName) + ", " + Trim$(tmPersonInfoRst!tntTitle)
                        End If
                    End If
                End If
            End If
             
            'get the personnel phone #
            If Trim$(tmPersonInfoRst!arttPhone) = "" Then         'does person phone exist?
                If Trim$(tmStatInfoRst!shttPhone) <> "" Then
                    slPersonInfo = slPersonInfo + ", SP:" + Trim$(tmStatInfoRst!shttPhone)     'no person phone, use station phone
                End If
            Else
                slPersonInfo = slPersonInfo + ", P:" + Trim$(tmPersonInfoRst!arttPhone)  'use person phone
            End If
            
            'get the personnel fax #
            If Trim$(tmPersonInfoRst!arttFax) = "" Then         'does person fax exist?
                If Trim$(tmStatInfoRst!shttFax) <> "" Then       'does station fax exist?
                    slPersonInfo = slPersonInfo + ", SF:" + Trim$(tmStatInfoRst!shttFax)     'no person fax, use station phone
                End If
            Else                                                'person fax
                slPersonInfo = slPersonInfo + ", F:" + Trim$(tmPersonInfoRst!arttFax)   'use person fax
            End If
            
             'get the email
            If Trim$(tmPersonInfoRst!arttEmail) = "" Then         'does person email exist?
            '10-4-11 remove this item to show on the report.  No longer applies
            'If the Agreement email doesnt exist, show NOne
'                If Trim$(tmStatInfoRst!shttEMail) <> "" Then        'does station email exist?
'                    slPersonInfo = slPersonInfo + ", SE:" + Trim$(tmStatInfoRst!shttEMail)     'no person email, use station email
'                End If
                ilFoundOne = ilFoundOne
            Else                                                'use person email
                slPersonInfo = slPersonInfo + ", E:" + Trim$(tmPersonInfoRst!arttEmail)   'use person email
            End If
            
            mBuildPerson llIndex, slPersonInfo
            tmPersonInfoRst.MoveNext
        Wend            'loop on personnel records
        If Not ilFoundOne Then       'didnt retrieve any personel info and they asked for affiliate contact to be shown
                                                                    'force "NONE" for the aff contact
            slPersonInfo = "A:None"
            mBuildPerson llIndex, slPersonInfo
            
           'get the web email
            slPersonInfo = ""
           '2-22-11 email is built into each personnel; plus stations email is always shown
'            'slEmailQuery = "Select * from emt where emtshttcode = " + Str(tmStatInfoRst!shttCode)
'            slEmailQuery = "Select * from artt where arttshttcode = " + Str(tmStatInfoRst!shttCode)
'
'            'get all the matching station personnel information for this station
'            Set tmEmailInfoRst = gSQLSelectCall(slEmailQuery)
'            While Not tmEmailInfoRst.EOF
'
'                'If Trim$(tmEmailInfoRst!emtEmail) <> "" Then   'does station web email exist?
'                If Trim$(tmEmailInfoRst!arttEmail) <> "" Then   'does station web email exist?
'                    If Trim$(slPersonInfo) = "" Then
'                        'slPersonInfo = Trim$(tmEmailInfoRst!emtEmail)
'                        slPersonInfo = Trim$(tmEmailInfoRst!arttEmail)
'                    Else
'                        'slPersonInfo = slPersonInfo & "," & Trim$(tmEmailInfoRst!emtEmail)   'station web email
'                        slPersonInfo = slPersonInfo & "," & Trim$(tmEmailInfoRst!arttEmail)   'station web email
'                    End If
'                End If
'                tmEmailInfoRst.MoveNext
'            Wend
'            If slPersonInfo = "" Then
'                slPersonInfo = "WE:None"
'            Else
'                slPersonInfo = "WE:" + Trim$(slPersonInfo)
'            End If
            
            If ckcInclWebPW = vbChecked Then
                If Trim$(tmStatInfoRst!shttWebPW) = "" Then   'does station web password exist?
                    slPersonInfo = slPersonInfo + "WP: None "     'no station web password"
                Else
                    slPersonInfo = slPersonInfo + "WP:" + Trim$(tmStatInfoRst!shttWebPW)   'station web password
                End If
            End If
            
            mBuildPerson llIndex, slPersonInfo
            mInsertIntoIVR sGenDate, sGenTime
        Else
        '2-22-11 email is built into each personnel; plus stations email is always shown
'           'get the web email
             slPersonInfo = ""
'
'            'slEmailQuery = "Select * from emt where emtshttcode = " + Str(tmStatInfoRst!shttCode)
'            slEmailQuery = "Select * from artt where arttshttcode = " + Str(tmStatInfoRst!shttCode)
'
'            'get all the matching station personnel information for this station
'            Set tmEmailInfoRst = gSQLSelectCall(slEmailQuery)
'            While Not tmEmailInfoRst.EOF
'
'                If Trim$(tmEmailInfoRst!arttEmail) <> "" Then   'does station web email exist?
'                    If Trim$(slPersonInfo) = "" Then
'                        slPersonInfo = Trim$(tmEmailInfoRst!arttEmail)
'                    Else
'                        slPersonInfo = slPersonInfo & "," & Trim$(tmEmailInfoRst!arttEmail)   'station web email
'                    End If
'                End If
'                tmEmailInfoRst.MoveNext
'            Wend
'            If slPersonInfo = "" Then
'                slPersonInfo = "WE:None"
'            Else
'                slPersonInfo = "WE:" + Trim$(slPersonInfo)
'            End If
            
            If ckcInclWebPW = vbChecked Then
                If Trim$(tmStatInfoRst!shttWebPW) = "" Then   'does station web password exist?
                    slPersonInfo = slPersonInfo + "WP: None "     'no station web password"
                Else
                    slPersonInfo = slPersonInfo + "WP:" + Trim$(tmStatInfoRst!shttWebPW)   'station web password
                End If
            End If
            mBuildPerson llIndex, slPersonInfo
            mInsertIntoIVR sGenDate, sGenTime
        End If
        
        tmStatInfoRst.MoveNext
        ilDoe = ilDoe + 1
        If ilDoe > 200 Then
            DoEvents
            ilDoe = 0
        End If
        
    Wend            'loop on station records
    
    If optSortby(3).Value = True Then           'if multicast option,  only print those that have multicast flag
        SQLQuery = "SELECT * FROM IVR_Invoice_Rpt, shtt   "
        SQLQuery = SQLQuery + " WHERE ivrchfcode = shttcode   and ivrSpotType = '*' and ivrGenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND ivrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "' Order by ivrPayAddr1, ivrPayAddr2, ivrPayAddr3, ivrpayAddr4, ivrSpotKeyNo"
    Else
        SQLQuery = "SELECT * FROM IVR_Invoice_Rpt, shtt   "
        SQLQuery = SQLQuery + " WHERE ivrchfcode = shttcode    and ivrGenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND ivrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "' Order by ivrPayAddr1, ivrPayAddr2, ivrPayAddr3, ivrpayAddr4, ivrSpotKeyNo"
    End If
    
    slRptName = "afStatin.rpt"
    slExportName = "StatRpt"
    gUserActivityLog "E", sgReportListName & ": Prepass"
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName
    
    
    
    ' Delete the info we stored in the IVR table
    SQLQuery = "DELETE FROM IVR_INVOICE_RPT"
    SQLQuery = SQLQuery & " WHERE (ivrGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and ivrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "StationRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    cmdReport.Enabled = True            'give user back control to gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
    
    Screen.MousePointer = vbDefault
    Exit Sub
    

    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "StationRpt-cmdReport_Click"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmStationRpt
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
    gSelectiveStationsFromImport lbcSortList, CkcAll, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.9
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmStationRpt
    gCenterForm frmStationRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim ilRet As Integer

    frmStationRpt.Caption = "Station Information Report - " & sgClientName
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.9
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    imckcAllIgnore = False
    lbcSortList.Height = frcSelection.Height - CkcAll.Height - 480  '255 for check box

    If optSortby(0).Value = True Then
        optSortby_Click 0           'activate the event
    Else
        optSortby(0).Value = True
    End If

    ilRet = gPopTitleNames()        'get all the title names
    ilRet = gPopOwnerNames()        'get all owner names from ARTT
    lbcTitles.Clear
    'place titles in list box for selection
    For iLoop = 0 To UBound(tgTitleInfo) - 1 Step 1
        lbcTitles.AddItem Trim$(tgTitleInfo(iLoop).sTitle)
        lbcTitles.ItemData(lbcTitles.NewIndex) = tgTitleInfo(iLoop).iCode
    Next iLoop
    
    gPopExportTypes cboFileType         '3-12-04
    cboFileType.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblExportDestination.ForeColor = vbBlack
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmStationArray
    tmStatInfoRst.Close
    tmPersonInfoRst.Close
    tmEmailInfoRst.Close
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmStationRpt = Nothing
End Sub

Private Sub frcSelection_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblExportDestination.ForeColor = vbBlack
End Sub

Private Sub lbcSortList_Click()
    If imckcAllIgnore Then
        Exit Sub
    End If
    If CkcAll.Value = vbChecked Then
        imckcAllIgnore = True
        CkcAll.Value = vbUnchecked
        imckcAllIgnore = False
    End If
End Sub

Private Sub lblExportDestination_Click()
    If lblExportDestination.Caption <> "" Then
        'Show exported file in explorer
        Shell "explorer.exe /select, " & Mid(lblExportDestination.Caption, 19), vbNormalFocus
    End If
End Sub

Private Sub lblExportDestination_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And Y > 0 Then
        lblExportDestination.ForeColor = vbBlue
    End If
End Sub

Private Sub optContact_Click(Index As Integer)
    If Index = 1 Then           'specfic cntact, showthe titles
        lbcSortList.Height = (frcSelection.Height - CkcAll.Height - lacTitles.Height - 480) / 2
        lbcTitles.Height = (frcSelection.Height - CkcAll.Height - lacTitles.Height - 480) / 2   'lbcSortList.Height
        lacTitles.Top = lbcSortList.Top + lbcSortList.Height + 30
        lbcTitles.Top = lacTitles.Top + lacTitles.Height + 30
        lacTitles.Visible = True
        lbcTitles.Visible = True
    Else
        lacTitles.Visible = False
        lbcTitles.Visible = False
        lbcSortList.Height = frcSelection.Height - CkcAll.Height - 480

    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to adobe
    Else
        cboFileType.Enabled = False
    End If
    
    If optRptDest(4).Value Then
        '10/8/2020 - CSV export
        frcSortBy.Visible = False
        frcContact.Visible = False
        frcSubsort.Visible = False
        ckcInclWebPW.Visible = False
    Else
        frcSortBy.Visible = True
        frcContact.Visible = True
        frcSubsort.Visible = True
        ckcInclWebPW.Visible = True
    End If
    lblExportDestination.Visible = False
End Sub
Private Sub optSortby_Click(Index As Integer)
Dim llLoop As Long
    CkcAll.Value = vbUnchecked          'uncheck the all selection for different list box selectivity
    ckcInclWebPW.Top = frcContact.Top + frcContact.Height + 120
    
    'TTP 9943
    cmdStationListFile.Visible = False
    If Index = 0 Then           'call letter
        cmdStationListFile.Visible = True
        frcSubsort.Visible = False
        lbcSortList.Clear
        For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            lbcSortList.AddItem Trim$(tgStationInfo(llLoop).sCallLetters) & ", " & Trim$(tgStationInfo(llLoop).sMarket)
            lbcSortList.ItemData(lbcSortList.NewIndex) = tgStationInfo(llLoop).iCode
        Next llLoop
    ElseIf Index = 1 Then      'market
        frcSubsort.Visible = False
        lbcSortList.Clear
        For llLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
            lbcSortList.AddItem Trim$(tgMarketInfo(llLoop).sName)
            lbcSortList.ItemData(lbcSortList.NewIndex) = tgMarketInfo(llLoop).lCode
        Next llLoop
    ElseIf Index = 2 Then       'owner
        lbcSortList.Clear
        frcSubsort.Visible = False
        For llLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
            lbcSortList.AddItem Trim$(tgOwnerInfo(llLoop).sName)
            lbcSortList.ItemData(lbcSortList.NewIndex) = tgOwnerInfo(llLoop).lCode
        Next llLoop
    ElseIf Index = 3 Then       'multicast has subsort
        lbcSortList.Clear
        frcSubsort.Visible = True
        ckcInclWebPW.Top = frcSubsort.Top + frcSubsort.Height + 120

        If optSubSort(0).Value Then     'owner subsort
            optSubSort_Click 0
        Else
            optSubSort(0).Value = True
        End If
    Else                                'msa market
        frcSubsort.Visible = False
        lbcSortList.Clear
        For llLoop = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
            lbcSortList.AddItem Trim$(tgMSAMarketInfo(llLoop).sName)
            lbcSortList.ItemData(lbcSortList.NewIndex) = tgMSAMarketInfo(llLoop).lCode
        Next llLoop
    End If
    
End Sub

Private Sub optSubSort_Click(Index As Integer)
Dim llLoop As Long
    CkcAll.Value = vbUnchecked
    If Index = 0 Then       'subsort by owner
        lbcSortList.Clear
        For llLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
            lbcSortList.AddItem Trim$(tgOwnerInfo(llLoop).sName)
            lbcSortList.ItemData(lbcSortList.NewIndex) = tgOwnerInfo(llLoop).lCode
        Next llLoop
    ElseIf Index = 1 Then      'subsort by market
        lbcSortList.Clear
        For llLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
            lbcSortList.AddItem Trim$(tgMarketInfo(llLoop).sName)
            lbcSortList.ItemData(lbcSortList.NewIndex) = tgMarketInfo(llLoop).lCode
        Next llLoop
    End If
End Sub
'
'       Build up the arrays for the station addresses & phone #s
'
'       <input> llIndex - index into array to build the rquired data
'               sField1 - field 1 of 2 that is concatenated with field 2, commas in between
'               sField2 - field 2 of 2 that is concatenated with field 1
'       <output> llIndex - next index to place more information
'
Public Sub mBuildLoc(llIndex As Long, sField1 As String, sField2 As String)
    If llIndex <= UBound(tmStationArray) Then
        If sField1 = "" And sField2 = "" Then       'do nothing if no address
            Exit Sub
        End If
        
        If sField2 = "" Then
            tmStationArray(llIndex).sAddress = sField1
        Else
            tmStationArray(llIndex).sAddress = sField1 + "  " + sField2
        End If
        llIndex = llIndex + 1
        ReDim Preserve tmStationArray(0 To llIndex) As STATIONARRAY
        'initialize for next time thru
        tmStationArray(llIndex).sAddress = ""
        tmStationArray(llIndex).sPersonInfo = ""
    End If
    
    Exit Sub
End Sub
'
'       Build up the arrays for the station personnel addresses & phone #s, fax, & email
'
'       <input> llIndex - index into array to build the rquired data
'               slperson - personnel information (name, title, phone #s, email)
'       <output> llIndex - next index to place more information
'
Public Sub mBuildPerson(llIndex As Long, slPersonInfo)
    If llIndex > UBound(tmStationArray) Then        'see if another entry needs to be created for this person
        llIndex = llIndex + 1
        ReDim Preserve tmStationArray(0 To llIndex) As STATIONARRAY
    End If
    
    tmStationArray(llIndex).sPersonInfo = slPersonInfo
    llIndex = llIndex + 1
    If llIndex > UBound(tmStationArray) Then
        ReDim Preserve tmStationArray(0 To llIndex) As STATIONARRAY
        'initialize for next time thru
        tmStationArray(llIndex).sAddress = ""
        tmStationArray(llIndex).sPersonInfo = ""
    End If
    
    Exit Sub
End Sub
'           mInsertIntoIVR - loop thru the StationArray table that has all the address and
'           personnel info.  Create 1 record for eachentry to be dumped
'           in Crystsl reports
'           <input>
'                    GenDate - generation date for table filtering to crystal
'                    GenTime - generation time for table filtering to crystal
'   ivrAddr - Address from shtt (address, city state,zip, country)
'   ivrKey - Personnel info from artt (phone #, fax, email)
'   ivrSpotKeyNo - running sequence # for all info to be printed for a single station,
'               init to 0 each new station
'   ivrChfCode - Stattion code (shttcode) to retrieve rank, dst, zone

'
'Sort option-->                Call Ltrs  |   DMA      |   Owner    |        Multicast       |    MSA
'                                            Market    |            |    Owner     market    |  Market
'-----------------------------------------|------------|------------|------------------------|----------
'sort keys major to minor:                                                                   |
' Key0   ivrpayaddr1              ----    |  -----     |   ----     |  MCastGroup  MCastGroup|   ------
' Key1  ivrpayaddr2           Call Ltrs   |  DMA Mkt   |   Owner    |   Owner      Market    |   MSA Mkt
' Key2  ivrpayaddr3           Market      |  Call Ltrs |   Mkt      |   Mkt        Owner     |   Call Ltrs
' Key3  ivrpayaddr4           Owner       |  Owner     |   Call Ltrs|   Call Ltrs  Call Ltrs |   Owner
'
'       9-28-06 build mktrank into prepass record to avoid outer joins which slows the sql call
Public Sub mInsertIntoIVR(sGenDate As String, sGenTime As String)
Dim llLoop As Integer
Dim slCallLetters As String
Dim slOwnerName As String
Dim slMktName As String
Dim slAddr As String
Dim slNumber As String
Dim llGroupID As Long
Dim slKey0 As String * 40
Dim slKey1 As String * 40
Dim slKey2 As String * 40
Dim slKey3 As String * 40
Dim slTemp As String * 40
Dim ilLen As Integer
Dim slMultiCastFlag As String * 1
Dim ilMktRank As Integer
Dim slMSAMktName As String
    
    slOwnerName = ""
    slMktName = ""
    llGroupID = 0
    ilMktRank = 0
    slMultiCastFlag = ""
    slMSAMktName = ""
    For llLoop = 0 To UBound(tmStationArray) - 1
        slCallLetters = Trim$(tmStatInfoRst!shttCallLetters)
        'Check owner
        If Not IsNull(tmStatInfoRst!arttLastName) Then
            slOwnerName = Trim$(tmStatInfoRst!arttLastName)
        Else
            If optSortby(2).Value = True Or (optSortby(3).Value = True And optSubSort(0).Value = True) Then       'sort by owner, or multicast and owner subsort and owner doesnt exist
                slOwnerName = "ZZZZZZZZZZZZZZZZZZZZ"    '20 Zs to sort to end
            Else                                    'leave blank if not by owner sort
                slOwnerName = ""
            End If
        End If
        'check MSA market name vs DMA market name
        If optSortby(4).Value = True Then           'msa MARKET
            If Not IsNull(tmStatInfoRst!metName) Then
                slMSAMktName = Trim$(tmStatInfoRst!metName)
                ilMktRank = tmStatInfoRst!metRank
             Else
                If optSortby(4).Value = True Or (optSortby(3).Value = True And optSubSort(1).Value = True) Then       'sort by MSA market, or multicast & MSA market subsort and MSA market doesnt exist
                    slMSAMktName = "ZZZZZZZZZZZZZZZZZZZZ"    '20 Zs to sort to end
                Else
                    slMSAMktName = ""                      'leave blank if not by MSA market
                End If
            End If
            
        Else                        'DMA Market
            If Not IsNull(tmStatInfoRst!mktName) Then
                slMktName = Trim$(tmStatInfoRst!mktName)
                ilMktRank = tmStatInfoRst!mktRank
            Else
                If optSortby(1).Value = True Or (optSortby(3).Value = True And optSubSort(1).Value = True) Then       'sort by market, or multicast & market subsort and market doesnt exist
                    slMktName = "ZZZZZZZZZZZZZZZZZZZZ"    '20 Zs to sort to end
                Else
                    slMktName = ""                      'leave blank if not by market
                End If
            End If
        End If
        'check groupid
        'If Not IsNull(tmStatInfoRst!mgtGroupID) Then
        '    llGroupID = tmStatInfoRst!mgtGroupID
        
        'If Not IsNull(tmStatInfoRst!shttMultiCastGroupID) Then
        If tmStatInfoRst!shttMultiCastGroupID > 0 Then
            llGroupID = tmStatInfoRst!shttMultiCastGroupID
            slMultiCastFlag = "*"                   'mark station as a multicast station on report
        End If
        
        slAddr = Trim$(tmStationArray(llLoop).sAddress)
        slNumber = Trim$(tmStationArray(llLoop).sPersonInfo)
        
        If optSortby(0).Value = True Then           'sort by Call ltrs
            slKey0 = ""
            slKey1 = slCallLetters
            slKey2 = slMktName
            slKey3 = slOwnerName
        ElseIf optSortby(1).Value = True Then       'sort by market
            slKey0 = ""
            slKey1 = slMktName
            slKey2 = slCallLetters
            slKey3 = slOwnerName
        ElseIf optSortby(2).Value = True Then      'owner
            slKey0 = ""
            slKey1 = slOwnerName
            slKey2 = slMktName
            slKey3 = slCallLetters
        ElseIf optSortby(3).Value = True Then      'multicast
            'make the group id a string to sort
            slKey0 = Trim$(Str$(llGroupID))
            ilLen = Len(Trim$(slKey0))
            Do While ilLen < 10
                slKey0 = "0" & slKey0
                ilLen = Len(Trim$(slKey0))
            Loop
             If optSubSort(0).Value = True Then            'subsort multicast by owner
                slKey1 = slOwnerName
                slKey2 = slMktName
                slKey3 = slCallLetters
            Else                                         'subsort multicast by market
                slKey1 = slMktName
                slKey2 = slOwnerName
                slKey3 = slCallLetters
            End If
        Else                                    'msa market
            slKey0 = ""
            slKey1 = slMSAMktName
            slKey2 = slCallLetters
            slKey3 = slOwnerName
        End If

        'handle any embedded single quotes in text
        slKey0 = gFixQuote(slKey0)
        slKey1 = gFixQuote(slKey1)
        slKey2 = gFixQuote(slKey2)
        slKey3 = gFixQuote(slKey3)
        slAddr = gFixQuote(slAddr)
        slNumber = gFixQuote(slNumber)
        
        SQLQuery = "INSERT INTO IVR_Invoice_Rpt"
        SQLQuery = SQLQuery & " (ivrPayAddr1, ivrPayAddr2 , ivrPayAddr3, ivrAddr1, ivrKey, ivrSpotKeyNo, ivrChfCode, ivrPayAddr4, ivrSpotType, ivrFormType, ivrGenDate, ivrGenTime) "

        SQLQuery = SQLQuery & " VALUES (" & "'" & Trim$(slKey0) & "', '" & Trim$(slKey1) & "', '" & Trim$(slKey2) & "', '" & slAddr & "', '" & slNumber + "', '" & llLoop & "', '" & tmStatInfoRst!shttCode & "', '" & Trim$(slKey3) & "', '" & slMultiCastFlag & "', " & ilMktRank & ", "
        SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "'," & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
        cnn.BeginTrans
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "StationRpt-mInsertIntoIVR"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
    Next llLoop
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "StationRpt-mInsertIntoIVR"
    Exit Sub
End Sub

Function mExportStationInfoCSV(slExportFilename)
    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    Dim olRs As ADODB.Recordset
    Dim slErrorMessage As String
    Dim blLogSuccess As Boolean
    Dim slLogSuccess As String
    
    lblExportDestination.Visible = True
    lblExportDestination.Caption = "Exporting..."
    DoEvents
        
    slErrorMessage = mQueryDatabaseCSV(olRs)
    If slErrorMessage = "No errors" Then
        slErrorMessage = mWriteCsv(olRs, slExportFilename)
    End If
    
    If slErrorMessage = "No errors" Then
        lblExportDestination.Caption = "StationInformation Export created. "
    Else
        lblExportDestination.Caption = "Errors writing 'StationInformation.csv' " & slErrorMessage & slLogSuccess
    End If
    
finish:
    Screen.MousePointer = vbDefault
    Set olRs = Nothing
    
    Exit Function

ERRORBOX:
    MsgBox "Errors exporting " & slExportFilename, vbOKOnly + vbInformation

End Function

Private Function mWriteCsv(ByRef olRs As Recordset, slExportFilename) As String
    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    'Create a CSV file, give me a recordset and a (fully qualified path\filename.ext) filename
    'Makes Headers from recordset Column Names
    'Makes rows from Data
    Dim slErrorMessage As String
    Dim olFileSys As FileSystemObject
    Dim olCsv As TextStream
    Dim slPath As String
    Dim slRowToWrite As String
    Dim slComma As String
    Dim slAppendLine As String
    Dim olField As Field
    Dim slFormattedString As String
    Dim slHeader As String
    slComma = ","
    Set olFileSys = New FileSystemObject
    
    On Error GoTo ERRORBOX
    If olFileSys.FolderExists(sgExportDirectory) Then
        Set olCsv = olFileSys.OpenTextFile(slExportFilename, ForWriting, True)
        'Get Header
        slHeader = mGetHeaderString(olRs)
        'Write Header
        olCsv.WriteLine slHeader
        'Wite Data rows
        If olRs.EOF And olRs.BOF Then
            mWriteCsv = "There are no records to write to 'StationInformation.csv'"
            GoTo finish
        End If
        olRs.MoveFirst
        Do While Not olRs.EOF
            slRowToWrite = ""
            For Each olField In olRs.Fields
                slFormattedString = mWriteField(olField)
                'not sure if need to test for first error string.
                If slFormattedString = "Error reading records in mWriteCsv" Or slFormattedString = "Error in function mWriteField" Then
                    mWriteCsv = slFormattedString
                    GoTo finish
                End If
                If slRowToWrite <> "" Then slRowToWrite = slRowToWrite & slComma
                slRowToWrite = slRowToWrite & slFormattedString
            Next olField
            olCsv.WriteLine slRowToWrite
            olRs.MoveNext
        Loop
        olRs.Close
        olCsv.Close
        slErrorMessage = "No errors"
    Else
        slErrorMessage = "Couldn't find Export Folder."
    End If
    mWriteCsv = slErrorMessage

finish:
    Set olField = Nothing
    Set olFileSys = Nothing
    Set olCsv = Nothing
    Exit Function

ERRORBOX:
    mWriteCsv = "Error reading records in mWriteCsv"
    GoTo finish
End Function

Private Function mGetHeaderString(ByRef olRs As Recordset) As String
    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    'Get the headers from the recordset columns (skipping the 1st Letter, which is a format indicator)
    mGetHeaderString = ""
    Dim slComma As String
    Dim olField As Field
    slComma = ","
    For Each olField In olRs.Fields
        If mGetHeaderString <> "" Then mGetHeaderString = mGetHeaderString & slComma
        mGetHeaderString = mGetHeaderString & Mid(olField.Name, 2)
    Next olField
End Function

Private Function mWriteField(olField As Field) As String
    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    'return a formatted string for the CSV export, based on the 1st letter (format indicator) of ColumnName
    Dim slReturnName As String
    Dim ilWriteToLineCode As Integer
    On Error GoTo ERRORBOX
    Select Case Left(olField.Name, 1)
        Case "s" 'String
            If IsNull(olField.Value) Then
                mWriteField = """"""
            Else
                mWriteField = Chr(34) & Trim(olField.Value) & Chr(34) 'Quotted
            End If
        Case "i" 'integer
            If IsNull(olField.Value) Then
                mWriteField = ""
            Else
                mWriteField = Trim(Str(Int(Val(olField.Value))))
            End If
        'could write date or other type handlers here, but the spec'd Station Information.csv export didnt have any other types to export
        Case Else
            mWriteField = Chr(34) & Trim(olField.Value) & Chr(34) 'Quotted
    End Select
    Exit Function
    
ERRORBOX:
    mWriteField = "Error in function mWriteField"
End Function

Private Function mQueryDatabaseCSV(ByRef olRs As Recordset) As String
    '10-08-20 - TTP 9985 - Export Station Information in CSV format
    'This returns the Column Names Exactly as they should appear in the export, with a prepended format character (like s=string, n=numeric); example: "sCall Letters"
    Dim slErrorMessage As String
    Dim SQLQuery As String
    Dim blNeedAnd As Boolean 'For Query building
    Dim lLoop As Integer
    SQLQuery = SQLQuery & " SELECT "
    SQLQuery = SQLQuery & "    shttCallLetters AS ""sCall Letters"","
    SQLQuery = SQLQuery & "    shttFrequency AS ""sFrequency"","
    SQLQuery = SQLQuery & "    MNT_Moniker.mntName AS ""sMoniker"","
    SQLQuery = SQLQuery & "    shttWatts AS ""iWatts"","
    SQLQuery = SQLQuery & "    shttPermStationID AS ""iPerm Station ID"","
    SQLQuery = SQLQuery & "    if(shttOnAir='N','N','Y') AS ""sOn Air"","
    SQLQuery = SQLQuery & "    if(shttStationType='N','N','C') AS ""sStation Type"","
    SQLQuery = SQLQuery & "    if(shttAckDaylight=0,'Y','N') AS ""sHonor DST"","
    SQLQuery = SQLQuery & "    shttAudP12Plus AS ""iP12+ Aud"","
    SQLQuery = SQLQuery & "    shttAddress1 AS ""sMail Address 1"","
    SQLQuery = SQLQuery & "    shttAddress2 AS ""sMail Address 2"","
    SQLQuery = SQLQuery & "    MNT_MailCity.mntName  AS ""sMail Address City"","
    SQLQuery = SQLQuery & "    shttState AS ""sMail Address State"","
    SQLQuery = SQLQuery & "    shttCountry AS ""sMail Address Country"","
    SQLQuery = SQLQuery & "    shttZip AS ""sMail Address Zip"","
    SQLQuery = SQLQuery & "    shttOnAddress1 AS ""sPhysical Address 1"","
    SQLQuery = SQLQuery & "    shttOnAddress2 AS ""sPhysical Address 2"","
    SQLQuery = SQLQuery & "    MNT_PhysCity.mntName as ""sPhysical Adress City"","
    SQLQuery = SQLQuery & "    shttOnState AS ""sPhysical Adress State"","
    SQLQuery = SQLQuery & "    shttOnZip AS ""sPhysical address Zip"","
    SQLQuery = SQLQuery & "    shttOnCountry AS ""sPhysical Address country"","
    SQLQuery = SQLQuery & "    MNT_CityLic.mntName AS ""sCity License"","
    SQLQuery = SQLQuery & "    MNT_CountyLic.mntName  AS ""sCounty License"","
    SQLQuery = SQLQuery & "    shttStateLic AS ""sState License"","
    SQLQuery = SQLQuery & "    shttPhone AS ""sPhone"","
    SQLQuery = SQLQuery & "    shttFax AS ""sFax"","
    SQLQuery = SQLQuery & "    FMT_Station_Format.fmtName AS ""sFormat"","
    SQLQuery = SQLQuery & "    artt.arttLastName as ""sOwner"","
    SQLQuery = SQLQuery & "    MKT.mktName as ""sMarket"","
    SQLQuery = SQLQuery & "    MKT.mktRank AS ""iMarket Rank"","
    SQLQuery = SQLQuery & "    MET.metName as ""sMetro name"","
    SQLQuery = SQLQuery & "    MET.metRank as ""iMetro Rank"","
    SQLQuery = SQLQuery & "    MNT_Operator.mntName as ""sOperator"","
    SQLQuery = SQLQuery & "    MNT_Territory.mntName as ""sTerritory"","
    SQLQuery = SQLQuery & "    MNT_Area.mntName as ""sArea"","
    SQLQuery = SQLQuery & "    UST_MarketRep.ustName AS ""sMarket Rep"","
    SQLQuery = SQLQuery & "    UST_ServiceRep.ustName AS ""sService Rep"","
    SQLQuery = SQLQuery & "    tzt.tztName as ""sTime Zone"","
    SQLQuery = SQLQuery & "    if(shttAgreementExist='Y','Y','N') AS ""sAgree Exist"","
    SQLQuery = SQLQuery & "    if(shttUsedForAtt='N','N','Y') AS ""sAgreement"","
    SQLQuery = SQLQuery & "    if(shttUsedForXDigital='Y','Y','N') AS ""sXDS"","
    SQLQuery = SQLQuery & "    if(shttUsedForWegener='Y','Y','N') AS ""sWegener"","
    SQLQuery = SQLQuery & "    if(shttUsedForOLA='Y','Y','N') AS ""sOLA"","
    SQLQuery = SQLQuery & "    if(shttPledgevsAir='Y','Y','N') AS ""sPledge vs Air"","
    SQLQuery = SQLQuery & "    shttStationID AS ""iXDS Station ID"","
    SQLQuery = SQLQuery & "    IF(shttWebNumber='2','2','1') AS ""iWeb version"","
    SQLQuery = SQLQuery & "    shttWebAddress AS ""sWeb Site"","
    SQLQuery = SQLQuery & "    shttWebPW AS ""sPassword"","
    SQLQuery = SQLQuery & "    shttVieroID AS ""sViero ID"","
    SQLQuery = SQLQuery & "    shttiPumpID AS ""siPump ID"","
    SQLQuery = SQLQuery & "    if(shttUsedForiPump='Y','Y','N') AS ""siPump"","
    SQLQuery = SQLQuery & "    shttSerialNo1 AS ""sStarguide Serial 1"","
    SQLQuery = SQLQuery & "    shttSerialNo2 AS ""sStarguide Serial 2"""
    SQLQuery = SQLQuery & " FROM "
    SQLQuery = SQLQuery & "    ""shtt"" "
    SQLQuery = SQLQuery & "    Left Join ""mnt"" AS ""MNT_Moniker"" on shttMonikerMntCode = MNT_Moniker.MntCode"
    SQLQuery = SQLQuery & "    Left Join ""mnt"" AS ""MNT_MailCity"" on shttCityMntCode = MNT_MailCity.MntCode"
    SQLQuery = SQLQuery & "    Left Join ""mnt"" AS ""MNT_PhysCity"" on shttOnCityMntCode = MNT_PhysCity.MntCode"
    SQLQuery = SQLQuery & "    Left Join ""mnt"" AS ""MNT_CityLic"" on shttCityLicMntCode = MNT_CityLic.MntCode"
    SQLQuery = SQLQuery & "    Left Join ""mnt"" AS ""MNT_CountyLic"" on shttCountyLicMntCode = MNT_CountyLic.MntCode"
    SQLQuery = SQLQuery & "    Left Join ""FMT_Station_Format"" on shttfmtCode = FMT_Station_Format.FmtCode"
    SQLQuery = SQLQuery & "    Left Join ""artt"" on shttOwnerArttCode = artt.arttCode"
    SQLQuery = SQLQuery & "    Left Join ""mkt"" AS ""MKT"" on Shttmktcode = MKT.mktCode"
    SQLQuery = SQLQuery & "    Left Join ""met"" AS ""MET"" on Shttmetcode = MET.metCode"
    SQLQuery = SQLQuery & "    Left Join ""mnt"" AS ""MNT_Operator"" on shttOperatorMntcode = MNT_Operator.MntCode"
    SQLQuery = SQLQuery & "    Left Join ""mnt"" AS ""MNT_Territory"" on shttmntCode = MNT_Territory.MntCode"
    SQLQuery = SQLQuery & "    Left Join ""mnt"" AS ""MNT_Area"" on shttAreaMntcode = MNT_Area.MntCode"
    SQLQuery = SQLQuery & "    Left Join ""ust"" AS ""UST_MarketRep"" on shttMktRepUstCode = UST_MarketRep.ustCode"
    SQLQuery = SQLQuery & "    Left Join ""ust"" AS ""UST_ServiceRep"" on shttServRepUstCode = UST_ServiceRep.ustCode"
    SQLQuery = SQLQuery & "    Left Join ""tzt"" on shttTztCode = tzt.tztCode"

    
    If optSP(0).Value Then
        SQLQuery = SQLQuery & " WHERE "
        SQLQuery = SQLQuery & " shttType = 0" 'Station
        blNeedAnd = True
    End If
    If optSP(1).Value Then
        SQLQuery = SQLQuery & " WHERE "
        SQLQuery = SQLQuery & " shttType = 1" 'People
        blNeedAnd = True
    End If
    If optSP(2).Value Then
        'both
    End If
    
    If CkcAll.Value = False Then
        If blNeedAnd = True Then
            SQLQuery = SQLQuery & " AND "
            blNeedAnd = False
        Else
            SQLQuery = SQLQuery & " WHERE "
        End If
        'Specific Advertisers
        SQLQuery = SQLQuery & " shttCode IN ("
        'Filter to Specific ID's
        For lLoop = 0 To lbcSortList.ListCount - 1 Step 1
            If lbcSortList.Selected(lLoop) Then
                If blNeedAnd = True Then
                    blNeedAnd = False
                    SQLQuery = SQLQuery & ","
                End If
                SQLQuery = SQLQuery & lbcSortList.ItemData(lLoop)
                blNeedAnd = True
            End If
        Next lLoop
        SQLQuery = SQLQuery & ")"
    Else
        'All Advertisers
    End If
    
    On Error GoTo ERRORBOX
    Set olRs = gSQLSelectCall(SQLQuery)
    mQueryDatabaseCSV = "No errors"
    Exit Function
    
ERRORBOX:
    mQueryDatabaseCSV = "Problem with query in mQueryDatabase. "
End Function

