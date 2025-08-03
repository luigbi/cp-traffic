VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRgAssignRpt 
   Caption         =   "Regional Copy Assignment"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "AffRgAssgnRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   9165
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
      FormDesignHeight=   5775
      FormDesignWidth =   9165
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
      Height          =   3780
      Left            =   195
      TabIndex        =   28
      Top             =   1830
      Width           =   8760
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   8280
         Picture         =   "AffRgAssgnRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Select Stations from File.."
         Top             =   1800
         Width           =   360
      End
      Begin VB.CheckBox ckcShowOnlyRegAsgn 
         Caption         =   "Show Only Regional Copy Assigned"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2070
         Value           =   1  'Checked
         Width           =   3045
      End
      Begin V81Affiliate.CSI_Calendar CalWeek 
         Height          =   285
         Left            =   1125
         TabIndex        =   7
         Top             =   225
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
         font            =   "AffRgAssgnRpt.frx":0E34
         csi_daynamefont =   "AffRgAssgnRpt.frx":0E60
         csi_monthnamefont=   "AffRgAssgnRpt.frx":0E8E
      End
      Begin V81Affiliate.CSI_Calendar CalEndDate 
         Height          =   285
         Left            =   2565
         TabIndex        =   8
         Top             =   225
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
         font            =   "AffRgAssgnRpt.frx":0EBC
         csi_daynamefont =   "AffRgAssgnRpt.frx":0EE8
         csi_monthnamefont=   "AffRgAssgnRpt.frx":0F16
      End
      Begin VB.CheckBox ckcExclSpotsLackReg 
         Caption         =   "Exclude Spots Lacking Regional Copy"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1455
         Width           =   3495
      End
      Begin VB.CheckBox ckcIncludeMiss 
         Caption         =   "Include Not Aired"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1770
         Width           =   1845
      End
      Begin VB.TextBox edcContract 
         Height          =   285
         Left            =   1020
         MaxLength       =   9
         TabIndex        =   16
         Top             =   2400
         Width           =   930
      End
      Begin VB.CheckBox ckcAllAdvt 
         Caption         =   "All Advertisers"
         Height          =   195
         Left            =   3780
         TabIndex        =   17
         Top             =   195
         Width           =   1380
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   195
         Left            =   6225
         TabIndex        =   23
         Top             =   1815
         Width           =   1380
      End
      Begin VB.CheckBox ckcAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   195
         Left            =   3750
         TabIndex        =   21
         Top             =   1800
         Width           =   1380
      End
      Begin VB.ListBox lbcVehicles 
         Height          =   1425
         ItemData        =   "AffRgAssgnRpt.frx":0F44
         Left            =   3720
         List            =   "AffRgAssgnRpt.frx":0F4B
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ListBox lbcStations 
         Height          =   1425
         ItemData        =   "AffRgAssgnRpt.frx":0F53
         Left            =   6210
         List            =   "AffRgAssgnRpt.frx":0F5A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Frame ShowSpotAired 
         Caption         =   "Sort by"
         Height          =   825
         Left            =   120
         TabIndex        =   10
         Top             =   555
         Width           =   3540
         Begin VB.OptionButton optSortby 
            Caption         =   "Station, Vehicle, Date Time"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   12
            Top             =   480
            Width           =   2310
         End
         Begin VB.OptionButton optSortby 
            Caption         =   "Vehicle, Station, Date Time"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   2505
         End
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All Contracts"
         Height          =   195
         Left            =   6180
         TabIndex        =   19
         Top             =   180
         Width           =   1860
      End
      Begin VB.ListBox lbcContract 
         Height          =   1230
         ItemData        =   "AffRgAssgnRpt.frx":0F61
         Left            =   6180
         List            =   "AffRgAssgnRpt.frx":0F63
         MultiSelect     =   2  'Extended
         TabIndex        =   20
         Top             =   480
         Width           =   2415
      End
      Begin VB.ListBox lbcAdvertiser 
         Height          =   1230
         ItemData        =   "AffRgAssgnRpt.frx":0F65
         Left            =   3750
         List            =   "AffRgAssgnRpt.frx":0F67
         MultiSelect     =   2  'Extended
         TabIndex        =   18
         Top             =   480
         Width           =   2340
      End
      Begin VB.Label Label3 
         Caption         =   "Contract #"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Dates- Start"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "End"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4845
      TabIndex        =   27
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4605
      TabIndex        =   26
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4410
      TabIndex        =   25
      Top             =   225
      Width           =   2685
   End
   Begin VB.Frame frame1 
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
         ItemData        =   "AffRgAssgnRpt.frx":0F69
         Left            =   840
         List            =   "AffRgAssgnRpt.frx":0F6B
         TabIndex        =   4
         Top             =   795
         Width           =   1935
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1185
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   840
         Width           =   690
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   2130
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmRgAssignRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'* Advertiser Posting (Clearance) Detail
'* 9/9/99 dh
'*
'*  Created July,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
 
Private smFWkDate As String
Private smLWkDate As String
Private imAdfCode As Integer
Private imAllClick As Integer
Private imAllVehClick As Integer
Private imAllAdvtClick As Integer
Private imAllStationsClick As Integer
Private hmAst As Integer

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
    If lbcContract.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcContract.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcContract.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If
End Sub

Private Sub ckcAllAdvt_Click()
 Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllAdvtClick Then
        Exit Sub
    End If
    If ckcAllAdvt.Value = vbChecked Then
        iValue = True
        lbcContract.Clear
        lbcContract.Visible = False
        chkAll.Visible = False
    Else
        iValue = False
        lbcContract.Visible = True
        chkAll.Visible = True
    End If
    If lbcAdvertiser.ListCount > 0 Then
        imAllAdvtClick = True
        lRg = CLng(lbcAdvertiser.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcAdvertiser.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllAdvtClick = False
    End If
End Sub

Private Sub ckcAllStations_Click()
 Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllStationsClick Then
        Exit Sub
    End If
    If ckcAllStations.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcStations.ListCount > 0 Then
        imAllStationsClick = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllStationsClick = False
    End If
End Sub

Private Sub ckcAllVehicles_Click()
 Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllVehClick Then
        Exit Sub
    End If
    If ckcAllVehicles.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllVehClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllVehClick = False
        ckcAllStations.Value = vbUnchecked
'        mPopStationsAndHide
    End If
End Sub

Private Sub cmdDone_Click()
    Unload frmRgAssignRpt
End Sub

Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim sCode As String
    Dim bm As Variant
    Dim sName, sVehicles, sStations As String
    Dim sDateRange As String
    Dim sContracts As String
    Dim sStatus, sCPStatus As String    'spot status and posting status flags
    Dim sStationType As String
    Dim sOutput As String
    Dim ilRet As Integer
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim sGenDate As String              'prepass generation date for crystal filtering
    Dim sGenTime As String              'prepass generation time
    Dim llGenTime As Long
    Dim ilAdvt As Integer
    'Dim NewForm As New frmViewReport
    Dim ilIncludeMissed As Integer
    Dim llSelectedContracts() As Long
    Dim llSingleCntr As Long
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String


    On Error GoTo ErrHand
    
    sGenDate = Format$(gNow(), "m/d/yyyy")
    sGenTime = Format$(gNow(), sgShowTimeWSecForm)
   
    If lbcAdvertiser.ListIndex < 0 Then
        gMsgBox "Advertiser must be specified.", vbOKOnly
        Exit Sub
    End If
    If CalWeek.Text = "" Then
        gMsgBox "Start Date must be specified.", vbOKOnly
        CalWeek.SetFocus
        Exit Sub
    End If
    If Trim$(CalEndDate.Text) = "" Then
        gMsgBox "End Date must be specified.", vbOKOnly
        CalEndDate.SetFocus
        Exit Sub
    End If
    If gIsDate(CalWeek.Text) = False Then       'is start date valid?
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        CalWeek.SetFocus
    Else
        smFWkDate = Format(CalWeek.Text, "m/d/yy")
    End If
    

    smLWkDate = Format$(CalEndDate.Text, "m/d/yy")
    
    Screen.MousePointer = vbHourglass
    gUserActivityLog "S", sgReportListName & ": Prepass"
    
    'debugging only for timing tests
    Dim sGenStartTime As String
    Dim sGenEndTime As String
    sGenStartTime = Format$(gNow(), sgShowTimeWSecForm)
    
    If ckcIncludeMiss.Value = vbChecked Then      'include missed
        ilIncludeMissed = True
    Else
        ilIncludeMissed = False
    End If
    
    ReDim llSelectedContracts(0 To 0) As Long
    llSingleCntr = Val(edcContract.Text)
    llSelectedContracts(0) = 0
    If llSingleCntr > 0 Then
        'SQLQuery = "Select chfCode FROM CHF_Contract_Header where chfCntrno = " & llSingleCntr & " and chfDelete <> 'Y' and (chfStatus = 'H' or chfstatus = 'O')"
        'Set rst = gSQLSelectCall(SQLQuery)
        'While Not rst.EOF
            'llSelectedContracts(0) = rst!chfcode
            llSelectedContracts(0) = llSingleCntr
            ReDim Preserve llSelectedContracts(0 To 1) As Long
            'rst.MoveNext
        'Wend
    Else
         For i = 0 To lbcContract.ListCount - 1
            If lbcContract.Selected(i) Then
                'get the contract code from the contract #
                'SQLQuery = "Select chfCode FROM CHF_Contract_Header where chfCntrno = " & lbcContract.List(i) & " and chfDelete <> 'Y' and (chfStatus = 'H' or chfstatus = 'O')"
                'Set rst = gSQLSelectCall(SQLQuery)
                'While Not rst.EOF
                    'llSelectedContracts(UBound(llSelectedContracts)) = rst!chfcode
                    llSelectedContracts(UBound(llSelectedContracts)) = lbcContract.List(i)
                    ReDim Preserve llSelectedContracts(0 To UBound(llSelectedContracts) + 1) As Long
                    'rst.MoveNext
                'Wend
            End If
        Next i
    End If
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False
    
    '4-17-08 always exclude spots not carried, bBuildAstStnClr tests for the status to exclude them.  last parameter is a flag to show exact station feed (exclude not carried if true)
    gGenRegionalCopyRept hmAst, smFWkDate, smLWkDate, lbcVehicles, lbcStations, lbcAdvertiser, sGenDate, sGenTime, True, False, ilIncludeMissed, llSelectedContracts()
    
    'CRpt1.Connect = "DSN = " & sgDatabaseName
    If optRptDest(0).Value = True Then
        'CRpt1.Destination = crptToWindow
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        'CRpt1.Destination = crptToPrinter
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        'gOutputMethod frmRgAssignRpt, "AfAdvClr", sOutput
        'ilExportType = cboFileType.ItemData(cboFileType.ListIndex)
        ilExportType = cboFileType.ListIndex    'select the user input
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
  
    Screen.MousePointer = vbHourglass
    
    If lbcAdvertiser.ListIndex < 0 Then
        Screen.MousePointer = vbDefault
        gMsgBox "Advertiser must be selected.", vbOKOnly
        lbcAdvertiser.SetFocus
        Exit Sub
    End If
    sAdvtDates = Trim$(lbcAdvertiser.List(lbcAdvertiser.ListIndex))
    sContracts = ""
    sStatus = ""
    sCPStatus = ""
   
    
    'The SQL statementis currently not used in the selection in V5.7
    'Only the selection is used filtering on grfGenDate and grfGenTime.
    'All the joins are used from the crystal report

    '3-8-17 reinstate the sql calls; otherwise might get too much data or bad data
    SQLQuery = "SELECT grfrdfCode, grfDateType, grfBktType, grfCode2, grfCode4,"
    SQLQuery = SQLQuery & "grfYear, grfGenDesc, grfDate, grfTime, grfLong, "
    SQLQuery = SQLQuery & "grfPer1, grfPer2, grfPer4, grfPer1Genl, grfPer2Genl, grfPer4Genl, "
    SQLQuery = SQLQuery & "shttCallLetters, shttState, chfcntrno, chfProduct, "
    SQLQuery = SQLQuery & "vefName, RafName, cifName, cifcpfCode, crfInOut, crfBkOutInstAdfCode, "
    SQLQuery = SQLQuery & "mcfName, cpf_Copy_Prodct_ISCI.cpfISCI, cpf_Copy_Prodct_ISCI.cpfCreative, cpf_Copy_Prodct_ISCI.cpfName, AnfName, Adf_Advertisers.AdfName, "
    SQLQuery = SQLQuery & " Adf_BlkOutAdvertisers.adfName, mktname, attLoad, "
    SQLQuery = SQLQuery & " fmtName, tztName, metName, vefName, vpfAllowSplitCopy "
    SQLQuery = SQLQuery & " VEF_RotVehicles.vefname, VEF_RotVehicles.sType "
    SQLQuery = SQLQuery & " from grf_generic_report "
    SQLQuery = SQLQuery & "INNER JOIN VEF_Vehicles on grfvefCode = vefCode "
    SQLQuery = SQLQuery & "INNER JOIN VPF_Vehicle_Options on vefcode = vpfvefkcode "
    SQLQuery = SQLQuery & "INNER JOIN Chf_Contract_Header on grfchfcode = chfcode "
    SQLQuery = SQLQuery & "INNER JOIN Adf_Advertisers on chfAdfCode = adfcode "
    SQLQuery = SQLQuery & "INNER JOIN ast on grfPer4 = astcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN cif_Copy_Inventory on grfcode4 = cifcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN cpf_Copy_Prodct_ISCI on cifcpfCode = cpfcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN mcf_Media_Code on cifmcfcode = mcfcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN crf_Copy_Rot_Header on grfLong = crfcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN anf_Avail_Names on crfanfCode = anfCode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN Raf_Region_Area on grfPer1 = rafcode "
    SQLQuery = SQLQuery & "INNER JOIN Shtt on grfSofCode = shttcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN mkt on shttMktCode = mktcode "
    SQLQuery = SQLQuery & "INNER JOIN att on grfPer3 = attcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN fmt_Station_Format on shttfmtCode = fmtcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN met on shttMetCode = metcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN tzt on shtttztCode = tztcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN ADF_Advertisers ADF_BlkoutAdvertisers on crfBkoutInstAdfCode = adf_blkoutAdvertisers.adfcode "
    SQLQuery = SQLQuery & "LEFT OUTER JOIN VEF_Vehicles VEF_RotVehicles ON CRF_Copy_Rot_Header.crfvefCode = VEF_RotVehicles.vefCode "
    SQLQuery = SQLQuery & "INNER JOIN CPF_Copy_Prodct_ISCI CPF_Copy_ChgProdct_ISCI_1 ON astCpfCode = CPF_Copy_ChgProdct_ISCI_1.cpfCode"

    SQLQuery = SQLQuery & " where ( grfGenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & ")"

   ' SQLQuery = "Select * from grf_generic_Rpt where ( grfGenDate = " & "'" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = " & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & ")"

    sgCrystlFormula1 = "'" & Format$(smFWkDate, "mm/dd/yy") & " - " & Format$(smLWkDate, "mm/dd/yy") & "'"
                    
    If optSortby(0).Value Then         'Sort by vehicle
        sgCrystlFormula2 = "'V'"
    Else                                'Sort by station
        sgCrystlFormula2 = "'S'"
    End If
    
    If ckcShowOnlyRegAsgn.Value = vbChecked Then
        sgCrystlFormula5 = "'Y'"
    Else
        sgCrystlFormula5 = "'N'"
    End If
      
    llGenTime = Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False)))))
    sgCrystlFormula3 = llGenTime
    slYear = Format(sGenDate, "yyyy")
    slMonth = Format(sGenDate, "mm")
    slDay = Format(sGenDate, "dd")
    sgCrystlFormula4 = " Date(" & Trim$(slYear) & "," & Trim$(slMonth) & "," & Trim$(slDay) & ")"
    'SQLQuery = ""
    
    gUserActivityLog "E", sgReportListName & ": Prepass"
    If igRptIndex = REGIONASSIGN_RPT Then
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfRegCopy.rpt", "AfRegCopy"               'external report
    Else
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfRegCopyTrace.rpt", "AfRegCopyTrace"     'internal report
    End If
    
    'debugging only for timing tests
    sGenEndTime = Format$(gNow(), sgShowTimeWSecForm)
    'gMsgBox sGenStartTime & "-" & sGenEndTime

     'remove all the records just printed
    SQLQuery = "DELETE FROM grf_Generic_Report "
    SQLQuery = SQLQuery & " WHERE (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
    cnn.BeginTrans
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "RgAssignRpt-cmdReport_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    
    cmdReport.Enabled = True               're-enable Gen, done buttons
    cmdDone.Enabled = True
    cmdReturn.Enabled = True
   
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRgAssignRpt-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmRgAssignRpt
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
    gSelectiveStationsFromImport lbcStations, ckcAllStations, Trim$(CommonDialog1.fileName)
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
    gSetFonts frmRgAssignRpt
    gCenterForm frmRgAssignRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim lRg As Long
    Dim lRet As Long
        
    'chkAll.Visible = False
    igRptIndex = frmReports!lbcReports.ItemData(frmReports!lbcReports.ListIndex)
    If igRptIndex = REGIONASSIGN_RPT Then
        frmRgAssignRpt.Caption = "Regional Affiliate Copy Assignment Report - " & sgClientName
    Else
        frmRgAssignRpt.Caption = "Regional Affiliate Copy Tracing Report - " & sgClientName
    End If
    smFWkDate = ""
    smLWkDate = ""
    imAllClick = False
    imAllVehClick = False
    imAllAdvtClick = False
    
'    'populate the Stations, Vehicles & Advertisers (currently only advertisers are selectable)
'    lbcStations.Clear
'    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
'            If tgStationInfo(iLoop).iType = 0 Then
'                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
'                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
'            End If
'        End If
'    Next iLoop
'    'force all stations selected
'    lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
'    lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, True, lRg)
    
    lbcVehicles.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    'force all vehicles selected
    'lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
    'lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, True, lRg)
    
    mFillAdvt
    mPopStationsAndHide

    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
      
    gPopExportTypes cboFileType         '3-15-04
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    cboFileType.Enabled = False         'disable the export types since display mode is default
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path

    Set frmRgAssignRpt = Nothing
End Sub



Private Sub lbcAdvertiser_Click()
    Dim sCode As String
    Dim ilLoop As Integer
    On Error GoTo ErrHand
    
    lbcContract.Clear
    chkAll.Value = 0        'chged from False to 0 10-22-99

    If lbcAdvertiser.ListIndex < 0 Then
        Exit Sub
    End If
    If CalWeek.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        CalWeek.SetFocus
        Exit Sub
    End If
    
    If gIsDate(CalWeek.Text) = False Then
        Beep
        gMsgBox "Please enter a valid start date (m/d/yy).", vbCritical
        CalWeek.SetFocus
    Else
        smFWkDate = Format(CalWeek.Text, sgShowDateForm)
    End If
    
    If CalWeek.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        CalEndDate.SetFocus
        Exit Sub
    End If
    
    If gIsDate(CalEndDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid end date (m/d/yy).", vbCritical
        CalEndDate.SetFocus
    Else
        smLWkDate = Format(CalEndDate.Text, sgShowDateForm)
    End If
    
    Screen.MousePointer = vbHourglass
    'smFWkDate & smLWkDAte = earliest/latest requested dates
    
    For ilLoop = 0 To lbcAdvertiser.ListCount - 1
        If lbcAdvertiser.Selected(ilLoop) Then
            imAdfCode = lbcAdvertiser.ItemData(ilLoop)
            
'            SQLQuery = "SELECT DISTINCT lstCntrNo from lst"
'            SQLQuery = SQLQuery + " WHERE ((lstLogDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
'            SQLQuery = SQLQuery + " AND lstAdfCode = " & imAdfCode & ")"
'            SQLQuery = SQLQuery + " ORDER BY lstCntrNo"
            
            '7-24-19 change access of unique contract #s from lst to sdf for speedup
            SQLQuery = "SELECT DISTINCT chfCntrNo from sdf_spot_detail inner join chf_contract_header on sdfchfcode = chfcode "
            SQLQuery = SQLQuery + " WHERE ((sdfDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND sdfDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')"
            SQLQuery = SQLQuery + " AND sdfAdfCode = " & imAdfCode & ")"
            SQLQuery = SQLQuery + " ORDER BY chfCntrNo"

            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                chkAll.Visible = True
                lbcContract.Visible = True
                ckcAllAdvt.Value = vbUnchecked
            End If
            While Not rst.EOF
'                lbcContract.AddItem rst!lstCntrNo  ', " & rst(1).Value & ""
                lbcContract.AddItem rst!chfCntrNo  ', " & rst(1).Value & ""
                rst.MoveNext
            Wend
        End If
    Next ilLoop
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRgAssignRpt-lbcAdvertiser"
End Sub

Private Sub lbcContract_Click()
  If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0        'chged from False to 0 10-22-99
        imAllClick = False
    End If
End Sub

Private Sub lbcStations_Click()
 If imAllStationsClick Then
        Exit Sub
    End If
    If ckcAllStations.Value = 1 Then
        imAllStationsClick = True
        ckcAllStations.Value = 0        'chged from False to 0 10-22-99
        imAllStationsClick = False
    End If
End Sub

Private Sub lbcVehicles_Click()
Dim ilLoop As Integer
Dim ilVefCode As Integer
If imAllVehClick Then
        Exit Sub
    End If
'    lbcStations.Clear
    If ckcAllVehicles.Value = vbChecked Then
        imAllVehClick = True
        ckcAllVehicles.Value = 0        'chged from False to 0 10-22-99
        imAllVehClick = False
        ckcAllStations.Value = vbUnchecked
        'populate all
'        mPopStationsAndHide
    Else                        'not all vehicles selected, or multiple vehicles selected
        If lbcVehicles.SelCount > 1 Then       'more than 1 vehicle selected, assume all stations
            'populate all
'            mPopStationsAndHide
        Else                                    'only one vehicle, populate the associated affiliates
            lbcStations.Visible = True
            ckcAllStations.Visible = True
            ckcAllStations.Value = vbUnchecked
            For ilLoop = 0 To lbcVehicles.ListCount - 1
                If lbcVehicles.Selected(ilLoop) Then
                    ilVefCode = lbcVehicles.ItemData(ilLoop)
                    gFillStations ilVefCode, lbcStations
                End If
            Next ilLoop
        End If
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0           'default to PDF
    Else
        cboFileType.Enabled = False
    End If
End Sub



Private Sub mFillAdvt()
    Dim iNoWeeks As Integer
    Dim iFound As Integer
    Dim iLoop As Integer
    On Error GoTo ErrHand
    
    lbcAdvertiser.Clear
    lbcContract.Clear
    'chkAll.Value = False
    chkAll.Value = 0        'chged from False to 0 10-22-99
    'SQLQuery = "SELECT adf.adfName, adf.adfCode from ADF_Advertisers adf"
    SQLQuery = "SELECT adfName, adfCode"
    SQLQuery = SQLQuery & " FROM ADF_Advertisers"
    SQLQuery = SQLQuery + " ORDER BY adfName"
  
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        iFound = False
    
        If Not iFound Then
            lbcAdvertiser.AddItem rst!adfName '& ", " & rst(1).Value
            lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = rst!adfCode
        End If
        rst.MoveNext
    Wend
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRgAssignRpt-mFillAdvt"
End Sub

Public Sub mPopStationsAndHide()
Dim iLoop As Integer
Dim lRg As Long
Dim lRet As Long
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
    'force all stations selected
    'lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
    'lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, True, lRg)
    'lbcStations.Visible = False
    'ckcAllStations.Value = vbChecked
    'ckcAllStations.Visible = False
End Sub
