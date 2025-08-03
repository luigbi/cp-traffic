VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWebVendorRpt 
   Caption         =   "Web Vendor"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "AffWebVendorRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3510
      Top             =   1080
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6630
      FormDesignWidth =   9360
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4590
      TabIndex        =   4
      Top             =   255
      Width           =   1935
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
      Height          =   4750
      Left            =   240
      TabIndex        =   1
      Top             =   1785
      Width           =   8895
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   8280
         Picture         =   "AffWebVendorRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Select Stations from File.."
         Top             =   360
         Width           =   360
      End
      Begin VB.ListBox lbcVendors 
         Height          =   1230
         ItemData        =   "AffWebVendorRpt.frx":0E34
         Left            =   4440
         List            =   "AffWebVendorRpt.frx":0E36
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2085
      End
      Begin VB.CheckBox ckcAllVendors 
         Caption         =   "All Vendors"
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ListBox lbcStation 
         Height          =   2400
         ItemData        =   "AffWebVendorRpt.frx":0E38
         Left            =   6600
         List            =   "AffWebVendorRpt.frx":0E3F
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   720
         Width           =   2085
      End
      Begin VB.ListBox lbcVehicle 
         Height          =   2400
         ItemData        =   "AffWebVendorRpt.frx":0E46
         Left            =   4440
         List            =   "AffWebVendorRpt.frx":0E48
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   2085
      End
      Begin VB.CheckBox chkAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   6600
         TabIndex        =   21
         Top             =   360
         Width           =   1245
      End
      Begin VB.CheckBox chkAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4440
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin V81Affiliate.CSI_Calendar CalToDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   960
         Width           =   915
         _extentx        =   1614
         _extenty        =   503
         text            =   "12/13/2022"
         borderstyle     =   1
         csi_showdropdownonfocus=   -1  'True
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_caldateformat=   1
         font            =   "AffWebVendorRpt.frx":0E4A
         csi_daynamefont =   "AffWebVendorRpt.frx":0E76
         csi_monthnamefont=   "AffWebVendorRpt.frx":0EA4
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0   'False
         csi_allowblankdate=   -1  'True
         csi_allowtfn    =   0   'False
         csi_defaultdatetype=   0
      End
      Begin V81Affiliate.CSI_Calendar CalFromDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   915
         _extentx        =   1614
         _extenty        =   503
         text            =   "12/13/2022"
         borderstyle     =   1
         csi_showdropdownonfocus=   -1  'True
         csi_inputboxboxalignment=   0
         csi_calbackcolor=   16777130
         csi_caldateformat=   1
         font            =   "AffWebVendorRpt.frx":0ED2
         csi_daynamefont =   "AffWebVendorRpt.frx":0EFE
         csi_monthnamefont=   "AffWebVendorRpt.frx":0F2C
         csi_curdaybackcolor=   16777215
         csi_curdayforecolor=   0
         csi_forcemondayselectiononly=   0   'False
         csi_allowblankdate=   -1  'True
         csi_allowtfn    =   -1  'True
         csi_defaultdatetype=   0
      End
      Begin VB.Frame frcSortBy 
         Caption         =   "Sort by"
         Height          =   1140
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   2835
         Begin VB.OptionButton rbcSortby 
            Caption         =   "Vendor, Station, Vehicle"
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   19
            Top             =   760
            Width           =   2565
         End
         Begin VB.OptionButton rbcSortby 
            Caption         =   "Station, Vehicle, Vendor"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   18
            Top             =   210
            Value           =   -1  'True
            Width           =   2280
         End
         Begin VB.OptionButton rbcSortby 
            Caption         =   "Vehicle, Station, Vendor"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   17
            Top             =   480
            Width           =   2565
         End
      End
      Begin VB.Frame frcShowBy 
         Caption         =   "Show Vendor"
         Height          =   585
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2100
         Begin VB.OptionButton rbcShowBy 
            Caption         =   "Import"
            Height          =   210
            Index           =   1
            Left            =   1080
            TabIndex        =   11
            Top             =   240
            Width           =   945
         End
         Begin VB.OptionButton rbcShowBy 
            Caption         =   "Export"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.Label LabTo 
         Caption         =   "End:"
         Height          =   240
         Left            =   2400
         TabIndex        =   14
         Top             =   960
         Width           =   330
      End
      Begin VB.Label labFrom 
         Caption         =   "Log Date- Start:"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4590
      TabIndex        =   2
      Top             =   720
      Width           =   1935
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
         ItemData        =   "AffWebVendorRpt.frx":0F5A
         Left            =   1050
         List            =   "AffWebVendorRpt.frx":0F5C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   825
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmWebVendorRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hmFrom As Integer

Private Type VENDOR_INFO
    sGenDate As String
    sGenTime As String
    lAttCode As Long
    iVendorCode As Integer
    sLogDate As String * 10
    lSpotCount As Long
    sProcessDateTime As String * 30
End Type

Private tmVendor_Info As VENDOR_INFO

'Date: 8/2/2018 FYM
Private Type VEHICLE_LIST
    iVehicleCode As Integer
End Type
Private tmVehicleList() As VEHICLE_LIST

'constant index within export/import line
Private Const ATTCODEINDEX = 1
Private Const VENDORINDEX = 2  '1
Private Const LOGDATE = 3  '2
Private Const SPOTCOUNT = 4
Private Const PROCESSDATETIME = 5

Private imSort1 As Integer           '0 = station, 1 = vehicle,2 = Vendor
Private imSort2 As Integer           '0 = none, 1 = station, 2 =vehicle, 3 = vendor
Private imSort3 As Integer           '0 = none, 1 = station, 2 =vehicle, 3 = vendor
Private imChkAllVehiclesIgnore As Integer
Private imChkAllStationsIgnore As Integer
Private imCkcAllVendorsIgnore As Integer
Private smUsingUnivision As String * 1

Private tmVendorList() As VendorInfo
Private tmSelectedVendors() As VendorInfo

Private rstATT As ADODB.Recordset
Private imIncludeCodes As Integer   'include or exclude the code list
Private imUseCodes() As Integer     'array of stations to include

Private Function mEnableGenerateReportButton()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    If (CalFromDate.Text <> "" And CalToDate.Text <> "" And _
        ((lbcVehicle.SelCount > 0 And lbcStation.SelCount > 0 And lbcVendors.SelCount > 0) Or _
        (chkAllStations.Value = vbChecked And chkAllVehicles.Value = vbChecked And ckcAllVendors.Value = vbChecked))) Then
        cmdReport.Enabled = True
    Else
        cmdReport.Enabled = False
    End If
End Function


Private Sub CalFromDate_Change()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub CalToDate_Change()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub chkAllStations_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkAllStationsIgnore Then
        Exit Sub
    End If
    If chkAllStations.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStation.ListCount > 0 Then
        imChkAllStationsIgnore = True
        lRg = CLng(lbcStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkAllStationsIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton
    Screen.MousePointer = vbDefault

End Sub

Private Sub chkAllVehicles_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkAllVehiclesIgnore Then
        Exit Sub
    End If
    If chkAllVehicles.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehicle.ListCount > 0 Then
        imChkAllVehiclesIgnore = True
        lRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicle.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkAllVehiclesIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton

    Screen.MousePointer = vbDefault

End Sub

Private Sub ckcAllVendors_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imCkcAllVendorsIgnore Then
        Exit Sub
    End If
    If ckcAllVendors.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVendors.ListCount > 0 Then
        imCkcAllVendorsIgnore = True
        lRg = CLng(lbcVendors.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVendors.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imCkcAllVendorsIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdReport_Click()
 
    Dim ilRptDest As Integer        'output to display, print, save to
    Dim slExportName As String      'name given to a SAVE-TO file
    Dim ilExportType As Integer     'SAVE-TO output type
    Dim slRptName As String         'full report name of crystal .rpt
    Dim slFromFile As String
    Dim llFromDate As Long
    Dim sFromDate As String
    Dim llToDate As Long
    Dim sToDate As String
    Dim slGenDate As String
    Dim slGenTime As String
    Dim ilOk As Integer
    Dim dFWeek As Date
    Dim blIsExport As Boolean
    Dim slFilePath As String
    Dim llCount As Long
    Dim ilLoop As Integer
    Dim ilTemp As Integer
    Dim llVefCode As Long
    Dim ilVefCode As Integer
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim blVendorFound As Boolean
    Dim blStationFound As Boolean
    Dim ilVendorId As Integer
    Dim ilShttCode As Integer
    Dim llAttCode As Long
    Dim slService As String
    Dim slType As String * 1
    Dim sGenDate As String      'generation date for filtering prepass records
    Dim sGenTime As String      'generation time for filtering prepass records
    
    On Error GoTo ErrHand
    '9747
    If Not gWebAccessTestedOk Then
        Beep
        gMsgBox "You do not have access to the Web.  Report cannot be run.", vbCritical
        Exit Sub
    End If
    sStartDate = Trim$(CalFromDate.Text)
    If sStartDate = "" Then
        sStartDate = "1/1/1970"
    End If
    sEndDate = Trim$(CalToDate.Text)
    If sEndDate = "" Then
        sEndDate = "12/31/2069"
    End If
    If gIsDate(sStartDate) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalFromDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sEndDate) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalToDate.SetFocus
        Exit Sub
    End If
        
    If optRptDest(0).Value = True Then
        ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilExportType = cboFileType.ListIndex
        ilRptDest = 2
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gUserActivityLog "S", sgReportListName & ": Prepass"
    'guide gets to see key codes
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        sgCrystlFormula1 = 1
    Else
        sgCrystlFormula1 = 0
    End If
    slRptName = "AfWebVendor.rpt"
    slExportName = "AfWebVendor"
       
    'gUserActivityLog "S", sgReportListName & ": Prepass"
    
    sStartDate = Format(sStartDate, "m/d/yyyy")
    sEndDate = Format(sEndDate, "m/d/yyyy")
    sDateRange = " (attOffAir >=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'" & " And attDropDate >=" & "'" + Format$(sStartDate, sgSQLDateForm) & "'" & " And attOnAir <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "')"

    slGenDate = Format$(gNow(), "m/d/yyyy")
    slGenTime = Format$(gNow(), sgShowTimeWSecForm)

    If CalFromDate.Text = "" Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalFromDate.SetFocus
        Exit Sub
    End If
    
    sFromDate = CalFromDate.Text
    
    llFromDate = DateValue(gAdjYear(sFromDate))
    sToDate = CalToDate.Text
    If Trim$(sToDate) = "" Then         'no end date enterd, make same as from date
        sToDate = sFromDate
    End If
    llToDate = DateValue(gAdjYear(sToDate))
    
    If gIsDate(sFromDate) = False Or (Len(Trim$(sFromDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalFromDate.SetFocus
        Exit Sub
    End If
    If gIsDate(sToDate) = False Or (Len(Trim$(sToDate)) = 0) Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
        CalToDate.SetFocus
        Exit Sub
    End If
    
    dFWeek = CDate(sFromDate)
    sgCrystlFormula1 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
 
    dFWeek = CDate(sToDate)
    sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
    sgCrystlFormula3 = False            'assume txt file exists
    
    Screen.MousePointer = vbHourglass
    
      
    If rbcShowBy(0).Value Then          'Export
        sgCrystlFormula1 = "E"
    Else
        sgCrystlFormula1 = "I"          'import
    End If
    
    sgCrystlFormula2 = sFromDate & "-" & sToDate        'user entered start/end log dates for crystl heading
    
    If rbcSortby(0).Value Then          'Station
        sgCrystlFormula3 = "S"
    ElseIf rbcSortby(1).Value Then
        sgCrystlFormula3 = "V"          'vehicle
    Else
        sgCrystlFormula3 = "N"          'vendor
    End If

    blIsExport = True
    If rbcShowBy(1).Value Then          'import
        blIsExport = False
    End If
    
    cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
    cmdDone.Enabled = False
    cmdReturn.Enabled = False

    ' create list array for selected VENDORS    Date: 8/2/2018    FYM
    ReDim tmSelectedVendors(0 To 0) As VendorInfo
    llCount = 0
    For ilTemp = 0 To lbcVendors.ListCount - 1
        If lbcVendors.Selected(ilTemp) Then
            For ilLoop = 0 To UBound(tmVendorList) - 1
                If tmVendorList(ilLoop).iIdCode = lbcVendors.ItemData(ilTemp) Then
                    LSet tmSelectedVendors(llCount) = tmVendorList(ilLoop)
                    llCount = llCount + 1
                    ReDim Preserve tmSelectedVendors(0 To llCount) As VendorInfo
                    Exit For
                End If
            Next ilLoop
        End If
    Next ilTemp
    
    'create list array for selected STATIONS    Date: 8/2/2018    FYM
    ReDim imUseCodes(0 To 0) As Integer
    gObtainCodes lbcStation, imIncludeCodes, imUseCodes()        'build array of which codes to incl/excl

    'create the array for the selected VEHICLES Date: 8/2/2018 FYM
    ReDim tmVehicleList(0 To 0) As VEHICLE_LIST
    llCount = 0
    For ilTemp = 0 To lbcVehicle.ListCount - 1
        If lbcVehicle.Selected(ilTemp) Then
            tmVehicleList(UBound(tmVehicleList)).iVehicleCode = lbcVehicle.ItemData(ilTemp)
            ReDim Preserve tmVehicleList(0 To ilTemp) As VEHICLE_LIST
        End If
    Next ilTemp
    
    slFilePath = gVendorReportInfoToText(blIsExport, sFromDate, sToDate)
    ilOk = mReadFile(slFilePath, llFromDate, llToDate, slGenDate, slGenTime)
    If Not ilOk Then
        'error with web import log file, does not exist in message folder
        'gMsg = "VendorFile.Txt does not exist in " & sgImportDirectory & " folder"
        gMsg = "No data exists for selected parameters"
        gMsgBox gMsg, vbCritical
        sgCrystlFormula3 = True         'file does not exist
        cmdReport.Enabled = True            'give user back control to gen, done buttons
        cmdDone.Enabled = True
        cmdReturn.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    gUserActivityLog "E", sgReportListName & ": Prepass"
        
    SQLQuery = "Select * from afr "
    SQLQuery = SQLQuery & "INNER JOIN att ON afrattCode = attCode "
    SQLQuery = SQLQuery & "INNER JOIN wvt_Vendor_Table on afrseqno = wvtVendorID "
    SQLQuery = SQLQuery & "INNER JOIN VEF_Vehicles ON attVefCode =vefCode "
    SQLQuery = SQLQuery & "INNER JOIN shtt ON attshfCode = shttCode"
    SQLQuery = SQLQuery + " where ( afrGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' AND afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"

    'dan todo change for rollback
    'frmCrystal.gCrystlReports "", ilExportType, ilRptDest, slRptName, slExportName, True
    frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName

    'remove all the records just printed
    SQLQuery = "DELETE FROM afr "
    SQLQuery = SQLQuery & " WHERE (afrGenDate = '" & Format$(slGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(slGenTime, False))))) & "')"
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "", "frmWebVendorRpt" & "-cmdReport"
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
    gHandleError "AffErrorLog.txt", "frmWebVendorRpt-Click"
End Sub

Private Sub cmdDone_Click()
    Unload frmWebVendorRpt

End Sub


Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmWebVendorRpt
    
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
    gSelectiveStationsFromImport lbcStation, chkAllStations, Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim sNowDate As String
    ReDim tmVendorList(0 To 0) As VendorInfo

    frmWebVendorRpt.Caption = "Web Vendor Report - " & sgClientName

    imChkAllStationsIgnore = False
    chkAllStations.Value = vbUnchecked
    lbcStation.Clear
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).sUsedForATT = "Y" Then
            If tgStationInfo(iLoop).iType = 0 Then
                lbcStation.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStation.ItemData(lbcStation.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        End If
    Next iLoop
    chkAllStations.Value = vbUnchecked

    imChkAllVehiclesIgnore = False
    chkAllVehicles.Value = vbUnchecked
    lbcVehicle.Clear
    For iLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicle.AddItem Trim$(tgVehicleInfo(iLoop).sVehicleName)
        lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    chkAllVehicles.Value = vbUnchecked
    
    imCkcAllVendorsIgnore = False
    tmVendorList = gGetActiveDeliveryVendors()
    For iLoop = 0 To UBound(tmVendorList) - 1 Step 1
        lbcVendors.AddItem Trim$(tmVendorList(iLoop).sName)
        lbcVendors.ItemData(lbcVendors.NewIndex) = tmVendorList(iLoop).iIdCode
    Next iLoop
    ckcAllVendors.Value = vbUnchecked

End Sub
Sub mInit()
    
    Me.Width = Screen.Width / 1.3
    Me.Height = Screen.Height / 1.3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    gSetFonts frmWebVendorRpt
    gCenterForm frmWebVendorRpt
    gPopExportTypes cboFileType
    cboFileType.Enabled = True
    cmdReport.Enabled = False   'enable only after ALL filters are set  8/2/2018    FYM
    
End Sub
Private Sub Form_Initialize()
    mInit

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmWebVendorRpt = Nothing

End Sub

'****************************************************************
'*                                                              *
'*      Procedure Name:mReadFile                                *
'*      <input>   slFromFile - full path and file name          *
'                 of web Vendor filename (VendorFile-Guide.Txt  *
'                 llFromDate - user requested start date        *
'                 llToDAte - user requested end date            *
'*                                                              *
'*                                                              *
'*                                                              *
'****************************************************************
Private Function mReadFile(slFromFile As String, llFromDate As Long, llToDate As Long, slGenDate As String, slGenTime As String) As Integer
    
    Dim ilRet As Integer
    Dim slLine As String
    Dim slStr As String
    Dim ilEof As Integer
    Dim llDate As Long
    Dim slTempMsg As String
    Dim llline As Long
    Dim slLogDate As String
    Dim iICounter As Integer
    Dim blFound As Boolean          'used for checking included stations    Date: 8/2/2018  FYM
    Dim iIVehicleCounter As Integer  'used of Vendor list counter            Date: 8/2/2018  FYM
    
    ilRet = 0
    On Error GoTo mReadFileErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        mReadFile = False
        Exit Function
    End If
        
    
    Do
        ilRet = 0
        'On Error GoTo mReadFileErr:
        If EOF(hmFrom) Then
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        
llline = llline + 1         'debugging

        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                tmVendor_Info.sGenDate = slGenDate
                tmVendor_Info.sGenTime = slGenTime
                ilRet = gParseItem(slLine, ATTCODEINDEX, ",", slStr)            'agreement code
                slStr = mStripDoubleQuote(slStr)
                tmVendor_Info.lAttCode = Val(slStr)
                ilRet = gParseItem(slLine, VENDORINDEX, ",", slStr)             'Vendor code
                slStr = mStripDoubleQuote(slStr)
                tmVendor_Info.iVendorCode = Val(slStr)
                ilRet = gParseItem(slLine, LOGDATE, ",", slLogDate)             'log date
                slLogDate = mStripDoubleQuote(slLogDate)
                tmVendor_Info.sLogDate = Trim$(slLogDate)
                ilRet = gParseItem(slLine, SPOTCOUNT, ",", slStr)               'spot count
                slStr = mStripDoubleQuote(slStr)
                tmVendor_Info.lSpotCount = Val(slStr)
                ilRet = gParseItem(slLine, PROCESSDATETIME, ",", slStr)
                slStr = mStripDoubleQuote(slStr)
                tmVendor_Info.sProcessDateTime = Trim$(slStr)
               
                ' create the AFR records based on selected Vendors, Stations, Vehicles
                ' Date: 8/2/2018    FYM
                SQLQuery = "Select * from att where attCode = " & tmVendor_Info.lAttCode        'get the agreements
                Set rstATT = gSQLSelectCall(SQLQuery)
                While Not rstATT.EOF
                    For iICounter = 0 To UBound(tmSelectedVendors) - 1
                        'check for included Vendors
                        If (tmVendor_Info.iVendorCode = tmSelectedVendors(iICounter).iIdCode) Then
                            'check for included Vehicles
                            iIVehicleCounter = 0
                            Do While iIVehicleCounter < UBound(tmVehicleList)
                                If (rstATT!attvefCode = tmVehicleList(iIVehicleCounter).iVehicleCode) Then
                                    'check for included Stations
                                    blFound = gTestIncludeExclude(rstATT!attshfcode, imIncludeCodes, imUseCodes())
                                    If blFound Then
                                        mWriteVendorReport
                                        Exit Do
                                    End If
                                End If
                                iIVehicleCounter = iIVehicleCounter + 1
                            Loop
                        End If
                    Next
                    rstATT.MoveNext
                Wend
            End If                  'asc(slline) = 26            'eof
        End If                      '(len(slline) > 0
    Loop Until ilEof
    Close hmFrom
    If ilRet <> 0 Then
        mReadFile = False
    Else
        mReadFile = True
    End If
    Set rstATT = Nothing
    MousePointer = vbDefault
    Exit Function
mReadFileErr:
    ilRet = Err.Number
    Resume Next
End Function

'       mWRiteImportRecord
'       Create the prepass record from what was parsed from the Web Vendor text file
'       to send to crystal
'
Private Sub mWriteVendorReport()
Dim SQLQuery As String


       
        SQLQuery = "INSERT INTO afr (afrGenDate, afrGenTime, "      'gen date & time
        SQLQuery = SQLQuery & "afrAttCode, "                        'agreement code
        SQLQuery = SQLQuery & "afrSeqNo, "                          'vendor code
        SQLQuery = SQLQuery & "afrPledgeDate, "                     'log date
        SQLQuery = SQLQuery & "afrID, "                             'spot count
        SQLQuery = SQLQuery & "afrCreative) "                       'Process date & time
       
        SQLQuery = SQLQuery & " Values ( '" & Format$(tmVendor_Info.sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(tmVendor_Info.sGenTime, False))))) & "',"
        SQLQuery = SQLQuery & tmVendor_Info.lAttCode & ", "
        SQLQuery = SQLQuery & tmVendor_Info.iVendorCode & ","
        SQLQuery = SQLQuery & "'" & Format$(tmVendor_Info.sLogDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & tmVendor_Info.lSpotCount & ","
        SQLQuery = SQLQuery & "'" & tmVendor_Info.sProcessDateTime & "'"
        SQLQuery = SQLQuery & ")"

        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "", "frmWebVendorRpt" & "-mWriteVendorReport"
            Exit Sub
        End If
        cnn.CommitTrans
        Exit Sub
        
ErrHand:
    Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "frmWebVendorRpt-mWriteVendorReport"
    Exit Sub
End Sub

Public Function mStripDoubleQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    Dim slQuote As String * 1
    slQuote = """"
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = slQuote Then
                sOutStr = sOutStr & " "
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    mStripDoubleQuote = sOutStr
    Exit Function
End Function







Private Sub lbcStation_Click()
  If imChkAllStationsIgnore Then
        Exit Sub
    End If
    If chkAllStations.Value = vbChecked Then
        imChkAllStationsIgnore = True
        'chkListBox.Value = False
        chkAllStations.Value = vbUnchecked
        imChkAllStationsIgnore = False
    End If
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub lbcVehicle_Click()
    If imChkAllVehiclesIgnore Then
        Exit Sub
    End If
    If chkAllVehicles.Value = vbChecked Then
        imChkAllVehiclesIgnore = True
        'chkListBox.Value = False
        chkAllVehicles.Value = vbUnchecked
        imChkAllVehiclesIgnore = False
    End If
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub lbcVendors_Click()
  If imCkcAllVendorsIgnore Then
        Exit Sub
    End If
    If ckcAllVendors.Value = vbChecked Then
        imCkcAllVendorsIgnore = True
        'chkListBox.Value = False
        ckcAllVendors.Value = vbUnchecked
        imCkcAllVendorsIgnore = False
    End If
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/3/2018  FYM
    mEnableGenerateReportButton
End Sub

