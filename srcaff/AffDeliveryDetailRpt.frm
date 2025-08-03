VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDeliveryDetailRpt 
   Caption         =   "Log Type Report"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "AffDeliveryDetailRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   375
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
      Height          =   4600
      Left            =   255
      TabIndex        =   13
      Top             =   1710
      Width           =   8895
      Begin VB.CommandButton cmdStationListFile 
         Height          =   240
         Left            =   8280
         Picture         =   "AffDeliveryDetailRpt.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Select Stations from File.."
         Top             =   165
         Width           =   360
      End
      Begin V81Affiliate.CSI_Calendar CalEndDate 
         Height          =   270
         Left            =   2520
         TabIndex        =   12
         Top             =   600
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "8/8/2023"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
         CSI_CurDayForeColor=   51200
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   -1  'True
         CSI_AllowTFN    =   -1  'True
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar CalStartDate 
         Height          =   270
         Left            =   600
         TabIndex        =   10
         Top             =   600
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "8/8/2023"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
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
         CSI_CurDayForeColor=   51200
         CSI_ForceMondaySelectionOnly=   -1  'True
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin VB.CheckBox ckcSkip1 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   1560
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.ComboBox cbcSort3 
         Height          =   315
         ItemData        =   "AffDeliveryDetailRpt.frx":0E34
         Left            =   1320
         List            =   "AffDeliveryDetailRpt.frx":0E36
         TabIndex        =   24
         Top             =   2280
         Width           =   1365
      End
      Begin VB.ComboBox cbcSort2 
         Height          =   315
         ItemData        =   "AffDeliveryDetailRpt.frx":0E38
         Left            =   1320
         List            =   "AffDeliveryDetailRpt.frx":0E3A
         TabIndex        =   22
         Top             =   1920
         Width           =   1365
      End
      Begin VB.ComboBox cbcSort1 
         Height          =   315
         ItemData        =   "AffDeliveryDetailRpt.frx":0E3C
         Left            =   1320
         List            =   "AffDeliveryDetailRpt.frx":0E3E
         TabIndex        =   19
         Top             =   1560
         Width           =   1365
      End
      Begin VB.ListBox lbcVendors 
         Height          =   1230
         ItemData        =   "AffDeliveryDetailRpt.frx":0E40
         Left            =   4440
         List            =   "AffDeliveryDetailRpt.frx":0E42
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2085
      End
      Begin VB.CheckBox ckcAllVendors 
         Caption         =   "All Vendors"
         Height          =   255
         Left            =   4440
         TabIndex        =   29
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ListBox lbcStation 
         Height          =   2400
         ItemData        =   "AffDeliveryDetailRpt.frx":0E44
         Left            =   6600
         List            =   "AffDeliveryDetailRpt.frx":0E4B
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Width           =   2085
      End
      Begin VB.CheckBox chkAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   6600
         TabIndex        =   27
         Top             =   165
         Width           =   1245
      End
      Begin VB.ListBox lbcVehicle 
         Height          =   2400
         ItemData        =   "AffDeliveryDetailRpt.frx":0E52
         Left            =   4440
         List            =   "AffDeliveryDetailRpt.frx":0E54
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   2085
      End
      Begin VB.CheckBox chkAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   165
         Width           =   1935
      End
      Begin VB.Label lacEnd 
         Caption         =   "End"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lacStart 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lacAttText 
         Caption         =   "Agreement Dates-"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label lacSkip 
         Alignment       =   1  'Right Justify
         Caption         =   "Skip"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   17
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label lacSkip 
         Alignment       =   1  'Right Justify
         Caption         =   "Page"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lacSort3 
         Caption         =   "Sort Field #3"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2340
         Width           =   1065
      End
      Begin VB.Label lacSort2 
         Caption         =   "Sort Field #2"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lacSort1 
         Caption         =   "Sort Field #1"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1620
         Width           =   945
      End
      Begin VB.Label lacSortSeq2 
         Caption         =   "Major to Minor:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label lacSortSeq 
         Caption         =   "Enter sort sequence-"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   5355
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   5115
      TabIndex        =   6
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   225
      Width           =   2685
   End
   Begin VB.Frame frcOutput 
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
      Left            =   255
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffDeliveryDetailRpt.frx":0E56
         Left            =   1335
         List            =   "AffDeliveryDetailRpt.frx":0E58
         TabIndex        =   4
         Top             =   765
         Width           =   2040
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   825
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frmDeliveryDetailRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*  frmDeliveryDetailRpt - Shows what log exports are defined for each Vehicle
'                                       (conventional, airing, selling (if an agreement is assoc;
'                                       it was orig. a conventional changed to selling), Game, Log
'
'*
'*  Created July,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'
'****************************************************************************
Option Explicit

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

Private Const SORT1_STATION = 0
Private Const SORT1_VEHICLE = 1
Private Const SORT1_VENDOR = 2
Private Const SORT2_NONE = 0
Private Const SORT2_STATION = 1
Private Const SORT2_VEHICLE = 2
Private Const SORT2_VENDOR = 3
Private Const SORT3_NONE = 0
Private Const SORT3_STATION = 1
Private Const SORT3_VEHICLE = 2
Private Const SORT3_VENDOR = 3


Private Sub cbcSort1_Click()
Dim blOk As Boolean
    blOk = mTestDuplicateSort1()
    If Not blOk Then
        MsgBox "Cannot have same sort defined for more than 1 sort field"
        imSort1 = cbcSort1.ListIndex
    Else
        imSort1 = cbcSort1.ListIndex
    End If
End Sub

Private Sub cbcSort2_Click()
Dim blOk As Boolean
    blOk = mTestDuplicateSort2()
    If Not blOk Then
        MsgBox "Cannot have same sort defined for more than 1 sort field"
        imSort2 = cbcSort2.ListIndex
    Else
        imSort2 = cbcSort2.ListIndex
    End If
End Sub

Private Sub cbcSort3_Click()
Dim blOk As Boolean
    blOk = mTestDuplicateSort3()
    If Not blOk Then
        MsgBox "Cannot have same sort defined for more than 1 sort field"
        imSort3 = cbcSort3.ListIndex
    Else
        imSort3 = cbcSort3.ListIndex
    End If

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
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload frmDeliveryDetailRpt
End Sub
'
'       Generate a report of vehicles (Conventional, airing, selling (if agreement exists, changed from conv to selling), Log, Game
'       Keep count of how many agreements there are for each vehicle, including only those agreements that are active as of the
'       report generation.  Ignore all future agreements so as not to duplicate counts.
'       Allow sorting of any report column, with subsort always alphabetical by vehicle name
Private Sub cmdReport_Click()
        Dim ilTemp As Integer
        Dim ilExportType As Integer
        Dim ilRptDest As Integer
        Dim slRptName As String
        Dim slExportName As String
        Dim sGenDate As String      'generation date for filtering prepass records
        Dim sGenTime As String      'generation time for filtering prepass records
        Dim ilVefCode As Integer
        Dim SQLQuery As String
        Dim llVefCode As Long
        Dim llCount As Long
        Dim ilLoop As Integer
        Dim sStartDate As String
        Dim sEndDate As String
        Dim sDateRange As String
        Dim ilVendorId As Integer
        Dim ilShttCode As Integer
        Dim llAttCode As Long
        Dim slService As String
        Dim slType As String * 1
        Dim blVendorFound As Boolean
        Dim blStationFound As Boolean
        
        On Error GoTo ErrHand
    
        Screen.MousePointer = vbHourglass
        
        sStartDate = Trim$(CalStartDate.Text)
        If sStartDate = "" Then
            sStartDate = "1/1/1970"
        End If
        sEndDate = Trim$(CalEndDate.Text)
        If sEndDate = "" Then
            sEndDate = "12/31/2069"
        End If
        If gIsDate(sStartDate) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalStartDate.SetFocus
            Exit Sub
        End If
        If gIsDate(sEndDate) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy)", vbCritical
            CalEndDate.SetFocus
            Exit Sub
        End If
        sStartDate = Format(sStartDate, "m/d/yyyy")
        sEndDate = Format(sEndDate, "m/d/yyyy")
'        sDateRange = " (attAgreeStart <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'" & " And attAgreeEnd >= " & "'" + Format$(sStartDate, sgSQLDateForm) & "') "
        sDateRange = " (attOffAir >=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'" & " And attDropDate >=" & "'" + Format$(sStartDate, sgSQLDateForm) & "'" & " And attOnAir <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "')"

        sGenDate = Format$(gNow(), "m/d/yyyy")
        sGenTime = Format$(gNow(), sgShowTimeWSecForm)
        
        If optRptDest(0).Value = True Then
            'CRpt1.Destination = crptToWindow
            ilRptDest = 0
        ElseIf optRptDest(1).Value = True Then
            'CRpt1.Destination = crptToPrinter
            ilRptDest = 1
        ElseIf optRptDest(2).Value = True Then
            ilRptDest = 2
            ilExportType = cboFileType.ListIndex    '3-15-04
        End If
        
        gUserActivityLog "S", sgReportListName & ": Prepass"
        cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = False
        cmdReturn.Enabled = False
        
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
        
        ReDim imUseCodes(0 To 0) As Integer
        gObtainCodes lbcStation, imIncludeCodes, imUseCodes()        'build array of which codes to incl/excl
                      
        For ilTemp = 0 To lbcVehicle.ListCount - 1 Step 1
            If lbcVehicle.Selected(ilTemp) Then
                llVefCode = lbcVehicle.ItemData(ilTemp)
                ilVefCode = gBinarySearchVef(llVefCode)
                If ilVefCode <> -1 Then
                    'include active vehicles that are conventional, airing, selling (only if agreement defined since it was changed from conventional to selling), Game or Log
                    If (tgVehicleInfo(ilVefCode).sState <> "D") And (tgVehicleInfo(ilVefCode).sVehType = "C" Or tgVehicleInfo(ilVefCode).sVehType = "A" Or tgVehicleInfo(ilVefCode).sVehType = "S" Or tgVehicleInfo(ilVefCode).sVehType = "G" Or tgVehicleInfo(ilVefCode).sVehType = "L") Then
                        'obtain the agreements
                        SQLQuery = "Select * from att where attvefcode = " & Str$(llVefCode) & " and attServiceAgreement <> 'Y' and " & sDateRange & " order by attshfcode, attAgreeStart"
                        Set rstATT = gSQLSelectCall(SQLQuery)
                        While Not rstATT.EOF
                            blStationFound = gTestIncludeExclude(rstATT!attshfcode, imIncludeCodes, imUseCodes())
                            'TTP 10717 - Affiliate Delivery Report - Does not generate data for Agreements set to Manual
                            'attExportType was historically used for Univision (2=Univision), [but is not used for univision any more]. (0=Manual, 1=Web).  The statement probably should have been If (rstATT!attExportType <> 2)
                            'If (rstATT!attExportType <> 0) And (blStationFound) Then
                            If (blStationFound) Then
                                '5-23-18 Univision is ignored due to its unique nature
                                '7701
                                SQLQuery = "Select vatWvtVendorId as ID from VAT_Vendor_Agreement where vatattcode = " & rstATT!attCode
                                Set rst = gSQLSelectCall(SQLQuery)
                                Do While Not rst.EOF
                                    ilVendorId = rst!ID
                                    ilShttCode = rstATT!attshfcode
                                    llAttCode = rstATT!attCode
                                    blVendorFound = False
                                    For ilLoop = 0 To UBound(tmSelectedVendors) - 1
                                        If ilVendorId = tmSelectedVendors(ilLoop).iIdCode Then
                                            slService = Trim$(tmSelectedVendors(ilLoop).sName)
                                            slType = tmSelectedVendors(ilLoop).sDeliveryType            ' A = audio, L = Log
                                            blVendorFound = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If blVendorFound Then   'if vendor not found, wasnt selected
                                        'table of services built for selected vehicles & stations
                                        'Write prepass record for .rpt
                                        'grfGenDate = generation date for filtering
                                        'grfGenTime = generation time for filtering in .rpt
                                        'grfBktType = A = audio, L = Log delivery
                                        'grfvefcode - vehicle code
                                        'grfCode = Station (shttcode).
                                        'Retrieve time zone, format & market from station in crystal reports
                                        SQLQuery = "INSERT INTO " & "GRF_Generic_Report"
                                        SQLQuery = SQLQuery & " (grfgenDate, grfGenTime, "           'gen date & time
                                        SQLQuery = SQLQuery & " grfBktType, grfVefCode,grfCode2, grfCode4, grfGenDesc) "                 'ServiceType, vehicle code, and service name
                                      
                                        SQLQuery = SQLQuery & " VALUES ('" & Format$(sGenDate, sgSQLDateForm) & "', " & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "', "
                                        SQLQuery = SQLQuery & "'" & slType & "', " & llVefCode & ", " & ilShttCode & ", " & llAttCode & ", " & "'" & slService & "' )"
    
                                        cnn.BeginTrans
                                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                            '6/11/16: Replaced GoSub
                                            'GoSub ErrHand:
                                            Screen.MousePointer = vbDefault
                                            gHandleError "AffErrorLog.Txt", "DeliveryDetailRpt-cmdReport_Click"
                                            cnn.RollbackTrans
                                            Exit Sub
                                        End If
                                        cnn.CommitTrans
                                    End If

                                    rst.MoveNext
                                Loop
                            End If
                            rstATT.MoveNext
                        Wend
                    End If                              '(tgVehicleInfo(ilVefCode).sState <> "D").......
                End If                                  'ilVefCode <> -1
            End If
        Next ilTemp
        
 
                    
'        If imSort1 = SORT1_VENDOR Then              'vendor major sort
'            slRptName = "AfDelVendorDet.rpt"
'            slExportName = "AfDelVendorDet"
'        Else                            'station or vehicle sort
        'changed to use one rpt instead of 2
            slRptName = "AfDelDetail.rpt"
            slExportName = "AfDelDetail"
'        End If
        
        sgCrystlFormula1 = Trim$(Str(imSort1))
        sgCrystlFormula2 = Trim$(Str(imSort2))
        sgCrystlFormula3 = Trim$(Str(imSort3))
        sgCrystlFormula4 = "'" & sStartDate & " - " & sEndDate & "'"
        If ckcSkip1.Value = vbChecked Then
            sgCrystlFormula5 = "'Y'"
        Else
            sgCrystlFormula5 = "'N'"
        End If
        
        gUserActivityLog "E", sgReportListName & ": Prepass"
        
        'Prepare records to pass to Crystal
        SQLQuery = "SELECT * from GRF_Generic_Report "
        SQLQuery = SQLQuery & "INNER JOIN VEF_Vehicles on grfvefCode = vefCode "
        SQLQuery = SQLQuery & "INNER JOIN shtt on grfCode2 = shttCode "
        SQLQuery = SQLQuery & "INNER JOIN att on grfcode4 = attcode "
        SQLQuery = SQLQuery & "INNER JOIN mkt on shttmktcode = mktcode "
        SQLQuery = SQLQuery & "LEFT OUTER JOIN fmt_Station_format on shttfmtcode = fmtcode "
        SQLQuery = SQLQuery + " where (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' AND grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"

        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, slRptName, slExportName

        
        'remove all the records just printed
        SQLQuery = "DELETE FROM grf_Generic_Report "
        SQLQuery = SQLQuery & " WHERE (grfGenDate = '" & Format$(sGenDate, sgSQLDateForm) & "' " & "and grfGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"
        cnn.BeginTrans
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "DeliveryDetailRpt-cmdReport_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        
        cmdReport.Enabled = True               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = True
        cmdReturn.Enabled = True

        Screen.MousePointer = vbDefault
        Erase tmSelectedVendors, imUseCodes
        Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "frmDeliveryDetailRpt" & "-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmDeliveryDetailRpt
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

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmDeliveryDetailRpt
    gCenterForm frmDeliveryDetailRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim sNowDate As String
    ReDim tmVendorList(0 To 0) As VendorInfo
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    frmDeliveryDetailRpt.Caption = "Affiliate Delivery Detail Report - " & sgClientName
'    SQLQuery = "SELECT * From Site Where siteCode = 1"
'    Set rst = gSQLSelectCall(SQLQuery)
'    smUsingUnivision = "N"
'    If Not rst.EOF Then
'        If rst!siteMarketron = "1" Then
'            smUsingUnivision = "Y"
'        End If
'    End If

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

    cbcSort1.AddItem "Station"
    cbcSort1.AddItem "Vehicle"
    cbcSort1.AddItem "Vendor"
    cbcSort1.ListIndex = 2          'default to vendor
    
    cbcSort2.AddItem "None"
    cbcSort2.AddItem "Station"
    cbcSort2.AddItem "Vehicle"
    cbcSort2.AddItem "Vendor"
    cbcSort2.ListIndex = 0
    
    cbcSort3.AddItem "None"
    cbcSort3.AddItem "Station"
    cbcSort3.AddItem "Vehicle"
    cbcSort3.AddItem "Vendor"
    cbcSort3.ListIndex = 0
    
    sNowDate = Format$(gNow(), "m/d/yy")

    sNowDate = Format(sNowDate, "m/d/yyyy")
    'backup to Monday since all CPTTS are by week
    Do While Weekday(sNowDate, vbSunday) <> vbMonday
        sNowDate = DateAdd("d", -1, sNowDate)
    Loop
    CalStartDate.Text = sNowDate
    CalEndDate.Text = DateAdd("d", 6, sNowDate)
    imSort1 = 2
    imSort2 = 0
    imSort3 = 0
    gPopExportTypes cboFileType     '3-15-04
    cboFileType.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmVendorList
    rstATT.Close
    rst.Close
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmDeliveryDetailRpt = Nothing
End Sub

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
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub

Private Function mTestDuplicateSort1() As Boolean
    mTestDuplicateSort1 = True
    If cbcSort1.ListIndex = SORT1_STATION Then
        If cbcSort2.ListIndex = SORT2_STATION Or cbcSort3.ListIndex = SORT3_STATION Then
            mTestDuplicateSort1 = False
            'MsgBox "Cannot have Station sort defined for more than 1 sort field"
            Exit Function
        End If
    ElseIf cbcSort1.ListIndex = SORT1_VEHICLE Then
        If cbcSort2.ListIndex = SORT2_VEHICLE Or cbcSort3.ListIndex = SORT3_VEHICLE Then
            mTestDuplicateSort1 = False
            'MsgBox "Cannot have Vehicle sort defined for more than 1 sort field"
            Exit Function
        End If
    Else
        If cbcSort1.ListIndex = SORT1_VENDOR Then
            If cbcSort2.ListIndex = SORT2_VENDOR Or cbcSort3.ListIndex = SORT3_VENDOR Then
                mTestDuplicateSort1 = False
                'MsgBox "Cannot have Vehicle sort defined for more than 1 sort field"
                Exit Function
            End If
        End If
    End If
End Function
Private Function mTestDuplicateSort2() As Boolean
    mTestDuplicateSort2 = True
    If cbcSort2.ListIndex = SORT2_STATION Then
        If cbcSort1.ListIndex = SORT1_STATION Or cbcSort3.ListIndex = SORT3_STATION Then
            mTestDuplicateSort2 = False
            'MsgBox "Cannot have Station sort defined for more than 1 sort field"
            Exit Function
        End If
    ElseIf cbcSort2.ListIndex = SORT2_VEHICLE Then
        If cbcSort1.ListIndex = SORT1_VEHICLE Or cbcSort3.ListIndex = SORT3_VEHICLE Then
            mTestDuplicateSort2 = False
            'MsgBox "Cannot have Vehicle sort defined for more than 1 sort field"
            Exit Function
        End If
    Else
        If cbcSort2.ListIndex = SORT2_VENDOR Then
            If cbcSort1.ListIndex = SORT1_VENDOR Or cbcSort3.ListIndex = SORT3_VENDOR Then
                mTestDuplicateSort2 = False
                'MsgBox "Cannot have Vehicle sort defined for more than 1 sort field"
                Exit Function
            End If
        End If
    End If
End Function
Private Function mTestDuplicateSort3() As Boolean
    mTestDuplicateSort3 = True
    If cbcSort3.ListIndex = SORT3_STATION Then
        If cbcSort1.ListIndex = SORT1_STATION Or cbcSort2.ListIndex = SORT2_STATION Then
            mTestDuplicateSort3 = False
            'MsgBox "Cannot have Station sort defined for more than 1 sort field"
            Exit Function
        End If
    ElseIf cbcSort3.ListIndex = SORT3_VEHICLE Then
        If cbcSort1.ListIndex = SORT1_VEHICLE Or cbcSort2.ListIndex = SORT2_VEHICLE Then
            mTestDuplicateSort3 = False
            'MsgBox "Cannot have Vehicle sort defined for more than 1 sort field"
            Exit Function
        End If
    Else
        If cbcSort3.ListIndex = SORT3_VENDOR Then
            If cbcSort1.ListIndex = SORT1_VENDOR Or cbcSort2.ListIndex = SORT2_VENDOR Then
                mTestDuplicateSort3 = False
                'MsgBox "Cannot have Vehicle sort defined for more than 1 sort field"
                Exit Function
            End If
        End If
    End If
End Function

