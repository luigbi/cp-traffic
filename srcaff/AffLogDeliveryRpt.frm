VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmLogDeliveryRpt 
   Caption         =   "Affiliate Delivery Summary"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "AffLogDeliveryRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9360
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
      Height          =   4600
      Left            =   255
      TabIndex        =   8
      Top             =   1710
      Width           =   8895
      Begin VB.Frame Frame3 
         Caption         =   "Delivery by"
         Height          =   570
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2640
         Begin VB.OptionButton rbcDeliveryBy 
            Caption         =   "Log"
            Height          =   225
            Index           =   1
            Left            =   1320
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton rbcDeliveryBy 
            Caption         =   "Audio"
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   10
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.ListBox lbcStation 
         Height          =   3570
         ItemData        =   "AffLogDeliveryRpt.frx":08CA
         Left            =   4440
         List            =   "AffLogDeliveryRpt.frx":08D1
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   495
         Width           =   4125
      End
      Begin VB.CheckBox chkAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   165
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Frame frcLogBy 
         Caption         =   "Log Delivery by"
         Height          =   570
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   2640
         Begin VB.OptionButton rbcLogBy 
            Caption         =   "Vehicle"
            Height          =   225
            Index           =   1
            Left            =   1320
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton rbcLogBy 
            Caption         =   "Station"
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   15
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.ComboBox cbcSort 
         Height          =   315
         ItemData        =   "AffLogDeliveryRpt.frx":08D8
         Left            =   1320
         List            =   "AffLogDeliveryRpt.frx":08DA
         TabIndex        =   13
         Top             =   960
         Width           =   2340
      End
      Begin VB.ListBox lbcVehicle 
         Height          =   3570
         ItemData        =   "AffLogDeliveryRpt.frx":08DC
         Left            =   4440
         List            =   "AffLogDeliveryRpt.frx":08DE
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   495
         Width           =   4125
      End
      Begin VB.CheckBox chkAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Top             =   165
         Width           =   1935
      End
      Begin VB.Label lacSort 
         Caption         =   "Major Sort by"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   990
         Width           =   1065
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
      Left            =   255
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AffLogDeliveryRpt.frx":08E0
         Left            =   1335
         List            =   "AffLogDeliveryRpt.frx":08E2
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
Attribute VB_Name = "frmLogDeliveryRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*  frmLogDeliveryRpt - Shows what log exports are defined for each Vehicle
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


Private imChkAllVehiclesIgnore As Integer
Private imChkAllStationsIgnore As Integer
Private smUsingUnivision As String * 1
Private rstATT As ADODB.Recordset

Private Type VEHICLELIST
    iVefCode As Integer
    lShfCode As Long
    sType As String * 1
    'counts of agreements that are active as of report generation.  Future agreements are ignored
    iCBS As Integer
    iCC As Integer
    iCumulus As Integer
    iWeb As Integer
    iMktrn As Integer
    iUniv As Integer
    iNone As Integer
    iWideOrbit As Integer   '10-27-15
    iJelli As Integer       '10-27-15
    iIHeart As Integer       '10-27-15
    iAquira As Integer      '10-27-15  not implemented
    iRCS As Integer         '8-9-17
    iRadioTraffic As Integer    '8-9-17
    iMrMaster As Integer    '8-9-17
    iSynchronocity As Integer   '8-9-17
    iBSI As Integer             '8-9-17
    iWegeneriPump As Integer    '8-10-17
    iWegenerCompel As Integer   '8-10-17
    iXDSBreak As Integer       '8-10-17
    iXDSISCI As Integer         '8-10-17
    iIDC As Integer             '8-11-17
    iRadioWorkflow As Integer
End Type

Private tmVehicleList() As VEHICLELIST

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

Private Sub cmdDone_Click()
    Unload frmLogDeliveryRpt
End Sub

'       Generate a report of vehicles (Conventional, airing, selling (if agreement exists, changed from conv to selling), Log, Game
'       Keep count of how many agreements there are for each vehicle, including only those agreements that are active as of the
'       report generation.  Ignore all future agreements so as not to duplicate counts.
'       Allow sorting of any report column, with subsort always alphabetical by vehicle name
Private Sub cmdReport_Click()
        Dim ilTemp As Integer
        Dim iRet As Integer
        Dim sOutput As String
        Dim ilRet As Integer
        Dim dFWeek As Date
        Dim ilExportType As Integer
        Dim ilRptDest As Integer
        Dim slRptName As String
        Dim slExportName As String
        Dim slEnteredRange As String
        Dim sGenDate As String      'generation date for filtering prepass records
        Dim sGenTime As String      'generation time for filtering prepass records
        Dim ilUpper As Integer
        Dim ilVefCode As Integer
        Dim SQLQuery As String
        Dim slDropDate As String
        Dim slOffAir As String
        Dim slAttEndDate As String
        Dim llTodayDate As Long
        Dim llEndDate As Long
        Dim ilGetStation As Integer
        Dim llPrevShfCode As Long
        Dim llVefCode As Long
        Dim llCount As Long
        Dim ilSort As Integer
        Dim slSort As String * 13
        Dim ilColumns(0 To 14) As Integer    'each integer is equivalent to the items in the combo list box
        Dim ilTempColumns(0 To 14) As Integer    'copy of ilColumns for the sort indicator, but modified for each vehicle for update purposes if there are no agreements.  they need to
        Dim ilLoop As Integer
        Dim ilAtLeast1Att As Integer
        Dim ilSeparate As Integer
        Dim blAtLeast1VAT4Att As Boolean
        Dim blAtLeast1None As Boolean
        
        On Error GoTo ErrHand
    
        Screen.MousePointer = vbHourglass
        'CRpt1.Connect = "DSN = " & sgDatabaseName
        sGenDate = Format$(gNow(), "m/d/yyyy")
        sGenTime = Format$(gNow(), sgShowTimeWSecForm)
        llTodayDate = gDateValue(sGenDate)
        
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
        
        slSort = "VTBLJMDSRCUWON"        'this list matches the list below in order, changed to move status down in the list
        'Delivery by Log Sort Codes
        '1="V" = Vehicle
        '2="T" = Vehicle Type
        '3="B" = CBS
        '4="L" = iHeart
        '5="J" = Jelli
        '6="M" = Marketron
        '7="D" = Radio Traffic
        '8="S" = Radio Workflow
        '9="R" = RCS
        '10="C" = Stratus/Cumilus
        '11="U" = Univision
        '12="W" = Web
        '13="O" = Wide Orbit
        '14="N" = None / Manual
        If rbcDeliveryBy(0).Value Then          '8-11-17 audio or log delivery option
            'Delivery by Audio Sort Codes
            slSort = "VTBIMSXN"
            '1="V" = Vehicle
            '2="T" = Vehicle Type
            '3="B" = BSI
            '4="I" = IDC
            '5="M" = Mr. Master
            '6="S" = Synchonicity
            '7="X" = XDS-Break
            '8="N" = None
        End If
        If rbcLogBy(0).Value = True Then        'delivery by station (disabled.  if enabled, all sorts need to be adjusted in afLogDeliverySt.rpt, along with prepass code)
            'Delivery by Station Sort Codes
            slSort = "SBLCMW"
            '4/23/21 appears to be a not implemented feature, rbcLogBy / frcLogBy is not visible.
        End If
        
        'determine the sort criteria, which column
        'For vehicle option: 0 = vehicle name, 1 = vehicle type, 2= CBS log, 3 = clear channel , now iheart (8-10-17), 4 = cumulus, 5 = jelli, 6 = marketron, 7= web, 8 =wide orbit, 9 = manual (none)
        'For station option: 0 = station,  1= CBS log, 2 = clear channel, 3 = cumulus, 4 = marketron, 5 = uni, 6= web
        ilSort = cbcSort.ListIndex
        sgCrystlFormula1 = Mid$(slSort, ilSort + 1, 1)
        'the columnn selected for major sort will have a 1 in it, to sort to the top; all others will be a higher # just to sort after the selected sort column
        For ilTemp = 1 To 14
            ilColumns(ilTemp) = 2
        Next ilTemp
        'Each integer in array represents the sort column selected based on the combo box
        ilColumns(ilSort + 1) = 1             'used to sort the selected column to the top
        
        'setup the sort fields for crystal. cannot use the actual fields that the counts are in because we do not want to sort by count
        gUserActivityLog "S", sgReportListName & ": Prepass"
        cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = False
        cmdReturn.Enabled = False
                      
        ilUpper = 0
        ReDim tmVehicleList(0 To 0) As VEHICLELIST
        For ilTemp = 0 To lbcVehicle.ListCount - 1 Step 1
            If lbcVehicle.Selected(ilTemp) Then
                ilAtLeast1Att = False
                llVefCode = lbcVehicle.ItemData(ilTemp)
                ilVefCode = gBinarySearchVef(llVefCode)
                If ilVefCode <> -1 Then
                    'include active vehicles that are conventional, airing, selling (only if agreement defined since it was changed from conventional to selling), Game or Log
                    If (tgVehicleInfo(ilVefCode).sState <> "D") And (tgVehicleInfo(ilVefCode).sVehType = "C" Or tgVehicleInfo(ilVefCode).sVehType = "A" Or tgVehicleInfo(ilVefCode).sVehType = "S" Or tgVehicleInfo(ilVefCode).sVehType = "G" Or tgVehicleInfo(ilVefCode).sVehType = "L") Then
                        tmVehicleList(ilUpper).sType = tgVehicleInfo(ilVefCode).sVehType
                        tmVehicleList(ilUpper).iVefCode = llVefCode
                        'obtain the agreements
                        SQLQuery = "Select * from att where attvefcode = " & Str$(tmVehicleList(ilUpper).iVefCode) & " and attServiceAgreement <> 'Y' order by attshfcode, attAgreeStart"
                        Set rstATT = gSQLSelectCall(SQLQuery)
                        ilGetStation = True
                        While Not rstATT.EOF
                
                            slOffAir = Format$(rstATT!attOffAir, "mm/dd/yyyy")
                            slDropDate = Format$(rstATT!attDropDate, "mm/dd/yyyy")
                            'determine the earliest of 2 dates:  either drop date or off air date
                            If DateValue(gAdjYear(slDropDate)) < DateValue(gAdjYear(slOffAir)) Then
                                slAttEndDate = slDropDate
                            Else
                                slAttEndDate = slOffAir
                            End If
                            llEndDate = gDateValue(slAttEndDate)
                            If llEndDate >= llTodayDate Then            'ignore anything not active as of todays date
                                Do While ilGetStation
                                    If llPrevShfCode <> rstATT!attshfcode Then
                                        ilAtLeast1Att = True
                                        If rbcDeliveryBy(1).Value Then            'log delivery
                                            'do agreement count for the specific log export type (CSB, Cumulus, Marketron, Univision, CC, Web)
                                            If rstATT!attExportType <> 0 Then
                                                '5-10-16 resurrect Univision export, not using the Vendor method
                                                If rstATT!attExportToUnivision = "Y" And smUsingUnivision = "Y" Then        'using marketron has been converted to Using Univision
                                                    tmVehicleList(ilUpper).iUniv = tmVehicleList(ilUpper).iUniv + 1
                                                    tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                End If
                                                '7701
                                                blAtLeast1None = False
                                                blAtLeast1VAT4Att = False
                                                SQLQuery = "Select vatWvtVendorId as ID from VAT_Vendor_Agreement where vatattcode = " & rstATT!attCode
                                                Set rst = gSQLSelectCall(SQLQuery)
                                                Do While Not rst.EOF
                                                    If rst!ID = Vendors.cBs Then
                                                        tmVehicleList(ilUpper).iCBS = tmVehicleList(ilUpper).iCBS + 1
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                        blAtLeast1VAT4Att = True
                                                    ElseIf rst!ID = Vendors.iHeart Then
                                                        tmVehicleList(ilUpper).iCC = tmVehicleList(ilUpper).iCC + 1     '8-9-17 iCC (clear channel now iHeart)
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                        blAtLeast1VAT4Att = True
                                                    ElseIf rst!ID = Vendors.stratus Then            'cumulus
                                                        tmVehicleList(ilUpper).iCumulus = tmVehicleList(ilUpper).iCumulus + 1
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                         blAtLeast1VAT4Att = True
                                                   ElseIf rst!ID = Vendors.NetworkConnect Then
                                                        tmVehicleList(ilUpper).iMktrn = tmVehicleList(ilUpper).iMktrn + 1
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                        blAtLeast1VAT4Att = True
                                                    ElseIf rst!ID = Vendors.Jelli Then
                                                        tmVehicleList(ilUpper).iJelli = tmVehicleList(ilUpper).iJelli + 1
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                        blAtLeast1VAT4Att = True
                                                    ElseIf rst!ID = Vendors.WideOrbit Then
                                                        tmVehicleList(ilUpper).iWideOrbit = tmVehicleList(ilUpper).iWideOrbit + 1
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                    '8-9-17 following log delivery vendors added
                                                        blAtLeast1VAT4Att = True
                                                    ElseIf rst!ID = Vendors.RCS Then
                                                        tmVehicleList(ilUpper).iRCS = tmVehicleList(ilUpper).iRCS + 1
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                        blAtLeast1VAT4Att = True
                                                    ElseIf rst!ID = Vendors.RadioTraffic Then
                                                        tmVehicleList(ilUpper).iRadioTraffic = tmVehicleList(ilUpper).iRadioTraffic + 1
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                        blAtLeast1VAT4Att = True
                                                    ElseIf rst!ID = Vendors.RadioWorkflow Then  'TTP 10116
                                                        tmVehicleList(ilUpper).iRadioWorkflow = tmVehicleList(ilUpper).iRadioWorkflow + 1
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                        blAtLeast1VAT4Att = True
                                                    Else
                                                        blAtLeast1None = True
                                                    End If
                                                    rst.MoveNext
                                                Loop
                                                'Dan for 7701, changed else.  Should always be going to web if cumulus,cbs,etc. so don't need to test them
                                                If rstATT!attExportToWeb = "Y" Then
                                                    If Not blAtLeast1VAT4Att Then
                                                        tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
                                                    End If
                                                Else
                                                    tmVehicleList(ilUpper).iNone = tmVehicleList(ilUpper).iNone + 1
                                                End If
    '                                            If rstATT!attExportToCBS = "Y" Then
    '                                                tmVehicleList(ilUpper).iCBS = tmVehicleList(ilUpper).iCBS + 1
    '                                            End If
    '                                            If rstATT!attExportToClearCh = "Y" Then
    '                                                tmVehicleList(ilUpper).iCC = tmVehicleList(ilUpper).iCC + 1
    '                                            End If
    '                                            If rstATT!attWebInterface = "C" Then            'cumulus
    '                                                tmVehicleList(ilUpper).iCumulus = tmVehicleList(ilUpper).iCumulus + 1
    '                                            End If
    '                                            If rstATT!attExportToMarketron = "Y" Then
    '                                                tmVehicleList(ilUpper).iMktrn = tmVehicleList(ilUpper).iMktrn + 1
    '                                            End If
    '                                            If rstATT!attExportToUnivision = "Y" Then
    '                                                tmVehicleList(ilUpper).iUniv = tmVehicleList(ilUpper).iUniv + 1
    '                                            End If
    '                                            If rstATT!attExportToWeb = "Y" Then
    '                                                tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1
    '                                            End If
    '                                            If (rstATT!attExportToCBS <> "Y") And (rstATT!attExportToClearCh <> "Y") And (rstATT!attWebInterface <> "C") And (rstATT!attExportToMarketron <> "Y") And (rstATT!attExportToUnivision <> "Y") And (rstATT!attExportToWeb <> "Y") Then          'agreement doesnt have any log exports defined
    '                                                tmVehicleList(ilUpper).iNone = tmVehicleList(ilUpper).iNone + 1
    '                                            End If
                                                iRet = iRet
                                            Else                    'none defined
                                                tmVehicleList(ilUpper).iNone = tmVehicleList(ilUpper).iNone + 1
                                            End If
                                        Else                'audio delivery
                                            'If Trim$(rstATT!attAudioDelivery) <> "" Then

                                            blAtLeast1VAT4Att = False
                                            blAtLeast1None = False
                                            SQLQuery = "Select vatWvtVendorId as ID from VAT_Vendor_Agreement where vatattcode = " & rstATT!attCode
                                            Set rst = cnn.Execute(SQLQuery)
                                            Do While Not rst.EOF
                                                '8-24-17 ignore if any one of the log vendors
                                                If (rst!ID = Vendors.cBs) Or (rst!ID = Vendors.iHeart) Or (rst!ID = Vendors.stratus) Or (rst!ID = Vendors.NetworkConnect) Or (rst!ID = Vendors.Jelli) Or (rst!ID = Vendors.WideOrbit) Or (rst!ID = Vendors.RCS) Or (rst!ID = Vendors.RadioTraffic) Then
                                                    'ignore, do nothing
                                                    iRet = iRet
                                                Else
                                                    blAtLeast1VAT4Att = True        'at least 1 VAT exists for this agreement
                                                    If rst!ID = Vendors.BSI Then
                                                        tmVehicleList(ilUpper).iBSI = tmVehicleList(ilUpper).iBSI + 1
                                                    ElseIf rst!ID = Vendors.iDc Then
                                                        tmVehicleList(ilUpper).iIDC = tmVehicleList(ilUpper).iIDC + 1
                                                    ElseIf rst!ID = Vendors.MrMaster Then
                                                        tmVehicleList(ilUpper).iMrMaster = tmVehicleList(ilUpper).iMrMaster + 1
                                                    ElseIf rst!ID = Vendors.Synchronicity Then
                                                        tmVehicleList(ilUpper).iSynchronocity = tmVehicleList(ilUpper).iSynchronocity + 1
                                                    ElseIf rst!ID = Vendors.XDS_Break Then
                                                        tmVehicleList(ilUpper).iXDSBreak = tmVehicleList(ilUpper).iXDSBreak + 1
                                                    Else
                                                        tmVehicleList(ilUpper).iNone = tmVehicleList(ilUpper).iNone + 1
                                                        blAtLeast1None = True
                                                    End If
                                                End If
                                                rst.MoveNext
                                            Loop
                                            If Not blAtLeast1VAT4Att Then           'no VATs exist
                                                tmVehicleList(ilUpper).iNone = tmVehicleList(ilUpper).iNone + 1
                                            Else                                    'at least 1 VAT exist,
                                                If Not blAtLeast1None Then
                                                    tmVehicleList(ilUpper).iWeb = tmVehicleList(ilUpper).iWeb + 1           'keep track of at least 1 agreement for the vehicle/station; agreements could have up to 2 audio deliveries
                                                End If
                                            End If
                                            iRet = iRet
'                                            Else                    'none defined
'                                                tmVehicleList(ilUpper).iNone = tmVehicleList(ilUpper).iNone + 1
'                                            End If
                                        End If
                                    Else                'current station code  =  previous, only process the first active one
                                        iRet = iRet
                                    End If
                                    llPrevShfCode = rstATT!attshfcode
                                    ilGetStation = False
                                Loop
                            Else                'dates in the past
                                iRet = iRet
                            End If
                            ilGetStation = True
                            rstATT.MoveNext
                        Wend
                        'If ilAtLeast1Att Then
                            ReDim Preserve tmVehicleList(0 To ilUpper + 1) As VEHICLELIST
                            ilUpper = ilUpper + 1
                        'End If
                    End If                              '(tgVehicleInfo(ilVefCode).sState <> "D").......
                End If                                  'ilVefCode <> -1

            End If
        Next ilTemp
        
        '----------------------------------------------------
        'table of agreements built for selected vehicles
        'Write prepass record for .rpt

        'Audio Delivery
        '--------------------
        'grfgenDate = gen date
        'grfGenTime = gen  time
        'grfBktType = vehicle type
        'grfVefCode = vehicle code
        'Counts
        'grfPer1Genl = Counts: bsi
        'grfPer2Genl = Counts: idc
        'grfPer3Genl = Counts: mr master
        'grfPer4Genl = Counts: synchonicity
        'grfPer5Genl = Counts: xds-break
        'grfPer7Genl = Counts: none
        'sort indicators
        'grfPer1 = vehicle type
        'grfPer2 = vehicle code
        'grfPer3 = bsi
        'grfPer4 = idc
        'grfPer5 = mr master
        'grfPer6 = synchronicity
        'grfPer7 = xds-brek, none
        'grfPer8 = none
        'grfCode2 = ilseparate
        'grfLong = llcount
        
        'Log Delivery
        '--------------------
        'grfgenDate = gen date
        'grfGenTime = gen time
        'grfBktType = Vehicle type (Conventional, selling if converted from conv to selling, airing, games/sports/events, Log)
        'grfvefcode = vehicle code
        'Counts
        'grfPer1Genl = Counts: web
        'grfPer2Genl = Counts: cumulus
        'grfPer3Genl = Counts: mkt
        'grfPer4Genl = Counts: cbs
        'grfPer5Genl = Counts: iHeart (was cchannel)
        'grfPer6Genl = Counts: univ
        'grfPer7Genl = Counts: none
        'grfPer1 = counts: jelli
        'grfPer2 = counts: wide orbit
        'grfPer3 = counts: RCS
        'grfPer4 = counts: RadioTraffic
        'grfPer5 = counts: Radio Workflow
        'Sort indicator fields for crystal
        'grfPer8Genl = vehicle type
        'grfPer9Genl = vehicle code
        'grfPer10Genl = CBS
        'grfPer11Genl = iHeart(was cChannel)
        'grfPer12Genl = cumulus
        'grfPer10 = Jelli
        'grfPer13Genl = Mkt
        'grfPer14Genl = Univ
        'grfPer15Genl = Web
        'grfPer11 = Wide Orbig
        'grfPer16Genl = None
        'grfPer12 = RCS
        'grfPer13 = RadioTraffic
        'grfPer14 = RadioWorkflow
        'grfCode2 = ilseparate
        'grfLong = llcount
                       
        For ilTemp = 0 To UBound(tmVehicleList) - 1
            'only retain selling vehicles that have agreements, they have been converted from conventional to selling
            For ilLoop = 1 To 14
                ilTempColumns(ilLoop) = ilColumns(ilLoop)
            Next ilLoop
            ilSeparate = 0
            llCount = 0
            'use web count as each vendor used also gets added to the web count
            llCount = tmVehicleList(ilTemp).iWeb + tmVehicleList(ilTemp).iNone  '+ tmVehicleList(ilTemp).iCumulus + tmVehicleList(ilTemp).iMktrn + tmVehicleList(ilTemp).iCBS + tmVehicleList(ilTemp).iCC + tmVehicleList(ilTemp).iJelli + tmVehicleList(ilTemp).iWideOrbit + tmVehicleList(ilTemp).iIHeart + tmVehicleList(ilTemp).iAquira + tmVehicleList(ilTemp).iUniv + tmVehicleList(ilTemp).iNone
            If (tmVehicleList(ilTemp).sType <> "S") Or (tmVehicleList(ilTemp).sType = "S" And llCount > 0) Then
                If llCount = 0 Then
                    ilTempColumns(ilSort + 1) = 3     'modify sort sequence to sort after the major selected, and at the very end to show no agreements exist
                    ilSeparate = 3
                Else
                    ilSeparate = ilSeparate
                End If
                
                If rbcDeliveryBy(1).Value Then  '8-14-17 - Delivery By Log
                     '8-10-17 grfper3 was iHeart, became rcs; Clear Channel became iHeart
                     SQLQuery = "INSERT INTO " & "GRF_Generic_Report ("
                     SQLQuery = SQLQuery & " grfgenDate, "      'gen date
                     SQLQuery = SQLQuery & " grfGenTime, "      'gen time
                     'Counts
                     SQLQuery = SQLQuery & " grfBktType, "      'vehicle type
                     SQLQuery = SQLQuery & " grfVefCode, "      'vehicle code
                     SQLQuery = SQLQuery & " grfPer1Genl, "     'Counts: web
                     SQLQuery = SQLQuery & " grfPer2Genl, "     'Counts: cumulus
                     SQLQuery = SQLQuery & " grfPer3Genl, "     'Counts: mkt
                     SQLQuery = SQLQuery & " grfPer4Genl, "     'Counts: cbs
                     SQLQuery = SQLQuery & " grfPer5Genl, "     'Counts: iHeart (was cchannel)
                     SQLQuery = SQLQuery & " grfPer6Genl, "     'Counts: univ
                     SQLQuery = SQLQuery & " grfPer7Genl, "     'Counts: none
                     SQLQuery = SQLQuery & " grfPer1, "         'counts: jelli
                     SQLQuery = SQLQuery & " grfPer2, "         'counts: wide orbit
                     SQLQuery = SQLQuery & " grfPer3, "         'counts: RCS
                     SQLQuery = SQLQuery & " grfPer4, "         'counts: RadioTraffic
                     SQLQuery = SQLQuery & " grfPer5, "         'counts: Radio Workflow TTP 10116
                     'Sort indicator fields for crystal
                     SQLQuery = SQLQuery & " grfPer8Genl, "     'vehicle type
                     SQLQuery = SQLQuery & " grfPer9Genl, "     'vehicle code
                     SQLQuery = SQLQuery & " grfPer10Genl, "    'CBS
                     SQLQuery = SQLQuery & " grfPer11Genl, "    'iHeart(was cChannel)
                     SQLQuery = SQLQuery & " grfPer12Genl, "    'cumulus
                     SQLQuery = SQLQuery & " grfPer10, "        'Jelli
                     SQLQuery = SQLQuery & " grfPer13Genl, "    'Mkt
                     SQLQuery = SQLQuery & " grfPer14Genl, "    'Univ
                     SQLQuery = SQLQuery & " grfPer15Genl, "    'Web
                     SQLQuery = SQLQuery & " grfPer11, "        'Wide Orbig
                     SQLQuery = SQLQuery & " grfPer16Genl, "    'None
                     SQLQuery = SQLQuery & " grfPer12, "        'RCS
                     SQLQuery = SQLQuery & " grfPer13, "        'RadioTraffic
                     SQLQuery = SQLQuery & " grfPer14, "        'RadioWorkflow TTP 10116
                     SQLQuery = SQLQuery & " grfCode2, "        'ilseparate
                     SQLQuery = SQLQuery & " grfLong) "         'llcount
                     
                     SQLQuery = SQLQuery & " VALUES ("
                     SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', "   'gen date
                     SQLQuery = SQLQuery & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "', " 'gen time
                     'Counts:
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).sType & "', "        'vehicle type
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iVefCode & "', "     'vehicle code
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iWeb & "', "         'Counts: web
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iCumulus & "', "     'Counts: cumulus
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iMktrn & "', "       'Counts: mkt
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iCBS & "', "         'Counts: cbs
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iCC & "', "          'Counts: iHeart (was cchannel)
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iUniv & "', "        'Counts: univ
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iNone & "', "        'Counts: none
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iJelli & "', "       'counts: jelli
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iWideOrbit & "', "   'counts: wide orbit
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iRCS & "',"          'counts: RCS
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iRadioTraffic & "', " 'counts: RadioTraffic
                     SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iRadioWorkflow & "', "      'counts: Radio Workflow
                     
                     'Sort indicators  values for crystal
                     SQLQuery = SQLQuery & "'" & ilTempColumns(1) & "', "   'V=vehicle
                     SQLQuery = SQLQuery & "'" & ilTempColumns(2) & "', "   'T=vehicle type
                     SQLQuery = SQLQuery & "'" & ilTempColumns(3) & "', "   'B=CBS
                     SQLQuery = SQLQuery & "'" & ilTempColumns(4) & "', "   'L=iHeart(was cChannel)
                     SQLQuery = SQLQuery & "'" & ilTempColumns(5) & "', "   'J=Jelli
                     SQLQuery = SQLQuery & "'" & ilTempColumns(6) & "', "   'M=Marketron
                     SQLQuery = SQLQuery & "'" & ilTempColumns(7) & "', "   'D=Radio Traffic
                     SQLQuery = SQLQuery & "'" & ilTempColumns(8) & "', "   'S=Radio Workflow
                     SQLQuery = SQLQuery & "'" & ilTempColumns(9) & "', "   'R=RCS
                     SQLQuery = SQLQuery & "'" & ilTempColumns(10) & "', "  'C=Stratus/Cumilus
                     SQLQuery = SQLQuery & "'" & ilTempColumns(11) & "', "  'U=Univision
                     SQLQuery = SQLQuery & "'" & ilTempColumns(12) & "', "  'W=Web
                     SQLQuery = SQLQuery & "'" & ilTempColumns(13) & "', "  'O=Wide Orbit
                     SQLQuery = SQLQuery & "'" & ilTempColumns(14) & "', "  'None/Manual
                     SQLQuery = SQLQuery & "'" & ilSeparate & "', "         'ilseparate
                     SQLQuery = SQLQuery & "'" & llCount & "'"              'llcount
                     SQLQuery = SQLQuery & ")"
                    
                Else 'Delivery By Audio
                    SQLQuery = "INSERT INTO " & "GRF_Generic_Report ("
                    'counts
                    SQLQuery = SQLQuery & " grfgenDate, "   'gen date
                    SQLQuery = SQLQuery & " grfGenTime, "   'gen  time
                    SQLQuery = SQLQuery & " grfBktType, "   'vehicle type
                    SQLQuery = SQLQuery & " grfVefCode, "   'vehicle code
                    SQLQuery = SQLQuery & " grfPer1Genl, "  'Counts: bsi
                    SQLQuery = SQLQuery & " grfPer2Genl, "  'Counts: idc
                    SQLQuery = SQLQuery & " grfPer3Genl, "  'Counts: mr master
                    SQLQuery = SQLQuery & " grfPer4Genl, "  'Counts: synchonicity
                    SQLQuery = SQLQuery & " grfPer5Genl, "  'Counts: xds-break
                    SQLQuery = SQLQuery & " grfPer7Genl, "  'Counts: none
                    'sort indicators
                    SQLQuery = SQLQuery & " grfPer1, "      'vehicle type
                    SQLQuery = SQLQuery & " grfPer2, "      'vehicle code
                    SQLQuery = SQLQuery & " grfPer3, "      'bsi
                    SQLQuery = SQLQuery & " grfPer4, "      'idc
                    SQLQuery = SQLQuery & " grfPer5, "      'mr master
                    SQLQuery = SQLQuery & " grfPer6, "      'synchronicity
                    SQLQuery = SQLQuery & " grfPer7, "      'xds-brek, none
                    SQLQuery = SQLQuery & " grfPer8, "      'none
                    SQLQuery = SQLQuery & " grfCode2, "     'ilseparate
                    SQLQuery = SQLQuery & " grfLong"        'llcount
                    SQLQuery = SQLQuery & " ) "
                
                    SQLQuery = SQLQuery & " VALUES ('"
                    SQLQuery = SQLQuery & Format$(sGenDate, sgSQLDateForm) & "', "          'gen date
                    SQLQuery = SQLQuery & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "', " 'gen time
                    'Counts
                    SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).sType & "', "         'vehicle type
                    SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iVefCode & "', "      'vehicle code
                    SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iBSI & "', "          'counts: BSI
                    SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iIDC & "', "          'counts: IDC
                    SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iMrMaster & "', "     'counts: Mr.Master
                    SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iSynchronocity & "', " 'counts: synchonicity
                    SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iXDSBreak & "', "     'counts: xds-break
                    SQLQuery = SQLQuery & "'" & tmVehicleList(ilTemp).iNone & "', "         'Counts: none
                    
                    'Sort indicators  values for crystal
                    SQLQuery = SQLQuery & "'" & ilTempColumns(1) & "', " 'ilTemp(1) = V Vehiclename
                    SQLQuery = SQLQuery & "'" & ilTempColumns(2) & "', " 'iltemp(2) = T Vehicle type
                    SQLQuery = SQLQuery & "'" & ilTempColumns(3) & "', " 'iltemp(3) = B BSI
                    SQLQuery = SQLQuery & "'" & ilTempColumns(4) & "', " 'ilTemp(4) = I IDC
                    SQLQuery = SQLQuery & "'" & ilTempColumns(5) & "', " 'ilTemp(5) = M Mr Master
                    SQLQuery = SQLQuery & "'" & ilTempColumns(6) & "', " 'ilTemp(6) = S Synchronicity
                    SQLQuery = SQLQuery & "'" & ilTempColumns(7) & "',"  'ilTemp(7) = X XDS Break
                    SQLQuery = SQLQuery & "'" & ilTempColumns(8) & "', " 'ilTemp(8) = N None
                    SQLQuery = SQLQuery & "'" & ilSeparate & "', "
                    SQLQuery = SQLQuery & "'" & llCount & "'"
                    SQLQuery = SQLQuery & ")"
                End If

                cnn.BeginTrans
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.Txt", "LogDeliveryRpt-cmdReport_Click"
                    cnn.RollbackTrans
                    Exit Sub
                End If
                cnn.CommitTrans
            End If
        Next ilTemp
                    
        If rbcDeliveryBy(1).Value Then              'log delivery
            slRptName = "AfLogDeliveryVh.rpt"
            slExportName = "LogDeliveryVh"
        Else
            slRptName = "AfAudioDeliveryVh.rpt"
            slExportName = "LogDeliveryVh"
        End If
        
        gUserActivityLog "E", sgReportListName & ": Prepass"
        
        'Prepare records to pass to Crystal
        SQLQuery = "SELECT * from GRF_Generic_Report "
        SQLQuery = SQLQuery & "INNER JOIN VEF_Vehicles on grfvefCode = vefCode "
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
            gHandleError "AffErrorLog.Txt", "LogDeliveryRpt-cmdReport_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
        
        cmdReport.Enabled = True               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = True
        cmdReturn.Enabled = True


        Screen.MousePointer = vbDefault
        Erase tmVehicleList
        Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "frmLogDeliveryRpt" & "-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmLogDeliveryRpt
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmLogDeliveryRpt
    gCenterForm frmLogDeliveryRpt
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slStr As String
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    frmLogDeliveryRpt.Caption = "Affiliate Delivery Summary Report - " & sgClientName
    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    smUsingUnivision = "N"
    If Not rst.EOF Then
        If rst!siteMarketron = "1" Then
            smUsingUnivision = "Y"
        End If
    End If
    rbcLogBy_Click 1
    rbcDeliveryBy_Click 1 'Populate cbcSort
    
    gPopExportTypes cboFileType     '3-15-04
    cboFileType.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmVehicleList
    rstATT.Close
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmLogDeliveryRpt = Nothing
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

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub

Private Sub rbcDeliveryBy_Click(Index As Integer)
    If Index = 0 Then                   'audio
        cbcSort.Clear
        cbcSort.AddItem "Vehicle Name"
        cbcSort.AddItem "VehicleType"
        cbcSort.AddItem "BSI"
        cbcSort.AddItem "IDC"
        cbcSort.AddItem "Mr. Master"
        cbcSort.AddItem "Synchonicity"
        cbcSort.AddItem "XDS-Break"
        cbcSort.AddItem "None"
        cbcSort.ListIndex = 0
    Else
        cbcSort.Clear
        cbcSort.AddItem "Vehicle Name"
        cbcSort.AddItem "Vehicle Type"
        cbcSort.AddItem "CBS Log"
'        cbcSort.AddItem "Stratus Log"       '2-7-19
        cbcSort.AddItem "iHeart"            '8-10-17 was CChannel
        cbcSort.AddItem "Jelli Log"         '10-27-15
        cbcSort.AddItem "Marketron Log"
        cbcSort.AddItem "Radio Traffic Log"         '8-10-17
        cbcSort.AddItem "Radio Workflow"         'TTP 10116
        cbcSort.AddItem "RCS Log"                   '8-10-17
        cbcSort.AddItem "Stratus Log"       '2-7-19
        cbcSort.AddItem "Univision Log"
        cbcSort.AddItem "Web Log"
        cbcSort.AddItem "Wide Orbit Log"            '19-27-15
        cbcSort.AddItem "Manual"
        cbcSort.ListIndex = 0

    End If
End Sub

Private Sub rbcLogBy_Click(Index As Integer)
Dim iLoop As Integer
    If Index = 0 Then           'station option disabled
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
        'vehicle/station controls for list box and check box are on top of each other.  when moving the controls into place, the resize control screws up positioning
        'lbcStation.Move Frame2.Left + 4440, Frame2.Top + 495
        'chkAllStations.Move Frame2.Top + 4440, Frame2.Top + 165
        chkAllStations.Visible = True
        chkAllVehicles.Visible = False
        lbcStation.Visible = True
        lbcVehicle.Visible = False
        cbcSort.Clear
        cbcSort.AddItem "Station"
        cbcSort.AddItem "CBS Log"
        cbcSort.AddItem "iHeart Log"            '8-10-17 was Clear Channel
        cbcSort.AddItem "Stratus Log"           '2-7-19
        cbcSort.AddItem "Jelli Log"                     '10-27-15
        cbcSort.AddItem "Marketron Log"
        cbcSort.AddItem "Univision Log"            '10-12-15 remove for now, 5-10-16 resurrect
        cbcSort.AddItem "Web Log"
        cbcSort.AddItem "Wide Orbit Log"                '10-27-15
        cbcSort.AddItem "RCS Log"                       '8-10-17
        cbcSort.AddItem "Radio Traffic Log"             '8-10-17
        cbcSort.AddItem "Stratus Log"           '2-7-19
        cbcSort.ListIndex = 0

    Else
        imChkAllVehiclesIgnore = False
        chkAllVehicles.Value = vbUnchecked
        lbcVehicle.Clear
        For iLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
            lbcVehicle.AddItem Trim$(tgVehicleInfo(iLoop).sVehicleName)
            lbcVehicle.ItemData(lbcVehicle.NewIndex) = tgVehicleInfo(iLoop).iCode
        Next iLoop
        chkAllVehicles.Value = vbUnchecked
        'vehicle/station controls for list box and check box are on top of each other.  when moving the controls into place, the resize control screws up positioning
        'lbcVehicle.Move Frame2.Left + 4440, Frame2.Top + 495    '4440, 495
        'chkAllVehicles.Move Frame2.Left + 4440, Frame2.Top + 165    '4440, 165
        chkAllVehicles.Visible = True
        chkAllStations.Visible = False
        lbcStation.Visible = False
        lbcVehicle.Visible = True
        cbcSort.Clear
        cbcSort.AddItem "Vehicle Name"
        cbcSort.AddItem "Vehicle Type"
        cbcSort.AddItem "CBS Log"
        cbcSort.AddItem "Stratus Log"       '2-7-19
        cbcSort.AddItem "iHeart"            '8-10-17 was CChannel
        cbcSort.AddItem "Jelli Log"         '10-27-15
        cbcSort.AddItem "Marketron Log"
        cbcSort.AddItem "Radio Traffic Log"         '8-10-17
        cbcSort.AddItem "Radio Workflow"         'TTP 10116
        cbcSort.AddItem "RCS Log"                   '8-10-17
        cbcSort.AddItem "Stratus Log"       '2-7-19
        cbcSort.AddItem "Univision Log"
        cbcSort.AddItem "Web Log"
        cbcSort.AddItem "Wide Orbit Log"            '19-27-15
        cbcSort.AddItem "Manual"
        cbcSort.ListIndex = 0

    End If
End Sub
