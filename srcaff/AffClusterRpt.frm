VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmClusterRpt 
   Caption         =   "Affiliate Cluster Report"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   9360
   Icon            =   "AffClusterRpt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
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
      FormDesignHeight=   6780
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
      Height          =   5055
      Left            =   240
      TabIndex        =   8
      Top             =   1650
      Width           =   8895
      Begin VB.ComboBox cbcSort 
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Width           =   1485
      End
      Begin VB.ListBox lbcVehAff 
         Height          =   4155
         ItemData        =   "AffClusterRpt.frx":08CA
         Left            =   4560
         List            =   "AffClusterRpt.frx":08CC
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   4020
      End
      Begin VB.CheckBox chkListBox 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   270
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "5/4/2010"
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
         CSI_DefaultDateType=   1
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   270
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         Text            =   "5/4/2010"
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
         CSI_DefaultDateType=   1
      End
      Begin VB.Label lacEnd 
         Caption         =   "End"
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lacStart 
         Caption         =   "Active Dates - Start"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lacSortBy 
         Caption         =   "Sort by-"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   750
         Width           =   855
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
         ItemData        =   "AffClusterRpt.frx":08CE
         Left            =   1335
         List            =   "AffClusterRpt.frx":08D0
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
Attribute VB_Name = "frmClusterRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*  frmClusterRpt - compares vehicles and their affiliations, sorted by either one
'*
'*  Created July,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'
'   8-11-04 Add selectivity to gather agreements starting between X & Y dates;
'           Add selectivity to gather agreements ending between X & Y dates
'****************************************************************************
Option Explicit

Private smToFile As String
Private imChkStationIgnore As Integer
Private imChkListBoxIgnore As Integer
Private imChkListOtherIgnore As Integer
Private imSortBy As Integer
Private bmSortListTest As Boolean
Private tmAmr As AMR

Private rst_Agreement As ADODB.Recordset

Private Sub mEnableGenerateReportButton()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/13/2018  FYM
    If CalOnAirDate.Text = "" Or CalOffAirDate.Text = "" Then
        cmdReport.Enabled = False
    Else
        If (lbcVehAff.SelCount > 0) Then
            cmdReport.Enabled = True
        Else
            cmdReport.Enabled = False
        End If
    End If
    
End Sub

Private Sub CalOffAirDate_CalendarChanged()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub

Private Sub CalOnAirDate_CalendarChanged()
    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
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
    If chkListBox.Value = vbChecked Then
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

    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton

End Sub

Private Sub cmdDone_Click()
    Unload frmClusterRpt
End Sub

Private Sub cmdReport_Click()
    Dim i, j, X, Y, iPos As Integer
    Dim ilLoop As Integer
    Dim iRet As Integer
    Dim sCode As String
    Dim sName As String
    Dim sVehicles As String
    Dim sStartDate As String
    Dim sEndDate As String
    Dim sDateRange As String
    Dim sStationType As String
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim slDescription As String
    Dim slEnteredRange As String
    Dim slMulticastOnly As String
    Dim ilInclVehicleCodes As Integer
    Dim ilUseVehicleCodes() As Integer
    Dim llCount As Long
    Dim llValue As Long
    Dim llOwnerInx As Long
    Dim ilShttInx As Integer
    Dim llTemp As Long
    Dim slOwnerName As String
    Dim ilMktInx As Integer
    Dim llVefInx As Long
    Dim ilFmtInx As Integer
    Dim slSortOption As String * 5
        
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
            ilRptDest = 0
        ElseIf optRptDest(1).Value = True Then
            ilRptDest = 1
        ElseIf optRptDest(2).Value = True Then
            ilRptDest = 2
            ilExportType = cboFileType.ListIndex    '3-15-04
        End If
        
        slSortOption = "FMOST"           'format, market, owner, station, time zone        cmdReport.Enabled = False               'disallow user from clicking these buttons until report completed
        cmdDone.Enabled = False
        cmdReturn.Enabled = False

        gUserActivityLog "S", sgReportListName & ": Prepass"
        
        sStationType = "shttType = 0"  'agreement type is station , not people
        
        'get the Generation date and time to filter data for Crystal
        sgGenDate = Format$(gNow(), "m/d/yyyy")             '7-10-13 use global gen date/time for crystal filtering
        sgGenTime = Format$(gNow(), sgShowTimeWSecForm)
    
        sStartDate = Format(sStartDate, "m/d/yyyy")
        sEndDate = Format(sEndDate, "m/d/yyyy")
       
        sDateRange = " attOffAir >=" & "'" & Format$(sStartDate, sgSQLDateForm) & "'" & " And attDropDate >=" & "'" + Format$(sStartDate, sgSQLDateForm) & "'" & " And attOnAir <=" & "'" & Format$(sEndDate, sgSQLDateForm) & "'"
        sVehicles = ""
        
        ReDim ilUseVehicleCodes(0 To 0) As Integer
        gObtainCodes lbcVehAff, ilInclVehicleCodes, ilUseVehicleCodes()        'build array of which codes to incl/excl
        For ilLoop = LBound(ilUseVehicleCodes) To UBound(ilUseVehicleCodes) - 1
            If Trim$(sVehicles) = "" Then
                If ilInclVehicleCodes = True Then                          'include the list
                    sVehicles = "attvefcode IN (" & Str(ilUseVehicleCodes(ilLoop))
                Else                                                        'exclude the list
                    sVehicles = "attvefcode Not IN (" & Str(ilUseVehicleCodes(ilLoop))
                End If
            Else
                sVehicles = sVehicles & "," & Str(ilUseVehicleCodes(ilLoop))
            End If
        Next ilLoop
        If sVehicles <> "" Then
            sVehicles = sVehicles & ")"
        End If
    
        SQLQuery = "SELECT * from"
        SQLQuery = SQLQuery + " shtt INNER JOIN  att ON shttCode = attShfCode "
        SQLQuery = SQLQuery + "LEFT OUTER JOIN mkt ON shttMktCode = mktCode "
        SQLQuery = SQLQuery + " INNER JOIN   VEF_Vehicles ON attVefCode = vefCode "
        SQLQuery = SQLQuery + " Where (" & sDateRange & ")"         '& " and (" & slEnteredRange & ")"     and (" & slStartBetween & ")" & " AND (" & slEndBetween & ")"
'        If rbcInclExpired(1).Value Then  'If True don't show expired agreements
'            SQLQuery = SQLQuery + " AND " & "(attOffAir >=" & "'" & Format$(Date, sgSQLDateForm) & "'" & ")"
'            SQLQuery = SQLQuery + " AND " & "(attDropDate >=" & "'" & Format$(Date, sgSQLDateForm) & "'" & ")"
'        End If
        
'        If rbcDormVeh(1).Value Then  'If True don't show include dormant vehicles
            SQLQuery = SQLQuery + " AND " & " vefstate <> 'D'"
'        End If
    
        If sStationType <> "" Then
            SQLQuery = SQLQuery + " AND (" & sStationType & ")"
        End If
        If sVehicles <> "" Then     '12-13-00
            SQLQuery = SQLQuery + " AND (" & sVehicles & ")"
        End If
'        SQLQuery = SQLQuery + slMulticastOnly + slService
        
        Set rst_Agreement = gSQLSelectCall(SQLQuery)
        llCount = 0
        While Not rst_Agreement.EOF
               
                tmAmr.sOwner = ""
                tmAmr.sVehicleName = ""
                tmAmr.sMarket = ""
                tmAmr.iRank = 0
                tmAmr.sSalesRep = ""
                tmAmr.sServRep = ""
                
                llVefInx = gBinarySearchVef(CLng(rst_Agreement!vefCode))
                If llVefInx <> -1 Then
                    tmAmr.sVehicleName = Trim$(tgVehicleInfo(llVefInx).sVehicleName)
                End If
                
                ilShttInx = gBinarySearchStationInfoByCode(rst_Agreement!shttCode)
                tmAmr.sCallLetters = tgStationInfoByCode(ilShttInx).sCallLetters
                ilFmtInx = gBinarySearchFmt(CLng(tgStationInfoByCode(ilShttInx).iFormatCode))
                If ilFmtInx >= 0 Then
                    tmAmr.sFormat = Trim$(tgFormatInfo(ilFmtInx).sName)
                Else
                    tmAmr.sFormat = ""
                End If
                tmAmr.sSalesRep = Trim$(tgStationInfoByCode(ilShttInx).sZone)
                
                ilMktInx = gBinarySearchMkt(CLng(tgStationInfoByCode(ilShttInx).iMktCode))
                If ilMktInx <> -1 Then
                    tmAmr.sMarket = Trim$(tgMarketInfo(ilMktInx).sName)
                    tmAmr.iRank = tgMarketInfo(ilMktInx).iRank
                End If
    
                llOwnerInx = gBinarySearchOwner(CLng(tgStationInfoByCode(ilShttInx).lOwnerCode))
                If llOwnerInx >= 0 Then
                    tmAmr.sOwner = Trim(tgOwnerInfo(llOwnerInx).sName)
                End If
                               
                tmAmr.lSmtCode = rst_Agreement!attCode
                SQLQuery = "Insert Into amr ( "
                SQLQuery = SQLQuery & "amrGenDate, "
                SQLQuery = SQLQuery & "amrGenTime, "
                SQLQuery = SQLQuery & "amrSmtCode, "
                SQLQuery = SQLQuery & "amrRank, "
                SQLQuery = SQLQuery & "amrMarket, "
                SQLQuery = SQLQuery & "amrOwner, "
                SQLQuery = SQLQuery & "amrVehicleName, "
                SQLQuery = SQLQuery & "amrCallLetters, "
                SQLQuery = SQLQuery & "amrFormat, "
                SQLQuery = SQLQuery & "amrSalesRep "
         
                SQLQuery = SQLQuery & ") "
                SQLQuery = SQLQuery & "Values ( "
                SQLQuery = SQLQuery & "'" & Format$(sgGenDate, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & ", "
                SQLQuery = SQLQuery & tmAmr.lSmtCode & ", "
                SQLQuery = SQLQuery & tmAmr.iRank & ", "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sMarket)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sOwner)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sVehicleName)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sCallLetters)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sFormat)) & "', "
                SQLQuery = SQLQuery & "'" & Trim$(gFixQuote(tmAmr.sSalesRep)) & "' "
            
                SQLQuery = SQLQuery & ") "
                On Error GoTo ErrHand
                
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "AffClusterRpt-cmdReport_Click"
                    Exit Sub
                End If
                On Error GoTo 0
            llCount = llCount + 1
            rst_Agreement.MoveNext
        Wend

        
        iRet = cbcSort.ListIndex        'get the sort selected
        sgCrystlFormula1 = Mid(slSortOption, iRet + 1, 1)    'crystal code for sorting option
        SQLQuery = "SELECT * from amr "
        SQLQuery = SQLQuery + " Inner Join att on amrsmtcode = attcode  INNER JOIN  shtt ON shttCode = attShfCode "
        SQLQuery = SQLQuery & " Where (amrgenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' AND amrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
    
        dFWeek = CDate(sStartDate)
        'StartDate
        sgCrystlFormula2 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
        dFWeek = CDate(sEndDate)
        'EndDate
        sgCrystlFormula3 = "Date(" + Format$(dFWeek, "yyyy") + "," + Format$(dFWeek, "mm") + "," + Format$(dFWeek, "dd") + ")"
              
        gUserActivityLog "E", sgReportListName & ": Prepass"
        frmCrystal.gCrystlReports SQLQuery, ilExportType, ilRptDest, "AfCluster.rpt", "ClusterAgreements"
    
        gUserActivityLog "S", sgReportListName & ": Clear amr"
    
        'remove all the records just printed
        SQLQuery = "DELETE FROM amr "
        SQLQuery = SQLQuery & " WHERE (amrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and amrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
        cnn.BeginTrans
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "AffClusterRpt-cmdReport_Click"
            cnn.RollbackTrans
            Exit Sub
        End If
        cnn.CommitTrans
            
        cmdReport.Enabled = True            'give user back control to gen, done buttons
        cmdDone.Enabled = True
        cmdReturn.Enabled = True
        
        gUserActivityLog "E", sgReportListName & ": Clear amr"
    
        Screen.MousePointer = vbDefault
        
        Exit Sub

    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmClusterRpt-cmdReport"
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmClusterRpt
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
    'Change slSortOption field in cmdReport if new sort added
    cbcSort.AddItem "Format "
    cbcSort.AddItem "Market"
    cbcSort.AddItem "Owner"
    cbcSort.AddItem "Station"
    cbcSort.AddItem "Time Zone"

    cbcSort.ListIndex = 0              'default to Format
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmClusterRpt
    gCenterForm frmClusterRpt

    cmdReport.Enabled = False   'enable only after ALL filters are set  Date: 8/4/2018    FYM
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    
    'Me.Width = Screen.Width / 1.3
    'Me.Height = Screen.Height / 1.3
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    
    frmClusterRpt.Caption = "Affiliate Cluster Report - " & sgClientName
    imChkListBoxIgnore = False
    'SQLQuery = "SELECT vef.vefName from vef WHERE ((vef.vefvefCode = 0 AND vef.vefType = 'C') OR vef.vefType = 'L' OR vef.vefType = 'A')"
    'SQLQuery = SQLQuery + " ORDER BY vef.vefName"
    'Set rst = gSQLSelectCall(SQLQuery)
    'While Not rst.EOF
    '    grdVehAff.AddItem "" & rst(0).Value & ""
    '    rst.MoveNext
    'Wend
    slDate = Format$(gNow(), "m/d/yyyy")
    Do While Weekday(slDate, vbSunday) <> vbMonday
        slDate = DateAdd("d", -1, slDate)
    Loop
    CalOnAirDate.Text = slDate
'    CalOffAirDate.Text = DateAdd("d", 6, slDate)
    CalOffAirDate.Text = CalOnAirDate.Text              'default to todays date
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        ''grdVehAff.AddItem "" & Trim$(tgVehicleInfo(iLoop).sVehicle) & "|" & tgVehicleInfo(iLoop).iCode
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehAff.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehAff.ItemData(lbcVehAff.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
    
    chkListBox.Value = 0    'chged from false to 0 10-22-99
    
    CalOnAirDate.ZOrder (0)
    CalOffAirDate.ZOrder (0)

    gPopExportTypes cboFileType     '3-15-04

    cboFileType.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Agreement.Close
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    Set frmClusterRpt = Nothing
End Sub

Private Sub lbcVehAff_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If chkListBox.Value = vbChecked Then
        imChkListBoxIgnore = True
        chkListBox.Value = vbUnchecked    'chged from false to 0 10-22-99
        imChkListBoxIgnore = False
    End If

    'Enable GENERATE REPORT when ALL filters are set    Date: 8/14/2018  FYM
    mEnableGenerateReportButton
End Sub
Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to pdf
    Else
        cboFileType.Enabled = False
    End If
End Sub
